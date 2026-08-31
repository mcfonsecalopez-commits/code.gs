/**
 * Entregables.gs
 * Responsabilidad única: convertir consultoría(s) aprobada(s) en un
 * entregable para el cliente, y exportar el diagnóstico crudo.
 *
 * MVP: genera PDF (por defecto) u HTML, guardados como archivo en Drive
 * para tener una URL persistente y compartible. La conversión a PDF usa
 * Utilities.newBlob(html,...).getAs(MimeType.PDF) — soporta CSS básico,
 * por eso estilosEntregable_() evita flexbox/grid. La arquitectura queda
 * lista para sumar Google Docs / Slides sin tocar Consultoria.gs ni
 * Code.gs: basta con implementar generarEntregableDocs_() /
 * generarEntregableSlides_() siguiendo la misma firma y despachar por
 * `tipo` en generarEntregable().
 */

/**
 * Botón "GENERAR ENTREGABLE" (sección 14 del pedido).
 * @param {string} clienteId
 * @param {Array<string>} consultoriaIds  deben estar en estado Aprobada
 * @param {string} tipo  'PDF' (por defecto) | 'HTML'
 */
function generarEntregable(clienteId, consultoriaIds, tipo) {
  const cliente = obtenerCliente_(clienteId);
  const consultorias = consultoriaIds.map(function (id) { return obtenerConsultoria_(id); });

  const noAprobadas = consultorias.filter(function (c) { return c.estado !== ESTADOS_CONSULTORIA.APROBADA; });
  if (noAprobadas.length > 0) {
    throw new AppError('VALIDACION',
      'No se puede generar el entregable: hay consultorías sin aprobar (' +
      noAprobadas.map(function (c) { return c.modulo; }).join(', ') + ').');
  }

  const tipoFinal = tipo || 'PDF';
  let url;
  switch (tipoFinal) {
    case 'HTML':
      url = generarEntregableHTML_(cliente, consultorias);
      break;
    case 'PDF':
    default:
      url = generarEntregablePDF_(cliente, consultorias);
  }

  consultorias.forEach(function (c) { cambiarEstadoConsultoria_(c.id, ESTADOS_CONSULTORIA.ENTREGADA); });
  return guardarEntregable_(clienteId, consultoriaIds, tipoFinal, url);
}

function generarEntregableHTML_(cliente, consultorias) {
  const html = construirHtmlEntregable_(cliente, consultorias);
  const carpeta = obtenerCarpetaEntregables_();
  const nombre = 'Consultoria COE Talento y Cultura - ' + cliente.nombre + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-5', 'yyyy-MM-dd HHmm');
  const archivo = carpeta.createFile(nombre + '.html', html, MimeType.HTML);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return archivo.getUrl();
}

/**
 * Convierte el mismo HTML del entregable a PDF (Utilities.newBlob(...).getAs(MimeType.PDF))
 * y lo guarda como archivo en Drive. La conversión de Apps Script soporta CSS básico
 * (por eso estilosEntregable_() evita flexbox/grid), suficiente para un documento
 * de texto con secciones como este.
 */
function generarEntregablePDF_(cliente, consultorias) {
  const html = construirHtmlEntregable_(cliente, consultorias);
  return guardarHtmlComoPDF_(html, obtenerCarpetaEntregables_(),
    'Consultoria COE Talento y Cultura - ' + cliente.nombre + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-5', 'yyyy-MM-dd HHmm'));
}

function obtenerCarpetaEntregables_() {
  const nombreCarpeta = 'Entregables COE Talento y Cultura';
  const it = DriveApp.getFoldersByName(nombreCarpeta);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(nombreCarpeta);
}

/**
 * Botón "Descargar diagnóstico (PDF)" del Paso 5 del wizard.
 * A diferencia del entregable, este PDF se puede generar ANTES de llamar a la
 * IA: usa únicamente el diagnóstico calculado por reglas (madurez, necesidades,
 * dolores, oportunidades semilla) — útil si el COI quiere dejar registro del
 * diagnóstico aunque todavía no apruebe la consultoría completa.
 * @param {string} clienteId
 * @param {Array<{modulo: string, diagnosticoJson: Object}>} diagnosticosPorModulo
 */
function generarDiagnosticoPDF(clienteId, diagnosticosPorModulo) {
  if (!diagnosticosPorModulo || diagnosticosPorModulo.length === 0) {
    throw new AppError('DATOS_INCOMPLETOS', 'No hay diagnóstico para exportar. Completa el Paso 4 primero.');
  }
  const cliente = obtenerCliente_(clienteId);
  const html = construirHtmlDiagnostico_(cliente, diagnosticosPorModulo);
  const archivo = guardarHtmlComoPDF_(html, obtenerCarpetaDiagnosticos_(),
    'Diagnostico COE Talento y Cultura - ' + cliente.nombre + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-5', 'yyyy-MM-dd HHmm'));
  return archivo;
}

function obtenerCarpetaDiagnosticos_() {
  const nombreCarpeta = 'Diagnosticos COE Talento y Cultura';
  const it = DriveApp.getFoldersByName(nombreCarpeta);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(nombreCarpeta);
}

/** Helper compartido: convierte un string HTML a PDF y lo guarda en `carpeta`. Devuelve la URL. */
function guardarHtmlComoPDF_(html, carpeta, nombreSinExtension) {
  let blob;
  try {
    blob = Utilities.newBlob(html, 'text/html', nombreSinExtension + '.html').getAs(MimeType.PDF);
  } catch (e) {
    throw new AppError('ERROR_DESCONOCIDO', 'No se pudo convertir el documento a PDF: ' + e.message);
  }
  blob.setName(nombreSinExtension + '.pdf');
  const archivo = carpeta.createFile(blob);
  archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return archivo.getUrl();
}

function construirHtmlDiagnostico_(cliente, diagnosticosPorModulo) {
  const secciones = diagnosticosPorModulo.map(function (d) {
    const dj = d.diagnosticoJson || {};
    return '<section class="modulo">' +
      '<div class="modulo-encabezado"><h2>' + escaparHtml_(d.modulo) + '</h2></div>' +
      '<p><strong>Nivel de madurez:</strong> ' + escaparHtml_(dj.madurez || 'Sin determinar') + '</p>' +
      '<h3>Necesidades detectadas</h3>' + listaHtml_(dj.principales_necesidades) +
      '<h3>Dolores detectados</h3>' + listaHtml_(dj.dolores) +
      '<h3>Oportunidades (semilla)</h3>' + listaHtml_(dj.oportunidades) +
      '</section>';
  }).join('\n');

  const meta = escaparHtml_(cliente.industria) + ' &middot; ' + escaparHtml_(String(cliente.colaboradores)) +
    ' colaboradores &middot; ' + new Date().toLocaleDateString('es-CO');

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<title>Diagnóstico COE Talento y Cultura - ' + escaparHtml_(cliente.nombre) + '</title>' +
    '<style>' + estilosEntregable_() + '</style></head><body>' +
    construirPortada_('Diagnóstico de Consultoría', cliente, meta) +
    '<div class="contenido">' + secciones + '</div>' +
    '<footer><p>Diagnóstico calculado por reglas a partir de las respuestas del cliente (sin intervención de IA) el ' +
    new Date().toLocaleDateString('es-CO') + '.</p></footer>' +
    '</body></html>';
}

function listaHtml_(arr) {
  return '<ul>' + (arr || []).map(function (x) { return '<li>' + escaparHtml_(x) + '</li>'; }).join('') + '</ul>';
}

function construirHtmlEntregable_(cliente, consultorias) {
  const secciones = consultorias.map(function (c) { return construirSeccionModulo_(c.modulo, c.contenido_json); }).join('\n');
  const meta = escaparHtml_(cliente.industria) + ' &middot; ' + escaparHtml_(String(cliente.colaboradores)) +
    ' colaboradores &middot; ' + escaparHtml_(cliente.tamano) + ' &middot; ' + new Date().toLocaleDateString('es-CO');

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<title>Consultoría COE Talento y Cultura - ' + escaparHtml_(cliente.nombre) + '</title>' +
    '<style>' + estilosEntregable_() + '</style></head><body>' +
    construirPortada_('Consultoría de Implementación', cliente, meta) +
    '<div class="contenido">' +
    '<div class="titulo-seccion-general">Resumen por módulo</div>' +
    construirIndiceModulos_(consultorias) +
    secciones +
    '</div>' +
    '<footer><p>Generado por COE Talento y Cultura el ' + new Date().toLocaleDateString('es-CO') + '.</p></footer>' +
    '</body></html>';
}

/** Portada compartida por el diagnóstico y el entregable — misma marca, mismo estilo. */
function construirPortada_(tituloDocumento, cliente, lineaMeta) {
  return '<div class="portada">' +
    '<div class="portada-marca">COE Talento y Cultura</div>' +
    '<h1>' + escaparHtml_(tituloDocumento) + '</h1>' +
    '<div class="portada-cliente">' + escaparHtml_(cliente.nombre) + '</div>' +
    '<div class="portada-meta">' + lineaMeta + '</div>' +
    '</div>';
}

/** Tabla de un vistazo con el módulo y su prioridad, justo después de la portada. */
function construirIndiceModulos_(consultorias) {
  const filas = consultorias.map(function (c) {
    const prioridad = c.contenido_json.prioridad || '-';
    return '<tr><td>' + escaparHtml_(c.modulo) + '</td><td><span class="prioridad prioridad-' +
      escaparHtml_(prioridad.toLowerCase()) + '">' + escaparHtml_(prioridad) + '</span></td></tr>';
  }).join('');
  return '<table class="tabla-indice"><thead><tr><th>Módulo</th><th>Prioridad</th></tr></thead><tbody>' + filas + '</tbody></table>';
}

function construirSeccionModulo_(modulo, c) {
  const lista = function (arr) {
    return (arr && arr.length) ? '<ul>' + arr.map(function (x) { return '<li>' + escaparHtml_(x) + '</li>'; }).join('') + '</ul>' : '<p class="vacio">—</p>';
  };
  const propuestas = (c.propuestas_valor || []).map(function (p) {
    return '<div class="propuesta"><strong>' + escaparHtml_(p.necesidad || '') + '</strong><p>' + escaparHtml_(p.propuesta_completa || '') + '</p></div>';
  }).join('') || '<p class="vacio">—</p>';

  return '<section class="modulo">' +
    '<div class="modulo-encabezado"><h2>' + escaparHtml_(modulo) + '</h2>' +
    '<span class="prioridad prioridad-' + escaparHtml_((c.prioridad || '').toLowerCase()) + '">Prioridad ' + escaparHtml_(c.prioridad || '-') + '</span></div>' +
    '<h3>Diagnóstico</h3><p>' + escaparHtml_(c.diagnostico || '') + '</p>' +
    '<h3>Necesidades</h3>' + lista(c.necesidades) +
    '<h3>Oportunidades</h3>' + lista(c.oportunidades) +
    '<h3>Consultoría recomendada</h3><p>' + escaparHtml_(c.consultoria_recomendada || '') + '</p>' +
    '<h3>Configuración recomendada</h3>' + lista(c.configuracion_recomendada) +
    '<h3>Plantillas</h3>' + lista(c.plantillas_sugeridas) +
    '<h3>Comunicaciones</h3>' + lista(c.comunicaciones) +
    '<h3>Tareas</h3>' + lista(c.tareas) +
    '<h3>Flujo recomendado</h3><p>' + escaparHtml_(c.flujo_recomendado || '') + '</p>' +
    '<h3>Buenas prácticas</h3>' + lista(c.buenas_practicas) +
    '<h3>Primeros pasos</h3>' + lista(c.primeros_pasos) +
    '<h3>Propuestas de valor</h3>' + propuestas +
    '</section>';
}

function estilosEntregable_() {
  return '' +
    'body{font-family:"Helvetica Neue",Arial,sans-serif;margin:0;color:#334155;background:#fff;}' +
    '.portada{background-color:#1e3a8a;color:#ffffff;padding:70px 55px 50px;}' +
    '.portada-marca{font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#93c5fd;margin-bottom:26px;font-weight:bold;}' +
    '.portada h1{font-size:30px;margin:0 0 18px;font-weight:800;}' +
    '.portada-cliente{font-size:19px;font-weight:700;margin-bottom:6px;}' +
    '.portada-meta{font-size:12.5px;color:#bfdbfe;}' +
    '.contenido{max-width:760px;margin:0 auto;padding:40px 50px 10px;}' +
    '.titulo-seccion-general{font-size:13px;font-weight:700;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin:0 0 4px;}' +
    '.tabla-indice{width:100%;border-collapse:collapse;margin:14px 0 36px;}' +
    '.tabla-indice th{text-align:left;font-size:10.5px;text-transform:uppercase;letter-spacing:.4px;color:#94a3b8;padding:0 10px 8px;border-bottom:2px solid #e2e8f0;}' +
    '.tabla-indice td{padding:10px;font-size:13px;border-bottom:1px solid #eef1f7;}' +
    '.modulo{margin-bottom:38px;padding-bottom:26px;border-bottom:1px solid #e4e7eb;}' +
    '.modulo-encabezado{border-left:5px solid #2563eb;padding-left:14px;margin-bottom:6px;}' +
    '.modulo-encabezado h2{display:inline-block;margin:0 10px 6px 0;font-size:19px;color:#1e3a8a;}' +
    'h3{font-size:11.5px;text-transform:uppercase;letter-spacing:.5px;color:#2563eb;margin:16px 0 6px;border-bottom:1px solid #dbe4fb;padding-bottom:4px;}' +
    'p{font-size:13.5px;line-height:1.65;margin:0 0 8px;}' +
    'ul{padding-left:18px;margin:4px 0 10px;} li{font-size:13.5px;line-height:1.6;margin-bottom:4px;}' +
    '.prioridad{display:inline-block;font-size:11px;font-weight:700;padding:3px 12px;border-radius:12px;color:#fff;vertical-align:middle;}' +
    '.prioridad-alta{background:#dc2626;} .prioridad-media{background:#d97706;} .prioridad-baja{background:#16a34a;}' +
    '.propuesta{background:#eff6ff;border-left:4px solid #2563eb;padding:14px 16px;border-radius:0 8px 8px 0;margin-bottom:10px;}' +
    '.propuesta strong{color:#1e3a8a;display:block;margin-bottom:4px;font-size:13.5px;} .propuesta p{margin:0;}' +
    '.vacio{color:#94a3b8;font-style:italic;}' +
    'footer{padding:26px 50px;background:#f7f9fe;color:#94a3b8;font-size:11px;text-align:center;border-top:1px solid #e4e7eb;}';
}

function escaparHtml_(texto) {
  return String(texto == null ? '' : texto)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

// --- Ganchos para futuras implementaciones (no usados en el MVP) ---
// function generarEntregableDocs_(cliente, consultorias) { /* DocumentApp.create(...) */ }
// function generarEntregableSlides_(cliente, consultorias) { /* SlidesApp.create(...) */ }
