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

/**
 * Botón "📝" (columna "Diagnóstico Comercial") del Historial: a diferencia de
 * generarDiagnosticoPDF() del Paso 5 (que recibe el diagnóstico desde el estado
 * del wizard abierto en el navegador), este genera el mismo PDF a partir de lo
 * que YA quedó guardado en Sheets — así funciona para cualquier cliente desde
 * el Historial, sin tener que reabrir su wizard.
 * @param {string} clienteId
 * @return {string} URL del PDF guardado en Drive
 */
function generarDiagnosticoComercialPDF(clienteId) {
  const cliente = obtenerCliente_(clienteId);
  const diagnosticos = obtenerDiagnosticosPorCliente_(clienteId);
  if (!diagnosticos || diagnosticos.length === 0) {
    throw new AppError('DATOS_INCOMPLETOS', 'Este cliente todavía no tiene ningún módulo diagnosticado en Diagnóstico Comercial.');
  }
  const diagnosticosPorModulo = diagnosticos.map(function (d) {
    return { modulo: d.modulo, diagnosticoJson: d.diagnostico_json };
  });
  const html = construirHtmlDiagnostico_(cliente, diagnosticosPorModulo);
  return guardarHtmlComoPDF_(html, obtenerCarpetaDiagnosticos_(),
    'Diagnostico Comercial COE Talento y Cultura - ' + cliente.nombre + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-5', 'yyyy-MM-dd HHmm'));
}

function obtenerCarpetaDiagnosticos_() {
  const nombreCarpeta = 'Diagnosticos COE Talento y Cultura';
  const it = DriveApp.getFoldersByName(nombreCarpeta);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(nombreCarpeta);
}

/**
 * Botón "📄 Descargar Radiografía (PDF)" (dentro del resultado de Kick Off).
 * Formato inspirado en la plantilla "Informe Pre Kick Off" de Buk que compartió
 * la usuaria: portada con el nombre del cliente, y el cuerpo en cajas con una
 * barra de título en navy Buk y contenido con bordes finos — ver
 * estilosRadiografia_() más abajo. La conversión HTML->PDF de Apps Script
 * solo soporta CSS básico (igual que el resto de Entregables.gs), así que se
 * evita flexbox/grid también acá.
 * @param {string} clienteId
 * @return {string} URL del PDF guardado en Drive
 */
function generarRadiografiaPDF(clienteId) {
  const cliente = obtenerCliente_(clienteId);
  const radiografiaRow = obtenerRadiografiaPorCliente_(clienteId);
  if (!radiografiaRow || !radiografiaRow.radiografia) {
    throw new AppError('DATOS_INCOMPLETOS', 'Este cliente todavía no tiene una Radiografía generada.');
  }
  const html = construirHtmlRadiografia_(cliente, radiografiaRow.radiografia);
  return guardarHtmlComoPDF_(html, obtenerCarpetaRadiografias_(),
    'Radiografia COE Talento y Cultura - ' + cliente.nombre + ' - ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'GMT-5', 'yyyy-MM-dd HHmm'));
}

function obtenerCarpetaRadiografias_() {
  const nombreCarpeta = 'Radiografias COE Talento y Cultura';
  const it = DriveApp.getFoldersByName(nombreCarpeta);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(nombreCarpeta);
}

/**
 * Radiografía condensada, formato "Informe Pre Kick Off" de Buk: cajas simples
 * con barra de título centrada en navy, agrupando la información relacionada
 * en menos secciones (antes eran 11 cajas independientes; la usuaria pidió
 * bajarle el volumen porque "es muy extensa / muy redundante"). Usa su propia
 * portada y hoja de estilos (estilosRadiografia_(), navy #002B49 / azul #005691,
 * misma paleta que el Diagnóstico Comercial) en vez de la portada genérica
 * compartida con el Entregable de Consultoría.
 */
function construirHtmlRadiografia_(cliente, r) {
  const meta = escaparHtml_(cliente.industria) + ' &middot; ' + escaparHtml_(cliente.tamano) +
    ' &middot; ' + new Date().toLocaleDateString('es-CO');

  const parrafo = function (texto) {
    return texto ? '<p>' + escaparHtml_(texto) + '</p>' : '<p class="vacio">Sin información.</p>';
  };
  const listaCerrada = function (arr) {
    return (arr && arr.length) ? '<ul>' + arr.map(function (x) { return '<li>' + escaparHtml_(x) + '</li>'; }).join('') + '</ul>' : '<p class="vacio">Sin información.</p>';
  };
  const sub = function (titulo, contenidoHtml) {
    return '<p class="rad-subtitulo">' + escaparHtml_(titulo) + '</p>' + contenidoHtml;
  };
  const caja = function (titulo, innerHtml, clase) {
    return '<table class="rad-caja' + (clase ? ' ' + clase : '') + '"><tr><th>' + escaparHtml_(titulo) + '</th></tr>' +
      '<tr><td>' + innerHtml + '</td></tr></table>';
  };

  let alertasHtml = '';
  if (r.alertas_inconsistencia && r.alertas_inconsistencia.length) {
    alertasHtml = caja('⚠️ ALERTAS DE INCONSISTENCIA',
      '<ul>' + r.alertas_inconsistencia.map(function (a) { return '<li>' + escaparHtml_(a) + '</li>'; }).join('') + '</ul>',
      'rad-alerta');
  }

  // CONTEXTO Y ALCANCE: combina lo que antes eran 4 cajas separadas
  // (contexto_cliente, necesidades_principales, objetivos_implementacion, procesos_actuales).
  const contextoHtml = caja('🏢 CONTEXTO Y ALCANCE',
    sub('Contexto del cliente', parrafo(r.contexto_cliente)) +
    sub('Necesidades principales', listaCerrada(r.necesidades_principales)) +
    sub('Objetivos de implementación', listaCerrada(r.objetivos_implementacion)) +
    sub('Procesos actuales', listaCerrada(r.procesos_actuales))
  );

  const modulosHtml = (cliente.modulos_contratados || []).length
    ? caja('📦 MÓDULOS CONTRATADOS', '<table class="rad-tabla-modulos">' + (cliente.modulos_contratados || []).map(function (m) {
        return '<tr><td>' + escaparHtml_(m) + '</td></tr>';
      }).join('') + '</table>')
    : '';

  // HALLAZGOS Y DECISIONES DEL KICK OFF: combina hallazgos_kickoff + definiciones_tomadas + pendientes.
  const hallazgosHtml = caja('🔎 HALLAZGOS Y DECISIONES DEL KICK OFF',
    sub('Hallazgos', listaCerrada(r.hallazgos_kickoff)) +
    sub('Decisiones tomadas', listaCerrada(r.definiciones_tomadas)) +
    sub('Pendientes', listaCerrada(r.pendientes))
  );

  const responsablesHtml = (r.responsables && r.responsables.length)
    ? '<ul>' + r.responsables.map(function (p) { return '<li>' + escaparHtml_(p.nombre || '—') + (p.rol ? ' — ' + escaparHtml_(p.rol) : '') + '</li>'; }).join('') + '</ul>'
    : '<p class="vacio">Sin información.</p>';

  // RIESGOS Y RESPONSABLES: combina riesgos + responsables.
  const riesgosHtml = caja('⚠️ RIESGOS Y RESPONSABLES',
    sub('Riesgos', listaCerrada(r.riesgos)) +
    sub('Responsables', responsablesHtml)
  );

  const configuracionesHtml = caja('🧩 CONFIGURACIONES REQUERIDAS', listaCerrada(r.configuraciones_requeridas));

  // Información por módulo, ahora como una sola tabla de 3 columnas (Hallazgos | Decisiones |
  // Pendientes) en vez de una caja completa por cada módulo — mucho más compacto.
  let porModuloHtml = '';
  if (r.por_modulo) {
    const filas = Object.keys(r.por_modulo).map(function (m) {
      const b = r.por_modulo[m];
      if (!b || (!b.hallazgos.length && !b.decisiones.length && !b.pendientes.length)) return '';
      return '<tr><td class="rad-modulo-nombre">' + escaparHtml_(m) + '</td>' +
        '<td>' + listaCerrada(b.hallazgos) + '</td>' +
        '<td>' + listaCerrada(b.decisiones) + '</td>' +
        '<td>' + listaCerrada(b.pendientes) + '</td></tr>';
    }).join('');
    if (filas) {
      porModuloHtml = caja('INFORMACIÓN POR MÓDULO',
        '<table class="rad-tabla-por-modulo"><thead><tr><th>Módulo</th><th>Hallazgos</th><th>Decisiones</th><th>Pendientes</th></tr></thead><tbody>' + filas + '</tbody></table>');
    }
  }

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<title>Radiografía del Cliente - ' + escaparHtml_(cliente.nombre) + '</title>' +
    '<style>' + estilosRadiografia_() + '</style></head><body>' +
    '<div class="rad-portada"><div class="rad-marca">Célula Talento y Cultura</div>' +
    '<h1>Radiografía del Cliente</h1>' +
    '<div class="rad-portada-cliente">' + escaparHtml_(cliente.nombre) + '</div>' +
    '<div class="rad-portada-meta">' + meta + '</div></div>' +
    '<div class="contenido">' +
    alertasHtml +
    contextoHtml +
    modulosHtml +
    hallazgosHtml +
    riesgosHtml +
    configuracionesHtml +
    porModuloHtml +
    '</div>' +
    '<footer><p>Radiografía generada por Célula Talento y Cultura a partir del Diagnóstico Comercial y las notas de Kick Off, el ' +
    new Date().toLocaleDateString('es-CO') + '.</p></footer>' +
    '</body></html>';
}

/**
 * Línea gráfica propia de la Radiografía — navy #002B49 / azul #005691 (misma
 * paleta Buk que estilosDiagnosticoComercial_()), con barra de título CENTRADA
 * en cada caja, inspirada en la plantilla "Informe Pre Kick Off" que compartió
 * la usuaria. Autocontenida (no depende de estilosEntregable_()) para poder
 * tener su propia portada compacta sin afectar el PDF de Consultoría.
 */
function estilosRadiografia_() {
  return '' +
    'body{font-family:"Helvetica Neue",Arial,sans-serif;margin:0;color:#334155;background:#fff;}' +
    '.rad-portada{background-color:#002B49;color:#fff;padding:50px 55px 38px;text-align:center;}' +
    '.rad-portada .rad-marca{font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#7fb3d9;margin-bottom:18px;font-weight:bold;}' +
    '.rad-portada h1{font-size:26px;margin:0 0 14px;font-weight:800;}' +
    '.rad-portada-cliente{font-size:18px;font-weight:700;margin-bottom:6px;}' +
    '.rad-portada-meta{font-size:12px;color:#bcd6ea;}' +
    '.contenido{max-width:760px;margin:0 auto;padding:34px 50px 10px;}' +
    '.rad-caja{width:100%;border-collapse:collapse;border:1px solid #c9d8e8;border-radius:6px;overflow:hidden;margin-bottom:16px;}' +
    '.rad-caja th{background:#002B49;color:#fff;font-size:12px;letter-spacing:.4px;text-align:center;padding:9px 12px;}' +
    '.rad-caja td{padding:14px 16px;border-top:1px solid #eef1f7;}' +
    '.rad-caja p{margin:0 0 6px;} .rad-caja ul{margin:0 0 6px;padding-left:18px;} .rad-caja li{margin-bottom:5px;font-size:13px;}' +
    '.rad-caja .vacio{color:#94a3b8;font-style:italic;}' +
    '.rad-alerta th{background:#d99e00;}' +
    '.rad-alerta td{background:#fff8e8;}' +
    '.rad-tabla-modulos{width:100%;border-collapse:collapse;}' +
    '.rad-tabla-modulos td{padding:6px 4px;border-bottom:1px solid #eef1f7;font-weight:700;color:#002B49;}' +
    '.rad-tabla-modulos tr:last-child td{border-bottom:none;}' +
    '.rad-subtitulo{font-size:11px;font-weight:800;text-transform:uppercase;color:#005691;margin:10px 0 4px;}' +
    '.rad-subtitulo:first-child{margin-top:0;}' +
    '.rad-tabla-por-modulo{width:100%;border-collapse:collapse;}' +
    '.rad-tabla-por-modulo th{background:#005691;color:#fff;font-size:10.5px;text-transform:uppercase;letter-spacing:.3px;padding:7px 8px;text-align:left;}' +
    '.rad-tabla-por-modulo td{padding:8px;font-size:12.5px;border-bottom:1px solid #eef1f7;vertical-align:top;}' +
    '.rad-tabla-por-modulo td.rad-modulo-nombre{font-weight:700;color:#002B49;white-space:nowrap;}' +
    '.rad-tabla-por-modulo ul{margin:0;padding-left:14px;} .rad-tabla-por-modulo li{font-size:12px;margin-bottom:3px;}' +
    'footer{padding:22px 50px;background:#f7f9fe;color:#94a3b8;font-size:11px;text-align:center;border-top:1px solid #e4e7eb;}';
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

/**
 * PDF del Diagnóstico Comercial — con la línea gráfica oficial de Buk (navy #002B49 /
 * azul #005691, la misma paleta que GemaIntegracion.gs usa para el Playbook Técnico)
 * y trayendo toda la información comercial relevante del cliente (ficha, contexto,
 * procesos y herramientas actuales, equipo y liderazgo, alcance de implementación),
 * además del diagnóstico calculado por reglas de cada módulo. Solo se muestran las
 * cajas y campos que el cliente realmente tiene cargados, para no dejar secciones
 * vacías en el documento.
 */
function construirHtmlDiagnostico_(cliente, diagnosticosPorModulo) {
  const campo = function (etiqueta, valor) {
    if (!valor) return '';
    return '<div class="dc-campo"><strong>' + escaparHtml_(etiqueta) + '</strong><p>' + escaparHtml_(valor) + '</p></div>';
  };
  const cajaCampos = function (titulo, camposHtml) {
    if (!camposHtml) return '';
    return '<table class="dc-caja"><tr><th>' + escaparHtml_(titulo) + '</th></tr><tr><td>' + camposHtml + '</td></tr></table>';
  };
  const listaConVacio = function (arr) {
    return (arr && arr.length) ? '<ul>' + arr.map(function (x) { return '<li>' + escaparHtml_(x) + '</li>'; }).join('') + '</ul>' : '<p class="vacio">Sin información.</p>';
  };

  const fichaHtml =
    campo('Cliente', cliente.nombre) +
    campo('País / Ciudad', cliente.pais) +
    campo('Industria', cliente.industria) +
    campo('Categoría', cliente.categoria) +
    campo('Tamaño', cliente.tamano) +
    campo('Cuenta ejecutiva (Account Manager)', cliente.account_manager) +
    campo('Fecha de registro', cliente.fecha_creacion ? new Date(cliente.fecha_creacion).toLocaleDateString('es-CO') : '');

  const modulosContratados = cliente.modulos_contratados || [];
  const modulosHtml = modulosContratados.length
    ? '<table class="dc-caja"><tr><th>Módulos contratados</th></tr><tr><td><ul class="dc-badges">' +
      modulosContratados.map(function (m) { return '<li>' + escaparHtml_(m) + '</li>'; }).join('') + '</ul></td></tr></table>'
    : '';

  const contextoHtml =
    campo('Contexto comercial', cliente.contexto_comercial) +
    campo('Objetivos y dolores', cliente.objetivos) +
    campo('Motivo de compra de Buk', cliente.motivo_compra_buk) +
    campo('Frase de cultura', cliente.cultura_frase) +
    campo('Información adicional', cliente.info_adicional);

  const procesosHtml =
    campo('Procesos actuales', cliente.procesos_actuales) +
    campo('Herramientas utilizadas', cliente.herramientas_utilizadas) +
    campo('Nivel de automatización', cliente.nivel_automatizacion) +
    campo('¿Usaba otra plataforma?', cliente.usaba_otra_plataforma === 'Sí'
      ? ('Sí — ' + (cliente.otra_plataforma_cual || 'sin detalle'))
      : cliente.usaba_otra_plataforma);

  const equipoHtml =
    campo('Líder de Personas', [cliente.lider_nombre, cliente.lider_celular, cliente.lider_correo].filter(Boolean).join(' · ')) +
    campo('Líder estratégico', [cliente.lider_estrategico_nombre, cliente.lider_estrategico_celular, cliente.lider_estrategico_correo].filter(Boolean).join(' · ')) +
    campo('Sedes / Ubicación', cliente.sedes_ubicacion) +
    campo('Estructura del equipo', cliente.estructura_equipo);

  const alcanceHtml =
    campo('Integración SSO', cliente.integracion_sso) +
    campo('Integración API', cliente.integracion_api) +
    campo('Módulo de inicio deseado', cliente.modulo_inicio_deseado) +
    campo('Fecha fin esperada', cliente.fecha_fin_esperada) +
    campo('Comentarios de tiempos', cliente.comentarios_tiempos) +
    campo('Sesiones grupales', cliente.sesiones_grupales);

  const seccionesModulo = diagnosticosPorModulo.map(function (d) {
    const dj = d.diagnosticoJson || {};
    const madurez = dj.madurez || '';
    return '<table class="dc-caja"><tr><th>' + escaparHtml_(d.modulo) +
      (madurez ? '<span class="dc-madurez dc-madurez-' + escaparHtml_(madurez.toLowerCase()) + '">Madurez ' + escaparHtml_(madurez) + '</span>' : '') +
      '</th></tr><tr><td>' +
      '<div class="dc-campo"><strong>Necesidades detectadas</strong>' + listaConVacio(dj.principales_necesidades) + '</div>' +
      '<div class="dc-campo"><strong>Dolores detectados</strong>' + listaConVacio(dj.dolores) + '</div>' +
      '<div class="dc-campo"><strong>Oportunidades (semilla)</strong>' + listaConVacio(dj.oportunidades) + '</div>' +
      '</td></tr></table>';
  }).join('\n');

  const meta = escaparHtml_(cliente.industria) + ' &middot; ' + escaparHtml_(cliente.tamano) +
    ' &middot; ' + new Date().toLocaleDateString('es-CO');

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<title>Diagnóstico Comercial - ' + escaparHtml_(cliente.nombre) + '</title>' +
    '<style>' + estilosDiagnosticoComercial_() + '</style></head><body>' +
    '<div class="dc-portada"><div class="dc-marca">Célula Talento y Cultura</div>' +
    '<h1>Diagnóstico Comercial</h1>' +
    '<div class="dc-cliente">' + escaparHtml_(cliente.nombre) + '</div>' +
    '<div class="dc-meta">' + meta + '</div></div>' +
    '<div class="contenido">' +
    cajaCampos('🏢 FICHA DEL CLIENTE', fichaHtml) +
    modulosHtml +
    cajaCampos('🎯 CONTEXTO COMERCIAL', contextoHtml) +
    cajaCampos('⚙️ PROCESOS Y HERRAMIENTAS ACTUALES', procesosHtml) +
    cajaCampos('👥 EQUIPO Y LIDERAZGO', equipoHtml) +
    cajaCampos('🧩 ALCANCE DE IMPLEMENTACIÓN', alcanceHtml) +
    (seccionesModulo ? '<div class="dc-seccion-general">DIAGNÓSTICO POR MÓDULO</div>' + seccionesModulo : '') +
    '</div>' +
    '<footer><p>Diagnóstico Comercial calculado por reglas a partir de las respuestas del cliente (sin intervención de IA) el ' +
    new Date().toLocaleDateString('es-CO') + '.</p></footer>' +
    '</body></html>';
}

/**
 * Línea gráfica oficial de Buk (navy #002B49 / azul #005691 — misma paleta que
 * estilosPlaybook_() en GemaIntegracion.gs) para el PDF de Diagnóstico Comercial.
 * Evita flexbox/grid, igual que estilosEntregable_(), porque la conversión
 * HTML→PDF de Apps Script solo soporta CSS básico.
 */
function estilosDiagnosticoComercial_() {
  return '' +
    'body{font-family:"Helvetica Neue",Arial,sans-serif;margin:0;color:#334155;background:#fff;}' +
    '.dc-portada{background-color:#002B49;color:#fff;padding:60px 55px 46px;}' +
    '.dc-portada .dc-marca{font-size:11px;letter-spacing:3px;text-transform:uppercase;color:#7fb3d9;margin-bottom:22px;font-weight:bold;}' +
    '.dc-portada h1{font-size:28px;margin:0 0 16px;font-weight:800;}' +
    '.dc-portada .dc-cliente{font-size:18px;font-weight:700;margin-bottom:6px;}' +
    '.dc-portada .dc-meta{font-size:12.5px;color:#bcd7ec;}' +
    '.contenido{max-width:760px;margin:0 auto;padding:40px 50px 10px;}' +
    '.dc-caja{width:100%;border-collapse:collapse;border:1px solid #cfe0ec;border-radius:6px;overflow:hidden;margin-bottom:18px;}' +
    '.dc-caja th{background:#002B49;color:#fff;font-size:12px;letter-spacing:.4px;text-align:left;padding:10px 14px;}' +
    '.dc-caja td{padding:14px 16px;border-top:1px solid #eef4f9;}' +
    '.dc-caja ul{margin:4px 0 0;padding-left:18px;} .dc-caja li{font-size:13.5px;line-height:1.55;margin-bottom:4px;}' +
    '.dc-caja .vacio{color:#94a3b8;font-style:italic;margin:4px 0 0;}' +
    '.dc-campo{margin-bottom:12px;} .dc-campo:last-child{margin-bottom:0;}' +
    '.dc-campo strong{color:#005691;display:block;font-size:11px;text-transform:uppercase;letter-spacing:.3px;margin-bottom:3px;}' +
    '.dc-campo p{margin:0;font-size:13.5px;line-height:1.55;}' +
    '.dc-badges{margin:0;padding:0;list-style:none;}' +
    '.dc-badges li{display:inline-block;background:#eaf2f8;color:#002B49;font-weight:700;font-size:12px;padding:5px 12px;border-radius:12px;margin:0 6px 6px 0;}' +
    '.dc-madurez{display:inline-block;font-size:11px;font-weight:700;padding:3px 12px;border-radius:12px;color:#fff;vertical-align:middle;margin-left:8px;}' +
    '.dc-madurez-baja{background:#dc2626;} .dc-madurez-media{background:#d97706;} .dc-madurez-alta{background:#16a34a;}' +
    '.dc-seccion-general{font-size:13px;font-weight:800;text-transform:uppercase;letter-spacing:.5px;color:#64748b;margin:10px 0 12px;}' +
    'footer{padding:26px 50px;background:#f7f9fe;color:#94a3b8;font-size:11px;text-align:center;border-top:1px solid #e4e7eb;}';
}

function listaHtml_(arr) {
  return '<ul>' + (arr || []).map(function (x) { return '<li>' + escaparHtml_(x) + '</li>'; }).join('') + '</ul>';
}

function construirHtmlEntregable_(cliente, consultorias) {
  const secciones = consultorias.map(function (c) { return construirSeccionModulo_(c.modulo, c.contenido_json); }).join('\n');
  const meta = escaparHtml_(cliente.industria) + ' &middot; ' + escaparHtml_(cliente.tamano) +
    ' &middot; ' + new Date().toLocaleDateString('es-CO');

  return '<!DOCTYPE html><html lang="es"><head><meta charset="UTF-8">' +
    '<title>Consultoría Célula Talento y Cultura - ' + escaparHtml_(cliente.nombre) + '</title>' +
    '<style>' + estilosEntregable_() + '</style></head><body>' +
    construirPortada_('Consultoría de Implementación', cliente, meta) +
    '<div class="contenido">' +
    '<div class="titulo-seccion-general">Resumen por módulo</div>' +
    construirIndiceModulos_(consultorias) +
    secciones +
    '</div>' +
    '<footer><p>Generado por Célula Talento y Cultura el ' + new Date().toLocaleDateString('es-CO') + '.</p></footer>' +
    '</body></html>';
}

/** Portada compartida por el diagnóstico y el entregable — misma marca, mismo estilo. */
function construirPortada_(tituloDocumento, cliente, lineaMeta) {
  return '<div class="portada">' +
    '<div class="portada-marca">Célula Talento y Cultura</div>' +
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
