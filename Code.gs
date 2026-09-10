/**
 * Code.gs
 * Responsabilidad única: punto de entrada web (doGet) y la capa delgada
 * de funciones expuestas a la UI vía google.script.run.
 *
 * Convención: TODA función api* devuelve siempre { success: true, data }
 * o { success: false, codigo, mensaje }. La UI nunca recibe una excepción
 * cruda de Apps Script (que perdería el mensaje claro) — ver JsClient.html
 * -> llamarServidor().
 */

function doGet(e) {
  return HtmlService.createTemplateFromFile('UI')
    .evaluate()
    .setTitle('Célula Talento y Cultura')
    .setFaviconUrl('https://www.gstatic.com/images/branding/product/1x/apps_script_48dp.png')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(nombreArchivo) {
  return HtmlService.createHtmlOutputFromFile(nombreArchivo).getContent();
}

function ejecutarSeguro_(fn) {
  try {
    // sanearParaCliente_ convierte cualquier Date anidado a texto ISO antes de
    // cruzar hacia el navegador: google.script.run puede perder la respuesta
    // completa (llega undefined, sin ningún error) cuando el objeto devuelto
    // trae un Date anidado — ver Database.gs, todos los registros con fecha.
    return { success: true, data: sanearParaCliente_(fn()) };
  } catch (e) {
    if (e instanceof AppError) {
      return { success: false, codigo: e.codigo, mensaje: e.message };
    }
    // Error no anticipado: no exponemos stack trace crudo, pero sí un mensaje útil.
    return { success: false, codigo: 'ERROR_DESCONOCIDO', mensaje: 'Ocurrió un error inesperado: ' + e.message };
  }
}

/**
 * Recorre recursivamente cualquier valor devuelto a la UI y convierte los
 * objetos Date a texto ISO (JsClient.html los vuelve a interpretar con
 * `new Date(texto)` donde hace falta mostrarlos — ver formatearFecha()).
 * Necesario porque google.script.run puede devolver `undefined` al navegador,
 * sin lanzar ningún error, cuando el valor de retorno trae un Date anidado.
 */
function sanearParaCliente_(valor) {
  if (valor instanceof Date) return valor.toISOString();
  if (Array.isArray(valor)) return valor.map(sanearParaCliente_);
  if (valor && typeof valor === 'object') {
    const limpio = {};
    Object.keys(valor).forEach(function (k) { limpio[k] = sanearParaCliente_(valor[k]); });
    return limpio;
  }
  return valor;
}

// --------------------------- Setup / catálogos ---------------------------

function apiGetModulos() {
  return ejecutarSeguro_(function () { return getModulosActivos(); });
}

function apiGetArbolPreguntas(modulo) {
  return ejecutarSeguro_(function () { return getArbolPreguntas(modulo); });
}

// -------------------------------- Cliente --------------------------------

function apiGuardarCliente(datosFormulario) {
  return ejecutarSeguro_(function () {
    validarCamposCliente_(datosFormulario);
    return guardarCliente_(datosFormulario);
  });
}

function apiObtenerCliente(clienteId) {
  return ejecutarSeguro_(function () { return obtenerCliente_(clienteId); });
}

/** Alimenta el selector de clientes de la pestaña "Consultoría y Gestión de Entregables". */
function apiGetClientesConResumen() {
  return ejecutarSeguro_(function () { return obtenerClientesConResumen_(); });
}

/** Diagnósticos ya guardados de un cliente — para retomar su consultoría sin repetir el Paso 4. */
function apiObtenerDiagnosticosCliente(clienteId) {
  return ejecutarSeguro_(function () { return obtenerDiagnosticosPorCliente_(clienteId); });
}

function validarCamposCliente_(d) {
  const requeridos = ['nombre', 'pais', 'industria', 'categoria', 'tamano', 'culturaFrase'];
  const faltantes = requeridos.filter(function (campo) { return !d[campo] && d[campo] !== 0; });
  if (faltantes.length > 0) {
    throw new AppError('DATOS_INCOMPLETOS', 'Faltan campos obligatorios de información comercial: ' + faltantes.join(', '));
  }
  if (!d.modulosContratados || d.modulosContratados.length === 0) {
    throw new AppError('VALIDACION', 'Debes seleccionar al menos un módulo contratado.');
  }
}

// ------------------------------ Diagnóstico -------------------------------

function apiProcesarDiagnostico(clienteId, modulo, respuestas) {
  return ejecutarSeguro_(function () {
    const diagnosticoJson = procesarRespuestasModulo(modulo, respuestas);
    return guardarDiagnostico_(clienteId, modulo, respuestas, diagnosticoJson);
  });
}

// ------------------------------ Consultoría -------------------------------

function apiGenerarConsultoria(clienteId, diagnosticosPorModulo) {
  return ejecutarSeguro_(function () { return generarConsultoriaCompleta(clienteId, diagnosticosPorModulo); });
}

function apiGuardarEdicionConsultoria(consultoriaId, contenidoEditado) {
  return ejecutarSeguro_(function () { return guardarEdicionConsultoria(consultoriaId, contenidoEditado); });
}

function apiAprobarConsultoria(consultoriaId, contenidoFinal) {
  return ejecutarSeguro_(function () { return aprobarConsultoria(consultoriaId, contenidoFinal); });
}

function apiObtenerConsultoriasCliente(clienteId) {
  return ejecutarSeguro_(function () { return obtenerConsultoriasPorCliente_(clienteId); });
}

// ------------------------------- Entregable --------------------------------

function apiGenerarEntregable(clienteId, consultoriaIds, tipo) {
  return ejecutarSeguro_(function () { return generarEntregable(clienteId, consultoriaIds, tipo || 'PDF'); });
}

/** Paso 5 del wizard: descargar el diagnóstico calculado por reglas (sin IA) en PDF. */
function apiGenerarDiagnosticoPDF(clienteId, diagnosticosPorModulo) {
  return ejecutarSeguro_(function () { return generarDiagnosticoPDF(clienteId, diagnosticosPorModulo); });
}

/** Checklist de seguimiento de la consultoría guiada (sección 5 del rediseño). */
function apiGuardarSeguimientoConsultoria(consultoriaId, seguimiento) {
  return ejecutarSeguro_(function () { return guardarSeguimientoConsultoria(consultoriaId, seguimiento); });
}

// -------------------------------- Kick Off COE ---------------------------------
// Nueva etapa intermedia Diagnóstico -> Kick Off COE -> Consultoría (sección 3).

/** Botón "Estructurar con IA": solo interpreta el texto, no lo guarda todavía. */
function apiEstructurarKickOffIA(notasGemini) {
  return ejecutarSeguro_(function () { return estructurarKickOffConIA(notasGemini); });
}

/** Botón "Guardar Kick Off": guarda la versión que el COE ya revisó/editó. */
function apiGuardarKickOff(clienteId, datos, notasGemini) {
  return ejecutarSeguro_(function () { return guardarKickOff(clienteId, datos, notasGemini); });
}

/** Para prellenar la vista si el cliente ya tiene un Kick Off guardado (null si no hay ninguno). */
function apiObtenerKickOffCliente(clienteId) {
  return ejecutarSeguro_(function () { return obtenerKickOff(clienteId); });
}

// ------------------------- Radiografía del Cliente -------------------------
// Flujo Información Comercial -> Preguntas Orientadoras -> Kick Off -> Notas
// Gemini -> Radiografía (ver KickOff.gs).

/** Botón "🩻 Generar Radiografía": cruza Diagnóstico Comercial + Notas de Kick Off. */
function apiGenerarRadiografia(clienteId, notasGemini, notasGeminiUrl) {
  return ejecutarSeguro_(function () { return generarRadiografia(clienteId, notasGemini, notasGeminiUrl); });
}

/** Para prellenar la vista si el cliente ya tiene una Radiografía guardada (null si no hay ninguna). */
function apiObtenerRadiografiaCliente(clienteId) {
  return ejecutarSeguro_(function () { return obtenerRadiografia(clienteId); });
}

/** Trae el texto de un Google Doc por su URL (notas de Gemini de Google Meet). */
function apiObtenerNotasDesdeUrl(url) {
  return ejecutarSeguro_(function () { return obtenerNotasDesdeUrl(url); });
}

/** Auto-detección de cliente a partir del texto/título de las notas — nunca decide sola, solo sugiere. */
function apiBuscarClienteParaNotas(texto, titulo) {
  return ejecutarSeguro_(function () { return buscarClienteParaNotas(texto, titulo); });
}

/** Botón "📄 Descargar Radiografía (PDF)" dentro del resultado de Kick Off. */
function apiGenerarRadiografiaPDF(clienteId) {
  return ejecutarSeguro_(function () { return generarRadiografiaPDF(clienteId); });
}

/** Columna "Diagnóstico Comercial" (📝) del Historial: genera el PDF a partir de lo ya guardado en Sheets. */
function apiGenerarDiagnosticoComercialPDF(clienteId) {
  return ejecutarSeguro_(function () { return generarDiagnosticoComercialPDF(clienteId); });
}

// ----------------------------- Biblioteca de Magia -----------------------------
// Catálogo de AppScripts del equipo COE (sección 7 del rediseño).

function apiGetHerramientas() {
  return ejecutarSeguro_(function () { return obtenerHerramientas_(); });
}

function apiGetHerramientasPorModulos(modulos) {
  return ejecutarSeguro_(function () { return obtenerHerramientasPorModulos_(modulos); });
}

// ------------------- Playbook y Presentación (Gema COE) ---------------------
// Paso 9 (Entregables): además del PDF de la consultoría, el COE puede generar
// un Playbook Técnico (y, si quiere, también la Presentación Ejecutiva a partir
// de una plantilla de Slides) con el contenido que trae de su Gema de Gemini
// — ver GemaIntegracion.gs.

/** Botón "🚀 Generar Playbook y Enviarlo a mi Correo". */
function apiGenerarPlaybookYNotificar(clienteId, datosGema) {
  return ejecutarSeguro_(function () { return generarPlaybookYNotificar(clienteId, datosGema); });
}

/** Botón "⬇ Descargar Entregables (PDF + PPT)". */
function apiGenerarEntregablesDuales(clienteId, datosGema) {
  return ejecutarSeguro_(function () { return generarEntregablesDuales(clienteId, datosGema); });
}

// -------------------------------- Historial ---------------------------------

function apiGetHistorial() {
  return ejecutarSeguro_(function () { return obtenerHistorialClientes_(); });
}

function apiGetConsultoria(consultoriaId) {
  return ejecutarSeguro_(function () { return obtenerConsultoria_(consultoriaId); });
}
