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
    .setTitle('COE Talento y Cultura')
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
  const requeridos = ['nombre', 'pais', 'industria', 'categoria', 'colaboradores', 'tamano'];
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

// -------------------------------- Historial ---------------------------------

function apiGetHistorial() {
  return ejecutarSeguro_(function () { return obtenerHistorial_(); });
}

function apiGetConsultoria(consultoriaId) {
  return ejecutarSeguro_(function () { return obtenerConsultoria_(consultoriaId); });
}
