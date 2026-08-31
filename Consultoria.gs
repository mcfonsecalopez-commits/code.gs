/**
 * Consultoria.gs
 * Responsabilidad única: orquestar la generación de consultoría.
 * Es el único archivo que conoce el orden completo:
 *   validar módulos -> tomar diagnóstico -> consultar KB -> armar prompt
 *   -> llamar IA -> validar/parsear -> guardar -> (más tarde) aprobar.
 *
 * No accede a Sheets directamente (usa Database.gs) ni construye prompts
 * directamente (usa Prompts.gs) ni llama HTTP directamente (usa AI.gs).
 */

/**
 * Botón "GENERAR CONSULTORÍA COI" (sección 13 del pedido).
 * @param {string} clienteId
 * @param {Array<{modulo: string, diagnosticoId: string}>} diagnosticosPorModulo
 * @return {Array<Object>} una consultoría (registro de CONSULTORIAS con contenido_json) por módulo
 */
function generarConsultoriaCompleta(clienteId, diagnosticosPorModulo) {
  const cliente = obtenerCliente_(clienteId); // AppError VALIDACION si no existe

  validarModulosContratados_(cliente, diagnosticosPorModulo);

  const diagnosticoGeneral = construirDiagnosticoGeneral(cliente);
  const resultados = [];
  const errores = [];

  diagnosticosPorModulo.forEach(function (item) {
    try {
      resultados.push(generarConsultoriaModulo_(cliente, diagnosticoGeneral, item.modulo, item.diagnosticoId));
    } catch (e) {
      errores.push({ modulo: item.modulo, error: (e instanceof AppError) ? e.message : String(e) });
    }
  });

  if (resultados.length === 0) {
    throw new AppError('IA_ERROR', 'No se pudo generar la consultoría de ningún módulo. Detalle: ' +
      errores.map(function (e) { return e.modulo + ': ' + e.error; }).join(' | '));
  }

  return { generadas: resultados, errores: errores };
}

function validarModulosContratados_(cliente, diagnosticosPorModulo) {
  if (!cliente.modulos_contratados || cliente.modulos_contratados.length === 0) {
    throw new AppError('VALIDACION', 'El cliente "' + cliente.nombre + '" no tiene módulos contratados registrados.');
  }
  if (!diagnosticosPorModulo || diagnosticosPorModulo.length === 0) {
    throw new AppError('DATOS_INCOMPLETOS', 'No hay diagnósticos para generar consultoría. Completa el paso de diagnóstico primero.');
  }
  const noContratados = diagnosticosPorModulo
    .map(function (d) { return d.modulo; })
    .filter(function (m) { return cliente.modulos_contratados.indexOf(m) === -1; });
  if (noContratados.length > 0) {
    throw new AppError('VALIDACION',
      'Se intentó generar consultoría para módulo(s) no contratados por el cliente: ' + noContratados.join(', ') +
      '. La aplicación no genera consultoría para módulos no contratados (regla de negocio).');
  }
}

function generarConsultoriaModulo_(cliente, diagnosticoGeneral, modulo, diagnosticoId) {
  const diagnosticoRow = obtenerDiagnostico_(diagnosticoId);
  const diagnosticoModulo = diagnosticoRow.diagnostico_json;
  if (!diagnosticoModulo || !diagnosticoModulo.madurez) {
    throw new AppError('DATOS_INCOMPLETOS', 'El diagnóstico del módulo ' + modulo + ' está incompleto.');
  }

  const conocimiento = getConocimientoRelevante(modulo, cliente.industria, cliente.tamano, diagnosticoModulo.madurez);
  const plantillas = getPlantillasRelevantes(modulo, cliente.industria);
  const propuestasBase = getPropuestasValorBase(modulo, diagnosticoModulo.principales_necesidades);

  const prompt = construirPromptModulo(cliente, diagnosticoGeneral, diagnosticoModulo, conocimiento, plantillas, propuestasBase);
  const textoIA = llamarIA(prompt);
  const contenido = parsearRespuestaJSON(textoIA);

  validarContenidoConsultoria_(contenido, modulo);

  // Conservamos el diagnóstico de reglas dentro del contenido guardado: así
  // la UI y el entregable pueden mostrar "lo que dijo el cliente" separado
  // de "lo que generó la IA", sin una segunda consulta.
  contenido.diagnostico_estructurado = diagnosticoModulo;
  contenido.modulo = modulo;

  return guardarConsultoria_(cliente.id, modulo, diagnosticoId, contenido, 1);
}

/**
 * Valida que la IA haya devuelto el esquema mínimo. Los campos de texto/array
 * ausentes se rellenan con vacío en vez de fallar duro, PERO los campos
 * estructuralmente críticos (propuestas_valor, trazabilidad) si vienen mal
 * formados sí se consideran error de generación (sección 22 del pedido).
 */
function validarContenidoConsultoria_(contenido, modulo) {
  if (!contenido || typeof contenido !== 'object') {
    throw new AppError('IA_ERROR', 'La IA no devolvió un objeto de consultoría válido para ' + modulo + '.');
  }

  const camposTexto = ['diagnostico', 'consultoria_recomendada', 'flujo_recomendado', 'prioridad'];
  camposTexto.forEach(function (c) { if (typeof contenido[c] !== 'string') contenido[c] = contenido[c] ? String(contenido[c]) : ''; });

  const camposArray = ['necesidades', 'oportunidades', 'configuracion_recomendada', 'plantillas_sugeridas',
    'comunicaciones', 'tareas', 'buenas_practicas', 'primeros_pasos'];
  camposArray.forEach(function (c) { if (!Array.isArray(contenido[c])) contenido[c] = []; });

  if (!Array.isArray(contenido.propuestas_valor) || contenido.propuestas_valor.length < CONFIG.MIN_PROPUESTAS_VALOR) {
    throw new AppError('IA_ERROR',
      'La IA no generó las ' + CONFIG.MIN_PROPUESTAS_VALOR + '-' + CONFIG.MAX_PROPUESTAS_VALOR +
      ' propuestas de valor requeridas para ' + modulo + '. Intenta generar nuevamente.');
  }
  if (contenido.propuestas_valor.length > CONFIG.MAX_PROPUESTAS_VALOR) {
    contenido.propuestas_valor = contenido.propuestas_valor.slice(0, CONFIG.MAX_PROPUESTAS_VALOR);
  }

  if (!Array.isArray(contenido.trazabilidad)) contenido.trazabilidad = [];

  if (['Alta', 'Media', 'Baja'].indexOf(contenido.prioridad) === -1) contenido.prioridad = 'Media';
}

// -------------------------- Revisión y aprobación --------------------------

/** Paso 7 del wizard: el COI edita el contenido antes de aprobar. */
function guardarEdicionConsultoria(consultoriaId, contenidoEditado) {
  validarContenidoConsultoria_(contenidoEditado, contenidoEditado.modulo || '');
  return actualizarConsultoria_(consultoriaId, contenidoEditado, ESTADOS_CONSULTORIA.EN_REVISION);
}

/** Paso 8 del wizard: aprobación final. Guarda la "versión final". */
function aprobarConsultoria(consultoriaId, contenidoFinal) {
  if (contenidoFinal) {
    validarContenidoConsultoria_(contenidoFinal, contenidoFinal.modulo || '');
    return actualizarConsultoria_(consultoriaId, contenidoFinal, ESTADOS_CONSULTORIA.APROBADA);
  }
  cambiarEstadoConsultoria_(consultoriaId, ESTADOS_CONSULTORIA.APROBADA);
  return obtenerConsultoria_(consultoriaId);
}
