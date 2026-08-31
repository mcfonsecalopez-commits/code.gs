/**
 * Database.gs
 * Responsabilidad única: leer y escribir en Google Sheets.
 * Expone (a) helpers genéricos basados en encabezados, y (b) funciones
 * específicas por entidad (Clientes, Diagnósticos, Consultorías, Entregables).
 *
 * Ningún otro archivo debe llamar SpreadsheetApp directamente para las
 * hojas transaccionales: todo pasa por acá, así el manejo de errores de
 * Sheets (hoja bloqueada, fuera de cuota, hoja inexistente) está en un
 * solo lugar (sección 22 del pedido).
 */

// ------------------------- Helpers genéricos --------------------------

function getSheet_(nombreHoja, spreadsheet) {
  try {
    const ss = spreadsheet || getSpreadsheet();
    const sheet = ss.getSheetByName(nombreHoja);
    if (!sheet) {
      throw new AppError('SHEETS_ERROR',
        'No existe la hoja "' + nombreHoja + '". Ejecuta inicializarBaseDeDatos() ' +
        'desde SetupSheets.gs o revisa el nombre de la pestaña.');
    }
    return sheet;
  } catch (e) {
    if (e instanceof AppError) throw e;
    throw new AppError('SHEETS_ERROR', 'Error accediendo a la hoja "' + nombreHoja + '": ' + e.message);
  }
}

/** Lee toda la hoja y la devuelve como array de objetos {encabezado: valor}. */
function getAllRows_(nombreHoja, spreadsheet) {
  const sheet = getSheet_(nombreHoja, spreadsheet);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    if (values[i].every(function (c) { return c === '' || c === null; })) continue; // saltar filas vacías
    const obj = {};
    headers.forEach(function (h, idx) { obj[h] = values[i][idx]; });
    obj._row = i + 1; // 1-indexed, útil para updates
    rows.push(obj);
  }
  return rows;
}

function getRowById_(nombreHoja, id, spreadsheet) {
  const rows = getAllRows_(nombreHoja, spreadsheet);
  return rows.find(function (r) { return String(r.id) === String(id); }) || null;
}

/** Agrega una fila respetando el orden de encabezados existente. */
function appendRow_(nombreHoja, objeto, spreadsheet) {
  const sheet = getSheet_(nombreHoja, spreadsheet);
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const fila = headers.map(function (h) { return objeto.hasOwnProperty(h) ? objeto[h] : ''; });
  sheet.appendRow(fila);
  return objeto;
}

/** Actualiza una fila existente por id, solo las columnas presentes en `cambios`. */
function updateRowById_(nombreHoja, id, cambios, spreadsheet) {
  const sheet = getSheet_(nombreHoja, spreadsheet);
  const values = sheet.getDataRange().getValues();
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const idCol = headers.indexOf('id');
  if (idCol === -1) throw new AppError('SHEETS_ERROR', 'La hoja "' + nombreHoja + '" no tiene columna "id".');

  for (let i = 1; i < values.length; i++) {
    if (String(values[i][idCol]) === String(id)) {
      Object.keys(cambios).forEach(function (key) {
        const col = headers.indexOf(key);
        if (col !== -1) sheet.getRange(i + 1, col + 1).setValue(cambios[key]);
      });
      return true;
    }
  }
  throw new AppError('SHEETS_ERROR', 'No se encontró id="' + id + '" en la hoja "' + nombreHoja + '" para actualizar.');
}

function generarId_(prefijo) {
  return prefijo + '_' + Utilities.getUuid().substring(0, 8);
}

// ------------------------------ Clientes -------------------------------

function guardarCliente_(datos) {
  const id = generarId_('cli');
  const registro = {
    id: id,
    nombre: datos.nombre,
    pais: datos.pais, // el formulario ahora pide "Ciudad" — se guarda en esta misma columna para no romper hojas ya creadas
    url: datos.url || '',

    industria: datos.industria,
    categoria: datos.categoria,
    colaboradores: datos.colaboradores,
    tamano: datos.tamano,
    modulos_contratados: JSON.stringify(datos.modulosContratados || []),
    contexto_comercial: datos.contextoComercial || '',
    objetivos: datos.objetivos || '',
    dolores_iniciales: datos.doloresIniciales || '',
    info_adicional: datos.infoAdicional || '',
    procesos_actuales: datos.procesosActuales || '',
    herramientas_utilizadas: datos.herramientasUtilizadas || '',
    nivel_automatizacion: datos.nivelAutomatizacion || '',
    fecha_creacion: new Date(),
    usuario: getUsuarioActual_()
  };
  appendRow_(SHEETS.CLIENTES, registro);
  return registro;
}

function obtenerCliente_(clienteId) {
  const row = getRowById_(SHEETS.CLIENTES, clienteId);
  if (!row) throw new AppError('VALIDACION', 'Cliente no encontrado: ' + clienteId);
  row.modulos_contratados = safeParseJSON_(row.modulos_contratados, []);
  return row;
}

// ---------------------------- Diagnósticos ------------------------------

function guardarDiagnostico_(clienteId, modulo, respuestas, diagnosticoJson) {
  const id = generarId_('diag');
  const paraGuardar = {
    id: id,
    cliente_id: clienteId,
    modulo: modulo,
    respuestas_json: JSON.stringify(respuestas),
    diagnostico_json: JSON.stringify(diagnosticoJson),
    fecha: new Date()
  };
  appendRow_(SHEETS.DIAGNOSTICOS, paraGuardar);
  // Devolvemos las versiones "objeto" (no las stringificadas que se guardaron
  // en la hoja), porque quien llama a esta función las necesita en memoria
  // de inmediato (Code.gs -> respuesta al cliente; Consultoria.gs -> uso interno).
  return {
    id: id, cliente_id: clienteId, modulo: modulo,
    respuestas_json: respuestas, diagnostico_json: diagnosticoJson, fecha: paraGuardar.fecha
  };
}

function obtenerDiagnostico_(diagnosticoId) {
  const row = getRowById_(SHEETS.DIAGNOSTICOS, diagnosticoId);
  if (!row) throw new AppError('VALIDACION', 'Diagnóstico no encontrado: ' + diagnosticoId);
  row.respuestas_json = safeParseJSON_(row.respuestas_json, {});
  row.diagnostico_json = safeParseJSON_(row.diagnostico_json, {});
  return row;
}

function obtenerDiagnosticosPorCliente_(clienteId) {
  // Igual que obtenerDiagnostico_: hay que parsear respuestas_json/diagnostico_json,
  // si no quien llama recibe un string en vez de un objeto.
  return getAllRows_(SHEETS.DIAGNOSTICOS)
    .filter(function (r) { return String(r.cliente_id) === String(clienteId); })
    .map(function (r) {
      return {
        id: r.id, cliente_id: r.cliente_id, modulo: r.modulo,
        respuestas_json: safeParseJSON_(r.respuestas_json, {}),
        diagnostico_json: safeParseJSON_(r.diagnostico_json, {}),
        fecha: r.fecha
      };
    });
}

// ---------------------------- Consultorías ------------------------------

function guardarConsultoria_(clienteId, modulo, diagnosticoId, contenidoJson, version) {
  const id = generarId_('cons');
  const paraGuardar = {
    id: id,
    cliente_id: clienteId,
    modulo: modulo,
    diagnostico_id: diagnosticoId,
    contenido_json: JSON.stringify(contenidoJson),
    version: version || 1,
    estado: ESTADOS_CONSULTORIA.GENERADA,
    fecha_generacion: new Date(),
    fecha_aprobacion: '',
    usuario_aprobador: ''
  };
  appendRow_(SHEETS.CONSULTORIAS, paraGuardar);
  // Igual que en guardarDiagnostico_: devolvemos contenido_json como objeto
  // (no stringificado) porque el llamador lo usa de inmediato en memoria.
  return {
    id: id, cliente_id: clienteId, modulo: modulo, diagnostico_id: diagnosticoId,
    contenido_json: contenidoJson, version: paraGuardar.version, estado: paraGuardar.estado,
    fecha_generacion: paraGuardar.fecha_generacion, fecha_aprobacion: '', usuario_aprobador: ''
  };
}

function obtenerConsultoria_(consultoriaId) {
  const row = getRowById_(SHEETS.CONSULTORIAS, consultoriaId);
  if (!row) throw new AppError('VALIDACION', 'Consultoría no encontrada: ' + consultoriaId);
  row.contenido_json = safeParseJSON_(row.contenido_json, {});
  return row;
}

function obtenerConsultoriasPorCliente_(clienteId) {
  return getAllRows_(SHEETS.CONSULTORIAS)
    .filter(function (r) { return String(r.cliente_id) === String(clienteId); })
    .map(function (r) { r.contenido_json = safeParseJSON_(r.contenido_json, {}); return r; });
}

function actualizarConsultoria_(consultoriaId, contenidoJson, nuevoEstado) {
  const cambios = { contenido_json: JSON.stringify(contenidoJson) };
  if (nuevoEstado) cambios.estado = nuevoEstado;
  if (nuevoEstado === ESTADOS_CONSULTORIA.APROBADA) {
    cambios.fecha_aprobacion = new Date();
    cambios.usuario_aprobador = getUsuarioActual_();
  }
  updateRowById_(SHEETS.CONSULTORIAS, consultoriaId, cambios);
  return obtenerConsultoria_(consultoriaId);
}

function cambiarEstadoConsultoria_(consultoriaId, nuevoEstado) {
  const cambios = { estado: nuevoEstado };
  if (nuevoEstado === ESTADOS_CONSULTORIA.APROBADA) {
    cambios.fecha_aprobacion = new Date();
    cambios.usuario_aprobador = getUsuarioActual_();
  }
  updateRowById_(SHEETS.CONSULTORIAS, consultoriaId, cambios);
}

// ----------------------------- Entregables -------------------------------

function guardarEntregable_(clienteId, consultoriaIds, tipo, url) {
  const id = generarId_('entr');
  const registro = {
    id: id,
    cliente_id: clienteId,
    consultoria_ids: JSON.stringify(consultoriaIds),
    tipo: tipo,
    url: url,
    fecha: new Date(),
    usuario: getUsuarioActual_()
  };
  appendRow_(SHEETS.ENTREGABLES, registro);
  return registro;
}

// ------------------------------ Historial --------------------------------

/** Vista consolidada para la pantalla de historial (sección 15 del pedido). */
function obtenerHistorial_() {
  const clientes = {};
  getAllRows_(SHEETS.CLIENTES).forEach(function (c) { clientes[c.id] = c; });

  return getAllRows_(SHEETS.CONSULTORIAS).map(function (c) {
    const cliente = clientes[c.cliente_id] || {};
    return {
      id: c.id,
      fecha: c.fecha_generacion,
      usuario: cliente.usuario || '',
      cliente: cliente.nombre || '(cliente eliminado)',
      clienteId: c.cliente_id,
      modulo: c.modulo,
      version: c.version,
      estado: c.estado
    };
  }).sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
}

/**
 * Alimenta el selector de clientes de la pestaña "Consultoría y Gestión de
 * Entregables": un cliente cargado en la pestaña "Diagnóstico Comercial"
 * aparece acá con un resumen de cuántos módulos ya diagnosticó y en qué
 * estado va su consultoría, sin tener que abrir cada uno para saberlo.
 */
function obtenerClientesConResumen_() {
  const clientes = getAllRows_(SHEETS.CLIENTES);
  const diagnosticos = getAllRows_(SHEETS.DIAGNOSTICOS);
  const consultorias = getAllRows_(SHEETS.CONSULTORIAS);

  return clientes.map(function (c) {
    const diagsCliente = diagnosticos.filter(function (d) { return String(d.cliente_id) === String(c.id); });
    const consultoriasCliente = consultorias.filter(function (k) { return String(k.cliente_id) === String(c.id); });
    const conteoEstados = {};
    consultoriasCliente.forEach(function (k) { conteoEstados[k.estado] = (conteoEstados[k.estado] || 0) + 1; });

    return {
      id: c.id,
      nombre: c.nombre,
      pais: c.pais,
      url: c.url,
      industria: c.industria,
      tamano: c.tamano,
      colaboradores: c.colaboradores,
      modulos_contratados: safeParseJSON_(c.modulos_contratados, []),
      fecha_creacion: c.fecha_creacion,
      modulos_diagnosticados: diagsCliente.map(function (d) { return d.modulo; }),
      cantidad_consultorias: consultoriasCliente.length,
      conteo_estados: conteoEstados
    };
  }).sort(function (a, b) { return new Date(b.fecha_creacion) - new Date(a.fecha_creacion); });
}

function safeParseJSON_(texto, valorPorDefecto) {
  if (texto === '' || texto === null || texto === undefined) return valorPorDefecto;
  try {
    return typeof texto === 'string' ? JSON.parse(texto) : texto;
  } catch (e) {
    return valorPorDefecto;
  }
}
