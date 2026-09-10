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
    nombre: datos.nombre, // "Nombre de la empresa"
    pais: datos.pais, // el formulario ahora pide "Ciudad" — se guarda en esta misma columna para no romper hojas ya creadas
    url: datos.url || '',
    // Reemplaza al número de versión en el Historial ("Kick Off" / "Traspaso" / "Upselling") — se
    // define una vez, al registrar al cliente, y clasifica el TIPO de proyecto que es (sección de
    // ajustes de Historial a pedido de la usuaria).
    tipo_proyecto: datos.tipoProyecto || 'Kick Off',

    industria: datos.industria,
    categoria: datos.categoria,
    tamano: datos.tamano,
    // 'colaboradores' se deja de recoger (columna se conserva vacía en registros nuevos, solo para no
    // romper hojas con historial previo): el tamaño de la organización ya cubre ese dato, sin duplicidad.
    modulos_contratados: JSON.stringify(datos.modulosContratados || []),
    contexto_comercial: datos.contextoComercial || '',
    // Objetivos y dolores ahora se preguntan como una sola pregunta comercial ("¿Cuáles son los
    // objetivos principales del cliente y sus principales dolores?"); se guarda en 'objetivos' y se
    // deja 'dolores_iniciales' vacío en registros nuevos (se conserva la columna por compatibilidad).
    objetivos: datos.objetivosYDolores || datos.objetivos || '',
    dolores_iniciales: datos.doloresIniciales || '',
    info_adicional: datos.infoAdicional || '',
    procesos_actuales: datos.procesosActuales || '',
    herramientas_utilizadas: datos.herramientasUtilizadas || '',
    nivel_automatizacion: datos.nivelAutomatizacion || '',
    usa_herramientas_ofimaticas: datos.usaHerramientasOfimaticas || '',

    account_manager: datos.accountManager || '',
    ae_acompana_kickoff: datos.aeAcompanaKickoff || '',
    dominio_cliente: datos.dominioCliente || '',
    lider_nombre: datos.liderNombre || '',
    lider_celular: datos.liderCelular || '',
    lider_correo: datos.liderCorreo || '',
    lider_estrategico_nombre: datos.liderEstrategicoNombre || '',
    lider_estrategico_celular: datos.liderEstrategicoCelular || '',
    lider_estrategico_correo: datos.liderEstrategicoCorreo || '',
    sedes_ubicacion: datos.sedesUbicacion || '',
    usaba_otra_plataforma: datos.usabaOtraPlataforma || '',
    otra_plataforma_cual: datos.otraPlataformaCual || '',
    motivo_compra_buk: datos.motivoCompraBuk || '',
    cultura_frase: datos.culturaFrase || '',
    sesiones_grupales: datos.sesionesGrupales || '',
    estructura_equipo: datos.estructuraEquipo || '',

    integracion_sso: datos.integracionSso || '',
    integracion_api: datos.integracionApi || '',
    dominio_correo: datos.dominioCorreo || '',
    modulo_inicio_deseado: datos.moduloInicioDeseado || '',
    fecha_fin_esperada: datos.fechaFinEsperada || '',
    comentarios_tiempos: datos.comentariosTiempos || '',
    observaciones_modulos_json: JSON.stringify(datos.observacionesModulos || {}),

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
  row.observaciones_modulos_json = safeParseJSON_(row.observaciones_modulos_json, {});
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
  row.seguimiento_json = safeParseJSON_(row.seguimiento_json, {}); // undefined si la columna aún no existe en la hoja — seguro
  return row;
}

function obtenerConsultoriasPorCliente_(clienteId) {
  return getAllRows_(SHEETS.CONSULTORIAS)
    .filter(function (r) { return String(r.cliente_id) === String(clienteId); })
    .map(function (r) {
      r.contenido_json = safeParseJSON_(r.contenido_json, {});
      r.seguimiento_json = safeParseJSON_(r.seguimiento_json, {});
      return r;
    });
}

/**
 * Checklist "aplica / no aplica / requiere ajuste" que el COE completa junto
 * al cliente durante la consultoría guiada (sección 5 del pedido de rediseño).
 * Se guarda en una columna ADITIVA (seguimiento_json) que no afecta el resto
 * del contenido_json ya validado — ver agregarFaseKickOffYMagia() en
 * SetupSheets.gs para agregar esa columna a una hoja CONSULTORIAS existente.
 */
function guardarSeguimientoConsultoria_(consultoriaId, seguimiento) {
  updateRowById_(SHEETS.CONSULTORIAS, consultoriaId, { seguimiento_json: JSON.stringify(seguimiento || {}) });
  return obtenerConsultoria_(consultoriaId);
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
 * Vista consolidada de Historial: una fila por CLIENTE (no por consultoría/módulo
 * como antes), con acceso directo de descarga a los 3 documentos que ya genera la
 * app para ese cliente — Diagnóstico Comercial, Consultoría (último entregable) y
 * Radiografía — todo centralizado, tal como pidió la usuaria.
 */
function obtenerHistorialClientes_() {
  const diagnosticos = getAllRows_(SHEETS.DIAGNOSTICOS);
  const entregables = getAllRows_(SHEETS.ENTREGABLES);

  return getAllRows_(SHEETS.CLIENTES).map(function (c) {
    const tieneDiagnostico = diagnosticos.some(function (d) { return String(d.cliente_id) === String(c.id); });

    const entregablesCliente = entregables.filter(function (e) { return String(e.cliente_id) === String(c.id); });
    entregablesCliente.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
    const ultimoEntregable = entregablesCliente[0] || null;

    const radiografia = obtenerRadiografiaPorCliente_(c.id);

    return {
      id: c.id,
      nombre: c.nombre,
      tipoProyecto: c.tipo_proyecto || '',
      fechaRegistro: c.fecha_creacion,
      tieneDiagnostico: tieneDiagnostico,
      entregableUrl: ultimoEntregable ? ultimoEntregable.url : '',
      entregableFecha: ultimoEntregable ? ultimoEntregable.fecha : '',
      tieneRadiografia: !!radiografia
    };
  }).sort(function (a, b) { return new Date(b.fechaRegistro) - new Date(a.fechaRegistro); });
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
      modulos_contratados: safeParseJSON_(c.modulos_contratados, []),
      fecha_creacion: c.fecha_creacion,
      modulos_diagnosticados: diagsCliente.map(function (d) { return d.modulo; }),
      cantidad_consultorias: consultoriasCliente.length,
      conteo_estados: conteoEstados
    };
  }).sort(function (a, b) { return new Date(b.fecha_creacion) - new Date(a.fecha_creacion); });
}

// ------------------------------ Kick Off COE ------------------------------
// Etapa intermedia Diagnóstico -> Kick Off COE -> Consultoría (sección 3 del
// pedido de rediseño). Un cliente puede tener varios registros de KICKOFF
// (por ejemplo, si se repite la sesión); obtenerKickOffPorCliente_ siempre
// devuelve el más reciente, igual que el patrón usado para diagnósticos.

function guardarKickOff_(clienteId, datos, notasGemini) {
  const id = generarId_('ko');
  const registro = {
    id: id,
    cliente_id: clienteId,
    notas_gemini: notasGemini || '',
    contexto: (datos && datos.contexto) || '',
    necesidades_json: JSON.stringify((datos && datos.necesidades) || []),
    procesos_json: JSON.stringify((datos && datos.procesos) || []),
    expectativas_json: JSON.stringify((datos && datos.expectativas) || []),
    fecha: new Date(),
    usuario: getUsuarioActual_()
  };
  appendRow_(SHEETS.KICKOFF, registro);
  return {
    id: id, cliente_id: clienteId, notas_gemini: registro.notas_gemini, contexto: registro.contexto,
    necesidades: (datos && datos.necesidades) || [], procesos: (datos && datos.procesos) || [],
    expectativas: (datos && datos.expectativas) || [], fecha: registro.fecha
  };
}

function obtenerKickOffPorCliente_(clienteId) {
  const rows = getAllRows_(SHEETS.KICKOFF).filter(function (r) { return String(r.cliente_id) === String(clienteId); });
  if (!rows.length) return null;
  rows.sort(function (a, b) { return new Date(b.fecha) - new Date(a.fecha); });
  const r = rows[0];
  return {
    id: r.id, cliente_id: r.cliente_id, notas_gemini: r.notas_gemini || '', contexto: r.contexto || '',
    necesidades: safeParseJSON_(r.necesidades_json, []),
    procesos: safeParseJSON_(r.procesos_json, []),
    expectativas: safeParseJSON_(r.expectativas_json, []),
    fecha: r.fecha
  };
}

// ------------------------- Radiografía del Cliente -------------------------
// Flujo Información Comercial -> Preguntas Orientadoras -> Kick Off -> Notas
// Gemini -> Radiografía. Se guarda como un registro más de KICKOFF (misma
// hoja, columnas aditivas) para no duplicar el concepto de "sesión de Kick
// Off" — obtenerRadiografiaPorCliente_ siempre devuelve la más reciente,
// igual que el resto de entidades de esta app.

function guardarRadiografia_(clienteId, notasGemini, notasGeminiUrl, fuenteNotas, radiografiaJson) {
  const id = generarId_('rad');
  const registro = {
    id: id,
    cliente_id: clienteId,
    notas_gemini: notasGemini || '',
    notas_gemini_url: notasGeminiUrl || '',
    fuente_notas: fuenteNotas || 'pegado', // 'pegado' | 'url'
    radiografia_json: JSON.stringify(radiografiaJson || {}),
    fecha_radiografia: new Date(),
    fecha: new Date(),
    usuario: getUsuarioActual_()
  };
  appendRow_(SHEETS.KICKOFF, registro);
  return {
    id: id, cliente_id: clienteId, notas_gemini: registro.notas_gemini, notas_gemini_url: registro.notas_gemini_url,
    fuente_notas: registro.fuente_notas, radiografia: radiografiaJson || {}, fecha_radiografia: registro.fecha_radiografia
  };
}

function obtenerRadiografiaPorCliente_(clienteId) {
  const rows = getAllRows_(SHEETS.KICKOFF)
    .filter(function (r) { return String(r.cliente_id) === String(clienteId) && r.radiografia_json; });
  if (!rows.length) return null;
  rows.sort(function (a, b) { return new Date(b.fecha_radiografia || b.fecha) - new Date(a.fecha_radiografia || a.fecha); });
  const r = rows[0];
  return {
    id: r.id, cliente_id: r.cliente_id, notas_gemini: r.notas_gemini || '', notas_gemini_url: r.notas_gemini_url || '',
    fuente_notas: r.fuente_notas || 'pegado', radiografia: safeParseJSON_(r.radiografia_json, {}),
    fecha_radiografia: r.fecha_radiografia || r.fecha
  };
}

/**
 * Búsqueda de candidatos para la auto-detección de cliente a partir del texto
 * (o el título) de un documento de notas de Kick Off (sección 4 del pedido).
 * NO decide por sí sola: devuelve todos los clientes cuyo nombre o dominio
 * aparece mencionado en el texto, para que la UI decida si hay un único
 * candidato claro (auto-confirmar) o si debe pedir selección manual.
 */
function buscarClientesPorTexto_(texto) {
  if (!texto) return [];
  const textoNormalizado = String(texto).toLowerCase();
  return getAllRows_(SHEETS.CLIENTES)
    .filter(function (c) {
      const nombre = String(c.nombre || '').toLowerCase();
      const dominio = String(c.dominio_cliente || '').toLowerCase();
      return (nombre.length > 2 && textoNormalizado.indexOf(nombre) !== -1) ||
        (dominio.length > 2 && textoNormalizado.indexOf(dominio) !== -1);
    })
    .map(function (c) { return { id: c.id, nombre: c.nombre, dominio_cliente: c.dominio_cliente || '' }; });
}

// -------------------------- Biblioteca de Magia (HERRAMIENTAS) --------------------------
// Catálogo de AppScripts internos del equipo COE. `url` vacío = herramienta
// "pendiente de link" (arquitectura lista para completarla sin tocar código
// — sección 13.8 del pedido): basta con escribir el link en la hoja HERRAMIENTAS.

function obtenerHerramientas_() {
  return getAllRows_(SHEETS.HERRAMIENTAS, getKnowledgeSpreadsheet()).map(function (r) {
    return {
      id: r.id, nombre: r.nombre, url: r.url || '', modulo: r.modulo || 'General',
      tipo: r.tipo || '', etapa: r.etapa || '', uso_recomendado: r.uso_recomendado || '',
      estado: r.url ? 'Disponible' : 'Pendiente de link'
    };
  });
}

function obtenerHerramientasPorModulos_(modulos) {
  if (!modulos || !modulos.length) return [];
  return obtenerHerramientas_().filter(function (h) { return modulos.indexOf(h.modulo) !== -1; });
}

function safeParseJSON_(texto, valorPorDefecto) {
  if (texto === '' || texto === null || texto === undefined) return valorPorDefecto;
  try {
    return typeof texto === 'string' ? JSON.parse(texto) : texto;
  } catch (e) {
    return valorPorDefecto;
  }
}
