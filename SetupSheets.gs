/**
 * SetupSheets.gs
 * Script de inicialización, para correr UNA VEZ desde el editor de Apps
 * Script (seleccionar la función `inicializarBaseDeDatos` y presionar
 * Ejecutar). Crea todas las hojas necesarias con sus encabezados y carga
 * datos semilla:
 *   - Módulo Onboarding completo, replicando el ejemplo condicional del
 *     encargo (Manual / Parcialmente automatizado / Automatizado / No existe).
 *   - Un árbol básico (root + una rama condicional) para los otros 5
 *     módulos iniciales, listo para que el equipo COI lo enriquezca.
 *   - Conocimiento, plantillas, propuestas de valor y oportunidades de
 *     ejemplo para poder probar el flujo de punta a punta.
 *
 * Es seguro volver a ejecutarlo: si una hoja ya existe, se limpia y se
 * vuelve a poblar (útil en desarrollo). En producción, una vez cargada la
 * metodología real del equipo, NO lo vuelvas a correr o perderás los cambios.
 */

function inicializarBaseDeDatos() {
  const ss = getSpreadsheet();

  crearHojaConEncabezados_(ss, SHEETS.MODULOS, ['id', 'modulo', 'descripcion', 'activo', 'orden']);
  crearHojaConEncabezados_(ss, SHEETS.PREGUNTAS, ['id', 'modulo', 'pregunta', 'tipo', 'orden', 'obligatoria', 'pregunta_padre_id', 'valor_padre_dispara', 'ayuda']);
  crearHojaConEncabezados_(ss, SHEETS.OPCIONES, ['id', 'pregunta_id', 'valor', 'etiqueta_necesidad', 'etiqueta_dolor', 'peso_madurez']);
  crearHojaConEncabezados_(ss, SHEETS.CONOCIMIENTO, ['id', 'modulo', 'industria', 'tamano', 'madurez', 'categoria', 'tipo', 'contenido', 'activo']);
  crearHojaConEncabezados_(ss, SHEETS.PLANTILLAS, ['id', 'modulo', 'industria', 'tipo', 'nombre', 'contenido', 'activo']);
  crearHojaConEncabezados_(ss, SHEETS.PROPUESTAS_VALOR, ['id', 'modulo', 'necesidad', 'propuesta', 'activo']);
  crearHojaConEncabezados_(ss, SHEETS.OPORTUNIDADES, ['id', 'necesidad', 'oportunidad_base']);
  crearHojaConEncabezados_(ss, SHEETS.HERRAMIENTAS, ['id', 'nombre', 'url', 'modulo', 'tipo', 'etapa', 'uso_recomendado']);

  crearHojaConEncabezados_(ss, SHEETS.CLIENTES, ['id', 'nombre', 'pais', 'url', 'industria', 'categoria', 'colaboradores', 'tamano',
    'modulos_contratados', 'contexto_comercial', 'objetivos', 'dolores_iniciales', 'info_adicional',
    'procesos_actuales', 'herramientas_utilizadas', 'nivel_automatizacion', 'usa_herramientas_ofimaticas',
    'account_manager', 'ae_acompana_kickoff', 'dominio_cliente',
    'lider_nombre', 'lider_celular', 'lider_correo',
    'lider_estrategico_nombre', 'lider_estrategico_celular', 'lider_estrategico_correo',
    'sedes_ubicacion', 'usaba_otra_plataforma', 'otra_plataforma_cual',
    'motivo_compra_buk', 'cultura_frase', 'sesiones_grupales', 'estructura_equipo',
    'integracion_sso', 'integracion_api', 'dominio_correo',
    'modulo_inicio_deseado', 'fecha_fin_esperada', 'comentarios_tiempos', 'observaciones_modulos_json',
    'fecha_creacion', 'usuario']);
  crearHojaConEncabezados_(ss, SHEETS.DIAGNOSTICOS, ['id', 'cliente_id', 'modulo', 'respuestas_json', 'diagnostico_json', 'fecha']);
  crearHojaConEncabezados_(ss, SHEETS.KICKOFF, ['id', 'cliente_id', 'notas_gemini', 'contexto', 'necesidades_json',
    'procesos_json', 'expectativas_json', 'fecha', 'usuario']);
  crearHojaConEncabezados_(ss, SHEETS.CONSULTORIAS, ['id', 'cliente_id', 'modulo', 'diagnostico_id', 'contenido_json', 'version',
    'estado', 'fecha_generacion', 'fecha_aprobacion', 'usuario_aprobador', 'seguimiento_json']);
  crearHojaConEncabezados_(ss, SHEETS.ENTREGABLES, ['id', 'cliente_id', 'consultoria_ids', 'tipo', 'url', 'fecha', 'usuario']);

  poblarModulos_();
  poblarOnboardingCompleto_();
  ['Gestión del Desempeño', 'Selección', 'Encuestas / Clima', 'Reconocimiento', 'Comunicaciones',
    'Beneficios', 'Canal de Denuncias', 'Servicio al Colaborador', 'API'].forEach(poblarModuloBasico_);
  poblarPropuestasValorYOportunidades_();
  poblarHerramientas_();

  Logger.log('Base de datos inicializada correctamente. Revisa las pestañas del spreadsheet.');
}

function crearHojaConEncabezados_(ss, nombre, headers) {
  let sheet = ss.getSheetByName(nombre);
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(nombre); }
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
  return sheet;
}

function poblarModulos_() {
  const modulos = [
    ['mod_onboarding', 'Onboarding', 'Proceso de bienvenida e integración de nuevos colaboradores.', true, 1],
    ['mod_desempeno', 'Gestión del Desempeño', 'Evaluaciones de desempeño, objetivos y seguimiento continuo.', true, 2],
    ['mod_seleccion', 'Selección', 'Procesos de atracción y selección de talento.', true, 3],
    ['mod_encuestas', 'Encuestas / Clima', 'Medición de clima organizacional y experiencia del colaborador.', true, 4],
    ['mod_reconocimiento', 'Reconocimiento', 'Programas de reconocimiento entre colaboradores.', true, 5],
    ['mod_comunicaciones', 'Comunicaciones', 'Comunicación interna y difusión de información a colaboradores.', true, 6],
    ['mod_beneficios', 'Beneficios', 'Gestión, comunicación y administración de beneficios para colaboradores.', true, 7],
    ['mod_denuncias', 'Canal de Denuncias', 'Canal de reporte y gestión de denuncias e investigaciones internas.', true, 8],
    ['mod_servicio_colaborador', 'Servicio al Colaborador', 'Atención de solicitudes y casos del colaborador (mesa de ayuda de RRHH).', true, 9],
    ['mod_workflow', 'Workflow', 'Automatización de flujos de aprobación y procesos entre áreas.', true, 10]
  ];
  modulos.forEach(function (m) {
    appendRow_(SHEETS.MODULOS, { id: m[0], modulo: m[1], descripcion: m[2], activo: m[3], orden: m[4] });
  });
}

/**
 * Para una base de datos que YA fue inicializada con inicializarBaseDeDatos() y que ya
 * tiene clientes/diagnósticos/consultorías cargados: agrega los 4 módulos nuevos
 * (Beneficios, Canal de Denuncias, Servicio al Colaborador, API) con su árbol básico de
 * preguntas, SIN borrar nada de lo existente (a diferencia de inicializarBaseDeDatos,
 * que limpia las hojas antes de repoblarlas).
 * Para correrla: selecciona "agregarModulosCOE_" en el desplegable de funciones de arriba
 * y presiona "Ejecutar". Ejecútala una sola vez — si la corres dos veces, los módulos
 * quedarán duplicados en la hoja MODULOS.
 */
function agregarModulosCOE_() {
  const nuevos = [
    ['mod_beneficios', 'Beneficios', 'Gestión, comunicación y administración de beneficios para colaboradores.', true, 7],
    ['mod_denuncias', 'Canal de Denuncias', 'Canal de reporte y gestión de denuncias e investigaciones internas.', true, 8],
    ['mod_servicio_colaborador', 'Servicio al Colaborador', 'Atención de solicitudes y casos del colaborador (mesa de ayuda de RRHH).', true, 9],
    ['mod_api', 'API', 'Integraciones y automatizaciones vía API con otros sistemas del cliente.', true, 10]
  ];
  nuevos.forEach(function (m) {
    appendRow_(SHEETS.MODULOS, { id: m[0], modulo: m[1], descripcion: m[2], activo: m[3], orden: m[4] });
  });
  nuevos.forEach(function (m) { poblarModuloBasico_(m[1]); });

  Logger.log('Módulos nuevos agregados: Beneficios, Canal de Denuncias, Servicio al Colaborador, API.');
}

/**
 * Onboarding: replica exactamente el árbol condicional descrito en el
 * encargo (sección 4) como referencia de cómo modelar preguntas nuevas.
 */
function poblarOnboardingCompleto_() {
  const M = 'Onboarding';

  pregunta_('p_ob_1', M, '¿Cómo gestiona actualmente el proceso de onboarding?', 'radio', 1, true, null, null,
    'Selecciona la opción que mejor describe el estado actual.');
  opciones_('p_ob_1', [
    ['Manual', 'Automatización', 'Alta operatividad manual', 1],
    ['Parcialmente automatizado', 'Estandarización', 'Procesos inconsistentes', 2],
    ['Automatizado', 'Optimización continua', '', 4],
    ['No existe un proceso definido', 'Definición de proceso', 'Falta de trazabilidad', 0]
  ]);

  // --- Rama: Manual ---
  pregunta_('p_ob_2', M, '¿Cuenta con plantillas?', 'radio', 2, true, 'p_ob_1', 'Manual', '');
  opciones_('p_ob_2', [['Sí', '', '', 2], ['No', 'Estandarización', 'Falta de plantillas', 0]]);

  pregunta_('p_ob_3', M, '¿Cuenta con correos estandarizados?', 'radio', 3, true, 'p_ob_1', 'Manual', '');
  opciones_('p_ob_3', [['Sí', '', '', 2], ['No', 'Mejora de comunicación', 'Comunicación inconsistente', 0]]);

  pregunta_('p_ob_4', M, '¿Tiene tareas definidas?', 'radio', 4, true, 'p_ob_1', 'Manual', '');
  opciones_('p_ob_4', [['Sí', '', '', 2], ['No', 'Estandarización', 'Falta de trazabilidad', 0]]);

  pregunta_('p_ob_5', M, '¿Existe un responsable del proceso?', 'radio', 5, true, 'p_ob_1', 'Manual', '');
  opciones_('p_ob_5', [['Sí', '', '', 2], ['No', 'Gobernanza del proceso', 'Falta de responsable claro', 0]]);

  pregunta_('p_ob_6', M, '¿Existe seguimiento?', 'radio', 6, true, 'p_ob_1', 'Manual', '');
  opciones_('p_ob_6', [['Sí', '', '', 2], ['No', 'Trazabilidad', 'Falta de trazabilidad', 0]]);

  pregunta_('p_ob_7', M, '¿Cómo se comunica actualmente el proceso?', 'texto', 7, false, 'p_ob_1', 'Manual',
    'Describe el canal (correo, WhatsApp, verbal, etc.)');

  pregunta_('p_ob_8', M, '¿Qué dificultades presenta el proceso actual?', 'texto', 8, false, 'p_ob_1', 'Manual', '');

  // --- Rama: Parcialmente automatizado ---
  pregunta_('p_ob_9', M, '¿Qué partes del proceso siguen siendo manuales?', 'texto', 9, true, 'p_ob_1', 'Parcialmente automatizado', '');
  pregunta_('p_ob_10', M, '¿Qué automatización o integración planean agregar próximamente?', 'texto', 10, false, 'p_ob_1', 'Parcialmente automatizado', '');

  // --- Rama: Automatizado (orientadas a optimización, no a implementación básica) ---
  pregunta_('p_ob_11', M, '¿Qué tan satisfecho está el equipo con la herramienta/proceso actual?', 'radio', 11, true, 'p_ob_1', 'Automatizado', '');
  opciones_('p_ob_11', [
    ['Muy satisfecho', '', '', 4],
    ['Parcialmente satisfecho', 'Optimización de configuración', 'Fricciones puntuales', 3],
    ['Poco satisfecho', 'Rediseño del flujo', 'Herramienta subutilizada', 2]
  ]);
  pregunta_('p_ob_12', M, '¿Ha identificado cuellos de botella en el proceso automatizado?', 'texto', 12, false, 'p_ob_1', 'Automatizado', '');
  pregunta_('p_ob_13', M, '¿Qué métricas de onboarding monitorean actualmente?', 'texto', 13, false, 'p_ob_1', 'Automatizado', '');

  // --- Rama: No existe un proceso definido ---
  pregunta_('p_ob_14', M, '¿Quién realiza hoy el onboarding de nuevos colaboradores?', 'texto', 14, true, 'p_ob_1', 'No existe un proceso definido', '');
  pregunta_('p_ob_15', M, '¿Cuánto tiempo toma aproximadamente que un nuevo colaborador sea productivo?', 'texto', 15, false, 'p_ob_1', 'No existe un proceso definido', '');

  // Conocimiento de ejemplo (General + específico por industria/tamaño/madurez)
  conocimiento_(M, 'General', 'General', 'General', 'metodologia', 'metodologia',
    'La metodología COI de onboarding se estructura en 3 fases: Pre-ingreso (documentación y expectativas), ' +
    'Primer día/semana (bienvenida, accesos, plan de aprendizaje) y Seguimiento (check-ins a 30/60/90 días).');
  conocimiento_(M, 'General', 'General', 'Baja', 'buena_practica', 'buena_practica',
    'Cuando la madurez es baja, priorizar primero estandarizar comunicaciones (correos y checklist) antes de ' +
    'introducir automatización: reduce riesgo de adopción y da resultados rápidos y visibles.');
  conocimiento_(M, 'Retail', 'General', 'General', 'recomendacion_industria', 'recomendacion_industria',
    'En Retail, considerar procesos de onboarding con alta rotación y múltiples puntos geográficos: priorizar ' +
    'plantillas replicables por tienda/sede y un flujo que no dependa de presencialidad centralizada.');
  conocimiento_(M, 'General', 'L (más de 1.000 colaboradores)', 'General', 'recomendacion_tamano', 'recomendacion_tamano',
    'En organizaciones grandes, la gobernanza del proceso (dueño claro, SLA de tareas por área) es tan ' +
    'importante como la automatización: sin gobernanza, la herramienta se subutiliza.');
  conocimiento_(M, 'General', 'General', 'General', 'error_frecuente', 'error_frecuente',
    'Error frecuente: automatizar el envío de correos sin antes definir el contenido y tono, generando una ' +
    'experiencia despersonalizada para el nuevo colaborador.');

  plantilla_(M, 'General', 'email', 'Correo de bienvenida', 'Asunto: ¡Bienvenido/a a {empresa}! — Cuerpo: información de primer día, accesos y contacto de referencia.');
  plantilla_(M, 'General', 'tarea', 'Checklist primer día', 'Entrega de equipo, accesos a sistemas, presentación al equipo, agenda primera semana.');
}

/**
 * Árbol básico (1 pregunta raíz + 1 rama condicional) para los módulos
 * que aún no tienen metodología detallada cargada. Sirve como plantilla
 * de partida para que el equipo COI la reemplace/expanda desde Sheets,
 * sin tocar código — ver GUIA_INSTALACION.md, sección "Agregar un módulo".
 */
function poblarModuloBasico_(modulo) {
  const idBase = 'p_' + slug_(modulo);

  pregunta_(idBase + '_1', modulo, '¿Cómo gestiona actualmente el proceso de ' + modulo.toLowerCase() + '?', 'radio', 1, true, null, null, '');
  opciones_(idBase + '_1', [
    ['Manual', 'Automatización', 'Alta operatividad manual', 1],
    ['Parcialmente automatizado', 'Estandarización', 'Procesos inconsistentes', 2],
    ['Automatizado', 'Optimización continua', '', 4],
    ['No existe un proceso definido', 'Definición de proceso', 'Falta de trazabilidad', 0]
  ]);

  pregunta_(idBase + '_2', modulo, '¿Qué dificultades presenta el proceso actual?', 'texto', 2, false, idBase + '_1', 'Manual', '');
  pregunta_(idBase + '_3', modulo, '¿Existe un responsable definido del proceso?', 'radio', 3, true, idBase + '_1', 'Manual', '');
  opciones_(idBase + '_3', [['Sí', '', '', 2], ['No', 'Gobernanza del proceso', 'Falta de responsable claro', 0]]);

  conocimiento_(modulo, 'General', 'General', 'General', 'metodologia', 'metodologia',
    'Metodología pendiente de carga detallada por el equipo COI para el módulo ' + modulo + '. ' +
    'Completar en la hoja CONOCIMIENTO siguiendo el mismo patrón usado en Onboarding.');
}

// ------------------------------- Helpers de carga -------------------------------

function pregunta_(id, modulo, texto, tipo, orden, obligatoria, padreId, valorPadre, ayuda) {
  appendRow_(SHEETS.PREGUNTAS, {
    id: id, modulo: modulo, pregunta: texto, tipo: tipo, orden: orden, obligatoria: obligatoria,
    pregunta_padre_id: padreId || '', valor_padre_dispara: valorPadre || '', ayuda: ayuda || ''
  });
}

function opciones_(preguntaId, filas) {
  filas.forEach(function (f, i) {
    appendRow_(SHEETS.OPCIONES, {
      id: preguntaId + '_o' + (i + 1), pregunta_id: preguntaId, valor: f[0],
      etiqueta_necesidad: f[1] || '', etiqueta_dolor: f[2] || '', peso_madurez: f[3]
    });
  });
}

function conocimiento_(modulo, industria, tamano, madurez, categoria, tipo, contenido) {
  appendRow_(SHEETS.CONOCIMIENTO, {
    id: generarId_('kb'), modulo: modulo, industria: industria, tamano: tamano, madurez: madurez,
    categoria: categoria, tipo: tipo, contenido: contenido, activo: true
  });
}

function plantilla_(modulo, industria, tipo, nombre, contenido) {
  appendRow_(SHEETS.PLANTILLAS, {
    id: generarId_('pl'), modulo: modulo, industria: industria, tipo: tipo, nombre: nombre, contenido: contenido, activo: true
  });
}

function poblarPropuestasValorYOportunidades_() {
  const oportunidades = [
    ['Automatización', 'Reducir la carga operativa manual del equipo de Personas mediante flujos automatizados.'],
    ['Estandarización', 'Estandarizar comunicaciones y tareas para asegurar una experiencia consistente.'],
    ['Mejora de comunicación', 'Ordenar y estandarizar los canales de comunicación del proceso.'],
    ['Definición de proceso', 'Definir un proceso base antes de automatizar, evitando digitalizar el caos.'],
    ['Gobernanza del proceso', 'Asignar responsables claros y SLAs para sostener el proceso en el tiempo.'],
    ['Trazabilidad', 'Incorporar seguimiento y trazabilidad de cada caso en curso.'],
    ['Optimización continua', 'Revisar y ajustar el proceso automatizado con base en métricas de uso.'],
    ['Optimización de configuración', 'Ajustar la configuración actual de la plataforma para resolver fricciones puntuales.'],
    ['Rediseño del flujo', 'Rediseñar el flujo automatizado actual para aumentar su adopción.']
  ];
  oportunidades.forEach(function (o) {
    appendRow_(SHEETS.OPORTUNIDADES, { id: generarId_('op'), necesidad: o[0], oportunidad_base: o[1] });
  });

  const propuestas = [
    ['Onboarding', 'Automatización', 'Automatizar tareas y comunicaciones del onboarding para reducir la carga operativa manual del equipo de Personas y liberar tiempo para acompañamiento real a los nuevos colaboradores.'],
    ['Onboarding', 'Estandarización', 'Estandarizar el proceso de onboarding (plantillas, correos y tareas) para asegurar una experiencia consistente a todos los nuevos colaboradores, independientemente de quién la ejecute.'],
    ['Onboarding', 'Trazabilidad', 'Incorporar seguimiento estructurado del proceso de onboarding para dar visibilidad al equipo de Personas y anticipar riesgos de fuga temprana.']
  ];
  propuestas.forEach(function (p) {
    appendRow_(SHEETS.PROPUESTAS_VALOR, { id: generarId_('pv'), modulo: p[0], necesidad: p[1], propuesta: p[2], activo: true });
  });
}

/**
 * Catálogo inicial de la Biblioteca de Magia (sección 7 y 8 del rediseño):
 * los 4 AppScripts confirmados + 4 filas "pendiente de link" para que el
 * equipo solo tenga que pegar la URL cuando la tenga — sin tocar código
 * (arquitectura lista, sección 13.8 del pedido).
 */
function poblarHerramientas_() {
  const herramientas = [
    ['herr_objetivos', 'Asistente Objetivos', 'https://script.google.com/a/macros/buk.co/s/AKfycbz0bQH7Jtjh7XB9yCVniF1q5_SxUs9rYg6u9TXFmIRGlnbOdN4Ro3LlK9Hst3pgM5ml/exec',
      'Gestión del Desempeño', 'Generador de documento', 'Consultoría', 'Estructura objetivos y metas sugeridas para clientes que no los tienen definidos.'],
    ['herr_competencias', 'Asistente Competencias Buk', 'https://script.google.com/a/macros/buk.co/s/AKfycbw-9gxfIemoUpor9Zix8ie0T9K7kaZwAum7yYpqrTRXzSRN7OC-8I1fpJnr9YGQbBTpdQ/exec',
      'Gestión del Desempeño', 'Generador de documento', 'Consultoría', 'Genera un marco de competencias, indicadores conductuales y niveles de desarrollo.'],
    ['herr_onboarding', 'Asistente Onboarding Buk', 'https://script.google.com/a/macros/buk.co/s/AKfycby7g_h2bDWxLjHI1OB7YjtcYzxN4XuIUaFgv8uk9LyKtF4RPb_cbcM8Qfhyjcx7_ErK/exec',
      'Onboarding', 'Generador de documento/plan', 'Consultoría', 'Sugiere el flujo de onboarding, campos del formulario de pre-ingreso, tareas y correos.'],
    ['herr_encuestas', 'Encuestas', 'https://script.google.com/a/macros/buk.co/s/AKfycbykgKb0apSsVH3DrcJO5-imnkr_SU2y80tPO4B8rsPHtr936zCEJYHJAgIdU7L3ni6T/exec',
      'Encuestas / Clima', '', '', ''],
    ['herr_servicio_colaborador', 'Servicio al Colaborador', '', 'Servicio al Colaborador', '', '', ''],
    ['herr_comunicaciones', 'Comunicaciones y Reconocimientos', '', 'Comunicaciones', '', '', ''],
    ['herr_beneficios', 'Beneficios', '', 'Beneficios', '', '', ''],
    ['herr_workflow', 'Workflow', '', 'API', '', '', '']
  ];
  herramientas.forEach(function (h) {
    appendRow_(SHEETS.HERRAMIENTAS, { id: h[0], nombre: h[1], url: h[2], modulo: h[3], tipo: h[4], etapa: h[5], uso_recomendado: h[6] });
  });
}

/**
 * Para una base de datos que YA fue inicializada y tiene datos de clientes
 * cargados: agrega la etapa Kick Off COE y la Biblioteca de Magia SIN BORRAR
 * NADA de lo existente (mismo patrón seguro que agregarModulosCOE_).
 * Es la única función nueva que hay que correr en un proyecto que ya está
 * en uso — reemplaza el intento anterior (agregarHerramientasCOE_, revertido).
 *
 * Para correrla: en el editor de Apps Script, abre este archivo
 * (SetupSheets.gs), en el desplegable de funciones (junto al botón
 * "Ejecutar", arriba) selecciona "agregarFaseKickOffYMagia" y presiona
 * Ejecutar. Es segura de correr más de una vez: si una hoja ya existe o la
 * columna ya fue agregada, no la vuelve a tocar.
 */
function agregarFaseKickOffYMagia() {
  const ss = getSpreadsheet();

  const yaExistiaKickoff = !!ss.getSheetByName(SHEETS.KICKOFF);
  if (!yaExistiaKickoff) {
    crearHojaConEncabezados_(ss, SHEETS.KICKOFF, ['id', 'cliente_id', 'notas_gemini', 'contexto', 'necesidades_json',
      'procesos_json', 'expectativas_json', 'fecha', 'usuario']);
  }

  const yaExistiaHerramientas = !!ss.getSheetByName(SHEETS.HERRAMIENTAS);
  if (!yaExistiaHerramientas) {
    crearHojaConEncabezados_(ss, SHEETS.HERRAMIENTAS, ['id', 'nombre', 'url', 'modulo', 'tipo', 'etapa', 'uso_recomendado']);
    poblarHerramientas_();
  }

  agregarColumnaSiFalta_(ss, SHEETS.CONSULTORIAS, 'seguimiento_json');

  Logger.log('Listo: Kick Off COE y Biblioteca de Magia agregados (o ya existían). Nada existente fue modificado ni borrado.');
}

/** Agrega `nombreColumna` al final de los encabezados de `nombreHoja` si todavía no existe. No toca filas ya cargadas. */
function agregarColumnaSiFalta_(ss, nombreHoja, nombreColumna) {
  const sheet = ss.getSheetByName(nombreHoja);
  if (!sheet) return; // si la hoja no existe todavía, no hay nada que agregar
  const ultimaCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, ultimaCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (headers.indexOf(nombreColumna) !== -1) return; // ya existe
  sheet.getRange(1, ultimaCol + 1).setValue(nombreColumna).setFontWeight('bold').setBackground('#1e3a8a').setFontColor('#ffffff');
}

function slug_(texto) {
  return texto.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}

// =====================================================================
// ACTUALIZACIÓN "DIAGNÓSTICO COMERCIAL" (ajuste de contenido y campos,
// pedido posterior al rediseño visual — NO toca Kick Off, Consultoría
// ni Biblioteca de Magia). Para una base de datos que YA está en uso:
// agrega columnas nuevas a CLIENTES sin borrar nada, agrega el módulo
// "Workflow", desactiva "API" como módulo contratable (pasa a ser un
// checkbox de integración, sección "Módulos especiales"), y REEMPLAZA
// el árbol de preguntas de 9 módulos por la versión superficial en
// tercera persona pedida para esta etapa (la profundización queda para
// el Kick Off). Es seguro volver a correrla: las columnas no se
// duplican y el árbol de cada módulo se limpia antes de recrearse.
//
// Para correrla: en el editor de Apps Script, selecciona
// "actualizarDiagnosticoComercial" en el desplegable de funciones y
// presiona Ejecutar.
// =====================================================================

function actualizarDiagnosticoComercial() {
  const ss = getSpreadsheet();

  [
    'usa_herramientas_ofimaticas',
    'account_manager', 'ae_acompana_kickoff', 'dominio_cliente',
    'lider_nombre', 'lider_celular', 'lider_correo',
    'lider_estrategico_nombre', 'lider_estrategico_celular', 'lider_estrategico_correo',
    'sedes_ubicacion', 'usaba_otra_plataforma', 'otra_plataforma_cual',
    'motivo_compra_buk', 'cultura_frase', 'sesiones_grupales', 'estructura_equipo',
    'integracion_sso', 'integracion_api', 'dominio_correo',
    'modulo_inicio_deseado', 'fecha_fin_esperada', 'comentarios_tiempos', 'observaciones_modulos_json'
  ].forEach(function (col) { agregarColumnaSiFalta_(ss, SHEETS.CLIENTES, col); });

  agregarModuloWorkflowYDesactivarApi_();

  reemplazarArbolPreguntas_('Gestión del Desempeño', poblarPreguntasDesempenoComercial_);
  reemplazarArbolPreguntas_('Selección', poblarPreguntasSeleccionComercial_);
  reemplazarArbolPreguntas_('Encuestas / Clima', poblarPreguntasEncuestasComercial_);
  reemplazarArbolPreguntas_('Onboarding', poblarPreguntasOnboardingComercial_);
  reemplazarArbolPreguntas_('Reconocimiento', poblarPreguntasReconocimientoComercial_);
  reemplazarArbolPreguntas_('Comunicaciones', poblarPreguntasComunicacionesComercial_);
  reemplazarArbolPreguntas_('Beneficios', poblarPreguntasBeneficiosComercial_);
  reemplazarArbolPreguntas_('Servicio al Colaborador', poblarPreguntasServicioColaboradorComercial_);
  reemplazarArbolPreguntas_('Canal de Denuncias', poblarPreguntasCanalDenunciasComercial_);

  agregarOportunidadesDiagnosticoComercial_();

  Logger.log('Diagnóstico Comercial actualizado: columnas nuevas en CLIENTES, módulo Workflow agregado ' +
    '(API pasa a checkbox de integración), y árbol de preguntas de 9 módulos reemplazado por la versión ' +
    'superficial en tercera persona. Nada de Kick Off, Consultoría o Biblioteca de Magia fue modificado.');
}

/** Agrega el módulo "Workflow" al checklist de módulos contratables (si no existe) y desactiva
 * "API" como módulo contratable: en el nuevo diseño, API es un checkbox de integración especial
 * dentro de la pestaña "Módulos", no un módulo que se contrata. No se borra la fila de MODULOS ni
 * las preguntas ya cargadas para "API", para no perder diagnósticos ya guardados con ese módulo. */
function agregarModuloWorkflowYDesactivarApi_() {
  const ss = getKnowledgeSpreadsheet();
  const modulos = getAllRows_(SHEETS.MODULOS, ss);

  const yaExisteWorkflow = modulos.some(function (m) { return m.modulo === 'Workflow'; });
  if (!yaExisteWorkflow) {
    appendRow_(SHEETS.MODULOS, {
      id: 'mod_workflow', modulo: 'Workflow',
      descripcion: 'Automatización de flujos de aprobación y procesos entre áreas.', activo: true, orden: 10
    }, ss);
    poblarModuloBasico_('Workflow');
  }

  const filaApi = modulos.find(function (m) { return m.modulo === 'API'; });
  if (filaApi && filaApi.activo !== false && filaApi.activo !== 'FALSE') {
    updateRowById_(SHEETS.MODULOS, filaApi.id, { activo: false }, ss);
  }
}

/** Borra todas las preguntas (y sus opciones) de `modulo` y vuelve a cargarlas con `construirFn`.
 * Solo toca PREGUNTAS/OPCIONES — no borra CONOCIMIENTO, PLANTILLAS ni diagnósticos ya guardados
 * de clientes (esos quedan tal cual se respondieron, con el árbol vigente en su momento). */
function reemplazarArbolPreguntas_(modulo, construirFn) {
  const ss = getSpreadsheet();
  const idsExistentes = getAllRows_(SHEETS.PREGUNTAS, ss)
    .filter(function (p) { return p.modulo === modulo; })
    .map(function (p) { return String(p.id); });

  idsExistentes.forEach(function (id) { eliminarFilasDondeIgual_(SHEETS.OPCIONES, 'pregunta_id', id, ss); });
  eliminarFilasDondeIgual_(SHEETS.PREGUNTAS, 'modulo', modulo, ss);

  construirFn();
}

/** Elimina (de abajo hacia arriba, para no desfasar índices) las filas de `nombreHoja` donde
 * `columna` sea igual a `valor`. Helper genérico de limpieza — usado solo por actualizaciones
 * de contenido que necesitan reemplazar filas ya cargadas, nunca por el flujo normal de la app. */
function eliminarFilasDondeIgual_(nombreHoja, columna, valor, spreadsheet) {
  const sheet = getSheet_(nombreHoja, spreadsheet);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return;
  const headers = values[0].map(function (h) { return String(h).trim(); });
  const col = headers.indexOf(columna);
  if (col === -1) return;
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][col]) === String(valor)) sheet.deleteRow(i + 1);
  }
}

/** Agrega una necesidad/oportunidad semilla a OPORTUNIDADES solo si esa `necesidad` no existe
 * todavía (evita duplicados si la función de actualización se corre más de una vez). */
function agregarOportunidadSiFalta_(necesidad, oportunidadBase) {
  const ss = getKnowledgeSpreadsheet();
  const yaExiste = getAllRows_(SHEETS.OPORTUNIDADES, ss).some(function (o) { return o.necesidad === necesidad; });
  if (yaExiste) return;
  appendRow_(SHEETS.OPORTUNIDADES, { id: generarId_('op'), necesidad: necesidad, oportunidad_base: oportunidadBase }, ss);
}

function agregarOportunidadesDiagnosticoComercial_() {
  agregarOportunidadSiFalta_('Definición de competencias y objetivos',
    'Construir un marco de competencias y objetivos claros por cargo y área, como base de un proceso de evaluación de desempeño estructurado.');
  agregarOportunidadSiFalta_('Implementación de evaluación de desempeño',
    'Formalizar un ciclo de evaluación de desempeño con metodología, periodicidad y evaluadores definidos.');
  agregarOportunidadSiFalta_('Implementación de encuesta de clima',
    'Levantar una primera medición de clima organizacional para contar con una línea base antes de intervenir.');
  agregarOportunidadSiFalta_('Implementación de programa de reconocimiento',
    'Diseñar un programa de reconocimiento formal, alineado a los valores corporativos que la organización quiere destacar.');
  agregarOportunidadSiFalta_('Gestión de beneficios centralizada',
    'Centralizar la administración, aprobación y comunicación de beneficios en una sola plataforma, dejando atrás procesos manuales o en Excel.');
  agregarOportunidadSiFalta_('Segmentación de consultas del colaborador',
    'Organizar los canales de atención al colaborador por tipo de consulta para reducir tiempos de respuesta y dar trazabilidad.');
  agregarOportunidadSiFalta_('Protocolo de canal de denuncias',
    'Formalizar un protocolo de escalamiento y un comité de ética/denuncias, alineado con los requerimientos legales del sector.');
}

// ---------------- Árboles de preguntas — Diagnóstico Comercial (superficial) ----------------
// Todas formuladas como pregunta directa en tercera persona (la responde el Comercial sobre el
// cliente). Donde existe una versión más profunda para el Kick Off, acá solo va la superficial.

function poblarPreguntasDesempenoComercial_() {
  const M = 'Gestión del Desempeño';

  pregunta_('p_desc_1', M, '¿Tienen definidas competencias y objetivos estructurados por cargo y área?', 'radio', 1, true, null, null, '');
  opciones_('p_desc_1', [
    ['Sí', '', '', 4],
    ['No', 'Definición de competencias y objetivos', 'Falta de objetivos y competencias estructuradas por cargo', 0]
  ]);

  pregunta_('p_desc_2', M, '¿Han realizado evaluaciones de desempeño antes?', 'radio', 2, true, null, null, '');
  opciones_('p_desc_2', [
    ['Sí', '', '', 3],
    ['No', 'Implementación de evaluación de desempeño', 'Ausencia de evaluaciones formales de desempeño', 0]
  ]);
  pregunta_('p_desc_2b', M, '¿Cómo realizan actualmente las evaluaciones de desempeño (plataforma)?', 'texto', 3, false, 'p_desc_2', 'Sí', '');

  pregunta_('p_desc_3', M, '¿Qué tipo de evaluadores participarían en la evaluación?', 'checkbox', 4, true, null, null, 'Selecciona todas las que apliquen.');
  opciones_('p_desc_3', [['Autoevaluación', '', '', 0], ['Pares', '', '', 0], ['Descendente', '', '', 0], ['Ascendente', '', '', 0]]);
}

function poblarPreguntasSeleccionComercial_() {
  const M = 'Selección';

  pregunta_('p_sel_1', M, '¿Cómo llevan a cabo sus procesos de selección hoy?', 'texto', 1, true, null, null, '');
  pregunta_('p_sel_2', M, '¿Utilizan Elempleo, Computrabajo, LinkedIn u otra plataforma?', 'checkbox', 2, false, null, null, 'Selecciona todas las que apliquen.');
  opciones_('p_sel_2', [['Elempleo', '', '', 0], ['Computrabajo', '', '', 0], ['LinkedIn', '', '', 0], ['Otra', '', '', 0]]);
  pregunta_('p_sel_3', M, '¿Cuál es el dominio del correo electrónico que manejan para agendar entrevistas?', 'texto', 3, false, null, null, '');
}

function poblarPreguntasEncuestasComercial_() {
  const M = 'Encuestas / Clima';

  pregunta_('p_enc_1', M, '¿Han realizado encuestas de clima anteriormente?', 'radio', 1, true, null, null, '');
  opciones_('p_enc_1', [
    ['Sí', '', '', 3],
    ['No', 'Implementación de encuesta de clima', 'Sin medición formal de clima organizacional', 0]
  ]);
  pregunta_('p_enc_1b', M, '¿Qué metodología han utilizado para sus encuestas de clima?', 'texto', 2, false, 'p_enc_1', 'Sí', '');
  pregunta_('p_enc_2', M, '¿Qué tipo de encuestas libres (adicionales a clima) desean realizar en Buk?', 'texto', 3, false, null, null, '');
}

function poblarPreguntasOnboardingComercial_() {
  const M = 'Onboarding';

  pregunta_('p_ob2_1', M, '¿Cuánto dura el proceso de onboarding actual, si existe?', 'texto', 1, true, null, null, '');
  pregunta_('p_ob2_2', M, '¿Quién lidera el onboarding hoy?', 'radio', 2, true, null, null, '');
  opciones_('p_ob2_2', [['RR. HH.', '', '', 0], ['Jefe directo', '', '', 0], ['Ambos', '', '', 0]]);
}

function poblarPreguntasReconocimientoComercial_() {
  const M = 'Reconocimiento';

  pregunta_('p_rec_1', M, '¿Actualmente manejan algún programa de reconocimiento (formal o informal)?', 'radio', 1, true, null, null, '');
  opciones_('p_rec_1', [
    ['Sí', '', '', 3],
    ['No', 'Implementación de programa de reconocimiento', 'Sin programa de reconocimiento formal', 0]
  ]);
  pregunta_('p_rec_1a', M, '¿De qué tipo es el programa de reconocimiento que manejan?', 'texto', 2, false, 'p_rec_1', 'Sí', '');
  pregunta_('p_rec_1b', M, '¿Qué valores corporativos les gustaría destacar al reconocer a las personas?', 'texto', 3, false, 'p_rec_1', 'No', '');
}

function poblarPreguntasComunicacionesComercial_() {
  const M = 'Comunicaciones';

  pregunta_('p_com_1', M, '¿Qué tipo de comunicaciones desean gestionar en Buk (fechas especiales, procesos, reglamentos)?', 'texto', 1, true, null, null, '');
  pregunta_('p_com_2', M, '¿Qué tan frecuente es la comunicación desde el liderazgo hacia los colaboradores?', 'radio', 2, true, null, null, '');
  opciones_('p_com_2', [
    ['Frecuente', '', '', 4],
    ['Ocasional', '', '', 2],
    ['Poco frecuente', 'Mejora de comunicación', 'Comunicación poco frecuente desde el liderazgo', 0]
  ]);
}

function poblarPreguntasBeneficiosComercial_() {
  const M = 'Beneficios';

  pregunta_('p_ben_1', M, '¿Qué beneficios ofrecen hoy en día, cómo se aprueban y cómo los administran (manual, Excel, otra plataforma)?', 'texto', 1, true, null, null, '');
  pregunta_('p_ben_2', M, '¿Tienen alianzas o convenios vigentes?', 'radio', 2, false, null, null, '');
  opciones_('p_ben_2', [['Sí', '', '', 0], ['No', '', '', 0]]);
  pregunta_('p_ben_3', M, '¿El presupuesto de beneficios está definido o se construirá desde cero?', 'radio', 3, true, null, null, '');
  opciones_('p_ben_3', [
    ['Definido', '', '', 3],
    ['Se construirá desde cero', 'Gestión de beneficios centralizada', 'Presupuesto de beneficios sin definir', 0]
  ]);
}

function poblarPreguntasServicioColaboradorComercial_() {
  const M = 'Servicio al Colaborador';

  pregunta_('p_sc_1', M, '¿Cuántos asientos se contrataron?', 'texto', 1, true, null, null, '');
  pregunta_('p_sc_2', M, '¿Cuentan hoy en día con un chat ya organizado por cada flujo de conversación?', 'radio', 2, true, null, null, '');
  opciones_('p_sc_2', [
    ['Sí', '', '', 3],
    ['No', 'Segmentación de consultas del colaborador', 'Canales de atención sin organizar por flujo de conversación', 0]
  ]);
  pregunta_('p_sc_3', M, '¿Qué tipo de consultas reciben con mayor frecuencia?', 'texto', 3, false, null, null, '');
}

function poblarPreguntasCanalDenunciasComercial_() {
  const M = 'Canal de Denuncias';

  pregunta_('p_cd_1', M, '¿Manejan hoy en día procesos de denuncia y cómo los escalan actualmente?', 'texto', 1, true, null, null, '');
  pregunta_('p_cd_2', M, '¿Cuentan con algún protocolo de escalamiento, comité ya conformado o requerimiento legal específico del sector?', 'radio', 2, true, null, null, '');
  opciones_('p_cd_2', [
    ['Sí', '', '', 3],
    ['No', 'Protocolo de canal de denuncias', 'Sin protocolo de escalamiento formal ni comité conformado', 0]
  ]);
}

// =====================================================================
// FLUJO "RADIOGRAFÍA DEL CLIENTE" (Kick Off): Información Comercial →
// Preguntas Orientadoras → Kick Off → Notas Gemini → Radiografía.
// La Radiografía es el resultado de cruzar el Diagnóstico Comercial con
// las notas de la sesión de Kick Off (pegadas o traídas desde un Google
// Doc por URL) — ver KickOff.gs -> generarRadiografia(). Esto NO
// reemplaza el flujo anterior de "Estructurar con IA" (contexto/
// necesidades/procesos/expectativas, columnas ya existentes en KICKOFF):
// se agregan columnas nuevas, aditivas, a la misma hoja.
//
// Para correrla: en el editor de Apps Script, selecciona
// "agregarRadiografiaKickOff" en el desplegable de funciones y
// presiona Ejecutar. Segura de correr más de una vez.
// =====================================================================

function agregarRadiografiaKickOff() {
  const ss = getSpreadsheet();
  if (!ss.getSheetByName(SHEETS.KICKOFF)) {
    crearHojaConEncabezados_(ss, SHEETS.KICKOFF, ['id', 'cliente_id', 'notas_gemini', 'contexto', 'necesidades_json',
      'procesos_json', 'expectativas_json', 'fecha', 'usuario']);
  }
  ['notas_gemini_url', 'fuente_notas', 'radiografia_json', 'alertas_json', 'fecha_radiografia']
    .forEach(function (col) { agregarColumnaSiFalta_(ss, SHEETS.KICKOFF, col); });

  Logger.log('Listo: columnas de Radiografía agregadas a KICKOFF (o ya existían). Nada existente fue modificado ni borrado.');
}
