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

  crearHojaConEncabezados_(ss, SHEETS.CLIENTES, ['id', 'nombre', 'pais', 'url', 'industria', 'categoria', 'colaboradores', 'tamano',
    'modulos_contratados', 'contexto_comercial', 'objetivos', 'dolores_iniciales', 'info_adicional',
    'procesos_actuales', 'herramientas_utilizadas', 'nivel_automatizacion', 'fecha_creacion', 'usuario']);
  crearHojaConEncabezados_(ss, SHEETS.DIAGNOSTICOS, ['id', 'cliente_id', 'modulo', 'respuestas_json', 'diagnostico_json', 'fecha']);
  crearHojaConEncabezados_(ss, SHEETS.CONSULTORIAS, ['id', 'cliente_id', 'modulo', 'diagnostico_id', 'contenido_json', 'version',
    'estado', 'fecha_generacion', 'fecha_aprobacion', 'usuario_aprobador']);
  crearHojaConEncabezados_(ss, SHEETS.ENTREGABLES, ['id', 'cliente_id', 'consultoria_ids', 'tipo', 'url', 'fecha', 'usuario']);

  poblarModulos_();
  poblarOnboardingCompleto_();
  ['Gestión del Desempeño', 'Selección', 'Encuestas / Clima', 'Reconocimiento', 'Comunicaciones',
    'Beneficios', 'Canal de Denuncias', 'Servicio al Colaborador', 'API'].forEach(poblarModuloBasico_);
  poblarPropuestasValorYOportunidades_();

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
    ['mod_api', 'API', 'Integraciones y automatizaciones vía API con otros sistemas del cliente.', true, 10]
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

function slug_(texto) {
  return texto.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, '_').replace(/^_|_$/g, '');
}
