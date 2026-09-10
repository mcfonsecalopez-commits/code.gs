/**
 * Prompts.gs
 * Responsabilidad única: construir los prompts que se envían a la IA.
 * No hace llamadas HTTP (eso es AI.gs) ni lee Sheets directamente (eso
 * ya vino resuelto desde KnowledgeBase.gs / Diagnostico.gs como parámetros).
 *
 * El contrato de salida (JSON) está definido acá y DEBE mantenerse en
 * sincronía con lo que Consultoria.gs valida y con lo que la UI renderiza
 * (JsClient.html -> renderConsultoriaModulo).
 */

const PERSONA_SISTEMA =
  'Eres un Consultor Senior de Talento y Cultura y especialista en implementación ' +
  'de HR Tech, trabajando para el equipo COE Talento y Cultura ' +
  'de una plataforma de gestión de talento. No eres un chatbot genérico: tu trabajo es producir ' +
  'una consultoría de implementación específica, accionable y estandarizada para UN módulo ' +
  'contratado por UN cliente concreto.\n\n' +
  'REGLAS ESTRICTAS QUE DEBES RESPETAR SIEMPRE:\n' +
  '1. Solo puedes usar tres fuentes de información: (a) los datos del cliente, sus respuestas ' +
  'de diagnóstico y la información recogida durante el Kick Off (contexto, necesidades, procesos y ' +
  'expectativas conversadas con el cliente) que se te entregan textualmente, (b) los fragmentos de la ' +
  'Base de Conocimiento de COE Talento y Cultura que se te entregan, y (c) tu propio criterio experto para inferir recomendaciones razonables. ' +
  'NUNCA inventes datos específicos (cifras, nombres de herramientas, políticas) y los presentes como ' +
  'si fueran hechos del cliente o de la Base de Conocimiento.\n' +
  '2. Toda afirmación que generes debe poder clasificarse en uno de estos tres orígenes, y DEBES ' +
  'declararlo en el campo "trazabilidad" de tu respuesta: CLIENTE (viene literalmente de lo que el ' +
  'cliente o el comercial ingresaron), BASE_CONOCIMIENTO (viene de los fragmentos de metodología ' +
  'entregados), INFERENCIA_IA (es tu criterio experto o una buena práctica general, no un hecho ' +
  'confirmado de este cliente ni de la Base de Conocimiento).\n' +
  '3. Prioriza siempre el contenido de la Base de Conocimiento por sobre tu conocimiento general ' +
  'cuando ambos apliquen. Úsalo como fuente principal de metodología, plantillas y buenas prácticas.\n' +
  '4. Cuando uses conocimiento general del sector (no confirmado por el cliente ni por la Base de ' +
  'Conocimiento), trátalo explícitamente como una recomendación o hipótesis razonable, nunca como un ' +
  'hecho. Evita frases que suenen a dato verificado si no lo es.\n' +
  '5. Personaliza SIEMPRE por industria y por tamaño de la organización. No repitas el mismo texto ' +
  'para clientes de distinta industria o tamaño: ajusta complejidad, gobernanza y nivel de ' +
  'automatización recomendado según corresponda.\n' +
  '6. Evita absolutamente las recomendaciones genéricas del estilo "mejorar la experiencia de los ' +
  'colaboradores". Cada recomendación y cada propuesta de valor debe conectarse explícitamente con ' +
  'el diagnóstico entregado (necesidades, dolores, madurez) de este cliente.\n' +
  '7. Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional antes o después, sin bloques ' +
  'de markdown (```). Si no tienes información suficiente para un campo, usa un array vacío o string ' +
  'vacío — nunca inventes contenido de relleno.\n' +
  '8. El campo "diagnostico" debe ser un análisis completo de 3 a 5 párrafos (no una sola frase), que ' +
  'cubra explícitamente: (a) qué está pasando hoy en este proceso según las respuestas del cliente, ' +
  '(b) por qué está pasando así (causas probables, conectadas a las respuestas), (c) qué impacto tiene ' +
  'para el negocio y para el equipo de Personas si no se interviene, y (d) cómo se compara el nivel de ' +
  'madurez detectado con lo esperable para el tamaño y la industria de este cliente. Sé específico: cita ' +
  'las respuestas concretas del diagnóstico en vez de hablar en abstracto.\n' +
  '9. Los campos "necesidades", "oportunidades" y "buenas_practicas" deben tener entre 3 y 6 elementos ' +
  'cada uno, cada elemento con una frase completa y específica (no una sola palabra ni una etiqueta).';

/**
 * @param {Object} cliente
 * @param {Object} diagnosticoGeneral  { contexto_comercial, objetivos, dolores_iniciales, info_adicional }
 * @param {Object} diagnosticoModulo   salida de procesarRespuestasModulo()
 * @param {Array}  conocimiento        salida de getConocimientoRelevante()
 * @param {Array}  plantillas          salida de getPlantillasRelevantes()
 * @param {Array}  propuestasBase      salida de getPropuestasValorBase()
 * @param {Object} [kickOff]           salida de obtenerKickOff() — null si el cliente aún no tiene Kick Off (flujo clásico) registrado
 * @param {Object} [radiografia]       salida de obtenerRadiografia() — null si el cliente aún no tiene Radiografía generada; cuando existe, tiene prioridad sobre `kickOff` (ver construirContextoKickOffParaModulo_)
 * @return {{system: string, user: string}}
 */
function construirPromptModulo(cliente, diagnosticoGeneral, diagnosticoModulo, conocimiento, plantillas, propuestasBase, kickOff, radiografia) {
  const esquema = {
    diagnostico: 'string — 3 a 5 párrafos: qué pasa hoy, por qué, qué impacto tiene, y cómo se compara con lo esperable para su tamaño/industria (ver regla 8)',
    necesidades: ['string (3 a 6 elementos, frases completas y específicas)'],
    oportunidades: ['string (3 a 6 elementos, frases completas y específicas)'],
    consultoria_recomendada: 'string — qué debería implementarse y por qué',
    configuracion_recomendada: ['string — qué configurar en la plataforma'],
    plantillas_sugeridas: ['string'],
    comunicaciones: ['string — correos/comunicaciones a implementar'],
    tareas: ['string'],
    flujo_recomendado: 'string — cómo debería funcionar el proceso paso a paso',
    buenas_practicas: ['string (3 a 6 elementos) — relevantes para esta industria y tamaño'],
    prioridad: 'Alta | Media | Baja',
    primeros_pasos: ['string — qué debe hacer el equipo de COE Talento y Cultura en las primeras sesiones'],
    propuestas_valor: [
      { necesidad: 'string', solucion: 'string', beneficio_esperado: 'string', propuesta_completa: 'string (Necesidad -> Solución -> Beneficio en un párrafo accionable, NO genérico)' }
    ],
    trazabilidad: [
      { campo: 'string — nombre del campo de este JSON al que aplica', origen: 'CLIENTE | BASE_CONOCIMIENTO | INFERENCIA_IA', detalle: 'string breve' }
    ]
  };

  const contexto = {
    cliente: {
      nombre: cliente.nombre, pais: cliente.pais, industria: cliente.industria,
      categoria: cliente.categoria, tamano: cliente.tamano
    },
    contexto_comercial: diagnosticoGeneral.contexto_comercial,
    objetivos_cliente: diagnosticoGeneral.objetivos,
    dolores_iniciales: diagnosticoGeneral.dolores_iniciales,
    info_adicional: diagnosticoGeneral.info_adicional,
    procesos_actuales_generales: diagnosticoGeneral.procesos_actuales,
    herramientas_utilizadas_actualmente: diagnosticoGeneral.herramientas_utilizadas,
    nivel_automatizacion_percibido: diagnosticoGeneral.nivel_automatizacion,
    informacion_kickoff: construirContextoKickOffParaModulo_(diagnosticoModulo.modulo, kickOff, radiografia),
    modulo: diagnosticoModulo.modulo,
    diagnostico_estructurado_por_reglas: {
      madurez: diagnosticoModulo.madurez,
      necesidades_detectadas: diagnosticoModulo.principales_necesidades,
      dolores_detectados: diagnosticoModulo.dolores,
      oportunidades_semilla: diagnosticoModulo.oportunidades
    },
    respuestas_textuales_del_diagnostico: diagnosticoModulo.respuestas_originales,
    fragmentos_base_de_conocimiento: conocimiento.map(function (c) { return { tipo: c.tipo, categoria: c.categoria, contenido: c.contenido }; }),
    plantillas_disponibles: plantillas.map(function (p) { return { tipo: p.tipo, nombre: p.nombre, contenido: p.contenido }; }),
    propuestas_valor_semilla: propuestasBase
  };

  const user =
    'Genera la consultoría de implementación para el módulo "' + diagnosticoModulo.modulo + '" ' +
    'del siguiente cliente, usando exclusivamente la información entregada abajo y respetando ' +
    'las reglas del sistema. Si "informacion_kickoff" no es null, dale prioridad: es información que el ' +
    'cliente conversó directamente con el equipo COE y debe personalizar el diagnóstico y las ' +
    'recomendaciones (trátala como origen CLIENTE en el campo "trazabilidad"). Si "informacion_kickoff" ' +
    'trae "alertas_inconsistencia" con contenido, menciónalas explícitamente en el campo "diagnostico" ' +
    'como puntos que el equipo COE debe validar con el cliente antes de configurar — nunca las ignores ' +
    'ni asumas cuál de las dos fuentes (comercial o Kick Off) es la correcta.\n\n' +
    'DATOS Y DIAGNÓSTICO (JSON):\n' + JSON.stringify(contexto, null, 2) + '\n\n' +
    'Genera entre ' + CONFIG.MIN_PROPUESTAS_VALOR + ' y ' + CONFIG.MAX_PROPUESTAS_VALOR + ' propuestas de valor ' +
    'en "propuestas_valor", cada una conectada al diagnóstico (no genéricas).\n\n' +
    'Responde EXCLUSIVAMENTE con un JSON que siga esta forma exacta (los valores son descripciones ' +
    'de lo que va en cada campo, no los copies literalmente):\n' + JSON.stringify(esquema, null, 2);

  return { system: PERSONA_SISTEMA, user: user };
}

/**
 * Arma el bloque "informacion_kickoff" del prompt de Consultoría — el punto donde
 * el flujo de Radiografía (Kick Off -> Notas Gemini -> Radiografía, ver KickOff.gs)
 * se conecta con la Consultoría existente sin cambiar su UI ni su flujo visible.
 *
 * Prioridad de fuentes:
 *  1. Radiografía (si existe): ya cruzó Diagnóstico Comercial + Notas de Kick Off
 *     y clasificó la información por módulo — se le pasa a la IA el bloque
 *     específico de ESTE módulo (hallazgos/decisiones/pendientes), más el
 *     contexto general y las alertas de inconsistencia detectadas.
 *  2. Kick Off "clásico" (flujo anterior a este ajuste, sin Radiografía): se usa
 *     como respaldo para no romper clientes que ya venían de ese flujo.
 *  3. Ninguno de los dos: devuelve null (comportamiento idéntico al que ya existía
 *     antes de la Radiografía — la consultoría se genera solo con el Diagnóstico).
 * @param {string} modulo
 * @param {Object} [kickOff]      salida de obtenerKickOff()
 * @param {Object} [radiografia]  salida de obtenerRadiografia()
 * @return {Object|null}
 */
function construirContextoKickOffParaModulo_(modulo, kickOff, radiografia) {
  if (radiografia && radiografia.radiografia) {
    const r = radiografia.radiografia;
    const bloqueModulo = (r.por_modulo && (r.por_modulo[modulo] || r.por_modulo['Información transversal / general'])) || {};
    return {
      fuente: 'RADIOGRAFIA',
      contexto_cliente: r.contexto_cliente || '',
      necesidades_identificadas_en_kickoff: r.necesidades_principales || [],
      procesos_conversados_en_kickoff: r.procesos_actuales || [],
      expectativas_del_cliente: r.objetivos_implementacion || [],
      hallazgos_especificos_de_este_modulo: bloqueModulo.hallazgos || [],
      decisiones_tomadas_para_este_modulo: bloqueModulo.decisiones || [],
      pendientes_de_este_modulo: bloqueModulo.pendientes || [],
      riesgos_generales_del_cliente: r.riesgos || [],
      configuraciones_requeridas_generales: r.configuraciones_requeridas || [],
      alertas_inconsistencia: r.alertas_inconsistencia || []
    };
  }
  if (kickOff) {
    return {
      fuente: 'KICKOFF_CLASICO',
      contexto_conversado_con_el_cliente: kickOff.contexto,
      necesidades_identificadas_en_kickoff: kickOff.necesidades,
      procesos_conversados_en_kickoff: kickOff.procesos,
      expectativas_del_cliente: kickOff.expectativas
    };
  }
  return null;
}

// ---------------------------------------------------------------------------
// Kick Off COE (sección 3 del rediseño): estructura las notas/transcripción
// que Gemini genera durante la sesión con el cliente. Es un prompt mucho más
// acotado que el de consultoría: solo EXTRAE lo que el texto realmente dice,
// nunca completa huecos con inferencia — si algo no aparece en las notas,
// el array/campo correspondiente queda vacío.
// ---------------------------------------------------------------------------

const PERSONA_KICKOFF =
  'Eres un asistente que ayuda al equipo COE Talento y Cultura a estructurar las notas de una ' +
  'sesión de Kick Off con un cliente. Tu única tarea es EXTRAER y ORGANIZAR información que ya está ' +
  'presente en el texto que te entregan — nunca inventar, suponer ni completar con conocimiento general. ' +
  'Si el texto no menciona algo, el campo correspondiente debe quedar vacío (string vacío o array vacío). ' +
  'Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional antes o después, sin bloques de markdown.';

/**
 * @param {string} notasGemini  texto crudo (notas o transcripción) pegado por el COE
 * @return {{system: string, user: string}}
 */
function construirPromptKickOff(notasGemini) {
  const esquema = {
    contexto: 'string — 1 a 2 párrafos resumiendo el contexto del cliente tal como se conversó en la sesión (situación actual, quién participó, qué se buscaba con la sesión)',
    necesidades: ['string — necesidades específicas que el cliente mencionó o que se desprenden explícitamente de la conversación'],
    procesos: ['string — procesos actuales del cliente que se describieron en la sesión'],
    expectativas: ['string — expectativas o resultados que el cliente espera de la implementación, tal como las expresó']
  };

  const user =
    'Estas son las notas o la transcripción (generadas por Gemini) de una sesión de Kick Off entre el ' +
    'equipo COE Talento y Cultura y un cliente:\n\n---\n' + notasGemini + '\n---\n\n' +
    'Estructura esta información EXCLUSIVAMENTE con lo que dice el texto de arriba, respondiendo con un ' +
    'JSON que siga esta forma exacta (los valores son descripciones de lo que va en cada campo, no los ' +
    'copies literalmente):\n' + JSON.stringify(esquema, null, 2);

  return { system: PERSONA_KICKOFF, user: user };
}

// ---------------------------------------------------------------------------
// Radiografía del Cliente: flujo Información Comercial → Preguntas
// Orientadoras → Kick Off → Notas Gemini → Radiografía. Cruza DOS insumos
// (Diagnóstico Comercial ya guardado + Notas Gemini de la sesión de Kick
// Off) para producir una síntesis ejecutiva por módulo contratado, sin
// inventar información y señalando cualquier contradicción entre ambas
// fuentes en vez de decidir cuál es la correcta (regla explícita del
// pedido — sección 6).
// ---------------------------------------------------------------------------

const PERSONA_RADIOGRAFIA =
  'Eres un analista del equipo COE Talento y Cultura encargado de producir la "Radiografía del Cliente": ' +
  'una síntesis ejecutiva que cruza DOS fuentes de información — (a) el Diagnóstico Comercial, ya ' +
  'registrado en el sistema antes del Kick Off, y (b) las Notas de la sesión de Kick Off (generadas por ' +
  'Gemini), pegadas o traídas por el equipo COE. NO son la misma cosa: el Diagnóstico Comercial es lo que ' +
  'el comercial registró ANTES de reunirse con el cliente; las Notas de Kick Off son lo que se conversó ' +
  'DURANTE la reunión.\n\n' +
  'REGLAS ESTRICTAS:\n' +
  '1. La Radiografía es una síntesis ejecutiva, NUNCA una copia o transcripción de las notas. Sintetiza, no ' +
  'repitas literalmente párrafos completos.\n' +
  '2. NUNCA inventes información que no esté en el Diagnóstico Comercial ni en las Notas de Kick Off. Si un ' +
  'campo no tiene información suficiente en ninguna de las dos fuentes, el array o el texto correspondiente ' +
  'debe quedar vacío — nunca lo rellenes con supuestos ni con conocimiento general.\n' +
  '3. Debes CRUZAR ambas fuentes, nunca reemplazar una con la otra: si el Diagnóstico Comercial dice algo y ' +
  'las Notas de Kick Off agregan información nueva sobre lo mismo, consolida ambas en un solo hallazgo. Si ' +
  'las dos fuentes se CONTRADICEN entre sí sobre el mismo tema, NO decidas cuál es la correcta: repórtalo ' +
  'como un elemento del array "alertas_inconsistencia", describiendo con precisión qué dice cada fuente, ' +
  'para que el COE lo valide directamente con el cliente.\n' +
  '4. Debes clasificar cada hallazgo, decisión o pendiente identificado en las Notas de Kick Off dentro del ' +
  'módulo contratado al que corresponde (de la lista de módulos contratados que se te entrega). Si un ' +
  'hallazgo no corresponde a ningún módulo contratado específico (por ejemplo, gobernanza general del ' +
  'proyecto, disponibilidad del equipo, aprobaciones internas), clasifícalo en el módulo especial ' +
  '"Información transversal / general".\n' +
  '5. Identifica específicamente, cuando la información esté presente en las notas: decisiones tomadas ' +
  'durante el Kick Off, responsables (nombre y rol si se mencionan), pendientes que el cliente debe ' +
  'entregar o validar, riesgos o posibles bloqueos, e información o configuraciones que aún faltan para ' +
  'poder parametrizar la plataforma.\n' +
  '6. Todo elemento que generes debe poder clasificarse en el campo "trazabilidad" según su origen: ' +
  'DIAGNOSTICO_COMERCIAL (viene del diagnóstico registrado antes del Kick Off), NOTAS_KICKOFF (viene de las ' +
  'notas de la sesión), o CRUCE (surge de combinar ambas fuentes).\n' +
  '7. Responde ÚNICAMENTE con un objeto JSON válido, sin texto adicional antes o después, sin bloques de ' +
  'markdown (```).';

/**
 * @param {Object} cliente
 * @param {Object} diagnosticoGeneral  salida de construirDiagnosticoGeneral()
 * @param {Array<Object>} diagnosticosPorModulo  [{modulo, diagnostico_json}] — diagnósticos ya calculados por reglas
 * @param {Array<string>} modulosContratados
 * @param {string} notasGemini  texto de las notas de Kick Off (pegadas o traídas desde el Google Doc)
 * @param {Object|null} radiografiaPrevia  si ya existía una Radiografía anterior para este cliente, se
 *   entrega como contexto adicional (por ejemplo, si se está regenerando con notas ampliadas)
 * @return {{system: string, user: string}}
 */
function construirPromptRadiografia(cliente, diagnosticoGeneral, diagnosticosPorModulo, modulosContratados, notasGemini, radiografiaPrevia) {
  const esquema = {
    contexto_cliente: 'string — 2 a 4 frases: quién es el cliente y cómo funciona actualmente (industria, tamaño, estructura relevante)',
    necesidades_principales: ['string — problemas que el cliente busca resolver, cruzando diagnóstico y notas'],
    procesos_actuales: ['string — cómo realiza hoy sus procesos, cruzando diagnóstico y notas'],
    objetivos_implementacion: ['string — qué espera conseguir el cliente con Buk'],
    hallazgos_kickoff: ['string — información NUEVA que solo aparece en las notas de Kick Off (no estaba en el diagnóstico comercial)'],
    riesgos: ['string — posibles bloqueos, dependencias o alertas para la implementación'],
    definiciones_tomadas: ['string — decisiones concretas obtenidas durante el Kick Off'],
    pendientes: ['string — información, documentos o decisiones que el cliente aún debe entregar o validar'],
    responsables: [{ nombre: 'string', rol: 'string — ej. Líder de implementación, Usuario administrador, Aprobador' }],
    configuraciones_requeridas: ['string — definiciones que deben llevarse posteriormente a la parametrización de la plataforma'],
    por_modulo: 'objeto — una clave por cada módulo contratado (usa exactamente los nombres de la lista de módulos contratados que se te entrega, más la clave "Información transversal / general"), cada una con la forma: { "hallazgos": ["string"], "decisiones": ["string"], "pendientes": ["string"] }',
    alertas_inconsistencia: ['string — descripción de cada contradicción encontrada entre el diagnóstico comercial y las notas de Kick Off, indicando qué dice cada fuente'],
    trazabilidad: [{ campo: 'string — nombre del campo de este JSON', origen: 'DIAGNOSTICO_COMERCIAL | NOTAS_KICKOFF | CRUCE', detalle: 'string breve' }]
  };

  const contexto = {
    cliente: { nombre: cliente.nombre, industria: cliente.industria, tamano: cliente.tamano, categoria: cliente.categoria },
    modulos_contratados: modulosContratados,
    diagnostico_comercial: diagnosticoGeneral,
    diagnosticos_por_modulo: (diagnosticosPorModulo || []).map(function (d) {
      return { modulo: d.modulo, madurez: d.diagnostico_json && d.diagnostico_json.madurez, necesidades_detectadas: d.diagnostico_json && d.diagnostico_json.principales_necesidades };
    }),
    radiografia_anterior: radiografiaPrevia || null
  };

  const user =
    'INSUMO 1 — Diagnóstico Comercial y datos ya registrados del cliente (JSON):\n' +
    JSON.stringify(contexto, null, 2) + '\n\n' +
    'INSUMO 2 — Notas de la sesión de Kick Off (texto tal como las entregó Gemini o el equipo COE):\n' +
    '---\n' + notasGemini + '\n---\n\n' +
    'Cruza ambos insumos y genera la Radiografía del Cliente siguiendo TODAS las reglas del sistema. ' +
    'Responde EXCLUSIVAMENTE con un JSON que siga esta forma exacta (los valores son descripciones de lo ' +
    'que va en cada campo, no los copies literalmente):\n' + JSON.stringify(esquema, null, 2);

  return { system: PERSONA_RADIOGRAFIA, user: user };
}
