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
  '1. Solo puedes usar tres fuentes de información: (a) los datos del cliente y sus respuestas ' +
  'de diagnóstico que se te entregan textualmente, (b) los fragmentos de la Base de Conocimiento ' +
  'de COE Talento y Cultura que se te entregan, y (c) tu propio criterio experto para inferir recomendaciones razonables. ' +
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
 * @return {{system: string, user: string}}
 */
function construirPromptModulo(cliente, diagnosticoGeneral, diagnosticoModulo, conocimiento, plantillas, propuestasBase) {
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
      categoria: cliente.categoria, colaboradores: cliente.colaboradores, tamano: cliente.tamano
    },
    contexto_comercial: diagnosticoGeneral.contexto_comercial,
    objetivos_cliente: diagnosticoGeneral.objetivos,
    dolores_iniciales: diagnosticoGeneral.dolores_iniciales,
    info_adicional: diagnosticoGeneral.info_adicional,
    procesos_actuales_generales: diagnosticoGeneral.procesos_actuales,
    herramientas_utilizadas_actualmente: diagnosticoGeneral.herramientas_utilizadas,
    nivel_automatizacion_percibido: diagnosticoGeneral.nivel_automatizacion,
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
    'las reglas del sistema.\n\n' +
    'DATOS Y DIAGNÓSTICO (JSON):\n' + JSON.stringify(contexto, null, 2) + '\n\n' +
    'Genera entre ' + CONFIG.MIN_PROPUESTAS_VALOR + ' y ' + CONFIG.MAX_PROPUESTAS_VALOR + ' propuestas de valor ' +
    'en "propuestas_valor", cada una conectada al diagnóstico (no genéricas).\n\n' +
    'Responde EXCLUSIVAMENTE con un JSON que siga esta forma exacta (los valores son descripciones ' +
    'de lo que va en cada campo, no los copies literalmente):\n' + JSON.stringify(esquema, null, 2);

  return { system: PERSONA_SISTEMA, user: user };
}
