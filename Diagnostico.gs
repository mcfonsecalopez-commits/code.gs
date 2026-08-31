/**
 * Diagnostico.gs
 * Responsabilidad única: transformar respuestas crudas del wizard en un
 * diagnóstico estructurado, usando ÚNICAMENTE reglas (sin IA — ver
 * decisión de arquitectura en la sección 23 del pedido). La IA entra
 * después, en Consultoria.gs, para interpretar y redactar — nunca acá.
 *
 * Esto garantiza el requisito "no debe inventar respuestas que el
 * usuario no haya proporcionado": el diagnóstico estructurado es un
 * cálculo puro sobre datos que el usuario efectivamente ingresó.
 */

/**
 * @param {string} modulo
 * @param {Object} respuestas  { preguntaId: valorSeleccionado }
 * @return {Object} diagnostico estructurado, ver forma exacta abajo.
 */
function procesarRespuestasModulo(modulo, respuestas) {
  if (!respuestas || Object.keys(respuestas).length === 0) {
    throw new AppError('DATOS_INCOMPLETOS', 'No se recibieron respuestas para el módulo ' + modulo + '.');
  }

  const arbol = getArbolPreguntas(modulo);
  const preguntasPorId = {};
  arbol.forEach(function (p) { preguntasPorId[p.id] = p; });

  let puntajeMadurez = 0;
  let preguntasConPeso = 0;
  const necesidades = {};
  const dolores = {};

  Object.keys(respuestas).forEach(function (preguntaId) {
    const pregunta = preguntasPorId[preguntaId];
    if (!pregunta) return; // pregunta de otro módulo o inexistente: se ignora, no se inventa
    const valor = respuestas[preguntaId];
    const valores = Array.isArray(valor) ? valor : [valor]; // soporta checkbox multi-valor

    valores.forEach(function (v) {
      const opcion = pregunta.opciones.find(function (o) { return o.valor === v; });
      if (!opcion) return; // valor de texto libre: no aporta al cálculo de reglas, pero se conserva en respuestas_originales
      if (opcion.etiqueta_necesidad) necesidades[opcion.etiqueta_necesidad] = true;
      if (opcion.etiqueta_dolor) dolores[opcion.etiqueta_dolor] = true;
      if (opcion.peso_madurez !== undefined) {
        puntajeMadurez += opcion.peso_madurez;
        preguntasConPeso++;
      }
    });
  });

  const necesidadesArr = Object.keys(necesidades);
  const doloresArr = Object.keys(dolores);
  const madurez = calcularMadurez_(puntajeMadurez, preguntasConPeso);
  const oportunidadesBase = getOportunidadesBase(necesidadesArr);

  return {
    modulo: modulo,
    madurez: madurez,
    puntaje_madurez: puntajeMadurez,
    principales_necesidades: necesidadesArr,
    dolores: doloresArr,
    oportunidades: oportunidadesBase, // texto semilla desde OPORTUNIDADES; la IA las enriquece luego, marcadas como INFERENCIA_IA
    respuestas_originales: respuestas, // se conserva tal cual la entregó el usuario — nunca se sobreescribe
    origen: 'reglas'
  };
}

/**
 * Promedia el puntaje y lo mapea a un nivel de madurez usando los
 * umbrales de Config.gs. Editable sin tocar lógica.
 */
function calcularMadurez_(puntajeTotal, cantidadPreguntasConPeso) {
  if (cantidadPreguntasConPeso === 0) return 'Sin determinar';
  // Normalizamos a una escala comparable independientemente de cuántas
  // preguntas con peso respondió el usuario (árbol condicional = distinto
  // número de preguntas por caso).
  const promedio = puntajeTotal / cantidadPreguntasConPeso; // 0-4
  const escalado = promedio * 3; // aprox. a escala 0-12 para usar los mismos umbrales que un cuestionario de 3 preguntas
  if (escalado <= CONFIG.MADUREZ_THRESHOLDS.BAJA_MAX) return 'Baja';
  if (escalado <= CONFIG.MADUREZ_THRESHOLDS.MEDIA_MAX) return 'Media';
  return 'Alta';
}

/**
 * Diagnóstico general (Etapa 1, no específico de módulo): no hay reglas
 * que "calcular" acá, es información que el comercial ya escribió — se
 * pasa tal cual para que la IA la use como contexto, jamás para
 * reinterpretarla como si fuera del cliente sin que lo sea.
 */
function construirDiagnosticoGeneral(cliente) {
  return {
    contexto_comercial: cliente.contexto_comercial || '',
    objetivos: cliente.objetivos || '',
    dolores_iniciales: cliente.dolores_iniciales || '',
    info_adicional: cliente.info_adicional || '',
    procesos_actuales: cliente.procesos_actuales || '',
    herramientas_utilizadas: cliente.herramientas_utilizadas || '',
    nivel_automatizacion: cliente.nivel_automatizacion || ''
  };
}
