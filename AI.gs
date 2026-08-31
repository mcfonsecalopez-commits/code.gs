/**
 * KnowledgeBase.gs
 * Responsabilidad única: leer y filtrar la Base de Conocimiento COI
 * (hojas MODULOS, PREGUNTAS, OPCIONES, CONOCIMIENTO, PLANTILLAS,
 * PROPUESTAS_VALOR, OPORTUNIDADES).
 *
 * Este archivo es el que hay que entender para agregar un módulo nuevo:
 * ninguna función de acá referencia el nombre de un módulo específico,
 * todo se resuelve por lo que exista cargado en las hojas.
 */

function getModulosActivos() {
  return getAllRows_(SHEETS.MODULOS, getKnowledgeSpreadsheet())
    .filter(function (m) { return m.activo === true || m.activo === 'TRUE' || m.activo === 'VERDADERO'; })
    .sort(function (a, b) { return (a.orden || 0) - (b.orden || 0); })
    .map(function (m) { return { id: m.id, modulo: m.modulo, descripcion: m.descripcion }; });
}

/**
 * Devuelve el árbol de preguntas de un módulo: preguntas raíz con sus
 * opciones, y para cada opción, las preguntas hijas que dispara.
 * La UI recorre este árbol y solo pinta lo que corresponde según lo
 * que el usuario ya respondió (ver JsClient.html -> renderPreguntaActual).
 */
function getArbolPreguntas(modulo) {
  const ss = getKnowledgeSpreadsheet();
  const preguntas = getAllRows_(SHEETS.PREGUNTAS, ss).filter(function (p) { return p.modulo === modulo; });
  const opciones = getAllRows_(SHEETS.OPCIONES, ss);

  const opcionesPorPregunta = {};
  opciones.forEach(function (o) {
    if (!opcionesPorPregunta[o.pregunta_id]) opcionesPorPregunta[o.pregunta_id] = [];
    opcionesPorPregunta[o.pregunta_id].push({
      id: o.id,
      valor: o.valor,
      etiqueta_necesidad: o.etiqueta_necesidad,
      etiqueta_dolor: o.etiqueta_dolor,
      peso_madurez: Number(o.peso_madurez) || 0
    });
  });

  return preguntas
    .sort(function (a, b) { return (a.orden || 0) - (b.orden || 0); })
    .map(function (p) {
      return {
        id: p.id,
        pregunta: p.pregunta,
        tipo: p.tipo,
        obligatoria: !!p.obligatoria,
        ayuda: p.ayuda || '',
        preguntaPadreId: p.pregunta_padre_id || null,
        valorPadreDispara: p.valor_padre_dispara || null,
        opciones: opcionesPorPregunta[p.id] || []
      };
    });
}

/**
 * Filtra CONOCIMIENTO priorizando coincidencias específicas sobre "General".
 * Ejemplo: si existe una fila (modulo=Onboarding, industria=Retail,
 * tamaño=General, madurez=Baja) y otra (modulo=Onboarding, industria=General,
 * tamaño=General, madurez=Baja), se devuelven ambas, pero marcadas con su
 * nivel de especificidad para que Prompts.gs pueda darle más peso a la
 * primera. No se inventa contenido: si no hay filas, se devuelve [].
 */
function getConocimientoRelevante(modulo, industria, tamano, madurez) {
  const rows = getAllRows_(SHEETS.CONOCIMIENTO, getKnowledgeSpreadsheet())
    .filter(function (r) {
      if (r.activo === false || r.activo === 'FALSE') return false;
      if (r.modulo !== modulo) return false;
      const okIndustria = r.industria === 'General' || r.industria === industria;
      const okTamano = r.tamano === 'General' || r.tamano === tamano;
      const okMadurez = r.madurez === 'General' || r.madurez === madurez;
      return okIndustria && okTamano && okMadurez;
    });

  return rows.map(function (r) {
    const especificidad =
      (r.industria !== 'General' ? 1 : 0) +
      (r.tamano !== 'General' ? 1 : 0) +
      (r.madurez !== 'General' ? 1 : 0);
    return {
      id: r.id, tipo: r.tipo, categoria: r.categoria, contenido: r.contenido, especificidad: especificidad
    };
  }).sort(function (a, b) { return b.especificidad - a.especificidad; });
}

function getPlantillasRelevantes(modulo, industria) {
  return getAllRows_(SHEETS.PLANTILLAS, getKnowledgeSpreadsheet())
    .filter(function (r) {
      if (r.activo === false || r.activo === 'FALSE') return false;
      return r.modulo === modulo && (r.industria === 'General' || r.industria === industria);
    })
    .map(function (r) { return { id: r.id, tipo: r.tipo, nombre: r.nombre, contenido: r.contenido }; });
}

/** Propuestas de valor "semilla" para las necesidades detectadas por el motor de reglas. */
function getPropuestasValorBase(modulo, necesidades) {
  if (!necesidades || necesidades.length === 0) return [];
  return getAllRows_(SHEETS.PROPUESTAS_VALOR, getKnowledgeSpreadsheet())
    .filter(function (r) {
      if (r.activo === false || r.activo === 'FALSE') return false;
      return r.modulo === modulo && necesidades.indexOf(r.necesidad) !== -1;
    })
    .map(function (r) { return { necesidad: r.necesidad, propuesta: r.propuesta }; });
}

function getOportunidadesBase(necesidades) {
  if (!necesidades || necesidades.length === 0) return [];
  const todas = getAllRows_(SHEETS.OPORTUNIDADES, getKnowledgeSpreadsheet());
  const resultado = [];
  necesidades.forEach(function (n) {
    todas.filter(function (r) { return r.necesidad === n; })
      .forEach(function (r) { resultado.push(r.oportunidad_base); });
  });
  return resultado;
}
