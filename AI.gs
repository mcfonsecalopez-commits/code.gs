/**
 * AI.gs
 * Responsabilidad única: hablar con la API de IA por HTTP.
 * No construye prompts (eso es Prompts.gs) ni conoce el esquema de la
 * consultoría (eso lo valida Consultoria.gs). Este archivo solo sabe
 * mandar {system, user} y devolver el texto de respuesta, con manejo
 * de errores robusto (sección 22 del pedido).
 *
 * Proveedor: Google Gemini API (generativelanguage.googleapis.com), elegido
 * porque tiene un nivel gratuito real (sin tarjeta de crédito) vía Google AI
 * Studio — ver GUIA_INSTALACION.md sección 5. Si en el futuro se quiere
 * volver a Anthropic u otro proveedor, este es el ÚNICO archivo que hay que
 * tocar: construye {system, user} en Prompts.gs y todo lo demás no cambia,
 * porque todos consumen llamarIA().
 */

const GEMINI_ENDPOINT_BASE = 'https://generativelanguage.googleapis.com/v1beta/models/';

/**
 * @param {{system: string, user: string}} prompt
 * @return {string} texto crudo devuelto por el modelo (JSON, porque se
 *                   fuerza responseMimeType: application/json más abajo).
 * @throws {AppError} con codigo IA_ERROR | CONFIG_ERROR
 */
function llamarIA(prompt) {
  const apiKey = getApiKey_(); // lanza CONFIG_ERROR si no está configurada
  const modelo = getAiModel_();
  const endpoint = GEMINI_ENDPOINT_BASE + modelo + ':generateContent';

  const payload = {
    system_instruction: { parts: [{ text: prompt.system }] },
    contents: [{ role: 'user', parts: [{ text: prompt.user }] }],
    generationConfig: {
      maxOutputTokens: CONFIG.AI_MAX_TOKENS,
      temperature: 0.4,
      responseMimeType: 'application/json' // le pedimos a Gemini que devuelva JSON garantizado
    }
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  let intento = 0;
  let ultimoError = null;

  while (intento <= CONFIG.AI_MAX_REINTENTOS) {
    intento++;
    let response;
    try {
      response = UrlFetchApp.fetch(endpoint, options);
    } catch (e) {
      // Errores de red / DNS / servicio caído (UrlFetchApp lanza excepción en esos casos)
      ultimoError = new AppError('IA_ERROR', 'No se pudo contactar el servicio de IA: ' + e.message);
      continue;
    }

    const codigo = response.getResponseCode();
    const texto = response.getContentText();

    if (codigo === 200) {
      return extraerTextoRespuesta_(texto);
    }

    if (codigo === 400) {
      throw new AppError('IA_ERROR', 'La API de Gemini rechazó la solicitud (posible bloqueo de contenido o payload inválido). Detalle: ' + resumir_(texto));
    }
    if (codigo === 401 || codigo === 403) {
      throw new AppError('CONFIG_ERROR', 'La API Key de IA fue rechazada (código ' + codigo + '). Verifica GEMINI_API_KEY en Propiedades del script.');
    }
    if (codigo === 429) {
      ultimoError = new AppError('IA_ERROR', 'Se alcanzó el límite de solicitudes del nivel gratuito de Gemini. Espera un minuto e intenta de nuevo.');
      Utilities.sleep(2000);
      continue; // reintentable
    }
    if (codigo >= 500) {
      ultimoError = new AppError('IA_ERROR', 'El servicio de IA no está disponible en este momento (código ' + codigo + ').');
      continue; // reintentable
    }
    // otros códigos: no tiene sentido reintentar
    throw new AppError('IA_ERROR', 'La API de IA devolvió un error (código ' + codigo + '): ' + resumir_(texto));
  }

  throw ultimoError || new AppError('IA_ERROR', 'No se pudo obtener respuesta de la IA tras varios intentos.');
}

function extraerTextoRespuesta_(jsonTexto) {
  let data;
  try {
    data = JSON.parse(jsonTexto);
  } catch (e) {
    throw new AppError('IA_ERROR', 'La respuesta de la IA no es JSON válido a nivel HTTP.');
  }

  if (data.promptFeedback && data.promptFeedback.blockReason) {
    throw new AppError('IA_ERROR', 'Gemini bloqueó la generación por su filtro de contenido (motivo: ' + data.promptFeedback.blockReason + ').');
  }
  const candidato = data.candidates && data.candidates[0];
  const texto = candidato && candidato.content && candidato.content.parts && candidato.content.parts[0] && candidato.content.parts[0].text;
  if (!texto) {
    const razon = candidato ? candidato.finishReason : 'sin candidatos';
    throw new AppError('IA_ERROR', 'La respuesta de la IA no tiene el formato esperado (finishReason: ' + razon + ').');
  }
  return texto;
}

/**
 * Extrae y parsea el JSON que la IA debió devolver como único contenido.
 * Con responseMimeType: 'application/json' ya viene garantizado por Gemini,
 * pero igual toleramos que se haya "colado" un bloque ```json ... ``` y
 * damos un error claro si no es parseable en vez de fallar silenciosamente.
 */
function parsearRespuestaJSON(texto) {
  let limpio = texto.trim();
  const fenceMatch = limpio.match(/```(?:json)?\s*([\s\S]*?)```/);
  if (fenceMatch) limpio = fenceMatch[1].trim();

  try {
    return JSON.parse(limpio);
  } catch (e) {
    throw new AppError('IA_ERROR',
      'La IA devolvió una respuesta que no se pudo interpretar como JSON. ' +
      'Intenta generar la consultoría nuevamente. Detalle técnico: ' + e.message);
  }
}

function resumir_(texto) {
  return (texto || '').substring(0, 300);
}
