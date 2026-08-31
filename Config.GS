/**
 * Config.gs
 * Responsabilidad única: configuración central de la aplicación.
 * Nada de lógica de negocio vive acá — solo constantes, acceso a
 * PropertiesService y helpers de bajo nivel para obtener el Spreadsheet.
 *
 * NINGÚN OTRO ARCHIVO debe leer PropertiesService directamente:
 * todo pasa por las funciones de este archivo, así el día de mañana
 * cambiamos de proveedor de IA o de spreadsheet sin tocar el resto.
 */

// ---------------------------------------------------------------------
// Nombres de hoja. Si cambian los encabezados de las hojas, este es el
// único lugar que hay que tocar (además de SetupSheets.gs).
// ---------------------------------------------------------------------
const SHEETS = {
  MODULOS: 'MODULOS',
  PREGUNTAS: 'PREGUNTAS',
  OPCIONES: 'OPCIONES',
  CONOCIMIENTO: 'CONOCIMIENTO',
  PLANTILLAS: 'PLANTILLAS',
  PROPUESTAS_VALOR: 'PROPUESTAS_VALOR',
  OPORTUNIDADES: 'OPORTUNIDADES',
  CLIENTES: 'CLIENTES',
  DIAGNOSTICOS: 'DIAGNOSTICOS',
  CONSULTORIAS: 'CONSULTORIAS',
  ENTREGABLES: 'ENTREGABLES'
};

// Hojas que forman la Base de Conocimiento (pueden vivir en un
// spreadsheet separado del transaccional — ver getKnowledgeSpreadsheet()).
const KNOWLEDGE_SHEETS = [
  SHEETS.MODULOS, SHEETS.PREGUNTAS, SHEETS.OPCIONES,
  SHEETS.CONOCIMIENTO, SHEETS.PLANTILLAS, SHEETS.PROPUESTAS_VALOR,
  SHEETS.OPORTUNIDADES
];

const ESTADOS_CONSULTORIA = {
  BORRADOR: 'Borrador',
  GENERADA: 'Consultoría generada',
  EN_REVISION: 'En revisión',
  APROBADA: 'Aprobada',
  ENTREGADA: 'Entregada'
};

const ORIGEN = {
  CLIENTE: 'CLIENTE',
  BASE_CONOCIMIENTO: 'BASE_CONOCIMIENTO',
  INFERENCIA_IA: 'INFERENCIA_IA'
};

const MADUREZ_NIVELES = ['Baja', 'Media', 'Alta'];

const CONFIG = {
  // Umbrales de puntaje acumulado (OPCIONES.peso_madurez) -> nivel de madurez.
  // Editable sin tocar lógica: ver Diagnostico.gs -> calcularMadurez_()
  MADUREZ_THRESHOLDS: { BAJA_MAX: 4, MEDIA_MAX: 8 }, // > MEDIA_MAX => Alta

  AI_PROVIDER: 'google', // Gemini API — tiene nivel gratuito sin tarjeta; ver AI.gs
  AI_MODEL_DEFAULT: 'gemini-3.6-flash', // gemini-2.5-flash quedó deprecado para cuentas nuevas; verificar vigencia en ai.google.dev/gemini-api/docs/pricing
  AI_MAX_TOKENS: 7000, // subido para permitir un diagnóstico más completo (3-5 párrafos) por módulo
  AI_TIMEOUT_MS: 55000, // UrlFetchApp no tiene timeout configurable; documental,
                         // usamos MuteHttpExceptions + validación de duración manual.
  AI_MAX_REINTENTOS: 1,

  MIN_PROPUESTAS_VALOR: 2,
  MAX_PROPUESTAS_VALOR: 3
};

/**
 * Devuelve el Spreadsheet transaccional (Clientes/Diagnósticos/Consultorías).
 * Prioridad: Script Property SPREADSHEET_ID -> spreadsheet activo (si el
 * script está vinculado a uno) -> error explicativo.
 */
function getSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      throw new Error(
        'No se pudo abrir el Spreadsheet configurado en SPREADSHEET_ID (' + id + '). ' +
        'Verifica el ID en Propiedades del proyecto. Detalle: ' + e.message
      );
    }
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  throw new Error(
    'No hay Spreadsheet configurado. Define la Script Property SPREADSHEET_ID ' +
    'o ejecuta este proyecto vinculado (bound) a una hoja de cálculo.'
  );
}

/**
 * Devuelve el Spreadsheet de Base de Conocimiento. Por defecto es el mismo
 * que el transaccional (pestañas distintas). Si el equipo COI quiere
 * administrar la metodología en un archivo separado (p. ej. para dar
 * permisos de edición distintos), basta con setear la Script Property
 * KNOWLEDGE_SPREADSHEET_ID. Ningún otro archivo necesita cambios.
 */
function getKnowledgeSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('KNOWLEDGE_SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      throw new Error('No se pudo abrir KNOWLEDGE_SPREADSHEET_ID (' + id + '): ' + e.message);
    }
  }
  return getSpreadsheet();
}

function getApiKey_() {
  const key = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  if (!key) {
    throw new AppError(
      'CONFIG_ERROR',
      'No se encontró la API Key. Configúrala en Extensiones > Apps Script > ' +
      'Configuración del proyecto > Propiedades del script, con la clave GEMINI_API_KEY ' +
      '(consíguela gratis en aistudio.google.com — ver GUIA_INSTALACION.md sección 5).'
    );
  }
  return key;
}

function getAiModel_() {
  return PropertiesService.getScriptProperties().getProperty('AI_MODEL') || CONFIG.AI_MODEL_DEFAULT;
}

function getUsuarioActual_() {
  try {
    const email = Session.getActiveUser().getEmail();
    return email || 'usuario_desconocido';
  } catch (e) {
    return 'usuario_desconocido';
  }
}

/**
 * Error tipado para poder distinguir en la UI errores de configuración,
 * de validación, de IA, o de Sheets, y mostrar mensajes claros (sección 22
 * del pedido). Ver AI.gs, Consultoria.gs, Code.gs.
 */
function AppError(codigo, mensaje) {
  this.name = 'AppError';
  this.codigo = codigo; // CONFIG_ERROR | VALIDACION | IA_ERROR | IA_TIMEOUT | SHEETS_ERROR | DATOS_INCOMPLETOS
  this.message = mensaje;
}
AppError.prototype = Object.create(Error.prototype);
