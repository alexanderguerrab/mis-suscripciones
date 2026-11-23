// ===== CONFIGURA ESTO =====
const CONFIG = {
  TOKEN: 'cambia-esta-clave',     // la misma clave que pondrás en la app
  EMAIL: 'tucorreo@gmail.com',
  TELEGRAM_BOT_TOKEN: '',         // de @BotFather (opcional)
  TELEGRAM_CHAT_ID: '',           // tu chat id (opcional)
  HORA_AVISO: 8,                  // hora del aviso diario
  SHEET_ID: ''                    // déjalo vacío si el script se creó desde la hoja (Extensiones > Apps Script)
};
const HOJA = 'Suscripciones';
const COLS = ['id', 'servicio', 'costo', 'vencimiento', 'avisar', 'notas', 'periodo', 'categoria'];

// Ejecuta setup() UNA vez: crea la hoja y el aviso diario
function setup() {
  const sh = hoja_();
  sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  sh.getRange('D:D').setNumberFormat('@');
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'revisar').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('revisar').timeBased().everyDays(1).atHour(CONFIG.HORA_AVISO).create();
}

const ok_ = (email, token) => token === CONFIG.TOKEN && String(email || '').trim().toLowerCase() === CONFIG.EMAIL.toLowerCase();
function hoja_() {
  const ss = CONFIG.SHEET_ID ? SpreadsheetApp.openById(CONFIG.SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA);
  if (!sh) {
    sh = ss.insertSheet(HOJA);
    sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
    sh.getRange('D:D').setNumberFormat('@');
  }
  // Hojas creadas antes de las categorías: agrega el encabezado que falta
  if (sh.getLastColumn() < COLS.length) sh.getRange(1, 1, 1, COLS.length).setValues([COLS]);
  return sh;
}
const fmt_ = d => d instanceof Date ? Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(d);
const dias_ = iso => Math.round((new Date(iso + 'T00:00:00Z') - new Date(fmt_(new Date()) + 'T00:00:00Z')) / 864e5);
const responder_ = o => ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);

function leer_() {
  const v = hoja_().getDataRange().getValues(); v.shift();
  return v.filter(r => r[0]).map(r => {
    const o = {}; COLS.forEach((c, i) => o[c] = r[i]);
    o.vencimiento = fmt_(o.vencimiento);
    return o;
  });
}

function doGet_(e) {
  if (!ok_(e.parameter.email, e.parameter.token)) return responder_({ error: 'credenciales' });
  return responder_({ items: leer_() });
}

function doPost_(e) {
  const b = JSON.parse(e.postData.contents);
  if (!ok_(b.email, b.token)) return responder_({ error: 'credenciales' });
  const lock = LockService.getScriptLock(); lock.waitLock(10000);
  try {
    const sh = hoja_();
    const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues().flat();
    const id = b.id || (b.item && b.item.id);
    const fila = ids.indexOf(id) + 1;
    if (b.action === 'delete') {
      if (fila > 1) sh.deleteRow(fila);
    } else {
      const it = b.item;
      const row = COLS.map(c => it[c] ?? '');
      if (fila > 1) sh.getRange(fila, 1, 1, COLS.length).setValues([row]); else sh.appendRow(row);
    }
    return responder_({ ok: true });
  } finally { lock.releaseLock(); }
}

// Se ejecuta cada día: avisa por correo y Telegram
function revisar() {
  const lista = leer_().map(s => ({ ...s, d: dias_(s.vencimiento) }));
  const vencidas = lista.filter(s => s.d < 0);
  // Avisos escalonados: el día del plazo, y luego a 14, 7, 3, 1 y 0 días
  const hitos = s => [Number(s.avisar || 7), 14, 7, 3, 1, 0].filter(h => h <= Number(s.avisar || 7));
  const pronto = lista.filter(s => s.d >= 0 && hitos(s).includes(s.d));
  if (!vencidas.length && !pronto.length) return;
  const linea = s => `• ${s.servicio} (${s.costo || 's/costo'}): ` +
    (s.d < 0 ? `venció hace ${-s.d} día(s)` : s.d === 0 ? 'vence HOY' : `vence en ${s.d} día(s)`) + ` [${s.vencimiento}]`;
  const texto = [
    vencidas.length ? '⛔ VENCIDAS\n' + vencidas.map(linea).join('\n') : '',
    pronto.length ? '⏰ POR VENCER\n' + pronto.map(linea).join('\n') : ''
  ].filter(Boolean).join('\n\n');
  MailApp.sendEmail(CONFIG.EMAIL, `Suscripciones: ${vencidas.length + pronto.length} requieren atención`, texto);
  if (CONFIG.TELEGRAM_BOT_TOKEN && CONFIG.TELEGRAM_CHAT_ID) {
    UrlFetchApp.fetch(`https://api.telegram.org/bot${CONFIG.TELEGRAM_BOT_TOKEN}/sendMessage`,
      { method: 'post', payload: { chat_id: CONFIG.TELEGRAM_CHAT_ID, text: texto } });
  }
}

// Si el script falla, responde el motivo en JSON (en vez de una página de error de Google)
function doGet(e) { try { return doGet_(e); } catch (err) { return responder_({ error: 'script', detalle: String(err) }); } }
function doPost(e) { try { return doPost_(e); } catch (err) { return responder_({ error: 'script', detalle: String(err) }); } }
