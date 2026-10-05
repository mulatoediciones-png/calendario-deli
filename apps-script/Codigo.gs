/**
 * Calendario de Eventos Deli Snacks — Mulato Ediciones
 * API que guarda los eventos en la hoja de Google Sheets.
 * La página vive en GitHub (index.html). Este código se configura UNA sola vez:
 * los campos nuevos que se agreguen en index.html se crean solos como columnas.
 */

const SHEET_ID = '1iZWw_dl-ZH298enTw-B4gv_krfhzSXK1BSxHH9YF2Es';
const SHEET_NAME = 'Eventos';
const OLD_LIST_FIELDS = ['dinamicas', 'contenido']; // filas antiguas guardadas como "a | b"

function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') return json_(function () { return getEvents(); });
  return ContentService.createTextOutput('API del Calendario de Eventos Deli Snacks activa.');
}

function doPost(e) {
  return json_(function () {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (body.action === 'save') return saveEvent(body.event);
    if (body.action === 'delete') return deleteEvent(body.id);
    throw new Error('Acción no válida.');
  });
}

function json_(fn) {
  let out;
  try { out = { ok: true, events: fn() }; }
  catch (err) { out = { ok: false, error: String(err && err.message || err) }; }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function getSheet_() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    const first = ss.getSheets()[0];
    if (first && String(first.getRange(1, 1).getValue()) === 'id') { first.setName(SHEET_NAME); sh = first; }
    else sh = ss.insertSheet(SHEET_NAME);
  }
  if (sh.getLastRow() === 0) { sh.appendRow(['id']); sh.setFrozenRows(1); }
  return sh;
}

function header_(sh) {
  return sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
}

function toCell_(v) {
  if (v === undefined || v === null) return '';
  if (Array.isArray(v) || typeof v === 'object') return JSON.stringify(v);
  return String(v);
}

function fromCell_(field, v) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (/^hora/.test(field)) return Utilities.formatDate(v, tz, 'HH:mm');
    if (field === 'fecha') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
  }
  const s = v === null || v === undefined ? '' : String(v);
  if (s.charAt(0) === '[') { try { return JSON.parse(s); } catch (e) {} }
  if (OLD_LIST_FIELDS.indexOf(field) > -1) return s ? s.split(' | ').map(x => x.trim()).filter(Boolean) : [];
  return s;
}

function getEvents() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(String);
  const out = [];
  for (let r = 1; r < values.length; r++) {
    const ev = {};
    head.forEach((h, i) => { if (h) ev[h] = fromCell_(h, values[r][i]); });
    if (ev.id) out.push(ev);
  }
  return out;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

function saveEvent(ev) {
  if (!ev || !ev.id) throw new Error('Falta el id del evento.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    ev.updatedAt = new Date().toISOString();
    let head = header_(sh);
    const nuevos = Object.keys(ev).filter(k => head.indexOf(k) === -1);
    if (nuevos.length) {
      sh.getRange(1, head.length + 1, 1, nuevos.length).setValues([nuevos]);
      head = head.concat(nuevos);
    }
    let r = findRow_(sh, ev.id);
    const prev = r > -1 ? sh.getRange(r, 1, 1, head.length).getValues()[0] : [];
    if (r === -1) r = sh.getLastRow() + 1;
    const row = head.map((h, i) => (h in ev ? toCell_(ev[h]) : (prev[i] === undefined ? '' : prev[i])));
    const range = sh.getRange(r, 1, 1, head.length);
    range.setNumberFormat('@'); // texto plano: Sheets no cambia fechas ni teléfonos
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return getEvents();
}

function deleteEvent(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const r = findRow_(sh, id);
    if (r > -1) sh.deleteRow(r);
  } finally {
    lock.releaseLock();
  }
  return getEvents();
}

/* ============ Recordatorio por correo (7 a.m., eventos del día siguiente) ============ */

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function fechaLarga_(iso) {
  const p = iso.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2]);
  return DIAS[d.getDay()] + ' ' + p[2] + ' de ' + MESES[p[1] - 1];
}

function enviarRecordatorios() {
  const tz = Session.getScriptTimeZone();
  const manana = Utilities.formatDate(new Date(Date.now() + 24 * 3600 * 1000), tz, 'yyyy-MM-dd');
  const evs = getEvents().filter(e => e.fecha === manana);
  if (!evs.length) return;
  const items = evs.map(e =>
    '<li><b>' + esc_(e.comercio || 'Evento') + '</b> · ' + esc_(e.ciudad || '') +
    '<br>Llegada: ' + esc_(e.horaLlegada || 'por definir') + ' · Evento: ' + esc_(e.horaInicio || 'por definir') +
    (e.horaFin ? ' a ' + esc_(e.horaFin) : '') + ' · Retorno: ' + esc_(e.horaRetorno || 'por definir') +
    (e.direccion ? '<br>Dirección: ' + esc_(e.direccion) : '') +
    (e.responsable ? '<br>Responsable: ' + esc_(e.responsable) + ' ' + esc_(e.telefono || '') : '') +
    (e.maps ? '<br><a href="' + esc_(e.maps) + '">Ver en Google Maps</a>' : '') + '</li>').join('');
  MailApp.sendEmail({
    to: Session.getEffectiveUser().getEmail(),
    subject: 'Recordatorio: ' + evs.length + ' evento' + (evs.length > 1 ? 's' : '') + ' mañana, ' + fechaLarga_(manana),
    htmlBody: '<p>Mañana hay ' + evs.length + ' evento' + (evs.length > 1 ? 's' : '') + ' en el calendario de Deli Snacks:</p><ul>' + items + '</ul>',
    name: 'Calendario Deli Snacks'
  });
}

/** Ejecutar UNA vez desde el editor para activar el recordatorio diario. */
function activarRecordatorio() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'enviarRecordatorios').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarRecordatorios').timeBased().everyDays(1).atHour(7).create();
  return 'Recordatorio diario activado';
}
