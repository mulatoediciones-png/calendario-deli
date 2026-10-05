/**
 * Calendario de Eventos Deli Snacks — Mulato Ediciones
 * Backend en Google Apps Script. Guarda los eventos en la pestaña "Eventos"
 * de la hoja de Google Sheets donde está este script.
 */

const SHEET_ID = '1iZWw_dl-ZH298enTw-B4gv_krfhzSXK1BSxHH9YF2Es';
const SHEET_NAME = 'Eventos';
const FIELDS = ['id', 'estado', 'comercio', 'instagram', 'tipoCliente', 'responsable', 'telefono',
  'ciudad', 'direccion', 'referencia', 'maps', 'tipoEvento', 'tipoEventoOtro', 'descripcion',
  'fecha', 'horaInicio', 'horaFin', 'mascota', 'premios', 'dinamicas', 'dinamicaOtra',
  'contenido', 'notas', 'updatedAt',
  'horaLlegada', 'horaRetorno', 'materialPop', 'dinamicaPersonalizada', 'mensajeGerencia', 'reporteUrl'];
const LIST_FIELDS = ['dinamicas', 'contenido'];
const LIST_SEP = ' | ';

/** GET: ?action=list devuelve los eventos en JSON; sin parámetros muestra la app. */
function doGet(e) {
  const action = e && e.parameter && e.parameter.action;
  if (action === 'list') return json_(function () { return getEvents(); });
  return HtmlService.createHtmlOutput(APP_HTML)
    .setTitle('Calendario de Eventos Deli Snacks')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/** POST: {action:'save', event:{...}} o {action:'delete', id:'...'} */
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
    if (first && String(first.getRange(1, 1).getValue()) === 'id') {
      first.setName(SHEET_NAME);
      sh = first;
    } else {
      sh = ss.insertSheet(SHEET_NAME);
    }
  }
  if (sh.getLastRow() === 0) {
    sh.appendRow(FIELDS);
    sh.setFrozenRows(1);
  } else {
    const head = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
    if (FIELDS.some((f, i) => head[i] !== f)) sh.getRange(1, 1, 1, FIELDS.length).setValues([FIELDS]);
  }
  return sh;
}

function cellToString_(field, v) {
  if (v instanceof Date) {
    const tz = Session.getScriptTimeZone();
    if (['horaInicio', 'horaFin', 'horaLlegada', 'horaRetorno'].indexOf(field) > -1) return Utilities.formatDate(v, tz, 'HH:mm');
    if (field === 'fecha') return Utilities.formatDate(v, tz, 'yyyy-MM-dd');
    return Utilities.formatDate(v, tz, "yyyy-MM-dd'T'HH:mm:ss");
  }
  return v === null || v === undefined ? '' : String(v);
}

/** Devuelve todos los eventos. */
function getEvents() {
  const sh = getSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const header = values[0].map(String);
  const events = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const ev = {};
    header.forEach((h, i) => {
      if (FIELDS.indexOf(h) === -1) return;
      const s = cellToString_(h, row[i]);
      ev[h] = LIST_FIELDS.indexOf(h) > -1 ? (s ? s.split(LIST_SEP).map(x => x.trim()).filter(Boolean) : []) : s;
    });
    if (ev.id) events.push(ev);
  }
  return events;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return -1;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return -1;
}

/** Crea o actualiza un evento y devuelve la lista actualizada. */
function saveEvent(ev) {
  if (!ev || !ev.id) throw new Error('Falta el id del evento.');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    ev.updatedAt = new Date().toISOString();
    const row = FIELDS.map(f => {
      const v = ev[f];
      if (LIST_FIELDS.indexOf(f) > -1) return Array.isArray(v) ? v.join(LIST_SEP) : (v || '');
      return v === undefined || v === null ? '' : String(v);
    });
    let r = findRow_(sh, ev.id);
    if (r === -1) r = sh.getLastRow() + 1;
    const range = sh.getRange(r, 1, 1, FIELDS.length);
    range.setNumberFormat('@'); // texto plano: evita que Sheets cambie fechas y teléfonos
    range.setValues([row]);
  } finally {
    lock.releaseLock();
  }
  return getEvents();
}

/** Elimina un evento y devuelve la lista actualizada. */
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

/* ===================== Reporte para gerencia ===================== */

const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];

function getEvent_(id) {
  const ev = getEvents().filter(e => String(e.id) === String(id))[0];
  if (!ev) throw new Error('No se encontró el evento.');
  return ev;
}

function fechaLarga_(iso) {
  if (!iso) return 'Por definir';
  const p = iso.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2]);
  return DIAS[d.getDay()] + ' ' + p[2] + ' de ' + MESES[p[1] - 1] + ' de ' + p[0];
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
}

function lista_(texto) {
  const items = String(texto || '').split(/\n|;/).map(x => x.replace(/^[-•*\s]+/, '').trim()).filter(Boolean);
  if (!items.length) return '<p class="muted">No aplica.</p>';
  return '<ul>' + items.map(i => '<li>' + esc_(i) + '</li>').join('') + '</ul>';
}

function reporteHtml_(ev) {
  const tipo = ev.tipoEvento === 'Otro' ? (ev.tipoEventoOtro || 'Otro') : (ev.tipoEvento || 'Por definir');
  const ubic = [ev.ciudad, ev.direccion].filter(Boolean).join(' · ') + (ev.referencia ? ' (Ref.: ' + ev.referencia + ')' : '');
  const horario = ev.horaInicio ? ev.horaInicio + (ev.horaFin ? ' a ' + ev.horaFin : '') : 'Por definir';
  const din = (ev.dinamicas || []).concat(ev.dinamicaOtra ? [ev.dinamicaOtra] : []);
  const hoy = new Date();
  const fila = (k, v) => '<tr><th>' + k + '</th><td>' + esc_(v || 'Por definir') + '</td></tr>';
  return '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1E1A2B;font-size:11pt;margin:0}' +
    '.top{border-top:8px solid #A58CF4;padding:18px 0 6px}' +
    '.brand{font-size:9pt;letter-spacing:2px;color:#5B3FD1;font-weight:bold}' +
    'h1{font-size:20pt;margin:6px 0 2px}.sub{color:#6A6480;margin:0 0 14px}' +
    'h2{font-size:11pt;color:#5B3FD1;text-transform:uppercase;letter-spacing:1px;border-bottom:1px solid #E4DEF4;padding-bottom:4px;margin:18px 0 8px}' +
    'table{width:100%;border-collapse:collapse}th,td{text-align:left;padding:7px 8px;border-bottom:1px solid #EEEAF8;vertical-align:top}' +
    'th{width:34%;color:#6A6480;font-weight:normal}td{font-weight:bold}' +
    '.hl td{background:#F1ECFE}ul{margin:4px 0 0 18px;padding:0}li{margin:3px 0}.muted{color:#6A6480}' +
    '.msg{background:#F6F4FC;border-left:4px solid #A58CF4;padding:10px 12px;line-height:1.5}' +
    '.firma{margin-top:40px;display:flex}.firma div{width:45%;border-top:1px solid #1E1A2B;padding-top:4px;font-size:9pt;color:#6A6480}' +
    '.foot{margin-top:26px;font-size:8.5pt;color:#6A6480}' +
    '</style></head><body>' +
    '<div class="top"><div class="brand">DELI SNACKS · SALTY</div>' +
    '<h1>Planificación de actividad</h1>' +
    '<p class="sub">' + esc_(tipo) + ' en ' + esc_(ev.comercio || 'comercio aliado') + '</p></div>' +
    '<h2>Datos de la jornada</h2><table>' +
    fila('Día', fechaLarga_(ev.fecha)) +
    '<tr class="hl"><th>Hora de llegada</th><td>' + esc_(ev.horaLlegada || 'Por definir') + '</td></tr>' +
    fila('Horario del evento', horario) +
    '<tr class="hl"><th>Hora de retorno</th><td>' + esc_(ev.horaRetorno || 'Por definir') + '</td></tr>' +
    fila('Comercio aliado', ev.comercio) +
    fila('Ubicación', ubic) +
    fila('Responsable Deli Snacks', [ev.responsable, ev.telefono].filter(Boolean).join(' · ')) +
    '</table>' +
    '<h2>Material POP</h2>' + lista_(ev.materialPop) +
    '<h2>Dinámicas del evento</h2>' + (din.length ? '<ul>' + din.map(d => '<li>' + esc_(d) + '</li>').join('') + '</ul>' : '<p class="muted">Por definir.</p>') +
    (ev.dinamicaPersonalizada ? '<h2>Dinámica personalizada</h2><p>' + esc_(ev.dinamicaPersonalizada).replace(/\n/g, '<br>') + '</p>' : '') +
    (ev.premios ? '<h2>Premios u obsequios</h2>' + lista_(ev.premios) : '') +
    '<h2>Solicitud a la gerencia</h2><div class="msg">' + esc_(ev.mensajeGerencia || '').replace(/\n/g, '<br>') + '</div>' +
    '<div class="firma"><div>Departamento de Marketing</div><div style="margin-left:10%">Aprobado por Gerencia</div></div>' +
    '<p class="foot">Emitido el ' + hoy.getDate() + ' de ' + MESES[hoy.getMonth()] + ' de ' + hoy.getFullYear() + ' · Preparado por Mulato Ediciones para Deli Snacks</p>' +
    '</body></html>';
}

function reportePdf_(ev) {
  const nombre = 'Reporte ' + (ev.fecha || '') + ' - ' + (ev.comercio || 'evento') + '.pdf';
  return Utilities.newBlob(reporteHtml_(ev), 'text/html', 'reporte.html').getAs('application/pdf').setName(nombre);
}

function carpetaReportes_() {
  const nombre = 'Reportes Calendario Deli Snacks';
  const it = DriveApp.getFoldersByName(nombre);
  return it.hasNext() ? it.next() : DriveApp.createFolder(nombre);
}

/** Crea el PDF del reporte en Drive y devuelve su link. */
function generarReporte(id) {
  const ev = getEvent_(id);
  const file = carpetaReportes_().createFile(reportePdf_(ev));
  try { file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); } catch (e) {}
  ev.reporteUrl = file.getUrl();
  saveEvent(ev);
  return { url: file.getUrl(), nombre: file.getName() };
}

/** Envía el reporte en PDF por correo. */
function enviarReporteCorreo(id, email) {
  email = String(email || '').trim();
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email)) throw new Error('Revisa el correo de la gerente.');
  const ev = getEvent_(id);
  PropertiesService.getScriptProperties().setProperty('GERENTE_EMAIL', email);
  MailApp.sendEmail({
    to: email,
    subject: 'Planificación de actividad: ' + (ev.comercio || 'evento') + ' · ' + fechaLarga_(ev.fecha),
    htmlBody: '<p>Buen día,</p><p>' + esc_(ev.mensajeGerencia || 'Adjunto la planificación de la actividad.').replace(/\n/g, '<br>') +
      '</p><p>En el PDF adjunto encontrará el detalle de la jornada.</p><p>Saludos,<br>Departamento de Marketing · Deli Snacks</p>',
    attachments: [reportePdf_(ev)],
    name: 'Marketing Deli Snacks'
  });
  return 'ok';
}

function getConfig() {
  return { gerenteEmail: PropertiesService.getScriptProperties().getProperty('GERENTE_EMAIL') || '' };
}

/* ===================== Recordatorio por correo ===================== */

/** Envía un correo con los eventos de mañana. Lo ejecuta un activador diario. */
function enviarRecordatorios() {
  const tz = Session.getScriptTimeZone();
  const manana = Utilities.formatDate(new Date(Date.now() + 24 * 3600 * 1000), tz, 'yyyy-MM-dd');
  const evs = getEvents().filter(e => e.fecha === manana);
  if (!evs.length) return;
  const to = Session.getEffectiveUser().getEmail();
  const items = evs.map(e =>
    '<li><b>' + esc_(e.comercio || 'Evento') + '</b> · ' + esc_(e.ciudad || '') +
    '<br>Llegada: ' + esc_(e.horaLlegada || 'por definir') + ' · Evento: ' + esc_(e.horaInicio || 'por definir') + (e.horaFin ? ' a ' + esc_(e.horaFin) : '') +
    ' · Retorno: ' + esc_(e.horaRetorno || 'por definir') +
    (e.direccion ? '<br>Dirección: ' + esc_(e.direccion) : '') +
    (e.responsable ? '<br>Responsable: ' + esc_(e.responsable) + ' ' + esc_(e.telefono || '') : '') +
    (e.maps ? '<br><a href="' + esc_(e.maps) + '">Ver en Google Maps</a>' : '') + '</li>').join('');
  MailApp.sendEmail({
    to: to,
    subject: 'Recordatorio: ' + evs.length + ' evento' + (evs.length > 1 ? 's' : '') + ' mañana (' + fechaLarga_(manana) + ')',
    htmlBody: '<p>Mañana hay ' + evs.length + ' evento' + (evs.length > 1 ? 's' : '') + ' en el calendario de Deli Snacks:</p><ul>' + items + '</ul>',
    name: 'Calendario Deli Snacks'
  });
}

/** Ejecutar UNA vez desde el editor para activar el recordatorio diario (7 a.m.). */
function instalarRecordatorio() {
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'enviarRecordatorios').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('enviarRecordatorios').timeBased().everyDays(1).atHour(7).create();
  carpetaReportes_();
  return 'Recordatorio diario activado';
}

/** Página de la app (calendario + formulario). */
const APP_HTML = "<!doctype html>\n<html lang=\"es\">\n<head>\n<meta charset=\"utf-8\">\n<meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">\n<base target=\"_blank\">\n<title>Calendario de Eventos Deli Snacks</title>\n<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=Bricolage+Grotesque:opsz,wght@12..96,500;12..96,700;12..96,800&family=Figtree:wght@400;500;600;700&display=swap\">\n<style>\n/* Layout: month calendar on the left, a working panel on the right (day's events / event form); stacks on phones. */\n:root{\n  --bg:#FAF9FE; --surface:#FFFFFF; --ink:#1E1A2B; --muted:#6A6480; --line:#E4DEF4;\n  --lila:#A58CF4; --lila-ink:#5B3FD1; --lila-soft:#F1ECFE; --on-lila:#1E1A2B;\n  --warn:#B7791F; --warn-soft:#FDF3E1; --ok:#2F855A; --ok-soft:#E5F5EC; --done:#6A6480; --done-soft:#EEECF3;\n  --danger:#C53030;\n  --f-display:\"Bricolage Grotesque\", \"Trebuchet MS\", system-ui, sans-serif;\n  --f-body:\"Figtree\", system-ui, -apple-system, \"Segoe UI\", sans-serif;\n}\n@media (prefers-color-scheme: dark){ :root:not([data-theme=\"light\"]){\n  --bg:#14121B; --surface:#1D1A27; --ink:#EEEAF8; --muted:#A39DB8; --line:#2E2940;\n  --lila:#B39FF7; --lila-ink:#C4B4FA; --lila-soft:#2A2340; --on-lila:#14121B;\n  --warn:#F0B95A; --warn-soft:#3A2E17; --ok:#6FD39C; --ok-soft:#173326; --done:#A39DB8; --done-soft:#26222F;\n  --danger:#F28B8B; color-scheme:dark; } }\n:root[data-theme=\"dark\"]{\n  --bg:#14121B; --surface:#1D1A27; --ink:#EEEAF8; --muted:#A39DB8; --line:#2E2940;\n  --lila:#B39FF7; --lila-ink:#C4B4FA; --lila-soft:#2A2340; --on-lila:#14121B;\n  --warn:#F0B95A; --warn-soft:#3A2E17; --ok:#6FD39C; --ok-soft:#173326; --done:#A39DB8; --done-soft:#26222F;\n  --danger:#F28B8B; color-scheme:dark; }\n\n*{box-sizing:border-box}\nbody{margin:0;background:var(--bg);color:var(--ink);font-family:var(--f-body);font-size:15px;line-height:1.45}\n.wrap{max-width:1180px;margin:0 auto;padding-inline:20px;padding-block:22px 40px}\nheader{display:flex;flex-wrap:wrap;align-items:flex-end;justify-content:space-between;gap:12px;margin-bottom:20px}\n.brand{font-family:var(--f-display);font-weight:800;letter-spacing:.08em;font-size:13px;color:var(--lila-ink)}\n.brand span{color:var(--lila)}\nh1{font-family:var(--f-display);font-weight:700;font-size:clamp(24px,4vw,34px);line-height:1.1;margin:4px 0 0;text-wrap:balance}\n.sub{color:var(--muted);margin:4px 0 0}\n.legend{display:flex;gap:12px;flex-wrap:wrap;font-size:13px;color:var(--muted)}\n.legend i{display:inline-block;width:9px;height:9px;border-radius:50%;margin-right:5px;vertical-align:middle}\n\n.grid{display:grid;grid-template-columns:minmax(0,1.6fr) minmax(0,1fr);gap:20px;align-items:start}\n@media (max-width:860px){.grid{grid-template-columns:1fr}}\n.card{background:var(--surface);border:1px solid var(--line);border-radius:14px}\n\n/* calendar */\n.cal-head{display:flex;align-items:center;justify-content:space-between;padding:14px 16px;border-bottom:1px solid var(--line)}\n.cal-head h2{font-family:var(--f-display);font-size:20px;margin:0;text-transform:capitalize}\n.nav{display:flex;gap:6px}\nbutton{font:inherit;cursor:pointer}\n.btn{border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:9px;padding:7px 12px;font-weight:600;font-size:14px}\n.btn:hover{border-color:var(--lila)}\n.btn-primary{background:var(--lila);border-color:var(--lila);color:var(--on-lila)}\n.btn-primary:hover{filter:brightness(1.05)}\n.btn-danger{color:var(--danger)}\n.btn:focus-visible,.day:focus-visible,input:focus-visible,select:focus-visible,textarea:focus-visible,.ev:focus-visible{outline:2px solid var(--lila-ink);outline-offset:2px}\n.btn:disabled{opacity:.5;cursor:not-allowed}\n.dow,.days{display:grid;grid-template-columns:repeat(7,minmax(0,1fr))}\n.dow div{padding:8px 6px;font-size:11px;font-weight:700;letter-spacing:.08em;text-transform:uppercase;color:var(--muted);text-align:center}\n.day{min-height:92px;border:0;border-top:1px solid var(--line);border-right:1px solid var(--line);background:transparent;color:var(--ink);text-align:left;padding:6px;display:flex;flex-direction:column;gap:3px;min-width:0}\n.days .day:nth-child(7n){border-right:0}\n.day.out{color:var(--muted);opacity:.45}\n.day:hover{background:var(--lila-soft)}\n.day.sel{background:var(--lila-soft);box-shadow:inset 0 0 0 2px var(--lila)}\n.num{font-size:13px;font-weight:600;font-variant-numeric:tabular-nums}\n.day.today .num{background:var(--lila);color:var(--on-lila);border-radius:6px;padding:0 6px;align-self:flex-start}\n.chip{font-size:11px;line-height:1.3;padding:2px 6px;border-radius:5px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;max-width:100%}\n.s-pend{background:var(--warn-soft);color:var(--warn)}\n.s-conf{background:var(--ok-soft);color:var(--ok)}\n.s-done{background:var(--done-soft);color:var(--done)}\n.more{font-size:11px;color:var(--muted)}\n@media (max-width:560px){.day{min-height:58px;padding:4px}.chip{font-size:0;padding:0;height:6px;border-radius:3px}.more{display:none}}\n\n/* panel */\n.panel{padding:18px}\n.panel h3{font-family:var(--f-display);font-size:19px;margin:0 0 2px;text-transform:capitalize}\n.panel .meta{color:var(--muted);font-size:13px;margin:0 0 14px}\n.list{display:flex;flex-direction:column;gap:8px;margin-bottom:14px}\n.ev{display:block;width:100%;text-align:left;border:1px solid var(--line);background:var(--surface);color:var(--ink);border-radius:10px;padding:10px 12px}\n.ev:hover{border-color:var(--lila)}\n.ev strong{display:block;font-size:15px}\n.ev small{color:var(--muted);font-size:13px}\n.pill{display:inline-block;font-size:11px;font-weight:700;padding:2px 8px;border-radius:99px;margin-top:6px}\n.empty{border:1px dashed var(--line);border-radius:10px;padding:16px;color:var(--muted);font-size:14px;margin-bottom:14px}\n.notice{background:var(--warn-soft);color:var(--warn);border-radius:10px;padding:10px 12px;font-size:13px;margin-bottom:16px}\n\n/* form */\nform{display:flex;flex-direction:column;gap:18px}\nfieldset{border:0;padding:0;margin:0;display:flex;flex-direction:column;gap:10px;min-width:0}\nlegend{font-family:var(--f-display);font-weight:700;font-size:13px;letter-spacing:.06em;text-transform:uppercase;color:var(--lila-ink);padding:0;margin-bottom:6px}\nlabel.f{display:flex;flex-direction:column;gap:4px;font-size:13px;font-weight:600}\ninput[type=text],input[type=tel],input[type=url],input[type=date],input[type=time],textarea,select{font:inherit;font-size:14px;font-weight:400;color:var(--ink);background:var(--bg);border:1px solid var(--line);border-radius:8px;padding:8px 10px;width:100%;min-width:0}\ntextarea{min-height:64px;resize:vertical}\n.row{display:grid;grid-template-columns:1fr 1fr;gap:10px}\n@media (max-width:420px){.row{grid-template-columns:1fr}}\n.opts{display:flex;flex-wrap:wrap;gap:6px}\n.opt{display:inline-flex;align-items:center;gap:6px;border:1px solid var(--line);border-radius:8px;padding:6px 10px;font-size:13px;font-weight:500;cursor:pointer;background:var(--bg)}\n.opt:has(input:checked){border-color:var(--lila);background:var(--lila-soft)}\n.opt input{accent-color:var(--lila-ink);margin:0}\n.qlabel{font-size:13px;font-weight:600}\n.actions{display:flex;gap:8px;flex-wrap:wrap;position:sticky;bottom:0;background:var(--surface);padding-top:10px;border-top:1px solid var(--line)}\n.actions .spacer{flex:1}\n.confirm{display:flex;flex-wrap:wrap;align-items:center;gap:8px;font-size:13px;color:var(--danger)}\n.toast{position:fixed;left:50%;bottom:calc(20px + env(safe-area-inset-bottom,0px));transform:translateX(-50%);background:var(--ink);color:var(--bg);padding:10px 16px;border-radius:10px;font-weight:600;font-size:14px;opacity:0;transition:opacity .2s;pointer-events:none}\n.toast.show{opacity:1}\n@media (prefers-reduced-motion:reduce){.toast{transition:none}}\nfooter{margin-top:22px;color:var(--muted);font-size:12px}\n.tools{display:flex;flex-wrap:wrap;gap:8px;margin:0 0 16px}\na.btn{text-decoration:none;display:inline-flex;align-items:center}\n.btn-wa{background:#1F9D55;border-color:#1F9D55;color:#fff}\n.report{background:var(--lila-soft);border:1px solid var(--line);border-radius:12px;padding:14px;margin:0 0 16px;display:flex;flex-direction:column;gap:10px}\n.report h4{font-family:var(--f-display);margin:0;font-size:16px}\n.report p{margin:0;font-size:13px;color:var(--muted)}\n.report .row2{display:flex;gap:8px;flex-wrap:wrap}\n.report .row2 input{flex:1;min-width:180px}\n.hint{font-size:12px;font-weight:400;color:var(--muted)}\n</style>\n</head>\n<body>\n\n<div class=\"wrap\">\n  <header>\n    <div>\n      <div class=\"brand\">MULATO <span>/</span> EDICIONES</div>\n      <h1>Eventos y degustaciones Deli Snacks</h1>\n      <p class=\"sub\">Elige un día del calendario para ver o crear la planificación del evento. <span id=\"sync\"></span></p>\n    </div>\n    <div class=\"legend\" aria-label=\"Estados\">\n      <span><i style=\"background:var(--warn)\"></i>Por confirmar</span>\n      <span><i style=\"background:var(--ok)\"></i>Confirmado</span>\n      <span><i style=\"background:var(--done)\"></i>Realizado</span>\n    </div>\n  </header>\n\n  <div class=\"grid\">\n    <section class=\"card\" aria-label=\"Calendario\">\n      <div class=\"cal-head\">\n        <h2 id=\"monthLabel\"></h2>\n        <div class=\"nav\">\n          <button class=\"btn\" id=\"prev\" aria-label=\"Mes anterior\">‹</button>\n          <button class=\"btn\" id=\"todayBtn\">Hoy</button>\n          <button class=\"btn\" id=\"next\" aria-label=\"Mes siguiente\">›</button>\n        </div>\n      </div>\n      <div class=\"dow\"><div>Lun</div><div>Mar</div><div>Mié</div><div>Jue</div><div>Vie</div><div>Sáb</div><div>Dom</div></div>\n      <div class=\"days\" id=\"days\"></div>\n    </section>\n\n    <aside class=\"card panel\" id=\"panel\" aria-live=\"polite\"></aside>\n  </div>\n  <footer>Planificación preparada por Mulato Ediciones para el equipo de Deli Snacks.</footer>\n</div>\n<div class=\"toast\" id=\"toast\" role=\"status\"></div>\n\n<script>\nconst MONTHS=[\"enero\",\"febrero\",\"marzo\",\"abril\",\"mayo\",\"junio\",\"julio\",\"agosto\",\"septiembre\",\"octubre\",\"noviembre\",\"diciembre\"];\nconst DAYS=[\"domingo\",\"lunes\",\"martes\",\"miércoles\",\"jueves\",\"viernes\",\"sábado\"];\nconst STATUS={pend:{t:\"Por confirmar\",c:\"s-pend\"},conf:{t:\"Confirmado\",c:\"s-conf\"},done:{t:\"Realizado\",c:\"s-done\"}};\nconst TIPO_EVENTO=[\"Degustación\",\"Aniversario del comercio\",\"Inauguración\",\"Activación en punto de venta\",\"Otro\"];\nconst DINAMICAS=[\"Degustación de productos\",\"Ruleta de premios\",\"Síguenos en Instagram y gana\",\"Votación del sabor favorito\",\"Adivina el sabor\",\"Compra mínima de $3: recibe un snack (minorista)\",\"Compra mínima de un bulto: recibe un obsequio (mayorista)\",\"Interacción con la mascota de la marca\",\"Fotografías con clientes\"];\nconst DEFAULT_MSG=\"Por medio de la presente, solicitamos su aprobación para la realización de esta actividad, así como la asignación del material POP, los productos para degustación y los premios indicados, y la autorización del traslado del equipo en el horario señalado.\\n\\nQuedamos atentos a sus comentarios.\";\nlet report={id:null,state:\"idle\",url:\"\",email:\"\"};\nconst CONTENIDO=[\"Publicación en colaboración con el comercio\",\"Stories en vivo durante el evento\",\"Videos de las dinámicas\",\"Reacciones y entrevistas al público\",\"Fotografías para redes\",\"Video resumen del evento\"];\n\nconst pad=n=>String(n).padStart(2,\"0\");\nconst iso=d=>d.getFullYear()+\"-\"+pad(d.getMonth()+1)+\"-\"+pad(d.getDate());\nconst esc=s=>String(s??\"\").replace(/[&<>\"']/g,c=>({\"&\":\"&amp;\",\"<\":\"&lt;\",\">\":\"&gt;\",'\"':\"&quot;\",\"'\":\"&#39;\"}[c]));\nconst today=iso(new Date());\n\nlet view=new Date(); view.setDate(1);\nlet selected=today;\nlet events={};\nlet mode={type:\"day\"};\nconst canWrite=true; let dbState=\"loading\", busy=false;\nlet confirmDelete=false;\n\nconst $=id=>document.getElementById(id);\n\nfunction toast(msg){const t=$(\"toast\");t.textContent=msg;t.classList.add(\"show\");clearTimeout(toast._t);toast._t=setTimeout(()=>t.classList.remove(\"show\"),2200);}\n\nfunction byDate(date){return Object.entries(events).filter(([,e])=>e.fecha===date).sort((a,b)=>(a[1].horaInicio||\"\").localeCompare(b[1].horaInicio||\"\"));}\nfunction title(e){return e.comercio||\"Evento sin nombre\";}\n\nfunction renderCalendar(){\n  $(\"monthLabel\").textContent=MONTHS[view.getMonth()]+\" \"+view.getFullYear();\n  const first=new Date(view.getFullYear(),view.getMonth(),1);\n  const offset=(first.getDay()+6)%7;\n  const start=new Date(first); start.setDate(1-offset);\n  let html=\"\";\n  for(let i=0;i<42;i++){\n    const d=new Date(start); d.setDate(start.getDate()+i);\n    const k=iso(d), evs=byDate(k);\n    const cls=[\"day\"]; if(d.getMonth()!==view.getMonth())cls.push(\"out\"); if(k===today)cls.push(\"today\"); if(k===selected)cls.push(\"sel\");\n    const label=d.getDate()+\" de \"+MONTHS[d.getMonth()]+(evs.length?\", \"+evs.length+\" evento\"+(evs.length>1?\"s\":\"\"):\"\");\n    html+=`<button class=\"${cls.join(\" \")}\" data-date=\"${k}\" aria-label=\"${label}\"><span class=\"num\">${d.getDate()}</span>`+\n      evs.slice(0,2).map(([,e])=>`<span class=\"chip ${STATUS[e.estado||\"pend\"].c}\">${esc(title(e))}</span>`).join(\"\")+\n      (evs.length>2?`<span class=\"more\">+${evs.length-2} más</span>`:\"\")+`</button>`;\n    if(i===34){const n=new Date(start);n.setDate(start.getDate()+35);if(n.getMonth()!==view.getMonth())break;}\n  }\n  $(\"days\").innerHTML=html;\n}\n\nfunction prettyDate(k){const [y,m,d]=k.split(\"-\").map(Number);const dt=new Date(y,m-1,d);return DAYS[dt.getDay()]+\" \"+d+\" de \"+MONTHS[m-1];}\n\nfunction noticeHtml(){\n  if(dbState===\"off\") return `<div class=\"notice\">No se pudo conectar con la hoja de eventos. Revisa tu conexión; la página lo vuelve a intentar sola.</div>`;\n  return \"\";\n}\n\nfunction renderDay(){\n  const evs=byDate(selected);\n  let html=noticeHtml()+`<h3>${prettyDate(selected)}</h3><p class=\"meta\">${evs.length?evs.length+\" evento\"+(evs.length>1?\"s\":\"\")+\" este día\":\"Sin eventos este día\"}</p>`;\n  if(dbState===\"loading\") html+=`<div class=\"empty\">Cargando eventos…</div>`;\n  else if(evs.length){\n    html+=`<div class=\"list\">`+evs.map(([id,e])=>{\n      const st=STATUS[e.estado||\"pend\"];\n      const hora=e.horaInicio?(e.horaInicio+(e.horaFin?\" a \"+e.horaFin:\"\")):\"Hora por definir\";\n      const lugar=[e.tipoEvento===\"Otro\"?(e.tipoEventoOtro||\"Otro\"):e.tipoEvento,e.ciudad].filter(Boolean).join(\" · \");\n      return `<button class=\"ev\" data-id=\"${id}\"><strong>${esc(title(e))}</strong><small>${esc(hora)}${lugar?\" · \"+esc(lugar):\"\"}</small><br><span class=\"pill ${st.c}\">${st.t}</span></button>`;\n    }).join(\"\")+`</div>`;\n  } else html+=`<div class=\"empty\">Aún no hay eventos para este día. Crea uno para llenar su planificación: comercio, ubicación, horario, dinámicas y contenido.</div>`;\n  html+=`<button class=\"btn btn-primary\" id=\"newEv\" ${canWrite?\"\":\"disabled\"}>Nuevo evento</button>`;\n  $(\"panel\").innerHTML=html;\n}\n\nfunction opts(name,list,type,values){\n  return `<div class=\"opts\">`+list.map((o,i)=>{const id=name+\"_\"+i;const on=type===\"radio\"?values===o:(values||[]).includes(o);\n    return `<label class=\"opt\" for=\"${id}\"><input type=\"${type}\" id=\"${id}\" name=\"${name}\" value=\"${esc(o)}\" ${on?\"checked\":\"\"}>${esc(o)}</label>`;}).join(\"\")+`</div>`;\n}\n\nfunction renderForm(){\n  const isNew=!mode.id; const e=isNew?{fecha:selected,estado:\"pend\",mascota:\"No\",mensajeGerencia:DEFAULT_MSG}:events[mode.id]||{};\n  const ro=canWrite?\"\":\"disabled\";\n  $(\"panel\").innerHTML=noticeHtml()+`\n  <h3>${isNew?\"Nuevo evento\":esc(title(e))}</h3><p class=\"meta\">${prettyDate(e.fecha||selected)}</p>\n  ${isNew?\"\":toolsHtml(mode.id,e)}\n  <form id=\"evForm\" novalidate><fieldset ${ro} style=\"display:contents\">\n    <fieldset><legend>Estado</legend>${opts(\"estado\",Object.keys(STATUS).map(k=>STATUS[k].t),\"radio\",STATUS[e.estado||\"pend\"].t)}</fieldset>\n    <fieldset><legend>1. Información del cliente</legend>\n      <label class=\"f\" for=\"comercio\">Nombre del comercio<input type=\"text\" id=\"comercio\" value=\"${esc(e.comercio)}\" required></label>\n      <label class=\"f\" for=\"instagram\">Instagram del comercio<input type=\"text\" id=\"instagram\" value=\"${esc(e.instagram)}\" placeholder=\"@usuario\"></label>\n      <div><div class=\"qlabel\">Tipo de cliente</div>${opts(\"tipoCliente\",[\"Mayorista\",\"Minorista\",\"Mixto\"],\"radio\",e.tipoCliente)}</div>\n      <div class=\"row\">\n        <label class=\"f\" for=\"responsable\">Responsable de Deli Snacks<input type=\"text\" id=\"responsable\" value=\"${esc(e.responsable)}\"></label>\n        <label class=\"f\" for=\"telefono\">Número de contacto<input type=\"tel\" id=\"telefono\" value=\"${esc(e.telefono)}\"></label>\n      </div>\n    </fieldset>\n    <fieldset><legend>2. Ubicación</legend>\n      <label class=\"f\" for=\"ciudad\">Ciudad<input type=\"text\" id=\"ciudad\" value=\"${esc(e.ciudad)}\"></label>\n      <label class=\"f\" for=\"direccion\">Dirección exacta<input type=\"text\" id=\"direccion\" value=\"${esc(e.direccion)}\"></label>\n      <label class=\"f\" for=\"referencia\">Punto de referencia<input type=\"text\" id=\"referencia\" value=\"${esc(e.referencia)}\"></label>\n      <label class=\"f\" for=\"maps\">Enlace de Google Maps<input type=\"url\" id=\"maps\" value=\"${esc(e.maps)}\" placeholder=\"Pega aquí el enlace de Google Maps\"></label>\n      ${e.maps&&e.maps.indexOf(\"http\")===0?`<a href=\"${esc(e.maps)}\" target=\"_blank\" rel=\"noopener\" style=\"color:var(--lila-ink);font-size:13px;font-weight:600\">Abrir ubicación en Google Maps ↗</a>`:\"\"}\n    </fieldset>\n    <fieldset><legend>3. Evento o degustación</legend>\n      <div><div class=\"qlabel\">Tipo de evento</div>${opts(\"tipoEvento\",TIPO_EVENTO,\"radio\",e.tipoEvento)}</div>\n      <label class=\"f\" for=\"tipoEventoOtro\">Si es otro, ¿cuál?<input type=\"text\" id=\"tipoEventoOtro\" value=\"${esc(e.tipoEventoOtro)}\"></label>\n      <label class=\"f\" for=\"descripcion\">Descripción del evento<textarea id=\"descripcion\">${esc(e.descripcion)}</textarea></label>\n      <label class=\"f\" for=\"fecha\">Fecha<input type=\"date\" id=\"fecha\" value=\"${esc(e.fecha||selected)}\"></label>\n      <div class=\"row\">\n        <label class=\"f\" for=\"horaInicio\">Hora de inicio<input type=\"time\" id=\"horaInicio\" value=\"${esc(e.horaInicio)}\"></label>\n        <label class=\"f\" for=\"horaFin\">Hora de cierre<input type=\"time\" id=\"horaFin\" value=\"${esc(e.horaFin)}\"></label>\n      </div>\n      <div><div class=\"qlabel\">¿Participará la mascota de la marca?</div>${opts(\"mascota\",[\"Sí\",\"No\"],\"radio\",e.mascota)}</div>\n      <label class=\"f\" for=\"premios\">Productos para premios u obsequios<input type=\"text\" id=\"premios\" value=\"${esc(e.premios)}\"></label>\n    </fieldset>\n    <fieldset><legend>4. Dinámicas y actividades</legend>\n      ${opts(\"dinamicas\",DINAMICAS,\"checkbox\",e.dinamicas)}\n      <label class=\"f\" for=\"dinamicaOtra\">Otra dinámica<input type=\"text\" id=\"dinamicaOtra\" value=\"${esc(e.dinamicaOtra)}\"></label>\n    </fieldset>\n    <fieldset><legend>5. Contenido y cobertura</legend>\n      ${opts(\"contenido\",CONTENIDO,\"checkbox\",e.contenido)}\n      <label class=\"f\" for=\"notas\">Observaciones<textarea id=\"notas\">${esc(e.notas)}</textarea></label>\n    </fieldset>\n    <fieldset><legend>6. Reporte para gerencia</legend>\n      <div class=\"row\">\n        <label class=\"f\" for=\"horaLlegada\">Hora de llegada<input type=\"time\" id=\"horaLlegada\" value=\"${esc(e.horaLlegada)}\"></label>\n        <label class=\"f\" for=\"horaRetorno\">Hora de retorno<input type=\"time\" id=\"horaRetorno\" value=\"${esc(e.horaRetorno)}\"></label>\n      </div>\n      <label class=\"f\" for=\"materialPop\">Material POP <span class=\"hint\">Uno por línea: pendón, exhibidor, mantel, volantes…</span><textarea id=\"materialPop\">${esc(e.materialPop)}</textarea></label>\n      <label class=\"f\" for=\"dinamicaPersonalizada\">Dinámica personalizada <span class=\"hint\">Si este comercio necesita algo especial</span><textarea id=\"dinamicaPersonalizada\">${esc(e.dinamicaPersonalizada)}</textarea></label>\n      <label class=\"f\" for=\"mensajeGerencia\">Mensaje de solicitud para la gerencia<textarea id=\"mensajeGerencia\" style=\"min-height:110px\">${esc(e.mensajeGerencia)}</textarea></label>\n    </fieldset>\n  </fieldset>\n    <div class=\"actions\" id=\"actions\">\n      ${canWrite?`<button type=\"submit\" class=\"btn btn-primary\">${isNew?\"Guardar evento\":\"Guardar cambios\"}</button>`:\"\"}\n      <button type=\"button\" class=\"btn\" id=\"back\">Volver</button>\n      <span class=\"spacer\"></span>\n      ${!isNew&&canWrite?(confirmDelete?`<span class=\"confirm\">¿Eliminar este evento? <button type=\"button\" class=\"btn btn-danger\" id=\"delYes\">Eliminar</button><button type=\"button\" class=\"btn\" id=\"delNo\">No</button></span>`:`<button type=\"button\" class=\"btn btn-danger\" id=\"del\">Eliminar</button>`):\"\"}\n    </div>\n  </form>`;\n}\n\n\nconst WA=\"https:\"+\"/\"+\"/wa.me/?text=\";\nfunction waText(e){\n  const L=[];\n  L.push(\"*Evento Deli Snacks*\");\n  L.push(\"📅 \"+prettyDate(e.fecha||selected));\n  L.push(\"🏪 \"+(e.comercio||\"Por definir\"));\n  const ub=[e.ciudad,e.direccion].filter(Boolean).join(\" - \");\n  if(ub) L.push(\"📍 \"+ub+(e.referencia?\" (Ref.: \"+e.referencia+\")\":\"\"));\n  L.push(\"🕐 Llegada: \"+(e.horaLlegada||\"por definir\")+\" · Evento: \"+(e.horaInicio||\"por definir\")+(e.horaFin?\" a \"+e.horaFin:\"\")+\" · Retorno: \"+(e.horaRetorno||\"por definir\"));\n  if(e.responsable) L.push(\"👤 Responsable: \"+e.responsable+(e.telefono?\" (\"+e.telefono+\")\":\"\"));\n  if(e.maps) L.push(\"🗺️ \"+e.maps);\n  return L.join(\"\\n\");\n}\nfunction toolsHtml(id,e){\n  let h=`<div class=\"tools\"><a class=\"btn btn-wa\" href=\"${WA+encodeURIComponent(waText(e))}\" target=\"_blank\" rel=\"noopener\">Enviar por WhatsApp</a><button type=\"button\" class=\"btn btn-primary\" id=\"repBtn\">Reporte para gerencia</button></div>`;\n  if(report.id===id && report.state!==\"idle\"){\n    if(report.state===\"loading\") h+=`<div class=\"report\"><h4>Generando el reporte…</h4><p>Esto tarda unos segundos.</p></div>`;\n    else if(report.state===\"error\") h+=`<div class=\"report\"><h4>No se pudo generar el reporte</h4><p>${esc(report.msg||\"Revisa tu conexión e intenta de nuevo.\")}</p></div>`;\n    else{\n      const waG=WA+encodeURIComponent(\"Buen día. Le comparto la planificación de la actividad en \"+(e.comercio||\"el comercio aliado\")+\" para el \"+prettyDate(e.fecha||selected)+\": \"+report.url);\n      h+=`<div class=\"report\"><h4>Reporte listo</h4><p>Se guardó en tu Drive, en la carpeta \"Reportes Calendario Deli Snacks\".</p>\n      <div class=\"row2\"><a class=\"btn\" href=\"${esc(report.url)}\" target=\"_blank\" rel=\"noopener\">Abrir PDF</a><a class=\"btn btn-wa\" href=\"${waG}\" target=\"_blank\" rel=\"noopener\">Enviar a la gerente por WhatsApp</a></div>\n      <label class=\"f\" for=\"gerEmail\">Enviar por correo a la gerente</label>\n      <div class=\"row2\"><input type=\"text\" id=\"gerEmail\" placeholder=\"correo@empresa.com\" value=\"${esc(report.email)}\"><button type=\"button\" class=\"btn btn-primary\" id=\"mailBtn\" ${report.sending?\"disabled\":\"\"}>${report.sending?\"Enviando…\":\"Enviar correo\"}</button></div></div>`;\n    }\n  }\n  return h;\n}\nfunction callFn(fn,...a){return new Promise((res,rej)=>google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...a));}\nasync function makeReport(){\n  const id=mode.id; const data=collect();\n  if(!data.comercio){toast(\"Escribe el nombre del comercio\");return;}\n  report={id,state:\"loading\",url:\"\",email:report.email};\n  events[id]=data; renderForm();\n  try{\n    await callFn(\"saveEvent\",Object.assign({id},data));\n    const r=await callFn(\"generarReporte\",id);\n    report.state=\"ready\"; report.url=r.url;\n    if(!report.email){try{const c=await callFn(\"getConfig\");report.email=c.gerenteEmail||\"\";}catch(e){}}\n  }catch(err){report.state=\"error\";report.msg=err&&err.message;}\n  if(mode.id===id) renderForm();\n}\nasync function mailReport(){\n  const em=$(\"gerEmail\").value.trim(); report.email=em;\n  if(!/^[^@\\s]+@[^@\\s]+\\.[^@\\s]+$/.test(em)){toast(\"Escribe un correo válido\");$(\"gerEmail\").focus();return;}\n  report.sending=true; renderForm();\n  try{await callFn(\"enviarReporteCorreo\",report.id,em);toast(\"Reporte enviado a \"+em);}\n  catch(err){toast((err&&err.message)||\"No se pudo enviar el correo\");}\n  report.sending=false; if(mode.id===report.id) renderForm();\n}\n\nfunction render(){renderCalendar(); mode.type===\"form\"?renderForm():renderDay();}\n\nfunction collect(){\n  const v=id=>$(id).value.trim();\n  const radio=n=>{const x=document.querySelector(`input[name=\"${n}\"]:checked`);return x?x.value:\"\";};\n  const checks=n=>[...document.querySelectorAll(`input[name=\"${n}\"]:checked`)].map(x=>x.value);\n  const estT=radio(\"estado\"); const estado=Object.keys(STATUS).find(k=>STATUS[k].t===estT)||\"pend\";\n  return {estado,comercio:v(\"comercio\"),instagram:v(\"instagram\"),tipoCliente:radio(\"tipoCliente\"),responsable:v(\"responsable\"),telefono:v(\"telefono\"),\n    ciudad:v(\"ciudad\"),direccion:v(\"direccion\"),referencia:v(\"referencia\"),maps:v(\"maps\"),\n    tipoEvento:radio(\"tipoEvento\"),tipoEventoOtro:v(\"tipoEventoOtro\"),descripcion:v(\"descripcion\"),fecha:v(\"fecha\")||selected,\n    horaInicio:v(\"horaInicio\"),horaFin:v(\"horaFin\"),mascota:radio(\"mascota\"),premios:v(\"premios\"),\n    dinamicas:checks(\"dinamicas\"),dinamicaOtra:v(\"dinamicaOtra\"),contenido:checks(\"contenido\"),notas:v(\"notas\"),horaLlegada:v(\"horaLlegada\"),horaRetorno:v(\"horaRetorno\"),materialPop:v(\"materialPop\"),dinamicaPersonalizada:v(\"dinamicaPersonalizada\"),mensajeGerencia:v(\"mensajeGerencia\"),reporteUrl:(events[mode.id]||{}).reporteUrl||\"\",updatedAt:new Date().toISOString()};\n}\n\nfunction call(fn,...args){\n  return new Promise((res,rej)=>{\n    if(!(window.google&&google.script&&google.script.run)){rej(new Error(\"sin conexión\"));return;}\n    google.script.run.withSuccessHandler(res).withFailureHandler(rej)[fn](...args);\n  });\n}\n\nfunction apply(list){\n  const next={}; (list||[]).forEach(e=>{if(e&&e.id){const id=e.id; const c=Object.assign({},e); delete c.id; next[id]=c;}});\n  events=next; dbState=\"ready\";\n  const s=$(\"sync\"); if(s){const d=new Date(); s.textContent=\"Actualizado \"+pad(d.getHours())+\":\"+pad(d.getMinutes());}\n  if(mode.type===\"form\") renderCalendar(); else render();\n}\n\nasync function load(){\n  if(busy) return;\n  try{ apply(await call(\"getEvents\")); }\n  catch(e){ if(dbState!==\"ready\"){dbState=\"off\"; if(mode.type!==\"form\") render();} }\n}\n\nasync function save(){\n  const data=collect();\n  if(!data.comercio){$(\"comercio\").focus();toast(\"Escribe el nombre del comercio\");return;}\n  const id=mode.id||(\"ev\"+Date.now().toString(36)+Math.random().toString(36).slice(2,6));\n  busy=true; toast(\"Guardando…\");\n  try{\n    const list=await call(\"saveEvent\",Object.assign({id},data));\n    selected=data.fecha; const [y,m]=data.fecha.split(\"-\").map(Number); view=new Date(y,m-1,1);\n    mode={type:\"day\"}; apply(list); toast(\"Evento guardado\");\n  }catch(e){ toast(\"No se pudo guardar. Revisa tu conexión e intenta de nuevo.\"); }\n  finally{ busy=false; }\n}\n\nasync function remove(){\n  const id=mode.id; busy=true;\n  try{\n    const list=await call(\"deleteEvent\",id);\n    confirmDelete=false; mode={type:\"day\"}; apply(list); toast(\"Evento eliminado\");\n  }catch(e){ toast(\"No se pudo eliminar. Intenta de nuevo.\"); }\n  finally{ busy=false; }\n}\n\ndocument.addEventListener(\"click\",ev=>{\n  const t=ev.target.closest(\"button\"); if(!t) return;\n  if(t.id===\"prev\"){view=new Date(view.getFullYear(),view.getMonth()-1,1);renderCalendar();}\n  else if(t.id===\"next\"){view=new Date(view.getFullYear(),view.getMonth()+1,1);renderCalendar();}\n  else if(t.id===\"todayBtn\"){view=new Date();view.setDate(1);selected=today;mode={type:\"day\"};render();}\n  else if(t.dataset.date){selected=t.dataset.date;const [y,m]=selected.split(\"-\").map(Number);if(m-1!==view.getMonth())view=new Date(y,m-1,1);mode={type:\"day\"};confirmDelete=false;render();}\n  else if(t.id===\"newEv\"){mode={type:\"form\"};confirmDelete=false;render();$(\"comercio\").focus();}\n  else if(t.dataset.id){mode={type:\"form\",id:t.dataset.id};confirmDelete=false;if(report.id!==t.dataset.id)report={id:null,state:\"idle\",url:\"\",email:report.email};render();}\n  else if(t.id===\"back\"){mode={type:\"day\"};confirmDelete=false;render();}\n  else if(t.id===\"del\"){confirmDelete=true;renderForm();}\n  else if(t.id===\"delNo\"){confirmDelete=false;renderForm();}\n  else if(t.id===\"delYes\"){remove();}\n  else if(t.id===\"repBtn\"){makeReport();}\n  else if(t.id===\"mailBtn\"){mailReport();}\n});\ndocument.addEventListener(\"submit\",ev=>{ev.preventDefault();if(canWrite)save();});\n\nrender();\nload();\nsetInterval(()=>{ if(!document.hidden) load(); }, 10000);\ndocument.addEventListener(\"visibilitychange\",()=>{ if(!document.hidden) load(); });\n</script>\n\n\n</body>\n</html>\n";
