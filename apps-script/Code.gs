/**
 * Anotador diario — Backend (Google Apps Script vinculado a la planilla)
 * Versión 1.0 — 2026-09-30 12:00 ARG
 * Autor: Germán Rodríguez
 *
 * Instalación (una sola vez):
 *   1. Pegar este archivo y appsscript.json en el proyecto de Apps Script de la planilla.
 *   2. Ejecutar setup() y autorizar.
 *   3. Implementar > Nueva implementación > Aplicación web
 *      Ejecutar como: Yo — Quién tiene acceso: Cualquier usuario.
 *   4. Copiar la URL /exec en index.html (API_DEFAULT) y la clave (menú Anotador > Ver clave de acceso).
 */

const TZ = 'America/Argentina/Buenos_Aires';
const SH = { T: 'Tareas', H: 'Hojas', C: 'Config', E: 'Envios' };
const COLS_T = ['id', 'hojaId', 'texto', 'hecha', 'fecha', 'persona', 'nota', 'prioridad', 'creada', 'completada'];
const COLS_H = ['id', 'nombre', 'color', 'orden'];
const COLS_E = ['fecha', 'para', 'cc', 'hojas', 'pendientes', 'origen'];
const CONFIG_DEFAULTS = {
  destinatario: 'amelife@amelife-911.com',
  ccSugeridos: 'ezzy@ame-life.net',
  personas: 'Ezzy,Maia,Maca,Lucia',
  autoEnvio: 'no',
  horaAutoEnvio: '19'
};
const GRUPOS = [
  ['venc', 'Vencidas'], ['hoy', 'Hoy'], ['man', 'Mañana'], ['semana', 'Esta semana'],
  ['prox', 'Semana entrante'], ['mas', 'Más adelante'], ['sin', 'Sin fecha']
];
const DIAS = ['domingo', 'lunes', 'martes', 'miércoles', 'jueves', 'viernes', 'sábado'];
const MESES = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];

/* ================= Setup y menú ================= */

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Ejecutá setup() desde el proyecto de Apps Script vinculado a la planilla.');
  PropertiesService.getScriptProperties().setProperty('SS_ID', ss.getId());

  ensureSheet_(ss, SH.T, COLS_T);
  ensureSheet_(ss, SH.H, COLS_H);
  ensureSheet_(ss, SH.C, ['clave', 'valor']);
  ensureSheet_(ss, SH.E, COLS_E);

  if (getHojas_().length === 0) {
    upsertMany_(SH.H, COLS_H, [{ id: 'h_general', nombre: 'General', color: '#3E5C76', orden: 1 }]);
  }
  const cfg = getConfig_();
  writeConfig_(cfg);
  instalarTrigger_(cfg);

  const def = ss.getSheetByName('Hoja 1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1 && def.getLastRow() === 0) ss.deleteSheet(def);

  const key = getOrCreateKey_();
  Logger.log('Clave de acceso: ' + key);
  return key;
}

function onOpen() {
  SpreadsheetApp.getUi().createMenu('Anotador')
    .addItem('Ver clave de acceso', 'verClave')
    .addItem('Enviar resumen ahora (todas las hojas)', 'enviarAhora')
    .addToUi();
}

function verClave() {
  SpreadsheetApp.getUi().alert('Clave de acceso del anotador:\n\n' + getOrCreateKey_());
}

function enviarAhora() {
  const cfg = getConfig_();
  sendReport_({ hojaIds: getHojas_().map(h => h.id), to: cfg.destinatario, cc: [], incluirHechas: true, origen: 'menú' });
  SpreadsheetApp.getUi().alert('Resumen enviado a ' + cfg.destinatario);
}

function getOrCreateKey_() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('API_KEY');
  if (!k) {
    k = 'an-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    p.setProperty('API_KEY', k);
  }
  return k;
}

/* ================= API web ================= */

function doGet() {
  return json_({ ok: true, app: 'Anotador diario', version: '1.0' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Solicitud inválida' });
  }
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key || req.key !== key) return json_({ ok: false, auth: true, error: 'Clave de acceso incorrecta' });

  const fn = ACTIONS[req.action];
  if (!fn) return json_({ ok: false, error: 'Acción desconocida: ' + req.action });

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(25000);
    return json_({ ok: true, data: fn(req.payload || {}) });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

const ACTIONS = {
  bootstrap: () => ({ hojas: getHojas_(), tareas: getTareas_(), config: getConfig_(), hoy: today_() }),
  upsertTask: p => { const t = cleanTask_(p); upsertMany_(SH.T, COLS_T, [t]); return t; },
  upsertTasks: p => { const ts = (p.tareas || []).map(cleanTask_); upsertMany_(SH.T, COLS_T, ts); return ts.length; },
  deleteTask: p => { deleteMany_(SH.T, [p.id]); return true; },
  deleteTasks: p => { deleteMany_(SH.T, p.ids || []); return true; },
  upsertHoja: p => { const h = cleanHoja_(p); upsertMany_(SH.H, COLS_H, [h]); return h; },
  deleteHoja: p => deleteHoja_(p.id),
  saveConfig: p => saveConfig_(p),
  sendReport: p => sendReport_(Object.assign({}, p, { origen: 'app' }))
};

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Datos ================= */

function ss_() {
  const id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Falta la pestaña "' + name + '". Ejecutá setup().');
  return sh;
}

function ensureSheet_(ss, name, headers) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  sh.getRange(1, 1, sh.getMaxRows(), headers.length).setNumberFormat('@');
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#F5F7F9').setFontColor('#1C2630');
  sh.setFrozenRows(1);
  return sh;
}

function toCell_(v) {
  if (v === true) return 'true';
  if (v === false) return 'false';
  if (v === null || v === undefined) return '';
  return String(v);
}

function fromCell_(v) {
  if (v instanceof Date) {
    const s = Utilities.formatDate(v, TZ, "yyyy-MM-dd'T'HH:mm");
    return s.endsWith('T00:00') ? s.slice(0, 10) : s;
  }
  return v === null || v === undefined ? '' : String(v);
}

function readTable_(name, cols) {
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, cols.length).getValues()
    .filter(r => String(r[0]).trim() !== '')
    .map(r => {
      const o = {};
      cols.forEach((c, i) => { o[c] = fromCell_(r[i]); });
      return o;
    });
}

function upsertMany_(name, cols, objs) {
  if (!objs.length) return;
  const sh = sheet_(name);
  const last = sh.getLastRow();
  const ids = last > 1 ? sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0])) : [];
  let next = Math.max(last, 1) + 1;
  objs.forEach(o => {
    const row = cols.map(c => toCell_(o[c]));
    const i = ids.indexOf(String(o.id));
    const r = i >= 0 ? i + 2 : next++;
    sh.getRange(r, 1, 1, cols.length).setNumberFormat('@').setValues([row]);
    if (i < 0) ids.push(String(o.id));
  });
}

function deleteMany_(name, idList) {
  const want = new Set((idList || []).map(String));
  if (!want.size) return;
  const sh = sheet_(name);
  const last = sh.getLastRow();
  if (last < 2) return;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues().map(r => String(r[0]));
  const rows = [];
  ids.forEach((id, i) => { if (want.has(id)) rows.push(i + 2); });
  rows.sort((a, b) => b - a).forEach(r => sh.deleteRow(r));
}

function isTrue_(v) { return v === true || String(v).toLowerCase() === 'true'; }

function getTareas_() {
  return readTable_(SH.T, COLS_T).map(t => Object.assign(t, { hecha: isTrue_(t.hecha), prioridad: isTrue_(t.prioridad) }));
}

function getHojas_() {
  return readTable_(SH.H, COLS_H)
    .map(h => Object.assign(h, { orden: Number(h.orden) || 0 }))
    .sort((a, b) => a.orden - b.orden);
}

function cleanTask_(p) {
  if (!p || !p.id) throw new Error('Tarea sin id');
  const fecha = /^\d{4}-\d{2}-\d{2}$/.test(p.fecha || '') ? p.fecha : '';
  return {
    id: String(p.id).slice(0, 40),
    hojaId: String(p.hojaId || '').slice(0, 40),
    texto: String(p.texto || '').slice(0, 2000),
    hecha: !!p.hecha,
    fecha: fecha,
    persona: String(p.persona || '').slice(0, 40),
    nota: String(p.nota || '').slice(0, 4000),
    prioridad: !!p.prioridad,
    creada: String(p.creada || nowIso_()),
    completada: p.hecha ? String(p.completada || nowIso_()) : ''
  };
}

function cleanHoja_(p) {
  if (!p || !p.id) throw new Error('Hoja sin id');
  const color = /^#[0-9a-fA-F]{6}$/.test(p.color || '') ? p.color : '#3E5C76';
  return { id: String(p.id).slice(0, 40), nombre: String(p.nombre || 'Sin nombre').slice(0, 60), color: color, orden: Number(p.orden) || 0 };
}

function deleteHoja_(id) {
  if (getHojas_().length <= 1) throw new Error('Tiene que quedar al menos una hoja.');
  const tareas = getTareas_().filter(t => t.hojaId === id).map(t => t.id);
  deleteMany_(SH.T, tareas);
  deleteMany_(SH.H, [id]);
  return true;
}

/* ================= Configuración ================= */

function getConfig_() {
  const out = Object.assign({}, CONFIG_DEFAULTS);
  const sh = ss_().getSheetByName(SH.C);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues().forEach(r => {
      const k = String(r[0]).trim();
      if (k && Object.prototype.hasOwnProperty.call(CONFIG_DEFAULTS, k)) out[k] = String(r[1]);
    });
  }
  return out;
}

function writeConfig_(cfg) {
  const sh = sheet_(SH.C);
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 2).clearContent();
  const rows = Object.keys(CONFIG_DEFAULTS).map(k => [k, String(cfg[k])]);
  sh.getRange(2, 1, rows.length, 2).setNumberFormat('@').setValues(rows);
}

function saveConfig_(p) {
  const cur = getConfig_();
  let trig = false;
  Object.keys(CONFIG_DEFAULTS).forEach(k => {
    if (p[k] === undefined) return;
    const v = String(p[k]).slice(0, 2000);
    if ((k === 'autoEnvio' || k === 'horaAutoEnvio') && v !== cur[k]) trig = true;
    cur[k] = v;
  });
  writeConfig_(cur);
  if (trig) instalarTrigger_(cur);
  return cur;
}

function instalarTrigger_(cfg) {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'envioAutomatico')
    .forEach(t => ScriptApp.deleteTrigger(t));
  if (cfg.autoEnvio !== 'si') return;
  const h = Math.min(23, Math.max(0, parseInt(cfg.horaAutoEnvio, 10) || 19));
  ScriptApp.newTrigger('envioAutomatico').timeBased().everyDays(1).atHour(h).inTimezone(TZ).create();
}

/** Disparador diario: envía todas las hojas al destinatario, de lunes a viernes. */
function envioAutomatico() {
  const cfg = getConfig_();
  if (cfg.autoEnvio !== 'si') return;
  const dow = Number(Utilities.formatDate(new Date(), TZ, 'u'));
  if (dow > 5) return;
  sendReport_({ hojaIds: getHojas_().map(h => h.id), to: cfg.destinatario, cc: [], incluirHechas: true, comentario: 'Envío automático de fin del día.', origen: 'automático' });
}

/* ================= Fechas ================= */

function today_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function nowIso_() { return Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'HH:mm"); }
function d_(s) { const p = s.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2], 12); }
function ymd_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); }
function addDays_(s, n) { const d = d_(s); d.setDate(d.getDate() + n); return ymd_(d); }
function nextMonday_(hoy) { const dow = Number(Utilities.formatDate(d_(hoy), TZ, 'u')); return addDays_(hoy, 8 - dow); }
function fmtCorta_(s) { const d = d_(s); return DIAS[d.getDay()].slice(0, 3) + ' ' + d.getDate() + '/' + (d.getMonth() + 1); }
function fmtLarga_(s) { const d = d_(s); const w = DIAS[d.getDay()]; return w.charAt(0).toUpperCase() + w.slice(1) + ' ' + d.getDate() + ' de ' + MESES[d.getMonth()] + ' de ' + d.getFullYear(); }

function bucket_(f, hoy) {
  if (!f) return 'sin';
  if (f < hoy) return 'venc';
  if (f === hoy) return 'hoy';
  if (f === addDays_(hoy, 1)) return 'man';
  const nm = nextMonday_(hoy);
  if (f < nm) return 'semana';
  if (f <= addDays_(nm, 6)) return 'prox';
  return 'mas';
}

function cuando_(t, hoy) {
  const b = bucket_(t.fecha, hoy);
  if (b === 'sin') return '';
  if (b === 'hoy') return 'HOY';
  if (b === 'man') return 'Mañana';
  if (b === 'venc') return 'Vencida ' + fmtCorta_(t.fecha);
  return fmtCorta_(t.fecha);
}

/* ================= Informe y envío ================= */

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

function isEmail_(s) { return /^[^\s@,;]+@[^\s@,;]+\.[^\s@,;]+$/.test(String(s).trim()); }

function ordenar_(a, b) {
  if (a.prioridad !== b.prioridad) return a.prioridad ? -1 : 1;
  if ((a.fecha || '') !== (b.fecha || '')) return (a.fecha || '9') < (b.fecha || '9') ? -1 : 1;
  return String(a.creada) < String(b.creada) ? -1 : 1;
}

function reportData_(tareas, hoy, incluirHechas) {
  const pend = tareas.filter(t => !t.hecha);
  const hechasHoy = tareas.filter(t => t.hecha && String(t.completada).slice(0, 10) === hoy)
    .sort((a, b) => String(a.completada) < String(b.completada) ? -1 : 1);
  const grupos = GRUPOS.map(g => ({ key: g[0], label: g[1], items: pend.filter(t => bucket_(t.fecha, hoy) === g[0]).sort(ordenar_) }))
    .filter(g => g.items.length);
  const n = k => pend.filter(t => bucket_(t.fecha, hoy) === k).length;
  return { grupos: grupos, hechasHoy: incluirHechas ? hechasHoy : [], pendientes: pend.length, hoy: n('hoy'), vencidas: n('venc'), terminadasHoy: hechasHoy.length };
}

function pdfHtml_(hoja, r, hoy) {
  const fila = (t, cls) => '<tr class="' + cls + '"><td>' +
    (t.prioridad ? '<b>' + esc_(t.texto) + '</b> <span class="prio">Prioridad</span>' : esc_(t.texto)) +
    (t.nota ? '<div class="nota">' + esc_(t.nota).replace(/\n/g, '<br>') + '</div>' : '') +
    '</td><td class="c">' + (cls === 'hoy' ? '<span class="pill">HOY</span>' : esc_(cuando_(t, hoy))) +
    '</td><td class="c">' + esc_(t.persona) + '</td></tr>';
  const tabla = (items, cls) => '<table class="t"><tr><th>Tarea</th><th style="width:22%">Cuándo</th><th style="width:18%">Hablar con</th></tr>' +
    items.map(t => fila(t, typeof cls === 'function' ? cls(t) : cls)).join('') + '</table>';

  let html = '<html><head><meta charset="utf-8"><style>' +
    'body{font-family:Arial,Helvetica,sans-serif;color:#1C2630;font-size:10.5pt;margin:0}' +
    '.bar{height:6px;background:' + hoja.color + '}' +
    'h1{font-family:Georgia,serif;font-weight:normal;font-size:22pt;margin:16px 0 2px}' +
    '.sub{color:#56626D;font-size:10pt;margin:0 0 14px}' +
    '.stats{border-collapse:collapse;margin-bottom:6px}.stats td{padding:0 22px 0 0;color:#56626D;font-size:9pt}' +
    '.stats .v{font-size:16pt;color:#1C2630;font-weight:bold}' +
    'h2{font-size:11pt;margin:20px 0 6px;padding-bottom:4px;border-bottom:1px solid #E3E7EB}' +
    '.t{width:100%;border-collapse:collapse}' +
    '.t th{text-align:left;font-size:8.5pt;color:#8B959E;padding:4px 6px;border-bottom:1px solid #E3E7EB}' +
    '.t td{padding:6px;border-bottom:1px solid #EEF1F4;vertical-align:top}.t td.c{font-size:9.5pt;color:#56626D}' +
    'tr.hoy td{background:#EEF4FA}tr.venc td{background:#FBF0EE}tr.done td{color:#8B959E}' +
    '.pill{background:#1F4E79;color:#fff;font-weight:bold;font-size:8pt;padding:1px 6px}' +
    '.prio{color:#8A6D3B;font-size:8pt;font-weight:bold}' +
    '.nota{color:#56626D;font-size:9pt;margin-top:2px}' +
    '.foot{margin-top:28px;color:#8B959E;font-size:8pt}' +
    '</style></head><body><div class="bar"></div>' +
    '<h1>' + esc_(hoja.nombre) + '</h1><p class="sub">' + fmtLarga_(hoy) + '</p>' +
    '<table class="stats"><tr>' +
    '<td><div class="v">' + r.pendientes + '</div>pendientes</td>' +
    '<td><div class="v">' + r.hoy + '</div>para hoy</td>' +
    '<td><div class="v">' + r.vencidas + '</div>vencidas</td>' +
    '<td><div class="v">' + r.terminadasHoy + '</div>terminadas hoy</td>' +
    '</tr></table>';

  if (!r.grupos.length && !r.hechasHoy.length) html += '<p class="sub">Sin tareas en esta hoja.</p>';
  r.grupos.forEach(g => {
    const cls = g.key === 'hoy' ? 'hoy' : g.key === 'venc' ? 'venc' : '';
    html += '<h2>' + g.label + ' (' + g.items.length + ')</h2>' + tabla(g.items, cls);
  });
  if (r.hechasHoy.length) {
    html += '<h2>Terminadas hoy (' + r.hechasHoy.length + ')</h2><table class="t"><tr><th>Tarea</th><th style="width:22%">Terminada</th><th style="width:18%">Hablar con</th></tr>' +
      r.hechasHoy.map(t => '<tr class="done"><td>' + esc_(t.texto) + (t.nota ? '<div class="nota">' + esc_(t.nota) + '</div>' : '') +
        '</td><td class="c">' + esc_(String(t.completada).slice(11, 16)) + ' h</td><td class="c">' + esc_(t.persona) + '</td></tr>').join('') + '</table>';
  }
  html += '<p class="foot">Anotador diario — generado el ' + Utilities.formatDate(new Date(), TZ, 'dd/MM/yyyy HH:mm') + ' h</p></body></html>';
  return html;
}

function emailHtml_(resumen, hoy, comentario) {
  const td = 'padding:8px 10px;border-bottom:1px solid #EEF1F4;font-size:13px';
  const th = 'padding:6px 10px;border-bottom:1px solid #E3E7EB;font-size:11px;color:#8B959E;text-align:left';
  let paraHoy = [];
  resumen.forEach(x => x.r.grupos.filter(g => g.key === 'hoy' || g.key === 'venc')
    .forEach(g => g.items.forEach(t => paraHoy.push({ t: t, hoja: x.hoja, venc: g.key === 'venc' }))));

  let h = '<div style="font-family:Arial,Helvetica,sans-serif;color:#1C2630;max-width:640px">' +
    '<h2 style="font-family:Georgia,serif;font-weight:normal;font-size:24px;margin:0 0 4px">Anotador diario</h2>' +
    '<p style="color:#56626D;margin:0 0 18px;font-size:14px">' + fmtLarga_(hoy) + '</p>';
  if (comentario) h += '<p style="font-size:14px;line-height:1.5;border-left:3px solid #1F4E79;padding:4px 12px;margin:0 0 18px">' + esc_(comentario).replace(/\n/g, '<br>') + '</p>';

  h += '<table style="border-collapse:collapse;width:100%"><tr><th style="' + th + '">Hoja</th><th style="' + th + '">Pendientes</th><th style="' + th + '">Para hoy</th><th style="' + th + '">Vencidas</th><th style="' + th + '">Terminadas hoy</th></tr>' +
    resumen.map(x => '<tr><td style="' + td + ';border-left:4px solid ' + x.hoja.color + '"><b>' + esc_(x.hoja.nombre) + '</b></td><td style="' + td + '">' + x.r.pendientes +
      '</td><td style="' + td + '">' + x.r.hoy + '</td><td style="' + td + '">' + x.r.vencidas + '</td><td style="' + td + '">' + x.r.terminadasHoy + '</td></tr>').join('') + '</table>';

  if (paraHoy.length) {
    h += '<h3 style="font-size:14px;margin:24px 0 8px">Pendiente para hoy</h3><table style="border-collapse:collapse;width:100%">' +
      paraHoy.map(x => '<tr><td style="' + td + ';background:' + (x.venc ? '#FBF0EE' : '#EEF4FA') + '">' +
        (x.t.prioridad ? '<b>' + esc_(x.t.texto) + '</b>' : esc_(x.t.texto)) +
        (x.t.persona ? ' <span style="color:#56626D">— hablar con ' + esc_(x.t.persona) + '</span>' : '') +
        (x.venc ? ' <span style="color:#983A33;font-size:11px;font-weight:bold">VENCIDA</span>' : '') +
        '</td><td style="' + td + ';color:#8B959E;width:120px;background:' + (x.venc ? '#FBF0EE' : '#EEF4FA') + '">' + esc_(x.hoja.nombre) + '</td></tr>').join('') + '</table>';
  }
  h += '<p style="color:#8B959E;font-size:12px;margin-top:24px">Adjunto el detalle de cada hoja en PDF.</p></div>';
  return h;
}

function safeName_(s) { return s.replace(/[\\/:*?"<>|]/g, '-'); }

function sendReport_(p) {
  const hojas = getHojas_();
  const sel = hojas.filter(h => (p.hojaIds || []).indexOf(h.id) >= 0);
  if (!sel.length) throw new Error('Elegí al menos una hoja para enviar.');
  const to = String(p.to || '').trim();
  if (!isEmail_(to)) throw new Error('El destinatario no es un email válido: ' + to);
  const cc = (p.cc || []).map(s => String(s).trim()).filter(Boolean);
  const bad = cc.filter(x => !isEmail_(x));
  if (bad.length) throw new Error('CC inválido: ' + bad.join(', '));

  const tareas = getTareas_();
  const hoy = today_();
  const incluir = p.incluirHechas !== false;
  const resumen = [];
  const attachments = sel.map(h => {
    const r = reportData_(tareas.filter(t => t.hojaId === h.id), hoy, incluir);
    resumen.push({ hoja: h, r: r });
    return Utilities.newBlob(pdfHtml_(h, r, hoy), MimeType.HTML, 'hoja.html')
      .getAs(MimeType.PDF).setName(safeName_('Anotador - ' + h.nombre + ' - ' + hoy + '.pdf'));
  });

  const d = d_(hoy);
  const opts = {
    to: to,
    subject: 'Anotador diario – ' + Utilities.formatDate(d, TZ, 'dd/MM/yyyy'),
    htmlBody: emailHtml_(resumen, hoy, String(p.comentario || '').slice(0, 3000)),
    attachments: attachments,
    name: 'Anotador diario'
  };
  if (cc.length) opts.cc = cc.join(',');
  MailApp.sendEmail(opts);

  const pend = resumen.reduce((a, x) => a + x.r.pendientes, 0);
  const sh = ss_().getSheetByName(SH.E);
  if (sh) {
    const r = sh.getLastRow() + 1;
    sh.getRange(r, 1, 1, COLS_E.length).setNumberFormat('@')
      .setValues([[nowIso_(), to, cc.join(', '), sel.map(h => h.nombre).join(', '), String(pend), p.origen || '']]);
  }
  return { enviados: attachments.length, to: to, cc: cc };
}
