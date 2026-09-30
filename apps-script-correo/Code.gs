/**
 * Anotador diario — Correo con IA
 * Versión 1.0 — 2026-09-30 10:45 ARG
 * Autor: Germán Rodríguez
 *
 * Proyecto de Apps Script SEPARADO, creado con la cuenta amelife@amelife-911.com
 * (es la que tiene acceso a esa bandeja). Lee los hilos recientes, descarta los que
 * ya respondiste y le pide a Claude que detecte qué quedó pendiente.
 *
 * Instalación:
 *   1. Con la cuenta de AMELIFE: script.google.com > Nuevo proyecto. Pegar este archivo.
 *   2. Configuración del proyecto > Propiedades del script > agregar
 *      ANTHROPIC_API_KEY = (tu clave de console.anthropic.com).
 *   3. Ejecutar setupCorreo() y autorizar. En el registro aparece la clave del correo.
 *   4. Implementar > Nueva implementación > Aplicación web.
 *      Ejecutar como: Yo — Quién tiene acceso: Cualquier usuario.
 *   5. En el anotador: Ajustes > Correo con IA > pegar la URL /exec y la clave.
 */

const TZ = 'America/Argentina/Buenos_Aires';
const MODELO = 'claude-haiku-4-5-20251001';
const MAX_HILOS = 40;
const DIAS_RECORDAR_DESCARTES = 60;

/* ================= Setup ================= */

function setupCorreo() {
  const p = PropertiesService.getScriptProperties();
  let k = p.getProperty('API_KEY');
  if (!k) {
    k = 'co-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10);
    p.setProperty('API_KEY', k);
  }
  GmailApp.getInboxUnreadCount();
  Logger.log('Cuenta: ' + Session.getEffectiveUser().getEmail());
  Logger.log('Clave del correo: ' + k);
  Logger.log(p.getProperty('ANTHROPIC_API_KEY') ? 'Clave de IA: configurada' : 'FALTA ANTHROPIC_API_KEY en Propiedades del script');
  return k;
}

/* ================= API web ================= */

function doGet() {
  return json_({ ok: true, app: 'Anotador correo', version: '1.0' });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'Solicitud inválida' });
  }
  const key = PropertiesService.getScriptProperties().getProperty('API_KEY');
  if (!key || req.key !== key) return json_({ ok: false, auth: true, error: 'Clave del correo incorrecta' });
  try {
    const p = req.payload || {};
    if (req.action === 'pendientes') return json_({ ok: true, data: pendientes_(p) });
    if (req.action === 'descartar') return json_({ ok: true, data: descartar_(p.ids || []) });
    return json_({ ok: false, error: 'Acción desconocida: ' + req.action });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/* ================= Descartes ================= */

function getDescartes_() {
  let d = {};
  try { d = JSON.parse(PropertiesService.getScriptProperties().getProperty('DESCARTES') || '{}'); } catch (e) { d = {}; }
  const limite = Date.now() - DIAS_RECORDAR_DESCARTES * 86400000;
  Object.keys(d).forEach(function (id) { if (d[id] < limite) delete d[id]; });
  return d;
}

function descartar_(ids) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const d = getDescartes_();
    ids.slice(0, 200).forEach(function (id) { d[String(id)] = Date.now(); });
    PropertiesService.getScriptProperties().setProperty('DESCARTES', JSON.stringify(d));
    return Object.keys(d).length;
  } finally {
    lock.releaseLock();
  }
}

/* ================= Lectura del correo ================= */

function limpiar_(txt) {
  const lineas = String(txt || '').replace(/\r/g, '').split('\n');
  const out = [];
  for (let i = 0; i < lineas.length; i++) {
    const l = lineas[i];
    if (/^\s*>/.test(l)) continue;
    if (/^(El|On)\s.+(escribió|wrote):\s*$/i.test(l.trim())) break;
    if (/^-{2,}\s*(Original Message|Mensaje original|Forwarded message|Mensaje reenviado)/i.test(l.trim())) break;
    if (/^(De|From):\s.+/i.test(l.trim()) && out.length > 3) break;
    out.push(l);
  }
  return out.join('\n').replace(/\n{3,}/g, '\n\n').replace(/[ \t]{2,}/g, ' ').trim();
}

function esMio_(from, yo) {
  return String(from || '').toLowerCase().indexOf(yo) >= 0;
}

function pendientes_(p) {
  const dias = Math.min(30, Math.max(1, parseInt(p.dias, 10) || 7));
  const yo = Session.getEffectiveUser().getEmail().toLowerCase();
  const descartes = getDescartes_();
  const q = 'in:inbox newer_than:' + dias + 'd -category:promotions -category:social -category:forums -category:updates';
  const hilos = GmailApp.search(q, 0, MAX_HILOS * 2);
  const cand = [];

  for (let i = 0; i < hilos.length && cand.length < MAX_HILOS; i++) {
    const th = hilos[i];
    const id = th.getId();
    if (descartes[id]) continue;
    const msgs = th.getMessages();
    const last = msgs[msgs.length - 1];
    const from = last.getFrom();
    if (esMio_(from, yo)) continue;
    if (/no-?reply|notifications?@|notificaciones@|mailer-daemon|postmaster@/i.test(from)) continue;
    cand.push({
      id: id,
      asunto: th.getFirstMessageSubject() || '(sin asunto)',
      de: from,
      fecha: Utilities.formatDate(last.getDate(), TZ, 'dd/MM HH:mm'),
      mensajes: msgs.length,
      participe: msgs.some(function (m) { return esMio_(m.getFrom(), yo); }),
      adjuntos: last.getAttachments({ includeInlineImages: false }).map(function (a) { return a.getName(); }).slice(0, 5),
      texto: limpiar_(last.getPlainBody()).slice(0, 1800)
    });
  }

  if (!cand.length) return { items: [], revisados: 0, cuenta: yo };

  const porId = {};
  cand.forEach(function (c) { porId[c.id] = c; });
  const items = analizar_(cand)
    .filter(function (x) { return x && x.requiere && porId[x.id]; })
    .map(function (x) {
      const c = porId[x.id];
      return {
        id: c.id,
        titulo: String(x.titulo || c.asunto).slice(0, 200),
        resumen: String(x.resumen || '').slice(0, 1200),
        contacto: String(x.contacto || '').slice(0, 30),
        urgencia: ['alta', 'normal', 'baja'].indexOf(x.urgencia) >= 0 ? x.urgencia : 'normal',
        asunto: c.asunto,
        de: c.de.replace(/\s*<[^>]+>/, ''),
        fecha: c.fecha,
        link: 'https://mail.google.com/mail/?authuser=' + encodeURIComponent(yo) + '#all/' + c.id
      };
    });
  const orden = { alta: 0, normal: 1, baja: 2 };
  items.sort(function (a, b) { return orden[a.urgencia] - orden[b.urgencia]; });
  return { items: items, revisados: cand.length, cuenta: yo };
}

/* ================= IA ================= */

function analizar_(cand) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('ANTHROPIC_API_KEY');
  if (!apiKey) throw new Error('Falta ANTHROPIC_API_KEY en las propiedades del script de correo.');

  const system = [
    'Sos el asistente de Germán Rodríguez, gerente comercial de AME LIFE 911 (distribuidor de equipamiento médico para Latinoamérica).',
    'Recibís el último mensaje de hilos recientes de su bandeja amelife@amelife-911.com que todavía no respondió.',
    'Decidí en cada uno si le requiere una respuesta o una acción concreta (cotizar, enviar documentación, confirmar, coordinar, pagar, reclamar).',
    'Marcá requiere=false para newsletters, notificaciones automáticas, avisos informativos sin pregunta, confirmaciones de recepción y mensajes que solo agradecen o cierran el tema.',
    'Respondé SOLO con un array JSON válido, sin texto antes ni después y sin bloques de código. Un objeto por cada id recibido, con esta forma:',
    '{"id":"...","requiere":true,"titulo":"acción concreta en español rioplatense, máximo 70 caracteres, empezando con verbo (Responder a…, Enviar…, Cotizar…, Confirmar…)","resumen":"1 a 3 líneas con qué pide y los datos clave: cantidades, modelos, fechas, montos","contacto":"nombre corto de la persona o empresa","urgencia":"alta|normal|baja"}',
    'urgencia alta: plazos cercanos, reclamos, licitaciones o clientes esperando respuesta hace días.'
  ].join('\n');

  const partes = cand.map(function (c) {
    return 'ID: ' + c.id +
      '\nAsunto: ' + c.asunto +
      '\nDe: ' + c.de +
      '\nFecha: ' + c.fecha +
      '\nMensajes en el hilo: ' + c.mensajes + (c.participe ? ' (Germán ya escribió antes en este hilo)' : '') +
      (c.adjuntos.length ? '\nAdjuntos: ' + c.adjuntos.join(', ') : '') +
      '\n---\n' + c.texto;
  });

  const res = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-api-key': apiKey, 'anthropic-version': '2023-06-01' },
    payload: JSON.stringify({
      model: MODELO,
      max_tokens: 6000,
      system: system,
      messages: [{ role: 'user', content: 'Correos a analizar:\n\n' + partes.join('\n\n=====\n\n') }]
    }),
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code !== 200) {
    let msg = body.slice(0, 300);
    try { msg = JSON.parse(body).error.message; } catch (e) {}
    throw new Error('La IA respondió con error ' + code + ': ' + msg);
  }
  const data = JSON.parse(body);
  const texto = (data.content || []).filter(function (b) { return b.type === 'text'; }).map(function (b) { return b.text; }).join('');
  const a = texto.indexOf('['), b = texto.lastIndexOf(']');
  if (a < 0 || b < a) throw new Error('La IA no devolvió una lista válida.');
  return JSON.parse(texto.slice(a, b + 1));
}
