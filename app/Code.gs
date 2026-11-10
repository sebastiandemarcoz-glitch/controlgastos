/**
 * Control de Gastos — servidor (Google Apps Script)
 * Los datos se guardan en este mismo Google Sheet, en las pestañas:
 *   Movimientos · Cuentas · Conceptos · Config · Notificaciones · Reglas
 * Las pestañas se crean solas la primera vez que se abre la app.
 *
 * Se publica DOS veces como aplicación web (Implementar → Nueva implementación):
 *   1) "App":      Ejecutar como: yo · Acceso: solo yo            → la abrís en el celular y la PC
 *   2) "Receptor": Ejecutar como: yo · Acceso: cualquier persona  → la usa MacroDroid para mandar
 *                  las notificaciones (exige la clave que muestra la app en Ajustes).
 */

const VERSION_APP = '5.2.0';
const REPO = 'sebastiandemarcoz-glitch/controlgastos';   // de acá se bajan las actualizaciones
const RAW = 'https://raw.githubusercontent.com/' + REPO + '/main/app/';

const HDR = {
  Movimientos: ['ID', 'Fecha', 'Tipo', 'Cuenta', 'Cuenta destino', 'Concepto', 'Etiqueta', 'Descripción', 'Importe', 'Importe destino', 'Creado', 'Grupo', 'Cuota', 'Conciliado'],
  Cuentas: ['Nombre', 'Moneda', 'Saldo inicial', 'Archivada'],
  Conceptos: ['Nombre', 'Tipo', 'Categoría', 'Presupuesto mensual'],
  Config: ['Clave', 'Valor'],
  Notificaciones: ['ID', 'Recibida', 'App', 'Título', 'Texto', 'Estado', 'Importe', 'Moneda', 'Comercio', 'Cuenta', 'Concepto', 'Movimiento', 'Tipo', 'Cuotas'],
  Reglas: ['Tipo', 'Palabra', 'Valor', 'App'],
  PresupuestoAnual: ['ID', 'Nombre', 'Concepto', 'Mes', 'Año', 'Importe'],
  Fijos: ['ID', 'Nombre', 'Tipo', 'Cuenta', 'Concepto', 'Importe', 'Día', 'Desde', 'Hasta', 'Activo', 'Último mes'],
  Conciliacion: ['ID', 'Estado de cuenta', 'Cuenta', 'Diferencia', 'Fecha', 'Descripción', 'Importe', 'Moneda', 'Movimiento', 'Importe app', 'Cuota', 'Estado', 'Creado']
};
const TIPOS = { G: 'Gasto', I: 'Ingreso', T: 'Transferencia', A: 'Ajuste' };
const TIPO_R = { gasto: 'G', ingreso: 'I', transferencia: 'T', ajuste: 'A' };

/* Palabras clave precargadas: comercio → concepto. Se guardan en la pestaña Reglas y se pueden editar. */
const REGLAS_INICIALES = [
  ['Víveres', 'disco devoto geant tata ta-ta tienda inglesa el dorado macromercado kinko supermercado almacen carniceria frigorifico autoservicio'],
  ['Delivery', 'pedidosya pedidos ya rappi'],
  ['Auto', 'ancap axion petrobras disa peaje sucive estacionamiento parking telepeaje'],
  ['Streaming', 'netflix spotify dgo directv disney hbo prime video youtube paramount'],
  ['ANTEL', 'antel'], ['UTE', 'ute'], ['OSE', 'ose'], ['Gas', 'montevideo gas conecta'],
  ['Salud', 'farmacia farmashop san roque mutualista casmu medica uruguaya hospital'],
  ['Ferretería', 'ferreteria sodimac'], ['Mascotas', 'veterinaria'],
  ['Impuestos', 'dgi bps intendencia'],
  ['Esparcimiento', 'cine movie teatro']
];
const BANCOS = ['brou', 'bbva', 'santander', 'itau', 'scotiabank', 'scotia', 'hsbc', 'oca', 'prex', 'midinero', 'bandes', 'heritage', 'citibank'];

/* ---------- entrada web ---------- */
function doGet(e) {
  if (!esDueno_() && !(e && e.parameter && e.parameter.k === token_())) {
    return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Acceso no autorizado.</p>');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Control de Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

/** Recibe notificaciones desde el celular (MacroDroid u otra app). */
function doPost(e) {
  const P = (e && e.parameter) || {};
  const body = (e && e.postData && e.postData.contents) || '';
  let d = {};
  if (body.trim().charAt(0) === '{') { try { d = JSON.parse(body); } catch (x) { d = {}; } }
  if ((d.k || P.k) !== token_()) return txt_('clave incorrecta');
  let app = d.app || P.app || '', titulo = d.titulo || P.titulo || '', texto = d.texto || P.texto || '';
  if (!texto && body && body.trim().charAt(0) !== '{' && !P.texto) {
    // Cuerpo en texto plano: línea 1 = app, línea 2 = título, resto = texto
    const L = body.split(/\r?\n/).map(function (s) { return s.trim(); });
    if (L.length >= 3) { app = app || L.shift(); titulo = titulo || L.shift(); }
    else if (L.length === 2) { titulo = titulo || L.shift(); }
    texto = L.join(' ').trim();
  }
  if (!texto && !titulo) return txt_('vacia');
  // El receptor solo guarda la notificación tal cual llega; la app la procesa (cada 5 minutos y al abrirse).
  withLock_(function () {
    const s = sh_('Notificaciones'), n = s.getLastRow();
    s.getRange(n + 1, 1, 1, 6).setValues([['n' + Utilities.getUuid().replace(/-/g, '').slice(0, 12), new Date(), app, titulo, texto, 'nueva']]);
  });
  return txt_('recibida');
}
function txt_(s) { return ContentService.createTextOutput(String(s)); }

/* ---------- seguridad ---------- */
function esDueno_() {
  const yo = Session.getEffectiveUser().getEmail();
  const quien = Session.getActiveUser().getEmail();
  return !!quien && quien === yo;
}
function dueno_() { if (!esDueno_()) throw new Error('No autorizado'); }
function token_() {
  const p = PropertiesService.getScriptProperties();
  let t = p.getProperty('TOKEN');
  if (!t) { t = Utilities.getUuid().replace(/-/g, '').slice(0, 20); p.setProperty('TOKEN', t); }
  return t;
}

/* ---------- acceso a la planilla ---------- */
let _book = null, _tz = null;
function book_() {
  if (_book) return _book;
  _book = SpreadsheetApp.getActiveSpreadsheet();
  if (!_book) {
    const id = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    if (!id) throw new Error('El script no está dentro de un Google Sheet. Crealo desde Extensiones → Apps Script.');
    _book = SpreadsheetApp.openById(id);
  }
  return _book;
}
function tz_() { return _tz || (_tz = book_().getSpreadsheetTimeZone()); }
const _hdrOk = {};
function sh_(name) {
  const b = book_();
  let s = b.getSheetByName(name);
  if (s && !_hdrOk[name]) {   // agrega columnas nuevas a pestañas creadas por versiones anteriores
    const w = HDR[name].length, cur = s.getRange(1, 1, 1, w).getValues()[0];
    if (cur.some(function (v, i) { return String(v) !== HDR[name][i]; })) s.getRange(1, 1, 1, w).setValues([HDR[name]]).setFontWeight('bold');
    _hdrOk[name] = true;
  }
  if (!s) {
    _hdrOk[name] = true;
    s = b.insertSheet(name);
    s.getRange(1, 1, 1, HDR[name].length).setValues([HDR[name]]).setFontWeight('bold');
    s.setFrozenRows(1);
    if (name === 'Reglas') {
      const rows = [];
      REGLAS_INICIALES.forEach(function (r) {
        // Las palabras de dos términos van juntas (ej. "pedidos ya", "tienda inglesa")
        partir_(r[1]).forEach(function (w) { rows.push(['Concepto', w, r[0], '']); });
      });
      s.getRange(2, 1, rows.length, 4).setValues(rows);
    }
  }
  return s;
}
function partir_(s) {
  const dobles = ['tienda inglesa', 'el dorado', 'pedidos ya', 'prime video', 'montevideo gas', 'san roque', 'medica uruguaya'];
  let resto = ' ' + s + ' ';
  const out = [];
  dobles.forEach(function (d) { if (resto.indexOf(' ' + d + ' ') > -1) { out.push(d); resto = resto.replace(' ' + d + ' ', ' '); } });
  resto.trim().split(/\s+/).forEach(function (w) { if (w && w !== 'gas') out.push(w); });
  return out;
}
function body_(name) {
  const s = sh_(name), n = s.getLastRow();
  return n < 2 ? [] : s.getRange(2, 1, n - 1, HDR[name].length).getValues();
}
function writeBody_(name, data) {
  const s = sh_(name), n = s.getLastRow(), w = HDR[name].length;
  if (n > 1) s.getRange(2, 1, n - 1, w).clearContent();
  if (data.length) s.getRange(2, 1, data.length, w).setValues(data);
}
function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { return fn(); } finally { lock.releaseLock(); }
}
function pad_(n) { return ('0' + n).slice(-2); }
function hoy_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }

/* ---------- conversión fila <-> movimiento ---------- */
function fecha_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = String(v || '').trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return m[1] + '-' + pad_(m[2]) + '-' + pad_(m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + pad_(m[2]) + '-' + pad_(m[1]);
  return '';
}
function aFecha_(f) { const p = f.split('-'); return new Date(+p[0], +p[1] - 1, +p[2], 12); }
function num_(v) {
  if (typeof v === 'number') return v;
  let s = String(v || '').replace(/[^\d,.\-]/g, '').replace(/[.,]$/, '');
  if (!s) return 0;
  const lc = s.lastIndexOf(','), ld = s.lastIndexOf('.');
  if (lc > -1 && ld > -1) s = lc > ld ? s.replace(/\./g, '').replace(',', '.') : s.replace(/,/g, '');
  else if (lc > -1) s = s.replace(/\./g, '').replace(',', '.');
  else if (/^-?\d{1,3}(\.\d{3})+$/.test(s)) s = s.replace(/\./g, '');
  return parseFloat(s) || 0;
}
function movToRow_(x) {
  const imp = x.t === 'G' ? -Math.abs(x.m) : x.m;
  return [x.id, aFecha_(x.f), TIPOS[x.t] || 'Gasto', x.c || '', x.c2 || '', x.k || '', x.e || '', x.d || '',
    imp, x.m2 == null ? '' : x.m2, new Date(x.ts || Date.now()), x.g || '', x.q || '', x.cc || ''];
}
function rowToMov_(r) {
  const t = TIPO_R[String(r[2]).trim().toLowerCase()] || (num_(r[8]) < 0 ? 'G' : 'I');
  const m = num_(r[8]);
  const x = { id: String(r[0]), f: fecha_(r[1]), t: t, c: String(r[3]).trim(), m: t === 'A' ? m : Math.abs(m) };
  if (r[4] !== '') x.c2 = String(r[4]).trim();
  if (r[5] !== '') x.k = String(r[5]).trim();
  if (r[6] !== '') x.e = String(r[6]);
  if (r[7] !== '') x.d = String(r[7]);
  if (r[9] !== '' && r[9] != null) x.m2 = Math.abs(num_(r[9]));
  x.ts = r[10] instanceof Date ? r[10].getTime() : (Number(r[10]) || 0);
  if (r[11] !== '' && r[11] != null) x.g = String(r[11]);
  if (r[12] !== '' && r[12] != null) x.q = String(r[12]);
  if (r[13] !== '' && r[13] != null) x.cc = String(r[13]);
  return x;
}
function upsertMov_(x) {
  const s = sh_('Movimientos'), n = s.getLastRow(), row = movToRow_(x);
  if (n > 1) {
    const ids = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === x.id) { s.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
    }
  }
  s.getRange(Math.max(n, 1) + 1, 1, 1, row.length).setValues([row]);
}
function leerCfg_() {
  const cfg = { cuentas: [], conceptos: [] };
  body_('Cuentas').forEach(function (r) {
    if (String(r[0]).trim() === '') return;
    const mon = String(r[1]).toUpperCase();
    const a = String(r[3]).toLowerCase();
    cfg.cuentas.push({ n: String(r[0]).trim(), mon: (mon.indexOf('US') > -1 || mon.indexOf('LAR') > -1) ? 'USD' : 'UYU',
      si: num_(r[2]), arch: r[3] === true || a === 'true' || a === 'si' || a === 'sí' });
  });
  body_('Conceptos').forEach(function (r) {
    if (String(r[0]).trim() === '') return;
    cfg.conceptos.push({ n: String(r[0]).trim(), tipo: /^ing/i.test(String(r[1])) ? 'I' : 'G', cat: String(r[2] || '').trim(), pres: num_(r[3]) || 0 });
  });
  cfg.preguntarCuotas = getConfig_('preguntar_cuotas') || 'tc';
  return cfg;
}
function getConfig_(k) {
  const r = body_('Config').filter(function (x) { return x[0] === k; })[0];
  return r ? r[1] : '';
}
function setConfig_(k, v) {
  const data = body_('Config');
  const i = data.findIndex(function (x) { return x[0] === k; });
  if (i > -1) data[i][1] = v; else data.push([k, v]);
  writeBody_('Config', data);
}

/* ---------- cotización BROU (Dólar eBROU, venta) ---------- */
function cotizacion_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('cot');
  if (hit) return JSON.parse(hit);
  let o = { v: num_(getConfig_('cotizacion')) || 0, f: String(getConfig_('cotizacion_fecha') || ''), ok: false };
  const v = brou_();
  if (v) {
    o = { v: v, f: Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm'), ok: true };
    withLock_(function () {
      setConfig_('cotizacion', v); setConfig_('cotizacion_fecha', o.f); setConfig_('cotizacion_fuente', 'BROU · Dólar eBROU · venta');
    });
  }
  cache.put('cot', JSON.stringify(o), v ? 1800 : 300);
  return o;
}
function brou_() {
  const urls = ['https://www.brou.com.uy/cotizaciones', 'https://brou.com.uy/en/cotizaciones'];
  for (let i = 0; i < urls.length; i++) {
    try {
      const r = UrlFetchApp.fetch(urls[i], { muteHttpExceptions: true, followRedirects: true, headers: { 'User-Agent': 'Mozilla/5.0' } });
      if (r.getResponseCode() !== 200) continue;
      const v = parseBrou_(r.getContentText());
      if (v) return v;
    } catch (e) { /* se intenta la siguiente */ }
  }
  return 0;
}
function parseBrou_(html) {
  const t = String(html)
    .replace(/<script[\s\S]*?<\/script>/gi, ' ').replace(/<style[\s\S]*?<\/style>/gi, ' ')
    .replace(/<[^>]+>/g, ' ').replace(/&nbsp;/gi, ' ').replace(/&oacute;/gi, 'ó').replace(/&#0*243;/g, 'ó')
    .replace(/\s+/g, ' ');
  const m = t.match(/eBROU\s*\|?\s*(\d{1,3}[.,]\d+)\s*\|?\s*(\d{1,3}[.,]\d+)/i);
  if (!m) return 0;
  const v = parseFloat(m[2].replace(',', '.'));
  return v > 10 && v < 300 ? Math.round(v * 100) / 100 : 0;
}

/* ---------- lectura de notificaciones ---------- */
function normTxt_(s) { return String(s || '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase(); }
function palabras_(s) { return ' ' + normTxt_(s).replace(/[^a-z0-9]+/g, ' ').trim() + ' '; }

function parseNotif_(app, titulo, texto) {
  const full = [titulo, texto].filter(function (x) { return x; }).join('. ').replace(/\s+/g, ' ');
  const n = normTxt_(full);
  const p = { app: app || '', texto: full, importe: 0, mon: 'UYU', comercio: '', tipo: 'G', digitos: '', esMov: false };
  if (/rechazad|denegad|no aprobad|no fue aprobad|codigo de (verificacion|seguridad)|tu clave|token|inicio de sesion|iniciaste sesion/.test(n)) return p;
  let m = /(u\$s|us\$|usd|u\$d|uyu|\$u|\$)\s*([0-9][0-9.,]*)/i.exec(full);
  if (m) { p.importe = num_(m[2]); p.mon = /^(u\$s|us\$|usd|u\$d)$/i.test(m[1]) ? 'USD' : 'UYU'; }
  else {
    m = /([0-9][0-9.,]*)\s*(pesos|d[oó]lares|uyu|usd)/i.exec(full);
    if (m) { p.importe = num_(m[1]); p.mon = /d[oó]lar|usd/i.test(m[2]) ? 'USD' : 'UYU'; }
  }
  p.esMov = p.importe > 0;
  if (/acredit|recibiste|te transfiri|transferencia recibida|deposito|reintegro|devolucion|cobraste/.test(n)) p.tipo = 'I';
  if (/pago (de |a )?(tu |la )?tarjeta|pago tarjeta|pagaste tu tarjeta|transferencia (enviada|realizada)|transferiste/.test(n)) p.tipo = 'T';
  const cm = full.match(/comercio:?\s*([^.,;]+)/i);
  if (cm) p.comercio = cm[1].trim();
  else {
    const re = /\ben\s+(.+?)(?=\s+(?:con|por|en|el dia|el d[ií]a|el\s+\d+|a las|desde|usando|mediante|fue|ha sido|hoy|aprobad[ao]|autorizad[ao]|realizad[ao]|exitosa)\b|[,;]|\.\s|\.$|$)/gi;
    let x;
    while ((x = re.exec(full))) {
      const c = x[1].trim();
      if (/^\d/.test(c) || /^(tu|su|sus|tus|cuenta|caja|tarjeta|cuotas?|u\$s|usd|uyu|\$)\b/i.test(c) || /^(la|el|los|las|una?)\s+(cuenta|tarjeta|caja|fecha|dia|día|cuotas?)\b/i.test(c)) continue;
      p.comercio = c.slice(0, 40); break;
    }
  }
  const cq = full.match(/\b(\d{1,2})\s*cuotas\b/i);
  p.cuotas = cq ? +cq[1] : 1;
  if (!p.comercio && m) {   // "Compra U$S 20,00 ANTHROPIC CLAUDE tarjeta 4321"
    const resto = full.slice(full.indexOf(m[0]) + m[0].length);
    const cx = resto.match(/^\s*[,.]?\s*([A-Za-z][A-Za-z0-9*.&'\/ -]{2,40}?)(?=\s+(?:tarjeta|con|el|en|por|fecha|hora|aprobad[ao])\b|[.,;]|$)/i);
    if (cx && !/^(en|con|por|pesos|d[oó]lares)$/i.test(cx[1].trim())) p.comercio = cx[1].trim();
  }
  const dg = full.match(/(?:terminad[ao] en|finalizad[ao] en|final|\*+|x{3,4})\s*(\d{4})\b/i);
  if (dg) p.digitos = dg[1];
  return p;
}
function leerReglas_() {
  return body_('Reglas').filter(function (r) { return String(r[2]).trim(); }).map(function (r) {
    return { tipo: /^cuen/i.test(String(r[0])) ? 'Cuenta' : 'Concepto', palabra: String(r[1]).trim(), valor: String(r[2]).trim(), app: String(r[3] || '').trim() };
  });
}
function hermana_(a, mon, cfg) {
  const base = function (n) { return palabras_(n).replace(/ (pesos|dolares|uyu|usd|u s) /g, ' ').trim(); };
  const b = base(a.n);
  return cfg.cuentas.filter(function (c) { return c.mon === mon && base(c.n) === b; })[0] || null;
}
function resolver_(p, cfg, reglas) {
  const W = palabras_(p.app + ' ' + p.texto + ' ' + p.comercio);
  // concepto: la palabra más larga que aparezca como palabra completa
  let best = null;
  if (p.tipo === 'G' || p.tipo === 'I') {
    reglas.forEach(function (r) {
      if (r.tipo !== 'Concepto') return;
      const w = palabras_(r.palabra);
      if (w.trim() && W.indexOf(w) > -1 && cfg.conceptos.some(function (c) { return c.tipo === p.tipo && c.n === r.valor; }) &&
          (!best || w.length > best.w.length)) best = { w: w, v: r.valor };
    });
  }
  p.concepto = best ? best.v : '';
  // cuenta: primero reglas con palabra (más específicas), después reglas solo por app
  let cu = '';
  const rc = reglas.filter(function (r) { return r.tipo === 'Cuenta'; })
    .sort(function (a, b) { return (b.palabra ? 1 : 0) - (a.palabra ? 1 : 0); });
  for (let i = 0; i < rc.length && !cu; i++) {
    const r = rc[i];
    if (r.app && normTxt_(p.app).indexOf(normTxt_(r.app)) < 0) continue;
    if (!r.palabra && !r.app) continue;
    if (!r.palabra || W.indexOf(palabras_(r.palabra)) > -1) cu = r.valor;
  }
  if (cu) {
    const a = cfg.cuentas.filter(function (c) { return c.n === cu; })[0];
    if (!a) cu = '';
    else if (a.mon !== p.mon) { const h = hermana_(a, p.mon, cfg); cu = h ? h.n : ''; }
  }
  if (!cu) cu = heuristica_(p, cfg);
  p.cuenta = cu;
  return p;
}
function heuristica_(p, cfg) {
  const T = palabras_(p.app + ' ' + p.texto);
  const banco = BANCOS.filter(function (b) { return T.indexOf(' ' + b + ' ') > -1; })[0];
  if (!banco) return '';
  let c = cfg.cuentas.filter(function (a) { return !a.arch && a.mon === p.mon && palabras_(a.n).indexOf(' ' + banco + ' ') > -1; });
  if (c.length > 1) {
    const n = normTxt_(p.texto);
    if (/debito|caja de ahorro|cuenta corriente/.test(n)) c = c.filter(function (a) { return /^ca\b|caja/i.test(normTxt_(a.n)); });
    else if (/credito|visa|master|recompensa|tarjeta/.test(n)) c = c.filter(function (a) { return /^tc\b|tarjeta/i.test(normTxt_(a.n)); });
  }
  return c.length === 1 ? c[0].n : '';
}
/** Canal del aviso: el campo App viene como "APP|BBVA", "SMS|BROU" o "MAIL|remitente" (MacroDroid: sin prefijo = APP). */
function canal_(app) { const m = String(app || '').match(/^(APP|SMS|MAIL)\|/); return m ? m[1] : 'APP'; }
function sinCanal_(app) { return String(app || '').replace(/^(APP|SMS|MAIL)\|/, ''); }
const VENTANA_DUP = 3 * 3600 * 1000;   // mismo gasto por dos medios: hasta 3 horas de diferencia
function mismoComercio_(a, b) {
  if (!a || !b) return true;
  const A = palabras_(a).trim().split(' ').filter(function (w) { return w.length > 2 && !/^\d+$/.test(w); });
  const B = palabras_(b);
  return !A.length || A.some(function (w) { return B.indexOf(' ' + w + ' ') > -1; });
}
/** Procesa las notificaciones que dejó el receptor (estado "nueva"). */
function procesarNuevas_() {
  const s = sh_('Notificaciones'), n = s.getLastRow();
  if (n < 2) return 0;
  const w = HDR.Notificaciones.length;
  const data = s.getRange(2, 1, n - 1, w).getValues();
  if (!data.some(function (r) { return r[5] === 'nueva'; })) return 0;
  return withLock_(function () {
    const d2 = s.getRange(2, 1, s.getLastRow() - 1, w).getValues();
    const cfg = leerCfg_(), reglas = leerReglas_(), preg = getConfig_('preguntar_cuotas');
    let k = 0;
    d2.forEach(function (r, i) {
      if (r[5] !== 'nueva') return;
      k++;
      const t0 = r[1] instanceof Date ? r[1].getTime() : Date.now();
      // la misma notificación repetida en los 10 minutos anteriores
      for (let j = Math.max(0, i - 30); j < i; j++) {
        const q = d2[j];
        if (q[5] !== 'duplicada' && String(q[2]) === String(r[2]) && String(q[3]) === String(r[3]) && String(q[4]) === String(r[4]) &&
            q[1] instanceof Date && Math.abs(t0 - q[1].getTime()) < 600000) { r[5] = 'duplicada'; return; }
      }
      const p = parseNotif_(sinCanal_(r[2]), String(r[3]), String(r[4]));
      // el mismo gasto que ya llegó por otro medio (notificación, SMS o mail)
      if (p.esMov) {
        const can = canal_(r[2]);
        for (let j = 0; j < d2.length; j++) {
          const q = d2[j];
          if (j === i || q[5] === 'nueva' || q[5] === 'duplicada' || q[5] === 'ignorada' || !q[6]) continue;
          if (canal_(q[2]) === can || !(q[1] instanceof Date) || Math.abs(t0 - q[1].getTime()) > VENTANA_DUP) continue;
          if (String(q[7] || 'UYU') !== p.mon || Math.abs(num_(q[6]) - p.importe) > 0.01) continue;
          if (!mismoComercio_(p.comercio, String(q[8])) && !mismoComercio_(String(q[8]), p.comercio)) continue;
          d2[i] = [r[0], r[1], r[2], r[3], r[4], 'duplicada', p.importe, p.mon, p.comercio, q[9], q[10], q[11], p.tipo, ''];
          return;
        }
      }
      let estado = 'ignorada', mid = '', cuotas = '';
      if (p.esMov) {
        resolver_(p, cfg, reglas);
        estado = 'pendiente';
        if (p.cuenta && p.concepto && (p.tipo === 'G' || p.tipo === 'I')) {
          mid = 'm' + String(r[0]).slice(1);
          upsertMov_({ id: mid, f: Utilities.formatDate(new Date(t0), tz_(), 'yyyy-MM-dd'), t: p.tipo, c: p.cuenta, k: p.concepto,
            m: Math.round(p.importe * 100) / 100, d: p.comercio || '', ts: t0 });
          estado = 'asignada';
          if (p.tipo === 'G' && (p.cuotas > 1 || preg === 'todos' || esTarjeta_(p.cuenta))) cuotas = 'preguntar';
        }
      }
      d2[i] = [r[0], r[1], r[2], r[3], r[4], estado, p.importe || '', p.mon, p.comercio, p.cuenta || '', p.concepto || '', mid, p.tipo, cuotas];
    });
    s.getRange(2, 1, d2.length, w).setValues(d2);
    return k;
  });
}
/** Lo ejecuta un activador cada 5 minutos, así los gastos se asignan aunque no abras la app. */
function procesarNotificaciones() { leerGmail_(); procesarNuevas_(); }

/* ---------- mails de los bancos (Gmail) ---------- */
const REMITENTES_DEF = 'bbva, itau, brou, sistarbanc, oca, prex, santander, scotiabank, hsbc, midinero, heritage, bandes';
function leerGmail_() {
  if (getConfig_('leer_gmail') !== 'si') return 0;
  try {
    const desde = new Date(String(getConfig_('gmail_desde') || new Date().toISOString()));
    const rem = String(getConfig_('remitentes_mail') || REMITENTES_DEF).split(/[,;\s]+/).filter(String);
    const dia = Utilities.formatDate(new Date(desde.getTime() - 86400000), 'UTC', 'yyyy/MM/dd');
    const q = 'after:' + dia + ' (' + rem.map(function (r) { return 'from:' + r; }).join(' OR ') + ')';
    const s = sh_('Notificaciones'), n = s.getLastRow();
    const ids = {};
    if (n > 1) s.getRange(2, 1, n - 1, 1).getValues().forEach(function (r) { ids[String(r[0])] = 1; });
    const nuevas = [];
    GmailApp.search(q, 0, 40).forEach(function (th) {
      th.getMessages().forEach(function (m) {
        const id = 'g' + m.getId();
        if (ids[id] || m.getDate() < desde) return;
        ids[id] = 1;
        const asunto = m.getSubject() || '';
        if (/promo|beneficio|descuento exclusivo|sorte|gan[aá]|newsletter|oferta|invita|encuesta|novedad/i.test(asunto)) return;
        const cuerpo = (m.getPlainBody() || '').replace(/\s+/g, ' ').trim();
        const k = cuerpo.search(/(U\$S|US\$|USD|UYU|\$)\s?-?\d/i);
        const trozo = k < 0 ? '' : cuerpo.slice(Math.max(0, k - 220), k + 220);
        const texto = (trozo || cuerpo.slice(0, 300)).trim();
        if (!/(U\$S|US\$|USD|UYU|\$)\s?-?\d/i.test(asunto + ' ' + texto)) return;
        if (!/compra|consumo|d[eé]bito|transacci|transfer|acredit|pago|retiro|cobro|dep[oó]sito|cargo|operaci/i.test(asunto + ' ' + texto)) return;
        const deQuien = String(m.getFrom() || '').replace(/<.*>/, '').replace(/"/g, '').trim() || String(m.getFrom());
        nuevas.push([id, m.getDate(), 'MAIL|' + deQuien, asunto, texto, 'nueva']);
      });
    });
    if (nuevas.length) withLock_(function () {
      const s2 = sh_('Notificaciones'), n2 = s2.getLastRow();
      s2.getRange(n2 + 1, 1, nuevas.length, 6).setValues(nuevas);
    });
    setConfig_('gmail_estado', 'ok ' + Utilities.formatDate(new Date(), tz_(), 'dd/MM HH:mm') + (nuevas.length ? ' · ' + nuevas.length + ' nuevos' : ''));
    return nuevas.length;
  } catch (e) {
    setConfig_('gmail_estado', 'error: ' + (e.message || e));
    return 0;
  }
}
function activarGmail(si, remitentes) {
  dueno_();
  return withLock_(function () {
    setConfig_('leer_gmail', si ? 'si' : 'no');
    if (si) setConfig_('gmail_desde', new Date().toISOString());
    if (remitentes != null) setConfig_('remitentes_mail', String(remitentes));
    return 'ok';
  });
}
/** Ejecutala una vez desde el editor si Google pide permiso para leer Gmail. */
function autorizar() { GmailApp.getInboxUnreadCount(); return 'ok'; }

/* ---------- dirección del receptor (para configurar la app de Android) ---------- */
function urlReceptor_() {
  const guardada = String(getConfig_('url_receptor') || '');
  if (guardada) return guardada;
  try {
    const deps = api_('get', '/deployments').deployments || [];
    for (let i = 0; i < deps.length; i++) {
      const web = (deps[i].entryPoints || []).filter(function (e) { return e.entryPointType === 'WEB_APP'; })[0];
      const cfg = web && web.webApp && web.webApp.entryPointConfig;
      if (cfg && cfg.access === 'ANYONE_ANONYMOUS' && web.webApp.url) { setConfig_('url_receptor', web.webApp.url); return web.webApp.url; }
    }
  } catch (e) { /* sin API: se carga a mano */ }
  return '';
}
function guardarUrlReceptor(u) { dueno_(); return withLock_(function () { setConfig_('url_receptor', String(u || '').trim()); return 'ok'; }); }

/* ---------- actualización automática ---------- */
function cmpVer_(a, b) {
  const x = String(a).split('.').map(Number), y = String(b).split('.').map(Number);
  for (let i = 0; i < 3; i++) { if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) - (y[i] || 0); }
  return 0;
}
function versionRemota_(sinCache) {
  const c = CacheService.getScriptCache();
  const hit = !sinCache && c.get('verRemota');
  if (hit) return JSON.parse(hit);
  let o = null;
  try {
    const r = UrlFetchApp.fetch(RAW + 'version.json?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() === 200) o = JSON.parse(r.getContentText());
  } catch (e) { o = null; }
  c.put('verRemota', JSON.stringify(o), 3600);
  return o;
}
function api_(metodo, ruta, cuerpo) {
  const r = UrlFetchApp.fetch('https://script.googleapis.com/v1/projects/' + ScriptApp.getScriptId() + ruta, {
    method: metodo, contentType: 'application/json', muteHttpExceptions: true,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: cuerpo ? JSON.stringify(cuerpo) : undefined
  });
  const txt = r.getContentText();
  if (r.getResponseCode() >= 300) {
    if (/has not been used|is disabled|SERVICE_DISABLED|User has not enabled the Apps Script API/i.test(txt))
      throw new Error('Falta activar la “API de Google Apps Script” en script.google.com/home/usersettings');
    throw new Error('Google respondió ' + r.getResponseCode() + ': ' + txt.slice(0, 300));
  }
  return txt ? JSON.parse(txt) : {};
}
function actualizar_() {
  const rem = versionRemota_(true);
  if (!rem || !rem.version) throw new Error('No pude leer la versión publicada en GitHub');
  if (cmpVer_(rem.version, VERSION_APP) <= 0) return { ok: true, version: VERSION_APP, nada: true };
  const bajar = function (f) {
    const r = UrlFetchApp.fetch(RAW + f + '?t=' + Date.now(), { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error('No pude bajar ' + f);
    return r.getContentText();
  };
  const files = [
    { name: 'appsscript', type: 'JSON', source: bajar('appsscript.json') },
    { name: 'Code', type: 'SERVER_JS', source: bajar('Code.gs') },
    { name: 'Index', type: 'HTML', source: bajar('Index.html') }
  ];
  if (files[1].source.indexOf("VERSION_APP = '" + rem.version + "'") < 0) throw new Error('Los archivos de GitHub no coinciden con la versión ' + rem.version);
  api_('put', '/content', { files: files });
  const ver = api_('post', '/versions', { description: 'Control de Gastos ' + rem.version });
  // se actualizan solo las publicaciones de la app (acceso "solo yo"); el receptor queda como está
  const deps = (api_('get', '/deployments').deployments || []);
  let n = 0;
  deps.forEach(function (d) {
    if (!d.deploymentConfig || !d.deploymentConfig.versionNumber) return;
    const web = (d.entryPoints || []).filter(function (e) { return e.entryPointType === 'WEB_APP'; })[0];
    const acceso = web && web.webApp && web.webApp.entryPointConfig && web.webApp.entryPointConfig.access;
    if (acceso && acceso !== 'MYSELF') return;
    api_('put', '/deployments/' + d.deploymentId, { deploymentConfig: { scriptId: ScriptApp.getScriptId(), versionNumber: ver.versionNumber,
      manifestFileName: 'appsscript', description: 'Control de Gastos ' + rem.version } });
    n++;
  });
  setConfig_('ultima_actualizacion', Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm') + ' → ' + rem.version);
  CacheService.getScriptCache().remove('verRemota');
  return { ok: true, version: rem.version, publicaciones: n };
}
/** Botón "Actualizar" de la app. */
function actualizarApp() {
  dueno_();
  return JSON.stringify(withLock_(function () { return actualizar_(); }));
}
function buscarActualizacion() { dueno_(); versionRemota_(true); return JSON.stringify(infoVersion_()); }
/** Activador diario: instala sola la versión nueva si está activada la actualización automática. */
function autoActualizar() {
  if (getConfig_('auto_actualizar') === 'no') return;
  try { withLock_(function () { actualizar_(); }); } catch (e) { setConfig_('error_actualizacion', String(e.message || e)); }
}
function instalarActivadores_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('ACTIVADORES') === '2') return;
  try {
    const ex = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
    if (ex.indexOf('procesarNotificaciones') < 0) ScriptApp.newTrigger('procesarNotificaciones').timeBased().everyMinutes(5).create();
    if (ex.indexOf('autoActualizar') < 0) ScriptApp.newTrigger('autoActualizar').timeBased().everyDays(1).atHour(4).create();
    p.setProperty('ACTIVADORES', '2');
  } catch (e) { /* se reintenta la próxima vez */ }
}
function esTarjeta_(n) { return /^tc\b|tarjeta|cr[eé]dito|visa|master|\boca\b|amex/i.test(String(n || '')); }
function notifs_() {
  const tz = tz_();
  const all = body_('Notificaciones').filter(function (r) { return String(r[0]); }).map(function (r) {
    return { id: String(r[0]), fecha: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd HH:mm') : String(r[1]),
      app: String(r[2]), titulo: String(r[3]), texto: String(r[4]), estado: String(r[5]), importe: num_(r[6]), mon: String(r[7]) || 'UYU',
      comercio: String(r[8]), cuenta: String(r[9]), concepto: String(r[10]), mov: String(r[11]), tipo: String(r[12] || 'G'),
      pregunta: String(r[13]), digitos: '', cuotas: 1 };
  });
  all.forEach(function (x) { if (x.estado === 'pendiente' || x.pregunta === 'preguntar') { const p = parseNotif_(sinCanal_(x.app), x.titulo, x.texto); x.digitos = p.digitos; x.cuotas = p.cuotas; } });
  const pend = all.filter(function (x) { return x.estado === 'pendiente'; });
  const cuotas = all.filter(function (x) { return x.estado === 'asignada' && x.pregunta === 'preguntar'; });
  const rec = all.filter(function (x) { return x.estado !== 'pendiente'; }).slice(-40).reverse();
  return { pend: pend, rec: rec, cuotas: cuotas };
}
function setNotif_(id, cambios) {
  const s = sh_('Notificaciones'), n = s.getLastRow();
  if (n < 2) return;
  const ids = s.getRange(2, 1, n - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === id) {
      Object.keys(cambios).forEach(function (col) { s.getRange(i + 2, HDR.Notificaciones.indexOf(col) + 1).setValue(cambios[col]); });
      return;
    }
  }
}
function agregarReglas_(nuevas) {
  if (!nuevas || !nuevas.length) return;
  const data = body_('Reglas');
  nuevas.forEach(function (r) {
    if (!r.valor || (!r.palabra && !r.app)) return;
    const tipo = r.tipo === 'Cuenta' ? 'Cuenta' : 'Concepto';
    const i = data.findIndex(function (x) { return x[0] === tipo && normTxt_(x[1]) === normTxt_(r.palabra || '') && normTxt_(x[3]) === normTxt_(r.app || ''); });
    const row = [tipo, r.palabra || '', r.valor, r.app || ''];
    if (i > -1) data[i] = row; else data.push(row);
  });
  writeBody_('Reglas', data);
}

/* ---------- API que usa la app ---------- */
function getData() {
  dueno_();
  instalarActivadores_();
  leerGmail_();
  procesarNuevas_();
  const cfg = leerCfg_();
  const cot = cotizacion_();
  cfg.cot = cot.v || 40; cfg.cotFecha = cot.f; cfg.cotOk = cot.ok;
  const raw = body_('Movimientos');
  const movs = [];
  let fix = false;
  raw.forEach(function (r) {
    if (String(r[1]) === '' && String(r[3]) === '') return;
    if (String(r[0]) === '') { r[0] = Utilities.getUuid().replace(/-/g, '').slice(0, 14); fix = true; }
    const x = rowToMov_(r);
    if (x.f) movs.push(x);
  });
  if (fix) withLock_(function () { writeBody_('Movimientos', raw); });
  const nuevosFijos = generarFijos_();
  nuevosFijos.forEach(function (x) { movs.push(x); });
  return JSON.stringify({ cfg: cfg, movs: movs, sheet: book_().getUrl(), notifs: notifs_(), reglas: leerReglas_(), token: token_(),
    fijos: leerFijos_(), difs: leerDifs_(), extras: leerExtras_(), app: infoVersion_(), url: ScriptApp.getService().getUrl(),
    avisos: { gmail: getConfig_('leer_gmail') === 'si', gmailEstado: String(getConfig_('gmail_estado') || ''),
      remitentes: String(getConfig_('remitentes_mail') || REMITENTES_DEF), receptor: urlReceptor_() } });
}
function infoVersion_() {
  const rem = versionRemota_();
  return { version: VERSION_APP, remota: rem && rem.version || '', notas: rem && rem.notas || '', hay: !!(rem && cmpVer_(rem.version, VERSION_APP) > 0),
    auto: getConfig_('auto_actualizar') !== 'no', ultima: String(getConfig_('ultima_actualizacion') || ''), error: String(getConfig_('error_actualizacion') || ''), repo: REPO };
}
function saveMov(json) {
  dueno_();
  const x = JSON.parse(json);
  return withLock_(function () { upsertMov_(x); return 'ok'; });
}
function deleteMov(id) {
  dueno_();
  return withLock_(function () {
    const s = sh_('Movimientos'), n = s.getLastRow();
    if (n < 2) return 'ok';
    const ids = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) s.deleteRow(i + 2);
    return 'ok';
  });
}
/** Guarda cuentas y conceptos. ren = { c: {viejo: nuevo}, k: {'G|viejo': nuevo} } */
function saveCfg(json) {
  dueno_();
  const o = JSON.parse(json);
  return withLock_(function () {
    writeBody_('Cuentas', o.cuentas.map(function (c) { return [c.n, c.mon, c.si || 0, !!c.arch]; }));
    writeBody_('Conceptos', o.conceptos.map(function (c) { return [c.n, c.tipo === 'I' ? 'Ingreso' : 'Gasto', c.cat || '', c.pres || '']; }));
    const ren = o.ren || {}, rc = ren.c || {}, rk = ren.k || {};
    if (Object.keys(rc).length || Object.keys(rk).length) {
      const data = body_('Movimientos');
      data.forEach(function (r) {
        if (rc[r[3]] != null) r[3] = rc[r[3]];
        if (rc[r[4]] != null) r[4] = rc[r[4]];
        const key = TIPO_R[String(r[2]).toLowerCase()] + '|' + r[5];
        if (rk[key] != null) r[5] = rk[key];
      });
      writeBody_('Movimientos', data);
      const reg = body_('Reglas');
      reg.forEach(function (r) {
        if (/^cuen/i.test(r[0]) && rc[r[2]] != null) r[2] = rc[r[2]];
        if (/^conc/i.test(r[0]) && (rk['G|' + r[2]] != null || rk['I|' + r[2]] != null)) r[2] = rk['G|' + r[2]] != null ? rk['G|' + r[2]] : rk['I|' + r[2]];
      });
      writeBody_('Reglas', reg);
      const fj = body_('Fijos');
      fj.forEach(function (r) {
        if (rc[r[3]] != null) r[3] = rc[r[3]];
        const t = String(r[2]).toLowerCase().indexOf('ing') === 0 ? 'I' : 'G';
        if (rk[t + '|' + r[4]] != null) r[4] = rk[t + '|' + r[4]];
      });
      writeBody_('Fijos', fj);
      const ex = body_('PresupuestoAnual');
      ex.forEach(function (r) { if (rk['G|' + r[2]] != null) r[2] = rk['G|' + r[2]]; });
      writeBody_('PresupuestoAnual', ex);
    }
    return 'ok';
  });
}
function importMovs(json) {
  dueno_();
  const items = JSON.parse(json);
  return withLock_(function () {
    const data = body_('Movimientos');
    const idx = {};
    data.forEach(function (r, i) { idx[String(r[0])] = i; });
    let nuevos = 0;
    items.forEach(function (x) {
      const row = movToRow_(x);
      if (idx[x.id] != null) data[idx[x.id]] = row;
      else { idx[x.id] = data.length; data.push(row); nuevos++; }
    });
    writeBody_('Movimientos', data);
    return nuevos;
  });
}
/** Clasifica una notificación: guarda el movimiento y, si se pidió, aprende reglas. */
function asignarNotif(json) {
  dueno_();
  const o = JSON.parse(json);
  return withLock_(function () {
    upsertMov_(o.mov);
    setNotif_(o.id, { Estado: 'asignada', Cuenta: o.mov.c, Concepto: o.mov.k || '', Movimiento: o.mov.id });
    agregarReglas_(o.reglas);
    return 'ok';
  });
}
function estadoNotif(id, estado) {
  dueno_();
  return withLock_(function () { setNotif_(id, { Estado: estado }); return 'ok'; });
}
function saveReglas(json) {
  dueno_();
  const list = JSON.parse(json);
  return withLock_(function () {
    writeBody_('Reglas', list.map(function (r) { return [r.tipo, r.palabra || '', r.valor, r.app || '']; }));
    return 'ok';
  });
}
/** Prueba una notificación sin guardar nada. */
function probarNotif(app, texto) {
  dueno_();
  const p = parseNotif_(app, '', texto);
  if (p.esMov) resolver_(p, leerCfg_(), leerReglas_());
  return JSON.stringify(p);
}
/** Procesa de nuevo las pendientes con las reglas actuales. */
function reprocesar() {
  dueno_();
  return withLock_(function () {
    const cfg = leerCfg_(), reglas = leerReglas_(), s = sh_('Notificaciones'), n = s.getLastRow();
    if (n < 2) return 0;
    const data = s.getRange(2, 1, n - 1, HDR.Notificaciones.length).getValues();
    let k = 0;
    data.forEach(function (r) {
      if (r[5] !== 'pendiente') return;
      const p = resolver_(parseNotif_(sinCanal_(r[2]), r[3], r[4]), cfg, reglas);
      r[9] = p.cuenta || ''; r[10] = p.concepto || '';
      if (p.cuenta && p.concepto && (p.tipo === 'G' || p.tipo === 'I')) {
        const mid = 'm' + String(r[0]).slice(1);
        upsertMov_({ id: mid, f: Utilities.formatDate(r[1] instanceof Date ? r[1] : new Date(), tz_(), 'yyyy-MM-dd'), t: p.tipo, c: p.cuenta, k: p.concepto,
          m: Math.round(p.importe * 100) / 100, d: p.comercio || '', ts: Date.now() });
        r[5] = 'asignada'; r[11] = mid; k++;
      }
    });
    s.getRange(2, 1, data.length, HDR.Notificaciones.length).setValues(data);
    return k;
  });
}
function borrarTodo() {
  dueno_();
  return withLock_(function () {
    ['Movimientos', 'Cuentas', 'Conceptos'].forEach(function (n) { writeBody_(n, []); });
    return 'ok';
  });
}

/* ---------- varios movimientos a la vez ---------- */
function deleteMovs(json) {
  dueno_();
  const ids = JSON.parse(json), set = {};
  ids.forEach(function (i) { set[i] = 1; });
  return withLock_(function () {
    writeBody_('Movimientos', body_('Movimientos').filter(function (r) { return !set[String(r[0])]; }));
    return 'ok';
  });
}
function setConfig(k, v) { dueno_(); return withLock_(function () { setConfig_(k, v); return 'ok'; }); }

/* ---------- cuotas de compras que entraron por notificación ---------- */
function cuotasNotif(json) {
  dueno_();
  const o = JSON.parse(json);   // {id, borrar: idOriginal, movs: [...]}
  return withLock_(function () {
    if (o.movs && o.movs.length) {
      const set = {}; set[o.borrar] = 1;
      const data = body_('Movimientos').filter(function (r) { return !set[String(r[0])]; });
      o.movs.forEach(function (x) { data.push(movToRow_(x)); });
      writeBody_('Movimientos', data);
    }
    setNotif_(o.id, { Cuotas: 'ok' });
    return 'ok';
  });
}

/* ---------- gastos fijos ---------- */
function leerFijos_() {
  return body_('Fijos').filter(function (r) { return String(r[0]); }).map(function (r) {
    return { id: String(r[0]), n: String(r[1]), t: /^ing/i.test(String(r[2])) ? 'I' : 'G', c: String(r[3]), k: String(r[4]), m: num_(r[5]),
      dia: Number(r[6]) || 1, desde: mesStr_(r[7]), hasta: mesStr_(r[8]), activo: !(r[9] === false || /^(false|no)$/i.test(String(r[9]))), ult: mesStr_(r[10]) };
  });
}
function mesStr_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, tz_(), 'yyyy-MM');
  const m = String(v || '').match(/(\d{4})-(\d{1,2})/);
  return m ? m[1] + '-' + pad_(m[2]) : '';
}
function sumMes_(ym, n) { let y = +ym.slice(0, 4), m = +ym.slice(5, 7) + n; while (m > 12) { m -= 12; y++; } while (m < 1) { m += 12; y--; } return y + '-' + pad_(m); }
function diaMes_(ym, dia) { const y = +ym.slice(0, 4), m = +ym.slice(5, 7); const max = new Date(y, m, 0).getDate(); return ym + '-' + pad_(Math.min(dia, max)); }
/** Crea los movimientos de los gastos fijos hasta el mes actual (cada mes una sola vez). */
function generarFijos_() {
  const fijos = leerFijos_();
  if (!fijos.length) return [];
  const actual = hoy_().slice(0, 7), nuevos = [];
  let cambio = false;
  fijos.forEach(function (f) {
    if (!f.activo || !f.desde || !f.c || !f.m) return;
    let ym = f.ult ? sumMes_(f.ult, 1) : f.desde;
    if (ym < f.desde) ym = f.desde;
    let n = 0;
    while (ym <= actual && (!f.hasta || ym <= f.hasta) && n < 24) {
      nuevos.push({ id: 'f' + f.id + '-' + ym, f: diaMes_(ym, f.dia), t: f.t, c: f.c, k: f.k, m: f.m, d: f.n, e: 'Fijo', ts: Date.now() });
      f.ult = ym; ym = sumMes_(ym, 1); n++; cambio = true;
    }
  });
  if (cambio) withLock_(function () {
    const data = body_('Movimientos'), idx = {};
    data.forEach(function (r) { idx[String(r[0])] = 1; });
    nuevos.forEach(function (x) { if (!idx[x.id]) data.push(movToRow_(x)); });
    writeBody_('Movimientos', data);
    escribirFijos_(fijos);
  });
  return nuevos;
}
function escribirFijos_(list) {
  writeBody_('Fijos', list.map(function (f) { return [f.id, f.n, f.t === 'I' ? 'Ingreso' : 'Gasto', f.c, f.k, f.m, f.dia, f.desde, f.hasta || '', !!f.activo, f.ult || '']; }));
}
function saveFijos(json) {
  dueno_();
  const list = JSON.parse(json);
  withLock_(function () { escribirFijos_(list); });
  generarFijos_();
  return 'ok';
}

/* ---------- conciliación con estados de cuenta ---------- */
function leerDifs_() {
  const tz = tz_();
  return body_('Conciliacion').filter(function (r) { return String(r[0]) && r[11] === 'pendiente'; }).map(function (r) {
    return { id: String(r[0]), ec: String(r[1]), c: String(r[2]), tipo: String(r[3]), f: fecha_(r[4]), d: String(r[5]), m: num_(r[6]),
      mon: String(r[7]) || 'UYU', mov: String(r[8]), mApp: num_(r[9]), q: String(r[10]) };
  });
}
/** o = {ec, cuenta, conciliados: [ids], difs: [{tipo,f,d,m,mon,mov,mApp,q,c}]} */
function guardarConciliacion(json) {
  dueno_();
  const o = JSON.parse(json);
  return withLock_(function () {
    const set = {};
    (o.conciliados || []).forEach(function (i) { set[i] = 1; });
    const data = body_('Movimientos');
    data.forEach(function (r) { if (set[String(r[0])]) r[13] = o.ec; });
    writeBody_('Movimientos', data);
    // se reemplazan las diferencias pendientes de una carga anterior del mismo estado de cuenta
    const dif = body_('Conciliacion').filter(function (r) { return !(String(r[1]) === o.ec && r[11] === 'pendiente'); });
    (o.difs || []).forEach(function (x, i) {
      dif.push(['d' + Utilities.getUuid().replace(/-/g, '').slice(0, 12), o.ec, x.c || o.cuenta, x.tipo, aFecha_(x.f), x.d || '', x.m, x.mon || 'UYU',
        x.mov || '', x.mApp == null ? '' : x.mApp, x.q || '', 'pendiente', new Date()]);
    });
    writeBody_('Conciliacion', dif);
    return 'ok';
  });
}
function estadoDif(id, estado) {
  dueno_();
  return withLock_(function () {
    const s = sh_('Conciliacion'), n = s.getLastRow();
    if (n < 2) return 'ok';
    const ids = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === id) { s.getRange(i + 2, 12).setValue(estado); break; }
    return 'ok';
  });
}

/* ---------- gastos anuales previstos (presupuesto puntual por mes) ---------- */
function leerExtras_() {
  return body_('PresupuestoAnual').filter(function (r) { return String(r[0]) && String(r[2]); }).map(function (r) {
    return { id: String(r[0]), n: String(r[1]), k: String(r[2]), mes: Number(r[3]) || 1, anio: Number(r[4]) || 0, m: num_(r[5]) };
  });
}
function saveExtras(json) {
  dueno_();
  const list = JSON.parse(json);
  return withLock_(function () {
    writeBody_('PresupuestoAnual', list.map(function (x) { return [x.id, x.n, x.k, x.mes, x.anio || '', x.m]; }));
    return 'ok';
  });
}
