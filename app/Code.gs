/**
 * Control de Gastos — lógica de la app (la baja y ejecuta el Cargador.gs desde GitHub).
 * Los datos se guardan en este mismo Google Sheet, en las pestañas:
 *   Movimientos · Cuentas · Conceptos · Config · Notificaciones · Reglas
 * Las pestañas se crean solas la primera vez que se abre la app.
 *
 * Se publica DOS veces como aplicación web (Implementar → Nueva implementación):
 *   1) "App":      Ejecutar como: yo · Acceso: solo yo            → la abrís en el celular y la PC
 *   2) "Receptor": Ejecutar como: yo · Acceso: cualquier persona  → la usa MacroDroid para mandar
 *                  las notificaciones (exige la clave que muestra la app en Ajustes).
 */

const VERSION_APP = '5.5.0';
const REPO = 'sebastiandemarcoz-glitch/controlgastos';   // de acá se bajan las actualizaciones
const RAW = 'https://raw.githubusercontent.com/' + REPO + '/main/app/';

const HDR = {
  Movimientos: ['ID', 'Fecha', 'Tipo', 'Cuenta', 'Cuenta destino', 'Concepto', 'Etiqueta', 'Descripción', 'Importe', 'Importe destino', 'Creado', 'Grupo', 'Cuota', 'Conciliado'],
  Cuentas: ['Nombre', 'Moneda', 'Saldo inicial', 'Archivada'],
  Conceptos: ['Nombre', 'Tipo', 'Categoría', 'Presupuesto mensual'],
  Config: ['Clave', 'Valor'],
  Notificaciones: ['ID', 'Recibida', 'App', 'Título', 'Texto', 'Estado', 'Importe', 'Moneda', 'Comercio', 'Cuenta', 'Concepto', 'Movimiento', 'Tipo', 'Cuotas'],
  Reglas: ['Tipo', 'Palabra', 'Valor', 'App'],
  PresupuestoAnual: ['ID', 'Nombre', 'Concepto', 'Mes', 'Año', 'Importe', 'Tipo'],
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
/** La pantalla solo se muestra al dueño (o con la clave). La arma el Cargador. */
function puedeVer(e) { return esDueno_() || !!(e && e.parameter && e.parameter.k === token_()); }

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
    setConfig_('ultimo_aviso', Utilities.formatDate(new Date(), tz_(), 'dd/MM HH:mm') + ' · ' + sinCanal_(app));
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
      if (/PRUEBA GASTOS AVISOS/i.test(String(r[4]))) { d2[i] = [r[0], r[1], r[2], r[3], r[4], 'prueba', '', '', '', '', '', '', '', '']; return; }
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
function procesarNotificaciones() { leerGmail_(); procesarNuevas_(); try { cotizacion_(); } catch (e) { } }

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
function urlReceptor_() { return String(getConfig_('url_receptor') || ''); }
/** Manda un aviso de prueba al receptor desde el servidor y explica qué pasa. */
function probarReceptor() {
  dueno_();
  const url = urlReceptor_();
  if (!url) return JSON.stringify({ ok: false, msg: 'Falta cargar la dirección del receptor (la implementación con acceso “Cualquier persona”).' });
  if (!/^https:\/\/script\.google\.com\/macros\/s\/[^/]+\/exec$/.test(url))
    return JSON.stringify({ ok: false, msg: 'La dirección no tiene el formato correcto: tiene que ser https://script.google.com/macros/s/…/exec' });
  const marca = 'PRUEBA GASTOS AVISOS ' + Utilities.getUuid().slice(0, 6);
  const cuerpo = JSON.stringify({ k: token_(), app: 'APP|Prueba del receptor', titulo: 'Prueba', texto: 'Compra por $ 1,00 en ' + marca });
  let r;
  try { r = UrlFetchApp.fetch(url, { method: 'post', contentType: 'application/json', payload: cuerpo, followRedirects: false, muteHttpExceptions: true }); }
  catch (e) { return JSON.stringify({ ok: false, msg: 'No pude conectarme con el receptor: ' + e.message }); }
  const code = r.getResponseCode(), loc = String((r.getAllHeaders() || {}).Location || (r.getAllHeaders() || {}).location || '');
  if (/accounts\.google\.com|ServiceLogin/i.test(loc) || (code === 200 && /accounts\.google\.com|Iniciar sesi/i.test(r.getContentText().slice(0, 3000))))
    return JSON.stringify({ ok: false, msg: 'Esa dirección pide iniciar sesión: es la implementación “Solo yo”. Usá la de acceso “Cualquier persona” (o cambiá el acceso de esa implementación a “Cualquier persona”).' });
  let texto = '';
  try { texto = code === 302 && loc ? UrlFetchApp.fetch(loc, { muteHttpExceptions: true }).getContentText().slice(0, 300) : r.getContentText().slice(0, 300); } catch (e) { }
  Utilities.sleep(1500);
  const filas = body_('Notificaciones');
  const llego = filas.some(function (f) { return String(f[4]).indexOf(marca) > -1; });
  if (llego) return JSON.stringify({ ok: true, msg: 'El receptor funciona: el aviso de prueba llegó a la planilla.' });
  if (/clave incorrecta/.test(texto)) return JSON.stringify({ ok: false, msg: 'El receptor respondió “clave incorrecta”: la implementación “Cualquier persona” usa otra planilla o una versión vieja. Editala y elegí “Nueva versión”.' });
  return JSON.stringify({ ok: false, msg: 'El receptor respondió (' + code + ') pero el aviso no llegó a la planilla. Respuesta: ' + (texto.replace(/<[^>]+>/g, ' ').replace(/\s+/g, ' ').trim().slice(0, 200) || 'vacía') + '. Editá la implementación “Cualquier persona” y elegí “Nueva versión”.' });
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
    else o = { error: 'GitHub respondió ' + r.getResponseCode() };
  } catch (e) { o = { error: String(e.message || e) }; }
  c.put('verRemota', JSON.stringify(o), o && o.version ? 3600 : 60);
  return o;
}
/** Botón "Actualizar ahora": lo hace el Cargador (baja la versión nueva de GitHub y la guarda en la planilla). */
function actualizarApp() {
  dueno_();
  const r = cgActualizar_(true);
  if (r.ok && !r.nada) {
    setConfig_('ultima_actualizacion', Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm') + ' → ' + r.version);
    setConfig_('error_actualizacion', '');
  }
  CacheService.getScriptCache().remove('verRemota');
  return JSON.stringify(r);
}
function buscarActualizacion() { dueno_(); versionRemota_(true); return JSON.stringify(infoVersion_()); }
function autoActualizarPermitido() { return getConfig_('auto_actualizar') !== 'no'; }
function anotarActualizacion(version, error) {
  if (error) setConfig_('error_actualizacion', error);
  else { setConfig_('ultima_actualizacion', Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm') + ' → ' + version); setConfig_('error_actualizacion', ''); }
}
function instalarActivadores_() {
  const p = PropertiesService.getScriptProperties();
  if (p.getProperty('ACTIVADORES') === '3') return;
  try {
    const ex = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
    if (ex.indexOf('procesarNotificaciones') < 0) ScriptApp.newTrigger('procesarNotificaciones').timeBased().everyMinutes(5).create();
    if (ex.indexOf('autoActualizar') < 0) ScriptApp.newTrigger('autoActualizar').timeBased().everyDays(1).atHour(4).create();
    p.setProperty('ACTIVADORES', '3');
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
  const rec = all.filter(function (x) { return x.estado !== 'pendiente' && x.estado !== 'descartada' && x.estado !== 'nueva'; }).slice(-40).reverse();
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
  try { procesarNuevas_(); } catch (e) { /* se reintenta en segundo plano */ }
  const cfg = leerCfg_();
  // la cotización guardada (se actualiza en segundo plano con extrasCarga y el activador)
  cfg.cot = num_(getConfig_('cotizacion')) || 40; cfg.cotFecha = String(getConfig_('cotizacion_fecha') || ''); cfg.cotOk = true;
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
    fijos: leerFijos_(), difs: leerDifs_(), extras: leerExtras_(), app: infoVersion_(true), url: ScriptApp.getService().getUrl(),
    avisos: { gmail: getConfig_('leer_gmail') === 'si', gmailEstado: String(getConfig_('gmail_estado') || ''),
      remitentes: String(getConfig_('remitentes_mail') || REMITENTES_DEF), receptor: String(getConfig_('url_receptor') || ''),
      ultimo: String(getConfig_('ultimo_aviso') || '') } });
}
/** Lo que necesita internet (BROU, GitHub, Gmail, Google): se pide después de mostrar la app, sin hacerla esperar. */
function extrasCarga() {
  dueno_();
  const r = {};
  try { instalarActivadores_(); } catch (e) { }
  try { const c = cotizacion_(); r.cot = c.v; r.cotFecha = c.f; r.cotOk = c.ok; } catch (e) { r.cotOk = false; }
  try { r.app = infoVersion_(); } catch (e) { r.app = null; }
  try { r.receptor = urlReceptor_(); } catch (e) { r.receptor = ''; }
  try { r.mails = leerGmail_(); if (r.mails) procesarNuevas_(); } catch (e) { r.mails = 0; }
  r.gmailEstado = String(getConfig_('gmail_estado') || '');
  return JSON.stringify(r);
}
function infoVersion_(soloCache) {
  const hit = soloCache && CacheService.getScriptCache().get('verRemota');
  const rem = soloCache ? (hit ? JSON.parse(hit) : { version: '' }) : versionRemota_();
  return { version: VERSION_APP, remota: rem && rem.version || '', notas: rem && rem.notas || '', hay: !!(rem && rem.version && cmpVer_(rem.version, VERSION_APP) > 0),
    errorRemota: rem && rem.error || (rem ? '' : 'sin respuesta de GitHub'), pendiente: !!(soloCache && rem && !rem.version && !rem.error),
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
      ex.forEach(function (r) { const t = /^ing/i.test(String(r[6])) ? 'I' : 'G'; if (rk[t + '|' + r[2]] != null) r[2] = rk[t + '|' + r[2]]; });
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
function borrarNotif(id) {
  dueno_();
  return withLock_(function () {
    const s = sh_('Notificaciones'), n = s.getLastRow();
    if (n < 2) return 'ok';
    const ids = s.getRange(2, 1, n - 1, 1).getValues();
    for (let i = ids.length - 1; i >= 0; i--) if (String(ids[i][0]) === id) s.deleteRow(i + 2);
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
    return { id: String(r[0]), n: String(r[1]), k: String(r[2]), mes: Number(r[3]) || 1, anio: Number(r[4]) || 0, m: num_(r[5]), t: /^ing/i.test(String(r[6])) ? 'I' : 'G' };
  });
}
function saveExtras(json) {
  dueno_();
  const list = JSON.parse(json);
  return withLock_(function () {
    writeBody_('PresupuestoAnual', list.map(function (x) { return [x.id, x.n, x.k, x.mes, x.anio || '', x.m, x.t === 'I' ? 'Ingreso' : 'Gasto']; }));
    return 'ok';
  });
}

/* ---------- funciones que usa el Cargador ---------- */
var API_ = {
  pantalla: function () { return JSON.stringify({ js: PANTALLA_JS, largo: PANTALLA_JS.length }); },
  puedeVer: puedeVer, doPost: doPost, procesarNotificaciones: procesarNotificaciones,
  autoActualizarPermitido: autoActualizarPermitido, anotarActualizacion: anotarActualizacion,
  getData: getData, extrasCarga: extrasCarga, saveMov: saveMov, deleteMov: deleteMov, deleteMovs: deleteMovs,
  saveCfg: saveCfg, importMovs: importMovs, borrarTodo: borrarTodo, asignarNotif: asignarNotif, estadoNotif: estadoNotif,
  borrarNotif: borrarNotif, saveReglas: saveReglas, probarNotif: probarNotif, reprocesar: reprocesar, cuotasNotif: cuotasNotif,
  saveFijos: saveFijos, guardarConciliacion: guardarConciliacion, estadoDif: estadoDif, setConfig: setConfig,
  saveExtras: saveExtras, activarGmail: activarGmail, guardarUrlReceptor: guardarUrlReceptor, probarReceptor: probarReceptor,
  buscarActualizacion: buscarActualizacion, actualizarApp: actualizarApp, autorizar: autorizar
};

/* ---------- código de la pantalla (lo pide Index.html con rpc("pantalla")) ---------- */
var PANTALLA_JS = "\n/* ===================== Estado ===================== */\nconst S={cfg:null,movs:[],ready:false,err:null,sheet:'',notifs:{pend:[],rec:[],cuotas:[]},reglas:[],token:'',fijos:[],difs:[]};\nconst V={tab:'inicio',per:ym(hoy()),f:{c:'',k:'',t:'',q:'',e:''},show:300,\n  inf:{per:ym(hoy()),tipo:'G',by:'cat'},verArch:false};\nlet F=null;\nconst MESES=['enero','febrero','marzo','abril','mayo','junio','julio','agosto','septiembre','octubre','noviembre','diciembre'];\nconst MES3=['ene','feb','mar','abr','may','jun','jul','ago','sep','oct','nov','dic'];\nconst DIAS=['domingo','lunes','martes','mi\u00e9rcoles','jueves','viernes','s\u00e1bado'];\n\n/* ===================== Utilidades ===================== */\nfunction p2(n){return String(n).padStart(2,'0')}\nfunction hoy(){const d=new Date();return d.getFullYear()+'-'+p2(d.getMonth()+1)+'-'+p2(d.getDate())}\nfunction ym(f){return f.slice(0,7)}\nfunction uid(){return Date.now().toString(36)+Math.random().toString(36).slice(2,7)}\nfunction h(s){return String(s??'').replace(/[&<>\"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','\"':'&quot;',\"'\":'&#39;'}[c]))}\nfunction js(s){return h(JSON.stringify(String(s??'')))}\nfunction norm(s){return String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\\s+/g,' ').trim()}\nfunction clone(o){return JSON.parse(JSON.stringify(o))}\nconst NF=new Intl.NumberFormat('es-UY',{minimumFractionDigits:2,maximumFractionDigits:2});\nfunction sym(mon){return mon==='USD'?'U$S':'$'}\nfunction fmt(v,mon){return sym(mon)+' '+NF.format(Math.abs(v)<0.005?0:v)}\nfunction fmtS(v,mon){return (v<-0.004?'\u2212':'')+fmt(Math.abs(v),mon)}\nfunction parseNum(v){\n  if(typeof v==='number')return v;\n  let s=String(v??'').trim();const par=/^\\(.*\\)$/.test(s);\n  s=s.replace(/[^\\d,.\\-]/g,'');if(!s||s==='-')return NaN;\n  const lc=s.lastIndexOf(','),ld=s.lastIndexOf('.');\n  if(lc>-1&&ld>-1){s=lc>ld?s.replace(/\\./g,'').replace(',','.'):s.replace(/,/g,'')}\n  else if(lc>-1){s=s.replace(/\\./g,'').replace(',','.')}\n  else if(/^-?\\d{1,3}(\\.\\d{3})+$/.test(s)){s=s.replace(/\\./g,'')}\n  const n=parseFloat(s);return par?-Math.abs(n):n;\n}\nfunction r2(n){return Math.round(n*100)/100}\nfunction perLabel(p){if(!p)return 'todo el per\u00edodo';if(p.length===4)return 'a\u00f1o '+p;const[y,m]=p.split('-');return MESES[+m-1]+' '+y}\nfunction perStep(p,n){if(p.length===4)return String(+p+n);let[y,m]=p.split('-').map(Number);m+=n;while(m>12){m-=12;y++}while(m<1){m+=12;y--}return y+'-'+p2(m)}\nfunction fechaLarga(f){const d=new Date(f+'T12:00');return DIAS[d.getDay()]+' '+d.getDate()+' de '+MESES[d.getMonth()]+(f.slice(0,4)!==hoy().slice(0,4)?' '+f.slice(0,4):'')}\nfunction toast(msg){const t=document.getElementById('toast');t.textContent=msg;t.classList.add('on');clearTimeout(toast.t);toast.t=setTimeout(()=>t.classList.remove('on'),3000)}\nfunction lsGet(k){try{return localStorage.getItem(k)}catch(e){return null}}\nfunction lsSet(k,v){try{localStorage.setItem(k,v)}catch(e){}}\n\n/* ===================== Datos derivados ===================== */\nlet _movs=null,_acc=null,_con=null;\nfunction invalidate(){_movs=null;_acc=null;_con=null}\nfunction movs(){if(!_movs){_movs=[...S.movs].sort((x,y)=>x.f<y.f?1:x.f>y.f?-1:((y.ts||0)-(x.ts||0)))}return _movs}\nfunction acc(n){if(!_acc){_acc={};for(const c of S.cfg.cuentas)_acc[c.n]=c}return _acc[n]}\nfunction con(t,n){if(!_con){_con={};for(const c of S.cfg.conceptos)_con[c.tipo+'|'+c.n]=c}return _con[t+'|'+n]}\nfunction catOf(x){const c=con(x.t,x.k);return (c&&c.cat)||'Sin categor\u00eda'}\nfunction monOf(n){const a=acc(n);return a?a.mon:'UYU'}\nfunction cot(){return +S.cfg.cot||1}\nfunction toUYU(v,mon){return mon==='USD'?v*cot():v}\nfunction saldos(hasta){\n  hasta=hasta||hoy();const r={};\n  for(const c of S.cfg.cuentas)r[c.n]={s:+c.si||0,fut:0};\n  const add=(n,v,f)=>{const o=r[n];if(!o)return;if(f<=hasta)o.s+=v;else o.fut+=v};\n  for(const x of S.movs){\n    if(x.t==='T'){add(x.c,-x.m,x.f);add(x.c2,x.m2??x.m,x.f)}\n    else add(x.c,x.t==='G'?-x.m:x.m,x.f);\n  }\n  return r;\n}\nfunction efecto(x,n){\n  if(x.t==='T'){let v=0;if(x.c===n)v-=x.m;if(x.c2===n)v+=(x.m2??x.m);return v}\n  return x.c===n?(x.t==='G'?-x.m:x.m):0;\n}\nfunction titulo(x){if(x.t==='T')return 'Transferencia';if(x.t==='A')return 'Ajuste de saldo';return x.k||'Sin concepto'}\nfunction enPer(x,p){return !p||x.f.startsWith(p)}\nfunction sumGI(list){let g=0,i=0;for(const x of list){if(x.t==='G')g+=toUYU(x.m,monOf(x.c));else if(x.t==='I')i+=toUYU(x.m,monOf(x.c))}return{g,i}}\nfunction sortN(a){return [...a].sort((x,y)=>x.n.localeCompare(y.n,'es'))}\n\n/* ===================== Servidor (Google Sheet) ===================== */\nlet chain=Promise.resolve(),pend=0;\nfunction call(fn,...a){return new Promise((res,rej)=>google.script.run.withSuccessHandler(res).withFailureHandler(rej).rpc(fn,JSON.stringify(a)))}\nfunction setSync(){\n  const el=document.getElementById('sync');\n  el.textContent=pend?'Guardando\u2026':S.ready?'\u2713 Guardado en Google Sheets \u27f3':'Conectando\u2026';\n}\nfunction op(fn,...a){\n  pend++;setSync();\n  const p=chain.then(()=>call(fn,...a)).then(r=>{pend--;setSync();return r},e=>{pend--;setSync();toast('No se pudo guardar en el Sheet: '+(e&&e.message||e));load();throw e});\n  chain=p.catch(()=>{});return p;\n}\nfunction saveCfg(ren){return op('saveCfg',JSON.stringify({cot:S.cfg.cot,cuentas:S.cfg.cuentas,conceptos:S.cfg.conceptos,ren:ren||{}}))}\nlet _lento=null;\nasync function load(){\n  if(pend)return;\n  clearTimeout(_lento);\n  if(!S.ready)_lento=setTimeout(()=>{if(!S.ready){document.getElementById('view').innerHTML=`<div class=\"card welcome\"><h3 style=\"margin-top:0\">Est\u00e1 tardando m\u00e1s de lo normal</h3>\n    <p class=\"mu sm\">Puede ser la conexi\u00f3n o que Google est\u00e9 pidiendo un permiso. Prob\u00e1 recargar. Si sigue igual, en Apps Script abr\u00ed \u201cEjecuciones\u201d (men\u00fa de la izquierda) y mandame el error de <b>getData</b>.</p>\n    <button class=\"btn pri\" onclick=\"location.reload()\">Recargar</button></div>`}},25000);\n  try{\n    const o=JSON.parse(await call('getData'));\n    clearTimeout(_lento);\n    if(pend)return;\n    S.cfg=o.cfg;S.movs=o.movs;S.sheet=o.sheet||'';S.notifs=o.notifs||{pend:[],rec:[],cuotas:[]};S.notifs.cuotas=S.notifs.cuotas||[];S.reglas=o.reglas||[];S.token=o.token||'';S.fijos=o.fijos||[];S.difs=o.difs||[];S.extras=o.extras||[];S.app=o.app||{};S.url=o.url||'';S.avisos=o.avisos||{};\n    if(!load._extras){load._extras=true;cargarExtras()}S.ready=true;S.err=null;invalidate();render();\n  }catch(e){clearTimeout(_lento);S.err=(e&&e.message)||String(e);render()}\n}\nasync function cargarExtras(){\n  try{\n    const r=JSON.parse(await call('extrasCarga'));\n    if(r.cot){S.cfg.cot=r.cot;S.cfg.cotFecha=r.cotFecha}S.cfg.cotOk=r.cotOk;\n    if(r.app)S.app=r.app;\n    if(r.receptor)S.avisos.receptor=r.receptor;\n    S.avisos.gmailEstado=r.gmailEstado;\n    invalidate();if(document.getElementById('modal').hidden)render();\n    if(r.mails)load();\n  }catch(e){/* no es grave: se reintenta la pr\u00f3xima vez */}\n}\ndocument.addEventListener('visibilitychange',()=>{if(document.visibilityState==='visible'&&S.ready&&document.getElementById('modal').hidden)load()});\n\n/* ===================== Navegaci\u00f3n ===================== */\nconst ICONS={\n inicio:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M3 11l9-7 9 7v9a1 1 0 0 1-1 1h-5v-6H9v6H4a1 1 0 0 1-1-1z\"/></svg>',\n movimientos:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M8 6h13M8 12h13M8 18h13M3 6h.01M3 12h.01M3 18h.01\"/></svg>',\n cuentas:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><rect x=\"2\" y=\"5\" width=\"20\" height=\"14\" rx=\"2\"/><path d=\"M2 10h20\"/></svg>',\n informes:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M4 20V10M10 20V4M16 20v-7M22 20H2\"/></svg>',\n clasificar:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M6 8a6 6 0 0 1 12 0c0 7 3 9 3 9H3s3-2 3-9M10.3 21a1.9 1.9 0 0 0 3.4 0\"/></svg>',\n presupuesto:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><circle cx=\"12\" cy=\"12\" r=\"9\"/><path d=\"M12 3v9l6.4 6.4\"/></svg>',\n fijos:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><rect x=\"3\" y=\"4\" width=\"18\" height=\"17\" rx=\"2\"/><path d=\"M3 9h18M8 2v4M16 2v4M8 14l2.5 2.5L16 12\"/></svg>',\n mas:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><circle cx=\"5\" cy=\"12\" r=\"1.5\"/><circle cx=\"12\" cy=\"12\" r=\"1.5\"/><circle cx=\"19\" cy=\"12\" r=\"1.5\"/></svg>',\n ajustes:'<svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><circle cx=\"12\" cy=\"12\" r=\"3\"/><path d=\"M12 2v3M12 19v3M4.2 4.2l2.1 2.1M17.7 17.7l2.1 2.1M2 12h3M19 12h3M4.2 19.8l2.1-2.1M17.7 6.3l2.1-2.1\"/></svg>'\n};\nconst TABS=[['inicio','Inicio'],['movimientos','Movimientos'],['clasificar','Por clasificar'],['presupuesto','Presupuesto'],['cuentas','Cuentas'],['informes','Informes'],['fijos','Gastos fijos'],['ajustes','Ajustes']];\nconst TABS_CEL=[['inicio','Inicio'],['movimientos','Movimientos'],['clasificar','Clasificar'],['presupuesto','Presupuesto'],['mas','M\u00e1s']];\nfunction buildNav(){\n  document.getElementById('tabs').innerHTML=TABS_CEL.map(([k,l])=>`<button data-t=\"${k}\" onclick=\"go('${k}')\"><span class=\"ib\">${ICONS[k]}${k==='clasificar'?'<em class=\"badge\" hidden></em>':''}</span><span>${l}</span></button>`).join('');\n  document.getElementById('side').innerHTML='<div class=\"brand\">\ud83d\udcb0 Control de Gastos</div><button class=\"newbtn\" onclick=\"openMov()\">+ Nuevo movimiento</button>'+\n    TABS.map(([k,l])=>`<button class=\"nv\" data-t=\"${k}\" onclick=\"go('${k}')\"><span class=\"ib\">${ICONS[k]}${k==='clasificar'?'<em class=\"badge\" hidden></em>':''}</span><span>${l}</span></button>`).join('');\n}\nfunction go(t){V.tab=t;window.scrollTo(0,0);render()}\n\n/* ===================== Render ===================== */\nfunction render(){\n  setSync();\n  const enMas=['cuentas','informes','fijos','ajustes'].includes(V.tab);\n  document.querySelectorAll('[data-t]').forEach(b=>b.classList.toggle('on',b.dataset.t===V.tab||(b.dataset.t==='mas'&&enMas&&b.closest('#tabs'))));\n  const np=S.notifs.pend.length+S.notifs.cuotas.length+S.difs.length;document.querySelectorAll('.badge').forEach(b=>{b.hidden=!np;b.textContent=np>99?'99+':np});\n  const view=document.getElementById('view'),title=document.getElementById('title');\n  const vacio=S.cfg&&!S.cfg.cuentas.length&&!S.cfg.conceptos.length&&!S.movs.length;\n  document.getElementById('fab').style.visibility=S.ready&&!vacio?'visible':'hidden';\n  if(S.err&&!S.ready){view.innerHTML=`<div class=\"card welcome\"><h3 style=\"margin-top:0\">No pude leer la planilla</h3><p class=\"mu sm\">${h(S.err)}</p><button class=\"btn pri\" onclick=\"S.err=null;render();load()\">Reintentar</button></div>`;return}\n  if(!S.ready){view.innerHTML='<div class=\"loading\">Cargando tus datos del Google Sheet\u2026</div>';return}\n  if(vacio){title.textContent='Bienvenido';view.innerHTML=welcome();return}\n  const T={inicio:['Inicio',vInicio],movimientos:['Movimientos',vMovs],cuentas:['Cuentas',vCuentas],informes:['Informes',vInformes],ajustes:['Ajustes',vAjustes],clasificar:['Por clasificar',vClasificar],presupuesto:['Presupuesto',vPresupuesto],fijos:['Gastos fijos',vFijos],mas:['M\u00e1s',vMas]}[V.tab];\n  title.textContent=T[0];view.innerHTML=T[1]();\n  if(V.tab==='movimientos')renderList();\n}\nfunction welcome(){\n  return `<div class=\"welcome card\"><h3 style=\"margin-top:0\">Empecemos</h3>\n  <p class=\"mu\">Tus datos se guardan en tu Google Sheet (pesta\u00f1as Movimientos, Cuentas y Conceptos). Lo que cargues se ve igual en el celular y en la PC.</p>\n  <div class=\"btns\" style=\"margin-top:14px\">\n   <button class=\"btn pri\" onclick=\"document.getElementById('fileIn').click()\">Importar desde ZetaCuentas</button>\n   <button class=\"btn\" onclick=\"startBasic()\">Empezar de cero</button>\n  </div>\n  <p class=\"xs mu\" style=\"margin-top:14px\">Sirve el Excel <b>Suma de ingresos y egresos</b> o el <b>listado de comprobantes</b> de ZetaCuentas, o cualquier Excel/CSV con fecha, cuenta e importe.</p></div>`;\n}\nconst CONC_BASE=[['V\u00edveres','Alimentaci\u00f3n'],['Delivery','Alimentaci\u00f3n'],['Auto','Transporte, locomoci\u00f3n y veh\u00edculos'],['ANTEL','Servicios y tarifas'],['UTE','Servicios y tarifas'],['OSE','Servicios y tarifas'],['Gas','Servicios y tarifas'],['Streaming','Servicios y tarifas'],['Impuestos','Servicios y tarifas'],['Salud','Salud y gastos m\u00e9dicos'],['Vestimenta','Vestimenta'],['Casa','Vivienda, mejoras y arreglos'],['Mascotas','Vivienda, mejoras y arreglos'],['Esparcimiento','Esparcimiento'],['Vacaciones','Esparcimiento'],['Regalos','Regalos, donaciones y aportes'],['Gastos sin categor\u00eda','Gastos varios']];\nfunction startBasic(){\n  S.cfg.cuentas=[{n:'Efectivo',mon:'UYU',si:0},{n:'Efectivo d\u00f3lares',mon:'USD',si:0}];\n  S.cfg.conceptos=[...CONC_BASE.map(([n,cat])=>({n,tipo:'G',cat})),{n:'Sueldo',tipo:'I',cat:'Sueldos'},{n:'Otros ingresos',tipo:'I',cat:'Otros ingresos'}];\n  invalidate();saveCfg();render();\n}\n\n/* ---------- Inicio ---------- */\nfunction vInicio(){\n  const sal=saldos(),hy=hoy(),mes=ym(hy);\n  let tU=0,tD=0,fut=0;\n  const cs=S.cfg.cuentas.filter(c=>!c.arch);\n  for(const c of S.cfg.cuentas){const s=sal[c.n];if(c.mon==='USD')tD+=s.s;else tU+=s.s;fut+=toUYU(s.fut,c.mon)}\n  const ms=movs().filter(x=>x.f.startsWith(mes));const{g,i}=sumGI(ms);\n  const by={};for(const x of ms)if(x.t==='G'){const c=catOf(x);by[c]=(by[c]||0)+toUYU(x.m,monOf(x.c))}\n  const top=Object.entries(by).sort((a,b)=>b[1]-a[1]).slice(0,7);const mx=top.length?top[0][1]:1;\n  const ult=movs().filter(x=>x.f<=hy).slice(0,8);\n  const np=S.notifs.pend.length+S.notifs.cuotas.length+S.difs.length;\n  const pr=presResumen(mes);\n  return `${S.app&&S.app.hay?`<div class=\"card banner\" onclick=\"go('ajustes')\"><b>\u2b06\ufe0f Hay una versi\u00f3n nueva de la app (${h(S.app.remota)})</b><span>${S.app.auto?'Se instala sola esta noche \u203a':'Actualizar \u203a'}</span></div>`:''}${np?`<div class=\"card banner\" onclick=\"go('clasificar')\"><b>\ud83d\udd14 Ten\u00e9s ${np} pendiente${np>1?'s':''} para revisar</b><span>Ver \u203a</span></div>`:''}\n  ${pr.pres?`<div class=\"card\" style=\"cursor:pointer\" onclick=\"go('presupuesto')\"><h2>Presupuesto del mes<span class=\"num\" style=\"text-transform:none;font-size:20px\">${pctTxt(pr.gas,pr.pres)}</span></h2>\n    <div class=\"sm\">Gastado <b class=\"num\">${fmt0(pr.gas)}</b> de <b class=\"num\">${fmt0(pr.pres)}</b> \u00b7 disponible <b class=\"num ${pr.pres-pr.gas<0?'neg':''}\">${fmt0(pr.pres-pr.gas)}</b></div></div>`:''}\n  <div class=\"grid2\">\n  <div class=\"card\"><h2>Saldo total</h2>\n    <div class=\"big num ${tU+toUYU(tD,'USD')<0?'neg':''}\">${fmtS(tU+toUYU(tD,'USD'),'UYU')}</div>\n    <div class=\"kv sm\"><span class=\"mu\">En pesos</span><span class=\"num\">${fmtS(tU,'UYU')}</span></div>\n    <div class=\"kv sm\"><span class=\"mu\">En d\u00f3lares</span><span class=\"num\">${fmtS(tD,'USD')}</span></div>\n    <div class=\"xs mu\">D\u00f3lar eBROU venta: $ ${NF.format(cot())}${S.cfg.cotFecha?' \u00b7 '+h(S.cfg.cotFecha):''}</div>\n    ${Math.abs(fut)>0.005?`<div class=\"kv sm\" style=\"margin-top:6px\"><span class=\"mu\">Movimientos futuros (cuotas)</span><span class=\"num ${fut<0?'neg':'pos'}\">${fmtS(fut,'UYU')}</span></div>`:''}\n  </div>\n  <div class=\"card\"><h2>${h(perLabel(mes))}</h2>\n    <div class=\"kv\"><span>Ingresos</span><span class=\"num pos\">${fmt(i,'UYU')}</span></div>\n    <div class=\"kv\"><span>Gastos</span><span class=\"num neg\">${fmt(g,'UYU')}</span></div>\n    <div class=\"kv\" style=\"border-top:1px solid var(--ln);margin-top:4px;padding-top:8px\"><b>Balance</b><b class=\"num ${i-g<0?'neg':'pos'}\">${fmtS(i-g,'UYU')}</b></div>\n  </div>\n  <div class=\"card\"><h2>Cuentas <span class=\"btns\"><button class=\"btn sm\" onclick=\"fijarSaldos()\">Fijar saldos</button><button class=\"btn sm\" onclick=\"go('cuentas')\">Ver todas</button></span></h2>\n    ${cs.length?cs.map(c=>`<div class=\"row\" onclick=\"verCuenta(${js(c.n)})\"><div class=\"mid\"><div class=\"t1\">${h(c.n)}</div></div><div class=\"num ${sal[c.n].s<0?'neg':''}\">${fmtS(sal[c.n].s,c.mon)}</div></div>`).join(''):'<div class=\"empty\">Sin cuentas</div>'}\n  </div>\n  <div class=\"card\"><h2>Gastos del mes por categor\u00eda <button class=\"btn sm\" onclick=\"V.inf.per='${mes}';V.inf.tipo='G';go('informes')\">Informe</button></h2>\n    ${top.length?top.map(([k,v])=>`<div style=\"padding:5px 0;cursor:pointer\" onclick=\"V.inf.per='${mes}';V.inf.tipo='G';V.inf.by='k';V.inf.cat=${js(k)};go('informes')\"><div class=\"kv sm\" style=\"padding:0\"><span>${h(k)}</span><span class=\"num\">${fmt(v,'UYU')}</span></div><div class=\"bar G\"><i style=\"width:${(v/mx*100).toFixed(1)}%\"></i></div></div>`).join(''):'<div class=\"empty\">Todav\u00eda no hay gastos este mes</div>'}\n  </div>\n  </div>\n  <div class=\"card\" style=\"margin-top:12px\"><h2>\u00daltimos movimientos <button class=\"btn sm\" onclick=\"go('movimientos')\">Ver todos</button></h2>\n    ${ult.length?ult.map(x=>rowHTML(x,null,true)).join(''):'<div class=\"empty\">Toc\u00e1 + para cargar tu primer movimiento</div>'}\n  </div>`;\n}\nfunction rowHTML(x,an,showDate){\n  const sub=[];\n  if(showDate)sub.push(x.f.slice(8)+' '+MES3[+x.f.slice(5,7)-1]);\n  sub.push(x.t==='T'?(h(x.c)+' \u2192 '+h(x.c2||'?')):h(x.c));\n  if(x.e)sub.push('#'+h(x.e));\n  if(x.d)sub.push(h(x.d));\n  let v,mon;\n  if(an){v=efecto(x,an);mon=monOf(an)}\n  else if(x.t==='T'){v=x.m;mon=monOf(x.c)}\n  else{v=x.t==='G'?-x.m:x.m;mon=monOf(x.c)}\n  const cls=x.t==='T'?'tr':x.t==='A'?'aj':v<0?'neg':'pos';\n  const letra=x.t==='T'?'\u21c4':x.t==='A'?'\u00b1':titulo(x).charAt(0).toUpperCase();\n  const fut=(x.q?'<span class=\"chip\">cuota '+h(x.q)+'</span>':'')+(x.f>hoy()?'<span class=\"chip\">futuro</span>':'')+(x.cc?'<span class=\"chip\" title=\"Conciliado con el estado de cuenta\">\u2713</span>':'');\n  return `<div class=\"row\" onclick=\"openMov(${js(x.id)})\"><div class=\"ico ${x.t}\">${h(letra)}</div>\n   <div class=\"mid\"><div class=\"t1\">${h(titulo(x))}${fut}</div><div class=\"t2\">${sub.join(' \u00b7 ')}</div></div>\n   <div class=\"num ${cls}\">${x.t==='T'&&!an?fmt(v,mon):fmtS(v,mon)}</div></div>`;\n}\n\n/* ---------- Movimientos ---------- */\nfunction filtrar(f,per){V.f={c:'',k:'',t:'',q:'',e:'',cat:'',...f};V.per=per??V.per;V.show=300;go('movimientos')}\nfunction verCuenta(n){filtrar({c:n},V.per)}\nfunction vMovs(){\n  const opts=(arr,sel)=>arr.map(([v,l])=>`<option value=\"${h(v)}\"${v===sel?' selected':''}>${h(l)}</option>`).join('');\n  const cs=[['','Todas las cuentas'],...S.cfg.cuentas.filter(c=>!c.arch||c.n===V.f.c).map(c=>[c.n,c.n])];\n  const ks=[['','Todos los conceptos'],...sortN(S.cfg.conceptos).map(c=>[c.tipo+'|'+c.n,c.n+(c.tipo==='I'?' (ingreso)':'')])];\n  const ts=[['','Todos los tipos'],['G','Gastos'],['I','Ingresos'],['T','Transferencias'],['A','Ajustes']];\n  return `<div class=\"card\">\n   <div style=\"display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px\">\n    ${pnav('V.per',V.per)}\n    <div class=\"seg\"><button class=\"${V.per.length===7?'on':''}\" onclick=\"V.per=ym(hoy());render()\">Mes</button><button class=\"${V.per.length===4?'on':''}\" onclick=\"V.per=hoy().slice(0,4);render()\">A\u00f1o</button><button class=\"${!V.per?'on':''}\" onclick=\"V.per='';render()\">Todo</button></div>\n   </div>\n   <div class=\"filters\">\n    <select onchange=\"V.f.c=this.value;renderList()\">${opts(cs,V.f.c)}</select>\n    <select onchange=\"V.f.k=this.value;renderList()\">${opts(ks,V.f.k)}</select>\n    <select onchange=\"V.f.t=this.value;renderList()\">${opts(ts,V.f.t)}</select>\n    <input type=\"search\" placeholder=\"Buscar descripci\u00f3n o etiqueta\u2026\" value=\"${h(V.f.q)}\" oninput=\"V.f.q=this.value;clearTimeout(window._qt);window._qt=setTimeout(renderList,200)\">\n   </div>\n   ${V.f.e||V.f.cat?`<div class=\"sm\" style=\"margin-top:8px\">${V.f.e?'Etiqueta: <b>#'+h(V.f.e)+'</b> ':''}${V.f.cat?'Categor\u00eda: <b>'+h(V.f.cat)+'</b> ':''}<button class=\"btn sm\" onclick=\"V.f.e='';V.f.cat='';render()\">Quitar</button></div>`:''}\n   <div class=\"totals\" id=\"totals\"></div>\n  </div>\n  <div class=\"card\" id=\"list\"></div>`;\n}\nfunction pnav(varName,p){\n  if(!p)return '<div class=\"pnav\"><b>Todo el per\u00edodo</b></div>';\n  return `<div class=\"pnav\"><button onclick=\"${varName}=perStep(${varName},-1);render()\">\u2039</button><b>${h(perLabel(p))}</b><button onclick=\"${varName}=perStep(${varName},1);render()\">\u203a</button></div>`;\n}\nfunction filtered(){\n  const f=V.f,q=norm(f.q);\n  return movs().filter(x=>enPer(x,V.per)&&(!f.c||x.c===f.c||x.c2===f.c)&&(!f.k||(x.t+'|'+x.k)===f.k)&&(!f.t||x.t===f.t)&&(!f.e||x.e===f.e)&&\n    (!f.cat||((x.t==='G'||x.t==='I')&&catOf(x)===f.cat))&&\n    (!q||norm((x.d||'')+' '+(x.e||'')+' '+titulo(x)+' '+x.c+' '+(x.c2||'')).includes(q)||String(x.m).includes(q)));\n}\nfunction renderList(){\n  const L=document.getElementById('list');if(!L)return;\n  const list=filtered(),an=V.f.c||null;\n  let tt;\n  if(an){const mon=monOf(an);let ing=0,egr=0;for(const x of list){const v=efecto(x,an);if(v>0)ing+=v;else egr-=v}\n    const s=saldos()[an]||{s:0};\n    tt=`<span>Entradas <b class=\"num pos\">${fmt(ing,mon)}</b></span><span>Salidas <b class=\"num neg\">${fmt(egr,mon)}</b></span><span>Saldo hoy <b class=\"num\">${fmtS(s.s,mon)}</b></span>\n    <button class=\"btn sm\" onclick=\"ajustarSaldo(${js(an)})\">Ajustar saldo</button>`;\n  }else{const{g,i}=sumGI(list);tt=`<span>Ingresos <b class=\"num pos\">${fmt(i,'UYU')}</b></span><span>Gastos <b class=\"num neg\">${fmt(g,'UYU')}</b></span><span>Balance <b class=\"num ${i-g<0?'neg':'pos'}\">${fmtS(i-g,'UYU')}</b></span><span class=\"mu\">${list.length} mov.</span>`}\n  document.getElementById('totals').innerHTML=tt;\n  if(!list.length){L.innerHTML='<div class=\"empty\">No hay movimientos con estos filtros</div>';return}\n  const dayTot={};for(const y of list){let v;if(an)v=efecto(y,an);else if(y.t==='G')v=-toUYU(y.m,monOf(y.c));else if(y.t==='I')v=toUYU(y.m,monOf(y.c));else v=0;dayTot[y.f]=(dayTot[y.f]||0)+v}\n  let out='',day='',n=0;\n  for(const x of list){\n    if(n>=V.show)break;n++;\n    if(x.f!==day){day=x.f;const dv=dayTot[day];\n      out+=`<div class=\"day\"><span>${fechaLarga(day)}</span><span class=\"num\">${dv?fmtS(dv,an?monOf(an):'UYU'):''}</span></div>`}\n    out+=rowHTML(x,an);\n  }\n  if(list.length>V.show)out+=`<div style=\"text-align:center;padding:10px\"><button class=\"btn\" onclick=\"V.show+=300;renderList()\">Mostrar m\u00e1s (${list.length-V.show} restantes)</button></div>`;\n  L.innerHTML=out;\n}\n\n/* ---------- Cuentas ---------- */\nfunction vCuentas(){\n  const sal=saldos();\n  const grp=mon=>{\n    const cs=S.cfg.cuentas.filter(c=>c.mon===mon&&(V.verArch||!c.arch));\n    if(!cs.length)return '';\n    let t=0;for(const c of cs)t+=sal[c.n].s;\n    return `<div class=\"card\"><h2>${mon==='USD'?'D\u00f3lares':'Pesos'}<span class=\"num\" style=\"text-transform:none;color:var(--tx)\">${fmtS(t,mon)}</span></h2>\n      ${cs.map(c=>`<div class=\"row\" onclick=\"verCuenta(${js(c.n)})\"><div class=\"mid\"><div class=\"t1\">${h(c.n)}${c.arch?'<span class=\"chip\">archivada</span>':''}</div>\n       ${Math.abs(sal[c.n].fut)>0.005?`<div class=\"t2\">Futuro: ${fmtS(sal[c.n].fut,mon)}</div>`:''}</div>\n       <div class=\"num ${sal[c.n].s<0?'neg':''}\">${fmtS(sal[c.n].s,mon)}</div>\n       <button class=\"btn sm\" onclick=\"event.stopPropagation();openCuenta(${js(c.n)})\">Editar</button></div>`).join('')}</div>`;\n  };\n  return `<div class=\"btns\" style=\"margin-bottom:12px\"><button class=\"btn pri\" onclick=\"fijarSaldos()\">Fijar saldos de hoy</button><button class=\"btn\" onclick=\"openCuenta()\">+ Nueva cuenta</button>\n   <button class=\"btn\" onclick=\"cargarPDF()\">\ud83d\udcc4 Conciliar estado de cuenta</button>\n   <button class=\"btn\" onclick=\"V.verArch=!V.verArch;render()\">${V.verArch?'Ocultar':'Mostrar'} archivadas</button></div>\n   ${grp('UYU')}${grp('USD')}\n   ${!S.cfg.cuentas.length?'<div class=\"card empty\">Cre\u00e1 tu primera cuenta (efectivo, caja de ahorro, tarjeta\u2026)</div>':''}`;\n}\nlet _editAcc=null;\nfunction openCuenta(n){\n  _editAcc=n||null;\n  const c=n?acc(n):{n:'',mon:'UYU',si:0,arch:false};\n  modal(`<h3>${n?'Editar cuenta':'Nueva cuenta'}</h3>\n   <form onsubmit=\"event.preventDefault();saveCuenta()\">\n   <div class=\"fld\"><label>Nombre</label><input type=\"text\" id=\"cu_n\" value=\"${h(c.n)}\" required></div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Moneda</label><select id=\"cu_m\"><option value=\"UYU\"${c.mon==='UYU'?' selected':''}>Pesos ($)</option><option value=\"USD\"${c.mon==='USD'?' selected':''}>D\u00f3lares (U$S)</option></select></div>\n   <div class=\"fld\"><label>Saldo de partida</label><input type=\"text\" inputmode=\"decimal\" id=\"cu_s\" value=\"${c.si?String(c.si).replace('.',','):''}\" placeholder=\"0\"></div></div>\n   <div class=\"fld\"><label><input type=\"checkbox\" id=\"cu_a\"${c.arch?' checked':''}> Archivada (no se muestra en listas)</label></div>\n   <div class=\"btns\" style=\"justify-content:space-between\"><div>${n?`<button type=\"button\" class=\"btn dan\" onclick=\"delCuenta()\">Eliminar</button>`:''}</div>\n   <div class=\"btns\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\">Guardar</button></div></div></form>`);\n  setTimeout(()=>document.getElementById('cu_n').focus(),50);\n}\nfunction saveCuenta(){\n  const old=_editAcc;\n  const n=document.getElementById('cu_n').value.trim();if(!n)return;\n  if(n!==old&&acc(n)){toast('Ya existe una cuenta con ese nombre');return}\n  const o={n,mon:document.getElementById('cu_m').value,si:r2(parseNum(document.getElementById('cu_s').value)||0),arch:document.getElementById('cu_a').checked};\n  const ren={};\n  if(old){Object.assign(acc(old),o);if(old!==n){ren.c={[old]:n};for(const x of S.movs){if(x.c===old)x.c=n;if(x.c2===old)x.c2=n}}}\n  else S.cfg.cuentas.push(o);\n  invalidate();saveCfg(ren);closeModal();render();toast('Cuenta guardada');\n}\nfunction delCuenta(){\n  const n=_editAcc,cant=S.movs.filter(x=>x.c===n||x.c2===n).length;\n  if(cant){toast(`Tiene ${cant} movimientos: archivala en lugar de eliminarla`);return}\n  confirmar('\u00bfEliminar esta cuenta?',()=>{S.cfg.cuentas=S.cfg.cuentas.filter(c=>c.n!==n);invalidate();saveCfg();closeModal();render()});\n}\nfunction ajustarSaldo(an){\n  const a=acc(an),f=hoy(),s=saldos(f)[an].s;\n  modal(`<h3>Ajustar saldo</h3><p class=\"mu sm\">${h(a.n)} \u00b7 saldo seg\u00fan la app: <b class=\"num\">${fmtS(s,a.mon)}</b></p>\n  <form onsubmit=\"event.preventDefault();doAjuste(${js(an)})\">\n  <div class=\"fld\"><label>Saldo real</label><input type=\"text\" inputmode=\"decimal\" class=\"amount\" id=\"aj_v\" required></div>\n  <div class=\"fld\"><label>Fecha</label><input type=\"date\" id=\"aj_f\" value=\"${f}\"></div>\n  <div class=\"btns\" style=\"justify-content:flex-end\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\">Ajustar</button></div></form>`);\n  setTimeout(()=>document.getElementById('aj_v').focus(),50);\n}\nfunction doAjuste(an){\n  const real=parseNum(document.getElementById('aj_v').value);if(isNaN(real))return;\n  const f=document.getElementById('aj_f').value||hoy();\n  const d=r2(real-saldos(f)[an].s);\n  if(Math.abs(d)<0.005){closeModal();toast('El saldo ya coincide');return}\n  const x={id:uid(),t:'A',f,c:an,m:d,ts:Date.now(),d:'Ajuste de saldo'};\n  S.movs.push(x);invalidate();op('saveMov',JSON.stringify(x));\n  closeModal();render();toast('Saldo ajustado');\n}\n\n/* ---------- Fijar saldos de hoy ---------- */\nfunction esTC(n){return /^tc\\b|tarjeta|cr[e\u00e9]dito|visa|master|\\boca\\b/i.test(norm(n))}\nfunction fijarSaldos(){\n  const sal=saldos(),f=hoy();\n  const cs=S.cfg.cuentas.filter(c=>!c.arch);\n  const fila=c=>{const tc=esTC(c.n),s=sal[c.n].s,fut=sal[c.n].fut;\n    return `<tr><td>${h(c.n)}<div class=\"xs mu\">${tc?'Deuda seg\u00fan la app':'Saldo seg\u00fan la app'}: <span class=\"num\">${tc?fmt(Math.max(0,-s),c.mon)+(s>0.004?' (a favor '+fmt(s,c.mon)+')':''):fmtS(s,c.mon)}</span>${Math.abs(fut)>0.004?` \u00b7 cuotas futuras: <span class=\"num\">${fmt(Math.abs(fut),c.mon)}</span>`:''}</div></td>\n      <td style=\"text-align:right\"><input type=\"text\" inputmode=\"decimal\" data-acc=\"${h(c.n)}\" data-tc=\"${tc?1:0}\" placeholder=\"${tc?'deuda':'saldo'}\" style=\"width:130px;text-align:right\"><div class=\"xs mu\">${sym(c.mon)}</div></td></tr>`};\n  modal(`<h3>Fijar saldos de hoy</h3>\n   <p class=\"sm mu\" style=\"margin-top:0\">Escrib\u00ed lo que ten\u00e9s <b>hoy</b> en cada cuenta. En las tarjetas pon\u00e9 <b>lo que deb\u00e9s hoy</b> (lo que todav\u00eda no pagaste, sin las cuotas futuras). Dej\u00e1 vac\u00edo lo que no quieras cambiar. Tu historial no se modifica: la app recalcula el saldo de partida.</p>\n   <div style=\"max-height:56vh;overflow:auto\"><table class=\"pv fs\" style=\"font-size:13.5px\">\n    ${cs.filter(c=>!esTC(c.n)).length?`<tr><th colspan=\"2\">Cuentas bancarias y efectivo</th></tr>${cs.filter(c=>!esTC(c.n)).map(fila).join('')}`:''}\n    ${cs.filter(c=>esTC(c.n)).length?`<tr><th colspan=\"2\" style=\"padding-top:12px\">Tarjetas de cr\u00e9dito (deuda de hoy)</th></tr>${cs.filter(c=>esTC(c.n)).map(fila).join('')}`:''}\n   </table></div>\n   <p class=\"xs mu\">Las cuotas que vencen en meses pr\u00f3ximos quedan aparte, como \u201cfuturo\u201d, y se van sumando a la deuda cuando llega su mes.</p>\n   <div class=\"btns\" style=\"justify-content:flex-end\"><button class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\" onclick=\"guardarSaldos()\">Guardar saldos</button></div>`);\n}\nfunction guardarSaldos(){\n  const sal=saldos();let n=0;\n  document.querySelectorAll('input[data-acc]').forEach(i=>{\n    if(!i.value.trim())return;\n    let v=parseNum(i.value);if(isNaN(v))return;\n    if(i.dataset.tc==='1')v=-Math.abs(v);\n    const c=acc(i.dataset.acc);c.si=r2((+c.si||0)+(v-sal[c.n].s));n++;\n  });\n  if(!n){closeModal();return}\n  invalidate();saveCfg();closeModal();render();toast(`Saldos fijados en ${n} cuenta${n>1?'s':''}`);\n}\n\n/* ---------- Informes ---------- */\nfunction vInformes(){\n  const I=V.inf,p=I.per;\n  const list=movs().filter(x=>x.t===I.tipo&&enPer(x,p)&&(!I.cat||catOf(x)===I.cat));\n  const by={};let tot=0;\n  const keyOf=x=>I.by==='cat'?catOf(x):I.by==='k'?(x.k||''):I.by==='e'?(x.e||''):x.c;\n  for(const x of list){const key=keyOf(x);const v=toUYU(x.m,monOf(x.c));by[key]=(by[key]||0)+v;tot+=v}\n  const rows=Object.entries(by).sort((a,b)=>b[1]-a[1]);const mx=rows.length?rows[0][1]:1;\n  const name=k=>I.by==='e'?(k?'#'+k:'Sin etiqueta'):(k||'Sin concepto');\n  const click=k=>I.by==='cat'?`V.inf.by='k';V.inf.cat=${js(k)};render()`:I.by==='k'?`filtrar({k:${js(I.tipo+'|'+k)}},'${p}')`:I.by==='e'?`filtrar({e:${js(k)},t:'${I.tipo}'},'${p}')`:`filtrar({c:${js(k)},t:'${I.tipo}'},'${p}')`;\n  const end=p?(p.length===7?p:(p===hoy().slice(0,4)?ym(hoy()):p+'-12')):ym(hoy());\n  const ms=[];for(let n=11;n>=0;n--)ms.push(perStep(end,-n));\n  const ev=ms.map(m=>({m,...sumGI(movs().filter(x=>x.f.startsWith(m)))}));\n  const emx=Math.max(1,...ev.map(e=>Math.max(e.g,e.i)));\n  const segBy=[['cat','Categor\u00eda'],['k','Concepto'],['e','Etiqueta'],['c','Cuenta']].map(([k,l])=>`<button class=\"${I.by===k?'on':''}\" onclick=\"V.inf.by='${k}';V.inf.cat='';render()\">${l}</button>`).join('');\n  return `<div class=\"card\">\n   <div style=\"display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px\">\n    ${pnav('V.inf.per',p)}\n    <div class=\"seg\"><button class=\"${p.length===7?'on':''}\" onclick=\"V.inf.per=ym(hoy());render()\">Mes</button><button class=\"${p.length===4?'on':''}\" onclick=\"V.inf.per=hoy().slice(0,4);render()\">A\u00f1o</button><button class=\"${!p?'on':''}\" onclick=\"V.inf.per='';render()\">Todo</button></div>\n   </div>\n   <div class=\"filters\">\n    <div class=\"seg\" style=\"flex:1 1 200px\"><button class=\"${I.tipo==='G'?'on G':''}\" onclick=\"V.inf.tipo='G';V.inf.cat='';render()\">Gastos</button><button class=\"${I.tipo==='I'?'on I':''}\" onclick=\"V.inf.tipo='I';V.inf.cat='';render()\">Ingresos</button></div>\n    <div class=\"seg\" style=\"flex:1 1 300px\">${segBy}</div>\n   </div></div>\n  <div class=\"card\"><h2><span>${I.tipo==='G'?'Gastos':'Ingresos'} \u00b7 ${h(perLabel(p))}${I.cat?' \u00b7 '+h(I.cat):''}</span><span class=\"num\" style=\"color:var(--tx)\">${fmt(tot,'UYU')}</span></h2>\n   ${I.cat?`<div style=\"margin-bottom:6px\"><button class=\"btn sm\" onclick=\"V.inf.cat='';V.inf.by='cat';render()\">\u2039 Todas las categor\u00edas</button> <button class=\"btn sm\" onclick=\"filtrar({cat:${js(I.cat)},t:'${I.tipo}'},'${p}')\">Ver movimientos</button></div>`:''}\n   ${rows.length?rows.map(([k,v])=>`<div style=\"padding:6px 0;cursor:pointer\" onclick=\"${h(click(k))}\"><div class=\"kv sm\" style=\"padding:0\"><span>${h(name(k))}</span><span class=\"num\">${fmt(v,'UYU')} <span class=\"mu xs\">${(v/tot*100).toFixed(1)}%</span></span></div><div class=\"bar ${I.tipo}\"><i style=\"width:${(v/mx*100).toFixed(1)}%\"></i></div></div>`).join(''):'<div class=\"empty\">Sin datos para este per\u00edodo</div>'}\n   <div class=\"xs mu\" style=\"margin-top:8px\">Montos en pesos; d\u00f3lares convertidos a $ ${NF.format(cot())}.</div>\n  </div>\n  <div class=\"card\"><h2>12 meses</h2>\n   <div class=\"cols\">${ev.map(e=>`<div class=\"c\" title=\"${perLabel(e.m)}: ingresos ${fmt(e.i,'UYU')}, gastos ${fmt(e.g,'UYU')}\" onclick=\"V.inf.per='${e.m}';render()\"><div class=\"pair\"><i style=\"height:${(e.i/emx*100).toFixed(1)}%;background:var(--verde)\"></i><i style=\"height:${(e.g/emx*100).toFixed(1)}%;background:var(--rojo)\"></i></div><span>${MES3[+e.m.slice(5)-1]}</span></div>`).join('')}</div>\n   <div class=\"legend\"><span><i style=\"background:var(--verde)\"></i>Ingresos</span><span><i style=\"background:var(--rojo)\"></i>Gastos</span></div>\n   <div style=\"margin-top:10px\">${ev.slice().reverse().map(e=>`<div class=\"kv sm\"><span>${perLabel(e.m)}</span><span class=\"num\"><span class=\"pos\">${fmt(e.i,'UYU')}</span> \u00b7 <span class=\"neg\">${fmt(e.g,'UYU')}</span> \u00b7 <b class=\"${e.i-e.g<0?'neg':'pos'}\">${fmtS(e.i-e.g,'UYU')}</b></span></div>`).join('')}</div>\n  </div>`;\n}\n\n/* ---------- Ajustes ---------- */\nfunction vAjustes(){\n  const uso={};for(const x of S.movs)if(x.k){const k=x.t+'|'+x.k;uso[k]=(uso[k]||0)+1}\n  const lst=t=>{\n    const cats={};for(const c of S.cfg.conceptos.filter(c=>c.tipo===t)){(cats[c.cat||'Sin categor\u00eda']||(cats[c.cat||'Sin categor\u00eda']=[])).push(c)}\n    return Object.keys(cats).sort((a,b)=>a.localeCompare(b,'es')).map(cat=>`<div class=\"catH\">${h(cat)}</div>`+sortN(cats[cat]).map(c=>`<div class=\"row\" onclick=\"openConcepto(${js(c.tipo)},${js(c.n)})\"><div class=\"mid\"><div class=\"t1\">${h(c.n)}</div></div><span class=\"mu xs\">${uso[c.tipo+'|'+c.n]||0} mov.</span></div>`).join('')).join('');\n  };\n  return `<div class=\"grid2\">\n  <div class=\"card\"><h2>Saldos de las cuentas</h2>\n   <p class=\"sm mu\" style=\"margin-top:0\">Pon\u00e9 lo que ten\u00e9s hoy en cada cuenta y lo que deb\u00e9s hoy en cada tarjeta. La app ajusta el punto de partida sin tocar tu historial.</p>\n   <button class=\"btn pri\" onclick=\"fijarSaldos()\">Fijar saldos de hoy</button></div>\n  <div class=\"card\"><h2>Cotizaci\u00f3n del d\u00f3lar</h2>\n   <div class=\"big num\">$ ${NF.format(cot())}</div>\n   <div class=\"sm mu\">BROU \u00b7 D\u00f3lar eBROU \u00b7 venta${S.cfg.cotFecha?' \u00b7 actualizada '+h(S.cfg.cotFecha):''}</div>\n   ${S.cfg.cotOk===false?'<div class=\"xs aj\" style=\"margin-top:4px\">No se pudo leer la p\u00e1gina del BROU ahora; se muestra el \u00faltimo valor obtenido.</div>':''}\n   <div class=\"xs mu\" style=\"margin-top:6px\">Se toma sola del sitio del BROU (se actualiza cada 30 minutos) y se usa para pasar los d\u00f3lares a pesos en totales e informes.</div></div>\n  <div class=\"card\"><h2>Importar de ZetaCuentas</h2>\n   <p class=\"sm mu\" style=\"margin-top:0\">Excel <b>Suma de ingresos y egresos</b>, <b>listado de comprobantes</b> o cualquier Excel/CSV. Antes de importar pod\u00e9s elegir qu\u00e9 columna es cada dato. Importar el mismo archivo dos veces no duplica.</p>\n   <button class=\"btn pri\" onclick=\"document.getElementById('fileIn').click()\">Elegir archivo\u2026</button></div>\n  </div>\n  <div class=\"grid2\" style=\"margin-top:12px\">\n   <div class=\"card\"><h2>Conceptos de gasto <button class=\"btn sm\" onclick=\"openConcepto('G')\">+ Agregar</button></h2>${lst('G')||'<div class=\"empty\">\u2014</div>'}</div>\n   <div class=\"card\"><h2>Conceptos de ingreso <button class=\"btn sm\" onclick=\"openConcepto('I')\">+ Agregar</button></h2>${lst('I')||'<div class=\"empty\">\u2014</div>'}</div>\n  </div>\n  ${vNotifCfg()}\n  <div class=\"card\" style=\"margin-top:12px\"><h2>Datos</h2>\n   <div class=\"sm mu\" style=\"margin-bottom:10px\">${S.movs.length} movimientos \u00b7 ${S.cfg.cuentas.length} cuentas \u00b7 guardados en tu Google Sheet${S.sheet?` (<a href=\"${h(S.sheet)}\" target=\"_blank\" rel=\"noopener\">abrir planilla</a>)`:''}.</div>\n   <div class=\"btns\"><button class=\"btn\" onclick=\"load();toast('Datos actualizados')\">Actualizar desde la planilla</button><button class=\"btn dan\" onclick=\"borrarTodo()\">Borrar todos los datos</button></div>\n  </div>\n  ${vVersion()}`;\n}\nlet _editCon=null;\nfunction openConcepto(tipo,n){\n  _editCon=n?{tipo,n}:{tipo,n:null};\n  const c=n?con(tipo,n):{n:'',tipo,cat:''};\n  const cats=[...new Set(S.cfg.conceptos.map(x=>x.cat).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));\n  modal(`<h3>${n?'Editar concepto':'Nuevo concepto'} de ${tipo==='I'?'ingreso':'gasto'}</h3><form onsubmit=\"event.preventDefault();saveConcepto()\">\n  <div class=\"fld\"><label>Nombre</label><input type=\"text\" id=\"co_n\" value=\"${h(c.n)}\" required></div>\n  <div class=\"fld\"><label>Categor\u00eda</label><input type=\"text\" id=\"co_c\" list=\"dl_cats\" value=\"${h(c.cat||'')}\" autocomplete=\"off\"><datalist id=\"dl_cats\">${cats.map(x=>`<option value=\"${h(x)}\">`).join('')}</datalist></div>\n  ${n?`<div class=\"fld\"><label>Pasar sus movimientos a otro concepto y eliminar este</label><select id=\"co_mv\"><option value=\"\">\u2014 no \u2014</option>${sortN(S.cfg.conceptos.filter(x=>x.tipo===tipo&&x.n!==n)).map(x=>`<option value=\"${h(x.n)}\">${h(x.n)}</option>`).join('')}</select></div>`:''}\n  <div class=\"btns\" style=\"justify-content:flex-end\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\">Guardar</button></div></form>`);\n  setTimeout(()=>document.getElementById('co_n').focus(),50);\n}\nfunction saveConcepto(){\n  const {tipo,n:old}=_editCon;\n  const n=document.getElementById('co_n').value.trim(),cat=document.getElementById('co_c').value.trim();if(!n)return;\n  const mv=old?document.getElementById('co_mv').value:'';\n  const ren={};\n  if(!mv&&n!==old&&con(tipo,n)){toast('Ya existe ese concepto');return}\n  if(old){\n    const to=mv||n;\n    if(to!==old){ren.k={[tipo+'|'+old]:to};for(const x of S.movs)if(x.t===tipo&&x.k===old)x.k=to;(S.extras||[]).forEach(e=>{if(e.k===old&&(e.t||'G')===tipo)e.k=to})}\n    if(mv)S.cfg.conceptos=S.cfg.conceptos.filter(c=>!(c.tipo===tipo&&c.n===old));\n    else Object.assign(con(tipo,old),{n,cat});\n  }else S.cfg.conceptos.push({n,tipo,cat});\n  invalidate();saveCfg(ren);closeModal();render();toast('Concepto guardado');\n}\nfunction borrarTodo(){\n  confirmar('Se borran todas las cuentas, conceptos y movimientos del Google Sheet. \u00bfSeguro?',()=>{\n    S.movs=[];S.cfg.cuentas=[];S.cfg.conceptos=[];invalidate();op('borrarTodo');closeModal();render();toast('Datos borrados')});\n}\n\n/* ---------- Por clasificar ---------- */\nconst NF0=v=>NF.format(v||0);\nfunction vClasificar(){\n  const P=S.notifs.pend,R=S.notifs.rec;\n  const lbl={asignada:'<span class=\"chip\">asignada</span>',descartada:'<span class=\"chip\" style=\"background:var(--bg);color:var(--mu)\">descartada</span>',ignorada:'<span class=\"chip\" style=\"background:var(--bg);color:var(--mu)\">no es gasto</span>',duplicada:'<span class=\"chip\" style=\"background:var(--bg);color:var(--mu)\">ya lleg\u00f3 por otro medio</span>',prueba:'<span class=\"chip\">\u2713 prueba recibida</span>'};\n  const CQ=S.notifs.cuotas.filter(n=>S.movs.some(x=>x.id===n.mov)),D=S.difs;\n  const grupos={};D.forEach(d=>(grupos[d.c]||(grupos[d.c]=[])).push(d));\n  return `<div class=\"btns\" style=\"margin-bottom:12px\"><button class=\"btn\" onclick=\"cargarPDF()\">\ud83d\udcc4 Conciliar estado de cuenta (PDF)</button></div>\n  ${CQ.length?`<div class=\"card\"><h2>\u00bfEn cu\u00e1ntas cuotas? <span class=\"chip\">${CQ.length}</span></h2><p class=\"xs mu\" style=\"margin-top:0\">Compras que entraron solas por notificaci\u00f3n. Si fue en cuotas, se divide en un movimiento por mes.</p>${CQ.map(cardCuotas).join('')}</div>`:''}\n  ${D.length?`<div class=\"catH\" style=\"margin:4px 0 8px\">Diferencias con estados de cuenta (${D.length})</div>${D.map(cardDif).join('')}`:''}\n  ${P.length?`<div class=\"catH\" style=\"margin:4px 0 8px\">Notificaciones sin asignar (${P.length})</div>`+P.map(cardNotif).join(''):(!CQ.length&&!D.length?`<div class=\"card empty\">No hay nada por clasificar \ud83c\udf89<br><span class=\"xs\">Ac\u00e1 aparecen las notificaciones que la app no supo asignar, las compras para confirmar cuotas y las diferencias con los estados de cuenta.</span></div>`:'')}\n  ${P.length>1?`<div class=\"btns\" style=\"margin-bottom:12px\"><button class=\"btn\" onclick=\"reprocesar()\">Volver a aplicar las reglas a todas</button></div>`:''}\n  <div class=\"card\"><h2>\u00daltimas notificaciones recibidas</h2>\n   ${R.length?R.map(n=>`<div class=\"row\" style=\"cursor:default\"><div class=\"mid\"><div class=\"t1\" style=\"white-space:normal\">${h(n.comercio||n.titulo||sinCanal(n.app)||'Aviso')} ${lbl[n.estado]||''}</div>\n     <div class=\"t2\">${h(n.fecha)} \u00b7 ${canalIco(n.app)} ${h(sinCanal(n.app))}${n.estado==='asignada'?' \u00b7 '+h(n.concepto)+' \u00b7 '+h(n.cuenta):''} \u00b7 ${h(n.texto)}</div></div>\n     ${n.importe?`<div class=\"num\">${fmt(n.importe,n.mon)}</div>`:''}\n     ${n.estado==='ignorada'?`<button class=\"btn sm\" onclick=\"estadoNotif(${js(n.id)},'pendiente')\">Clasificar</button>`:''}${n.estado!=='asignada'?`<button class=\"btn sm\" title=\"Borrar\" onclick=\"borrarNotif(${js(n.id)})\">\ud83d\uddd1</button>`:''}</div>`).join(''):'<div class=\"empty\">Todav\u00eda no lleg\u00f3 ninguna notificaci\u00f3n. Configur\u00e1 el celular en Ajustes \u2192 Notificaciones del celular.</div>'}\n  </div>`;\n}\nfunction canalIco(a){const c=(String(a||'').match(/^(APP|SMS|MAIL)\\|/)||[])[1]||'APP';return c==='MAIL'?'\u2709\ufe0f':c==='SMS'?'\ud83d\udcac':'\ud83d\udd14'}\nfunction sinCanal(a){return String(a||'').replace(/^(APP|SMS|MAIL)\\|/,'')}\nfunction cardNotif(n){\n  const i=n.id,t=['G','I','T'].includes(n.tipo)?n.tipo:'G';\n  const act=S.cfg.cuentas.filter(c=>!c.arch);\n  const accO=(sel,blank)=>(blank?'<option value=\"\">\u2014 elegir cuenta \u2014</option>':'')+act.map(c=>`<option value=\"${h(c.n)}\"${c.n===sel?' selected':''}>${h(c.n)} (${sym(c.mon)})</option>`).join('');\n  const conO=tp=>{const cats={};for(const c of S.cfg.conceptos.filter(c=>c.tipo===tp))(cats[c.cat||'Sin categor\u00eda']||(cats[c.cat||'Sin categor\u00eda']=[])).push(c);\n    return '<option value=\"\">\u2014 elegir concepto \u2014</option>'+Object.keys(cats).sort((a,b)=>a.localeCompare(b,'es')).map(cat=>`<optgroup label=\"${h(cat)}\">${sortN(cats[cat]).map(c=>`<option value=\"${h(c.n)}\"${c.n===n.concepto?' selected':''}>${h(c.n)}</option>`).join('')}</optgroup>`).join('')};\n  const cuenta=n.cuenta||'';\n  return `<div class=\"card pc\" id=\"nc_${h(i)}\">\n   <h2><span style=\"text-transform:none\">${canalIco(n.app)} ${h(sinCanal(n.app)||'Aviso')} \u00b7 ${h(n.fecha)}</span></h2>\n   <div class=\"nt\">${h(n.texto)}</div>\n   <div class=\"seg\" style=\"margin-bottom:10px\">${[['G','Gasto'],['I','Ingreso'],['T','Transferencia']].map(([k,l])=>`<button type=\"button\" data-k=\"${k}\" class=\"${t===k?'on '+k:''}\" onclick=\"ncTipo(${js(i)},'${k}')\">${l}</button>`).join('')}</div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Importe</label><input type=\"text\" inputmode=\"decimal\" class=\"nc_m\" value=\"${n.importe?String(n.importe).replace('.',','):''}\"></div>\n    <div class=\"fld\"><label>Fecha</label><input type=\"date\" class=\"nc_f\" value=\"${h((n.fecha||hoy()).slice(0,10))}\"></div></div>\n   <div class=\"fld\"><label class=\"nc_cl\">${t==='T'?'Desde':'Cuenta'}</label><select class=\"nc_c\">${accO(cuenta,true)}</select></div>\n   <div class=\"fld nc_c2w\"${t==='T'?'':' hidden'}><label>Hacia</label><select class=\"nc_c2\">${accO('',true)}</select></div>\n   <div class=\"fld nc_kw\"${t==='T'?' hidden':''}><label>Concepto</label><select class=\"nc_k\" data-tipo=\"${t}\">${conO(t==='I'?'I':'G')}</select></div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Descripci\u00f3n</label><input type=\"text\" class=\"nc_d\" value=\"${h(n.comercio)}\"></div>\n    <div class=\"fld nc_nw\"${t==='G'?'':' hidden'}><label>Cuotas</label><input type=\"number\" class=\"nc_n\" min=\"1\" max=\"60\" value=\"${n.cuotas||1}\"></div></div>\n   <div class=\"fld\"><label><input type=\"checkbox\" class=\"nc_rk\" checked> Recordar: cuando diga <input type=\"text\" class=\"nc_pk\" value=\"${h(n.comercio)}\" style=\"width:auto;display:inline-block;padding:3px 6px\"> usar este concepto</label></div>\n   <div class=\"fld\"><label><input type=\"checkbox\" class=\"nc_rc\"${n.cuenta?'':' checked'}> Recordar la cuenta para ${n.digitos?`la tarjeta terminada en <b>${h(n.digitos)}</b>`:`los avisos de <b>${h(sinCanal(n.app)||'esta app')}</b>`}</label></div>\n   <div class=\"btns\" style=\"justify-content:space-between\"><button class=\"btn dan\" onclick=\"estadoNotif(${js(i)},'descartada')\">Descartar</button>\n    <button class=\"btn pri\" onclick=\"asignar(${js(i)})\">Guardar</button></div>\n  </div>`;\n}\nfunction ncTipo(id,t){\n  const el=document.getElementById('nc_'+id);\n  el.querySelectorAll('.seg button').forEach(b=>b.className=b.dataset.k===t?'on '+t:'');\n  el.querySelector('.nc_c2w').hidden=t!=='T';el.querySelector('.nc_kw').hidden=t==='T';el.querySelector('.nc_nw').hidden=t!=='G';\n  el.querySelector('.nc_cl').textContent=t==='T'?'Desde':'Cuenta';\n  const k=el.querySelector('.nc_k');\n  if(t!=='T'&&k.dataset.tipo!==t){const tmp=document.createElement('div');const n=S.notifs.pend.find(x=>x.id===id);tmp.innerHTML=cardNotif({...n,tipo:t,concepto:''});k.innerHTML=tmp.querySelector('.nc_k').innerHTML;k.dataset.tipo=t}\n  el.dataset.t=t;\n}\nasync function asignar(id){\n  const el=document.getElementById('nc_'+id),n=S.notifs.pend.find(x=>x.id===id);\n  const q=c=>el.querySelector(c);\n  const t=el.dataset.t||(['G','I','T'].includes(n.tipo)?n.tipo:'G');\n  const m=Math.abs(parseNum(q('.nc_m').value)),c=q('.nc_c').value,c2=q('.nc_c2').value,k=q('.nc_k').value,d=q('.nc_d').value.trim(),f=q('.nc_f').value||hoy();\n  if(!(m>0)){toast('Ingres\u00e1 el importe');return}\n  if(!c){toast('Eleg\u00ed la cuenta');return}\n  if(t==='T'&&(!c2||c2===c)){toast('Eleg\u00ed la cuenta de destino');return}\n  if(t!=='T'&&!k){toast('Eleg\u00ed el concepto');return}\n  const mov={id:'m'+id.slice(1),t,f,c,m:r2(m),ts:Date.now()};\n  if(t==='T'){mov.c2=c2;if(monOf(c)!==monOf(c2)){toast('Para transferencias entre monedas, cargala desde Movimientos');return}}else mov.k=k;\n  if(d)mov.d=d;\n  const reglas=[];\n  if(t!=='T'&&q('.nc_rk').checked&&q('.nc_pk').value.trim())reglas.push({tipo:'Concepto',palabra:q('.nc_pk').value.trim(),valor:k});\n  if(q('.nc_rc').checked)reglas.push({tipo:'Cuenta',palabra:n.digitos||'',app:n.digitos?'':sinCanal(n.app),valor:c});\n  S.movs=S.movs.filter(x=>x.id!==mov.id);\n  const N=t==='G'?Math.max(1,Math.min(60,parseInt(q('.nc_n').value)||1)):1;\n  if(N>1){const cs=crearCuotas(mov,N,mov.m);S.movs.push(...cs);op('importMovs',JSON.stringify(cs)).catch(()=>{});mov.id=cs[0].id}else S.movs.push(mov);\n  const movN=N>1?null:mov;\n  S.notifs.pend=S.notifs.pend.filter(x=>x.id!==id);S.notifs.rec.unshift({...n,estado:'asignada',cuenta:c,concepto:k||'Transferencia'});\n  reglas.forEach(r=>{S.reglas=S.reglas.filter(x=>!(x.tipo===r.tipo&&norm(x.palabra)===norm(r.palabra)&&norm(x.app||'')===norm(r.app||'')));S.reglas.push(r)});\n  invalidate();render();toast('Guardado');\n  await op('asignarNotif',JSON.stringify({id,mov:movN||S.movs.find(x=>x.id===mov.id),reglas})).catch(()=>{});\n  if(reglas.length&&S.notifs.pend.length)reprocesar(true);\n}\nasync function estadoNotif(id,estado){\n  const all=[...S.notifs.pend,...S.notifs.rec];const n=all.find(x=>x.id===id);if(!n)return;\n  S.notifs.pend=S.notifs.pend.filter(x=>x.id!==id);S.notifs.rec=S.notifs.rec.filter(x=>x.id!==id);\n  n.estado=estado;if(estado==='pendiente')S.notifs.pend.push(n);else if(estado!=='descartada')S.notifs.rec.unshift(n);\n  render();if(estado==='descartada')toast('Descartado');\n  await op('estadoNotif',id,estado).catch(()=>{});\n}\nasync function borrarNotif(id){\n  S.notifs.pend=S.notifs.pend.filter(x=>x.id!==id);S.notifs.rec=S.notifs.rec.filter(x=>x.id!==id);S.notifs.cuotas=S.notifs.cuotas.filter(x=>x.id!==id);\n  render();toast('Borrado');await op('borrarNotif',id).catch(()=>{});\n}\nasync function reprocesar(silencioso){\n  const k=await op('reprocesar').catch(()=>null);\n  if(k){toast(`Se asignaron ${k} m\u00e1s con las reglas nuevas`);load()}else if(!silencioso)toast('Ninguna coincide todav\u00eda con las reglas');\n}\n\n/* ---------- Presupuesto ---------- */\nfunction sem(p){p=Math.round(p);return p<=75?'v':p<=100?'a':'r'}\n/* en ingresos el sem\u00e1foro va al rev\u00e9s: llegar al 100% es lo bueno */\nfunction semT(p,t){if(t!=='I')return sem(p);p=Math.round(p);return p>=100?'v':p>=75?'a':'r'}\nfunction pctTxt(gas,pres){if(!pres)return '';const p=gas/pres*100;return `<span class=\"tx-${sem(p)}\">${Math.round(p)}%</span>`}\nfunction extrasMes(k,ymx,t){const y=+ymx.slice(0,4),m=+ymx.slice(5,7);t=t||'G';return (S.extras||[]).filter(x=>x.k===k&&(x.t||'G')===t&&x.mes===m&&(!x.anio||x.anio===y))}\nfunction presMes(c,ymx){return (c.pres||0)+extrasMes(c.n,ymx,c.tipo).reduce((s,x)=>s+x.m,0)}\nfunction presPeriodo(c,P){const y=P.per.slice(0,4),n=P.modo==='mes'?null:P.modo==='corriente'?+P.per.slice(5,7):12;\n  if(!n)return presMes(c,P.per);let t=0;for(let i=1;i<=n;i++)t+=presMes(c,y+'-'+p2(i));return t}\nfunction movConcepto(t,filtro){const r={};for(const x of S.movs)if(x.t===t&&filtro(x)){r[x.k]=(r[x.k]||0)+toUYU(x.m,monOf(x.c))}return r}\nfunction gastosConcepto(filtro){return movConcepto('G',filtro)}\nfunction presResumen(mes){\n  const g=gastosConcepto(x=>x.f.startsWith(mes));let pres=0,gas=0;\n  for(const c of S.cfg.conceptos){if(c.tipo!=='G')continue;const p=presMes(c,mes);if(p>0){pres+=p;gas+=g[c.n]||0}}\n  return {pres,gas};\n}\nfunction fmt0(v){const r=Math.round(v);return (r<0?'\u2212':'')+'$ '+Math.abs(r).toLocaleString('es-UY')}\nfunction pctCelda(g,p,t){if(!p)return '<td class=\"pct mu\">\u2014</td>';const x=g/p*100;return `<td class=\"pct tx-${semT(x,t)}\">${Math.round(x)}%</td>`}\nfunction vPresupuesto(){\n  const P=V.pres||(V.pres={per:ym(hoy()),modo:'mes',edit:false,tipo:'G'});\n  P.tipo=P.tipo||'G';\n  const y=P.per.slice(0,4),mesN=+P.per.slice(5,7),t=P.tipo;\n  const tabs=`<div class=\"seg\" style=\"margin-bottom:12px\">${[['G','Gastos'],['I','Ingresos'],['R','Excedente']].map(([k,l])=>`<button class=\"${t===k?'on '+(k==='R'?'':k):''}\" onclick=\"V.pres.tipo='${k}';V.pres.edit=false;render()\">${l}</button>`).join('')}</div>`;\n  if(t==='R')return tabs+vExcedente(P);\n  const fil=P.modo==='mes'?(x=>x.f.startsWith(P.per)):P.modo==='corriente'?(x=>x.f.startsWith(y)&&x.f.slice(0,7)<=P.per):(x=>x.f.startsWith(y));\n  const gas=movConcepto(t,fil);\n  const per=P.modo==='mes'?P.per:y;\n  const G=t==='G';\n  const cab=`${tabs}<div class=\"card\">\n   <div style=\"display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px\">\n    ${pnav('V.pres.per',P.per)}\n    <div class=\"seg\">${[['mes','Mes'],['corriente','Al corriente'],['anio','A\u00f1o']].map(([k,l])=>`<button class=\"${P.modo===k?'on':''}\" onclick=\"V.pres.modo='${k}';render()\">${l}</button>`).join('')}</div>\n   </div>\n   <div class=\"xs mu\" style=\"margin-top:8px\">${P.modo==='mes'?(G?'Gastos':'Ingresos')+' del mes contra lo presupuestado.':P.modo==='corriente'?`De enero a ${MESES[mesN-1]}.`:`Todo ${y}.`}\n    ${G?'Avance: <b class=\"tx-v\">verde</b> hasta 75% \u00b7 <b class=\"tx-a\">amarillo</b> 76\u2013100% \u00b7 <b class=\"tx-r\">rojo</b> m\u00e1s de 100%.':'Cobrado: <b class=\"tx-v\">verde</b> 100% o m\u00e1s \u00b7 <b class=\"tx-a\">amarillo</b> 75\u201399% \u00b7 <b class=\"tx-r\">rojo</b> menos de 75%.'}</div>\n  </div>`;\n  if(P.edit){\n    const cats={};for(const c of S.cfg.conceptos.filter(c=>c.tipo===t))(cats[c.cat||'Sin categor\u00eda']||(cats[c.cat||'Sin categor\u00eda']=[])).push(c);\n    return cab+`<div class=\"card\"><h2>${G?'Gasto':'Ingreso'} mensual previsto por concepto</h2>\n     <p class=\"sm mu\" style=\"margin-top:0\">${G?'Pon\u00e9 cu\u00e1nto quer\u00e9s gastar por mes en cada concepto':'Pon\u00e9 cu\u00e1nto esper\u00e1s cobrar por mes en cada concepto (sueldo, alquileres\u2026)'} en pesos. Dejalo vac\u00edo si no lo quer\u00e9s controlar.</p>\n     <table class=\"pr ed\">${Object.keys(cats).sort((a,b)=>a.localeCompare(b,'es')).map(cat=>`<tr class=\"cat\"><td colspan=\"2\">${h(cat)}</td></tr>`+sortN(cats[cat]).map(c=>`<tr><td>${h(c.n)}</td><td><input type=\"text\" inputmode=\"decimal\" data-con=\"${h(c.n)}\" value=\"${c.pres?String(c.pres).replace('.',','):''}\" placeholder=\"0\"></td></tr>`).join('')).join('')||'<tr><td class=\"empty\">No hay conceptos de este tipo. Agregalos en Ajustes.</td></tr>'}</table>\n     <div class=\"btns\" style=\"justify-content:flex-end;margin-top:10px\"><button class=\"btn\" onclick=\"V.pres.edit=false;render()\">Cancelar</button><button class=\"btn pri\" onclick=\"guardarPres('${t}')\">Guardar</button></div></div>`+vExtras(P);\n  }\n  const cats={},sin=[];\n  for(const c of S.cfg.conceptos.filter(c=>c.tipo===t)){\n    const g=gas[c.n]||0,pp=presPeriodo(c,P);\n    if(pp>0)(cats[c.cat||'Sin categor\u00eda']||(cats[c.cat||'Sin categor\u00eda']=[])).push({c,g,p:pp,ex:P.modo==='mes'?extrasMes(c.n,P.per,t):[]});\n    else if(g>0.5)sin.push({c,g});\n  }\n  let TP=0,TG=0,filas='';\n  for(const cat of Object.keys(cats).sort((a,b)=>a.localeCompare(b,'es'))){\n    const L=cats[cat].sort((a,b)=>b.p-a.p);\n    const cp=L.reduce((s,r)=>s+r.p,0),cg=L.reduce((s,r)=>s+r.g,0);TP+=cp;TG+=cg;\n    const dif=v=>G?v:-v;\n    if(L.length>1)filas+=`<tr class=\"cat\"><td>${h(cat)}</td><td class=\"num\">${fmt0(cp)}</td><td class=\"num\">${fmt0(cg)}</td>${pctCelda(cg,cp,t)}<td class=\"disp num ${dif(cp-cg)<0&&G?'neg':''}\">${fmt0(cp-cg)}</td></tr>`;\n    for(const r of L)filas+=`<tr class=\"${L.length>1?'sub':'cat1'}\" onclick=\"filtrar({k:${js(t+'|'+r.c.n)}},'${per}')\"><td>${h(r.c.n)}${L.length===1&&norm(cat)!==norm(r.c.n)?`<div class=\"xs mu\">${h(cat)}</div>`:''}${r.ex.length?`<div class=\"xs mu\">incluye ${r.ex.map(x=>h(x.n)).join(', ')}</div>`:''}</td>\n      <td class=\"num\">${fmt0(r.p)}</td><td class=\"num\">${fmt0(r.g)}</td>${pctCelda(r.g,r.p,t)}<td class=\"disp num ${G&&r.p-r.g<0?'neg':''}\">${fmt0(r.p-r.g)}</td></tr>`;\n  }\n  const totSin=sin.reduce((s,r)=>s+r.g,0);\n  const L1=G?['Presupuesto','Gastado','Disponible']:['Previsto','Cobrado','Falta cobrar'];\n  return cab+`<div class=\"card\">\n   <h2>${G?'Ejecuci\u00f3n del presupuesto':'Ingresos previstos y cobrados'} <button class=\"btn sm\" onclick=\"V.pres.edit=true;render()\">Editar</button></h2>\n   ${TP?`<div class=\"resumen\"><div><div class=\"xs mu\">${L1[0]}</div><b class=\"num\">${fmt0(TP)}</b></div><div><div class=\"xs mu\">${L1[1]}</div><b class=\"num\">${fmt0(TG)}</b></div>\n     <div><div class=\"xs mu\">Avance</div><b class=\"num tx-${semT(TG/TP*100,t)}\" style=\"font-size:20px\">${Math.round(TG/TP*100)}%</b></div><div><div class=\"xs mu\">${L1[2]}</div><b class=\"num ${G&&TP-TG<0?'neg':''}\">${fmt0(G?TP-TG:Math.max(0,TP-TG))}</b></div></div>\n   <table class=\"pr\"><tr><th>Concepto</th><th>${G?'Presup.':'Previsto'}</th><th>${G?'Gastado':'Cobrado'}</th><th>%</th><th class=\"disp\">${G?'Disponible':'Diferencia'}</th></tr>${filas}</table>`\n   :`<div class=\"empty\">Todav\u00eda no cargaste ${G?'presupuestos de gastos':'ingresos previstos'}.<br><button class=\"btn pri\" style=\"margin-top:10px\" onclick=\"V.pres.edit=true;render()\">${G?'Cargar presupuestos':'Cargar ingresos previstos'}</button></div>`}\n   ${sin.length?`<details class=\"sinp\"><summary><span>${G?'Gastos':'Ingresos'} sin presupuesto (${sin.length})</span><b class=\"num\">${fmt0(totSin)}</b></summary>\n     <table class=\"pr\">${sin.sort((a,b)=>b.g-a.g).map(r=>`<tr class=\"sub\" onclick=\"filtrar({k:${js(t+'|'+r.c.n)}},'${per}')\"><td>${h(r.c.n)}<div class=\"xs mu\">${h(r.c.cat||'')}</div></td><td class=\"num\">${fmt0(r.g)}</td></tr>`).join('')}</table></details>`:''}\n   <div class=\"xs mu\" style=\"margin-top:8px\">En pesos; d\u00f3lares a $ ${NF.format(cot())}. Toc\u00e1 un concepto para ver sus movimientos.</div>\n  </div>\n  ${vExtras(P)}`;\n}\n/* ---------- Excedente: ingresos previstos \u2212 gastos previstos, mes a mes ---------- */\nfunction vExcedente(P){\n  const y=P.per.slice(0,4),actual=ym(hoy());\n  const CG=S.cfg.conceptos.filter(c=>c.tipo==='G'),CI=S.cfg.conceptos.filter(c=>c.tipo==='I');\n  let T={pi:0,pg:0,ri:0,rg:0,proy:0},filas='',hayPres=false;\n  for(let m=1;m<=12;m++){\n    const mm=y+'-'+p2(m);\n    const pi=CI.reduce((s,c)=>s+presMes(c,mm),0),pg=CG.reduce((s,c)=>s+presMes(c,mm),0);\n    const {g:rg,i:ri}=sumGI(S.movs.filter(x=>x.f.startsWith(mm)));\n    const pasado=mm<actual,esActual=mm===actual,futuro=mm>actual;\n    if(pi||pg)hayPres=true;\n    T.pi+=pi;T.pg+=pg;if(!futuro){T.ri+=ri;T.rg+=rg}\n    T.proy+=pasado?(ri-rg):esActual?Math.max(pi,ri)-Math.max(pg,rg):(pi-pg);\n    const ep=pi-pg,er=ri-rg;\n    filas+=`<tr class=\"${esActual?'cat':''}\" onclick=\"V.pres.per='${mm}';V.pres.modo='mes';V.pres.tipo='G';render()\"><td>${MESES[m-1].charAt(0).toUpperCase()+MESES[m-1].slice(1)}${esActual?' <span class=\"chip\">hoy</span>':''}</td>\n      <td class=\"num disp\">${fmt0(pi)}</td><td class=\"num disp\">${fmt0(pg)}</td><td class=\"num ${ep<0?'neg':'pos'}\"><b>${fmt0(ep)}</b></td>\n      <td class=\"num ${futuro?'mu':er<0?'neg':'pos'}\">${futuro?'\u2014':fmt0(er)}</td></tr>`;\n  }\n  const exP=T.pi-T.pg,exR=T.ri-T.rg;\n  const sinPresG=sumGI(S.movs.filter(x=>x.f.startsWith(y)&&x.t==='G'&&!(con('G',x.k)&&presMes(con('G',x.k),x.f.slice(0,7))>0))).g;\n  return `<div class=\"card\">\n   <div style=\"display:flex;justify-content:space-between;align-items:center;flex-wrap:wrap;gap:8px\">\n    <div class=\"pnav\"><button onclick=\"V.pres.per=perStep(V.pres.per.slice(0,4),-1)+'-12';render()\">\u2039</button><b>A\u00f1o ${y}</b><button onclick=\"V.pres.per=perStep(V.pres.per.slice(0,4),1)+'-01';render()\">\u203a</button></div>\n   </div>\n   <div class=\"resumen\" style=\"margin-top:12px\"><div><div class=\"xs mu\">Ingresos previstos</div><b class=\"num\">${fmt0(T.pi)}</b></div><div><div class=\"xs mu\">Gastos previstos</div><b class=\"num\">${fmt0(T.pg)}</b></div>\n    <div><div class=\"xs mu\">Excedente previsto</div><b class=\"num ${exP<0?'neg':'pos'}\" style=\"font-size:20px\">${fmt0(exP)}</b></div>\n    <div><div class=\"xs mu\">${y===actual.slice(0,4)?'Proyecci\u00f3n al cierre':'Excedente real'}</div><b class=\"num ${(y===actual.slice(0,4)?T.proy:exR)<0?'neg':'pos'}\" style=\"font-size:20px\">${fmt0(y===actual.slice(0,4)?T.proy:exR)}</b></div></div>\n   ${y===actual.slice(0,4)?`<div class=\"sm\" style=\"margin-bottom:8px\">Excedente real hasta hoy: <b class=\"num ${exR<0?'neg':'pos'}\">${fmt0(exR)}</b> <span class=\"mu xs\">(ingresos ${fmt0(T.ri)} \u2212 gastos ${fmt0(T.rg)})</span></div>`:''}\n   ${hayPres?`<table class=\"pr\"><tr><th>Mes</th><th class=\"disp\">Ingresos prev.</th><th class=\"disp\">Gastos prev.</th><th>Excedente prev.</th><th>Excedente real</th></tr>${filas}\n     <tr class=\"cat\"><td>Total</td><td class=\"num disp\">${fmt0(T.pi)}</td><td class=\"num disp\">${fmt0(T.pg)}</td><td class=\"num ${exP<0?'neg':''}\">${fmt0(exP)}</td><td class=\"num ${exR<0?'neg':''}\">${fmt0(exR)}</td></tr></table>`\n   :`<div class=\"empty\">Carg\u00e1 los ingresos previstos (pesta\u00f1a Ingresos) y el presupuesto de gastos (pesta\u00f1a Gastos) para ver el excedente del a\u00f1o.</div>`}\n   <div class=\"xs mu\" style=\"margin-top:8px\">El previsto incluye los montos mensuales y los anuales (aguinaldo, patente\u2026). El real incluye todos los movimientos.\n    ${sinPresG>0.5?` Este a\u00f1o hay <b>${fmt0(sinPresG)}</b> de gastos en conceptos sin presupuesto: si se repiten, conviene presupuestarlos para que el excedente previsto sea realista.`:''}\n    La proyecci\u00f3n suma lo real de los meses pasados y lo previsto de los que faltan.</div>\n  </div>`;\n}\nfunction vExtras(P){\n  const t=P.tipo==='I'?'I':'G',G=t==='G';\n  const L=[...(S.extras||[])].filter(x=>(x.t||'G')===t).sort((a,b)=>a.mes-b.mes||a.n.localeCompare(b.n,'es'));\n  const y=+P.per.slice(0,4);\n  const tot=L.filter(x=>!x.anio||x.anio===y).reduce((s,x)=>s+x.m,0);\n  return `<div class=\"card\"><h2>${G?'Gastos':'Ingresos'} anuales previstos <button class=\"btn sm\" onclick=\"openExtra(null,'${t}')\">+ Agregar</button></h2>\n   <p class=\"xs mu\" style=\"margin-top:0\">${G?'Gastos grandes que caen una vez al a\u00f1o en un mes fijo (patente, contribuci\u00f3n, escolaridad\u2026).':'Ingresos que llegan en meses fijos (aguinaldo en junio y diciembre, salario vacacional\u2026).'} Ese mes se suman a lo previsto de su concepto.</p>\n   ${L.length?L.map(x=>`<div class=\"row\" onclick=\"openExtra(${js(x.id)})\"><div class=\"ico ${t}\" style=\"font-size:11px\">${MES3[x.mes-1]}</div><div class=\"mid\"><div class=\"t1\">${h(x.n)}</div>\n     <div class=\"t2\">${h(x.k)} \u00b7 ${x.anio?'solo '+x.anio:'todos los a\u00f1os'} \u00b7 ${MESES[x.mes-1]}</div></div><div class=\"num\">${fmt(x.m,'UYU')}</div></div>`).join('')\n     +`<div class=\"kv\" style=\"margin-top:6px\"><b>Total previsto ${y}</b><b class=\"num\">${fmt(tot,'UYU')}</b></div>`:'<div class=\"empty\">Todav\u00eda no cargaste ninguno.</div>'}\n  </div>`;\n}\nfunction openExtra(id,tipo){\n  const x=id?S.extras.find(e=>e.id===id):{n:'',k:'',mes:+hoy().slice(5,7),anio:0,m:'',t:tipo||'G'};\n  const t=x.t||'G',G=t==='G';\n  const conc=sortN(S.cfg.conceptos.filter(c=>c.tipo===t));\n  modal(`<h3>${id?'Editar':'Nuevo'} ${G?'gasto':'ingreso'} anual previsto</h3><form onsubmit=\"event.preventDefault();guardarExtra(${js(id||'')},'${t}')\">\n   <div class=\"fld\"><label>Nombre</label><input type=\"text\" id=\"ex_n\" value=\"${h(x.n)}\" placeholder=\"${G?'Ej. Patente del auto, Contribuci\u00f3n, Escolaridad':'Ej. Aguinaldo, Salario vacacional'}\" required></div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Mes</label><select id=\"ex_mes\">${MESES.map((m,i)=>`<option value=\"${i+1}\"${i+1===x.mes?' selected':''}>${m}</option>`).join('')}</select></div>\n   <div class=\"fld\"><label>Monto ($)</label><input type=\"text\" inputmode=\"decimal\" id=\"ex_m\" value=\"${x.m?String(x.m).replace('.',','):''}\" required></div></div>\n   <div class=\"fld\"><label>Concepto</label><input type=\"text\" id=\"ex_k\" list=\"dl_exk\" value=\"${h(x.k)}\" placeholder=\"Eleg\u00ed o escrib\u00ed uno nuevo\" autocomplete=\"off\" required>\n    <datalist id=\"dl_exk\">${conc.map(c=>`<option value=\"${h(c.n)}\">`).join('')}</datalist>\n    <div class=\"xs mu\" style=\"margin-top:3px\">Si escrib\u00eds un concepto que no existe (ej. \u201c${G?'Escolaridad':'Aguinaldo'}\u201d), se crea en la categor\u00eda \u201c${G?'Gastos':'Ingresos'} anuales\u201d.</div></div>\n   <div class=\"fld\"><label>Se repite</label><select id=\"ex_a\"><option value=\"0\"${!x.anio?' selected':''}>Todos los a\u00f1os</option>${[0,1,2].map(d=>{const a=+hoy().slice(0,4)+d;return `<option value=\"${a}\"${x.anio===a?' selected':''}>Solo en ${a}</option>`}).join('')}</select></div>\n   <div class=\"btns\" style=\"justify-content:space-between\"><div>${id?`<button type=\"button\" class=\"btn dan\" onclick=\"borrarExtra(${js(id)})\">Eliminar</button>`:''}</div>\n   <div class=\"btns\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\">Guardar</button></div></div></form>`);\n  setTimeout(()=>document.getElementById('ex_n').focus(),50);\n}\nfunction guardarExtra(id,t){\n  t=t||'G';\n  const g=i=>document.getElementById(i).value;\n  const x={id:id||uid(),n:g('ex_n').trim(),k:g('ex_k').trim(),mes:+g('ex_mes'),anio:+g('ex_a')||0,m:r2(Math.abs(parseNum(g('ex_m')))),t};\n  if(!x.n||!x.k||!(x.m>0)){toast('Complet\u00e1 nombre, concepto y monto');return}\n  const ex=S.cfg.conceptos.find(c=>c.tipo===t&&norm(c.n)===norm(x.k));\n  if(ex)x.k=ex.n;else{S.cfg.conceptos.push({n:x.k,tipo:t,cat:t==='G'?'Gastos anuales':'Ingresos anuales'});invalidate();saveCfg()}\n  S.extras=(S.extras||[]).filter(e=>e.id!==x.id);S.extras.push(x);\n  op('saveExtras',JSON.stringify(S.extras)).catch(()=>{});closeModal();render();toast('Guardado');\n}\nfunction borrarExtra(id){confirmar('\u00bfEliminar este monto anual previsto?',()=>{S.extras=S.extras.filter(e=>e.id!==id);op('saveExtras',JSON.stringify(S.extras)).catch(()=>{});closeModal();render()})}\nfunction dispTxt(v){return `<span class=\"num ${v<0?'neg':''}\">${fmtS(v,'UYU')}</span>`}\nfunction guardarPres(t){\n  t=t||'G';\n  document.querySelectorAll('input[data-con]').forEach(i=>{const c=con(t,i.dataset.con);if(c)c.pres=r2(parseNum(i.value)||0)});\n  V.pres.edit=false;saveCfg();render();toast('Guardado');\n}\n\n/* ---------- M\u00e1s (celular) ---------- */\nfunction vMas(){\n  return `<div class=\"card mas\"><a onclick=\"fijarSaldos()\"><svg viewBox=\"0 0 24 24\" fill=\"none\" stroke=\"currentColor\" stroke-width=\"2\"><path d=\"M12 20V10M18 20V4M6 20v-4\"/></svg><span><b>Fijar saldos de hoy</b><br><span class=\"xs mu\">Lo que ten\u00e9s en cada cuenta y lo que deb\u00e9s en cada tarjeta</span></span></a>${[['cuentas','Cuentas'],['informes','Informes'],['fijos','Gastos fijos'],['ajustes','Ajustes']].map(([k,l])=>`<a onclick=\"go('${k}')\">${ICONS[k]||''}<span>${l}</span></a>`).join('')}</div>`;\n}\n\n/* ---------- Gastos fijos ---------- */\nfunction vFijos(){\n  const L=S.fijos;\n  return `<div class=\"btns\" style=\"margin-bottom:12px\"><button class=\"btn pri\" onclick=\"openFijo()\">+ Nuevo gasto fijo</button></div>\n  <div class=\"card\"><h2>Gastos e ingresos fijos</h2>\n   <p class=\"xs mu\" style=\"margin-top:0\">Cada mes se crea solo el movimiento en el d\u00eda indicado (aparece como \u201cfuturo\u201d hasta esa fecha). Si un mes cambia el importe, edit\u00e1 ese movimiento.</p>\n   ${L.length?L.map(f=>`<div class=\"row\" onclick=\"openFijo(${js(f.id)})\"><div class=\"ico ${f.t}\">${f.t==='I'?'+':'\u27f3'}</div><div class=\"mid\"><div class=\"t1\">${h(f.n)}${f.activo?'':'<span class=\"chip\" style=\"background:var(--bg);color:var(--mu)\">pausado</span>'}</div>\n     <div class=\"t2\">D\u00eda ${f.dia} \u00b7 ${h(f.k)} \u00b7 ${h(f.c)} \u00b7 desde ${h(perLabel(f.desde))}${f.hasta?' hasta '+h(perLabel(f.hasta)):''}</div></div><div class=\"num ${f.t==='I'?'pos':'neg'}\">${fmt(f.m,monOf(f.c))}</div></div>`).join(''):'<div class=\"empty\">Agreg\u00e1 la hipoteca, ANTEL, UTE, el sueldo\u2026 lo que se repite todos los meses.</div>'}\n  </div>`;\n}\nfunction openFijo(id){\n  const f=id?S.fijos.find(x=>x.id===id):{n:'',t:'G',c:(S.cfg.cuentas.find(c=>!c.arch)||{}).n||'',k:'',m:'',dia:1,desde:ym(hoy()),hasta:'',activo:true};\n  V._fj={...f};\n  const accO=S.cfg.cuentas.filter(c=>!c.arch||c.n===f.c).map(c=>`<option value=\"${h(c.n)}\"${c.n===f.c?' selected':''}>${h(c.n)} (${sym(c.mon)})</option>`).join('');\n  const kO=t=>'<option value=\"\">\u2014 elegir \u2014</option>'+sortN(S.cfg.conceptos.filter(c=>c.tipo===t)).map(c=>`<option value=\"${h(c.n)}\"${c.n===f.k?' selected':''}>${h(c.n)}</option>`).join('');\n  modal(`<h3>${id?'Editar':'Nuevo'} gasto fijo</h3><form onsubmit=\"event.preventDefault();guardarFijo(${js(id||'')})\">\n   <div class=\"seg\" style=\"margin-bottom:12px\"><button type=\"button\" class=\"${f.t==='G'?'on G':''}\" onclick=\"V._fj.t='G';fjTipo('G')\">Gasto</button><button type=\"button\" class=\"${f.t==='I'?'on I':''}\" onclick=\"fjTipo('I')\">Ingreso</button></div>\n   <div class=\"fld\"><label>Nombre</label><input type=\"text\" id=\"fj_n\" value=\"${h(f.n)}\" placeholder=\"Ej. Hipoteca, ANTEL, Sueldo\" required></div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Importe</label><input type=\"text\" inputmode=\"decimal\" id=\"fj_m\" value=\"${f.m?String(f.m).replace('.',','):''}\" required></div>\n   <div class=\"fld\"><label>D\u00eda del mes</label><input type=\"number\" id=\"fj_d\" min=\"1\" max=\"31\" value=\"${f.dia}\"></div></div>\n   <div class=\"fld\"><label>Cuenta</label><select id=\"fj_c\">${accO}</select></div>\n   <div class=\"fld\"><label>Concepto</label><select id=\"fj_k\" data-t=\"${f.t}\">${kO(f.t)}</select></div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Desde (mes)</label><input type=\"month\" id=\"fj_desde\" value=\"${h(f.desde)}\"></div>\n   <div class=\"fld\"><label>Hasta (opcional)</label><input type=\"month\" id=\"fj_hasta\" value=\"${h(f.hasta||'')}\"></div></div>\n   <div class=\"fld\"><label><input type=\"checkbox\" id=\"fj_a\"${f.activo?' checked':''}> Activo</label></div>\n   <div class=\"btns\" style=\"justify-content:space-between\"><div>${id?`<button type=\"button\" class=\"btn dan\" onclick=\"borrarFijo(${js(id)})\">Eliminar</button>`:''}</div>\n   <div class=\"btns\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\">Guardar</button></div></div></form>`);\n}\nfunction fjTipo(t){V._fj.t=t;document.querySelectorAll('#sheet .seg button').forEach((b,i)=>b.className=(i===0&&t==='G')?'on G':(i===1&&t==='I')?'on I':'');\n  document.getElementById('fj_k').innerHTML='<option value=\"\">\u2014 elegir \u2014</option>'+sortN(S.cfg.conceptos.filter(c=>c.tipo===t)).map(c=>`<option value=\"${h(c.n)}\">${h(c.n)}</option>`).join('')}\nfunction guardarFijo(id){\n  const g=i=>document.getElementById(i).value;\n  const f={id:id||uid(),n:g('fj_n').trim(),t:V._fj.t,c:g('fj_c'),k:g('fj_k'),m:r2(Math.abs(parseNum(g('fj_m')))),dia:Math.max(1,Math.min(31,+g('fj_d')||1)),\n    desde:g('fj_desde')||ym(hoy()),hasta:g('fj_hasta')||'',activo:document.getElementById('fj_a').checked,ult:(S.fijos.find(x=>x.id===id)||{}).ult||''};\n  if(!f.n||!(f.m>0)||!f.c||!f.k){toast('Complet\u00e1 nombre, importe, cuenta y concepto');return}\n  S.fijos=S.fijos.filter(x=>x.id!==f.id);S.fijos.push(f);\n  closeModal();render();\n  op('saveFijos',JSON.stringify(S.fijos)).then(()=>load()).catch(()=>{});toast('Gasto fijo guardado');\n}\nfunction borrarFijo(id){confirmar('\u00bfEliminar este gasto fijo? Los movimientos ya creados quedan.',()=>{S.fijos=S.fijos.filter(x=>x.id!==id);op('saveFijos',JSON.stringify(S.fijos)).catch(()=>{});closeModal();render()})}\n\n/* ---------- Cuotas de compras que entraron solas ---------- */\nfunction cardCuotas(n){\n  const mv=S.movs.find(x=>x.id===n.mov);\n  if(!mv)return '';\n  return `<div class=\"row\" style=\"cursor:default;flex-wrap:wrap\" id=\"cq_${h(n.id)}\"><div class=\"mid\" style=\"min-width:160px\"><div class=\"t1\">${h(mv.d||mv.k)} \u00b7 <span class=\"num\">${fmt(mv.m,monOf(mv.c))}</span></div>\n    <div class=\"t2\">${h(mv.f.split('-').reverse().join('/'))} \u00b7 ${h(mv.c)} \u00b7 ${h(mv.k)}</div></div>\n    <div class=\"btns\">${[1,2,3,6,12].map(k=>`<button class=\"btn sm qbtn${k===(n.cuotas||1)?' pri':''}\" onclick=\"aplicarCuotas(${js(n.id)},${k})\">${k===1?'1 pago':k}</button>`).join('')}\n    <input type=\"number\" min=\"2\" max=\"60\" placeholder=\"otro\" style=\"width:70px;padding:6px\" onchange=\"if(+this.value>1)aplicarCuotas(${js(n.id)},+this.value)\"></div></div>`;\n}\nasync function aplicarCuotas(id,N){\n  const n=S.notifs.cuotas.find(x=>x.id===id);const mv=S.movs.find(x=>x.id===n.mov);\n  S.notifs.cuotas=S.notifs.cuotas.filter(x=>x.id!==id);\n  let movs=[];\n  if(mv&&N>1){movs=crearCuotas(mv,N,mv.m);S.movs=S.movs.filter(x=>x.id!==mv.id);S.movs.push(...movs);invalidate()}\n  render();toast(N>1?`Dividido en ${N} cuotas`:'Listo: 1 pago');\n  await op('cuotasNotif',JSON.stringify({id,borrar:mv?mv.id:'',movs})).catch(()=>{});\n}\n\n/* ---------- Diferencias de estados de cuenta ---------- */\nconst DIF_TXT={falta_app:'Est\u00e1 en el estado de cuenta y no en la app',falta_estado:'Est\u00e1 en la app y no en el estado de cuenta',importe:'El importe no coincide'};\nfunction cardDif(d){\n  const mon=d.mon||monOf(d.c);\n  let cuerpo='',aprobar='Aprobar',rechazar='Rechazar';\n  if(d.tipo==='falta_app'){\n    const t=d.m<0?'I':'G';const sug=conceptoPorReglas(d.d,t);\n    const qq=(d.q||'').split('/').map(Number);\n    cuerpo=`<div class=\"fld2\"><div class=\"fld\"><label>Tipo</label><select class=\"df_t\" onchange=\"difTipo(${js(d.id)},this.value)\"><option value=\"G\"${t==='G'?' selected':''}>Gasto</option><option value=\"I\"${t==='I'?' selected':''}>Ingreso / devoluci\u00f3n</option></select></div>\n      <div class=\"fld\"><label>Concepto</label><select class=\"df_k\">${optsConcepto(t,sug)}</select></div></div>\n      ${qq[1]>1&&qq[0]<qq[1]?`<div class=\"fld\"><label><input type=\"checkbox\" class=\"df_q\" checked> Crear tambi\u00e9n las cuotas siguientes (${qq[0]+1} a ${qq[1]})</label></div>`:''}`;\n    aprobar='Agregar a la app';\n  }else if(d.tipo==='falta_estado'){aprobar='Eliminar de la app';rechazar='Mantener'}\n  else{cuerpo=`<div class=\"sm\">En la app: <b class=\"num\">${fmt(d.mApp,mon)}</b> \u00b7 En el estado: <b class=\"num\">${fmt(Math.abs(d.m),mon)}</b></div>`;aprobar='Usar importe del estado';rechazar='Dejar como est\u00e1'}\n  return `<div class=\"card pc\" id=\"df_${h(d.id)}\"><h2><span style=\"text-transform:none\">${DIF_TXT[d.tipo]||d.tipo}</span></h2>\n   <div class=\"kv\"><span><b>${h(d.d||'\u2014')}</b>${d.q?' <span class=\"chip\">cuota '+h(d.q)+'</span>':''}<br><span class=\"xs mu\">${h(d.f.split('-').reverse().join('/'))} \u00b7 ${h(d.c)}</span></span><b class=\"num ${d.m<0?'pos':''}\">${fmtS(-d.m,mon).replace('\u2212','')}${d.m<0?' (cr\u00e9dito)':''}</b></div>\n   ${cuerpo}\n   <div class=\"btns\" style=\"justify-content:space-between;margin-top:6px\"><button class=\"btn\" onclick=\"resolverDif(${js(d.id)},false)\">${rechazar}</button><button class=\"btn pri\" onclick=\"resolverDif(${js(d.id)},true)\">${aprobar}</button></div></div>`;\n}\nfunction optsConcepto(t,sel){const cats={};for(const c of S.cfg.conceptos.filter(c=>c.tipo===t))(cats[c.cat||'Sin categor\u00eda']||(cats[c.cat||'Sin categor\u00eda']=[])).push(c);\n  return '<option value=\"\">\u2014 elegir concepto \u2014</option>'+Object.keys(cats).sort((a,b)=>a.localeCompare(b,'es')).map(cat=>`<optgroup label=\"${h(cat)}\">${sortN(cats[cat]).map(c=>`<option value=\"${h(c.n)}\"${c.n===sel?' selected':''}>${h(c.n)}</option>`).join('')}</optgroup>`).join('')}\nfunction difTipo(id,t){document.querySelector('#df_'+id+' .df_k').innerHTML=optsConcepto(t,conceptoPorReglas(S.difs.find(x=>x.id===id).d,t))}\nfunction palabrasC(s){return ' '+norm(s).replace(/[^a-z0-9]+/g,' ').trim()+' '}\nfunction conceptoPorReglas(txt,t){\n  const W=palabrasC(txt);let best=null;\n  for(const r of S.reglas){if(r.tipo!=='Concepto')continue;const w=palabrasC(r.palabra);if(w.trim()&&W.includes(w)&&con(t,r.valor)&&(!best||w.length>best.w.length))best={w,v:r.valor}}\n  return best?best.v:'';\n}\nasync function resolverDif(id,aprobar){\n  const d=S.difs.find(x=>x.id===id),el=document.getElementById('df_'+id);\n  if(aprobar){\n    if(d.tipo==='falta_app'){\n      const t=el.querySelector('.df_t').value,k=el.querySelector('.df_k').value;\n      if(!k){toast('Eleg\u00ed el concepto');return}\n      const base={t,f:d.f,c:d.c,k,m:r2(Math.abs(d.m)),d:d.d,cc:d.ec,ts:Date.now()};\n      const qq=(d.q||'').split('/').map(Number),q=el.querySelector('.df_q');\n      let movs;\n      if(qq[1]>1){base.f=addMeses(d.f,qq[0]-1);movs=crearCuotas(base,qq[1],r2(Math.abs(d.m))*qq[1],qq[0]);movs.forEach((x,i)=>{x.m=r2(Math.abs(d.m));if(i>0)delete x.cc});if(!(q&&q.checked))movs=movs.slice(0,1)}\n      else movs=[{...base,id:uid()}];\n      S.movs.push(...movs);op('importMovs',JSON.stringify(movs)).catch(()=>{});\n    }else if(d.tipo==='falta_estado'){S.movs=S.movs.filter(x=>x.id!==d.mov);op('deleteMov',d.mov).catch(()=>{})}\n    else{const mv=S.movs.find(x=>x.id===d.mov);if(mv){mv.m=r2(Math.abs(d.m));mv.cc=d.ec;op('saveMov',JSON.stringify(mv)).catch(()=>{})}}\n    invalidate();\n  }\n  S.difs=S.difs.filter(x=>x.id!==id);render();toast(aprobar?'Aplicado':'Rechazado');\n  op('estadoDif',id,aprobar?'aprobada':'rechazada').catch(()=>{});\n}\n\n/* ---------- Carga de PDF del estado de cuenta ---------- */\nlet EC=null;\nfunction cargarPDF(){document.getElementById('pdfIn').click()}\nfunction cargarScript(src){return new Promise((res,rej)=>{const s=document.createElement('script');s.src=src;s.onload=res;s.onerror=rej;document.head.appendChild(s)})}\nasync function onPDF(inp){\n  const file=inp.files[0];inp.value='';if(!file)return;\n  toast('Leyendo el PDF\u2026');\n  try{\n    if(!window.pdfjsLib){await cargarScript('https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.min.js');\n      pdfjsLib.GlobalWorkerOptions.workerSrc='https://cdnjs.cloudflare.com/ajax/libs/pdf.js/3.11.174/pdf.worker.min.js'}\n    const pdf=await pdfjsLib.getDocument({data:new Uint8Array(await file.arrayBuffer())}).promise;\n    const lines=[];\n    for(let pn=1;pn<=pdf.numPages;pn++){\n      const page=await pdf.getPage(pn);const tc=await page.getTextContent();\n      const its=tc.items.filter(i=>i.str&&i.str.trim()).map(i=>({s:i.str.trim(),x:i.transform[4],y:Math.round(i.transform[5]*2)/2,w:i.width||0}));\n      its.sort((a,b)=>b.y-a.y||a.x-b.x);\n      let cur=null;\n      for(const it of its){if(!cur||Math.abs(cur.y-it.y)>2.5){cur={y:it.y,p:pn,items:[]};lines.push(cur)}cur.items.push(it)}\n      lines.forEach(l=>l.items.sort((a,b)=>a.x-b.x));\n    }\n    EC=parseEstado(lines,file.name);\n    if(!EC.items.length){toast('No encontr\u00e9 movimientos en el PDF. Mandame un ejemplo para ajustar la lectura.');return}\n    renderEC();\n  }catch(e){toast('No pude leer el PDF: '+(e.message||e))}\n}\nconst RX_IMP=/^\\(?-?\\s?(?:U\\$S|US\\$|USD|\\$)?\\s?-?(?:\\d+(?:[.,]\\d{3})*)?[.,]\\d{2}\\)?-?(?:\\s?CR)?$/i;\nconst RX_FEC=/^(\\d{1,2})[\\/\\-.](\\d{1,2})(?:[\\/\\-.](\\d{2,4}))?$/;\nconst RX_FEC2=/^(\\d{1,2})[\\s\\-\\/]?(ene|feb|mar|abr|may|jun|jul|ago|sep|set|oct|nov|dic)[a-z]*\\.?(?:[\\s\\-\\/](\\d{2,4}))?$/i;\nconst RX_FIN=/saldo contado|total a pagar|cargos del mes|ud\\. ha generado/i;\nconst RX_INI=/saldo (del estado de cuenta )?anterior/i;\nconst RX_SKIP=/saldo (anterior|actual|a pagar|deudor anterior|pendiente|contado)|saldo del estado|^total\\b|pago m[i\u00ed]nimo|l[i\u00ed]mite de|fecha de (cierre|vencimiento)|^vencimiento|cr[e\u00e9]dito disponible/i;\nfunction fechaTok(s,refY,refM){\n  let m=s.match(RX_FEC),d,mo,y;\n  if(m){d=+m[1];mo=+m[2];y=m[3]?(+m[3]<100?2000+ +m[3]:+m[3]):null}\n  else if((m=s.match(/^(\\d{2})\\s(\\d{2})\\s(\\d{2,4})$/))){d=+m[1];mo=+m[2];y=+m[3]<100?2000+ +m[3]:+m[3]}\n  else if((m=s.match(RX_FEC2))){d=+m[1];mo=MES3.indexOf(m[2].toLowerCase().replace('set','sep').slice(0,3))+1;y=m[3]?(+m[3]<100?2000+ +m[3]:+m[3]):null}\n  if(!m||d<1||d>31||mo<1||mo>12)return null;\n  if(!y)y=mo>refM?refY-1:refY;\n  return y+'-'+p2(mo)+'-'+p2(d);\n}\nfunction formatoEC(txt){const m=txt.match(/cuenta:?\\s*(\\d{5,})/i)||txt.match(/\\b(\\d{6,})-\\d\\b/);return m?m[1]:'x'}\nfunction parseEstado(lines,fname){\n  const txt=lines.map(l=>l.items.map(i=>i.s).join(' ')).join('\\n');\n  // fecha de cierre (se usa para cargos sin fecha, como seguros)\n  let cierre=null;\n  const mc=txt.match(/fecha de cierre:?\\s*(\\d{1,2})[\\/\\-.](\\d{1,2})[\\/\\-.](\\d{2,4})/i)||txt.match(/\\b(\\d{1,2})\\/(\\d{1,2})\\/(\\d{2,4})\\b/);\n  if(mc){let y=+mc[3];if(y<100)y+=2000;cierre=y+'-'+p2(mc[2])+'-'+p2(mc[1])}\n  const refY=cierre?+cierre.slice(0,4):+hoy().slice(0,4),refM=cierre?+cierre.slice(5,7):+hoy().slice(5,7);\n  // columnas de moneda\n  let colU=null,colD=null;\n  for(const l of lines){const a=l.items.find(i=>/^(\\$|\\$U|UYU|PESOS|M\\/N|IMPORTE \\$)$/i.test(i.s)),b=l.items.find(i=>/^(U\\$S|US\\$|USD|D[O\u00d3]LARES|IMPORTE U\\$S)$/i.test(i.s));if(a&&b&&b.x>a.x){colU=a.x+a.w/2;colD=b.x+b.w/2;break}}\n  // bordes derechos t\u00edpicos de cada columna, tomados de las l\u00edneas que traen las dos monedas\n  const rU=[],rD=[];\n  for(const l of lines){\n    const lin=l.items.map(i=>i.s).join(' ').trim();\n    if(!(/^saldo|^total/i.test(lin)||fechaTok((l.items[0]||{s:''}).s.trim(),2000,1)))continue;\n    const t=l.items.filter(i=>RX_IMP.test(i.s.trim())&&/\\d/.test(i.s));\n    if(t.length===2){rU.push(t[0].x+t[0].w);rD.push(t[1].x+t[1].w)}}\n  const prom=a=>a.length?a.reduce((x,y)=>x+y,0)/a.length:null;\n  const bordeU=prom(rU),bordeD=prom(rD);\n  const items=[];let enTabla=false,tarjeta='';\n  for(const l of lines){\n    const linea=l.items.map(i=>i.s).join(' ').replace(/\\s+/g,' ').trim();\n    if(RX_INI.test(linea))enTabla=true;\n    const tj=linea.match(/^tarjeta\\s*(?:no:?\\s*)?([\\d*]{4,})\\s*(?:-\\s*(.+))?$/i);if(tj){tarjeta=tj[1].slice(-4)+(tj[2]?' '+tj[2].trim():'');continue}\n    if(RX_FIN.test(linea)){enTabla=false;continue}\n    let tk=[];for(const it of l.items){const parts=it.s.split(/\\s{2,}/);let x=it.x;for(const s of parts){tk.push({s:s.trim(),x,w:it.w/parts.length});x+=it.w/parts.length}}\n    tk=tk.filter(t=>t.s);\n    if(tk.length>=3&&/^\\d{2}$/.test(tk[0].s)&&/^\\d{2}$/.test(tk[1].s)&&/^\\d{2,4}$/.test(tk[2].s))tk.splice(0,3,{s:tk[0].s+' '+tk[1].s+' '+tk[2].s,x:tk[0].x,w:0});\n    // importes: la secuencia de n\u00fameros al final de la l\u00ednea\n    const amts=[];let k=tk.length-1;\n    while(k>=0){const s=tk[k].s;\n      if(RX_IMP.test(s)&&/\\d/.test(s)){amts.unshift({...tk[k]});k--;continue}\n      if(s==='-'&&amts.length&&!amts[0].neg){amts[0].neg=true;k--;continue}\n      break}\n    if(!amts.length)continue;\n    let fi=-1,f=null;\n    for(let i=0;i<Math.min(2,k+1);i++){f=fechaTok(tk[i].s,refY,refM);if(f){fi=i;break}}\n    if(!f){if(!enTabla||!cierre)continue;f=cierre}\n    let i0=fi+1;if(tk[i0]&&/^\\d{4,}$/.test(tk[i0].s)&&i0<k)i0++;   // c\u00f3digo de comercio (Ita\u00fa)\n    let d=tk.slice(i0,k+1).map(t=>t.s).join(' ').replace(/\\s+/g,' ').trim();\n    if(!d)continue;\n    let q='';\n    const cq=d.match(/(?:\\bcuota\\s*)?(\\d{1,2})\\s*\\/\\s*(\\d{1,2})\\s*$/i);\n    if(cq&&+cq[1]>=1&&+cq[1]<=+cq[2]&&+cq[2]>1&&+cq[2]<=60){q=p2(+cq[1])+'/'+p2(+cq[2]);d=d.slice(0,cq.index).trim()}\n    d=d.replace(/\\s*cuota\\s*$/i,'').replace(/(\\d)cuota$/i,'$1').trim();\n    if(RX_SKIP.test(d))continue;\n    amts.forEach((a,ix)=>{\n      let v=parseNum(a.s.replace(/CR/i,'').replace(/[()]/g,'').replace(/\\s/g,''));\n      if(a.neg||/^\\(|\\)$|-$|CR$/i.test(a.s)||a.s.trim().startsWith('-'))v=-Math.abs(v);\n      if(!v||isNaN(v))return;\n      let mon='UYU';\n      if(amts.length>=2)mon=ix===amts.length-1?'USD':'UYU';\n      else if(bordeU!=null&&bordeD!=null&&Math.abs(bordeD-bordeU)>8){const r=a.x+a.w;mon=Math.abs(r-bordeD)<Math.abs(r-bordeU)?'USD':'UYU'}\n      else if(colU!=null&&colD!=null){const r=a.x+a.w;mon=r>colD-2?'USD':'UYU'}\n      else if(/U\\$S|US\\$|\\bUSD\\b/i.test(linea))mon='USD';\n      const esPago=v<0&&/\\bpagos?\\b|su pago|pago recibido|pago efectuado/i.test(d);\n      items.push({f,d,m:r2(v),mon,q,usar:!esPago,pago:esPago,tj:tarjeta});\n    });\n  }\n  const fs=items.filter(x=>!x.q||x.q.startsWith('01/')).map(x=>x.f).sort();\n  const T=palabrasC(txt+' '+fname);\n  const cand=S.cfg.cuentas.filter(c=>!c.arch&&c.mon==='UYU'&&/^tc\\b|tarjeta|cr[e\u00e9]dito|visa|master/i.test(norm(c.n))&&['brou','bbva','santander','itau','scotiabank','hsbc','oca','prex'].some(b=>T.includes(' '+b+' ')&&palabrasC(c.n).includes(' '+b+' ')));\n  const fmtId=formatoEC(txt);\n  let cuenta=cand.length===1?cand[0].n:'';\n  const recordada=lsGet('cg_ec_'+fmtId);if(!cuenta&&recordada&&acc(recordada))cuenta=recordada;\n  // per\u00edodo: del d\u00eda siguiente al cierre anterior hasta el cierre\n  const desde=cierre?addDias(addMeses(cierre,-1),1):(fs[0]||'');\n  return {items,desde,hasta:cierre||fs[fs.length-1]||'',cuenta,fname,fmt:fmtId};\n}\nfunction renderEC(){\n  const E=EC;\n  const accO=S.cfg.cuentas.filter(c=>!c.arch&&c.mon==='UYU').map(c=>`<option value=\"${h(c.n)}\"${c.n===E.cuenta?' selected':''}>${h(c.n)}</option>`).join('');\n  const usados=E.items.filter(x=>x.usar);\n  const tot=m=>usados.filter(x=>x.mon===m).reduce((s,x)=>s+x.m,0);\n  modal(`<h3>Estado de cuenta: ${h(E.fname)}</h3>\n   <div class=\"fld\"><label>Tarjeta (cuenta en pesos; los d\u00f3lares van a su cuenta \u201chermana\u201d en d\u00f3lares)</label><select id=\"ec_c\" onchange=\"EC.cuenta=this.value\"><option value=\"\">\u2014 elegir \u2014</option>${accO}</select></div>\n   <div class=\"fld2\"><div class=\"fld\"><label>Desde</label><input type=\"date\" id=\"ec_d\" value=\"${E.desde}\" onchange=\"EC.desde=this.value\"></div><div class=\"fld\"><label>Hasta</label><input type=\"date\" id=\"ec_h\" value=\"${E.hasta}\" onchange=\"EC.hasta=this.value\"></div></div>\n   <div class=\"sm\" style=\"margin-bottom:6px\"><b>${E.items.length}</b> l\u00edneas le\u00eddas \u00b7 se van a comparar <b>${usados.length}</b> \u00b7 Total $ <span class=\"num\">${NF.format(tot('UYU'))}</span> \u00b7 U$S <span class=\"num\">${NF.format(tot('USD'))}</span></div>\n   <div style=\"max-height:40vh;overflow:auto;border:1px solid var(--ln);border-radius:10px\"><table class=\"pv\"><tr><th></th><th>Fecha</th><th>Descripci\u00f3n</th><th>Cuota</th><th style=\"text-align:right\">Importe</th></tr>\n   ${E.items.map((x,i)=>`<tr><td><input type=\"checkbox\"${x.usar?' checked':''} onchange=\"EC.items[${i}].usar=this.checked\"></td><td>${x.f.slice(8)}/${x.f.slice(5,7)}</td><td title=\"${h(x.d)}\">${h(x.d)}${x.pago?' <span class=\"chip\">pago</span>':''}${x.tj&&E.items.some(y=>y.tj&&y.tj!==x.tj)?' <span class=\"chip\" style=\"background:var(--bg);color:var(--mu)\">'+h(x.tj.split(' ')[0])+'</span>':''}</td><td>${h(x.q)}</td><td class=\"num\" style=\"text-align:right\">${sym(x.mon)} ${NF.format(x.m)}</td></tr>`).join('')}</table></div>\n   <p class=\"xs mu\">Los pagos de la tarjeta vienen sin marcar porque ya son transferencias. Revis\u00e1 que las l\u00edneas y los importes est\u00e9n bien antes de seguir.</p>\n   <div class=\"btns\" style=\"justify-content:flex-end\"><button class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\" onclick=\"conciliar()\">Conciliar</button></div>`);\n}\nfunction valorEn(x,acct){ // positivo = cargo en la tarjeta\n  if(x.t==='G')return x.c===acct?x.m:0;\n  if(x.t==='I')return x.c===acct?-x.m:0;\n  if(x.t==='A')return x.c===acct?-x.m:0;\n  if(x.t==='T'){if(x.c2===acct)return -(x.m2??x.m);if(x.c===acct)return x.m}\n  return 0;\n}\nfunction dias(a,b){return Math.abs((new Date(a+'T12:00')-new Date(b+'T12:00'))/864e5)}\nfunction conciliar(){\n  const E=EC;E.cuenta=document.getElementById('ec_c').value;if(E.cuenta)lsSet('cg_ec_'+E.fmt,E.cuenta);E.desde=document.getElementById('ec_d').value;E.hasta=document.getElementById('ec_h').value;\n  if(!E.cuenta){toast('Eleg\u00ed la tarjeta');return}\n  const A=acc(E.cuenta),B=(()=>{const base=n=>palabrasC(n).replace(/ (pesos|dolares|uyu|usd|u s) /g,' ').trim();return S.cfg.cuentas.find(c=>c.mon==='USD'&&base(c.n)===base(A.n))})();\n  const ec='ec'+hash(E.cuenta+'|'+E.desde+'|'+E.hasta);\n  const lin=E.items.filter(x=>x.usar).map((x,i)=>({...x,i,acct:x.mon==='USD'?(B?B.n:A.n):A.n}));\n  const ini=addDias(E.desde,-10),fin=addDias(E.hasta,10);\n  const cands=S.movs.filter(x=>x.f>=ini&&x.f<=fin&&(!x.cc||x.cc===ec)).flatMap(x=>[A.n,B&&B.n].filter(Boolean).map(a=>({x,a,v:valorEn(x,a)})).filter(o=>o.v));\n  const usado=new Set(),match=[],difs=[];\n  const esperado=l=>{const q=(l.q||'').split('/').map(Number);return q[1]>1&&q[0]>1?addMeses(l.f,q[0]-1):l.f};\n  // 1) mismo importe y fecha cercana\n  for(const l of lin){let best=null;\n    for(const c of cands){if(usado.has(c)||c.a!==l.acct||Math.abs(c.v-l.m)>0.011)continue;\n      let dd=Math.min(dias(c.x.f,l.f),dias(c.x.f,esperado(l)));\n      // movimientos importados de ZetaCuentas sin d\u00eda (fecha 1 del mes): alcanza con que sea el mismo mes\n      if(/^z/.test(c.x.id)&&c.x.f.endsWith('-01')&&(c.x.f.slice(0,7)===l.f.slice(0,7)||c.x.f.slice(0,7)===esperado(l).slice(0,7)))dd=Math.min(dd,9.5);\n      if(dd<=10&&(!best||dd<best.dd))best={c,dd}}\n    if(best){usado.add(best.c);l.ok=true;match.push(best.c.x.id)}}\n  // 2) misma fecha, importe parecido o mismo comercio\n  for(const l of lin){if(l.ok)continue;let best=null;\n    for(const c of cands){if(usado.has(c)||c.a!==l.acct||Math.sign(c.v)!==Math.sign(l.m))continue;const dd=Math.min(dias(c.x.f,l.f),dias(c.x.f,esperado(l)));if(dd>4)continue;\n      const pal=palabrasC(l.d).split(' ').filter(w=>w.length>3),comun=pal.some(w=>palabrasC((c.x.d||'')+' '+(c.x.k||'')).includes(' '+w+' '));\n      const imp=/^z/.test(c.x.id)&&c.x.f.endsWith('-01')&&(c.x.f.slice(0,7)===l.f.slice(0,7)||c.x.f.slice(0,7)===esperado(l).slice(0,7));\n      if(imp&&Math.abs(c.v-l.m)/Math.abs(l.m)<=0.015){const sc=0.5;if(!best||sc<best.sc)best={c,sc};continue}\n      if(comun||(!l.q&&!c.x.q&&Math.abs(c.v-l.m)/Math.abs(l.m)<=0.10&&dd<=2)){const sc=dd+(comun?0:3);if(!best||sc<best.sc)best={c,sc}}}\n    if(best){usado.add(best.c);l.ok=true;difs.push({tipo:'importe',f:l.f,d:l.d,m:l.m,mon:l.mon,mov:best.c.x.id,mApp:Math.abs(best.c.v),q:l.q,c:l.acct})}}\n  for(const l of lin)if(!l.ok)difs.push({tipo:'falta_app',f:l.f,d:l.d,m:l.m,mon:l.mon,q:l.q,c:l.acct});\n  for(const c of cands)if(!usado.has(c)&&c.x.f>=E.desde&&c.x.f<=E.hasta&&!c.x.cc&&!(/^z/.test(c.x.id)&&c.x.f.endsWith('-01')))difs.push({tipo:'falta_estado',f:c.x.f,d:c.x.d||c.x.k||titulo(c.x),m:c.v,mon:monOf(c.a),mov:c.x.id,q:c.x.q||'',c:c.a});\n  const n={falta_app:0,falta_estado:0,importe:0};difs.forEach(d=>n[d.tipo]++);\n  E.res={ec,match,difs};\n  modal(`<h3>Resultado de la conciliaci\u00f3n</h3>\n   <div class=\"kv\"><span>Coinciden</span><b class=\"pos\">${match.length}</b></div>\n   <div class=\"kv\"><span>En el estado y no en la app</span><b>${n.falta_app}</b></div>\n   <div class=\"kv\"><span>En la app y no en el estado</span><b>${n.falta_estado}</b></div>\n   <div class=\"kv\"><span>Importe distinto</span><b>${n.importe}</b></div>\n   <p class=\"sm mu\">Las coincidencias quedan marcadas como conciliadas. Las diferencias van a <b>Por clasificar</b> para que apruebes o rechaces cada una.</p>\n   <div class=\"btns\" style=\"justify-content:flex-end\"><button class=\"btn\" onclick=\"renderEC()\">\u2039 Volver</button><button class=\"btn pri\" id=\"btnEC\" onclick=\"guardarEC()\">Guardar</button></div>`);\n}\nfunction addDias(f,n){const d=new Date(f+'T12:00');d.setDate(d.getDate()+n);return d.getFullYear()+'-'+p2(d.getMonth()+1)+'-'+p2(d.getDate())}\nasync function guardarEC(){\n  const {ec,match,difs}=EC.res;document.getElementById('btnEC').disabled=true;\n  S.movs.forEach(x=>{if(match.includes(x.id))x.cc=ec});\n  S.difs=S.difs.filter(d=>d.ec!==ec).concat(difs.map(d=>({...d,id:'tmp'+uid(),ec})));\n  closeModal();go('clasificar');\n  try{await op('guardarConciliacion',JSON.stringify({ec,cuenta:EC.cuenta,conciliados:match,difs}));load()}catch(e){}\n  toast(`Conciliado: ${match.length} coinciden, ${difs.length} diferencias`);\n}\n\n/* ---------- Versi\u00f3n y actualizaciones ---------- */\nfunction vVersion(){\n  const A=S.app||{};\n  return `<div class=\"card\" style=\"margin-top:12px\"><h2>Versi\u00f3n de la app</h2>\n   <div class=\"kv\"><span>Instalada</span><b>${h(A.version||'?')}</b></div>\n   <div class=\"kv\"><span>Publicada</span><b>${h(A.remota||'\u2014')}</b></div>\n   ${A.errorRemota?`<div class=\"xs neg\">No pude consultar las versiones publicadas: ${h(A.errorRemota)}</div>`:''}\n   ${A.hay?`<div class=\"nt\"><b>Hay una versi\u00f3n nueva.</b> ${h(A.notas||'')}</div>`:'<div class=\"sm pos\" style=\"margin:4px 0 8px\">\u2713 Est\u00e1s en la \u00faltima versi\u00f3n</div>'}\n   ${A.ultima?`<div class=\"xs mu\">\u00daltima actualizaci\u00f3n: ${h(A.ultima)}</div>`:''}\n   ${A.error?`<div class=\"xs neg\">\u00daltimo error al actualizar: ${h(A.error)}</div>`:''}\n   <div class=\"fld\" style=\"margin-top:10px\"><label><input type=\"checkbox\"${A.auto!==false?' checked':''} onchange=\"S.app.auto=this.checked;op('setConfig','auto_actualizar',this.checked?'si':'no').catch(()=>{});toast('Guardado')\"> Instalar las versiones nuevas autom\u00e1ticamente (todas las noches)</label></div>\n   <div class=\"btns\"><button class=\"btn\" onclick=\"buscarAct()\">Buscar actualizaciones</button>${A.hay?'<button class=\"btn pri\" id=\"btnAct\" onclick=\"actualizarYa()\">Actualizar ahora</button>':''}</div>\n   <div id=\"actOut\" class=\"sm\" style=\"margin-top:8px\"></div>\n  </div>`;\n}\nasync function buscarAct(){\n  try{S.app=JSON.parse(await call('buscarActualizacion'));render();toast(S.app.errorRemota?'No pude consultar GitHub: '+S.app.errorRemota:S.app.hay?'Hay una versi\u00f3n nueva':'Est\u00e1s en la \u00faltima versi\u00f3n')}catch(e){toast('No pude buscar: '+(e.message||e))}\n}\nasync function actualizarYa(){\n  const b=document.getElementById('btnAct'),out=document.getElementById('actOut');b.disabled=true;b.textContent='Actualizando\u2026';\n  try{const r=JSON.parse(await call('actualizarApp'));\n    out.innerHTML=r.nada?'Ya estabas en la \u00faltima versi\u00f3n.':`<span class=\"pos\">\u2713 Actualizada a la versi\u00f3n ${h(r.version)}.</span> <a href=\"${h(S.url)}\" target=\"_top\">Toc\u00e1 ac\u00e1 para recargar la app</a>.`;\n  }catch(e){b.disabled=false;b.textContent='Actualizar ahora';out.innerHTML='<span class=\"neg\">'+h(e.message||e)+'</span>'}\n}\n\n/* ---------- Ajustes: notificaciones y reglas ---------- */\nfunction vNotifCfg(){\n  const rc=S.reglas.filter(r=>r.tipo==='Concepto'),ra=S.reglas.filter(r=>r.tipo==='Cuenta');\n  const byV={};rc.forEach(r=>(byV[r.valor]||(byV[r.valor]=[])).push(r.palabra));\n  const AV=S.avisos||{};\n  const apk='https://github.com/sebastiandemarcoz-glitch/controlgastos/releases/latest/download/ControlGastos-Avisos.apk';\n  const link=AV.receptor?'intent://config?u='+encodeURIComponent(AV.receptor)+'&k='+encodeURIComponent(S.token)+'#Intent;scheme=cgavisos;package=uy.cg.avisos;end':'';\n  return `<div class=\"card\" style=\"margin-top:12px\"><h2>Avisos de gastos</h2>\n   <p class=\"sm\" style=\"margin-top:0\">La app se entera sola de tus gastos por <b>mail</b>, <b>SMS</b> y <b>notificaciones</b> de las apps de los bancos. Asigna las que reconoce y deja el resto en <b>Por clasificar</b>. Si el mismo gasto llega por dos medios, queda uno solo.</p>\n   <div class=\"catH\">1 \u00b7 Mails (Gmail)</div>\n   <div class=\"fld\"><label><input type=\"checkbox\"${AV.gmail?' checked':''} onchange=\"activarMail(this.checked)\"> Leer los mails de aviso de mis bancos</label></div>\n   ${AV.gmail?`<div class=\"xs ${/error/.test(AV.gmailEstado)?'neg':'mu'}\">${/error/.test(AV.gmailEstado)?'No pude leer Gmail: '+h(AV.gmailEstado.replace(/^error: /,''))+'. Abr\u00ed la app de nuevo y acept\u00e1 el permiso de Gmail que te pide Google.':'\u00daltimo control: '+h((AV.gmailEstado||'todav\u00eda no').replace(/^ok /,''))}</div>`:''}\n   <div class=\"fld\" style=\"margin-top:6px\"><label>Remitentes (palabras del mail del banco)</label><input type=\"text\" id=\"av_rem\" value=\"${h(AV.remitentes||'')}\" onchange=\"activarMail(${AV.gmail?'true':'false'},this.value)\"></div>\n   <div class=\"xs mu\">Revis\u00e1 en el home banking de cada banco que tengas activados los avisos de compra por mail.</div>\n   <div class=\"catH\" style=\"margin-top:14px\">2 \u00b7 SMS y notificaciones (app de Android)</div>\n   <ol class=\"sm\" style=\"padding-left:18px;margin:6px 0\">\n    <li>Desde el celular, <a href=\"${apk}\" target=\"_blank\" rel=\"noopener\">descarg\u00e1 la app \u201cGastos \u00b7 Avisos\u201d</a> e instalala. Si Android pregunta, permit\u00ed instalar apps de esta fuente.</li>\n    <li>${link?`Toc\u00e1 <a class=\"btn sm pri\" href=\"${h(link)}\" target=\"_top\" style=\"text-decoration:none\">Configurar la app de avisos</a> (abre la app ya conectada).`:'Abr\u00ed la app y peg\u00e1 la direcci\u00f3n del receptor y la clave de abajo.'}</li>\n    <li>En la app, toc\u00e1 <b>Abrir permiso de notificaciones</b> y <b>Permitir SMS</b>. Si dice \u201cconfiguraci\u00f3n restringida\u201d, segu\u00ed la ayuda que muestra la app.</li>\n   </ol>\n   <div class=\"xs mu\">Direcci\u00f3n del receptor: ${AV.receptor?`<code class=\"tok\">${h(AV.receptor)}</code> <button class=\"btn sm\" onclick=\"copiar(${js(AV.receptor)})\">Copiar</button>`:`<input type=\"text\" id=\"av_url\" placeholder=\"https://script.google.com/macros/s/\u2026/exec (implementaci\u00f3n con acceso Cualquier persona)\" style=\"margin-top:4px\"><button class=\"btn sm\" style=\"margin-top:4px\" onclick=\"guardarReceptor()\">Guardar</button>`}</div>\n   <div class=\"xs mu\" style=\"margin-top:4px\">Clave: <code class=\"tok\">${h(S.token)}</code> <button class=\"btn sm\" onclick=\"copiar(${js(S.token)})\">Copiar</button></div>\n   <div class=\"sm\" style=\"margin-top:10px\">\u00daltimo aviso recibido: <b>${h(AV.ultimo||'ninguno todav\u00eda')}</b></div>\n   ${AV.receptor?`<div class=\"btns\" style=\"margin-top:8px\"><button class=\"btn\" id=\"btnPR\" onclick=\"probarRec()\">Probar el receptor</button><button class=\"btn sm\" onclick=\"S.avisos.receptor='';render()\">Cambiar direcci\u00f3n</button></div><div id=\"prOut\" class=\"sm\" style=\"margin-top:6px\"></div>`:''}\n   <div class=\"catH\" style=\"margin-top:14px\">Opciones</div>\n   <div class=\"fld\"><label>Preguntar en cu\u00e1ntas cuotas fue la compra</label><select onchange=\"S.cfg.preguntarCuotas=this.value;op('setConfig','preguntar_cuotas',this.value).catch(()=>{});toast('Guardado')\">\n     <option value=\"tc\"${S.cfg.preguntarCuotas!=='todos'?' selected':''}>Solo en compras con tarjeta de cr\u00e9dito</option><option value=\"todos\"${S.cfg.preguntarCuotas==='todos'?' selected':''}>En todos los gastos que entran solos</option></select></div>\n   <div class=\"fld\" style=\"margin-top:12px\"><label>Probar una notificaci\u00f3n (peg\u00e1 el texto como llega del banco)</label>\n    <input type=\"text\" id=\"pr_app\" placeholder=\"App (ej. BBVA)\" style=\"margin-bottom:6px\"><textarea id=\"pr_txt\" placeholder=\"Compra aprobada por $ 1.250,00 en DISCO 15 con tu tarjeta terminada en 1234\"></textarea></div>\n   <button class=\"btn\" onclick=\"probar()\">Probar</button><div id=\"pr_out\" class=\"sm\" style=\"margin-top:8px\"></div>\n  </div>\n  <div class=\"grid2\" style=\"margin-top:12px\">\n   <div class=\"card\"><h2>Reglas de concepto <button class=\"btn sm\" onclick=\"nuevaRegla('Concepto')\">+ Agregar</button></h2>\n    <p class=\"xs mu\" style=\"margin-top:0\">Si la notificaci\u00f3n contiene la palabra (completa), se usa ese concepto.</p>\n    ${Object.keys(byV).sort((a,b)=>a.localeCompare(b,'es')).map(v=>`<div style=\"padding:6px 0;border-bottom:1px solid var(--ln)\"><b class=\"sm\">${h(v)}</b><div>${byV[v].map(w=>`<span class=\"chip\" style=\"margin:3px 4px 0 0;cursor:pointer\" title=\"Quitar\" onclick=\"borrarRegla('Concepto',${js(w)},'')\">${h(w)} \u2715</span>`).join('')}</div></div>`).join('')||'<div class=\"empty\">Sin reglas</div>'}\n   </div>\n   <div class=\"card\"><h2>Reglas de cuenta <button class=\"btn sm\" onclick=\"nuevaRegla('Cuenta')\">+ Agregar</button></h2>\n    <p class=\"xs mu\" style=\"margin-top:0\">Por app y/o palabra (por ejemplo, los 4 \u00faltimos d\u00edgitos de la tarjeta). Si el gasto es en otra moneda se usa la cuenta \u201chermana\u201d (Pesos \u2194 D\u00f3lares).</p>\n    ${ra.map(r=>`<div class=\"row\" style=\"cursor:default\"><div class=\"mid\"><div class=\"t1\">${h(r.valor)}</div><div class=\"t2\">${r.app?'App: '+h(r.app):''}${r.app&&r.palabra?' \u00b7 ':''}${r.palabra?'Contiene: '+h(r.palabra):''}</div></div><button class=\"btn sm dan\" onclick=\"borrarRegla('Cuenta',${js(r.palabra)},${js(r.app||'')})\">Quitar</button></div>`).join('')||'<div class=\"empty\">Se aprenden al clasificar con \u201cRecordar la cuenta\u201d.</div>'}\n   </div>\n  </div>`;\n}\nasync function activarMail(si,rem){\n  S.avisos.gmail=si;if(rem!=null)S.avisos.remitentes=rem;render();\n  await op('activarGmail',si,rem==null?null:rem).catch(()=>{});\n  toast(si?'Listo: desde ahora la app lee los mails de tus bancos':'Lectura de mails desactivada');\n  if(si)load();\n}\nasync function probarRec(){\n  const b=document.getElementById('btnPR'),o=document.getElementById('prOut');b.disabled=true;b.textContent='Probando\u2026';\n  try{const r=JSON.parse(await call('probarReceptor'));o.innerHTML=`<span class=\"${r.ok?'pos':'neg'}\">${r.ok?'\u2713 ':'\u2717 '}${h(r.msg)}</span>`;toast(r.ok?'El receptor funciona':'El receptor tiene un problema')}\n  catch(e){o.innerHTML='<span class=\"neg\">'+h(e.message||e)+'</span>'}\n  b.disabled=false;b.textContent='Probar el receptor';\n}\nasync function guardarReceptor(){const u=document.getElementById('av_url').value.trim();if(!/^https:\\/\\/script\\.google\\.com\\//.test(u)){toast('Peg\u00e1 la direcci\u00f3n que termina en /exec');return}\n  S.avisos.receptor=u;await op('guardarUrlReceptor',u).catch(()=>{});render();toast('Guardado')}\nasync function copiar(t){try{await navigator.clipboard.writeText(t);toast('Copiado')}catch(e){toast('Manten\u00e9 apretado el texto para copiarlo')}}\nasync function probar(){\n  const app=document.getElementById('pr_app').value,txt=document.getElementById('pr_txt').value;if(!txt.trim())return;\n  const out=document.getElementById('pr_out');out.textContent='Probando\u2026';\n  try{const p=JSON.parse(await call('probarNotif',app,txt));\n    out.innerHTML=!p.esMov?'<span class=\"aj\">No encontr\u00e9 un importe: se registrar\u00eda como \u201cno es gasto\u201d.</span>':\n     `Importe: <b>${fmt(p.importe,p.mon)}</b> \u00b7 Tipo: <b>${{G:'Gasto',I:'Ingreso',T:'Transferencia'}[p.tipo]}</b> \u00b7 Comercio: <b>${h(p.comercio||'\u2014')}</b>${p.digitos?' \u00b7 Tarjeta: <b>'+h(p.digitos)+'</b>':''}<br>\n      Cuenta: <b>${h(p.cuenta||'\u2014 (preguntar\u00eda)')}</b> \u00b7 Concepto: <b>${h(p.concepto||(p.tipo==='T'?'\u2014':'\u2014 (preguntar\u00eda)'))}</b><br>\n      ${p.cuenta&&p.concepto&&p.tipo!=='T'?'<span class=\"pos\">\u2713 Se asignar\u00eda autom\u00e1ticamente</span>':'<span class=\"aj\">Quedar\u00eda en Por clasificar</span>'}`;\n  }catch(e){out.textContent='Error: '+(e.message||e)}\n}\nfunction nuevaRegla(tipo){\n  const opts=tipo==='Concepto'?sortN(S.cfg.conceptos).map(c=>`<option value=\"${h(c.n)}\">${h(c.n)}${c.tipo==='I'?' (ingreso)':''}</option>`).join(''):S.cfg.cuentas.filter(c=>!c.arch).map(c=>`<option value=\"${h(c.n)}\">${h(c.n)}</option>`).join('');\n  modal(`<h3>Nueva regla de ${tipo.toLowerCase()}</h3><form onsubmit=\"event.preventDefault();guardarRegla('${tipo}')\">\n   <div class=\"fld\"><label>${tipo==='Concepto'?'Palabra del comercio (ej. disco, netflix)':'Palabra o \u00faltimos 4 d\u00edgitos de la tarjeta (opcional)'}</label><input type=\"text\" id=\"rg_p\"${tipo==='Concepto'?' required':''}></div>\n   ${tipo==='Cuenta'?'<div class=\"fld\"><label>Nombre de la app que manda la notificaci\u00f3n (opcional, ej. BBVA)</label><input type=\"text\" id=\"rg_a\"></div>':''}\n   <div class=\"fld\"><label>${tipo}</label><select id=\"rg_v\">${opts}</select></div>\n   <div class=\"btns\" style=\"justify-content:flex-end\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\">Guardar</button></div></form>`);\n}\nfunction guardarRegla(tipo){\n  const palabra=document.getElementById('rg_p').value.trim(),app=tipo==='Cuenta'?document.getElementById('rg_a').value.trim():'',valor=document.getElementById('rg_v').value;\n  if(!palabra&&!app){toast('Pon\u00e9 una palabra o una app');return}\n  S.reglas=S.reglas.filter(x=>!(x.tipo===tipo&&norm(x.palabra)===norm(palabra)&&norm(x.app||'')===norm(app)));S.reglas.push({tipo,palabra,app,valor});\n  op('saveReglas',JSON.stringify(S.reglas)).then(()=>reprocesar(true)).catch(()=>{});closeModal();render();toast('Regla guardada');\n}\nfunction borrarRegla(tipo,palabra,app){\n  S.reglas=S.reglas.filter(x=>!(x.tipo===tipo&&x.palabra===palabra&&(x.app||'')===app));\n  op('saveReglas',JSON.stringify(S.reglas)).catch(()=>{});render();\n}\n\n/* ===================== Modal / confirmar ===================== */\nfunction modal(html){document.getElementById('sheet').innerHTML=html;document.getElementById('modal').hidden=false}\nfunction closeModal(){document.getElementById('modal').hidden=true;F=null}\nlet _cb=null;\nfunction confirmar(msg,cb){_cb=cb;modal(`<h3>Confirmar</h3><p>${h(msg)}</p><div class=\"btns\" style=\"justify-content:flex-end\"><button class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\" onclick=\"_cb()\">S\u00ed, continuar</button></div>`)}\ndocument.addEventListener('keydown',e=>{if(e.key==='Escape'&&!document.getElementById('modal').hidden)closeModal()});\n\n/* ===================== Alta / edici\u00f3n de movimientos ===================== */\nfunction openMov(id){\n  if(!S.cfg)return;\n  if(!S.cfg.cuentas.length){toast('Primero cre\u00e1 una cuenta');go('cuentas');return}\n  const act=S.cfg.cuentas.filter(c=>!c.arch);\n  const last=lsGet('cg_lastacc');\n  const defAcc=(V.tab==='movimientos'&&V.f.c)||(act.find(c=>c.n===last)?last:(act[0]||S.cfg.cuentas[0]).n);\n  if(id){const x=S.movs.find(y=>y.id===id);if(!x)return;F=clone(x)}\n  else F={t:'G',f:hoy(),c:defAcc,c2:'',k:'',e:'',d:'',m:''};\n  renderForm();\n  setTimeout(()=>{const el=document.getElementById('fm_m');if(el&&!id)el.focus()},60);\n}\nfunction readForm(){\n  if(!F)return;\n  for(const [k,i] of [['m','fm_m'],['f','fm_f'],['c','fm_c'],['c2','fm_c2'],['k','fm_k'],['e','fm_e'],['d','fm_d'],['m2','fm_m2'],['n','fm_n']]){const el=document.getElementById(i);if(el)F[k]=el.value}\n}\nfunction setTipo(t){readForm();F.t=t;if(F.k&&!con(t,F.k))F.k='';renderForm()}\nfunction renderForm(){\n  const x=F,act=S.cfg.cuentas.filter(c=>!c.arch||c.n===x.c||c.n===x.c2);\n  const accOpts=(sel,blank)=>(blank?'<option value=\"\">\u2014 elegir \u2014</option>':'')+act.map(c=>`<option value=\"${h(c.n)}\"${c.n===sel?' selected':''}>${h(c.n)} (${sym(c.mon)})</option>`).join('');\n  const cats={};for(const c of S.cfg.conceptos.filter(c=>c.tipo===x.t))(cats[c.cat||'Sin categor\u00eda']||(cats[c.cat||'Sin categor\u00eda']=[])).push(c);\n  const kOpts='<option value=\"\">\u2014 elegir concepto \u2014</option>'+Object.keys(cats).sort((a,b)=>a.localeCompare(b,'es')).map(cat=>`<optgroup label=\"${h(cat)}\">${sortN(cats[cat]).map(c=>`<option value=\"${h(c.n)}\"${c.n===x.k?' selected':''}>${h(c.n)}</option>`).join('')}</optgroup>`).join('');\n  const fmtIn=v=>v===''||v==null?'':(typeof v==='number'?String(v).replace('.',','):v);\n  const tags={};for(const y of S.movs)if(y.e)tags[y.e]=(tags[y.e]||0)+1;\n  const tagList=Object.entries(tags).sort((a,b)=>b[1]-a[1]).slice(0,80).map(([t])=>`<option value=\"${h(t)}\">`).join('');\n  const difMon=x.t==='T'&&x.c2&&monOf(x.c)!==monOf(x.c2);\n  const edit=!!x.id;\n  modal(`<h3>${edit?'Editar movimiento':'Nuevo movimiento'}</h3>\n  <form onsubmit=\"event.preventDefault();saveMov(false)\">\n  <div class=\"seg\" style=\"margin-bottom:12px\">${[['G','Gasto'],['I','Ingreso'],['T','Transferencia'],['A','Ajuste']].map(([k,l])=>`<button type=\"button\" class=\"${x.t===k?'on '+k:''}\" onclick=\"setTipo('${k}')\">${l}</button>`).join('')}</div>\n  ${x.g?`<div class=\"nt\">Cuota <b>${h(x.q)}</b> de una compra en cuotas. Los cambios se aplican solo a esta cuota. <button type=\"button\" class=\"btn sm dan\" onclick=\"delGrupo(${js(x.g)})\">Eliminar todas las cuotas</button></div>`:''}\n  <div class=\"fld\"><label><span id=\"fm_ml\">${x.t==='A'?'Importe (negativo para restar)':(x.t==='G'&&!edit&&+x.n>1?'Importe total':'Importe')}</span>${x.c?' en '+sym(monOf(x.c)):''}</label><input type=\"text\" inputmode=\"decimal\" class=\"amount\" id=\"fm_m\" value=\"${h(fmtIn(x.m))}\" placeholder=\"0,00\" autocomplete=\"off\"></div>\n  ${x.t==='T'?`\n   <div class=\"fld2\"><div class=\"fld\"><label>Desde</label><select id=\"fm_c\" onchange=\"readForm();renderForm()\">${accOpts(x.c)}</select></div>\n   <div class=\"fld\"><label>Hacia</label><select id=\"fm_c2\" onchange=\"readForm();renderForm()\">${accOpts(x.c2,true)}</select></div></div>\n   ${difMon?`<div class=\"fld\"><label>Importe recibido en ${sym(monOf(x.c2))}</label><input type=\"text\" inputmode=\"decimal\" id=\"fm_m2\" value=\"${h(fmtIn(x.m2))}\" placeholder=\"0,00\"></div>`:''}`\n  :`<div class=\"fld\"><label>Cuenta</label><select id=\"fm_c\" onchange=\"readForm();renderForm()\">${accOpts(x.c)}</select></div>`}\n  ${x.t==='G'||x.t==='I'?`<div class=\"fld\"><label>Concepto</label><select id=\"fm_k\">${kOpts}</select></div>`:''}\n  ${x.t==='G'&&!edit?`<div class=\"fld\"><label>Cuotas</label><input type=\"number\" id=\"fm_n\" min=\"1\" max=\"60\" value=\"${h(x.n||1)}\" inputmode=\"numeric\" oninput=\"document.getElementById('fm_ml').textContent=(+this.value>1?'Importe total':'Importe')\"><div class=\"xs mu\" style=\"margin-top:3px\">Con m\u00e1s de 1 cuota se crea un movimiento por mes, empezando en la fecha de la compra.</div></div>`:''}\n  <div class=\"fld2\"><div class=\"fld\"><label>Fecha</label><input type=\"date\" id=\"fm_f\" value=\"${h(x.f)}\"></div>\n  <div class=\"fld\"><label>Etiqueta</label><input type=\"text\" id=\"fm_e\" list=\"dl_tags\" value=\"${h(x.e||'')}\" autocomplete=\"off\"><datalist id=\"dl_tags\">${tagList}</datalist></div></div>\n  <div class=\"fld\"><label>Descripci\u00f3n</label><input type=\"text\" id=\"fm_d\" value=\"${h(x.d||'')}\" autocomplete=\"off\"></div>\n  <div class=\"btns\" style=\"justify-content:space-between;margin-top:6px\">\n   <div class=\"btns\">${edit?`<button type=\"button\" class=\"btn dan\" onclick=\"delMov()\">Eliminar</button><button type=\"button\" class=\"btn\" onclick=\"dupMov()\">Duplicar</button>`:''}</div>\n   <div class=\"btns\"><button type=\"button\" class=\"btn\" onclick=\"closeModal()\">Cancelar</button>\n   ${edit?'':'<button type=\"button\" class=\"btn\" onclick=\"saveMov(true)\">Guardar y otro</button>'}\n   <button class=\"btn pri\">Guardar</button></div></div>\n  </form>`);\n}\nfunction saveMov(otro){\n  readForm();const x=F;\n  let m=parseNum(x.m);\n  if(isNaN(m)||m===0){toast('Ingres\u00e1 un importe');return}\n  if(x.t!=='A')m=Math.abs(m);\n  if(!x.c){toast('Eleg\u00ed la cuenta');return}\n  if(x.t==='T'&&(!x.c2||x.c2===x.c)){toast('Eleg\u00ed cuentas distintas de origen y destino');return}\n  if((x.t==='G'||x.t==='I')&&!x.k){\n    const def=S.cfg.conceptos.find(c=>c.tipo===x.t&&/sin categor|otros ingresos/i.test(c.n));\n    if(def)x.k=def.n;else{toast('Eleg\u00ed un concepto');return}\n  }\n  const o={id:x.id||uid(),t:x.t,f:x.f||hoy(),c:x.c,m:r2(m),ts:x.ts||Date.now()};\n  if(x.t==='T'){o.c2=x.c2;if(monOf(x.c)!==monOf(x.c2)){const m2=parseNum(x.m2);if(!(m2>0)){toast('Ingres\u00e1 el importe recibido');return}o.m2=r2(m2)}}\n  if(x.t==='G'||x.t==='I')o.k=x.k;\n  if(x.e&&x.e.trim())o.e=x.e.trim();\n  if(x.d&&x.d.trim())o.d=x.d.trim();\n  if(x.g){o.g=x.g;o.q=x.q}if(x.cc)o.cc=x.cc;\n  const N=Math.max(1,Math.min(60,parseInt(x.n)||1));\n  if(!x.id&&o.t==='G'&&N>1){\n    const cs=crearCuotas(o,N,o.m);S.movs.push(...cs);invalidate();op('importMovs',JSON.stringify(cs));lsSet('cg_lastacc',o.c);\n    toast(`Se crearon ${N} cuotas de ${fmt(cs[0].m,monOf(o.c))}`);\n    if(otro){F={t:o.t,f:o.f,c:o.c,c2:'',k:'',e:'',d:'',m:'',n:1};renderForm()}else closeModal();\n    render();return;\n  }\n  const i=S.movs.findIndex(y=>y.id===o.id);if(i>=0)S.movs[i]=o;else S.movs.push(o);\n  invalidate();op('saveMov',JSON.stringify(o));lsSet('cg_lastacc',o.c);\n  toast('Guardado');\n  if(otro){F={t:o.t,f:o.f,c:o.c,c2:o.c2||'',k:'',e:'',d:'',m:''};renderForm();setTimeout(()=>document.getElementById('fm_m').focus(),60)}\n  else closeModal();\n  render();\n}\nfunction dupMov(){readForm();const c={...F};delete c.id;delete c.ts;delete c.g;delete c.q;delete c.cc;c.f=hoy();F=c;renderForm();toast('Copia: revis\u00e1 y guard\u00e1')}\nfunction addMeses(f,n){const ymx=perStep(f.slice(0,7),n);const max=new Date(+ymx.slice(0,4),+ymx.slice(5,7),0).getDate();return ymx+'-'+p2(Math.min(+f.slice(8,10),max))}\nfunction crearCuotas(base,N,total,desde){\n  desde=desde||1;const g='g'+uid(),cu=Math.floor(total/N*100)/100,out=[];\n  for(let k=desde;k<=N;k++){const m=k<N?cu:r2(total-cu*(N-1));const o={...base,id:g+'-'+k,f:addMeses(base.f,k-desde),m,g,q:k+'/'+N,ts:Date.now()+k};delete o.n;out.push(o)}\n  return out;\n}\nfunction delGrupo(g){\n  const ids=S.movs.filter(x=>x.g===g).map(x=>x.id);\n  confirmar(`\u00bfEliminar las ${ids.length} cuotas de esta compra?`,()=>{S.movs=S.movs.filter(x=>x.g!==g);invalidate();op('deleteMovs',JSON.stringify(ids));closeModal();render();toast('Cuotas eliminadas')});\n}\nfunction delMov(){\n  const id=F.id;confirmar('\u00bfEliminar este movimiento?',()=>{S.movs=S.movs.filter(y=>y.id!==id);invalidate();op('deleteMov',id);closeModal();render();toast('Eliminado')});\n}\n\n/* ===================== Importador ===================== */\nlet IMPF=null,IMP=null;\nfunction hash(s){let h1=0xdeadbeef,h2=0x41c6ce57;for(let i=0;i<s.length;i++){const c=s.charCodeAt(i);h1=Math.imul(h1^c,2654435761);h2=Math.imul(h2^c,1597334677)}\n  h1=Math.imul(h1^(h1>>>16),2246822507)^Math.imul(h2^(h2>>>13),3266489909);h2=Math.imul(h2^(h2>>>16),2246822507)^Math.imul(h1^(h1>>>13),3266489909);\n  return (4294967296*(2097151&h2)+(h1>>>0)).toString(36)}\nfunction parseFecha(v){\n  if(typeof v==='number'&&v>20000&&v<80000)return new Date(Math.round((v-25569)*864e5)).toISOString().slice(0,10);\n  if(v instanceof Date)return v.getFullYear()+'-'+p2(v.getMonth()+1)+'-'+p2(v.getDate());\n  const s=norm(v);let m;\n  if((m=s.match(/^(\\d{1,2})[\\/\\-.](\\d{1,2})[\\/\\-.](\\d{2,4})/))){let y=+m[3];if(y<100)y+=2000;return y+'-'+p2(m[2])+'-'+p2(m[1])}\n  if((m=s.match(/^(\\d{4})[\\/\\-.](\\d{1,2})[\\/\\-.](\\d{1,2})/)))return m[1]+'-'+p2(m[2])+'-'+p2(m[3]);\n  if((m=s.match(/^(\\d{4})[\\/\\-.](\\d{1,2})$/)))return m[1]+'-'+p2(m[2])+'-01';           // A\u00f1o/Mes\n  if((m=s.match(/^(\\d{1,2})[\\/\\-.](\\d{4})$/)))return m[2]+'-'+p2(m[1])+'-01';           // Mes/A\u00f1o\n  if((m=s.match(/^(\\d{1,2})[\\s\\-\\/]([a-z]{3})[a-z.]*[\\s\\-\\/](\\d{2,4})/))){const i=MES3.indexOf(m[2]);if(i>=0){let y=+m[3];if(y<100)y+=2000;return y+'-'+p2(i+1)+'-'+p2(m[1])}}\n  return null;\n}\nconst FIELDS=[['f','Fecha o A\u00f1o/Mes',true],['t','Tipo (gasto/ingreso/transferencia)'],['c','Cuenta',true],['k','Concepto'],['cat','Categor\u00eda'],['e','Etiqueta'],['d','Descripci\u00f3n'],['m','Importe (con signo)'],['mi','Importe ingreso (columna aparte)'],['mg','Importe gasto (columna aparte)'],['mon','Moneda']];\nconst GUESS={f:['fecha','ano/mes','a\u00f1o/mes','periodo','mes'],t:['tipo'],c:['cuenta'],k:['concepto','subcategoria','rubro'],cat:['categoria'],e:['etiqueta','tag'],d:['descrip','detalle','nota','comentario','observ'],m:['importe','monto','valor','total'],mi:['ingreso','haber','credito'],mg:['gasto','egreso','debe','debito'],mon:['moneda','divisa']};\nasync function onFile(inp){\n  const file=inp.files[0];inp.value='';if(!file)return;\n  if(!S.cfg){toast('Esper\u00e1 a que carguen tus datos');return}\n  if(!window.XLSX){toast('No carg\u00f3 el lector de Excel. Recarg\u00e1 la p\u00e1gina.');return}\n  try{\n    const buf=await file.arrayBuffer();let wb;\n    if(/\\.(csv|txt)$/i.test(file.name)){\n      let txt;try{txt=new TextDecoder('utf-8',{fatal:true}).decode(buf)}catch(e){txt=new TextDecoder('windows-1252').decode(buf)}\n      wb=XLSX.read(txt,{type:'string',raw:true});\n    }else wb=XLSX.read(buf,{type:'array'});\n    // busca la hoja y la fila de encabezados con m\u00e1s coincidencias\n    let best=null;\n    for(const sn of wb.SheetNames){\n      const rows=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,raw:true,defval:''});\n      for(let i=0;i<Math.min(rows.length,40);i++){\n        const sc=rows[i].filter(c=>{const n=norm(c);return n&&Object.values(GUESS).some(g=>g.some(w=>n.startsWith(w)))}).length;\n        if(sc>=2&&(!best||sc>best.sc))best={sc,rows,hi:i,sn};\n      }\n    }\n    if(!best)throw new Error('no encontr\u00e9 una fila de encabezados (Fecha, Cuenta, Importe\u2026)');\n    const H=best.rows[best.hi].map(c=>String(c).trim());\n    const map={};const usado=new Set();\n    for(const [k] of FIELDS){\n      const i=H.findIndex((x,j)=>!usado.has(j)&&GUESS[k].some(w=>norm(x).startsWith(w)));\n      map[k]=i;if(i>=0)usado.add(i);\n    }\n    if(map.m>=0){map.mi=-1;map.mg=-1}\n    IMPF={rows:best.rows,hi:best.hi,H,map,fname:file.name,sheet:best.sn,cotUSD:cot()};\n    renderMapeo();\n  }catch(e){toast('No pude leer el archivo: '+(e.message||e))}\n}\nfunction parseImport(){\n  const {rows,hi,map:M}=IMPF;const R=[];const seen={};let malas=0;\n  const g=(r,k)=>M[k]>=0?r[M[k]]:'';\n  for(let i=hi+1;i<rows.length;i++){\n    const r=rows[i];if(!r||r.every(c=>String(c).trim()===''))continue;\n    const f=parseFecha(g(r,'f'));\n    let v=M.m>=0?parseNum(g(r,'m')):NaN;\n    if(isNaN(v)&&(M.mi>=0||M.mg>=0)){const a=parseNum(g(r,'mi')),b=parseNum(g(r,'mg'));if(a)v=Math.abs(a);else if(b)v=-Math.abs(b)}\n    const cn=String(g(r,'c')??'').replace(/\\s+/g,' ').trim();\n    if(!f||isNaN(v)||!cn||v===0){malas++;continue}\n    const tt=norm(g(r,'t'));\n    let t=/^(gas|egr|deb)/.test(tt)?'G':/^(ing|cre)/.test(tt)?'I':/^tra/.test(tt)?'T':/^aj/.test(tt)?'A':tt==='g'?'G':tt==='i'?'I':(v<0?'G':'I');\n    const kn=String(g(r,'k')??'').replace(/\\s+/g,' ').trim(),cat=String(g(r,'cat')??'').replace(/\\s+/g,' ').trim();\n    const e=String(g(r,'e')??'').trim(),d=String(g(r,'d')??'').trim(),mon=norm(g(r,'mon'));\n    if(t!=='T'&&/^transferencia/i.test(kn)&&M.t<0)t='T';\n    const key=[f,t,cn,kn,e,d,v].join('|');seen[key]=(seen[key]||0)+1;\n    R.push({i,f,t,tt,cn,kn,cat,e,d,v,mon,id:'z'+hash(key+'#'+seen[key])});\n  }\n  return {R,malas};\n}\nfunction renderMapeo(){\n  const I=IMPF,{R,malas}=parseImport();\n  const sel=k=>`<select onchange=\"IMPF.map['${k}']=+this.value;renderMapeo()\"><option value=\"-1\">\u2014 ninguna \u2014</option>${I.H.map((x,j)=>x?`<option value=\"${j}\"${I.map[k]===j?' selected':''}>${h(x)}</option>`:'').join('')}</select>`;\n  const sinDia=R.length&&R.every(x=>x.f.endsWith('-01'))&&I.map.f>=0&&/mes/.test(norm(I.H[I.map.f]));\n  modal(`<h3>Importar: ${h(I.fname)}</h3>\n   <p class=\"sm mu\" style=\"margin-top:0\">Hoja \u201c${h(I.sheet)}\u201d, encabezados en la fila ${I.hi+1}. Revis\u00e1 qu\u00e9 columna corresponde a cada dato.</p>\n   <div class=\"fld2\">${FIELDS.map(([k,l,req])=>`<div class=\"fld\"><label>${h(l)}${req?' *':''}</label>${sel(k)}</div>`).join('')}</div>\n   <div class=\"card\" style=\"padding:10px;margin:6px 0 10px;overflow:auto\"><div class=\"sm\" style=\"margin-bottom:6px\"><b>${R.length}</b> filas v\u00e1lidas${malas?` \u00b7 <span class=\"neg\">${malas} sin fecha, cuenta o importe (se ignoran)</span>`:''}</div>\n   <table class=\"pv\"><tr><th>Fecha</th><th>Tipo</th><th>Cuenta</th><th>Concepto</th><th>Categor\u00eda</th><th>Importe</th></tr>\n   ${R.slice(0,6).map(x=>`<tr><td>${x.f.split('-').reverse().join('/')}</td><td>${{G:'Gasto',I:'Ingreso',T:'Transf.',A:'Ajuste'}[x.t]}</td><td>${h(x.cn)}</td><td>${h(x.kn)}</td><td>${h(x.cat)}</td><td class=\"num\">${NF.format(x.v)}</td></tr>`).join('')}</table></div>\n   ${sinDia?'<p class=\"xs mu\">Este archivo trae solo a\u00f1o y mes: cada movimiento se carga con fecha d\u00eda 1 de su mes.</p>':''}\n   <div class=\"btns\" style=\"justify-content:flex-end\"><button class=\"btn\" onclick=\"closeModal()\">Cancelar</button><button class=\"btn pri\" ${R.length?'':'disabled'} onclick=\"analizar()\">Continuar</button></div>`);\n}\nfunction analizar(){\n  const {R}=parseImport();\n  const esMN=IMPF.map.m>=0&&/m\\/n|moneda nacional/.test(norm(IMPF.H[IMPF.map.m]));\n  const TR=R.filter(x=>x.t==='T');const anyNeg=TR.some(x=>x.v<0);\n  for(const x of TR){const txt=norm(x.tt+' '+x.kn+' '+x.d);\n    x.role=anyNeg?(x.v<0?'o':'d'):(/hacia|enviad|salida/.test(txt)?'o':/desde|recib|entrada/.test(txt)?'d':'?')}\n  const used=new Set(),pairs=[],solos=[];\n  for(let a=0;a<TR.length;a++){const x=TR[a];if(used.has(a)||x.role==='d')continue;\n    let best=-1,bs=1e9;\n    for(let b=Math.max(0,a-8);b<Math.min(TR.length,a+9);b++){if(b===a||used.has(b))continue;const y=TR[b];\n      if(y.f!==x.f||y.cn===x.cn)continue;if(y.role==='o')continue;\n      const sc=Math.abs(b-a)+(Math.abs(Math.abs(y.v)-Math.abs(x.v))<0.005?0:5);if(sc<bs){bs=sc;best=b}}\n    if(best>=0){used.add(a);used.add(best);const y=TR[best];pairs.push(x.role==='?'&&y.role==='?'&&best<a?[y,x]:[x,y])}\n  }\n  TR.forEach((x,a)=>{if(!used.has(a))solos.push(x)});\n  // cuentas y conceptos (se comparan sin importar may\u00fasculas, tildes ni espacios)\n  const accBy={};for(const c of S.cfg.cuentas)accBy[norm(c.n)]=c;\n  const newAcc=[],newCon=[],catUpd=[];\n  const getAcc=(n,mon)=>{const k=norm(n);if(!accBy[k]){const c={n,mon:/usd|u\\$s|dolar/.test(mon||'')||/d[o\u00f3]lar|u\\$s|usd|us\\$/i.test(n)?'USD':'UYU',si:0};accBy[k]=c;newAcc.push(c)}return accBy[k]};\n  const conBy={};for(const c of S.cfg.conceptos)conBy[c.tipo+norm(c.n)]=c;\n  const getCon=(n,t,cat)=>{if(!n)n=t==='G'?'Gastos sin categor\u00eda':'Otros ingresos';const k=t+norm(n);\n    if(!conBy[k]){const c={n,tipo:t,cat:cat||''};conBy[k]=c;newCon.push(c)}else if(cat&&!conBy[k].cat){conBy[k].cat=cat;catUpd.push(conBy[k])}\n    return conBy[k].n};\n  const usd=a=>a.mon==='USD';\n  const conv=(v,a)=>esMN&&usd(a)?r2(v/IMPF.cotUSD):r2(v);\n  const items=[];let nUSD=0;const T0=Date.now();\n  for(const x of R){\n    if(x.t==='G'||x.t==='I'){const a=getAcc(x.cn,x.mon);if(esMN&&usd(a))nUSD++;\n      const o={id:x.id,t:x.t,f:x.f,c:a.n,k:getCon(x.kn,x.t,x.cat),m:conv(Math.abs(x.v),a),ts:T0+x.i};if(x.e)o.e=x.e;if(x.d)o.d=x.d;items.push(o)}\n    else if(x.t==='A'){const a=getAcc(x.cn,x.mon);items.push({id:x.id,t:'A',f:x.f,c:a.n,m:conv(x.v,a),ts:T0+x.i,d:x.d||'Ajuste de saldo'})}\n  }\n  for(const [o,d] of pairs){const a=getAcc(o.cn,o.mon),b=getAcc(d.cn,d.mon);\n    const it={id:'t'+hash(o.id+d.id),t:'T',f:o.f,c:a.n,c2:b.n,m:conv(Math.abs(o.v),a),ts:T0+o.i};\n    if(a.mon!==b.mon||Math.abs(Math.abs(d.v)-Math.abs(o.v))>0.004)it.m2=conv(Math.abs(d.v),b);\n    const desc=o.d||d.d;if(desc&&!/^(desde|hacia)\\b/.test(norm(desc)))it.d=desc;if(o.e)it.e=o.e;items.push(it)}\n  for(const x of solos){const a=getAcc(x.cn,x.mon);const sg=x.role==='o'?-1:x.role==='d'?1:Math.sign(x.v)||1;\n    items.push({id:x.id,t:'A',f:x.f,c:a.n,m:conv(sg*Math.abs(x.v),a),ts:T0+x.i,d:'Transferencia sin pareja'+(x.d?': '+x.d:'')})}\n  const ex=new Set(S.movs.map(x=>x.id));const nuevos=items.filter(x=>!ex.has(x.id)).length;\n  const fs=items.map(x=>x.f).sort();\n  const cnt={G:0,I:0,T:0,A:0};items.forEach(x=>cnt[x.t]++);\n  let tg=0,ti=0;for(const x of items){const a=accBy[norm(x.c)];const v=a&&a.mon==='USD'?x.m*IMPF.cotUSD:x.m;if(x.t==='G')tg+=v;else if(x.t==='I')ti+=v}\n  IMP={items,newAcc,newCon,catUpd};\n  modal(`<h3>Importar de ZetaCuentas</h3>\n   <div class=\"kv\"><span>Movimientos</span><b>${items.length}</b></div>\n   <div class=\"kv sm\"><span class=\"mu\">Per\u00edodo</span><span>${fs[0].split('-').reverse().join('/')} \u2013 ${fs[fs.length-1].split('-').reverse().join('/')}</span></div>\n   <div class=\"kv sm\"><span class=\"mu\">Gastos \u00b7 Ingresos \u00b7 Transferencias \u00b7 Ajustes</span><span>${cnt.G} \u00b7 ${cnt.I} \u00b7 ${cnt.T} \u00b7 ${cnt.A}</span></div>\n   <div class=\"kv sm\"><span class=\"mu\">Total gastos / ingresos</span><span class=\"num\"><span class=\"neg\">${fmt(tg,'UYU')}</span> / <span class=\"pos\">${fmt(ti,'UYU')}</span></span></div>\n   ${solos.length?`<div class=\"kv sm\"><span class=\"aj\">Transferencias sin pareja (entran como ajuste)</span><span>${solos.length}</span></div>`:''}\n   <div class=\"kv sm\"><span class=\"mu\">Nuevos (no estaban cargados)</span><b>${nuevos}</b></div>\n   ${nUSD?`<div class=\"card\" style=\"padding:10px;margin:8px 0\"><div class=\"sm\">Hay <b>${nUSD}</b> movimientos de cuentas en d\u00f3lares, pero el archivo trae los importes ya pasados a pesos. Para guardarlos en d\u00f3lares se dividen por esta cotizaci\u00f3n:</div>\n     <div class=\"fld\" style=\"margin:8px 0 0\"><input type=\"text\" inputmode=\"decimal\" value=\"${String(IMPF.cotUSD).replace('.',',')}\" onchange=\"const v=parseNum(this.value);if(v>0){IMPF.cotUSD=v;analizar()}\"></div></div>`:''}\n   ${newAcc.length?`<p class=\"sm\" style=\"margin-bottom:4px\"><b>Cuentas nuevas (${newAcc.length}):</b></p><div class=\"sm mu\">${newAcc.map(c=>h(c.n)+' <span class=\"chip\">'+sym(c.mon)+'</span>').join(', ')}</div>`:''}\n   ${newCon.length?`<p class=\"sm\" style=\"margin-bottom:4px\"><b>Conceptos nuevos (${newCon.length}):</b></p><div class=\"sm mu\">${newCon.map(c=>h(c.n)+(c.cat?' <span class=\"xs\">('+h(c.cat)+')</span>':'')).join(', ')}</div>`:''}\n   <p class=\"xs mu\">La moneda de cada cuenta sale del nombre (\u201cD\u00f3lares\u201d, \u201cU$S\u201d). Se puede cambiar despu\u00e9s en Cuentas.</p>\n   <div class=\"btns\" style=\"justify-content:flex-end\"><button class=\"btn\" onclick=\"renderMapeo()\">\u2039 Volver</button><button class=\"btn pri\" id=\"btnImp\" onclick=\"aplicarImport()\">Importar</button></div>`);\n}\nasync function aplicarImport(){\n  const {items,newAcc,newCon}=IMP;IMP=null;\n  const b=document.getElementById('btnImp');b.disabled=true;b.textContent='Guardando en el Sheet\u2026';\n  S.cfg.cuentas.push(...newAcc);S.cfg.conceptos.push(...newCon);\n  const idx={};S.movs.forEach((x,i)=>idx[x.id]=i);\n  for(const x of items){if(idx[x.id]!=null)S.movs[idx[x.id]]=x;else{idx[x.id]=S.movs.length;S.movs.push(x)}}\n  invalidate();\n  try{\n    await saveCfg();\n    for(let i=0;i<items.length;i+=1000){b.textContent=`Guardando\u2026 ${Math.min(i+1000,items.length)}/${items.length}`;await op('importMovs',JSON.stringify(items.slice(i,i+1000)))}\n    closeModal();V.per=ym(hoy());go('inicio');toast(`Importados ${items.length} movimientos`);\n  }catch(e){closeModal()}\n}\n\n/* ===================== Arranque ===================== */\nbuildNav();render();\nif(window.google&&google.script)load();else{S.err='Esta p\u00e1gina tiene que abrirse desde la aplicaci\u00f3n web de Google Apps Script.';render()}\n";
