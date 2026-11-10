/**
 * Control de Gastos — CARGADOR
 *
 * Este es el único archivo del proyecto y casi nunca cambia.
 * La app (pantalla y lógica) se baja de GitHub y se guarda en la pestaña oculta "_app" de esta planilla.
 * Las versiones nuevas se instalan solas todas las noches, o con el botón "Actualizar ahora" de la app,
 * sin tener que volver a pegar código ni publicar en Apps Script.
 *
 * Primera vez: elegí la función "instalar" arriba y tocá ▶ Ejecutar.
 */
const CG_URL = 'https://raw.githubusercontent.com/sebastiandemarcoz-glitch/controlgastos/main/app/';
const CG_HOJA = '_app';
let CG_COD = null, CG_API = null;

/* ---------- entradas de la aplicación web ---------- */
function doGet(e) {
  const api = cgApi_();
  if (!api.puedeVer(e)) return HtmlService.createHtmlOutput('<p style="font-family:sans-serif">Acceso no autorizado.</p>');
  return HtmlService.createHtmlOutput(cgCodigo_().html)
    .setTitle('Control de Gastos')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}
function doPost(e) { return cgApi_().doPost(e); }
/** Todo lo que pide la pantalla pasa por acá. */
function rpc(nombre, args) {
  const f = cgApi_()[nombre];
  if (typeof f !== 'function') throw new Error('Función desconocida: ' + nombre);
  return f.apply(null, JSON.parse(args || '[]'));
}
/* ---------- activadores ---------- */
function procesarNotificaciones() { cgApi_().procesarNotificaciones(); }
function autoActualizar() {
  try {
    if (!cgApi_().autoActualizarPermitido()) return;
    const r = cgActualizar_(false);
    if (r.ok && !r.nada) cgApi_().anotarActualizacion(r.version, '');
  } catch (e) {
    try { cgApi_().anotarActualizacion('', String(e.message || e)); } catch (x) { }
  }
}

/* ---------- instalación y actualización ---------- */
/** Ejecutala desde el editor la primera vez (o si algo se rompe): baja la última versión de GitHub. */
function instalar() {
  const o = cgBajar_();
  cgGuardar_(o);
  console.log('Instalada la versión ' + o.version + '. Ya podés abrir la app.');
  return o.version;
}
function cgActualizar_(forzar) {
  const actual = cgCodigo_().version;
  const rem = cgJson_('version.json');
  if (!rem || !rem.version) throw new Error('No pude leer la versión publicada en GitHub');
  if (cgCmp_(rem.version, actual) <= 0) return { ok: true, nada: true, version: actual };
  const o = cgBajar_();
  cgGuardar_(o);
  return { ok: true, version: o.version };
}
function cgBajar_() {
  const rem = cgJson_('version.json');
  if (!rem || !rem.version) throw new Error('No pude leer la versión publicada en GitHub');
  const server = cgTexto_('Code.gs'), html = cgTexto_('Index.html');
  if (server.indexOf("VERSION_APP = '" + rem.version + "'") < 0) throw new Error('Los archivos de GitHub no coinciden con la versión ' + rem.version);
  new Function(server + '\n;return API_;')();   // verifica que el código cargue antes de guardarlo
  return { version: rem.version, server: server, html: html };
}
function cgTexto_(f) {
  const r = UrlFetchApp.fetch(CG_URL + f + '?t=' + Date.now(), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) throw new Error('No pude bajar ' + f + ' de GitHub (' + r.getResponseCode() + ')');
  return r.getContentText('UTF-8');
}
function cgJson_(f) { try { return JSON.parse(cgTexto_(f)); } catch (e) { return null; } }
function cgCmp_(a, b) {
  const x = String(a).split('.').map(Number), y = String(b || '0').split('.').map(Number);
  for (let i = 0; i < 3; i++) if ((x[i] || 0) !== (y[i] || 0)) return (x[i] || 0) - (y[i] || 0);
  return 0;
}

/* ---------- guardado en la planilla (pestaña oculta "_app") ---------- */
function cgGuardar_(o) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sh = ss.getSheetByName(CG_HOJA) || ss.insertSheet(CG_HOJA);
    const filas = [['version', 0, 'x' + o.version]];
    ['server', 'html'].forEach(function (k) {
      const t = o[k];
      for (let i = 0, n = 0; i < t.length; i += 40000, n++) filas.push([k, n, 'x' + t.slice(i, i + 40000)]);
    });
    sh.clear();
    sh.getRange(1, 1, filas.length, 3).setNumberFormat('@').setValues(filas);
    if (!sh.isSheetHidden()) sh.hideSheet();
    SpreadsheetApp.flush();
    CacheService.getScriptCache().remove('cg_ver');
    CG_COD = o; CG_API = null;
  } finally { lock.releaseLock(); }
}
function cgLeer_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CG_HOJA);
  if (!sh || sh.getLastRow() < 2) return null;
  const v = sh.getRange(1, 1, sh.getLastRow(), 3).getValues();
  const o = { version: '', server: '', html: '' };
  v.forEach(function (r) { const t = String(r[2]).slice(1); if (r[0] === 'version') o.version = t; else if (o[r[0]] !== undefined) o[r[0]] += t; });
  return o.server ? o : null;
}
/** Copia en memoria (caché) para que la app abra rápido. */
function cgCodigo_() {
  if (CG_COD) return CG_COD;
  const c = CacheService.getScriptCache();
  const ver = c.get('cg_ver');
  if (ver) {
    const meta = JSON.parse(ver);
    const partes = c.getAll(meta.claves);
    if (meta.claves.every(function (k) { return partes[k] != null; })) {
      const o = { version: meta.version, server: '', html: '' };
      meta.claves.forEach(function (k) { o[k.split('_')[1]] += partes[k]; });
      return (CG_COD = o);
    }
  }
  let o = cgLeer_();
  if (!o) { o = cgBajar_(); cgGuardar_(o); }
  const val = {}, claves = [];
  ['server', 'html'].forEach(function (k) {
    for (let i = 0, n = 0; i < o[k].length; i += 90000, n++) { const key = 'cg_' + k + '_' + n; val[key] = o[k].slice(i, i + 90000); claves.push(key); }
  });
  try { c.putAll(val, 21600); c.put('cg_ver', JSON.stringify({ version: o.version, claves: claves }), 21600); } catch (e) { }
  return (CG_COD = o);
}
function cgApi_() {
  if (!CG_API) CG_API = new Function(cgCodigo_().server + '\n;return API_;')();
  return CG_API;
}
