/**
 * PUENTE "Shipping Orders 26-27" → Shipping App (Harvest Tek México)
 *
 * Lee la hoja de programación de Ventas (solo lectura, con los permisos de
 * quien publica este script) y entrega a la app las órdenes agrupadas por SO.
 *
 * Seguridad: solo responde si la app envía un token de sesión válido de un
 * usuario @harvesttekmex.com o @sfcos.com (el mismo login de Google de la app).
 *
 * Publicar como: Implementar → Nueva implementación → Aplicación web
 *   - Ejecutar como: Yo
 *   - Quién tiene acceso: Cualquier persona
 */

const SHEET_ID = '1_5J4Gdh07EFzre8SzSkqzauCdkKvE0iPDneN_HXscgY';
const FIREBASE_API_KEY = 'AIzaSyCy_vgYncZv48qTf5sPXgZeRdOSu_CgUE8';
const APP_URL = 'https://shippingapp-eight.vercel.app/';
const DOMINIOS = ['harvesttekmex.com', 'sfcos.com'];
// Pestañas que repiten datos de otras (no se leen):
const EXCLUIR = ['QUERY'];
// Encabezados mínimos para considerar una pestaña como "hoja de programación":
const REQUERIDOS = ['SO', 'Harvest Date', 'Product', 'Client', 'Pallets'];
const CACHE_SEG = 60;

// Prueba rápida desde el navegador: abre la URL /exec y debe decir ok:true
function doGet() {
  return json_({ ok: true, servicio: 'shipping-ventas', hora: new Date().toISOString() });
}

function doPost(e) {
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const email = verificarToken_(body.idToken);
    const desde = body.desde || '';
    const hasta = body.hasta || '9999-12-31';
    const datos = leerOrdenesCache_(desde, hasta);
    return json_({ ok: true, usuario: email, generado: new Date().toISOString(), hojas: datos.hojas, ordenes: datos.ordenes });
  } catch (err) {
    return json_({ ok: false, error: String((err && err.message) || err) });
  }
}

// ── Verificación del login (token de Firebase) ──
function verificarToken_(idToken) {
  if (!idToken) throw new Error('Sin sesión');
  const cache = CacheService.getScriptCache();
  const ck = 'tk_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, idToken)).slice(0, 40);
  const hit = cache.get(ck);
  if (hit) return hit;
  const r = UrlFetchApp.fetch('https://identitytoolkit.googleapis.com/v1/accounts:lookup?key=' + FIREBASE_API_KEY, {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify({ idToken: idToken }),
    headers: { Referer: APP_URL },
    muteHttpExceptions: true
  });
  if (r.getResponseCode() !== 200) throw new Error('Sesión inválida o vencida');
  const u = (JSON.parse(r.getContentText()).users || [])[0];
  const email = (u && u.email || '').toLowerCase();
  if (!u || !u.emailVerified || DOMINIOS.indexOf(email.split('@')[1]) < 0) throw new Error('Cuenta no autorizada');
  cache.put(ck, email, 600);
  return email;
}

// ── Lectura de la hoja ──
function leerOrdenesCache_(desde, hasta) {
  const cache = CacheService.getScriptCache();
  const ck = 'ord_' + desde + '_' + hasta;
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);
  const datos = leerOrdenes_(desde, hasta);
  const txt = JSON.stringify(datos);
  if (txt.length < 95000) cache.put(ck, txt, CACHE_SEG);
  return datos;
}

function leerOrdenes_(desde, hasta) {
  // Lectura con la API de Sheets (servicio avanzado "Sheets"): funciona con permiso de solo lectura.
  const q = function (n) { return "'" + n.replace(/'/g, "''") + "'"; };
  const meta = Sheets.Spreadsheets.get(SHEET_ID, { fields: 'sheets.properties.title' });
  const nombres = meta.sheets.map(function (s) { return s.properties.title; })
    .filter(function (n) { return EXCLUIR.indexOf(n.trim()) < 0; });
  const encs = Sheets.Spreadsheets.Values.batchGet(SHEET_ID, { ranges: nombres.map(function (n) { return q(n) + '!1:1'; }) }).valueRanges;
  const validas = [];
  encs.forEach(function (vr, i) {
    const enc = ((vr.values || [])[0] || []).map(function (h) { return String(h).trim(); });
    if (REQUERIDOS.every(function (h) { return enc.indexOf(h) >= 0; })) validas.push({ nombre: nombres[i], enc: enc });
  });
  if (!validas.length) return { hojas: [], ordenes: [] };
  const datos = Sheets.Spreadsheets.Values.batchGet(SHEET_ID, {
    ranges: validas.map(function (v) { return q(v.nombre) + '!A2:AZ'; }),
    valueRenderOption: 'UNFORMATTED_VALUE',
    dateTimeRenderOption: 'SERIAL_NUMBER'
  }).valueRanges;
  const FECHAS = ['Harvest Date', 'Ship Date', 'Deliver Date'];
  const filas = [];
  datos.forEach(function (vr, i) {
    const nombre = validas[i].nombre, enc = validas[i].enc;
    (vr.values || []).forEach(function (v, k) {
      const reg = { _hoja: nombre, _fila: k + 2 };
      enc.forEach(function (h, j) {
        if (!h) return;
        let x = v[j];
        // Fechas de Sheets vienen como número de serie (días desde 30/12/1899)
        if (FECHAS.indexOf(h) >= 0 && typeof x === 'number') x = new Date(Math.round((x - 25569) * 864e5));
        reg[h] = x;
      });
      filas.push(reg);
    });
  });
  return { hojas: validas.map(function (v) { return v.nombre; }), ordenes: agrupar_(filas, 'UTC', desde, hasta) };
}

// ── Lógica de agrupación (función pura, probada fuera de Apps Script) ──
function agrupar_(filas, tz, desde, hasta) {
  function txt(x) { return x === null || x === undefined ? '' : String(x).trim(); }
  function num(x) { const n = typeof x === 'number' ? x : parseFloat(String(x).replace(/,/g, '')); return isNaN(n) ? 0 : n; }
  function fecha(x) {
    if (x instanceof Date && !isNaN(x)) return Utilities.formatDate(x, tz, 'yyyy-MM-dd');
    return '';
  }
  function soTxt(x) { return typeof x === 'number' ? String(Math.round(x)) : txt(x); }

  const grupos = {};
  const orden = [];
  const contPend = {};
  let bloqueAnterior = null; // para agrupar SO "PENDIENTE" por bloque de filas contiguas

  filas.forEach(function (r) {
    const cosecha = fecha(r['Harvest Date']);
    const producto = txt(r['Product']);
    const cliente = txt(r['Client']);
    const tarimas = num(r['Pallets']);
    const vacia = !cosecha || !producto || !cliente;
    if (vacia) { bloqueAnterior = null; return; }
    if (cosecha < desde || cosecha > hasta) { bloqueAnterior = null; return; }

    let so = soTxt(r['SO']);
    const pendiente = !so || /^pend/i.test(so);
    const destino = txt(r['Destination']);
    let key;
    if (!pendiente) {
      key = 'SO:' + so;
    } else {
      // Mismo bloque contiguo en la misma hoja = misma orden pendiente
      const base = 'PEND:' + r._hoja + '|' + cosecha + '|' + cliente + '|' + destino;
      if (bloqueAnterior && bloqueAnterior.base === base && bloqueAnterior.fila === r._fila - 1) {
        key = bloqueAnterior.key;
      } else {
        contPend[base] = (contPend[base] || 0) + 1;
        key = base + '|' + contPend[base];
      }
      bloqueAnterior = { base: base, key: key, fila: r._fila };
      so = 'PENDIENTE';
    }
    if (!pendiente) bloqueAnterior = null;

    if (!grupos[key]) {
      grupos[key] = {
        key: key, so: so, pendiente: pendiente, cliente: cliente,
        po: txt(r['Client PO']), cosecha: cosecha,
        embarque: fecha(r['Ship Date']), entrega: fecha(r['Deliver Date']),
        destino: destino, prioridad: txt(r['Priority']), tipoTarima: txt(r['Pallet Type']),
        hojas: [], productos: [], totalTarimas: 0
      };
      orden.push(key);
    }
    const g = grupos[key];
    if (g.hojas.indexOf(r._hoja) < 0) g.hojas.push(r._hoja);
    if (tarimas <= 0) return; // líneas en 0 no se cargan
    g.productos.push({
      producto: producto, tarimas: tarimas,
      cantidad: num(r['Qty']), unidades: txt(r['Units']),
      famousId: txt(r['FAMOUS ID'] || r['Ship Out'])
    });
    g.totalTarimas += tarimas;
  });

  return orden.map(function (k) { return grupos[k]; }).filter(function (g) { return g.productos.length > 0; });
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Para probar dentro del editor de Apps Script (botón ▶ Ejecutar con esta función):
function probarLectura() {
  const hoy = new Date();
  const d = Utilities.formatDate(new Date(hoy.getTime() - 7 * 864e5), 'America/Tijuana', 'yyyy-MM-dd');
  const h = Utilities.formatDate(new Date(hoy.getTime() + 21 * 864e5), 'America/Tijuana', 'yyyy-MM-dd');
  const datos = leerOrdenes_(d, h);
  Logger.log('Hojas leídas: ' + datos.hojas.join(', '));
  Logger.log('Órdenes: ' + datos.ordenes.length);
  datos.ordenes.slice(0, 10).forEach(function (o) {
    Logger.log(o.cosecha + ' | ' + o.so + ' | ' + o.cliente + ' | ' + o.destino + ' | ' + o.totalTarimas + ' tarimas | ' + o.productos.map(function (p) { return p.producto + ' (' + p.tarimas + ')'; }).join(', '));
  });
}
