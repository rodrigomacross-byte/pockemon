/**
 * Ventas feria: recibe las ventas desde el formulario del celular.
 * - Guarda las fotos de cada venta en una carpeta de Drive ("Ventas feria / <ID de la venta>").
 * - Anota cada venta en una hoja de cálculo ("Ventas feria").
 * - Si una venta llega dos veces (mismo ID), la ignora: nunca se duplica.
 *
 * Antes de implementarlo, cambia CLAVE por la clave que te entregó Claude.
 */
const CLAVE = 'CAMBIAR-ESTA-CLAVE';
const NOMBRE_CARPETA = 'Ventas feria';
const NOMBRE_HOJA = 'Ventas feria';
const ENCABEZADOS = [
  'ID', 'Dispositivo', 'Creada', 'Cerrada', 'Recibida', 'Acción',
  'Descuento venta', 'Descuentos por foto', 'Medio de pago', 'Total cobrado',
  'N° boleta', 'Nota', 'N° fotos', 'Carpeta', 'Estado', 'Juegos (según la app)', 'Total estimado'
];
// Tiendas desde donde se lee el catálogo público (products.json). Se prueba en orden.
const TIENDAS = ['https://chileboardgames.com', 'https://chileboardgames.cl'];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.k !== CLAVE) return salida_({ ok: false, error: 'clave' });
  if (p.catalog) return salida_(catalogo_());
  return salida_({ ok: true, ping: true });
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const body = JSON.parse(e.postData.contents);
    if (body.key !== CLAVE) return salida_({ ok: false, error: 'clave' });
    if (body.action !== 'sale') return salida_({ ok: false, error: 'accion' });

    const s = body.sale || {};
    const fotos = body.photos || [];
    if (!/^V-[A-Z0-9-]+$/.test(String(s.id || ''))) return salida_({ ok: false, error: 'id' });
    if (!fotos.length) return salida_({ ok: false, error: 'sin fotos' });

    const hoja = hoja_();
    const ultima = hoja.getLastRow();
    if (ultima > 1) {
      const hallada = hoja.getRange(2, 1, ultima - 1, 1).createTextFinder(s.id).matchEntireCell(true).findNext();
      if (hallada) {
        const n = hoja.getRange(hallada.getRow(), 13).getValue();
        return salida_({ ok: true, id: s.id, photos: Number(n), duplicate: true });
      }
    }

    // Si quedó una carpeta a medias de un intento anterior, se descarta.
    const raiz = carpetaRaiz_();
    const previas = raiz.getFoldersByName(s.id);
    while (previas.hasNext()) previas.next().setTrashed(true);

    const carpeta = raiz.createFolder(s.id);
    fotos.forEach(function (f) {
      const nombre = 'foto-' + ('0' + f.n).slice(-2) + '.jpg';
      carpeta.createFile(Utilities.newBlob(Utilities.base64Decode(f.data), 'image/jpeg', nombre));
    });

    const descFotos = fotos
      .filter(function (f) { return f.discount; })
      .map(function (f) { return 'foto ' + f.n + ': ' + (f.discount.type === 'pct' ? f.discount.value + '%' : '$' + f.discount.value); })
      .join(' | ');
    const descVenta = s.saleDiscount && String(s.saleDiscount.value).trim()
      ? (s.saleDiscount.type === 'pct' ? s.saleDiscount.value + '%' : '$' + s.saleDiscount.value)
      : '';

    const juegos = fotos.map(function (f) {
      return f.game ? ('foto ' + f.n + ': ' + f.game.t.split(/\s[–-]\s/)[0] + ' [' + (f.game.s || 's/sku') + '] $' + f.game.p) : ('foto ' + f.n + ': sin identificar');
    }).join(' | ');

    const ficha = {
      sale: s,
      photos: fotos.map(function (f) { return { n: f.n, file: 'foto-' + ('0' + f.n).slice(-2) + '.jpg', discount: f.discount || null, game: f.game || null }; })
    };
    carpeta.createFile('venta.json', JSON.stringify(ficha, null, 2), MimeType.PLAIN_TEXT);

    // La fila se escribe al final: su existencia indica que la venta llegó completa.
    hoja.appendRow([
      s.id, s.dev || '', s.createdAt || '', s.closedAt || '', new Date().toISOString(),
      s.action === 'stock_boleta' ? 'Rebajar stock y emitir boleta' : 'Solo rebajar stock',
      descVenta, descFotos, s.pay || '', s.total || '', s.boleta || '', s.note || '',
      fotos.length, carpeta.getUrl(), 'Pendiente', juegos, s.est || ''
    ]);

    return salida_({ ok: true, id: s.id, photos: fotos.length });
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** Catálogo público de la tienda: nombre, SKU, precio con IVA y si está disponible. */
function catalogo_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('catalogo');
  if (hit) return JSON.parse(hit);
  let ultimoError = 'sin respuesta';
  for (let t = 0; t < TIENDAS.length; t++) {
    try {
      const items = [];
      for (let pagina = 1; pagina <= 10; pagina++) {
        const r = UrlFetchApp.fetch(TIENDAS[t] + '/products.json?limit=250&page=' + pagina, { muteHttpExceptions: true, followRedirects: true });
        if (r.getResponseCode() !== 200) throw new Error('HTTP ' + r.getResponseCode());
        const prods = JSON.parse(r.getContentText()).products || [];
        if (!prods.length) break;
        prods.forEach(function (pr) {
          (pr.variants || []).forEach(function (v) {
            const extra = v.title && v.title !== 'Default Title' ? ' – ' + v.title : '';
            items.push({ t: pr.title + extra, s: v.sku || '', p: Math.round(Number(v.price)), a: !!v.available });
          });
        });
        if (prods.length < 250) break;
      }
      if (items.length) {
        const out = { ok: true, at: new Date().toISOString(), items: items };
        try { cache.put('catalogo', JSON.stringify(out), 300); } catch (e) { /* si es muy grande, sin caché */ }
        return out;
      }
    } catch (err) { ultimoError = String(err && err.message ? err.message : err); }
  }
  return { ok: false, error: 'catalogo: ' + ultimoError };
}

function salida_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function carpetaRaiz_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('CARPETA_ID');
  if (id) {
    try { return DriveApp.getFolderById(id); } catch (e) { /* si la borraron, se crea de nuevo */ }
  }
  const c = DriveApp.createFolder(NOMBRE_CARPETA);
  props.setProperty('CARPETA_ID', c.getId());
  return c;
}

function hoja_() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('HOJA_ID');
  let ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(NOMBRE_HOJA);
    props.setProperty('HOJA_ID', ss.getId());
    const h = ss.getSheets()[0];
    h.appendRow(ENCABEZADOS);
    h.setFrozenRows(1);
  }
  const hoja = ss.getSheets()[0];
  if (hoja.getRange(1, 16).getValue() === '') hoja.getRange(1, 16, 1, 2).setValues([[ENCABEZADOS[15], ENCABEZADOS[16]]]);
  return hoja;
}

/** Opcional: ejecútala una vez desde el editor para autorizar Drive y Hojas antes de implementar. */
function autorizar() {
  carpetaRaiz_();
  hoja_();
}
