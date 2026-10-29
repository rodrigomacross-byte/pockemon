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
  'N° boleta', 'Nota', 'N° fotos', 'Carpeta', 'Estado'
];

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.k !== CLAVE) return salida_({ ok: false, error: 'clave' });
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

    const ficha = {
      sale: s,
      photos: fotos.map(function (f) { return { n: f.n, file: 'foto-' + ('0' + f.n).slice(-2) + '.jpg', discount: f.discount || null }; })
    };
    carpeta.createFile('venta.json', JSON.stringify(ficha, null, 2), MimeType.PLAIN_TEXT);

    // La fila se escribe al final: su existencia indica que la venta llegó completa.
    hoja.appendRow([
      s.id, s.dev || '', s.createdAt || '', s.closedAt || '', new Date().toISOString(),
      s.action === 'stock_boleta' ? 'Rebajar stock y emitir boleta' : 'Solo rebajar stock',
      descVenta, descFotos, s.pay || '', s.total || '', s.boleta || '', s.note || '',
      fotos.length, carpeta.getUrl(), 'Pendiente'
    ]);

    return salida_({ ok: true, id: s.id, photos: fotos.length });
  } catch (err) {
    return salida_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
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
  return ss.getSheets()[0];
}

/** Opcional: ejecútala una vez desde el editor para autorizar Drive y Hojas antes de implementar. */
function autorizar() {
  carpetaRaiz_();
  hoja_();
}
