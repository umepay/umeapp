/*** UmeGas → Planilla de pedidos *********************************************
 * Conector (Apps Script) que recibe los pedidos de la app UmeGas y los anota
 * en la planilla, guardando el comprobante en la carpeta de Drive.
 *
 * IMPORTANTE: escribe cada dato buscando la columna POR SU NOMBRE (encabezado),
 * así se pueden mover/reordenar las columnas sin romper nada (ej. ENTREGADO
 * puede estar primera, última, donde sea).
 *
 * Pestaña destino (gid):  1815523778
 * Carpeta de comprobantes: COMPROBANTES (File responses) — carpeta oficial
 * Encabezados que usa: Marca temporal · RESPONSABLE · WHATSAPP · BARRIO/ZONA ·
 *   LOTE · TIPO · MODO · 45 K · 30 K · 15 K · 10 K · COMPROBANTES ·
 *   EXPRESIÓN ESCRITA (aclaración) · ID_APP · ENTREGADO
 ****************************************************************************/

const SHEET_GID = 1815523778;
const CARPETA_COMPROBANTES = '1LgCi59vD6s5i-yAIx_43HPzgppN7JsPXFartuqBZj0KfN1wVICTWElqxh9KYf_jaOj8P9xyM';

function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents);
    if(data.accion === 'pedido')      return guardarPedido(data);
    if(data.accion === 'comprobante') return guardarComprobante(data);
    return ok();
  }catch(err){
    return ContentService.createTextOutput('ERROR: ' + err).setMimeType(ContentService.MimeType.TEXT);
  }
}

// Consulta de estado del pedido (para "Mis pedidos" en la app). Responde JSONP.
function doGet(e){
  const p = (e && e.parameter) || {};
  if(p.accion === 'estado'){
    const est = estadoPedido(p.id);
    const cb = p.callback || 'callback';
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(est) + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput('UmeGas OK').setMimeType(ContentService.MimeType.TEXT);
}

function hoja(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hs = ss.getSheets();
  for(let i = 0; i < hs.length; i++){
    if(hs[i].getSheetId() === SHEET_GID) return hs[i];
  }
  return ss.getSheets()[0];
}

function guardarPedido(d){
  const h = hoja();
  h.appendRow(['']);                 // reserva una fila nueva (evita choques)
  const fila = h.getLastRow();
  setCel(h, fila, 'Marca temporal', new Date());
  setCel(h, fila, 'RESPONSABLE', d.nombre || '');
  setCel(h, fila, 'BARRIO/ZONA', d.barrio || '');
  setCel(h, fila, 'LOTE', d.lote || '');
  setCel(h, fila, 'TIPO', d.tipoPedido === 'nuevo' ? 'Tubo nuevo' : 'Recarga');
  setCel(h, fila, 'MODO', d.urgencia === 'urgente' ? '🔴 URGENTE' : '💚 TRANCA');
  if(d.c45) setCel(h, fila, '45 K', d.c45);
  if(d.c30) setCel(h, fila, '30 K', d.c30);
  if(d.c15) setCel(h, fila, '15 K', d.c15);
  if(d.c10) setCel(h, fila, '10 K', d.c10);
  setCel(h, fila, 'ID_APP', d.id);
  const cEnt = colPorHeader(h, 'ENTREGADO', true);
  h.getRange(fila, cEnt).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
  h.getRange(fila, cEnt).setValue(false);
  PropertiesService.getScriptProperties().setProperty('row_' + d.id, String(fila));
  return ok();
}

function guardarComprobante(d){
  const h = hoja();
  let link = '';
  if(d.imagen){
    const carpeta = carpetaComprobantes();
    const bytes = Utilities.base64Decode(d.imagen);
    const blob = Utilities.newBlob(bytes, d.tipo || 'image/jpeg', 'comprobante_' + (d.id || '') + '.jpg');
    const archivo = carpeta.createFile(blob);
    try{ archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); }catch(err){}
    link = 'https://drive.google.com/open?id=' + archivo.getId();
  }
  let fila = parseInt(PropertiesService.getScriptProperties().getProperty('row_' + d.id), 10);
  if(!fila) fila = filaPorId(h, d.id);
  if(!fila){
    h.appendRow(['']); fila = h.getLastRow();
    if(d.id) setCel(h, fila, 'ID_APP', d.id);
  }
  if(d.whatsapp)   setCel(h, fila, 'WHATSAPP', d.whatsapp);
  if(link)         setCel(h, fila, 'COMPROBANTES', link);
  if(d.aclaracion) h.getRange(fila, colAclaracion(h)).setValue(d.aclaracion);
  return ok();
}

function estadoPedido(id){
  const res = { id: id, entregado: false, encontrado: false };
  if(!id) return res;
  const h = hoja();
  const fila = filaPorId(h, id);
  if(!fila) return res;
  res.encontrado = true;
  const cEnt = colPorHeader(h, 'ENTREGADO', false);
  if(cEnt){
    const v = h.getRange(fila, cEnt).getValue();
    const t = String(v).toLowerCase();
    res.entregado = (v === true || t === 'si' || t === 'true' || t === '✓' || t === 'x');
  }
  return res;
}

// Devuelve el número de fila del pedido con ese ID_APP (0 si no está).
function filaPorId(h, id){
  if(!id) return 0;
  const cId = colPorHeader(h, 'ID_APP', false);
  if(!cId) return 0;
  const ult = h.getLastRow();
  if(ult < 2) return 0;
  const ids = h.getRange(2, cId, ult - 1, 1).getValues();
  for(let i = 0; i < ids.length; i++){ if(String(ids[i][0]) === String(id)) return i + 2; }
  return 0;
}

// Escribe un valor en la fila, buscando la columna por su encabezado (la crea si falta).
function setCel(h, fila, header, val){
  const c = colPorHeader(h, header, true);
  if(c) h.getRange(fila, c).setValue(val);
}

// Busca una columna por su título exacto. Si crear=true y no existe, la agrega al final.
function colPorHeader(h, nombre, crear){
  const ult = Math.max(1, h.getLastColumn());
  const headers = h.getRange(1, 1, 1, ult).getValues()[0];
  for(let i = 0; i < headers.length; i++){
    if(String(headers[i]).trim().toUpperCase() === nombre.toUpperCase()) return i + 1;
  }
  if(crear){ const c = h.getLastColumn() + 1; h.getRange(1, c).setValue(nombre); return c; }
  return 0;
}

// La aclaración de entrega va en la columna "EXPRESIÓN ESCRITA".
function colAclaracion(h){
  const ult = Math.max(1, h.getLastColumn());
  const headers = h.getRange(1, 1, 1, ult).getValues()[0];
  for(let i = 0; i < headers.length; i++){
    if(String(headers[i]).toUpperCase().indexOf('ESCRITA') !== -1) return i + 1;
  }
  const c = h.getLastColumn() + 1;
  h.getRange(1, c).setValue('EXPRESIÓN ESCRITA');
  return c;
}

function carpetaComprobantes(){
  if(CARPETA_COMPROBANTES){
    try{ return DriveApp.getFolderById(CARPETA_COMPROBANTES); }catch(err){}
  }
  const nombre = 'Comprobantes UmeGas';
  const it = DriveApp.getFoldersByName(nombre);
  if(it.hasNext()) return it.next();
  return DriveApp.createFolder(nombre);
}

function ok(){
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}
