/*** UmeGas → Planilla de pedidos *********************************************
 * Conector (Apps Script) que recibe los pedidos de la app UmeGas y los anota
 * en la planilla, guardando el comprobante en la carpeta de Drive.
 *
 * Pestaña destino (gid):  1815523778
 * Carpeta de comprobantes: 13AnwBMvqJ5GzPWCRCdJldVqerKMzf0J4
 *
 * Columnas:
 *  A Marca temporal · B RESPONSABLE · C WHATSAPP · D BARRIO/ZONA · E LOTE
 *  F (vacía) · G (vacía) · H TIPO · I MODO · J 45K · K 30K · L 15K · M 10K
 *  N COMPROBANTES
 ****************************************************************************/

const SHEET_GID = 1815523778;
const CARPETA_COMPROBANTES = '13AnwBMvqJ5GzPWCRCdJldVqerKMzf0J4';

function doPost(e){
  try{
    const data = JSON.parse(e.postData.contents);
    if(data.accion === 'pedido')      return guardarPedido(data);
    if(data.accion === 'comprobante') return guardarComprobante(data);
    return ok();
  }catch(err){
    return ContentService.createTextOutput('ERROR: ' + err)
      .setMimeType(ContentService.MimeType.TEXT);
  }
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
  const tipo = d.tipoPedido === 'nuevo' ? 'Tubo nuevo' : 'Recarga';
  const modo = d.urgencia === 'urgente' ? '🔴 URGENTE' : '💚 TRANCA';
  const fila = [
    new Date(),          // A Marca temporal
    d.nombre || '',      // B RESPONSABLE (nombre de quien pide)
    '',                  // C WHATSAPP (llega con el comprobante)
    d.barrio || '',      // D BARRIO/ZONA
    d.lote || '',        // E LOTE
    '',                  // F
    '',                  // G
    tipo,                // H TIPO
    modo,                // I MODO
    d.c45 ? d.c45 : '',  // J 45 K
    d.c30 ? d.c30 : '',  // K 30 K
    d.c15 ? d.c15 : '',  // L 15 K
    d.c10 ? d.c10 : '',  // M 10 K
    ''                   // N COMPROBANTES (llega con el comprobante)
  ];
  h.appendRow(fila);
  PropertiesService.getScriptProperties()
    .setProperty('row_' + d.id, String(h.getLastRow()));
  return ok();
}

function guardarComprobante(d){
  const h = hoja();
  const props = PropertiesService.getScriptProperties();
  const filaNum = parseInt(props.getProperty('row_' + d.id), 10);

  let link = '';
  if(d.imagen){
    const carpeta = carpetaComprobantes();
    const bytes = Utilities.base64Decode(d.imagen);
    const blob = Utilities.newBlob(bytes, d.tipo || 'image/jpeg',
      'comprobante_' + (d.id || '') + '.jpg');
    const archivo = carpeta.createFile(blob);
    try{
      archivo.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    }catch(err){}
    link = 'https://drive.google.com/open?id=' + archivo.getId();
  }

  if(!filaNum){
    // Si por algo no se encontró el pedido, igual no perdemos el comprobante.
    h.appendRow([new Date(), '', d.whatsapp || '', '', '', '', '',
                 '', '', '', '', '', '', link]);
    if(d.aclaracion) h.getRange(h.getLastRow(), colAclaracion(h)).setValue(d.aclaracion);
    return ok();
  }
  if(d.whatsapp)   h.getRange(filaNum, 3).setValue(d.whatsapp);   // C WHATSAPP
  if(link)         h.getRange(filaNum, 14).setValue(link);        // N COMPROBANTES
  if(d.aclaracion) h.getRange(filaNum, colAclaracion(h)).setValue(d.aclaracion);
  return ok();
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

// Usa la carpeta oficial si hay acceso; si no, una carpeta propia "Comprobantes UmeGas".
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
  return ContentService.createTextOutput('OK')
    .setMimeType(ContentService.MimeType.TEXT);
}
