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
  if(p.accion === 'lotes'){
    const cb = p.callback || 'callback';
    return ContentService.createTextOutput(cb + '(' + JSON.stringify(leerLotes()) + ')')
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
  const lock = LockService.getScriptLock();
  try{ lock.waitLock(20000); }catch(e){}
  try{
    const h = hoja();
    const fila = h.getLastRow() + 1;   // primera fila libre debajo de los datos
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
    SpreadsheetApp.flush();
  } finally {
    try{ lock.releaseLock(); }catch(e){}
  }
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
  const lock = LockService.getScriptLock();
  try{ lock.waitLock(20000); }catch(e){}
  try{
    let fila = parseInt(PropertiesService.getScriptProperties().getProperty('row_' + d.id), 10);
    if(!fila) fila = filaPorId(h, d.id);
    if(!fila){
      // fila nueva: se crea recién acá (el pedido ya no se escribe en el paso 1)
      fila = h.getLastRow() + 1;
      setCel(h, fila, 'Marca temporal', new Date());
      if(d.id) setCel(h, fila, 'ID_APP', d.id);
      const cEnt = colPorHeader(h, 'ENTREGADO', true);
      h.getRange(fila, cEnt).setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
      h.getRange(fila, cEnt).setValue(false);
    }
    // Datos del pedido (llegan junto con el comprobante)
    if(d.nombre != null)     setCel(h, fila, 'RESPONSABLE', d.nombre);
    if(d.barrio != null)     setCel(h, fila, 'BARRIO/ZONA', d.barrio);
    if(d.lote != null)       setCel(h, fila, 'LOTE', d.lote);
    if(d.tipoPedido != null) setCel(h, fila, 'TIPO', d.tipoPedido === 'nuevo' ? 'Tubo nuevo' : 'Recarga');
    if(d.urgencia != null)   setCel(h, fila, 'MODO', d.urgencia === 'urgente' ? '🔴 URGENTE' : '💚 TRANCA');
    if(d.c45) setCel(h, fila, '45 K', d.c45);
    if(d.c30) setCel(h, fila, '30 K', d.c30);
    if(d.c15) setCel(h, fila, '15 K', d.c15);
    if(d.c10) setCel(h, fila, '10 K', d.c10);
    // Comprobante / whatsapp / aclaración
    if(d.whatsapp)   setCel(h, fila, 'WHATSAPP', d.whatsapp);
    if(link)         setCel(h, fila, 'COMPROBANTES', link);
    if(d.aclaracion) h.getRange(fila, colAclaracion(h)).setValue(d.aclaracion);
    if(d.id) PropertiesService.getScriptProperties().setProperty('row_' + d.id, String(fila));
    SpreadsheetApp.flush();
  } finally {
    try{ lock.releaseLock(); }catch(e){}
  }
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

// ===== Lotes editables desde la pestaña "LOTES" (columnas: Lote | Barrio) =====
function hojaLotes(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let h = ss.getSheetByName('LOTES');
  if(!h){ h = ss.insertSheet('LOTES'); h.getRange(1, 1, 1, 2).setValues([['Lote', 'Barrio']]); }
  return h;
}
function leerLotes(){
  const h = hojaLotes();
  if(h.getLastRow() < 2) seedLotes();   // primera vez: carga sola la lista actual
  const ult = h.getLastRow();
  if(ult < 2) return [];
  const vals = h.getRange(2, 1, ult - 1, 2).getValues();
  const out = [];
  for(let i = 0; i < vals.length; i++){
    const l = String(vals[i][0]).trim();
    const b = String(vals[i][1] == null ? '' : vals[i][1]).trim();
    if(l) out.push({ l: l, b: b });
  }
  return out;
}
// EJECUTAR UNA SOLA VEZ para cargar la lista actual en la pestaña LOTES.
// No pisa nada si la pestaña ya tiene datos.
function seedLotes(){
  const h = hojaLotes();
  if(h.getLastRow() > 1) return;
  const LOTES = [{"l": "1. Mauro y Lucy", "b": "Aldea"}, {"l": "2. Eugenia y Sergio", "b": "Aldea"}, {"l": "3. Barro Tal Vez (Batá)", "b": "Aldea"}, {"l": "4. Acros (Coni+Rodri+Bambú)", "b": "Aldea"}, {"l": "5. Casa Kurry", "b": "Aldea"}, {"l": "6. Cris Porto", "b": "Aldea"}, {"l": "7. Danila Tōshin y Noe", "b": "Aldea"}, {"l": "8. Amrit", "b": "Aldea"}, {"l": "9. Ceci Estenssoro", "b": "Aldea"}, {"l": "10. Ana Estenssoro", "b": "Aldea"}, {"l": "11. Santi Calvo", "b": "Aldea"}, {"l": "12. Jorge Estenssoro", "b": "Aldea"}, {"l": "13. Julieta y Robert", "b": "Aldea"}, {"l": "14. Maru y Eze (Domo)", "b": "Aldea"}, {"l": "15. Luciana (Ex Nacho)", "b": "Aldea"}, {"l": "16. Roberto Gallelli", "b": "Aldea"}, {"l": "17. Ademir", "b": "Aldea"}, {"l": "18. María Elena", "b": "Aldea"}, {"l": "19. Stephy y Mauro", "b": "Aldea"}, {"l": "20. Nati Tevelez", "b": "Aldea"}, {"l": "21. Mica Jere y Quimey", "b": "Aldea"}, {"l": "22. Casa Colibrí", "b": "Aldea"}, {"l": "23. Lau Petrolo (TheoLab)", "b": "Aldea"}, {"l": "24. Mati Echeguren", "b": "Aldea"}, {"l": "25. Caro y Ger (DeJa)", "b": "Aldea"}, {"l": "26. Ana Altavista", "b": "Aldea"}, {"l": "27. Bavali", "b": "Nogales Alto"}, {"l": "28. Obrador 100T", "b": "Nogales Alto"}, {"l": "29. Fran Fabre", "b": "Nogales Alto"}, {"l": "30. Caro Martin y Aaron", "b": "Nogales Alto"}, {"l": "31. Tonga", "b": "Nogales Alto"}, {"l": "32. Daphne (Emilse)", "b": "Nogales Alto"}, {"l": "33. Ceci Gus Haruki Akemi", "b": "Nogales Alto"}, {"l": "34. Vivi", "b": "Nogales Alto"}, {"l": "35. Nico Villa", "b": "Nogales Alto"}, {"l": "36. Pau (Titi Pañalera)", "b": "Nogales Alto"}, {"l": "37. Tristán", "b": "Nogales Alto"}, {"l": "38. Mario Paixao", "b": "Nogales Alto"}, {"l": "39. Maga de Masi", "b": "Nogales Alto"}, {"l": "40. Carla Lucas Raiza", "b": "Nogales Alto"}, {"l": "41. Gina y Noe", "b": "Nogales Alto"}, {"l": "42. Vale Bai", "b": "Nogales Alto"}, {"l": "43. Casa 4F", "b": "Nogales Bajo"}, {"l": "44. Fran Herni y Feli (Renoleta)", "b": "Nogales Bajo"}, {"l": "45. Marcelo y Mariana (ikonicoff)", "b": "Nogales Bajo"}, {"l": "46. Eli Lucho Luz y Tomoteo", "b": "Nogales Bajo"}, {"l": "47. Ro Ari Andina Marino", "b": "Nogales Bajo"}, {"l": "48. Graciela Varela (Gaby)", "b": "Nogales Bajo"}, {"l": "49. Rodri Rasta", "b": "Nogales Bajo"}, {"l": "50. Casa Aurora", "b": "Nogales Bajo"}, {"l": "51. Marina Britos (Ford Roja)", "b": "Nogales Bajo"}, {"l": "52. Agos Espeche (Debora)", "b": "Acacias"}, {"l": "53. Daniela Dewey", "b": "Acacias"}, {"l": "54. Juan Dixon", "b": "Acacias"}, {"l": "55. Eze Molina", "b": "Acacias"}, {"l": "56. Magda Chattah", "b": "Acacias"}, {"l": "57. Julieta Nacho Silvestre", "b": "Acacias"}, {"l": "58. Guille Morano", "b": "Espinillos"}, {"l": "59. Martin Corral", "b": "Espinillos"}, {"l": "60. Cami Joaco Selva", "b": "Espinillos"}, {"l": "61. Clari Canale", "b": "Espinillos"}, {"l": "62. Lau Dillon (Aixa)", "b": "Espinillos"}, {"l": "63. Ger Vidal Haan", "b": "Espinillos"}, {"l": "64. Nina+Tomy (Coni Soria)", "b": "Espinillos"}, {"l": "65. Anabella Osky y Aloe", "b": "Espinillos"}, {"l": "66. Sen de Campo (Despensa)", "b": "Cosecha"}, {"l": "67. Cintya y Orión", "b": "Cosecha"}, {"l": "68. Luna (Fondo izquierda)", "b": "Cosecha"}, {"l": "69. Seba Panero (Fondo derecha)", "b": "Cosecha"}, {"l": "70. Maru Mati y León", "b": "Arroyo del Sauce"}, {"l": "71. Carlitos", "b": "Arroyo del Sauce"}, {"l": "72. Belén", "b": "Arroyo del Sauce"}, {"l": "73. Adri", "b": "Arroyo del Sauce"}, {"l": "74. Ulises", "b": "Arroyo del Sauce"}, {"l": "75. Ornella (La Vieja Osada)", "b": "Arroyo del Sauce"}, {"l": "76. Erci", "b": "Arroyo del Sauce"}, {"l": "77. Natacha", "b": "Arroyo del Sauce"}, {"l": "78. Siembra Dicha", "b": "Nogales Alto"}, {"l": "79. Oficina 100T", "b": "Nogales Alto"}, {"l": "80. Ceci Sonzini", "b": "Castaños"}, {"l": "81. Mara", "b": "Castaños"}, {"l": "82. Leo Miraglia", "b": "Castaños"}, {"l": "83. Carolina (3R)", "b": "Tres Rios"}, {"l": "84. Lindsay y Mati (Dutto)", "b": "Tres Rios"}, {"l": "85. Milly Hirschon (3R)", "b": "Tres Rios"}, {"l": "86. Cabaña Pileta (La Vic)", "b": "La Victoria"}, {"l": "87. Cabaña Medio (La Vic)", "b": "La Victoria"}, {"l": "88. Cabaña Fondo (La Vic)", "b": "La Victoria"}, {"l": "92. Coles", "b": "Nogales Bajo"}, {"l": "93. Magui Riachi", "b": "Nogales Alto"}, {"l": "94. Rochi Juampi Awara", "b": "Nogales Alto"}, {"l": "95. Laura Kalauz (3R)", "b": "Tres Rios"}, {"l": "96. Clara Terán (3R)", "b": "Tres Rios"}, {"l": "97. Fran Barlett", "b": "Castaños"}, {"l": "98. Ikonikoff Acacias", "b": "Acacias"}, {"l": "99. La Audelina (Casa Principal fondo)", "b": "La Audelina"}, {"l": "100. Nati Damo Nina Simón", "b": "Aldea"}, {"l": "101. Tiny House (Ikonikoff Acacias)", "b": "Acacias"}, {"l": "102. Escuela Umepay", "b": "Nogales Alto"}, {"l": "103. Lu y Fede (Entrando a Kurry)", "b": "Aldea"}, {"l": "104. Johy (entre Bavali y Tristan)", "b": "Nogales Alto"}, {"l": "105. Magda Zucchi", "b": "Aldea"}, {"l": "106. Chacra Alimento", "b": "Nogales Alto"}, {"l": "107. Vicente Nadal", "b": "Aldea"}, {"l": "108. Gise (Nueva en 3 Ríos)", "b": "Tres Rios"}, {"l": "109. Rosa María (Lucio)", "b": ""}, {"l": "110. Oficina Pueblo Vivo", "b": "Cosecha"}, {"l": "111. Martín Corral (Flor/Pela)", "b": "Espinillos"}, {"l": "112. Mercedes Taranto", "b": "Aldea"}, {"l": "113. Robert (Casa Amor)", "b": "Aldea"}, {"l": "114. Mariano Quiroga", "b": "Tres Rios"}, {"l": "115. Mario Chacón (Lote Hugo)", "b": ""}, {"l": "116. JuliZyk", "b": "Aldea"}, {"l": "117. Coni Guevara", "b": "Paraje de los Guachos"}, {"l": "118. Nico Moner", "b": "Paraje de los Guachos"}];
  const filas = LOTES.map(function(x){ return [x.l, x.b]; });
  h.getRange(2, 1, filas.length, 2).setValues(filas);
}

function ok(){
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}
