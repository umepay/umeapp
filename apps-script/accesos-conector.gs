/*** Umepay · Accesos (llaves Cooperativa y Admin) ***************************
 * Conector (Apps Script) que protege ciertas capas del mapa del agua con
 * DOS llaves fijas y compartidas:
 *   - "Cooperativa de agua": desbloquea las capas operativas (puntos, mangueras…)
 *   - "Administrador": por ahora desbloquea lo mismo; queda listo para sumarle
 *     permisos (ej. editar datos) más adelante.
 *
 * POR QUÉ ACÁ Y NO EN LA APP: si la contraseña estuviera en el código de la
 * página, cualquiera podría verla con "ver código fuente". Acá vive en tu
 * cuenta de Google (server-side), donde solo vos la ves.
 *
 * No necesita planilla: las sesiones se firman y se validan solas (sin guardar
 * nada). Cuando alguien inicia sesión recibe un "token" firmado que la app
 * guarda en el teléfono; el candado del mapa lo revalida contra este script.
 *
 * >>> CAMBIÁ estas 5 líneas por tus valores reales antes de desplegar <<<
 ****************************************************************************/
const COOP_USUARIO  = 'cooperativa';
const COOP_CLAVE    = 'CAMBIAME-clave-de-la-cooperativa';
const ADMIN_USUARIO = 'admin';
const ADMIN_CLAVE   = 'CAMBIAME-clave-de-administrador';
const SECRET        = 'CAMBIAME-por-una-frase-larga-y-secreta-para-firmar-2026';
const TOKEN_DIAS    = 30;   // cuántos días queda abierta la sesión en el teléfono

/* DESPLIEGUE (una sola vez, igual que UmeGas):
 *   1) script.google.com → Nuevo proyecto (o Extensiones→Apps Script en una hoja).
 *   2) Pegá este archivo y cambiá las 5 líneas de arriba.
 *   3) Implementar → Nueva implementación → Aplicación web:
 *        Ejecutar como: Yo   ·   Quién accede: Cualquiera
 *      Implementar, autorizar, y copiar la URL que termina en /exec.
 *   4) Pasame esa URL para enchufarla en el mapa.
 */

function doPost(e){
  try{
    const d = JSON.parse(e.postData.contents);
    let res = {ok:false, error:'Acción desconocida'};
    if(d.accion === 'login') res = login(d);
    if(d.reqid) CacheService.getScriptCache().put('r_'+d.reqid, JSON.stringify(res), 120);
    return txt('OK');
  }catch(err){
    return txt('ERROR: '+err);
  }
}

function doGet(e){
  const p  = (e && e.parameter) || {};
  const cb = p.callback || 'callback';
  let out;
  if(p.accion === 'resultado'){          // la app viene a buscar el resultado del login
    const c = CacheService.getScriptCache().get('r_'+p.reqid);
    out = c ? JSON.parse(c) : {ok:false, pendiente:true};
  } else if(p.accion === 'validar'){     // ¿la sesión sigue válida? (candado del mapa)
    out = validarToken(p.token);
  } else {
    return txt('Umepay Accesos OK');
  }
  return jsonp(cb, out);
}

/* ---------- login ---------- */
function login(d){
  const u = String(d.usuario||'').trim().toLowerCase();
  const c = String(d.clave||'');
  let rol = '';
  if(u === COOP_USUARIO.toLowerCase()  && c === COOP_CLAVE)  rol = 'coop';
  else if(u === ADMIN_USUARIO.toLowerCase() && c === ADMIN_CLAVE) rol = 'admin';
  if(!rol) return {ok:false, error:'Usuario o contraseña incorrectos.'};
  return {ok:true, rol:rol, token:crearToken(rol)};
}

/* ---------- sesiones firmadas (sin planilla) ---------- */
function crearToken(rol){
  const payload = rol + '|' + (new Date().getTime() + TOKEN_DIAS*86400000);
  return Utilities.base64EncodeWebSafe(payload) + '.' + firma(payload);
}
function validarToken(tk){
  if(!tk || tk.indexOf('.') < 0) return {ok:false};
  const parts = tk.split('.');
  let payload;
  try{ payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[0])).getDataAsString(); }
  catch(e){ return {ok:false}; }
  if(firma(payload) !== parts[1]) return {ok:false};
  const trozos = payload.split('|'), rol = trozos[0], exp = Number(trozos[1]);
  if(!exp || exp < new Date().getTime()) return {ok:false, error:'Sesión vencida'};
  return {ok:true, rol:rol};
}
function firma(payload){
  const raw = Utilities.computeHmacSha256Signature(payload, SECRET);
  return raw.map(function(b){ return ('0'+(b & 0xff).toString(16)).slice(-2); }).join('');
}

/* ---------- utilidades ---------- */
function txt(s){ return ContentService.createTextOutput(s).setMimeType(ContentService.MimeType.TEXT); }
function jsonp(cb,obj){
  return ContentService.createTextOutput(cb+'('+JSON.stringify(obj)+')')
    .setMimeType(ContentService.MimeType.JAVASCRIPT);
}
