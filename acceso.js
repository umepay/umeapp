/* acceso.js · Login compartido Cooperativa/Admin para la app Umepay.
 * Lo usan varias páginas (index, mapa, gas). Expone window.Acceso.
 * Valida contra el conector accesos-conector.gs (Apps Script).
 *
 * API:
 *   Acceso.onCambio(fn)  → fn(sesion) se llama al cargar y cada vez que cambia
 *                          (sesion = {rol:'coop'|'admin'} o null)
 *   Acceso.abrirLogin()  → abre el cartel de ingreso
 *   Acceso.cerrarSesion()→ cierra sesión
 *   Acceso.haySesion()   Acceso.rol()   Acceso.esAdmin()   Acceso.nombreRol(r)
 */
(function(){
  'use strict';
  var URL='https://script.google.com/macros/s/AKfycbygcye6l4qghtUGheeSViPuWdyNY_4JrNg7jC51hKIpG3VXsC54vZ8vI0T-8FTzTOkH/exec';
  var TOKEN_KEY='ume_acceso_token', ROL_KEY='ume_acceso_rol';
  var sesion=null, escuchas=[];

  /* ---- red (mismo mecanismo que UmeGas: POST no-cors + leer por JSONP) ---- */
  function jsonp(u){
    return new Promise(function(resolve,reject){
      var cb='ume_cb_'+Math.random().toString(36).slice(2);
      var s=document.createElement('script'), done=false;
      function limpiar(){ try{delete window[cb];}catch(e){window[cb]=undefined;} if(s.parentNode)s.parentNode.removeChild(s); }
      window[cb]=function(d){ done=true; resolve(d); limpiar(); };
      s.onerror=function(){ if(!done){ limpiar(); reject(new Error('red')); } };
      s.src=u+(u.indexOf('?')<0?'?':'&')+'callback='+cb+'&t='+Date.now();
      document.body.appendChild(s);
      setTimeout(function(){ if(!done){ limpiar(); reject(new Error('timeout')); } },12000);
    });
  }
  function esperar(ms){ return new Promise(function(r){ setTimeout(r,ms); }); }
  async function loginRemoto(usuario,clave){
    var reqid='q'+Math.random().toString(36).slice(2)+Date.now();
    try{
      await fetch(URL,{method:'POST',mode:'no-cors',headers:{'Content-Type':'text/plain;charset=utf-8'},
        body:JSON.stringify({accion:'login',usuario:usuario,clave:clave,reqid:reqid})});
    }catch(e){ return {ok:false,error:'Sin conexión. Revisá internet e intentá de nuevo.'}; }
    for(var i=0;i<8;i++){
      await esperar(i===0?700:750);
      try{
        var res=await jsonp(URL+'?accion=resultado&reqid='+encodeURIComponent(reqid));
        if(res && !res.pendiente) return res;
      }catch(e){}
    }
    return {ok:false,error:'El servidor no respondió. Probá de nuevo.'};
  }
  async function validarRemoto(token){
    try{ return await jsonp(URL+'?accion=validar&token='+encodeURIComponent(token)); }
    catch(e){ return {ok:false,offline:true}; }
  }

  /* ---- estado ---- */
  function rol(){ return sesion?sesion.rol:null; }
  function esAdmin(){ return rol()==='admin'; }
  function haySesion(){ return !!sesion; }
  function nombreRol(r){ return r==='admin'?'Administrador':'Cooperativa'; }
  function avisar(){ escuchas.forEach(function(fn){ try{ fn(sesion); }catch(e){} }); }
  function onCambio(fn){ escuchas.push(fn); try{ fn(sesion); }catch(e){} }
  function setSesion(r,token){ sesion={rol:r}; if(token){ localStorage.setItem(TOKEN_KEY,token); localStorage.setItem(ROL_KEY,r); } avisar(); }
  function cerrarSesion(){ localStorage.removeItem(TOKEN_KEY); localStorage.removeItem(ROL_KEY); sesion=null; avisar(); }
  async function restaurar(){
    var t=localStorage.getItem(TOKEN_KEY); if(!t){ avisar(); return; }
    var res=await validarRemoto(t);
    if(res && res.ok) setSesion(res.rol,null);
    else if(res && res.offline) setSesion(localStorage.getItem(ROL_KEY)||'coop',null);
    else cerrarSesion();
  }

  /* ---- cartel de ingreso (se inyecta solo, una vez) ---- */
  function inyectar(){
    if(document.getElementById('accModal')) return;
    var st=document.createElement('style');
    st.textContent=
      '.acc-bg{position:fixed;inset:0;background:rgba(0,0,0,.5);display:none;align-items:center;justify-content:center;z-index:2000;padding:20px;font-family:system-ui,sans-serif}'
     +'.acc-bg.show{display:flex}'
     +'.acc-modal{background:#fff;border-radius:14px;padding:20px 18px 18px;width:100%;max-width:330px;position:relative;box-shadow:0 10px 40px rgba(0,0,0,.35)}'
     +'.acc-modal h2{font-size:18px;color:#2d5a3d;margin-bottom:4px}'
     +'.acc-sub{font-size:13px;color:#666;margin-bottom:14px;line-height:1.35}'
     +'.acc-modal label{display:block;font-size:12px;color:#555;font-weight:600;margin-bottom:10px}'
     +'.acc-modal input{display:block;width:100%;margin-top:4px;padding:10px;border:1.5px solid #cdd9cd;border-radius:8px;font-size:16px;font-family:inherit}'
     +'.acc-modal input:focus{outline:none;border-color:#2d5a3d}'
     +'.acc-error{color:#c92a2a;font-size:13px;min-height:18px;margin:2px 0 8px}'
     +'.acc-btn{width:100%;background:#2d5a3d;color:#fff;border:none;border-radius:8px;padding:12px;font-size:15px;font-weight:600;cursor:pointer;font-family:inherit}'
     +'.acc-btn:disabled{opacity:.6}'
     +'.acc-x{position:absolute;top:8px;right:12px;background:none;border:none;color:#999;font-size:22px;cursor:pointer;line-height:1}';
    document.head.appendChild(st);
    var d=document.createElement('div');
    d.id='accModal'; d.className='acc-bg';
    d.innerHTML=
      '<div class="acc-modal">'
     +'<button class="acc-x" id="accX" aria-label="Cerrar">&times;</button>'
     +'<h2>Ingresar</h2>'
     +'<p class="acc-sub">Acceso de la Cooperativa / Administración.</p>'
     +'<label>Usuario<input id="accUsuario" type="text" autocomplete="username" autocapitalize="none" autocorrect="off" spellcheck="false"></label>'
     +'<label>Contraseña<input id="accClave" type="password" autocomplete="current-password"></label>'
     +'<div class="acc-error" id="accError"></div>'
     +'<button class="acc-btn" id="accEntrar">Entrar</button>'
     +'</div>';
    document.body.appendChild(d);
    document.getElementById('accX').addEventListener('click',cerrarLogin);
    d.addEventListener('click',function(e){ if(e.target===d) cerrarLogin(); });
    document.getElementById('accEntrar').addEventListener('click',hacerLogin);
    document.getElementById('accClave').addEventListener('keydown',function(e){ if(e.key==='Enter') hacerLogin(); });
  }
  function abrirLogin(){ inyectar(); document.getElementById('accError').textContent=''; document.getElementById('accModal').classList.add('show'); setTimeout(function(){ document.getElementById('accUsuario').focus(); },50); }
  function cerrarLogin(){ var m=document.getElementById('accModal'); if(m){ m.classList.remove('show'); document.getElementById('accClave').value=''; } }
  async function hacerLogin(){
    var u=document.getElementById('accUsuario').value.trim();
    var c=document.getElementById('accClave').value;
    var err=document.getElementById('accError'); err.textContent='';
    if(!u||!c){ err.textContent='Completá usuario y contraseña.'; return; }
    var btn=document.getElementById('accEntrar'); btn.disabled=true; var prev=btn.textContent; btn.textContent='Entrando…';
    var res=await loginRemoto(u,c);
    btn.disabled=false; btn.textContent=prev;
    if(res && res.ok){ setSesion(res.rol,res.token); cerrarLogin(); }
    else err.textContent=(res&&res.error)||'No se pudo ingresar.';
  }

  window.Acceso={ onCambio:onCambio, abrirLogin:abrirLogin, cerrarSesion:cerrarSesion, rol:rol, esAdmin:esAdmin, haySesion:haySesion, nombreRol:nombreRol, URL:URL };

  if(document.body) inyectar(); else document.addEventListener('DOMContentLoaded',inyectar);
  restaurar();
})();
