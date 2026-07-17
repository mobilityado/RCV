
(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const API_URL = String(window.REPORTIA_CONFIG?.API_URL || "").trim();
  const SESSION_KEY = "reportia_rcv_session_v19_1";

  function jsonp(params) {
    return new Promise((resolve, reject) => {
      if (!API_URL || API_URL.includes("PEGA_AQUI")) return reject(new Error("Configura la URL de Apps Script en config.js"));
      const callback = "__reportia_cb_" + Date.now() + "_" + Math.random().toString(36).slice(2);
      const script = document.createElement("script");
      const query = new URLSearchParams({...params, callback});
      const timer = setTimeout(() => { cleanup(); reject(new Error("Tiempo de espera agotado al conectar con Apps Script.")); }, 15000);
      function cleanup(){ clearTimeout(timer); try{delete window[callback]}catch(_){ } script.remove(); }
      window[callback] = data => { cleanup(); resolve(data); };
      script.onerror = () => { cleanup(); reject(new Error("No fue posible conectar con Apps Script.")); };
      script.src = API_URL + (API_URL.includes("?") ? "&" : "?") + query.toString();
      document.head.appendChild(script);
    });
  }

  function initials(name){
    const parts=String(name||"AD").trim().split(/\s+/).filter(Boolean);
    return ((parts[0]?.[0]||"A")+(parts[1]?.[0]||"D")).toUpperCase();
  }

  function applySession(session){
    document.body.classList.add("authenticated");
    $("secureLogin").classList.add("hidden");
    $("sessionUser").textContent=session.usuario;
    $("sessionRole").textContent=session.tipo==="ADMINISTRADOR"?"Administrador del sistema":"Usuario del sistema";
    $("sessionAvatar").textContent=initials(session.usuario);
    $("welcomeUser").textContent="BIENVENIDO, "+session.usuario.toUpperCase();
    const settings=document.querySelector('[data-view="settings"]');
    if(settings) settings.style.display=session.tipo==="ADMINISTRADOR"?"":"none";
  }

  function logout(){
    sessionStorage.removeItem(SESSION_KEY);
    location.reload();
  }

  async function loadUsers(){
    const select=$("loginUser"), status=$("apiStatus");
    try{
      const data=await jsonp({accion:"usuarios"});
      if(!data?.ok) throw new Error(data?.mensaje||"No fue posible cargar los usuarios.");
      const users=Array.isArray(data.usuarios)?data.usuarios:[];
      select.innerHTML='<option value="">Selecciona un usuario</option>'+users.map(u=>`<option value="${String(u.usuario).replace(/"/g,"&quot;")}">${u.usuario} — ${u.tipo}</option>`).join("");
      status.textContent=`${users.length} usuario(s) disponibles.`;
      status.className="secure-api-status ok";
    }catch(err){
      select.innerHTML='<option value="">No se pudieron cargar los usuarios</option>';
      status.textContent=err.message;
      status.className="secure-api-status error";
    }
  }

  async function login(){
    const usuario=$("loginUser").value.trim(), contrasena=$("loginPassword").value;
    const msg=$("loginMessage"), btn=$("loginBtn");
    if(!usuario||!contrasena){msg.textContent="Selecciona un usuario e ingresa la contraseña.";msg.className="secure-login-message error";return;}
    btn.disabled=true;btn.textContent="Validando acceso…";msg.textContent="";
    try{
      const data=await jsonp({accion:"login",usuario,contrasena});
      if(!data?.ok) throw new Error(data?.mensaje||"Usuario o contraseña incorrectos.");
      const session={usuario:data.usuario,tipo:String(data.tipo||"USUARIO").toUpperCase(),loginAt:new Date().toISOString()};
      sessionStorage.setItem(SESSION_KEY,JSON.stringify(session));
      applySession(session);
    }catch(err){msg.textContent=err.message;msg.className="secure-login-message error";}
    finally{btn.disabled=false;btn.textContent="Entrar a REPORT.IA";}
  }

  $("togglePassword").addEventListener("click",()=>{$("loginPassword").type=$("loginPassword").type==="password"?"text":"password";});
  $("loginBtn").addEventListener("click",login);
  $("loginPassword").addEventListener("keydown",e=>{if(e.key==="Enter")login();});
  $("logoutBtn").addEventListener("click",logout);

  let session=null;try{session=JSON.parse(sessionStorage.getItem(SESSION_KEY)||"null")}catch(_){}
  if(session?.usuario) applySession(session); else loadUsers();
})();
