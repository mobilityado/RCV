(() => {
  "use strict";
  const API_URL=String(window.REPORTIA_CONFIG?.API_URL||'').trim();let messaging=null,swReg=null,currentToken='';
  const $=id=>document.getElementById(id);
  function jsonp(params,timeout=15000){return new Promise((resolve,reject)=>{const cb='__r51push_'+Date.now()+'_'+Math.random().toString(36).slice(2),sc=document.createElement('script'),q=new URLSearchParams({...params,callback:cb}),tm=setTimeout(()=>{clean();reject(new Error('Tiempo de espera agotado.'))},timeout);function clean(){clearTimeout(tm);try{delete window[cb]}catch(_){}sc.remove()}window[cb]=d=>{clean();resolve(d)};sc.onerror=()=>{clean();reject(new Error('Sin conexión con notificaciones.'))};sc.src=API_URL+(API_URL.includes('?')?'&':'?')+q;document.head.appendChild(sc)})}
  function post(params){return fetch(API_URL,{method:'POST',mode:'no-cors',headers:{'Content-Type':'application/x-www-form-urlencoded;charset=UTF-8'},body:new URLSearchParams(params)})}
  function configured(){const c=window.REPORTIA_FIREBASE_CONFIG||{};return c.apiKey&&c.projectId&&c.messagingSenderId&&c.appId&&window.REPORTIA_VAPID_KEY&&!String(c.apiKey).includes('PEGA_')}
  async function registerSW(){if(!('serviceWorker'in navigator))throw new Error('Este navegador no admite Service Worker.');if(!swReg)swReg=await navigator.serviceWorker.register('./firebase-messaging-sw.js?v=53.3',{scope:'./',updateViaCache:'none'});return swReg}
  async function initMessaging(){if(!configured())throw new Error('Firebase todavía no está configurado en firebase-config.js.');await registerSW();if(!window.firebase)throw new Error('No se cargó Firebase Messaging.');if(!firebase.apps.length)firebase.initializeApp(window.REPORTIA_FIREBASE_CONFIG);messaging=firebase.messaging();messaging.onMessage(payload=>{showForeground(payload);window.dispatchEvent(new CustomEvent('reportia:push',{detail:payload}))});return messaging}
  async function showForeground(payload){const data=payload?.data||{},title=data.title||'REPORT.IA',body=data.body||'Tienes una nueva notificación.';try{const r=await registerSW();if(Notification.permission==='granted')await r.showNotification(title,{body,icon:'./icon-192.png',badge:'./icon-192.png',tag:data.messageId||('reportia-fg-'+Date.now()),renotify:true,data:{url:data.url||'./?reportiaNotification=1',module:data.module||'',messageId:data.messageId||''}})}catch(_){}const toast=document.createElement('div');toast.className='v51-push-toast';toast.innerHTML=`<b>🔔 ${title}</b><span>${body}</span>`;document.body.appendChild(toast);setTimeout(()=>toast.classList.add('show'),20);setTimeout(()=>{toast.classList.remove('show');setTimeout(()=>toast.remove(),350)},6500);try{window.REPORTIA_REFRESH_NOTIFICATIONS?.()}catch(_){} }
  async function enable(){const s=window.REPORTIA_SESSION;if(!s?.token)throw new Error('Inicia sesión primero.');if(!('Notification'in window))throw new Error('Este navegador no admite notificaciones.');const permission=await Notification.requestPermission();if(permission!=='granted')throw new Error('Permiso de notificaciones no concedido.');const m=await initMessaging();currentToken=await m.getToken({vapidKey:window.REPORTIA_VAPID_KEY,serviceWorkerRegistration:swReg});if(!currentToken)throw new Error('No fue posible obtener el token de notificaciones.');await post({accion:'v51_register_push',token:s.token,pushToken:currentToken,deviceLabel:navigator.userAgent.slice(0,120),platform:navigator.platform||''});localStorage.setItem('reportia_push_token_v51',currentToken);await new Promise(r=>setTimeout(r,600));await refreshStatus();return currentToken}
  async function unregister(){const s=window.REPORTIA_SESSION;const pt=currentToken||localStorage.getItem('reportia_push_token_v51')||'';if(s?.token)try{await post({accion:'v51_unregister_push',token:s.token,pushToken:pt})}catch(_){}try{if(messaging&&pt)await messaging.deleteToken()}catch(_){}localStorage.removeItem('reportia_push_token_v51');currentToken='';return true}
  async function syncCurrentToken(pt,force=false){
    const s=window.REPORTIA_SESSION;if(!s?.token||!pt)return false;
    const key='reportia_push_sync_v533',now=Date.now(),last=Number(localStorage.getItem(key)||0);
    if(!force&&now-last<120000)return true;
    await post({accion:'v51_register_push',token:s.token,pushToken:pt,deviceLabel:navigator.userAgent.slice(0,120),platform:navigator.platform||''});
    localStorage.setItem(key,String(now));return true;
  }
  async function deviceStatus(){
    const s=window.REPORTIA_SESSION;
    const permission=('Notification'in window)?Notification.permission:'unsupported';
    let pt=currentToken||localStorage.getItem('reportia_push_token_v51')||'',serviceWorker=false,registered=false,activeUserDevices=0,device=null;
    try{const r=await registerSW();serviceWorker=!!r}catch(_){}
    if(s?.token&&configured()&&permission==='granted'){
      try{const m=await initMessaging();pt=await m.getToken({vapidKey:window.REPORTIA_VAPID_KEY,serviceWorkerRegistration:swReg})||pt;if(pt){currentToken=pt;localStorage.setItem('reportia_push_token_v51',pt);await syncCurrentToken(pt)}}catch(_){}
    }
    if(s?.token){
      try{const d=await jsonp({accion:'v51_push_status',token:s.token});activeUserDevices=Number(d.activos||0)}catch(_){}
      if(pt)try{let d=await jsonp({accion:'v51_push_device_status',token:s.token,pushToken:pt});registered=!!d?.activo;device=d?.device||null;if(!registered&&permission==='granted'){await syncCurrentToken(pt,true);await new Promise(r=>setTimeout(r,700));d=await jsonp({accion:'v51_push_device_status',token:s.token,pushToken:pt});registered=!!d?.activo;device=d?.device||null}}catch(_){}
    }
    return{permission,permissionLabel:permission==='granted'?'Permitido':permission==='denied'?'Bloqueado':permission==='default'?'Pendiente':'No compatible',serviceWorker,registered,activeUserDevices,platform:navigator.platform||navigator.userAgent||'',device};
  }
  async function testLocal(){
    if(!('Notification'in window))throw new Error('Este navegador no admite notificaciones.');
    let p=Notification.permission;if(p!=='granted')p=await Notification.requestPermission();if(p!=='granted')throw new Error('Permiso de notificaciones no concedido.');
    const r=await registerSW();await r.showNotification('REPORT.IA RGI · Prueba local',{body:'Si ves este aviso, el sistema operativo y el navegador permiten notificaciones en este dispositivo.',icon:'./icon-192.png',badge:'./icon-192.png',tag:'reportia-local-test-'+Date.now(),data:{url:'./?reportiaNotification=1'}});return true;
  }
  async function refreshStatus(){const s=window.REPORTIA_SESSION,el=$('v51PushStatus'),btn=$('v51EnablePush');if(!s?.token||!el)return;try{const d=await deviceStatus();if(!configured()){el.textContent='Pendiente de configurar Firebase';el.className='v51-push-status pending';if(btn)btn.textContent='Configurar después';return}if(d.permission==='granted'&&d.registered){el.textContent='Este dispositivo ACTIVO · '+String(s.gerencia||'').trim();el.className='v51-push-status ok';if(btn)btn.textContent='Reactivar en este dispositivo'}else if(d.permission==='denied'){el.textContent='Bloqueadas en este dispositivo';el.className='v51-push-status error';if(btn)btn.textContent='Revisar permisos'}else if(d.permission==='granted'){el.textContent='Permiso concedido · dispositivo no registrado';el.className='v51-push-status pending';if(btn)btn.textContent='Registrar este dispositivo'}else{el.textContent='No activadas en este dispositivo';el.className='v51-push-status pending';if(btn)btn.textContent='Activar notificaciones'}}catch(e){el.textContent=e.message;el.className='v51-push-status error'}}
  async function handleEnable(){const btn=$('v51EnablePush'),msg=$('v51PushMsg');if(btn)btn.disabled=true;if(msg){msg.textContent='Solicitando permiso…';msg.className='v51-msg'}try{await enable();if(msg){msg.textContent='Listo. Este dispositivo recibirá avisos de tu gerencia.';msg.className='v51-msg ok'}}catch(e){if(msg){msg.textContent=e.message;msg.className='v51-msg error'}}finally{if(btn)btn.disabled=false;refreshStatus()}}

  async function testPush(){
    const s=window.REPORTIA_SESSION;
    const box=$('v516AdminPushTest'),btn=$('v516TestPush'),msg=$('v516TestPushMsg');
    if(!s?.token)throw new Error('Inicia sesión primero.');
    if(String(s.tipo||'').toUpperCase()!=='ADMINISTRADOR')throw new Error('Solo disponible para administrador.');
    if(btn)btn.disabled=true;
    if(msg)msg.textContent='Enviando prueba…';
    try{
      const d=await jsonp({accion:'v51_test_push',token:s.token},25000);
      if(!d?.ok)throw new Error(d?.mensaje||'No fue posible enviar la prueba.');
      if(msg){const dep=Number(d.depurados||0),fall=Number(d.fallidos||0),env=Number(d.enviados||0),dest=Number(d.destinatarios||0),admins=Number(d.administradores||0);let t=`Servidor ${d.serverVersion||'v53.3'} · administradores ${admins} · dispositivos ${dest} · aceptados ${env} · vencidos ${dep} · otros errores ${fall}.`;if(dep)t+=` Los tokens vencidos fueron desactivados automáticamente.`;if(dest===0)t+=` Abre REPORT.IA al menos una vez en cada dispositivo administrador para registrarlo.`;msg.textContent=t}
      return d;
    }finally{if(btn)btn.disabled=false}
  }
  function refreshAdminTestVisibility(){
    const s=window.REPORTIA_SESSION,box=$('v516AdminPushTest');
    if(box)box.hidden=String(s?.tipo||'').toUpperCase()!=='ADMINISTRADOR';
  }

  async function autoInit(){if(!window.REPORTIA_SESSION)return;try{await registerSW()}catch(_){}if(configured()&&Notification?.permission==='granted'){try{await initMessaging();currentToken=await messaging.getToken({vapidKey:window.REPORTIA_VAPID_KEY,serviceWorkerRegistration:swReg});if(currentToken){localStorage.setItem('reportia_push_token_v51',currentToken);await syncCurrentToken(currentToken,true)}}catch(_){}}}
  window.addEventListener('reportia:session',()=>{refreshAdminTestVisibility();autoInit()});document.addEventListener('DOMContentLoaded',()=>{$('v51EnablePush')?.addEventListener('click',handleEnable);$('v516TestPush')?.addEventListener('click',()=>testPush().catch(e=>{const m=$('v516TestPushMsg');if(m)m.textContent=e.message}));refreshAdminTestVisibility();registerSW().catch(()=>{})});
  window.REPORTIA_PUSH={enable,unregister,refreshStatus,configured,testPush,deviceStatus,testLocal};
})();
