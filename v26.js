(() => {
  "use strict";
  const $ = id => document.getElementById(id);
  const API_URL = String(window.REPORTIA_CONFIG?.API_URL || "").trim();
  const STATUS_LABEL = {PENDIENTE:"Pendiente",EN_ATENCION:"En atención",POR_VALIDAR:"Por validar",CERRADA:"Cerrada"};
  const STATUS_CLASS = {PENDIENTE:"pending",EN_ATENCION:"working",POR_VALIDAR:"validate",CERRADA:"closed"};
  let workflowCache = [];
  let cloudCache = null;

  function session(){ return window.REPORTIA_SESSION || null; }
  function isAdmin(){ return String(session()?.tipo||"").toUpperCase()==="ADMINISTRADOR"; }
  // Periodo mensual que debe consultar el centro de seguimiento en nube.
  // v26 llamaba a cloudPeriod() pero la función no existía, lo que provocaba
  // "No fue posible actualizar el centro de pendientes" y dejaba datos obsoletos en pantalla.
  function cloudPeriod(){
    const apiPeriod=window.REPORTIA_APP?.getCloudPeriod?.();
    if(apiPeriod!==null && apiPeriod!==undefined && Number.isFinite(Number(apiPeriod))) return Number(apiPeriod);
    const sel=document.getElementById("periodSelect");
    if(sel && Number.isFinite(Number(sel.value))) return Number(sel.value);
    return new Date().getMonth();
  }
  function esc(v){ return String(v??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c])); }
  function fmt(v){ const d=new Date(v); return Number.isNaN(d.getTime())?String(v||""):d.toLocaleString("es-MX"); }
  function roleName(tipo){
    const map={ADMINISTRADOR:"Administrador",ADMVA:"Administración",COMERCIAL:"Comercial",ESPECIALIZADO:"Especializado",GENERAL:"General","JEF SERV":"Jefatura de Servicio",MANTTO:"Mantenimiento","OP INTERM":"Operación Intermedia","OP PRIMERA":"Operación Primera","REC HUM":"Recursos Humanos"};
    return map[String(tipo||"").toUpperCase()]||tipo||"Usuario";
  }
  function jsonp(params){
    return new Promise((resolve,reject)=>{
      if(!API_URL)return reject(new Error("No está configurada la URL de Apps Script."));
      const callback="__reportia_v26_"+Date.now()+"_"+Math.random().toString(36).slice(2);
      const script=document.createElement("script");
      const query=new URLSearchParams({...params,callback});
      const timer=setTimeout(()=>{cleanup();reject(new Error("Tiempo de espera agotado."));},20000);
      function cleanup(){clearTimeout(timer);try{delete window[callback]}catch(_){ }script.remove();}
      window[callback]=data=>{cleanup();resolve(data)};
      script.onerror=()=>{cleanup();reject(new Error("No fue posible conectar con Apps Script."));};
      script.src=API_URL+(API_URL.includes("?")?"&":"?")+query.toString();
      document.head.appendChild(script);
    });
  }
  async function post(params){
    if(!API_URL)throw new Error("No está configurada la URL de Apps Script.");
    await fetch(API_URL,{method:"POST",mode:"no-cors",headers:{"Content-Type":"application/x-www-form-urlencoded;charset=UTF-8"},body:new URLSearchParams(params)});
  }
  async function cloudMeta(){
    const s=session(); if(!s?.token)return null;
    const d=await jsonp({accion:"cloud_meta",token:s.token,periodIndex:cloudPeriod()});
    if(!d?.ok)throw new Error(d?.mensaje||"No fue posible consultar la publicación.");
    cloudCache=d.disponible?d.cloud:null;
    return d;
  }
  function ageText(v){
    const d=new Date(v); if(Number.isNaN(d.getTime()))return "";
    const days=Math.max(0,Math.floor((Date.now()-d.getTime())/86400000));
    if(days===0)return "Abierta hoy"; if(days===1)return "Abierta hace 1 día"; return `Abierta hace ${days} días`;
  }
  function setCounts(resumen={}){
    $("v26Pending").textContent=resumen.PENDIENTE||0;
    $("v26Working").textContent=resumen.EN_ATENCION||0;
    $("v26Validate").textContent=resumen.POR_VALIDAR||0;
    $("v26Closed").textContent=resumen.CERRADA||0;
  }
  function renderInbox(items){
    const box=$("v26Inbox"), s=session(); if(!box)return;
    let relevant=[...items];
    if(isAdmin()) relevant.sort((a,b)=>({POR_VALIDAR:0,PENDIENTE:1,EN_ATENCION:2,CERRADA:3}[a.estado]??9)-({POR_VALIDAR:0,PENDIENTE:1,EN_ATENCION:2,CERRADA:3}[b.estado]??9));
    else relevant.sort((a,b)=>({PENDIENTE:0,EN_ATENCION:1,POR_VALIDAR:2,CERRADA:3}[a.estado]??9)-({PENDIENTE:0,EN_ATENCION:1,POR_VALIDAR:2,CERRADA:3}[b.estado]??9));
    const visible=relevant.filter(x=>x.estado!=="CERRADA").slice(0,8);
    if(!visible.length){box.innerHTML='<div class="empty">No tienes incidencias pendientes en la publicación vigente.</div>';return;}
    box.innerHTML=visible.map(x=>`<article class="v26-inbox-item ${STATUS_CLASS[x.estado]||"pending"}" data-v26-manager="${esc(x.gerencia)}"><span class="dot"></span><div><strong>${esc(x.gerencia)}</strong><p>${esc(STATUS_LABEL[x.estado]||x.estado)} · Prioridad ${esc(String(x.prioridad||"MEDIA").toLowerCase())}</p></div><em>${esc(ageText(x.creada))}</em></article>`).join("");
    $("v26InboxTitle").textContent=isAdmin()?"Pendientes de seguimiento":"Mis incidencias por atender";
  }
  async function loadWorkflow(){
    const s=session(); if(!s?.token)return;
    try{
      const meta=await cloudMeta();
      if(!meta?.disponible){ setCounts({}); workflowCache=[]; renderInbox([]); $("v26LastPublication").textContent="Sin publicación"; $("v26PublicationBy").textContent="El administrador debe publicar información."; return; }
      $("v26LastPublication").textContent=fmt(meta.cloud?.uploadedAt||"");
      $("v26PublicationBy").textContent=`Publicado por ${meta.cloud?.uploadedBy||"Administrador"}`;
      const d=await jsonp({accion:"workflow",token:s.token,snapshotId:meta.cloud?.snapshotId||"",periodIndex:cloudPeriod()});
      if(!d?.ok)throw new Error(d?.mensaje||"No fue posible consultar el seguimiento.");
      workflowCache=Array.isArray(d.incidencias)?d.incidencias:[];
      setCounts(d.resumen||{}); renderInbox(workflowCache);
      const role=roleName(s.tipo);
      $("v26Greeting").textContent=isAdmin()?`Hola, ${s.usuario}. Este es el estado de atención de RCV.`:`Hola, ${s.usuario}. Esto es lo que ${role} necesita atender.`;
      const open=(d.resumen?.PENDIENTE||0)+(d.resumen?.EN_ATENCION||0)+(d.resumen?.POR_VALIDAR||0);
      $("v26RoleSummary").textContent=isAdmin()?`${open} incidencia(s) siguen abiertas en la publicación vigente.`:`Tu vista está filtrada para ${role}. Tienes ${open} incidencia(s) abiertas.`;
    }catch(err){
      $("v26Inbox").innerHTML=`<div class="empty">${esc(err.message)}</div>`;
      $("v26RoleSummary").textContent="No fue posible actualizar el centro de pendientes.";
    }
  }
  function currentIncident(){
    const manager=$("incidentModal")?.dataset.manager||"";
    return workflowCache.find(x=>x.gerencia===manager)||null;
  }
  async function renderIncidentWorkflow(){
    const modal=$("incidentModal"), manager=modal?.dataset.manager; if(!manager||!modal.classList.contains("open"))return;
    const s=session(); if(!s?.token)return;
    try{
      const meta=cloudCache?{disponible:true,cloud:cloudCache}:await cloudMeta();
      if(!meta?.disponible)return;
      const d=await jsonp({accion:"workflow",token:s.token,snapshotId:meta.cloud?.snapshotId||"",gerencia:manager,periodIndex:cloudPeriod()});
      if(!d?.ok)throw new Error(d?.mensaje||"No fue posible consultar el estado.");
      const item=(d.incidencias||[])[0]||{gerencia:manager,estado:"PENDIENTE",prioridad:"MEDIA",creada:new Date().toISOString()};
      const ix=workflowCache.findIndex(x=>x.gerencia===manager); if(ix>=0)workflowCache[ix]=item; else workflowCache.push(item);
      $("v26IncidentWorkflowState").textContent=STATUS_LABEL[item.estado]||item.estado;
      $("v26IncidentAge").textContent=ageText(item.creada);
      document.querySelectorAll("#v26Workflow [data-flow]").forEach(el=>{el.classList.remove("active","done");});
      const order=["PENDIENTE","EN_ATENCION","POR_VALIDAR","CERRADA"], idx=order.indexOf(item.estado);
      order.forEach((st,i)=>{const el=document.querySelector(`#v26Workflow [data-flow="${st}"]`); if(el){if(i<idx)el.classList.add("done"); if(i===idx)el.classList.add("active");}});
      renderWorkflowActions(item);
    }catch(err){$("v26WorkflowActions").innerHTML=`<span class="empty">${esc(err.message)}</span>`;}
  }
  function renderWorkflowActions(item){
    const box=$("v26WorkflowActions"); if(!box)return;
    if(item.estado==="CERRADA"){box.innerHTML='<span class="empty">✓ Incidencia cerrada y validada.</span>';return;}
    if(isAdmin()){
      if(item.estado==="POR_VALIDAR") box.innerHTML='<button class="success" data-v26-state="CERRADA">✓ Aprobar cierre</button><button class="warning" data-v26-state="EN_ATENCION">↩ Solicitar corrección</button>';
      else if(item.estado==="PENDIENTE") box.innerHTML='<span class="empty">Esperando que el área inicie la atención.</span>';
      else if(item.estado==="EN_ATENCION") box.innerHTML='<span class="empty">El área está trabajando en la acción correctiva.</span>';
      else box.innerHTML='';
    }else{
      if(item.estado==="PENDIENTE") box.innerHTML='<button class="primary" data-v26-state="EN_ATENCION">▶ Iniciar atención</button>';
      else if(item.estado==="EN_ATENCION") box.innerHTML='<button class="primary" data-v26-state="POR_VALIDAR">➜ Enviar para validación</button>';
      else if(item.estado==="POR_VALIDAR") box.innerHTML='<span class="empty">En espera de validación del administrador.</span>';
      else box.innerHTML='';
    }
  }
  async function changeState(newState){
    const s=session(), manager=$("incidentModal")?.dataset.manager; if(!s?.token||!manager)return;
    const item=currentIncident()||{};
    try{
      const meta=cloudCache?{cloud:cloudCache}:await cloudMeta();
      const btn=document.querySelector(`[data-v26-state="${newState}"]`); if(btn){btn.disabled=true;btn.textContent="Guardando…";}
      await post({accion:"actualizar_estado",token:s.token,snapshotId:meta.cloud?.snapshotId||"",gerencia:manager,estado:newState,prioridad:item.prioridad||"MEDIA"});
      await new Promise(r=>setTimeout(r,750));
      await loadWorkflow(); await renderIncidentWorkflow();
    }catch(err){alert("No fue posible cambiar el estado: "+err.message);}
  }
  async function loadPublicationHistory(){
    const s=session(); if(!s?.token||!isAdmin())return;
    const box=$("v26PublicationHistory"), current=$("v26PubCurrent");
    try{
      const meta=await cloudMeta();
      current.innerHTML=meta?.disponible?`<strong>${esc(fmt(meta.cloud?.uploadedAt||""))}</strong><span>Publicada por ${esc(meta.cloud?.uploadedBy||"Administrador")} · periodo ${esc(meta.month||"—")}</span>`:'<div class="empty">Aún no hay publicación vigente.</div>';
      const d=await jsonp({accion:"publication_history",token:s.token});
      if(!d?.ok)throw new Error(d?.mensaje||"No fue posible consultar el historial.");
      const rows=Array.isArray(d.publicaciones)?d.publicaciones:[];
      box.innerHTML=rows.length?rows.map((x,i)=>`<article class="v26-history-item"><div><strong>${i===0?"Publicación vigente":"Publicación anterior"}</strong><span>${esc(fmt(x.fecha))} · ${esc(x.usuario)}</span></div><em>${esc(x.periodo||"—")} · ${esc(x.modo||"—")}</em></article>`).join(""):'<div class="empty">Todavía no hay publicaciones registradas.</div>';
    }catch(err){box.innerHTML=`<div class="empty">${esc(err.message)}</div>`;}
  }
  function openPublication(){ if(!isAdmin())return; $("v26PublicationModal").classList.add("open"); loadPublicationHistory(); }
  function closePublication(){ $("v26PublicationModal").classList.remove("open"); }

  document.addEventListener("click", async e=>{
    if(e.target.closest("#v26ToggleDetail")){
      const box=$("v26LegacyHome"), btn=$("v26ToggleDetail"); box.classList.toggle("open"); btn.textContent=box.classList.contains("open")?"Ocultar análisis ejecutivo ↑":"Mostrar análisis ejecutivo completo ↓"; return;
    }
    if(e.target.closest("#v26RefreshWorkflow")){await loadWorkflow();return;}
    if(e.target.closest("#v26OpenPublication")){openPublication();return;}
    if(e.target.closest("#v26RefreshHistory")){await loadPublicationHistory();return;}
    if(e.target.closest("[data-close-v26-publication]")){closePublication();return;}
    const wi=e.target.closest("[data-v26-manager]"); if(wi){window.REPORTIA_APP?.openIncident?.(wi.dataset.v26Manager); setTimeout(renderIncidentWorkflow,120); return;}
    const st=e.target.closest("[data-v26-state]"); if(st){await changeState(st.dataset.v26State);return;}
    if(e.target.closest("#saveIncidentNote")){setTimeout(async()=>{await loadWorkflow();await renderIncidentWorkflow();},1200);}
  });

  const obs=new MutationObserver(()=>{const m=$("incidentModal");if(m?.classList.contains("open"))renderIncidentWorkflow();});
  if($("incidentModal"))obs.observe($("incidentModal"),{attributes:true,attributeFilter:["class","data-manager"]});
  function renderProcessedModel(){
    const health=window.REPORTIA_APP?.getManagerHealth?.()||[];
    if(!health.length){
      $("v26RoleSummary").textContent="La información fue procesada, pero no se encontraron gerencias utilizables en el modelo.";
      return;
    }
    const critical=health.filter(x=>x.status==="Crítico");
    const attention=health.filter(x=>x.status==="Atención");
    const favorable=health.filter(x=>x.status==="Favorable");
    const pending=[...critical,...attention];
    // Antes de publicar, mostramos lo recién detectado sin confundirlo con el flujo ya guardado en nube.
    $("v26Pending").textContent=pending.length;
    $("v26Working").textContent="0";
    $("v26Validate").textContent="0";
    $("v26Closed").textContent="0";
    $("v26Greeting").textContent=`Información procesada correctamente · ${health.length} gerencia(s) analizadas.`;
    $("v26RoleSummary").textContent=pending.length
      ? `${critical.length} crítica(s) y ${attention.length} en atención fueron detectadas. Publica en la nube para convertirlas en incidencias de seguimiento.`
      : `${favorable.length} gerencia(s) se encuentran en condición favorable. Puedes revisar el análisis ejecutivo o publicar la información.`;
    const box=$("v26Inbox");
    $("v26InboxTitle").textContent="Resultado recién procesado";
    if(!pending.length){
      box.innerHTML='<div class="empty">✓ No se detectaron incidencias que requieran seguimiento con los umbrales actuales.</div>';
    }else{
      box.innerHTML=pending.slice(0,8).map(x=>`<article class="v26-inbox-item ${x.status==="Crítico"?"pending":"working"}" data-manager-detail="${encodeURIComponent(x.manager)}"><span class="dot"></span><div><strong>${esc(x.manager)}</strong><p>${esc(x.status)} · presión ${Number.isFinite(x.pressure)?(x.pressure*100).toFixed(1)+"%":"—"}</p></div><em>Pendiente de publicar</em></article>`).join("");
    }
    const badge=$("cloudStatusBadge");
    if(badge){badge.textContent="● Datos procesados localmente · pendientes de publicar";badge.className="cloud-status-badge working";}
  }

  window.addEventListener("reportia:model-processed",()=>setTimeout(renderProcessedModel,50));
  window.addEventListener("reportia:cloud-published",async()=>{await loadWorkflow();if(isAdmin())await loadPublicationHistory();});
  window.addEventListener("reportia:session",()=>setTimeout(loadWorkflow,900));
  window.addEventListener("focus",()=>{if(session()?.token)loadWorkflow();});
  if(session()?.token)setTimeout(loadWorkflow,900);
})();
