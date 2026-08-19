(()=>{
  const $=id=>document.getElementById(id);
  const qa=(s,r=document)=>[...r.querySelectorAll(s)];
  const session=()=>window.REPORTIA_SESSION||null;
  const isAdmin=()=>String(session()?.tipo||'').toUpperCase()==='ADMINISTRADOR';
  let currentTab='realtime';

  function healthFor(source){
    const app=window.REPORTIA_APP;
    if(!app)return [];
    const wanted=source==='cloud'?'cloud':'local';
    const active=app.getActiveSource?.();
    let restore=active;
    if(wanted==='local') app.useLocalModel?.(); else app.useCloudModel?.();
    const h=app.getManagerHealth?.()||[];
    if(restore==='local') app.useLocalModel?.();
    else if(restore==='cloud') app.useCloudModel?.();
    return h;
  }

  function updateWorkspace(source=currentTab){
    const isCloud=source==='cloud';
    const model=isCloud?window.REPORTIA_APP?.getCloudModel?.():window.REPORTIA_APP?.getLocalModel?.();
    const h=model?healthFor(source):[];
    const c=h.filter(x=>x.status==='Crítico').length;
    const a=h.filter(x=>x.status==='Atención').length;
    const f=h.filter(x=>x.status==='Favorable').length;

    const dot=$(isCloud?'v27CloudDot':'v27RealtimeDot');
    if(isCloud){
      const cloud=window.REPORTIA_APP?.getCloud?.()||{};
      if($('v27CloudDot')){$('v27CloudDot').textContent=cloud.snapshotId?'PUBLICADA':'SIN PUBLICACIÓN';$('v27CloudDot').className=cloud.snapshotId?'cloud':'';}
    }else if($('v27RealtimeDot')){
      $('v27RealtimeDot').textContent=model?'EN VIVO':'SIN CARGA';
      $('v27RealtimeDot').className=model?'live':'';
    }

    if($('v27LocalManagers'))$('v27LocalManagers').textContent=h.length;
    if($('v27LocalCritical'))$('v27LocalCritical').textContent=c;
    if($('v27LocalAttention'))$('v27LocalAttention').textContent=a;
    if($('v27LocalFavorable'))$('v27LocalFavorable').textContent=f;

    const live=$('v27LiveStatus');
    const title=$('v27RealtimeTitle');
    const text=$('v27RealtimeText');
    const kicker=document.querySelector('.v27-realtime-title>span');
    if(kicker) kicker.textContent=isCloud?'CLOUD ANALYTICS · PUBLICACIÓN OFICIAL':'LIVE ANALYTICS · DATOS LOCALES';

    if(!model){
      if(live)live.textContent=isCloud?'Sin publicación seleccionada':'Esperando procesamiento';
      if(title)title.textContent=isCloud?'Selecciona un periodo publicado en la nube':'Carga y procesa los Excel para iniciar';
      if(text)text.textContent=isCloud?'La misma vista ejecutiva mostrará aquí la información oficial publicada, sin cambiar de menú.':'Aquí verás inmediatamente el resultado completo del archivo procesado, antes de publicarlo.';
    }else{
      if(live)live.textContent=isCloud?'Publicación oficial cargada':'Modelo local procesado y listo para revisar';
      if(title)title.textContent=`${isCloud?'Información oficial en la nube':'Análisis en tiempo real'} · ${h.length} gerencia${h.length===1?'':'s'} procesadas`;
      if(text)text.textContent=`${isCloud?'Resultado publicado':'Resultado local listo'}: ${c} crítica${c===1?'':'s'}, ${a} en atención y ${f} favorable${f===1?'':'s'}.`;
    }

    const shared=$('v27CloudShared');
    if(shared) shared.style.display=isCloud?'block':'none';
    const liveActions=document.querySelector('.v27-live-actions');
    if(liveActions) liveActions.style.display=isCloud?'none':'flex';
  }

  function setTab(name,{switchModel=true}={}){
    if(name==='realtime'&&!isAdmin()) name='cloud';
    currentTab=name;
    qa('[data-v27-tab]').forEach(b=>b.classList.toggle('active',b.dataset.v27Tab===name));
    // v30: ambos orígenes usan el MISMO tablero. Ya no cambiamos de diseño ni movemos menús.
    qa('[data-v27-panel]').forEach(p=>p.classList.toggle('active',p.dataset.v27Panel==='realtime'));
    if(switchModel){
      if(name==='realtime') window.REPORTIA_APP?.useLocalModel?.();
      else window.REPORTIA_APP?.useCloudModel?.();
    }
    document.getElementById('v27Command')?.setAttribute('data-source',name);
    localStorage.setItem('reportia_v27_tab',name);
    updateWorkspace(name);
    if(name==='cloud') setTimeout(()=>document.getElementById('v26RefreshWorkflow')?.click(),40);
  }

  document.addEventListener('click',e=>{
    const tab=e.target.closest('[data-v27-tab]');
    if(tab){setTab(tab.dataset.v27Tab);return;}
    if(e.target.closest('#v27PublishQuick')){window.REPORTIA_APP?.publishCloud?.();return;}
  });

  window.addEventListener('reportia:model-processed',()=>{setTab('realtime',{switchModel:false});});
  window.addEventListener('reportia:cloud-published',()=>{if(currentTab==='cloud')setTab('cloud',{switchModel:false});else updateWorkspace('realtime');});
  // v31: al cargar/cambiar un mes publicado, refrescar también los KPIs superiores
  // del tablero unificado. Antes la zona inferior cambiaba al modelo nube,
  // pero las tarjetas superiores conservaban los 0 del estado previo.
  window.addEventListener('reportia:cloud-month-changed',()=>{
    if(currentTab==='cloud'){
      setTimeout(()=>updateWorkspace('cloud'),80);
    }
  });
  window.addEventListener('reportia:session',()=>{
    setTimeout(()=>{
      const preferred=localStorage.getItem('reportia_v27_tab');
      setTab(isAdmin()?(preferred==='cloud'?'cloud':'realtime'):'cloud',{switchModel:true});
    },1100);
  });

  const oldToggle=$('v26ToggleDetail'); if(oldToggle)oldToggle.style.display='none';
  setTimeout(()=>{if(session())setTab(isAdmin()?'realtime':'cloud',{switchModel:!isAdmin()});else updateWorkspace('realtime');},1200);
})();
