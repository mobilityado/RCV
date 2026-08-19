(()=>{
  const $=id=>document.getElementById(id);
  const qa=(s,r=document)=>[...r.querySelectorAll(s)];
  const session=()=>window.REPORTIA_SESSION||null;
  const isAdmin=()=>String(session()?.tipo||'').toUpperCase()==='ADMINISTRADOR';

  function setTab(name,{switchModel=true}={}){
    if(name==='realtime'&&!isAdmin()) name='cloud';
    qa('[data-v27-tab]').forEach(b=>b.classList.toggle('active',b.dataset.v27Tab===name));
    qa('[data-v27-panel]').forEach(p=>p.classList.toggle('active',p.dataset.v27Panel===name));
    if(switchModel){
      if(name==='realtime') window.REPORTIA_APP?.useLocalModel?.();
      else window.REPORTIA_APP?.useCloudModel?.();
    }
    localStorage.setItem('reportia_v27_tab',name);
    if(name==='cloud') setTimeout(()=>document.getElementById('v26RefreshWorkflow')?.click(),40);
  }

  function localHealth(){
    if(!window.REPORTIA_APP?.getLocalModel?.()) return [];
    const active=window.REPORTIA_APP?.getActiveSource?.();
    let restore=false;
    if(active!=='local'){restore=true;window.REPORTIA_APP?.useLocalModel?.();}
    const h=window.REPORTIA_APP?.getManagerHealth?.()||[];
    if(restore) window.REPORTIA_APP?.useCloudModel?.();
    return h;
  }

  function updateRealtime(){
    const has=!!window.REPORTIA_APP?.getLocalModel?.();
    const dot=$('v27RealtimeDot'), live=$('v27LiveStatus');
    if(!has){
      if(dot){dot.textContent='SIN CARGA';dot.className='';}
      if(live)live.textContent='Esperando procesamiento';
      ['v27LocalManagers','v27LocalCritical','v27LocalAttention','v27LocalFavorable'].forEach(id=>{if($(id))$(id).textContent='0';});
      if($('v27RealtimeTitle'))$('v27RealtimeTitle').textContent='Carga y procesa los Excel para iniciar';
      if($('v27RealtimeText'))$('v27RealtimeText').textContent='Aquí verás inmediatamente el resultado completo del archivo procesado, antes de publicarlo.';
      return;
    }
    const h=localHealth();
    const c=h.filter(x=>x.status==='Crítico').length;
    const a=h.filter(x=>x.status==='Atención').length;
    const f=h.filter(x=>x.status==='Favorable').length;
    if(dot){dot.textContent='EN VIVO';dot.className='live';}
    if(live)live.textContent='Modelo local procesado y listo para revisar';
    if($('v27LocalManagers'))$('v27LocalManagers').textContent=h.length;
    if($('v27LocalCritical'))$('v27LocalCritical').textContent=c;
    if($('v27LocalAttention'))$('v27LocalAttention').textContent=a;
    if($('v27LocalFavorable'))$('v27LocalFavorable').textContent=f;
    if($('v27RealtimeTitle'))$('v27RealtimeTitle').textContent=`Análisis en tiempo real · ${h.length} gerencia${h.length===1?'':'s'} procesadas`;
    if($('v27RealtimeText'))$('v27RealtimeText').textContent=`Resultado local listo: ${c} crítica${c===1?'':'s'}, ${a} en atención y ${f} favorable${f===1?'':'s'}. Revisa KPIs, gráficas y hallazgos antes de publicar.`;
  }

  function updateCloudBadge(){
    const cloud=window.REPORTIA_APP?.getCloud?.()||{};
    const dot=$('v27CloudDot');
    if(!dot)return;
    if(cloud.snapshotId){dot.textContent='PUBLICADA';dot.className='cloud';}
    else {dot.textContent='SIN PUBLICACIÓN';dot.className='';}
  }

  document.addEventListener('click',e=>{
    const tab=e.target.closest('[data-v27-tab]');
    if(tab){setTab(tab.dataset.v27Tab);return;}
    if(e.target.closest('#v27PublishQuick')){window.REPORTIA_APP?.publishCloud?.();return;}
  });

  window.addEventListener('reportia:model-processed',()=>{
    updateRealtime();
    setTab('realtime',{switchModel:false});
  });
  window.addEventListener('reportia:cloud-published',()=>{updateCloudBadge();});
  window.addEventListener('reportia:session',()=>{
    setTimeout(()=>{
      updateRealtime();updateCloudBadge();
      const preferred=localStorage.getItem('reportia_v27_tab');
      setTab(isAdmin()?(preferred==='cloud'?'cloud':'realtime'):'cloud',{switchModel:true});
    },1100);
  });

  // La vista local ya contiene el tablero validado completo: no se colapsa en v27.
  const oldToggle=$('v26ToggleDetail'); if(oldToggle)oldToggle.style.display='none';
  setTimeout(()=>{updateRealtime();updateCloudBadge();if(session())setTab(isAdmin()?'realtime':'cloud',{switchModel:true});},1200);
})();
