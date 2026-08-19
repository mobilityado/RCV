(()=>{
  'use strict';
  const $=id=>document.getElementById(id);
  const session=()=>window.REPORTIA_SESSION||null;
  const isAdmin=()=>String(session()?.tipo||'').toUpperCase()==='ADMINISTRADOR';
  const API_URL=()=>String(window.REPORTIA_CONFIG?.API_URL||'').trim();
  const MONTHS=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];

  function jsonp(params){
    return new Promise((resolve,reject)=>{
      const url=API_URL(); if(!url)return reject(new Error('No está configurada la URL de Apps Script.'));
      const cb='__reportia_v28_'+Date.now()+'_'+Math.random().toString(36).slice(2);
      const sc=document.createElement('script');
      const timer=setTimeout(()=>{cleanup();reject(new Error('Tiempo de espera agotado.'));},25000);
      function cleanup(){clearTimeout(timer);try{delete window[cb]}catch(_){}sc.remove();}
      window[cb]=d=>{cleanup();resolve(d)};
      sc.onerror=()=>{cleanup();reject(new Error('No fue posible consultar Apps Script.'));};
      sc.src=url+(url.includes('?')?'&':'?')+new URLSearchParams({...params,callback:cb}).toString();
      document.head.appendChild(sc);
    });
  }

  function setStatus(text,type=''){
    const el=$('v28MonthStatus');if(!el)return;
    el.textContent=text;el.className='v28-month-status'+(type?' '+type:'');
  }

  async function refreshMonths({autoload=false}={}){
    const s=session(); if(!s?.token)return;
    try{
      setStatus('Consultando disponibilidad mensual…','working');
      const d=await jsonp({accion:'cloud_months',token:s.token});
      if(!d?.ok)throw new Error(d?.mensaje||'No fue posible consultar los meses.');
      const rows=Array.isArray(d.meses)?d.meses:[];
      const select=$('v28CloudMonth'); if(!select)return;
      const current=Number(select.value||0);
      rows.forEach(x=>{
        const opt=select.querySelector(`option[value="${x.periodIndex}"]`);
        if(!opt)return;
        opt.textContent=x.disponible?`${MONTHS[x.periodIndex]} · ${x.anio||'publicado'}`:`${MONTHS[x.periodIndex]} · sin publicación`;
        opt.dataset.available=x.disponible?'1':'0';
      });
      const available=rows.filter(x=>x.disponible);
      // Usuarios operativos reciben automáticamente la publicación más reciente.
      // El administrador NO cambia de fuente mientras carga/procesa archivos locales.
      if(autoload&&available.length&&!isAdmin()){
        const newest=available.slice().sort((a,b)=>new Date(b.fecha)-new Date(a.fecha))[0];
        if(newest){select.value=String(newest.periodIndex);await loadSelectedMonth();return;}
      }
      select.value=String(current);
      const count=available.length;
      setStatus(`${count} de 12 meses tienen información publicada. Selecciona un mes para consultarlo.`,count?'ok':'');
    }catch(err){setStatus(err.message,'error');}
  }

  async function loadSelectedMonth(){
    const s=session();if(!s?.token)return;
    const idx=Number($('v28CloudMonth')?.value||0);
    setStatus(`Cargando ${MONTHS[idx]} desde la nube…`,'working');
    await window.REPORTIA_APP?.loadCloudForSession?.(s,idx);
    const cloud=window.REPORTIA_APP?.getCloud?.()||{};
    if(cloud.snapshotId){
      const when=cloud.uploadedAt?new Date(cloud.uploadedAt).toLocaleString('es-MX'):'';
      setStatus(`${MONTHS[idx]} cargado correctamente${when?' · publicado '+when:''}.`,'ok');
      window.dispatchEvent(new CustomEvent('reportia:cloud-month-changed',{detail:{periodIndex:idx}}));
      document.getElementById('v26RefreshWorkflow')?.click();
    }else setStatus(`${MONTHS[idx]} todavía no tiene una publicación guardada.`,'');
  }

  async function downloadBackup(){
    const s=session();if(!s?.token)return alert('Inicia sesión nuevamente.');
    try{
      const btn=$('v28BackupExcel');if(btn){btn.disabled=true;btn.textContent='Preparando respaldo…';}
      const d=await jsonp({accion:'backup_url',token:s.token});
      if(!d?.ok||!d.url)throw new Error(d?.mensaje||'No fue posible preparar el respaldo.');
      const a=document.createElement('a');a.href=d.url;a.target='_blank';a.rel='noopener';document.body.appendChild(a);a.click();a.remove();
      setStatus('Se abrió la descarga del Excel completo con todas las pestañas mensuales.','ok');
    }catch(err){alert('No fue posible descargar el respaldo: '+err.message);}
    finally{const btn=$('v28BackupExcel');if(btn){btn.disabled=false;btn.textContent='⬇ Respaldo Excel completo';}}
  }

  document.addEventListener('click',e=>{
    if(e.target.closest('#v28LoadCloudMonth')){loadSelectedMonth();return;}
    if(e.target.closest('#v28BackupExcel')){downloadBackup();return;}
  });
  $('v28CloudMonth')?.addEventListener('change',loadSelectedMonth);

  window.addEventListener('reportia:session',()=>setTimeout(()=>refreshMonths({autoload:true}),1300));
  window.addEventListener('reportia:cloud-published',()=>setTimeout(()=>refreshMonths(),700));
  setTimeout(()=>{if(session())refreshMonths({autoload:true});},1500);
})();
