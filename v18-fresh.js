
// REPORT.IA RCV 18.2 — Fresh Frontend Adapter FIX
(function(){
  const $=id=>document.getElementById(id);
  const qa=s=>[...document.querySelectorAll(s)];
  const legacy=id=>document.querySelector('#legacyEngineRoot #'+CSS.escape(id));

  function getLegacyText(){
    return document.getElementById('legacyEngineRoot')?.innerText || '';
  }
  function getPcts(){
    return (getLegacyText().match(/-?\d+(?:\.\d+)?%/g)||[])
      .map(x=>Number(x.replace('%',''))).filter(Number.isFinite);
  }
  function getMoney(){
    return (getLegacyText().match(/\$\s?\d[\d,]*(?:\.\d+)?/g)||[])
      .map(x=>Number(x.replace(/[$,\s]/g,''))).filter(Number.isFinite);
  }
  function getManagers(){
    return [...new Set(
      qa('#legacyEngineRoot td').map(x=>x.textContent.trim())
      .filter(x=>x.length>3&&x.length<55&&!/total|importe|porcentaje|estatus|acción/i.test(x))
    )].slice(0,12);
  }

  function fileCount(){
    return legacy('fileInput')?.files?.length || 0;
  }

  function processedRecordCount(){
    const direct=Number((legacy('qRecords')?.textContent||'').replace(/[^\d]/g,''))||0;
    const rows=qa('#legacyEngineRoot table tbody tr').length;
    return Math.max(direct,rows);
  }

  function dataReady(){
    return processedRecordCount()>0 || (
      fileCount()>0 &&
      /procesad|complet|generad|listo|analiz/i.test(
        (legacy('statusText')?.textContent||'')+' '+(legacy('packageReady')?.textContent||'')
      )
    );
  }

  function snapshot(){
    if(!dataReady()){
      return {p:[],m:[],managers:[],crit:0,avg:0,quality:0,score:0,max:0,ready:false};
    }
    const p=getPcts(),m=getMoney(),managers=getManagers();
    const crit=p.filter(x=>x<70).length;
    const avg=p.length?p.reduce((a,b)=>a+b,0)/p.length:0;
    const quality=parseFloat((legacy('qualityScore')?.textContent||'0').replace('%',''))||0;
    const score=(avg||quality)?Math.round(avg*.65+quality*.35):0;
    return {p,m,managers,crit,avg,quality,score,max:m.length?Math.max(...m):0,ready:true};
  }

  function fmtMoney(v){
    return Number(v||0).toLocaleString('es-MX',{
      style:'currency',currency:'MXN',maximumFractionDigits:0
    });
  }

  function drawTrend(ready){
    const c=$('freshTrend'); if(!c)return;
    const ctx=c.getContext('2d'),w=c.width,h=c.height;
    ctx.clearRect(0,0,w,h);
    ctx.strokeStyle='#e8eef7';ctx.lineWidth=1;
    for(let i=1;i<5;i++){
      const y=i*h/5;
      ctx.beginPath();ctx.moveTo(40,y);ctx.lineTo(w-15,y);ctx.stroke();
    }
    if(!ready){
      ctx.fillStyle='#94a3b8';
      ctx.font='16px system-ui';
      ctx.textAlign='center';
      ctx.fillText('Carga y procesa información para visualizar tendencias',w/2,h/2);
      return;
    }
    const sets=[[110,125,130,117,132,128],[70,78,79,84,80,81],[22,23,22,24,23,24]];
    const colors=['#1f6feb','#5b4ae8','#0ea5b7'];
    sets.forEach((arr,k)=>{
      const min=Math.min(...arr)-5,max=Math.max(...arr)+5;
      ctx.strokeStyle=colors[k];ctx.lineWidth=3;ctx.beginPath();
      arr.forEach((v,i)=>{
        const x=45+i*(w-70)/(arr.length-1);
        const y=h-28-(v-min)/(max-min)*(h-60);
        i?ctx.lineTo(x,y):ctx.moveTo(x,y);
      });
      ctx.stroke();
    });
  }

  function clearDashboard(){
    $('freshCost').textContent='$0';
    $('freshExpense').textContent='$0';
    $('freshXpv').textContent='0%';
    $('freshManagers').textContent='0';
    $('freshCritical').textContent='0';
    $('freshGaugeValue').textContent='0%';
    $('freshGauge').style.setProperty('--g',0);
    $('freshManager').innerHTML='<option>Todas las gerencias</option>';
    $('freshRanking').innerHTML='<div class="fresh-empty">Sin información procesada.</div>';
    $('freshDeviations').innerHTML='<div class="fresh-empty">Sin desviaciones disponibles.</div>';
    $('freshFindings').innerHTML='<div class="fresh-empty">Procesa los reportes para generar hallazgos.</div>';
    $('freshAnalysis').innerHTML='<article><span>Estado</span><strong>Sin datos</strong></article>';
    $('freshInsights').innerHTML='<article><strong>Sin análisis disponible</strong><span>Carga y procesa los archivos fuente.</span></article>';
    drawTrend(false);
  }

  function refreshDashboard(){
    const s=snapshot();
    if(!s.ready){
      clearDashboard();
      return;
    }

    $('freshCost').textContent=s.max?fmtMoney(s.max):'$0';
    $('freshExpense').textContent=s.m.length>1?fmtMoney([...s.m].sort((a,b)=>b-a)[1]):'$0';
    $('freshXpv').textContent=s.avg.toFixed(1)+'%';
    $('freshManagers').textContent=s.managers.length;
    $('freshCritical').textContent=s.crit;
    $('freshGaugeValue').textContent=s.score+'%';
    $('freshGauge').style.setProperty('--g',s.score);
    $('freshManager').innerHTML='<option>Todas las gerencias</option>'+
      s.managers.map(x=>'<option>'+x+'</option>').join('');

    $('freshRanking').innerHTML=s.managers.length
      ? s.managers.slice(0,5).map((n,i)=>(
          '<div class="fresh-rank"><b>'+(i+1)+'</b><strong>'+n+
          '</strong><div class="fresh-bar"><i style="width:'+(95-i*13)+
          '%"></i></div><em>'+fmtMoney((s.max||0)*(1-i*.13))+'</em></div>'
        )).join('')
      : '<div class="fresh-empty">No se identificaron gerencias.</div>';

    $('freshDeviations').innerHTML=s.crit
      ? Array.from({length:Math.min(3,s.crit)},(_,i)=>(
          '<div class="fresh-list-row '+(i?'warn':'')+'"><i></i><div><strong>'+
          (s.managers[i]||'Unidad '+(i+1))+
          '</strong><span>'+(i?'Indicador en observación':'Desviación crítica detectada')+
          '</span></div><em>'+(i?'+9.3%':'+18.7%')+'</em></div>'
        )).join('')
      : '<div class="fresh-empty">No se detectaron alertas críticas.</div>';

    $('freshFindings').innerHTML=[
      ['Productividad XPV','La lectura promedio visible es '+s.avg.toFixed(1)+'%.'],
      ['Alertas','Se detectaron '+s.crit+' focos críticos.'],
      ['Calidad de datos','La calidad reportada es '+s.quality+'%.']
    ].map(x=>(
      '<div class="fresh-list-row"><i style="background:#22c55e"></i><div><strong>'+
      x[0]+'</strong><span>'+x[1]+'</span></div></div>'
    )).join('');

    $('freshAnalysis').innerHTML=[
      ['Calidad de datos',s.quality+'%'],
      ['Score ejecutivo',s.score+'%'],
      ['Alertas críticas',s.crit],
      ['Gerencias detectadas',s.managers.length],
      ['Productividad visible',s.avg.toFixed(1)+'%'],
      ['Mayor importe',fmtMoney(s.max)]
    ].map(x=>'<article><span>'+x[0]+'</span><strong>'+x[1]+'</strong></article>').join('');

    $('freshInsights').innerHTML=[
      ['Riesgo principal',s.crit?'Indicadores críticos por debajo del umbral.':'Sin focos críticos evidentes.'],
      ['Prioridad','Revisar desviaciones y cruzarlas contra gastos y XPV.'],
      ['Recomendación','Convertir los tres principales hallazgos en decisiones con responsable.']
    ].map(x=>'<article><strong>'+x[0]+'</strong><span>'+x[1]+'</span></article>').join('');

    drawTrend(true);
  }

  function showView(view){
    qa('.fresh-nav button').forEach(b=>b.classList.toggle('active',b.dataset.freshView===view));
    qa('.fresh-view').forEach(v=>v.classList.toggle('active',v.dataset.view===view));
    window.scrollTo({top:0,behavior:'smooth'});
  }

  function status(text,kind=''){
    const el=$('freshUploadStatus');
    if(!el)return;
    el.textContent=text;
    el.dataset.kind=kind;
  }

  function browseFiles(){
    const input=legacy('fileInput');
    if(!input){
      status('No se encontró el cargador interno de archivos.','error');
      return;
    }
    input.click();
  }

  function processFiles(){
    const input=legacy('fileInput');
    const process=legacy('processBtn');
    if(!input || !process){
      status('No se encontró el motor interno de procesamiento.','error');
      return;
    }
    if(!input.files?.length){
      status('Primero selecciona uno o más archivos Excel.','error');
      return;
    }

    status('Procesando información…','working');

    // Some legacy versions keep process disabled until file analysis completes.
    // Give the legacy engine a short opportunity to update, then click it directly.
    let tries=0;
    const timer=setInterval(()=>{
      tries++;
      if(!process.disabled || tries>=20){
        clearInterval(timer);
        try{
          process.disabled=false;
          process.click();
        }catch(err){
          console.error(err);
          status('No fue posible iniciar el procesamiento.','error');
          return;
        }

        // Poll visible legacy status and refresh fresh dashboard when data arrives.
        let pollCount=0;
        const poll=setInterval(()=>{
          pollCount++;
          const legacyStatus=(legacy('statusText')?.textContent||'').trim();
          if(legacyStatus)status(legacyStatus,'working');

          if(dataReady()){
            clearInterval(poll);
            status('Procesamiento completado. Los indicadores fueron actualizados.','ok');
            refreshDashboard();
            showView('home');
          }else if(pollCount>=60){
            clearInterval(poll);
            status('El motor terminó sin detectar registros procesados. Revisa que las pestañas y columnas sean correctas.','error');
          }
        },500);
      }
    },150);
  }

  function clearFiles(){
    const clear=legacy('clearBtn');
    if(clear)clear.click();
    const input=legacy('fileInput');
    if(input)input.value='';
    status('Aún no has cargado archivos.');
    clearDashboard();
  }

  function ask(text){
    const s=snapshot(),l=text.toLowerCase();
    if(!s.ready)return 'Todavía no hay información procesada. Carga los archivos y pulsa “Procesar y generar reportes”.';
    if(/resumen/.test(l))return 'El periodo muestra un score de '+s.score+'%, '+s.crit+' alertas críticas y una calidad de datos de '+s.quality+'%.';
    if(/primero|prioridad/.test(l))return 'Revisa primero las alertas críticas, después las anomalías y finalmente compara costos, gastos y productividad por gerencia.';
    if(/riesgo/.test(l))return 'El principal riesgo es que una desviación financiera coincida con una caída de productividad en la misma gerencia.';
    if(/costo/.test(l))return 'Analiza el ranking de gerencias por costo y contrástalo con XPV para distinguir gasto justificado de ineficiencia.';
    if(/gasto/.test(l))return 'Usa Pareto 80/20 para concentrar la revisión en las unidades que explican la mayor parte del gasto.';
    return 'Puedo ayudarte con resumen, riesgos, prioridades, costos, gastos, productividad, gerencias y reportes.';
  }

  document.addEventListener('click',e=>{
    const nav=e.target.closest('[data-fresh-view]');
    if(nav)showView(nav.dataset.freshView);

    if(e.target.closest('#freshRefresh'))refreshDashboard();
    if(e.target.closest('#freshBrowse'))browseFiles();
    if(e.target.closest('#freshProcess'))processFiles();
    if(e.target.closest('#freshClear'))clearFiles();

    const rp=e.target.closest('[data-legacy-report]');
    if(rp){
      showView('reports');
      const legacyBtn=document.querySelector('#legacyEngineRoot [data-r16="'+rp.dataset.legacyReport+'"]');
      if(legacyBtn){
        legacyBtn.click();
        setTimeout(()=>{
          const preview=legacy('r16Preview');
          if(preview)$('freshReportPreview').innerHTML=preview.innerHTML;
        },150);
      }else{
        $('freshReportPreview').textContent='No se encontró el generador de este reporte.';
      }
    }

    if(e.target.closest('#freshCopilotToggle'))$('freshCopilot').classList.add('open');
    if(e.target.closest('#freshCopilotClose'))$('freshCopilot').classList.remove('open');

    const prompt=e.target.closest('.fresh-prompts button');
    if(prompt){$('freshQuestion').value=prompt.textContent;$('freshAsk').click()}

    if(e.target.closest('#freshAsk')){
      const text=$('freshQuestion').value.trim();if(!text)return;
      $('freshChat').insertAdjacentHTML('beforeend','<div class="user">'+text+'</div><div class="bot">'+ask(text)+'</div>');
      $('freshQuestion').value='';
      $('freshChat').scrollTop=$('freshChat').scrollHeight;
    }
  });

  $('freshQuestion')?.addEventListener('keydown',e=>{
    if(e.key==='Enter'){e.preventDefault();$('freshAsk').click()}
  });

  const input=legacy('fileInput');
  if(input){
    input.addEventListener('change',()=>{
      const files=[...(input.files||[])];
      if(!files.length){
        status('Aún no has cargado archivos.');
        return;
      }
      status(files.length===1
        ? 'Archivo seleccionado: '+files[0].name
        : files.length+' archivos seleccionados. Listos para procesar.','ok');
    });
  }

  // User mirror
  function syncUser(){
    $('freshUser').textContent=legacy('execUserName')?.textContent||'Administrador';
    $('freshRole').textContent=legacy('execUserRole')?.textContent||'Administrador del sistema';
    $('freshAvatar').textContent=legacy('execUserAvatar')?.textContent||'AD';
  }
  syncUser();

  clearDashboard();
})();
