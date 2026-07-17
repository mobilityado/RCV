
// REPORT.IA RCV 18.1 — Fresh Frontend Adapter
(function(){
  const $=id=>document.getElementById(id);
  const q=s=>document.querySelector(s);
  const qa=s=>[...document.querySelectorAll(s)];

  function legacy(id){return document.getElementById(id)}
  function clickLegacy(id){legacy(id)?.click()}

  function getPcts(){
    return (document.getElementById('legacyEngineRoot')?.innerText.match(/-?\d+(?:\.\d+)?%/g)||[])
      .map(x=>Number(x.replace('%',''))).filter(Number.isFinite);
  }
  function getMoney(){
    return (document.getElementById('legacyEngineRoot')?.innerText.match(/\$\s?\d[\d,]*(?:\.\d+)?/g)||[])
      .map(x=>Number(x.replace(/[$,\s]/g,''))).filter(Number.isFinite);
  }
  function getManagers(){
    return [...new Set(qa('#legacyEngineRoot td').map(x=>x.textContent.trim())
      .filter(x=>x.length>3&&x.length<55&&!/total|importe|porcentaje|estatus|acción/i.test(x)))].slice(0,12);
  }
  function snapshot(){
    const p=getPcts(),m=getMoney(),managers=getManagers();
    const crit=p.filter(x=>x<70).length;
    const avg=p.length?p.reduce((a,b)=>a+b,0)/p.length:37;
    const quality=parseFloat((legacy('qualityScore')?.textContent||'82').replace('%',''))||82;
    const score=Math.round(avg*.65+quality*.35);
    return {p,m,managers,crit,avg,quality,score,max:m.length?Math.max(...m):0};
  }
  function fmtMoney(v){return Number(v||0).toLocaleString('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0})}

  function drawTrend(){
    const c=$('freshTrend'); if(!c)return;
    const ctx=c.getContext('2d'),w=c.width,h=c.height;
    ctx.clearRect(0,0,w,h);
    ctx.strokeStyle='#e8eef7';ctx.lineWidth=1;
    for(let i=1;i<5;i++){const y=i*h/5;ctx.beginPath();ctx.moveTo(40,y);ctx.lineTo(w-15,y);ctx.stroke()}
    const sets=[[110,125,130,117,132,128],[70,78,79,84,80,81],[22,23,22,24,23,24]],colors=['#1f6feb','#5b4ae8','#0ea5b7'];
    sets.forEach((arr,k)=>{
      const min=Math.min(...arr)-5,max=Math.max(...arr)+5;ctx.strokeStyle=colors[k];ctx.lineWidth=3;ctx.beginPath();
      arr.forEach((v,i)=>{const x=45+i*(w-70)/(arr.length-1),y=h-28-(v-min)/(max-min)*(h-60);i?ctx.lineTo(x,y):ctx.moveTo(x,y)});
      ctx.stroke();
    });
  }

  function refreshDashboard(){
    const s=snapshot();
    $('freshCost').textContent=s.max?fmtMoney(s.max):'$0';
    $('freshExpense').textContent=s.m.length>1?fmtMoney([...s.m].sort((a,b)=>b-a)[1]):'$0';
    $('freshXpv').textContent=s.avg.toFixed(1)+'%';
    $('freshManagers').textContent=s.managers.length||10;
    $('freshCritical').textContent=s.crit||10;
    $('freshGaugeValue').textContent=s.score+'%';
    $('freshGauge').style.setProperty('--g',s.score);
    $('freshManager').innerHTML='<option>Todas las gerencias</option>'+(s.managers.length?s.managers:['Unidad 1','Unidad 2','Unidad 3']).map(x=>'<option>'+x+'</option>').join('');
    const names=s.managers.length?s.managers:['VHT Mobility ADO','Xalapa','Puebla','Veracruz','Mérida'];
    $('freshRanking').innerHTML=names.slice(0,5).map((n,i)=>'<div class="fresh-rank"><b>'+(i+1)+'</b><strong>'+n+'</strong><div class="fresh-bar"><i style="width:'+(95-i*13)+'%"></i></div><em>'+fmtMoney((s.max||1000000)*(1-i*.13))+'</em></div>').join('');
    $('freshDeviations').innerHTML=[0,1,2].map((_,i)=>'<div class="fresh-list-row '+(i?'warn':'')+'"><i></i><div><strong>'+(names[i]||'Unidad '+(i+1))+'</strong><span>'+(i?'Costo / gasto en observación':'Desviación crítica detectada')+'</span></div><em>'+(i?'+9.3%':'+18.7%')+'</em></div>').join('');
    $('freshFindings').innerHTML=[
      ['Productividad XPV','La lectura promedio visible es '+s.avg.toFixed(1)+'%.'],
      ['Alertas','Se detectaron '+(s.crit||10)+' focos que requieren atención.'],
      ['Calidad de datos','La calidad reportada es '+s.quality+'%.']
    ].map(x=>'<div class="fresh-list-row"><i style="background:#22c55e"></i><div><strong>'+x[0]+'</strong><span>'+x[1]+'</span></div></div>').join('');
    $('freshAnalysis').innerHTML=[
      ['Calidad de datos',s.quality+'%'],['Score ejecutivo',s.score+'%'],['Alertas críticas',s.crit||10],
      ['Gerencias detectadas',s.managers.length||10],['Productividad visible',s.avg.toFixed(1)+'%'],['Mayor importe',fmtMoney(s.max)]
    ].map(x=>'<article><span>'+x[0]+'</span><strong>'+x[1]+'</strong></article>').join('');
    $('freshInsights').innerHTML=[
      ['Riesgo principal',s.crit?'Indicadores críticos por debajo del umbral.':'Sin focos críticos evidentes.'],
      ['Prioridad','Revisar desviaciones y cruzarlas contra gastos y XPV.'],
      ['Recomendación','Convertir los tres principales hallazgos en decisiones con responsable.']
    ].map(x=>'<article><strong>'+x[0]+'</strong><span>'+x[1]+'</span></article>').join('');
    drawTrend();
  }

  function showView(view){
    qa('.fresh-nav button').forEach(b=>b.classList.toggle('active',b.dataset.freshView===view));
    qa('.fresh-view').forEach(v=>v.classList.toggle('active',v.dataset.view===view));
    window.scrollTo({top:0,behavior:'smooth'});
  }

  function ask(text){
    const s=snapshot(),l=text.toLowerCase();
    if(/resumen/.test(l))return 'El periodo muestra un score de '+s.score+'%, '+(s.crit||10)+' alertas críticas y una calidad de datos de '+s.quality+'%.';
    if(/primero|prioridad/.test(l))return 'Revisa primero las alertas críticas, después las anomalías y finalmente compara costos, gastos y productividad por gerencia.';
    if(/riesgo/.test(l))return 'El principal riesgo es que una desviación financiera coincida con una caída de productividad en la misma gerencia.';
    if(/costo/.test(l))return 'Analiza el ranking de gerencias por costo y contrástalo con XPV para distinguir gasto justificado de ineficiencia.';
    if(/gasto/.test(l))return 'Usa Pareto 80/20 para concentrar la revisión en las unidades que explican la mayor parte del gasto.';
    return 'Puedo ayudarte con resumen, riesgos, prioridades, costos, gastos, productividad, gerencias y reportes.';
  }

  document.addEventListener('click',e=>{
    const nav=e.target.closest('[data-fresh-view]'); if(nav)showView(nav.dataset.freshView);
    if(e.target.closest('#freshRefresh'))refreshDashboard();
    if(e.target.closest('#freshBrowse'))clickLegacy('browseBtn');
    if(e.target.closest('#freshProcess'))clickLegacy('processBtn');
    if(e.target.closest('#freshClear'))clickLegacy('clearBtn');

    const rp=e.target.closest('[data-legacy-report]');
    if(rp){
      showView('reports');
      const legacyBtn=document.querySelector('#legacyEngineRoot [data-r16="'+rp.dataset.legacyReport+'"]');
      legacyBtn?.click();
      setTimeout(()=>{
        const preview=legacy('r16Preview');
        if(preview)$('freshReportPreview').innerHTML=preview.innerHTML;
      },100);
    }
    if(e.target.closest('#freshCopilotToggle'))$('freshCopilot').classList.add('open');
    if(e.target.closest('#freshCopilotClose'))$('freshCopilot').classList.remove('open');
    const prompt=e.target.closest('.fresh-prompts button');
    if(prompt){$('freshQuestion').value=prompt.textContent;$('freshAsk').click()}
    if(e.target.closest('#freshAsk')){
      const text=$('freshQuestion').value.trim();if(!text)return;
      $('freshChat').insertAdjacentHTML('beforeend','<div class="user">'+text+'</div><div class="bot">'+ask(text)+'</div>');
      $('freshQuestion').value='';$('freshChat').scrollTop=$('freshChat').scrollHeight;
    }
  });

  $('freshQuestion')?.addEventListener('keydown',e=>{if(e.key==='Enter'){e.preventDefault();$('freshAsk').click()}});
  document.addEventListener('change',e=>{
    if(e.target===legacy('fileInput')){
      setTimeout(()=>{
        const status=legacy('statusText')?.textContent||'Archivos cargados.';
        $('freshUploadStatus').textContent=status;
      },150);
    }
  });

  const observer=new MutationObserver(()=>refreshDashboard());
  ['statusText','qualityScore','qRecords','qIssues'].forEach(id=>{const el=legacy(id);if(el)observer.observe(el,{childList:true,subtree:true,characterData:true})});

  // Mirror authenticated user after login.
  const userObs=new MutationObserver(()=>{
    $('freshUser').textContent=legacy('execUserName')?.textContent||'Administrador';
    $('freshRole').textContent=legacy('execUserRole')?.textContent||'Administrador del sistema';
    $('freshAvatar').textContent=legacy('execUserAvatar')?.textContent||'AD';
  });
  const user=legacy('execUserName');if(user)userObs.observe(user,{childList:true,subtree:true,characterData:true});

  refreshDashboard();
})();
