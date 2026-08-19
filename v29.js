(()=>{
  const $v=id=>document.getElementById(id);
  const monthNames=["Enero","Febrero","Marzo","Abril","Mayo","Junio","Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"];
  function safeRows(){try{return typeof managerHealth==="function"?managerHealth():[]}catch{return[]}}
  function safeAudit(){try{return typeof auditMetrics==="function"?auditMetrics():null}catch{return null}}
  function sourceLabel(){try{return state?.activeDataSource==="cloud"?"Nube oficial":"Tiempo real"}catch{return"REPORT.IA"}}
  function refreshCommand(){
    const rows=safeRows(), audit=safeAudit(), top=rows[0];
    const critical=rows.filter(x=>x.status==="Crítico").length, attention=rows.filter(x=>x.status==="Atención").length;
    const overall=!rows.length?"SIN DATOS":critical?"REQUIERE ATENCIÓN":attention?"BAJO VIGILANCIA":"OPERACIÓN ESTABLE";
    const headline=!rows.length?"La operación está lista para convertirse en decisiones.":critical?`${critical} gerencia${critical>1?"s":""} requiere${critical>1?"n":""} atención prioritaria.`:attention?`La operación es estable, con ${attention} señal${attention>1?"es":""} por vigilar.`:"Los indicadores principales se mantienen en zona favorable.";
    const narrative=!rows.length?"Procesa tus archivos para obtener prioridades, riesgos y señales ejecutivas en un solo vistazo.":top?`El mayor foco del periodo es ${top.manager}, con presión ejecutiva de ${(top.pressure*100).toFixed(1)}%. Usa el semáforo para documentar seguimiento.`:"Sin focos relevantes.";
    if($v("v29OverallStatus"))$v("v29OverallStatus").textContent=overall;
    if($v("v29OverallHint"))$v("v29OverallHint").textContent=critical?"Hay focos rojos activos":attention?"Revisar señales amarillas":"Sin presión relevante";
    if($v("v29TopRisk"))$v("v29TopRisk").textContent=top?.manager||"—";
    if($v("v29TopRiskHint"))$v("v29TopRiskHint").textContent=top?`${top.status} · ${(top.pressure*100).toFixed(1)}% presión`:"Sin gerencia prioritaria";
    if($v("v29Quality"))$v("v29Quality").textContent=audit?`${audit.confidence}%`:"—";
    let periodText="PERIODO";
    try{
      const realtimeActive=!!document.querySelector('[data-v27-tab="realtime"].active');
      if(realtimeActive){
        // Fuente principal: el nombre de mes ya calculado dentro del modelo local.
        // Así la tarjeta no depende del selector de nube ni de un valor inicial.
        const localModel=window.REPORTIA_APP?.getLocalModel?.();
        const modelMonth=String(localModel?.month||"").trim();
        if(modelMonth){
          periodText=modelMonth.toUpperCase();
        }else{
          const lp=window.REPORTIA_APP?.getLocalPeriod?.();
          if(lp!==null && lp!==undefined && Number.isFinite(Number(lp)))
            periodText=(monthNames[Number(lp)]||"Periodo").toUpperCase();
        }
      }else{
        const cp=window.REPORTIA_APP?.getCloudPeriod?.();
        const fallback=(typeof state!=="undefined"?state?.periodIndex:null);
        const p=(cp!==null&&cp!==undefined)?Number(cp):Number(fallback);
        if(Number.isFinite(p)) periodText=(monthNames[p]||"Periodo").toUpperCase();
      }
    }catch{}
    if($v("v29Period"))$v("v29Period").textContent=periodText;
    if($v("v29Source"))$v("v29Source").textContent=sourceLabel();
    if($v("v29Headline"))$v("v29Headline").textContent=headline;
    if($v("v29Narrative"))$v("v29Narrative").textContent=narrative;
    [$v("v29CommandCenter")].forEach(el=>{if(el){el.classList.remove("v29-pulse");void el.offsetWidth;el.classList.add("v29-pulse")}});
    refreshCloudFocus();refreshActivity();
  }
  function refreshCloudFocus(){
    const rows=safeRows(),top=rows[0], cloud=(window.REPORTIA_APP?.getCloud?.()||{});
    if($v("v29CloudFocus"))$v("v29CloudFocus").textContent=top?top.manager:(cloud.snapshotId?"Publicación disponible":"Sin publicación seleccionada");
    if($v("v29CloudFocusText"))$v("v29CloudFocusText").textContent=top?`${top.status}. Presión ejecutiva ${(top.pressure*100).toFixed(1)}%. Revisa la incidencia y el seguimiento del área.`:(cloud.snapshotId?`Publicada por ${cloud.uploadedBy||"Administrador"}. Abre el semáforo para revisar incidencias.`:"Selecciona un mes para revisar su principal señal ejecutiva.");
  }
  function refreshActivity(){
    const box=$v("v29ActivityFeed");if(!box)return;
    const cloud=(window.REPORTIA_APP?.getCloud?.()||{}),rows=safeRows();const events=[];
    if(cloud.snapshotId)events.push({ico:"☁",title:`Publicación oficial disponible`,desc:`${cloud.uploadedBy||"Administrador"} publicó el periodo seleccionado`,time:cloud.uploadedAt?new Date(cloud.uploadedAt).toLocaleDateString("es-MX"):"vigente"});
    rows.filter(r=>r.status!=="Favorable").slice(0,3).forEach(r=>events.push({ico:r.status==="Crítico"?"🔴":"🟡",title:`${r.manager} · ${r.status}`,desc:`Presión ${(r.pressure*100).toFixed(1)}% · requiere seguimiento`,time:"Ahora"}));
    if(!events.length){box.innerHTML='<div class="empty">La actividad aparecerá cuando existan publicaciones o incidencias.</div>';return}
    box.innerHTML=events.map(e=>`<div class="v29-event"><i>${e.ico}</i><div><strong>${e.title}</strong><span>${e.desc}</span></div><time>${e.time}</time></div>`).join("");
  }
  function toggleBoard(){document.body.classList.toggle("v29-board-mode");const b=$v("v29BoardMode");if(b)b.textContent=document.body.classList.contains("v29-board-mode")?"✕ Salir modo junta":"▣ Modo junta"}
  document.addEventListener("click",e=>{if(e.target.closest("#v29BoardMode"))toggleBoard();if(e.target.closest("#v29RefreshActivity"))refreshCommand();});
  document.addEventListener("keydown",e=>{if(e.key==="Escape"&&document.body.classList.contains("v29-board-mode"))toggleBoard();});
  ["reportia:model-processed","reportia:cloud-published","reportia:cloud-loaded","reportia:source-changed"].forEach(n=>window.addEventListener(n,()=>setTimeout(refreshCommand,100)));
  const oldRender=window.renderAll||null;
  if(typeof renderAll==="function"){
    const base=renderAll;
    window.renderAll=function(){const r=base.apply(this,arguments);setTimeout(refreshCommand,0);return r};
  }
  setTimeout(refreshCommand,350);
})();
