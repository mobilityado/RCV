
/* REPORT.IA RCV Intelligence Executive 9.0 - non-destructive enhancement layer */
(function(){
  'use strict';

  const V9 = {
    text(){ return document.body.innerText || ''; },
    nums(){
      const matches = this.text().match(/-?\$?\s?\d[\d,]*(?:\.\d+)?%?/g) || [];
      return matches.map(x => ({
        raw:x.trim(),
        value:Number(x.replace(/[$,%\s]/g,'').replace(/,/g,'')) || 0,
        pct:x.includes('%'),
        money:x.includes('$')
      })).filter(x=>Number.isFinite(x.value));
    },
    headings(){
      return [...document.querySelectorAll('h1,h2,h3,h4,.card-title,.kpi-title,.section-title')]
        .map(e=>e.textContent.trim()).filter(Boolean).slice(0,30);
    },
    visibleCards(){
      return [...document.querySelectorAll('.card,.kpi-card,.metric-card,.stat-card,[class*="kpi"],[class*="card"]')]
        .filter(e=>e.offsetParent!==null).slice(0,20);
    }
  };

  function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}

  function buildUI(){
    if(document.getElementById('rcvV9Toolbar')) return;
    document.body.insertAdjacentHTML('beforeend', `
      <div class="rcv-v9-toolbar" id="rcvV9Toolbar">
        <button class="rcv-v9-btn" data-v9="alerts">🔔 Alertas</button>
        <button class="rcv-v9-btn" data-v9="compare">⚖️ Comparador</button>
        <button class="rcv-v9-btn" data-v9="copilot">✦ Copiloto RCV</button>
        <button class="rcv-v9-btn" data-v9="present">▣ Presentar</button>
      </div>
      <aside class="rcv-v9-panel" id="rcvV9Panel">
        <button class="rcv-v9-close" data-v9="close">✕</button>
        <div id="rcvV9PanelBody"></div>
      </aside>
      <div class="rcv-v9-modal" id="rcvV9Modal">
        <div class="rcv-v9-presentation">
          <button class="rcv-v9-btn" data-v9="closepresent" style="float:right">Salir</button>
          <div class="rcv-v9-slide" id="rcvV9Slide"></div>
        </div>
      </div>
    `);
  }

  function openPanel(html){
    document.getElementById('rcvV9PanelBody').innerHTML=html;
    document.getElementById('rcvV9Panel').classList.add('open');
  }

  function smartSummary(){
    const n=V9.nums();
    const money=n.filter(x=>x.money).sort((a,b)=>b.value-a.value);
    const pct=n.filter(x=>x.pct);
    const headings=V9.headings();
    let lines=[
      'Resumen ejecutivo automático',
      '',
      `Se detectaron ${headings.length} bloques de análisis visibles y ${n.length} indicadores numéricos en la vista actual.`
    ];
    if(money[0]) lines.push(`El importe de mayor magnitud visible es ${money[0].raw}.`);
    if(pct.length){
      const best=[...pct].sort((a,b)=>b.value-a.value)[0];
      const low=[...pct].sort((a,b)=>a.value-b.value)[0];
      lines.push(`El porcentaje más alto visible es ${best.raw}; el menor es ${low.raw}.`);
    }
    lines.push('Prioridad sugerida: revisar variaciones atípicas, gerencias con deterioro y rubros con crecimiento sostenido antes de cerrar el periodo.');
    return lines.join('\n');
  }

  function alerts(){
    const n=V9.nums(), pct=n.filter(x=>x.pct);
    let cards=[];
    pct.forEach(x=>{
      if(x.value<70) cards.push(`<div class="rcv-v9-card rcv-v9-danger"><strong>Indicador crítico: ${esc(x.raw)}</strong>Requiere revisión prioritaria por encontrarse debajo del umbral ejecutivo de 70%.</div>`);
      else if(x.value<90) cards.push(`<div class="rcv-v9-card rcv-v9-alert"><strong>Indicador en observación: ${esc(x.raw)}</strong>Conviene revisar su evolución frente al periodo anterior.</div>`);
    });
    if(!cards.length) cards.push(`<div class="rcv-v9-card rcv-v9-good"><strong>Sin alertas críticas visibles</strong>No se detectaron porcentajes visibles debajo de los umbrales configurados en la vista actual.</div>`);
    openPanel(`<h2>Centro de Alertas</h2><p>Análisis automático de la información actualmente visible.</p>${cards.slice(0,12).join('')}`);
  }

  function compare(){
    const hs=V9.headings();
    openPanel(`<h2>Comparador avanzado</h2>
      <p>Selecciona dos bloques visibles para analizarlos lado a lado. Esta herramienta no modifica los cálculos originales.</p>
      <label>Elemento A</label><select class="rcv-v9-select" id="v9a">${hs.map(x=>`<option>${esc(x)}</option>`).join('')}</select>
      <label>Elemento B</label><select class="rcv-v9-select" id="v9b">${hs.map(x=>`<option>${esc(x)}</option>`).join('')}</select>
      <button class="rcv-v9-btn" id="v9CompareRun">Comparar</button>
      <div id="v9CompareResult"></div>`);
    setTimeout(()=>{
      const b=document.getElementById('v9CompareRun');
      if(b)b.onclick=()=>{
        const a=document.getElementById('v9a').value, c=document.getElementById('v9b').value;
        document.getElementById('v9CompareResult').innerHTML=
          `<div class="rcv-v9-grid"><div class="rcv-v9-card"><strong>${esc(a)}</strong>Utiliza los filtros originales del portal para aislar este elemento y contrastar sus KPIs.</div><div class="rcv-v9-card"><strong>${esc(c)}</strong>Compara variación, tendencia, costos, gastos y productividad contra el elemento A.</div></div>`;
      };
    },0);
  }

  function copilot(){
    openPanel(`<h2>Copiloto RCV</h2><p>Asistente de análisis sobre el contenido cargado en REPORT.IA RCV.</p>
      <div class="rcv-v9-card"><div class="rcv-v9-answer">${esc(smartSummary())}</div></div>
      <input class="rcv-v9-input" id="v9q" placeholder="Ej. ¿Qué debo revisar primero?">
      <button class="rcv-v9-btn" id="v9Ask">Analizar</button>
      <div class="rcv-v9-card rcv-v9-answer" id="v9Ans">Puedes preguntar por gerencias, costos, gastos, productividad, tendencias o desviaciones.</div>`);
    setTimeout(()=>{
      const ask=document.getElementById('v9Ask');
      if(ask) ask.onclick=()=>{
        const q=(document.getElementById('v9q').value||'').toLowerCase();
        let ans=smartSummary();
        if(q.includes('primero')||q.includes('prioridad')) ans='Prioriza los indicadores con deterioro porcentual, después los rubros de mayor impacto económico y finalmente las gerencias con caída simultánea de productividad y aumento de gasto.';
        else if(q.includes('gerencia')) ans='Usa el módulo Gerencias para identificar extremos del ranking y después abre el detalle de la gerencia seleccionada. Conviene contrastar costo, gasto y productividad en el mismo periodo.';
        else if(q.includes('gasto')||q.includes('costo')) ans='Revisa primero los importes de mayor magnitud y las variaciones contra el periodo anterior. Un incremento de gasto acompañado de menor productividad debe tratarse como desviación prioritaria.';
        else if(q.includes('productividad')) ans='Compara XPV entre periodos y gerencias. Las caídas sostenidas deben analizarse junto con costos y gastos para distinguir un efecto temporal de una pérdida estructural de eficiencia.';
        document.getElementById('v9Ans').textContent=ans;
      };
    },0);
  }

  function present(){
    const modal=document.getElementById('rcvV9Modal'), slide=document.getElementById('rcvV9Slide');
    const cards=V9.visibleCards();
    const snippets=cards.slice(0,6).map(c=>`<div class="rcv-v9-mini">${esc(c.innerText.slice(0,220))}</div>`).join('');
    slide.innerHTML=`<h1>REPORT.IA RCV</h1><p>Intelligence Executive 9.0 · Vista de presentación</p>
      <div class="rcv-v9-grid">${snippets || `<div class="rcv-v9-mini">${esc(smartSummary())}</div>`}</div>
      <div class="rcv-v9-card" style="margin-top:20px"><strong>Resumen Inteligente</strong><div class="rcv-v9-answer">${esc(smartSummary())}</div></div>`;
    modal.classList.add('open');
  }

  document.addEventListener('click',e=>{
    const a=e.target.closest('[data-v9]');
    if(!a)return;
    const k=a.dataset.v9;
    if(k==='alerts')alerts();
    if(k==='compare')compare();
    if(k==='copilot')copilot();
    if(k==='present')present();
    if(k==='close')document.getElementById('rcvV9Panel').classList.remove('open');
    if(k==='closepresent')document.getElementById('rcvV9Modal').classList.remove('open');
  });

  function init(){buildUI();}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
