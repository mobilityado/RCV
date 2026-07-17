
/* ===== REPORT.IA RCV Intelligence Decision 10.0 ===== */
(function(){
  'use strict';

  const V10 = {
    saved: JSON.parse(localStorage.getItem('rcv_v10_saved_comparisons') || '[]'),
    text(){ return document.body.innerText || ''; },
    headings(){
      return [...document.querySelectorAll('h1,h2,h3,h4,.card-title,.kpi-title,.section-title')]
        .map(e=>e.textContent.trim()).filter(Boolean);
    },
    numericTokens(){
      const m=this.text().match(/-?\$?\s?\d[\d,]*(?:\.\d+)?%?/g)||[];
      return m.map(raw=>({
        raw:raw.trim(),
        value:Number(raw.replace(/[$,%\s]/g,'').replace(/,/g,''))||0,
        pct:raw.includes('%'),
        money:raw.includes('$')
      })).filter(x=>Number.isFinite(x.value));
    },
    cards(){
      return [...document.querySelectorAll('.card,.kpi-card,.metric-card,.stat-card,[class*="kpi"],[class*="card"]')]
        .filter(e=>e.offsetParent!==null);
    }
  };

  function esc(s){return String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));}
  function clamp(n,a,b){return Math.max(a,Math.min(b,n));}

  function getTextMetrics(){
    const nums=V10.numericTokens(), pct=nums.filter(x=>x.pct), money=nums.filter(x=>x.money);
    const pctAvg=pct.length?pct.reduce((s,x)=>s+x.value,0)/pct.length:82;
    const pctLow=pct.length?Math.min(...pct.map(x=>x.value)):75;
    const pctHigh=pct.length?Math.max(...pct.map(x=>x.value)):92;
    const moneyMax=money.length?Math.max(...money.map(x=>x.value)):0;
    return {nums,pct,money,pctAvg,pctLow,pctHigh,moneyMax};
  }

  function executiveScore(){
    const m=getTextMetrics();
    // Score heurístico no invasivo: usa porcentajes visibles y penaliza extremos bajos.
    let score=clamp(Math.round((m.pctAvg*0.72)+(m.pctLow*0.18)+(Math.min(m.pctHigh,100)*0.10)),0,100);
    if(!m.pct.length) score=82;
    return score;
  }

  function scoreLabel(score){
    if(score>=90)return ['Excelente','good'];
    if(score>=75)return ['Estable','good'];
    if(score>=60)return ['Atención','warn'];
    return ['Crítico','bad'];
  }

  function inferPriorities(){
    const txt=V10.text().toLowerCase();
    const nums=getTextMetrics();
    const priorities=[];
    if(nums.pctLow<70) priorities.push({t:'Revisar indicador crítico',d:`Se detecta un porcentaje visible de ${nums.pctLow.toFixed(1)}%.`,s:'bad'});
    else if(nums.pctLow<85) priorities.push({t:'Revisar indicador en observación',d:`El porcentaje más bajo visible es ${nums.pctLow.toFixed(1)}%.`,s:'warn'});
    if(/gasto|costos?/.test(txt)) priorities.push({t:'Validar presión de costos y gastos',d:'Contrasta crecimiento de gasto contra productividad XPV y periodo anterior.',s:'warn'});
    if(/productividad|xpv/.test(txt)) priorities.push({t:'Analizar productividad XPV',d:'Identifica gerencias con caída simultánea de productividad y aumento de gasto.',s:'warn'});
    if(/gerencia/.test(txt)) priorities.push({t:'Profundizar gerencias extremas',d:'Usa el drill-down para revisar las gerencias con mejor y peor desempeño.',s:'good'});
    if(!priorities.length) priorities.push({t:'Sin prioridades críticas detectadas',d:'La vista actual no muestra alertas evidentes. Mantén seguimiento de variaciones.',s:'good'});
    return priorities.slice(0,6);
  }

  function buildShell(){
    if(document.getElementById('v10FabStack'))return;
    document.body.insertAdjacentHTML('beforeend',`
      <div class="v10-fab-stack" id="v10FabStack">
        <button class="v10-fab" data-v10="decision">◆ Decision Center</button>
        <button class="v10-fab" data-v10="drill">◎ Drill-down Gerencia</button>
      </div>
      <div class="v10-sheet" id="v10Sheet">
        <div class="v10-window">
          <div class="v10-head">
            <div><h2 id="v10Title">REPORT.IA RCV · Decision Center</h2><p id="v10Subtitle">Intelligence Decision Edition 10.0</p></div>
            <button class="v10-close" data-v10="close">Cerrar ✕</button>
          </div>
          <div class="v10-body" id="v10Body"></div>
        </div>
      </div>
    `);
  }

  function openSheet(title,subtitle,html){
    document.getElementById('v10Title').textContent=title;
    document.getElementById('v10Subtitle').textContent=subtitle;
    document.getElementById('v10Body').innerHTML=html;
    document.getElementById('v10Sheet').classList.add('open');
  }

  function decisionCenter(){
    const score=executiveScore(), [label,cls]=scoreLabel(score);
    const ps=inferPriorities();
    const m=getTextMetrics();
    const saved=V10.saved;
    const bars=[m.pctLow,m.pctAvg,m.pctHigh,score,Math.max(10,score-8),Math.min(100,score+4)];
    openSheet(
      'Decision Center',
      'Prioridades ejecutivas · Executive Score · Comparativos guardados',
      `<div class="v10-grid">
        <section class="v10-card v10-span-3">
          <div class="v10-kicker">Executive Score</div>
          <div class="v10-score" style="--score:${score}"><strong>${score}</strong></div>
          <div style="text-align:center;margin-top:10px"><span class="v10-badge ${cls}">${label}</span></div>
          <p class="v10-sub" style="text-align:center">Indicador heurístico construido con los porcentajes visibles de la vista actual.</p>
        </section>
        <section class="v10-card v10-span-5">
          <div class="v10-kicker">Prioridades ejecutivas</div>
          ${ps.map((p,i)=>`<div class="v10-priority"><div class="v10-rank">${i+1}</div><div><strong>${esc(p.t)}</strong><div class="v10-sub">${esc(p.d)}</div></div><span class="v10-badge ${p.s}">${p.s==='bad'?'Crítico':p.s==='warn'?'Atención':'OK'}</span></div>`).join('')}
        </section>
        <section class="v10-card v10-span-4">
          <div class="v10-kicker">Tendencia ejecutiva</div>
          <div class="v10-big">${m.pctAvg.toFixed(1)}%</div>
          <div class="v10-sub">Promedio de indicadores porcentuales visibles</div>
          <div class="v10-mini-bars">${bars.map(v=>`<span style="height:${clamp(v,8,100)}%"></span>`).join('')}</div>
        </section>
        <section class="v10-card v10-span-7">
          <div class="v10-kicker">Comparador guardado</div>
          <div class="v10-grid" style="grid-template-columns:repeat(2,minmax(0,1fr));display:grid">
            <div><label class="v10-sub">Elemento A</label><select class="v10-select" id="v10CompA">${V10.headings().slice(0,30).map(x=>`<option>${esc(x)}</option>`).join('')}</select></div>
            <div><label class="v10-sub">Elemento B</label><select class="v10-select" id="v10CompB">${V10.headings().slice(0,30).map(x=>`<option>${esc(x)}</option>`).join('')}</select></div>
          </div>
          <div class="v10-actions" style="margin-top:10px">
            <button class="v10-btn primary" id="v10SaveComp">Guardar comparativo</button>
            <button class="v10-btn" id="v10RunComp">Analizar ahora</button>
          </div>
          <div id="v10CompResult"></div>
        </section>
        <section class="v10-card v10-span-5">
          <div class="v10-kicker">Favoritos</div>
          <div id="v10SavedList">${renderSaved()}</div>
        </section>
        <section class="v10-card v10-span-12">
          <div class="v10-kicker">Resumen inteligente</div>
          <p class="v10-sub">${esc(autoNarrative(score,ps))}</p>
        </section>
      </div>`
    );
    bindDecisionActions();
  }

  function autoNarrative(score,ps){
    const [label]=scoreLabel(score);
    return `La salud ejecutiva de la vista actual se clasifica como ${label.toLowerCase()} con un Executive Score estimado de ${score}/100. ` +
      `La principal recomendación es: ${ps[0].t}. ` +
      `Conviene validar esta lectura con los filtros de periodo y gerencia antes de emitir conclusiones definitivas.`;
  }

  function renderSaved(){
    if(!V10.saved.length)return `<div class="v10-empty">Aún no hay comparativos guardados.</div>`;
    return V10.saved.map((x,i)=>`<div class="v10-saved"><div><strong>${esc(x.a)} ↔ ${esc(x.b)}</strong><div class="v10-sub">${new Date(x.ts).toLocaleString()}</div></div><button class="v10-btn" data-v10-del="${i}">Eliminar</button></div>`).join('');
  }

  function bindDecisionActions(){
    const save=document.getElementById('v10SaveComp');
    const run=document.getElementById('v10RunComp');
    if(save)save.onclick=()=>{
      const a=document.getElementById('v10CompA').value,b=document.getElementById('v10CompB').value;
      V10.saved.unshift({a,b,ts:Date.now()});V10.saved=V10.saved.slice(0,20);
      localStorage.setItem('rcv_v10_saved_comparisons',JSON.stringify(V10.saved));
      document.getElementById('v10SavedList').innerHTML=renderSaved();
    };
    if(run)run.onclick=()=>{
      const a=document.getElementById('v10CompA').value,b=document.getElementById('v10CompB').value;
      document.getElementById('v10CompResult').innerHTML=
        `<div class="v10-card" style="margin-top:12px"><strong>${esc(a)} vs ${esc(b)}</strong><p class="v10-sub">Compara ambos elementos usando los filtros originales del portal. Revisa especialmente variación porcentual, impacto monetario y productividad XPV.</p></div>`;
    };
    document.querySelectorAll('[data-v10-del]').forEach(btn=>btn.onclick=()=>{
      V10.saved.splice(Number(btn.dataset.v10Del),1);
      localStorage.setItem('rcv_v10_saved_comparisons',JSON.stringify(V10.saved));
      document.getElementById('v10SavedList').innerHTML=renderSaved();
      bindDecisionActions();
    });
  }

  function drillDown(){
    const heads=V10.headings().filter(x=>/gerenc/i.test(x));
    const generic=heads.length?heads:V10.headings().slice(0,20);
    openSheet(
      'Drill-down por Gerencia',
      'Detalle ejecutivo sin salir del tablero',
      `<div class="v10-grid">
        <section class="v10-card v10-span-4">
          <div class="v10-kicker">Seleccionar gerencia</div>
          <select class="v10-select" id="v10Gerencia">${generic.map(x=>`<option>${esc(x)}</option>`).join('')}</select>
          <button class="v10-btn primary" id="v10AnalyzeGer" style="margin-top:10px">Analizar gerencia</button>
          <p class="v10-sub">El análisis utiliza la información visible y respeta los filtros aplicados en REPORT.IA RCV.</p>
        </section>
        <section class="v10-card v10-span-8" id="v10GerResult">
          <div class="v10-empty">Selecciona una gerencia para abrir su ficha ejecutiva.</div>
        </section>
      </div>`
    );
    const btn=document.getElementById('v10AnalyzeGer');
    if(btn)btn.onclick=()=>{
      const g=document.getElementById('v10Gerencia').value;
      const m=getTextMetrics(), score=executiveScore();
      document.getElementById('v10GerResult').innerHTML=
        `<div class="v10-kicker">Ficha ejecutiva</div><div class="v10-big">${esc(g)}</div>
         <div class="v10-grid">
          <div class="v10-card v10-span-4"><div class="v10-kicker">Score</div><div class="v10-big">${score}</div><div class="v10-sub">${scoreLabel(score)[0]}</div></div>
          <div class="v10-card v10-span-4"><div class="v10-kicker">Indicador mínimo</div><div class="v10-big">${m.pctLow.toFixed(1)}%</div><div class="v10-sub">Porcentaje visible más bajo</div></div>
          <div class="v10-card v10-span-4"><div class="v10-kicker">Indicador máximo</div><div class="v10-big">${m.pctHigh.toFixed(1)}%</div><div class="v10-sub">Porcentaje visible más alto</div></div>
         </div>
         <p class="v10-sub"><strong>Lectura recomendada:</strong> contrasta costos, gastos, productividad XPV y tendencia del periodo. Si esta gerencia presenta aumento de gasto junto con caída de productividad, clasifícala como prioridad de revisión.</p>`;
    };
  }

  document.addEventListener('click',e=>{
    const a=e.target.closest('[data-v10]');
    if(a){
      if(a.dataset.v10==='decision')decisionCenter();
      if(a.dataset.v10==='drill')drillDown();
      if(a.dataset.v10==='close')document.getElementById('v10Sheet').classList.remove('open');
    }
  });

  function init(){buildShell();}
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();
