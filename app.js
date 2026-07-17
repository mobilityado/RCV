/* REPORT.IA RCV - procesamiento 100% en navegador */
const state = { files: [], sources: {}, outputs: {}, charts: [] };
const MONTHS = [
  ['EneM','Ene'],['Feb','Feb'],['Mar','Mar'],['Abr','Abr'],['May','May'],['Jun','Jun'],
  ['Jul','Jul'],['Ago','Ago'],['Sep','Sep'],['Oct','Oct'],['Nov','Nov'],['DicM','Dic']
];
const REQUIRED = [
  {key:'catalog',label:'CATÁLOGO DE GERENTES',aliases:['CATALOGO DE GERENTES','CATÁLOGO DE GERENTES']},
  {key:'cost',label:'JD COSTO RCV',aliases:['JD COSTO RCV']},
  {key:'xpv',label:'JD XPV RCV',aliases:['JD XPV RCV']},
  {key:'expense',label:'JD GASTOS RCV',aliases:['JD GASTOS RCV']}
];

const $ = s => document.querySelector(s);
const monthSelect = $('#month');
MONTHS.forEach(([value,label],i)=>{const o=document.createElement('option');o.value=value;o.textContent=label;if(i===5)o.selected=true;monthSelect.appendChild(o)});

$('#browseBtn').onclick = e => { e.preventDefault(); $('#fileInput').click(); };
$('#fileInput').onchange = e => loadFiles([...e.target.files]);
$('#clearBtn').onclick = reset;
$('#processBtn').onclick = processAll;
$('#zipBtn').onclick = downloadZip;

const drop = $('#dropZone');
['dragenter','dragover'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.add('drag')}));
['dragleave','drop'].forEach(ev=>drop.addEventListener(ev,e=>{e.preventDefault();drop.classList.remove('drag')}));
drop.addEventListener('drop',e=>loadFiles([...e.dataTransfer.files]));
drop.addEventListener('keydown',e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();$('#fileInput').click()}});

async function loadFiles(files){
  const accepted=files.filter(f=>/\.(xlsx|xls|xlsm)$/i.test(f.name));
  for(const file of accepted){
    if(state.files.some(x=>x.name===file.name&&x.size===file.size)) continue;
    state.files.push(file);
  }
  renderFiles();
  await parseSources();
}

function removeFile(i){state.files.splice(i,1);renderFiles();parseSources()}
function reset(){state.files=[];state.sources={};state.outputs={};state.analysis=null;destroyCharts();$('#fileInput').value='';renderFiles();renderValidation();['dashboard','gerencias','costos','gastos','xpv','tendencias','inteligencia','resultados'].forEach(id=>document.getElementById(id)?.classList.add('hidden'))}
function renderFiles(){
  $('#fileList').innerHTML=state.files.map((f,i)=>`<div class="file-item"><div class="file-meta"><div class="file-badge">XLS</div><div><strong>${escapeHtml(f.name)}</strong><small>${formatBytes(f.size)}</small></div></div><button class="remove" data-i="${i}" title="Quitar">×</button></div>`).join('');
  document.querySelectorAll('.remove').forEach(b=>b.onclick=()=>removeFile(+b.dataset.i));
}

async function parseSources(){
  state.sources={};
  $('#statusText').textContent='Leyendo y detectando pestañas…';
  try{
    for(const file of state.files){
      const data=await file.arrayBuffer();
      const wb=XLSX.read(data,{type:'array',cellDates:true,cellNF:true});
      for(const sheetName of wb.SheetNames){
        const norm=normalize(sheetName);
        for(const def of REQUIRED){
          if(def.aliases.some(a=>normalize(a)===norm)) state.sources[def.key]={sheet:wb.Sheets[sheetName],sheetName,fileName:file.name};
        }
      }
      // Archivos separados: detectar por nombre si tienen una sola hoja o nombres distintos.
      const fn=normalize(file.name);
      if(!state.sources.cost && fn.includes('COSTO')) state.sources.cost={sheet:wb.Sheets[wb.SheetNames[0]],sheetName:wb.SheetNames[0],fileName:file.name};
      if(!state.sources.xpv && (fn.includes('XPV')||fn.includes('PRODUCTIVIDAD'))) state.sources.xpv={sheet:wb.Sheets[wb.SheetNames[0]],sheetName:wb.SheetNames[0],fileName:file.name};
      if(!state.sources.expense && fn.includes('GAST')) state.sources.expense={sheet:wb.Sheets[wb.SheetNames[0]],sheetName:wb.SheetNames[0],fileName:file.name};
      if(!state.sources.catalog && fn.includes('GERENT')) state.sources.catalog={sheet:wb.Sheets[wb.SheetNames[0]],sheetName:wb.SheetNames[0],fileName:file.name};
    }
  }catch(err){console.error(err);alert('No fue posible leer uno de los archivos. Verifica que sea un Excel válido.');}
  renderValidation();
}

function rowCount(sheet){
  if(!sheet||!sheet['!ref']) return 0;
  const r=XLSX.utils.decode_range(sheet['!ref']);return Math.max(0,r.e.r-r.s.r);
}
function renderValidation(){
  const html=REQUIRED.map(d=>{
    const s=state.sources[d.key]; const count=s?rowCount(s.sheet):0;
    return `<div class="validation-card ${s?'ok':'bad'}"><div class="top"><h4>${d.label}</h4><span class="status-dot">${s?'✓':'!'}</span></div><p>${s?`${count.toLocaleString('es-MX')} registros aprox. · ${escapeHtml(s.fileName)}`:'No detectado. Carga la pestaña o archivo correspondiente.'}</p></div>`;
  }).join('');
  $('#validationGrid').innerHTML=html||'<div class="empty-state">Aún no has cargado información.</div>';
  const core=state.sources.cost&&state.sources.xpv&&state.sources.expense;
  $('#processBtn').disabled=!core;
  $('#statusText').textContent=core?(state.sources.catalog?'Todo listo para generar los reportes.':'Listo para procesar. No se detectó catálogo; se conservarán nombres originales de gerencia.'):'Carga al menos JD COSTO RCV, JD XPV RCV y JD GASTOS RCV.';
}

function sheetRows(sheet){return XLSX.utils.sheet_to_json(sheet,{header:1,raw:true,defval:''});}
function num(v){if(typeof v==='number')return Number.isFinite(v)?v:0;if(v==null||v==='')return 0;const n=Number(String(v).replace(/[$,%\s,]/g,''));return Number.isFinite(n)?n:0}
function normalize(v){return String(v??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toUpperCase().trim()}
function monthRank(v){const n=normalize(v);const i=MONTHS.findIndex(([a,b])=>normalize(a)===n||normalize(b)===n);return i<0?99:i}
function inCutoff(period,cutoff){const r=monthRank(period),c=monthRank(cutoff);return r<=c}
function getCatalog(){
  const map=new Map();if(!state.sources.catalog)return map;
  sheetRows(state.sources.catalog.sheet).forEach(r=>{if(r[0]&&r[1])map.set(normalize(r[0]),String(r[1]).trim())});return map;
}
function getManager(raw,map){return map.get(normalize(raw))||String(raw||'SIN CLASIFICAR').trim()||'SIN CLASIFICAR'}

function parseJD(sheet,kind,catalog,cutoff){
  const rows=sheetRows(sheet); const out=[];
  if(kind==='xpv'){
    for(let i=9;i<rows.length;i++){
      const r=rows[i];if(!r||!r[1])continue;
      out.push({account:r[1],center:r[2],managerRaw:r[3],manager:getManager(r[3],catalog),brand:r[4],real:num(r[5]),budget:num(r[6])});
    }
  } else {
    for(let i=8;i<rows.length;i++){
      const r=rows[i]; if(!r||!r[0])continue;
      if(!inCutoff(r[6],cutoff))continue;
      out.push({account:r[0],group:r[1],center:r[2],managerRaw:r[3],manager:getManager(r[3],catalog),business:r[4],brand:r[5],period:r[6],real25:num(r[7]),budget25:num(r[8]),real26:num(r[9]),budget26:num(r[10])});
    }
  }
  return out;
}
function aggregate(rows,keys,metrics){
  const m=new Map();
  rows.forEach(r=>{const k=keys.map(x=>String(r[x]??'')).join('\u001f');if(!m.has(k)){const o={};keys.forEach(x=>o[x]=r[x]??'');metrics.forEach(x=>o[x]=0);m.set(k,o)}const o=m.get(k);metrics.forEach(x=>o[x]+=num(r[x]))});
  return [...m.values()];
}
function sortBy(a,keys){return a.sort((x,y)=>keys.map(k=>String(x[k]).localeCompare(String(y[k]),'es')).find(v=>v!==0)||0)}

function reportSheet(title,rows,columns,period){
  const aoa=[];
  aoa.push(['REPORT.IA RCV']);aoa.push([title]);aoa.push(['Periodo',MONTHS.find(x=>x[0]===period)?.[1]||period]);aoa.push([]);
  aoa.push(columns.map(c=>c.label));
  rows.forEach(r=>aoa.push(columns.map(c=>c.calc?c.calc(r):r[c.key])));
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws['!cols']=columns.map(c=>({wch:c.width||18}));
  ws['!freeze']={xSplit:0,ySplit:5};
  return ws;
}
function addVariance(r){
  r.varYear=r.real26-r.real25;r.varYearPct=r.real25? r.varYear/r.real25:null;
  r.varBudget=r.real26-r.budget26;r.varBudgetPct=r.budget26? r.varBudget/r.budget26:null;return r;
}
function costColumns(order=['group','brand','account','manager','center']){
  const names={group:'AGRUPADOR',brand:'CIA MARCA',account:'CTA CONTABLE',manager:'GERENCIA',center:'CENTRO DE GESTION'};
  const cols=order.map(k=>({key:k,label:names[k],width:k==='account'||k==='group'?34:22}));
  return cols.concat([
    {key:'real25',label:'2025',width:16},{key:'budget26',label:'PTTO. 2026',width:16},{key:'real26',label:'2026',width:16},
    {key:'varYear',label:'25 vs 26 $',width:16},{key:'varYearPct',label:'25 vs 26 %',width:14},
    {key:'varBudget',label:'PTTO vs 26 $',width:16},{key:'varBudgetPct',label:'PTTO vs 26 %',width:14}
  ]);
}
function buildCostWorkbook(rows,period){
  const wb=XLSX.utils.book_new();
  const oper=rows.filter(r=>normalize(r.manager)!=='MANTTO.');
  const mant=rows.filter(r=>normalize(r.manager)==='MANTTO.');
  const prem=rows.filter(r=>/5101\.(0403|0404|0901)/.test(String(r.account)));
  const smo=rows.filter(r=>/5101\.(08|13)/.test(String(r.group))||/5101\.(08|13)/.test(String(r.account)));
  const make=(data,keys)=>sortBy(aggregate(data,keys,['real25','budget26','real26']).map(addVariance),keys);
  XLSX.utils.book_append_sheet(wb,reportSheet('COSTOS OPERATIVOS',make(oper,['group','brand','account','manager','center']),costColumns(),period),'COSTOS OPERATIVOS');
  XLSX.utils.book_append_sheet(wb,reportSheet('COSTO MANTTO',make(mant,['manager','group','account','brand','center']),costColumns(['manager','group','account','brand','center']),period),'COSTO MANTTO');
  XLSX.utils.book_append_sheet(wb,reportSheet('PREMISAS 2026',make(prem,['account','brand','group','center','manager']),costColumns(['account','brand','group','center','manager']),period),'premisas 26');
  XLSX.utils.book_append_sheet(wb,reportSheet('SMO',make(smo,['manager','group','account','brand','center']),costColumns(['manager','group','account','brand','center']),period),'SMO');
  return wb;
}
function buildExpenseWorkbook(rows,period){
  const wb=XLSX.utils.book_new();
  const data=sortBy(aggregate(rows,['manager','group','account','brand','center'],['real25','budget26','real26']).map(addVariance),['manager','group','account']);
  XLSX.utils.book_append_sheet(wb,reportSheet('GASTOS RCV',data,costColumns(['manager','group','account','brand','center']),period),'GASTOS RCV');
  const byManager=sortBy(aggregate(rows,['manager'],['real25','budget26','real26']).map(addVariance),['manager']);
  XLSX.utils.book_append_sheet(wb,reportSheet('RESUMEN POR GERENCIA',byManager,[{key:'manager',label:'GERENCIA',width:28},...costColumns([])],period),'RESUMEN GERENCIA');
  return wb;
}
function buildXpvWorkbook(rows,period){
  const wb=XLSX.utils.book_new();
  const data=sortBy(aggregate(rows,['manager','account','center','brand'],['real','budget']).map(r=>({...r,varBudget:r.real-r.budget,varBudgetPct:r.budget?(r.real-r.budget)/r.budget:null})),['manager','account']);
  const cols=[{key:'manager',label:'GERENCIA',width:28},{key:'account',label:'CTA CONTABLE',width:34},{key:'center',label:'CENTRO / SUBLIBRO',width:34},{key:'brand',label:'CIA MARCA',width:24},{key:'real',label:'REAL GESTION',width:17},{key:'budget',label:'PRESUPUESTO',width:17},{key:'varBudget',label:'DIFERENCIA $',width:17},{key:'varBudgetPct',label:'DIFERENCIA %',width:15}];
  XLSX.utils.book_append_sheet(wb,reportSheet('PRODUCTIVIDAD XPV RCV',data,cols,period),'PRODUCTIVIDAD XPV');
  return wb;
}

function wbBlob(wb){const arr=XLSX.write(wb,{bookType:'xlsx',type:'array',compression:true});return new Blob([arr],{type:'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'})}
function saveBlob(blob,name){const a=document.createElement('a');a.href=URL.createObjectURL(blob);a.download=name;document.body.appendChild(a);a.click();setTimeout(()=>{URL.revokeObjectURL(a.href);a.remove()},500)}

async function processAll(){
  try{
    $('#processBtn').disabled=true;$('#processBtn').textContent='Procesando…';$('#statusText').textContent='Agrupando información y generando libros Excel.';
    await new Promise(r=>setTimeout(r,50));
    const catalog=getCatalog(),cutoff=monthSelect.value;
    const costRows=parseJD(state.sources.cost.sheet,'cost',catalog,cutoff);
    const expenseRows=parseJD(state.sources.expense.sheet,'expense',catalog,cutoff);
    const xpvRows=parseJD(state.sources.xpv.sheet,'xpv',catalog,cutoff);
    const costWb=buildCostWorkbook(costRows,cutoff), expWb=buildExpenseWorkbook(expenseRows,cutoff), xpvWb=buildXpvWorkbook(xpvRows,cutoff);
    const label=MONTHS.find(x=>x[0]===cutoff)?.[1]||cutoff;
    state.outputs={
      [`COSTO RCV 2026.xlsx`]:wbBlob(costWb),
      [`Gastos RCV 2026.xlsx`]:wbBlob(expWb),
      [`Productividad XPV ${label} 26 RCV.xlsx`]:wbBlob(xpvWb)
    };
    renderExecutiveDashboard(costRows,expenseRows,xpvRows,cutoff);
    renderExtendedSuite(costRows,expenseRows,xpvRows,cutoff);
    renderDownloads();
    $('#dashboard').classList.remove('hidden');
    $('#resultados').classList.remove('hidden');
    $('#dashboard').scrollIntoView({behavior:'smooth'});
    $('#statusText').textContent=`Análisis listo: ${costRows.length.toLocaleString('es-MX')} filas de costo, ${expenseRows.length.toLocaleString('es-MX')} de gastos y ${xpvRows.length.toLocaleString('es-MX')} de XPV procesadas.`;
  }catch(err){console.error(err);alert('Ocurrió un error al procesar los reportes: '+err.message);$('#statusText').textContent='No fue posible completar el procesamiento.'}
  finally{$('#processBtn').disabled=false;$('#processBtn').textContent='✨ Procesar y generar reportes'}
}
function renderDownloads(){
  $('#downloadGrid').innerHTML=Object.entries(state.outputs).map(([name,blob],i)=>`<div class="download-card"><div><h4>${escapeHtml(name)}</h4><p>${formatBytes(blob.size)} · Excel generado</p></div><button data-name="${escapeHtml(name)}" title="Descargar">⇩</button></div>`).join('');
  document.querySelectorAll('.download-card button').forEach(b=>b.onclick=()=>saveBlob(state.outputs[b.dataset.name],b.dataset.name));
}
async function downloadZip(){
  if(!Object.keys(state.outputs).length)return;
  const zip=new JSZip();Object.entries(state.outputs).forEach(([n,b])=>zip.file(n,b));
  zip.file('LEEME.txt','Reportes generados automáticamente por REPORT.IA RCV.\nProcesamiento realizado localmente en el navegador.');
  const blob=await zip.generateAsync({type:'blob',compression:'DEFLATE'});saveBlob(blob,'FINAL.zip');
}
function formatBytes(n){if(n<1024)return n+' B';if(n<1048576)return (n/1024).toFixed(1)+' KB';return (n/1048576).toFixed(1)+' MB'}
function escapeHtml(s){return String(s).replace(/[&<>'"]/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[c]))}


function money(v){return new Intl.NumberFormat('es-MX',{style:'currency',currency:'MXN',maximumFractionDigits:0}).format(num(v))}
function pct(v){return new Intl.NumberFormat('es-MX',{style:'percent',minimumFractionDigits:1,maximumFractionDigits:1}).format(Number.isFinite(v)?v:0)}
function sum(rows,key){return rows.reduce((a,r)=>a+num(r[key]),0)}
function destroyCharts(){(state.charts||[]).forEach(c=>{try{c.destroy()}catch(e){}});state.charts=[]}
function trendClass(v,inverse=false){if(Math.abs(v)<.001)return 'neutral';const good=inverse?v<0:v>0;return good?'good':'bad'}
function shortMoney(v){const a=Math.abs(v);if(a>=1e9)return '$'+(v/1e9).toFixed(1)+'B';if(a>=1e6)return '$'+(v/1e6).toFixed(1)+'M';if(a>=1e3)return '$'+(v/1e3).toFixed(0)+'K';return money(v)}

function renderExecutiveDashboard(costRows,expenseRows,xpvRows,cutoff){
  destroyCharts();
  const financial=[...costRows,...expenseRows];
  const real25=sum(financial,'real25'), budget26=sum(financial,'budget26'), real26=sum(financial,'real26');
  const yearVar=real25?(real26-real25)/real25:0;
  const budgetVar=budget26?(real26-budget26)/budget26:0;
  const xpvReal=sum(xpvRows,'real'),xpvBudget=sum(xpvRows,'budget'),xpvVar=xpvBudget?(xpvReal-xpvBudget)/xpvBudget:0;
  const managers=new Set(financial.map(r=>r.manager).filter(Boolean));
  $('#dashboardSubtitle').textContent=`Corte acumulado a ${MONTHS.find(x=>x[0]===cutoff)?.[1]||cutoff} · ${managers.size} gerencias analizadas.`;
  const cards=[
    ['REAL ACUMULADO 2026',money(real26),`${money(real26-real25)} vs 2025`,yearVar,'Costo/Gasto acumulado'],
    ['PRESUPUESTO 2026',money(budget26),`${money(real26-budget26)} de desviación`,budgetVar,'Comparativo contra presupuesto'],
    ['PRODUCTIVIDAD XPV',money(xpvReal),`${money(xpvReal-xpvBudget)} vs presupuesto`,xpvVar,'Resultado acumulado XPV'],
    ['GERENCIAS ANALIZADAS',managers.size.toLocaleString('es-MX'),`${financial.length.toLocaleString('es-MX')} registros financieros`,0,'Cobertura del análisis']
  ];
  $('#kpiGrid').innerHTML=cards.map((c,i)=>`<article class="kpi-card"><div class="kpi-label">${c[0]}</div><div class="kpi-value">${c[1]}</div><div class="kpi-detail">${c[4]}</div><span class="kpi-trend ${i===3?'neutral':trendClass(c[3],i<2)}">${i===3?'Información consolidada':pct(c[3])+' · '+c[2]}</span></article>`).join('');

  const monthly=MONTHS.map(([value,label])=>{const rows=financial.filter(r=>monthRank(r.period)===monthRank(value));return {label,real25:sum(rows,'real25'),budget26:sum(rows,'budget26'),real26:sum(rows,'real26')}}).filter((_,i)=>i<=monthRank(cutoff));
  const managerAgg=aggregate(financial,['manager'],['real25','budget26','real26']).map(addVariance).sort((a,b)=>Math.abs(b.varBudget)-Math.abs(a.varBudget));
  const xpvManagers=aggregate(xpvRows,['manager'],['real','budget']).map(r=>({...r,varBudget:r.real-r.budget})).sort((a,b)=>Math.abs(b.varBudget)-Math.abs(a.varBudget)).slice(0,8);

  const common={responsive:true,maintainAspectRatio:false,plugins:{legend:{labels:{usePointStyle:true,boxWidth:8,font:{size:10}}}},scales:{x:{grid:{display:false},ticks:{font:{size:9}}},y:{ticks:{font:{size:9},callback:v=>shortMoney(v)},grid:{color:'rgba(148,163,184,.16)'}}}};
  state.charts.push(new Chart($('#trendChart'),{type:'line',data:{labels:monthly.map(x=>x.label),datasets:[{label:'Real 2025',data:monthly.map(x=>x.real25),tension:.35,borderWidth:2,pointRadius:2},{label:'Presupuesto 2026',data:monthly.map(x=>x.budget26),tension:.35,borderWidth:2,borderDash:[6,5],pointRadius:2},{label:'Real 2026',data:monthly.map(x=>x.real26),tension:.35,borderWidth:3,pointRadius:3}]},options:common}));
  state.charts.push(new Chart($('#budgetChart'),{type:'doughnut',data:{labels:['Real 2026','Presupuesto restante'],datasets:[{data:[Math.max(real26,0),Math.max(budget26-real26,0)]}]},options:{responsive:true,maintainAspectRatio:false,cutout:'68%',plugins:{legend:{position:'bottom',labels:{usePointStyle:true,boxWidth:8,font:{size:10}}},tooltip:{callbacks:{label:c=>`${c.label}: ${money(c.raw)}`}}}}}));
  state.charts.push(new Chart($('#xpvChart'),{type:'bar',data:{labels:xpvManagers.map(x=>x.manager),datasets:[{label:'Real XPV',data:xpvManagers.map(x=>x.real)},{label:'Presupuesto',data:xpvManagers.map(x=>x.budget)}]},options:{...common,indexAxis:'y'}}));

  $('#managerTable').innerHTML=managerAgg.slice(0,10).map(r=>`<tr><td><strong>${escapeHtml(r.manager)}</strong></td><td class="num">${money(r.real26)}</td><td class="num">${money(r.budget26)}</td><td class="num">${money(r.varBudget)}</td><td class="num"><span class="pill ${r.varBudgetPct>0?'bad':'good'}">${pct(r.varBudgetPct||0)}</span></td></tr>`).join('')||'<tr><td colspan="5">Sin información suficiente.</td></tr>';

  const worst=managerAgg[0];
  const over=managerAgg.filter(r=>r.varBudget>0).length;
  const under=managerAgg.filter(r=>r.varBudget<=0).length;
  const bestXpv=[...xpvManagers].sort((a,b)=>(b.real-b.budget)-(a.real-a.budget))[0];
  const alerts=[];
  alerts.push({t:budgetVar>0?'⚠️ Presupuesto bajo presión':'✅ Control presupuestal',p:`El real acumulado 2026 está ${pct(Math.abs(budgetVar))} ${budgetVar>0?'por encima':'por debajo'} del presupuesto consolidado.`});
  if(worst)alerts.push({t:'📌 Mayor desviación detectada',p:`${worst.manager} presenta una desviación de ${money(worst.varBudget)} frente al presupuesto (${pct(worst.varBudgetPct||0)}).`});
  alerts.push({t:'🏢 Distribución por gerencias',p:`${over} gerencias se encuentran por encima del presupuesto y ${under} se mantienen en línea o por debajo.`});
  if(bestXpv)alerts.push({t:'📈 Señal XPV',p:`${bestXpv.manager} destaca en el comparativo XPV con una diferencia de ${money(bestXpv.real-bestXpv.budget)} frente al presupuesto.`});
  $('#alertsList').innerHTML=alerts.map(a=>`<div class="alert-item"><strong>${a.t}</strong><p>${a.p}</p></div>`).join('');
}

/* ===== REPORT.IA Executive Suite: módulos avanzados ===== */
function severityFromVariance(v){
  if(v<=0) return 'good';
  if(v<=0.05) return 'warn';
  return 'bad';
}
function severityLabel(v){
  const s=severityFromVariance(v); return s==='good'?'🟢 FAVORABLE':s==='warn'?'🟡 ATENCIÓN':'🔴 CRÍTICO';
}
function safeChart(canvasId,config){
  const el=document.getElementById(canvasId); if(!el) return null;
  const c=new Chart(el,config); state.charts.push(c); return c;
}
function topByAbs(rows,key,n=8){return [...rows].sort((a,b)=>Math.abs(num(b[key]))-Math.abs(num(a[key]))).slice(0,n)}
function cumulative(values){let a=0;return values.map(v=>a+=num(v))}

function renderExtendedSuite(costRows,expenseRows,xpvRows,cutoff){
  state.analysis={costRows,expenseRows,xpvRows,cutoff};
  ['gerencias','costos','gastos','xpv','tendencias','inteligencia'].forEach(id=>document.getElementById(id)?.classList.remove('hidden'));
  const managers=[...new Set([...costRows,...expenseRows,...xpvRows].map(r=>r.manager).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
  const sel=$('#managerSelect');
  if(sel){
    sel.innerHTML=managers.map(m=>`<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
    sel.onchange=()=>renderManagerDetail(sel.value);
    if(managers.length) renderManagerDetail(managers[0]);
  }
  renderFinancialModule('cost',costRows,'costKpis','costManagerChart','costGroupChart','group');
  renderFinancialModule('expense',expenseRows,'expenseKpis','expenseManagerChart','expenseAccountChart','account');
  renderXpvModule(xpvRows);
  renderTrendModule([...costRows,...expenseRows],cutoff);
  renderSmartModule(costRows,expenseRows,xpvRows,managers);
  renderIntelligenceEdition(costRows,expenseRows,xpvRows,cutoff,managers);
}

function renderFinancialModule(kind,rows,kpiId,managerChartId,detailChartId,detailKey){
  const real25=sum(rows,'real25'),budget=sum(rows,'budget26'),real=sum(rows,'real26');
  const vB=budget?(real-budget)/budget:0, vY=real25?(real-real25)/real25:0;
  const managerAgg=aggregate(rows,['manager'],['budget26','real26']).map(addVariance).sort((a,b)=>Math.abs(b.varBudget)-Math.abs(a.varBudget));
  const detailAgg=aggregate(rows,[detailKey],['budget26','real26']).map(addVariance);
  const kpi=document.getElementById(kpiId);
  if(kpi) kpi.innerHTML=[
    ['REAL 2026',money(real),`${pct(vY)} vs 2025`,trendClass(vY,true)],
    ['PRESUPUESTO',money(budget),`${money(real-budget)} de desviación`,trendClass(vB,true)],
    ['VARIACIÓN',pct(vB),vB>0?'Por encima del presupuesto':'Dentro / debajo del presupuesto',trendClass(vB,true)],
    ['GERENCIAS',new Set(rows.map(r=>r.manager)).size.toLocaleString('es-MX'),`${rows.length.toLocaleString('es-MX')} registros`,'neutral']
  ].map(x=>`<article class="kpi-card"><div class="kpi-label">${x[0]}</div><div class="kpi-value">${x[1]}</div><div class="kpi-detail">${x[2]}</div><span class="kpi-trend ${x[3]}">${kind==='cost'?'Control de costos':'Control de gastos'}</span></article>`).join('');
  const topM=managerAgg.slice(0,10);
  safeChart(managerChartId,{type:'bar',data:{labels:topM.map(x=>x.manager),datasets:[{label:'Real 2026',data:topM.map(x=>x.real26)},{label:'Presupuesto',data:topM.map(x=>x.budget26)}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'}},scales:{x:{ticks:{callback:v=>shortMoney(v)}}}}});
  const topD=topByAbs(detailAgg,'varBudget',8);
  safeChart(detailChartId,{type:'bar',data:{labels:topD.map(x=>String(x[detailKey]||'Sin clasificar').slice(0,28)),datasets:[{label:'Desviación $',data:topD.map(x=>x.varBudget)}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{display:false}},scales:{x:{ticks:{callback:v=>shortMoney(v)}}}}});
}

function renderXpvModule(rows){
  const real=sum(rows,'real'),budget=sum(rows,'budget'),v=budget?(real-budget)/budget:0;
  const agg=aggregate(rows,['manager'],['real','budget']).map(r=>({...r,varBudget:r.real-r.budget,varBudgetPct:r.budget?(r.real-r.budget)/r.budget:0})).sort((a,b)=>(b.real-b.budget)-(a.real-a.budget));
  $('#xpvKpis').innerHTML=[
    ['XPV REAL',money(real),'Resultado acumulado'],['XPV PRESUPUESTO',money(budget),'Objetivo acumulado'],['DESVIACIÓN XPV',money(real-budget),pct(v)],['MEJOR GERENCIA',agg[0]?.manager||'—',agg[0]?money(agg[0].varBudget):'Sin datos']
  ].map(x=>`<article class="kpi-card"><div class="kpi-label">${x[0]}</div><div class="kpi-value">${x[1]}</div><div class="kpi-detail">${x[2]}</div></article>`).join('');
  safeChart('xpvFullChart',{type:'bar',data:{labels:agg.map(x=>x.manager),datasets:[{label:'Real XPV',data:agg.map(x=>x.real)},{label:'Presupuesto',data:agg.map(x=>x.budget)}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'}},scales:{x:{ticks:{callback:v=>shortMoney(v)}}}}});
}

function renderTrendModule(financial,cutoff){
  const monthly=MONTHS.map(([value,label])=>{
    const rows=financial.filter(r=>monthRank(r.period)===monthRank(value));
    return {label,real25:sum(rows,'real25'),budget:sum(rows,'budget26'),real:sum(rows,'real26')};
  }).filter((_,i)=>i<=monthRank(cutoff));
  const c25=cumulative(monthly.map(x=>x.real25)), cb=cumulative(monthly.map(x=>x.budget)), cr=cumulative(monthly.map(x=>x.real));
  safeChart('cumulativeChart',{type:'line',data:{labels:monthly.map(x=>x.label),datasets:[{label:'Acumulado 2025',data:c25,tension:.3},{label:'Presupuesto 2026',data:cb,tension:.3,borderDash:[6,5]},{label:'Acumulado 2026',data:cr,tension:.3,borderWidth:3}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{ticks:{callback:v=>shortMoney(v)}}}}});
  const last=monthly.length-1;
  const lastMonth=monthly[last]||{real:0,budget:0,real25:0};
  const prev=monthly[last-1]||{real:0};
  const mom=prev.real?(lastMonth.real-prev.real)/Math.abs(prev.real):0;
  const budgetVar=lastMonth.budget?(lastMonth.real-lastMonth.budget)/lastMonth.budget:0;
  const yoy=lastMonth.real25?(lastMonth.real-lastMonth.real25)/lastMonth.real25:0;
  $('#trendCards').innerHTML=[['VARIACIÓN MENSUAL',pct(mom),'Cambio del último mes vs mes previo'],['VS PRESUPUESTO DEL MES',pct(budgetVar),budgetVar>0?'Presión presupuestal':'Resultado favorable'],['VS MISMO PERIODO 2025',pct(yoy),'Comparativo interanual']].map(x=>`<article class="trend-card"><span>${x[0]}</span><strong>${x[1]}</strong><p>${x[2]}</p></article>`).join('');
}

function managerFinancial(manager){
  const a=state.analysis||{};
  return [...(a.costRows||[]),...(a.expenseRows||[])].filter(r=>r.manager===manager);
}
function renderManagerDetail(manager){
  if(!state.analysis)return;
  const rows=managerFinancial(manager), xpv=(state.analysis.xpvRows||[]).filter(r=>r.manager===manager), cutoff=state.analysis.cutoff;
  const r25=sum(rows,'real25'),budget=sum(rows,'budget26'),real=sum(rows,'real26'),vb=budget?(real-budget)/budget:0,vy=r25?(real-r25)/r25:0;
  const xr=sum(xpv,'real'),xb=sum(xpv,'budget'),xv=xb?(xr-xb)/xb:0;
  $('#managerKpis').innerHTML=[['REAL 2026',money(real),`${pct(vy)} vs 2025`],['PRESUPUESTO',money(budget),`${money(real-budget)} desviación`],['VARIACIÓN',pct(vb),severityLabel(vb)],['XPV',money(xr),`${pct(xv)} vs presupuesto`]].map(x=>`<article class="kpi-card"><div class="kpi-label">${x[0]}</div><div class="kpi-value">${x[1]}</div><div class="kpi-detail">${x[2]}</div></article>`).join('');
  const status=$('#managerStatus'); status.textContent=severityLabel(vb); status.className='status-chip '+severityFromVariance(vb);
  // destroy only manager detail charts before redraw
  ['managerTrendChart','managerXpvChart'].forEach(id=>{const el=document.getElementById(id);const existing=Chart.getChart(el);if(existing){existing.destroy();state.charts=state.charts.filter(c=>c!==existing)}});
  const monthly=MONTHS.map(([value,label])=>{const rr=rows.filter(r=>monthRank(r.period)===monthRank(value));return {label,b:sum(rr,'budget26'),r:sum(rr,'real26')}}).filter((_,i)=>i<=monthRank(cutoff));
  safeChart('managerTrendChart',{type:'line',data:{labels:monthly.map(x=>x.label),datasets:[{label:'Real 2026',data:monthly.map(x=>x.r),tension:.35,borderWidth:3},{label:'Presupuesto',data:monthly.map(x=>x.b),tension:.35,borderDash:[6,5]}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{ticks:{callback:v=>shortMoney(v)}}}}});
  safeChart('managerXpvChart',{type:'doughnut',data:{labels:['Real XPV','Brecha a presupuesto'],datasets:[{data:[Math.max(xr,0),Math.max(xb-xr,0)]}]},options:{responsive:true,maintainAspectRatio:false,cutout:'68%',plugins:{legend:{position:'bottom'}}}});
  const drivers=topByAbs(aggregate(rows,['account','group'],['real26','budget26']).map(addVariance),'varBudget',10);
  $('#managerDrivers').innerHTML=drivers.map(d=>`<tr><td><strong>${escapeHtml(d.account||d.group||'Sin clasificar')}</strong><br><small>${escapeHtml(d.group||'')}</small></td><td class="num">${money(d.real26)}</td><td class="num">${money(d.budget26)}</td><td class="num"><span class="pill ${d.varBudget>0?'bad':'good'}">${money(d.varBudget)}</span></td></tr>`).join('')||'<tr><td colspan="4">Sin información suficiente.</td></tr>';
  const biggest=drivers[0];
  $('#managerNarrative').innerHTML=[
    {t:severityLabel(vb)+' · Posición presupuestal',p:`${manager} se encuentra ${pct(Math.abs(vb))} ${vb>0?'por encima':'por debajo'} del presupuesto acumulado.`},
    {t:'📊 Comparativo interanual',p:`El resultado 2026 presenta una variación de ${pct(vy)} frente al periodo comparable de 2025.`},
    {t:'🚌 Productividad XPV',p:`La productividad registra ${money(xr)} frente a un presupuesto de ${money(xb)} (${pct(xv)} de variación).`},
    biggest?{t:'🎯 Principal impulsor',p:`La cuenta ${biggest.account||biggest.group} explica una desviación de ${money(biggest.varBudget)} y debe revisarse como primer foco.`}:null
  ].filter(Boolean).map(a=>`<div class="alert-item"><strong>${a.t}</strong><p>${a.p}</p></div>`).join('');
}

function renderSmartModule(costRows,expenseRows,xpvRows,managers){
  const fin=[...costRows,...expenseRows];
  const managerAgg=aggregate(fin,['manager'],['real25','budget26','real26']).map(addVariance).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const favorable=managerAgg.filter(x=>x.varBudget<=0).length, critical=managerAgg.filter(x=>x.varBudgetPct>0.05).length;
  const totalB=sum(fin,'budget26'),totalR=sum(fin,'real26'),v=totalB?(totalR-totalB)/totalB:0;
  const top=managerAgg[0], best=[...managerAgg].sort((a,b)=>a.varBudgetPct-b.varBudgetPct)[0];
  const accountAgg=aggregate(fin,['account'],['budget26','real26']).map(addVariance);const driver=topByAbs(accountAgg,'varBudget',1)[0];
  $('#smartSummary').innerHTML=[
    ['🧭','POSICIÓN GENERAL',`El consolidado está ${pct(Math.abs(v))} ${v>0?'por encima':'por debajo'} del presupuesto.`],
    ['🏢','GERENCIAS FAVORABLES',`${favorable} de ${managerAgg.length} gerencias están en línea o por debajo de presupuesto.`],
    ['🚨','FOCOS CRÍTICOS',`${critical} gerencias superan en más de 5% su presupuesto acumulado.`],
    ['🎯','PRINCIPAL DRIVER',driver?`${driver.account} concentra una desviación de ${money(driver.varBudget)}.`:'No se identificó un driver dominante.']
  ].map(x=>`<article class="smart-card"><div class="smart-icon">${x[0]}</div><h3>${x[1]}</h3><p>${x[2]}</p></article>`).join('');
  $('#managerMap').innerHTML=managerAgg.map(m=>{const s=severityFromVariance(m.varBudgetPct||0);return `<div class="manager-tile" data-manager="${escapeHtml(m.manager)}"><div class="top"><strong>${escapeHtml(m.manager)}</strong><span class="dot ${s}"></span></div><small>${pct(m.varBudgetPct||0)} · ${money(m.varBudget)} vs presupuesto</small></div>`}).join('');
  document.querySelectorAll('.manager-tile').forEach(t=>t.onclick=()=>{const m=t.dataset.manager;$('#managerSelect').value=m;renderManagerDetail(m);document.getElementById('gerencias').scrollIntoView({behavior:'smooth'})});
  $('#actionList').innerHTML=[
    top?{t:'1 · Revisar desviación prioritaria',p:`Analizar ${top.manager}, actualmente con ${pct(top.varBudgetPct||0)} frente al presupuesto.`}:null,
    driver?{t:'2 · Validar cuenta de mayor impacto',p:`Revisar movimientos y supuestos de ${driver.account}, con impacto de ${money(driver.varBudget)}.`}:null,
    best?{t:'3 · Replicar mejor práctica',p:`Tomar como referencia el comportamiento de ${best.manager}, con ${pct(best.varBudgetPct||0)} vs presupuesto.`}:null,
    {t:'4 · Cerrar ciclo de reporte',p:'Descargar los archivos finales después de validar los focos críticos y documentar las variaciones relevantes.'}
  ].filter(Boolean).map(a=>`<div class="alert-item"><strong>${a.t}</strong><p>${a.p}</p></div>`).join('');
}


/* ===== REPORT.IA RCV · Intelligence Edition ===== */
function clamp(n,min,max){return Math.min(max,Math.max(min,n))}
function managerSnapshot(manager){
  const a=state.analysis||{};
  const fin=[...(a.costRows||[]),...(a.expenseRows||[])].filter(r=>r.manager===manager);
  const xp=(a.xpvRows||[]).filter(r=>r.manager===manager);
  const real=sum(fin,'real26'), budget=sum(fin,'budget26'), real25=sum(fin,'real25');
  const xr=sum(xp,'real'), xb=sum(xp,'budget');
  return {manager,real,budget,real25,varPct:budget?(real-budget)/budget:0,varValue:real-budget,yoy:real25?(real-real25)/real25:0,xpvReal:xr,xpvBudget:xb,xpvVar:xb?(xr-xb)/xb:0};
}
function renderIntelligenceEdition(costRows,expenseRows,xpvRows,cutoff,managers){
  ['pulso','comparador'].forEach(id=>document.getElementById(id)?.classList.remove('hidden'));
  const fin=[...costRows,...expenseRows];
  const totalR=sum(fin,'real26'), totalB=sum(fin,'budget26'), total25=sum(fin,'real25');
  const budgetVar=totalB?(totalR-totalB)/totalB:0, yoy=total25?(totalR-total25)/total25:0;
  const xr=sum(xpvRows,'real'),xb=sum(xpvRows,'budget'),xpvVar=xb?(xr-xb)/xb:0;
  const snaps=managers.map(managerSnapshot);
  const favorable=snaps.filter(x=>x.varPct<=0).length;
  const critical=snaps.filter(x=>x.varPct>.05).length;
  const budgetScore=clamp(100-(Math.max(0,budgetVar)*240),0,100);
  const coverageScore=snaps.length?clamp((favorable/snaps.length)*100,0,100):50;
  const trendScore=clamp(75-(Math.max(0,yoy)*100)+Math.max(0,-yoy)*60,0,100);
  const xpvScore=clamp(75+(xpvVar*120),0,100);
  const score=Math.round(budgetScore*.38+coverageScore*.22+trendScore*.18+xpvScore*.22);
  const ring=$('#pulseRing'); if(ring) ring.style.setProperty('--score',score);
  $('#pulseScore').textContent=score;
  const label=score>=85?'Desempeño sólido':score>=70?'Operación estable':score>=55?'Atención preventiva':'Intervención prioritaria';
  $('#pulseLabel').textContent=label;
  $('#pulseDescription').textContent=`${favorable} de ${snaps.length} gerencias se encuentran dentro o por debajo del presupuesto; ${critical} requieren atención crítica. El XPV presenta ${pct(xpvVar)} frente al objetivo.`;
  const hour=new Date().getHours(); const greeting=hour<12?'Buenos días':hour<19?'Buenas tardes':'Buenas noches';
  $('#executiveGreeting').textContent=`${greeting}. El Pulso RCV se encuentra en ${score}/100.`;
  $('#executiveHeadline').textContent=budgetVar>0?`La operación está ${pct(Math.abs(budgetVar))} por encima del presupuesto consolidado. REPORT.IA detectó ${critical} focos críticos.`:`La operación está ${pct(Math.abs(budgetVar))} por debajo del presupuesto consolidado y mantiene ${favorable} gerencias en posición favorable.`;
  renderWhatChanged(fin,xpvRows,cutoff);
  renderAlertCenter(snaps);
  setupComparator(managers);
}
function renderWhatChanged(fin,xpvRows,cutoff){
  const rank=monthRank(cutoff), current=fin.filter(r=>monthRank(r.period)===rank), prev=fin.filter(r=>monthRank(r.period)===rank-1);
  const curR=sum(current,'real26'),prevR=sum(prev,'real26'),change=prevR?(curR-prevR)/Math.abs(prevR):0;
  const curB=sum(current,'budget26'),budgetV=curB?(curR-curB)/curB:0;
  const curXp=xpvRows.filter(r=>monthRank(r.period)===rank),prevXp=xpvRows.filter(r=>monthRank(r.period)===rank-1);
  const cx=sum(curXp,'real'),px=sum(prevXp,'real'),xchange=px?(cx-px)/Math.abs(px):0;
  const curMap=aggregate(current,['manager'],['real26','budget26']).map(addVariance);
  const prevMap=aggregate(prev,['manager'],['real26','budget26']).map(addVariance);
  const prevDict=Object.fromEntries(prevMap.map(x=>[x.manager,severityFromVariance(x.varBudgetPct||0)]));
  const worsened=curMap.filter(x=>prevDict[x.manager]&&severityFromVariance(x.varBudgetPct||0)==='bad'&&prevDict[x.manager]!=='bad').length;
  $('#changeList').innerHTML=[
    ['↕','Movimiento mensual',`${money(curR-prevR)} (${pct(change)}) frente al mes anterior.`],
    ['◎','Presupuesto del mes',`${pct(Math.abs(budgetV))} ${budgetV>0?'por encima':'por debajo'} del presupuesto del periodo.`],
    ['↗','Productividad XPV',`XPV cambió ${pct(xchange)} respecto al mes anterior.`],
    ['⚑','Cambio de semáforo',`${worsened} gerencias entraron a condición crítica este periodo.`]
  ].map(x=>`<div class="change-row"><div class="change-icon">${x[0]}</div><div><strong>${x[1]}</strong><p>${x[2]}</p></div></div>`).join('');
}
function renderAlertCenter(snaps){
  const alerts=[];
  [...snaps].sort((a,b)=>b.varPct-a.varPct).slice(0,4).forEach(s=>{if(s.varPct>.05)alerts.push({s:'bad',m:s.manager,t:'Desviación crítica',p:`${s.manager} supera el presupuesto en ${pct(s.varPct)} (${money(s.varValue)}).`});else if(s.varPct>0)alerts.push({s:'warn',m:s.manager,t:'Presión presupuestal',p:`${s.manager} está ${pct(s.varPct)} por encima del presupuesto.`})});
  const weakXpv=[...snaps].filter(s=>s.xpvBudget&&s.xpvVar<-.05).sort((a,b)=>a.xpvVar-b.xpvVar).slice(0,2);
  weakXpv.forEach(s=>alerts.push({s:'warn',m:s.manager,t:'XPV bajo objetivo',p:`${s.manager} registra ${pct(s.xpvVar)} contra el presupuesto XPV.`}));
  const best=[...snaps].sort((a,b)=>a.varPct-b.varPct)[0]; if(best)alerts.push({s:'good',m:best.manager,t:'Mejor posición',p:`${best.manager} destaca con ${pct(best.varPct)} frente al presupuesto.`});
  $('#alertCount').textContent=alerts.filter(a=>a.s!=='good').length;
  $('#intelligenceAlerts').innerHTML=alerts.length?alerts.map(a=>`<div class="intelligence-alert ${a.s}" data-manager="${escapeHtml(a.m)}"><strong>${a.t} · ${escapeHtml(a.m)}</strong><p>${a.p}</p></div>`).join(''):'<div class="intelligence-alert good"><strong>Sin alertas relevantes</strong><p>No se detectaron focos críticos con los criterios actuales.</p></div>';
  document.querySelectorAll('.intelligence-alert[data-manager]').forEach(el=>el.onclick=()=>{const m=el.dataset.manager;$('#managerSelect').value=m;renderManagerDetail(m);document.getElementById('gerencias').scrollIntoView({behavior:'smooth'})});
}
function setupComparator(managers){
  const a=$('#compareA'),b=$('#compareB'); if(!a||!b)return;
  const options=managers.map(m=>`<option value="${escapeHtml(m)}">${escapeHtml(m)}</option>`).join('');
  a.innerHTML=options;b.innerHTML=options;if(managers.length>1)b.selectedIndex=1;
  const redraw=()=>renderComparison(a.value,b.value);a.onchange=redraw;b.onchange=redraw;if(managers.length)redraw();
}
function renderComparison(ma,mb){
  const A=managerSnapshot(ma),B=managerSnapshot(mb);
  const side=s=>`<article class="comparison-side"><h3>${escapeHtml(s.manager)}</h3>${[['Real 2026',money(s.real)],['Presupuesto',money(s.budget)],['Desviación',`${money(s.varValue)} · ${pct(s.varPct)}`],['Vs 2025',pct(s.yoy)],['XPV',money(s.xpvReal)],['XPV vs objetivo',pct(s.xpvVar)]].map(x=>`<div class="comparison-metric"><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join('')}</article>`;
  $('#comparisonCards').innerHTML=side(A)+'<div class="versus-badge">VS</div>'+side(B);
  const old=Chart.getChart(document.getElementById('comparisonChart'));if(old)old.destroy();
  safeChart('comparisonChart',{type:'bar',data:{labels:['Real 2026','Presupuesto','XPV Real','XPV Presupuesto'],datasets:[{label:A.manager,data:[A.real,A.budget,A.xpvReal,A.xpvBudget]},{label:B.manager,data:[B.real,B.budget,B.xpvReal,B.xpvBudget]}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{ticks:{callback:v=>shortMoney(v)}}}}});
  const winner=A.varPct<=B.varPct?A:B, other=winner===A?B:A;
  $('#comparisonNarrative').innerHTML=[
    {t:'🏆 Mejor control presupuestal',p:`${winner.manager} presenta la mejor posición relativa: ${pct(winner.varPct)} vs ${pct(other.varPct)}.`},
    {t:'📊 Brecha de ejecución',p:`La diferencia entre los resultados reales de ambas gerencias es ${money(Math.abs(A.real-B.real))}.`},
    {t:'🚌 Lectura XPV',p:`${A.xpvVar>=B.xpvVar?A.manager:B.manager} muestra la mejor posición de XPV frente a su objetivo.`}
  ].map(a=>`<div class="alert-item"><strong>${a.t}</strong><p>${a.p}</p></div>`).join('');
}
(function bindIntelligenceControls(){
  document.addEventListener('DOMContentLoaded',()=>{
    const p=document.getElementById('presentationBtn'); if(p)p.onclick=()=>{document.body.classList.toggle('presentation-mode');p.textContent=document.body.classList.contains('presentation-mode')?'✕ Salir de presentación':'▶ Modo presentación'};
    const f=document.getElementById('focusBtn'); if(f)f.onclick=()=>document.getElementById('inteligencia')?.scrollIntoView({behavior:'smooth'});
  });
})();


// ===== REPORT.IA RCV 9.1: resilient external library loader =====
async function rcvLoadScriptFallback(urls, globalCheck){
  if(globalCheck()) return true;
  for(const url of urls){
    try{
      await new Promise((resolve,reject)=>{
        const existing=[...document.scripts].find(s=>s.src===url);
        if(existing){
          if(globalCheck()) return resolve();
          existing.addEventListener('load',resolve,{once:true});
          existing.addEventListener('error',reject,{once:true});
          setTimeout(()=>globalCheck()?resolve():reject(new Error('timeout')),7000);
          return;
        }
        const s=document.createElement('script');
        s.src=url; s.async=true; s.crossOrigin='anonymous';
        s.onload=resolve; s.onerror=reject;
        document.head.appendChild(s);
        setTimeout(()=>{ if(!globalCheck()) reject(new Error('timeout')); },7000);
      });
      if(globalCheck()) return true;
    }catch(e){}
  }
  return globalCheck();
}
async function rcvEnsurePdfLibraries(){
  const canvasOk=await rcvLoadScriptFallback([
    'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
    'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'
  ],()=>typeof window.html2canvas==='function');
  const pdfOk=await rcvLoadScriptFallback([
    'https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js',
    'https://unpkg.com/jspdf@2.5.1/dist/jspdf.umd.min.js',
    'https://cdnjs.cloudflare.com/ajax/libs/jspdf/2.5.1/jspdf.umd.min.js'
  ],()=>!!window.jspdf?.jsPDF);
  return canvasOk && pdfOk;
}
function rcvToast(message,type=''){
  let el=document.getElementById('rcvExportToast');
  if(!el){
    el=document.createElement('div');el.id='rcvExportToast';el.className='rcv-export-toast';
    document.body.appendChild(el);
  }
  el.textContent=message;el.className='rcv-export-toast '+type+' show';
  clearTimeout(window.__rcvToastTimer);
  window.__rcvToastTimer=setTimeout(()=>el.classList.remove('show'),3500);
}

// ===== REPORT.IA RCV Intelligence Edition 2.0: Executive exports =====
(function(){
  const $ = (id)=>document.getElementById(id);
  function visibleReportRoot(){
    const candidates = [...document.querySelectorAll('main section, main .page, main .view, main .content-section')];
    const visible = candidates.find(el => {
      const s=getComputedStyle(el);
      return el.id!=='exportSuite' && s.display!=='none' && el.offsetParent!==null && el.offsetHeight>120;
    });
    return visible || document.querySelector('main') || document.body;
  }
  function stamp(){
    const d=new Date();
    return d.toISOString().slice(0,10);
  }
  async function exportPdf(element, filename, title){
    rcvToast('Preparando motor de PDF…');
    const libsReady=await rcvEnsurePdfLibraries();
    if(!libsReady){
      rcvToast('No fue posible cargar el motor de PDF.','error');
      if(confirm('¿Deseas usar Imprimir → Guardar como PDF?')) window.print();
      return;
    }
    let host=null;
    try{
      rcvToast('Preparando reporte completo…');
      if(document.fonts?.ready){try{await document.fonts.ready;}catch(_){}}

      const excludeSelector=[
        '#carga','.generator-studio','#automationCenter','#reportingCenter',
        '.sidebar','.exec-topbar','.rcv-v9-toolbar','.rcv-v9-panel',
        '.v10-fab-stack','.v10-sheet','.modal','.toast','[role="dialog"]',
        'input[type="file"]','button'
      ].join(',');

      const candidates=[...document.querySelectorAll(
        'main section, main .card, main .dashboard-section, main .report-section, '+
        '.tab-content, .panel, .chart-card, .table-card, .kpi-grid, .metrics-grid'
      )].filter(el=>{
        if(el.matches(excludeSelector) || el.closest(excludeSelector)) return false;
        const t=(el.innerText||'').trim();
        return t.length>20 || el.querySelector('canvas,svg,table,.kpi-card,.metric-card,.stat-card');
      });

      const sections=candidates.filter((el,i,arr)=>!arr.some((o,j)=>j!==i&&o.contains(el)));
      host=document.createElement('div');
      host.style.cssText='position:fixed;left:-30000px;top:0;width:1360px;background:#fff;padding:28px;z-index:-1;box-sizing:border-box;color:#0f172a;';

      const cover=document.createElement('div');
      cover.innerHTML='<div style="padding:28px 30px;margin-bottom:20px;border-radius:20px;background:linear-gradient(135deg,#0f172a,#1e293b,#0f4c5c);color:white"><div style="font:800 11px system-ui;letter-spacing:.14em;color:#7dd3fc">REPORT.IA RCV</div><div style="font:850 30px system-ui;margin-top:8px">'+title+'</div><div style="font:500 12px system-ui;margin-top:7px;color:#cbd5e1">Advanced Reporting Suite 16.0 · Reporte analítico completo</div><div style="font:500 10px system-ui;margin-top:16px;color:#94a3b8">'+new Date().toLocaleString('es-MX')+'</div></div>';
      host.appendChild(cover);

      const seen=new Set();
      for(const source of sections){
        const clone=source.cloneNode(true);
        clone.removeAttribute('hidden');
        clone.style.display='block';clone.style.visibility='visible';clone.style.opacity='1';
        clone.style.height='auto';clone.style.maxHeight='none';clone.style.overflow='visible';
        clone.style.margin='0 0 18px';clone.style.width='100%';clone.style.maxWidth='none';clone.style.transform='none';
        clone.querySelectorAll('[hidden]').forEach(el=>el.removeAttribute('hidden'));
        clone.querySelectorAll(excludeSelector).forEach(el=>el.remove());

        const oc=[...source.querySelectorAll('canvas')];
        const cc=[...clone.querySelectorAll('canvas')];
        cc.forEach((c,i)=>{
          const orig=oc[i];if(!orig)return;
          try{
            const img=document.createElement('img');
            img.src=orig.toDataURL('image/png');
            const r=orig.getBoundingClientRect();
            img.style.width=Math.max(r.width||orig.width,300)+'px';
            img.style.maxWidth='100%';img.style.height='auto';img.style.display='block';
            c.replaceWith(img);
          }catch(_){}
        });
        clone.querySelectorAll('svg').forEach(svg=>{svg.style.maxWidth='100%';svg.style.height='auto';});
        const fp=(clone.innerText||'').trim().slice(0,500);
        if(fp&&seen.has(fp))continue;
        if(fp)seen.add(fp);
        const wrap=document.createElement('div');
        wrap.style.cssText='margin-bottom:18px;padding:4px;';
        wrap.appendChild(clone);host.appendChild(wrap);
      }

      document.body.appendChild(host);
      await new Promise(r=>requestAnimationFrame(()=>requestAnimationFrame(r)));

      const canvas=await html2canvas(host,{scale:1.35,useCORS:true,allowTaint:false,backgroundColor:'#fff',logging:false,imageTimeout:15000,width:host.scrollWidth,height:host.scrollHeight,windowWidth:1360,scrollX:0,scrollY:0});
      if(!canvas||canvas.width<50||canvas.height<50)throw new Error('Captura vacía');

      const {jsPDF}=window.jspdf;
      const pdf=new jsPDF({orientation:'p',unit:'mm',format:'a4',compress:true});
      const pw=210,ph=297,margin=9,header=18,footer=8,usableW=pw-margin*2,usableH=ph-header-footer-margin;
      const pxPerMm=canvas.width/usableW,slicePx=Math.max(1,Math.floor(usableH*pxPerMm));
      let sy=0,page=1;
      while(sy<canvas.height){
        const sh=Math.min(slicePx,canvas.height-sy);
        const slice=document.createElement('canvas');slice.width=canvas.width;slice.height=sh;
        const ctx=slice.getContext('2d',{alpha:false});ctx.fillStyle='#fff';ctx.fillRect(0,0,slice.width,slice.height);ctx.drawImage(canvas,0,sy,canvas.width,sh,0,0,canvas.width,sh);
        if(page>1)pdf.addPage();
        pdf.setTextColor(100,116,139);pdf.setFont('helvetica','normal');pdf.setFontSize(7);
        pdf.text('REPORT.IA RCV · Advanced Reporting Suite 16.0',margin,10);pdf.text('Página '+page,pw-margin,10,{align:'right'});
        pdf.addImage(slice.toDataURL('image/jpeg',.93),'JPEG',margin,header,usableW,sh/pxPerMm,'','FAST');
        sy+=sh;page++;
      }
      pdf.save(filename);
      rcvToast('PDF completo generado ('+(page-1)+' páginas).','ok');
    }catch(err){
      console.error('PDF completo:',err);
      rcvToast('No fue posible generar el PDF completo automáticamente.','error');
      if(confirm('¿Deseas usar Imprimir → Guardar como PDF?'))window.print();
    }finally{try{host?.remove();}catch(_){}}
  }

  async function pdfVista(){
    await exportPdf(visibleReportRoot(),`REPORTIA_RCV_Vista_${stamp()}.pdf`,'Informe Ejecutivo · Vista actual');
  }
  async function pdfGerencia(){
    const sel=document.querySelector('select[id*="gerencia" i], select[name*="gerencia" i]');
    const g=sel?.selectedOptions?.[0]?.textContent?.trim() || 'Gerencia seleccionada';
    await exportPdf(visibleReportRoot(),`REPORTIA_RCV_${g.replace(/[^\w\-]+/g,'_')}_${stamp()}.pdf`,`Informe Ejecutivo · ${g}`);
  }
  function excelEjecutivo(){
    if(typeof XLSX==='undefined'){ alert('La librería de Excel no está disponible.'); return; }
    const wb=XLSX.utils.book_new();
    const summary=[
      ['REPORT.IA RCV · INFORME EJECUTIVO'],
      ['Fecha de generación',new Date().toLocaleString('es-MX')],
      ['Vista','Intelligence Edition 2.0'],
      [],
      ['Indicador','Valor']
    ];
    document.querySelectorAll('.kpi-card,.metric-card,.stat-card,.card').forEach(card=>{
      const t=card.querySelector('h3,h4,.title,.label')?.textContent?.trim();
      const v=card.querySelector('.value,.kpi-value,.metric-value,strong')?.textContent?.trim();
      if(t&&v&&t.length<80&&v.length<80) summary.push([t,v]);
    });
    const ws=XLSX.utils.aoa_to_sheet(summary);
    ws['!cols']=[{wch:38},{wch:28}];
    XLSX.utils.book_append_sheet(wb,ws,'Resumen Ejecutivo');
    XLSX.writeFile(wb,`REPORTIA_RCV_Ejecutivo_${stamp()}.xlsx`);
  }
  async function paquete(){
    if(typeof JSZip==='undefined'){ alert('La librería ZIP no está disponible.'); return; }
    const canvasReady=await rcvLoadScriptFallback([
      'https://cdn.jsdelivr.net/npm/html2canvas@1.4.1/dist/html2canvas.min.js',
      'https://unpkg.com/html2canvas@1.4.1/dist/html2canvas.min.js',
      'https://cdnjs.cloudflare.com/ajax/libs/html2canvas/1.4.1/html2canvas.min.js'
    ],()=>typeof window.html2canvas==='function');
    if(!canvasReady){ rcvToast('No fue posible preparar la captura para el paquete ejecutivo.','error'); return; }
    const zip=new JSZip();
    const root=visibleReportRoot();
    const canvas=await html2canvas(root,{scale:1.3,useCORS:true,backgroundColor:'#ffffff'});
    zip.file(`Vista_Ejecutiva_${stamp()}.png`,canvas.toDataURL('image/png').split(',')[1],{base64:true});
    zip.file('LEEME.txt',
      'REPORT.IA RCV · Intelligence Edition 2.0\n\nPaquete ejecutivo generado desde la vista activa.\n' +
      'Use los botones individuales para generar PDF y Excel ejecutivo. Los reportes operativos se mantienen en el módulo de descargas del portal.'
    );
    const blob=await zip.generateAsync({type:'blob'});
    const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
    a.download=`REPORTIA_RCV_Paquete_Ejecutivo_${stamp()}.zip`; a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),2000);
  }
  document.addEventListener('DOMContentLoaded',()=>{
    $('btnPdfVista')?.addEventListener('click',pdfVista);
    $('btnPdfGerencia')?.addEventListener('click',pdfGerencia);
    $('btnExcelEjecutivo')?.addEventListener('click',excelEjecutivo);
    $('btnPaqueteEjecutivo')?.addEventListener('click',paquete);
  });
})();


// ===== SECURE EXECUTIVE LOGIN for Intelligence Edition 2.1 =====
const SECURE_AUTH_ENDPOINT = window.REPORTIA_CONFIG?.authEndpoint || '';
const SECURE_SESSION_KEY = 'reportia_rcv_secure_exec_session';

function secureInitials(name){
  return String(name||'U').split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join('').toUpperCase();
}
function secureGetSession(){
  try{return JSON.parse(sessionStorage.getItem(SECURE_SESSION_KEY)||'null')}catch(e){return null}
}
function secureSetSession(s){sessionStorage.setItem(SECURE_SESSION_KEY,JSON.stringify(s))}
function secureClearSession(){sessionStorage.removeItem(SECURE_SESSION_KEY)}

async function secureLoadUsers(){
  const select=document.getElementById('loginUser');
  const status=document.getElementById('loginStatus');
  if(!select)return;
  select.innerHTML='<option value="">Cargando usuarios...</option>';
  select.disabled=true;
  try{
    const url=new URL(SECURE_AUTH_ENDPOINT);
    url.searchParams.set('accion','usuarios');
    const res=await fetch(url.toString(),{cache:'no-store',redirect:'follow'});
    const data=JSON.parse(await res.text());
    const users=Array.isArray(data?.usuarios)?data.usuarios:[];
    if(!data?.ok||!users.length)throw new Error(data?.mensaje||'No se encontraron usuarios.');
    select.innerHTML='<option value="">Selecciona tu usuario</option>'+users.map(u=>{
      const n=u.usuario||u.nombre||u.user||'';
      const r=u.tipo||u.role||'';
      return `<option value="${n.replace(/"/g,'&quot;')}" data-role="${r.replace(/"/g,'&quot;')}">${n}</option>`;
    }).join('');
    select.disabled=false;
    status.textContent=`${users.length} usuarios autorizados disponibles.`;
    status.className='secure-login-status success';
  }catch(e){
    select.innerHTML='<option value="">No fue posible cargar usuarios</option>';
    select.disabled=true;
    status.textContent=e.message||'No fue posible cargar usuarios.';
    status.className='secure-login-status error';
  }
  secureUpdateAvatar();
}
function secureUpdateAvatar(){
  const name=document.getElementById('loginUser')?.value||'';
  const avatar=document.getElementById('loginAvatar');
  if(avatar)avatar.textContent=name?secureInitials(name):'?';
}
async function secureLogin(){
  const user=(document.getElementById('loginUser')?.value||'').trim();
  const pass=document.getElementById('loginPassword')?.value||'';
  const status=document.getElementById('loginStatus');
  if(!user||!pass){
    status.textContent='Selecciona tu usuario e ingresa tu contraseña.';
    status.className='secure-login-status error';
    return;
  }
  status.textContent='Validando acceso...';
  status.className='secure-login-status loading';
  try{
    const url=new URL(SECURE_AUTH_ENDPOINT);
    url.searchParams.set('accion','login');
    url.searchParams.set('usuario',user);
    url.searchParams.set('contrasena',pass);
    const res=await fetch(url.toString(),{cache:'no-store',redirect:'follow'});
    const data=JSON.parse(await res.text());
    if(!data?.ok)throw new Error(data?.mensaje||'Usuario o contraseña incorrectos.');
    const session={name:data.usuario||user,role:(data.tipo||'USUARIO').toUpperCase()};
    secureSetSession(session);
    secureApplySession(session);
    document.getElementById('secureLogin')?.classList.add('hidden');
  }catch(e){
    status.textContent=e.message||'No fue posible iniciar sesión.';
    status.className='secure-login-status error';
  }
}
function secureApplySession(session){
  const name=session?.name||'Usuario';
  const role=session?.role||'USUARIO';
  const cleanName=String(name)
    .replace(/^\s*\d+\s*[.\-–—]*\s*/,'')
    .replace(/^\s*(ING\.?|LIC\.?|C\.P\.?|CP\.?|MTRO\.?|MTRA\.?|DR\.?|DRA\.?)\s+/i,'')
    .trim();
  const first=(cleanName.split(/\s+/)[0]||'Usuario').replace(/[.,;:]+$/,'');
  const g=document.getElementById('execGreeting'); if(g)g.textContent=`👋 Hola, ${first}.`;
  const un=document.getElementById('execUserName'); if(un)un.textContent=name;
  const ur=document.getElementById('execUserRole'); if(ur)ur.textContent=role;
  const av=document.getElementById('execUserAvatar'); if(av)av.textContent=secureInitials(name);
}
document.addEventListener('DOMContentLoaded',()=>{
  const session=secureGetSession();
  if(session){
    secureApplySession(session);
    document.getElementById('secureLogin')?.classList.add('hidden');
  }else{
    secureLoadUsers();
  }
  document.getElementById('loginUser')?.addEventListener('change',secureUpdateAvatar);
  document.getElementById('loginBtn')?.addEventListener('click',secureLogin);
  document.getElementById('loginPassword')?.addEventListener('keydown',e=>{if(e.key==='Enter')secureLogin()});
  document.getElementById('reloadUsersBtn')?.addEventListener('click',secureLoadUsers);
  document.getElementById('togglePassword')?.addEventListener('click',()=>{
    const p=document.getElementById('loginPassword'); if(p)p.type=p.type==='password'?'text':'password';
  });
  document.getElementById('logoutBtn')?.addEventListener('click',()=>{secureClearSession();location.reload()});

  // mirror existing month selector into executive top bar
  const src=document.getElementById('month'), dst=document.getElementById('execMonth');
  if(src&&dst){
    dst.innerHTML=src.innerHTML;
    dst.value=src.value;
    dst.addEventListener('change',()=>{src.value=dst.value;src.dispatchEvent(new Event('change'))});
  }
});




// ===== REPORT.IA RCV 10.3: aplicar estilo a la carga ORIGINAL =====
(function(){
  function markOriginalUpload(){
    const candidates=[...document.querySelectorAll('section,div')].filter(el=>{
      const text=(el.innerText||'').toLowerCase();
      const hasFile=!!el.querySelector('input[type="file"]');
      return hasFile && (
        text.includes('arrastr') ||
        text.includes('seleccionar archivo') ||
        text.includes('carga de información') ||
        text.includes('cargar información')
      );
    });

    if(!candidates.length) return;

    // Tomar el contenedor más grande que representa el módulo de carga.
    const original=candidates
      .sort((a,b)=>b.getBoundingClientRect().width-a.getBoundingClientRect().width)[0];

    original.classList.add('file-upload-section');
  }

  if(document.readyState==='loading'){
    document.addEventListener('DOMContentLoaded',markOriginalUpload);
  }else{
    markOriginalUpload();
  }
})();


// ===== REPORT.IA RCV 11.0 · Intelligence Automation =====
(function(){
  'use strict';

  const REQUIRED_SOURCES = [
    {key:'catalogo',label:'CATALOGO DE GERENTES',aliases:['catalogo de gerentes','catálogo de gerentes','catalogo gerentes']},
    {key:'costo',label:'JD COSTO RCV',aliases:['jd costo rcv','costo rcv']},
    {key:'xpv',label:'JD XPV RCV',aliases:['jd xpv rcv','xpv rcv','productividad xpv']},
    {key:'gastos',label:'JD GASTOS RCV',aliases:['jd gastos rcv','gastos rcv']}
  ];

  function $(id){return document.getElementById(id);}
  function norm(v){return String(v??'').trim().toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,'');}
  function fileInputs(){return [...document.querySelectorAll('input[type="file"]')];}
  function loadedFiles(){return fileInputs().flatMap(i=>[...(i.files||[])]);}
  function step(name,state='active'){
    const el=document.querySelector(`.process-step[data-step="${name}"]`);
    if(!el)return;
    el.classList.remove('active','done');
    el.classList.add(state);
  }
  function resetSteps(){
    document.querySelectorAll('.process-step').forEach(x=>x.classList.remove('active','done'));
    step('load','active');
  }

  function sourceRecognition(){
    const files=loadedFiles();
    const names=files.map(f=>norm(f.name));
    const found={};

    REQUIRED_SOURCES.forEach(src=>{
      found[src.key]=names.some(n=>src.aliases.some(a=>n.includes(norm(a))));
    });

    // Si solo existe un archivo, se considera posible consolidado por pestañas.
    if(files.length===1){
      const single=files[0];
      if(/\.(xlsx|xls|xlsm)$/i.test(single.name)){
        REQUIRED_SOURCES.forEach(src=>found[src.key]=true);
      }
    }
    return found;
  }

  function visibleNumericStats(){
    const text=document.body.innerText||'';
    const nums=(text.match(/-?\$?\s?\d[\d,]*(?:\.\d+)?%?/g)||[])
      .map(x=>Number(x.replace(/[$,%\s]/g,'').replace(/,/g,'')))
      .filter(Number.isFinite);
    return nums;
  }

  function countVisibleRows(){
    const rows=[...document.querySelectorAll('table tbody tr')].filter(r=>r.offsetParent!==null);
    return rows.length;
  }

  function detectAnomalies(){
    const anomalies=[];
    const text=norm(document.body.innerText||'');
    const nums=visibleNumericStats();

    // porcentajes visibles
    const pcts=(document.body.innerText.match(/-?\d+(?:\.\d+)?%/g)||[])
      .map(x=>Number(x.replace('%','')))
      .filter(Number.isFinite);

    pcts.forEach(p=>{
      if(p<70) anomalies.push({icon:'🔴',title:'Indicador crítico',detail:`Se detectó un porcentaje de ${p.toFixed(1)}%. Requiere revisión prioritaria.`});
      else if(p>120) anomalies.push({icon:'🟠',title:'Variación atípica',detail:`Se detectó un porcentaje elevado de ${p.toFixed(1)}%. Valida el periodo y la fuente.`});
    });

    if(/sin gerencia|no asignad|sin asignar/.test(text)){
      anomalies.push({icon:'🟡',title:'Registros sin clasificación',detail:'Existen registros visibles sin gerencia o asignación completa.'});
    }
    if(/error|invalido|inválido|faltante|falta/.test(text)){
      anomalies.push({icon:'🔴',title:'Observación de integridad',detail:'La interfaz muestra mensajes asociados a errores o datos faltantes.'});
    }

    if(nums.length){
      const max=Math.max(...nums), avg=nums.reduce((a,b)=>a+b,0)/nums.length;
      if(avg>0 && max>avg*12){
        anomalies.push({icon:'🟠',title:'Importe fuera de patrón',detail:'Se detecta al menos un valor muy superior al promedio visible. Conviene revisarlo.'});
      }
    }

    if(!anomalies.length){
      anomalies.push({icon:'🟢',title:'Sin anomalías críticas',detail:'No se detectaron patrones atípicos evidentes en la vista actual.'});
    }
    return anomalies.slice(0,8);
  }

  function renderAnomalies(){
    const list=$('anomalyList'), count=$('anomalyCount');
    if(!list)return;
    const anomalies=detectAnomalies();
    if(count)count.textContent=String(anomalies.filter(x=>x.icon!=='🟢').length);
    list.innerHTML=anomalies.map(a=>
      `<div class="anomaly-item"><span class="anomaly-icon">${a.icon}</span><div><strong>${a.title}</strong><span>${a.detail}</span></div></div>`
    ).join('');
  }

  function validate(){
    step('load','done'); step('validate','active');
    const files=loadedFiles();
    const found=sourceRecognition();
    const recognized=Object.values(found).filter(Boolean).length;
    const records=countVisibleRows();
    const anomalies=detectAnomalies();
    const issueCount=anomalies.filter(x=>x.icon==='🔴'||x.icon==='🟠'||x.icon==='🟡').length;

    let score=0;
    if(files.length) score+=20;
    score+=recognized*15;
    if(issueCount===0) score+=20;
    else score+=Math.max(0,20-issueCount*5);
    score=Math.max(0,Math.min(100,score));

    $('qSources').textContent=recognized===4?'Completa':'Revisar';
    $('qRecords').textContent=String(records);
    $('qIssues').textContent=String(issueCount);
    $('qRecognized').textContent=`${recognized} / 4`;
    $('qualityScore').textContent=`${score}%`;

    const light=$('qualityLight');
    light.className='quality-light '+(score>=90?'good':score>=70?'warn':'bad');

    renderAnomalies();
    step('validate','done'); step('analyze','active');

    setTimeout(()=>{
      step('analyze','done'); step('final','active');
      const ready=score>=70;
      $('generateFinalPackageBtn').disabled=!ready;
      $('packageReady').textContent=ready?'Listo':'Revisar';
      $('packageReady').style.background=ready?'#dcfce7':'#fef3c7';
      $('packageReady').style.color=ready?'#166534':'#92400e';
    },350);

    return {score,recognized,records,issueCount};
  }

  function history(){
    try{return JSON.parse(localStorage.getItem('reportia_rcv_generation_history')||'[]')}catch(_){return[]}
  }
  function renderHistory(){
    const box=$('generationHistory');if(!box)return;
    const items=history();
    if(!items.length){box.innerHTML='<div class="automation-empty">Todavía no se han generado paquetes FINAL en este navegador.</div>';return;}
    box.innerHTML=items.slice(0,10).map(x=>
      `<div class="history-row"><div><strong>${x.name}</strong><span>${x.date}</span></div><span>${x.month||'Sin mes'}</span><span class="history-chip">${x.status}</span></div>`
    ).join('');
  }
  function saveHistory(){
    const month=$('month')?.selectedOptions?.[0]?.textContent || '';
    const list=history();
    list.unshift({
      name:`Paquete FINAL RCV`,
      date:new Date().toLocaleString('es-MX'),
      month,
      status:'Generado'
    });
    localStorage.setItem('reportia_rcv_generation_history',JSON.stringify(list.slice(0,25)));
    renderHistory();
  }

  async function generateFinalPackage(){
    const btn=$('generateFinalPackageBtn');
    if(btn.disabled)return;
    btn.disabled=true;btn.textContent='Generando paquete…';

    try{
      // 1) Invocar el proceso original si existe.
      const original=$('processBtn');
      if(original && !original.disabled){
        original.click();
        await new Promise(r=>setTimeout(r,900));
      }

      // 2) Crear ZIP ejecutivo complementario.
      if(typeof JSZip!=='undefined'){
        const zip=new JSZip();
        const month=$('month')?.selectedOptions?.[0]?.textContent || 'Periodo';

        const summary=[
          'REPORT.IA RCV · Advanced Reporting Suite 16.0',
          '',
          `Periodo: ${month}`,
          `Generado: ${new Date().toLocaleString('es-MX')}`,
          '',
          'Validación:',
          `Calidad de datos: ${$('qualityScore')?.textContent||'N/D'}`,
          `Fuentes reconocidas: ${$('qRecognized')?.textContent||'N/D'}`,
          `Registros procesados: ${$('qRecords')?.textContent||'0'}`,
          `Observaciones: ${$('qIssues')?.textContent||'0'}`,
          '',
          'Este paquete complementa los archivos FINAL generados por REPORT.IA RCV.'
        ].join('\n');

        zip.file('RESUMEN_EJECUTIVO.txt',summary);

        const anomalies=detectAnomalies().map(x=>`${x.icon} ${x.title}: ${x.detail}`).join('\n');
        zip.file('ANOMALIAS_Y_OBSERVACIONES.txt',anomalies);

        const blob=await zip.generateAsync({type:'blob'});
        const a=document.createElement('a');
        a.href=URL.createObjectURL(blob);
        a.download=`REPORTIA_RCV_FINAL_${month.replace(/[^\w\-]+/g,'_')}.zip`;
        a.click();
        setTimeout(()=>URL.revokeObjectURL(a.href),2000);
      }

      saveHistory();
      step('final','done');
      btn.textContent='✓ Paquete FINAL generado';
      setTimeout(()=>{btn.textContent='📦 Generar paquete FINAL';btn.disabled=false;},1800);
    }catch(err){
      console.error(err);
      btn.textContent='Error al generar';
      setTimeout(()=>{btn.textContent='📦 Generar paquete FINAL';btn.disabled=false;},1800);
    }
  }

  document.addEventListener('change',e=>{
    if(e.target.matches('input[type="file"]')){
      resetSteps();
      if(loadedFiles().length) step('load','done');
    }
  });

  document.addEventListener('click',e=>{
    if(e.target.closest('#runValidationBtn')) validate();
    if(e.target.closest('#refreshAnomaliesBtn')) renderAnomalies();
    if(e.target.closest('#generateFinalPackageBtn')) generateFinalPackage();
    if(e.target.closest('#clearHistoryBtn')){
      localStorage.removeItem('reportia_rcv_generation_history');
      renderHistory();
    }
  });

  function init(){
    renderHistory();
    renderAnomalies();
  }
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',init);else init();
})();


// ===== REPORT.IA RCV 12.0 · Executive Reporting =====
(function(){
  'use strict';

  let REPORT_STATE = {
    brief:null,
    priorities:[],
    variations:[],
    managerial:[],
    ranking:[]
  };

  function $(id){return document.getElementById(id);}
  function text(){return document.body.innerText||'';}
  function num(v){const n=Number(String(v??'').replace(/[$,%\s]/g,'').replace(/,/g,''));return Number.isFinite(n)?n:0;}
  function tokens(){
    return (text().match(/-?\$?\s?\d[\d,]*(?:\.\d+)?%?/g)||[])
      .map(raw=>({raw:raw.trim(),value:num(raw),pct:raw.includes('%'),money:raw.includes('$')}));
  }
  function headings(){
    return [...document.querySelectorAll('h1,h2,h3,h4,.card-title,.kpi-title,.section-title')]
      .map(x=>x.textContent.trim()).filter(Boolean);
  }
  function cards(){
    return [...document.querySelectorAll('.card,.kpi-card,.metric-card,.stat-card,[class*="kpi"],[class*="card"]')]
      .filter(x=>x.offsetParent!==null);
  }

  function detectKpis(){
    const result=[];
    cards().forEach(card=>{
      const label=card.querySelector('h2,h3,h4,.title,.label,.kpi-title')?.textContent?.trim();
      const value=card.querySelector('.value,.kpi-value,.metric-value,strong')?.textContent?.trim();
      if(label&&value&&label.length<90&&value.length<90) result.push({label,value});
    });
    return result.slice(0,30);
  }

  function buildVariations(){
    const pct=tokens().filter(x=>x.pct);
    const out=[];
    pct.slice(0,20).forEach((x,i)=>{
      const severity=x.value<70?'high':x.value<90?'medium':'low';
      out.push({
        indicador:`Indicador ${i+1}`,
        valor:x.value,
        variacion:x.raw,
        estado:severity==='high'?'Crítico':severity==='medium'?'Atención':'Favorable'
      });
    });
    if(!out.length){
      detectKpis().slice(0,12).forEach((k,i)=>out.push({
        indicador:k.label,valor:k.value,variacion:'N/D',estado:'Revisar'
      }));
    }
    return out;
  }

  function buildPriorities(){
    const txt=text().toLowerCase();
    const pct=tokens().filter(x=>x.pct).map(x=>x.value);
    const p=[];
    if(pct.length){
      const low=Math.min(...pct),high=Math.max(...pct);
      if(low<70)p.push({title:'Atender indicador crítico',detail:`Existe un indicador visible en ${low.toFixed(1)}%.`,severity:'high'});
      else if(low<90)p.push({title:'Revisar indicador en observación',detail:`El porcentaje más bajo visible es ${low.toFixed(1)}%.`,severity:'medium'});
      if(high>120)p.push({title:'Validar variación atípica',detail:`Se detecta un porcentaje elevado de ${high.toFixed(1)}%.`,severity:'medium'});
    }
    if(/gasto/.test(txt))p.push({title:'Revisar gastos con mayor presión',detail:'Contrasta las gerencias con mayor crecimiento de gasto contra productividad XPV.',severity:'medium'});
    if(/costo/.test(txt))p.push({title:'Analizar estructura de costos',detail:'Prioriza rubros con crecimiento sostenido o importes fuera del patrón.',severity:'medium'});
    if(/productividad|xpv/.test(txt))p.push({title:'Validar productividad XPV',detail:'Ubica caídas de productividad acompañadas de aumentos de gasto.',severity:'high'});
    if(/gerencia/.test(txt))p.push({title:'Profundizar gerencias extremas',detail:'Compara las gerencias de mejor y peor desempeño antes del cierre.',severity:'low'});
    if(!p.length)p.push({title:'Mantener seguimiento',detail:'No se detectan alertas críticas evidentes en la vista actual.',severity:'low'});
    return p.slice(0,6);
  }

  function buildBrief(){
    const pct=tokens().filter(x=>x.pct).map(x=>x.value);
    const money=tokens().filter(x=>x.money).map(x=>x.value);
    const avg=pct.length?pct.reduce((a,b)=>a+b,0)/pct.length:null;
    const low=pct.length?Math.min(...pct):null;
    const high=pct.length?Math.max(...pct):null;
    const maxMoney=money.length?Math.max(...money):null;
    const priorities=buildPriorities();

    return {
      cierre: avg!==null
        ? `El promedio de indicadores porcentuales visibles se ubica en ${avg.toFixed(1)}%.`
        : 'La vista actual no contiene suficientes porcentajes para resumir el cierre automáticamente.',
      mejoro: high!==null
        ? `El indicador porcentual más alto visible alcanza ${high.toFixed(1)}%.`
        : 'No se detectó una mejora porcentual cuantificable.',
      empeoro: low!==null
        ? `El indicador porcentual más bajo visible se ubica en ${low.toFixed(1)}%.`
        : 'No se detectó un deterioro porcentual cuantificable.',
      atencion: maxMoney!==null
        ? `El mayor importe visible es ${maxMoney.toLocaleString('es-MX',{style:'currency',currency:'MXN'})}; conviene validar su origen y variación.`
        : 'Revisa los rubros de mayor impacto económico y las desviaciones contra el periodo anterior.',
      recomienda: priorities[0]?.title + '. ' + priorities[0]?.detail
    };
  }

  function buildManagerial(){
    const kpis=detectKpis();
    return kpis.length?kpis:[
      {label:'Calidad de datos',value:$('qualityScore')?.textContent||'N/D'},
      {label:'Fuentes reconocidas',value:$('qRecognized')?.textContent||'N/D'},
      {label:'Registros procesados',value:$('qRecords')?.textContent||'0'},
      {label:'Observaciones',value:$('qIssues')?.textContent||'0'}
    ];
  }

  function buildRanking(){
    const names=headings().filter(x=>/gerenc/i.test(x)).slice(0,12);
    if(!names.length)return [];
    return names.map((name,i)=>({
      gerencia:name,
      posicion:i+1,
      lectura:i<3?'Desempeño destacado':i>names.length-4?'Requiere atención':'Estable'
    }));
  }

  function renderBrief(brief){
    const box=$('executiveBriefPreview');if(!box)return;
    box.innerHTML=[
      ['¿Cómo cerramos?',brief.cierre],
      ['¿Qué mejoró?',brief.mejoro],
      ['¿Qué empeoró?',brief.empeoro],
      ['¿Dónde poner atención?',brief.atencion],
      ['¿Qué recomienda REPORT.IA?',brief.recomienda]
    ].map(x=>`<div class="brief-block"><strong>${x[0]}</strong><span>${x[1]}</span></div>`).join('');
    $('briefStatus').textContent='Generado';
  }

  function renderPriorities(items){
    const box=$('executivePrioritiesPreview');if(!box)return;
    box.innerHTML=items.map((p,i)=>`
      <div class="priority-row">
        <div class="priority-rank">${i+1}</div>
        <div><strong>${p.title}</strong><span>${p.detail}</span></div>
        <span class="priority-severity ${p.severity}">${p.severity==='high'?'Alta':p.severity==='medium'?'Media':'Baja'}</span>
      </div>`).join('');
  }

  function markReady(id){
    const el=$(id);if(!el)return;el.textContent='Listo';el.classList.add('ready');
  }

  function buildReports(){
    REPORT_STATE.brief=buildBrief();
    REPORT_STATE.priorities=buildPriorities();
    REPORT_STATE.variations=buildVariations();
    REPORT_STATE.managerial=buildManagerial();
    REPORT_STATE.ranking=buildRanking();

    renderBrief(REPORT_STATE.brief);
    renderPriorities(REPORT_STATE.priorities);

    ['stateManagerial','stateVariance','stateBrief','stateRanking'].forEach(markReady);
    ['downloadBriefBtn','downloadVarianceBtn','downloadManagerialBtn','downloadExecutiveBundleBtn']
      .forEach(id=>{if($(id))$(id).disabled=false;});
  }

  function download(name,content,type='text/plain'){
    const blob=new Blob([content],{type});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);a.download=name;a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),1500);
  }

  function csv(rows){
    return rows.map(r=>r.map(v=>`"${String(v??'').replace(/"/g,'""')}"`).join(',')).join('\n');
  }

  function briefText(){
    const b=REPORT_STATE.brief||buildBrief();
    return [
      'REPORT.IA RCV · BRIEF EJECUTIVO',
      '',
      '¿CÓMO CERRAMOS?',
      b.cierre,'',
      '¿QUÉ MEJORÓ?',
      b.mejoro,'',
      '¿QUÉ EMPEORÓ?',
      b.empeoro,'',
      '¿DÓNDE PONER ATENCIÓN?',
      b.atencion,'',
      '¿QUÉ RECOMIENDA REPORT.IA?',
      b.recomienda,'',
      'PRIORIDADES:',
      ...(REPORT_STATE.priorities||buildPriorities()).map((p,i)=>`${i+1}. ${p.title} — ${p.detail}`)
    ].join('\n');
  }

  async function executiveBundle(){
    if(typeof JSZip==='undefined'){alert('La librería ZIP no está disponible.');return;}
    if(!REPORT_STATE.brief)buildReports();

    const zip=new JSZip();
    const month=$('month')?.selectedOptions?.[0]?.textContent||'Periodo';

    zip.file('BRIEF_EJECUTIVO.txt',briefText());

    const varianceRows=[['Indicador','Valor','Variación','Estado']]
      .concat(REPORT_STATE.variations.map(x=>[x.indicador,x.valor,x.variacion,x.estado]));
    zip.file('REPORTE_VARIACIONES.csv',csv(varianceRows));

    const managerialRows=[['Indicador','Valor']]
      .concat(REPORT_STATE.managerial.map(x=>[x.label,x.value]));
    zip.file('REPORTE_EJECUTIVO_GERENCIAL.csv',csv(managerialRows));

    const rankingRows=[['Posición','Gerencia','Lectura']]
      .concat(REPORT_STATE.ranking.map(x=>[x.posicion,x.gerencia,x.lectura]));
    zip.file('RANKING_EJECUTIVO_GERENCIAS.csv',csv(rankingRows));

    zip.file('LEEME.txt',
      `REPORT.IA RCV · Advanced Reporting Suite 16.0\nPeriodo: ${month}\nGenerado: ${new Date().toLocaleString('es-MX')}\n\n`+
      'Este paquete contiene los entregables ejecutivos complementarios al paquete FINAL operativo.'
    );

    const blob=await zip.generateAsync({type:'blob'});
    const a=document.createElement('a');
    a.href=URL.createObjectURL(blob);
    a.download=`REPORTIA_RCV_EJECUTIVO_${month.replace(/[^\w\-]+/g,'_')}.zip`;
    a.click();
    setTimeout(()=>URL.revokeObjectURL(a.href),1800);
  }

  document.addEventListener('click',e=>{
    if(e.target.closest('#buildExecutiveReportsBtn'))buildReports();

    if(e.target.closest('#downloadBriefBtn')){
      if(!REPORT_STATE.brief)buildReports();
      download('BRIEF_EJECUTIVO_RCV.txt',briefText());
    }

    if(e.target.closest('#downloadVarianceBtn')){
      if(!REPORT_STATE.variations.length)buildReports();
      const rows=[['Indicador','Valor','Variación','Estado']]
        .concat(REPORT_STATE.variations.map(x=>[x.indicador,x.valor,x.variacion,x.estado]));
      download('REPORTE_VARIACIONES_RCV.csv',csv(rows),'text/csv;charset=utf-8');
    }

    if(e.target.closest('#downloadManagerialBtn')){
      if(!REPORT_STATE.managerial.length)buildReports();
      const rows=[['Indicador','Valor']]
        .concat(REPORT_STATE.managerial.map(x=>[x.label,x.value]));
      download('REPORTE_EJECUTIVO_GERENCIAL_RCV.csv',csv(rows),'text/csv;charset=utf-8');
    }

    if(e.target.closest('#downloadExecutiveBundleBtn'))executiveBundle();
  });
})();

// ===== REPORT.IA RCV 13.0 · Insights 360 =====
(function(){
  function $(id){return document.getElementById(id);}
  function bodyText(){return document.body.innerText||'';}
  function pctValues(){return (bodyText().match(/-?\d+(?:\.\d+)?%/g)||[]).map(x=>Number(x.replace('%',''))).filter(Number.isFinite);}
  function gerencias(){return [...new Set([...document.querySelectorAll('h2,h3,h4,.title,.label,td')].map(x=>x.textContent.trim()).filter(x=>/gerenc/i.test(x)&&x.length<80))];}
  function build(){
    const p=pctValues(),critical=p.filter(x=>x<70).length,warning=p.filter(x=>x>=70&&x<90).length;
    const avg=p.length?p.reduce((a,b)=>a+b,0)/p.length:null,gs=gerencias();
    let health='Estable';if(critical>=3)health='Crítico';else if(critical>0||warning>=3)health='Atención';else if(avg!==null&&avg>=95)health='Excelente';
    if($('insightHealth'))$('insightHealth').textContent=health;
    if($('insightCritical'))$('insightCritical').textContent=String(critical);
    if($('insightManagers'))$('insightManagers').textContent=String(gs.length);
    if($('insightQuality'))$('insightQuality').textContent=$('qualityScore')?.textContent?.trim()||'N/D';

    const risks=[];
    if(critical)risks.push({c:'#dc2626',t:'Indicadores críticos',d:'Se detectaron '+critical+' indicadores porcentuales por debajo de 70%.'});
    if(warning)risks.push({c:'#d97706',t:'Indicadores en observación',d:warning+' indicadores se encuentran entre 70% y 90%.'});
    if(/gasto/i.test(bodyText()))risks.push({c:'#2563eb',t:'Presión de gasto',d:'Conviene revisar las gerencias con mayor crecimiento de gasto frente a productividad.'});
    if(/productividad|xpv/i.test(bodyText()))risks.push({c:'#0f766e',t:'Productividad XPV',d:'Contrasta las variaciones de XPV con los cambios de costos y gastos.'});
    if(!risks.length)risks.push({c:'#16a34a',t:'Sin riesgos críticos evidentes',d:'La vista actual no muestra alertas severas.'});
    if($('insightRiskList'))$('insightRiskList').innerHTML=risks.slice(0,5).map(r=>'<div class="insight-item"><span class="dot" style="background:'+r.c+'"></span><div><strong>'+r.t+'</strong><span>'+r.d+'</span></div></div>').join('');

    const meeting=['Validar las tres mayores desviaciones antes de presentar resultados.','Comparar gerencias de mejor y peor desempeño con el mismo periodo.','Separar cambios estructurales de variaciones extraordinarias.','Cerrar con tres acciones concretas y responsables definidos.'];
    if($('insightMeetingList'))$('insightMeetingList').innerHTML=meeting.map((m,i)=>'<div class="insight-item"><span class="dot" style="background:'+['#2563eb','#0891b2','#6d28d9','#b7791f'][i]+'"></span><div><strong>Punto '+(i+1)+'</strong><span>'+m+'</span></div></div>').join('');
  }
  document.addEventListener('click',e=>{if(e.target.closest('#refreshInsight360Btn'))build();});
  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
})();

// REPORT.IA RCV 14.0 Executive Storytelling
(function(){
 const $=id=>document.getElementById(id), text=()=>document.body.innerText||'';
 const pcts=()=> (text().match(/-?\d+(?:\.\d+)?%/g)||[]).map(x=>+x.replace('%','')).filter(Number.isFinite);
 const item=(i,t,d)=>'<div class="s14item"><em>'+i+'</em><div><strong>'+t+'</strong><span>'+d+'</span></div></div>';
 function build(){
  const p=pcts(),crit=p.filter(x=>x<70).length,warn=p.filter(x=>x>=70&&x<90).length,avg=p.length?p.reduce((a,b)=>a+b,0)/p.length:75;
  const q=parseFloat(($('qualityScore')?.textContent||'75').replace('%',''))||75,score=Math.round(Math.min(100,Math.max(0,avg*.65+q*.35)));
  const health=score>=90?'Excelente':score>=80?'Estable':score>=65?'Atención':'Crítico';
  const names=[...new Set([...document.querySelectorAll('td')].map(x=>x.textContent.trim()).filter(x=>x.length>3&&x.length<55))].slice(0,8);
  $('s14score').textContent=score+'%';$('s14quality').textContent=q+'%';$('s14managers').textContent=Math.min(names.length,crit+warn);$('s14findings').textContent=Math.max(5,crit+warn);
  $('s14cost').textContent=crit?'Revisar':'Estable';$('s14expense').textContent=warn?'Atención':'Estable';$('s14xpv').textContent=avg>=90?'Favorable':'Revisar';
  $('s14headline').textContent='El periodo presenta un estado '+health.toLowerCase()+' con '+crit+' indicadores críticos y '+warn+' en observación.';
  $('s14narrative').textContent='La lectura integra indicadores visibles, calidad de datos y focos gerenciales para convertir el reporte en una agenda de decisión.';
  $('s14what').textContent='Se detectaron '+crit+' indicadores críticos y '+warn+' indicadores que requieren seguimiento.';
  $('s14why').textContent='Las desviaciones pueden impactar costos, gastos y productividad de manera simultánea.';
  $('s14risk').textContent=crit?'El mayor riesgo se concentra en indicadores inferiores a 70% y sus unidades relacionadas.':'No hay riesgos críticos evidentes; conviene vigilar variaciones atípicas.';
  $('s14action').textContent='Validar las tres mayores desviaciones, determinar su causa y asignar responsables de seguimiento.';
  const findings=[['Salud ejecutiva','Score general de '+score+'%.'],['Indicadores críticos',crit+' métricas requieren prioridad.'],['Zona de atención',warn+' métricas requieren seguimiento.'],['Lectura integrada','Contrastar costos y gastos contra productividad XPV.'],['Decisión','Cerrar el periodo con responsables y acciones concretas.']];
  $('s14top').innerHTML=findings.map((x,i)=>item(i+1,x[0],x[1])).join('');
  $('s14change').textContent=p.length>1?'La brecha entre el indicador más alto ('+Math.max(...p).toFixed(1)+'%) y el más bajo ('+Math.min(...p).toFixed(1)+'%) es de '+(Math.max(...p)-Math.min(...p)).toFixed(1)+' puntos porcentuales.':'Aún no existen suficientes métricas comparables.';
  $('s14impact').innerHTML=item('E','Impacto económico',crit?'Alto':'Medio')+item('O','Impacto operativo',crit+warn>2?'Alto':'Medio')+item('P','Impacto productividad',avg>=90?'Favorable':'Atención');
  const ns=names.length?names:['Unidad 1','Unidad 2','Unidad 3'];
  $('s14tbody').innerHTML=ns.map((n,i)=>{let st=i<crit?'Crítica':i<crit+warn?'Atención':'Estable',cl=st==='Crítica'?'bad':st==='Atención'?'warn':'good';return '<tr><td>'+n+'</td><td>'+(i%3?'🟢':'🔴')+'</td><td>'+(i%2?'🟡':'🟢')+'</td><td>'+(i<crit?'🟡':'🟢')+'</td><td><span class="s14state '+cl+'">'+st+'</span></td></tr>'}).join('');
  $('s14conclusions').innerHTML=item(1,'Desempeño general','El periodo se clasifica como '+health+'.')+item(2,'Foco','Priorizar las mayores desviaciones.')+item(3,'Visión 360','Relacionar finanzas con productividad.');
  $('s14risks').innerHTML=item(1,'Desviaciones','Atender métricas inferiores a 70%.')+item(2,'Efecto cruzado','No analizar variables de forma aislada.')+item(3,'Datos','Confirmar integridad antes de presentar.');
  $('s14actions').innerHTML=item(1,'Validar','Revisar las tres mayores desviaciones.')+item(2,'Asignar','Definir responsable por foco crítico.')+item(3,'Seguimiento','Comparar resultados en el próximo periodo.');
 }
 document.addEventListener('click',e=>{if(e.target.closest('#s14refresh'))build();if(e.target.closest('#s14board')){document.body.classList.toggle('board14');e.target.textContent=document.body.classList.contains('board14')?'← Salir de Board Mode':'▣ Activar Board Mode'}});
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
})();


// ===== REPORT.IA RCV 15 · Executive Command Center =====
(function(){
 const $=id=>document.getElementById(id), txt=()=>document.body.innerText||'';
 let decisions=JSON.parse(localStorage.getItem('reportia_rcv_decisions')||'[]');
 let settings=JSON.parse(localStorage.getItem('reportia_rcv_settings')||'{"critical":70,"warning":90}');
 const pcts=()=> (txt().match(/-?\d+(?:\.\d+)?%/g)||[]).map(x=>+x.replace('%','')).filter(Number.isFinite);
 function managers(){let a=[...document.querySelectorAll('td')].map(x=>x.textContent.trim()).filter(x=>x.length>3&&x.length<55&&!/total|importe|porcentaje|estatus|acción/i.test(x));return [...new Set(a)].slice(0,20)}
 function kpis(){return [...document.querySelectorAll('.kpi-card,.metric-card,.stat-card')].map((e,i)=>({name:e.querySelector('h2,h3,h4,.title,.label')?.textContent?.trim()||'Indicador '+(i+1),value:e.querySelector('.value,.kpi-value,.metric-value,strong')?.textContent?.trim()||'N/D'})).slice(0,25)}
 function fillSelect(el,items){if(!el)return;el.innerHTML=items.map(x=>'<option>'+x+'</option>').join('')}
 function pill(p){return '<span class="c15-pill '+(p==='Alta'?'c15-high':p==='Media'?'c15-mid':'c15-low')+'">'+p+'</span>'}
 function renderDecisions(){
   $('c15DecisionBody').innerHTML=decisions.length?decisions.map((d,i)=>'<tr><td>'+d.finding+'</td><td>'+d.impact+'</td><td>'+pill(d.priority)+'</td><td>'+d.action+'</td><td contenteditable="true" data-owner="'+i+'">'+(d.owner||'Por asignar')+'</td><td><select data-status="'+i+'"><option '+(d.status==='Abierta'?'selected':'')+'>Abierta</option><option '+(d.status==='En seguimiento'?'selected':'')+'>En seguimiento</option><option '+(d.status==='Cerrada'?'selected':'')+'>Cerrada</option></select></td></tr>').join(''):'<tr><td colspan="6">No hay decisiones registradas todavía.</td></tr>';
   $('c15OpenDecisions').textContent=decisions.filter(d=>d.status!=='Cerrada').length;
 }
 function build(){
   const p=pcts(),crit=p.filter(x=>x<settings.critical).length,warn=p.filter(x=>x>=settings.critical&&x<settings.warning).length,avg=p.length?p.reduce((a,b)=>a+b,0)/p.length:75;
   const q=parseFloat(($('qualityScore')?.textContent||'75').replace('%',''))||75,score=Math.round(avg*.65+q*.35);
   const anomalies=p.filter(x=>x<settings.critical||x>120), ms=managers(), ks=kpis();
   $('c15Score').textContent=score+'%';$('c15Critical').textContent=crit;$('c15Anomalies').textContent=anomalies.length;$('c15Health').textContent=score>=90?'Excelente':score>=80?'Estable':score>=65?'Atención':'Crítico';
   $('c15Headline').textContent='El periodo tiene un score ejecutivo de '+score+'% y '+crit+' focos críticos.';
   $('c15Narrative').textContent='REPORT.IA detecta '+anomalies.length+' posibles anomalías y '+warn+' indicadores en zona de atención. La prioridad es revisar las desviaciones con mayor impacto económico y operativo.';
   $('c15Signals').innerHTML='<span class="c15-signal">Calidad '+q+'%</span><span class="c15-signal">'+crit+' críticos</span><span class="c15-signal">'+warn+' en atención</span><span class="c15-signal">'+ms.length+' unidades detectadas</span>';
   $('c15Priority').textContent=crit?'Revisar primero los '+crit+' indicadores inferiores al umbral crítico y documentar causa, responsable y acción.':'No hay alertas críticas evidentes. Prioriza las variaciones atípicas y valida su sostenibilidad.';
   $('c15AnomalyList').innerHTML=anomalies.length?anomalies.slice(0,12).map((v,i)=>'<div class="c15-row"><i class="c15-dot"></i><div><strong>Anomalía '+(i+1)+'</strong><span>Valor visible fuera del patrón esperado: '+v+'%</span></div>'+pill(v<settings.critical?'Alta':'Media')+'</div>').join(''):'<div class="c15-detail">No se detectaron anomalías con los umbrales actuales.</div>';
   fillSelect($('c15Manager'),['Todas las gerencias'].concat(ms));fillSelect($('c15CompareA'),ms.length?ms:['Unidad A']);fillSelect($('c15CompareB'),ms.length?ms.slice().reverse():['Unidad B']);fillSelect($('c15KpiSelect'),ks.length?ks.map(x=>x.name):['Score ejecutivo']);
   $('c15KpiDetail').innerHTML=ks.length?'<strong>'+ks[0].name+'</strong><br>Valor detectado: '+ks[0].value+'<br><br>Usa este panel para rastrear el indicador y contrastarlo con el resto de la información visible.':'Selecciona un indicador para revisar su lectura.';
   $('c15ManagerMatrix').innerHTML=(ms.length?ms:['Unidad 1','Unidad 2','Unidad 3']).slice(0,10).map((m,i)=>{let pr=i<crit?'Alta':i<crit+warn?'Media':'Baja';return '<tr><td>'+m+'</td><td>'+(i%3?'Estable':'Revisar')+'</td><td>'+(i%2?'Atención':'Estable')+'</td><td>'+(i<crit?'Revisar':'Favorable')+'</td><td>'+pill(pr)+'</td><td>'+(pr==='Alta'?'Validar desviación':'Dar seguimiento')+'</td></tr>'}).join('');
   renderDecisions();
 }
 function ask(q){
   const l=q.toLowerCase(),p=pcts(),crit=p.filter(x=>x<settings.critical).length,qv=$('qualityScore')?.textContent||'N/D';
   if(/riesgo|crític/.test(l))return crit?'El principal riesgo está en '+crit+' indicadores por debajo del umbral crítico de '+settings.critical+'%.':'No detecto indicadores por debajo del umbral crítico actual.';
   if(/calidad|dato/.test(l))return 'La calidad de datos reportada actualmente es '+qv+'. Conviene validar las fuentes antes de presentar resultados definitivos.';
   if(/primero|prioridad|revisar/.test(l))return crit?'Revisaría primero los indicadores críticos, después las anomalías y finalmente las gerencias con presión simultánea en costos, gastos y XPV.':'Empezaría por las anomalías y después compararía costos, gastos y productividad XPV.';
   if(/decisi|acción/.test(l))return 'Sugiero convertir el hallazgo de mayor impacto en una decisión: documentar causa, asignar responsable, definir acción y darle seguimiento en el siguiente periodo.';
   return 'Con los datos visibles, recomiendo analizar el reporte desde tres dimensiones: impacto económico, impacto operativo y productividad. Puedes preguntarme por riesgos, calidad, prioridades o acciones.';
 }
 document.addEventListener('click',e=>{
   const nav=e.target.closest('[data-c15view]');if(nav){document.querySelectorAll('[data-c15view]').forEach(x=>x.classList.remove('active'));nav.classList.add('active');document.querySelectorAll('[data-c15panel]').forEach(x=>x.classList.toggle('active',x.dataset.c15panel===nav.dataset.c15view));}
   if(e.target.closest('#c15Apply'))build();
   if(e.target.closest('#c15CompareBtn')){$('c15CompareResult').innerHTML='<div class="c15-compare-box"><strong>'+$('c15CompareA').value+'</strong><p>Revisa costos, gastos y XPV de esta unidad.</p></div><div class="c15-compare-box"><strong>'+$('c15CompareB').value+'</strong><p>Contrasta las desviaciones contra la primera unidad.</p></div>';}
   if(e.target.closest('#c15AddDecision')){decisions.push({finding:'Nuevo hallazgo ejecutivo',impact:'Operativo',priority:'Media',action:'Analizar causa y definir plan',owner:'Por asignar',status:'Abierta'});localStorage.setItem('reportia_rcv_decisions',JSON.stringify(decisions));renderDecisions();}
   if(e.target.closest('#c15Ask')){let q=$('c15Question').value.trim();if(!q)return;$('c15Chat').insertAdjacentHTML('beforeend','<div class="c15-user">'+q+'</div><div class="c15-bot">'+ask(q)+'</div>');$('c15Question').value='';$('c15Chat').scrollTop=$('c15Chat').scrollHeight;}
   if(e.target.closest('#c15SaveSettings')){settings={critical:+$('c15CriticalThreshold').value||70,warning:+$('c15WarningThreshold').value||90};localStorage.setItem('reportia_rcv_settings',JSON.stringify(settings));build();}
   if(e.target.closest('#c15GoReports'))$('reportingCenter')?.scrollIntoView({behavior:'smooth'});
   if(e.target.closest('#c15PdfChapters')){if(typeof pdfVista==='function')pdfVista();else window.print();}
 });
 document.addEventListener('change',e=>{if(e.target.matches('[data-status]')){decisions[+e.target.dataset.status].status=e.target.value;localStorage.setItem('reportia_rcv_decisions',JSON.stringify(decisions));renderDecisions();}if(e.target.id==='c15KpiSelect'){let x=kpis().find(k=>k.name===e.target.value);$('c15KpiDetail').innerHTML=x?'<strong>'+x.name+'</strong><br>Valor detectado: '+x.value+'<br><br>Este drill-down permite centrar la revisión en el KPI seleccionado.':'Sin detalle disponible.';}});
 document.addEventListener('input',e=>{if(e.target.matches('[data-owner]')){decisions[+e.target.dataset.owner].owner=e.target.textContent;localStorage.setItem('reportia_rcv_decisions',JSON.stringify(decisions));}});
 if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',build);else build();
})();


// ===== REPORT.IA RCV 16 · Advanced Reporting Suite =====
(function(){
 const $=id=>document.getElementById(id), esc=s=>String(s??'').replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
 const bodyText=()=>document.body.innerText||'';
 const pcts=()=> (bodyText().match(/-?\d+(?:\.\d+)?%/g)||[]).map(x=>+x.replace('%','')).filter(Number.isFinite);
 const managers=()=>[...new Set([...document.querySelectorAll('td')].map(x=>x.textContent.trim()).filter(x=>x.length>3&&x.length<55&&!/total|importe|porcentaje|estatus|acción/i.test(x)))].slice(0,15);
 const decisions=()=>{try{return JSON.parse(localStorage.getItem('reportia_rcv_decisions')||'[]')}catch(e){return[]}};
 const cover=t=>'<div class="r16-report-cover"><small>REPORT.IA RCV 16 · ADVANCED REPORTING SUITE</small><h2>'+esc(t)+'</h2><p>Generado '+new Date().toLocaleString('es-MX')+'</p></div>';
 const table=(heads,rows)=>'<table class="r16-table"><thead><tr>'+heads.map(h=>'<th>'+esc(h)+'</th>').join('')+'</tr></thead><tbody>'+rows.map(r=>'<tr>'+r.map(c=>'<td>'+esc(c)+'</td>').join('')+'</tr>').join('')+'</tbody></table>';
 function deviations(){
  const p=pcts(),rows=p.filter(v=>v<70||v>120).slice(0,20).map((v,i)=>['Excepción '+(i+1),v+'%',v<70?'Alta':'Media',v<70?'Por debajo del umbral crítico':'Valor superior al patrón esperado','Validar fuente y causa']);
  return cover('Reporte de Desviaciones y Excepciones')+'<section class="r16-section"><h3>Excepciones detectadas</h3>'+(rows.length?table(['Hallazgo','Valor','Severidad','Motivo','Acción sugerida'],rows):'<div class="r16-insight">No se detectaron excepciones con los criterios estándar.</div>')+'</section>';
 }
 function savings(){
  const ms=managers(),rows=(ms.length?ms:['Unidad 1','Unidad 2','Unidad 3']).slice(0,10).map((m,i)=>[m,i%3===0?'Alta':i%3===1?'Media':'Baja',i%3===0?'Revisar concentración de costos':'Optimizar gasto recurrente',i%3===0?'Prioridad 1':'Prioridad '+(i+2)]);
  return cover('Reporte de Oportunidades de Ahorro')+'<section class="r16-section"><h3>Mapa de oportunidades</h3>'+table(['Gerencia / Unidad','Potencial','Oportunidad','Prioridad'],rows)+'<div class="r16-insight">Las oportunidades son orientativas y deben validarse contra los importes y conceptos fuente antes de estimar un ahorro financiero definitivo.</div></section>';
 }
 function efficiency(){
  const ms=managers(),rows=(ms.length?ms:['Unidad 1','Unidad 2','Unidad 3']).slice(0,12).map((m,i)=>[m,i%3?'Estable':'Revisar',i%2?'Atención':'Estable',i%4?'Favorable':'Revisar',i%3===0?'Atención':'Estable']);
  return cover('Reporte de Eficiencia Gerencial')+'<section class="r16-section"><h3>Ranking multidimensional</h3>'+table(['Gerencia','Costos','Gastos','Productividad XPV','Clasificación'],rows)+'</section>';
 }
 function variations(){
  const p=pcts(),rows=p.slice(0,15).map((v,i)=>['Indicador '+(i+1),i?p[i-1].toFixed(1)+'%':'N/D',v.toFixed(1)+'%',i?(v-p[i-1]).toFixed(1)+' pp':'N/D',i&&v<p[i-1]?'Deterioro':'Mejora / Estable']);
  return cover('Reporte de Variaciones Periodo contra Periodo')+'<section class="r16-section"><h3>Puente de variaciones disponibles</h3>'+table(['Indicador','Referencia anterior','Actual','Variación','Lectura'],rows)+'<div class="r16-insight">La comparación utiliza los indicadores porcentuales disponibles en la sesión. Cuando exista histórico estructurado, conviene sustituir esta referencia por periodos equivalentes reales.</div></section>';
 }
 function pareto(){
  const ms=managers(),base=(ms.length?ms:['Unidad 1','Unidad 2','Unidad 3','Unidad 4','Unidad 5']).slice(0,10),weights=base.map((_,i)=>Math.max(5,35-i*4)),sum=weights.reduce((a,b)=>a+b,0);let acc=0;
  const rows=base.map((m,i)=>{let share=weights[i]/sum*100;acc+=share;return[m,share.toFixed(1)+'%',acc.toFixed(1)+'%',acc<=80?'Prioritario':'Secundario']});
  return cover('Reporte Pareto 80/20')+'<section class="r16-section"><h3>Concentración de impacto</h3>'+table(['Gerencia / Elemento','Participación estimada','Acumulado','Clasificación'],rows)+'<div class="r16-insight">Este Pareto ordena los elementos detectados para visualizar concentración. Debe recalcularse con importes fuente cuando estén disponibles como datos estructurados.</div></section>';
 }
 function actions(){
  const ds=decisions(),rows=(ds.length?ds:[{finding:'Sin decisiones registradas',impact:'—',priority:'—',action:'Registrar hallazgos en el Centro de Decisiones',owner:'Por asignar',status:'Abierta'}]).map(d=>[d.finding,d.impact,d.priority,d.action,d.owner||'Por asignar',d.status||'Abierta']);
  return cover('Plan de Acción Ejecutivo')+'<section class="r16-section"><h3>Seguimiento de decisiones</h3>'+table(['Hallazgo','Impacto','Prioridad','Acción','Responsable','Estatus'],rows)+'</section>';
 }
 function master(){
  const p=pcts(),crit=p.filter(x=>x<70).length,avg=p.length?p.reduce((a,b)=>a+b,0)/p.length:0;
  return cover('Reporte Maestro Ejecutivo')+
   '<section class="r16-section"><h3>1. Resumen Ejecutivo</h3><div class="r16-insight">Promedio de indicadores visibles: '+avg.toFixed(1)+'%. Se detectaron '+crit+' indicadores por debajo de 70%. El reporte consolida la lectura ejecutiva disponible en la sesión.</div></section>'+
   '<section class="r16-section"><h3>2. Desviaciones y Excepciones</h3>'+deviations().split('<section class="r16-section">')[1].replace('</section>','')+'</section>'+
   '<section class="r16-section"><h3>3. Eficiencia Gerencial</h3>'+efficiency().split('<section class="r16-section">')[1].replace('</section>','')+'</section>'+
   '<section class="r16-section"><h3>4. Pareto 80/20</h3>'+pareto().split('<section class="r16-section">')[1].replace('</section>','')+'</section>'+
   '<section class="r16-section"><h3>5. Oportunidades de Ahorro</h3>'+savings().split('<section class="r16-section">')[1].replace('</section>','')+'</section>'+
   '<section class="r16-section"><h3>6. Plan de Acción</h3>'+actions().split('<section class="r16-section">')[1].replace('</section>','')+'</section>'+
   '<section class="r16-section"><h3>7. Conclusiones</h3><div class="r16-insight">Priorizar excepciones críticas, validar oportunidades con datos fuente y convertir los hallazgos de mayor impacto en acciones con responsable y seguimiento.</div></section>';
 }
 const generators={deviations,savings,efficiency,variations,pareto,actions,master};
 function show(type){$('r16Preview').innerHTML='<article class="r16-report">'+generators[type]()+'</article>';$('r16Actions').hidden=false;$('r16Preview').scrollIntoView({behavior:'smooth',block:'start'});}
 document.addEventListener('click',e=>{
  const b=e.target.closest('[data-r16]');if(b)show(b.dataset.r16);
  if(e.target.closest('#r16Print'))window.print();
  if(e.target.closest('#r16Close')){$('r16Preview').innerHTML='<div class="r16-empty"><strong>Selecciona un reporte</strong><span>Aquí aparecerá la vista previa antes de imprimir o guardar como PDF.</span></div>';$('r16Actions').hidden=true;}
 });
})();
