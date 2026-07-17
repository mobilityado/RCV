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
function reset(){state.files=[];state.sources={};state.outputs={};state.analysis=null;destroyCharts();$('#fileInput').value='';renderFiles();renderValidation();['dashboard','gerencias','costos','gastos','xpv','tendencias','inteligencia','comparador','prioridades','copiloto','explorador','historial','informePro','cierreEjecutivo','resultados'].forEach(id=>document.getElementById(id)?.classList.add('hidden'))}
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
    renderProEdition(costRows,expenseRows,xpvRows,cutoff);
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
    if(!window.html2canvas || !window.jspdf){ alert('No se pudieron cargar las librerías de PDF. Revisa tu conexión a internet.'); return; }
    const canvas=await html2canvas(element,{scale:1.6,useCORS:true,backgroundColor:'#ffffff',logging:false});
    const img=canvas.toDataURL('image/jpeg',0.92);
    const {jsPDF}=window.jspdf;
    const pdf=new jsPDF('p','mm','a4');
    const pw=210, ph=297, margin=10, header=18;
    pdf.setFont('helvetica','bold'); pdf.setFontSize(15); pdf.text(title,margin,11);
    pdf.setFont('helvetica','normal'); pdf.setFontSize(8); pdf.text('REPORT.IA RCV · Intelligence Edition 2.0',margin,16);
    const iw=pw-margin*2, ih=canvas.height*iw/canvas.width;
    let y=header, remaining=ih, sy=0;
    while(remaining>0){
      const pageH=ph-y-margin;
      const sliceH=Math.min(pageH,remaining);
      // Use same image with negative y clipping per page
      pdf.addImage(img,'JPEG',margin,y-sy,iw,ih);
      remaining-=sliceH; sy+=sliceH;
      if(remaining>0){ pdf.addPage(); y=margin; }
    }
    pdf.save(filename);
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


// ===== REPORT.IA RCV · PRO EDITION 4.0 =====
const PRO_HISTORY_KEY='reportia_rcv_history_v4';

function proFinancialRows(){
  const a=state.analysis||{};
  return [...(a.costRows||[]),...(a.expenseRows||[])];
}
function proManagers(){
  return [...new Set(proFinancialRows().map(r=>r.manager).filter(Boolean))].sort((a,b)=>a.localeCompare(b,'es'));
}
function proSnapshot(){
  if(!state.analysis)return null;
  const fin=proFinancialRows(), xpv=state.analysis.xpvRows||[];
  const real=sum(fin,'real26'), budget=sum(fin,'budget26'), real25=sum(fin,'real25');
  const xpvReal=sum(xpv,'real'), xpvBudget=sum(xpv,'budget');
  const managers=aggregate(fin,['manager'],['real26','budget26']).map(addVariance);
  const favorable=managers.filter(x=>x.varBudget<=0).length;
  const critical=managers.filter(x=>(x.varBudgetPct||0)>.10).length;
  const budgetVar=budget?(real-budget)/budget:0;
  const xpvVar=xpvBudget?(xpvReal-xpvBudget)/xpvBudget:0;
  let score=100-Math.max(0,budgetVar*180)-critical*6+Math.max(-8,Math.min(8,xpvVar*35));
  score=Math.max(0,Math.min(100,Math.round(score)));
  const label=MONTHS.find(x=>x[0]===state.analysis.cutoff)?.[1]||state.analysis.cutoff;
  return {
    id:`${new Date().getFullYear()}-${state.analysis.cutoff}`,
    label:`${label} 2026`,
    cutoff:state.analysis.cutoff,
    createdAt:new Date().toISOString(),
    real,budget,real25,xpvReal,xpvBudget,budgetVar,xpvVar,
    favorable,critical,totalManagers:managers.length,score
  };
}
function loadProHistory(){
  try{return JSON.parse(localStorage.getItem(PRO_HISTORY_KEY)||'[]')}catch(e){return[]}
}
function saveProHistory(snapshot){
  if(!snapshot)return;
  let list=loadProHistory().filter(x=>x.id!==snapshot.id);
  list.push(snapshot);
  list=list.slice(-18);
  localStorage.setItem(PRO_HISTORY_KEY,JSON.stringify(list));
  renderHistoryModule();
}
function renderProEdition(costRows,expenseRows,xpvRows,cutoff){
  ['explorador','historial','informePro','comparador'].forEach(id=>document.getElementById(id)?.classList.remove('hidden'));
  document.getElementById('proCommandStrip')?.classList.remove('hidden');
  setupDrilldown();
  renderHistoryModule();
  renderProNarrative();
  enrichComparator();
  renderPriorityCenter();
  renderExecutiveClose();
}
function setupDrilldown(){
  const manager=document.getElementById('drillManager'),group=document.getElementById('drillGroup'),account=document.getElementById('drillAccount');
  if(!manager||!group||!account)return;
  const managers=proManagers();
  manager.innerHTML='<option value="">Todas</option>'+managers.map(m=>`<option>${escapeHtml(m)}</option>`).join('');
  const updateGroups=()=>{
    const rows=proFinancialRows().filter(r=>!manager.value||r.manager===manager.value);
    const groups=[...new Set(rows.map(r=>r.group).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'es'));
    group.innerHTML='<option value="">Todos</option>'+groups.map(g=>`<option>${escapeHtml(g)}</option>`).join('');
    updateAccounts();
  };
  const updateAccounts=()=>{
    const rows=proFinancialRows().filter(r=>(!manager.value||r.manager===manager.value)&&(!group.value||r.group===group.value));
    const accounts=[...new Set(rows.map(r=>r.account).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b),'es'));
    account.innerHTML='<option value="">Todas</option>'+accounts.map(a=>`<option>${escapeHtml(a)}</option>`).join('');
    renderDrilldown();
  };
  manager.onchange=updateGroups; group.onchange=updateAccounts; account.onchange=renderDrilldown;
  updateGroups();
}
function renderDrilldown(){
  if(!state.analysis)return;
  const manager=document.getElementById('drillManager')?.value||'',group=document.getElementById('drillGroup')?.value||'',account=document.getElementById('drillAccount')?.value||'';
  const rows=proFinancialRows().filter(r=>(!manager||r.manager===manager)&&(!group||r.group===group)&&(!account||r.account===account));
  const real=sum(rows,'real26'),budget=sum(rows,'budget26'),real25=sum(rows,'real25');
  const varValue=real-budget,varPct=budget?varValue/budget:0,yoy=real25?(real-real25)/real25:0;
  const box=document.getElementById('drillKpis');
  if(box)box.innerHTML=[
    ['REAL 2026',money(real),`${rows.length.toLocaleString('es-MX')} registros`],
    ['PRESUPUESTO',money(budget),'Base comparable'],
    ['DESVIACIÓN',money(varValue),pct(varPct)],
    ['VS 2025',pct(yoy),money(real-real25)]
  ].map(x=>`<article class="kpi-card"><div class="kpi-label">${x[0]}</div><div class="kpi-value">${x[1]}</div><div class="kpi-detail">${x[2]}</div></article>`).join('');
  const detail=topByAbs(aggregate(rows,['account','group'],['real26','budget26']).map(addVariance),'varBudget',12);
  const tbody=document.getElementById('drillTable');
  if(tbody)tbody.innerHTML=detail.map(d=>`<tr><td>${escapeHtml(d.account||'—')}</td><td>${escapeHtml(d.group||'—')}</td><td class="num">${money(d.real26)}</td><td class="num">${money(d.budget26)}</td><td class="num"><span class="pill ${d.varBudget>0?'bad':'good'}">${money(d.varBudget)}</span></td></tr>`).join('')||'<tr><td colspan="5">Sin datos para los filtros seleccionados.</td></tr>';
  const old=Chart.getChart(document.getElementById('drillChart')); if(old)old.destroy();
  safeChart('drillChart',{type:'bar',data:{labels:detail.slice(0,8).map(d=>String(d.account||d.group||'Sin clasificar').slice(0,24)),datasets:[{label:'Real',data:detail.slice(0,8).map(d=>d.real26)},{label:'Presupuesto',data:detail.slice(0,8).map(d=>d.budget26)}]},options:{responsive:true,maintainAspectRatio:false,indexAxis:'y',plugins:{legend:{position:'bottom'}},scales:{x:{ticks:{callback:v=>shortMoney(v)}}}}});
}
function renderHistoryModule(){
  const history=loadProHistory();
  const current=proSnapshot();
  const k=document.getElementById('historyKpis');
  if(k){
    const prev=history.length?history[history.length-1]:null;
    const delta=prev&&current&&prev.real?((current.real-prev.real)/Math.abs(prev.real)):0;
    k.innerHTML=[
      ['CORTES GUARDADOS',history.length.toLocaleString('es-MX'),'Historial local'],
      ['ÚLTIMO CORTE',history.at(-1)?.label||'—','Periodo más reciente'],
      ['CAMBIO REAL',prev&&current?pct(delta):'—','Vs corte guardado anterior'],
      ['MEJOR PULSO',history.length?Math.max(...history.map(x=>x.score))+'/100':'—','Máximo histórico']
    ].map(x=>`<article class="kpi-card"><div class="kpi-label">${x[0]}</div><div class="kpi-value">${x[1]}</div><div class="kpi-detail">${x[2]}</div></article>`).join('');
  }
  const body=document.getElementById('historyTable');
  if(body)body.innerHTML=history.slice().reverse().map(x=>`<tr><td>${escapeHtml(x.label)}</td><td class="num">${money(x.real)}</td><td class="num">${money(x.budget)}</td><td class="num">${pct(x.budgetVar)}</td><td class="num">${x.score}/100</td></tr>`).join('')||'<tr><td colspan="5">Aún no has guardado cortes.</td></tr>';
  const el=document.getElementById('historyChart'); if(!el)return;
  const old=Chart.getChart(el);if(old)old.destroy();
  safeChart('historyChart',{type:'line',data:{labels:history.map(x=>x.label),datasets:[{label:'Real',data:history.map(x=>x.real),tension:.35,borderWidth:3},{label:'Presupuesto',data:history.map(x=>x.budget),tension:.35,borderDash:[6,5]}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{position:'bottom'}},scales:{y:{ticks:{callback:v=>shortMoney(v)}}}}});
}
function renderProNarrative(){
  const s=proSnapshot(); if(!s)return;
  const fin=proFinancialRows();
  const managers=aggregate(fin,['manager'],['real26','budget26']).map(addVariance).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const worst=managers[0],best=managers[managers.length-1];
  document.getElementById('proReportTitle').textContent=`Informe Ejecutivo RCV · ${s.label}`;
  document.getElementById('proReportSubtitle').textContent=`Generado automáticamente con ${s.totalManagers} gerencias analizadas.`;
  document.getElementById('proReportScore').textContent=`${s.score}/100`;
  const blocks=[
    ['Resumen general',s.budgetVar>0?`La operación se encuentra ${pct(Math.abs(s.budgetVar))} por encima del presupuesto consolidado, con una desviación de ${money(s.real-s.budget)}.`:`La operación mantiene una posición favorable de ${pct(Math.abs(s.budgetVar))} por debajo del presupuesto, equivalente a ${money(Math.abs(s.real-s.budget))}.`],
    ['Pulso RCV',`El Pulso RCV se ubica en ${s.score}/100. ${s.favorable} de ${s.totalManagers} gerencias están dentro o por debajo del presupuesto y ${s.critical} se encuentran en condición crítica.`],
    ['Principal foco',worst?`${worst.manager} presenta la mayor presión presupuestal con ${pct(worst.varBudgetPct||0)} de desviación y ${money(worst.varBudget)} de impacto.`:'No se identificó un foco principal.'],
    ['Mejor posición',best?`${best.manager} presenta la mejor posición relativa frente al presupuesto con ${pct(best.varBudgetPct||0)}.`:'Sin información suficiente.'],
    ['Productividad XPV',`La productividad XPV registra ${money(s.xpvReal)} frente a un objetivo de ${money(s.xpvBudget)}, con una variación de ${pct(s.xpvVar)}.`],
    ['Recomendación',s.critical>0?'Priorizar la revisión de las gerencias críticas y utilizar el Explorador para identificar cuentas y agrupadores que explican la desviación.':'Mantener seguimiento preventivo y documentar las prácticas de las gerencias con mejor desempeño para replicarlas.']
  ];
  document.getElementById('proNarrative').innerHTML=blocks.map((b,i)=>`<div class="pro-narrative-block"><span>${String(i+1).padStart(2,'0')}</span><div><strong>${b[0]}</strong><p>${b[1]}</p></div></div>`).join('');
}
function enrichComparator(){
  const box=document.getElementById('comparisonCards');
  if(box) box.classList.add('pro-comparison');
}
async function downloadNarrativePdf(){
  const s=proSnapshot();if(!s||!window.jspdf){alert('Primero procesa los reportes.');return;}
  const {jsPDF}=window.jspdf;
  const pdf=new jsPDF('p','mm','a4');
  const margin=16,maxW=178;
  pdf.setFont('helvetica','bold');pdf.setFontSize(20);pdf.text('REPORT.IA RCV',margin,20);
  pdf.setFontSize(13);pdf.text(`Informe Ejecutivo · ${s.label}`,margin,30);
  pdf.setFont('helvetica','normal');pdf.setFontSize(9);pdf.text(`Pulso RCV: ${s.score}/100`,margin,38);
  const text=(document.getElementById('proNarrative')?.innerText||'').replace(/\n{3,}/g,'\n\n');
  const lines=pdf.splitTextToSize(text,maxW);
  pdf.text(lines,margin,50);
  pdf.save(`REPORTIA_RCV_Informe_Ejecutivo_${s.label.replace(/\s+/g,'_')}.pdf`);
}
document.addEventListener('DOMContentLoaded',()=>{
  document.querySelectorAll('.pro-command[data-target]').forEach(b=>b.addEventListener('click',()=>document.getElementById(b.dataset.target)?.scrollIntoView({behavior:'smooth'})));
  document.getElementById('saveSnapshotBtn')?.addEventListener('click',()=>{const s=proSnapshot();if(!s){alert('Primero procesa los reportes.');return;}saveProHistory(s)});
  document.getElementById('clearHistoryBtn')?.addEventListener('click',()=>{if(confirm('¿Deseas limpiar el historial local de REPORT.IA?')){localStorage.removeItem(PRO_HISTORY_KEY);renderHistoryModule();}});
  document.getElementById('downloadNarrativePdf')?.addEventListener('click',downloadNarrativePdf);
});


// ===== REPORT.IA RCV · PRO EDITION 5.0 =====
function renderPriorityCenter(){
  if(!state.analysis)return;
  const fin=proFinancialRows();
  const xpv=state.analysis.xpvRows||[];
  const managers=aggregate(fin,['manager'],['real26','budget26']).map(addVariance);
  const critical=managers.filter(x=>(x.varBudgetPct||0)>.10).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const attention=managers.filter(x=>(x.varBudgetPct||0)>0&&(x.varBudgetPct||0)<=.10).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const favorable=managers.filter(x=>(x.varBudgetPct||0)<=0).sort((a,b)=>a.varBudgetPct-b.varBudgetPct);
  const accounts=topByAbs(aggregate(fin,['account','group'],['real26','budget26']).map(addVariance),'varBudget',5);
  const xpvReal=sum(xpv,'real'),xpvBudget=sum(xpv,'budget'),xpvVar=xpvBudget?(xpvReal-xpvBudget)/xpvBudget:0;
  const items=[
    {
      cls:'critical',icon:'!',title:`${critical.length} gerencia${critical.length===1?'':'s'} crítica${critical.length===1?'':'s'}`,
      text:critical[0]?`${critical[0].manager} encabeza la presión presupuestal con ${pct(critical[0].varBudgetPct)}.`:'No hay gerencias en estado crítico.',
      target:'gerencias',cta:'Revisar gerencias'
    },
    {
      cls:'attention',icon:'⌁',title:`${attention.length} foco${attention.length===1?'':'s'} de atención`,
      text:accounts[0]?`La cuenta ${accounts[0].account||'sin clasificar'} concentra un impacto de ${money(accounts[0].varBudget)}.`:'No se detectaron desviaciones relevantes.',
      target:'explorador',cta:'Abrir explorador'
    },
    {
      cls:'good',icon:'↗',title:'Productividad XPV',
      text:`XPV se encuentra ${xpvVar>=0?pct(xpvVar)+' arriba':pct(Math.abs(xpvVar))+' abajo'} del objetivo del periodo.`,
      target:'xpv',cta:'Ver productividad'
    },
    {
      cls:'insight',icon:'✦',title:'Oportunidad ejecutiva',
      text:favorable[0]?`${favorable[0].manager} presenta la mejor posición frente al presupuesto con ${pct(favorable[0].varBudgetPct)}.`:'Aún no hay una oportunidad destacada.',
      target:'comparador',cta:'Comparar desempeño'
    }
  ];
  const grid=document.getElementById('priorityGrid');
  if(grid)grid.innerHTML=items.map(i=>`<article class="priority-card ${i.cls}" data-target="${i.target}">
    <div class="priority-icon">${i.icon}</div>
    <div><span>${i.title}</span><p>${i.text}</p><button>${i.cta} →</button></div>
  </article>`).join('');
  grid?.querySelectorAll('[data-target]').forEach(el=>el.addEventListener('click',()=>document.getElementById(el.dataset.target)?.scrollIntoView({behavior:'smooth'})));
}
function executiveSummaryText(){
  const s=proSnapshot();if(!s)return null;
  const fin=proFinancialRows();
  const managers=aggregate(fin,['manager'],['real26','budget26']).map(addVariance).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const worst=managers[0], best=managers[managers.length-1];
  let intro=s.budgetVar>0
    ?`El periodo cierra con una desviación desfavorable de ${money(s.real-s.budget)}, equivalente a ${pct(Math.abs(s.budgetVar))} sobre presupuesto.`
    :`El periodo mantiene una posición favorable de ${money(Math.abs(s.real-s.budget))}, equivalente a ${pct(Math.abs(s.budgetVar))} por debajo del presupuesto.`;
  let middle=` ${s.favorable} de ${s.totalManagers} gerencias se encuentran dentro del objetivo.`;
  let focus=worst?` El principal foco se concentra en ${worst.manager}, con una desviación de ${pct(worst.varBudgetPct||0)}.`:'';
  let opportunity=best?` La mejor posición corresponde a ${best.manager}.`:'';
  return intro+middle+focus+opportunity;
}
function renderExecutiveClose(){
  const s=proSnapshot();if(!s)return;
  const text=executiveSummaryText();
  const el=document.getElementById('closeNarrative');if(el)el.textContent=text;
  document.getElementById('closeTitle').textContent=`Conclusión ejecutiva · ${s.label}`;
  document.getElementById('closeScore').textContent=`${s.score}/100`;
  document.getElementById('closeStatus').textContent=s.score>=85?'Desempeño sólido':s.score>=70?'Desempeño estable':s.score>=55?'Requiere atención':'Prioridad crítica';
}
function copilotAnswer(question){
  if(!state.analysis)return 'Primero procesa los reportes para que pueda analizar la información.';
  const q=(question||'').toLowerCase();
  const fin=proFinancialRows(),xpv=state.analysis.xpvRows||[];
  const managers=aggregate(fin,['manager'],['real26','budget26']).map(addVariance).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const worst=managers[0],best=managers[managers.length-1];
  const xpvReal=sum(xpv,'real'),xpvBudget=sum(xpv,'budget'),xpvVar=xpvBudget?(xpvReal-xpvBudget)/xpvBudget:0;
  if(q.includes('peor')||q.includes('crít')||q.includes('presupuesto')){
    return worst?`${worst.manager} presenta actualmente la mayor presión presupuestal: ${pct(worst.varBudgetPct||0)} de desviación, equivalente a ${money(worst.varBudget)}.`:'No encontré información suficiente.';
  }
  if(q.includes('mejor')){
    return best?`${best.manager} tiene la mejor posición frente al presupuesto con una variación de ${pct(best.varBudgetPct||0)}.`:'No encontré información suficiente.';
  }
  if(q.includes('xpv')||q.includes('productividad')){
    return `La productividad XPV registra ${money(xpvReal)} frente a un objetivo de ${money(xpvBudget)}. La variación es de ${pct(xpvVar)}.`;
  }
  if(q.includes('prioridad')||q.includes('revisar')||q.includes('atención')){
    const critical=managers.filter(x=>(x.varBudgetPct||0)>.10);
    return critical.length?`Yo priorizaría ${critical.slice(0,3).map(x=>x.manager).join(', ')}. Son las gerencias con mayor desviación relativa frente al presupuesto.`:'No hay gerencias críticas. Conviene revisar las mayores desviaciones absolutas desde el Explorador.';
  }
  const matched=managers.find(m=>q.includes(String(m.manager||'').toLowerCase()));
  if(matched){
    return `${matched.manager}: Real ${money(matched.real26)}, Presupuesto ${money(matched.budget26)}, desviación ${money(matched.varBudget)} (${pct(matched.varBudgetPct||0)}).`;
  }
  return executiveSummaryText()||'No pude interpretar esa pregunta. Prueba con una gerencia específica, presupuesto, XPV o prioridades.';
}
function addCopilotMessage(text,type='user'){
  const box=document.getElementById('copilotConversation');if(!box)return;
  const wrap=document.createElement('div');wrap.className=`copilot-message ${type}`;
  wrap.innerHTML=type==='bot'
    ?`<div class="copilot-avatar">R</div><div><strong>Copiloto REPORT.IA</strong><p>${escapeHtml(text)}</p></div>`
    :`<div><strong>Tú</strong><p>${escapeHtml(text)}</p></div>`;
  box.appendChild(wrap);box.scrollTop=box.scrollHeight;
}
function askCopilot(q){
  const text=(q||document.getElementById('copilotInput')?.value||'').trim();if(!text)return;
  addCopilotMessage(text,'user');
  setTimeout(()=>addCopilotMessage(copilotAnswer(text),'bot'),150);
  const input=document.getElementById('copilotInput');if(input)input.value='';
}
function enterCommitteeMode(){
  document.body.classList.add('committee-mode');
  const sections=['dashboard','prioridades','gerencias','tendencias','cierreEjecutivo'];
  document.querySelectorAll('main > section').forEach(s=>{
    if(s.id && !sections.includes(s.id))s.classList.add('committee-hidden');
  });
  document.getElementById('dashboard')?.scrollIntoView({behavior:'smooth'});
}
function exitCommitteeMode(){
  document.body.classList.remove('committee-mode');
  document.querySelectorAll('.committee-hidden').forEach(s=>s.classList.remove('committee-hidden'));
}
function toggleCommitteeMode(){
  if(document.body.classList.contains('committee-mode'))exitCommitteeMode();else enterCommitteeMode();
}
document.addEventListener('DOMContentLoaded',()=>{
  document.getElementById('copilotSendBtn')?.addEventListener('click',()=>askCopilot());
  document.getElementById('copilotInput')?.addEventListener('keydown',e=>{if(e.key==='Enter')askCopilot();});
  document.querySelectorAll('.copilot-suggestions button').forEach(b=>b.addEventListener('click',()=>askCopilot(b.dataset.q)));
  document.getElementById('committeeBtn')?.addEventListener('click',toggleCommitteeMode);
  document.getElementById('closeCommitteeBtn')?.addEventListener('click',enterCommitteeMode);
  document.getElementById('closePdfBtn')?.addEventListener('click',downloadNarrativePdf);
  document.querySelectorAll('#closeActions [data-target]').forEach(b=>b.addEventListener('click',()=>document.getElementById(b.dataset.target)?.scrollIntoView({behavior:'smooth'})));
  document.addEventListener('keydown',e=>{
    if(e.key==='Escape'&&document.body.classList.contains('committee-mode'))exitCommitteeMode();
  });
});


// ===== REPORT.IA RCV · ENTERPRISE EDITION 6.0 =====
const AUTH_ENDPOINT=window.REPORTIA_CONFIG?.authEndpoint||"https://script.google.com/macros/s/AKfycbxUhENeMAGaVJx2Gs4yR_qncJJxHyq8NlFFSfa9qu7XBDgcDu4L9HasfrzZSQBOKgwp/exec";
const SESSION_KEY='reportia_rcv_enterprise_session_v1';

function getSession(){
  try{return JSON.parse(sessionStorage.getItem(SESSION_KEY)||'null')}catch(e){return null}
}
function setSession(s){sessionStorage.setItem(SESSION_KEY,JSON.stringify(s));}
function clearSession(){sessionStorage.removeItem(SESSION_KEY);}

function normalizeLoginResponse(data,username){
  const ok=Boolean(data?.ok ?? data?.success ?? data?.acceso ?? data?.autorizado ?? data?.valid);
  const user=data?.usuario ?? data?.user ?? data?.nombre ?? data?.name ?? username;
  const role=data?.tipo ?? data?.role ?? data?.perfil ?? data?.rango ?? 'USUARIO';
  return {ok,user:String(user||username),role:String(role||'USUARIO').toUpperCase()};
}

async function loginEnterprise(){
  const user=(document.getElementById('loginUser')?.value||'').trim();
  const pass=document.getElementById('loginPassword')?.value;
  const status=document.getElementById('loginStatus');
  if(!user||!pass){status.textContent='Ingresa usuario y contraseña.';status.className='login-status error';return;}
  status.textContent='Validando acceso...';status.className='login-status loading';
  try{
    const url=new URL(AUTH_ENDPOINT);
    url.searchParams.set('accion','login');
    url.searchParams.set('usuario',user);
    url.searchParams.set('contrasena',pass);
    const res=await fetch(url.toString(),{method:'GET',cache:'no-store'});
    const text=await res.text();
    let data; try{data=JSON.parse(text)}catch(e){throw new Error('El Apps Script no devolvió JSON válido.');}
    const auth=normalizeLoginResponse(data,user);
    if(!auth.ok)throw new Error(data?.mensaje||data?.message||'Usuario o contraseña incorrectos.');
    const session={name:auth.user,role:auth.role,loginAt:new Date().toISOString()};
    setSession(session);applySession(session);showApp();
  }catch(err){
    status.textContent=err.message||'No fue posible validar el acceso.';
    status.className='login-status error';
  }
}

function initials(name){
  return String(name||'U').split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join('').toUpperCase();
}
function applySession(session){
  const name=session?.name||'Usuario',role=session?.role||'USUARIO';
  const sn=document.getElementById('sessionName'); if(sn)sn.textContent=name;
  const sr=document.getElementById('sessionRole'); if(sr)sr.textContent=role;
  const av=document.getElementById('sessionAvatar'); if(av)av.textContent=initials(name);
  const sideName=document.querySelector('.mockup-user strong'); if(sideName)sideName.textContent=name;
  const sideRole=document.querySelector('.mockup-user span'); if(sideRole)sideRole.textContent=role;
  const sideAvatar=document.querySelector('.mockup-avatar'); if(sideAvatar)sideAvatar.textContent=initials(name);
  const greet=document.querySelector('.mockup-greeting h1'); if(greet)greet.textContent=`👋 Hola, ${name.split(/\s+/)[0]}.`;
  applyRolePermissions(role);
}
function applyRolePermissions(role){
  const isAdmin=role.includes('ADMIN');
  document.body.dataset.role=role;
  document.querySelectorAll('[data-admin-only]').forEach(el=>el.classList.toggle('role-hidden',!isAdmin));
  // Enterprise policy: non-admin can analyze and export, but uploader remains available unless a future role policy says otherwise.
}
function showApp(){
  document.getElementById('enterpriseLogin')?.classList.add('login-hidden');
  document.body.classList.add('authenticated');
  showRoute('enterpriseHome');
}
function showLogin(){
  document.getElementById('enterpriseLogin')?.classList.remove('login-hidden');
  document.body.classList.remove('authenticated');
}
function showRoute(id){
  const routeIds=['enterpriseHome','dashboard','prioridades','gerencias','costos','gastos','xpv','tendencias','comparador','inteligencia','explorador','historial','copiloto','informePro','cierreEjecutivo','resultados'];
  routeIds.forEach(r=>{
    const el=document.getElementById(r); if(!el)return;
    el.classList.toggle('enterprise-route-hidden',r!==id);
  });
  document.querySelectorAll('.mockup-nav a[data-route]').forEach(a=>a.classList.toggle('active',a.dataset.route===id));
  window.scrollTo({top:0,behavior:'smooth'});
}
function renderEnterpriseHome(){
  const s=typeof proSnapshot==='function'?proSnapshot():null;
  if(!s)return;
  const fin=typeof proFinancialRows==='function'?proFinancialRows():[];
  const managers=aggregate(fin,['manager'],['real26','budget26']).map(addVariance).sort((a,b)=>b.varBudgetPct-a.varBudgetPct);
  const worst=managers[0],second=managers[1],best=managers[managers.length-1];
  document.getElementById('enterpriseGreeting').textContent=`${s.score>=80?'Excelente lectura del periodo':'Resumen ejecutivo del periodo'} · ${s.label}`;
  document.getElementById('enterpriseSummary').textContent=executiveSummaryText()||'El análisis ya está disponible.';
  document.getElementById('homePulse').textContent=`${s.score}/100`;
  document.getElementById('homePulseText').textContent=s.score>=85?'Desempeño sólido':s.score>=70?'Desempeño estable':s.score>=55?'Requiere atención':'Prioridad crítica';
  document.getElementById('homePriority1').textContent=worst?`${worst.manager} · ${pct(worst.varBudgetPct||0)}`:'Sin focos críticos';
  document.getElementById('homePriority2').textContent=second?`${second.manager} · ${pct(second.varBudgetPct||0)}`:'Sin segunda prioridad';
  document.getElementById('homeOpportunity').textContent=best?`${best.manager} · ${pct(best.varBudgetPct||0)}`:'Sin oportunidad destacada';
}
const _oldRenderProEdition=typeof renderProEdition==='function'?renderProEdition:null;
if(_oldRenderProEdition){
  renderProEdition=function(...args){
    _oldRenderProEdition(...args);
    renderEnterpriseHome();
    document.getElementById('enterpriseHome')?.classList.remove('hidden');
  }
}

document.addEventListener('DOMContentLoaded',()=>{
  const session=getSession();
  if(session){applySession(session);showApp();} else showLogin();

  document.getElementById('loginBtn')?.addEventListener('click',loginEnterprise);
  document.getElementById('loginPassword')?.addEventListener('keydown',e=>{if(e.key==='Enter')loginEnterprise();});
  document.getElementById('togglePassword')?.addEventListener('click',()=>{
    const p=document.getElementById('loginPassword'); if(p)p.type=p.type==='password'?'text':'password';
  });
  document.getElementById('sessionUserBtn')?.addEventListener('click',()=>document.getElementById('sessionMenu')?.classList.toggle('hidden'));
  document.getElementById('logoutBtn')?.addEventListener('click',()=>{clearSession();location.reload();});

  document.querySelectorAll('.mockup-nav a[data-route]').forEach(a=>a.addEventListener('click',e=>{e.preventDefault();showRoute(a.dataset.route);}));
  document.querySelectorAll('.home-launch[data-go]').forEach(b=>b.addEventListener('click',()=>showRoute(b.dataset.go)));
});


// ===== ENTERPRISE EDITION 6.1 · LISTA DESPLEGABLE DE USUARIOS =====
async function loadAuthorizedUsers(){
  const select=document.getElementById('loginUser');
  const status=document.getElementById('loginStatus');
  if(!select)return;

  select.innerHTML='<option value="">Cargando usuarios...</option>';
  select.disabled=true;
  if(status){
    status.textContent='Cargando usuarios autorizados...';
    status.className='login-status loading';
  }

  try{
    let data=null;

    // 1) Intento normal con fetch.
    try{
      const url=new URL(AUTH_ENDPOINT);
      url.searchParams.set('accion','usuarios');
      const res=await fetch(url.toString(),{method:'GET',cache:'no-store',redirect:'follow'});
      if(res.ok){
        const text=await res.text();
        try{ data=JSON.parse(text); }catch(e){}
      }
    }catch(fetchError){
      console.warn('Fetch directo no disponible; usando JSONP.',fetchError);
    }

    // 2) Fallback JSONP: evita bloqueos CORS/redirect de Apps Script.
    if(!data || !data.ok){
      data=await loadUsersViaJsonp();
    }

    const users=Array.isArray(data?.usuarios)?data.usuarios:[];
    if(!data?.ok || users.length===0){
      throw new Error(data?.mensaje||'No se encontraron usuarios autorizados.');
    }

    select.innerHTML='<option value="">Selecciona tu usuario</option>'+
      users.map(u=>{
        const username=typeof u==='string'?u:(u.usuario||u.nombre||u.user||'');
        const role=typeof u==='string'?'':(u.tipo||u.role||u.rango||'');
        return `<option value="${escapeHtml(username)}" data-role="${escapeHtml(role)}">${escapeHtml(username)}</option>`;
      }).join('');

    select.disabled=false;

    if(status){
      status.textContent=`${users.length} usuario${users.length===1?'':'s'} autorizado${users.length===1?'':'s'} disponible${users.length===1?'':'s'}.`;
      status.className='login-status success';
    }
  }catch(err){
    select.innerHTML='<option value="">No fue posible cargar usuarios</option>';
    select.disabled=true;
    if(status){
      status.textContent=(err && err.message) ? err.message : 'No fue posible cargar la lista de usuarios.';
      status.className='login-status error';
    }
  }

  updateLoginSelectedAvatar();
}

function loadUsersViaJsonp(){
  return new Promise((resolve,reject)=>{
    const callback='reportiaUsersCallback_'+Date.now()+'_'+Math.random().toString(36).slice(2);
    const script=document.createElement('script');
    const timeout=setTimeout(()=>{
      cleanup();
      reject(new Error('Tiempo de espera agotado al cargar usuarios.'));
    },12000);

    function cleanup(){
      clearTimeout(timeout);
      delete window[callback];
      script.remove();
    }

    window[callback]=(data)=>{
      cleanup();
      resolve(data);
    };

    script.onerror=()=>{
      cleanup();
      reject(new Error('No fue posible conectar con el servicio de usuarios.'));
    };

    const url=new URL(AUTH_ENDPOINT);
    url.searchParams.set('accion','usuarios');
    url.searchParams.set('callback',callback);
    script.src=url.toString();
    document.head.appendChild(script);
  });
}

function updateLoginSelectedAvatar(){
  const select=document.getElementById('loginUser');
  const avatar=document.getElementById('loginSelectedAvatar');
  if(!select||!avatar)return;
  const name=select.value||'';
  avatar.textContent=name?initials(name):'?';
  avatar.classList.toggle('has-user',Boolean(name));
}

document.addEventListener('DOMContentLoaded',()=>{
  loadAuthorizedUsers();
  document.getElementById('loginUser')?.addEventListener('change',updateLoginSelectedAvatar);
  document.getElementById('reloadUsersBtn')?.addEventListener('click',loadAuthorizedUsers);
});
