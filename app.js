/* REPORT.IA RCV - procesamiento 100% en navegador */
const state = { files: [], sources: {}, outputs: {} };
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
function reset(){state.files=[];state.sources={};state.outputs={};$('#fileInput').value='';renderFiles();renderValidation();$('#resultados').classList.add('hidden')}
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
    renderDownloads();$('#resultados').classList.remove('hidden');$('#resultados').scrollIntoView({behavior:'smooth'});
    $('#statusText').textContent=`Listo: ${costRows.length.toLocaleString('es-MX')} filas de costo, ${expenseRows.length.toLocaleString('es-MX')} de gastos y ${xpvRows.length.toLocaleString('es-MX')} de XPV procesadas.`;
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
