(() => {
"use strict";

const SHEETS = {
  CATALOG: "CATALOGO DE GERENTES",
  COST: "JD COSTO RCV",
  EXPENSE: "JD GASTOS RCV",
  XPV: "JD XPV RCV"
};
const MONTHS = ["EneM","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","DicM"];
const MONTH_LABELS = ["Ene","Feb","Mar","Abr","May","Jun","Jul","Ago","Sep","Oct","Nov","Dic"];
const PREMISES = new Set(["5101.0403.Servicios de comedor","5101.0404.GASTOS Y PREVISIÓN SOCIAL","5101.0901.Consumo de diesel"]);

const state = {
  sources: {},
  model: null,
  periodIndex: 5,
  finalZipFile: null,
  threshold: 10,
  decisions: JSON.parse(localStorage.getItem("reportia_decisions_v19") || "[]"),
  inputMode: null,
  rpWorkbooks: []
};

const $ = id => document.getElementById(id);
const qa = sel => [...document.querySelectorAll(sel)];

function cleanText(v) {
  return String(v ?? "").replace(/\u00a0/g, " ").trim();
}
function outputText(v) {
  const s = cleanText(v);
  return s || "\u00a0";
}
function parseNumber(v) {
  if (typeof v === "number") return Number.isFinite(v) ? v : 0;
  let s = String(v ?? "").replace(/\u00a0/g,"").replace(/\$/g,"").replace(/,/g,"").replace(/\s+/g,"").trim();
  if (!s || s === "-") return 0;
  const n = Number(s);
  return Number.isFinite(n) ? n : 0;
}
function pct(diff, base) { return base ? diff / base : null; }
function money(v) { return Math.abs(v || 0).toLocaleString("es-MX",{style:"currency",currency:"MXN",maximumFractionDigits:0}); }
function signedMoney(v) { return (v || 0).toLocaleString("es-MX",{style:"currency",currency:"MXN",maximumFractionDigits:0}); }
function pctText(v) { return v == null || !Number.isFinite(v) ? "—" : (v*100).toFixed(2)+"%"; }
function key(parts) { return parts.map(x => String(x ?? "")).join("\u001f"); }

function showView(name) {
  qa("[data-view].view").forEach(v => v.classList.toggle("active", v.dataset.view === name));
  qa("#nav button[data-view]").forEach(b => b.classList.toggle("active", b.dataset.view === name));
  window.scrollTo({top:0,behavior:"smooth"});
}
function setStatus(id, text, type="") {
  const el=$(id); if(!el) return;
  el.textContent=text; el.className="status"+(type ? " "+type : "");
}

function initPeriods() {
  const options = MONTH_LABELS.map((m,i)=>`<option value="${i}">${m} 2026</option>`).join("");
  $("periodSelect").innerHTML = options;
  $("settingsPeriod").innerHTML = options;
  $("periodSelect").value = state.periodIndex;
  $("settingsPeriod").value = state.periodIndex;
}
initPeriods();

async function readWorkbookFile(file) {
  const data = await file.arrayBuffer();
  return XLSX.read(data,{type:"array",raw:true,cellDates:true});
}
function sheetArray(wb, name) {
  const actual = wb.SheetNames.find(n => cleanText(n).toUpperCase() === name);
  if (!actual) return null;
  return XLSX.utils.sheet_to_json(wb.Sheets[actual],{header:1,defval:"",raw:true});
}

async function expandInputFiles(files){
  const expanded=[];
  for(const file of files){
    if(/\.zip$/i.test(file.name)){
      const zip=await JSZip.loadAsync(await file.arrayBuffer());
      for(const entry of Object.values(zip.files)){
        if(entry.dir||!/\.(xlsx|xls)$/i.test(entry.name))continue;
        const blob=await entry.async("blob");
        expanded.push(new File([blob],entry.name.split("/").pop(),{type:blob.type||"application/octet-stream"}));
      }
    }else{
      expanded.push(file);
    }
  }
  return expanded;
}

function findHeaderRow(arr, requiredTokens){
  for(let i=0;i<Math.min(arr.length,30);i++){
    const row=(arr[i]||[]).map(v=>cleanText(v).toUpperCase());
    if(requiredTokens.every(t=>row.some(v=>v.includes(t)))) return i;
  }
  return -1;
}
function headerIndex(headers, alternatives){
  const hs=headers.map(v=>cleanText(v).toUpperCase());
  for(const alt of alternatives){
    const idx=hs.findIndex(v=>v===alt||v.includes(alt));
    if(idx>=0)return idx;
  }
  return -1;
}
function parseFinalLikeSheet(ws, kind){
  if(!ws)return [];
  const arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
  if(!arr.length)return [];

  // Locate the most likely header row. RP files may use merged titles and different column names.
  let hr=-1,bestScore=-1;
  const tokens=["GERENCIA","UNIDAD","CENTRO","CUENTA","CTA","AGRUPADOR","MARCA","REAL","PRESUPUESTO","PTTO","2025","2026","XPV","PRODUCTIVIDAD"];
  for(let i=0;i<Math.min(arr.length,60);i++){
    const row=(arr[i]||[]).map(v=>cleanText(v).toUpperCase());
    const nonempty=row.filter(Boolean).length;
    const score=tokens.reduce((n,t)=>n+(row.some(v=>v.includes(t))?1:0),0)+(nonempty>=3?1:0);
    if(score>bestScore){bestScore=score;hr=i;}
  }
  if(hr<0||bestScore<1)return [];

  const headers=(arr[hr]||[]).map(v=>cleanText(v));
  const idx=(alts)=>{
    const hs=headers.map(v=>v.toUpperCase());
    for(const a of alts){
      let p=hs.findIndex(v=>v===a);
      if(p>=0)return p;
      p=hs.findIndex(v=>v.includes(a));
      if(p>=0)return p;
    }
    return -1;
  };
  const idxManager=idx(["GERENCIA","UNIDAD","REGION","REGIÓN","DIVISION","DIVISIÓN"]);
  const idxAgr=idx(["AGRUPADOR","CONCEPTO","RUBRO"]);
  const idxAccount=idx(["CTA CONTABLE","CUENTA CONTABLE","CUENTA","CTA"]);
  const idxBrand=idx(["CIA MARCA","MARCA","CIA"]);
  const idxCenter=idx(["CENTRO DE GESTION","CENTRO DE GESTIÓN","CENTRO / SUBLIBRO","CENTRO","SUBLIBRO"]);
  const idx2025=idx(["2025"]);
  const idxBudget=idx(["PTTO. 2026","PRESUPUESTO 2026","PRESUPUESTO","PTTO"]);
  let idx2026=headers.findIndex(v=>v.toUpperCase()==="2026");
  if(idx2026<0)idx2026=idx(["REAL 2026","REAL"]);
  const idxReal=idx(["REAL GESTION","REAL GESTIÓN","REAL","XPV"]);
  const numericCandidates=[];
  headers.forEach((h,i)=>{
    const H=h.toUpperCase();
    if(/2025|2026|REAL|PRESUPUESTO|PTTO|TOTAL|IMPORTE|MONTO|XPV/.test(H))numericCandidates.push(i);
  });

  const data=arr.slice(hr+1).filter(r=>r.some(v=>cleanText(v)!==""));
  const labelAt=(r,preferred)=>{
    if(preferred>=0&&cleanText(r[preferred]))return cleanText(r[preferred]);
    for(let i=0;i<r.length;i++){
      const v=r[i];
      if(typeof v==="string"&&cleanText(v)&&!/^[-+]?\d+([.,]\d+)?$/.test(cleanText(v)))return cleanText(v);
    }
    return "SIN CLASIFICAR";
  };
  const numAt=(r,i)=>i>=0?parseNumber(r[i]):0;
  const fallbackNums=(r)=>r.map((v,i)=>({i,n:parseNumber(v),raw:v})).filter(x=>typeof x.raw==="number"&&Number.isFinite(x.raw));

  return data.map(r=>{
    const nums=fallbackNums(r);
    const manager=labelAt(r,idxManager);
    const account=idxAccount>=0?outputText(r[idxAccount]):(idxAgr>=0?outputText(r[idxAgr]):manager);
    const center=idxCenter>=0?outputText(r[idxCenter]):"";
    const brand=idxBrand>=0?outputText(r[idxBrand]):"";
    if(kind==="XPV"){
      let real=numAt(r,idxReal),budget=numAt(r,idxBudget);
      if(idxReal<0&&nums.length)real=nums[nums.length-1].n;
      if(idxBudget<0&&nums.length>1)budget=nums[nums.length-2].n;
      return {manager,account,center,brand,real,budget,diff:real-budget,pct:pct(real-budget,budget)};
    }
    let y2025=numAt(r,idx2025),budget2026=numAt(r,idxBudget),y2026=numAt(r,idx2026);
    if(idx2026<0&&nums.length)y2026=nums[nums.length-1].n;
    if(idxBudget<0&&nums.length>1)budget2026=nums[nums.length-2].n;
    if(idx2025<0&&nums.length>2)y2025=nums[nums.length-3].n;
    return {agrupador:idxAgr>=0?outputText(r[idxAgr]):"",brand,account,manager,center,
      y2025,budget2026,y2026,diff2526:y2026-y2025,pct2526:pct(y2026-y2025,y2025),
      diffBudget:y2026-budget2026,pctBudget:pct(y2026-budget2026,budget2026)};
  }).filter(r=>{
    if(kind==="XPV")return r.manager!=="SIN CLASIFICAR"||r.real!==0||r.budget!==0;
    return r.manager!=="SIN CLASIFICAR"||r.y2025!==0||r.y2026!==0||r.budget2026!==0;
  });
}
function chooseWorkbookByName(items,token){
  const t=token.toLowerCase();
  return items.find(x=>x.file.name.toLowerCase().includes(t));
}
async function detectAndLoadInputs(files){
  const expanded=await expandInputFiles(files);
  const found={};
  const rpItems=[];
  for(const file of expanded){
    const wb=await readWorkbookFile(file);
    let hasJd=false;
    for(const required of Object.values(SHEETS)){
      const rows=sheetArray(wb,required);
      if(rows){found[required]=rows;hasJd=true;}
    }
    rpItems.push({file,wb});
  }
  const allJd=Object.values(SHEETS).every(s=>found[s]);
  if(allJd){
    state.sources=found;
    state.inputMode="JD";
    state.rpWorkbooks=[];
    return {mode:"JD",files:expanded};
  }

  const costItem=chooseWorkbookByName(rpItems,"costo rcv");
  const expenseCandidates=rpItems.filter(x=>x.file.name.toLowerCase().includes("gastos rcv"));
  const expenseItem=expenseCandidates.find(x=>!x.file.name.toLowerCase().includes("plantilla"))||expenseCandidates[0];
  const xpvItem=chooseWorkbookByName(rpItems,"productividad xpv");
  if(costItem&&expenseItem&&xpvItem){
    state.inputMode="RP";
    state.rpWorkbooks=[costItem,expenseItem,xpvItem];
    state.sources={};
    return {mode:"RP",files:expanded};
  }
  state.inputMode=null;
  state.rpWorkbooks=rpItems;
  return {mode:null,files:expanded};
}

function processRpModel(){
  const costItem=chooseWorkbookByName(state.rpWorkbooks,"costo rcv");
  const expenseCandidates=state.rpWorkbooks.filter(x=>x.file.name.toLowerCase().includes("gastos rcv"));
  const expenseItem=expenseCandidates.find(x=>!x.file.name.toLowerCase().includes("plantilla"))||expenseCandidates[0];
  const xpvItem=chooseWorkbookByName(state.rpWorkbooks,"productividad xpv");
  if(!costItem||!expenseItem||!xpvItem)throw new Error("No se localizaron los tres archivos RP requeridos.");

  const getSheet=(wb,names)=>{
    for(const wanted of names){
      const actual=wb.SheetNames.find(n=>cleanText(n).toUpperCase()===wanted.toUpperCase());
      if(actual)return wb.Sheets[actual];
    }
    return null;
  };

  const costOperating=parseFinalLikeSheet(getSheet(costItem.wb,["COSTOS OPERATIVOS"])||costItem.wb.Sheets[costItem.wb.SheetNames[0]],"COST");
  const costMaintenance=parseFinalLikeSheet(getSheet(costItem.wb,["COSTO MANTTO"])||costItem.wb.Sheets[costItem.wb.SheetNames[1]||costItem.wb.SheetNames[0]],"COST");
  const premises=parseFinalLikeSheet(getSheet(costItem.wb,["premisas 26","PREMISAS 2026"])||costItem.wb.Sheets[costItem.wb.SheetNames[2]||costItem.wb.SheetNames[0]],"COST");
  const smo=parseFinalLikeSheet(getSheet(costItem.wb,["SMO"])||costItem.wb.Sheets[costItem.wb.SheetNames[3]||costItem.wb.SheetNames[0]],"COST");

  const expenseCandidatesSheets=expenseItem.wb.SheetNames.map(n=>expenseItem.wb.Sheets[n]);
  let expenses=[];
  for(const ws of expenseCandidatesSheets){
    const parsed=parseFinalLikeSheet(ws,"COST");
    if(parsed.length>expenses.length)expenses=parsed;
  }

  const xpvSheet=getSheet(xpvItem.wb,["PRODUCTIVIDAD XPV"])||xpvItem.wb.Sheets[xpvItem.wb.SheetNames[0]];
  const xpv=parseFinalLikeSheet(xpvSheet,"XPV");

  if(!costOperating.length&&!costMaintenance.length)throw new Error("El archivo de Costos RP no contiene una estructura reconocible.");
  if(!expenses.length)throw new Error("No se encontraron filas de datos utilizables en el archivo de Gastos RP. Revisa que el archivo no esté protegido o vacío.");
  if(!xpv.length)throw new Error("El archivo de Productividad XPV no contiene una estructura reconocible.");

  const expenseSummary=aggregateByManager(expenses);
  const costSummary=aggregateByManager([...costOperating,...costMaintenance]);
  const zeroMonths=Array.from({length:12},()=>({y2025:0,budget2026:0,y2026:0}));
  state.model={
    catalog:new Map(),costOperating,costMaintenance,premises,smo,expenses,expenseSummary,xpv,costSummary,
    costMonthly:zeroMonths,expenseMonthly:zeroMonths,month:MONTH_LABELS[state.periodIndex],sourceMode:"RP"
  };
  renderAll();
  setStatus("uploadStatus",`Archivos RP procesados correctamente. Dashboard actualizado desde Costos, Gastos y Productividad XPV.`,"ok");
}

async function loadSources(files) {
  const found = {};
  for (const file of files) {
    const wb = await readWorkbookFile(file);
    for (const required of Object.values(SHEETS)) {
      if (!found[required]) {
        const rows = sheetArray(wb, required);
        if (rows) found[required] = rows;
      }
    }
  }
  state.sources = found;
  renderChecklist();
  return found;
}
function renderChecklist() {
  if(state.inputMode==="RP"){
    const names=state.rpWorkbooks.map(x=>x.file.name.toLowerCase());
    const labels=[
      [names.some(n=>n.includes("costo rcv")),"Costo RCV"],
      [names.some(n=>n.includes("gastos rcv")&&!n.includes("plantilla"))||names.some(n=>n.includes("gastos rcv")),"Gastos RCV"],
      [names.some(n=>n.includes("productividad xpv")),"Productividad XPV"]
    ];
    $("sourceChecklist").innerHTML=labels.map(([ok,l])=>`<div class="check ${ok?"ok":""}">${ok?"✓":"○"} ${l}</div>`).join("");
    $("processBtn").disabled=labels.some(x=>!x[0]);
    $("inputMode").innerHTML='Modo de entrada: <b>Archivos RP sueltos / ZIP</b>';
    return;
  }
  const labels = [
    [SHEETS.CATALOG,"Catálogo de Gerentes"],
    [SHEETS.COST,"JD Costo RCV"],
    [SHEETS.EXPENSE,"JD Gastos RCV"],
    [SHEETS.XPV,"JD XPV RCV"]
  ];
  $("sourceChecklist").innerHTML = labels.map(([k,l])=>`<div class="check ${state.sources[k]?"ok":""}">${state.sources[k]?"✓":"○"} ${l}</div>`).join("");
  $("processBtn").disabled = labels.some(([k])=>!state.sources[k]);
  $("inputMode").innerHTML='Modo de entrada: <b>'+(state.inputMode==="JD"?"Concentrado reportes JD":"Sin detectar")+'</b>';
}

function detectPeriodFromXpv() {
  const rows = state.sources[SHEETS.XPV];
  if (!rows || !rows[4]) return 5;
  const v = rows[4][2];
  if (v instanceof Date) return v.getMonth();
  if (typeof v === "number") {
    const d = XLSX.SSF.parse_date_code(v);
    if (d && d.m) return d.m-1;
  }
  return 5;
}

function buildCatalog() {
  const rows = state.sources[SHEETS.CATALOG] || [];
  const map = new Map();
  rows.forEach(r => {
    const a=cleanText(r[0]), b=cleanText(r[1]);
    if(a && b) map.set(a,b);
  });
  return map;
}

function aggregateJd(sheetRows, catalog, periodIndex) {
  const allowed = new Set(MONTHS.slice(0,periodIndex+1));
  const map = new Map();
  const monthly = Array.from({length:12},()=>({y2025:0,budget2026:0,y2026:0}));
  for (let i=8;i<sheetRows.length;i++) {
    const r=sheetRows[i] || [];
    const period=cleanText(r[6]);
    if (!allowed.has(period)) continue;
    const sourceManager=cleanText(r[3]);
    const manager=catalog.get(sourceManager);
    if (!manager) continue;
    const agrupador=cleanText(r[1]), brand=cleanText(r[5]), account=cleanText(r[0]), center=cleanText(r[2]);
    const k=key([agrupador,brand,account,manager,center]);
    if(!map.has(k)) map.set(k,{agrupador,brand,account,manager,center,y2025:0,budget2026:0,y2026:0});
    const o=map.get(k);
    o.y2025 += parseNumber(r[7]);
    o.budget2026 += parseNumber(r[10]);
    o.y2026 += parseNumber(r[9]);
    const mi=MONTHS.indexOf(period);
    if(mi>=0){
      monthly[mi].y2025 += parseNumber(r[7]);
      monthly[mi].budget2026 += parseNumber(r[10]);
      monthly[mi].y2026 += parseNumber(r[9]);
    }
  }
  const rows=[...map.values()].map(o=>{
    o.diff2526=o.y2026-o.y2025; o.pct2526=pct(o.diff2526,o.y2025);
    o.diffBudget=o.y2026-o.budget2026; o.pctBudget=pct(o.diffBudget,o.budget2026);
    return o;
  });
  return {rows,monthly};
}

function aggregateXpv(rows) {
  const map=new Map();
  for(let i=9;i<rows.length;i++){
    const r=rows[i]||[];
    let manager=cleanText(r[3]);
    const account=outputText(r[1]), center=outputText(r[2]), brand=outputText(r[4]);
    if(!manager) manager="SIN CLASIFICAR";
    const k=key([manager,account,center,brand]);
    if(!map.has(k)) map.set(k,{manager,account,center,brand,real:0,budget:0});
    const o=map.get(k); o.real+=parseNumber(r[5]); o.budget+=parseNumber(r[6]);
  }
  return [...map.values()].map(o=>{o.diff=o.real-o.budget;o.pct=pct(o.diff,o.budget);return o;});
}

function aggregateByManager(rows) {
  const map=new Map();
  rows.forEach(r=>{
    if(!map.has(r.manager)) map.set(r.manager,{manager:r.manager,y2025:0,budget2026:0,y2026:0});
    const o=map.get(r.manager);o.y2025+=r.y2025;o.budget2026+=r.budget2026;o.y2026+=r.y2026;
  });
  return [...map.values()].map(o=>{o.diff2526=o.y2026-o.y2025;o.pct2526=pct(o.diff2526,o.y2025);o.diffBudget=o.y2026-o.budget2026;o.pctBudget=pct(o.diffBudget,o.budget2026);return o;});
}

function processModel() {
  const catalog=buildCatalog();
  const costAgg=aggregateJd(state.sources[SHEETS.COST],catalog,state.periodIndex);
  const expAgg=aggregateJd(state.sources[SHEETS.EXPENSE],catalog,state.periodIndex);
  const xpv=aggregateXpv(state.sources[SHEETS.XPV]);

  const costOperating=costAgg.rows.filter(r=>r.manager==="OP PRIMERA"||r.manager==="OP INTERM");
  const costMaintenance=costAgg.rows.filter(r=>r.manager==="MANTTO.");
  const premises=costAgg.rows.filter(r=>(r.manager==="OP PRIMERA"||r.manager==="OP INTERM")&&PREMISES.has(r.account));
  const smo=costAgg.rows.filter(r=>(r.agrupador==="5101.0800.Capacitación"&&["OP PRIMERA","OP INTERM","MANTTO."].includes(r.manager))||(r.agrupador==="5101.1300.Mecánicos"&&r.manager==="MANTTO."));
  const expenses=expAgg.rows;
  const expenseSummary=aggregateByManager(expenses);
  const costSummary=aggregateByManager([...costOperating,...costMaintenance]);

  state.model={
    catalog,costOperating,costMaintenance,premises,smo,expenses,expenseSummary,xpv,costSummary,
    costMonthly:costAgg.monthly,expenseMonthly:expAgg.monthly,
    month:MONTH_LABELS[state.periodIndex]
  };
  renderAll();
  setStatus("uploadStatus",`Procesamiento completado para ${state.model.month} 2026. Los indicadores y reportes fueron recalculados.`,"ok");
}

function modelTotals() {
  const m=state.model;
  if(!m) return null;
  const cost2026=[...m.costOperating,...m.costMaintenance].reduce((s,r)=>s+r.y2026,0);
  const expense2026=m.expenses.reduce((s,r)=>s+r.y2026,0);
  const xpReal=m.xpv.reduce((s,r)=>s+r.real,0);
  const xpBudget=m.xpv.reduce((s,r)=>s+r.budget,0);
  const managers=new Set([...m.costSummary.map(x=>x.manager),...m.expenseSummary.map(x=>x.manager)]);
  const deviations=buildDeviations();
  return {cost2026,expense2026,xpReal,xpBudget,xpCompliance:xpBudget?xpReal/xpBudget:0,managers:managers.size,deviations};
}
function buildDeviations() {
  if(!state.model) return [];
  const threshold=state.threshold/100;
  const out=[];
  state.model.costSummary.forEach(r=>{if(r.pct2526!=null&&r.pct2526>threshold)out.push({type:"Costo",manager:r.manager,y2025:r.y2025,y2026:r.y2026,pct:r.pct2526,impact:Math.abs(r.diff2526)});});
  state.model.expenseSummary.forEach(r=>{if(r.pct2526!=null&&r.pct2526>threshold)out.push({type:"Gasto",manager:r.manager,y2025:r.y2025,y2026:r.y2026,pct:r.pct2526,impact:Math.abs(r.diff2526)});});
  return out.sort((a,b)=>b.impact-a.impact);
}

function renderDashboard(){
  const t=modelTotals();
  if(!t){clearDashboard();return;}
  $("kpiCost").textContent=money(t.cost2026);
  $("kpiExpense").textContent=money(t.expense2026);
  $("kpiXpv").textContent=(t.xpCompliance*100).toFixed(2)+"%";
  $("kpiXpvSub").textContent=`Real ${money(t.xpReal)} / Ptto ${money(t.xpBudget)}`;
  $("kpiManagers").textContent=t.managers;
  $("kpiAlerts").textContent=new Set(t.deviations.map(d=>d.manager)).size;
  $("gaugeValue").textContent=(t.xpCompliance*100).toFixed(1)+"%";
  $("gauge").style.setProperty("--g",Math.max(0,Math.min(100,t.xpCompliance*100)));
  $("gaugeFoot").textContent=`Real ${money(t.xpReal)} · Presupuesto ${money(t.xpBudget)}`;
  $("periodLabel").textContent=`Acumulado a ${state.model.month} 2026`;

  const ranks=state.model.costSummary.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026));
  const max=Math.abs(ranks[0]?.y2026||1);
  $("costRanking").innerHTML=ranks.length?ranks.slice(0,5).map((r,i)=>`<div class="rank-row"><b>${i+1}</b><strong>${r.manager}</strong><div class="bar"><i style="width:${Math.abs(r.y2026)/max*100}%"></i></div><em>${money(r.y2026)}</em></div>`).join(""):'<div class="empty">Sin datos.</div>';

  $("deviationsList").innerHTML=t.deviations.length?t.deviations.slice(0,3).map((d,i)=>`<div class="list-row ${i?"warn":""}"><i></i><div><strong>${d.manager}</strong><span>${d.type} · 2025 vs 2026</span></div><em>+${(d.pct*100).toFixed(1)}%</em></div>`).join(""):'<div class="empty">No hay desviaciones mayores al umbral.</div>';

  const topExpense=state.model.expenseSummary.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026))[0];
  const findings=[
    [`Cumplimiento XPV`,`${(t.xpCompliance*100).toFixed(2)}% del presupuesto del periodo.`],
    [`Mayor costo`,`${ranks[0]?.manager||"—"} concentra ${money(ranks[0]?.y2026||0)}.`],
    [`Mayor gasto`,`${topExpense?.manager||"—"} registra ${money(topExpense?.y2026||0)}.`]
  ];
  $("findingsList").innerHTML=findings.map(x=>`<div class="list-row good"><i></i><div><strong>${x[0]}</strong><span>${x[1]}</span></div></div>`).join("");
  drawTrend();
}

function clearDashboard(){
  ["kpiCost","kpiExpense"].forEach(id=>$(id).textContent="$0");
  $("kpiXpv").textContent="0%";$("kpiManagers").textContent="0";$("kpiAlerts").textContent="0";$("gaugeValue").textContent="0%";$("gauge").style.setProperty("--g",0);$("gaugeFoot").textContent="Real $0 · Presupuesto $0";
  $("costRanking").innerHTML='<div class="empty">Sin datos procesados.</div>';$("deviationsList").innerHTML='<div class="empty">Sin datos procesados.</div>';$("findingsList").innerHTML='<div class="empty">Procesa los archivos para generar hallazgos.</div>';
  drawTrend();
}
function drawTrend(){
  const c=$("trendChart"),ctx=c.getContext("2d"),w=c.width,h=c.height;
  ctx.clearRect(0,0,w,h);
  ctx.strokeStyle="#e8eef7";ctx.lineWidth=1;
  for(let i=1;i<5;i++){const y=i*h/5;ctx.beginPath();ctx.moveTo(45,y);ctx.lineTo(w-15,y);ctx.stroke();}
  if(!state.model){
    ctx.fillStyle="#94a3b8";ctx.font="16px system-ui";ctx.textAlign="center";
    ctx.fillText("Procesa los archivos para visualizar tendencias reales",w/2,h/2);return;
  }

  // JD mode: monthly trend. RP mode: real 2025 vs 2026 by top managers.
  if(state.inputMode==="JD"){
    const upto=state.periodIndex+1;
    const cost=state.model.costMonthly.slice(0,upto).map(x=>Math.abs(x.y2026));
    const exp=state.model.expenseMonthly.slice(0,upto).map(x=>Math.abs(x.y2026));
    const all=[...cost,...exp];const max=Math.max(...all,1);
    const series=[[cost,"#2563eb"],[exp,"#7c3aed"]];
    series.forEach(([arr,color])=>{
      ctx.strokeStyle=color;ctx.lineWidth=3;ctx.beginPath();
      arr.forEach((v,i)=>{const x=55+i*(w-100)/Math.max(1,arr.length-1),y=h-35-v/max*(h-70);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});
      ctx.stroke();
      arr.forEach((v,i)=>{const x=55+i*(w-100)/Math.max(1,arr.length-1),y=h-35-v/max*(h-70);ctx.beginPath();ctx.fillStyle=color;ctx.arc(x,y,4,0,Math.PI*2);ctx.fill();});
    });
    ctx.fillStyle="#64748b";ctx.font="11px system-ui";ctx.textAlign="center";
    MONTH_LABELS.slice(0,upto).forEach((m,i)=>{const x=55+i*(w-100)/Math.max(1,upto-1);ctx.fillText(m,x,h-12);});
  }else{
    const rows=state.model.costSummary.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026)).slice(0,6);
    const max=Math.max(...rows.flatMap(r=>[Math.abs(r.y2025),Math.abs(r.y2026)]),1);
    const groupW=(w-110)/Math.max(1,rows.length);
    rows.forEach((r,i)=>{
      const x=60+i*groupW;
      const h25=Math.abs(r.y2025)/max*(h-95);
      const h26=Math.abs(r.y2026)/max*(h-95);
      ctx.fillStyle="#22c55e";ctx.fillRect(x,h-38-h25,groupW*.26,h25);
      ctx.fillStyle="#f97316";ctx.fillRect(x+groupW*.31,h-38-h26,groupW*.26,h26);
      ctx.fillStyle="#64748b";ctx.font="10px system-ui";ctx.textAlign="center";
      ctx.fillText((r.manager||"").slice(0,12),x+groupW*.28,h-14);
    });
    ctx.fillStyle="#22c55e";ctx.fillRect(55,12,12,12);ctx.fillStyle="#475569";ctx.textAlign="left";ctx.fillText("2025",72,22);
    ctx.fillStyle="#f97316";ctx.fillRect(115,12,12,12);ctx.fillStyle="#475569";ctx.fillText("2026",132,22);
  }
}
function renderFilters(){
  if(!state.model)return;
  const managers=[...new Set([...state.model.costSummary.map(x=>x.manager),...state.model.expenseSummary.map(x=>x.manager)])].sort();
  $("managerFilter").innerHTML='<option value="">Todas las gerencias</option>'+managers.map(x=>`<option>${x}</option>`).join("");
  const brands=[...new Set([...state.model.costOperating,...state.model.costMaintenance,...state.model.expenses].map(x=>x.brand).filter(Boolean))].sort();
  $("brandFilter").innerHTML='<option value="">Todas</option>'+brands.map(x=>`<option>${x}</option>`).join("");
}
function renderProcessStats(){
  const m=state.model;if(!m)return;
  const vals=[m.costOperating.length,m.costMaintenance.length,m.premises.length,m.smo.length,m.expenses.length,m.xpv.length];
  qa("#processStats strong").forEach((el,i)=>el.textContent=vals[i]??0);
}
function renderAnalysis(){
  const t=modelTotals();
  if(!t){
    $("analysisCards").innerHTML='<article><span>Estado</span><strong>Sin datos</strong></article>';
    $("analysisTable").innerHTML='<tr><td colspan="5">Sin información procesada.</td></tr>';
    return;
  }

  const cards=[
    ["Costo total",money(t.cost2026)],
    ["Gasto total",money(t.expense2026)],
    ["XPV real",money(t.xpReal)],
    ["XPV presupuesto",money(t.xpBudget)],
    ["Cumplimiento XPV",(t.xpCompliance*100).toFixed(2)+"%"],
    ["Gerencias activas",t.managers]
  ];
  $("analysisCards").innerHTML=cards.map(x=>`<article><span>${x[0]}</span><strong>${x[1]}</strong></article>`).join("");

  // If there are critical deviations, show them.
  // Otherwise, show the most relevant real 2025 vs 2026 variations.
  let rows=t.deviations.slice();

  if(!rows.length && state.model){
    rows=[
      ...state.model.costSummary.map(x=>({
        type:"Costo",
        manager:x.manager,
        y2025:x.y2025,
        y2026:x.y2026,
        pct:x.pct2526,
        impact:Math.abs(x.diff2526||0)
      })),
      ...state.model.expenseSummary.map(x=>({
        type:"Gasto",
        manager:x.manager,
        y2025:x.y2025,
        y2026:x.y2026,
        pct:x.pct2526,
        impact:Math.abs(x.diff2526||0)
      }))
    ]
    .filter(x=>Number.isFinite(x.pct))
    .sort((a,b)=>Math.abs(b.pct)-Math.abs(a.pct));
  }

  if(!rows.length){
    $("analysisTable").innerHTML='<tr><td colspan="5">No hay variaciones calculables para el periodo.</td></tr>';
    return;
  }

  $("analysisTable").innerHTML=rows.slice(0,30).map(d=>{
    const p=Number.isFinite(d.pct)?d.pct:0;
    const cls=p>0?"var-up":p<0?"var-down":"var-flat";
    const label=p>0?"Aumento":p<0?"Reducción":"Sin cambio";
    return `<tr>
      <td>${d.type}</td>
      <td>${d.manager}</td>
      <td>${signedMoney(d.y2025)}</td>
      <td>${signedMoney(d.y2026)}</td>
      <td><span class="variation-pill ${cls}">${label} ${Math.abs(p*100).toFixed(2)}%</span></td>
    </tr>`;
  }).join("");
}
function renderComparisons(){
  if(!state.model){$("comparisonTable").innerHTML="";return;}
  const cost=new Map(state.model.costSummary.map(x=>[x.manager,x])),exp=new Map(state.model.expenseSummary.map(x=>[x.manager,x]));
  const managers=[...new Set([...cost.keys(),...exp.keys()])].sort();
  $("comparisonTable").innerHTML=managers.map(m=>{const c=cost.get(m)||{},e=exp.get(m)||{};return `<tr><td>${m}</td><td>${signedMoney(c.y2025||0)}</td><td>${signedMoney(c.y2026||0)}</td><td>${pctText(c.pct2526)}</td><td>${signedMoney(e.y2025||0)}</td><td>${signedMoney(e.y2026||0)}</td><td>${pctText(e.pct2526)}</td></tr>`;}).join("");
}
function renderManagers(){
  if(!state.model){$("managerTable").innerHTML="";return;}
  const cost=new Map(state.model.costSummary.map(x=>[x.manager,x.y2026])),exp=new Map(state.model.expenseSummary.map(x=>[x.manager,x.y2026]));
  const managers=[...new Set([...cost.keys(),...exp.keys()])].sort();
  $("managerTable").innerHTML=managers.map(m=>{const c=cost.get(m)||0,e=exp.get(m)||0,total=Math.abs(c)+Math.abs(e);return `<tr><td>${m}</td><td>${money(c)}</td><td>${money(e)}</td><td>${money(total)}</td><td><span class="pill low">Procesado</span></td></tr>`;}).join("");
}

function trendSummary(){
  if(!state.model)return null;
  const sum=(arr,k)=>arr.reduce((s,x)=>s+(x[k]||0),0);
  const cost25=sum(state.model.costSummary,"y2025"),cost26=sum(state.model.costSummary,"y2026");
  const exp25=sum(state.model.expenseSummary,"y2025"),exp26=sum(state.model.expenseSummary,"y2026");
  const costPct=pct(cost26-cost25,cost25),expPct=pct(exp26-exp25,exp25);
  const combined=[...state.model.costSummary.map(x=>({...x,type:"Costo"})),...state.model.expenseSummary.map(x=>({...x,type:"Gasto"}))].filter(x=>x.pct2526!=null);
  const best=combined.slice().sort((a,b)=>a.pct2526-b.pct2526)[0];
  const worst=combined.slice().sort((a,b)=>b.pct2526-a.pct2526)[0];
  return {cost25,cost26,exp25,exp26,costPct,expPct,best,worst,combined};
}
function renderTrends(){
  const t=trendSummary();
  if(!t){
    ["trendCostPct","trendExpensePct"].forEach(id=>$(id).textContent="0%");
    $("trendBestManager").textContent="—";$("trendWorstManager").textContent="—";
    $("trendMovers").innerHTML='<div class="empty">Procesa información para ver tendencias.</div>';
    drawTrendDetail();return;
  }
  $("trendCostPct").textContent=pctText(t.costPct);
  $("trendExpensePct").textContent=pctText(t.expPct);
  $("trendCostDir").textContent=t.costPct==null?"Sin base":(t.costPct>0?"↑ Aumenta vs 2025":"↓ Disminuye vs 2025");
  $("trendExpenseDir").textContent=t.expPct==null?"Sin base":(t.expPct>0?"↑ Aumenta vs 2025":"↓ Disminuye vs 2025");
  $("trendBestManager").textContent=t.best?.manager||"—";
  $("trendBestValue").textContent=t.best?`${t.best.type} ${pctText(t.best.pct2526)}`:"Sin datos";
  $("trendWorstManager").textContent=t.worst?.manager||"—";
  $("trendWorstValue").textContent=t.worst?`${t.worst.type} ${pctText(t.worst.pct2526)}`:"Sin datos";
  $("trendsSubtitle").textContent=state.inputMode==="JD"?"Tendencia mensual real desde los JD y comparativo 2025 vs 2026.":"Tendencia comparativa real 2025 vs 2026 a partir de los archivos RP.";
  const movers=t.combined.slice().sort((a,b)=>Math.abs(b.pct2526)-Math.abs(a.pct2526)).slice(0,8);
  $("trendMovers").innerHTML=movers.map(x=>`<div class="trend-mover ${x.pct2526>0?"up":"down"}"><span>${x.type}</span><strong>${x.manager}</strong><em>${pctText(x.pct2526)}</em></div>`).join("");
  drawTrendDetail();
}
function drawTrendDetail(){
  const c=$("trendDetailChart");if(!c)return;
  const ctx=c.getContext("2d"),w=c.width,h=c.height;ctx.clearRect(0,0,w,h);
  ctx.strokeStyle="#e8eef7";for(let i=1;i<5;i++){let y=i*h/5;ctx.beginPath();ctx.moveTo(50,y);ctx.lineTo(w-20,y);ctx.stroke();}
  if(!state.model){ctx.fillStyle="#94a3b8";ctx.font="16px system-ui";ctx.textAlign="center";ctx.fillText("Sin datos procesados",w/2,h/2);return;}
  if(state.inputMode==="JD"){
    $("trendMainTitle").textContent="Evolución mensual acumulada";
    const upto=state.periodIndex+1;
    const cost=state.model.costMonthly.slice(0,upto).map(x=>Math.abs(x.y2026));
    const exp=state.model.expenseMonthly.slice(0,upto).map(x=>Math.abs(x.y2026));
    const max=Math.max(...cost,...exp,1);
    [[cost,"#2563eb"],[exp,"#a855f7"]].forEach(([arr,color])=>{
      ctx.strokeStyle=color;ctx.lineWidth=4;ctx.beginPath();
      arr.forEach((v,i)=>{let x=60+i*(w-110)/Math.max(1,arr.length-1),y=h-40-v/max*(h-85);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.stroke();
    });
  }else{
    $("trendMainTitle").textContent="Comparativo 2025 vs 2026 por gerencia";
    const rows=trendSummary().combined.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026)).slice(0,8);
    const max=Math.max(...rows.flatMap(r=>[Math.abs(r.y2025),Math.abs(r.y2026)]),1);
    const gw=(w-100)/Math.max(1,rows.length);
    rows.forEach((r,i)=>{
      const x=55+i*gw,h25=Math.abs(r.y2025)/max*(h-90),h26=Math.abs(r.y2026)/max*(h-90);
      ctx.fillStyle="#14b8a6";ctx.fillRect(x,h-38-h25,gw*.28,h25);
      ctx.fillStyle="#f97316";ctx.fillRect(x+gw*.32,h-38-h26,gw*.28,h26);
      ctx.fillStyle="#64748b";ctx.font="9px system-ui";ctx.textAlign="center";ctx.fillText((r.manager||"").slice(0,10),x+gw*.28,h-12);
    });
  }
}


function auditMetrics(){
  if(!state.model)return null;
  const cost=(state.model.costSummary||[]).length, exp=(state.model.expenseSummary||[]).length, xpv=(state.model.xpvSummary||[]).length;
  const rows=cost+exp+xpv;
  const issues=[];
  if(!cost)issues.push("Sin registros de costo");
  if(!exp)issues.push("Sin registros de gasto");
  if(!xpv)issues.push("Sin registros de productividad XPV");
  const totals=modelTotals();
  if(totals && !Number.isFinite(totals.cost2026))issues.push("Costo total no calculable");
  if(totals && !Number.isFinite(totals.expense2026))issues.push("Gasto total no calculable");
  const sources=state.inputMode==="JD"?1:3;
  const confidence=Math.max(0,Math.min(100,100-issues.length*20));
  return {cost,exp,xpv,rows,issues,sources,confidence};
}
function renderAudit(){
  const a=auditMetrics();
  if(!a){$("auditConfidence").textContent="0%";$("auditStatus").textContent="Sin validar";$("auditSources").textContent="0";$("auditRows").textContent="0";$("auditIssues").textContent="0";$("auditChecks").innerHTML='<div class="empty">Procesa información para ejecutar la validación.</div>';$("auditLog").innerHTML='<div class="empty">Sin procesamiento registrado.</div>';return;}
  $("auditConfidence").textContent=a.confidence+"%";$("auditStatus").textContent=a.confidence===100?"Procesamiento validado":a.confidence>=80?"Validado con observaciones":"Requiere revisión";$("auditSources").textContent=a.sources;$("auditRows").textContent=a.rows.toLocaleString();$("auditIssues").textContent=a.issues.length;
  const checks=[
    ["Fuente de costos",a.cost>0,`${a.cost} registros`],
    ["Fuente de gastos",a.exp>0,`${a.exp} registros`],
    ["Fuente XPV",a.xpv>0,`${a.xpv} registros`],
    ["Modelo ejecutivo",a.rows>0,"Indicadores calculados desde datos procesados"]
  ];
  $("auditChecks").innerHTML=checks.map(x=>`<div class="audit-check ${x[1]?"ok":"bad"}"><b>${x[1]?"✓":"!"}</b><div><strong>${x[0]}</strong><span>${x[2]}</span></div></div>`).join("");
  $("auditLog").innerHTML=`<div class="log-row"><span>Modo de entrada</span><strong>${state.inputMode||"—"}</strong></div><div class="log-row"><span>Periodo</span><strong>${MONTH_LABELS[state.periodIndex]||"Periodo actual"}</strong></div><div class="log-row"><span>Resultado</span><strong>${a.confidence===100?"Validado":"Con observaciones"}</strong></div><div class="log-row"><span>Incidencias</span><strong>${a.issues.length?a.issues.join(", "):"Ninguna detectada"}</strong></div>`;
}
function managerHealth(){
  if(!state.model)return [];
  const map={};
  const add=(arr,type)=>arr.forEach(x=>{const k=x.manager||"SIN CLASIFICAR";map[k]??={manager:k,cost:null,expense:null,xpv:null};map[k][type]=x;});
  add(state.model.costSummary||[],"cost");add(state.model.expenseSummary||[],"expense");add(state.model.xpvSummary||[],"xpv");
  return Object.values(map).map(m=>{
    const cp=m.cost?.pct2526??0, ep=m.expense?.pct2526??0;
    const xp=m.xpv?(m.xpv.budget?((m.xpv.real-m.xpv.budget)/Math.abs(m.xpv.budget)):0):0;
    const pressure=(Math.max(0,cp)+Math.max(0,ep)+Math.max(0,-xp))/3;
    const status=pressure>.15?"Crítico":pressure>.05?"Atención":"Favorable";
    return {...m,cp,ep,xp,pressure,status};
  }).sort((a,b)=>b.pressure-a.pressure);
}
function renderSemaphore(){
  const rows=managerHealth();
  if(!rows.length){$("semaphoreTable").innerHTML='<tr><td colspan="6">Sin datos procesados.</td></tr>';return;}
  $("semaphoreTable").innerHTML=rows.map(r=>`<tr><td><strong>${r.manager}</strong></td><td>${r.cost?pctText(r.cp):"—"}</td><td>${r.expense?pctText(r.ep):"—"}</td><td>${r.xpv?pctText(r.xp):"—"}</td><td>${pctText(r.pressure)}</td><td><span class="health ${r.status==="Crítico"?"red":r.status==="Atención"?"yellow":"green"}">${r.status}</span></td></tr>`).join("");
}
function actionItems(){
  const rows=managerHealth(); const out=[];
  rows.slice(0,12).forEach(r=>{
    if(r.status==="Crítico")out.push({priority:"Alta",manager:r.manager,impact:pctText(r.pressure),finding:"Variaciones combinadas requieren revisión prioritaria.",action:"Revisar los rubros con mayor incremento y documentar causa raíz."});
    else if(r.status==="Atención")out.push({priority:"Media",manager:r.manager,impact:pctText(r.pressure),finding:"Se detecta presión moderada en los indicadores.",action:"Dar seguimiento al siguiente periodo y validar desviaciones relevantes."});
    else out.push({priority:"Oportunidad",manager:r.manager,impact:pctText(r.pressure),finding:"Comportamiento dentro de parámetros favorables.",action:"Identificar prácticas replicables y mantener seguimiento."});
  }); return out;
}
function renderActions(){
  const items=actionItems();
  const hi=items.filter(x=>x.priority==="Alta").length, med=items.filter(x=>x.priority==="Media").length, opp=items.filter(x=>x.priority==="Oportunidad").length;
  $("actionHigh").textContent=hi;$("actionMedium").textContent=med;$("actionOpportunities").textContent=opp;
  $("actionList").innerHTML=items.length?items.map(x=>`<div class="action-item ${x.priority.toLowerCase()}"><div><span>${x.priority}</span><strong>${x.manager}</strong></div><p>${x.finding}</p><em>Impacto ${x.impact}</em><b>${x.action}</b></div>`).join(""):'<div class="empty">Procesa información para generar el plan de acción.</div>';
}


function allManagers(){
  if(!state.model)return [];
  return [...new Set([...(state.model.costSummary||[]).map(x=>x.manager),...(state.model.expenseSummary||[]).map(x=>x.manager),...(state.model.xpvSummary||[]).map(x=>x.manager)].filter(Boolean))].sort();
}
function managerRecord(name){
  if(!state.model||!name)return null;
  return {name,cost:(state.model.costSummary||[]).find(x=>x.manager===name)||null,expense:(state.model.expenseSummary||[]).find(x=>x.manager===name)||null,xpv:(state.model.xpvSummary||[]).find(x=>x.manager===name)||null};
}
function managerStatus(rec){
  const cp=rec?.cost?.pct2526??0,ep=rec?.expense?.pct2526??0,xp=rec?.xpv?.budget?((rec.xpv.real-rec.xpv.budget)/Math.abs(rec.xpv.budget)):0;
  const pressure=(Math.max(0,cp)+Math.max(0,ep)+Math.max(0,-xp))/3;
  return pressure>.15?{label:"Crítico",cls:"red",pressure}:pressure>.05?{label:"Atención",cls:"yellow",pressure}:{label:"Favorable",cls:"green",pressure};
}
function renderManager360Options(){
  const el=$("manager360Select"); if(!el)return;
  el.innerHTML='<option value="">Selecciona una gerencia</option>'+allManagers().map(m=>`<option value="${m}">${m}</option>`).join("");
}
function renderManager360(name){
  const rec=managerRecord(name);
  if(!rec){$("manager360Content").hidden=true;$("manager360Empty").hidden=false;return;}
  $("manager360Content").hidden=false;$("manager360Empty").hidden=true;
  $("m360Name").textContent=name;
  const st=managerStatus(rec);$("m360Status").textContent=st.label;$("m360Status").className="manager360-status "+st.cls;
  const c26=rec.cost?.y2026||0,e26=rec.expense?.y2026||0,xreal=rec.xpv?.real||0,xbudget=rec.xpv?.budget||0;
  $("m360Cost").textContent=money(c26);$("m360Expense").textContent=money(e26);$("m360XpvReal").textContent=money(xreal);$("m360XpvBudget").textContent="Ptto "+money(xbudget);$("m360Impact").textContent=money(Math.abs(c26)+Math.abs(e26));
  $("m360CostVar").textContent=rec.cost?pctText(rec.cost.pct2526):"—";$("m360ExpenseVar").textContent=rec.expense?pctText(rec.expense.pct2526):"—";$("m360Summary").textContent=`Estado ${st.label}. Impacto combinado ${money(Math.abs(c26)+Math.abs(e26))}.`;
  const xp=xbudget?xreal/xbudget:0;
  $("m360Diagnosis").innerHTML=[
    rec.cost?`Costo ${rec.cost.pct2526>0?"aumenta":"disminuye"} ${Math.abs((rec.cost.pct2526||0)*100).toFixed(2)}% vs 2025.`:"Sin información de costo.",
    rec.expense?`Gasto ${rec.expense.pct2526>0?"aumenta":"disminuye"} ${Math.abs((rec.expense.pct2526||0)*100).toFixed(2)}% vs 2025.`:"Sin información de gasto.",
    rec.xpv?`Cumplimiento XPV ${(xp*100).toFixed(2)}% del presupuesto.`:"Sin información XPV."
  ].map(x=>`<div class="diag-item">${x}</div>`).join("");
  const rows=[...(state.model.costOperating||[]).filter(x=>x.manager===name).map(x=>({...x,type:"Costo"})),...(state.model.costMaintenance||[]).filter(x=>x.manager===name).map(x=>({...x,type:"Costo"})),...(state.model.expenses||[]).filter(x=>x.manager===name).map(x=>({...x,type:"Gasto"}))]
    .map(x=>({...x,impact:Math.abs((x.y2026||0)-(x.y2025||0))})).sort((a,b)=>b.impact-a.impact).slice(0,8);
  $("m360Drivers").innerHTML=rows.length?rows.map((x,i)=>`<div class="driver-row"><b>${i+1}</b><div><strong>${x.account||x.agrupador||x.type}</strong><span>${x.type} · ${signedMoney(x.y2025||0)} → ${signedMoney(x.y2026||0)}</span></div><em>${money(x.impact)}</em></div>`).join(""):'<div class="empty">No hay detalle disponible.</div>';
  const recs=[];if((rec.cost?.pct2526||0)>.1)recs.push("Revisar los conceptos que explican el incremento de costo.");if((rec.expense?.pct2526||0)>.1)recs.push("Validar gastos recurrentes y variaciones extraordinarias.");if(rec.xpv&&xp<1)recs.push("Analizar productividad XPV y su relación con costo/gasto.");if(!recs.length)recs.push("Mantener seguimiento y documentar prácticas favorables.");
  $("m360Recommendations").innerHTML=recs.map((x,i)=>`<div class="recommend-row"><b>${i+1}</b><span>${x}</span></div>`).join("");
  const c=$("m360Chart"),ctx=c.getContext("2d"),w=c.width,h=c.height;ctx.clearRect(0,0,w,h);
  const data=[{label:"Costo",v25:Math.abs(rec.cost?.y2025||0),v26:Math.abs(rec.cost?.y2026||0),c1:"#60a5fa",c2:"#2563eb"},{label:"Gasto",v25:Math.abs(rec.expense?.y2025||0),v26:Math.abs(rec.expense?.y2026||0),c1:"#c084fc",c2:"#7c3aed"}];
  const max=Math.max(...data.flatMap(x=>[x.v25,x.v26]),1);ctx.strokeStyle="#e8eef7";for(let i=1;i<5;i++){let y=i*h/5;ctx.beginPath();ctx.moveTo(50,y);ctx.lineTo(w-20,y);ctx.stroke();}
  data.forEach((d,i)=>{const x=140+i*280,bw=70,h25=d.v25/max*(h-100),h26=d.v26/max*(h-100);ctx.fillStyle=d.c1;ctx.fillRect(x,h-45-h25,bw,h25);ctx.fillStyle=d.c2;ctx.fillRect(x+bw+18,h-45-h26,bw,h26);ctx.fillStyle="#475569";ctx.font="14px system-ui";ctx.textAlign="center";ctx.fillText(d.label,x+bw+9,h-15);});
}
function lineageHtml(type){
  if(!state.model)return '<div class="empty">Procesa información para ver la trazabilidad.</div>';
  const t=modelTotals();
  if(type==="cost")return `<div class="lineage-kpi"><span>INDICADOR</span><strong>${money(t.cost2026)}</strong><p>Costo total 2026</p></div><div class="lineage-step"><b>1</b><div><strong>Fuente</strong><span>${state.inputMode==="JD"?"JD COSTO RCV":"Archivo RP COSTO RCV"}</span></div></div><div class="lineage-step"><b>2</b><div><strong>Registros incluidos</strong><span>${((state.model.costOperating||[]).length+(state.model.costMaintenance||[]).length).toLocaleString()}</span></div></div><div class="lineage-step"><b>3</b><div><strong>Cálculo</strong><span>Costos Operativos + Costo Mantto</span></div></div><div class="lineage-step"><b>4</b><div><strong>Resultado</strong><span>${money(t.cost2026)}</span></div></div>`;
  if(type==="expense")return `<div class="lineage-kpi"><span>INDICADOR</span><strong>${money(t.expense2026)}</strong><p>Gasto total 2026</p></div><div class="lineage-step"><b>1</b><div><strong>Fuente</strong><span>${state.inputMode==="JD"?"JD GASTOS RCV":"Archivo RP Gastos RCV"}</span></div></div><div class="lineage-step"><b>2</b><div><strong>Registros incluidos</strong><span>${(state.model.expenses||[]).length.toLocaleString()}</span></div></div><div class="lineage-step"><b>3</b><div><strong>Cálculo</strong><span>Suma del modelo de gastos 2026</span></div></div><div class="lineage-step"><b>4</b><div><strong>Resultado</strong><span>${money(t.expense2026)}</span></div></div>`;
  return `<div class="lineage-kpi"><span>INDICADOR</span><strong>${(t.xpCompliance*100).toFixed(2)}%</strong><p>Cumplimiento XPV</p></div><div class="lineage-step"><b>1</b><div><strong>Fuente</strong><span>${state.inputMode==="JD"?"JD XPV RCV":"Archivo RP Productividad XPV"}</span></div></div><div class="lineage-step"><b>2</b><div><strong>Real</strong><span>${money(t.xpReal)}</span></div></div><div class="lineage-step"><b>3</b><div><strong>Presupuesto</strong><span>${money(t.xpBudget)}</span></div></div><div class="lineage-step"><b>4</b><div><strong>Fórmula</strong><span>Real / Presupuesto × 100</span></div></div>`;
}

function renderAll(){renderDashboard();renderFilters();renderProcessStats();renderAnalysis();renderComparisons();renderManagers();renderTrends();buildExecutiveIntelligence();renderAudit();renderSemaphore();renderActions();renderManager360Options();}

function tableSheet(title, headers, data, period) {
  const aoa=[["REPORT.IA RCV"],[title],["Periodo",period],[],headers,...data];
  const ws=XLSX.utils.aoa_to_sheet(aoa);
  ws["!cols"]=headers.map((h,i)=>({wch:i<5?30:16}));
  return ws;
}
function commonRows(rows, order) {
  return rows.map(r=>order.map(k=>r[k]));
}
function costDataRows(rows, variant) {
  if(variant==="operating") return rows.map(r=>[r.agrupador,r.brand,r.account,r.manager,r.center,r.y2025,r.budget2026,r.y2026,r.diff2526,r.pct2526??"",r.diffBudget,r.pctBudget??""]);
  if(variant==="premises") return rows.map(r=>[r.account,r.brand,r.agrupador,r.center,r.manager,r.y2025,r.budget2026,r.y2026,r.diff2526,r.pct2526??"",r.diffBudget,r.pctBudget??""]);
  return rows.map(r=>[r.manager,r.agrupador,r.account,r.brand,r.center,r.y2025,r.budget2026,r.y2026,r.diff2526,r.pct2526??"",r.diffBudget,r.pctBudget??""]);
}
function makeCostWorkbook(){
  const m=state.model;if(!m)throw new Error("No hay datos procesados.");
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,tableSheet("COSTOS OPERATIVOS",["AGRUPADOR","CIA MARCA","CTA CONTABLE","GERENCIA","CENTRO DE GESTION","2025","PTTO. 2026","2026","25 vs 26 $","25 vs 26 %","PTTO vs 26 $","PTTO vs 26 %"],costDataRows(m.costOperating,"operating"),m.month),"COSTOS OPERATIVOS");
  XLSX.utils.book_append_sheet(wb,tableSheet("COSTO MANTTO",["GERENCIA","AGRUPADOR","CTA CONTABLE","CIA MARCA","CENTRO DE GESTION","2025","PTTO. 2026","2026","25 vs 26 $","25 vs 26 %","PTTO vs 26 $","PTTO vs 26 %"],costDataRows(m.costMaintenance,"standard"),m.month),"COSTO MANTTO");
  XLSX.utils.book_append_sheet(wb,tableSheet("PREMISAS 2026",["CTA CONTABLE","CIA MARCA","AGRUPADOR","CENTRO DE GESTION","GERENCIA","2025","PTTO. 2026","2026","25 vs 26 $","25 vs 26 %","PTTO vs 26 $","PTTO vs 26 %"],costDataRows(m.premises,"premises"),m.month),"premisas 26");
  XLSX.utils.book_append_sheet(wb,tableSheet("SMO",["GERENCIA","AGRUPADOR","CTA CONTABLE","CIA MARCA","CENTRO DE GESTION","2025","PTTO. 2026","2026","25 vs 26 $","25 vs 26 %","PTTO vs 26 $","PTTO vs 26 %"],costDataRows(m.smo,"standard"),m.month),"SMO");
  return wb;
}
function makeExpenseWorkbook(){
  const m=state.model;if(!m)throw new Error("No hay datos procesados.");
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,tableSheet("GASTOS RCV",["GERENCIA","AGRUPADOR","CTA CONTABLE","CIA MARCA","CENTRO DE GESTION","2025","PTTO. 2026","2026","25 vs 26 $","25 vs 26 %","PTTO vs 26 $","PTTO vs 26 %"],costDataRows(m.expenses,"standard"),m.month),"GASTOS RCV");
  const summary=m.expenseSummary.map(r=>[r.manager,r.y2025,r.budget2026,r.y2026,r.diff2526,r.pct2526??"",r.diffBudget,r.pctBudget??""]);
  XLSX.utils.book_append_sheet(wb,tableSheet("RESUMEN POR GERENCIA",["GERENCIA","2025","PTTO. 2026","2026","25 vs 26 $","25 vs 26 %","PTTO vs 26 $","PTTO vs 26 %"],summary,m.month),"RESUMEN GERENCIA");
  return wb;
}
function makeXpvWorkbook(){
  const m=state.model;if(!m)throw new Error("No hay datos procesados.");
  const wb=XLSX.utils.book_new();
  const data=m.xpv.map(r=>[r.manager,r.account,r.center,r.brand,r.real,r.budget,r.diff,r.pct??""]);
  XLSX.utils.book_append_sheet(wb,tableSheet("PRODUCTIVIDAD XPV RCV",["GERENCIA","CTA CONTABLE","CENTRO / SUBLIBRO","CIA MARCA","REAL GESTION","PRESUPUESTO","DIFERENCIA $","DIFERENCIA %"],data,m.month),"PRODUCTIVIDAD XPV");
  return wb;
}
function workbookBlob(wb){return new Blob([XLSX.write(wb,{bookType:"xlsx",type:"array"})],{type:"application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"});}
function saveBlob(blob,name){const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(a.href),1500);}
async function downloadAll(){
  if(!state.model){alert("Primero procesa los archivos JD.");return;}
  const zip=new JSZip();
  zip.file("COSTO RCV 2026.xlsx",workbookBlob(makeCostWorkbook()));
  zip.file("Gastos RCV 2026.xlsx",workbookBlob(makeExpenseWorkbook()));
  zip.file(`Productividad XPV ${state.model.month} 26 RCV.xlsx`,workbookBlob(makeXpvWorkbook()));
  zip.file("LEEME.txt","REPORT.IA RCV 23.0 · Archivos generados directamente desde los JD procesados.");
  saveBlob(await zip.generateAsync({type:"blob"}),`FINAL_${state.model.month}_2026.zip`);
}

function reportHtml(type){
  if(!state.model)return '<div class="empty">Primero procesa los archivos JD.</div>';
  const t=modelTotals(),dev=t.deviations;
  const table=(heads,rows)=>`<table><thead><tr>${heads.map(x=>`<th>${x}</th>`).join("")}</tr></thead><tbody>${rows.map(r=>`<tr>${r.map(x=>`<td>${x}</td>`).join("")}</tr>`).join("")}</tbody></table>`;
  if(type==="deviations") return `<h3>Desviaciones y Excepciones</h3>${table(["Tipo","Gerencia","2025","2026","Variación"],dev.map(d=>[d.type,d.manager,signedMoney(d.y2025),signedMoney(d.y2026),(d.pct*100).toFixed(2)+"%"]))}`;
  if(type==="savings"){
    const rows=[...state.model.costSummary.map(x=>({...x,type:"Costo"})),...state.model.expenseSummary.map(x=>({...x,type:"Gasto"}))].filter(x=>Math.abs(x.y2026)<Math.abs(x.y2025)).map(x=>[x.type,x.manager,money(x.y2025),money(x.y2026),money(Math.abs(x.y2025)-Math.abs(x.y2026))]).sort((a,b)=>parseFloat(b[4].replace(/[^\d.]/g,""))-parseFloat(a[4].replace(/[^\d.]/g,"")));
    return `<h3>Ahorro Observado</h3><p>Reducciones reales 2026 contra 2025; no son proyecciones.</p>${table(["Tipo","Gerencia","2025","2026","Reducción observada"],rows)}`;
  }
  if(type==="variations") return `<h3>Variaciones 2025 vs 2026</h3>${table(["Tipo","Gerencia","2025","2026","Δ %"],[...state.model.costSummary.map(x=>["Costo",x.manager,money(x.y2025),money(x.y2026),pctText(x.pct2526)]),...state.model.expenseSummary.map(x=>["Gasto",x.manager,money(x.y2025),money(x.y2026),pctText(x.pct2526)])])}`;
  if(type==="pareto"){
    const map=new Map();[...state.model.costSummary,...state.model.expenseSummary].forEach(x=>map.set(x.manager,(map.get(x.manager)||0)+Math.abs(x.y2026)));const arr=[...map.entries()].sort((a,b)=>b[1]-a[1]);const total=arr.reduce((s,x)=>s+x[1],0);let acc=0;return `<h3>Pareto 80/20</h3>${table(["Gerencia","Impacto","Participación","Acumulado"],arr.map(([m,v])=>{acc+=v;return[m,money(v),(v/total*100).toFixed(1)+"%",(acc/total*100).toFixed(1)+"%"]}))}`;
  }
  if(type==="efficiency"){
    const cost=new Map(state.model.costSummary.map(x=>[x.manager,Math.abs(x.y2026)])),exp=new Map(state.model.expenseSummary.map(x=>[x.manager,Math.abs(x.y2026)]));const ms=[...new Set([...cost.keys(),...exp.keys()])];return `<h3>Eficiencia Gerencial</h3>${table(["Gerencia","Costos","Gastos","Impacto total"],ms.map(m=>[m,money(cost.get(m)||0),money(exp.get(m)||0),money((cost.get(m)||0)+(exp.get(m)||0))]).sort((a,b)=>parseFloat(b[3].replace(/[^\d.]/g,""))-parseFloat(a[3].replace(/[^\d.]/g,""))))}`;
  }
  if(type==="actions"){
    if(!state.decisions.length)state.decisions=dev.slice(0,5).map(d=>({finding:`${d.type}: ${d.manager}`,action:"Analizar causa y definir plan",owner:"Por asignar",status:"Abierta"}));
    localStorage.setItem("reportia_decisions_v19",JSON.stringify(state.decisions));
    return `<h3>Plan de Acción Ejecutivo</h3>${table(["Hallazgo","Acción","Responsable","Estatus"],state.decisions.map(x=>[x.finding,x.action,x.owner,x.status]))}`;
  }
  if(type==="master") return `<h3>Reporte Maestro Ejecutivo</h3><p><b>Periodo:</b> ${state.model.month} 2026</p><p><b>Costo total:</b> ${money(t.cost2026)} · <b>Gasto total:</b> ${money(t.expense2026)} · <b>Cumplimiento XPV:</b> ${(t.xpCompliance*100).toFixed(2)}% · <b>Alertas:</b> ${new Set(dev.map(x=>x.manager)).size}</p>${reportHtml("deviations")}`;
  return '<div class="empty">Reporte no disponible.</div>';
}

async function validateAgainstFinal(){
  if(!state.model||!state.finalZipFile)return;
  setStatus("validationStatus","Validando contra FINAL.zip…","working");
  try{
    const zip=await JSZip.loadAsync(await state.finalZipFile.arrayBuffer());
    const files=Object.values(zip.files).filter(f=>!f.dir&&f.name.toLowerCase().endsWith(".xlsx"));
    const refs={};
    for(const f of files){const ab=await f.async("arraybuffer");refs[f.name]=XLSX.read(ab,{type:"array",raw:true});}
    const checks=[];
    const compare=(label,genCount,wbName,sheetName,totalCol,genTotal)=>{
      const entry=Object.entries(refs).find(([n])=>n.toLowerCase().includes(wbName.toLowerCase()));
      if(!entry){checks.push([label,false,"Archivo no encontrado"]);return;}
      const ws=entry[1].Sheets[sheetName];if(!ws){checks.push([label,false,"Hoja no encontrada"]);return;}
      const arr=XLSX.utils.sheet_to_json(ws,{header:1,defval:"",raw:true});
      const data=arr.slice(5).filter(r=>r.some(v=>cleanText(v)!==""));
      const refTotal=data.reduce((s,r)=>s+parseNumber(r[totalCol]),0);
      checks.push([label,data.length===genCount&&Math.abs(refTotal-genTotal)<0.02,`Filas ${genCount}/${data.length} · Total ${signedMoney(genTotal)}/${signedMoney(refTotal)}`]);
    };
    compare("COSTOS OPERATIVOS",state.model.costOperating.length,"COSTO RCV","COSTOS OPERATIVOS",7,state.model.costOperating.reduce((s,r)=>s+r.y2026,0));
    compare("COSTO MANTTO",state.model.costMaintenance.length,"COSTO RCV","COSTO MANTTO",7,state.model.costMaintenance.reduce((s,r)=>s+r.y2026,0));
    compare("GASTOS RCV",state.model.expenses.length,"Gastos RCV","GASTOS RCV",7,state.model.expenses.reduce((s,r)=>s+r.y2026,0));
    compare("PRODUCTIVIDAD XPV",state.model.xpv.length,"Productividad XPV","PRODUCTIVIDAD XPV",4,state.model.xpv.reduce((s,r)=>s+r.real,0));
    const ok=checks.every(x=>x[1]);
    setStatus("validationStatus",(ok?"✓ Validación correcta. ":"⚠ Se encontraron diferencias. ")+checks.map(x=>`${x[0]}: ${x[2]}`).join(" | "),ok?"ok":"error");
  }catch(e){console.error(e);setStatus("validationStatus","No fue posible validar el ZIP: "+e.message,"error");}
}

function answer(q){
  if(!state.model)return "Todavía no hay información procesada.";
  const t=modelTotals(),l=q.toLowerCase();
  if(/resumen/.test(l))return `Periodo ${state.model.month} 2026: costo ${money(t.cost2026)}, gasto ${money(t.expense2026)}, cumplimiento XPV ${(t.xpCompliance*100).toFixed(2)}% y ${new Set(t.deviations.map(d=>d.manager)).size} gerencias con alertas.`;
  if(/riesgo/.test(l)){const d=t.deviations[0];return d?`La mayor desviación es ${d.type} en ${d.manager}: ${(d.pct*100).toFixed(2)}% respecto a 2025.`:"No hay desviaciones mayores al umbral configurado.";}
  if(/mayor gasto|gerencia.*gasto/.test(l)){const x=state.model.expenseSummary.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026))[0];return `${x.manager} registra el mayor gasto 2026: ${money(x.y2026)}.`;}
  if(/productividad|xpv/.test(l))return `XPV real ${money(t.xpReal)} contra presupuesto ${money(t.xpBudget)}; cumplimiento ${(t.xpCompliance*100).toFixed(2)}%.`;
  if(/junta|mencionar|discurso/.test(l))return prepareSpeech();
  if(/80%|pareto|concentra/.test(l))return paretoAnswer();
  return "Puedo responder sobre resumen, principal riesgo, mayor gasto y productividad XPV usando el modelo procesado.";
}


function buildExecutiveIntelligence(){
  if(!state.model){
    $("directorBrief").className="exec-content empty";
    $("directorBrief").textContent="Procesa los archivos para generar el brief ejecutivo.";
    ["execConclusions","execRisks","execActions"].forEach(id=>$(id).innerHTML="");
    $("alertsList").innerHTML='<div class="empty">Sin alertas disponibles.</div>';
    $("alertsCount").textContent="0";
    return;
  }
  const t=modelTotals();
  const dev=t.deviations;
  const topCost=state.model.costSummary.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026))[0];
  const topExpense=state.model.expenseSummary.slice().sort((a,b)=>Math.abs(b.y2026)-Math.abs(a.y2026))[0];
  const positive=[...state.model.costSummary,...state.model.expenseSummary].filter(x=>x.pct2526!=null&&x.pct2526<0).sort((a,b)=>a.pct2526-b.pct2526)[0];

  const conclusions=[
    `El cumplimiento XPV del periodo es ${(t.xpCompliance*100).toFixed(2)}%.`,
    `${topCost?.manager||"—"} concentra el mayor costo: ${money(topCost?.y2026||0)}.`,
    `${topExpense?.manager||"—"} registra el mayor gasto: ${money(topExpense?.y2026||0)}.`
  ];
  const risks=[
    dev[0]?`${dev[0].manager} presenta la mayor desviación: ${(dev[0].pct*100).toFixed(2)}%.`:"No hay desviaciones superiores al umbral.",
    `${new Set(dev.map(d=>d.manager)).size} gerencias requieren atención por variación adversa.`,
    t.xpCompliance<1?`XPV se encuentra por debajo de presupuesto en ${((1-t.xpCompliance)*100).toFixed(2)} puntos porcentuales.`:"XPV cumple o supera el presupuesto."
  ];
  const actions=[
    dev[0]?`Revisar causa raíz de ${dev[0].type.toLowerCase()} en ${dev[0].manager}.`:"Mantener monitoreo de variaciones.",
    `Analizar el 80% del impacto concentrado antes de ampliar la revisión.`,
    `Asignar responsable y fecha compromiso a las tres principales desviaciones.`
  ];

  $("execConclusions").innerHTML=conclusions.map(x=>`<div class="mini-item">${x}</div>`).join("");
  $("execRisks").innerHTML=risks.map(x=>`<div class="mini-item">${x}</div>`).join("");
  $("execActions").innerHTML=actions.map(x=>`<div class="mini-item">${x}</div>`).join("");

  $("directorBrief").className="exec-content";
  $("directorBrief").innerHTML=`<strong>Resumen Ejecutivo — ${state.model.month} 2026</strong><br><br>
  El periodo presenta un cumplimiento XPV de <b>${(t.xpCompliance*100).toFixed(2)}%</b>.
  Se identificaron <b>${new Set(dev.map(d=>d.manager)).size} gerencias con alertas</b>.
  La mayor presión en costos se concentra en <b>${topCost?.manager||"—"}</b> con ${money(topCost?.y2026||0)},
  mientras que la mayor presión en gastos corresponde a <b>${topExpense?.manager||"—"}</b> con ${money(topExpense?.y2026||0)}.
  ${positive?`Como señal positiva, ${positive.manager} muestra una reducción de ${Math.abs(positive.pct2526*100).toFixed(2)}% respecto a 2025.`:""}
  <br><br><b>Recomendación:</b> priorizar las tres mayores desviaciones, validar su causa y asignar responsables de seguimiento.`;

  const alerts=[];
  dev.slice(0,5).forEach((d,i)=>alerts.push({type:i===0?"critical":"warning",title:`${d.type} · ${d.manager}`,text:`Variación ${(d.pct*100).toFixed(2)}% vs 2025.`}));
  if(positive)alerts.push({type:"positive",title:`Mejora · ${positive.manager}`,text:`Reducción ${Math.abs(positive.pct2526*100).toFixed(2)}% vs 2025.`});
  alerts.push({type:"info",title:"Productividad XPV",text:`Cumplimiento ${(t.xpCompliance*100).toFixed(2)}% del presupuesto.`});
  $("alertsList").innerHTML=alerts.map(a=>`<div class="alert-item"><i class="alert-dot ${a.type}"></i><div><strong>${a.title}</strong><span>${a.text}</span></div></div>`).join("");
  $("alertsCount").textContent=alerts.length;

  $("slideHeadline").textContent=`${state.model.month} 2026 · Estado Ejecutivo`;
  $("slideSummary").textContent=`Cumplimiento XPV ${(t.xpCompliance*100).toFixed(2)}%, ${new Set(dev.map(d=>d.manager)).size} alertas gerenciales y foco principal en ${dev[0]?.manager||"sin desviaciones críticas"}.`;
  $("slideCost").textContent=money(t.cost2026);$("slideExpense").textContent=money(t.expense2026);$("slideXpv").textContent=(t.xpCompliance*100).toFixed(1)+"%";$("slideAlerts").textContent=new Set(dev.map(d=>d.manager)).size;
  $("slideRisks").innerHTML=risks.map(x=>`<div class="presentation-item">${x}</div>`).join("");
  $("slideFindings").innerHTML=conclusions.map(x=>`<div class="presentation-item">${x}</div>`).join("");
  $("slideActions").innerHTML=actions.map(x=>`<div class="presentation-item">${x}</div>`).join("");
}
let currentSlide=0;
function showSlide(n){
  const slides=qa(".presentation-slide");
  currentSlide=(n+slides.length)%slides.length;
  slides.forEach((s,i)=>s.classList.toggle("active",i===currentSlide));
  $("slideIndicator").textContent=`${currentSlide+1} / ${slides.length}`;
}
function prepareSpeech(){
  if(!state.model)return "Procesa los archivos para generar el discurso.";
  const t=modelTotals(),d=t.deviations[0];
  return `Buenos días. Para el periodo ${state.model.month} 2026, el cumplimiento de productividad XPV es ${(t.xpCompliance*100).toFixed(2)}%. El costo total acumulado es ${money(t.cost2026)} y el gasto total ${money(t.expense2026)}. ${d?`Nuestra principal desviación se encuentra en ${d.manager}, con un incremento de ${(d.pct*100).toFixed(2)}%.`:"No identificamos desviaciones superiores al umbral configurado."} La recomendación es concentrar la revisión en las tres mayores desviaciones y asignar responsables para su seguimiento.`;
}
function paretoAnswer(){
  if(!state.model)return "Procesa los archivos primero.";
  const map=new Map();[...state.model.costSummary,...state.model.expenseSummary].forEach(x=>map.set(x.manager,(map.get(x.manager)||0)+Math.abs(x.y2026)));
  const arr=[...map.entries()].sort((a,b)=>b[1]-a[1]),total=arr.reduce((s,x)=>s+x[1],0);let acc=0,names=[];
  for(const [m,v] of arr){if(acc/total<.8){names.push(m);acc+=v;}}
  return `El 80% aproximado del impacto se concentra en ${names.length} gerencias: ${names.join(", ")}.`;
}


const reconciliationState={rpZip:null,finalZip:null,rpFiles:[],finalFiles:[]};

async function inspectZipExcelFiles(file){
  const zip=await JSZip.loadAsync(await file.arrayBuffer());
  const result=[];
  for(const entry of Object.values(zip.files)){
    if(entry.dir||!/\.(xlsx|xls)$/i.test(entry.name))continue;
    try{
      const ab=await entry.async("arraybuffer");
      const wb=XLSX.read(ab,{type:"array",raw:true});
      const sheets=[];
      for(const sn of wb.SheetNames){
        const arr=XLSX.utils.sheet_to_json(wb.Sheets[sn],{header:1,defval:"",raw:true});
        const nonempty=arr.filter(r=>r.some(v=>cleanText(v)!==""));
        let numericTotal=0,numericCells=0;
        nonempty.forEach(r=>r.forEach(v=>{
          if(typeof v==="number"&&Number.isFinite(v)){numericTotal+=v;numericCells++;}
        }));
        sheets.push({name:sn,rows:nonempty.length,numericTotal,numericCells});
      }
      result.push({name:entry.name,workbook:wb,sheets});
    }catch(err){
      result.push({name:entry.name,error:err.message,sheets:[]});
    }
  }
  return result;
}
function normalizeName(s){return cleanText(s).toLowerCase().replace(/\.(xlsx|xls)$/,"").replace(/plantilla/g,"").replace(/\s+/g," ").trim();}
function sheetSimilarity(a,b){
  const x=cleanText(a).toLowerCase(),y=cleanText(b).toLowerCase();
  if(x===y)return 100;
  const ax=new Set(x.split(/\s+/)),by=new Set(y.split(/\s+/));let common=0;ax.forEach(t=>{if(by.has(t))common++});
  return common/Math.max(ax.size,by.size,1)*100;
}
function bestFinalFile(rpFile,finalFiles){
  const rn=normalizeName(rpFile.name);
  let best=null,score=-1;
  for(const f of finalFiles){
    const fn=normalizeName(f.name);
    let s=0;
    if(fn===rn)s=100;
    else{
      const ra=new Set(rn.split(/\s+/)),fa=new Set(fn.split(/\s+/));let c=0;ra.forEach(t=>{if(fa.has(t))c++});
      s=c/Math.max(ra.size,fa.size,1)*100;
    }
    if(s>score){score=s;best=f;}
  }
  return score>=25?best:null;
}
function executeReconciliation(){
  const rows=[];
  for(const rpFile of reconciliationState.rpFiles){
    const finalFile=bestFinalFile(rpFile,reconciliationState.finalFiles);
    for(const rpSheet of rpFile.sheets){
      let finalSheet=null;
      if(finalFile?.sheets?.length){
        finalSheet=finalFile.sheets.slice().sort((a,b)=>sheetSimilarity(rpSheet.name,b.name)-sheetSimilarity(rpSheet.name,a.name))[0];
      }
      const finalRows=finalSheet?.rows??0,finalTotal=finalSheet?.numericTotal??0;
      const diff=rpSheet.numericTotal-finalTotal;
      const rowMatch=finalSheet?rpSheet.rows===finalRows:false;
      const totalMatch=finalSheet?Math.abs(diff)<0.02:false;
      const state=!finalSheet?"REVISAR":(rowMatch&&totalMatch?"CUADRA":"DIFERENCIA");
      rows.push({rp:rpFile.name,sheet:rpSheet.name,final:finalFile?.name||"Sin coincidencia",rpRows:rpSheet.rows,finalRows,rpTotal:rpSheet.numericTotal,finalTotal,diff,state});
    }
  }
  $("reconChecks").textContent=rows.length;
  $("reconOk").textContent=rows.filter(x=>x.state==="CUADRA").length;
  $("reconDiff").textContent=rows.filter(x=>x.state==="DIFERENCIA").length;
  $("reconTable").innerHTML=rows.length?rows.map(r=>`<tr><td>${r.rp}</td><td>${r.sheet}</td><td>${r.final}</td><td>${r.rpRows}</td><td>${r.finalRows}</td><td>${signedMoney(r.rpTotal)}</td><td>${signedMoney(r.finalTotal)}</td><td>${signedMoney(r.diff)}</td><td><span class="recon-state ${r.state==="CUADRA"?"ok":r.state==="DIFERENCIA"?"diff":"review"}">${r.state}</span></td></tr>`).join(""):'<tr><td colspan="9">No se encontraron archivos Excel comparables.</td></tr>';
  const ok=rows.filter(x=>x.state==="CUADRA").length,diff=rows.filter(x=>x.state==="DIFERENCIA").length,review=rows.filter(x=>x.state==="REVISAR").length;
  $("reconDiagnosis").className="exec-content";
  $("reconDiagnosis").innerHTML=`Se ejecutaron <b>${rows.length}</b> comparaciones estructurales. <b>${ok}</b> cuadran exactamente en filas y suma numérica global; <b>${diff}</b> presentan diferencias y <b>${review}</b> requieren revisión manual.<br><br><b>Importante:</b> una diferencia aquí no demuestra por sí sola que el FINAL esté mal. Las plantillas RP pueden contener fórmulas, hojas auxiliares o transformaciones específicas. Usa esta matriz para localizar dónde debemos replicar la lógica exacta.`;
}

document.addEventListener("click",async e=>{
  const nav=e.target.closest("[data-view]");if(nav)showView(nav.dataset.view);
  const act=e.target.closest("[data-action]");
  if(act){
    const a=act.dataset.action;
    if(a==="downloadAll")return downloadAll();
    if(!state.model)return alert("Primero procesa los archivos JD.");
    if(a==="downloadCost")saveBlob(workbookBlob(makeCostWorkbook()),"COSTO RCV 2026.xlsx");
    if(a==="downloadExpense")saveBlob(workbookBlob(makeExpenseWorkbook()),"Gastos RCV 2026.xlsx");
    if(a==="downloadXpv")saveBlob(workbookBlob(makeXpvWorkbook()),`Productividad XPV ${state.model.month} 26 RCV.xlsx`);
  }
  const rep=e.target.closest("[data-report]");if(rep){showView("reports");$("reportPreview").innerHTML=reportHtml(rep.dataset.report);}
  if(e.target.closest("#browseRpZip"))$("rpZipInput").click();
  if(e.target.closest("#browseReconFinal"))$("reconFinalZipInput").click();
  if(e.target.closest("#runRpReconciliation"))executeReconciliation();
  if(e.target.closest("#browseBtn"))$("jdFiles").click();
  if(e.target.closest("#processBtn")){
    state.periodIndex=Number($("periodSelect").value);
    try{
      if(state.inputMode==="JD")processModel();
      else if(state.inputMode==="RP")processRpModel();
      else throw new Error("No se reconoció el formato de entrada.");
      showView("home");
    }catch(err){
      console.error(err);
      setStatus("uploadStatus","No fue posible procesar: "+err.message,"error");
    }
  }
  if(e.target.closest("#clearBtn")){state.sources={};state.rpWorkbooks=[];state.inputMode=null;state.model=null;$("jdFiles").value="";renderChecklist();clearDashboard();setStatus("uploadStatus","Aún no has seleccionado archivos.");qa("#processStats strong").forEach(x=>x.textContent="0");}
  if(e.target.closest("#browseFinalBtn"))$("finalZip").click();
  if(e.target.closest("#validateBtn"))validateAgainstFinal();
  if(e.target.closest("#refreshBtn")){state.periodIndex=Number($("periodSelect").value);if(state.model)processModel();}
  if(e.target.closest("#applySettings")){state.threshold=Number($("criticalThreshold").value)||10;state.periodIndex=Number($("settingsPeriod").value);$("periodSelect").value=state.periodIndex;if(state.model)processModel();}
  if(e.target.closest("#presentationBtn")){$("presentationOverlay").classList.add("open");showSlide(0);}
  if(e.target.closest("#closePresentation"))$("presentationOverlay").classList.remove("open");
  if(e.target.closest("#nextSlide"))showSlide(currentSlide+1);
  if(e.target.closest("#prevSlide"))showSlide(currentSlide-1);
  if(e.target.closest("#alertsBtn"))$("alertsDrawer").classList.add("open");
  const lineage=e.target.closest("[data-lineage]");
  if(lineage){$("lineageContent").innerHTML=lineageHtml(lineage.dataset.lineage);$("lineageDrawer").classList.add("open");}
  if(e.target.closest("#closeLineage"))$("lineageDrawer").classList.remove("open");
  if(e.target.closest("#openTopManager360")){const rows=managerHealth();if(rows.length){$("manager360Select").value=rows[0].manager;renderManager360(rows[0].manager);}}
  if(e.target.closest("#closeAlerts"))$("alertsDrawer").classList.remove("open");
  const exec=e.target.closest("[data-exec]");
  if(exec){
    if(!state.model)return alert("Primero procesa los archivos JD.");
    if(exec.dataset.exec==="brief"){$("directorBrief").scrollIntoView({behavior:"smooth"});}
    if(exec.dataset.exec==="speech")alert(prepareSpeech());
    if(exec.dataset.exec==="recommendations")alert($("execActions").innerText||"Sin recomendaciones.");
  }
  if(e.target.closest("#copilotToggle"))$("copilot").classList.add("open");
  if(e.target.closest("#copilotClose"))$("copilot").classList.remove("open");
  const prompt=e.target.closest(".prompts button");if(prompt){$("question").value=prompt.textContent;$("askBtn").click();}
  if(e.target.closest("#askBtn")){const q=$("question").value.trim();if(!q)return;$("chat").insertAdjacentHTML("beforeend",`<div class="user-msg">${q}</div><div class="bot">${answer(q)}</div>`);$("question").value="";$("chat").scrollTop=$("chat").scrollHeight;}
});
$("jdFiles").addEventListener("change",async()=>{
  const files=[...$("jdFiles").files];if(!files.length)return;
  setStatus("uploadStatus","Analizando archivos y detectando modo de entrada…","working");
  $("processBtn").disabled=true;
  try{
    const result=await detectAndLoadInputs(files);
    renderChecklist();
    if(result.mode==="JD"){
      state.periodIndex=detectPeriodFromXpv();
      $("periodSelect").value=state.periodIndex;$("settingsPeriod").value=state.periodIndex;
      setStatus("uploadStatus",`${result.files.length} archivo(s) leído(s). Detectado concentrado reportes JD. Listo para procesar.`,"ok");
    }else if(result.mode==="RP"){
      setStatus("uploadStatus",`${result.files.length} archivo(s) detectado(s). Modo RP: Costos, Gastos y Productividad listos para procesar.`,"ok");
    }else{
      setStatus("uploadStatus","No se reconoció un concentrado reportes JD completo ni el conjunto RP requerido (Costos, Gastos y Productividad).","error");
    }
  }catch(e){
    console.error(e);
    state.inputMode=null;
    renderChecklist();
    setStatus("uploadStatus","Error al leer los archivos: "+e.message,"error");
  }
});
$("finalZip").addEventListener("change",()=>{state.finalZipFile=$("finalZip").files[0]||null;$("validateBtn").disabled=!state.finalZipFile;setStatus("validationStatus",state.finalZipFile?`Referencia seleccionada: ${state.finalZipFile.name}`:"Sin archivo FINAL de referencia.",state.finalZipFile?"ok":"");});
$("periodSelect").addEventListener("change",()=>{$("settingsPeriod").value=$("periodSelect").value;});
$("question").addEventListener("keydown",e=>{if(e.key==="Enter")$("askBtn").click();});


$("rpZipInput")?.addEventListener("change",async()=>{
  const file=$("rpZipInput").files[0];if(!file)return;
  reconciliationState.rpZip=file;$("rpZipName").textContent=file.name;
  $("rpDetectedFiles").innerHTML='<div class="source-file">Leyendo paquete RP…</div>';
  try{
    reconciliationState.rpFiles=await inspectZipExcelFiles(file);
    $("reconRpCount").textContent=reconciliationState.rpFiles.length;
    $("rpDetectedFiles").innerHTML=reconciliationState.rpFiles.map(f=>`<div class="source-file detected">✓ ${f.name} · ${f.sheets.length} hoja(s)</div>`).join("");
    $("runRpReconciliation").disabled=!(reconciliationState.rpFiles.length&&reconciliationState.finalFiles.length);
  }catch(err){$("rpDetectedFiles").innerHTML=`<div class="source-file">Error: ${err.message}</div>`;}
});
$("reconFinalZipInput")?.addEventListener("change",async()=>{
  const file=$("reconFinalZipInput").files[0];if(!file)return;
  reconciliationState.finalZip=file;$("reconFinalName").textContent=file.name;
  try{
    reconciliationState.finalFiles=await inspectZipExcelFiles(file);
    $("reconFinalCount").textContent=reconciliationState.finalFiles.length;
    $("runRpReconciliation").disabled=!(reconciliationState.rpFiles.length&&reconciliationState.finalFiles.length);
  }catch(err){$("reconFinalName").textContent="Error: "+err.message;}
});

$("manager360Select")?.addEventListener("change",e=>renderManager360(e.target.value));
renderChecklist();
clearDashboard();
buildExecutiveIntelligence();
})();