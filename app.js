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
  decisions: JSON.parse(localStorage.getItem("reportia_decisions_v19") || "[]")
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
  const labels = [
    [SHEETS.CATALOG,"Catálogo de Gerentes"],
    [SHEETS.COST,"JD Costo RCV"],
    [SHEETS.EXPENSE,"JD Gastos RCV"],
    [SHEETS.XPV,"JD XPV RCV"]
  ];
  $("sourceChecklist").innerHTML = labels.map(([k,l])=>`<div class="check ${state.sources[k]?"ok":""}">${state.sources[k]?"✓":"○"} ${l}</div>`).join("");
  $("processBtn").disabled = labels.some(([k])=>!state.sources[k]);
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
  const c=$("trendChart"),ctx=c.getContext("2d"),w=c.width,h=c.height;ctx.clearRect(0,0,w,h);
  ctx.strokeStyle="#e8eef7";ctx.lineWidth=1;for(let i=1;i<5;i++){const y=i*h/5;ctx.beginPath();ctx.moveTo(45,y);ctx.lineTo(w-15,y);ctx.stroke();}
  if(!state.model){ctx.fillStyle="#94a3b8";ctx.font="16px system-ui";ctx.textAlign="center";ctx.fillText("Procesa los JD para visualizar tendencias reales",w/2,h/2);return;}
  const upto=state.periodIndex+1;
  const cost=state.model.costMonthly.slice(0,upto).map(x=>Math.abs(x.y2026));
  const exp=state.model.expenseMonthly.slice(0,upto).map(x=>Math.abs(x.y2026));
  const all=[...cost,...exp];const max=Math.max(...all,1);
  [[cost,"#1f6feb"],[exp,"#5b4ae8"]].forEach(([arr,color])=>{ctx.strokeStyle=color;ctx.lineWidth=3;ctx.beginPath();arr.forEach((v,i)=>{const x=55+i*(w-100)/Math.max(1,arr.length-1),y=h-35-v/max*(h-70);i?ctx.lineTo(x,y):ctx.moveTo(x,y);});ctx.stroke();});
  ctx.fillStyle="#64748b";ctx.font="11px system-ui";ctx.textAlign="center";MONTH_LABELS.slice(0,upto).forEach((m,i)=>{const x=55+i*(w-100)/Math.max(1,upto-1);ctx.fillText(m,x,h-12);});
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
  const t=modelTotals(); if(!t){$("analysisCards").innerHTML='<article><span>Estado</span><strong>Sin datos</strong></article>';$("analysisTable").innerHTML="";return;}
  const cards=[["Costo total",money(t.cost2026)],["Gasto total",money(t.expense2026)],["XPV real",money(t.xpReal)],["XPV presupuesto",money(t.xpBudget)],["Cumplimiento XPV",(t.xpCompliance*100).toFixed(2)+"%"],["Gerencias activas",t.managers]];
  $("analysisCards").innerHTML=cards.map(x=>`<article><span>${x[0]}</span><strong>${x[1]}</strong></article>`).join("");
  $("analysisTable").innerHTML=t.deviations.slice(0,30).map(d=>`<tr><td>${d.type}</td><td>${d.manager}</td><td>${signedMoney(d.y2025)}</td><td>${signedMoney(d.y2026)}</td><td>${(d.pct*100).toFixed(2)}%</td></tr>`).join("");
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
function renderAll(){renderDashboard();renderFilters();renderProcessStats();renderAnalysis();renderComparisons();renderManagers();}

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
  zip.file("LEEME.txt","REPORT.IA RCV 19.0 · Archivos generados directamente desde los JD procesados.");
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
  return "Puedo responder sobre resumen, principal riesgo, mayor gasto y productividad XPV usando el modelo procesado.";
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
  if(e.target.closest("#browseBtn"))$("jdFiles").click();
  if(e.target.closest("#processBtn")){state.periodIndex=Number($("periodSelect").value);processModel();showView("home");}
  if(e.target.closest("#clearBtn")){state.sources={};state.model=null;$("jdFiles").value="";renderChecklist();clearDashboard();setStatus("uploadStatus","Aún no has seleccionado archivos.");qa("#processStats strong").forEach(x=>x.textContent="0");}
  if(e.target.closest("#browseFinalBtn"))$("finalZip").click();
  if(e.target.closest("#validateBtn"))validateAgainstFinal();
  if(e.target.closest("#refreshBtn")){state.periodIndex=Number($("periodSelect").value);if(state.model)processModel();}
  if(e.target.closest("#applySettings")){state.threshold=Number($("criticalThreshold").value)||10;state.periodIndex=Number($("settingsPeriod").value);$("periodSelect").value=state.periodIndex;if(state.model)processModel();}
  if(e.target.closest("#copilotToggle"))$("copilot").classList.add("open");
  if(e.target.closest("#copilotClose"))$("copilot").classList.remove("open");
  const prompt=e.target.closest(".prompts button");if(prompt){$("question").value=prompt.textContent;$("askBtn").click();}
  if(e.target.closest("#askBtn")){const q=$("question").value.trim();if(!q)return;$("chat").insertAdjacentHTML("beforeend",`<div class="user-msg">${q}</div><div class="bot">${answer(q)}</div>`);$("question").value="";$("chat").scrollTop=$("chat").scrollHeight;}
});
$("jdFiles").addEventListener("change",async()=>{
  const files=[...$("jdFiles").files];if(!files.length)return;
  setStatus("uploadStatus","Leyendo archivos…","working");
  try{await loadSources(files);const missing=Object.values(SHEETS).filter(s=>!state.sources[s]);if(missing.length)setStatus("uploadStatus","Faltan fuentes: "+missing.join(", "),"error");else{state.periodIndex=detectPeriodFromXpv();$("periodSelect").value=state.periodIndex;$("settingsPeriod").value=state.periodIndex;setStatus("uploadStatus",`${files.length} archivo(s) leído(s). Las cuatro fuentes están listas.`,"ok");}}catch(e){console.error(e);setStatus("uploadStatus","Error al leer Excel: "+e.message,"error");}
});
$("finalZip").addEventListener("change",()=>{state.finalZipFile=$("finalZip").files[0]||null;$("validateBtn").disabled=!state.finalZipFile;setStatus("validationStatus",state.finalZipFile?`Referencia seleccionada: ${state.finalZipFile.name}`:"Sin archivo FINAL de referencia.",state.finalZipFile?"ok":"");});
$("periodSelect").addEventListener("change",()=>{$("settingsPeriod").value=$("periodSelect").value;});
$("question").addEventListener("keydown",e=>{if(e.key==="Enter")$("askBtn").click();});

renderChecklist();
clearDashboard();
})();