const state={base:null,baseLines:[],files:[],results:[]};
const $=s=>document.querySelector(s);
const els={base:$("#base-file"),parts:$("#parts-files"),list:$("#file-list"),validate:$("#validate"),message:$("#global-message"),results:$("#results"),cards:$("#result-cards"),summary:$("#summary")};
const humanSize=b=>b>1048576?`${(b/1048576).toFixed(1)} MB`:`${Math.max(1,Math.round(b/1024))} KB`;
const escapeHtml=s=>String(s??"").replace(/[&<>"']/g,c=>({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));

async function readTxt(file){
  const bytes=await file.arrayBuffer(),utf8=new TextDecoder("utf-8").decode(bytes);
  const broken=(utf8.match(/\uFFFD|ÿý/g)||[]).length;
  return broken?new TextDecoder("windows-1252").decode(bytes):utf8;
}
els.base.addEventListener("change",async e=>{const f=e.target.files[0];if(!f)return;state.base=f;state.baseLines=ValidaRU.parseBase(await readTxt(f));$("#base-zone").classList.add("loaded");$("#base-title").textContent=f.name;$("#base-meta").textContent=`${humanSize(f.size)} · ${state.baseLines.length.toLocaleString("es-EC")} líneas`;refresh()});
els.parts.addEventListener("change",e=>{[...e.target.files].forEach(f=>{if(!state.files.some(x=>x.name===f.name&&x.size===f.size))state.files.push(f)});e.target.value="";refresh()});
$("#clear").addEventListener("click",()=>{state.base=null;state.baseLines=[];state.files=[];state.results=[];els.base.value="";$("#base-zone").classList.remove("loaded");$("#base-title").textContent="Seleccionar TXT base";$("#base-meta").textContent="Se recomienda el archivo de mayor tamaño";els.results.classList.add("hidden");els.message.textContent="";refresh()});
els.validate.addEventListener("click",runValidation);
$("#export-xls").addEventListener("click",exportXls);
$("#export-pdf").addEventListener("click",exportPdf);

function refresh(){
  els.list.innerHTML="";
  state.files.forEach((f,i)=>{const node=$("#file-template").content.cloneNode(true);node.querySelector(".file-kind").textContent=(ValidaRU.detectType([],f.name).slice(0,3)||"XLS").toUpperCase();node.querySelector(".file-name").textContent=f.name;node.querySelector(".file-info").textContent=humanSize(f.size);node.querySelector("button").onclick=()=>{state.files.splice(i,1);refresh()};els.list.appendChild(node)});
  els.validate.disabled=!(state.base&&state.files.length);
}
async function readWorkbook(file){
  const data=await file.arrayBuffer(),book=XLSX.read(data,{type:"array"}),sheet=book.Sheets[book.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet,{defval:"",raw:true}).filter(r=>Object.values(r).some(v=>String(v).trim()));
}
async function runValidation(){
  els.validate.disabled=true;els.validate.textContent="Validando…";els.message.textContent="";
  try{
    const expectedOp=$("#op").value;state.results=[];
    for(const file of state.files){
      const rows=await readWorkbook(file),headers=rows[0]?Object.keys(rows[0]):[],type=ValidaRU.detectType(headers,file.name);
      const report=ValidaRU.validate(rows,state.baseLines,type,expectedOp);
      state.results.push({file:file.name,type,report});
      await new Promise(r=>setTimeout(r,0));
    }
    renderResults();
  }catch(err){console.error(err);els.message.textContent="No se pudo leer uno de los archivos. Revise que sea un Excel válido."}
  finally{els.validate.disabled=false;els.validate.textContent="Validar archivos"}
}
function renderResults(){
  const reports=state.results.map(x=>x.report),approved=reports.filter(x=>x.status==="APROBADO").length,total=reports.reduce((a,x)=>a+x.total,0),errors=reports.reduce((a,x)=>a+x.errors,0);
  els.summary.innerHTML=metric(state.results.length,"Archivos analizados")+metric(approved,"Archivos aprobados")+metric(total.toLocaleString("es-EC"),"Registros revisados")+metric(errors.toLocaleString("es-EC"),"Observaciones");
  els.cards.innerHTML=state.results.map(({file,type,report:r})=>{
    const css=r.status==="APROBADO"?"approved":r.status==="RECHAZADO"?"rejected":"observed";
    const important=r.issues.length?r.issues:r.comparisons;
    const groups=groupComparisons(important);
    const groupOptions=[...groups.keys()].map(k=>`<option value="${escapeHtml(k)}">${escapeHtml(k.replace("|"," · "))}</option>`).join("");
    const rows=important.slice(0,100).map(detailRow).join("");
    const label=r.issues.length?`Revisar diferencias importantes (${r.issues.length})`:"Revisar comparación TXT vs XLS";
    const note=r.notComparable?`<p class="validation-note">${r.notComparable.toLocaleString("es-EC")} registros auxiliares de herrajes no tienen una línea directa en el TXT y no afectan la aprobación.</p>`:"";
    return `<article class="result-card ${css}" data-result="${state.results.findIndex(x=>x.file===file)}"><div class="result-head"><div><h3>${escapeHtml(file)}</h3><small>${escapeHtml(type.toUpperCase())} · validación independiente</small></div><span class="status">${r.status}</span></div><div class="result-stats">${stat(r.total,"Registros")}${stat(r.matched,"Coincidencias directas")}${stat(r.missing,"Faltantes importantes")}${stat(r.differences,"Diferencias importantes")}${stat(r.errors,"Total alertas")}</div>${note}<details><summary>${label} (${important.length.toLocaleString("es-EC")} registros)</summary><div class="detail-tools"><label>Agrupar/filtrar por número y código de mueble<select class="group-filter"><option value="">Todos los muebles</option>${groupOptions}</select></label><input class="detail-search" placeholder="Buscar código, pieza o texto"><span class="shown-count">Mostrando ${Math.min(100,important.length)} de ${important.length}</span></div><div class="issues"><table><thead><tr><th>Fila XLS</th><th>N.º mueble</th><th>Código mueble</th><th>Resultado</th><th>Código/pieza</th><th>Valor XLS</th><th>Línea TXT relacionada</th></tr></thead><tbody>${rows}</tbody></table></div><button class="btn load-more" type="button">Mostrar 100 más</button></details></article>`;
  }).join("");
  bindDetailControls();
  els.results.classList.remove("hidden");els.results.scrollIntoView({behavior:"smooth",block:"start"});
}
const detailRow=x=>`<tr><td>${x.row}</td><td>${escapeHtml(x.numero||"—")}</td><td>${escapeHtml(x.mueble||"—")}</td><td>${escapeHtml(x.kind)}</td><td>${escapeHtml(x.reference||"—")}</td><td>${escapeHtml(x.xls||x.detail||"—")}</td><td class="txt-cell">${x.txtLine?`Línea ${x.txtLine}: `:""}${escapeHtml(x.txt||"—")}</td></tr>`;
function groupComparisons(items){const map=new Map();items.forEach(x=>{const k=x.groupKey||`${x.numero||"—"}|${x.mueble||"—"}`;if(!map.has(k))map.set(k,[]);map.get(k).push(x)});return new Map([...map].sort((a,b)=>a[0].localeCompare(b[0],undefined,{numeric:true})))}
function bindDetailControls(){
  document.querySelectorAll(".result-card").forEach(card=>{
    const result=state.results[Number(card.dataset.result)]?.report;if(!result)return;
    const all=result.issues.length?result.issues:result.comparisons,tbody=card.querySelector("tbody"),filter=card.querySelector(".group-filter"),search=card.querySelector(".detail-search"),count=card.querySelector(".shown-count"),more=card.querySelector(".load-more");let limit=100;
    const draw=()=>{const q=ValidaRU.norm(search.value),g=filter.value;const selected=all.filter(x=>(!g||x.groupKey===g)&&(!q||ValidaRU.norm(`${x.numero} ${x.mueble} ${x.reference} ${x.xls} ${x.txt}`).includes(q)));tbody.innerHTML=selected.slice(0,limit).map(detailRow).join("");count.textContent=`Mostrando ${Math.min(limit,selected.length)} de ${selected.length}`;more.hidden=limit>=selected.length};
    filter.onchange=()=>{limit=100;draw()};search.oninput=()=>{limit=100;draw()};more.onclick=()=>{limit+=100;draw()};draw();
  })
}
const metric=(v,l)=>`<div class="metric"><strong>${v}</strong><span>${l}</span></div>`;
const stat=(v,l)=>`<div class="result-stat"><b>${v}</b><span>${l}</span></div>`;
const reportMeta=()=>({
  year:$("#year").value.trim()||"Sin especificar",lot:$("#lot").value.trim()||"Sin especificar",
  op:ValidaRU.op($("#op").value)||"Sin especificar",project:$("#project").value.trim()||"Sin especificar",
  base:state.base?.name||"Sin archivo",date:new Intl.DateTimeFormat("es-EC",{dateStyle:"medium",timeStyle:"short"}).format(new Date())
});
const downloadBlob=(blob,name)=>{const url=URL.createObjectURL(blob),a=document.createElement("a");a.href=url;a.download=name;document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),1000)};
const xmlEscape=v=>escapeHtml(v).replace(/\r?\n/g,"<br>");
function exportXls(){
  if(!state.results.length)return;
  const meta=reportMeta(),totals=state.results.reduce((a,x)=>({records:a.records+x.report.total,alerts:a.alerts+x.report.errors,approved:a.approved+(x.report.status==="APROBADO")}),{records:0,alerts:0,approved:0});
  const blocks=state.results.map(({file,type,report:r})=>{
    const important=r.issues.length?r.issues:r.comparisons;
    const detail=important.map(i=>`<tr><td>${i.row}</td><td>${xmlEscape(i.numero||"—")}</td><td>${xmlEscape(i.mueble||"—")}</td><td>${xmlEscape(i.kind)}</td><td>${xmlEscape(i.reference||"—")}<br><small>${xmlEscape(i.xls||"—")}</small></td><td>${xmlEscape(i.txtLine?`Línea ${i.txtLine}: ${i.txt}`:(i.txt||"—"))}</td></tr>`).join("");
    return `<tr class="section"><td colspan="6">${xmlEscape(file)} &nbsp; | &nbsp; ${xmlEscape(type.toUpperCase())} &nbsp; | &nbsp; ${r.status}</td></tr><tr class="stats"><td><b>${r.total}</b><small>Registros</small></td><td><b>${r.matched}</b><small>Coincidencias</small></td><td><b>${r.missing}</b><small>Faltantes</small></td><td><b>${r.differences}</b><small>Diferencias</small></td><td colspan="2"><b>${r.errors}</b><small>Total alertas</small></td></tr><tr class="columns"><th>Fila XLS</th><th>N.º mueble</th><th>Código mueble</th><th>Resultado</th><th>Código/pieza y valor XLS</th><th>Registro TXT enfrentado</th></tr>${detail}`;
  }).join("");
  const html=`<!doctype html><html><head><meta charset="utf-8"><style>body{font-family:Calibri,Arial;color:#173042}table{border-collapse:collapse;width:100%}td,th{border:1px solid #d9e1e5;padding:7px;vertical-align:top}.title td{background:#173f5f;color:#fff;font-size:22px;font-weight:bold;padding:16px}.accent td{background:#c89b3c;height:5px;padding:0}.meta th{background:#eef2f4;color:#173f5f;text-align:left;width:16%}.meta td{width:34%}.summary td{background:#eef2f4;text-align:center;font-size:18px;color:#173f5f}.summary small,.stats small{display:block;color:#60717c;font-size:10px}.section td{background:#173f5f;color:#fff;font-size:14px;font-weight:bold;padding:10px}.stats td{text-align:center;background:#f8fafb}.stats b{font-size:16px;color:#173f5f}.columns th{background:#c89b3c;color:#fff;text-align:left}.ok-row{color:#24735b;font-weight:bold}.footer td{border:0;color:#60717c;font-size:10px;padding-top:14px}</style></head><body><table><tr class="title"><td colspan="6">VALIDA RU - REPORTE DE INGENIERÍA</td></tr><tr class="accent"><td colspan="6"></td></tr><tr class="meta"><th>Año</th><td>${xmlEscape(meta.year)}</td><th>Lote</th><td>${xmlEscape(meta.lot)}</td><th>OP</th><td>${xmlEscape(meta.op)}</td></tr><tr class="meta"><th>Proyecto</th><td colspan="3">${xmlEscape(meta.project)}</td><th>Fecha</th><td>${xmlEscape(meta.date)}</td></tr><tr class="meta"><th>TXT base</th><td colspan="5">${xmlEscape(meta.base)}</td></tr><tr class="summary"><td colspan="2"><b>${state.results.length}</b><small>Archivos analizados</small></td><td><b>${totals.approved}</b><small>Archivos aprobados</small></td><td><b>${totals.records}</b><small>Registros revisados</small></td><td colspan="2"><b>${totals.alerts}</b><small>Observaciones</small></td></tr><tr><td colspan="6" style="border:0;height:12px"></td></tr>${blocks}<tr class="footer"><td colspan="6">Generado por VALIDA RU · Control de ingeniería · Procesamiento local</td></tr></table></body></html>`;
  downloadBlob(new Blob(["\ufeff",html],{type:"application/vnd.ms-excel;charset=utf-8"}),`reporte_VALIDA_RU_OP_${meta.op}.xls`);
}
function exportPdf(){
  if(!state.results.length)return;
  if(!window.jspdf?.jsPDF){els.message.textContent="No se pudo cargar el generador PDF. Revise su conexión e inténtelo nuevamente.";return}
  const {jsPDF}=window.jspdf,doc=new jsPDF({orientation:"landscape",unit:"mm",format:"a4"}),meta=reportMeta();
  const navy=[23,63,95],gold=[200,155,60],light=[238,242,244],green=[36,115,91],red=[169,57,57],pageWidth=doc.internal.pageSize.getWidth(),pageHeight=doc.internal.pageSize.getHeight();
  doc.setFillColor(...navy);doc.rect(0,0,pageWidth,29,"F");doc.setFillColor(...gold);doc.rect(0,29,pageWidth,2,"F");doc.setTextColor(255);doc.setFont("helvetica","bold");doc.setFontSize(19);doc.text("VALIDA RU",14,13);doc.setFontSize(10);doc.text("REPORTE TÉCNICO DE VALIDACIÓN",14,21);doc.setFont("helvetica","normal");doc.setFontSize(9);doc.text(`OP ${meta.op}  ·  Lote ${meta.lot}  ·  ${meta.year}`,pageWidth-14,17,{align:"right"});
  doc.autoTable({startY:38,margin:{left:14,right:14},theme:"plain",styles:{fontSize:9,cellPadding:2.5,textColor:navy},body:[[{content:"PROYECTO",styles:{fontStyle:"bold",fillColor:light}},meta.project,{content:"FECHA",styles:{fontStyle:"bold",fillColor:light}},meta.date],[{content:"TXT BASE",styles:{fontStyle:"bold",fillColor:light}},{content:meta.base,colSpan:3}]],columnStyles:{0:{cellWidth:26},1:{cellWidth:95},2:{cellWidth:22},3:{cellWidth:"auto"}}});
  const totals=state.results.reduce((a,x)=>({records:a.records+x.report.total,alerts:a.alerts+x.report.errors,approved:a.approved+(x.report.status==="APROBADO")}),{records:0,alerts:0,approved:0});
  doc.autoTable({startY:doc.lastAutoTable.finalY+5,margin:{left:14,right:14},theme:"grid",styles:{halign:"center",fontSize:9,cellPadding:3,lineColor:[217,225,229]},headStyles:{fillColor:navy,textColor:255},head:[["Archivos analizados","Archivos aprobados","Registros revisados","Observaciones"]],body:[[state.results.length,totals.approved,totals.records,totals.alerts]]});
  state.results.forEach(({file,type,report:r})=>{let y=doc.lastAutoTable.finalY+8;if(y>pageHeight-55){doc.addPage();y=18}doc.setFillColor(...navy);doc.roundedRect(14,y,pageWidth-28,13,2,2,"F");doc.setTextColor(255);doc.setFont("helvetica","bold");doc.setFontSize(10);doc.text(file,18,y+5.5);doc.setFont("helvetica","normal");doc.setFontSize(8);doc.text(`${type.toUpperCase()} · Validación independiente`,18,y+10);doc.setFillColor(...(r.status==="APROBADO"?green:red));doc.roundedRect(pageWidth-56,y+3,38,7,3,3,"F");doc.setFont("helvetica","bold");doc.setTextColor(255);doc.text(r.status,pageWidth-37,y+7.8,{align:"center"});doc.autoTable({startY:y+15,margin:{left:14,right:14},theme:"grid",styles:{halign:"center",fontSize:8,cellPadding:2,lineColor:[217,225,229]},headStyles:{fillColor:gold,textColor:255},head:[["Registros","Coincidencias","Faltantes","Diferencias","Total alertas"]],body:[[r.total,r.matched,r.missing,r.differences,r.errors]]});const important=(r.issues.length?r.issues:r.comparisons).slice(0,50);const rows=important.map(i=>[i.row,i.kind,i.reference||"—",i.xls||"—",i.txt||"—"]);doc.autoTable({startY:doc.lastAutoTable.finalY+3,margin:{left:14,right:14,bottom:16},theme:"striped",styles:{fontSize:6.8,cellPadding:1.8,overflow:"linebreak",lineColor:[217,225,229]},headStyles:{fillColor:navy,textColor:255},alternateRowStyles:{fillColor:light},head:[["Fila XLS","Resultado","Código/pieza","Valor XLS","Registro TXT enfrentado"]],body:rows,columnStyles:{0:{cellWidth:16,halign:"center"},1:{cellWidth:27},2:{cellWidth:32},3:{cellWidth:58},4:{cellWidth:"auto"}}})});
  const pages=doc.internal.getNumberOfPages();for(let page=1;page<=pages;page++){doc.setPage(page);doc.setDrawColor(...gold);doc.line(14,pageHeight-11,pageWidth-14,pageHeight-11);doc.setFontSize(8);doc.setTextColor(96,113,124);doc.text("VALIDA RU · Control de ingeniería",14,pageHeight-6);doc.text(`Página ${page} de ${pages}`,pageWidth-14,pageHeight-6,{align:"right"})}
  doc.save(`reporte_VALIDA_RU_OP_${meta.op}.pdf`);
}
["base-zone","parts-zone"].forEach(id=>{const z=$("#"+id);["dragenter","dragover"].forEach(ev=>z.addEventListener(ev,e=>{e.preventDefault();z.classList.add("drag")}));["dragleave","drop"].forEach(ev=>z.addEventListener(ev,e=>{e.preventDefault();z.classList.remove("drag")}));z.addEventListener("drop",e=>{const input=z.querySelector("input");const dt=new DataTransfer();[...e.dataTransfer.files].forEach(f=>dt.items.add(f));input.files=dt.files;input.dispatchEvent(new Event("change"))})});
