const STORAGE_KEY = "fama_app_usage_history_v1";
const EMP_KEY = "fama_app_employees_v1";

const state = {
  employees: [],
  history: {},
  selectedDate: "",
  pendingLogs: null,
  historyChart: null,
  individualChart: null
};

const DRIVER_ROLES = new Set([
  "MOTORISTA - ÔNIBUS",
  "MOTORISTA - MICRO ÔNIBUS",
  "MOTORISTA - VAN"
]);

const $ = (id) => document.getElementById(id);

function normalize(value){
  return String(value ?? "")
    .normalize("NFD").replace(/[\u0300-\u036f]/g,"")
    .replace(/\s+/g," ").trim().toUpperCase();
}
function formatDate(dateKey){
  if(!dateKey) return "—";
  const [y,m,d] = dateKey.split("-");
  return `${d}/${m}/${y}`;
}
function toDateKey(value){
  if(value instanceof Date && !isNaN(value)){
    const y=value.getFullYear(), m=String(value.getMonth()+1).padStart(2,"0"), d=String(value.getDate()).padStart(2,"0");
    return `${y}-${m}-${d}`;
  }
  const s=String(value ?? "").trim();
  const match=s.match(/^(\d{2})[\/-](\d{2})[\/-](\d{4})/);
  if(match) return `${match[3]}-${match[2]}-${match[1]}`;
  const dt=new Date(s);
  return isNaN(dt) ? "" : toDateKey(dt);
}
function pct(n,d){ return d ? (n/d)*100 : 0; }
function esc(v){
  return String(v ?? "").replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]));
}
function showToast(msg){
  $("toast").textContent=msg; $("toast").classList.remove("hidden");
  clearTimeout(showToast.t); showToast.t=setTimeout(()=>$("toast").classList.add("hidden"),3200);
}
function loadData(){
  try{ state.history=JSON.parse(localStorage.getItem(STORAGE_KEY)||"{}"); }catch{ state.history={}; }
  try{ state.employees=JSON.parse(localStorage.getItem(EMP_KEY)||"[]"); }catch{ state.employees=[]; }
}
function saveData(){
  localStorage.setItem(STORAGE_KEY, JSON.stringify(state.history));
  localStorage.setItem(EMP_KEY, JSON.stringify(state.employees));
}
function getDriverEmployees(){
  return state.employees.filter(e => DRIVER_ROLES.has(normalize(e.cargo)));
}
function historyDates(){ return Object.keys(state.history).sort().reverse(); }
function currentRecord(){
  if(state.selectedDate && state.history[state.selectedDate]) return state.history[state.selectedDate];
  const d=historyDates()[0]; state.selectedDate=d||""; return d?state.history[d]:null;
}

function parseWorkbook(file, callback){
  const reader=new FileReader();
  reader.onload=e=>{
    try{
      const wb=XLSX.read(new Uint8Array(e.target.result), {type:"array", cellDates:true});
      let rows=[];
      for(const sheetName of wb.SheetNames){
        const ws=wb.Sheets[sheetName];
        const json=XLSX.utils.sheet_to_json(ws,{defval:""});
        if(json.length) rows.push(...json);
      }
      callback(null,rows);
    }catch(err){ callback(err); }
  };
  reader.onerror=()=>callback(new Error("Não foi possível ler o arquivo."));
  reader.readAsArrayBuffer(file);
}

function importEmployees(file){
  parseWorkbook(file,(err,rows)=>{
    if(err) return showToast(err.message);
    const mapped=rows.map(r=>{
      const get=(name)=>r[name] ?? r[Object.keys(r).find(k=>normalize(k)===normalize(name))] ?? "";
      return {
        id:String(get("ID")||""),
        name:String(get("NOME")||"").trim(),
        cargo:String(get("CARGO")||"").trim(),
        empresa:String(get("EMPRESA")||"").trim(),
        matricula:String(get("MATRÍCULA")||"").trim()
      };
    }).filter(e=>e.name && DRIVER_ROLES.has(normalize(e.cargo)));
    if(!mapped.length) return showToast("Nenhum dos 3 cargos de motorista foi encontrado.");
    state.employees=mapped;
    saveData();
    updateUI();
    const bus=mapped.filter(e=>normalize(e.cargo)==="MOTORISTA - ÔNIBUS").length;
    const micro=mapped.filter(e=>normalize(e.cargo)==="MOTORISTA - MICRO ÔNIBUS").length;
    const van=mapped.filter(e=>normalize(e.cargo)==="MOTORISTA - VAN").length;
    showToast(`${mapped.length} motoristas importados: ${bus} ônibus + ${micro} micro-ônibus + ${van} vans.`);
  });
}

function importLogs(file){
  parseWorkbook(file,(err,rows)=>{
    if(err) return showToast(err.message);
    const mapped=rows.map(r=>{
      const keys=Object.keys(r);
      const find=(label)=>r[label] ?? r[keys.find(k=>normalize(k)===normalize(label))] ?? "";
      return {
        datetime:find("DATA/HORA"),
        driver:String(find("MOTORISTA")||"").trim(),
        action:String(find("AÇÃO")||"").trim(),
        screen:String(find("TELA")||"").trim()
      };
    }).filter(x=>x.driver);
    if(!mapped.length) return showToast("Não encontrei a coluna MOTORISTA nos logs.");
    const dates=[...new Set(mapped.map(x=>toDateKey(x.datetime)).filter(Boolean))].sort();
    if(!dates.length) return showToast("Não consegui identificar a data nos logs.");
    state.pendingLogs={rows:mapped,dates};
    $("importDate").value=dates.length===1?dates[0]:"";
    $("detectedDates").textContent=dates.length===1
      ? `Data detectada automaticamente: ${formatDate(dates[0])}.`
      : `Foram encontradas várias datas: ${dates.map(formatDate).join(", ")}.`;
    $("dateModal").classList.remove("hidden");
  });
}

function buildSnapshot(dateKey, rows){
  const drivers=getDriverEmployees();
  const byName=new Map(drivers.map(d=>[normalize(d.name),d]));
  const matched=rows.filter(r=>toDateKey(r.datetime)===dateKey && byName.has(normalize(r.driver)));
  const counts={};
  const last={};
  matched.forEach(r=>{
    const key=normalize(r.driver);
    counts[key]=(counts[key]||0)+1;
    const old=last[key];
    const dt=new Date(r.datetime);
    if(!old || (dt instanceof Date && dt>old)) last[key]=dt;
  });
  const driverStates=drivers.map(d=>{
    const key=normalize(d.name);
    return {
      id:d.id,name:d.name,cargo:d.cargo,empresa:d.empresa,matricula:d.matricula,
      used:!!counts[key], events:counts[key]||0,
      last:last[key] && !isNaN(last[key]) ? last[key].toISOString() : ""
    };
  });
  return {
    date:dateKey,
    total:drivers.length,
    used:driverStates.filter(d=>d.used).length,
    unused:driverStates.filter(d=>!d.used).length,
    events:matched.length,
    adherence:pct(driverStates.filter(d=>d.used).length,drivers.length),
    drivers:driverStates
  };
}

function confirmLogImport(){
  if(!state.pendingLogs) return;
  const date=$("importDate").value;
  if(!date) return showToast("Informe a data do log.");
  if(!state.employees.length) return showToast("Importe primeiro a planilha de funcionários.");
  const snapshot=buildSnapshot(date,state.pendingLogs.rows);
  const existed=!!state.history[date];
  state.history[date]=snapshot;
  state.selectedDate=date;
  saveData();
  $("dateModal").classList.add("hidden");
  updateUI();
  showToast(existed?`Histórico de ${formatDate(date)} atualizado.`:`Dia ${formatDate(date)} adicionado ao histórico.`);
  state.pendingLogs=null;
}

function populateDates(){
  const dates=historyDates();
  const select=$("dateSelect");
  select.innerHTML=dates.length ? dates.map(d=>`<option value="${d}">${formatDate(d)}</option>`).join("") : `<option value="">Nenhuma data importada</option>`;
  if(state.selectedDate && dates.includes(state.selectedDate)) select.value=state.selectedDate;
  else if(dates.length){state.selectedDate=dates[0];select.value=dates[0];}
  else state.selectedDate="";
}
function populateDrivers(){
  const select=$("driverSelect");
  const old=select.value;
  select.innerHTML=`<option value="">Todos os motoristas</option>`+
    getDriverEmployees().sort((a,b)=>a.name.localeCompare(b.name,"pt-BR")).map(d=>`<option value="${esc(d.name)}">${esc(d.name)}</option>`).join("");
  if([...select.options].some(o=>o.value===old)) select.value=old;
}

function renderStats(){
  const rec=currentRecord();
  const total=rec?.total ?? getDriverEmployees().length;
  const used=rec?.used ?? 0, unused=rec?.unused ?? total, events=rec?.events ?? 0, adherence=rec?.adherence ?? 0;
  $("statAdherence").textContent=rec?`${adherence.toFixed(2)}%`:"—";
  $("statAdherenceMeta").textContent=rec?formatDate(rec.date):"Importe os logs";
  $("statTotal").textContent=total.toLocaleString("pt-BR");
  $("statUsed").textContent=used.toLocaleString("pt-BR");
  $("statUnused").textContent=unused.toLocaleString("pt-BR");
  $("statEvents").textContent=events.toLocaleString("pt-BR");
  $("indUsed").textContent=used.toLocaleString("pt-BR");
  $("indAdherence").textContent=rec?`${adherence.toFixed(2)}%`:"—";
  $("indUnused").textContent=unused.toLocaleString("pt-BR");
  $("indTotal").textContent=total.toLocaleString("pt-BR");
  $("indEvents").textContent=events.toLocaleString("pt-BR");
  $("indDate").textContent=rec?formatDate(rec.date):"—";
  const tag=$("indAdherenceTag");
  tag.textContent=adherence>=80?"Boa":adherence>=50?"Atenção":"Crítico";
  tag.className=`tag ${adherence>=80?"operation":adherence>=50?"action":"critical"}`;
  const deg=Math.max(0,Math.min(100,adherence))*3.6;
  $("donut").style.background=`conic-gradient(var(--blue) 0deg ${deg}deg, #dfe1e5 ${deg}deg 360deg)`;
  $("donutPercent").textContent=rec?`${Math.round(adherence)}%`:"—";
  $("dataStatus").textContent=state.employees.length?`${state.employees.length} motoristas cadastrados`:"Aguardando planilha";
}

function renderHistoryTable(){
  const dates=historyDates();
  const rows=dates.map(d=>{
    const r=state.history[d];
    return `<tr>
      <td><strong>${formatDate(d)}</strong></td><td>${r.total}</td><td>${r.used}</td><td>${r.unused}</td>
      <td><strong>${r.adherence.toFixed(2)}%</strong></td><td>${r.events}</td>
      <td><button class="btn btn-soft" onclick="selectHistoryDate('${d}')">Abrir</button></td>
    </tr>`;
  }).join("");
  $("historyBody").innerHTML=rows||`<tr><td class="empty" colspan="7">Nenhum dia importado ainda.</td></tr>`;
  $("recentHistoryBody").innerHTML=dates.slice(0,7).map(d=>{
    const r=state.history[d];
    return `<tr onclick="selectHistoryDate('${d}')" style="cursor:pointer"><td>${formatDate(d)}</td><td>${r.total}</td><td>${r.used}</td><td>${r.unused}</td><td>${r.adherence.toFixed(2)}%</td><td>${r.events}</td></tr>`;
  }).join("")||`<tr><td class="empty" colspan="6">Importe os logs do primeiro dia para começar.</td></tr>`;
}

function renderDriverLists(){
  const rec=currentRecord();
  const rows=rec?.drivers||getDriverEmployees().map(d=>({...d,used:false,events:0,last:""}));
  const used=[...rows].filter(d=>d.used).sort((a,b)=>b.events-a.events||a.name.localeCompare(b.name,"pt-BR"));
  const unused=[...rows].filter(d=>!d.used).sort((a,b)=>a.name.localeCompare(b.name,"pt-BR"));

  $("usedListCount").textContent=`${used.length} motorista${used.length===1?"":"s"}`;
  $("unusedListCount").textContent=`${unused.length} motorista${unused.length===1?"":"s"}`;

  $("usedDriversList").innerHTML=used.map(d=>`
    <div class="driver-row">
      <div><strong>${esc(d.name)}</strong><span>${esc(d.cargo)}</span></div>
      <div class="driver-log-count">${d.events} log${d.events===1?"":"s"}</div>
    </div>`).join("") || `<div class="empty">Nenhum motorista usou o app neste dia.</div>`;

  $("unusedDriversList").innerHTML=unused.map(d=>`
    <div class="driver-row">
      <div><strong>${esc(d.name)}</strong><span>${esc(d.cargo)}</span></div>
      <div class="status unused">Sem uso</div>
    </div>`).join("") || `<div class="empty">Todos os motoristas usaram o app neste dia.</div>`;
}

function renderRanking(){
  const rec=currentRecord();
  const rows=[...(rec?.drivers||[])].sort((a,b)=>b.events-a.events||a.name.localeCompare(b.name,"pt-BR"));
  const totalLogs=rows.reduce((sum,d)=>sum+d.events,0);
  $("rankingSummary").textContent=rec?`${rows.filter(d=>d.events>0).length} usuários · ${totalLogs.toLocaleString("pt-BR")} logs`:"Sem dados";

  $("rankingBody").innerHTML=rows.map((d,i)=>{
    const share=totalLogs?pct(d.events,totalLogs):0;
    return `<tr>
      <td class="rank-number ${i<3?"rank-top":""}">${i+1}</td>
      <td><strong>${esc(d.name)}</strong></td>
      <td>${esc(d.cargo)}</td>
      <td><strong>${d.events}</strong></td>
      <td>${share.toFixed(2)}%</td>
      <td><span class="status ${d.used?"used":"unused"}">${d.used?"Usando":"Sem uso"}</span></td>
      <td>${d.last?new Date(d.last).toLocaleString("pt-BR"):"—"}</td>
    </tr>`;
  }).join("") || `<tr><td class="empty" colspan="7">Nenhum dado disponível.</td></tr>`;
}

function renderDrivers(){
  const rec=currentRecord();
  const rows=(rec?.drivers||getDriverEmployees().map(d=>({...d,used:false,events:0,last:""})));
  const search=normalize($("driverSearch").value);
  const status=$("driverStatusFilter").value;
  const filtered=rows.filter(d=>{
    const match=!search||normalize(d.name).includes(search)||normalize(d.cargo).includes(search);
    const st=status==="all"||(status==="used"&&d.used)||(status==="unused"&&!d.used);
    return match&&st;
  });
  $("driversBody").innerHTML=filtered.map(d=>`<tr>
    <td><strong>${esc(d.name)}</strong></td><td>${esc(d.cargo)}</td><td>${esc(d.empresa||"—")}</td>
    <td><span class="status ${d.used?"used":"unused"}">${d.used?"Usando":"Sem uso"}</span></td>
    <td>${d.events}</td><td>${d.last?new Date(d.last).toLocaleString("pt-BR"):"—"}</td>
  </tr>`).join("")||`<tr><td class="empty" colspan="6">Nenhum motorista encontrado.</td></tr>`;
}

function renderHistoryChart(){
  const dates=historyDates().sort();
  const labels=dates.map(formatDate), values=dates.map(d=>state.history[d].adherence);
  if(state.historyChart) state.historyChart.destroy();
  state.historyChart=new Chart($("historyChart"),{
    type:"line",
    data:{labels,datasets:[{label:"Adesão",data:values,borderWidth:2,tension:.35,pointRadius:3,fill:false}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"}},x:{grid:{display:false}}}}
  });
}

function renderIndividual(){
  const selected=$("driverSelect").value;
  const rec=currentRecord();
  if(!selected){
    $("individualTitle").textContent="Selecione um motorista";
    $("individualSubtitle").textContent="Use o filtro acima para acompanhar a evolução individual.";
    $("individualSummary").innerHTML="";
    if(state.individualChart) state.individualChart.destroy();
    return;
  }
  const key=normalize(selected);
  const dates=historyDates().sort();
  const person=state.employees.find(e=>normalize(e.name)===key);
  const series=dates.map(d=>state.history[d].drivers.find(x=>normalize(x.name)===key)?.used?1:0);
  const events=dates.map(d=>state.history[d].drivers.find(x=>normalize(x.name)===key)?.events||0);
  const daysUsed=series.filter(Boolean).length;
  $("individualTitle").textContent=person?.name||selected;
  $("individualSubtitle").textContent=person?`${person.cargo} · ${person.empresa||"Empresa não informada"}`:"";
  $("individualSummary").innerHTML=[
    ["Dias usando",daysUsed],["Dias sem uso",Math.max(0,dates.length-daysUsed)],["Adesão individual",dates.length?((daysUsed/dates.length)*100).toFixed(1)+"%":"—"],["Eventos no histórico",events.reduce((a,b)=>a+b,0)]
  ].map(x=>`<div class="individual-box"><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join("");
  if(state.individualChart) state.individualChart.destroy();
  state.individualChart=new Chart($("individualChart"),{
    type:"bar",
    data:{labels:dates.map(formatDate),datasets:[{label:"Uso do app",data:series.map(v=>v*100),borderWidth:1}]},
    options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"}}}}
  });
}

function updateUI(){
  loadData();
  populateDates(); populateDrivers(); renderStats(); renderHistoryTable(); renderDrivers(); renderDriverLists(); renderRanking(); renderHistoryChart(); renderIndividual();
}

window.selectHistoryDate=(date)=>{
  state.selectedDate=date;
  $("dateSelect").value=date;
  updateUI();
  activateTab("overview");
};

function activateTab(name){
  document.querySelectorAll(".tab").forEach(b=>b.classList.toggle("active",b.dataset.tab===name));
  document.querySelectorAll(".tab-panel").forEach(p=>p.classList.toggle("active",p.id===`panel-${name}`));
}

function fullReset(){
  const confirmed=confirm(
    "RESET COMPLETO\n\nIsso vai apagar a base de funcionários, todos os logs processados e todo o histórico deste navegador.\n\nDepois você poderá importar novamente a planilha correta.\n\nDeseja continuar?"
  );
  if(!confirmed) return;
  const typed=prompt('Para confirmar, digite: RESET');
  if(typed!=="RESET") return showToast("Reset cancelado.");
  localStorage.removeItem(STORAGE_KEY);
  localStorage.removeItem(EMP_KEY);
  state.employees=[];
  state.history={};
  state.selectedDate="";
  state.pendingLogs=null;
  updateUI();
  showToast("Reset completo realizado. Importe a planilha de funcionários novamente.");
}

function exportHistoryCSV(){
  const rows=[["Data","Total motoristas","Usando","Sem uso","Adesão","Eventos"]];
  historyDates().sort().forEach(d=>{const r=state.history[d];rows.push([d,r.total,r.used,r.unused,r.adherence.toFixed(2),r.events]);});
  const csv=rows.map(row=>row.map(v=>`"${String(v).replaceAll('"','""')}"`).join(";")).join("\n");
  const blob=new Blob(["\ufeff"+csv],{type:"text/csv;charset=utf-8"});
  const a=document.createElement("a");a.href=URL.createObjectURL(blob);a.download="historico_adesao_motoristas.csv";a.click();URL.revokeObjectURL(a.href);
}

$("btnFullReset").onclick=fullReset;
$("btnImportEmployees").onclick=()=>$("employeesInput").click();
$("btnImportLogs").onclick=()=>$("logsInput").click();
$("employeesInput").onchange=e=>{if(e.target.files[0]) importEmployees(e.target.files[0]);e.target.value=""};
$("logsInput").onchange=e=>{if(e.target.files[0]) importLogs(e.target.files[0]);e.target.value=""};
$("confirmImport").onclick=confirmLogImport;
$("cancelImport").onclick=()=>{$("dateModal").classList.add("hidden");state.pendingLogs=null};
$("closeModal").onclick=()=>{$("dateModal").classList.add("hidden");state.pendingLogs=null};
$("dateSelect").onchange=e=>{state.selectedDate=e.target.value;updateUI()};
$("driverSelect").onchange=()=>{renderIndividual();activateTab("individual")};
$("driverSearch").oninput=renderDrivers;
$("driverStatusFilter").onchange=renderDrivers;
$("btnExportHistory").onclick=exportHistoryCSV;
$("periodSelect").onchange=e=>{
  if(e.target.value==="all"){state.selectedDate=historyDates().sort()[0]||""}else{state.selectedDate=historyDates()[0]||""}
  updateUI();
};
$("btnClearHistory").onclick=()=>{
  if(!confirm("Apagar todo o histórico salvo neste navegador?")) return;
  localStorage.removeItem(STORAGE_KEY);state.history={};state.selectedDate="";updateUI();showToast("Histórico apagado.");
};
document.querySelectorAll(".tab").forEach(btn=>btn.addEventListener("click",()=>activateTab(btn.dataset.tab)));

loadData();
updateUI();
