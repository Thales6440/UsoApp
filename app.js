const SUPABASE_URL = "https://tolhxaxlstbsnophypxi.supabase.co";
const SUPABASE_KEY = "sb_publishable_cXIx-TlykFu-5CC_Ypyt6w_HxUQ0Qb_";
const { createClient } = window.supabase;
const db = createClient(SUPABASE_URL, SUPABASE_KEY);

const state = {
  user:null, employees:[], history:{}, selectedDate:"",
  pendingLogs:null, pendingChecklist:null, checklistRows:[], historyChart:null, individualChart:null
};

const DRIVER_ROLES = new Set([
  normalize("MOTORISTA - ÔNIBUS"),
  normalize("MOTORISTA - MICRO ÔNIBUS"),
  normalize("MOTORISTA - VAN")
]);

function $(id){return document.getElementById(id)}
function normalize(value){return String(value??"").normalize("NFD").replace(/[\u0300-\u036f]/g,"").replace(/\.+$/,"").replace(/\s+/g," ").trim().toUpperCase()}
function formatDate(d){if(!d)return"—";const [y,m,day]=d.split("-");return`${day}/${m}/${y}`}
function toDateKey(v){
  if(v instanceof Date&&!isNaN(v))return `${v.getFullYear()}-${String(v.getMonth()+1).padStart(2,"0")}-${String(v.getDate()).padStart(2,"0")}`;
  if(typeof v==="number"&&isFinite(v)){const d=new Date(Date.UTC(1899,11,30)+Math.round(v*86400000));return isNaN(d)?"":`${d.getUTCFullYear()}-${String(d.getUTCMonth()+1).padStart(2,"0")}-${String(d.getUTCDate()).padStart(2,"0")}`;}
  const str=String(v??"").trim();
  const m=str.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if(m)return `${m[3]}-${String(m[2]).padStart(2,"0")}-${String(m[1]).padStart(2,"0")}`;
  const iso=str.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if(iso)return `${iso[1]}-${String(iso[2]).padStart(2,"0")}-${String(iso[3]).padStart(2,"0")}`;
  const dt=new Date(str);return isNaN(dt)?"":`${dt.getFullYear()}-${String(dt.getMonth()+1).padStart(2,"0")}-${String(dt.getDate()).padStart(2,"0")}`;
}
function pct(n,d){return d?(n/d)*100:0}
function esc(v){return String(v??"").replace(/[&<>"']/g,m=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[m]))}
function showToast(msg){$("toast").textContent=msg;$("toast").classList.remove("hidden");clearTimeout(showToast.t);showToast.t=setTimeout(()=>$("toast").classList.add("hidden"),3500)}
function historyDates(){return Object.keys(state.history).sort().reverse()}
function currentRecord(){const d=state.selectedDate||historyDates()[0]||"";return d?state.history[d]:null}

async function loadCloudData(){
  const {data:emps,error:e1}=await db.from("motoristas").select("id,matricula,nome,cargo,empresa,ativo").eq("ativo",true).order("nome");
  if(e1) throw e1;
  state.employees=emps||[];
  const {data:usage,error:e2}=await db.from("uso_diario").select("motorista_id,data,quantidade_logs,primeiro_log,ultimo_log").order("data",{ascending:false});
  if(e2) throw e2;
  const byDate={};
  for(const u of usage||[]){
    if(!byDate[u.data])byDate[u.data]=[];
    byDate[u.data].push(u);
  }
  state.history={};
  for(const [date,rows] of Object.entries(byDate)){
    const map=new Map(rows.map(r=>[r.motorista_id,r]));
    const drivers=state.employees.map(e=>{const u=map.get(e.id);return{id:e.id,matricula:e.matricula,name:e.nome,cargo:e.cargo,empresa:e.empresa,used:!!u&&u.quantidade_logs>0,events:u?.quantidade_logs||0,last:u?.ultimo_log||"",first:u?.primeiro_log||""}});
    const used=drivers.filter(d=>d.used).length;
    state.history[date]={date,total:drivers.length,used,unused:drivers.length-used,events:drivers.reduce((s,d)=>s+d.events,0),adherence:pct(used,drivers.length),drivers};
  }
  state.selectedDate=state.selectedDate&&state.history[state.selectedDate]?state.selectedDate:historyDates()[0]||"";
  await loadChecklistData();
}

async function importEmployees(file){
  parseWorkbook(file,async(err,rows)=>{
    if(err)return showToast(err.message);
    try{
      const mapped=rows.map(r=>({
        matricula:String(r["MATRÍCULA"]??r["MATRICULA"]??"").trim(),
        nome:String(r["NOME"]??"").trim(),
        cargo:String(r["CARGO"]??"").trim(),
        empresa:String(r["EMPRESA"]??"").trim(),
        ativo:true
      })).filter(e=>e.nome&&DRIVER_ROLES.has(normalize(e.cargo)))
        // Matrícula é opcional na planilha (ver README); quando faltar, geramos uma
        // chave estável a partir do nome para o upsert não colidir nem descartar o motorista.
        .map(e=>e.matricula?e:{...e,matricula:`AUTO-${normalize(e.nome).replace(/\s+/g,"-")}`});
      if(!mapped.length)return showToast("Nenhum motorista com os cargos configurados foi encontrado na planilha.");
      const {error}=await db.from("motoristas").upsert(mapped,{onConflict:"matricula"});
      if(error)throw error;
      await loadCloudData();updateUI();
      const bus=mapped.filter(e=>normalize(e.cargo)===normalize("MOTORISTA - ÔNIBUS")).length;
      const micro=mapped.filter(e=>normalize(e.cargo)===normalize("MOTORISTA - MICRO ÔNIBUS")).length;
      const van=mapped.filter(e=>normalize(e.cargo)===normalize("MOTORISTA - VAN")).length;
      showToast(`${mapped.length} motoristas sincronizados: ${bus} ônibus + ${micro} micro-ônibus + ${van} vans.`);
    }catch(e){console.error(e);showToast("Erro ao salvar funcionários: "+e.message)}
  });
}

function parseWorkbook(file,callback){
  const reader=new FileReader();
  reader.onload=e=>{try{const wb=XLSX.read(new Uint8Array(e.target.result),{type:"array",cellDates:true});let rows=[];for(const s of wb.SheetNames){const j=XLSX.utils.sheet_to_json(wb.Sheets[s],{defval:""});if(j.length)rows.push(...j)}callback(null,rows)}catch(err){callback(err)}};
  reader.onerror=()=>callback(new Error("Não foi possível ler o arquivo."));
  reader.readAsArrayBuffer(file);
}

function importLogs(file){
  parseWorkbook(file,(err,rows)=>{
    if(err)return showToast(err.message);
    const mapped=rows.map(r=>({datetime:r["DATA/HORA"]||"",driver:String(r["MOTORISTA"]||"").trim(),action:String(r["AÇÃO"]||"").trim(),screen:String(r["TELA"]||"").trim()})).filter(x=>x.driver);
    const dates=[...new Set(mapped.map(x=>toDateKey(x.datetime)).filter(Boolean))].sort();
    if(!mapped.length)return showToast("Não encontrei a coluna MOTORISTA.");
    if(!dates.length)return showToast("Não consegui identificar a data dos logs.");
    state.pendingLogs={rows:mapped,dates};
    $("importDate").value=dates.length===1?dates[0]:"";
    $("detectedDates").textContent=dates.length===1?`Data detectada: ${formatDate(dates[0])}.`:`Datas encontradas: ${dates.map(formatDate).join(", ")}.`;
    $("dateModal").classList.remove("hidden");
  });
}

async function confirmLogImport(){
  if(!state.pendingLogs)return;
  const date=$("importDate").value;
  if(!date)return showToast("Informe a data do log.");
  if(!state.employees.length)return showToast("Importe a planilha de funcionários primeiro.");
  try{
    const byName=new Map(state.employees.map(e=>[normalize(e.nome),e]));
    const dayRows=state.pendingLogs.rows.filter(r=>toDateKey(r.datetime)===date);
    const matched=dayRows.map(r=>({...r,employee:byName.get(normalize(r.driver))})).filter(r=>r.employee);
    const unknown=dayRows.filter(r=>!byName.has(normalize(r.driver)));
    // Replace the selected day so re-importing the same file never duplicates it.
    const {error:delU}=await db.from("uso_diario").delete().eq("data",date);
    if(delU)throw delU;
    const {error:delL}=await db.from("app_logs").delete().eq("data_log",date);
    if(delL)throw delL;

    const logs=matched.map(r=>({
      motorista_id:r.employee.id,data_log:date,
      data_hora:parseDateTime(r.datetime),nome_motorista_log:r.driver,
      acao:r.action||null,tela:r.screen||null
    }));
    for(let i=0;i<logs.length;i+=500){
      const {error}=await db.from("app_logs").insert(logs.slice(i,i+500));
      if(error)throw error;
    }

    const grouped=new Map();
    for(const r of matched){
      const id=r.employee.id,dt=parseDateTime(r.datetime);
      if(!grouped.has(id))grouped.set(id,{motorista_id:id,data:date,quantidade_logs:0,primeiro_log:dt,ultimo_log:dt});
      const g=grouped.get(id);g.quantidade_logs++;
      if(dt&&(!g.primeiro_log||dt<g.primeiro_log))g.primeiro_log=dt;
      if(dt&&(!g.ultimo_log||dt>g.ultimo_log))g.ultimo_log=dt;
    }
    const usage=[...grouped.values()];
    for(let i=0;i<usage.length;i+=500){
      const {error}=await db.from("uso_diario").insert(usage.slice(i,i+500));
      if(error)throw error;
    }
    await loadCloudData();updateUI();
    $("dateModal").classList.add("hidden");state.pendingLogs=null;
    showToast(`${matched.length} eventos salvos em ${formatDate(date)}. ${unknown.length} registros ignorados por não pertencerem à base de motoristas.`);
  }catch(e){console.error(e);showToast("Erro ao salvar logs: "+e.message)}
}

function parseDateTime(v){
  if(v instanceof Date&&!isNaN(v))return v.toISOString();
  const s=String(v??"").trim();
  const m=s.match(/^(\d{2})\/(\d{2})\/(\d{4})\s+(\d{2}):(\d{2})(?::(\d{2}))?/);
  if(m)return new Date(`${m[3]}-${m[2]}-${m[1]}T${m[4]}:${m[5]}:${m[6]||"00"}-03:00`).toISOString();
  const d=new Date(s);return isNaN(d) ? null : d.toISOString();
}


async function loadChecklistData(){
  const {data,error}=await db.from("checklist_registros").select("id,motorista_id,data,tipo,horario_inicial,horario_final,status,prefixo");
  if(error){
    console.error(error);
    state.checklistRows=[];
    showToast("Erro ao carregar checklists: "+error.message);
    return;
  }
  state.checklistRows=data||[];
}

function checklistRowsFromState(){
  return state.checklistRows || [];
}

function checklistAvailableDates(){
  return [...new Set(checklistRowsFromState().map(r=>String(r.data||"")).filter(Boolean))].sort();
}

function checklistDefaultRange(){
  const dates=checklistAvailableDates();
  return {start:dates[0]||"",end:dates[dates.length-1]||""};
}

function checklistReport(startDate,endDate){
  const employees=state.employees||[];
  const allDates=checklistAvailableDates();
  const start=startDate||allDates[0]||"";
  const end=endDate||allDates[allDates.length-1]||"";

  const raw=checklistRowsFromState().filter(r=>{
    const d=String(r.data||"");
    return d && (!start||d>=start) && (!end||d<=end);
  });

  // One checklist per type (INÍCIO/FIM) per motorista/day.
  const unique=new Map();
  for(const r of raw){
    const type=normalize(r.tipo).replace("INICIO","INÍCIO");
    if(type!=="INÍCIO" && type!=="FIM") continue;
    const key=`${r.motorista_id||normalize(r.nome_motorista||"")}|${r.data}|${type}`;
    if(!unique.has(key)) unique.set(key,r);
  }
  const rows=[...unique.values()];
  const days=[...new Set(rows.map(r=>r.data).filter(Boolean))].sort();
  const goal=days.length*2;

  const byDriver=new Map();
  for(const e of employees){
    byDriver.set(e.id,{
      id:e.id,name:e.nome,matricula:e.matricula,cargo:e.cargo,empresa:e.empresa,
      checklists:0,days:new Set(),dayCounts:{}
    });
  }

  const byName=new Map(employees.map(e=>[normalize(e.nome),e]));
  for(const r of rows){
    let e=r.motorista_id?employees.find(x=>x.id===r.motorista_id):null;
    if(!e) e=byName.get(normalize(r.nome_motorista||""));
    if(!e) continue;
    const g=byDriver.get(e.id);
    if(!g) continue;
    g.checklists++;
    g.days.add(r.data);
    g.dayCounts[r.data]=(g.dayCounts[r.data]||0)+1;
  }

  const report=[...byDriver.values()].map(g=>{
    const completeDays=days.filter(d=>(g.dayCounts[d]||0)>=2).length;
    const percent=goal?Math.min(100,pct(g.checklists,goal)):0;
    const status=goal>0 && g.checklists>=goal ? "inside" : "outside";
    return {...g,daysCompleted:completeDays,percent,status,daysUsed:g.days.size};
  });

  report.sort((a,b)=>b.checklists-a.checklists||a.name.localeCompare(b.name,"pt-BR"));
  return {start,end,days,goal,rows,report,rawRows:raw.length};
}

function renderChecklist(){
  const startInput=$("checklistStartDate");
  const endInput=$("checklistEndDate");
  if(!startInput||!endInput)return;

  const defaults=checklistDefaultRange();
  if(!startInput.value)startInput.value=defaults.start;
  if(!endInput.value)endInput.value=defaults.end;

  let start=startInput.value||defaults.start;
  let end=endInput.value||defaults.end;
  if(start && end && start>end){
    const tmp=start;start=end;end=tmp;
    startInput.value=start;endInput.value=end;
  }

  const data=checklistReport(start,end);
  $("checklistDays").textContent=data.days.length;
  $("checklistGoal").textContent=data.goal;
  $("checklistTotal").textContent=data.rows.length;

  const inside=data.report.filter(r=>r.status==="inside").length;
  const outside=data.report.filter(r=>r.status==="outside").length;
  $("checklistOnTarget").textContent=inside;
  $("checklistOutsideTarget").textContent=outside;

  const filter=$("checklistStatusFilter")?.value||"all";
  const search=normalize($("checklistSearch")?.value||"");
  const visible=data.report.filter(r=>{
    const okFilter=filter==="all"||r.status===filter;
    const okSearch=!search||normalize(r.name).includes(search)||normalize(r.matricula).includes(search);
    return okFilter&&okSearch;
  });

  $("checklistBody").innerHTML=visible.map((r,i)=>{
    const label=r.status==="inside"?"Dentro da meta":"Fora da meta";
    return `<tr>
      <td>${i+1}</td>
      <td><strong>${esc(r.name)}</strong></td>
      <td>${esc(r.cargo)}</td>
      <td>${r.daysCompleted}/${data.days.length}</td>
      <td><strong>${r.checklists}</strong></td>
      <td>${data.goal}</td>
      <td>${r.percent.toFixed(1)}%</td>
      <td><span class="status ${r.status==="inside"?"used":"unused"}">${label}</span></td>
    </tr>`;
  }).join("") || `<tr><td class="empty" colspan="8">Nenhum motorista encontrado.</td></tr>`;
}

function applyChecklistPeriod(){
  const start=$("checklistStartDate").value;
  const end=$("checklistEndDate").value;
  if(start && end && start>end){
    showToast("A data inicial não pode ser maior que a data final.");
    return;
  }
  renderChecklist();
  showToast(`Período aplicado: ${formatDate(start)} a ${formatDate(end)}.`);
}

// O relatório real do sistema de checklist traz DATA (só data), HORÁRIO INICIAL,
// HORÁRIO FINAL, TIPO, MOTORISTA, PREFIXO e STATUS CHECKLIST separados — não existe
// coluna DATA/HORA nem MATRÍCULA. Este parser lê o formato real do arquivo.
function combineDateTime(dateKey,timeVal){
  if(!dateKey || timeVal === "" || timeVal == null) return null;

  let h=0,m=0,s=0;

  if(timeVal instanceof Date && !isNaN(timeVal)){
    h=timeVal.getHours();
    m=timeVal.getMinutes();
    s=timeVal.getSeconds();
  }else{
    const value=String(timeVal).trim();

    // Aceita HH:MM ou HH:MM:SS
    const mt=value.match(/^(\\d{1,2}):(\\d{2})(?::(\\d{2}))?/);
    if(!mt) return null;

    h=Number(mt[1]);
    m=Number(mt[2]);
    s=Number(mt[3]||0);
  }

  if(h>23 || m>59 || s>59) return null;

  // A coluna do Supabase é do tipo TIME.
  // Portanto enviamos somente HH:MM:SS, e não um timestamp ISO.
  return `${String(h).padStart(2,"0")}:${String(m).padStart(2,"0")}:${String(s).padStart(2,"0")}`;
}

function parseChecklistRows(file){
  parseWorkbook(file,(err,rows)=>{
    if(err)return showToast("Erro ao ler o Excel: "+err.message);
    const mapped=rows.map(r=>({
      dateKey:toDateKey(rowValue(r,"DATA","DATA/HORA")),
      driver:String(rowValue(r,"MOTORISTA")||"").trim(),
      matricula:String(rowValue(r,"MATRÍCULA","MATRICULA")||"").trim(),
      tipo:String(rowValue(r,"TIPO","AÇÃO")||"").trim(),
      horaInicial:rowValue(r,"HORÁRIO INICIAL","HORARIO INICIAL"),
      horaFinal:rowValue(r,"HORÁRIO FINAL","HORARIO FINAL"),
      prefixo:String(rowValue(r,"PREFIXO")||"").trim(),
      status:String(rowValue(r,"STATUS CHECKLIST","STATUS")||"").trim()
    })).filter(x=>x.driver&&x.dateKey);
    if(!mapped.length)return showToast("Não encontrei registros. Verifique as colunas DATA e MOTORISTA da planilha.");
    const valid=mapped.filter(x=>normalizeChecklistType(x.tipo)).length;
    if(!valid)return showToast("Encontrei os motoristas, mas não encontrei TIPO com INÍCIO ou FIM.");
    const dates=[...new Set(mapped.map(x=>x.dateKey))].sort();
    state.pendingChecklist={rows:mapped};
    $("checklistDetected").textContent=`${mapped.length} linhas · ${valid} INÍCIO/FIM válidos · ${dates.length} dia(s)`;
    $("checklistSaveError").textContent="";$("checklistModal").classList.remove("hidden");
  });
}

async function confirmChecklistImport(){
  if(!state.pendingChecklist)return;
  const errorBox=$("checklistSaveError"),button=$("confirmChecklistImport"); errorBox.textContent="";
  if(!state.employees.length){errorBox.textContent="Importe a planilha de funcionários antes do checklist.";return;}
  button.disabled=true;button.textContent="Salvando...";
  try{
    const byName=new Map(state.employees.map(e=>[normalize(e.nome),e]));
    const byMatricula=new Map(state.employees.filter(e=>e.matricula).map(e=>[normalize(e.matricula),e]));
    const unmatched=new Set(),unique=new Map();
    for(const r of state.pendingChecklist.rows){
      const emp=(r.matricula&&byMatricula.get(normalize(r.matricula)))||byName.get(normalize(r.driver));
      const tipo=normalizeChecklistType(r.tipo);
      if(!emp){unmatched.add(r.driver);continue;} if(!r.dateKey||!tipo)continue;
      const horario=combineDateTime(r.dateKey,tipo==="INÍCIO"?r.horaInicial:r.horaFinal);
      const rec={motorista_id:emp.id,data:r.dateKey,tipo,horario_inicial:tipo==="INÍCIO"?horario:null,horario_final:tipo==="FIM"?horario:null,status:r.status||null,prefixo:r.prefixo||null};
      const key=`${emp.id}|${r.dateKey}|${tipo}`; if(!unique.has(key))unique.set(key,rec);
    }
    const rows=[...unique.values()]; if(!rows.length){errorBox.textContent="Nenhum registro válido foi cruzado com os motoristas. Confira MOTORISTA e TIPO INÍCIO/FIM.";return;}
    const dates=[...new Set(rows.map(r=>r.data))];
    for(const d of dates){const {error}=await db.from("checklist_registros").delete().eq("data",d);if(error)throw new Error(`Não foi possível limpar ${formatDate(d)}: ${error.message}`);}
    for(let i=0;i<rows.length;i+=500){const {error}=await db.from("checklist_registros").insert(rows.slice(i,i+500));if(error)throw new Error(`Supabase recusou o lote ${Math.floor(i/500)+1}: ${error.message}`);}
    await loadChecklistData();renderChecklist();$("checklistModal").classList.add("hidden");state.pendingChecklist=null;
    showToast(`${rows.length} checklists válidos salvos.${unmatched.size?` ${unmatched.size} motorista(s) não encontrados.`:""}`);
  }catch(e){console.error("CHECKLIST_IMPORT_ERROR",e);errorBox.textContent="Erro ao salvar no Supabase: "+(e.message||"erro desconhecido");}
  finally{button.disabled=false;button.textContent="Salvar no Supabase";}
}

async function clearAllChecklists(){
  if(!confirm("Isso apagará todos os checklists importados do Supabase para todos os usuários.\n\nDeseja continuar?"))return;
  try{
    const {error}=await db.from("checklist_registros").delete().not("id","is",null);
    if(error)throw error;
    await loadChecklistData();
    renderChecklist();
    showToast("Checklists removidos.");
  }catch(e){
    console.error(e);
    showToast("Erro ao limpar checklists: "+e.message);
  }
}

function exportChecklistReport(){
  const start=$("checklistStartDate").value;
  const end=$("checklistEndDate").value;
  const data=checklistReport(start,end);
  if(!data.report.length)return showToast("Nenhum dado de checklist para exportar.");
  const rows=[
    ["#","Motorista","Matrícula","Cargo","Dias completos","Total de dias","Checklists","Meta","%","Situação"],
    ...data.report.map((r,i)=>[i+1,r.name,r.matricula,r.cargo,r.daysCompleted,data.days.length,r.checklists,data.goal,r.percent.toFixed(1)+"%",r.status==="inside"?"Dentro da meta":"Fora da meta"])
  ];
  csvDownload(`relatorio_checklist_${data.start||"inicio"}_${data.end||"fim"}.csv`,rows);
}

function renderBaseCheck(){
  const d=state.employees,bus=d.filter(e=>normalize(e.cargo)===normalize("MOTORISTA - ÔNIBUS")).length,micro=d.filter(e=>normalize(e.cargo)===normalize("MOTORISTA - MICRO ÔNIBUS")).length,van=d.filter(e=>normalize(e.cargo)===normalize("MOTORISTA - VAN")).length;
  $("baseCheckText").textContent=d.length?`${d.length} motoristas ativos no Supabase.`:"Nenhum motorista cadastrado.";
  $("baseCheckGrid").innerHTML=[["Ônibus",bus],["Micro-ônibus",micro],["Van",van],["Total",d.length]].map(([l,n])=>`<div class="base-check-item"><span>${l}</span><strong>${n}</strong></div>`).join("");
}

function populateDates(){
  const ds=historyDates(),sel=$("dateSelect");
  sel.innerHTML=ds.length?ds.map(d=>`<option value="${d}">${formatDate(d)}</option>`).join(""):`<option value="">Nenhuma data importada</option>`;
  if(state.selectedDate&&ds.includes(state.selectedDate))sel.value=state.selectedDate;else if(ds.length){state.selectedDate=ds[0];sel.value=ds[0]}else state.selectedDate="";
}
function populateDrivers(){
  const sel=$("driverSelect"),old=sel.value;
  sel.innerHTML=`<option value="">Todos os motoristas</option>`+state.employees.map(d=>`<option value="${esc(d.nome)}">${esc(d.nome)}</option>`).join("");
  if([...sel.options].some(o=>o.value===old))sel.value=old;
}
function renderStats(){
  const r=currentRecord(),total=r?.total??state.employees.length,used=r?.used??0,unused=r?.unused??total,events=r?.events??0,a=r?.adherence??0;
  $("statAdherence").textContent=r?`${a.toFixed(2)}%`:"—";$("statAdherenceMeta").textContent=r?formatDate(r.date):"Importe os logs";
  $("statTotal").textContent=total.toLocaleString("pt-BR");$("statUsed").textContent=used.toLocaleString("pt-BR");$("statUnused").textContent=unused.toLocaleString("pt-BR");$("statEvents").textContent=events.toLocaleString("pt-BR");
  $("indUsed").textContent=used;$("indAdherence").textContent=r?`${a.toFixed(2)}%`:"—";$("indUnused").textContent=unused;$("indTotal").textContent=total;$("indEvents").textContent=events;$("indDate").textContent=r?formatDate(r.date):"—";
  const tag=$("indAdherenceTag");tag.textContent=a>=80?"Boa":a>=50?"Atenção":"Crítico";tag.className=`tag ${a>=80?"operation":a>=50?"action":"critical"}`;
  const deg=Math.max(0,Math.min(100,a))*3.6;$("donut").style.background=`conic-gradient(var(--blue) 0deg ${deg}deg,#dfe1e5 ${deg}deg 360deg)`;$("donutPercent").textContent=r?`${Math.round(a)}%`:"—";
  $("dataStatus").textContent=`${state.employees.length} motoristas cadastrados`;
}
function renderHistoryTable(){
  const ds=historyDates();
  $("historyBody").innerHTML=ds.map(d=>{const r=state.history[d];return`<tr><td><strong>${formatDate(d)}</strong></td><td>${r.total}</td><td>${r.used}</td><td>${r.unused}</td><td><strong>${r.adherence.toFixed(2)}%</strong></td><td>${r.events}</td><td><button class="btn btn-soft" onclick="selectHistoryDate('${d}')">Abrir</button></td></tr>`}).join("")||`<tr><td class="empty" colspan="7">Nenhum dia importado.</td></tr>`;
  $("recentHistoryBody").innerHTML=ds.slice(0,7).map(d=>{const r=state.history[d];return`<tr onclick="selectHistoryDate('${d}')" style="cursor:pointer"><td>${formatDate(d)}</td><td>${r.total}</td><td>${r.used}</td><td>${r.unused}</td><td>${r.adherence.toFixed(2)}%</td><td>${r.events}</td></tr>`}).join("")||`<tr><td class="empty" colspan="6">Importe os logs do primeiro dia.</td></tr>`;
}
function renderLists(){
  const r=currentRecord(),rows=r?.drivers||state.employees.map(e=>({...e,name:e.nome,used:false,events:0,last:""}));
  const used=[...rows].filter(d=>d.used).sort((a,b)=>b.events-a.events||a.name.localeCompare(b.name,"pt-BR")),unused=[...rows].filter(d=>!d.used).sort((a,b)=>a.name.localeCompare(b.name,"pt-BR"));
  $("usedListCount").textContent=`${used.length} motoristas`;$("unusedListCount").textContent=`${unused.length} motoristas`;
  $("usedDriversList").innerHTML=used.map(d=>`<div class="driver-row"><div><strong>${esc(d.name)}</strong><span>${esc(d.cargo)}</span></div><div class="driver-log-count">${d.events} logs</div></div>`).join("")||`<div class="empty">Nenhum motorista usou o app neste dia.</div>`;
  $("unusedDriversList").innerHTML=unused.map(d=>`<div class="driver-row"><div><strong>${esc(d.name)}</strong><span>${esc(d.cargo)}</span></div><div class="status unused">Sem uso</div></div>`).join("")||`<div class="empty">Todos usaram o app.</div>`;
}
function renderDrivers(){
  const r=currentRecord(),rows=r?.drivers||state.employees.map(d=>({...d,name:d.nome,used:false,events:0,last:""})),q=normalize($("driverSearch").value),st=$("driverStatusFilter").value;
  const f=rows.filter(d=>(!q||normalize(d.name).includes(q)||normalize(d.cargo).includes(q))&&(st==="all"||(st==="used"&&d.used)||(st==="unused"&&!d.used)));
  $("driversBody").innerHTML=f.map(d=>`<tr><td><strong>${esc(d.name)}</strong></td><td>${esc(d.cargo)}</td><td>${esc(d.empresa||"—")}</td><td><span class="status ${d.used?"used":"unused"}">${d.used?"Usando":"Sem uso"}</span></td><td>${d.events}</td><td>${d.last?new Date(d.last).toLocaleString("pt-BR"):"—"}</td></tr>`).join("")||`<tr><td class="empty" colspan="6">Nenhum motorista.</td></tr>`;
}
function renderRanking(){
  const r=currentRecord(),rows=[...(r?.drivers||[])].sort((a,b)=>b.events-a.events||a.name.localeCompare(b.name,"pt-BR")),total=rows.reduce((s,d)=>s+d.events,0);
  $("rankingSummary").textContent=r?`${rows.filter(d=>d.events>0).length} usuários · ${total.toLocaleString("pt-BR")} logs`:"Sem dados";
  $("rankingBody").innerHTML=rows.map((d,i)=>`<tr><td class="rank-number ${i<3?"rank-top":""}">${i+1}</td><td><strong>${esc(d.name)}</strong></td><td>${esc(d.cargo)}</td><td><strong>${d.events}</strong></td><td>${total?pct(d.events,total).toFixed(2):"0.00"}%</td><td><span class="status ${d.used?"used":"unused"}">${d.used?"Usando":"Sem uso"}</span></td><td>${d.last?new Date(d.last).toLocaleString("pt-BR"):"—"}</td></tr>`).join("")||`<tr><td class="empty" colspan="7">Sem dados.</td></tr>`;
}
function renderCharts(){
  const ds=historyDates().sort(),labels=ds.map(formatDate),values=ds.map(d=>state.history[d].adherence);
  if(state.historyChart)state.historyChart.destroy();
  state.historyChart=new Chart($("historyChart"),{type:"line",data:{labels,datasets:[{label:"Adesão",data:values,borderWidth:2,tension:.35,pointRadius:3,fill:false}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"}},x:{grid:{display:false}}}}});
}
function renderIndividual(){
  const selected=$("driverSelect").value,ds=historyDates().sort();
  if(!selected){$("individualTitle").textContent="Selecione um motorista";$("individualSubtitle").textContent="Use o filtro acima para acompanhar a evolução individual."; $("individualSummary").innerHTML="";if(state.individualChart)state.individualChart.destroy();return}
  const key=normalize(selected),person=state.employees.find(e=>normalize(e.nome)===key),series=ds.map(d=>state.history[d].drivers.find(x=>normalize(x.name)===key)?.used?1:0),events=ds.map(d=>state.history[d].drivers.find(x=>normalize(x.name)===key)?.events||0),used=series.filter(Boolean).length;
  $("individualTitle").textContent=person?.nome||selected;$("individualSubtitle").textContent=person?`${person.cargo} · ${person.empresa||"Empresa não informada"}`:"";
  $("individualSummary").innerHTML=[["Dias usando",used],["Dias sem uso",Math.max(0,ds.length-used)],["Adesão individual",ds.length?((used/ds.length)*100).toFixed(1)+"%":"—"],["Eventos no histórico",events.reduce((a,b)=>a+b,0)]].map(x=>`<div class="individual-box"><span>${x[0]}</span><strong>${x[1]}</strong></div>`).join("");
  if(state.individualChart)state.individualChart.destroy();
  state.individualChart=new Chart($("individualChart"),{type:"bar",data:{labels:ds.map(formatDate),datasets:[{label:"Uso do app",data:series.map(v=>v*100),borderWidth:1}]},options:{responsive:true,maintainAspectRatio:false,plugins:{legend:{display:false}},scales:{y:{beginAtZero:true,max:100,ticks:{callback:v=>v+"%"}}}}});
}

function csvDownload(filename,rows){
  const csv="\ufeff"+rows.map(r=>r.map(v=>`"${String(v??"").replaceAll('"','""')}"`).join(";")).join("\n");
  const a=document.createElement("a");a.href=URL.createObjectURL(new Blob([csv],{type:"text/csv;charset=utf-8"}));a.download=filename;a.click();URL.revokeObjectURL(a.href);
}
function exportList(kind){
  const r=currentRecord();if(!r)return showToast("Selecione um dia com dados.");
  const rows=r.drivers.filter(d=>kind==="used"?d.used:!d.used).sort((a,b)=>kind==="used"?b.events-a.events:a.name.localeCompare(b.name,"pt-BR"));
  csvDownload(`${kind==="used"?"motoristas_usando":"motoristas_sem_uso"}_${r.date}.csv`,[["Motorista","Matrícula","Cargo","Empresa","Logs","Último registro"],...rows.map(d=>[d.name,d.matricula,d.cargo,d.empresa,d.events,d.last?new Date(d.last).toLocaleString("pt-BR"):""])]);
}
function exportRanking(){
  const r=currentRecord();if(!r)return showToast("Selecione um dia com dados.");
  const rows=[...r.drivers].sort((a,b)=>b.events-a.events||a.name.localeCompare(b.name,"pt-BR")),total=r.events;
  csvDownload(`ranking_logs_${r.date}.csv`,[["Ranking","Motorista","Matrícula","Cargo","Empresa","Logs","% dos logs","Status","Último registro"],...rows.map((d,i)=>[i+1,d.name,d.matricula,d.cargo,d.empresa,d.events,total?pct(d.events,total).toFixed(2)+"%":"0.00%",d.used?"Usando":"Sem uso",d.last?new Date(d.last).toLocaleString("pt-BR"):""])]);
}
function activateTab(n){document.querySelectorAll(".tab").forEach(b=>b.classList.toggle("active",b.dataset.tab===n));document.querySelectorAll(".tab-panel").forEach(p=>p.classList.toggle("active",p.id===`panel-${n}`))}
function updateUI(){populateDates();populateDrivers();renderBaseCheck();renderStats();renderHistoryTable();renderLists();renderDrivers();renderRanking();renderChecklist();renderCharts();renderIndividual()}
window.selectHistoryDate=d=>{state.selectedDate=d;$("dateSelect").value=d;updateUI();activateTab("overview")};

async function fullReset(){
  if(!confirm("RESET COMPLETO\n\nIsso apagará motoristas, logs e histórico do Supabase para todos os usuários.\n\nDeseja continuar?"))return;
  if(prompt("Digite RESET para confirmar:")!=="RESET")return;
  try{
    const {error:e1}=await db.from("app_logs").delete().not("id","is",null);if(e1)throw e1;
    const {error:e2}=await db.from("uso_diario").delete().not("id","is",null);if(e2)throw e2;
    const {error:e3}=await db.from("motoristas").delete().not("id","is",null);if(e3)throw e3;
    await loadCloudData();updateUI();showToast("Reset completo realizado no banco.");
  }catch(e){console.error(e);showToast("Erro no reset: "+e.message)}
}

function setAuthView(session){
  const loginScreen=$("loginScreen");
  const appShell=document.querySelector(".app-shell");
  if(session){
    state.user=session.user;
    loginScreen.classList.add("hidden");
    appShell.classList.remove("auth-hidden");
  }else{
    state.user=null;
    state.employees=[];
    state.history={};
    state.selectedDate="";
    appShell.classList.add("auth-hidden");
    loginScreen.classList.remove("hidden");
  }
}

async function startApp(){
  const {data,error}=await db.auth.getSession();
  if(error){
    console.error(error);
    $("loginError").textContent=error.message;
    setAuthView(null);
    return;
  }
  if(data.session){
    setAuthView(data.session);
    try{
      await loadCloudData();
      updateUI();
    }catch(e){
      console.error(e);
      showToast("Erro ao carregar Supabase: "+e.message);
    }
  }else{
    setAuthView(null);
  }
}

db.auth.onAuthStateChange((event,session)=>{
  if(event==="SIGNED_OUT"){
    setAuthView(null);
    $("loginEmail").value="";
    $("loginPassword").value="";
    $("loginError").textContent="";
    return;
  }

  if(session){
    setAuthView(session);
    setTimeout(async()=>{
      try{
        await loadCloudData();
        updateUI();
      }catch(e){
        console.error(e);
        showToast("Erro ao carregar Supabase: "+e.message);
      }
    },0);
  }else{
    setAuthView(null);
  }
});

$("loginForm").onsubmit=async e=>{
  e.preventDefault();
  const email=$("loginEmail").value.trim();
  const password=$("loginPassword").value;
  const errorBox=$("loginError");
  const button=$("loginSubmit");

  errorBox.textContent="Entrando...";
  button.disabled=true;

  try{
    const {error}=await db.auth.signInWithPassword({email,password});
    if(error) throw error;
    errorBox.textContent="";
    showToast("Login realizado.");
  }catch(error){
    console.error(error);
    errorBox.textContent=error.message||"Não foi possível entrar.";
  }finally{
    button.disabled=false;
  }
};

$("btnLogout").onclick=async()=>{
  const button=$("btnLogout");
  button.disabled=true;
  button.textContent="Saindo...";

  try{
    const {error}=await db.auth.signOut({scope:"local"});
    if(error) throw error;

    // Fallback explícito caso o listener demore a atualizar a interface.
    setAuthView(null);
    $("loginEmail").value="";
    $("loginPassword").value="";
    $("loginError").textContent="";
    showToast("Sessão encerrada.");
  }catch(error){
    console.error(error);
    showToast("Erro ao sair: "+(error.message||"não foi possível encerrar a sessão."));
  }finally{
    button.disabled=false;
    button.textContent="Sair";
  }
};

$("btnImportChecklist").onclick=()=>$("checklistInput").click();
$("checklistInput").onchange=e=>{if(e.target.files[0])parseChecklistRows(e.target.files[0]);e.target.value=""};
$("confirmChecklistImport").onclick=confirmChecklistImport;
$("cancelChecklistImport").onclick=()=>{$("checklistModal").classList.add("hidden");state.pendingChecklist=null};
$("closeChecklistModal").onclick=()=>{$("checklistModal").classList.add("hidden");state.pendingChecklist=null};
$("checklistModal").onclick=e=>{if(e.target.id==="checklistModal"){e.currentTarget.classList.add("hidden");state.pendingChecklist=null}};
$("checklistStatusFilter").onchange=renderChecklist;
$("checklistSearch").oninput=renderChecklist;
$("exportChecklistReport").onclick=exportChecklistReport;
$("btnImportEmployees").onclick=()=>$("employeesInput").click();
$("btnImportLogs").onclick=()=>$("logsInput").click();
$("employeesInput").onchange=e=>{if(e.target.files[0])importEmployees(e.target.files[0]);e.target.value=""};
$("logsInput").onchange=e=>{if(e.target.files[0])importLogs(e.target.files[0]);e.target.value=""};
$("confirmImport").onclick=confirmLogImport;$("cancelImport").onclick=()=>{$("dateModal").classList.add("hidden");state.pendingLogs=null};$("closeModal").onclick=()=>{$("dateModal").classList.add("hidden");state.pendingLogs=null};
$("dateSelect").onchange=e=>{state.selectedDate=e.target.value;updateUI()};
$("driverSelect").onchange=()=>{renderIndividual();activateTab("individual")};
$("driverSearch").oninput=renderDrivers;$("driverStatusFilter").onchange=renderDrivers;
$("btnExportHistory").onclick=()=>{const rows=[["Data","Total","Usando","Sem uso","Adesão","Eventos"],...historyDates().sort().map(d=>{const r=state.history[d];return[d,r.total,r.used,r.unused,r.adherence.toFixed(2)+"%",r.events]})];csvDownload("historico_adesao_motoristas.csv",rows)};
$("exportUsed").onclick=()=>exportList("used");$("exportUnused").onclick=()=>exportList("unused");$("exportRanking").onclick=exportRanking;
$("btnFullReset").onclick=fullReset;
$("periodSelect").onchange=e=>{state.selectedDate=historyDates().sort()[e.target.value==="all"?0:Math.max(0,historyDates().length-1)]||"";updateUI()};
document.querySelectorAll(".tab").forEach(b=>b.addEventListener("click",e=>{
  e.preventDefault();
  activateTab(b.dataset.tab);
  if(b.dataset.tab==="checklist") renderChecklist();
  if(b.dataset.tab==="individual") renderIndividual();
}));

if($("applyChecklistPeriod"))$("applyChecklistPeriod").onclick=applyChecklistPeriod;
if($("btnClearChecklist"))$("btnClearChecklist").onclick=clearAllChecklists;

startApp();
