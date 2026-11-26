(() => {
  "use strict";
  const allowed=()=>window.VargaTenantLabEnabled===true&&["localhost","127.0.0.1"].includes(location.hostname)&&location.protocol==="http:"&&window.firebase?.apps?.length>0&&firebase.app().options.projectId==="demo-varga-multi-tenant";
  if(!allowed())return;
  const auth=firebase.auth(),el=id=>document.getElementById("varga-company-"+id);
  const open=el("open"),dialog=el("dialog"),entry=el("entry"),entryOpen=el("entry-open");if(!open||!dialog||!entry||!entryOpen)return;
  if(!window.VargaTenantSession)return;
  const session=window.VargaTenantSession.create(()=>auth.currentUser?.uid||null);
  let epoch=0,controller=null,active=null,kind=null,cursor=null,jobsCursor=null,moduleCursor=null,moduleQuery=null,membersCursor=null,workerCursor=null,myHoursCursor=null,hoursDraft=null,entryAllowed=false,entryGeneration=0;
  const message=text=>{el("status").textContent=text;};
  function cancel(){epoch++;controller?.abort();controller=null;}
  function clearModule(){clearOperation();clearEditor();el("new-record").hidden=true;moduleCursor=null;moduleQuery=null;el("module-title").textContent="";el("module-rows").replaceChildren();el("module-more").hidden=true;}
  function clearSelection(){clearWorker();clearMembers();el("plan").hidden=true;el("upgrade").reset();session.clear();clearModule();el("new-job").hidden=true;active=null;kind=null;jobsCursor=null;el("selected").hidden=true;el("actions").hidden=true;el("name").textContent="";el("role").textContent="";el("jobs").replaceChildren();el("jobs-more").hidden=true;}
  function clear(){cancel();clearSelection();el("join-form").reset();el("join").open=false;el("register").reset();cursor=null;el("list").replaceChildren();el("more").hidden=true;message("");}
  function hide(){entryGeneration++;entryAllowed=false;entry.hidden=true;el("entry-feedback").textContent="";entryOpen.disabled=false;clear();open.hidden=true;open.classList.add("hidden");if(dialog.open)dialog.close();}
  function current(n,uid){return allowed()&&n===epoch&&auth.currentUser?.uid===uid&&dialog.open;}
  async function call(action,payload,n,uid,signal){
    if(action!=="listMyCompanies"&&action!=="openCompany"&&action!=="registerCompany"&&action!=="accept"&&action!=="checkTenantEntry")session.requireTenant(payload.tenantId);
    const token=await auth.currentUser.getIdTokenResult();
    if(!current(n,uid)||token.claims.email_verified!==true||token.claims.apple_review===true)throw Error("session");
    const res=await fetch("/api/v2/companies",{method:"POST",headers:{"content-type":"application/json","authorization":"Bearer "+token.token},body:JSON.stringify({action,payload}),cache:"no-store",credentials:"omit",redirect:"error",signal});
    const data=await res.json();
    if(!current(n,uid))throw Error("session");
    if(action!=="listMyCompanies"&&action!=="openCompany"&&action!=="registerCompany"&&action!=="accept")session.requireTenant(payload.tenantId);
    if(!res.ok||data.ok!==true)throw Error(data.code||"request-failed");
    return data.result;
  }
  async function run(task){
    cancel();if(!allowed()||!dialog.open||(!entryAllowed&&open.hidden)||!auth.currentUser)return;
    const n=epoch,uid=auth.currentUser.uid,ctrl=new AbortController();controller=ctrl;
    const timer=setTimeout(()=>ctrl.abort(),12000);
    try{message("Caricamento…");await task((action,payload)=>call(action,payload,n,uid,ctrl.signal));}
    catch(e){if(current(n,uid)){clearSelection();if(e.message==="forbidden"||e.message==="unauthenticated"){el("list").replaceChildren();cursor=null;el("more").hidden=true;}message(e.message.startsWith("invalid-quota-")?"Limite raggiunto. Seleziona l’azienda e richiedi ampliamento.":"Operazione non confermata. Aggiorna l’elenco prima di riprovare.");}}
    finally{clearTimeout(timer);if(controller===ctrl)controller=null;}
  }
  function showCompany(data){
    active=data.company.id;el("name").textContent=data.company.name;el("role").textContent="Ruolo: "+data.company.role;
    el("selected").hidden=false;el("actions").hidden=!data.capabilities.readCommesse;el("worker-area").hidden=!data.capabilities.readAssignedPlants;el("grant-area").hidden=!data.plan||!["owner","admin"].includes(data.company.role);el("membership").hidden=!data.plan||!["owner","admin"].includes(data.company.role);el("invite-role").querySelector('option[value="admin"]').disabled=data.company.role!=="owner";
    message("Azienda verificata. L’area operativa completa non è ancora collegata.");
  }
  async function choose(id){clearSelection();await run(async request=>{const ticket=session.begin(id);const data=await request("openCompany",{tenantId:id});session.commit(ticket,data);showCompany(data);if(data.plan)renderPlan(data.plan);});}
  async function list(append=false){
    clearSelection();if(!append){el("list").replaceChildren();cursor=null;el("more").hidden=true;}
    await run(async request=>{
      const result=await request("listMyCompanies",{limit:25,after:append?cursor:null});
      for(const row of result.companies){const li=document.createElement("li"),b=document.createElement("button");b.type="button";b.className="btn";b.textContent=row.name+" · "+row.role;b.addEventListener("click",()=>void choose(row.id));li.append(b);el("list").append(li);}
      cursor=result.nextCursor||null;el("more").hidden=!cursor;
      message(el("list").children.length?"Seleziona un’azienda autorizzata.":cursor?"Continua per verificare altre appartenenze.":"Nessuna azienda autorizzata.");
    });
  }
  async function jobs(nextKind,append=false){
    const tenantId=active;if(!tenantId)return;
    clearModule();
    if(!append){kind=null;el("new-job").hidden=true;el("jobs").replaceChildren();jobsCursor=null;el("jobs-more").hidden=true;}
    await run(async request=>{
      const data=await request("listCompanyCommesse",{tenantId,kind:nextKind,limit:25,after:append?jobsCursor:null});
      if(active!==data.company.id||data.kind!==nextKind)throw Error("session");
      kind=nextKind;el("new-job").hidden=false;for(const row of data.commesse){
        const li=document.createElement("li"),name=document.createElement("span");name.textContent=row.name;const identifier=document.createElement("small");identifier.textContent=" · ID: "+row.id+" ";li.append(name,identifier);if(row.editable===true)addEdit(li,{module:"jobs",kind:nextKind,title:"Commessa"},row.id);
        for(const [module,label] of [["accounting","Contabilità"],["prices","Prezziario"],["plants","Cantieri / impianti"]]){
          const b=document.createElement("button");b.type="button";b.className="btn";b.textContent=label;
          b.addEventListener("click",()=>void showModule({module,kind:nextKind,commessaId:row.id,title:label+" · "+row.name}));li.append(b);
        }if(row.editable===true){for(const [operation,label] of [["plant","Nuovo impianto"],["team","Assegna squadra"],["hoursRequest","Richiedi ore"]]){const b=document.createElement("button");b.type="button";b.className="btn";b.textContent=label;b.addEventListener("click",()=>{cancel();clearEditor();openOperation({operation,kind:nextKind,commessaId:row.id,title:label+" · "+row.name});});li.append(b);}}el("jobs").append(li);
      }
      jobsCursor=data.nextCursor||null;el("jobs-more").hidden=!jobsCursor;
      message(el("jobs").children.length?"Commesse aziendali · sola lettura.":"Nessuna commessa in questa azienda.");
    });
  }
  const money=new Intl.NumberFormat("it-IT",{style:"currency",currency:"EUR"});
  function formatRow(row){
    const parts=[row.name,row.code,row.unit,row.category,row.status,row.date];
    if(typeof row.count==="number")parts.push(row.count+" elementi");if(row.operator)parts.push(row.operator);if(typeof row.hours==="number")parts.push(row.hours+" ore");
    if(typeof row.activitiesCount==="number")parts.push(row.activitiesCount+" attività");
    if(typeof row.base==="number")parts.push("Base "+money.format(row.base));
    if(typeof row.discounted==="number")parts.push("Ribassato "+money.format(row.discounted));
    if(typeof row.total==="number")parts.push("Totale "+money.format(row.total));
    return parts.filter(Boolean).join(" · ");
  }
  async function showModule(q,append=false){
    const tenantId=active;if(!tenantId)return;
    if(!append){clearModule();el("module-title").textContent=q.title;}
    await run(async request=>{
      const data=await request("listCompanyModule",{tenantId,module:q.module,kind:q.kind,commessaId:q.commessaId,month:q.month,limit:25,after:append?moduleCursor:null});
      if(active!==data.company.id||q.module!==data.module||q.kind!==data.kind||(q.commessaId??null)!==data.commessaId)throw Error("session");
      moduleQuery=q;el("new-record").hidden=!editableModules.has(q.module);for(const row of data.rows){const li=document.createElement("li");li.textContent=formatRow(row)+" · ID: "+row.id;if(editableModules.has(q.module)&&row.editable===true)addEdit(li,q,row.id);if(q.module==="approvals"&&row.reviewable===true)addHoursReview(li,row,q.kind,true);el("module-rows").append(li);}
      moduleCursor=data.nextCursor||null;el("module-more").hidden=!moduleCursor;
      message(el("module-rows").children.length?"Dati aziendali · sola lettura.":"Nessun dato presente.");
    });
  }
  for(const [module,label] of [["personnel","Personale"],["teams","Squadre"],["vehicles","Mezzi"],["hours","Ore registrate"],["approvals","Richieste ore"]]){
    el(module).addEventListener("click",()=>{el("jobs").replaceChildren();jobsCursor=null;kind=null;el("new-job").hidden=true;el("jobs-more").hidden=true;void showModule({module,kind:el("module-kind").value,title:label});});
  }
  const editableModules=new Set(["personnel","vehicles","prices"]);
  const definitions={
    jobs:[["nome","Nome"],["codice","Codice"],["parentCommessaId","ID commessa padre (facoltativo)"]],
    personnel:[["cognome","Cognome"],["nome","Nome"]],
    vehicles:[["nId","Identificativo"],["marca","Marca"],["modello","Modello"]],
    prices:[["codiceVoce","Codice voce"],["descrizione","Descrizione"],["unitaMisura","Unità di misura"],["prezzoBase","Prezzo base (€)"],["percentualeRibasso","Ribasso (0–1, facoltativo)"]]
  };
  let editing=null,saving=false;
  function clearEditor(){editing=null;saving=false;el("write").hidden=true;el("write-fields").replaceChildren();el("write-save").disabled=false;}
  function editForm(q,record=null){
    if(!active)return;clearOperation();clearEditor();
    editing={...q,tenantId:active,recordId:record?.recordId??null,revision:record?.revision??0,requestId:crypto.randomUUID()};
    el("write-title").textContent=(record?"Modifica ":"Nuovo · ")+q.title;
    for(const [field,label] of definitions[q.module]){
      const l=document.createElement("label"),input=document.createElement("input");l.textContent=label;input.name=field;input.value=record?.data?.[field]??"";
      if(["prezzoBase","percentualeRibasso"].includes(field)){input.type="number";input.min="0";input.step="any";if(field==="percentualeRibasso")input.max="1";}
      input.required=["nome","cognome","nId","codiceVoce","descrizione","unitaMisura","prezzoBase"].includes(field);
      if(field==="parentCommessaId"&&record)input.readOnly=true;
      l.append(input);el("write-fields").append(l);
    }
    el("write").hidden=false;
  }
  function addEdit(li,q,recordId){
    const b=document.createElement("button");b.type="button";b.className="btn";b.textContent="Modifica (lab)";
    b.addEventListener("click",()=>{const tenantId=active;if(!tenantId)return;clearEditor();void run(async request=>{const data=await request("getCompanyRecord",{tenantId,module:q.module,kind:q.kind,commessaId:q.commessaId,recordId});if(active!==data.company.id||q.module!==data.module)throw Error("session");editForm(q,data);message("Record di laboratorio verificato.");});});li.append(b);
  }
  el("new-job").addEventListener("click",()=>{if(kind){cancel();editForm({module:"jobs",kind,title:"Commessa"});}});
  el("new-record").addEventListener("click",()=>{if(moduleQuery&&editableModules.has(moduleQuery.module)){cancel();editForm(moduleQuery);}});
  el("write-cancel").addEventListener("click",()=>{if(saving)return;cancel();clearEditor();});
  el("write-form").addEventListener("submit",e=>{
    e.preventDefault();if(!editing||saving||active!==editing.tenantId)return;
    const q={...editing},data=Object.fromEntries(new FormData(el("write-form")));
    for(const key of ["prezzoBase","percentualeRibasso"])if(key in data)data[key]=data[key]===""?null:Number(data[key]);
    saving=true;el("write-save").disabled=true;
    void run(async request=>{
      const result=await request("saveCompanyRecord",{tenantId:q.tenantId,module:q.module,kind:q.kind,commessaId:q.commessaId,recordId:q.recordId,revision:q.revision,requestId:q.requestId,data});
      if(active!==result.company.id)throw Error("session");
      clearEditor();message("Salvataggio confermato. Aggiorna l’elenco per visualizzarlo.");
    });
  });
  let operationForm=null,operationBusy=false;
  const operationFields={plant:[["denominazione","Denominazione","text"],["idSap","ID SAP","text"],["comune","Comune","text"],["indirizzo","Indirizzo","text"],["gpsY","Latitudine","number"],["gpsX","Longitudine","number"]],team:[["date","Data","date"],["leaderId","ID caposquadra","text"],["personnelIds","ID operatori (separati da virgola)","text"],["vehicleIds","ID mezzi (separati da virgola)","text"],["plantIds","ID impianti (separati da virgola)","text"]],hoursRequest:[["date","Data","date"],["personnelId","ID operatore","text"],["hours","Ore (incrementi di 0,25)","number"]]};
  function clearOperation(){operationForm=null;operationBusy=false;el("operation").hidden=true;el("operation-fields").replaceChildren();el("operation-save").disabled=false;}
  function openOperation(q){
    clearOperation();if(!active)return;
    operationForm={...q,tenantId:active,requestId:crypto.randomUUID()};
    el("operation-title").textContent=q.title;
    for(const [key,label,type] of operationFields[q.operation]){
      const l=document.createElement("label"),i=document.createElement("input");l.textContent=label;i.name=key;i.type=type;if(type==="number")i.step="any";
      i.required=["denominazione","date","leaderId","personnelId","hours"].includes(key);
      l.append(i);el("operation-fields").append(l);
    }el("operation").hidden=false;
  }
  el("operation-cancel").addEventListener("click",()=>{if(!operationBusy){cancel();clearOperation();}});
  el("operation-form").addEventListener("submit",e=>{
    e.preventDefault();if(!operationForm||operationBusy||active!==operationForm.tenantId)return;
    const q={...operationForm},data=Object.fromEntries(new FormData(el("operation-form")));
    for(const k of ["personnelIds","vehicleIds","plantIds"])if(k in data)data[k]=data[k].split(",").map(x=>x.trim()).filter(Boolean);
    for(const k of ["gpsY","gpsX","hours"])if(k in data)data[k]=data[k]===""?null:Number(data[k]);
    operationBusy=true;el("operation-save").disabled=true;
    void run(async request=>{const r=await request("createCompanyOperation",{tenantId:q.tenantId,operation:q.operation,kind:q.kind,commessaId:q.commessaId,requestId:q.requestId,data});if(active!==r.company.id)throw Error("session");clearOperation();message("Registrazione confermata. ID: "+r.recordId);});
  });
  const initialMonth=new Date();el("month").value=initialMonth.getFullYear()+"-"+String(initialMonth.getMonth()+1).padStart(2,"0");
  for(const [module,label] of [["documents","Miei documenti"],["events","Miei eventi"],["calendar","Riepilogo calendario"],["notifications","Notifiche create da me"],["integrations","Stato integrazioni"]]){
    el(module).addEventListener("click",()=>{
      if(module==="calendar"&&!/^\d{4}-(0[1-9]|1[0-2])$/.test(el("month").value)){message("Seleziona un mese valido.");return;}
      el("jobs").replaceChildren();jobsCursor=null;kind=null;el("jobs-more").hidden=true;
      void showModule({module,kind:"ordinary",month:module==="calendar"?el("month").value:undefined,title:label});
    });
  }
  el("month").addEventListener("change",()=>{cancel();clearModule();});
  el("module-kind").addEventListener("change",()=>{cancel();clearModule();});
  el("module-more").addEventListener("click",()=>{if(moduleQuery&&moduleCursor)void showModule(moduleQuery,true);});
  function clearWorker(){clearHours();myHoursCursor=null;el("my-hours-list").replaceChildren();el("my-hours-more").hidden=true;el("bind-form").reset();workerCursor=null;el("worker-area").hidden=true;el("worker-list").replaceChildren();el("worker-more").hidden=true;el("grant-area").hidden=true;el("grant-form").reset();}
  async function loadWorker(append=false){
    if(!active)return;const tenantId=active;
    await run(async request=>{
      const result=await request("listMyPlants",{tenantId,limit:25,after:append?workerCursor:null});
      if(active!==result.company.id)throw Error("session");
      if(!append)el("worker-list").replaceChildren();
      for(const row of result.plants){const li=document.createElement("li"),label=document.createElement("span"),button=document.createElement("button");label.textContent=[row.name,row.commessa,row.comune,row.indirizzo,row.status].filter(Boolean).join(" · ");button.type="button";button.className="btn";button.textContent="Richiedi ore";button.addEventListener("click",()=>openHours(row));li.append(label,button);el("worker-list").append(li);}
      workerCursor=result.nextCursor||null;el("worker-more").hidden=!workerCursor;
      message(el("worker-list").children.length?"Cantieri autorizzati.":"Nessun cantiere assegnato.");
    });
  }
  el("worker-refresh").addEventListener("click",()=>void loadWorker());
  el("worker-more").addEventListener("click",()=>{if(workerCursor)void loadWorker(true);});
  el("grant-form").addEventListener("submit",e=>{
    e.preventDefault();if(!active)return;
    const tenantId=active,enabled=e.submitter?.value==="grant";
    if(!["grant","revoke"].includes(e.submitter?.value))return;
    const payload={tenantId,enabled,kind:el("grant-kind").value,memberUid:el("grant-user").value.trim(),commessaId:el("grant-job").value.trim(),plantId:el("grant-plant").value.trim()};
    void run(async request=>{await request("setPlantGrant",payload);el("grant-form").reset();message(enabled?"Accesso al cantiere autorizzato.":"Accesso al cantiere revocato.");});
  });
  function addHoursReview(li,row,kind,admin){
    if(row.status!=="pending_level1"&&!(admin&&row.status==="pending_admin"))return;
    const tenantId=active;
    for(const [decision,label] of [["approve",row.status==="pending_level1"?"Primo OK":"Approva definitivamente"],["reject","Rifiuta"]]){
      const button=document.createElement("button");button.type="button";button.className="btn";button.textContent=label;
      button.addEventListener("click",()=>{
        if(!active||active!==tenantId)return;
        let reason="";
        if(decision==="reject"){const value=window.prompt("Motivo del rifiuto (facoltativo)","");if(value===null)return;reason=value.trim();}
        if(!window.confirm(label+" per questa richiesta ore?"))return;
        void run(async request=>{const result=await request("reviewHours",{tenantId,kind,requestId:row.id,expectedStatus:row.status,decision,reason});if(active!==result.company.id)throw Error("session");li.remove();message("Esito salvato: "+({pending_admin:"in attesa admin",approved:"approvata",rejected:"rifiutata"}[result.status]||result.status)+". Aggiorna l’elenco.");});
      });li.append(button);
    }
  }
  function clearHours(){hoursDraft=null;el("my-hours-form").reset();el("my-hours-form").hidden=true;el("my-hours-plant").textContent="";}
  function openHours(row){
    if(!active)return;cancel();clearHours();
    hoursDraft={tenantId:active,kind:row.kind,commessaId:row.commessaId,plantId:row.id,requestId:crypto.randomUUID()};
    el("my-hours-plant").textContent=row.name+" · "+row.commessa;
    const now=new Date();el("my-hours-date").value=now.getFullYear()+"-"+String(now.getMonth()+1).padStart(2,"0")+"-"+String(now.getDate()).padStart(2,"0");el("my-hours-form").hidden=false;
  }
  el("my-hours-cancel").addEventListener("click",()=>{cancel();clearHours();});
  el("my-hours-form").addEventListener("submit",e=>{
    e.preventDefault();if(!hoursDraft||active!==hoursDraft.tenantId)return;
    const payload={...hoursDraft,date:el("my-hours-date").value,hours:Number(el("my-hours-value").value)};
    void run(async request=>{await request("submitMyHours",payload);clearHours();message("Richiesta ore inviata, in attesa di approvazione.");});
  });
  async function loadMyHours(append=false){
    if(!active)return;const tenantId=active,kind=el("my-hours-kind").value;
    if(!append){myHoursCursor=null;el("my-hours-list").replaceChildren();}
    await run(async request=>{
      const result=await request("listMyHours",{tenantId,kind,limit:25,after:append?myHoursCursor:null});
      if(active!==result.company.id||result.kind!==kind)throw Error("session");
      for(const row of result.rows){const li=document.createElement("li");li.textContent=[row.date,row.commessa,row.hours+" ore",({pending_level1:"In attesa primo OK",pending_admin:"In attesa admin",approved:"Approvata",rejected:"Rifiutata"}[row.status]||row.status)].join(" · ");addHoursReview(li,row,kind,false);el("my-hours-list").append(li);}
      myHoursCursor=result.nextCursor||null;el("my-hours-more").hidden=!myHoursCursor;
      message(el("my-hours-list").children.length?"Richieste personali aggiornate.":"Nessuna richiesta personale.");
    });
  }
  el("my-hours-kind").addEventListener("change",()=>{cancel();myHoursCursor=null;el("my-hours-list").replaceChildren();el("my-hours-more").hidden=true;});
  el("my-hours-refresh").addEventListener("click",()=>void loadMyHours());
  el("my-hours-more").addEventListener("click",()=>{if(myHoursCursor)void loadMyHours(true);});
  el("bind-form").addEventListener("submit",e=>{
    e.preventDefault();if(!active)return;
    const payload={tenantId:active,kind:el("bind-kind").value,memberUid:el("bind-user").value.trim(),personnelId:el("bind-person").value.trim()};
    void run(async request=>{await request("bindWorkerPersonnel",payload);el("bind-form").reset();message("Anagrafica collegata all’operatore.");});
  });
  function clearInvite(){el("invite-code").value="";el("invite-result").hidden=true;}
  function clearMembers(){clearInvite();membersCursor=null;el("membership").hidden=true;el("members-list").replaceChildren();el("members-more").hidden=true;el("invite-form").reset();}
  async function showMembers(append=false){
    if(!active)return;const tenantId=active;
    await run(async request=>{
      const result=await request("listCompanyMembers",{tenantId,limit:25,after:append?membersCursor:null});
      if(active!==result.company.id)throw Error("session");
      if(!append)el("members-list").replaceChildren();
      for(const row of result.members){const li=document.createElement("li");li.textContent=(row.email||row.id)+" · "+({owner:"Titolare",admin:"Admin",worker:"Operatore",manager:"Responsabile",team_lead:"Caposquadra",client:"Cliente"}[row.role]||row.role)+" · "+row.status+" · UID: "+row.id;el("members-list").append(li);}
      membersCursor=result.nextCursor||null;el("members-more").hidden=!membersCursor;message("Accessi aziendali aggiornati.");
    });
  }
  el("members-refresh").addEventListener("click",()=>void showMembers());
  el("members-more").addEventListener("click",()=>{if(membersCursor)void showMembers(true);});
  el("invite-hide").addEventListener("click",clearInvite);
  el("invite-copy").addEventListener("click",async()=>{
    const code=el("invite-code").value;if(!code||el("invite-result").hidden)return;
    try{await navigator.clipboard.writeText(code);if(el("invite-code").value===code)message("Codice copiato. Condividilo solo con il destinatario.");}
    catch{if(el("invite-code").value===code){el("invite-code").focus();el("invite-code").select();message("Seleziona e copia il codice manualmente.");}}
  });
  el("invite-form").addEventListener("submit",e=>{
    e.preventDefault();if(!active)return;const tenantId=active,recipientEmail=el("invite-email").value,role=el("invite-role").value;
    clearInvite();
    void run(async request=>{
      const result=await request("invite",{tenantId,recipientEmail,role});
      if(active!==result.tenantId||!(/^[a-f0-9]{64}$/.test(result.token)))throw Error("session");
      el("invite-form").reset();el("invite-code").value=result.tenantId+":"+result.token;el("invite-result").hidden=false;
      message("Invito creato. Il posto viene occupato solo quando il destinatario accetta.");
    });
  });
  el("join-form").addEventListener("submit",e=>{
    e.preventDefault();const raw=el("join-code").value.trim(),match=/^([A-Za-z0-9_-]{1,128}):([a-f0-9]{64})$/.exec(raw);
    el("join-code").value="";
    if(!match){message("Codice invito non valido.");return;}
    clearSelection();
    void run(async request=>{
      const result=await request("accept",{tenantId:match[1],inviteToken:match[2]});
      if(result.tenantId!==match[1])throw Error("session");
      const ticket=session.begin(result.tenantId),data=await request("openCompany",{tenantId:result.tenantId});
      session.commit(ticket,data);showCompany(data);if(data.plan)renderPlan(data.plan);el("join").open=false;message("Invito accettato. Sei entrato nell’azienda.");
    });
  });
  const quotaKeys=["jobs","plants","admins","operators"];
  function renderPlan(plan){
    el("plan").hidden=false;
    el("plan-summary").textContent="Commesse "+plan.usage.jobs+"/"+plan.limits.jobs+" · Impianti "+plan.usage.plants+"/"+plan.limits.plants+" · Admin "+plan.usage.admins+"/"+plan.limits.admins+" · Operatori "+plan.usage.operators+"/"+plan.limits.operators+(plan.pending?" · Richiesta in attesa":plan.lastDecision?" · Ultima richiesta: "+(plan.lastDecision.decision==="approve"?"approvata":"rifiutata"):"");
    el("upgrade").hidden=!!plan.pending;
    for(const k of quotaKeys){el("quota-"+k).min=String(plan.limits[k]);el("quota-"+k).value=String(plan.limits[k]);}
  }
  el("register").addEventListener("submit",e=>{e.preventDefault();const name=el("register-name").value;void run(async request=>{const result=await request("registerCompany",{name});el("register").reset();const id=result.tenantId;clearSelection();const ticket=session.begin(id);const data=await request("openCompany",{tenantId:id});session.commit(ticket,data);showCompany(data);if(data.plan)renderPlan(data.plan);message("Azienda pronta nel laboratorio.");});});
  el("upgrade").addEventListener("submit",e=>{e.preventDefault();if(!active)return;const tenantId=active,limits=Object.fromEntries(quotaKeys.map(k=>[k,Number(el("quota-"+k).value)])),reason=el("upgrade-reason").value;void run(async request=>{await request("requestPlanUpgrade",{tenantId,limits,reason});renderPlan((await request("getCompanyPlan",{tenantId})).plan);message("Richiesta inviata al Super Admin.");});});
  // Tenant dialog is a separate surface; the original approval lock remains in place.
  // Never toggle the legacy screen, Home, auth classes or a platformUsers profile.
  entryOpen.addEventListener("click",async()=>{
    if(!allowed()||entry.hidden||entryOpen.disabled||!auth.currentUser)return;
    const generation=++entryGeneration,uid=auth.currentUser.uid;
    entryOpen.disabled=true;el("entry-feedback").textContent="Verifica accesso aziendale…";
    try{
      const token=await auth.currentUser.getIdTokenResult();
      if(token.claims.email_verified!==true||token.claims.apple_review===true)throw Error("forbidden");
      const response=await fetch("/api/v2/companies",{method:"POST",headers:{"content-type":"application/json","authorization":"Bearer "+token.token},body:JSON.stringify({action:"checkTenantEntry",payload:{}}),cache:"no-store",credentials:"omit",redirect:"error",signal:AbortSignal.timeout(12000)});
      const result=await response.json();
      if(generation!==entryGeneration||auth.currentUser?.uid!==uid||!allowed())return;
      if(!response.ok||result.ok!==true||result.result?.eligible!==true)throw Error("forbidden");
      entryAllowed=true;el("entry-feedback").textContent="";clear();dialog.showModal();void list();
    }catch{if(generation===entryGeneration&&auth.currentUser?.uid===uid)el("entry-feedback").textContent="Accesso aziendale non disponibile. Verifica l’email o riprova.";}
    finally{if(generation===entryGeneration)entryOpen.disabled=false;}
  });
  open.addEventListener("click",()=>{if(open.hidden||!allowed())return;clear();dialog.showModal();void list();});
  el("close").addEventListener("click",()=>dialog.close());dialog.addEventListener("close",()=>{entryAllowed=false;clear();});
  el("refresh").addEventListener("click",()=>void list());
  el("more").addEventListener("click",()=>{if(cursor)void list(true);});
  el("ordinary").addEventListener("click",()=>void jobs("ordinary"));
  el("snow").addEventListener("click",()=>void jobs("snow"));
  el("jobs-more").addEventListener("click",()=>{if(kind&&jobsCursor)void jobs(kind,true);});
  auth.onIdTokenChanged(async user=>{
    hide();const n=epoch;if(!user||!allowed())return;
    try{const t=await user.getIdTokenResult();if(n===epoch&&auth.currentUser?.uid===user.uid&&t.claims.email_verified===true&&t.claims.apple_review!==true){open.hidden=false;open.classList.remove("hidden");entry.hidden=false;}}catch{}
  },hide);
})();
