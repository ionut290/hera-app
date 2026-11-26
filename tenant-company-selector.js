(() => {
  "use strict";
  const allowed=()=>window.VargaTenantLabEnabled===true&&["localhost","127.0.0.1"].includes(location.hostname)&&location.protocol==="http:"&&window.firebase?.apps?.length>0&&firebase.app().options.projectId==="demo-varga-multi-tenant";
  if(!allowed())return;
  const auth=firebase.auth(),el=id=>document.getElementById("varga-company-"+id);
  const open=el("open"),dialog=el("dialog");if(!open||!dialog)return;
  let epoch=0,controller=null,active=null,kind=null,cursor=null,jobsCursor=null,moduleCursor=null,moduleQuery=null;
  const message=text=>{el("status").textContent=text;};
  function cancel(){epoch++;controller?.abort();controller=null;}
  function clearModule(){moduleCursor=null;moduleQuery=null;el("module-title").textContent="";el("module-rows").replaceChildren();el("module-more").hidden=true;}
  function clearSelection(){clearModule();active=null;kind=null;jobsCursor=null;el("selected").hidden=true;el("actions").hidden=true;el("name").textContent="";el("role").textContent="";el("jobs").replaceChildren();el("jobs-more").hidden=true;}
  function clear(){cancel();clearSelection();cursor=null;el("list").replaceChildren();el("more").hidden=true;message("");}
  function hide(){clear();open.hidden=true;open.classList.add("hidden");if(dialog.open)dialog.close();}
  function current(n,uid){return allowed()&&n===epoch&&auth.currentUser?.uid===uid&&dialog.open;}
  async function call(action,payload,n,uid,signal){
    const token=await auth.currentUser.getIdTokenResult();
    if(!current(n,uid)||token.claims.email_verified!==true||token.claims.apple_review===true)throw Error("session");
    const res=await fetch("/api/v2/companies",{method:"POST",headers:{"content-type":"application/json","authorization":"Bearer "+token.token},body:JSON.stringify({action,payload}),cache:"no-store",credentials:"omit",redirect:"error",signal});
    const data=await res.json();
    if(!current(n,uid))throw Error("session");
    if(!res.ok||data.ok!==true)throw Error(data.code||"request-failed");
    return data.result;
  }
  async function run(task){
    cancel();if(!allowed()||!dialog.open||open.hidden||!auth.currentUser)return;
    const n=epoch,uid=auth.currentUser.uid,ctrl=new AbortController();controller=ctrl;
    const timer=setTimeout(()=>ctrl.abort(),12000);
    try{message("Caricamento…");await task((action,payload)=>call(action,payload,n,uid,ctrl.signal));}
    catch(e){if(current(n,uid)){clearSelection();if(e.message==="forbidden"||e.message==="unauthenticated"){el("list").replaceChildren();cursor=null;el("more").hidden=true;}message("Accesso non disponibile. Aggiorna l’elenco.");}}
    finally{clearTimeout(timer);if(controller===ctrl)controller=null;}
  }
  function showCompany(data){
    active=data.company.id;el("name").textContent=data.company.name;el("role").textContent="Ruolo: "+data.company.role;
    el("selected").hidden=false;el("actions").hidden=!data.capabilities.readCommesse;
    message("Azienda verificata. L’area operativa completa non è ancora collegata.");
  }
  async function choose(id){clearSelection();await run(async request=>showCompany(await request("openCompany",{tenantId:id})));}
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
    if(!append){el("jobs").replaceChildren();jobsCursor=null;el("jobs-more").hidden=true;}
    await run(async request=>{
      const data=await request("listCompanyCommesse",{tenantId,kind:nextKind,limit:25,after:append?jobsCursor:null});
      if(active!==data.company.id||data.kind!==nextKind)throw Error("session");
      kind=nextKind;for(const row of data.commesse){
        const li=document.createElement("li"),name=document.createElement("span");name.textContent=row.name+" ";li.append(name);
        for(const [module,label] of [["accounting","Contabilità"],["prices","Prezziario"]]){
          const b=document.createElement("button");b.type="button";b.className="btn";b.textContent=label;
          b.addEventListener("click",()=>void showModule({module,kind:nextKind,commessaId:row.id,title:label+" · "+row.name}));li.append(b);
        }el("jobs").append(li);
      }
      jobsCursor=data.nextCursor||null;el("jobs-more").hidden=!jobsCursor;
      message(el("jobs").children.length?"Commesse aziendali · sola lettura.":"Nessuna commessa in questa azienda.");
    });
  }
  const money=new Intl.NumberFormat("it-IT",{style:"currency",currency:"EUR"});
  function formatRow(row){
    const parts=[row.name,row.code,row.unit,row.category,row.status,row.date];
    if(typeof row.count==="number")parts.push(row.count+" elementi");
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
      moduleQuery=q;for(const row of data.rows){const li=document.createElement("li");li.textContent=formatRow(row);el("module-rows").append(li);}
      moduleCursor=data.nextCursor||null;el("module-more").hidden=!moduleCursor;
      message(el("module-rows").children.length?"Dati aziendali · sola lettura.":"Nessun dato presente.");
    });
  }
  for(const [module,label] of [["personnel","Personale"],["teams","Squadre"],["vehicles","Mezzi"],["hours","Ore registrate"],["approvals","Richieste ore"]]){
    el(module).addEventListener("click",()=>{el("jobs").replaceChildren();jobsCursor=null;kind=null;el("jobs-more").hidden=true;void showModule({module,kind:el("module-kind").value,title:label});});
  }
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
  open.addEventListener("click",()=>{if(open.hidden||!allowed())return;clear();dialog.showModal();void list();});
  el("close").addEventListener("click",()=>dialog.close());dialog.addEventListener("close",clear);
  el("refresh").addEventListener("click",()=>void list());
  el("more").addEventListener("click",()=>{if(cursor)void list(true);});
  el("ordinary").addEventListener("click",()=>void jobs("ordinary"));
  el("snow").addEventListener("click",()=>void jobs("snow"));
  el("jobs-more").addEventListener("click",()=>{if(kind&&jobsCursor)void jobs(kind,true);});
  auth.onIdTokenChanged(async user=>{
    hide();const n=epoch;if(!user||!allowed())return;
    try{const t=await user.getIdTokenResult();if(n===epoch&&auth.currentUser?.uid===user.uid&&t.claims.email_verified===true&&t.claims.apple_review!==true){open.hidden=false;open.classList.remove("hidden");}}catch{}
  },hide);
})();
