(() => {
  "use strict";
  const allowed=()=>window.VargaTenantLabEnabled===true&&["localhost","127.0.0.1"].includes(location.hostname)&&location.protocol==="http:"&&window.firebase?.apps?.length>0&&firebase.app().options.projectId==="demo-varga-multi-tenant";
  if(!allowed())return;
  const auth=firebase.auth(),el=id=>document.getElementById("varga-company-"+id);
  const open=el("open"),dialog=el("dialog");if(!open||!dialog)return;
  let epoch=0,controller=null,active=null,kind=null,cursor=null,jobsCursor=null;
  const message=text=>{el("status").textContent=text;};
  function cancel(){epoch++;controller?.abort();controller=null;}
  function clearSelection(){active=null;kind=null;jobsCursor=null;el("selected").hidden=true;el("actions").hidden=true;el("name").textContent="";el("role").textContent="";el("jobs").replaceChildren();el("jobs-more").hidden=true;}
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
    if(!append){el("jobs").replaceChildren();jobsCursor=null;el("jobs-more").hidden=true;}
    await run(async request=>{
      const data=await request("listCompanyCommesse",{tenantId,kind:nextKind,limit:25,after:append?jobsCursor:null});
      if(active!==data.company.id||data.kind!==nextKind)throw Error("session");
      kind=nextKind;for(const row of data.commesse){const li=document.createElement("li");li.textContent=row.name;el("jobs").append(li);}
      jobsCursor=data.nextCursor||null;el("jobs-more").hidden=!jobsCursor;
      message(el("jobs").children.length?"Commesse aziendali · sola lettura.":"Nessuna commessa in questa azienda.");
    });
  }
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
