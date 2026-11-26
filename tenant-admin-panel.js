(() => {
  "use strict";
  // Dormant outside an explicitly enabled loopback emulator harness.
  const allowed = () => window.VargaTenantLabEnabled === true &&
    ["localhost","127.0.0.1"].includes(location.hostname) &&
    location.protocol === "http:" &&
    window.firebase?.apps?.length > 0 &&
    firebase.app().options.projectId === "demo-varga-multi-tenant";
  if (!allowed()) return;
  const auth = firebase.auth();
  const el = key => document.getElementById("varga-tenant-admin-" + key);
  const open = el("open"), dialog = el("dialog");
  if (!open || !dialog) return;
  let epoch=0, busy=false, cursor=null, pending=null,planTarget=null;
  function message(value){el("status").textContent=value;}
  function reset(){
    epoch++; busy=false; cursor=null; pending=null;planTarget=null;el("plan").hidden=true;el("plan-form").reset();
    el("list").replaceChildren(); el("more").hidden=true;
    el("form").reset(); el("state").reset(); el("state").hidden=true;
    el("target").textContent=""; message("");
    el("refresh").disabled=false;
    for(const form of [el("form"),el("state")]) for(const b of form.querySelectorAll("button")) b.disabled=false;
  }
  function hide(){reset();open.hidden=true;open.classList.add("hidden");if(dialog.open)dialog.close();}
  function current(n,uid){return allowed() && n===epoch && auth.currentUser?.uid===uid;}
  async function check(user){
    hide(); const n=epoch;
    if(!user||!allowed())return;
    try{
      const result=await user.getIdTokenResult();
      if(current(n,user.uid)&&result.claims.platform_admin===true&&result.claims.email_verified===true&&result.claims.apple_review!==true){
        open.hidden=false;open.classList.remove("hidden");
      }
    }catch{ /* fail closed */ }
  }
  async function request(action,payload,n,uid){
    if(!current(n,uid))throw Error("session");
    const result=await auth.currentUser.getIdTokenResult();
    if(!current(n,uid)||result.claims.platform_admin!==true||result.claims.email_verified!==true||result.claims.apple_review===true)throw Error("forbidden");
    const res=await fetch("/api/v2/companies",{method:"POST",headers:{"content-type":"application/json","authorization":"Bearer "+result.token},body:JSON.stringify({action,payload}),cache:"no-store",credentials:"omit",redirect:"error",signal:AbortSignal.timeout(12000)});
    if(!current(n,uid))throw Error("session");
    const data=await res.json();
    if(!res.ok||data.ok!==true)throw Error(data.code||"request-failed");
    return data.result;
  }
  function render(rows){
    for(const row of rows){
      const li=document.createElement("li"),name=document.createElement("span"),button=document.createElement("button");
      name.textContent=row.name+" · "+(row.status==="active"?"Attiva":"Sospesa")+" ";
      button.type="button";button.className="btn";button.textContent=row.status==="active"?"Sospendi":"Riattiva";
      button.addEventListener("click",()=>{
        if(busy||!dialog.open)return;
        pending={tenantId:row.id,status:row.status==="active"?"suspended":"active"};
        el("target").textContent=button.textContent+" "+row.name+" (nessuna cancellazione)";
        el("reason").value="";el("state").hidden=false;
      });
      const quota=document.createElement("button");quota.type="button";quota.className="btn";quota.textContent="Piano / richieste";quota.addEventListener("click",()=>void perform(async(n,u)=>{planTarget=null;el("plan").hidden=true;const data=await request("getPlanForAdmin",{tenantId:row.id},n,u);if(!current(n,u))return;const p=data.plan;el("plan").hidden=false;el("plan-summary").textContent=row.name+" · "+JSON.stringify(p.limits)+(p.pending?" → "+JSON.stringify(p.pending.limits)+" · "+p.pending.reason:" · Nessuna richiesta");el("plan-form").hidden=!p.pending;if(p.pending)planTarget={tenantId:row.id,requestId:p.pending.id};}));li.append(name,button,quota);el("list").append(li);
    }
  }
  async function list(n,uid,append=false){
    const data=await request("listCompanies",{limit:25,after:append?cursor:null},n,uid);
    if(!current(n,uid))return;
    if(!append)el("list").replaceChildren();
    render(data.companies);cursor=data.nextCursor||null;el("more").hidden=!cursor;
    message(data.companies.length?"Elenco aggiornato.":append?"Nessun altro risultato.":"Nessuna azienda presente.");
  }
  async function perform(task){
    if(busy||!dialog.open||open.hidden||!auth.currentUser)return;
    const n=epoch,uid=auth.currentUser.uid;busy=true;
    el("refresh").disabled=true;el("more").disabled=true;
    for(const form of [el("form"),el("state")])for(const b of form.querySelectorAll("button"))b.disabled=true;
    try{message("Operazione in corso…");await task(n,uid);}
    catch(err){if(current(n,uid)){if(["forbidden","unauthenticated","session"].includes(err.message)){hide();return;}message("Operazione non confermata. Aggiorna l'elenco prima di riprovare.");}}
    finally{if(current(n,uid)){busy=false;el("refresh").disabled=false;el("more").disabled=false;for(const form of [el("form"),el("state")])for(const b of form.querySelectorAll("button"))b.disabled=false;}}
  }
  open.addEventListener("click",()=>{if(open.hidden||!allowed())return;reset();dialog.showModal();void perform(list);});
  el("close").addEventListener("click",()=>dialog.close());
  dialog.addEventListener("close",reset);
  el("refresh").addEventListener("click",()=>void perform(list));
  el("more").addEventListener("click",()=>void perform((n,u)=>list(n,u,true)));
  el("cancel").addEventListener("click",()=>{pending=null;el("state").hidden=true;});
  el("form").addEventListener("submit",e=>{e.preventDefault();const name=el("name").value.trim(),ownerEmail=el("email").value.trim();void perform(async(n,u)=>{await request("createCompany",{name,ownerEmail},n,u);if(!current(n,u))return;el("form").reset();await list(n,u);});});
  el("state").addEventListener("submit",e=>{e.preventDefault();if(!pending)return;const payload={...pending,reason:el("reason").value.trim()};void perform(async(n,u)=>{await request("setCompanyStatus",payload,n,u);if(!current(n,u))return;pending=null;el("state").hidden=true;await list(n,u);});});
  el("plan-form").addEventListener("submit",e=>{e.preventDefault();if(!planTarget)return;const decision=e.submitter?.value;if(!["approve","reject"].includes(decision))return;const payload={...planTarget,decision,reason:el("plan-reason").value};void perform(async(n,u)=>{await request("reviewPlanUpgrade",payload,n,u);planTarget=null;el("plan").hidden=true;el("plan-form").reset();await list(n,u);});});
  auth.onIdTokenChanged(user=>void check(user),hide);
})();
