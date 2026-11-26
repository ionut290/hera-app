(() => {
  "use strict";
  // A separate self-service registration surface, enabled only in the local 2.0 lab.
  const allowed=()=>window.VargaTenantLabEnabled===true&&["localhost","127.0.0.1"].includes(location.hostname)&&location.protocol==="http:"&&window.firebase?.apps?.length>0&&firebase.app().options.projectId==="demo-varga-multi-tenant";
  if(!allowed())return;
  const auth=firebase.auth(),el=k=>document.getElementById("varga-signup-"+k);
  const open=el("open"),dialog=el("dialog"),form=el("form");if(!open||!dialog||!form)return;
  let busy=false,generation=0;
  function message(value){el("status").textContent=value;}
  function reset(){generation++;form.reset();message("");el("password").value="";el("confirm").value="";}
  function lock(value){busy=value;for(const e of form.querySelectorAll("button,input"))e.disabled=value;el("cancel").disabled=value;open.disabled=value;}
  open.addEventListener("click",()=>{if(!allowed()||busy||auth.currentUser)return;reset();dialog.showModal();});
  el("cancel").addEventListener("click",()=>{if(!busy)dialog.close();});
  dialog.addEventListener("cancel",e=>{if(busy)e.preventDefault();});
  dialog.addEventListener("close",reset);
  auth.onIdTokenChanged(user=>{open.hidden=!!user||!allowed();},()=>{open.hidden=true;});
  form.addEventListener("submit",async e=>{
    e.preventDefault();if(!allowed()||busy||auth.currentUser||!dialog.open)return;
    const mode=e.submitter?.value||"register",email=el("email").value.trim().toLowerCase(),password=el("password").value;
    if(!["register","resend"].includes(mode)||!email||password.length<10){message("Inserisci un’email e una password di almeno 10 caratteri.");return;}
    if(mode==="register"&&password!==el("confirm").value){message("Le password non coincidono.");return;}
    const n=++generation;let createdOrSignedIn=false;lock(true);message("Operazione in corso…");
    // Auth only: no writes to platformUsers, no legacy approval or tenant creation.
    // A verification failure never deletes the newly-created account.
    try{
      const credential=mode==="register"?await auth.createUserWithEmailAndPassword(email,password):await auth.signInWithEmailAndPassword(email,password);
      createdOrSignedIn=true;
      if(credential.user.emailVerified!==true)await credential.user.sendEmailVerification();
      await auth.signOut();createdOrSignedIn=false;
      if(n===generation&&dialog.open){el("password").value="";el("confirm").value="";message("Controlla l’email, conferma l’indirizzo e poi torna al login. Se è già verificata, puoi accedere.");}
    }catch(error){
      if(createdOrSignedIn){try{await auth.signOut();}catch{}}
      if(n!==generation||!dialog.open)return;
      el("password").value="";el("confirm").value="";
      message(error?.code==="auth/email-already-in-use"?"Account già esistente. Accedi oppure usa Invia di nuovo la verifica.":"Operazione non confermata. Riprova o invia di nuovo la verifica.");
    }finally{lock(false);}
  });
})();
