(function(root,factory){
  "use strict";
  const api=factory();
  if(typeof module==="object"&&module.exports)module.exports=api;
  else root.VargaTenantSession=api;
})(typeof globalThis!=="undefined"?globalThis:this,function(){
  "use strict";
  const valid=s=>typeof s==="string"&&/^[A-Za-z0-9_-]{1,128}$/.test(s)&&!/^__.*__$/.test(s);
  function denied(){return Object.assign(new Error("session"),{code:"session"});}
  // Ephemeral selected-company context only. No localStorage, cached privileges or fallback.
  function create(getIdentity){
    if(typeof getIdentity!=="function")throw denied();
    let generation=0,selected=null;
    function clear(){generation++;selected=null;}
    function begin(tenantId){
      clear();const uid=getIdentity();
      if(!valid(uid)||!valid(tenantId))throw denied();
      return Object.freeze({generation,uid,tenantId});
    }
    function commit(ticket,server){
      if(!ticket||ticket.generation!==generation||getIdentity()!==ticket.uid||server?.company?.id!==ticket.tenantId||!["owner","admin","manager","team_lead","worker","client"].includes(server.company.role))throw denied();
      selected=Object.freeze({uid:ticket.uid,tenantId:ticket.tenantId,role:server.company.role,generation});
      return selected;
    }
    function requireTenant(tenantId){
      if(!selected||selected.generation!==generation||selected.uid!==getIdentity()||selected.tenantId!==tenantId)throw denied();
      return selected;
    }
    return Object.freeze({begin,commit,clear,requireTenant});
  }
  return Object.freeze({create});
});
