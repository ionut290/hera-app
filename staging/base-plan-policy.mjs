// Staging-only domain policy for the new-company base plan.
// This is NOT a Firestore rule or a client authorization gate.
// Call only from an authenticated backend transaction using authoritative
// persisted limits and usage, including reserved invitations. The transaction
// must atomically persist the resource, its counter, and any request decision.
// Existing Avola company, protected FATTO/WHAZZUP and production remain untouched.
export const BASE_PLAN=Object.freeze({"commesse":1,"impianti":10,"admin":1,"operatori":3});
export function validateLimits(limits){
 if(!limits||typeof limits!=="object"||Object.keys(limits).length!==4)throw Error("INVALID_LIMITS");
 for(const key of Object.keys(BASE_PLAN)){if(!Object.hasOwn(limits,key)||!Number.isSafeInteger(limits[key])||limits[key]<BASE_PLAN[key])throw Error("INVALID_LIMITS");}
 return Object.fromEntries(Object.keys(BASE_PLAN).map(key=>[key,limits[key]]));
}

export function validateUsage(usage){
 if(!usage||typeof usage!=="object"||Object.keys(usage).length!==4)throw Error("INVALID_USAGE");
 for(const key of Object.keys(BASE_PLAN)){if(!Object.hasOwn(usage,key)||!Number.isSafeInteger(usage[key])||usage[key]<0)throw Error("INVALID_USAGE");}
 return Object.fromEntries(Object.keys(BASE_PLAN).map(key=>[key,usage[key]]));
}

export function quotaAfterCreate({limits,usage,kind,amount=1}){
 const max=validateLimits(limits),current=validateUsage(usage);
 if(!Object.hasOwn(BASE_PLAN,kind)||!Number.isSafeInteger(amount)||amount<1)throw Error("INVALID_INCREMENT");
 const next=current[kind]+amount;
 if(!Number.isSafeInteger(next)||next>max[kind])throw Error("SUPER_ADMIN_APPROVAL_REQUIRED");
 return {...current,[kind]:next};
}

export function memberQuotaKind(role){
 if(role==="admin")return "admin";
 if(role==="employee"||role==="supervisor")return "operatori";
 throw Error("INVALID_COMPANY_ROLE");
}

export function validateExpansion({currentLimits,requestedLimits}){
 const current=validateLimits(currentLimits),requested=validateLimits(requestedLimits);
 for(const key of Object.keys(BASE_PLAN)){if(requested[key]<current[key])throw Error("LIMIT_DECREASE_NOT_SUPPORTED");}
 if(!Object.keys(BASE_PLAN).some(key=>requested[key]>current[key]))throw Error("NO_EXPANSION_REQUESTED");
 return requested;
}

export function decideExpansion({actor,tenantId,request,currentLimits,decision}){
 // actor MUST originate in verified server-side authentication, never an HTTP body.
 if(actor?.superAdmin!==true||actor?.email_verified!==true||actor?.apple_review===true||typeof actor?.uid!=="string"||!actor.uid)throw Error("SUPER_ADMIN_REQUIRED");
 if(typeof tenantId!=="string"||!tenantId||request?.tenantId!==tenantId||request?.status!=="pending")throw Error("INVALID_PENDING_REQUEST");
 if(decision!=="approved"&&decision!=="rejected")throw Error("INVALID_DECISION");
 const current=validateLimits(currentLimits);
 if(decision==="approved"){
  const baseline=validateLimits(request.baseLimits);
  if(Object.keys(BASE_PLAN).some(k=>baseline[k]!==current[k]))throw Error("STALE_REQUEST");
 }
 const limits=decision==="approved"?validateExpansion({currentLimits:current,requestedLimits:request.requestedLimits}):current;
 return {limits,requestPatch:{status:decision,reviewedBy:actor.uid}};
}
