// Oracle onboarding/data client: Firebase Auth supplies identity only.
// No Firestore fallback, no persisted bearer token, no automatic write retries.
export function createOracleClient({auth,origin,request=fetch}){
 const url=new URL(origin);
 const local=['127.0.0.1','localhost'].includes(url.hostname);
 if((url.protocol!=='https:'&&!(local&&url.protocol==='http:'))||url.username||url.password||url.pathname!=='/'||url.search||url.hash)throw Error('INVALID_API_ORIGIN');
 const originOnly=url.origin;
 const segment=v=>encodeURIComponent(String(v));
 async function call(method,path,data){
  const user=auth.currentUser;if(!user)throw Error('LOGIN_REQUIRED');
  const token=await user.getIdToken();
  if(auth.currentUser!==user)throw Error('SESSION_CHANGED');
  const response=await request(originOnly+'/v1/'+path,{method,headers:{Authorization:'Bearer '+token,'Content-Type':'application/json'},cache:'no-store',redirect:'error',signal:AbortSignal.timeout(20000),...(data===undefined?{}:{body:JSON.stringify(data)})});
  if(auth.currentUser!==user)throw Error('SESSION_CHANGED');
  const result=await response.json();
  if(auth.currentUser!==user)throw Error('SESSION_CHANGED');
  if(!response.ok){const e=new Error(result.error||'API_ERROR');e.code=result.error||'API_ERROR';e.status=response.status;throw e}
  return result;
 }
 const company=id=>'companies/'+segment(id);
 return Object.freeze({
  listCompanies:()=>call('GET','companies'),
  createCompany:(name,requestId)=>call('POST','companies',{name,requestId}),
  requestJoin:code=>call('POST','join',{code}),
  listRequests:id=>call('GET',company(id)+'/requests'),
  decideRequest:(id,uid,decision)=>call('POST',company(id)+'/requests/'+segment(uid),{decision}),
  listJobs:id=>call('GET',company(id)+'/jobs'),
  createJob:(id,name)=>call('POST',company(id)+'/jobs',{name}),
  listPlants:(id,job)=>call('GET',company(id)+'/jobs/'+segment(job)+'/plants'),
  createPlant:(id,job,name,lat=null,lng=null)=>call('POST',company(id)+'/jobs/'+segment(job)+'/plants',{name,lat,lng}),
  listAllCompanies:()=>call('GET','admin/companies'),
  setPlan:(id,plan,limits={})=>call('POST','admin/'+company(id)+'/plan',{plan,commesse:limits.commesse,impianti:limits.impianti,operatori:limits.operatori})
 });
}
