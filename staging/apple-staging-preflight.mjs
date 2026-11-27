// Apple metadata preflight only. No certificate/profile bytes, no POST/PATCH/DELETE.
import {sign} from 'node:crypto';
import {fileURLToPath} from 'node:url';
const BUNDLE='it.vargacantieri.staging',API='https://api.appstoreconnect.apple.com';
export async function checkAppleStaging({request=fetch,env=process.env}={}){
 const {IOS_APPSTORE_KEY_ID:k,IOS_APPSTORE_ISSUER_ID:i,IOS_APPSTORE_PRIVATE_KEY_BASE64:b}=env;
 if(!k||!i||!b)return {status:'missing-credentials'};
 const key=Buffer.from(b,'base64').toString('utf8');
 if(!key.includes('-----BEGIN PRIVATE KEY-----'))return {status:'invalid-credentials'};
 const now=Math.floor(Date.now()/1000),encode=x=>Buffer.from(JSON.stringify(x)).toString('base64url');
 const data=encode({alg:'ES256',kid:k,typ:'JWT'})+'.'+encode({iss:i,aud:'appstoreconnect-v1',iat:now,exp:now+300});
 const jwt=data+'.'+sign('sha256',Buffer.from(data),{key,dsaEncoding:'ieee-p1363'}).toString('base64url');
 async function get(path,params){
  const url=new URL(path,API);for(const [k,v] of Object.entries(params))url.searchParams.set(k,v);
  const response=await request(url.toString(),{method:'GET',headers:{Authorization:'Bearer '+jwt,Accept:'application/json'},signal:AbortSignal.timeout(20000)});
  if(!response.ok)throw Object.assign(Error('Apple metadata API HTTP '+response.status),{httpStatus:response.status});
  const json=await response.json();if(!Array.isArray(json.data))throw Error('Unexpected Apple metadata response');
  return json;
 }
 const ids=await get('/v1/bundleIds',{'filter[identifier]':BUNDLE,'fields[bundleIds]':'identifier,name,platform','limit':'10'});
 const matching=ids.data.filter(x=>x.type==='bundleIds'&&x.attributes?.identifier===BUNDLE);
 const apps=await get('/v1/apps',{'filter[bundleId]':BUNDLE,'fields[apps]':'bundleId,name','limit':'10'});
 const appMatches=apps.data.filter(x=>x.type==='apps'&&x.attributes?.bundleId===BUNDLE);
 let profiles=null;
 if(matching.length===1){
  const json=await get('/v1/bundleIds/'+encodeURIComponent(matching[0].id)+'/profiles',{'fields[profiles]':'name,profileState,profileType,expirationDate','limit':'200'});
  profiles=json.data.filter(x=>x.type==='profiles'&&x.attributes?.profileType==='IOS_APP_STORE'&&x.attributes?.profileState==='ACTIVE'&&Date.parse(x.attributes?.expirationDate)>Date.now()).length;
 }
 return {status:'checked',bundleId:BUNDLE,bundleRegistrations:matching.length,appStoreConnectRecords:appMatches.length,activeAppStoreProfiles:profiles};
}
if(process.argv[1]===fileURLToPath(import.meta.url)){
 try{const result=await checkAppleStaging();console.log('APPLE STAGING PREFLIGHT '+JSON.stringify(result));if(result.status!=='checked')process.exitCode=1;}
 catch(error){console.error('APPLE STAGING PREFLIGHT failed: '+(Number.isInteger(error.httpStatus)?'HTTP '+error.httpStatus:'metadata check unavailable'));process.exitCode=1;}
}
