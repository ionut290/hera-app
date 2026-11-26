// Read-only, staging-only Firebase Rules API audit. No operational records,
// service account keys, production access, or published rules text in logs.
import {createHash} from 'node:crypto';
import {fileURLToPath} from 'node:url';
import {assembleRules} from './assemble-candidate-rules.mjs';
const PROJECT='varga-cantieri-v2-staging';
const PROVIDER='projects/394791335182/locations/global/workloadIdentityPools/varga-v2-audit/providers/github';
const API='https://firebaserules.googleapis.com/v1/';
const RELEASE='projects/'+PROJECT+'/releases/cloud.firestore';
const sha=s=>createHash('sha256').update(s).digest('hex');
async function json(request,url,options,label){
 const response=await request(url,{...options,redirect:'error',signal:AbortSignal.timeout(20000)});
 if(!response.ok)throw Error(label+' HTTP '+response.status);
 return response.json();
}
export async function getActiveRules({token,request=fetch}){
 if(typeof token!=='string'||!token)throw Error('Read-only access token required');
 const get=path=>json(request,API+path,{method:'GET',headers:{Authorization:'Bearer '+token,Accept:'application/json'}},'Firebase Rules API');
 const release=await get(RELEASE);
 if(release.name!==RELEASE||typeof release.rulesetName!=='string'||!/^projects\/varga-cantieri-v2-staging\/rulesets\/[A-Za-z0-9_-]+$/.test(release.rulesetName))throw Error('Unexpected release identity');
 const ruleset=await get(release.rulesetName);
 if(ruleset.name!==release.rulesetName||!Array.isArray(ruleset.source?.files)||ruleset.source.files.length!==1||typeof ruleset.source.files[0].content!=='string')throw Error('Unexpected ruleset source');
 const content=ruleset.source.files[0].content;
 if(!content.trim())throw Error('Empty active rules');
 return Object.freeze({rulesetName:release.rulesetName,updateTime:release.updateTime||null,sha256:sha(content),matchesTestedCandidate:content===assembleRules(),lineCount:content.split('\n').length});
}
async function getFederatedToken(env,request=fetch){
 if(env.GITHUB_REPOSITORY!=='ionut290/hera-app'||env.GITHUB_REF!=='refs/heads/varga-cantieri-v2-staging'||env.GITHUB_WORKFLOW_REF!=='ionut290/hera-app/.github/workflows/staging-firebase-rules-audit.yml@refs/heads/varga-cantieri-v2-staging')throw Error('Unexpected audit workflow identity');
 if(env.VARGA_V2_RULES_AUDIT_READY!=='true')throw Error('Read-only federation not authorized');
 if(!env.ACTIONS_ID_TOKEN_REQUEST_URL||!env.ACTIONS_ID_TOKEN_REQUEST_TOKEN)throw Error('GitHub OIDC unavailable');
 const endpoint=new URL(env.ACTIONS_ID_TOKEN_REQUEST_URL);
 if(endpoint.protocol!=='https:'||!endpoint.hostname.endsWith('.actions.githubusercontent.com'))throw Error('Unexpected OIDC endpoint');
 endpoint.searchParams.set('audience','https://iam.googleapis.com/'+PROVIDER);
 const oidc=await json(request,endpoint.toString(),{method:'GET',headers:{Authorization:'Bearer '+env.ACTIONS_ID_TOKEN_REQUEST_TOKEN}},'GitHub OIDC');
 if(typeof oidc.value!=='string'||!oidc.value)throw Error('Invalid OIDC response');
 const body=new URLSearchParams({audience:'//iam.googleapis.com/'+PROVIDER,grant_type:'urn:ietf:params:oauth:grant-type:token-exchange',requested_token_type:'urn:ietf:params:oauth:token-type:access_token',scope:'https://www.googleapis.com/auth/cloud-platform',subject_token_type:'urn:ietf:params:oauth:token-type:jwt',subject_token:oidc.value});
 const sts=await json(request,'https://sts.googleapis.com/v1/token',{method:'POST',headers:{'Content-Type':'application/x-www-form-urlencoded'},body},'Google STS');
 if(typeof sts.access_token!=='string'||!sts.access_token)throw Error('Invalid STS response');
 return sts.access_token;
}
if(process.argv[1]===fileURLToPath(import.meta.url)){
 try{
  const token=await getFederatedToken(process.env);
  const result=await getActiveRules({token});
  console.log('ACTIVE STAGING RULES AUDIT '+JSON.stringify(result));
  if(!result.matchesTestedCandidate)console.log('Published rules differ from the tested candidate; deployment remains blocked.');
 }catch(e){console.error('Staging rules audit failed; no change made. '+(/^(?:Firebase Rules API|Google STS|GitHub OIDC) HTTP \d{3}$/.test(e.message)?e.message:'Verification unavailable.'));process.exitCode=1;}
}
