const test=require('node:test');
const assert=require('node:assert/strict');
const {generateKeyPairSync}=require('node:crypto');
const {fetchActiveProfile,EXPECTED_NAME,EXPECTED_BUNDLE}=require('./fetch-v2-appstore-profile.cjs');
const {privateKey}=generateKeyPairSync('ec',{namedCurve:'prime256v1'});
const env={IOS_APPSTORE_KEY_ID:'test',IOS_APPSTORE_ISSUER_ID:'test',IOS_APPSTORE_PRIVATE_KEY_BASE64:Buffer.from(privateKey.export({format:'pem',type:'pkcs8'})).toString('base64')};
const good=Buffer.alloc(128,42);
function mock({bundle=EXPECTED_BUNDLE,appBundle=EXPECTED_BUNDLE,profileName=EXPECTED_NAME,expiry='2099-01-01',duplicate=false}={}){
 let count=0;
 return async(url,opt)=>{
  assert.equal(opt.method,'GET');
  const u=new URL(url);assert.equal(u.origin,'https://api.appstoreconnect.apple.com');
  if(++count===1){assert.equal(u.pathname,'/v1/apps/6839540105');return {ok:true,json:async()=>({data:{type:'apps',id:'6839540105',attributes:{bundleId:appBundle}}})};}
  assert.equal(u.pathname,'/v1/profiles');assert.equal(u.searchParams.get('filter[name]'),EXPECTED_NAME);
  const p={type:'profiles',attributes:{name:profileName,profileState:'ACTIVE',profileType:'IOS_APP_STORE',expirationDate:expiry,profileContent:good.toString('base64')},relationships:{bundleId:{data:{id:'b'}}}};
  return {ok:true,json:async()=>({data:duplicate?[p,p]:[p],included:[{type:'bundleIds',id:'b',attributes:{identifier:bundle}}]})};
 };
}
test('exact V2 app and active profile',async()=>assert.deepEqual(await fetchActiveProfile(env,mock()),good));
test('reject other app',async()=>assert.rejects(fetchActiveProfile(env,mock({appBundle:'it.vargacantieri.hera'})),/mismatch/));
test('reject other profile bundle',async()=>assert.rejects(fetchActiveProfile(env,mock({bundle:'it.vargacantieri.hera'})),/found 0/));
test('reject other profile name',async()=>assert.rejects(fetchActiveProfile(env,mock({profileName:'VARGA CANTIERI - TestFlight'})),/found 0/));
test('reject expired profile',async()=>assert.rejects(fetchActiveProfile(env,mock({expiry:'2000-01-01'})),/found 0/));
test('reject ambiguous profile',async()=>assert.rejects(fetchActiveProfile(env,mock({duplicate:true})),/found 2/));
