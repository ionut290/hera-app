import test from 'node:test';
import assert from 'node:assert/strict';
import {generateKeyPairSync} from 'node:crypto';
import {checkAppleStaging} from './apple-staging-preflight.mjs';
const {privateKey}=generateKeyPairSync('ec',{namedCurve:'prime256v1'});
const env={IOS_APPSTORE_KEY_ID:'test-key',IOS_APPSTORE_ISSUER_ID:'test-issuer',IOS_APPSTORE_PRIVATE_KEY_BASE64:Buffer.from(privateKey.export({type:'pkcs8',format:'pem'})).toString('base64')};
test('only GET metadata for the exact staging bundle; never profile content',async()=>{
 const calls=[];
 const data=[[{type:'bundleIds',id:'staging-id',attributes:{identifier:'it.vargacantieri.staging'}}],[{type:'apps',id:'staging-app',attributes:{bundleId:'it.vargacantieri.staging'}}],[{type:'profiles',attributes:{profileType:'IOS_APP_STORE',profileState:'ACTIVE',expirationDate:'2099-01-01'}},{type:'profiles',attributes:{profileType:'IOS_APP_STORE',profileState:'INVALID',expirationDate:'2099-01-01'}}]];
 const result=await checkAppleStaging({env,request:async(url,options)=>{assert.equal(options.method,'GET');const u=new URL(url);assert.equal(u.hostname,'api.appstoreconnect.apple.com');calls.push(u);return {ok:true,json:async()=>({data:data[calls.length-1]})};}});
 assert.equal(calls[0].searchParams.get('filter[identifier]'),'it.vargacantieri.staging');assert.equal(calls[1].searchParams.get('filter[bundleId]'),'it.vargacantieri.staging');assert.equal(calls[2].pathname,'/v1/bundleIds/staging-id/profiles');assert.ok(!calls[2].searchParams.get('fields[profiles]').includes('profileContent'));assert.equal(result.activeAppStoreProfiles,1);assert.equal(result.appStoreConnectRecords,1);
});
test('missing credentials make no requests',async()=>{assert.deepEqual(await checkAppleStaging({env:{},request:()=>{throw Error('unexpected request');}}),{status:'missing-credentials'});});
test('never queries profiles for a different bundle',async()=>{let n=0;const result=await checkAppleStaging({env,request:async()=>{n++;return {ok:true,json:async()=>({data:[{type:'bundleIds',id:'wrong',attributes:{identifier:'it.vargacantieri.hera'}}]})};}});assert.equal(n,2);assert.equal(result.bundleRegistrations,0);assert.equal(result.activeAppStoreProfiles,null);});
