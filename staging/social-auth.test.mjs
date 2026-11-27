import {test} from 'node:test';
import assert from 'node:assert/strict';
import {nativeProviderBridge,getProviderCredential,socialIntent} from './portal/social-auth.mjs';
test('disabled or non-native builds never expose social providers',()=>{
 const plugin={signInWithGoogle(){},signInWithApple(){}};
 const cap={isNativePlatform:()=>true,getPlatform:()=>'ios',isPluginAvailable:()=>true,Plugins:{FirebaseAuthentication:plugin}};
 assert.equal(nativeProviderBridge(cap,false),null);
 assert.equal(nativeProviderBridge({...cap,getPlatform:()=>'web'},true),null);
 assert.equal(nativeProviderBridge({...cap,isPluginAvailable:()=>false},true),null);
 assert.equal(nativeProviderBridge(cap,true),plugin);
});
test('Google uses native token, Apple requires raw nonce, neither logs tokens',async()=>{
 const bridge={signInWithGoogle:async opts=>{assert.equal(opts.skipNativeAuth,true);return {credential:{idToken:'fake-google'}}},signInWithApple:async opts=>{assert.equal(opts.skipNativeAuth,true);return {credential:{idToken:'fake-apple',nonce:'fake-nonce'}}}};
 const firebase={auth:{GoogleAuthProvider:{credential:idToken=>({provider:'google',idToken})},OAuthProvider:class{constructor(provider){this.provider=provider}credential(input){return {provider:this.provider,...input}}}}};
 assert.deepEqual(await getProviderCredential({provider:'google',bridge,firebase}),{provider:'google',idToken:'fake-google'});
 assert.deepEqual(await getProviderCredential({provider:'apple',bridge,firebase}),{provider:'apple.com',idToken:'fake-apple',rawNonce:'fake-nonce'});
 await assert.rejects(getProviderCredential({provider:'apple',bridge:{signInWithApple:async()=>({credential:{idToken:'fake'}})},firebase}),/PROVIDER_NONCE_MISSING/);
 await assert.rejects(getProviderCredential({provider:'google',bridge:null,firebase}),/PROVIDER_UNAVAILABLE/);
});
test('company and join intent cannot assign roles; login has no creation intent',()=>{
 assert.deepEqual(socialIntent('company',' Impresa '),{type:'company',name:'Impresa'});
 assert.deepEqual(socialIntent('join','vc-ab12cd34'),{type:'join',code:'VC-AB12CD34'});
 assert.equal(socialIntent('login'),null);
 assert.throws(()=>socialIntent('join','admin'),/CODE_INVALID/);
 assert.throws(()=>socialIntent('company',' '),/COMPANY_NAME_INVALID/);
 assert.throws(()=>socialIntent('admin','Impresa'),/PROVIDER_INVALID/);
});
