// Native OAuth tokens never grant tenant permissions. Firebase JS Auth remains authoritative.
export function nativeProviderBridge(capacitor,enabled){
 if(!enabled||capacitor?.isNativePlatform?.()!==true||capacitor?.getPlatform?.()!=='ios'||capacitor?.isPluginAvailable?.('FirebaseAuthentication')!==true)return null;
 const plugin=capacitor.Plugins?.FirebaseAuthentication;
 return plugin&&typeof plugin.signInWithGoogle==='function'&&typeof plugin.signInWithApple==='function'?plugin:null;
}
export async function getProviderCredential({provider,bridge,firebase}){
 if(!bridge)throw Error('PROVIDER_UNAVAILABLE');
 if(provider!=='google'&&provider!=='apple')throw Error('PROVIDER_INVALID');
 const result=await (provider==='google'?bridge.signInWithGoogle({skipNativeAuth:true}):bridge.signInWithApple({skipNativeAuth:true}));
 const token=result?.credential?.idToken;
 if(typeof token!=='string'||!token)throw Error('PROVIDER_TOKEN_MISSING');
 if(provider==='google')return firebase.auth.GoogleAuthProvider.credential(token);
 const nonce=result?.credential?.nonce;
 if(typeof nonce!=='string'||!nonce)throw Error('PROVIDER_NONCE_MISSING');
 return new firebase.auth.OAuthProvider('apple.com').credential({idToken:token,rawNonce:nonce});
}
export function socialIntent(mode,value){
 if(mode==='login')return null;
 const clean=String(value||'').trim();
 if(mode==='company'){if(clean.length<2||clean.length>120)throw Error('COMPANY_NAME_INVALID');return {type:'company',name:clean};}
 if(mode==='join'){const code=clean.toUpperCase().replace(/\s/g,'');if(!/^VC-[A-Z0-9]{8}$/.test(code))throw Error('CODE_INVALID');return {type:'join',code};}
 throw Error('PROVIDER_INVALID');
}
