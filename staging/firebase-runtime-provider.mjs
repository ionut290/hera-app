// The launcher depends on a connection provider, not on a hardcoded Firebase initializer.
// For now only the already-tested local emulator is supported. No live or production fallback.
export const stagingProjectId='varga-cantieri-v2-staging';
export async function connectLocalFirebase({firebase,location}){
 if(!['localhost','127.0.0.1'].includes(location?.hostname)||!['http:','https:','capacitor:'].includes(location?.protocol))throw Error('Local emulator only');
 const app=firebase.initializeApp({apiKey:'demo-local-emulator-key',authDomain:'localhost',projectId:stagingProjectId,appId:'demo-local-emulator'});
 const auth=app.auth(),db=app.firestore();
 auth.useEmulator('http://127.0.0.1:9099',{disableWarnings:true});
 db.useEmulator('127.0.0.1',8080);
 await auth.setPersistence(firebase.auth.Auth.Persistence.NONE);
 const statics=Object.freeze({FieldValue:firebase.firestore.FieldValue,FieldPath:firebase.firestore.FieldPath,Timestamp:firebase.firestore.Timestamp,GeoPoint:firebase.firestore.GeoPoint});
 return Object.freeze({auth,db,statics});
}

// Explicit staging-only client connection. Firebase web API keys identify projects;
// they are not server credentials. Server-side security rules remain authoritative.
export function makeStagingWebConfig(apiKey){
 if(typeof apiKey!=='string'||!/^AIza[0-9A-Za-z_-]{35}$/.test(apiKey))throw Error('Verified staging web API key required');
 return Object.freeze({apiKey,authDomain:'varga-cantieri-v2-staging.firebaseapp.com',projectId:stagingProjectId,appId:'1:394791335182:web:6c31c93ac34e64c881935f',messagingSenderId:'394791335182'});
}
export async function connectStagingFirebase({firebase,location,apiKey}){
 if(!['localhost','127.0.0.1'].includes(location?.hostname)||!['https:','capacitor:'].includes(location?.protocol))throw Error('Private staging native candidate only');
 const config=makeStagingWebConfig(apiKey);
 const app=firebase.initializeApp(config);
 const auth=app.auth(),db=app.firestore();
 await auth.setPersistence(firebase.auth.Auth.Persistence.NONE);
 const statics=Object.freeze({FieldValue:firebase.firestore.FieldValue,FieldPath:firebase.firestore.FieldPath,Timestamp:firebase.firestore.Timestamp,GeoPoint:firebase.firestore.GeoPoint});
 return Object.freeze({auth,db,statics});
}
