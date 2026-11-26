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
