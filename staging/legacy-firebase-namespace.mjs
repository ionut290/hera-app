// Candidate compatibility namespace for legacy scripts that call firebase.firestore()
// directly. This is not an authorization boundary: Firebase rules are authoritative.
// Keep native SDK references in the trusted entry controller, never in this namespace.
export function createLegacyFirebaseNamespace({prepared,statics={}}){
 const projectId='varga-cantieri-v2-staging';
 if(prepared?.projectId!==projectId||prepared?.auth?.app?.options?.projectId!==projectId||!prepared.session?.firestore||typeof prepared.isActive!=='function')throw Error('Verified staging scope required');
 const guard=()=>{if(!prepared.isActive()||prepared.auth.currentUser?.uid!==prepared.session.context.uid)throw Error('Company session closed');};
 guard();
 const deny=()=>{throw Error('Unapproved external service');};
 const app=Object.freeze({name:'[DEFAULT]',options:Object.freeze({projectId}),firestore:()=>firestore(),auth:()=>auth(),functions:deny,storage:deny,messaging:deny});
 const userView=user=>{
  if(!user||user.uid!==prepared.session.context.uid)return null;
  return Object.freeze({uid:user.uid,email:user.email||'',displayName:user.displayName||'',emailVerified:user.emailVerified===true,photoURL:user.photoURL||null,
   getIdTokenResult:async(...args)=>{guard();const value=await user.getIdTokenResult(...args);guard();return value;}});
 };
 const authFacade=Object.freeze({app,get currentUser(){return prepared.isActive()?userView(prepared.auth.currentUser):null;},
  onAuthStateChanged(next,error,complete){guard();const observer=typeof next==='function'?{next,error,complete}:next;
   if(typeof observer?.next!=='function')throw Error('Auth observer required');
   return prepared.auth.onAuthStateChanged(user=>observer.next(prepared.isActive()?userView(user):null),err=>observer.error?.(err),()=>observer.complete?.());},
  signOut(){guard();return prepared.auth.signOut();},
  setPersistence(){guard();return Promise.resolve();} // No change to persistence in a company runtime.
 });
 const auth=()=>{guard();return authFacade;};
 const firestore=()=>{guard();return prepared.session.firestore;};
 for(const key of ['FieldValue','FieldPath','Timestamp','GeoPoint'])if(statics[key]!==undefined)Object.defineProperty(firestore,key,{value:statics[key],enumerable:true});
 Object.freeze(firestore);Object.freeze(auth);
 return Object.freeze({SDK_VERSION:'staging-compat',apps:Object.freeze([app]),app(name){guard();if(name!==undefined&&name!=='[DEFAULT]')throw Error('Unapproved Firebase app');return app;},initializeApp:deny,auth,firestore,storage:deny,functions:deny,messaging:deny});
}
