// Session-owned facade for original compat Firestore references. No raw Firestore
// references or snapshots are returned to callers. Server rules remain mandatory.
export function createRuntimeFirestoreScope({firestore,routerContext}){
 if(!firestore||!routerContext?.tenantId)throw Error('Verified runtime scope required');
 let closed=false;
 const nativeByFacade=new WeakMap(),facadeByNative=new WeakMap(),listeners=new Set();
 const guard=()=>{if(closed)throw Error('Runtime scope closed');};
 const prefix='tenants/'+routerContext.tenantId+'/';
 function unwrap(ref){guard();const native=nativeByFacade.get(ref);if(!native)throw Error('Foreign reference');return native;}
 function checkPath(path){if(typeof path!=='string'||!path.startsWith(prefix))throw Error('Foreign document');}
 function watch(native,args,convert){
  guard();const options=args[0]&&typeof args[0]==='object'&&!('next' in args[0])&&!('error' in args[0])&&!('complete' in args[0])?args.shift():undefined;
  const observer=typeof args[0]==='function'?{next:args[0],error:args[1],complete:args[2]}:args[0];
  if(!observer||typeof observer.next!=='function')throw Error('Listener callback required');
  let stopped=false,unsubscribe=null;
  const stop=()=>{if(stopped)return;stopped=true;listeners.delete(stop);if(unsubscribe)unsubscribe();};
  listeners.add(stop);
  const callbacks={next:value=>{if(!closed&&!stopped)observer.next(convert(value));},error:error=>{if(!closed&&!stopped)observer.error?.(error);stop();},complete:()=>{if(!closed&&!stopped)observer.complete?.();stop();}};
  try{unsubscribe=options===undefined?native.onSnapshot(callbacks):native.onSnapshot(options,callbacks);if(stopped)unsubscribe();}
  catch(e){stop();throw e;}
  return stop;
 }
 // The test iframe owns a distinct JS realm. Convert only plain record/array
 // literals, preserving SDK timestamps, sentinels, Date, and approved references.
 function normalize(value,seen=new WeakSet()){
  if(value===null||typeof value!=='object')return value;
  if(nativeByFacade.has(value))return unwrap(value);
  const proto=Object.getPrototypeOf(value);
  if(!Array.isArray(value)&&proto!==null&&Object.getPrototypeOf(proto)!==null)return value;
  if(seen.has(value))throw Error('Cyclic runtime data');
  seen.add(value);
  try{return Array.isArray(value)?Array.from(value,v=>normalize(v,seen)):Object.fromEntries(Object.entries(value).map(([k,v])=>[k,normalize(v,seen)]));}
  finally{seen.delete(value);}
 }
 const queryMethods=['where','orderBy','limit','limitToLast','startAt','startAfter','endAt','endBefore'];
 function wrapSnapshot(s){
  guard();
  if(Array.isArray(s.docs)){
   const docs=Object.freeze(s.docs.map(wrapSnapshot));
   return Object.freeze({docs,size:s.size,empty:s.empty,metadata:s.metadata,query:s.query?wrap(s.query):undefined,forEach(fn,ctx){guard();docs.forEach(x=>fn.call(ctx,x));},docChanges(...args){guard();return s.docChanges(...args).map(x=>({...x,doc:wrapSnapshot(x.doc)}));}});
  }
  checkPath(s.ref.path);
  const f=Object.freeze({id:s.id,exists:s.exists,ref:wrap(s.ref),metadata:s.metadata,data(...args){guard();return s.data(...args);},get(...args){guard();return s.get(...args);}});
  nativeByFacade.set(f,s);return f;
 }
 function wrap(native){
  guard();if(!native||typeof native!=='object')throw Error('Invalid SDK reference');
  if(facadeByNative.has(native))return facadeByNative.get(native);
  if(typeof native.path==='string')checkPath(native.path);
  const f={};
  if(typeof native.path==='string')f.path=native.path;
  if(typeof native.id==='string')f.id=native.id;
  const isDoc=typeof native.collection==='function';
  if(isDoc||typeof native.doc==='function')Object.defineProperty(f,'parent',{get(){guard();const parent=native.parent;return parent?.path?.startsWith(prefix)?wrap(parent):null;}});
  for(const method of queryMethods)if(typeof native[method]==='function')f[method]=(...args)=>{guard();return wrap(native[method](...args.map(a=>normalize(a))));};
  if(typeof native.doc==='function')f.doc=(...args)=>{guard();return wrap(native.doc(...args));};
  if(isDoc)f.collection=(name)=>{guard();return wrap(native.collection(name));};
  if(typeof native.get==='function')f.get=async(...args)=>{guard();const result=await native.get(...args);guard();return wrapSnapshot(result);};
  if(typeof native.onSnapshot==='function')f.onSnapshot=(...args)=>watch(native,args,wrapSnapshot);
  if(typeof native.add==='function')f.add=async(...args)=>{guard();const ref=await native.add(...args.map(a=>normalize(a)));guard();return wrap(ref);};
  for(const method of ['set','update'])if(isDoc&&typeof native[method]==='function')f[method]=async(...args)=>{guard();const result=await native[method](...args.map(a=>normalize(a)));guard();return result;};
  if(isDoc)f.delete=()=>{throw Error('Deletion not enabled');};
  if(typeof native.isEqual==='function')f.isEqual=other=>{guard();return native.isEqual(unwrap(other));};
  Object.freeze(f);nativeByFacade.set(f,native);facadeByNative.set(native,f);return f;
 }
 function batch(native,isTransaction){
  let failed=false;
  const ensure=()=>{guard();if(failed)throw Error('Operation closed');};
  function write(method,ref,...args){try{ensure();native[method](unwrap(ref),...args.map(a=>normalize(a)));return f;}catch(e){failed=true;throw e;}}
  const f=Object.freeze({set:(r,...a)=>write('set',r,...a),update:(r,...a)=>write('update',r,...a),delete(){failed=true;throw Error('Deletion not enabled');},...(isTransaction?{async get(r){ensure();const s=await native.get(unwrap(r));ensure();return wrapSnapshot(s);}}:{async commit(){ensure();const value=await native.commit();guard();return value;}})});
  return {f,ensure};
 }
 return Object.freeze({
  collection(path){guard();return wrap(firestore.collection(path));},
  doc(path){guard();return wrap(firestore.doc(path));},
  batch(){guard();return batch(firestore.batch(),false).f;},
  runTransaction(fn,...args){guard();return firestore.runTransaction(async tx=>{guard();const scoped=batch(tx,true);const value=await fn(scoped.f);scoped.ensure();return value;},...args);},
  close(){if(closed)return;closed=true;for(const stop of [...listeners]){try{stop();}catch{}}firestore.close();}
 });
}
