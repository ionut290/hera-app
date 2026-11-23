/* Varga Piano Neve 2.0: durable FIFO work queue.
 * IndexedDB is authoritative; never drop an item until the server acknowledges it.
 * All writes for a record are serialized through the same promise chain.
 */
(function (root) {
  'use strict';
  const DB_NAME = 'varga-piano-neve-operations';
  const STORE = 'outbox', SESSIONS='sessions';
  const DB_VERSION = 1;
  let dbPromise, chain = Promise.resolve(), flushing = null, activeUid = null;
  function requireUid(){if(!activeUid)throw new Error('Accesso richiesto per i dati offline');return activeUid;}
  function setUser(uid){activeUid=uid?String(uid):null;}
  function scopedKey(id){return requireUid()+'::'+id;}
  function database() {
    if (!root.indexedDB) return Promise.reject(new Error('IndexedDB non disponibile'));
    if (!dbPromise) dbPromise = new Promise((resolve, reject) => {
      const req = root.indexedDB.open(DB_NAME, DB_VERSION);
      req.onupgradeneeded = () => { const db = req.result; if (!db.objectStoreNames.contains(STORE)) db.createObjectStore(STORE, {keyPath:'id'}); if(!db.objectStoreNames.contains(SESSIONS))db.createObjectStore(SESSIONS,{keyPath:'id'}); };
      req.onsuccess = () => resolve(req.result);
      req.onerror = () => reject(req.error || new Error('Errore IndexedDB'));
    }).catch(error => { dbPromise = null; throw error; });
    return dbPromise;
  }
  async function operation(mode, task) {
    const db = await database();
    return new Promise((resolve, reject) => {
      const tx = db.transaction(STORE, mode), store = tx.objectStore(STORE);
      let output;
      try { output = task(store); } catch (e) { tx.abort(); reject(e); return; }
      tx.oncomplete = () => resolve(output);
      tx.onerror = () => reject(tx.error || new Error('Transazione IndexedDB fallita'));
      tx.onabort = () => reject(tx.error || new Error('Transazione IndexedDB annullata'));
    });
  }
  function enqueue(id, payload) {
    if (!id || !payload) return Promise.reject(new Error('Record non valido'));
    const key=scopedKey(id),owner=activeUid;
    const snapshot = JSON.parse(JSON.stringify(payload));
    const work = chain.catch(() => {}).then(async () => {
      const db = await database();
      await new Promise((resolve,reject) => {
        const tx = db.transaction(STORE,'readwrite');
        tx.objectStore(STORE).put({id:key,recordId:id,ownerUid:owner,payload:snapshot,queuedAt:Date.now()});
        tx.oncomplete = resolve;
        tx.onerror = () => reject(tx.error);
        tx.onabort = () => reject(tx.error);
      });
    });
    chain = work;
    return work;
  }
  async function entries() {
    await chain;
    const owner=requireUid(), db = await database();
    return new Promise((resolve,reject) => {
      const tx = db.transaction(STORE,'readonly');
      const req = tx.objectStore(STORE).getAll();
      req.onsuccess = () => resolve(req.result.filter(row=>row.ownerUid===owner).sort((a,b)=>a.queuedAt-b.queuedAt));
      req.onerror = () => reject(req.error);
    });
  }
  async function removeIfUnchanged(id, snapshot) {
    await chain;
    const owner=requireUid(),db = await database();
    return new Promise((resolve,reject) => {
      const tx = db.transaction(STORE,'readwrite'), store=tx.objectStore(STORE);
      const req=store.get(owner+'::'+id);
      req.onsuccess=()=> {
        const row=req.result;
        if(row && row.ownerUid===owner && JSON.stringify(row.payload) === JSON.stringify(snapshot)) store.delete(owner+'::'+id);
      };
      tx.oncomplete=resolve;
      tx.onerror=()=>reject(tx.error);
    });
  }
  function flush(sender) {
    if (flushing) return flushing;
    const owner=requireUid();
    flushing = (async()=> {
      const list=await entries();
      for (const entry of list) {
        // Never timeout a write and immediately retry it: the remote write may still be in flight.
        if(activeUid!==owner)throw new Error('Account cambiato durante la sincronizzazione');
        await sender(entry.recordId,entry.payload);
        await removeIfUnchanged(entry.recordId,entry.payload);
      }
      return (await entries()).length;
    })().finally(()=>{flushing=null;});
    return flushing;
  }
  async function count() { return (await entries()).length; }
  function saveSession(payload) {
    const key=scopedKey('active');
    const snapshot=JSON.parse(JSON.stringify(payload));
    return new Promise(async (resolve,reject)=>{
      try{
       const db=await database(), tx=db.transaction(SESSIONS,'readwrite');
       tx.objectStore(SESSIONS).put({id:key,payload:snapshot});
       tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);
      }catch(e){reject(e)}
    });
  }
  async function loadSession(){
    const key=scopedKey('active'),db=await database();
    return new Promise((resolve,reject)=>{
      const req=db.transaction(SESSIONS,'readonly').objectStore(SESSIONS).get(key);
      req.onsuccess=()=>resolve(req.result?.payload||null);req.onerror=()=>reject(req.error);
    });
  }
  async function clearSession(){
    const key=scopedKey('active'),db=await database();
    return new Promise((resolve,reject)=>{
      const tx=db.transaction(SESSIONS,'readwrite');
      tx.objectStore(SESSIONS).delete(key);tx.oncomplete=resolve;tx.onerror=()=>reject(tx.error);
    });
  }
  const api={enqueue,flush,count,ready:database,saveSession,loadSession,clearSession,setUser};
  root.VargaSnowOutbox=api;
  if(typeof module!=='undefined'&&module.exports) module.exports=api;
})(typeof window!=='undefined'?window:globalThis);
