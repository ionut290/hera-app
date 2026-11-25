import {commessePath,commessaDocumentPath} from './commesse-paths.mjs';
import {planCreateCommessa,planCreatePlant} from './commesse-write-plan.mjs';
export function createCommesseStore(db,context,sdk){
 if(db?.app?.options?.projectId!=='varga-cantieri-v2-staging') throw new Error('Staging Firebase required');
 context=Object.freeze({...context});
 const root=commessePath(context);
 const {collection,doc,getDocsFromServer,getDocFromServer,runTransaction,serverTimestamp}=sdk;
 const readCollection=path=>getDocsFromServer(collection(db,...path));
 const readDocument=async path=>{
  const snap=await getDocFromServer(doc(db,...path));
  return Object.freeze({id:snap.id,exists:snap.exists(),data:()=>snap.data()});
 };
 async function createCommessa(commessaId,data){
  const plan=planCreateCommessa({context,commessaId,data});
  const ref=doc(db,...plan.path);
  await runTransaction(db,async tx=>{const snap=await tx.get(ref);if(snap.exists())throw new Error('Commessa already exists');tx.set(ref,{...plan.data,createdAt:serverTimestamp()});});
  return commessaId;
 }
 async function createPlant(commessaId,plantId,data){
  const plan=planCreatePlant({context,commessaId,plantId,data});
  const ref=doc(db,...plan.path);
  await runTransaction(db,async tx=>{const snap=await tx.get(ref);if(snap.exists())throw new Error('Cantiere already exists');tx.set(ref,{...plan.data,createdAt:serverTimestamp()});});
  return plantId;
 }
 return Object.freeze({createCommessa,createPlant,collection(name){
  if(name!=='commesse') throw new Error('Collection denied');
  return Object.freeze({get:()=>readCollection(root),doc(commessaId){
   if(commessaId===undefined) throw new Error('Commessa required');
   const path=commessaDocumentPath(context,commessaId);
   return Object.freeze({get:()=>readDocument(path),collection(name){
    if(name!=='impianti') throw new Error('Collection denied');
    return Object.freeze({get:()=>readCollection(commessePath(context,commessaId)),doc(impiantoId){
     if(impiantoId===undefined) throw new Error('Impianto required');
     const path=commessePath(context,commessaId,impiantoId);
     return Object.freeze({get:()=>readDocument(path)});
    }});
   }});
  }});
 }});
}
