// Synthetic Auth, Firestore and Storage; no live resource fallback.
import assert from 'node:assert/strict';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {getFirestore,FieldValue} from 'firebase-admin/firestore';
import {getStorage} from 'firebase-admin/storage';
import {createStorageQuotaService,MAX_FILE_BYTES} from './storage-quota-service.mjs';
const projectId='varga-cantieri-v2-staging',bucketName=projectId+'.appspot.com';
assert.match(process.env.FIRESTORE_EMULATOR_HOST||'',/^127\.0\.0\.1:8080$/);
assert.match(process.env.FIREBASE_AUTH_EMULATOR_HOST||'',/^127\.0\.0\.1:9099$/);
assert.match(process.env.FIREBASE_STORAGE_EMULATOR_HOST||'',/^127\.0\.0\.1:9199$/);
const app=initializeApp({projectId,storageBucket:bucketName});
const db=getFirestore(app),auth=getAuth(app),bucket=getStorage(app).bucket();
const service=createStorageQuotaService({app});
assert.throws(()=>createStorageQuotaService({app:{options:{projectId:'hera-app-6cd2b',storageBucket:'hera-app-6cd2b.appspot.com'}}}),e=>e.code==='STAGING_ONLY');
const password='Only-Local-Test-42',A='demo-storage-a',B='demo-storage-b';
const user=async uid=>{
 const email=uid+'@example.test';await auth.createUser({uid,email,password,emailVerified:true});
 const res=await fetch('http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=fake-api-key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true})});
 assert.equal(res.ok,true);return (await res.json()).idToken;
};
async function member(t,uid,role){
 await db.doc('tenants/'+t+'/members/'+uid).set({role,status:'active',demo:true});
 await db.doc('users/'+uid+'/memberships/'+t).set({tenantId:t,role,status:'active'});
}
const pdf=Buffer.from('%PDF-1.4\nSynthetic test document\n%%EOF\n');
const payload=pdf.toString('base64');
const make=(token,tenantId,fileId,base64=payload)=>({token,tenantId,fileId,name:'test.pdf',type:'application/pdf',base64});
const err=(code,fn)=>assert.rejects(fn,e=>e.code===code);
try{
 const [admin,worker,other]=await Promise.all(['storage-admin','storage-worker','storage-other'].map(user));
 for(const t of [A,B])await db.doc('tenants/'+t).set({demo:true,status:'active',plan:'trial',trialDays:30,trialStartedAt:new Date(),limits:{storageBytes:100000000},usage:{storageBytes:0}});
 await Promise.all([member(A,'storage-admin','admin'),member(A,'storage-worker','employee'),member(B,'storage-other','admin')]);
 const first=make(worker,A,'local-storage-file-0001');
 const file=await service.upload(first);
 assert.equal(file.status,'ready');assert.equal(file.size,pdf.length);
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,pdf.length);
 assert.deepEqual(await service.upload(first),file);
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,pdf.length);
 assert.equal((await service.download({token:admin,tenantId:A,fileId:first.fileId})).base64,payload);
 await err('FORBIDDEN',()=>service.download({token:other,tenantId:A,fileId:first.fileId}));
 await err('FORBIDDEN',()=>service.upload(make(other,A,'local-storage-file-0002')));
 await err('INVALID_FILE_TYPE',()=>service.upload(make(worker,A,'local-storage-file-0003',Buffer.from('not pdf').toString('base64'))));
 await err('FILE_ID_CONFLICT',()=>service.upload({...first,name:'other.pdf'}));
 // Explicit sharing: a colleague may read shared documents, but cannot edit or remove them.
 const peer=await user('storage-peer');await member(A,'storage-peer','employee');
 await err('FORBIDDEN',()=>service.download({token:peer,tenantId:A,fileId:first.fileId}));
 await service.share({token:worker,tenantId:A,fileId:first.fileId,visibility:'company'});
 assert.equal((await service.download({token:peer,tenantId:A,fileId:first.fileId})).base64,payload);
 const shared=(await service.list({token:peer,tenantId:A})).files.find(f=>f.id===first.fileId);
 assert.equal(shared.visibility,'company');assert.equal(shared.canManage,false);
 await err('FORBIDDEN',()=>service.share({token:peer,tenantId:A,fileId:first.fileId,visibility:'private'}));
 await err('FORBIDDEN',()=>service.remove({token:peer,tenantId:A,fileId:first.fileId}));
 await err('FORBIDDEN',()=>service.download({token:other,tenantId:A,fileId:first.fileId}));
 await service.share({token:admin,tenantId:A,fileId:first.fileId,visibility:'private'});
 await err('FORBIDDEN',()=>service.download({token:peer,tenantId:A,fileId:first.fileId}));
 assert.equal((await service.list({token:peer,tenantId:A})).files.some(f=>f.id===first.fileId),false);
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,pdf.length);
 // Plant photo associations are tenant-scoped and bound to actual documents.
 const plant={archive:'commesse',jobId:'photo-job',plantId:'plant-1'};
 for(const t of [A,B]){await db.doc('tenants/'+t+'/commesse/photo-job').set({nome:'Synthetic job'});await db.doc('tenants/'+t+'/commesse/photo-job/impianti/plant-1').set({nome:'Synthetic plant'});}
 const png='iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mP8/x8AAwMCAO+a0i0AAAAASUVORK5CYII=';
 const photo={token:worker,tenantId:A,fileId:'local-plant-photo-0001',name:'plant.png',type:'image/png',base64:png,plant,visibility:'company'};
 await err('INVALID_PLANT',()=>service.upload({...photo,fileId:'local-plant-photo-0002',plant:{...plant,plantId:'not-present'}}));
 await err('INVALID_FILE_TYPE',()=>service.upload({...make(worker,A,'local-plant-photo-0003'),plant}));
 const saved=await service.upload(photo);assert.deepEqual(saved.plant,plant);
 assert.equal((await service.list({token:peer,tenantId:A,plant})).files.length,1);
 assert.equal((await service.list({token:other,tenantId:B,plant})).files.length,0);
 assert.equal((await service.download({token:peer,tenantId:A,fileId:photo.fileId})).base64,png);
 await err('FORBIDDEN',()=>service.download({token:other,tenantId:B,fileId:photo.fileId}));
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,pdf.length+Buffer.from(png,'base64').length);
 await service.remove({token:worker,tenantId:A,fileId:photo.fileId});
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,pdf.length);
 console.log('STORAGE SHARING AND PLANT PASS: private default, company access, revocation, no cross-company access or peer deletion, verified plant link and shared quota.');
 // Counter fixture represents previous reservations; exact byte boundary and concurrent requests.
 await db.doc('tenants/'+A).update({'usage.storageBytes':100000000-pdf.length});
 const pair=await Promise.allSettled([service.upload(make(worker,A,'local-storage-file-0004')),service.upload(make(admin,A,'local-storage-file-0005'))]);
 assert.equal(pair.filter(x=>x.status==='fulfilled').length,1);
 assert.equal(pair.filter(x=>x.status==='rejected'&&x.reason.code==='STORAGE_LIMIT').length,1);
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,100000000);
 assert.equal((await db.doc('tenants/'+B).get()).data().usage.storageBytes,0);
 // Listing respects ownership, and deletion releases exactly one reservation.
 const listing=await service.list({token:worker,tenantId:A});
 assert.ok(listing.files.some(f=>f.id===first.fileId));assert.equal(listing.used,100000000);
 assert.ok((await service.list({token:other,tenantId:B})).files.length===0);
 await err('FORBIDDEN',()=>service.remove({token:other,tenantId:A,fileId:first.fileId}));
 const deletion=make(worker,A,'local-storage-delete-0001');
 const winning=pair.find(x=>x.status==='fulfilled').value.id;
 await Promise.all([service.remove({token:admin,tenantId:A,fileId:winning}),service.remove({token:admin,tenantId:A,fileId:winning})]);
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,100000000-pdf.length);
 await service.upload(deletion);
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,100000000);
 await service.remove({token:worker,tenantId:A,fileId:deletion.fileId});
 await service.remove({token:worker,tenantId:A,fileId:deletion.fileId});
 assert.equal((await db.doc('tenants/'+A).get()).data().usage.storageBytes,100000000-pdf.length);
 await err('FORBIDDEN',()=>service.download({token:worker,tenantId:A,fileId:deletion.fileId}));
 assert.equal((await service.list({token:worker,tenantId:A})).files.some(f=>f.id===deletion.fileId),false);
 console.log('STORAGE DELETE PASS: concurrent deletion, once-only release, new upload after release, hidden tombstone.');
 await db.doc('tenants/'+A).update({trialStartedAt:new Date(Date.now()-31*86400000)});
 await err('PLAN_EXPIRED',()=>service.upload(make(worker,A,'local-storage-file-0006')));
 await err('PLAN_EXPIRED',()=>service.remove({token:worker,tenantId:A,fileId:first.fileId}));
 assert.equal((await service.download({token:worker,tenantId:A,fileId:first.fileId})).base64,payload);
 await db.doc('tenants/'+A+'/members/storage-worker').update({status:'suspended'});
 await err('FORBIDDEN',()=>service.download({token:worker,tenantId:A,fileId:first.fileId}));
 // Direct client requests must fail, even with a valid ID token.
 const endpoint='http://127.0.0.1:9199/v0/b/'+bucketName+'/o';
 const object=encodeURIComponent('tenants/'+A+'/private-files/'+first.fileId);
 const res=await fetch(endpoint+'/'+object+'?alt=media',{headers:{Authorization:'Firebase '+admin}});
 assert.ok([401,403].includes(res.status),'Direct bucket read denied');
 const upload=await fetch(endpoint+'?name='+encodeURIComponent('tenants/'+A+'/private-files/bypass-test'),{method:'POST',headers:{Authorization:'Firebase '+admin,'Content-Type':'application/pdf'},body:pdf});
 assert.ok([401,403].includes(upload.status),'Direct bucket write denied');
 console.log('STORAGE QUOTA PASS: real bucket upload and download, duplicate reservation, concurrent exact 100 MB cap, cross-tenant access denied, expiry read-only, direct client access denied.');
}finally{await deleteApp(app);}
