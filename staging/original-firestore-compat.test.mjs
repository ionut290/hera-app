import test from 'node:test';
import assert from 'node:assert/strict';
import {createOriginalFirestoreCompat} from './original-firestore-compat.mjs';
const projectId='varga-cantieri-v2-staging';
function make(id='demo-a',project=projectId){
 let closed=false;const calls=[];
 const router={projectId,context:{tenantId:id},path(n){if(closed||!['commesse','neve_commesse','personale','mezzi'].includes(n))throw Error('Denied');return 'tenants/'+id+'/'+n;},close(){closed=true;}};
 const db={app:{options:{projectId:project}},collection:p=>(calls.push(p),{path:p}),doc:p=>(calls.push(p),{path:p})};
 return {db,router,calls};
}
test('routes original collection names without modifying helper functions',()=>{const f=make(),d=createOriginalFirestoreCompat(f);assert.equal(d.collection('commesse').path,'tenants/demo-a/commesse');assert.equal(d.collection('neve_commesse').path,'tenants/demo-a/neve_commesse');assert.equal(d.collection('personale').path,'tenants/demo-a/personale');});
test('same company scoped paths are not double-prefixed',()=>{const d=createOriginalFirestoreCompat(make());assert.equal(d.collection('tenants/demo-a/commesse').path,'tenants/demo-a/commesse');});
test('nested collection and document paths preserve original IDs',()=>{const d=createOriginalFirestoreCompat(make());assert.equal(d.collection('commesse/c1/impianti').path,'tenants/demo-a/commesse/c1/impianti');assert.equal(d.doc('commesse/c1/impianti/p1').path,'tenants/demo-a/commesse/c1/impianti/p1');});
test('rejects other company, sensitive collection, traversal and parity errors before SDK',()=>{const f=make(),d=createOriginalFirestoreCompat(f);for(const p of ['tenants/demo-b/commesse','appConfig','users','commesse/../impianti','commesse//impianti','/commesse','commesse/c1','commesse\\x'])assert.throws(()=>d.collection(p));assert.throws(()=>d.doc('commesse'));assert.equal(f.calls.length,0);});
test('different companies have different native reference paths',()=>{const a=createOriginalFirestoreCompat(make()),b=createOriginalFirestoreCompat(make('demo-b'));assert.notEqual(a.doc('commesse/c1').path,b.doc('commesse/c1').path);});
test('rejects production SDK and closed scope',()=>{assert.throws(()=>createOriginalFirestoreCompat(make('demo-a','hera-app-6cd2b')));const d=createOriginalFirestoreCompat(make());d.close();assert.throws(()=>d.collection('commesse'));});
test('rejects incorrect router path',()=>{const f=make();f.router.path=()=> 'tenants/demo-b/commesse';assert.throws(()=>createOriginalFirestoreCompat(f).collection('commesse'));});
