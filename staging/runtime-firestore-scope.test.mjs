import test from 'node:test';import assert from 'node:assert/strict';import {createRuntimeFirestoreScope} from './runtime-firestore-scope.mjs';
function fixture(id='demo-a'){
 let next,stops=0,closed=false,reads=0,writes=0;const path='tenants/'+id+'/commesse';
 const ref={id:'x',path:path+'/x',collection(name){return {...query,path:this.path+'/'+name};},get:async()=>{reads++;return snapshot;},set:async()=>{writes++;},update:async()=>{writes++;},onSnapshot(o){next=o.next;return()=>stops++;}};
 const snapshot={id:'x',exists:true,ref,data:()=>({nome:id}),get:()=>id};
 const query={path,id:'commesse',doc:()=>ref,where:()=>query,get:async()=>({docs:[snapshot],size:1,empty:false,query}),onSnapshot(o){next=o.next;return()=>stops++;}};
 const native={collection:()=>query,doc:()=>ref,batch:()=>({set:()=>writes++,update:()=>writes++,commit:async()=>{}}),close:()=>closed=true};
 const scope=createRuntimeFirestoreScope({firestore:native,routerContext:{tenantId:id}});
 return {scope,emit:()=>next({docs:[snapshot],size:1,empty:false,query}),stats:()=>({stops,closed,reads,writes})};
}
test('existing refs, queries, and snapshots cannot access data after close',async()=>{const f=fixture(),c=f.scope.collection('commesse'),r=c.doc('x'),s=await r.get(),q=c.where('a','==',1);assert.equal(s.data().nome,'demo-a');assert.equal((await q.get()).docs[0].id,'x');f.scope.close();assert.throws(()=>c.doc('x'));await assert.rejects(r.get());await assert.rejects(r.set({}));assert.throws(()=>s.data());await assert.rejects(q.get());assert.equal(f.stats().writes,0);});
test('listener automatically stops on close and late events are ignored',()=>{const f=fixture();let n=0;const stop=f.scope.collection('commesse').onSnapshot(()=>n++);f.emit();assert.equal(n,1);f.scope.close();f.emit();stop();assert.equal(n,1);assert.equal(f.stats().stops,1);assert(f.stats().closed);});
test('batches reject cross-session refs and cannot commit after rejection',async()=>{const a=fixture(),b=fixture('demo-b'),batch=a.scope.batch();assert.throws(()=>batch.set(b.scope.doc('commesse/x'),{}),/Foreign/);await assert.rejects(batch.commit());assert.equal(a.stats().writes,0);});
test('query snapshots preserve refs without leaking raw SDK',async()=>{const f=fixture(),q=await f.scope.collection('commesse').get();assert.equal(q.docs[0].ref.path,'tenants/demo-a/commesse/x');assert.equal(q.docs[0].ref.firestore,undefined);assert.equal(q.query.firestore,undefined);});
