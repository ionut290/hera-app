import test from 'node:test';
import assert from 'node:assert/strict';
import {createVerdeReadStore} from './verde-read-store.mjs';
const names=['verdeLevatoRecords','verdeLevatoCommesse','verdeLevatoConfig'];
function fake(projectId='varga-cantieri-v2-staging'){
  const ref=path=>({collection:n=>ref([...path,n]),doc:n=>ref([...path,n]),get:async()=>path.join('/')});
  return {...ref([]),app:{options:{projectId}}};
}
for(const tenantId of ['demo-a','demo-b']) for(const name of names){
  test(tenantId+' '+name,async()=>{
    const ref=createVerdeReadStore(fake(),{tenantId}).collection(name);
    assert.equal(await ref.get(),'tenants/'+tenantId+'/'+name);
    assert.equal(await ref.doc('sample').get(),'tenants/'+tenantId+'/'+name+'/sample');
    for(const method of ['set','add','update','delete']) {
      assert.equal(ref[method],undefined);
      assert.equal(ref.doc('sample')[method],undefined);
    }
  });
}
test('reject production',()=>assert.throws(()=>createVerdeReadStore(fake('hera-app-6cd2b'),{tenantId:'demo-a'})));
test('reject missing project',()=>assert.throws(()=>createVerdeReadStore({}, {tenantId:'demo-a'})));
test('reject missing context',()=>assert.throws(()=>createVerdeReadStore(fake(),null)));
test('reject invalid tenant',()=>assert.throws(()=>createVerdeReadStore(fake(),{tenantId:'../other'})));
test('reject unapproved collection',()=>assert.throws(()=>createVerdeReadStore(fake(),{tenantId:'demo-a'}).collection('users')));
test('reject invalid document',()=>assert.throws(()=>createVerdeReadStore(fake(),{tenantId:'demo-a'}).collection(names[0]).doc('a/b')));
