import test from 'node:test';import assert from 'node:assert/strict';import {Script,createContext} from 'node:vm';import {extractOriginalPreview} from './generate-original-preview.mjs';import {projectPreviewWork,summarizePreviewWork} from './portal/src/work-price-projection.mjs';
const code=extractOriginalPreview().priceHandlers;
const core=new Script(code+'\ncore').runInContext(createContext({}));
const plant={id:'p1',denominazione:'Parco',codicePrezzo:'A11 | B21',done:false};
const prices=[{id:'a',codiceVoce:'A11',prezzoBase:12,unitaMisura:'mq',percentualeRibasso:null},{id:'b',codiceVoce:'B21',prezzoBase:25,unitaMisura:'AC',percentualeRibasso:.2}];
test('original calculations link quantities and flat-rate work independently, without changing plant state',()=>{
 const q=new Map([[JSON.stringify(['p1','A11']),100]]);
 const rows=projectPreviewWork([plant],prices,.1,q,core);
 assert.equal(rows.length,2);assert.equal(rows[0].totale,1080);assert.equal(rows[1].totale,20);
 assert.deepEqual(summarizePreviewWork(rows),{rows:2,verified:2,missing:0,estimated:1100});
 assert.equal(rows[0].stato,'DA FARE');assert.equal(plant.done,false);
});
test('missing and ambiguous prices do not count as zero or produce false totals',()=>{
 const q=new Map([[JSON.stringify(['p1','A11']),100]]);
 const result=projectPreviewWork([plant],[prices[0],{...prices[0],id:'dup'}],.1,q,core);
 assert.equal(result.length,2);assert(result.every(x=>x.totale===null));assert.equal(summarizePreviewWork(result).missing,2);
});
test('a different contract does not inherit quantity overrides or prices',()=>{
 const q=new Map([[JSON.stringify(['p1','A11']),100]]);
 assert.equal(projectPreviewWork([plant],prices,.1,q,core)[0].totale,1080);
 assert.equal(projectPreviewWork([plant],prices,.1,new Map(),core)[0].totale,null);
 assert.equal(projectPreviewWork([plant],[],.1,q,core)[0].totale,null);
});
