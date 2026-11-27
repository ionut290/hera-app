import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {createOriginalCollectionBridge} from './original-collection-bridge.mjs';
const source=readFileSync(new URL('../app.js',import.meta.url),'utf8');
const original=[...source.matchAll(/function (get\w+CollectionName)\(\)\s*\{ return isSnowServiceContext\(\) \? "([^"]+)" : "([^"]+)"; \}/g)];
function fixture(id){let closed=false;return createOriginalCollectionBridge({path(n){if(closed)throw Error('Closed');return 'tenants/'+id+'/'+n;},close(){closed=true;}});}
test('all seven exact source selectors match both original archive names',()=>{assert.equal(original.length,7);const b=fixture('demo-a');for(const [,name,snow,normal] of original){assert.equal(b[name](false),'tenants/demo-a/'+normal);assert.equal(b[name](true),'tenants/demo-a/'+snow);}});
test('same original selectors isolate companies',()=>{const a=fixture('demo-a'),b=fixture('demo-b');for(const [,name] of original)for(const snow of [false,true])assert.notEqual(a[name](snow),b[name](snow));});
test('closed bridge rejects every selector',()=>{const b=fixture('demo-a');b.close();for(const [,name] of original)assert.throws(()=>b[name]());});
test('reject invalid service mode and missing router',()=>{assert.throws(()=>createOriginalCollectionBridge(null));assert.throws(()=>fixture('demo-a').getCommesseCollectionName('neve'));});
