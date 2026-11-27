import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {assembleRules,fragments} from './assemble-candidate-rules.mjs';
test('all rule fragments are inserted verbatim, including regex dollar anchors',()=>{
 const rules=assembleRules();
 for(const f of ['superadmin',...fragments]){
  const source=readFileSync(new URL(f+'.rules.fragment',import.meta.url),'utf8');
  assert.ok(rules.includes(source),f+' must appear verbatim');
 }
 assert.equal((rules.match(/match \/\{document=\*\*\}/g)||[]).length,1);
 assert.ok(!rules.includes('// SUPER_ADMIN_FRAGMENT'));
 assert.ok(!rules.includes('// VERDE_FRAGMENT'));
});
