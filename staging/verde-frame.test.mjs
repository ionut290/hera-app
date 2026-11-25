import test from 'node:test';
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {Script} from 'node:vm';
import {buildVerdeFrame} from './portal/src/verde-frame.mjs';
const source=readFileSync(new URL('./verde-levato-pilot.mjs',import.meta.url),'utf8');
const html=buildVerdeFrame(source);
test('original module retained',()=>assert.ok(html.includes(source.replace('export function installVerdeLevatoPilot(','function installVerdeLevatoPilot('))));
test('inline script parses',()=>{const script=html.match(/<script>([\s\S]*)<\/script>/)[1];new Script(script);});
test('network denied by default',()=>{assert.match(html,/default-src 'none'/);assert.doesNotMatch(html,/connect-src/);});
test('empty archive',()=>assert.match(html,/const records = \[\];/));
test('reject unexpected source',()=>assert.throws(()=>buildVerdeFrame('')));
test('reject closing script injection',()=>assert.throws(()=>buildVerdeFrame(source+'</script>')));
test('close signal exists',()=>assert.match(html,/varga-verde-close/));

test('snapshot escapes script termination',()=>{const value='</script><script>alert(1)</script>';const output=buildVerdeFrame(source,[{id:'a',note:value}]);assert.ok(!output.includes(value));new Script(output.match(/<script>([\s\S]*)<\/script>/)[1]);});
test('invalid snapshot rejected',()=>assert.throws(()=>buildVerdeFrame(source,null)));
