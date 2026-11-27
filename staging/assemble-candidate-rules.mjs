// Generate only the candidate emulator rules. Never deploy or write root firestore.rules.
import {readFileSync,writeFileSync} from 'node:fs';
import {fileURLToPath} from 'node:url';
const read=name=>readFileSync(new URL(name,import.meta.url),'utf8');
export const fragments=Object.freeze(['verde-levato-readonly','commesse-readonly','commesse-write','original-create','original-manual-plant','original-price-list','original-archives-readonly','fatto-policy','cobo-write']);
export function assembleRules(){
 let rules=read('firestore-emulator-base.rules');
 for(const [marker,code] of [['// SUPER_ADMIN_FRAGMENT',read('superadmin.rules.fragment')],['// VERDE_FRAGMENT',fragments.map(f=>read(f+'.rules.fragment')).join('\n')]]){
  if(rules.split(marker).length!==2)throw Error('Rules marker mismatch: '+marker);
  rules=rules.replace(marker,()=>code);
 }
 return rules;
}
if(process.argv[1]===fileURLToPath(import.meta.url)){
 writeFileSync(new URL('candidate.rules',import.meta.url),assembleRules());
 console.log('STAGING RULES: candidate.rules assembled; no deployment performed.');
}
