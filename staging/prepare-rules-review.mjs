// Make an auditable, demo-only release review bundle. No Firebase CLI or network.
import {mkdirSync,readFileSync,writeFileSync} from 'node:fs';
import {createHash} from 'node:crypto';
import {resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {assembleRules,fragments} from './assemble-candidate-rules.mjs';
const dir=fileURLToPath(new URL('./rules-review-dist/',import.meta.url));
const sha=process.env.GITHUB_SHA||'local-review';
if(sha!=='local-review'&&!/^[a-f0-9]{40}$/.test(sha))throw Error('Invalid source revision');
const rules=assembleRules();
const digest=value=>createHash('sha256').update(value).digest('hex');
// Freeze the exact tested candidate. It is never a deployment command or config.
mkdirSync(dir,{recursive:true});
writeFileSync(resolve(dir,'candidate.rules'),rules);
const sources=['firestore-emulator-base.rules','superadmin.rules.fragment',...fragments.map(f=>f+'.rules.fragment')];
writeFileSync(resolve(dir,'manifest.json'),JSON.stringify({projectId:'varga-cantieri-v2-staging',sourceRevision:sha,sha256:digest(rules),mode:'review-only',approvedForDeploy:false,requiresActiveRulesComparison:true,requiresDraftPreservation:true,protectedCompletionEnabled:false,sources:sources.map(name=>({name,sha256:digest(readFileSync(new URL(name,import.meta.url)))}))},null,2)+'\n');
writeFileSync(resolve(dir,'README.txt'),'STAGING REVIEW ONLY. Exact candidate tested in CI. Do not publish to any Firebase project until the active rules and unpublished draft are recovered, compared and reviewed. Never use root firestore.rules. Demo-only operational scopes; protected FATTO/WhatsApp execution not authorized. No production changes or data migrations.\n');
console.log('Rules review bundle: SHA256 '+digest(rules)+'; project varga-cantieri-v2-staging; NOT DEPLOYED.');
