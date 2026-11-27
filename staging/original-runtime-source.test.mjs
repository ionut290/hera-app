import test from 'node:test';import assert from 'node:assert/strict';import {readFileSync} from 'node:fs';import {Script,createContext} from 'node:vm';
import {buildScopedOriginalAppCandidate,originalSdkInit,scopedSdkInit} from './original-runtime-source.mjs';
const source=readFileSync(new URL('../app.js',import.meta.url),'utf8');
const candidate=buildScopedOriginalAppCandidate(source);
test('complete original application compiles after replacing only SDK initialization',()=>{
 assert.equal(candidate.replace(scopedSdkInit,originalSdkInit),source);
 assert.doesNotThrow(()=>new Script(candidate));
 assert(!candidate.includes(originalSdkInit));
 assert(candidate.includes('enableFirestorePersistence();'));
});
function boot(prepared,configProject='varga-cantieri-v2-staging'){
 const end=candidate.indexOf('const PERSISTED_SESSION_KEY');
 assert(end>0);const prefix=candidate.slice(0,end);
 const firebase={initializeApp(){throw Error('Unscoped SDK forbidden');},firestore(){throw Error('Unscoped SDK forbidden');},auth(){throw Error('Unscoped SDK forbidden');},functions(){throw Error('Unscoped SDK forbidden');}};
 const window={firebase,firebaseConfig:{apiKey:'test-only',projectId:configProject,authDomain:'example.test',appId:'test'},__VARGA_ISOLATED_RUNTIME_CANDIDATE__:prepared};
 const context=createContext({window,firebase,console:{log(){},error(){}}});
 return new Script(prefix+'\n({firebaseApp,auth,db,functions,firebaseInitError})').runInContext(context);
}
const prepared=()=>{
 const projectId='varga-cantieri-v2-staging',app={options:{projectId}};
 return {projectId,auth:{app,currentUser:{uid:'u'}},session:{context:{uid:'u',tenantId:'demo-a'},firestore:{collection(){}}}};
};
test('original initialization accepts only the prepared staging company scope, not global Firestore',()=>{
 const p=prepared(),result=boot(p);
 assert.equal(result.db,p.session.firestore);assert.equal(result.auth,p.auth);assert.equal(result.functions,null);assert.equal(result.firebaseInitError,null);
 assert.equal(typeof result.db.enablePersistence,'undefined');
});
test('original initialization fails closed without verified session, wrong user or production config',()=>{
 for(const p of [null,{...prepared(),session:null},{...prepared(),auth:{app:{options:{projectId:'varga-cantieri-v2-staging'}},currentUser:{uid:'other'}}}]){const result=boot(p);assert.equal(result.db,null);assert(result.firebaseInitError);}
 assert.equal(boot(prepared(),'hera-app-6cd2b').db,null);
});
test('unexpected changes in original init fail the source-generation gate',()=>assert.throws(()=>buildScopedOriginalAppCandidate(source.replace(originalSdkInit,'  // altered initialization'))));
