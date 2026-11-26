// Inventory the actual original app, including indirectly loaded local scripts.
// Diagnostic only: regex candidates require review; no credential values or source snippets in output.
import {readFileSync,existsSync,statSync,readdirSync} from 'node:fs';
import {resolve,relative,dirname,extname} from 'node:path';
import {fileURLToPath} from 'node:url';
const root=fileURLToPath(new URL('../',import.meta.url));
const read=p=>readFileSync(resolve(root,p),'utf8');
export function auditOriginalRuntime(){
 const html=read('index.html'),config=read('firebase-config.js');
 const files=new Set(),missing=new Set(),external=new Set();
 function candidate(value,origin){
  const v=value.split(/[?#]/)[0];if(!v)return;
  if(/^(https?:)?\/\//i.test(v)){try{external.add(new URL(v.startsWith('//')?'https:'+v:v).hostname);}catch{}return;}
  if(/^[a-z]+:/i.test(v))return;
  const path=resolve(root,dirname(origin),v.replace(/^\//,''));
  const local=relative(root,path);
  if(local.startsWith('..')||!['.js','.css','.html'].includes(extname(local)))return;
  if(!existsSync(path)||!statSync(path).isFile())missing.add(local);else files.add(local);
 }
 for(const m of html.matchAll(/<(?:script|link)\b[^>]*?\b(?:src|href)\s*=\s*["']([^"']+)["']/gi))candidate(m[1],'index.html');
 for(const m of config.matchAll(/["']([a-z0-9_-]+\.js(?:\?[^"']*)?)["']/gi))candidate(m[1],'firebase-config.js');
 // Include all original top-level modules, and the original snow application tree.
 for(const ent of readdirSync(root,{withFileTypes:true}))if(ent.isFile()&&ent.name.endsWith('.js'))files.add(ent.name);
 function walk(dir){for(const ent of readdirSync(resolve(root,dir),{withFileTypes:true})){const p=dir+'/'+ent.name;if(ent.isDirectory())walk(p);else if(ent.isFile()&&/\.(?:js|html|css)$/.test(p))files.add(p);}}
 if(existsSync(resolve(root,'piano-neve')))walk('piano-neve');
 const archives=new Map(),modules=[];
 for(const file of [...files].sort()){
  if(!file.endsWith('.js'))continue;
  const source=read(file),counts={literal:0,dynamic:0,collectionGroup:0,storage:0,cache:0,network:0,functions:0};
  const roots=new Set();
  for(const m of source.matchAll(/\b(collection|doc|collectionGroup)\s*\(\s*([^)\n]{0,300})/g)){
   const [,method,arg]=m;if(method==='collectionGroup')counts.collectionGroup++;
   const match=arg.match(/^(['"])([A-Za-z0-9_-]+)(?:\/[^'"]*)?\1/);
   if(match){counts.literal++;roots.add(match[2]);}else counts.dynamic++;
  }
  for(const name of roots){if(!archives.has(name))archives.set(name,[]);archives.get(name).push(file);}
  const patterns={storage:/\b(?:firebase\.storage|storage\.ref|getStorage|uploadBytes)\b/g,cache:/\b(?:localStorage|sessionStorage|indexedDB|caches|enablePersistence)\b/g,network:/\b(?:fetch|XMLHttpRequest|WebSocket|sendBeacon)\s*\(/g,functions:/\b(?:httpsCallable|firebase\.functions|getFunctions)\b/g};
  for(const [key,re] of Object.entries(patterns))counts[key]=[...source.matchAll(re)].length;
  if(Object.values(counts).some(Boolean))modules.push({file,...counts});
 }
 return {originalAssets:files.size,missing:[...missing].sort(),externalHosts:[...external].sort(),archiveCandidates:Object.fromEntries([...archives].sort()),modules,productionDeploymentAllowed:false};
}
if(process.argv[1]===fileURLToPath(import.meta.url)){
 const report=auditOriginalRuntime();
 if(report.missing.length){console.error('Original source dependencies missing:',report.missing);process.exitCode=1;}
 console.log(JSON.stringify({assets:report.originalAssets,archiveCandidates:Object.keys(report.archiveCandidates).length,modulesRequiringReview:report.modules.length,missing:report.missing,productionDeploymentAllowed:false},null,2));
}
