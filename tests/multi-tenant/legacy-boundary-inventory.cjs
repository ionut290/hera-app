'use strict';
// Static inventory only. Does not read data, credentials, or connect to a service.
// It reports source locations and collection/key identifiers, never source lines.
const fs=require('node:fs'),path=require('node:path');
const root=path.resolve(__dirname,'../..');
const report={schema:1,scope:'static heuristic; not an isolation certification',filesScanned:0,findings:[]};
const roots=['functions','netlify','piano-neve'];
const allowed=/\.(?:js|cjs|mjs|rules)$/;
function walk(dir,base=''){
  for(const ent of fs.readdirSync(dir,{withFileTypes:true})){
    if(ent.name.startsWith('.')||['node_modules','www','tests','multi-tenant','dist','build'].includes(ent.name))continue;
    const rel=path.join(base,ent.name),file=path.join(dir,ent.name);
    if(ent.isDirectory())walk(file,rel);
    else if(ent.isFile()&&allowed.test(rel))scan(file,rel);
  }
}
function scan(file,relative){
  const source=fs.readFileSync(file,'utf8');report.filesScanned++;
  const patterns=[
    ['collection-literal',/\.collection\(\s*["']([A-Za-z0-9_/-]{1,120})["']/g],
    ['collection-resolver',/\.collection\(\s*([A-Za-z_$][\w$]*)(?:\(|\s*[,)]|\s*\.)/g],
    ['callable',/\.httpsCallable\(\s*["']([A-Za-z0-9_-]{1,100})["']/g],
    ['cache-constant',/\b(?:localStorage|sessionStorage)\s*\.\s*(?:getItem|setItem|removeItem)\(\s*([A-Za-z_$][\w$]*)\s*[,)]/g],
    ['cache-literal',/\b(?:localStorage|sessionStorage)\s*\.\s*(?:getItem|setItem|removeItem)\(\s*["']([A-Za-z0-9_:-]{1,120})["']/g],
    ['storage-reference',/\b(?:storage|firebase\.storage\(\))\s*\.\s*(ref|refFromURL)\s*\(/g],
    ['storage-global-clear',/\b(localStorage|sessionStorage)\s*\.\s*clear\s*\(/g],
    ['indexed-db',/\bindexedDB\s*\.\s*(open|deleteDatabase)\s*\(/g],
    ['firestore-trigger',/\b(onDocumentCreated|onDocumentUpdated|onDocumentWritten|onDocumentDeleted)\s*\(/g],
    ['external-endpoint',/https?:\/\/([A-Za-z0-9.-]+\.(?:googleapis\.com|cloudfunctions\.net|netlify\.app|firebaseapp\.com|web\.app|firebaseio\.com|run\.app))\b/g]
  ];
  for(const [category,re] of patterns){
    let last=0,line=1;
    for(const m of source.matchAll(re)){
      line+=(source.slice(last,m.index).match(/\n/g)||[]).length;last=m.index;
      // Hostnames are intentionally omitted, including deployment-specific project names.
      report.findings.push({file:relative.replaceAll('\\','/'),line,category,identifier:category==='external-endpoint'?'external-service':m[1]});
    }
  }
}
for(const ent of fs.readdirSync(root,{withFileTypes:true})){
  if(ent.isFile()&&allowed.test(ent.name)&&!/^tenant-/.test(ent.name))scan(path.join(root,ent.name),ent.name);
}
for(const dir of roots){const full=path.join(root,dir);if(fs.existsSync(full))walk(full,dir);}
report.findings.sort((a,b)=>a.file.localeCompare(b.file)||a.line-b.line||a.category.localeCompare(b.category));
const counts={};for(const f of report.findings)counts[f.category]=(counts[f.category]||0)+1;
report.counts=counts;
const assert=require('node:assert/strict');
assert.ok(report.filesScanned>100,'Original source inventory incomplete');
for(const [category,identifier] of [['collection-literal','commesse'],['collection-literal','platformUsers'],['cache-constant','COMMESSE_LOCAL_CACHE_KEY'],['collection-resolver','getCommesseCollectionName']])assert.ok(report.findings.some(f=>f.file==='app.js'&&f.category===category&&f.identifier===identifier),'Expected original boundary missing: '+identifier);
const out=process.env.VARGA_INVENTORY_OUTPUT;
if(out)fs.writeFileSync(out,JSON.stringify(report,null,2)+'\n',{mode:0o600});
console.log(JSON.stringify({filesScanned:report.filesScanned,counts,scope:report.scope}));
