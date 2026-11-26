import {readFileSync,writeFileSync,readdirSync,mkdirSync,cpSync,rmSync,existsSync} from 'node:fs';
import {resolve,dirname,extname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {build} from 'vite';
import {createRequire} from 'node:module';
import {buildIntegratedRuntimeCandidate} from '../original-runtime-source.mjs';
const here=dirname(fileURLToPath(import.meta.url)),root=resolve(here,'../..'),out=resolve(here,'full-runtime-dist');
const read=p=>readFileSync(resolve(root,p),'utf8');
const req=createRequire(import.meta.url);
rmSync(out,{recursive:true,force:true});mkdirSync(out,{recursive:true});
await build({root:here,configFile:false,build:{outDir:out,emptyOutDir:false,lib:{entry:resolve(here,'full-runtime-launcher.js'),name:'VargaLocalLauncher',formats:['iife'],fileName:()=> 'launcher.js'},minify:false}});
const original=read('index.html');
const oldCsp="default-src 'none'; base-uri 'none'; form-action 'none'";
const newCsp="default-src 'none'; script-src 'self' 'unsafe-inline'; style-src 'self' 'unsafe-inline'; img-src 'self' data: blob:; font-src 'self'; connect-src 'none'; frame-src 'none'; worker-src 'none'; base-uri 'none'; form-action 'none'";
if(original.split(oldCsp).length!==2)throw Error('Quarantine marker missing');
let html=original.replace(oldCsp,newCsp).replace(/<script\b[^>]*\bsrc=["']https:\/\/www\.gstatic\.com\/firebasejs\/[^"']+["'][^>]*><\/script>/gi,'')
 .replace('https://unpkg.com/leaflet@1.9.4/dist/leaflet.js','./vendor/leaflet.js').replace('https://unpkg.com/leaflet@1.9.4/dist/leaflet.css','./vendor/leaflet.css');
html=html.replace('</head>','</head>'); // Insert before any existing script.
html=html.replace(/<script\b/i,'<script src="runtime-bootstrap.js"></script>\n<script');
if(!html.includes('runtime-bootstrap.js'))throw Error('Bootstrap missing');
for(const e of readdirSync(root,{withFileTypes:true})){
 if(!e.isFile()||e.name.startsWith('apple-review-')||!['.js','.css','.webmanifest'].includes(extname(e.name))||e.name==='app.js'||e.name==='sw.js')continue;
 cpSync(resolve(root,e.name),resolve(out,e.name));
}
for(const dir of ['icons','piano-neve'])if(existsSync(resolve(root,dir)))cpSync(resolve(root,dir),resolve(out,dir),{recursive:true});
 // Do not activate any independently booting sub-page until its own security is reviewed.
function quarantinePages(dir){for(const e of readdirSync(dir,{withFileTypes:true})){const p=resolve(dir,e.name);if(e.isDirectory())quarantinePages(p);else if(e.name.endsWith('.html')){let s=readFileSync(p,'utf8');s=s.replace(/<head[^>]*>/i,m=>m+"<meta http-equiv=\"Content-Security-Policy\" content=\"default-src 'none'; base-uri 'none'; form-action 'none'\">");writeFileSync(p,s);}}}
if(existsSync(resolve(out,'piano-neve')))quarantinePages(resolve(out,'piano-neve'));
mkdirSync(resolve(out,'vendor'));for(const file of ['leaflet.js','leaflet.css'])cpSync(req.resolve('leaflet/dist/'+file),resolve(out,'vendor',file));
writeFileSync(resolve(out,'index.html'),readFileSync(resolve(here,'full-runtime-index.html')));
writeFileSync(resolve(out,'original.html'),html);
writeFileSync(resolve(out,'app.js'),buildIntegratedRuntimeCandidate(read('app.js')));
cpSync(resolve(here,'full-runtime-bootstrap.js'),resolve(out,'runtime-bootstrap.js'));
writeFileSync(resolve(out,'LOCAL_TEST_ONLY.txt'),'Local emulator candidate. No production, public deployment, signing or TestFlight upload authorized.\n');
console.log('FULL ORIGINAL RUNTIME PACKAGE: '+readdirSync(out).length+' top-level assets, launcher and original source.');
