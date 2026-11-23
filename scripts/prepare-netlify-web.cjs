"use strict";
// Safe production Netlify output: reuse tested native web shell, not the repo
// root, which contains docs, scripts, project and business tooling.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const path=require("node:path");
const root=path.resolve(__dirname,".."), out=path.join(root,"www");
assert.ok(fs.statSync(out).isDirectory(),"First run prepare-capacitor-web.js");
const allowed = new Set([".js",".css",".html",".webmanifest",".png",".jpg",".jpeg",".webp",".svg",".gif",".ico",".woff",".woff2",".ttf",".avif"]);
const unsafe = /^(?:firebase\.json|package(?:-lock)?\.json|google-services\.json|firestore\.indexes\.json)$/i;
for(const item of fs.readdirSync(out,{withFileTypes:true})){
  if(item.isDirectory()){
    assert.ok(item.name==="icons"||item.name==="piano-neve","Unexpected directory in web publish: "+item.name);
    continue;
  }
  assert.ok(item.isFile(),"Unexpected symlink or file type: "+item.name);
  if(unsafe.test(item.name)||!allowed.has(path.extname(item.name).toLowerCase())){
    fs.rmSync(path.join(out,item.name));
  }
}
// Fail-closed nested bundle: only the original Piano Neve code and the
// verbatim GPL license for the included leaflet-rotate vendor script.
const pianoRoot=path.join(out,"piano-neve");
assert.ok(fs.statSync(pianoRoot).isDirectory(),"Bundled Piano Neve missing");
function verifyPianoFolder(dir,relative=""){
  for(const file of fs.readdirSync(dir,{withFileTypes:true})){
    const rel=relative?relative+"/"+file.name:file.name;
    if(file.isDirectory()){
      assert.ok(relative===""&&(file.name==="vendor"||file.name==="commesse"),
        "Unknown Piano Neve directory: "+rel);
      verifyPianoFolder(path.join(dir,file.name),rel);
    }else{
      assert.ok(file.isFile(),"Unexpected symlink in Piano Neve: "+rel);
      assert.ok(/\.(?:js|css|html)$/i.test(file.name)
        || rel==="vendor/leaflet-rotate-LICENSE.txt",
        "Unexpected Piano Neve public file: "+rel);
    }
  }
}
verifyPianoFolder(pianoRoot);
// Root-level icons/fonts used by public web CSS can be copied with safe file
// extensions; never recurse through arbitrary folders.
for(const item of fs.readdirSync(root,{withFileTypes:true})){
  if(!item.isFile()) continue;
  const ext=path.extname(item.name).toLowerCase();
  if(![".png",".jpg",".jpeg",".webp",".svg",".gif",".ico",".woff",".woff2",".ttf",".avif"].includes(ext))continue;
  if(item.name.startsWith("."))continue;
  fs.copyFileSync(path.join(root,item.name),path.join(out,item.name));
}
const kept=fs.readdirSync(out);
assert.ok(kept.includes("index.html")&&kept.includes("delete-account.html")&&kept.includes("sw.js"));
assert.ok(!kept.some(x=>unsafe.test(x)||x==="docs"||x==="scripts"||x==="functions"||x==="netlify"||x===".github"));
console.log("PASS: Netlify web-only package ready; no repository root, package metadata or internal folders.");
