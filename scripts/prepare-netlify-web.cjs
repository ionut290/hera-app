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
    assert.equal(item.name,"icons","Unexpected directory in web publish: "+item.name);
    continue;
  }
  assert.ok(item.isFile(),"Unexpected symlink or file type: "+item.name);
  if(unsafe.test(item.name)||!allowed.has(path.extname(item.name).toLowerCase())){
    fs.rmSync(path.join(out,item.name));
  }
}
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
