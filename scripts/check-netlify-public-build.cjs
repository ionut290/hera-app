"use strict";
// Validate the exact Netlify site publish tree. This is a build contract,
// not proof of any completed production deployment.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const path=require("node:path");
const root=path.resolve(__dirname,".."),out=path.join(root,"www");
const toml=fs.readFileSync(path.join(root,"netlify.toml"),"utf8");
assert.match(toml,/publish = "www"/);
assert.match(toml,/prepare-capacitor-web\.js && node scripts\/prepare-netlify-web\.cjs && node scripts\/check-netlify-public-build\.cjs/);
const publicAssets=fs.readdirSync(out,{withFileTypes:true});
assert.ok(publicAssets.length>40,"Expected full public app, not a partial build");
for(const item of publicAssets)if(item.isDirectory())assert.equal(item.name,"icons","Unapproved public folder: "+item.name);
for(const item of publicAssets)if(item.isFile()){
  assert.ok(/\.(?:js|css|html|webmanifest|png|jpe?g|webp|svg|gif|ico|woff2?|ttf|avif)$/i.test(item.name),"Unapproved web file type: "+item.name);
}
for(const rel of ["index.html","delete-account.html","privacy.html","sw.js","manifest.webmanifest","app.js","firebase-config.js","timbratura-ios-notifications.js","account-privacy-controls.js","icons/varga-cantieri-192.png"]){
  assert.ok(fs.existsSync(path.join(out,rel)),"Required web asset missing: "+rel);
}
for(const file of ["firebase.json","package.json","package-lock.json","README.md","netlify.toml","firestore.rules","firestore.indexes.json","storage.rules",".firebaserc"]){
  assert.ok(!fs.existsSync(path.join(out,file)),"Internal config exposed by Netlify: "+file);
}
for(const folder of ["docs","scripts","functions","netlify",".github","ios","android","node_modules"]){
  assert.ok(!fs.existsSync(path.join(out,folder)),"Internal path exposed by Netlify: "+folder);
}
const html=fs.readFileSync(path.join(out,"index.html"),"utf8");
const references=[...html.matchAll(/(?:src|href)=["']([^"']+)["']/g)].map(x=>x[1]);
for(const value of references){
  if(/^(?:https?:|data:|blob:|#|mailto:|tel:|javascript:)/i.test(value))continue;
  const rel=value.split(/[?#]/)[0].replace(/^\.?\//,"");
  if(!rel)continue;
  // Skip placeholder/dynamic or external URL schemes.
  if(/[{}$]/.test(rel)||rel.includes("://"))continue;
  assert.ok(fs.existsSync(path.join(out,rel)),"Broken local public HTML asset: "+rel);
}
console.log("PASS: isolated Netlify publish folder contains app assets but no internal repo content.");
