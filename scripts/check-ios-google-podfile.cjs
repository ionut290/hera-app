"use strict";
// Generated Podfile fixture only; no CocoaPods install or Apple credentials.
const assert=require("node:assert/strict");
const fs=require("node:fs");
const os=require("node:os");
const path=require("node:path");
const cp=require("node:child_process");
const temp=fs.mkdtempSync(path.join(os.tmpdir(),"varga-google-pod-"));
const script=path.resolve("scripts/install-ios-google-signin.rb");
const podfile=path.join(temp,"Podfile");
try{
  const before=[
    "platform :ios, '15.0'",
    "require_relative '../../node_modules/@capacitor/ios/scripts/pods_helpers'",
    "target 'App' do",
    "  capacitor_pods",
    "  # Add your Pods here",
    "end",
    "post_install do |installer|",
    "  assertDeploymentTarget(installer)",
    "end",""
  ].join("\n");
  fs.writeFileSync(podfile,before);
  function run(){cp.execFileSync("ruby",[script,podfile],{encoding:"utf8",stdio:"pipe"});}
  run();
  const after=fs.readFileSync(podfile,"utf8");
  const dependency="pod 'CapacitorFirebaseAuthentication/Google', :path => '../../node_modules/@capacitor-firebase/authentication'";
  assert.equal(after.split(dependency).length-1,1);
  assert.match(after,/target 'App' do[\s\S]*capacitor_pods[\s\S]*# Add your Pods here[\s\S]*Google/);
  assert.match(after,/post_install do \|installer\|/);
  assert.match(after,/CODE_SIGNING_ALLOWED/);
  assert.match(after,/com\.apple\.product-type\.bundle/);
  assert.match(after,/assertDeploymentTarget\(installer\)/);
  run();
  assert.equal(fs.readFileSync(podfile,"utf8"),after,"Podfile patch must be idempotent");
  fs.writeFileSync(podfile,"target 'App' do\nend\n");
  assert.throws(()=>run(),"Unexpected Podfile format must fail closed");
  const privateWorkflow=fs.readFileSync(".github/workflows/ios-private-testflight.yml","utf8");
  const appStoreWorkflow=fs.readFileSync(".github/workflows/ios-appstore-candidate.yml","utf8");
  for(const [name,workflow] of [["TestFlight",privateWorkflow],["App Store",appStoreWorkflow]]){
    assert.match(workflow,/ruby -c scripts\/install-ios-google-signin\.rb/);
    assert.match(workflow,/npx cap add ios\n[\s\S]*?ruby scripts\/install-ios-google-signin\.rb\n\s*npx cap sync ios/);
    assert.match(workflow,/python3 scripts\/configure-ios-firebase\.py/);
    console.log("PASS: "+name+" includes Google native pod and config before signing.");
  }
  console.log("PASS: Google native pod added once, bundle signing safeguarded, invalid Podfiles rejected.");
}finally{fs.rmSync(temp,{recursive:true,force:true});}
