"use strict";

// Test the EXISTING Apple-review account candidate's Auth/profile SHAPE,
// not their session, token or production data. Entirely LOCAL emulator.
//
// Key distinction: normal active operator can READ significant live records
// and WRITE some records; an apple_review custom claim blocks those flows.
// Do NOT share this person's existing credentials as an App Review demo
// without data-controller clearance and a tested access boundary.

const assert=require("node:assert/strict");
const fs=require("node:fs");
const {doc,getDoc,setDoc,updateDoc}=require("firebase/firestore");
const {initializeTestEnvironment,assertSucceeds,assertFails}=
  require("@firebase/rules-unit-testing");

async function main(){
  const host=process.env.FIRESTORE_EMULATOR_HOST||"";
  assert.match(host,/^(localhost|127\.0\.0\.1):[0-9]+$/,"Only a LOCAL emulator is allowed");
  const [hostname,port]=host.split(":");
  const env=await initializeTestEnvironment({
    projectId:"demo-operator-access-review",
    firestore:{host:hostname,port:Number(port),rules:fs.readFileSync("firestore.rules","utf8")}
  });
  try{
    await env.withSecurityRulesDisabled(async context=>{
      const db=context.firestore();
      // All fixture content is fake and created in local emulator only.
      await setDoc(doc(db,"platformUsers/syntheticOperator"),{
        uid:"syntheticOperator",email:"aaziz.aziz@operatori.vargacantieri.app",
        statoAccount:"attivo",accountStatus:"attivo",role:"user",banned:false
      });
      await setDoc(doc(db,"platformUsers/syntheticCoworker"),{
        uid:"syntheticCoworker",email:"fake-coworker@example.test",
        statoAccount:"attivo",role:"user"
      });
      await setDoc(doc(db,"platformUsers/syntheticPending"),{
        uid:"syntheticPending",email:"pending@example.test",statoAccount:"in_attesa",role:"user"
      });
      for(const [path,data] of [
        ["commesse/fixtureJob",{name:"Fictional project"}],
        ["commesse/fixtureJob/impianti/fixturePlant",{name:"Fictional plant"}],
        ["squadre/fixtureTeam",{name:"Fictional team"}],
        ["mezzi/fixtureTruck",{name:"Fictional vehicle"}],
        ["personale/fixtureEmployee",{displayName:"Fictional employee"}],
        ["oreReports/fixtureColleagueHours",{date:"2026-01-01",userId:"syntheticCoworker"}],
        ["chatMessages/fixtureOtherChat",{senderId:"syntheticCoworker",message:"Fake private message"}],
        ["calendarEvents/fixtureEvent",{type:"intervento",title:"Fictional event"}],
        ["documents/fixtureGlobal",{id:"fixtureGlobal",visibility:"global",ownerUserId:"syntheticCoworker"}],
        ["globalCommesse/fixtureGlobalJob",{name:"Fictional global job"}],
        ["sharedStaticViews/fixtureView",{title:"Fictional view"}],
        ["appConfig/adminUsers",{emails:["syntheticAdmin@example.test"]}]
      ])await setDoc(doc(db,path),data);
    });
    const email="aaziz.aziz@operatori.vargacantieri.app";
    const normal=env.authenticatedContext("syntheticOperator",{email,email_verified:true}).firestore();
    const claimed=env.authenticatedContext("syntheticOperator",{email,email_verified:true,apple_review:true}).firestore();
    const pending=env.authenticatedContext("syntheticPending",{email:"pending@example.test",email_verified:true}).firestore();

    // Confirm the currently operating account SHAPE can read the standard
    // app, including datasets which could contain personal information.
    const reads=[
      "commesse/fixtureJob","commesse/fixtureJob/impianti/fixturePlant",
      "squadre/fixtureTeam","mezzi/fixtureTruck","personale/fixtureEmployee",
      "oreReports/fixtureColleagueHours","chatMessages/fixtureOtherChat",
      "calendarEvents/fixtureEvent","documents/fixtureGlobal",
      "globalCommesse/fixtureGlobalJob","sharedStaticViews/fixtureView"
    ];
    for(const p of reads)await assertSucceeds(getDoc(doc(normal,p)));
    // Auth's apple_review claim would STOP full-feature production access.
    for(const p of reads)await assertFails(getDoc(doc(claimed,p)));
    for(const p of ["commesse/fixtureJob","personale/fixtureEmployee","oreReports/fixtureColleagueHours"])
      await assertFails(getDoc(doc(pending,p)));

    // Existing normal operator is restricted from certain admin actions.
    for(const p of ["commesse/newFakeJob","squadre/newFakeTeam","mezzi/newFakeTruck","personale/newFakeEmployee"])
      await assertFails(setDoc(doc(normal,p),{name:"Fake"}));
    await assertFails(getDoc(doc(normal,"platformUsers/syntheticCoworker")));
    await assertFails(getDoc(doc(normal,"appConfig/adminUsers")));
    await assertFails(getDoc(doc(normal,"operatorPositions/syntheticCoworker")));

    // Important risk: operating user currently CAN create/edit some data.
    // These writes are executed ONLY in an isolated local emulator.
    await assertSucceeds(setDoc(doc(normal,"oreReports/fixtureEditable"),{
      userId:"syntheticOperator",date:"2026-01-02",hours:1
    }));
    await assertSucceeds(setDoc(doc(normal,"chatMessages/fixtureEditable"),{
      senderId:"syntheticOperator",message:"Fake test"
    }));
    await assertSucceeds(setDoc(doc(normal,"commesse/fixtureJob/impianti/fixtureEditable"),{
      name:"Fake mutable plant"
    }));
    await assertSucceeds(updateDoc(doc(normal,"commesse/fixtureJob/impianti/fixtureEditable"),{name:"Fake changed"}));
    for(const p of [
      "oreReports/forbiddenReview","chatMessages/forbiddenReview",
      "commesse/fixtureJob/impianti/forbiddenReview"
    ])await assertFails(setDoc(doc(claimed,p),{name:"Fake review write"}));
    console.log("PASS: 11 operator reads include real-data classes; 4 local-only writes allowed; admin writes blocked; apple_review claim blocks ordinary app. NO production access.");
  }finally{await env.cleanup();}
}
main().catch(e=>{console.error("Synthetic operator authorization test FAILED:",e?.message||"unknown");process.exitCode=1;});
