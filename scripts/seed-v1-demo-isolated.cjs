"use strict";
// V1 DEMO only. No Firebase keys, no IAM changes, no production or V2 access.
// Default: offline plan. Live writes require explicit --apply + exact project.
const assert = require("node:assert/strict");
const { execFileSync } = require("node:child_process");
const PROJECT = "varga-cantieri-v1-demo";
const EMAIL = "operatore.demo@example.com";
const ROOT = "https://firestore.googleapis.com/v1/projects/" + PROJECT + "/databases/(default)";
function fixtureRows(uid) {
  assert.match(uid, /^[A-Za-z0-9_-]{1,128}$/);
  const marker = { synthetic: true };
  const named = name => ({ name, nome: name, ...marker });
  return [
    ["platformUsers/" + uid, { email: EMAIL, displayName: "Operatore Demo", nome: "Operatore Demo", statoAccount: "attivo", accountStatus: "attivo", role: "user", banned: false, ...marker }],
    ["commesse/demoJobAlpha", named("Commessa Demo Verde")],
    ["commesse/demoSnow", named("SERVIZIO NEVE DEMO")],
    ["commesse/demoJobAlpha/impianti/demoPlantOne", { ...named("Impianto Demo 01"), lat: 0, lng: 0 }],
    ["personale/demoPersonOne", { ...named("Operatore Demo"), displayName: "Operatore Demo" }],
    ["mezzi/demoVehicle", named("Mezzo Demo")],
    ["squadre/demoTeam", named("Squadra Demo")]
  ];
}
function fields(data) {
  return Object.fromEntries(Object.entries(data).map(([k,v]) => [k,
    typeof v === "boolean" ? {booleanValue:v} :
    typeof v === "number" ? {doubleValue:v} : {stringValue:v}]));
}
function canonical(value) {
  if (value && typeof value === "object" && !Array.isArray(value))
    return Object.fromEntries(Object.keys(value).sort().map(k => [k,canonical(value[k])]));
  return value;
}
function validateProject(project) {
  if (project !== PROJECT) throw Error("Refused: only the dedicated V1 DEMO project is permitted.");
}
async function seed(request, project) {
  validateProject(project);
  const result = await request("POST", "https://identitytoolkit.googleapis.com/v1/projects/" + PROJECT + "/accounts:lookup", {email:[EMAIL]});
  const users = result.body?.users || [];
  if (result.status !== 200 || users.length !== 1 || users[0].email !== EMAIL || users[0].disabled === true)
    throw Error("Exact active synthetic account could not be verified.");
  const uid = users[0].localId;
  const rows = fixtureRows(uid);
  const writes = [];
  for (const [path, data] of rows) {
    const expected = fields(data);
    const existing = await request("GET", ROOT + "/documents/" + path);
    if (existing.status === 404) {
      writes.push({update:{name:"projects/" + PROJECT + "/databases/(default)/documents/" + path, fields:expected},currentDocument:{exists:false}});
    } else if (existing.status === 200) {
      if (JSON.stringify(canonical(existing.body.fields)) !== JSON.stringify(canonical(expected)))
        throw Error("Existing document differs. No writes performed.");
    } else throw Error("Read access denied or unexpected Firestore response: HTTP " + existing.status);
  }
  if (writes.length) {
    const committed = await request("POST", ROOT + "/documents:commit", {writes});
    if (committed.status !== 200 || committed.body?.writeResults?.length !== writes.length)
      throw Error("Atomic commit not confirmed; inspect demo before retrying.");
  }
  for (const [path,data] of rows) {
    const check = await request("GET",ROOT + "/documents/" + path);
    if (check.status !== 200 || JSON.stringify(canonical(check.body.fields)) !== JSON.stringify(canonical(fields(data))))
      throw Error("Post-write verification incomplete.");
  }
  return {project:PROJECT,created:writes.length,verified:rows.length};
}
async function selfTest() {
  const docs = new Map(), writes = [];
  const mock = async (method,url,body) => {
    if(url.includes("identitytoolkit")) return {status:200,body:{users:[{email:EMAIL,localId:"synthetic_uid"}]}};
    if(method==="POST") {
      assert.equal(url,ROOT+"/documents:commit");
      for(const w of body.writes) {
        assert.equal(w.currentDocument.exists,false);
        assert.ok(w.update.name.startsWith("projects/"+PROJECT+"/"));
        const path=w.update.name.split("/documents/")[1];
        assert.ok(!docs.has(path));
        docs.set(path,{fields:w.update.fields}); writes.push(path);
      }
      return {status:200,body:{writeResults:body.writes.map(()=>({}))}};
    }
    const path=url.split("/documents/")[1];
    return docs.has(path)?{status:200,body:docs.get(path)}:{status:404,body:{}};
  };
  await assert.rejects(seed(mock,"wrong-project"));
  assert.equal(writes.length,0);
  assert.deepEqual(await seed(mock,PROJECT),{project:PROJECT,created:7,verified:7});
  assert.deepEqual(await seed(mock,PROJECT),{project:PROJECT,created:0,verified:7});
  docs.get("mezzi/demoVehicle").fields.name={stringValue:"Unexpected"};
  await assert.rejects(seed(mock,PROJECT),/differs/);
  assert.equal(writes.length,7);
  console.log("PASS: exact-project guard, seven synthetic records, atomic create-only, idempotence, conflict rejection.");
}
async function main() {
  const args=process.argv.slice(2);
  if(args.length===1 && args[0]==="--self-test") return selfTest();
  if(!args.length) {
    console.log(JSON.stringify({mode:"OFFLINE_PLAN",project:PROJECT,documents:fixtureRows("SYNTHETIC_UID").map(([path,data])=>({path,data}))},null,2));
    return;
  }
  assert.deepEqual(args,["--apply","--project",PROJECT],"Explicit exact-project confirmation required.");
  validateProject(args[2]);
  const token=execFileSync("gcloud",["auth","print-access-token","--quiet"],{encoding:"utf8",stdio:["ignore","pipe","pipe"]}).trim();
  assert.ok(token.length>30,"No authenticated Google Cloud session.");
  const request=async (method,url,body)=>{
    if(!url.startsWith(ROOT+"/") && url!=="https://identitytoolkit.googleapis.com/v1/projects/"+PROJECT+"/accounts:lookup")
      throw Error("Endpoint outside demo allowlist.");
    const response=await fetch(url,{method,headers:{Authorization:"Bearer "+token,"Content-Type":"application/json"},...(body?{body:JSON.stringify(body)}:{}),signal:AbortSignal.timeout(30000)});
    return {status:response.status,body:await response.json().catch(()=>({}))};
  };
  console.log(JSON.stringify(await seed(request,PROJECT)));
}
main().catch(e=>{console.error("DEMO seed stopped:",e instanceof assert.AssertionError?"Validation failed":String(e.message).slice(0,200));process.exitCode=1;});
