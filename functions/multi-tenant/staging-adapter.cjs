'use strict';
// Server-side staging entry point, intentionally not exported by main.js.
// No HTTP endpoint or automatic deployment. Never pass credentials from client.
const {createTenantDispatcher}=require('./dispatcher.cjs');
const PROJECT='varga-cantieri-v2-staging';
function assertStagingEnvironment(env=process.env){
  const fail=()=>{throw Object.assign(new Error('staging-only'),{code:'staging-only'});};
  if(env.VARGA_MULTI_TENANT_STAGING!=='1'||env.GOOGLE_CLOUD_PROJECT!==PROJECT)fail();
  if((env.GCLOUD_PROJECT&&env.GCLOUD_PROJECT!==PROJECT)||env.FIRESTORE_EMULATOR_HOST||env.FIREBASE_AUTH_EMULATOR_HOST)fail();
  if(env.FIREBASE_CONFIG){
    let config;try{config=JSON.parse(env.FIREBASE_CONFIG);}catch{fail();}
    if(!config||config.projectId!==PROJECT)fail();
  }
}
function createStagingAdapter(){
  assertStagingEnvironment();
  const {initializeApp,applicationDefault,deleteApp}=require('firebase-admin/app');
  const {getAuth}=require('firebase-admin/auth');
  const {getFirestore}=require('firebase-admin/firestore');
  const app=initializeApp({projectId:PROJECT,credential:applicationDefault()},'varga-staging-'+require('node:crypto').randomUUID());
  const db=getFirestore(app),auth=getAuth(app);
  const assertBoundary=()=>{assertStagingEnvironment();if(app.options.projectId!==PROJECT)throw Object.assign(new Error('staging-only'),{code:'staging-only'});};
  const {dispatch}=createTenantDispatcher({auth,db,assertBoundary});
  return Object.freeze({dispatch,close:async()=>{await db.terminate();await deleteApp(app);}});
}
module.exports={createStagingAdapter,assertStagingEnvironment,PROJECT};
