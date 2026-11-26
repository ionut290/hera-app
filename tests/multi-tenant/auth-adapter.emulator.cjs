'use strict';
// Local integration adapter. Never import in production; fail closed outside dedicated emulators.
const {initializeApp,deleteApp}=require('firebase-admin/app');
const {getAuth}=require('firebase-admin/auth');
const {getFirestore}=require('firebase-admin/firestore');
const {createTenantDispatcher}=require('../../functions/multi-tenant/dispatcher.cjs');
const PROJECT='demo-varga-multi-tenant';
function fail(code){throw Object.assign(new Error(code),{code});}
function validateEmulators(){
  if(process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8180'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9199')fail('emulator-only');
}
function createLocalAdapter(){
  validateEmulators();
  const app=initializeApp({projectId:PROJECT},'tenant-adapter-'+require('node:crypto').randomUUID());
  const db=getFirestore(app),auth=getAuth(app);
  const {dispatch}=createTenantDispatcher({auth,db,assertBoundary:validateEmulators});
  return {dispatch,auth,db,close:async()=>{await db.terminate();await deleteApp(app);}};
}
module.exports={createLocalAdapter,PROJECT};
