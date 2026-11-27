import {Pool} from 'pg';
import {readFileSync} from 'node:fs';
import {initializeApp,applicationDefault} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {createApi} from './api.mjs';
const projectId='varga-cantieri-v2-staging';
if(process.env.FIREBASE_PROJECT_ID!==projectId || process.env.FIREBASE_AUTH_EMULATOR_HOST || process.env.FIRESTORE_EMULATOR_HOST)throw Error('Explicit V2 staging project required; no emulator or production fallback');
const container=process.env.VARGA_CONTAINER_BIND==='1';
if(process.env.DATABASE_URL_FILE&&process.env.DATABASE_URL)throw Error('Choose exactly one database credential source');
let dbUrl;
try{dbUrl=new URL(process.env.DATABASE_URL_FILE?readFileSync(process.env.DATABASE_URL_FILE,'utf8').trim():process.env.DATABASE_URL||'')}catch{throw Error('Invalid database connection configuration')}
if(container&&(dbUrl.hostname!=='db'||!process.env.DATABASE_URL_FILE||!process.env.GOOGLE_APPLICATION_CREDENTIALS))throw Error('Container requires dedicated private database and mounted Firebase credentials');
if(process.env.GOOGLE_APPLICATION_CREDENTIALS){
 let credentials;
 try{credentials=JSON.parse(readFileSync(process.env.GOOGLE_APPLICATION_CREDENTIALS,'utf8'))}catch{throw Error('Firebase credentials file unavailable or invalid')}
 if(credentials.type!=='service_account'||credentials.project_id!==projectId)throw Error('Only explicit V2 staging service account credentials accepted');
}
if(!['postgres:','postgresql:'].includes(dbUrl.protocol)||dbUrl.pathname!=='/varga_v2_oracle'||dbUrl.username!=='varga_app')throw Error('Dedicated Oracle PostgreSQL database and limited role required');
const pool=new Pool({connectionString:dbUrl.toString(),max:10,connectionTimeoutMillis:5000,statement_timeout:10000});
const role=(await pool.query("SELECT current_user AS name,rolsuper,rolbypassrls FROM pg_roles WHERE rolname=current_user")).rows[0];
if(role?.name!=='varga_app'||role.rolsuper||role.rolbypassrls)throw Error('Unsafe database role');
const app=initializeApp({credential:applicationDefault(),projectId});
const auth=getAuth(app);
const server=createApi({pool,verifyToken:token=>auth.verifyIdToken(token,true)});
server.requestTimeout=15000;
server.headersTimeout=10000;
server.listen(8080,container?'0.0.0.0':'127.0.0.1',()=>console.log('V2 Oracle candidate API started; host access must remain loopback-only'));
for(const signal of ['SIGINT','SIGTERM'])process.on(signal,()=>server.close(()=>pool.end()));
