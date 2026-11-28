// Separate browser portal. Strictly emulator + synthetic fixtures, no Stripe network.
import http from 'node:http';
import {readFile} from 'node:fs/promises';
import {pathToFileURL} from 'node:url';
import Stripe from 'stripe';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {getFirestore} from 'firebase-admin/firestore';
import {SUBSCRIPTION_PLANS} from '../subscription-plans.mjs';
import {createSandboxCheckout,createSandboxStripeProvider} from '../stripe-sandbox-checkout.mjs';
import {SANDBOX_PRICE_IDS,auditSandboxCatalog} from '../stripe-sandbox-catalog.mjs';
import {createStripeTestBilling} from '../stripe-test-billing.mjs';
const projectId='varga-cantieri-v2-staging',origin='http://127.0.0.1:4173',host='127.0.0.1:4173';
const secret='whsec_local_web_fixture_only_not_real';
const priceIds=Object.fromEntries(Object.keys(SUBSCRIPTION_PLANS).map(t=>[t,'price_web_fixture_'+t]));
const fail=code=>{throw Object.assign(Error(code),{code});};
const version='2024-10-28.acacia';
function fixtureProvider(){
 const sessions=new Map(),prices=new Map();
 for(const [tier,p] of Object.entries(SUBSCRIPTION_PLANS)){
  const id=priceIds[tier];prices.set(id,{id,livemode:false,active:true,currency:'eur',unit_amount:p.priceCents,type:'recurring',recurring:{interval:'day',interval_count:30,usage_type:'licensed'},billing_scheme:'per_unit'});
 }
 return {
  async price(id){return structuredClone(prices.get(id));},
  async create(params,key){
   const id='cs_test_web_'+key.replace(/[^A-Za-z0-9_]/g,'_');
   if(!sessions.has(id)){
    const price=prices.get(params.line_items?.[0]?.price);if(!price)fail('PRICE_MISMATCH');
    sessions.set(id,{id,livemode:false,mode:'subscription',status:'open',payment_status:'unpaid',subscription:null,client_reference_id:params.client_reference_id,metadata:params.metadata,url:'https://checkout.stripe.com/c/pay/'+id,price});
   }
   return structuredClone(sessions.get(id));
  },
  async session(id){return structuredClone(sessions.get(id));},
  async expire(id){const s=sessions.get(id);if(!s||s.status!=='open'||s.payment_status!=='unpaid'||s.subscription!==null)fail('CHECKOUT_RECONCILIATION_REQUIRED');s.status='expired';return structuredClone(s);},
  complete(id){
   const s=sessions.get(id);if(!s)fail('CHECKOUT_RECONCILIATION_REQUIRED');
   if(s.event)return structuredClone(s.event);
   if(s.status!=='open')fail('CHECKOUT_RECONCILIATION_REQUIRED');
   const p=s.price,suffix=id.slice('cs_test_'.length),customer='cus_'+suffix,subscription='sub_'+suffix,invoice='in_'+suffix,start=Math.floor(Date.now()/1000);
   Object.assign(s,{status:'complete',payment_status:'paid',currency:'eur',amount_total:p.unit_amount,customer,
    line_items:{has_more:false,data:[{quantity:1,amount_total:p.unit_amount,price:p}]},
    subscription:{id:subscription,livemode:false,status:'active',customer,metadata:s.metadata,items:{has_more:false,data:[{quantity:1,price:p}]}}});
   s.event={id:'evt_'+suffix,object:'event',api_version:version,livemode:false,type:'invoice.paid',created:start,data:{object:{id:invoice,object:'invoice',livemode:false,subscription,customer,status:'paid',paid:true,paid_out_of_band:false,payment_intent:'pi_'+suffix,currency:'eur',amount_paid:p.unit_amount,amount_due:p.unit_amount,amount_remaining:0,collection_method:'charge_automatically',billing_reason:'subscription_create',lines:{has_more:false,data:[{type:'subscription',proration:false,quantity:1,amount:p.unit_amount,currency:'eur',price:p,period:{start,end:start+30*86400}}]}}}};
   return structuredClone(s.event);
  }
 };
}
export function createLocalBillingServer({app,mode='fixture'}){
 if(app?.options?.projectId!==projectId||process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8080'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9099')fail('EMULATORS_REQUIRED');
 if(!['fixture','stripe-sandbox'].includes(mode))fail('INVALID_CONFIG');
 const external=mode==='stripe-sandbox';
 const provider=external?createSandboxStripeProvider({app,secretKey:process.env.VARGA_STRIPE_SANDBOX_CHECKOUT_KEY}):fixtureProvider(),catalog=external?SANDBOX_PRICE_IDS:priceIds;
 const checkout=createSandboxCheckout({app,provider,priceIds:catalog}),billing=createStripeTestBilling({app,webhookSecret:secret,priceIds:catalog,authoritativeProvider:external?provider:null}),stripe=new Stripe('sk_test_LOCAL_FIXTURE_ONLY',{apiVersion:version}),db=getFirestore(app),auth=getAuth(app);
 const headers={'Cache-Control':'no-store','X-Content-Type-Options':'nosniff','Referrer-Policy':'no-referrer','X-Frame-Options':'DENY','Content-Security-Policy':"default-src 'none'; script-src 'self'; style-src 'self'; connect-src 'self' http://127.0.0.1:9099; frame-ancestors 'none'; form-action 'none'; base-uri 'none'"};
 const assets={'/':['billing.html','text/html'],'/billing-test-return.html':['billing.html','text/html'],'/billing.css':['billing.css','text/css'],'/billing.js':['billing.js','application/javascript']};
 const actions=new Set(['catalog','companies','request-plan','checkout','simulate','cancel','verify']);
 async function simulate(token,tenantId){
  if(external)fail('SIMULATION_DISABLED');
  // Only a verified owner can complete this synthetic session. No external API exists here.
  const allowed=(await checkout.companies({token})).companies.some(c=>c.id===tenantId&&['open','bound'].includes(c.checkoutStatus));
  if(!allowed)fail('FORBIDDEN');
  const claims=await auth.verifyIdToken(token,true),attempt=(await db.doc('stripeSandboxCheckouts/'+tenantId).get()).data();
  if(!attempt||attempt.ownerUid!==claims.uid||!attempt.sessionId)fail('FORBIDDEN');
  const event=provider.complete(attempt.sessionId);
  await checkout.finalizeCheckout({token,tenantId});
  const payload=JSON.stringify(event),signature=stripe.webhooks.generateTestHeaderString({payload,secret});
  return billing.handleWebhook(Buffer.from(payload),signature);
 }
 return http.createServer(async(req,res)=>{
  const reply=(status,data)=>{res.writeHead(status,{...headers,'Content-Type':'application/json; charset=utf-8'});res.end(JSON.stringify(data));};
  try{
   if(req.headers.host!==host)return reply(403,{code:'FORBIDDEN'});
   if(req.method==='GET'&&Object.hasOwn(assets,req.url)){
    const [file,type]=assets[req.url],body=await readFile(new URL(file,import.meta.url));
    res.writeHead(200,{...headers,'Content-Type':type+'; charset=utf-8'});return res.end(body);
   }
   if(req.method!=='POST'||req.headers.origin!==origin||!String(req.headers['content-type']).startsWith('application/json'))return reply(403,{code:'FORBIDDEN'});
   const action=req.url.startsWith('/api/')?req.url.slice(5):'';if(!actions.has(action))return reply(404,{code:'NOT_FOUND'});
   const token=String(req.headers.authorization||'').match(/^Bearer ([A-Za-z0-9._-]+)$/)?.[1];if(!token)return reply(401,{code:'UNAUTHENTICATED'});
   let size=0;const chunks=[];for await(const chunk of req){size+=chunk.length;if(size>8192)return reply(413,{code:'INVALID_REQUEST'});chunks.push(chunk);}
   const data=JSON.parse(Buffer.concat(chunks).toString('utf8'));if(!data||typeof data!=='object'||Array.isArray(data))return reply(400,{code:'INVALID_REQUEST'});
   if(action==='catalog'){await checkout.companies({token});return reply(200,{plans:Object.entries(SUBSCRIPTION_PLANS).map(([tier,p])=>({tier,...p})),testMode:true,mode});}
   if(action==='companies')return reply(200,await checkout.companies({token}));
   if(action==='request-plan')return reply(200,await checkout.requestPlan({token,tenantId:data.tenantId,tier:data.tier}));
   if(action==='checkout'){
    if(external)await auditSandboxCatalog(provider);
    const result=await checkout.createCheckout({token,tenantId:data.tenantId,tier:data.tier,acceptTerms:data.acceptTerms});
    // Never return the fixture's placeholder external URL or any payment secret.
    return reply(200,{status:result.status,testMode:true,...(external?{url:result.url}:{})});
   }
   if(action==='cancel')return reply(200,await checkout.cancelCheckout({token,tenantId:data.tenantId}));
   if(action==='simulate')return reply(200,await simulate(token,data.tenantId));
   if(action==='verify'){
    if(!external)fail('RECONCILIATION_UNAVAILABLE');
    const bound=await checkout.finalizeCheckout({token,tenantId:data.tenantId});
    return reply(200,await billing.reconcileInitialInvoice(bound.subscriptionId));
   }
  }catch(e){const code=typeof e?.code==='string'&&/^[A-Z_]{3,48}$/.test(e.code)?e.code:'REQUEST_FAILED';return reply(code==='UNAUTHENTICATED'?401:code==='FORBIDDEN'?403:400,{code});}
 });
}
if(process.argv[1]&&import.meta.url===pathToFileURL(process.argv[1]).href){
 const app=initializeApp({projectId}),mode=process.argv.includes('--stripe-sandbox')?'stripe-sandbox':'fixture',server=createLocalBillingServer({app,mode});
 server.listen(4173,'127.0.0.1',()=>console.log('LOCAL BILLING '+mode+' ONLY: '+origin));
 process.once('SIGINT',()=>server.close(async()=>{await deleteApp(app);process.exit(0);}));
}
