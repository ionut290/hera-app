// Only localhost Auth/Firestore and an explicitly enabled Stripe sandbox.
import Stripe from 'stripe';
import {randomUUID} from 'node:crypto';
import {getFirestore,FieldValue} from 'firebase-admin/firestore';
import {getAuth} from 'firebase-admin/auth';
import {SUBSCRIPTION_PLANS} from './subscription-plans.mjs';
import {SANDBOX_PRICE_IDS,verifySandboxPrice} from './stripe-sandbox-catalog.mjs';
const PROJECT='varga-cantieri-v2-staging',VERSION='2024-10-28.acacia';
const fail=code=>{throw Object.assign(new Error(code),{code});};
const tenantOk=v=>typeof v==='string'&&/^demo-[a-z0-9-]{3,59}$/.test(v);
const id=(v,p)=>typeof v==='string'&&v.startsWith(p)&&/^[A-Za-z0-9_]{5,200}$/.test(v);
const fits=(u,l)=>Object.entries(l).every(([k,v])=>Number.isSafeInteger(u?.[k])&&u[k]>=0&&u[k]<=v);
const stamp=()=>FieldValue.serverTimestamp();
function guard(app){
 if(app?.options?.projectId!==PROJECT||process.env.FIRESTORE_EMULATOR_HOST!=='127.0.0.1:8080'||process.env.FIREBASE_AUTH_EMULATOR_HOST!=='127.0.0.1:9099')fail('EMULATORS_REQUIRED');
}
const checkedPrice=verifySandboxPrice;
function checkedUrl(value){
 let u;try{u=new URL(value);}catch{fail('INVALID_CHECKOUT_URL');}
 if(u.protocol!=='https:'||u.hostname!=='checkout.stripe.com'||u.username||u.password||u.port)fail('INVALID_CHECKOUT_URL');return u.href;
}
// Real provider construction is deliberately opt-in. No live key or connected account.
export function createSandboxStripeProvider({app,secretKey}){
 guard(app);
 if(process.env.VARGA_ENABLE_STRIPE_SANDBOX!=='yes'||typeof secretKey!=='string'||!/^rk_test_[A-Za-z0-9]+$/.test(secretKey))fail('SANDBOX_DISABLED');
 const stripe=new Stripe(secretKey,{apiVersion:VERSION,maxNetworkRetries:2,timeout:15000});
 return Object.freeze({
  price:id=>stripe.prices.retrieve(id),
  create:(params,key)=>stripe.checkout.sessions.create(params,{idempotencyKey:key}),
  session:id=>stripe.checkout.sessions.retrieve(id,{expand:['subscription','line_items.data.price']}),
  expire:id=>stripe.checkout.sessions.expire(id,{}, {idempotencyKey:'varga-expire-'+id}),
  async initialPaidInvoice(subscriptionId){
   const sub=await stripe.subscriptions.retrieve(subscriptionId);
   if(sub.livemode!==false||sub.status!=='active'||sub.id!==subscriptionId||!id(sub.latest_invoice,'in_'))fail('CHECKOUT_NOT_VERIFIED');
   const invoice=await stripe.invoices.retrieve(sub.latest_invoice);
   if(invoice.livemode!==false||invoice.subscription!==sub.id||invoice.customer!==sub.customer||invoice.billing_reason!=='subscription_create'||!id(invoice.payment_intent,'pi_'))fail('CHECKOUT_NOT_VERIFIED');
   const pi=await stripe.paymentIntents.retrieve(invoice.payment_intent);
   if(pi.livemode!==false||pi.id!==invoice.payment_intent||pi.status!=='succeeded'||pi.amount_received!==invoice.amount_paid||pi.currency!==invoice.currency||pi.customer!==invoice.customer)fail('CHECKOUT_NOT_VERIFIED');
   return invoice;
  },
 });
}
export function createSandboxCheckout({app,provider,priceIds=SANDBOX_PRICE_IDS}){
 guard(app);
 if(!provider||!['price','create','session'].every(k=>typeof provider[k]==='function')||!Object.keys(SUBSCRIPTION_PLANS).every(t=>id(priceIds?.[t],'price_'))||new Set(Object.values(priceIds)).size!==3)fail('INVALID_CONFIG');
 const db=getFirestore(app),auth=getAuth(app);
 async function uidFor(token){
  if(typeof token!=='string'||!token)fail('UNAUTHENTICATED');
  let c;try{c=await auth.verifyIdToken(token,true);}catch{fail('UNAUTHENTICATED');}
  // Synthetic users only: never send real email or billing information from emulator data.
  if(!c.uid||c.email_verified!==true||c.apple_review===true||!String(c.email||'').endsWith('@example.test'))fail('UNAUTHENTICATED');
  return c.uid;
 }
 function refs(tenantId,uid){
  if(!tenantOk(tenantId))fail('INVALID_TENANT');
  const tenant=db.doc('tenants/'+tenantId);
  return {tenant,member:tenant.collection('members').doc(uid),mirror:db.doc('users/'+uid+'/memberships/'+tenantId),attempt:db.doc('stripeSandboxCheckouts/'+tenantId)};
 }
 async function owner(tx,r,uid,tenantId){
  const [ts,ms,is]=await Promise.all([tx.get(r.tenant),tx.get(r.member),tx.get(r.mirror)]),t=ts.data(),m=ms.data(),i=is.data();
  if(!t||t.demo!==true||t.status!=='active'||t.ownerUid!==uid||m?.role!=='admin'||m.status!=='active'||i?.role!=='admin'||i.status!=='active'||i.tenantId!==tenantId)fail('FORBIDDEN');
  return t;
 }
 async function createCheckout({token,tenantId,tier,acceptTerms}){
  const uid=await uidFor(token),r=refs(tenantId,uid),p=SUBSCRIPTION_PLANS[tier];
  if(!p||!Object.hasOwn(SUBSCRIPTION_PLANS,tier)||acceptTerms!==true)fail('INVALID_REQUEST');
  const freshId=randomUUID();
  const a=await db.runTransaction(async tx=>{
   const t=await owner(tx,r,uid,tenantId),snap=await tx.get(r.attempt),old=snap.data();
   if(t.plan!=='trial'||t.billing||t.stripeTestSubscriptionId||t.requestedTier!==tier||t.upgradeRequested!==true)fail('BILLING_CONFLICT');
   if(!fits(t.usage,p.limits))fail('USAGE_EXCEEDS_PLAN');
   if(old){if(old.ownerUid!==uid||old.tier!==tier||t.stripeSandboxCheckoutId!==old.attemptId||!['creating','open'].includes(old.status))fail('CHECKOUT_CONFLICT');
    // Stripe may remove idempotency keys after 24h. Do not retry old ambiguous requests.
    if(Date.now()-old.createdAt.toMillis()>20*3600000)fail('CHECKOUT_RECONCILIATION_REQUIRED');return old;
   }
   if(t.stripeSandboxCheckoutId)fail('CHECKOUT_CONFLICT');
   const data={attemptId:freshId,ownerUid:uid,tier,status:'creating',createdAt:new Date(),priceId:priceIds[tier],termsVersion:'sandbox-v1'};
   tx.create(r.attempt,data);tx.update(r.tenant,{stripeSandboxCheckoutId:freshId});
   return data;
  });
  const price=await provider.price(a.priceId);checkedPrice(price,tier,a.priceId);
  const meta={tenantId,attemptId:a.attemptId,tier};
  const session=a.sessionId?await provider.session(a.sessionId):await provider.create({
   mode:'subscription',line_items:[{price:a.priceId,quantity:1}],
   client_reference_id:a.attemptId,metadata:meta,subscription_data:{metadata:meta},
   payment_method_types:['card'],billing_address_collection:'required',
   success_url:'http://127.0.0.1:4173/billing-test-return.html',cancel_url:'http://127.0.0.1:4173/billing-test-return.html'
  },'varga-sandbox-'+a.attemptId);
  if(session?.livemode!==false||!id(session.id,'cs_test_')||session.mode!=='subscription'||session.client_reference_id!==a.attemptId||session.metadata?.tenantId!==tenantId||session.metadata?.tier!==tier||session.metadata?.attemptId!==a.attemptId||session.status!=='open')fail('CHECKOUT_RECONCILIATION_REQUIRED');
  const url=checkedUrl(session.url);
  await db.runTransaction(async tx=>{
   const t=await owner(tx,r,uid,tenantId),d=(await tx.get(r.attempt)).data();
   if(t.plan!=='trial'||t.billing||t.stripeSandboxCheckoutId!==a.attemptId||d?.attemptId!==a.attemptId||!['creating','open'].includes(d.status)||d.sessionId&&d.sessionId!==session.id)fail('CHECKOUT_CONFLICT');
   tx.update(r.attempt,{sessionId:session.id,status:'open',updatedAt:stamp()});
  });
  return {status:'open',sessionId:session.id,url,testMode:true};
 }
 async function finalizeCheckout({token,tenantId}){
  const uid=await uidFor(token),r=refs(tenantId,uid);
  const a=await db.runTransaction(async tx=>{const t=await owner(tx,r,uid,tenantId),d=(await tx.get(r.attempt)).data();if(!d||d.ownerUid!==uid||t.stripeSandboxCheckoutId!==d.attemptId||!id(d.sessionId,'cs_test_'))fail('CHECKOUT_CONFLICT');return d;});
  // No client-provided subscription/customer/payment status. Retrieve stored session.
  const s=await provider.session(a.sessionId),sub=s?.subscription,p=SUBSCRIPTION_PLANS[a.tier],line=s?.line_items?.data?.[0];
  if(s?.livemode!==false||s.id!==a.sessionId||s.mode!=='subscription'||s.status!=='complete'||s.payment_status!=='paid'||s.client_reference_id!==a.attemptId||s.metadata?.tenantId!==tenantId||s.metadata?.attemptId!==a.attemptId||s.metadata?.tier!==a.tier||
   s.currency!=='eur'||s.amount_total!==p.priceCents||s.line_items?.has_more!==false||s.line_items?.data?.length!==1||line.quantity!==1||line.amount_total!==p.priceCents||
   !id(s.customer,'cus_')||sub?.livemode!==false||sub.status!=='active'||!id(sub.id,'sub_')||sub.customer!==s.customer||sub.metadata?.tenantId!==tenantId||sub.metadata?.attemptId!==a.attemptId||sub.metadata?.tier!==a.tier||sub.items?.has_more!==false||sub.items?.data?.length!==1||sub.items.data[0].quantity!==1)fail('CHECKOUT_NOT_VERIFIED');
  checkedPrice(line.price,a.tier,a.priceId);checkedPrice(sub.items.data[0].price,a.tier,a.priceId);
  const binding=db.doc('stripeTestSubscriptions/'+sub.id);
  return db.runTransaction(async tx=>{
   const t=await owner(tx,r,uid,tenantId),[as,bs]=await Promise.all([tx.get(r.attempt),tx.get(binding)]),d=as.data(),b=bs.data();
   if(b&&d?.status==='bound'&&t.stripeTestSubscriptionId===sub.id&&b.tenantId===tenantId&&b.customerId===s.customer&&b.tier===a.tier)return {status:'already-bound',subscriptionId:sub.id,testMode:true};
   if(b||t.plan!=='trial'||t.billing||t.stripeTestSubscriptionId||t.stripeSandboxCheckoutId!==a.attemptId||d?.sessionId!==s.id||d.status!=='open'||t.requestedTier!==a.tier||t.upgradeRequested!==true||!fits(t.usage,p.limits))fail('CHECKOUT_CONFLICT');
   tx.create(binding,{tenantId,customerId:s.customer,tier:a.tier,subscriptionId:sub.id,status:'pending',createdAt:stamp()});
   tx.update(r.tenant,{stripeTestSubscriptionId:sub.id});
   tx.update(r.attempt,{status:'bound',subscriptionId:sub.id,boundAt:stamp()});
   // Binding does NOT grant access. A separately verified paid invoice is required.
   return {status:'bound-awaiting-invoice',subscriptionId:sub.id,testMode:true};
  });
 }
 async function companies({token}){
  const uid=await uidFor(token);
  const memberships=await db.collection('users/'+uid+'/memberships').limit(100).get();
  const result=[];
  for(const item of memberships.docs){
   if(!tenantOk(item.id)||item.data().role!=='admin'||item.data().status!=='active')continue;
   try{
    const r=refs(item.id,uid);
    const v=await db.runTransaction(async tx=>{const t=await owner(tx,r,uid,item.id),a=(await tx.get(r.attempt)).data();
     return {id:item.id,name:String(t.name||item.id).slice(0,120),plan:t.plan,requestedTier:t.requestedTier||null,checkoutStatus:a?.status||null,billingTier:t.billing?.tier||null,paidThrough:t.billing?.periodStart?.toMillis?.()+t.billing?.periods*30*86400000||null,testMode:true};});
    result.push(v);
   }catch(e){if(e.code!=='FORBIDDEN')throw e;}
  }
  return {companies:result,testMode:true};
 }
 async function requestPlan({token,tenantId,tier}){
  const uid=await uidFor(token),r=refs(tenantId,uid);
  if(!Object.hasOwn(SUBSCRIPTION_PLANS,tier))fail('INVALID_REQUEST');
  return db.runTransaction(async tx=>{
   const t=await owner(tx,r,uid,tenantId),a=await tx.get(r.attempt);
   if(t.plan!=='trial'||t.billing||t.stripeTestSubscriptionId||t.stripeSandboxCheckoutId||a.exists)fail('BILLING_CONFLICT');
   if(!fits(t.usage,SUBSCRIPTION_PLANS[tier].limits))fail('USAGE_EXCEEDS_PLAN');
   tx.update(r.tenant,{requestedTier:tier,upgradeRequested:true,upgradeRequestedAt:stamp()});
   return {status:'requested',tier,testMode:true};
  });
 }
 // Only release a known, verified, terminal unpaid session. Never clear a
 // creating attempt that may have an unknown live session at the provider.
 async function cancelCheckout({token,tenantId}){
  const uid=await uidFor(token),r=refs(tenantId,uid);
  if(typeof provider.expire!=='function')fail('CANCEL_UNAVAILABLE');
  const a=await db.runTransaction(async tx=>{
   const t=await owner(tx,r,uid,tenantId),d=(await tx.get(r.attempt)).data();
   if(!d&&!t.stripeSandboxCheckoutId&&!t.stripeTestSubscriptionId&&t.plan==='trial'&&!t.billing)return null;
   if(t.plan!=='trial'||t.billing||t.stripeTestSubscriptionId||!d||d.ownerUid!==uid||t.stripeSandboxCheckoutId!==d.attemptId||d.status!=='open'||!id(d.sessionId,'cs_test_'))fail('CHECKOUT_RECONCILIATION_REQUIRED');
   return d;
  });
  if(!a)return {status:'already-cleared',testMode:true};
  const matches=s=>s?.livemode===false&&s.id===a.sessionId&&s.mode==='subscription'&&s.client_reference_id===a.attemptId&&s.metadata?.tenantId===tenantId&&s.metadata?.attemptId===a.attemptId&&s.metadata?.tier===a.tier;
  const unpaid=s=>s.payment_status==='unpaid'&&s.subscription===null;
  let session=await provider.session(a.sessionId);
  if(!matches(session)||!unpaid(session)||!['open','expired'].includes(session.status))fail('CHECKOUT_RECONCILIATION_REQUIRED');
  if(session.status==='open'){
   // An expire request can race payment or return an ambiguous error. A new
   // authoritative read is mandatory; a failed call alone never releases it.
   try{await provider.expire(a.sessionId);}catch{}
   session=await provider.session(a.sessionId);
  }
  if(!matches(session)||!unpaid(session)||session.status!=='expired')fail('CHECKOUT_RECONCILIATION_REQUIRED');
  return db.runTransaction(async tx=>{
   const t=await owner(tx,r,uid,tenantId),d=(await tx.get(r.attempt)).data();
   if(!d&&!t.stripeSandboxCheckoutId&&!t.stripeTestSubscriptionId&&t.plan==='trial'&&!t.billing)return {status:'already-cleared',testMode:true};
   if(t.plan!=='trial'||t.billing||t.stripeTestSubscriptionId||t.stripeSandboxCheckoutId!==a.attemptId||d?.attemptId!==a.attemptId||d.ownerUid!==uid||d.sessionId!==a.sessionId||d.status!=='open')fail('CHECKOUT_CONFLICT');
   const audit=db.doc('stripeSandboxExpiredAttempts/'+a.attemptId);
   tx.create(audit,{tenantId,ownerUid:uid,tier:a.tier,sessionId:a.sessionId,status:'expired-unpaid',recordedAt:stamp(),testMode:true});
   tx.delete(r.attempt);tx.update(r.tenant,{stripeSandboxCheckoutId:FieldValue.delete(),requestedTier:null,upgradeRequested:false});
   return {status:'cleared',testMode:true};
  });
 }
 return Object.freeze({createCheckout,finalizeCheckout,companies,requestPlan,cancelCheckout});
}
