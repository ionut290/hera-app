import assert from 'node:assert/strict';
import {initializeApp,deleteApp} from 'firebase-admin/app';
import {getAuth} from 'firebase-admin/auth';
import {getFirestore} from 'firebase-admin/firestore';
import {createSandboxCheckout,createSandboxStripeProvider} from './stripe-sandbox-checkout.mjs';
import {TRIAL_LIMITS,initialTrialUsage} from './trial-plan.mjs';
const app=initializeApp({projectId:'varga-cantieri-v2-staging'}),db=getFirestore(app),auth=getAuth(app);
const A='demo-checkout-a',B='demo-checkout-b',priceIds={starter:'price_fixture_starter',professional:'price_fixture_professional',business:'price_fixture_business'};
const price={id:priceIds.professional,livemode:false,active:true,currency:'eur',unit_amount:3990,type:'recurring',recurring:{interval:'day',interval_count:30,usage_type:'licensed'},billing_scheme:'per_unit'};
const err=(code,fn)=>assert.rejects(fn,e=>e.code===code);
let session,createCalls=0,keys=[],priceResponse=price,unpaid=true;
const provider={
 price:async()=>priceResponse,
 create:async(params,key)=>{
  keys.push(key);createCalls++;
  assert.equal(params.mode,'subscription');assert.deepEqual(params.line_items,[{price:price.id,quantity:1}]);assert.equal(params.billing_address_collection,'required');assert.equal(params.client_reference_id,params.metadata.attemptId);
  assert.deepEqual(params.metadata,params.subscription_data.metadata);
  if(!session)session={id:'cs_test_fixture_a',livemode:false,mode:'subscription',status:'open',url:'https://checkout.stripe.com/c/pay/cs_test_fixture_a',client_reference_id:params.client_reference_id,metadata:params.metadata};
  return structuredClone(session);
 },
 session:async()=>structuredClone(session)
};
const service=createSandboxCheckout({app,provider,priceIds});
async function user(uid){
 const email=uid+'@example.test',password='Only-Local-Test-42';
 await auth.createUser({uid,email,password,emailVerified:true});
 const res=await fetch('http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=fake-api-key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password,returnSecureToken:true})});
 assert.equal(res.ok,true);return (await res.json()).idToken;
}
async function member(t,uid,role){
 await db.doc('tenants/'+t+'/members/'+uid).set({role,status:'active'});
 await db.doc('users/'+uid+'/memberships/'+t).set({tenantId:t,role,status:'active'});
}
try{
 assert.throws(()=>createSandboxStripeProvider({app,secretKey:'sk_live_not_allowed'}),e=>e.code==='SANDBOX_DISABLED');
 const [owner,peer,other]=await Promise.all(['checkout-owner','checkout-peer','checkout-other'].map(user));
 const trial={demo:true,status:'active',plan:'trial',trial:true,trialDays:30,trialStartedAt:new Date(),limits:{...TRIAL_LIMITS},usage:initialTrialUsage(),requestedTier:'professional',upgradeRequested:true};
 await db.doc('tenants/'+A).set({...trial,ownerUid:'checkout-owner'});await db.doc('tenants/'+B).set({...trial,ownerUid:'checkout-other'});
 await Promise.all([member(A,'checkout-owner','admin'),member(A,'checkout-peer','admin'),member(B,'checkout-other','admin')]);
 const request={token:owner,tenantId:A,tier:'professional',acceptTerms:true};
 await err('INVALID_REQUEST',()=>service.createCheckout({...request,acceptTerms:false}));
 await err('FORBIDDEN',()=>service.createCheckout({...request,token:peer}));
 await err('FORBIDDEN',()=>service.createCheckout({...request,token:other}));
 await err('BILLING_CONFLICT',()=>service.createCheckout({...request,tier:'starter'}));
 priceResponse={...price,unit_amount:1};
 await err('PRICE_MISMATCH',()=>service.createCheckout(request));assert.equal(createCalls,0);
 priceResponse=price;
 const pair=await Promise.all([service.createCheckout(request),service.createCheckout(request)]);
 assert.equal(pair[0].sessionId,pair[1].sessionId);assert.equal(new Set(keys).size,1);assert.equal(pair[0].testMode,true);
 const calls=createCalls;assert.deepEqual(await service.createCheckout(request),pair[0]);assert.equal(createCalls,calls);
 await err('CHECKOUT_NOT_VERIFIED',()=>service.finalizeCheckout({token:owner,tenantId:A}));
 assert.equal((await db.doc('tenants/'+A).get()).data().plan,'trial');
 session={...session,status:'complete',payment_status:'paid',currency:'eur',amount_total:3990,customer:'cus_fixture_checkout',line_items:{has_more:false,data:[{quantity:1,amount_total:3990,price}]},subscription:{id:'sub_fixture_checkout',livemode:false,status:'active',customer:'cus_fixture_checkout',metadata:session.metadata,items:{has_more:false,data:[{quantity:1,price}]}}};
 session.payment_status='unpaid';await err('CHECKOUT_NOT_VERIFIED',()=>service.finalizeCheckout({token:owner,tenantId:A}));session.payment_status='paid';
 await err('FORBIDDEN',()=>service.finalizeCheckout({token:peer,tenantId:A}));
 assert.equal((await service.finalizeCheckout({token:owner,tenantId:A})).status,'bound-awaiting-invoice');
 assert.equal((await service.finalizeCheckout({token:owner,tenantId:A})).status,'already-bound');
 const t=(await db.doc('tenants/'+A).get()).data();
 assert.equal(t.plan,'trial');assert.equal(t.billing,undefined);assert.equal(t.stripeTestSubscriptionId,'sub_fixture_checkout');
 assert.equal((await db.doc('stripeTestSubscriptions/sub_fixture_checkout').get()).data().tenantId,A);
 assert.equal((await db.doc('tenants/'+B).get()).data().stripeTestSubscriptionId,undefined);
 console.log('STRIPE CHECKOUT PASS: verified owner, exact provider price, stable idempotency, same-company binding, unpaid and foreign access denied, no entitlement on checkout return.');
}finally{await deleteApp(app);}
