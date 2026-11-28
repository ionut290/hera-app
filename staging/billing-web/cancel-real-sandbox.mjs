// Cancel only the paid synthetic subscription created by a specified Stripe Sandbox workflow run.
// Refuses live-mode, unpaid, mismatched-price, or unrelated sessions.
import assert from 'node:assert/strict';
import {readFileSync} from 'node:fs';
import {join} from 'node:path';
import Stripe from 'stripe';
import {SANDBOX_PRICE_IDS} from './stripe-sandbox-catalog.mjs';
const secret=process.env.VARGA_STRIPE_SANDBOX_CHECKOUT_KEY;
assert.match(secret||'',/^rk_test_[A-Za-z0-9]+$/,'Sandbox restricted test key required');
const state=JSON.parse(readFileSync(join(process.cwd(),'.sandbox-state','state.json'),'utf8'));
assert.match(state.sessionId||'',/^cs_test_[A-Za-z0-9]+$/,'Test Checkout Session required');
assert.match(state.tenantId||'',/^demo-[a-z0-9-]{3,59}$/,'Synthetic tenant required');
const stripe=new Stripe(secret,{apiVersion:'2024-10-28.acacia',maxNetworkRetries:2,timeout:15000});
const session=await stripe.checkout.sessions.retrieve(state.sessionId,{expand:['subscription','line_items.data.price']});
assert.equal(session.livemode,false,'Refusing a live Checkout Session');
assert.equal(session.id,state.sessionId);
assert.equal(session.mode,'subscription');
assert.equal(session.status,'complete');
assert.equal(session.payment_status,'paid');
assert.equal(session.currency,'eur');
assert.equal(session.amount_total,1990);
assert.equal(session.metadata?.tenantId,state.tenantId);
assert.equal(session.metadata?.tier,'starter');
assert.ok(session.metadata?.attemptId);
const line=session.line_items?.data?.[0];
assert.equal(session.line_items?.data?.length,1);
assert.equal(line?.price?.id,SANDBOX_PRICE_IDS.starter);
assert.equal(line?.price?.livemode,false);
const sub=session.subscription;
assert.ok(sub&&typeof sub==='object','Expanded subscription required');
assert.equal(sub.livemode,false,'Refusing a live subscription');
assert.equal(sub.customer,session.customer);
assert.equal(sub.metadata?.tenantId,state.tenantId);
assert.equal(sub.metadata?.tier,'starter');
assert.equal(sub.metadata?.attemptId,session.metadata.attemptId);
if(sub.status==='canceled'){
 console.log('SANDBOX TEST SUBSCRIPTION WAS ALREADY CANCELED.');
}else{
 assert.equal(sub.status,'active','Refusing to change a subscription outside active test state');
 const canceled=await stripe.subscriptions.cancel(sub.id,{}, {idempotencyKey:'varga-sandbox-cleanup-'+sub.id});
 assert.equal(canceled.livemode,false);
 assert.equal(canceled.id,sub.id);
 assert.equal(canceled.status,'canceled');
 console.log('SANDBOX TEST SUBSCRIPTION CANCELED; no live transaction.');
}
