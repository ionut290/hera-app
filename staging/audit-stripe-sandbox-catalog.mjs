// Read-only catalog verification. No Checkout, customer or subscription creation.
import Stripe from 'stripe';
import {auditSandboxCatalog} from './stripe-sandbox-catalog.mjs';
const key=process.env.VARGA_STRIPE_SANDBOX_READ_KEY;
if(!key||!/^rk_test_[A-Za-z0-9]+$/.test(key)){
 console.error('A restricted Stripe test key with Prices read permission is required.');
 process.exit(1);
}
try{
 const stripe=new Stripe(key,{apiVersion:'2024-10-28.acacia',maxNetworkRetries:1,timeout:15000});
 const result=await auditSandboxCatalog({price:id=>stripe.prices.retrieve(id)});
 console.log(JSON.stringify(result,null,2));
 if(result.some(r=>r.warning))console.warn('STRIPE SANDBOX CATALOG WARNING: lookup aliases differ; exact price IDs are authoritative and all economic checks remain mandatory.');
 console.log('STRIPE SANDBOX CATALOG PASS: all three configured prices are active, test-mode, EUR, exact amount and 30-day recurrence.');
}catch(e){
 // Never print provider request/response, authorization headers or credentials.
 if(Array.isArray(e?.results))console.error(JSON.stringify(e.results,null,2));
 console.error('STRIPE SANDBOX CATALOG FAILED:', ['PRICE_MISMATCH','LOOKUP_KEY_MISMATCH'].includes(e?.code)?e.code:'PROVIDER_READ_FAILED');
 process.exitCode=1;
}
