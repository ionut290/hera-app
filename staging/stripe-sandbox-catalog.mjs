// Non-secret identifiers supplied by the owner from the Stripe Sandbox dashboard.
// Never use this catalog in production. Provider verification is still required.
import {SUBSCRIPTION_PLANS,BILLING_PERIOD_DAYS} from './subscription-plans.mjs';
export const SANDBOX_PRICE_IDS=Object.freeze({
 starter:'price_1UggHrL5SpVe43ZEAiAoaaUR',
 professional:'price_1UggLoL5SpVe43ZE4FrvAT3R',
 business:'price_1UggMzL5SpVe43ZESQxPdrxU'
});
export const SANDBOX_LOOKUP_KEYS=Object.freeze({
 starter:'varga_starter_30d',
 professional:'varga_professional_30d',
 business:'varga_business_30d'
});
export function verifySandboxPrice(price,tier,expected=SANDBOX_PRICE_IDS[tier]){
 if(!Object.hasOwn(SUBSCRIPTION_PLANS,tier)||!price||price.id!==expected||price.livemode!==false||price.active!==true||price.currency!=='eur'||price.unit_amount!==SUBSCRIPTION_PLANS[tier].priceCents||price.type!=='recurring'||price.recurring?.interval!=='day'||price.recurring?.interval_count!==BILLING_PERIOD_DAYS||price.recurring?.usage_type!=='licensed'||price.billing_scheme!=='per_unit')throw Object.assign(Error('PRICE_MISMATCH'),{code:'PRICE_MISMATCH'});
 return true;
}
export async function inspectSandboxCatalog(provider){
 const results=[];
 for(const tier of Object.keys(SUBSCRIPTION_PLANS)){
  try{
   const price=await provider.price(SANDBOX_PRICE_IDS[tier]);
   let problem=null;
   try{verifySandboxPrice(price,tier);}catch{problem='PRICE_MISMATCH';}
   // Checkout uses the immutable configured price ID, never a lookup key.
   const warning=price?.lookup_key===SANDBOX_LOOKUP_KEYS[tier]?null:'LOOKUP_KEY_MISMATCH';
   results.push({tier,amountCents:price?.unit_amount??null,currency:price?.currency??null,interval:price?.recurring?.interval??null,intervalCount:price?.recurring?.interval_count??null,testMode:price?.livemode===false,active:price?.active===true,lookupKeyMatches:price?.lookup_key===SANDBOX_LOOKUP_KEYS[tier],lookupKeyPresent:Boolean(price?.lookup_key),expectedLookupKey:SANDBOX_LOOKUP_KEYS[tier],warning,problem});
  }catch{results.push({tier,problem:'PROVIDER_READ_FAILED'});}
 }
 return results;
}
export async function auditSandboxCatalog(provider){
 const results=await inspectSandboxCatalog(provider),first=results.find(r=>r.problem);
 if(first)throw Object.assign(Error(first.problem),{code:first.problem,results});
 return results;
}
