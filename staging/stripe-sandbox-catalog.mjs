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
export async function auditSandboxCatalog(provider){
 const results=[];
 for(const tier of Object.keys(SUBSCRIPTION_PLANS)){
  const price=await provider.price(SANDBOX_PRICE_IDS[tier]);
  verifySandboxPrice(price,tier);
  if(price.lookup_key!==SANDBOX_LOOKUP_KEYS[tier])throw Object.assign(Error('LOOKUP_KEY_MISMATCH'),{code:'LOOKUP_KEY_MISMATCH'});
  results.push({tier,amountCents:price.unit_amount,currency:price.currency,interval:price.recurring.interval,intervalCount:price.recurring.interval_count,testMode:true});
 }
 return results;
}
