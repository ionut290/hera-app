// Proposed commercial catalog, EUR amounts in cents. No payment provider calls.
export const BILLING_PERIOD_DAYS=30;
const periodMs=BILLING_PERIOD_DAYS*86400000;
function tier(name,priceCents,commesse,impianti,admin,operatori,mezzi,squadre,storageBytes){
 return Object.freeze({name,priceCents,limits:Object.freeze({commesse,impianti,admin,operatori,personale:operatori,mezzi,squadre,storageBytes})});
}
export const SUBSCRIPTION_PLANS=Object.freeze({
 starter:tier('Starter',1990,10,300,1,15,15,5,1000000000),
 professional:tier('Professional',3990,50,2000,3,50,50,20,5000000000),
 business:tier('Business',7990,200,10000,10,200,200,100,20000000000)
});
export function periodEnd(tenant){
 if(tenant?.plan==='trial'){
  const start=tenant.trialStartedAt?.toMillis?.();
  return Number.isFinite(start)&&tenant.trialDays===30?start+periodMs:null;
 }
 if(tenant?.plan==='full'&&tenant.billing){
  const b=tenant.billing,start=b.periodStart?.toMillis?.();
  return Number.isFinite(start)&&Number.isSafeInteger(b.periods)&&b.periods>0?start+b.periods*periodMs:null;
 }
 return null;
}
export function subscriptionLabel(tenant,now=Date.now()){
 const end=periodEnd(tenant);
 if(tenant?.plan==='trial'||tenant?.billing){
  const name=tenant.plan==='trial'?'Piano prova':SUBSCRIPTION_PLANS[tenant.billing?.tier]?.name||'Abbonamento';
  if(end===null||end<=now)return name+' scaduto · consultazione';
  const days=Math.ceil((end-now)/86400000);
  return name+' · '+days+' giorni rimanenti'+(days<=7?' · scadenza vicina':'');
 }
 return tenant?.plan==='full'?'Piano completo precedente':'Azienda';
}
export function formatPrice(cents){return new Intl.NumberFormat('it-IT',{style:'currency',currency:'EUR'}).format(cents/100);}
