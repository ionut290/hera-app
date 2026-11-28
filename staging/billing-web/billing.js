'use strict';
const $=id=>document.getElementById(id);
let token=null,companies=[],plans=[],busy=false,epoch=0;
const selected=()=>companies.find(c=>c.id===$('company').value);
const tier=()=>document.querySelector('input[name=tier]:checked')?.value;
const money=n=>new Intl.NumberFormat('it-IT',{style:'currency',currency:'EUR'}).format(n/100);
const errors={UNAUTHENTICATED:'Accedi con un account di prova verificato.',FORBIDDEN:'Accesso riservato al titolare dell’azienda.',BILLING_CONFLICT:'Esiste già una richiesta o un abbonamento: aggiorna lo stato.',CHECKOUT_CONFLICT:'Checkout già presente o non disponibile.',CHECKOUT_RECONCILIATION_REQUIRED:'Tentativo da verificare. Non avviare un nuovo pagamento.',USAGE_EXCEEDS_PLAN:'Il piano non contiene le risorse già utilizzate.',INVALID_REQUEST:'Seleziona un piano e conferma le condizioni di prova.'};
function message(v){$('status').textContent=v;}
function node(tag,text,cls){const n=document.createElement(tag);if(text!==undefined)n.textContent=text;if(cls)n.className=cls;return n;}
async function api(action,data={}){
 const generation=epoch;if(!token)throw Error('UNAUTHENTICATED');
 const res=await fetch('/api/'+action,{method:'POST',headers:{'Content-Type':'application/json','Authorization':'Bearer '+token},body:JSON.stringify(data),cache:'no-store'});
 const out=await res.json();if(generation!==epoch)throw Error('SESSION_CHANGED');if(!res.ok)throw Error(out.code||'REQUEST_FAILED');return out;
}
function state(){
 const c=selected(),t=tier(),locked=!!c?.checkoutStatus,active=c?.plan==='full';
 $('company').disabled=busy;$('terms').disabled=busy||!c||active||locked;
 $('checkout').disabled=busy||!c||active||!t||(!locked&&!$('terms').checked)||locked&&!['creating','open'].includes(c.checkoutStatus);
 $('checkout').textContent=locked?'Riprendi checkout di prova':'Avvia checkout di prova';
 $('checkout-panel').hidden=!c||active||!['open','bound'].includes(c.checkoutStatus);
 $('simulate').disabled=busy||c?.checkoutStatus!=='open';$('verify').disabled=busy||!c;$('cancel').disabled=busy||c?.checkoutStatus!=='open';
 for(const n of document.querySelectorAll('input[name=tier]'))n.disabled=busy||!c||active||locked;
}
function drawPlans(){
 const root=$('plans');root.replaceChildren();
 for(const p of plans){
  const card=node('article',undefined,'plan'),label=node('label'),radio=node('input');radio.type='radio';radio.name='tier';radio.value=p.tier;radio.addEventListener('change',state);
  label.append(radio,node('strong',p.name));card.append(label,node('p',money(p.priceCents),'price'),node('p','Ogni 30 giorni · rinnovo simulato','subtle'));
  const ul=node('ul');for(const text of [p.limits.commesse+' commesse',p.limits.impianti+' impianti',p.limits.admin+' amministratori',p.limits.operatori+' operatori',p.limits.mezzi+' mezzi',p.limits.squadre+' squadre',p.limits.storageBytes/1e9+' GB di archiviazione'])ul.append(node('li',text));card.append(ul);root.append(card);
 }
}
function drawCompany(){
 const c=selected();$('terms').checked=false;
 for(const n of document.querySelectorAll('input[name=tier]'))n.checked=n.value===(c?.billingTier||c?.requestedTier||'starter');
 $('subscription-status').textContent=!c?'Nessuna azienda attiva di cui sei titolare.':c.plan==='full'?(c.paidThrough>Date.now()?'Piano attivo (simulazione): ':'Piano scaduto (simulazione): ')+(plans.find(p=>p.tier===c.billingTier)?.name||c.billingTier)+' · scadenza '+(c.paidThrough?new Date(c.paidThrough).toLocaleDateString('it-IT'):'non disponibile'):c.checkoutStatus==='bound'?'Pagamento in verifica: attendi conferma della fattura.':c.checkoutStatus?'Checkout di prova in corso.':'Piano prova · scegli il tuo abbonamento.';
 $('checkout-summary').textContent=c?.requestedTier?'Piano '+(plans.find(p=>p.tier===c.requestedTier)?.name||c.requestedTier):'';
 state();
}
async function refresh(){
 const old=$('company').value;const out=await api('companies');companies=out.companies;const select=$('company');select.replaceChildren();
 for(const c of companies){const option=node('option',c.name);option.value=c.id;select.append(option);}if(companies.some(c=>c.id===old))select.value=old;drawCompany();
}
async function run(fn){
 if(busy)return;const generation=epoch;busy=true;state();message('Operazione in corso…');
 try{await fn();}catch(e){if(epoch===generation)message(errors[e.message]||'Operazione non riuscita. Riprova o contatta l’assistenza.');}
 finally{if(epoch===generation){busy=false;state();}}
}
$('login').addEventListener('submit',e=>{e.preventDefault();run(async()=>{
 const email=$('email').value.trim().toLowerCase();if(!email.endsWith('@example.test'))throw Error('UNAUTHENTICATED');
 const response=await fetch('http://127.0.0.1:9099/identitytoolkit.googleapis.com/v1/accounts:signInWithPassword?key=fake-api-key',{method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({email,password:$('password').value,returnSecureToken:true})});
 $('password').value='';const out=await response.json();if(!response.ok||!out.idToken)throw Error('UNAUTHENTICATED');
 token=out.idToken;
 try{plans=(await api('catalog')).plans;drawPlans();await refresh();$('login-panel').hidden=true;$('account').hidden=false;message('Accesso al collaudo effettuato.');}catch(e){token=null;throw e;}
});});
$('logout').onclick=()=>{epoch++;token=null;companies=[];plans=[];busy=false;$('account').hidden=true;$('login-panel').hidden=false;$('company').replaceChildren();$('plans').replaceChildren();$('password').value='';$('terms').checked=false;message('Sessione chiusa.');state();};
$('company').onchange=drawCompany;$('terms').onchange=state;
$('checkout').onclick=()=>run(async()=>{
 const c=selected(),t=tier();if(!c||!t)throw Error('INVALID_REQUEST');
 if(!c.checkoutStatus){if(!$('terms').checked)throw Error('INVALID_REQUEST');await api('request-plan',{tenantId:c.id,tier:t});}
 try{await api('checkout',{tenantId:c.id,tier:t,acceptTerms:true});}catch(e){try{await refresh();}catch{}throw e;}await refresh();message('Checkout simulato pronto. Nessun pagamento effettuato.');
});
$('simulate').onclick=()=>run(async()=>{const c=selected();if(!c)throw Error('FORBIDDEN');await api('simulate',{tenantId:c.id});await refresh();message('Pagamento di prova confermato. Nessun addebito reale.');});
$('cancel').onclick=()=>{const c=selected();if(!c||!confirm('Annullare il tentativo di pagamento non ancora pagato?'))return;run(async()=>{await api('cancel',{tenantId:c.id});await refresh();message('Checkout annullato. Puoi scegliere un nuovo piano.');});};
$('verify').onclick=()=>run(async()=>{await refresh();message('Stato aggiornato dal server.');});
state();
