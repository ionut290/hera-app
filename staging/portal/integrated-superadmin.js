// Super Admin UI within the integrated app, not a replacement for the original company UI.
// Only verified claims + restrictive server rules grant privileges. Live staging writes remain disabled.
export function createAdminPanel({host,auth,db,statics,readOnly=false}){
 let generation=0,busy=false,uid=null;
 const clear=()=>{generation++;uid=null;busy=false;host.hidden=true;host.replaceChildren();};
 const el=(tag,text)=>{const n=document.createElement(tag);if(text!==undefined)n.textContent=text;return n;};
 const current=(g,u)=>generation===g&&uid===u&&auth.currentUser?.uid===u;
 async function authorize(g,u){
  if(!current(g,u))throw Error('Sessione amministrativa chiusa');
  const c=(await auth.currentUser.getIdTokenResult(true)).claims;
  if(!current(g,u)||c.email_verified!==true||c.apple_review===true||c.superAdmin!==true)throw Error('Super Admin non autorizzato');
 }
 async function open(user){
  clear();uid=user.uid;const g=generation,u=uid;host.hidden=false;
  const title=el('h2','Super Admin'),feedback=el('p','Caricamento aziende…'),list=el('div'),memberList=el('div');
  feedback.setAttribute('role','status');feedback.id='admin-status';host.append(title,feedback,list,memberList);
  const button=(name,action)=>{const b=el('button',name);b.type='button';b.onclick=action;return b;};
  async function mutate(action,refresh){
   if(readOnly||busy||!current(g,u))return;
   busy=true;
   try{await authorize(g,u);await action();if(current(g,u)){await refresh();feedback.textContent='Modifica salvata.';}}
   catch(e){if(current(g,u))feedback.textContent='Operazione non riuscita: '+(e.code||e.message);}
   finally{if(current(g,u))busy=false;}
  }
  async function showMembers(t){
   memberList.replaceChildren(el('h3','Utenti — '+t.name));
   try{
    await authorize(g,u);const rows=await db.collection('tenants').doc(t.id).collection('members').get({source:'server'});
    if(!current(g,u))return;
    for(const item of rows.docs){
     const m=item.data(),r=el('div');r.dataset.member=item.id;r.append(el('strong',m.displayName||item.id),el('p',m.role+' · '+m.status));
     if(!readOnly&&t.demo===true&&m.demo===true&&item.id!==u){
      const select=el('select');select.setAttribute('aria-label','Ruolo aziendale');
      for(const role of ['admin','supervisor','employee']){const o=el('option',role);o.value=role;select.append(o);}select.value=m.role;
      const patch=changes=>{const batch=db.batch();batch.update(item.ref,changes);batch.update(db.doc('users/'+item.id+'/memberships/'+t.id),changes);return batch.commit();};
      if(!['trial','full'].includes(t.plan))r.append(select,button('Salva ruolo',()=>{if(select.value!==m.role)mutate(()=>patch({role:select.value}),()=>showMembers(t));}));else r.append(el('p','Modifica ruolo non disponibile per aziende con quote.'));
      r.append(button(m.status==='active'?'Sospendi utente':'Riattiva utente',()=>mutate(()=>patch({status:m.status==='active'?'suspended':'active'}),()=>showMembers(t))));
     }
     memberList.append(r);
    }
   }catch(e){if(current(g,u))feedback.textContent='Utenti non disponibili: '+(e.code||e.message);}
  }
  async function refresh(){
   await authorize(g,u);const rows=await db.collection('tenants').get({source:'server'});
   if(!current(g,u))return;
   list.replaceChildren();memberList.replaceChildren();
   if(!readOnly){
    const name=el('input');name.placeholder='Nome nuova azienda demo';name.id='admin-new-company-name';name.maxLength=120;
    list.append(name,button('Nuova azienda',()=>{const n=name.value.trim();if(n.length<2||n.length>120){feedback.textContent='Nome azienda non valido.';return;}mutate(()=>db.collection('tenants').doc('demo-'+crypto.randomUUID()).set({name:n,status:'active',demo:true,createdAt:statics.FieldValue.serverTimestamp()}),refresh);}));
   }
   for(const item of rows.docs){
    const t={...item.data(),id:item.id},r=el('div');r.dataset.adminTenant=t.id;r.append(el('strong',t.name||t.id),el('p',t.status),button('Utenti',()=>showMembers(t)));
    if(['trial','full'].includes(t.plan)){
     r.append(el('p',(t.plan==='trial'?'Piano prova · 30 giorni':'Piano completo')+(t.upgradeRequested?' · Richiesta attivazione ricevuta':'')));
     const fields=[['commesse','Commesse'],['impianti','Impianti'],['operatori','Operatori'],['personale','Anagrafica personale'],['mezzi','Mezzi'],['squadre','Squadre'],['storageBytes','Archiviazione (byte)']];
     const inputs=new Map();
     for(const [key,label] of fields){
      r.append(el('p',label+': '+(t.usage?.[key]??'non verificato')+' / '+(t.limits?.[key]??'non configurato')));
      if(!readOnly&&t.demo===true&&Number.isSafeInteger(t.usage?.[key])&&Number.isSafeInteger(t.limits?.[key])){
       const input=el('input');input.type='number';input.min=String(Math.max(['commesse','impianti'].includes(key)?1:0,t.usage[key]));input.step='1';input.value=String(t.limits[key]);input.setAttribute('aria-label',label+' limite');inputs.set(key,input);r.append(input);
      }
     }
     if(inputs.size===fields.length&&t.usage?.admin===1){
      r.append(button(t.plan==='trial'?'Attiva piano completo':'Salva limiti piano completo',()=>{
       const limits={admin:1};for(const [key,input] of inputs){const n=Number(input.value);if(!input.value.trim()||!Number.isSafeInteger(n)||n<Number(input.min)){feedback.textContent='Limite non valido: '+key;return;}limits[key]=n;}
       mutate(()=>item.ref.update({plan:'full',trial:false,limits,upgradeRequested:false,planUpdatedAt:statics.FieldValue.serverTimestamp(),planUpdatedBy:u}),refresh);
      }));
     }
    }
    if(!readOnly&&t.demo===true){
     const name=el('input');name.value=t.name;name.maxLength=120;name.setAttribute('aria-label','Nome azienda');
     r.append(name,button('Rinomina',()=>{const n=name.value.trim();if(n.length>=2&&n.length<=120&&n!==t.name)mutate(()=>item.ref.update({name:n}),refresh);}),button(t.status==='active'?'Sospendi azienda':'Riattiva azienda',()=>mutate(()=>item.ref.update({status:t.status==='active'?'suspended':'active'}),refresh)));
    }
    list.append(r);
   }
   feedback.textContent=readOnly?'Consultazione amministrativa staging. Modifiche disabilitate.':rows.size+' aziende registrate.';
  }
  try{await refresh();}catch(e){if(current(g,u))feedback.textContent='Dashboard non disponibile: '+(e.code||e.message);}
 }
 return Object.freeze({open,clear});
}
