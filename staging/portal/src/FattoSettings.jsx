import {useState} from 'react';
import {validateFattoPolicy} from '../../fatto-policy.mjs';
const labels={saveCompletion:'Salva intervento',moveToCompleted:'Sposta nei FATTI',openWhatsApp:'Apri WhatsApp',exportSheet:'Aggiorna Google Sheets'};
export default function FattoSettings({tenant,context,initialPolicy,onSave,preview=true}){
 const [actions,setActions]=useState(()=>[...initialPolicy.actions]);
 const [version,setVersion]=useState(initialPolicy.version);
 const [busy,setBusy]=useState(false);
 const [message,setMessage]=useState('');
 if(context.role!=='admin'||tenant.demo!==true||tenant.fattoMode!=='configurable-demo'||context.tenantId!==tenant.id)return null;
 const available=a=>a!=='openWhatsApp'||tenant.whatsappConfigured===true;
 const enabled=a=>available(a)&&(a!=='exportSheet'||tenant.sheetConfigured===true);
 function move(i,d){const j=i+d;if(i<1||j<1||j>=actions.length)return;setActions(old=>{const next=[...old];[next[i],next[j]]=[next[j],next[i]];return next;});setMessage('');}
 async function save(){
 if(busy)return;
 setBusy(true);setMessage('');
 try{
 const next=validateFattoPolicy({tenant,context,policy:{version:version+1,actions}});
 if(!preview){if(typeof onSave!=='function')throw Error('Missing persistence');await onSave({actions:[...next.actions],expectedVersion:version});}
 setVersion(next.version);setMessage(preview?'Configurazione verificata localmente. Nessun dato salvato.':'Configurazione salvata.');
 }catch{setMessage('Configurazione non valida o modificata. Ricaricare e riprovare.');}finally{setBusy(false);}
 }
 return <section className="members-panel" aria-label="Impostazioni FATTO">
 <h2>Impostazioni FATTO</h2><p>{preview?'Anteprima locale: nessun salvataggio o azione reale.':tenant.name}</p>
 {Object.entries(labels).map(([a,label])=><div key={a}><label><input type="checkbox" checked={actions.includes(a)} disabled={busy||a==='saveCompletion'||!enabled(a)} onChange={e=>{setActions(old=>e.target.checked?[...old,a]:old.filter(x=>x!==a));setMessage('');}}/>{label}</label>{!enabled(a)&&<small> — Integrazione non configurata</small>}</div>)}
 <h3>Ordine di esecuzione</h3><ol>{actions.map((a,i)=><li key={a}>{labels[a]} {i>0&&<><button type="button" disabled={busy||i===1} onClick={()=>move(i,-1)} aria-label={'Sposta su '+labels[a]}>↑</button><button type="button" disabled={busy||i===actions.length-1} onClick={()=>move(i,1)} aria-label={'Sposta giù '+labels[a]}>↓</button></>}</li>)}</ol>
 <button type="button" disabled={busy} onClick={save}>{preview?'Verifica configurazione':'Salva configurazione'}</button>
 <p role="status">{message}</p></section>;
}
