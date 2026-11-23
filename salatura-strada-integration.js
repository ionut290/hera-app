/* Salatura Strada 1/2/3 integra la PWA Piano Neve originale.
 * Non modifica i percorsi KML di Castel Maggiore e non duplica i giri.
 */
(function(){
"use strict";
const routes=window.VARGA_SALATURA_STRADA_ROUTES||[];
if(routes.length!==3)return;
const APP=new URL("./piano-neve/index.html",document.baseURI).href;
const $=id=>document.getElementById(id);
const admin=()=>typeof canManageData==="function"&&!!canManageData();
const logged=()=>typeof currentUser!=="undefined"&&!!currentUser;
const ready=()=>typeof db!=="undefined"&&!!db;
const url=i=>APP+"?vargaMajorRoute="+i;
const number=n=>Number(n).toLocaleString("it-IT");
let last=-1,started=false,busy=false;
function selected(){
  const page=$("impianti-page");
  if(!page||page.hidden||page.classList.contains("hidden")||document.body.classList.contains("snow-management-context"))return -1;
  if(typeof selectedCommessaId==="undefined"||typeof commesseById==="undefined")return -1;
  const c=commesseById.get(selectedCommessaId);
  if(!c)return -1;
  const code=String(c.codice||"").trim().toUpperCase();
  const name=String(c.nome||"").trim().toLowerCase();
  return routes.findIndex(r=>r.code===code||r.name.toLowerCase()===name);
}
function status(s){
  for(const k of ["salatura-import-status","salatura-panel-status"]){
    const el=$(k);if(el)el.textContent=s;
  }
}
async function seed(){
 if(busy||!admin()||!logged()||!ready())return;
 busy=true;const btn=$("salatura-import");if(btn)btn.disabled=true;
 try{
   const snap=await db.collection("commesse").get();
   const pending=[];let existing=0;
   for(const r of routes){
     const match=snap.docs.find(d=>String(d.data().codice||"").trim().toUpperCase()===r.code||
       String(d.data().nome||"").trim().toLowerCase()===r.name.toLowerCase());
     if(match){
       if(String(match.data().codice||"").toUpperCase()!==r.code)
         throw Error("Nome occupato da un'altra commessa: "+r.name);
       existing++;continue;
     }
     const ref=db.collection("commesse").doc("normal-salatura-strada-"+(r.index+1));
     if((await ref.get()).exists)throw Error("ID commessa già occupato: "+ref.id);
     pending.push({ref,record:{nome:r.name,codice:r.code,tipo:"main",
       pianoNeveService:"salatura-viabilita-maggiore",pianoNeveMajorRouteIndex:r.index,
       createdAt:firebase.firestore.FieldValue.serverTimestamp(),
       createdBy:currentUser.email||""}});
   }
   if(pending.length){
     const batch=db.batch();
     pending.forEach(item=>batch.set(item.ref,item.record));
     await batch.commit();
   }
   status("Salatura: "+pending.length+" nuove commesse create; "+existing+" già presenti.");
 }catch(e){status("Creazione salatura non riuscita: "+(e.code||e.message));}
 finally{busy=false;if(btn)btn.disabled=false;}
}
function adminButton(){
 const tools=$("panel-commesse")?.querySelector(".commesse-management-actions");
 if(!tools)return;
 let button=$("salatura-import");
 if(!button){
   const wrap=document.createElement("div");wrap.className="salatura-import-wrap";
   button=document.createElement("button");button.id="salatura-import";button.className="btn";
   button.type="button";button.textContent="🧂 Inserisci Salatura Strada 1, 2, 3";
   const msg=document.createElement("p");msg.className="muted";msg.id="salatura-import-status";
   msg.setAttribute("role","status");wrap.append(button,msg);tools.appendChild(wrap);
   button.addEventListener("click",seed);
 }
 button.hidden=!admin();
 if(admin()&&logged()&&ready()&&!started){started=true;seed();}
}
function fullScreen(){
 let frameBox=$("salatura-fullscreen");
 if(frameBox)return frameBox;
 frameBox=document.createElement("section");frameBox.id="salatura-fullscreen";
 frameBox.className="salatura-fullscreen";frameBox.hidden=true;
 frameBox.setAttribute("role","dialog");frameBox.setAttribute("aria-modal","true");
 frameBox.innerHTML='<header><button type="button" id="salatura-back">← Torna alla commessa</button>'+
 '<strong id="salatura-full-title">Piano Neve</strong>'+
 '<a id="salatura-full-link" rel="noopener noreferrer" target="_self">Apri pagina completa</a></header>'+
 '<iframe id="salatura-iframe" title="Varga Piano Neve" allow="geolocation; fullscreen; clipboard-write"></iframe>'+
 '<p>Accesso condiviso con Varga Cantieri. Se il GPS è limitato nel riquadro, apri la pagina completa.</p>';
 document.body.appendChild(frameBox);
 $("salatura-back").addEventListener("click",()=>{
   if(!window.confirm("Chiudere la vista? Metti in pausa o termina un eventuale giro attivo prima di uscire."))return;
   frameBox.hidden=true;document.body.classList.remove("salatura-fullscreen-open");
   $("salatura-iframe").removeAttribute("src");
 });
 return frameBox;
}
function ensure(){
 const page=$("impianti-page");if(!page)return null;
 let panel=$("salatura-strada-panel");if(panel)return panel;
 panel=document.createElement("section");panel.id="salatura-strada-panel";
 panel.className="card salatura-strada-panel hidden";panel.setAttribute("aria-hidden","true");
 panel.innerHTML='<h2 id="salatura-title">🧂 Salatura Strada</h2>'+
 '<p id="salatura-meta"></p>'+
 '<p>Il servizio mantiene la PWA Piano Neve originale: NAVIGA, GPS, percorso lavorato, INIZIA, PAUSA, RIPRENDI, TERMINA, sale caricato, rapportini, Excel e gestione amministratore.</p>'+
 '<div class="salatura-actions"><button type="button" id="salatura-inline" class="btn btn-primary">🧭 NAVIGA · Apri Piano Neve</button>'+
 '<a id="salatura-launch" class="btn" target="_self" rel="noopener noreferrer">🗺️ Apri pagina completa</a></div>'+
 '<p class="muted">Piano Neve è integrato in Varga Cantieri e usa il tuo accesso attuale, senza secondo login. I giri e i registri di salatura rimangono separati.</p>'+
 '<p id="salatura-panel-status" role="status"></p>'+
 '<details><summary id="salatura-street-count">Strade del percorso</summary><ol id="salatura-street-list"></ol></details>';
 const anchor=$("commessa-operational-card");
 if(anchor)anchor.insertAdjacentElement("afterend",panel);else page.appendChild(panel);
 $("salatura-inline").addEventListener("click",()=>{
   const i=selected();if(i<0)return;
   const box=fullScreen();
   $("salatura-full-title").textContent=routes[i].name+" · "+routes[i].sourceName;
   $("salatura-full-link").href=url(i);
   box.hidden=false;document.body.classList.add("salatura-fullscreen-open");
   $("salatura-iframe").src=url(i);
 });
 return panel;
}
function render(){
 adminButton();const panel=ensure();if(!panel)return;
 const i=selected();panel.classList.toggle("hidden",i<0);panel.setAttribute("aria-hidden",String(i<0));
 if(i===last)return;last=i;
 if(i<0)return;
 const r=routes[i];
 $("salatura-title").textContent="🧂 "+r.name;
 $("salatura-meta").textContent=r.sourceName+" · "+r.streets.length+" strade · "+number(r.routeLengthMeters)+" m ufficiali";
 $("salatura-launch").href=url(i);
 $("salatura-street-count").textContent="Elenco ufficiale · "+r.streets.length+" voci stradali";
 const list=$("salatura-street-list");list.replaceChildren();
 for(const street of r.streets){
   const li=document.createElement("li");
   li.textContent=street.name+" · "+number(street.meters)+" m";list.appendChild(li);
 }
}
function init(){
 if(!ensure())return;
 window.addEventListener("hashchange",render);window.addEventListener("pageshow",render);
 let previous="";
 setInterval(()=>{
   const key=String(typeof selectedCommessaId==="undefined"?"":selectedCommessaId)+"|"+selected()+"|"+admin()+"|"+logged();
   if(key!==previous||!$("salatura-import")){previous=key;render();}
 },1000);
 render();
}
window.VargaSalaturaStrada={routes,seed,render,url};
if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
})();
