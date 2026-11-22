/* Varga Cantieri - percorsi neve nelle COMMESSE NORMALI. Non usa neve_commesse. */
(function () {
  "use strict";
  const C = window.VARGA_SNOW_CM_CATALOG;
  if (!C || !Array.isArray(C.groups) || !Array.isArray(C.sections)) return;
  const id = (s) => document.getElementById(s);
  const esc = (v) => String(v == null ? "" : v).replace(/[&<>"']/g, (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const dbReady = () => typeof db !== "undefined" && !!db;
  const userReady = () => typeof currentUser !== "undefined" && !!currentUser;
  const admin = () => typeof canManageData === "function" && !!canManageData();
  const sectionsFor = (group) => group.indices.flatMap((sectionIndex) => C.sections[sectionIndex].stops.map((stop) => ({...stop, sectionIndex, sectionName:C.sections[sectionIndex].name})));
  const today = () => { const d=new Date(); return [d.getFullYear(),String(d.getMonth()+1).padStart(2,"0"),String(d.getDate()).padStart(2,"0")].join("-"); };
  const state = {group:0, dayKey:today(), map:null, paths:null, point:null, lineLayers:[], records:[], unlisten:null, subscriptionKey:null, docs:new Map(), selected:null, busy:false, lastFitGroup:null, activeCommessaId:null};
  const normalPageVisible = () => {
    const page=id("impianti-page");
    return !!page && !page.hidden && !page.classList.contains("hidden") && !document.body.classList.contains("snow-management-context");
  };
  function activeGroup(){
    if(!normalPageVisible() || typeof selectedCommessaId==="undefined" || !selectedCommessaId || typeof commesseById==="undefined") return -1;
    const c=commesseById.get(selectedCommessaId);
    if(!c) return -1;
    const code=String(c.codice||"").trim().toUpperCase(),name=String(c.nome||"").trim().toLowerCase();
    return C.groups.findIndex(g=>code===g.code || name===g.name.toLowerCase() || (c.snowPlan==="castel-maggiore-viabilita-minore" && String(c.id||selectedCommessaId)===String(state.docs.get(g.id)||"")));
  }
  const stopByKey = (key) => C.sections.flatMap((s,i)=>s.stops.map((x)=>({...x,sectionIndex:i}))).find((s)=>s.key===key);
  function targetGroup(stop){ return C.groups.findIndex((g)=>g.indices.includes(stop.sectionIndex)); }
  function aggregate(stopKey){
    const a=state.records.filter((r)=>r.stopKey===stopKey);
    const cleared=a.some((r)=>r.cleared===true),salted=a.some((r)=>r.salted===true);
    if(cleared && salted)return "done";
    if(cleared)return "cleared";
    if(salted)return "salted";
    if(a.some((r)=>r.state==="in_progress"))return "active";
    if(a.some((r)=>r.state==="blocked"))return "blocked";
    return "todo";
  }
  const colors={done:"#16a34a",cleared:"#0d9488",salted:"#a855f7",active:"#f59e0b",blocked:"#dc2626",todo:"#2563eb"};
  const labels={done:"Pulizia e salatura completate",cleared:"Pulito · da salare",salted:"Salato · da pulire",active:"In corso",blocked:"Problema segnalato",todo:"Da fare"};
  function ensureDom(){
    const page=id("impianti-page");
    if(!page)return false;
    if(!id("snow-plan-panel")){
      const section=document.createElement("section");
      section.id="snow-plan-panel";section.className="card snow-plan-panel hidden";section.setAttribute("aria-hidden","true");
      section.innerHTML='<div class="snow-plan-heading"><div><h2>🗺️ Percorsi neve Castel Maggiore</h2><p class="muted">3 commesse · 4 sottopercorsi · 45 interventi. Tratti di strada, marciapiedi, ciclabili e accessi.</p></div></div><p id="snow-plan-feedback" role="status" aria-live="polite" class="muted"></p><div class="snow-plan-filters"><label>Commessa<select id="snow-plan-group"></select></label><label>Giorno<input type="date" id="snow-plan-day"></label></div><p class="snow-plan-rules">⚠️ Piano operativo: attivazione oltre 5 cm di neve a terra. I punti prioritari sono evidenziati; verificare sul posto marciapiedi, accessi e geometrie. Ogni intervento prevede passaggi pedonali e salatura.</p><div class="snow-plan-summary" id="snow-plan-summary"></div><div class="snow-plan-layout"><div class="snow-plan-map-wrap"><div id="snow-plan-map" role="application" aria-label="Mappa dei percorsi neve"></div><div class="snow-plan-map-legend"><span>🔵 Da fare</span><span>🟠 In corso</span><span>🟢 Completato</span><span>🔴 Problema</span></div></div><div class="snow-plan-details"><h3 id="snow-plan-list-title">Interventi da pulire</h3><div id="snow-plan-stops"></div></div></div><p class="muted snow-plan-source">Tracce ricostruite da OSM / DBTR e ortofoto, da confermare sul posto. © OpenStreetMap contributors (ODbL), Regione Emilia-Romagna (CC BY 4.0). Il percorso di trasferimento stradale non equivale al tratto da pulire.</p>';
      const mapCard=page.querySelector("#commessa-operational-card");
      if(mapCard)mapCard.insertAdjacentElement("afterend",section);else page.appendChild(section);
      id("snow-plan-group").innerHTML=C.groups.map((g,i)=>'<option value="'+i+'">'+esc(g.name)+'</option>').join("");
      id("snow-plan-group").disabled=true;
      id("snow-plan-day").value=state.dayKey;
      id("snow-plan-day").addEventListener("change",(e)=>{state.dayKey=e.target.value||today();watch();});

    }
    // Importazione delle 3 commesse dall'elenco ordinario di Gestione commesse.
    const tools=id("panel-commesse")?.querySelector(".commesse-management-actions");
    if(tools && !id("snow-plan-import")){
      const wrap=document.createElement("div");wrap.className="snow-normal-import-wrap";
      const button=document.createElement("button");button.type="button";button.className="btn";button.id="snow-plan-import";button.textContent="❄️ Inserisci Castel Maggiore 1, 2, 3";
      const feedback=document.createElement("p");feedback.id="snow-plan-seed-feedback";feedback.className="muted";feedback.setAttribute("role","status");
      wrap.append(button,feedback);tools.appendChild(wrap);button.addEventListener("click",seed);
    }
    if(id("snow-plan-import"))id("snow-plan-import").hidden=!admin();
    return true;
  }
  function mapInit(){
    if(state.map || !window.L || !id("snow-plan-map"))return;
    state.map=L.map("snow-plan-map",{zoomControl:true,preferCanvas:true}).setView([44.576,11.36],13);
    const base=L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png",{maxZoom:20,maxNativeZoom:19,attribution:"© OpenStreetMap contributors"});
    const photo=L.tileLayer("https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}",{maxZoom:20,maxNativeZoom:19,attribution:"Immagini © Esri e fornitori"});
    base.addTo(state.map);
    L.control.layers({"Strade":base,"Satellite":photo},null,{position:"topright",collapsed:true}).addTo(state.map);
    state.paths=L.layerGroup().addTo(state.map);
  }
  function mapDraw(stops){
    mapInit();if(!state.map)return;
    state.paths.clearLayers();state.lineLayers=[];
    const bounds=[];
    stops.forEach((stop)=>{
      const status=aggregate(stop.key),color=colors[status],lines=[];
      (stop.parts||[]).forEach((part)=>{
        const coords=(part.coords||[]).filter((p)=>Array.isArray(p)&&p.length===2&&Number.isFinite(+p[0])&&Number.isFinite(+p[1]));
        if(coords.length<2)return;
        bounds.push(...coords);
        const uncertain=part.matched!==true || !/verified|dbtr/i.test(stop.geometryStatus||"");
        const line=L.polyline(coords,{color,weight:status==="done"?6:5,opacity:.92,dashArray:uncertain?"6 6":null,lineCap:"round"});
        line.on("click",()=>select(stop.key,true));line.bindTooltip((stop.priority?"★ PRIORITÀ · ":"")+stop.sectionName+" · "+stop.number,{sticky:true});
        line.addTo(state.paths);lines.push(line);
      });
      if(stop.marker?.length===2){
        const marker=L.circleMarker(stop.marker,{radius:stop.priority?7:5,color:stop.priority?"#b91c1c":color,weight:2,fillOpacity:.9});
        marker.on("click",()=>select(stop.key,true));marker.addTo(state.paths);bounds.push(stop.marker);
        lines.push(marker);
      }
      state.lineLayers.push({key:stop.key,lines});
    });
    if(bounds.length && state.lastFitGroup !== state.group){state.map.fitBounds(bounds,{padding:[24,24],maxZoom:16,animate:false});state.lastFitGroup=state.group;}
  }
  function select(key,center){
    const stop=stopByKey(key);
    if(!stop)return;
    state.selected=stop.key;
    const group=targetGroup(stop);
    if(group!==state.group){state.group=group;id("snow-plan-group").value=String(group);draw();return;}
    const row=id("snow-plan-stops")?.querySelector('[data-stop-key="'+stop.key+'"]');
    id("snow-plan-stops")?.querySelectorAll(".snow-plan-stop").forEach(x=>x.classList.toggle("selected",x.dataset.stopKey===stop.key));
    if(row)row.scrollIntoView({block:"nearest",behavior:"smooth"});
    if(center&&state.map) {
      const coords=stop.parts.flatMap(p=>p.coords||[]);
      if(coords.length)state.map.fitBounds(coords,{padding:[40,40],maxZoom:18,animate:false});
      else if(stop.marker)state.map.setView(stop.marker,17);
    }
  }
  function navigation(stop){
    const pos=stop.marker||stop.parts?.[0]?.coords?.[0];
    if(!pos){message("Punto di accesso non disponibile: verifica l'indirizzo sulla mappa.");return;}
    const fallback="https://www.google.com/maps/dir/?api=1&destination="+encodeURIComponent(pos.join(","));
    const link=window.VargaNavigation?.buildUrl?.(pos[0],pos[1])||fallback;
    window.open(link,"_blank","noopener,noreferrer");
  }
  function message(value){
    for(const name of ["snow-plan-feedback","snow-plan-seed-feedback"]){const node=id(name);if(node)node.textContent=value||"";}
  }
  function draw(){
    if(!ensureDom() || activeGroup()<0)return;
    const group=C.groups[state.group],stops=sectionsFor(group);
    id("snow-plan-group").value=String(state.group);
    id("snow-plan-list-title").textContent="Interventi · "+group.name;
    const completed=stops.filter((s)=>aggregate(s.key)==="done").length;
    const begun=stops.filter((s)=>["done","cleared","salted","active"].includes(aggregate(s.key))).length;
    id("snow-plan-summary").textContent=completed+" / "+stops.length+" interventi completati · "+begun+" avviati · "+(stops.length-completed)+" ancora da completare";
    const list=id("snow-plan-stops");list.replaceChildren();
    stops.forEach((stop)=>{
      const st=aggregate(stop.key),userId=userReady()?currentUser.uid:null,rec=state.records.find(r=>r.uid===userId&&r.stopKey===stop.key);
      const article=document.createElement("article");article.className="snow-plan-stop"+(state.selected===stop.key?" selected":"");article.dataset.stopKey=stop.key;
      article.innerHTML='<div class="snow-plan-stop-head"><strong>'+(stop.priority?'🔴 PRIORITÀ · ':'')+esc(stop.sectionName)+' · '+stop.number+'</strong><span style="color:'+colors[st]+'">'+esc(labels[st])+'</span></div><p>'+esc(stop.note)+'</p><div class="snow-plan-stop-buttons"><button type="button" data-action="view">📍 Mappa</button><button type="button" data-action="navigate">🧭 Naviga</button><button type="button" data-action="start">▶ Inizia</button><button type="button" data-action="clear">'+(rec?.cleared?'✓ Pulito':'✓ Pulizia')+'</button><button type="button" data-action="salt">'+(rec?.salted?'✓ Salato':'🧂 Salatura')+'</button><button type="button" data-action="problem">⚠ Problema</button></div><small>Tratto '+(stop.parts?.length?"presente in cartografia":"non disponibile")+' · Verificare accessi e lati dei marciapiedi sul posto.</small>';
      article.querySelectorAll("button").forEach((button)=>button.addEventListener("click",()=>{
        const act=button.dataset.action;
        if(act==="view")return select(stop.key,true);
        if(act==="navigate")return navigation(stop);
        update(stop,act);
      }));
      list.appendChild(article);
    });
    mapDraw(stops);
  }
  function watch(){
    const key=normalPageVisible() && activeGroup()>=0 && dbReady() && userReady() ? state.dayKey : null;
    if(state.subscriptionKey===key && (!key || state.unlisten))return;
    if(state.unlisten){state.unlisten();state.unlisten=null;}
    state.subscriptionKey=key;
    state.records=[];if(key)draw();
    if(!key)return;
    state.unlisten=db.collection("snowRouteWork").where("dayKey","==",state.dayKey).onSnapshot((snapshot)=>{
      state.records=snapshot.docs.map(d=>({id:d.id,...d.data()}));
      draw();
    },(error)=>message("Impossibile sincronizzare gli interventi: "+(error.code||error.message)));
  }
  async function update(stop,action){
    if(!dbReady()||!userReady()){message("Accedi per registrare il lavoro.");return}
    if(state.busy)return;
    const uid=currentUser.uid,workId=stop.key+"_"+state.dayKey.replace(/-/g,"")+"_"+uid;
    const mine=state.records.find((r)=>r.uid===uid&&r.stopKey===stop.key);
    const next={uid,dayKey:state.dayKey,stopKey:stop.key,cleared:!!mine?.cleared,salted:!!mine?.salted,state:mine?.state||"todo",updatedAt:firebase.firestore.FieldValue.serverTimestamp()};
    if(action==="start")next.state="in_progress";
    else if(action==="clear"){next.cleared=true;next.state=next.salted?"done":"in_progress"}
    else if(action==="salt"){next.salted=true;next.state=next.cleared?"done":"in_progress"}
    else if(action==="problem")next.state="blocked";
    else return;
    state.busy=true;message("Salvataggio intervento…");
    try{await db.collection("snowRouteWork").doc(workId).set(next,{merge:true});message("Intervento registrato per "+state.dayKey+".")}
    catch(error){message("Salvataggio non riuscito: "+(error.message||error));}
    finally{state.busy=false}
  }
  async function seed(){
    if(state.busy || !admin() || !dbReady()) {message("Solo l'amministratore può preparare le commesse.");return;}
    state.busy=true; const btn=id("snow-plan-import");btn.disabled=true;
    message("Verifico le commesse e i 45 interventi, senza modificare quelli già presenti…");
    try{
      const all=await db.collection("commesse").get(), known=all.docs;
      const refs=[];
      for(const group of C.groups){
        const found=known.find((d)=>String(d.data().codice||"").toUpperCase()===group.code || String(d.data().nome||"").trim().toLowerCase()===group.name.toLowerCase());
        const ref=found?.ref||db.collection("commesse").doc("normal-"+group.id);
        const current=await ref.get();
        if(current.exists && !found){
          throw new Error("ID occupato da una commessa differente: "+ref.id);
        }
        if(!current.exists)await ref.set({
          nome:group.name,codice:group.code,tipo:"main",snowPlan:"castel-maggiore-viabilita-minore",
          createdAt:firebase.firestore.FieldValue.serverTimestamp(),createdBy:currentUser?.email||""
        });
        refs.push(ref);
      }
      let created=0,existing=0;
      for(let groupIndex=0;groupIndex<C.groups.length;groupIndex++){
        const group=C.groups[groupIndex],root=refs[groupIndex];
        const saved=await root.collection("impianti").get();
        const keys=new Set(saved.docs.map(d=>d.data().snowPlanStopKey).filter(Boolean));
        const ids=new Set(saved.docs.map(d=>d.id));
        const pending=[];
        for(const stop of sectionsFor(group)){
          const docId="piano-neve-"+stop.key;
          if(keys.has(stop.key)||ids.has(docId)){existing++;continue}
          const marker=stop.marker||stop.parts?.[0]?.coords?.[0]||null;
          const first=stop.parts.find(p=>p.coords?.length>1);
          const data={
            denominazione:stop.note.split(":")[0].slice(0,120),
            indirizzo:stop.note,
            descrizioneVia:stop.note,
            noteOperative:stop.note,
            comune:"Castel Maggiore",localita:stop.sectionName,
            tipologiaImpianto:"Via neve",
            codicePrezzo:"NEVE-STRADA",
            snowRoad:true,snowPlanStopKey:stop.key,snowPlanSection:stop.sectionIndex,
            snowPriority:!!stop.priority,snowPlanSource:C.version,
            routePath:(first?.coords||[]).map(p=>({lat:p[0],lng:p[1]})),
            roadStatus:"todo",done:false,doneAt:null,doneBy:"",
            sortOrder:stop.sectionIndex===3?6+stop.number:stop.number,createdAt:firebase.firestore.FieldValue.serverTimestamp(),createdBy:currentUser?.email||""
          };
          if(marker){data.latitudine=marker[0];data.longitudine=marker[1];data.lat=marker[0];data.lng=marker[1];}
          pending.push({ref:root.collection("impianti").doc(docId),data});
        }
        if(pending.length){const batch=db.batch();pending.forEach(({ref,data})=>batch.set(ref,data));await batch.commit();created+=pending.length}
      }
      state.docs=new Map(C.groups.map((g,i)=>[g.id,refs[i].id]));
      message("COMMESSE NORMALI: "+C.groups.length+" commesse pronte, "+created+" interventi creati, "+existing+" già presenti. Nessun lavoro precedente sovrascritto.");
      show();
    }catch(error){console.error("Importazione piano neve",error);message("Creazione non completata: "+(error.message||error)+". Riprova: la procedura evita i duplicati.");}
    finally{state.busy=false;btn.disabled=false}
  }
  function show(){
    if(!ensureDom())return;
    const index=activeGroup(),visible=index>=0;
    const panel=id("snow-plan-panel");
    panel.classList.toggle("hidden",!visible);panel.setAttribute("aria-hidden",String(!visible));
    if(visible){
      if(state.activeCommessaId!==selectedCommessaId || state.group!==index){
        state.activeCommessaId=selectedCommessaId;state.group=index;state.selected=null;state.lastFitGroup=null;
      }
      draw();watch();
      requestAnimationFrame(()=>setTimeout(()=>state.map?.invalidateSize(),180));
    }else{
      state.activeCommessaId=null;state.lastFitGroup=null;
      if(state.unlisten){state.unlisten();state.unlisten=null;state.subscriptionKey=null;}
    }
  }
  function init(){
    if(!ensureDom())return;
    window.addEventListener("hashchange",show);
    window.addEventListener("pageshow",show);
    window.addEventListener("focus",show);
    // Le pagine commessa vengono aperte senza cambiare sempre l'hash.
    // La verifica leggera non invia richieste se la commessa non cambia.
    let oldId=null,wasVisible=false;
    setInterval(()=>{
      const visible=normalPageVisible(),active=visible?selectedCommessaId:null;
      if(visible!==wasVisible || active!==oldId){wasVisible=visible;oldId=active;show();}
    },900);
    show();
  }
  window.VargaSnowPlanUI={show,refresh:draw,catalog:C,aggregate};
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
})();
