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
  const state = {group:0, dayKey:today(), map:null, paths:null, point:null, lineLayers:[], records:[], unlisten:null, subscriptionKey:null, docs:new Map(), selected:null, busy:false, lastFitGroup:null, activeCommessaId:null, geometries:{}, geometryUnlisten:null, runId:"initial", runUnlisten:null, runSubscriptionKey:null, editing:false, editParts:[], editPartIndex:0, editLayers:[], editSelectedSegment:-1, drawingLine:false, insertingPoint:false, autoPrecisionStarted:false};
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
      section.innerHTML='<div class="snow-plan-heading"><div><h2>🗺️ Percorsi neve Castel Maggiore</h2><p class="muted">3 commesse · 4 sottopercorsi · 45 interventi. Tratti di strada, marciapiedi, ciclabili e accessi.</p></div></div><div id="snow-plan-editor" class="snow-plan-editor hidden"><button id="snow-plan-edit-start" type="button">🗺️ Apri editor mappa</button><button id="snow-plan-edit-new-part" type="button">✏️ Traccia linea</button><button id="snow-plan-edit-insert-point" type="button">➕ Inserisci punto</button><label>Tratto <select id="snow-plan-edit-part"></select></label><button id="snow-plan-edit-delete-part" type="button">🗑 Elimina piccolo pezzo</button><button id="snow-plan-edit-undo" type="button">Annulla ultimo punto</button><button id="snow-plan-edit-save" type="button">Salva mappa</button><button id="snow-plan-edit-export-kml" type="button">Esporta KML</button><button id="snow-plan-edit-import-kml" type="button">Importa KML</button><input id="snow-plan-edit-kml-file" type="file" accept=".kml,application/vnd.google-earth.kml+xml" hidden><button id="snow-plan-edit-cancel" type="button">Annulla modifica</button><span id="snow-plan-edit-hint">Seleziona un intervento per correggerlo.</span></div><button id="snow-plan-new-run" type="button" class="btn">🔄 Nuovo giro · azzera stati</button><p id="snow-plan-feedback" role="status" aria-live="polite" class="muted"></p><div class="snow-plan-filters"><label>Commessa<select id="snow-plan-group"></select></label><label>Giorno<input type="date" id="snow-plan-day"></label></div><p class="snow-plan-rules">⚠️ Piano operativo: attivazione oltre 5 cm di neve a terra. I punti prioritari sono evidenziati; verificare sul posto marciapiedi, accessi e geometrie. Ogni intervento prevede passaggi pedonali e salatura.</p><p id="snow-plan-resources" class="snow-plan-resources"></p><div class="snow-plan-summary" id="snow-plan-summary"></div><div class="snow-plan-layout"><div class="snow-plan-map-wrap"><div id="snow-plan-map" role="application" aria-label="Mappa dei percorsi neve"></div><div class="snow-plan-map-legend"><span>🔵 Da fare</span><span>🟠 In corso</span><span>🟢 Completato</span><span>🔴 Problema</span></div></div><div class="snow-plan-details"><h3 id="snow-plan-list-title">Interventi da pulire</h3><div id="snow-plan-stops"></div></div></div><p class="muted snow-plan-source">Geometrie dalla mappa Google My Maps «VIABILITA MINORE» dove disponibili. Per 2 interventi senza corrispondenza certa restano le geometrie indicative OSM/DBTR. Verificare sul posto. © OpenStreetMap contributors (ODbL), Regione Emilia-Romagna (CC BY 4.0). Il percorso di trasferimento stradale non equivale al tratto da pulire.</p>';
      const mapCard=page.querySelector("#commessa-operational-card");
      if(mapCard)mapCard.insertAdjacentElement("afterend",section);else page.appendChild(section);
      id("snow-plan-group").innerHTML=C.groups.map((g,i)=>'<option value="'+i+'">'+esc(g.name)+'</option>').join("");
      id("snow-plan-group").disabled=true;
      id("snow-plan-day").value=state.dayKey;
      id("snow-plan-day").addEventListener("change",(e)=>{state.dayKey=e.target.value||today();watchRun();watch();});
      id("snow-plan-edit-start").addEventListener("click",startEdit);
      id("snow-plan-edit-undo").addEventListener("click",undoEditPoint);
      id("snow-plan-edit-new-part").addEventListener("click",newEditPart);
      id("snow-plan-edit-insert-point").addEventListener("click",toggleInsertPoint);
      id("snow-plan-edit-delete-part").addEventListener("click",deleteEditSegment);
      id("snow-plan-edit-save").addEventListener("click",saveEdit);
      id("snow-plan-edit-export-kml").addEventListener("click",exportKml);
      id("snow-plan-edit-import-kml").addEventListener("click",()=>id("snow-plan-edit-kml-file").click());
      id("snow-plan-edit-kml-file").addEventListener("change",importKml);
      id("snow-plan-edit-cancel").addEventListener("click",cancelEdit);
      id("snow-plan-edit-part").addEventListener("change",changeEditPart);
      id("snow-plan-new-run").addEventListener("click",resetRun);

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
      const saved=state.geometries[stop.key];
      const editingParts=state.editing&&state.selected===stop.key?state.editParts:null;
      const parts=editingParts!==null?editingParts.map(coords=>({coords,matched:true})):(saved&&Array.isArray(saved.parts)?saved.parts.map(coords=>({coords,matched:true})):stop.parts||[]);
      parts.forEach((part,partIndex)=>{
        const coords=(part.coords||[]).filter((p)=>Array.isArray(p)&&p.length===2&&Number.isFinite(+p[0])&&Number.isFinite(+p[1]));
        if(coords.length<2)return;
        bounds.push(...coords);
        const uncertain=part.matched!==true || !/verified|dbtr|source-original-mymaps-kml/i.test(stop.geometryStatus||"");
        const line=L.polyline(coords,{color,weight:status==="done"?6:5,opacity:.92,dashArray:uncertain?"6 6":null,lineCap:"round"});
        line.on("click",event=>{if(state.editing&&state.drawingLine){event.originalEvent?.stopPropagation?.();onEditMapClick(event)}else if(state.editing&&state.insertingPoint){event.originalEvent?.stopPropagation?.();insertEditPoint(stop.key,partIndex,event.latlng)}else if(state.editing)selectEditSegment(stop.key,partIndex,event.latlng);else select(stop.key,true)});line.bindTooltip((stop.priority?"★ PRIORITÀ · ":"")+stop.sectionName+" · "+stop.number,{sticky:true});
        line.addTo(state.paths);lines.push(line);
      });
      if(stop.marker?.length===2){
        const marker=L.circleMarker(stop.marker,{radius:stop.priority?7:5,color:stop.priority?"#b91c1c":color,weight:2,fillOpacity:.9});
        marker.on("click",()=>{if(!state.editing)select(stop.key,true)});marker.addTo(state.paths);bounds.push(stop.marker);
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
    id("snow-plan-resources").textContent="PERSONALE: "+group.people+" · ATTREZZATURE: "+(group.equipment||[]).join(" · ")+(state.group>0?" · Trattore con spandisale condiviso sui 3 percorsi.":"");
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
    const key=normalPageVisible() && activeGroup()>=0 && dbReady() && userReady() ? state.dayKey+"|"+C.groups[state.group].id+"|"+state.runId : null;
    if(state.subscriptionKey===key && (!key || state.unlisten))return;
    if(state.unlisten){state.unlisten();state.unlisten=null;}
    state.subscriptionKey=key;
    state.records=[];if(key)draw();
    if(!key)return;
    state.unlisten=db.collection("snowRouteWork").where("dayKey","==",state.dayKey).onSnapshot((snapshot)=>{
      state.records=snapshot.docs.map(d=>({id:d.id,...d.data()})).filter(r=>(r.runId||"initial")===state.runId);
      draw();
    },(error)=>message("Impossibile sincronizzare gli interventi: "+(error.code||error.message)));
  }
  async function update(stop,action){
    if(!dbReady()||!userReady()){message("Accedi per registrare il lavoro.");return}
    if(state.busy)return;
    const uid=currentUser.uid,workId=stop.key+"_"+state.dayKey.replace(/-/g,"")+"_"+state.runId+"_"+uid;
    const mine=state.records.find((r)=>r.uid===uid&&r.stopKey===stop.key);
    const next={uid,dayKey:state.dayKey,stopKey:stop.key,runId:state.runId,cleared:!!mine?.cleared,salted:!!mine?.salted,state:mine?.state||"todo",updatedAt:firebase.firestore.FieldValue.serverTimestamp()};
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
            denominazione:(stop.sectionName+" · "+stop.number+" · "+stop.note).slice(0,140),
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
  function watchGeometry(){
    if(!dbReady()||!userReady()||!normalPageVisible()||activeGroup()<0)return;
    if(state.geometryUnlisten)return;
    state.geometryUnlisten=db.collection("snowRouteGeometry").onSnapshot((snapshot)=>{
      const next={};snapshot.docs.forEach(d=>{const data=d.data();if(Array.isArray(data.parts))next[d.id]=data;});
      state.geometries=next;draw();
      // Un amministratore autorizzato avvia automaticamente il riallineamento
      // conservativo una volta per sessione. Nessun intervento esistente
      // viene sostituito senza una corrispondenza OSM molto vicina.
      if(!state.autoPrecisionStarted && admin() && window.VargaSnowAutomaticSidewalks){
        state.autoPrecisionStarted=true;
        window.VargaSnowAutomaticSidewalks.run({
          db,admin:admin(),uid:currentUser?.uid,existing:next,
          onProgress:message
        }).catch(error=>message("Allineamento automatico non riuscito: "+(error.message||error)+". Tracce originali conservate."));
      }
    },error=>message("Mappa non sincronizzata: "+(error.code||error.message)));
  }
  function watchRun(){
    if(!dbReady()||!userReady()||!normalPageVisible()||activeGroup()<0)return;
    const group=C.groups[state.group],key=group.id+"_"+state.dayKey;
    if(state.runSubscriptionKey===key&&state.runUnlisten)return;
    if(state.runUnlisten)state.runUnlisten();
    state.runSubscriptionKey=key;state.runId="initial";
    state.runUnlisten=db.collection("snowRouteRuns").doc(key).onSnapshot(snap=>{
      const next=snap.exists&&typeof snap.data().runId==="string"?snap.data().runId:"initial";
      if(next!==state.runId){state.runId=next;state.subscriptionKey=null;watch();}
      draw();
    },error=>message("Giro non sincronizzato: "+(error.code||error.message)));
  }
  function currentStop(){return stopByKey(state.selected);}
  function openEditorPage(stop){
    let page=id("snow-plan-dedicated-editor");
    if(!page){
      page=document.createElement("section");
      page.id="snow-plan-dedicated-editor";
      page.className="snow-map-editor-page";
      page.hidden=true;
      page.setAttribute("role","dialog");
      page.setAttribute("aria-modal","true");
      page.setAttribute("aria-label","Editor dedicato della mappa");
      page.innerHTML='<header class="snow-map-editor-header"><button id="snow-map-editor-back" type="button" aria-label="Torna alla mappa della commessa">← Indietro</button><div><h1>Modifica mappa</h1><p id="snow-map-editor-route"></p></div><span class="snow-map-editor-admin">Solo admin</span></header><div class="snow-map-editor-controls"></div><main class="snow-map-editor-canvas"></main>';
      document.body.appendChild(page);
      id("snow-map-editor-back").addEventListener("click",cancelEdit);
    }
    id("snow-map-editor-route").textContent=stop.sectionName+" · intervento "+stop.number;
    const editor=id("snow-plan-editor"),mapWrap=id("snow-plan-map")?.closest(".snow-plan-map-wrap");
    if(editor&&!page.contains(editor))page.querySelector(".snow-map-editor-controls").appendChild(editor);
    if(mapWrap&&!page.contains(mapWrap))page.querySelector(".snow-map-editor-canvas").appendChild(mapWrap);
    page.hidden=false;
    document.body.classList.add("snow-map-editor-open");
    requestAnimationFrame(()=>setTimeout(()=>state.map?.invalidateSize(),80));
  }
  function closeEditorPage(){
    const page=id("snow-plan-dedicated-editor"),panel=id("snow-plan-panel"),editor=id("snow-plan-editor"),mapWrap=id("snow-plan-map")?.closest(".snow-plan-map-wrap"),layout=panel?.querySelector(".snow-plan-layout");
    if(editor&&panel)panel.insertBefore(editor,id("snow-plan-new-run"));
    if(mapWrap&&layout)layout.prepend(mapWrap);
    if(page)page.hidden=true;
    document.body.classList.remove("snow-map-editor-open");
    requestAnimationFrame(()=>setTimeout(()=>state.map?.invalidateSize(),80));
  }
  function startEdit(){
    if(!admin())return message("Solo un amministratore può correggere la mappa.");
    const stop=currentStop();if(!stop)return message("Seleziona prima un intervento dall’elenco.");
    const saved=state.geometries[stop.key];
    state.editParts=(saved&&Array.isArray(saved.parts)?saved.parts:stop.parts.map(p=>p.coords||[])).map(p=>p.map(x=>[+x[0],+x[1]]));
    state.editPartIndex=state.editParts.length?0:-1;state.editSelectedSegment=-1;state.drawingLine=false;state.insertingPoint=false;state.editing=true;
    renderEditPartOptions();
    id("snow-plan-edit-hint").textContent="Traccia linea aggiunge un pezzo. Inserisci punto divide una linea: inserisci due punti ai bordi del tratto, poi tocca il piccolo segmento rosso e cancellalo. Da PC puoi esportare o importare KML per lavorare in Google Earth o QGIS; salva qui solo dopo aver controllato la mappa.";
    if(state.map){state.map.off("click",onEditMapClick);state.map.on("click",onEditMapClick);}
    openEditorPage(stop);drawEditPreview();
    window.VargaSnowPedestrianAssist?.attach?.({map:state.map,stop,parts:state.editParts,getParts:()=>state.editParts,onPreview:parts=>{state.editParts=parts;state.editSelectedSegment=-1;drawEditPreview();}});
    message("Editor mappa aperto per "+stop.sectionName+" · "+stop.number+".");
  }
  function renderEditPartOptions(){
    const select=id("snow-plan-edit-part");if(!select)return;
    select.innerHTML=state.editParts.length?state.editParts.map((_,i)=>'<option value="'+i+'">Linea '+(i+1)+'</option>').join(""):'<option value="-1">Nessuna linea</option>';
    select.value=String(state.editPartIndex);
    id("snow-plan-edit-delete-part").disabled=!state.editing||state.editPartIndex<0||state.editSelectedSegment<0;
    id("snow-plan-edit-undo").disabled=!state.editing||!state.drawingLine||state.editPartIndex<0||!state.editParts[state.editPartIndex]?.length;
    id("snow-plan-edit-new-part").textContent=state.drawingLine?"✅ Fine linea":"✏️ Traccia linea";
    id("snow-plan-edit-insert-point").textContent=state.insertingPoint?"✅ Tocca una linea":"➕ Inserisci punto";
  }
  function onEditMapClick(e){
    if(!state.editing||!state.drawingLine)return;
    if(state.editPartIndex<0||!state.editParts[state.editPartIndex])return;
    state.editParts[state.editPartIndex].push(window.VargaSnowPedestrianAssist?.snap?.([e.latlng.lat,e.latlng.lng]) || [e.latlng.lat,e.latlng.lng]);drawEditPreview();
  }
  function clearEditLayers(){state.editLayers.forEach(x=>state.map?.removeLayer(x));state.editLayers=[];}
  function drawEditPreview(){
    clearEditLayers();if(!state.map||!state.editing)return;
    mapDraw(sectionsFor(C.groups[state.group]));
    const points=state.editParts[state.editPartIndex]||[];
    if(points.length){
      const line=L.polyline(points,{color:"#f97316",weight:7,dashArray:"8 5"}).addTo(state.map);state.editLayers.push(line);
      points.forEach((point,index)=>{
        const marker=L.marker(point,{draggable:true}).addTo(state.map);
        marker.on("drag",e=>{const p=e.target.getLatLng();points[index]=[p.lat,p.lng];line.setLatLngs(points);});
        marker.on("dragend",e=>{const p=e.target.getLatLng();points[index]=window.VargaSnowPedestrianAssist?.snap?.([p.lat,p.lng]) || [p.lat,p.lng];drawEditPreview();});
        state.editLayers.push(marker);
      });
    }
    if(state.editSelectedSegment>=0&&points[state.editSelectedSegment+1]){
      const highlight=L.polyline([points[state.editSelectedSegment],points[state.editSelectedSegment+1]],{color:"#dc2626",weight:11,opacity:.95}).addTo(state.map);
      state.editLayers.push(highlight);
    }
    renderEditPartOptions();
  }
  function closestSegmentIndex(points,latlng){
    if(!Array.isArray(points)||points.length<2||!state.map)return -1;
    const p=state.map.latLngToLayerPoint(latlng);let best=-1,bestDistance=Infinity;
    for(let i=0;i<points.length-1;i++){
      const a=state.map.latLngToLayerPoint(points[i]),b=state.map.latLngToLayerPoint(points[i+1]);
      const dx=b.x-a.x,dy=b.y-a.y,lengthSquared=dx*dx+dy*dy;
      const t=lengthSquared?Math.max(0,Math.min(1,((p.x-a.x)*dx+(p.y-a.y)*dy)/lengthSquared)):0;
      const ex=p.x-(a.x+t*dx),ey=p.y-(a.y+t*dy),distance=ex*ex+ey*ey;
      if(distance<bestDistance){bestDistance=distance;best=i;}
    }
    return best;
  }
  function selectEditSegment(stopKey,partIndex,latlng){
    if(!state.editing)return;
    if(stopKey!==state.selected)return message("Chiudi o annulla la modifica corrente prima di selezionare un altro intervento.");
    if(partIndex<0||partIndex>=state.editParts.length)return;
    state.editPartIndex=partIndex;state.editSelectedSegment=closestSegmentIndex(state.editParts[partIndex],latlng);
    state.drawingLine=false;state.insertingPoint=false;drawEditPreview();
    message(state.editSelectedSegment>=0?"Piccolo pezzo selezionato. Premi “Elimina piccolo pezzo” per toglierlo.":"Linea selezionata.");
  }
  function toggleInsertPoint(){
    if(!admin()||!state.editing)return;
    state.drawingLine=false;state.insertingPoint=!state.insertingPoint;state.editSelectedSegment=-1;
    drawEditPreview();
    message(state.insertingPoint?"Modalità attiva: tocca la linea nel punto in cui vuoi aggiungere un vertice.":"Inserimento punto terminato.");
  }
  function insertEditPoint(stopKey,partIndex,latlng){
    if(!admin()||!state.editing||!state.insertingPoint||stopKey!==state.selected)return;
    const points=state.editParts[partIndex];if(!Array.isArray(points)||points.length<2)return;
    const nearest=closestSegmentIndex(points,latlng);if(nearest<0)return;
    points.splice(nearest+1,0,window.VargaSnowPedestrianAssist?.snap?.([latlng.lat,latlng.lng]) || [latlng.lat,latlng.lng]);
    state.editPartIndex=partIndex;state.editSelectedSegment=-1;state.insertingPoint=false;
    drawEditPreview();message("Punto aggiunto. Inserisci il secondo bordo e cancella solo il pezzetto.");
  }
  function newEditPart(){
    if(!admin()||!state.editing)return;
    if(state.drawingLine){
      const current=state.editParts[state.editPartIndex]||[];
      state.drawingLine=false;
      if(current.length<2){state.editParts.splice(state.editPartIndex,1);state.editPartIndex=state.editParts.length?state.editParts.length-1:-1;}
      state.editSelectedSegment=-1;drawEditPreview();
      message(current.length<2?"Linea vuota rimossa.":"Linea terminata.");
      return;
    }
    state.editParts.push([]);state.editPartIndex=state.editParts.length-1;state.editSelectedSegment=-1;state.insertingPoint=false;state.drawingLine=true;drawEditPreview();
    message("Tocca la mappa per aggiungere punti; premi Fine linea quando hai finito.");
  }
  function deleteEditSegment(){
    if(!admin()||!state.editing||state.editPartIndex<0||state.editSelectedSegment<0)return;
    if(!window.confirm("Eliminare solo il piccolo pezzo selezionato dalla mappa?"))return;
    const partIndex=state.editPartIndex,segmentIndex=state.editSelectedSegment,points=state.editParts[partIndex]||[];
    if(!points[segmentIndex+1])return;
    const left=points.slice(0,segmentIndex+1),right=points.slice(segmentIndex+1),replacement=[];
    if(left.length>=2)replacement.push(left);
    if(right.length>=2)replacement.push(right);
    state.editParts.splice(partIndex,1,...replacement);
    state.editPartIndex=state.editParts.length?Math.min(partIndex,state.editParts.length-1):-1;
    state.editSelectedSegment=-1;state.drawingLine=false;drawEditPreview();
    message("Piccolo pezzo rimosso dalla modifica. Premi Salva mappa per confermare.");
  }
  function undoEditPoint(){
    if(!state.editing||!state.drawingLine||state.editPartIndex<0)return;
    state.editParts[state.editPartIndex].pop();drawEditPreview();
  }
  function changeEditPart(e){
    if(!state.editing)return;
    state.editPartIndex=Number(e.target.value);state.editSelectedSegment=-1;state.drawingLine=false;state.insertingPoint=false;
    drawEditPreview();
  }
  function cancelEdit(){
    state.editing=false;state.drawingLine=false;state.insertingPoint=false;state.editSelectedSegment=-1;if(state.map)state.map.off("click",onEditMapClick);clearEditLayers();window.VargaSnowPedestrianAssist?.detach?.();closeEditorPage();
    const hint=id("snow-plan-edit-hint");if(hint)hint.textContent="Seleziona un intervento per correggerlo.";if(normalPageVisible()&&activeGroup()>=0)draw();
  }
  function exportKml(){
    const stop=currentStop();if(!stop)return message("Seleziona un intervento prima di esportare.");
    const saved=state.editing?state.editParts:(state.geometries[stop.key]?.parts||stop.parts.map(p=>p.coords||[]));
    const lines=saved.filter(part=>Array.isArray(part)&&part.length>=2);
    if(!lines.length)return message("Questo intervento non ha linee da esportare.");
    const placemarks=lines.map((part,index)=>{
      const coordinates=part.map(point=>Number(point[1])+","+Number(point[0])+",0").join(" ");
      return "<Placemark><name>"+esc(stop.sectionName+" · "+stop.number+" · Linea "+(index+1))+"</name><LineString><tessellate>1</tessellate><coordinates>"+coordinates+"</coordinates></LineString></Placemark>";
    }).join("");
    const xml='<?xml version="1.0" encoding="UTF-8"?><kml xmlns="http://www.opengis.net/kml/2.2"><Document><name>'+esc(stop.sectionName+" · intervento "+stop.number)+'</name>'+placemarks+'</Document></kml>';
    const blob=new Blob([xml],{type:"application/vnd.google-earth.kml+xml;charset=utf-8"}),url=URL.createObjectURL(blob),link=document.createElement("a");
    link.href=url;link.download=(stop.key||"percorso-neve").replace(/[^a-z0-9_-]+/gi,"-")+".kml";
    document.body.appendChild(link);link.click();link.remove();setTimeout(()=>URL.revokeObjectURL(url),1000);
    message("KML esportato. Aprilo in Google Earth Pro, Google Earth Web o QGIS.");
  }
  async function importKml(event){
    const input=event.target,file=input.files?.[0];if(!file)return;
    try{
      if(!state.editing||!admin())return message("Apri prima l’editor dedicato come amministratore.");
      const textContent=await file.text(),xml=new DOMParser().parseFromString(textContent,"application/xml");
      if(xml.querySelector("parsererror"))throw new Error("Il file KML non è valido.");
      let lineStrings=Array.from(xml.getElementsByTagNameNS("*","LineString"));
      if(!lineStrings.length)lineStrings=Array.from(xml.getElementsByTagName("LineString"));
      const imported=lineStrings.map(line=>{
        let node=line.getElementsByTagNameNS("*","coordinates")[0]||line.getElementsByTagName("coordinates")[0];
        if(!node)return [];
        return (node.textContent||"").trim().split(/\s+/).map(token=>{
          const values=token.split(",").map(Number);
          return values.length>=2&&Number.isFinite(values[0])&&Number.isFinite(values[1])?[values[1],values[0]]:null;
        }).filter(Boolean);
      }).filter(part=>part.length>=2);
      if(!imported.length)throw new Error("Nel KML non ho trovato linee con almeno due punti.");
      if(!window.confirm("Caricare "+imported.length+" linea/e KML al posto delle linee attualmente in modifica? Le modifiche diventano permanenti solo dopo “Salva mappa”."))return;
      state.editParts=imported;state.editPartIndex=0;state.editSelectedSegment=-1;state.drawingLine=false;state.insertingPoint=false;
      renderEditPartOptions();drawEditPreview();
      message(imported.length+" linee importate in anteprima. Controlla la mappa e premi “Salva mappa” per conservarle.");
    }catch(error){message("Importazione KML non riuscita: "+(error.message||error));}
    finally{input.value="";}
  }
  async function saveEdit(){
    if(!admin()||!state.editing)return;
    const stop=currentStop();if(!stop)return;
    if(state.editParts.some(part=>part.length<2))return message("Completa ogni linea con almeno due punti o rimuovi il tratto vuoto.");
    try{
      const savedParts=state.editParts.map(p=>p.map(x=>[Number(x[0]),Number(x[1])]));
      await db.collection("snowRouteGeometry").doc(stop.key).set({
        stopKey:stop.key,parts:savedParts,
        updatedAt:firebase.firestore.FieldValue.serverTimestamp(),updatedBy:currentUser?.email||""
      },{merge:true});
      state.geometries[stop.key]={stopKey:stop.key,parts:savedParts};
      cancelEdit();message("Percorso salvato in modo permanente."); 
    }catch(error){message("Salvataggio mappa non riuscito: "+(error.code||error.message));}
  }
  async function resetRun(){
    if(!admin())return;
    if(!window.confirm("Avviare un nuovo giro? Gli stati ripartono da zero. Percorsi e cronologia saranno conservati."))return;
    const group=C.groups[state.group],key=group.id+"_"+state.dayKey;
    const runId=Date.now().toString(36)+"-"+Math.random().toString(36).slice(2,8);
    try{
      await db.collection("snowRouteRuns").doc(key).set({groupId:group.id,dayKey:state.dayKey,runId,updatedAt:firebase.firestore.FieldValue.serverTimestamp(),updatedBy:currentUser?.email||""},{merge:true});
      message("Nuovo giro avviato. Gli interventi sono di nuovo da fare; la mappa e il giro precedente restano salvati.");
    }catch(error){message("Reset non riuscito: "+(error.code||error.message));}
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
      watchGeometry();watchRun();
      draw();watch();
      requestAnimationFrame(()=>setTimeout(()=>state.map?.invalidateSize(),180));
    }else{
      state.activeCommessaId=null;state.lastFitGroup=null;
      if(state.unlisten){state.unlisten();state.unlisten=null;state.subscriptionKey=null;}
      if(state.geometryUnlisten){state.geometryUnlisten();state.geometryUnlisten=null;}
      if(state.runUnlisten){state.runUnlisten();state.runUnlisten=null;state.runSubscriptionKey=null;}
      cancelEdit();
    }
  }
  function init(){
    if(!ensureDom())return;
    window.addEventListener("hashchange",show);
    window.addEventListener("pageshow",show);
    window.addEventListener("focus",show);
    // Le pagine commessa vengono aperte senza cambiare sempre l'hash.
    // La verifica leggera non invia richieste se la commessa non cambia.
    let oldId=null,oldGroup=-2,oldSignedIn=false,wasVisible=false;
    setInterval(()=>{
      const importButton=id("snow-plan-import");
      if(importButton)importButton.hidden=!admin();
      const editBox=id("snow-plan-editor"),runButton=id("snow-plan-new-run");
      if(editBox)editBox.classList.toggle("hidden",!admin());
      if(runButton)runButton.hidden=!admin();
      const visible=normalPageVisible(),active=visible && typeof selectedCommessaId!=="undefined" ? selectedCommessaId : null;
      const group=visible?activeGroup():-1,signedIn=userReady();
      if(visible!==wasVisible || active!==oldId || group!==oldGroup || signedIn!==oldSignedIn){
        wasVisible=visible;oldId=active;oldGroup=group;oldSignedIn=signedIn;show();
      }
    },900);
    show();
  }
  window.VargaSnowPlanUI={show,refresh:draw,catalog:C,aggregate};
  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",init,{once:true});else init();
})();
