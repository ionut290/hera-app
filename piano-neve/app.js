const map=L.map('map',{zoomControl:false,rotate:true,rotateControl:false,touchRotate:false,shiftKeyRotate:false,bearing:0}).setView([44.4949,11.3426],13);const streetMap=L.tileLayer('https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png',{maxZoom:21,maxNativeZoom:19,attribution:'© OpenStreetMap'}).addTo(map);
const satelliteMap=L.tileLayer('https://services.arcgisonline.com/ArcGIS/rest/services/World_Imagery/MapServer/tile/{z}/{y}/{x}',{maxZoom:21,maxNativeZoom:19,attribution:'Immagini © Esri, Maxar, Earthstar Geographics e fornitori'});
map.createPane('officialPaths',map.getPane('rotatePane'));map.getPane('officialPaths').style.zIndex=350;
const regionalOrtho=L.tileLayer.wms('https://servizigis.regione.emilia-romagna.it/wms/rer2023_24_rgb',{layers:'RER2023_24_RGB',format:'image/jpeg',transparent:false,version:'1.1.1',maxZoom:21,attribution:'Ortofoto 2023–24 © Regione Emilia-Romagna · CC BY 4.0'});
const regionalWms='https://servizigis.regione.emilia-romagna.it/wms/dbtr';
const officialOptions={format:'image/png',transparent:true,version:'1.1.1',minZoom:16,maxZoom:21,pane:'officialPaths',attribution:'DBTR © Regione Emilia-Romagna · CC BY 4.0'};
const officialFootways=L.tileLayer.wms(regionalWms,{...officialOptions,layers:'ACP_Area_di_circolazione_pedonale'});
const officialCycleways=L.tileLayer.wms(regionalWms,{...officialOptions,layers:'ACI_Area_di_circolazione_ciclabile,EPC_Elemento_di_percorso_ciclabile'});
L.control.layers({'Mappa':streetMap,'Satellite':satelliteMap,'Ortofoto Regione 2023–24 · 20 cm':regionalOrtho},{'Aree pedonali · Regione':officialFootways,'Piste ciclabili · Regione':officialCycleways},{position:'topright',collapsed:true}).addTo(map);
let officialErrorShown=false;
[officialFootways,officialCycleways].forEach(layer=>layer.on('tileerror',()=>{if(!officialErrorShown&&currentMode==='minor'){officialErrorShown=true;$('officialInfo').textContent='Cartografia regionale non disponibile: restano visibili il percorso e la mappa di base.'}}));

L.control.zoom({position:'topright'}).addTo(map);
function repairMapSize(){requestAnimationFrame(()=>{map.invalidateSize({pan:false,animate:false});setTimeout(()=>map.invalidateSize({pan:false,animate:false}),180)})}
window.addEventListener('resize',repairMapSize);window.addEventListener('orientationchange',()=>setTimeout(repairMapSize,250));document.addEventListener('visibilitychange',()=>{if(!document.hidden)repairMapSize()});
if(window.ResizeObserver)new ResizeObserver(repairMapSize).observe(document.getElementById('map'));
const rebuiltNames=window.ROUTE_INSTRUCTIONS?.names||[];const rebuiltNotes=window.ROUTE_INSTRUCTIONS?.notes||[];const routes=rebuiltNotes.map((notes,rix)=>({name:rebuiltNames[rix]||('Percorso '+(rix+1)),segments:notes.map((name,i)=>({name,coords:window.REBUILT_MINOR_ROUTES?.[rix]?.[i]?.coords||[]})),points:notes}));
const majorRoutes=window.MAJOR_ROUTES||[];
const APP_BUILD='74-dev-salatura1';const SW_BUILD='74-dev-salatura1';const savedMajorRoute=localStorage.getItem('majorSelectedRoute');
let ri=0,pi=0,majorRi=savedMajorRoute!==null&&Number.isInteger(+savedMajorRoute)&&+savedMajorRoute>=0&&+savedMajorRoute<majorRoutes.length?+savedMajorRoute:0,majorPi=0,currentMode="major",minorMapOpened=false;
const done=new Set(JSON.parse(localStorage.getItem('snowDone')||'[]')),skipped=new Set(JSON.parse(localStorage.getItem('snowSkipped')||'[]')),majorDone=new Set(JSON.parse(localStorage.getItem('majorDone')||'[]')),routeLayer=L.layerGroup().addTo(map),allRoutesLayer=L.layerGroup().addTo(map),geoCache={};let navTarget=null,navLine=null,navRoute=null,majorGuideLine=null,majorPauseMarker=null,majorVehicleMarker=null;const majorNav={mode:'idle',target:null,pausePoint:null,resumeMode:null,resumeTarget:null,path:[],maneuvers:[],pathIndex:0,lastVoiceKey:'',lastInsideStreet:null,lastRerouteAt:0,rerouting:false};
async function roadRoute(a,b){try{const u='https://router.project-osrm.org/route/v1/driving/'+a[1]+','+a[0]+';'+b[1]+','+b[0]+'?overview=full&geometries=geojson&steps=true';const j=await fetch(u).then(r=>r.json());if(j.routes&&j.routes[0])return j.routes[0]}catch(e){}return null}
function fmtDist(m){return m<1000?Math.round(m)+' m':(m/1000).toFixed(1)+' km'}
function navUpdate(ll){if(!navTarget)return;const meters=d(ll,navTarget);if($('routeInfo'))$('routeInfo').textContent='🧭 NAVIGAZIONE • '+fmtDist(meters)+' alla tappa '+(pi+1);if(meters<25){$('routeInfo').textContent='📍 ARRIVATO ALLA TAPPA '+(pi+1);navTarget=null;syncScreenAwake()}}
function clearMajorNavLine(){if(navLine){routeLayer.removeLayer(navLine);navLine=null}navRoute=null;majorNav.path=[];majorNav.maneuvers=[];majorNav.pathIndex=0}
function clearMajorGuide(){if(majorGuideLine){routeLayer.removeLayer(majorGuideLine);majorGuideLine=null}}
function setMajorNavHud(show,icon='↑',instruction='',distance='',street=''){const box=$('majorNavHud');if(!box)return;box.hidden=!show;if(!show)return;$('majorNavIcon').textContent=icon;$('majorNavInstruction').textContent=instruction;$('majorNavDistance').textContent=distance;$('majorNavStreet').textContent=street}
function updateMajorVehicle(ll,heading){if(currentMode!=='major'||!mapState.active){if(majorVehicleMarker){map.removeLayer(majorVehicleMarker);majorVehicleMarker=null}return}const deg=Number.isFinite(heading)?heading:0,icon=L.divIcon({className:'majorVehicleIcon',html:'<span style="transform:rotate('+deg+'deg)">▲</span>',iconSize:[38,38],iconAnchor:[19,19]});if(!majorVehicleMarker)majorVehicleMarker=L.marker(ll,{icon,interactive:false,rotateWithView:true,zIndexOffset:2000}).addTo(map);else{majorVehicleMarker.setLatLng(ll);majorVehicleMarker.setIcon(icon)}}
function restoreMajorNavVisuals(){if(!mapState.active)return;if(mapState.paused){setMajorNavHud(true,'⏸','Pausa','',majorNav.pausePoint?'Punto di ripresa memorizzato':'');return}if((majorNav.mode==='to-start'||majorNav.mode==='to-pause')&&majorNav.path?.length){if(navLine)routeLayer.removeLayer(navLine);navLine=L.polyline(majorNav.path,{color:'#1677ff',weight:8,opacity:.96}).addTo(routeLayer);if(mapState.current)transferGuidance(mapState.current)}else if(majorNav.mode==='inside'&&mapState.current)updateMajorInsideGuide(mapState.current,20)}
function majorIcon(type){return type==='left'?'↰':type==='right'?'↱':type==='uturn'?'↶':type==='arrive'?'⚑':'↑'}
function majorSpeak(text,key='',force=false){return window.MajorOfflineNav?.speak?.(text,key,force)||false}
function majorRouteStartPoint(){if(!majorTracking.lanes.length)return null;let lanes=majorTracking.lanes.filter(x=>x.streetIndex===0);if(!lanes.length){const m=Math.min(...majorTracking.lanes.map(x=>x.streetIndex));lanes=majorTracking.lanes.filter(x=>x.streetIndex===m)}return lanes[0]?.coords?.[0]||null}
function majorAllowedDistance(accuracy){return Math.max(14,Math.min(28,(Number.isFinite(accuracy)?accuracy:15)+6))}
function nearestPathIndex(ll,points,startAt=0){if(!points?.length)return 0;let bi=Math.max(0,startAt-8),bd=Infinity;for(let i=Math.max(0,startAt-8);i<points.length;i++){const dd=d(ll,points[i]);if(dd<bd){bd=dd;bi=i}}return bi}
function pathDistance(points,a,b){let m=0;for(let i=Math.max(1,a+1);i<=b&&i<points.length;i++)m+=d(points[i-1],points[i]);return m}
function distanceToTransferPath(ll){const pts=majorNav.path||[];if(pts.length<2)return Infinity;let best=Infinity;for(let i=Math.max(1,majorNav.pathIndex-12);i<pts.length;i++){const q=nearestOnSegment(ll,pts[i-1],pts[i]);if(q.dist<best)best=q.dist;if(i>majorNav.pathIndex+80&&best<60)break}return best}
const navigationHeading=window.VargaNavigationHeading();
const positionFollow={enabled:true};
function setPositionFollow(enabled){
 positionFollow.enabled=!!enabled;const button=$('followPosition');if(button){button.textContent=enabled?'◎ TI SEGUO':'◎ SEGUIMI';button.setAttribute('aria-pressed',String(!!enabled))}
 if(enabled&&mapState.current)autoMajorView(mapState.current,0,0,true)
}
map.on('dragstart',()=>setPositionFollow(false));
map.on('zoomstart',e=>{if(e.originalEvent)setPositionFollow(false)});
function autoMajorView(ll,speed=0,remaining=0,inside=false){
 if(!positionFollow.enabled)return;
 let z=18;if(!inside)z=remaining>1800?15:remaining>700?16:remaining>220?17:18;else if(Number.isFinite(speed)&&speed>13)z=17;
 // One GPS update produces one completed map move; no competing pan animation.
 map.stop();map.setBearing?.(-(navigationHeading.value??0));map.setView(ll,z,{animate:false});
}

function transferGuidance(ll){
 const pts=majorNav.path||[];if(!pts.length)return;majorNav.pathIndex=nearestPathIndex(ll,pts,majorNav.pathIndex);
 const next=majorNav.maneuvers.find(m=>m.index>majorNav.pathIndex+1)||majorNav.maneuvers.at(-1);
 if(!next)return;const remaining=d(ll,pts[majorNav.pathIndex]||ll)+pathDistance(pts,majorNav.pathIndex,next.index),distText=fmtDist(remaining);
 const label=next.type==='arrive'?(majorNav.mode==='to-pause'?'Raggiungi il punto di pausa':'Raggiungi il percorso'):next.text;
 setMajorNavHud(true,majorIcon(next.type),label,distText,next.name||'');
 const stage=remaining<=35?'now':remaining<=140?'soon':'';if(stage){const key=next.index+':'+stage+':'+majorNav.mode;if(key!==majorNav.lastVoiceKey){majorNav.lastVoiceKey=key;const msg=stage==='now'?(next.type==='arrive'?label:'Ora, '+label.toLowerCase()):'Tra '+Math.max(30,Math.round(remaining/10)*10)+' metri, '+label.toLowerCase();majorSpeak(msg,key,true)}}
}
async function navigateMajorTo(from,target,mode,label){
 if(!from||!target)return false;majorNav.mode=mode;majorNav.target=target;majorNav.lastVoiceKey='';clearMajorGuide();clearMajorNavLine();
 const rr=window.MajorOfflineNav?.route?.(from,target);if(!rr){setMajorNavHud(true,'⚠','Percorso offline non disponibile','','');$('majorRouteInfo').textContent='Navigazione offline non disponibile: rete locale non connessa a questa posizione';return false}
 majorNav.path=rr.points;majorNav.maneuvers=rr.maneuvers||[];majorNav.pathIndex=0;majorNav.lastRerouteAt=Date.now();navRoute=rr;navLine=L.polyline(rr.points,{color:'#1677ff',weight:8,opacity:.96}).addTo(routeLayer);$('majorRouteInfo').textContent='🧭 '+label+' • '+fmtDist(rr.distance)+' • navigazione offline';transferGuidance(from);majorSpeak(label.toLowerCase(),'start-'+mode,true);return true
}
function startMajorInsideNavigation(ll,accuracy=15){
 clearMajorNavLine();majorNav.mode='inside';majorNav.target=null;majorNav.lastInsideStreet=null;majorNav.lastVoiceKey='';resetTrackingJoin();updateMajorInsideGuide(ll,accuracy)
}
function updateMajorInsideGuide(ll,accuracy=15){
 const moveHeading=majorMovementHeading(ll),best=trackedMajorLanePoint(ll,moveHeading,accuracy);if(!best)return;
 const allowed=majorAllowedDistance(accuracy),lane=majorTracking.lanes[best.laneIndex],street=majorRoutes[majorRi]?.streets?.[best.streetIndex],streetName=street?.name||'percorso assegnato';
 if(best.dist>allowed*1.6){
  clearMajorGuide();
  if(best.stabilizing){setMajorNavHud(true,'⌛','Verifico la posizione',fmtDist(best.dist),'Mantengo la corsia corrente');$('majorRouteInfo').textContent='GPS in verifica • attendo prima di ricalcolare';return}
  setMajorNavHud(true,'⚠','Sei fuori dal percorso',fmtDist(best.dist),'Rientra sulla linea colorata');if(majorNav.lastVoiceKey!=='off-route'){majorNav.lastVoiceKey='off-route';majorSpeak('Sei fuori dal percorso. Rientra sulla linea assegnata.','off-route',true)}return
 }
 majorNav.lastVoiceKey=majorNav.lastVoiceKey==='off-route'?'':majorNav.lastVoiceKey;
 clearMajorGuide();const from=Math.max(0,best.segmentIndex),to=Math.min(lane.coords.length,from+10),ahead=[best.point,...lane.coords.slice(from+1,to)];if(ahead.length>1)majorGuideLine=L.polyline(ahead,{color:'#ff9800',weight:7,opacity:.92}).addTo(routeLayer);
 const nextStreet=majorRoutes[majorRi]?.streets?.[Math.min(best.streetIndex+1,(majorRoutes[majorRi]?.streets?.length||1)-1)]?.name||'';
 setMajorNavHud(true,'↑','Continua su '+streetName,'',nextStreet&&nextStreet!==streetName?'Prossimo tratto: '+nextStreet:'Percorso '+(majorRoutes[majorRi]?.code||''));
 if(majorNav.lastInsideStreet!==best.streetIndex){majorNav.lastInsideStreet=best.streetIndex;majorSpeak('Continua su '+streetName,'inside-'+best.streetIndex,true)}
 $('majorRouteInfo').textContent='🧭 SUL PERCORSO • continua su '+streetName
}
async function beginMajorNavigation(ll,accuracy=15){
 const best=nearestMajorLanePoint(ll,null),allowed=majorAllowedDistance(accuracy);
 if(best&&best.dist<=allowed){startMajorInsideNavigation(ll,accuracy);return}
 if(best?.point){await navigateMajorTo(ll,best.point,'to-start','Verso il punto più vicino del percorso');return}
 $('majorRouteInfo').textContent='Nessun tratto disponibile: impossibile trovare il punto più vicino';
}
async function updateMajorNavigation(ll,accuracy=15,speed=0){
 if(currentMode!=='major'||!mapState.active||mapState.paused||majorNav.mode==='loading')return;
 if(majorNav.mode==='paused'){setMajorNavHud(true,'⏸','Pausa','',majorNav.pausePoint?'Punto di ripresa memorizzato':'');return}
 if(majorNav.mode==='restore-pause'&&majorNav.pausePoint){await navigateMajorTo(ll,majorNav.pausePoint,'to-pause','Ritorno al punto di pausa');return}
 if(majorNav.mode==='restore'){await beginMajorNavigation(ll,accuracy);return}
 if(majorNav.mode==='to-start'||majorNav.mode==='to-pause'){
  if(!majorNav.target)return;const meters=d(ll,majorNav.target),off=distanceToTransferPath(ll),now=Date.now();
  if(off>90&&!majorNav.rerouting&&now-majorNav.lastRerouteAt>12000){majorNav.rerouting=true;majorSpeak('Ricalcolo percorso','reroute-'+now,true);const mode=majorNav.mode,target=[...majorNav.target],label=mode==='to-pause'?'Ritorno al punto di pausa':'Verso inizio percorso';await navigateMajorTo(ll,target,mode,label);majorNav.rerouting=false;return}
  transferGuidance(ll);
  if(meters<=25){const wasPause=majorNav.mode==='to-pause';majorNav.target=null;clearMajorNavLine();if(wasPause&&majorPauseMarker){routeLayer.removeLayer(majorPauseMarker);majorPauseMarker=null}if(wasPause&&majorNav.resumeMode==='to-start'&&majorNav.resumeTarget){const target=[...majorNav.resumeTarget];majorNav.resumeMode=null;majorNav.resumeTarget=null;await navigateMajorTo(ll,target,'to-start','Verso inizio percorso')}else{majorNav.resumeMode=null;majorNav.resumeTarget=null;startMajorInsideNavigation(ll,accuracy);majorSpeak(wasPause?'Sei tornato al punto di pausa. Riprendo il percorso.':'Sei arrivato al percorso. Inizio navigazione interna.','arrived-'+(wasPause?'pause':'route'),true)}}return
 }
 if(majorNav.mode==='inside')updateMajorInsideGuide(ll,accuracy)
}
function qFor(s){return s.replace(/^Trebbo:\s*/,'').replace(/^1° Maggio:\s*/,'').replace(/\s*[–-].*$/,'')+', Castel Maggiore, Bologna, Italia'}
async function geocode(label){const q=qFor(label);if(geoCache[q])return geoCache[q];try{const u='https://nominatim.openstreetmap.org/search?format=json&limit=1&countrycodes=it&q='+encodeURIComponent(q);const a=await fetch(u,{headers:{'Accept-Language':'it'}}).then(x=>x.json());if(a[0])return geoCache[q]=[+a[0].lat,+a[0].lon]}catch(e){}return null}
function stopInstruction(rix,i){return window.ROUTE_INSTRUCTIONS?.notes?.[rix]?.[i]||''}
function escapeText(value){return String(value).replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]))}
function instructionPopup(rix,i){const note=stopInstruction(rix,i);return note?'<br><br><b>Da pulire · elenco operativo</b><br>'+escapeText(note):''}
function renderInstructions(){ $('workInstruction').textContent=stopInstruction(ri,pi);const box=$('routeChecks');box.replaceChildren();const issues=window.ROUTE_INSTRUCTIONS?.issues?.[ri]||[];$('routeChecksDetails').hidden=!issues.length;issues.forEach(note=>{const item=document.createElement('li');item.textContent=note;box.appendChild(item)})}
function surfaceParts(rix,i){const rebuilt=window.REBUILT_MINOR_ROUTES?.[rix]?.[i];if(rebuilt?.parts?.length)return rebuilt.parts;return routes[rix].segments[i].coords?.length?[{coords:routes[rix].segments[i].coords,matched:false,source:'rebuild-pending'}]:[]}
function surfaceStart(rix,i){const parts=surfaceParts(rix,i);const p=parts.find(x=>!x.excluded&&x.coords?.length)||parts.find(x=>x.coords?.length);return p?.coords?.[0]||routes[rix].segments[i].coords?.[0]||null}
const stopEntrances={
 "0-2":[44.5772966,11.3582505],
 "0-3":[44.5777438,11.3580944],
 "0-4":[44.5790124,11.3587355],
 "0-5":[44.5777442,11.3569505],
 "0-6":[44.5782438,11.3568768],
 "0-7":[44.5795743,11.3558244],
 "0-9":[44.5802846,11.3585958], // PAS/Biblioteca: marker provvisorio sul tracciato, sede attuale via Bondanello 39
 "1-3":[44.5741556,11.3598662],
 "1-5":[44.5756664,11.3642595],
 "1-6":[44.575504,11.3649328],
 "1-8":[44.5801351,11.3626448],
 "1-9":[44.5746195,11.3645997],
 "1-12":[44.5750023,11.3654988],
 "2-1":[44.5526334,11.3249917],
 "2-2":[44.5565516,11.3185115],
 "2-4":[44.5545065,11.3179643],
 "2-9":[44.5543218,11.3526711] // tappa materna: ingresso non certificato dal DBTR, marker storico mantenuto finché verificato
};
function stopMarkerPoint(rix,i){return window.REBUILT_MINOR_ROUTES?.[rix]?.[i]?.marker||surfaceStart(rix,i)}
function drawSegment(layer,rix,i,options){
 const s=routes[rix].segments[i],minor=currentMode==='minor';
 const parts=minor?surfaceParts(rix,i):[{coords:s.coords,matched:true}];
 const excluded=window.REBUILT_MINOR_ROUTES?.[rix]?.[i]?.excluded;
 if(minor&&excluded)return;
 parts.forEach(part=>{const src=part.source||'geometria-verificata';const srcLabel=src==='dbtr-acp'?'DBTR • area pedonale':src==='dbtr-aci'?'DBTR • area ciclabile':src==='manual-ortho'?'Ortofoto Regione 2023–24':src==='regional-estimate'?'Regione • stima da area ufficiale':src==='osm'?'OSM • percorso pedonale/ciclabile':src==='civic-road-axis-estimate'?'Asse stradale ricostruito per intervento su carreggiata':src;L.polyline(part.coords,{...options,color:minor&&!part.matched?'#e53935':options.color,weight:minor?Math.min(options.weight,part.matched?3:5):options.weight,dashArray:minor?(part.source==='regional-estimate'?'2 3':!part.matched?'8 7':null):null}).addTo(layer).on('click',e=>{if(routeEditMode==='select-edit'||routeEditMode==='select-delete'){L.DomEvent.stopPropagation(e);selectRouteForAdmin('minor',i,s.name,part.coords)}}).bindPopup('<b>'+s.name+'</b><br>'+routes[rix].name+'<br>'+(minor?srcLabel:'Percorso operativo')+(minor?instructionPopup(rix,i):'')+adminRoutePopup(s.name,'minor',i,part.coords))});
}
function drawAllRoutes(){allRoutesLayer.clearLayers();if(currentMode!=='minor')return;routes.forEach((r,rix)=>r.segments.forEach((s,i)=>{const complete=done.has(rix+'-'+i);drawSegment(allRoutesLayer,rix,i,{color:complete?'#20bd62':'#1976d2',weight:7,opacity:.35})}))}
function drawRoute(focus=true){
 routeLayer.clearLayers();const r=routes[ri],bounds=[];
 const pieces=surfaceParts(ri,pi),matches=pieces.filter(p=>p.matched).length,regional=pieces.filter(p=>p.source==='regional-estimate').length,manual=pieces.some(p=>p.source==='manual-ortho');
 $('routeInfo').textContent=r.name+' • '+r.segments.length+' tratti • '+matches+'/'+pieces.length+' parti adattate nella tappa'+(manual?' • ricalcata su ortofoto 2023–24':'')+(regional?' • '+regional+' stimate da aree Regione':'');
 r.segments.forEach((s,i)=>{
  if(currentMode==='minor'&&window.REBUILT_MINOR_ROUTES?.[ri]?.[i]?.excluded)return;
  const complete=done.has(ri+'-'+i),active=i===pi;
  drawSegment(routeLayer,ri,i,{color:complete?'#20bd62':active?'#ff9800':'#71808a',weight:active?10:7,opacity:.92});
  const point=stopMarkerPoint(ri,i);
  if(point){
   surfaceParts(ri,i).forEach(p=>bounds.push(...(p.coords||[])));const marker=L.marker(point,{icon:L.divIcon({className:'routePin '+(complete?'done ':'')+(active?'active':''),html:'<span>'+(i+1)+'</span>',iconSize:[30,30],iconAnchor:[15,15]})}).addTo(routeLayer);
   marker.bindPopup('<b>'+(i+1)+'. '+s.name+'</b><br>'+r.name+instructionPopup(ri,i));marker.on('click',()=>{pi=i;drawRoute(false);focusStop()});
  }
 });
 if(focus&&bounds.length)map.fitBounds(bounds,{padding:[45,45],maxZoom:17});renderRoute();
}
function focusStop(){const parts=surfaceParts(ri,pi);const bounds=parts.flatMap(p=>p.coords||[]);if(bounds.length)map.fitBounds(bounds,{padding:[40,40],maxZoom:19});else $('routeInfo').textContent='Tappa in ricostruzione DBTR: geometria non ancora pubblicata'}
function renderStops(){const r=routes[ri],box=$('stops');box.innerHTML='';r.points.forEach((label,i)=>{const key=ri+'-'+i,b=document.createElement('button');b.className='stopItem'+(i===pi?' active':'')+(done.has(key)?' done':'');b.innerHTML='<span class="stopNo">'+(done.has(key)?'✓':i+1)+'</span><span class="stopText">'+label+'</span>';b.onclick=()=>{pi=i;drawRoute(false);focusStop()};box.appendChild(b)})}

function renderRoute(){const r=routes[ri],key=ri+'-'+pi,n=r.points.filter((_,i)=>done.has(ri+'-'+i)).length,pct=Math.round(n/r.points.length*100);$('pointName').textContent=r.points[pi];$('progressText').textContent=r.name+' • Tappa '+(pi+1)+'/'+r.points.length+' • '+(r.points.length-n)+' rimanenti';$('progressPct').textContent=pct+'%';$('progressBar').style.width=pct+'%';$('navBanner').textContent=done.has(key)?'🟢 COMPLETATA':skipped.has(key)?'🟡 SALTATA':'🔵 TAPPA ATTUALE • DA ESEGUIRE';renderInstructions();renderStops()}
function majorKey(rix=majorRi,i=majorPi){return rix+'-'+i}
function majorCompletedMeters(r,rix=majorRi){return r.streets.reduce((sum,s,i)=>sum+(majorDone.has(majorKey(rix,i))?s.meters:0),0)}
function renderMajorStreets(){const r=majorRoutes[majorRi],box=$('majorStreets');box.replaceChildren();if(!r)return;r.streets.forEach((s,i)=>{const key=majorKey(majorRi,i),b=document.createElement('button');b.className='stopItem'+(i===majorPi?' active':'')+(majorDone.has(key)?' done':'');b.innerHTML='<span class="stopNo">'+(majorDone.has(key)?'✓':i+1)+'</span><span class="stopText"><span>'+escapeText(s.name)+'</span><span class="streetMeters">'+fmtDist(s.meters)+' • '+fmtDist(s.meters*(r.passes||2))+' operativi</span></span>';b.onclick=()=>{majorPi=i;renderMajor()};box.appendChild(b)})}
function renderMajor(){const r=majorRoutes[majorRi];if(!r){$('majorRouteInfo').textContent='Percorsi maggiori non disponibili';return}if(majorPi>=r.streets.length)majorPi=0;const s=r.streets[majorPi],doneMeters=majorCompletedMeters(r),pct=Math.round(doneMeters/r.totalMeters*100),left=r.totalMeters-doneMeters;$('majorPointName').textContent=s.name;$('majorProgressText').textContent=r.name+' • '+fmtDist(doneMeters*(r.passes||2))+' / '+fmtDist(r.totalMeters*(r.passes||2))+' operativi';$('majorProgressPct').textContent=pct+'%';$('majorProgressBar').style.width=pct+'%';$('majorRouteInfo').textContent='Tratto '+(majorPi+1)+'/'+r.streets.length+' • '+fmtDist(s.meters)+' ufficiali • ~'+fmtDist(s.meters*(r.passes||2))+' da percorrere sulle 2 corsie • ~'+fmtDist(left*(r.passes||2))+' operativi rimanenti';renderMajorStreets()}
function openMajorChooser(canCancel=true){const box=$('majorChooser');if(!box)return;$('majorChooserCancel').hidden=!canCancel;box.hidden=false}
function closeMajorChooser(){const box=$('majorChooser');if(box)box.hidden=true}
async function chooseMajorRoute(index){
 const n=Number(index);if(!Number.isInteger(n)||n<0||n>=majorRoutes.length)return;
 const saved=readMajorWorkSession();if(saved?.active&&saved.routeIndex!==n){alert('Hai già un lavoro in corso su un altro percorso. Premi TERMINA prima di cambiarlo.');return false}
 majorRi=n;majorPi=0;localStorage.setItem('majorSelectedRoute',String(n));majorTracking.routeIndex=null;majorTracking.lanes=[];resetTrackingJoin();routeLayer.clearLayers();closeMajorChooser();renderMajor();await showMajorRoute();await restoreMajorWorkSession();return true
}
function initMajorRouteSelection(){document.querySelectorAll('[data-major-route]').forEach(b=>b.onclick=()=>chooseMajorRoute(b.dataset.majorRoute));$('majorChooserCancel').onclick=closeMajorChooser;$('majorChangeRoute').onclick=()=>{setMajorMore(false);if(mapState.active){alert('Termina il lavoro prima di cambiare percorso.');return}openMajorChooser(true)}}
let startupMode=null;
function showStartupModeStep(){$('startupModeStep').hidden=false;$('startupRouteStep').hidden=true;startupMode=null}
function showStartupRouteStep(mode){startupMode=mode;$('startupModeStep').hidden=true;$('startupRouteStep').hidden=false;const major=mode==='major';$('startupRouteTitle').textContent=major?'Scegli il percorso maggiore':'Scegli il percorso minore';$('startupRouteSubtitle').textContent=major?'Rosso, Blu o Giallo':'Squadra / zona assegnata';$('startupMajorRoutes').hidden=!major;$('startupMinorRoutes').hidden=major}
async function finishStartupSelection(mode,index){const box=$('startupChooser'),saved=readMajorWorkSession();if(saved?.active&&mode!=='major'){alert('Hai un lavoro di viabilità maggiore ancora in corso. Seleziona quel percorso e premi TERMINA prima di cambiare viabilità.');return}if(mode==='major'){const ok=await chooseMajorRoute(index);if(ok===false)return;box.hidden=true;document.querySelector('.tab[data-mode="major"]')?.click()}else{const n=Number(index);if(!Number.isInteger(n)||n<0||n>=routes.length)return;ri=n;pi=0;localStorage.setItem('minorSelectedRoute',String(n));box.hidden=true;document.querySelector('.tab[data-mode="minor"]')?.click();drawRoute(false)}}
function initStartupSelection(){document.querySelectorAll('[data-start-mode]').forEach(b=>{b.onclick=e=>{e.preventDefault();e.stopPropagation();showStartupRouteStep(b.dataset.startMode)};b.addEventListener('touchend',e=>{e.preventDefault();showStartupRouteStep(b.dataset.startMode)},{passive:false})});document.querySelectorAll('[data-start-major-route]').forEach(b=>b.onclick=()=>finishStartupSelection('major',b.dataset.startMajorRoute));document.querySelectorAll('[data-start-minor-route]').forEach(b=>b.onclick=()=>finishStartupSelection('minor',b.dataset.startMinorRoute));$('startupBack').onclick=showStartupModeStep;showStartupModeStep();$('startupChooser').hidden=false}

function majorSearchName(name){return name.replace(/\s*\+.*$/,'').replace(/\s*\(.*$/,'').replace(/^PARCH\.\s*/i,'').replace(/^PARCHEGGIO\s+/i,'').replace(/^PARCHEGGI\s+/i,'').trim()}
let majorRoadNetwork=[],majorBoundary=null,majorBoundaryLayer=null;
function normRoadName(name){return String(name||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\bf\.\s*lli\b/g,'fratelli').replace(/\bxxv\b/g,'25').replace(/\biv\b/g,'4').replace(/1\s*°?/g,'primo ').replace(/[^a-z0-9 ]+/g,' ').replace(/\b(via|viale|piazza|circonvallazione)\b/g,' ').replace(/\b[a-z]\b/g,' ').replace(/\s+/g,' ').trim()}
function roadNameScore(a,b){const A=new Set(normRoadName(a).split(' ').filter(Boolean)),B=new Set(normRoadName(b).split(' ').filter(Boolean));if(!A.size||!B.size)return 0;let hit=0;A.forEach(x=>{if(B.has(x))hit++});return hit/Math.max(A.size,B.size)}
async function loadMajorBoundary(){
 if(majorBoundary)return majorBoundary;
 const geometry=window.MajorOfflineNav?.data?.boundary||window.MAJOR_OFFLINE?.boundary;
 if(!geometry)throw new Error('Confine offline di Castel Maggiore non ancora incluso');
 majorBoundary=geometry;
 if(!majorBoundaryLayer){majorBoundaryLayer=L.geoJSON({type:'Feature',properties:{name:'Castel Maggiore'},geometry},{style:{color:'#111',weight:3,opacity:.8,fill:false,dashArray:'8 6'},interactive:false}).addTo(map)}
 return majorBoundary
}
function pointInRing(p,ring){const x=p[1],y=p[0];let inside=false;for(let i=0,j=ring.length-1;i<ring.length;j=i++){const xi=ring[i][0],yi=ring[i][1],xj=ring[j][0],yj=ring[j][1],cross=((yi>y)!==(yj>y))&&(x<(xj-xi)*(y-yi)/((yj-yi)||1e-12)+xi);if(cross)inside=!inside}return inside}
function pointInBoundary(p){if(!majorBoundary)return false;const polys=majorBoundary.type==='Polygon'?[majorBoundary.coordinates]:majorBoundary.coordinates;for(const poly of polys){if(!poly?.[0]||!pointInRing(p,poly[0]))continue;let inHole=false;for(let i=1;i<poly.length;i++)if(pointInRing(p,poly[i])){inHole=true;break}if(!inHole)return true}return false}
function boundaryEdges(){if(!majorBoundary)return[];const polys=majorBoundary.type==='Polygon'?[majorBoundary.coordinates]:majorBoundary.coordinates,edges=[];for(const poly of polys)for(const ring of poly)for(let i=1;i<ring.length;i++)edges.push([[ring[i-1][1],ring[i-1][0]],[ring[i][1],ring[i][0]]]);return edges}
function segmentIntersectionT(a,b,c,d){const ax=a[1],ay=a[0],bx=b[1],by=b[0],cx=c[1],cy=c[0],dx=d[1],dy=d[0],rx=bx-ax,ry=by-ay,sx=dx-cx,sy=dy-cy,den=rx*sy-ry*sx;if(Math.abs(den)<1e-12)return null;const qx=cx-ax,qy=cy-ay,t=(qx*sy-qy*sx)/den,u=(qx*ry-qy*rx)/den;return t>=0&&t<=1&&u>=0&&u<=1?t:null}
function interp(a,b,t){return[a[0]+(b[0]-a[0])*t,a[1]+(b[1]-a[1])*t]}
function clipMajorCoords(coords){
 if(!majorBoundary||!coords?.length)return[];const edges=boundaryEdges(),parts=[],pushPart=p=>{if(p.length>1)parts.push(p)};let current=[];
 for(let i=1;i<coords.length;i++){const a=coords[i-1],b=coords[i],ts=[0,1];for(const e of edges){const t=segmentIntersectionT(a,b,e[0],e[1]);if(t!==null&&t>1e-8&&t<1-1e-8)ts.push(t)}ts.sort((x,y)=>x-y);
  for(let k=1;k<ts.length;k++){const t0=ts[k-1],t1=ts[k],m=interp(a,b,(t0+t1)/2);if(pointInBoundary(m)){const p0=interp(a,b,t0),p1=interp(a,b,t1);if(!current.length)current.push(p0);else{const last=current[current.length-1];if(Math.abs(last[0]-p0[0])>1e-8||Math.abs(last[1]-p0[1])>1e-8){pushPart(current);current=[p0]}}current.push(p1)}else if(current.length){pushPart(current);current=[]}}
 }if(current.length)pushPart(current);return parts
}
async function loadMajorRoadNetwork(){
 await loadMajorBoundary();
 const roads=window.MajorOfflineNav?.data?.roads||window.MAJOR_OFFLINE?.roads||[];
 if(!roads.length)throw new Error('Reticolo stradale offline non ancora generato');
 majorRoadNetwork=roads;return majorRoadNetwork
}
function findMajorWaysByName(value){
 const network=majorRoadNetwork||[],target=majorSearchName(value),n=normRoadName(target);if(!network.length||!n)return[];
 let exact=network.filter(w=>normRoadName(w.name)===n);if(exact.length)return exact;
 const scored=network.map(w=>({w,score:roadNameScore(target,w.name)})).sort((a,b)=>b.score-a.score),best=scored[0]?.score||0;if(best<.5)return[];
 const winner=normRoadName(scored[0].w.name);return network.filter(w=>normRoadName(w.name)===winner)
}
function findMajorWays(street){
 if(street._adminDeleted)return [];
 if(street._adminCoords?.length)return [{id:'admin-'+street.name,coords:street._adminCoords}];
 const parking=window.MAJOR_PARKING?.[street.name],parkingOnly=/^PARCH(?:\.|EGGI)/i.test(street.name);
 const aliases=Array.isArray(street.searchNames)&&street.searchNames.length?street.searchNames:[street.searchName||street.name],seen=new Set(),out=[];
 if(!parkingOnly)for(const alias of aliases)for(const w of findMajorWaysByName(alias)){if(!seen.has(w.id)){seen.add(w.id);out.push(w)}}
 for(const w of parking?.ways||[])if(!seen.has(w.id)){seen.add(w.id);out.push(w)}
 return out
}
function majorParkingAreas(street){return street?._adminDeleted||street?._adminCoords?.length?[]:window.MAJOR_PARKING?.[street?.name]?.areas||[]}
function drawMajorParking(r,s,active=true){
 const bounds=[],color=majorColor(r);
 for(const area of majorParkingAreas(s)){L.polygon(area.coords,{color,weight:active?4:3,opacity:1,fillOpacity:.12}).addTo(routeLayer).bindPopup('<b>'+escapeText(s.name)+'</b><br>Parcheggio da lavorare: accessi e manovre inclusi');bounds.push(...area.coords)}
 return bounds
}
function nearestMajorParkingPoint(p){
 let best=null;
 for(let i=0;i<majorTracking.lanes.length;i++){const lane=majorTracking.lanes[i];if(!lane.area)continue;const ring=lane.coords.map(q=>[q[1],q[0]]);if(!pointInRing(p,ring))continue;
  const candidate={point:p.slice(),dist:0,score:0,bearing:0,laneIndex:i,segmentIndex:0,streetIndex:lane.streetIndex,direction:0,parking:true};
  if(!best||/^PARCH(?:\.|EGGI)/i.test(majorRoutes[majorRi]?.streets?.[lane.streetIndex]?.name||''))best=candidate;
 }return best
}
function offsetRoad(coords,meters){
 return coords.map((p,i)=>{const a=coords[Math.max(0,i-1)],b=coords[Math.min(coords.length-1,i+1)],lat=p[0]*Math.PI/180,dx=(b[1]-a[1])*111320*Math.cos(lat),dy=(b[0]-a[0])*110540,len=Math.hypot(dx,dy)||1,nx=-dy/len,ny=dx/len;return[p[0]+ny*meters/110540,p[1]+nx*meters/(111320*Math.cos(lat))]})
}
function roadBearing(coords){if(coords.length<2)return 0;const a=coords[Math.max(0,Math.floor(coords.length/2)-1)],b=coords[Math.min(coords.length-1,Math.floor(coords.length/2)+1)],y=Math.sin((b[1]-a[1])*Math.PI/180)*Math.cos(b[0]*Math.PI/180),x=Math.cos(a[0]*Math.PI/180)*Math.sin(b[0]*Math.PI/180)-Math.sin(a[0]*Math.PI/180)*Math.cos(b[0]*Math.PI/180)*Math.cos((b[1]-a[1])*Math.PI/180);return(Math.atan2(y,x)*180/Math.PI+360)%360}
function addMajorArrow(coords,color){
 if(coords.length<2)return;const i=Math.floor(coords.length/2),p=coords[i],deg=roadBearing(coords)-90;L.marker(p,{interactive:false,icon:L.divIcon({className:'majorArrow',html:'<span style="color:'+color+';transform:rotate('+deg+'deg)">➤</span>',iconSize:[24,24],iconAnchor:[12,12]})}).addTo(routeLayer)
}
function majorColor(r){return r.code==='rosso'?'#d83f3f':r.code==='blu'?'#1976d2':'#d4a900'}
function drawMajorWayPair(coords,r,s,active=true,parkingWay=false){
 if(coords.length<2)return [];
 if(parkingWay){const parts=clipMajorCoords(coords);for(const part of parts)L.polyline(part,{color:majorColor(r),weight:active?6:4,opacity:1}).addTo(routeLayer).bindPopup('<b>'+escapeText(s.name)+'</b><br>Accesso e corsello del parcheggio');return parts}
 const color=majorColor(r),streetIndex=majorRoutes[majorRi]?.streets?.indexOf(s)??-1,popup='<b>'+escapeText(s.name)+'</b><br>'+r.name+' • '+fmtDist(s.meters)+'<br><b>Doppio passaggio:</b> una corsia per senso di marcia<br><b>Limite:</b> solo Comune di Castel Maggiore',drawn=[];
 const centers=clipMajorCoords(coords);for(const center of centers){const aParts=clipMajorCoords(offsetRoad(center,-2.4)),bParts=clipMajorCoords(offsetRoad(center,2.4)).map(x=>x.slice().reverse());
  for(const a of aParts){L.polyline(a,{color,weight:active?6:4,opacity:active?1:.7}).addTo(routeLayer).on('click',e=>{if(routeEditMode==='select-edit'||routeEditMode==='select-delete'){L.DomEvent.stopPropagation(e);selectRouteForAdmin('major',majorRoutes[majorRi]?.streets?.indexOf(s)??-1,s.name,a)}}).bindPopup(popup+adminRoutePopup(s.name,'major',streetIndex,a));drawn.push(a);if(active)addMajorArrow(a,color)}
  for(const b of bParts){L.polyline(b,{color,weight:active?6:4,opacity:active?1:.7,dashArray:'10 7'}).addTo(routeLayer).on('click',e=>{if(routeEditMode==='select-edit'||routeEditMode==='select-delete'){L.DomEvent.stopPropagation(e);selectRouteForAdmin('major',majorRoutes[majorRi]?.streets?.indexOf(s)??-1,s.name,b)}}).bindPopup(popup+adminRoutePopup(s.name,'major',streetIndex,b));drawn.push(b);if(active)addMajorArrow(b,color)}
 }return drawn
}
async function fallbackMajorPoint(r,s){await loadMajorRoadNetwork();const target=majorSearchName(s.searchName||s.name),scored=majorRoadNetwork.map(w=>({w,score:roadNameScore(target,w.name)})).sort((a,b)=>b.score-a.score),w=scored[0]?.score>=.35?scored[0].w:null;if(!w?.coords?.length)return null;const p=w.coords[Math.floor(w.coords.length/2)];L.circleMarker(p,{radius:11,color:'#fff',weight:4,fillColor:majorColor(r),fillOpacity:1}).addTo(routeLayer).bindPopup('<b>'+escapeText(s.name)+'</b><br>'+r.name+' • '+fmtDist(s.meters)+'<br>Posizione ricavata dal reticolo offline');return p}
async function focusMajorStreet(){
 const r=majorRoutes[majorRi],s=r?.streets?.[majorPi];if(!s)return;routeLayer.clearLayers();$('majorRouteInfo').textContent='Carico la strada e preparo andata + ritorno…';
 try{await loadMajorRoadNetwork();const ways=findMajorWays(s),bounds=drawMajorParking(r,s,true);ways.forEach(w=>drawMajorWayPair(w.coords,r,s,true,w.parking).forEach(x=>bounds.push(...x)));if(bounds.length){map.fitBounds(bounds,{padding:[35,35],maxZoom:17});$('majorRouteInfo').textContent=s.name+' • '+fmtDist(s.meters)+(window.MAJOR_PARKING?.[s.name]?' • parcheggio e accessi inclusi':' • doppio passaggio su entrambe le corsie')}else{const p=await fallbackMajorPoint(r,s);if(p){map.setView(p,17);$('majorRouteInfo').textContent='Strada localizzata; asse OSM da verificare prima del doppio passaggio'}else $('majorRouteInfo').textContent='Strada non trovata automaticamente • '+s.name}}catch(e){const p=await fallbackMajorPoint(r,s);if(p)map.setView(p,17);$('majorRouteInfo').textContent='Reticolo stradale non disponibile: mostrata posizione indicativa'}
}
async function showMajorRoute(){
 const r=majorRoutes[majorRi];if(!r)return;routeLayer.clearLayers();$('majorRouteInfo').textContent='Carico '+r.name+' e preparo entrambe le corsie…';
 try{await loadMajorRoadNetwork();const bounds=[];let matched=0;
  r.streets.forEach((s,i)=>{if(s._adminDeleted)return;const ways=findMajorWays(s),areas=drawMajorParking(r,s,i===majorPi);bounds.push(...areas);if(ways.length||areas.length){matched++;ways.forEach(w=>drawMajorWayPair(w.coords,r,s,i===majorPi,w.parking).forEach(x=>bounds.push(...x))) }});
  if(bounds.length)map.fitBounds(bounds,{padding:[30,30],maxZoom:15});
  $('majorRouteInfo').textContent=r.name+' • '+matched+'/'+r.streets.length+' voci agganciate al reticolo stradale • corsie, accessi e parcheggi inclusi';
 }catch(e){$('majorRouteInfo').textContent='Reticolo offline non disponibile • '+(e?.message||'file mancante')}
}

map.createPane('majorCoveragePane',map.getPane('rotatePane'));map.getPane('majorCoveragePane').style.zIndex='450';map.getPane('majorCoveragePane').style.pointerEvents='none';
const majorCoverageLayer=L.layerGroup().addTo(map);
const MAJOR_SESSION_KEY='vargaMajorWorkSessionV1';let majorCoverageSegments=[],lastMajorSessionSave=0;
const majorTracking={routeIndex:null,lanes:[],lastRaw:null,lastSnap:null,lastLane:null,coveredMeters:0,lockedLane:null,lockedStreet:null,candidateLane:null,candidateHits:0,offRouteSince:0};
function headingDeg(a,b){const y=Math.sin((b[1]-a[1])*Math.PI/180)*Math.cos(b[0]*Math.PI/180),x=Math.cos(a[0]*Math.PI/180)*Math.sin(b[0]*Math.PI/180)-Math.sin(a[0]*Math.PI/180)*Math.cos(b[0]*Math.PI/180)*Math.cos((b[1]-a[1])*Math.PI/180);return(Math.atan2(y,x)*180/Math.PI+360)%360}
function angleDiff(a,b){const x=Math.abs(a-b)%360;return Math.min(x,360-x)}
function nearestOnSegment(p,a,b){const lat=p[0]*Math.PI/180,kx=111320*Math.cos(lat),ky=110540,ax=a[1]*kx,ay=a[0]*ky,bx=b[1]*kx,by=b[0]*ky,px=p[1]*kx,py=p[0]*ky,dx=bx-ax,dy=by-ay,l2=dx*dx+dy*dy,t=l2?Math.max(0,Math.min(1,((px-ax)*dx+(py-ay)*dy)/l2)):0,qx=ax+t*dx,qy=ay+t*dy;return{point:[qy/ky,qx/kx],dist:Math.hypot(px-qx,py-qy),bearing:headingDeg(a,b),t}}
function nearestMajorLanePoint(p,moveHeading=null){
 let best=null;
 const parking=typeof nearestMajorParkingPoint==='function'?nearestMajorParkingPoint(p):null;if(parking)return parking;
 for(let li=0;li<majorTracking.lanes.length;li++){const lane=majorTracking.lanes[li],pts=lane.coords;if(lane.area)continue;for(let i=1;i<pts.length;i++){const q=nearestOnSegment(p,pts[i-1],pts[i]);const dirPenalty=moveHeading==null||lane.parking?0:angleDiff(moveHeading,q.bearing)*.08,score=q.dist+dirPenalty+(lane.parking?3:0);if(!best||score<best.score)best={...q,score,laneIndex:li,segmentIndex:i-1,streetIndex:lane.streetIndex,direction:lane.direction,parking:!!lane.parking}}}
 return best
}
function nearestMajorPointOnLane(p,laneIndex,moveHeading=null){
 const lane=majorTracking.lanes[laneIndex];if(!lane?.coords?.length||lane.area)return null;let best=null,pts=lane.coords;
 for(let i=1;i<pts.length;i++){const q=nearestOnSegment(p,pts[i-1],pts[i]),dirPenalty=moveHeading==null?0:angleDiff(moveHeading,q.bearing)*.08,score=q.dist+dirPenalty+(lane.parking?3:0);if(!best||score<best.score)best={...q,score,laneIndex,segmentIndex:i-1,streetIndex:lane.streetIndex,direction:lane.direction,parking:!!lane.parking}}
 return best
}
function lockMajorLane(best){
 if(!best)return null;majorTracking.lockedLane=best.laneIndex;majorTracking.lockedStreet=best.streetIndex;majorTracking.candidateLane=null;majorTracking.candidateHits=0;majorTracking.offRouteSince=0;return best
}
function trackedMajorLanePoint(p,moveHeading=null,accuracy=15){
 const nearest=nearestMajorLanePoint(p,moveHeading);if(!nearest)return null;if(nearest.parking)return lockMajorLane(nearest);
 const locked=nearestMajorPointOnLane(p,majorTracking.lockedLane,moveHeading),allowed=majorAllowedDistance(accuracy);
 // Keep small GPS drift stable, but switch immediately on a turn or a better road match.
 if(locked&&locked.dist<=allowed&&(moveHeading==null||angleDiff(moveHeading,locked.bearing)<90)&&locked.score<=nearest.score+3)return locked;
 return lockMajorLane(nearest)
}
function majorMovementHeading(ll){
 const anchor=majorTracking.headingRaw||majorTracking.lastRaw;
 if(anchor&&d(anchor,ll)>=3){majorTracking.moveHeading=headingDeg(anchor,ll);majorTracking.headingRaw=ll}
 else if(!anchor)majorTracking.headingRaw=ll;
 return majorTracking.moveHeading??null
}
async function prepareMajorTrackingRoute(){
 const r=majorRoutes[majorRi];if(!r)return false;
 if(majorTracking.routeIndex===majorRi&&majorTracking.lanes.length)return true;
 await loadMajorRoadNetwork();const lanes=[];
 r.streets.forEach((s,streetIndex)=>{
  if(s._adminDeleted)return;
  findMajorWays(s).forEach(w=>clipMajorCoords(w.coords).forEach(center=>{
   if(w.parking){lanes.push({coords:center,streetIndex,direction:0,parking:true});return}
   clipMajorCoords(offsetRoad(center,-2.4)).forEach(coords=>lanes.push({coords,streetIndex,direction:1}));
   clipMajorCoords(offsetRoad(center,2.4)).forEach(coords=>lanes.push({coords:coords.slice().reverse(),streetIndex,direction:-1}));
  }));
  for(const area of majorParkingAreas(s))lanes.push({coords:area.coords,streetIndex,direction:0,area:true,parking:true});
 });
 majorTracking.routeIndex=majorRi;majorTracking.lanes=lanes;majorTracking.lastRaw=null;majorTracking.lastSnap=null;majorTracking.lastLane=null;return lanes.length>0
}
function resetTrackingJoin(){majorTracking.lastRaw=null;majorTracking.lastSnap=null;majorTracking.lastLane=null;majorTracking.lockedLane=null;majorTracking.lockedStreet=null;majorTracking.candidateLane=null;majorTracking.candidateHits=0;majorTracking.offRouteSince=0;majorTracking.headingRaw=null;majorTracking.moveHeading=null}

const STREET_PENDING_KEY='snowStreetWorkPendingV2';
function snowUserStorageKey(key){
 const uid=window.VargaSnowCloud?.user?.uid;
 return uid?key+'::'+uid:null;
}
let streetWriteChain=Promise.resolve(),streetFlushPromise=null,streetCloudError=null;
let streetWorkTourId='',streetWorkState={},activeStreetIndex=null,lastStreetFlush=0;
function newStreetTourId(){return 'tour-'+Date.now().toString(36)+'-'+Math.random().toString(36).slice(2,8)}
function beginStreetWorkTour(){if(!streetWorkTourId)streetWorkTourId=newStreetTourId();if(!streetWorkState||typeof streetWorkState!=='object')streetWorkState={}}
function streetTotalMeters(index){const route=majorRoutes[majorRi],street=route?.streets?.[index];return Math.max(1,(Number(street?.meters)||0)*(Number(route?.passes)||2))}
function streetRecord(index){
 beginStreetWorkTour();const route=majorRoutes[majorRi],street=route?.streets?.[index];if(!street)return null;
 const key=String(index);if(!streetWorkState[key])streetWorkState[key]={recordId:streetWorkTourId+'-'+majorRi+'-'+index,tourId:streetWorkTourId,routeIndex:majorRi,routeCode:route.code||'',routeName:route.name||('Percorso '+(majorRi+1)),streetIndex:index,streetName:street.name,startAtClient:new Date().toISOString(),endAtClient:null,coveredMeters:0,totalMeters:streetTotalMeters(index),completionPercent:0,pauseEvents:[],status:'in_corso',lastActivityAtClient:new Date().toISOString()};
 return streetWorkState[key]
}
function readStreetPending(){try{const key=snowUserStorageKey(STREET_PENDING_KEY);return key?JSON.parse(localStorage.getItem(key)||'{}'):{} }catch(e){return{}}}
function writeStreetPending(data){try{const key=snowUserStorageKey(STREET_PENDING_KEY);if(!key)throw new Error('Non autenticato');localStorage.setItem(key,JSON.stringify(data))}catch(e){console.warn('Street queue save',e);throw e}}
function queueStreetRecord(record){
 if(!record)return Promise.resolve();
 record.lastActivityAtClient=new Date().toISOString();
 const id=record.recordId, snapshot=JSON.parse(JSON.stringify(record));
 // Emergency journal: kept until IndexedDB has durably committed.
 const pending=readStreetPending();pending[id]=snapshot;writeStreetPending(pending);
 streetWriteChain=streetWriteChain.catch(()=>{}).then(async()=>{
  if(!window.VargaSnowOutbox)throw new Error('Archivio offline non inizializzato');
  await window.VargaSnowOutbox.enqueue(id,snapshot);
  const latest=readStreetPending();
  if(JSON.stringify(latest[id])===JSON.stringify(snapshot)){delete latest[id];writeStreetPending(latest)}
 }).catch(e=>{streetCloudError=e;console.error('Salvataggio offline non riuscito',e);$('gps').textContent='⚠ Registro non salvato: non chiudere l’app'});
 saveMajorWorkSession();
 return streetWriteChain;
}
async function flushStreetWorkQueue(force=false){
 if(streetFlushPromise)return streetFlushPromise;
 const now=Date.now();if(!force&&now-lastStreetFlush<15000)return false;
 lastStreetFlush=now;
 streetFlushPromise=(async()=>{
  await streetWriteChain;
  if(Object.keys(readStreetPending()).length){
   const pending=readStreetPending();
   for(const [id,payload] of Object.entries(pending)){
    await window.VargaSnowOutbox.enqueue(id,payload);
    const latest=readStreetPending();
    if(JSON.stringify(latest[id])===JSON.stringify(payload)){delete latest[id];writeStreetPending(latest)}
   }
  }
  const cloud=window.VargaSnowCloud;
  if(!cloud?.user||!cloud?.saveStreetProgress||!navigator.onLine)return false;
  const left=await window.VargaSnowOutbox.flush((id,payload)=>cloud.saveStreetProgress(id,payload));
  streetCloudError=null;return left===0&&Object.keys(readStreetPending()).length===0;
 })().catch(e=>{streetCloudError=e;console.warn('Registro in attesa di sincronizzazione',e);return false}).finally(()=>{streetFlushPromise=null});
 return streetFlushPromise;
}
function switchStreetWork(index){
 if(!Number.isInteger(index)||index<0)return;if(index===activeStreetIndex)return;
 if(activeStreetIndex!==null){const previous=streetRecord(activeStreetIndex);if(previous&&previous.status!=='completata'){previous.status=previous.completionPercent>=100?'completata':'parziale';if(previous.status==='completata'&&!previous.endAtClient)previous.endAtClient=new Date().toISOString();queueStreetRecord(previous)}}
 activeStreetIndex=index;const current=streetRecord(index);if(current&&current.status!=='completata'){current.status=mapState.paused?'in_pausa':'in_corso';queueStreetRecord(current)}flushStreetWorkQueue();
}
function recordStreetDistance(index,meters){
 const record=streetRecord(index);if(!record||!Number.isFinite(meters)||meters<=0)return;
 record.coveredMeters=+(Number(record.coveredMeters||0)+meters).toFixed(1);record.totalMeters=streetTotalMeters(index);record.completionPercent=Math.min(100,Math.round(record.coveredMeters/record.totalMeters*1000)/10);
 if(record.completionPercent>=100){record.status='completata';record.endAtClient=record.endAtClient||new Date().toISOString()}else record.status='in_corso';
 queueStreetRecord(record);flushStreetWorkQueue();
}
function pauseStreetWork(){
 if(activeStreetIndex===null)return;const record=streetRecord(activeStreetIndex);if(!record||record.status==='completata')return;
 const now=new Date().toISOString(),pauses=Array.isArray(record.pauseEvents)?record.pauseEvents:[];if(!pauses.length||pauses.at(-1).resumeAtClient)pauses.push({pauseAtClient:now,resumeAtClient:null});
 record.pauseEvents=pauses;record.status='in_pausa';queueStreetRecord(record);flushStreetWorkQueue(true)
}
function resumeStreetWork(){
 if(activeStreetIndex===null)return;const record=streetRecord(activeStreetIndex);if(!record||record.status==='completata')return;
 const pauses=Array.isArray(record.pauseEvents)?record.pauseEvents:[];if(pauses.length&&!pauses.at(-1).resumeAtClient)pauses.at(-1).resumeAtClient=new Date().toISOString();
 record.pauseEvents=pauses;record.status='in_corso';queueStreetRecord(record);flushStreetWorkQueue(true)
}
function closeStreetWork(){
 if(activeStreetIndex!==null){const record=streetRecord(activeStreetIndex);if(record){const pauses=Array.isArray(record.pauseEvents)?record.pauseEvents:[];if(pauses.length&&!pauses.at(-1).resumeAtClient)pauses.at(-1).resumeAtClient=new Date().toISOString();record.pauseEvents=pauses;record.status=record.completionPercent>=100?'completata':'parziale';if(record.status==='completata')record.endAtClient=record.endAtClient||new Date().toISOString();queueStreetRecord(record)}}
 activeStreetIndex=null;flushStreetWorkQueue(true)
}
function minutesBetween(a,b){const x=a?new Date(a):null,y=b?new Date(b):null;return x&&y&&Number.isFinite(y-x)?Math.max(0,Math.round((y-x)/60000)):0}
async function exportStreetWorkExcel(){
 if(window.VargaSnowCloud?.role!=='admin')return;
 if(typeof XLSX==='undefined'){alert('Esportazione Excel non disponibile: collega Internet e riapri l’app una volta.');return}
 try{
  await flushStreetWorkQueue(true);const rows=await window.VargaSnowCloud.loadStreetWork(),wb=XLSX.utils.book_new();
  for(let n=0;n<3;n++){
   const rr=rows.filter(x=>Number(x.routeIndex)===n).map(x=>{const a=x.startAtClient?new Date(x.startAtClient):null,b=x.endAtClient?new Date(x.endAtClient):null,pauses=Array.isArray(x.pauseEvents)?x.pauseEvents:[],pauseMin=pauses.reduce((sum,p)=>sum+minutesBetween(p.pauseAtClient,p.resumeAtClient),0),elapsed=a&&b?Math.max(0,Math.round((b-a)/60000)):'';return{'Data':a?a.toLocaleDateString('it-IT'):'','Percorso':x.routeName||'','Strada / tratta':x.streetName||'','Ora inizio':a?a.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit',second:'2-digit'}):'','Ora fine':b?b.toLocaleTimeString('it-IT',{hour:'2-digit',minute:'2-digit',second:'2-digit'}):'','N. pause':pauses.length,'Durata pause (min)':pauseMin,'Tempo effettivo (min)':elapsed===''?'':Math.max(0,elapsed-pauseMin),'Metri percorsi':Math.round(Number(x.coveredMeters)||0),'Completamento %':Number(x.completionPercent)||0,'Stato':x.status||'','Operatore':x.userEmail||''}});
   const empty={'Data':'','Percorso':'','Strada / tratta':'','Ora inizio':'','Ora fine':'','N. pause':'','Durata pause (min)':'','Tempo effettivo (min)':'','Metri percorsi':'','Completamento %':'','Stato':'','Operatore':''},ws=XLSX.utils.json_to_sheet(rr.length?rr:[empty]);ws['!cols']=[{wch:12},{wch:20},{wch:34},{wch:12},{wch:12},{wch:10},{wch:18},{wch:20},{wch:16},{wch:17},{wch:14},{wch:28}];XLSX.utils.book_append_sheet(wb,ws,['Rosso','Blu','Giallo'][n])
  }
  XLSX.writeFile(wb,'Registro-Piano-Neve-'+new Date().toISOString().slice(0,10)+'.xlsx')
 }catch(e){console.error(e);alert('Impossibile creare il registro Excel. Controlla la connessione e riprova.')}
}
window.addEventListener('online',()=>flushStreetWorkQueue(true));setInterval(()=>flushStreetWorkQueue(),20000);

function majorCoveragePath(best,rawDistance,allowed){
 const a=majorTracking.lastSnap,b=best.point;if(!a)return [];
 if(!Number.isFinite(rawDistance)||rawDistance>100||d(a,b)>rawDistance+allowed*2)return [];
 const lane=majorTracking.lanes[best.laneIndex];let points=[a,b];
 if(!best.parking&&majorTracking.lastLane===best.laneIndex&&lane?.coords?.length){
  const previous=nearestMajorPointOnLane(a,best.laneIndex),from=previous?.segmentIndex,to=best.segmentIndex;
  if(Number.isInteger(from)&&Number.isInteger(to)){
   const middle=from<=to?lane.coords.slice(from+1,to+1):lane.coords.slice(to+1,from+1).reverse();
   const curved=[a,...middle,b],length=curved.slice(1).reduce((sum,p,i)=>sum+d(curved[i],p),0);
   if(length<=rawDistance*1.8+12)points=curved;
  }
 }else if(!best.parking){
  // Connect adjoining street fragments through their shared junction when available.
  const previousLane=majorTracking.lanes[majorTracking.lastLane];
  const previous=nearestMajorPointOnLane(a,majorTracking.lastLane);
  if(previousLane?.coords?.length&&lane?.coords?.length&&previous){
   const tail=[a,...previousLane.coords.slice(previous.segmentIndex+1)],head=[...lane.coords.slice(0,best.segmentIndex+1),b];
   const joined=[...tail,...head],length=joined.slice(1).reduce((sum,p,i)=>sum+d(joined[i],p),0);
   if(d(tail.at(-1),head[0])<=12&&length<=rawDistance*1.8+12)points=joined;
  }
 }
 // Never bridge an off-route shortcut, paused travel, or a lost GPS position.
 for(let i=1;i<points.length;i++){
  const steps=Math.max(1,Math.ceil(d(points[i-1],points[i])/4));
  for(let n=1;n<steps;n++){
   const t=n/steps,p=points[i-1].map((v,k)=>v+(points[i][k]-v)*t),near=nearestMajorLanePoint(p);
   if(!near||near.dist>allowed)return [];
  }
 }
 return points
}
function markMajorPassed(ll,accuracy){
 if(currentMode!=='major'||!mapState.active||mapState.paused||majorNav.mode==='loading'||!majorTracking.lanes.length||!Number.isFinite(accuracy)||accuracy>35)return false;
 const previousRaw=majorTracking.lastRaw,moveHeading=majorMovementHeading(ll),best=trackedMajorLanePoint(ll,moveHeading,accuracy);majorTracking.lastRaw=ll;
 const allowed=majorAllowedDistance(accuracy);if(!best||best.dist>allowed){resetTrackingJoin();$('majorRouteInfo').textContent='⚪ Fuori percorso • spostamento NON segnato come passato';return false}
 const points=majorCoveragePath(best,previousRaw?d(previousRaw,ll):Infinity,allowed);let added=false;
 for(let i=1;i<points.length;i++){
  const a=points[i-1],b=points[i],meters=d(a,b);if(meters<.1)continue;
  L.polyline([a,b],{pane:'majorCoveragePane',color:'#20bd62',weight:7,opacity:1,lineCap:'round',interactive:false}).addTo(majorCoverageLayer);
  majorTracking.coveredMeters+=meters;recordStreetDistance(best.streetIndex,meters);majorCoverageSegments.push([majorRi,+a[0].toFixed(6),+a[1].toFixed(6),+b[0].toFixed(6),+b[1].toFixed(6)]);added=true;
 }
 majorTracking.lastSnap=best.point;majorTracking.lastLane=best.laneIndex;if(added)saveMajorWorkSession();
 switchStreetWork(best.streetIndex);const street=majorRoutes[majorRi]?.streets?.[best.streetIndex];$('majorRouteInfo').textContent='🧭 SUL PERCORSO • '+(street?.name||'tratto')+' • passato evidenziato in verde';
 return true
}
function setTrackingPausedMessage(){if(currentMode==='major'){$('majorRouteInfo').textContent='⏸ PAUSA • punto memorizzato • gli spostamenti non vengono segnati come passati';setMajorNavHud(true,'⏸','Pausa','','Punto di ripresa memorizzato');majorSpeak('Pausa. Punto di ripresa memorizzato.','paused',true)}}
function initMajorVoiceButton(){const b=$('majorVoiceBtn');if(!b)return;const refresh=()=>{b.textContent=window.MajorOfflineNav?.isVoiceEnabled?.()===false?'🔇':'🔊'};b.onclick=()=>{const next=!(window.MajorOfflineNav?.isVoiceEnabled?.()!==false);window.MajorOfflineNav?.setVoiceEnabled?.(next);refresh();if(next)majorSpeak('Voce attivata','voice-on',true)};refresh()}

const saltEstimate={loadedKg:0,setting:0};const SALT_KG_PER_KM_PER_SETTING=8.8;const mapState={active:false,paused:false,last:null,current:null,pts:[],meters:0,mark:null,watch:null};const line=L.polyline([],{color:'#78909c',weight:4,opacity:.35}).addTo(map),$=id=>document.getElementById(id);function d(a,b){const R=6371000,p=Math.PI/180,x=(b[0]-a[0])*p,y=(b[1]-a[1])*p,q=Math.sin(x/2)**2+Math.cos(a[0]*p)*Math.cos(b[0]*p)*Math.sin(y/2)**2;return 2*R*Math.asin(Math.sqrt(q))}
function currentSaltEstimate(){const loadedKg=Number(saltEstimate.loadedKg)||0,setting=Number(saltEstimate.setting)||0;if(loadedKg<=0||setting<=0)return null;const spreadKg=Math.min(loadedKg,(mapState.meters/1000)*SALT_KG_PER_KM_PER_SETTING*setting);const remainingKg=Math.max(0,loadedKg-spreadKg);return{loadedKg,setting,spreadKg,remainingKg,percent:Math.max(0,Math.min(100,Math.round(remainingKg/loadedKg*100))),rateKgKm:SALT_KG_PER_KM_PER_SETTING*setting}}
function renderSaltEstimate(){const el=$('saltEstimate');if(!el)return;const estimate=currentSaltEstimate();el.hidden=!estimate;if(!estimate)return;el.classList.toggle('lowSalt',estimate.percent<20);$('saltEstimateText').textContent='SALE STIMATO · '+estimate.percent+'% · ~'+estimate.remainingKg.toFixed(0)+' kg rimasti'}
let sessionBackupChain=Promise.resolve(),lastCloudSessionSave=0;
function readMajorWorkSession(){try{const key=snowUserStorageKey(MAJOR_SESSION_KEY);return key?JSON.parse(localStorage.getItem(key)||'null'):null}catch(e){return null}}
async function recoverOfflineSession(){
 try{return await window.VargaSnowOutbox?.loadSession?.()}catch(e){console.warn('Backup offline non disponibile',e);return null}
}
function saveMajorWorkSession(force=false){
 if(!mapState.active||currentMode!=='major')return;
 const now=Date.now();const saveCloud=force||now-lastCloudSessionSave>=10000;lastMajorSessionSave=now;
 const pts=mapState.pts.map(p=>[+p[0].toFixed(6),+p[1].toFixed(6)]);
 const payload={v:2,mode:'major',routeIndex:majorRi,pointIndex:majorPi,active:true,paused:!!mapState.paused,meters:+mapState.meters.toFixed(1),pts,coverage:majorCoverageSegments,tourId:streetWorkTourId,streetWork:streetWorkState,activeStreetIndex,pausePoint:majorNav.pausePoint,navMode:majorNav.mode,resumeMode:majorNav.resumeMode,resumeTarget:majorNav.resumeTarget,current:mapState.current,saltLoadKg:Number(saltEstimate.loadedKg)||0,saltSetting:Number(saltEstimate.setting)||0,updatedAt:now};
 try{localStorage.setItem(snowUserStorageKey(MAJOR_SESSION_KEY),JSON.stringify(payload));sessionBackupChain=sessionBackupChain.catch(()=>{}).then(()=>window.VargaSnowOutbox?.saveSession?.(payload));sessionBackupChain.catch(e=>{$('gps').textContent='⚠ Backup offline non disponibile';console.error(e)});if(saveCloud){lastCloudSessionSave=now;window.VargaSnowCloud?.saveSession(payload).catch(e=>{$('gps').textContent='⚠ Sincronizzazione cloud non riuscita';console.warn('Cloud session save failed',e)})}}catch(e){lastMajorSessionSave=0;$('gps').textContent='⚠ Salvataggio non riuscito: non chiudere l’app';console.warn('Session save failed',e);return false}return true
}
function clearMajorWorkSession(){
 localStorage.removeItem(snowUserStorageKey(MAJOR_SESSION_KEY));streetWorkTourId='';streetWorkState={};activeStreetIndex=null;saltEstimate.loadedKg=0;saltEstimate.setting=0;renderSaltEstimate();majorCoverageSegments=[];majorCoverageLayer.clearLayers();majorTracking.coveredMeters=0;mapState.meters=0;mapState.pts=[];mapState.last=null;line.setLatLngs([])
}
function drawSavedMajorCoverage(){
 majorCoverageLayer.clearLayers();
 for(const s of majorCoverageSegments){if(!Array.isArray(s)||s.length<5||s[0]!==majorRi)continue;L.polyline([[s[1],s[2]],[s[3],s[4]]],{pane:'majorCoveragePane',color:'#20bd62',weight:7,opacity:1,lineCap:'round',interactive:false}).addTo(majorCoverageLayer)}
}
async function restoreMajorWorkSession(){
 let s=readMajorWorkSession();const offline=await recoverOfflineSession();if(offline?.active&&(!s||Number(offline.updatedAt||0)>Number(s.updatedAt||0)))s=offline;try{const cloud=await window.VargaSnowCloud?.loadSession?.();if(cloud?.active&&cloud.mode==='major'&&(!s||Number(cloud.updatedAt?.toMillis?.()||cloud.updatedAt||0)>Number(s.updatedAt||0))){s=cloud;try{localStorage.setItem(snowUserStorageKey(MAJOR_SESSION_KEY),JSON.stringify(Object.assign({},cloud,{updatedAt:Date.now()})))}catch(_){}}}catch(e){console.warn('Cloud session restore failed',e)}if(!s?.active||s.mode!=='major'||s.routeIndex!==majorRi)return false;
 mapState.active=true;mapState.paused=!!s.paused;mapState.meters=Number(s.meters)||0;saltEstimate.loadedKg=Number(s.saltLoadKg)||0;saltEstimate.setting=Number(s.saltSetting)||0;mapState.pts=Array.isArray(s.pts)?s.pts:[];mapState.last=null;mapState.current=Array.isArray(s.current)?s.current:null;line.setLatLngs(mapState.pts);
 majorPi=Number.isInteger(s.pointIndex)&&s.pointIndex>=0&&s.pointIndex<(majorRoutes[majorRi]?.streets?.length||0)?s.pointIndex:0;
 majorCoverageSegments=Array.isArray(s.coverage)?s.coverage:[];streetWorkTourId=s.tourId||newStreetTourId();streetWorkState=s.streetWork&&typeof s.streetWork==='object'?s.streetWork:{};activeStreetIndex=Number.isInteger(s.activeStreetIndex)?s.activeStreetIndex:null;drawSavedMajorCoverage();
 majorNav.pausePoint=Array.isArray(s.pausePoint)?s.pausePoint:null;majorNav.resumeMode=s.resumeMode||null;majorNav.resumeTarget=Array.isArray(s.resumeTarget)?s.resumeTarget:null;majorNav.target=null;majorNav.mode=mapState.paused?'paused':(s.navMode==='to-pause'&&majorNav.pausePoint?'restore-pause':'restore');
 try{await prepareMajorTrackingRoute()}catch(e){$('gps').textContent='Giro recuperato; reticolo da ricaricare'}
 if(majorPauseMarker){routeLayer.removeLayer(majorPauseMarker);majorPauseMarker=null}
 if(mapState.paused&&majorNav.pausePoint)majorPauseMarker=L.circleMarker(majorNav.pausePoint,{radius:11,color:'#fff',weight:4,fillColor:'#ff9800',fillOpacity:1}).addTo(routeLayer).bindPopup('PUNTO DI PAUSA');
 renderMajor();ui();if(mapState.paused)setTrackingPausedMessage();else setMajorNavHud(true,'↻','Sessione ripristinata','','Riprendo dal percorso salvato');
 return true
}
function setMajorMore(open){const menu=$('majorMoreMenu'),btn=$('majorMoreBtn');if(!menu||!btn)return;menu.hidden=!open;btn.setAttribute('aria-expanded',open?'true':'false');btn.classList.toggle('active',open)}
function syncMajorPrimaryAction(){const active=mapState.active,paused=mapState.paused;if($('start')){$('start').hidden=active;$('start').disabled=active}if($('pause')){$('pause').hidden=!active||paused;$('pause').disabled=!active||paused}if($('resume')){$('resume').hidden=!active||!paused;$('resume').disabled=!active||!paused}if($('stop'))$('stop').disabled=!active;if($('return'))$('return').hidden=!(active&&paused&&majorNav.pausePoint)}
function syncScreenAwake(){window.VargaScreenAwake?.setActive(!!(mapState.active||navTarget))}
function ui(){
 syncScreenAwake();if(!mapState.active&&!navTarget){map.setBearing?.(0);navigationHeading.reset()}
 const navMode=currentMode==='major'&&mapState.active;document.body.classList.toggle('major-navigation',navMode);
 $('state').textContent=mapState.paused?'IN PAUSA':mapState.active?'IN LAVORAZIONE':'PRONTO';$('km').textContent=(mapState.meters/1000).toFixed(2)+' km';syncMajorPrimaryAction();renderSaltEstimate();
 const metrics=$('majorNavMetrics');if(metrics){metrics.hidden=!navMode;if(navMode){const r=majorRoutes[majorRi],total=Math.max(1,(r?.totalMeters||1)*(r?.passes||2)),pct=Math.max(0,Math.min(100,Math.round(mapState.meters/total*100))),salt=currentSaltEstimate();$('majorNavKm').textContent=(mapState.meters/1000).toFixed(2)+' km';$('majorNavProgress').textContent=pct+'%';$('majorNavSalt').textContent=salt?salt.percent+'%':'—'}}
 if(mapState.active)saveMajorWorkSession();setTimeout(repairMapSize,0)
}function onPos(p){const c=p.coords,ll=[c.latitude,c.longitude],prev=mapState.current;mapState.current=ll;$('gps').textContent='GPS attivo • '+Math.round(c.accuracy)+' m';$('accuracy').textContent=Math.round(c.accuracy)+' m';if(!mapState.mark)mapState.mark=L.circleMarker(ll,{radius:8,weight:4,color:'#fff',fillColor:'#1677ff',fillOpacity:1}).addTo(map);else mapState.mark.setLatLng(ll);mapState.mark.setStyle(currentMode==='major'&&mapState.active?{opacity:0,fillOpacity:0}:{opacity:1,fillOpacity:1});const moveHeading=navigationHeading.update(c,ll,d,headingDeg)??0;updateMajorVehicle(ll,moveHeading);if(currentMode==='major'&&mapState.active){autoMajorView(ll,Number.isFinite(c.speed)?c.speed:0,majorNav.target?d(ll,majorNav.target):0,mapState.paused||majorNav.mode!=='to-start'&&majorNav.mode!=='to-pause');updateMajorNavigation(ll,c.accuracy,Number.isFinite(c.speed)?c.speed:0);}else if(navTarget){navUpdate(ll);if(navTarget)autoMajorView(ll,Number.isFinite(c.speed)?c.speed:0,d(ll,navTarget),false);else{map.setBearing?.(0);navigationHeading.reset()}}if(mapState.active&&(!Number.isFinite(c.accuracy)||c.accuracy>35)){mapState.last=null;resetTrackingJoin()}if(mapState.active&&!mapState.paused&&c.accuracy<=35){const worked=currentMode==='major'?markMajorPassed(ll,c.accuracy):true;if(worked){if(mapState.last){const x=d(mapState.last,ll);if(x>2&&x<120){mapState.meters+=x;mapState.pts.push(ll);line.setLatLngs(mapState.pts)}}else mapState.pts.push(ll);mapState.last=ll}else mapState.last=null;ui()}}function gps(){if(!navigator.geolocation)return alert('GPS non disponibile');if(mapState.watch!==null)return;mapState.watch=navigator.geolocation.watchPosition(onPos,e=>$('gps').textContent='GPS: '+e.message,{enableHighAccuracy:true,maximumAge:1000,timeout:15000})}
function currentGpsOnce(){return new Promise(resolve=>{if(!navigator.geolocation){resolve(mapState.current?{ll:mapState.current,accuracy:20}:null);return}navigator.geolocation.getCurrentPosition(p=>resolve({ll:[p.coords.latitude,p.coords.longitude],accuracy:p.coords.accuracy}),()=>resolve(mapState.current?{ll:mapState.current,accuracy:20}:null),{enableHighAccuracy:true,maximumAge:1000,timeout:10000})})}
$('majorMoreBtn').onclick=()=>setMajorMore($('majorMoreMenu').hidden);$('majorMoreClose').onclick=()=>setMajorMore(false);if($('appBuild'))$('appBuild').textContent='Build '+APP_BUILD;$('forceUpdateApp').onclick=async()=>{if(!navigator.onLine){alert('Per aggiornare l’app serve una connessione Internet.');return}saveMajorWorkSession(true);setMajorMore(false);$('forceUpdateApp').textContent='AGGIORNAMENTO…';try{const regs=await navigator.serviceWorker?.getRegistrations?.()||[];await Promise.all(regs.map(x=>x.unregister()));if(window.caches){const keys=await caches.keys();await Promise.all(keys.map(k=>caches.delete(k)))}location.replace(location.pathname+'?update='+Date.now())}catch(e){location.reload()}};$('followPosition').onclick=()=>setPositionFollow(true);$('loc').onclick=()=>navigator.geolocation.getCurrentPosition(p=>map.setView([p.coords.latitude,p.coords.longitude],17),()=>alert('Consenti la posizione al browser'),{enableHighAccuracy:true});
$('start').onclick=()=>{setMajorMore(false);const lastLoad=localStorage.getItem('snowSaltLastLoadKg'),lastSetting=localStorage.getItem('snowSaltLastSetting');$('saltLoadKg').value=lastLoad||'';$('saltSetting').value=lastSetting||'5';$('saltSetup').hidden=false;setTimeout(()=>$('saltLoadKg').focus(),50)};$('saltSetupCancel').onclick=()=>{$('saltSetup').hidden=true};$('saltSetupConfirm').onclick=()=>{const loadKg=Number($('saltLoadKg').value),setting=Number($('saltSetting').value);if(!Number.isFinite(loadKg)||loadKg<=0){$('saltLoadKg').reportValidity();return}if(!Number.isInteger(setting)||setting<1||setting>10)return;saltEstimate.loadedKg=loadKg;saltEstimate.setting=setting;localStorage.setItem('snowSaltLastLoadKg',String(loadKg));localStorage.setItem('snowSaltLastSetting',String(setting));$('saltSetup').hidden=true;startMajorWork()};const openSaltAddDialog=()=>{if(!mapState.active||!currentSaltEstimate())return;$('saltAddKg').value='';$('saltAddDialog').hidden=false;setTimeout(()=>$('saltAddKg').focus(),50)};$('saltEstimate').onclick=openSaltAddDialog;$('saltEstimate').onkeydown=e=>{if(e.key==='Enter'||e.key===' '){e.preventDefault();openSaltAddDialog()}};$('saltAddCancel').onclick=()=>{$('saltAddDialog').hidden=true};$('saltAddConfirm').onclick=()=>{const addedKg=Number($('saltAddKg').value);if(!Number.isFinite(addedKg)||addedKg<=0){$('saltAddKg').reportValidity();return}saltEstimate.loadedKg=(Number(saltEstimate.loadedKg)||0)+addedKg;$('saltAddDialog').hidden=true;ui();saveMajorWorkSession(true)};async function startMajorWork(){setPositionFollow(true);gps();beginStreetWorkTour();mapState.active=true;mapState.paused=false;mapState.last=null;resetTrackingJoin();majorNav.pausePoint=null;majorNav.target=null;majorNav.resumeMode=null;majorNav.resumeTarget=null;majorNav.mode='loading';$('start').disabled=true;$('pause').disabled=false;$('stop').disabled=false;ui();if(currentMode==='major'){saveMajorWorkSession(true);try{const ready=await prepareMajorTrackingRoute();if(!ready){$('majorRouteInfo').textContent='Percorso non disponibile per la navigazione automatica';return}await showMajorRoute();const pos=await currentGpsOnce();if(!mapState.active||mapState.paused)return;if(pos){mapState.current=pos.ll;await beginMajorNavigation(pos.ll,pos.accuracy);saveMajorWorkSession(true)}else{majorNav.mode='restore';$('majorRouteInfo').textContent='In attesa del GPS per trovare il punto più vicino'}}catch(e){$('majorRouteInfo').textContent='Percorso non caricato: navigazione automatica non disponibile'}}};
$('pause').onclick=()=>{setMajorMore(false);mapState.paused=true;mapState.last=null;majorNav.pausePoint=mapState.current?[...mapState.current]:(majorTracking.lastSnap?[...majorTracking.lastSnap]:null);majorNav.resumeMode=majorNav.mode;majorNav.resumeTarget=majorNav.target?[...majorNav.target]:null;majorNav.mode='paused';majorNav.target=null;resetTrackingJoin();clearMajorNavLine();clearMajorGuide();if(majorPauseMarker){routeLayer.removeLayer(majorPauseMarker);majorPauseMarker=null}if(majorNav.pausePoint)majorPauseMarker=L.circleMarker(majorNav.pausePoint,{radius:11,color:'#fff',weight:4,fillColor:'#ff9800',fillOpacity:1}).addTo(routeLayer).bindPopup('PUNTO DI PAUSA');$('pause').hidden=true;$('resume').hidden=false;$('return').hidden=false;setTrackingPausedMessage();pauseStreetWork();ui();saveMajorWorkSession(true)};
async function resumeMajorWork(){
 setMajorMore(false);majorNav.mode='loading';mapState.paused=false;resumeStreetWork();mapState.last=null;resetTrackingJoin();setPositionFollow(true);
 $('resume').hidden=true;$('pause').hidden=false;$('return').hidden=true;ui();
 if(currentMode==='major'){
  try{
   const ready=await prepareMajorTrackingRoute();if(!ready)throw new Error('Reticolo non disponibile');
   const pos=await currentGpsOnce();if(!mapState.active||mapState.paused)return;
   if(!pos)throw new Error('GPS non disponibile');
   mapState.current=pos.ll;majorNav.pausePoint=null;majorNav.resumeMode=null;majorNav.resumeTarget=null;
   if(majorPauseMarker){routeLayer.removeLayer(majorPauseMarker);majorPauseMarker=null}
   // Start a new work segment here; paused travel is never joined to the green trace.
   await beginMajorNavigation(pos.ll,pos.accuracy);
  }catch(e){mapState.paused=true;majorNav.mode='paused';pauseStreetWork();ui();$('majorRouteInfo').textContent='Ripresa non disponibile: controlla il GPS e riprova'}
  saveMajorWorkSession(true)
 }
}
$('resume').onclick=resumeMajorWork;
$('return').onclick=async()=>{if(currentMode==='major'&&majorNav.pausePoint){const pos=await currentGpsOnce();if(pos)await navigateMajorTo(pos.ll,majorNav.pausePoint,'to-pause','RITORNO AL PUNTO DI PAUSA');return}if(mapState.pts.length){const p=mapState.pts.at(-1);window.open('https://www.google.com/maps/dir/?api=1&destination='+p[0]+','+p[1],'_blank')}};
async function stopMajorWork(){if(!mapState.active)return;const stopButton=$('stop');stopButton.disabled=true;closeStreetWork();const synced=await flushStreetWorkQueue(true);if(!synced){alert('I dati del giro non sono ancora sincronizzati. Il lavoro rimane attivo e salvato sul dispositivo. Riconnettiti e riprova.');stopButton.disabled=false;return}try{await sessionBackupChain;await window.VargaSnowCloud?.clearSession?.();await window.VargaSnowOutbox?.clearSession?.()}catch(e){alert('Impossibile confermare la chiusura sul server. Il giro resta disponibile.');stopButton.disabled=false;return}setMajorMore(false);mapState.active=false;mapState.paused=false;mapState.last=null;majorNav.mode='idle';majorNav.target=null;majorNav.pausePoint=null;majorNav.resumeMode=null;majorNav.resumeTarget=null;resetTrackingJoin();clearMajorNavLine();clearMajorGuide();window.MajorOfflineNav?.resetSpeech?.();setMajorNavHud(false);if(majorVehicleMarker){map.removeLayer(majorVehicleMarker);majorVehicleMarker=null}if(majorPauseMarker){routeLayer.removeLayer(majorPauseMarker);majorPauseMarker=null}clearMajorWorkSession();majorDone.clear();localStorage.removeItem('majorDone');$('start').disabled=false;$('pause').disabled=true;$('pause').hidden=false;$('resume').hidden=true;$('return').hidden=true;$('stop').disabled=true;ui()} $('stop').onclick=stopMajorWork;if($('adminArchiveTour'))$('adminArchiveTour').onclick=async()=>{
 if(window.VargaSnowCloud?.role!=='admin')return;
 if(!mapState.active)return alert('Nessun giro attivo da archiviare.');
 if(!confirm('Archiviare e terminare questo giro? Il giro resterà nello storico amministratore.'))return;
 const b=$('adminArchiveTour');b.disabled=true;b.textContent='ARCHIVIAZIONE…';
 try{
  saveMajorWorkSession(true);
  if(!await flushStreetWorkQueue(true))throw new Error('Registro strade non sincronizzato: archiviazione rinviata');
  const route=majorRoutes[majorRi]||{};
  await window.VargaSnowCloud.archiveTour({
   v:1,type:'piano-neve',tourId:streetWorkTourId,mode:'major',routeIndex:majorRi,routeCode:route.code||'',routeName:route.name||('Percorso '+(majorRi+1)),
   pointIndex:majorPi,meters:+mapState.meters.toFixed(1),coverage:majorCoverageSegments,track:mapState.pts,
   completedStreets:[...majorDone],lastPosition:mapState.current||null,saltEstimate:currentSaltEstimate(),endedAtClient:new Date().toISOString()
  });
  await stopMajorWork();if(mapState.active)throw new Error('Giro archiviato, ma la chiusura non è ancora confermata');alert('Giro archiviato e terminato.');
 }catch(e){console.error(e);alert('Impossibile archiviare il giro. Il lavoro attivo non è stato cancellato.');b.disabled=false;b.textContent='📦 ARCHIVIA E TERMINA GIRO'}
};
document.querySelectorAll('.tab').forEach(b=>b.onclick=async()=>{document.querySelectorAll('.tab').forEach(x=>x.classList.remove('active'));b.classList.add('active');const minor=b.dataset.mode==='minor';currentMode=minor?'minor':'major';drawAllRoutes();routeLayer.clearLayers();$('major').hidden=minor;$('minor').hidden=!minor;if(minor){setMajorNavHud(false);if(majorBoundaryLayer&&map.hasLayer(majorBoundaryLayer))map.removeLayer(majorBoundaryLayer);if(!minorMapOpened&&map.hasLayer(streetMap)){map.removeLayer(streetMap);regionalOrtho.addTo(map)}if(!map.hasLayer(officialFootways))officialFootways.addTo(map);if(!map.hasLayer(officialCycleways))officialCycleways.addTo(map);minorMapOpened=true;drawRoute()}else{if(map.hasLayer(officialFootways))map.removeLayer(officialFootways);if(map.hasLayer(officialCycleways))map.removeLayer(officialCycleways);if(!map.hasLayer(streetMap)){map.removeLayer(regionalOrtho);streetMap.addTo(map)}if(majorBoundaryLayer&&!map.hasLayer(majorBoundaryLayer))majorBoundaryLayer.addTo(map);renderMajor();await showMajorRoute();restoreMajorNavVisuals()}});$('majorDone').onclick=()=>{setMajorMore(false);const r=majorRoutes[majorRi];if(!r)return;majorDone.add(majorKey());localStorage.setItem('majorDone',JSON.stringify([...majorDone]));if(majorPi<r.streets.length-1)majorPi++;renderMajor()};$('minorStart').onclick=()=>{gps();$('navBanner').textContent='🟠 INTERVENTO IN CORSO';$('minorStart').textContent='● INTERVENTO ATTIVO';renderRoute()};$('minorDone').onclick=()=>{done.add(ri+'-'+pi);skipped.delete(ri+'-'+pi);localStorage.setItem('snowDone',JSON.stringify([...done]));localStorage.setItem('snowSkipped',JSON.stringify([...skipped]));if(pi<routes[ri].points.length-1)pi++;drawRoute();drawAllRoutes()};$('next').onclick=()=>{pi++;if(pi>=routes[ri].points.length){pi=0;ri=(ri+1)%routes.length}drawRoute(false);focusStop()};$('skip').onclick=()=>{const why=prompt('Motivo: strada chiusa, ostacolo, già pulita o altro?');if(!why)return;skipped.add(ri+'-'+pi);localStorage.setItem('snowSkipped',JSON.stringify([...skipped]));if(pi<routes[ri].points.length-1)pi++;drawRoute()};$('problem').onclick=()=>{const note=prompt('Descrivi il problema operativo:');if(note)alert('Segnalazione salvata sul dispositivo: '+note)};$('routePrev').onclick=()=>{ri=(ri+routes.length-1)%routes.length;pi=0;drawRoute()};$('routeNext').onclick=()=>{ri=(ri+1)%routes.length;pi=0;drawRoute()};$('navigate').onclick=async()=>{const s=routes[ri].segments[pi],p=surfaceStart(ri,pi);if(!p)return alert('Geometria nuova ancora in ricostruzione per questa tappa');gps();navigator.geolocation.getCurrentPosition(async pos=>{const here=[pos.coords.latitude,pos.coords.longitude];navTarget=p;setPositionFollow(true);syncScreenAwake();const rr=await roadRoute(here,p);if(navLine)routeLayer.removeLayer(navLine);if(rr){const pts=rr.geometry.coordinates.map(x=>[x[1],x[0]]);navLine=L.polyline(pts,{color:'#ff9800',weight:8,opacity:.95}).addTo(routeLayer);navRoute=rr;map.fitBounds(navLine.getBounds(),{padding:[35,35]});$('routeInfo').textContent='🧭 NAVIGAZIONE • '+fmtDist(rr.distance)+' • circa '+Math.max(1,Math.round(rr.duration/60))+' min'}else{$('routeInfo').textContent='Navigazione stradale non disponibile';map.setView(p,16)}L.circleMarker(p,{radius:15,color:'#fff',weight:4,fillColor:'#ff9800',fillOpacity:1}).addTo(routeLayer).bindPopup('DESTINAZIONE • '+s.name)},()=>alert('Attiva il GPS per iniziare la navigazione'),{enableHighAccuracy:true})};$('streetView').onclick=()=>{const p=surfaceStart(ri,pi);if(p)window.open('https://www.google.com/maps/@?api=1&map_action=pano&viewpoint='+p[0]+','+p[1],'_blank','noopener')};async function registerAppWorker(){if(window.__VARGA_CANTIERI_INTEGRATED__)return;if(!('serviceWorker'in navigator))return;let reloading=false;navigator.serviceWorker.addEventListener('controllerchange',()=>{if(reloading)return;saveMajorWorkSession(true);reloading=true;location.reload()});try{const reg=await navigator.serviceWorker.register('./sw.js?v='+SW_BUILD,{scope:'./',updateViaCache:'none'});await reg.update();if(reg.waiting)reg.waiting.postMessage({type:'SKIP_WAITING'})}catch(e){console.warn('Service worker update failed',e)}}
window.addEventListener('pagehide',()=>saveMajorWorkSession(true));document.addEventListener('visibilitychange',()=>{if(document.hidden)saveMajorWorkSession(true)});initMajorRouteSelection();initMajorVoiceButton();bootWorkSession();
async function bootWorkSession(){
 const saved=readMajorWorkSession();
 if(saved?.active&&saved.mode==='major'&&Number.isInteger(saved.routeIndex)&&majorRoutes[saved.routeIndex]){
  majorRi=saved.routeIndex;currentMode='major';$('startupChooser').hidden=true;
  await showMajorRoute();await restoreMajorWorkSession();restoreMajorNavVisuals();
 }else{
  initStartupSelection();
  // Collegamento diretto dalle commesse Salatura Strada 1/2/3 di Varga Cantieri.
  // NON inizia il lavoro e NON interrompe una sessione già attiva:
  // seleziona soltanto il percorso assegnato.
  const requested=new URLSearchParams(window.location.search).get('vargaMajorRoute');
  if(/^[012]$/.test(String(requested||''))){
    showStartupRouteStep('major');
    await finishStartupSelection('major',Number(requested));
  }
 }
 renderMajor();drawAllRoutes();ui();gps();registerAppWorker();
}



let cloudRouteEdits=[];
async function loadCloudRouteEdits(){try{cloudRouteEdits=await window.VargaSnowCloud?.loadRouteEdits?.()||[];applyCloudRouteEdits()}catch(e){console.warn('Route edits',e)}}
function applyCloudRouteEdits(){for(const e of cloudRouteEdits){if(e.mode==='minor'&&routes[e.routeIndex]){const r=routes[e.routeIndex];if(e.action==='delete'&&r.segments[e.segmentIndex]){r.segments[e.segmentIndex].coords=[];if(window.REBUILT_MINOR_ROUTES?.[e.routeIndex]?.[e.segmentIndex])window.REBUILT_MINOR_ROUTES[e.routeIndex][e.segmentIndex].excluded=true}else if(e.action==='modify'&&r.segments[e.segmentIndex]&&Array.isArray(e.coords)){r.segments[e.segmentIndex].coords=e.coords;if(window.REBUILT_MINOR_ROUTES?.[e.routeIndex]?.[e.segmentIndex])window.REBUILT_MINOR_ROUTES[e.routeIndex][e.segmentIndex].coords=e.coords}else if(e.action==='add'&&Array.isArray(e.coords)){r.segments.push({name:e.name||'Nuova strada',coords:e.coords});r.points.push(e.name||'Nuova strada')}}else if(e.mode==='major'&&majorRoutes[e.routeIndex]){const r=majorRoutes[e.routeIndex];if(e.action==='delete'&&r.streets[e.segmentIndex])r.streets[e.segmentIndex]._adminDeleted=true;else if(e.action==='modify'&&r.streets[e.segmentIndex])r.streets[e.segmentIndex]._adminCoords=e.coords;else if(e.action==='add')r.streets.push({name:e.name||'Nuova strada',searchName:e.name||'',meters:0,_adminCoords:e.coords})}}}

/* Admin route manager */
let routeEditMode=false,routeEditPoints=[],routeEditLine=null,routeEditSelected=null;
function openRouteManager(){if(window.VargaSnowCloud?.role!=='admin')return;routeEditMode=false;routeEditPoints=[];$('routeEditor').hidden=false;setMajorMore(false)}
function closeRouteManager(){routeEditMode=false;routeEditPoints=[];routeEditSelected=null;if(routeEditLine){map.removeLayer(routeEditLine);routeEditLine=null}$('routeEditor').hidden=true}
$('adminManageRoutes').onclick=openRouteManager;
$('routeEditorClose').onclick=closeRouteManager;
$('routeAddStreet').onclick=()=>{routeEditMode=true;routeEditPoints=[];if(routeEditLine){map.removeLayer(routeEditLine);routeEditLine=null}$('routeEditor').hidden=true;alert('Tocca sulla mappa il punto iniziale e poi il punto finale della strada.')};
map.on('click',e=>{if(!routeEditMode||window.VargaSnowCloud?.role!=='admin')return;if(routeEditMode==='select-edit'||routeEditMode==='select-delete')return;if(routeEditMode==='extend'){routeEditPoints.push([e.latlng.lat,e.latlng.lng]);if(routeEditLine)map.removeLayer(routeEditLine);routeEditLine=L.polyline(routeEditPoints,{weight:9}).addTo(map);routeEditMode=false;$('routeEditor').hidden=false;$('routeEditorStatus').textContent='Tratto allungato. Premi Salva modifica.';$('routeEditorSave').disabled=false;return}routeEditPoints.push([e.latlng.lat,e.latlng.lng]);L.circleMarker(e.latlng,{radius:7,weight:3}).addTo(map);if(routeEditPoints.length===2){routeEditLine=L.polyline(routeEditPoints,{weight:9}).addTo(map);routeEditMode=false;$('routeEditor').hidden=false;$('routeEditorStatus').textContent='Nuovo tratto selezionato. Premi Salva modifica.';$('routeEditorSave').disabled=false}});
$('routeEditorSave').onclick=async()=>{if(routeEditPoints.length<2)return;const name=prompt('Nome della strada (facoltativo):','')||'Nuova strada';const municipality=prompt('Comune (facoltativo):','Castel Maggiore')||'';try{await window.VargaSnowCloud.saveRouteEdit({action:routeEditSelected?'modify':'add',mode:currentMode,routeIndex:currentMode==='major'?majorRi:ri,name,municipality,segmentIndex:routeEditSelected?.index??null,coords:routeEditPoints,createdAtClient:new Date().toISOString()});alert('Strada salvata.');closeRouteManager();location.reload()}catch(e){console.error(e);alert('Salvataggio non riuscito.')}};
function selectRouteForAdmin(kind,index,name,coords){if(window.VargaSnowCloud?.role!=='admin')return;routeEditSelected={kind,index,name,coords:(coords||[]).map(p=>[p[0],p[1]])};$('routeEditor').hidden=false;$('routeEditorStatus').textContent='Selezionato: '+name;$('routeEditorSave').disabled=true}
function adminRoutePopup(name,kind,index,coords){if(window.VargaSnowCloud?.role!=='admin')return '';const id='arp_'+Date.now()+'_'+Math.floor(Math.random()*9999);setTimeout(()=>{const e=document.getElementById(id);if(!e)return;e.querySelector('[data-edit]').onclick=()=>{map.closePopup();routeEditMode='select-edit';selectRouteForAdmin(kind,index,name,coords)};e.querySelector('[data-del]').onclick=()=>{map.closePopup();routeEditMode='select-delete';selectRouteForAdmin(kind,index,name,coords)}},0);return '<div id="'+id+'" class="adminRoutePopup"><b>Amministratore</b><div><button data-edit>✏️ MODIFICA / ALLUNGA</button><button data-del class="danger">🗑 ELIMINA</button></div></div>'}
function adminSelectableLine(coords,kind,index,name){if(window.VargaSnowCloud?.role!=='admin'||!coords?.length)return;const hit=L.polyline(coords,{weight:24,opacity:0,interactive:true}).addTo(routeLayer);hit.on('click',e=>{if(!routeEditMode)return;L.DomEvent.stopPropagation(e);selectRouteForAdmin(kind,index,name,coords)})}
$('routeEditSelected').onclick=()=>{routeEditMode='select-edit';$('routeEditor').hidden=true;alert('Tocca la strada del percorso che vuoi modificare o allungare.')};
$('routeDeleteSelected').onclick=()=>{routeEditMode='select-delete';$('routeEditor').hidden=true;alert('Tocca la strada del percorso che vuoi eliminare.')};
$('routeEditorStatus').onclick=async()=>{if(!routeEditSelected)return;if(routeEditMode==='select-delete'){if(!confirm('Eliminare '+routeEditSelected.name+' dal percorso?'))return;try{await window.VargaSnowCloud.saveRouteEdit({action:'delete',mode:currentMode,routeIndex:currentMode==='major'?majorRi:ri,segmentIndex:routeEditSelected.index,name:routeEditSelected.name,coords:routeEditSelected.coords,createdAtClient:new Date().toISOString()});alert('Eliminazione salvata.');closeRouteManager();location.reload()}catch(e){alert('Eliminazione non riuscita.')}return}if(routeEditMode==='select-edit'){routeEditPoints=routeEditSelected.coords.slice();routeEditLine=L.polyline(routeEditPoints,{weight:10}).addTo(map);routeEditMode='extend';$('routeEditor').hidden=true;alert('Tocca sulla mappa il nuovo punto finale per allungare il tratto.')}};

window.addEventListener('varga-auth-ready',async e=>{const admin=e.detail?.role==='admin';const b=$('adminArchiveTour');if(b)b.hidden=!admin;const rm=$('adminManageRoutes');if(rm)rm.hidden=!admin;const ex=$('adminExportStreetWork');if(ex){ex.hidden=!admin;ex.onclick=exportStreetWorkExcel;}repairMapSize();setTimeout(repairMapSize,500)});

