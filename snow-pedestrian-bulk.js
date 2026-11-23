/* Varga Cantieri - allineamento automatico conservativo marciapiedi.
 * Esegue solo per amministratori autenticati, sulle 3 commesse normali.
 * Si basa esclusivamente sui percorsi OSM separati e marcati
 * highway=footway + footway=sidewalk; NON usa la mezzeria delle strade.
 * Mantiene tutti i 45 interventi e salva un backup per ogni correzione.
 */
(function (global) {
  "use strict";
  const VERSION = "20261123-pedestrian-safe1";
  const MAX_METERS = 2.5;
  const ENDPOINT_MAX = 55;
  const MAX_WAYS = 1800;
  const OVERPASS = "https://overpass.private.coffee/api/interpreter";
  const DIST_EARTH = 6371000;
  let running = false;
  const toRadians = a => a * Math.PI / 180;
  function metres(a,b) {
    const dLat=toRadians(b[0]-a[0]),dLon=toRadians(b[1]-a[1]);
    const h=Math.sin(dLat/2)**2+Math.cos(toRadians(a[0]))*Math.cos(toRadians(b[0]))*Math.sin(dLon/2)**2;
    return 2*DIST_EARTH*Math.atan2(Math.sqrt(h),Math.sqrt(Math.max(0,1-h)));
  }
  function xy(point,refLat) {
    return [toRadians(point[1])*DIST_EARTH*Math.cos(toRadians(refLat)),toRadians(point[0])*DIST_EARTH];
  }
  function project(point,a,b) {
    const lat = point[0], p = xy(point,lat), aa=xy(a,lat), bb=xy(b,lat);
    const vx=bb[0]-aa[0],vy=bb[1]-aa[1],length=vx*vx+vy*vy;
    const t=length?Math.min(1,Math.max(0,((p[0]-aa[0])*vx+(p[1]-aa[1])*vy)/length)):0;
    return [a[0]+t*(b[0]-a[0]),a[1]+t*(b[1]-a[1])];
  }
  function normalizePart(part) {
    return Array.isArray(part)
      ? part.map(p=>[Number(p?.[0]),Number(p?.[1])]).filter(p=>p.every(Number.isFinite))
      : [];
  }
  function sections() {
    return global.VARGA_SNOW_CM_CATALOG?.sections || [];
  }
  function getBox(section) {
    const pts = section.stops.flatMap(s=>(s.parts||[]).flatMap(p=>p.coords||[]))
      .filter(p=>Array.isArray(p)&&p.length>=2);
    if (!pts.length) return null;
    const lats=pts.map(p=>p[0]),lons=pts.map(p=>p[1]);
    const south=Math.min(...lats)-.001,north=Math.max(...lats)+.001;
    const west=Math.min(...lons)-.0015,east=Math.max(...lons)+.0015;
    // Non eseguire ricerche Overpass regionali per geometrie corrotte.
    if (north-south>.065 || east-west>.085) return null;
    return [south,west,north,east].map(n=>n.toFixed(6)).join(",");
  }
  async function requestSidewalks(bbox) {
    const query = '[out:json][timeout:22];way["highway"="footway"]["footway"="sidewalk"]('+bbox+');out geom;';
    const controller=new AbortController();
    const timer=setTimeout(()=>controller.abort(),25000);
    try {
      const response=await fetch(OVERPASS,{method:"POST",
        headers:{"Content-Type":"application/x-www-form-urlencoded;charset=UTF-8"},
        body:new URLSearchParams({data:query}),signal:controller.signal});
      if (!response.ok) throw new Error("Overpass HTTP "+response.status);
      const data=await response.json();
      return (data.elements||[]).slice(0,MAX_WAYS).filter(e=>e.type==="way" && Array.isArray(e.geometry))
        .map(e=>({id:e.id,points:e.geometry.map(p=>[+p.lat,+p.lon])
          .filter(p=>p.every(Number.isFinite))})).filter(e=>e.points.length>=2);
    } finally {clearTimeout(timer);}
  }
  function nearestSidewalk(point, ways) {
    let best=null,bestDistance=MAX_METERS;
    for (const way of ways) {
      for (let i=1;i<way.points.length;i++) {
        const p=project(point,way.points[i-1],way.points[i]);
        const d=metres(point,p);
        if (d<bestDistance) {
          best={point:p,wayId:way.id,segment:i-1,dist:d,
            bearing:[way.points[i][0]-way.points[i-1][0],way.points[i][1]-way.points[i-1][1]]};
          bestDistance=d;
        }
      }
    }
    return best;
  }
  function automaticPart(part, ways) {
    const original=normalizePart(part);
    if (original.length<2 || original.length>700) return null;
    // Evita tracciati con vertici a centinaia di metri fra loro: uno snap
    // dei soli capi non descriverebbe il marciapiede intermedio.
    if (original.some((p,i)=>i&&metres(p,original[i-1])>ENDPOINT_MAX)) return null;
    const matching=original.map(p=>nearestSidewalk(p,ways));
    // Richiede un percorso separato OSM CONTINUO e nessun vertice ambiguo.
    if (matching.some(p=>!p)) return null;
    const sameWay=matching.every(m=>m.wayId===matching[0].wayId);
    if (!sameWay) return null;
    // Protegge l'orientamento: i vertici devono seguire lo stesso
    // marciapiede senza attraversamenti o inversioni improvvise.
    const segments=matching.map(m=>m.segment);
    const isNonDecreasing=segments.every((x,i)=>!i||x>=segments[i-1]);
    const isNonIncreasing=segments.every((x,i)=>!i||x<=segments[i-1]);
    if (!isNonDecreasing&&!isNonIncreasing) return null;
    const candidate=matching.map(m=>m.point);
    let before=0,after=0,moved=0,distSum=0;
    for (let i=0;i<candidate.length;i++) {
      const shift=metres(original[i],candidate[i]);
      if (shift>MAX_METERS+.01) return null;
      if (shift>.25) {moved++;distSum+=shift;}
      if (i) {before+=metres(original[i-1],original[i]);after+=metres(candidate[i-1],candidate[i]);}
    }
    if (moved<2 || before<2 || after<2 || Math.abs(before-after)>Math.max(2,before*.08)) return null;
    // Non salvare correzioni che allontanano il percorso originale.
    if (distSum/moved>2.25) return null;
    return {original,candidate,moved};
  }
  async function run({db,admin,uid,onProgress,existing={}}) {
    if (!admin || !db || running || !global.VARGA_SNOW_CM_CATALOG) return null;
    const key="varga-cm-snow-auto-"+VERSION;
    try {
      const last=Number(global.localStorage.getItem(key)||0);
      if (Date.now()-last<24*3600*1000) return null;
    } catch (_) {}
    running=true;
    let corrected=0, unchanged=0, unavailable=0, saved=0;
    const progress=msg=>{if(typeof onProgress==="function")onProgress(msg);};
    try {
      for (const section of sections()) {
        const bbox=getBox(section);
        if (!bbox) {unavailable+=section.stops.length;continue;}
        let ways;
        try {ways=await requestSidewalks(bbox);}
        catch (error) {
          unavailable+=section.stops.length;
          progress("OSM non disponibile per "+section.name+": mantengo tutte le tracce originali.");
          continue;
        }
        for (const stop of section.stops) {
          // Non sovrascrivere disegni che un amministratore ha già salvato.
          if (existing[stop.key]?.parts?.length) {unchanged++;continue;}
          const proposed=(stop.parts||[]).map(p=>automaticPart(p.coords,ways));
          let moved=0;
          const newParts=(stop.parts||[]).map((p,i)=>{
            if (proposed[i]) {moved+=proposed[i].moved;return proposed[i].candidate;}
            return normalizePart(p.coords);
          });
          if (!moved || newParts.some(p=>p.length<2)) {unchanged++;continue;}
          const doc=db.collection("snowRouteGeometry").doc(stop.key);
          try {
            // Transazione anti-concorrenza: un edit manuale prevale sempre.
            const changed=await db.runTransaction(async tx=>{
              const snap=await tx.get(doc);
              if (snap.exists && Array.isArray(snap.data()?.parts) && snap.data().parts.length) return false;
              tx.set(doc,{
                stopKey:stop.key,parts:newParts,
                originalParts:stop.parts.map(p=>normalizePart(p.coords)),
                updatedAt:global.firebase.firestore.FieldValue.serverTimestamp(),
                updatedBy:"snow-pedestrian-auto/"+String(uid||"admin"),
                autoPrecision:{version:VERSION,source:"OpenStreetMap footway=sidewalk",movedVertices:moved,
                  maxMoveMeters:MAX_METERS,confidence:"conservative",requiresFieldVerification:true}
              },{merge:true});
              return true;
            });
            if (changed) {corrected++;saved++;} else unchanged++;
          } catch(error) {
            unavailable++;
            progress("Salvataggio automatico non riuscito per "+stop.key+": tracciato originale conservato.");
          }
        }
        progress("Allineamento marciapiedi: "+corrected+" interventi aggiornati in modo conservativo; "+unchanged+" invariati.");
      }
      // Il fallimento di Overpass non blocca nuovi tentativi; preserva originali.
      if (unavailable===0) try {global.localStorage.setItem(key,String(Date.now()));} catch(_) {}
      progress("Controllo automatico terminato: "+saved+" interventi aggiornati, "+unchanged+" invariati, "+unavailable+" non verificabili. Le parti non riconosciute restano originali; controllo sul posto necessario.");
      return {saved,unchanged,unavailable};
    } finally {running=false;}
  }
  global.VargaSnowAutomaticSidewalks={run,automaticPart,metres};
})(window);
