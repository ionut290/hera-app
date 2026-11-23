/* Varga Cantieri — allineamento automatico ai percorsi pedonali OSM.
 * Importante: mantieni KML originale nel catalogo e qualsiasi modifica
 * manuale Firebase. Scrivi SOLO i 4 campi permessi dalle regole Firestore.
 */
(function installSnowSidewalkAlignment(global) {
  "use strict";
  const VERSION = "20261123-nearest-footway2";
  const MAX_SNAP_METERS = 12;
  const SAMPLE_METERS = 6;
  const MAX_PART_SAMPLES = 1500;
  const MAX_STOP_POINTS = 5000;
  const MAX_WAYS = 3500;
  const EARTH = 6371000;
  const overpassURLs = [
    "https://overpass.private.coffee/api/interpreter",
    "https://overpass-api.de/api/interpreter"
  ];
  const rad = value => value * Math.PI / 180;
  const metres = (a,b) => {
    const deltaLat=rad(b[0]-a[0]), deltaLng=rad(b[1]-a[1]);
    const x=Math.sin(deltaLat/2)**2 + Math.cos(rad(a[0]))*Math.cos(rad(b[0]))*Math.sin(deltaLng/2)**2;
    return EARTH*2*Math.atan2(Math.sqrt(x),Math.sqrt(Math.max(0,1-x)));
  };
  const xy=(p,lat)=>[rad(p[1])*EARTH*Math.cos(rad(lat)), rad(p[0])*EARTH];
  const cleanPart=part=>Array.isArray(part)
    ?part.map(p=>[Number(p?.[0]),Number(p?.[1])]).filter(p=>p.every(Number.isFinite)):[];
  const cloneParts=stop=>(stop.parts||[]).map(p=>cleanPart(p.coords));
  function project(point,a,b) {
    const ref=point[0],p=xy(point,ref),a0=xy(a,ref),b0=xy(b,ref);
    const dx=b0[0]-a0[0],dy=b0[1]-a0[1],len=dx*dx+dy*dy;
    const t=len?Math.min(1,Math.max(0,((p[0]-a0[0])*dx+(p[1]-a0[1])*dy)/len)):0;
    const target=[a[0]+t*(b[0]-a[0]),a[1]+t*(b[1]-a[1])];
    return {point:target,distance:metres(point,target)};
  }
  function samplePart(source) {
    const points=cleanPart(source); if(points.length<2) return [];
    const samples=[points[0]];
    for(let i=1;i<points.length;i++) {
      const before=points[i-1],after=points[i],distance=metres(before,after);
      const steps=Math.max(1,Math.ceil(distance/SAMPLE_METERS));
      if(samples.length+steps>MAX_PART_SAMPLES) return [];
      for(let k=1;k<=steps;k++) {
        const ratio=k/steps;
        samples.push([before[0]+ratio*(after[0]-before[0]),before[1]+ratio*(after[1]-before[1])]);
      }
    }
    return samples;
  }
  function kindOf(way) {
    const tags=way.tags||{};
    if(tags.footway==="sidewalk") return "sidewalk";
    if(tags.footway==="crossing") return "crossing";
    if(tags.highway==="pedestrian") return "pedestrian";
    if(tags.highway==="footway" || !tags.highway) return "footway";
    if(tags.highway==="path" && tags.foot!=="no") return "path";
    return "excluded";
  }
  // Griglia geografica da circa 65 m: non confrontare ogni vertice con
  // migliaia di segmenti, soprattutto sui telefoni.
  const cellKey=(lat,lng)=>Math.floor(lat/.00058)+":"+Math.floor(lng/.0008);
  function indexWays(ways) {
    const cells=new Map();
    for(const way of ways) {
      const kind=kindOf(way);
      if(kind==="excluded")continue;
      const pts=Array.isArray(way.points)?way.points:[];
      for(let i=1;i<pts.length;i++) {
        const a=pts[i-1],b=pts[i];
        if(!Array.isArray(a)||!Array.isArray(b)||metres(a,b)>160)continue;
        const seg={a,b,wayId:way.id,kind};
        const lat0=Math.floor((Math.min(a[0],b[0])-.00018)/.00058);
        const lat1=Math.floor((Math.max(a[0],b[0])+.00018)/.00058);
        const lon0=Math.floor((Math.min(a[1],b[1])-.00023)/.0008);
        const lon1=Math.floor((Math.max(a[1],b[1])+.00023)/.0008);
        if((lat1-lat0+1)*(lon1-lon0+1)>50)continue;
        for(let lat=lat0;lat<=lat1;lat++)for(let lon=lon0;lon<=lon1;lon++){
          const key=lat+":"+lon;
          if(!cells.has(key))cells.set(key,[]);
          cells.get(key).push(seg);
        }
      }
    }
    return cells;
  }
  const penalty={sidewalk:0,footway:1,pedestrian:1.5,crossing:2,path:3};
  function closest(point,tangent,index,previous) {
    const key=cellKey(point[0],point[1]);
    // Inline cell lookup for speed, including neighboring cells: 12m margin.
    const row=Math.floor(point[0]/.00058),col=Math.floor(point[1]/.0008);
    const seen=new Set(), options=[];
    const tangentNorm=Math.hypot(tangent[0],tangent[1]);
    for(let dr=-1;dr<=1;dr++)for(let dc=-1;dc<=1;dc++){
      for(const seg of (index.get((row+dr)+":"+(col+dc))||[])) {
        if(seen.has(seg))continue;
        seen.add(seg);
        const candidate=project(point,seg.a,seg.b);
        if(candidate.distance>MAX_SNAP_METERS)continue;
        const x=xy(seg.a,point[0]),y=xy(seg.b,point[0]);
        const sx=y[0]-x[0],sy=y[1]-x[1],sn=Math.hypot(sx,sy);
        if(sn<.5)continue;
        const parallel=tangentNorm?
          Math.abs((sx*tangent[0]+sy*tangent[1])/(sn*tangentNorm)):1;
        // Una traversa perpendicolare non è il marciapiede parallelo
        // richiesto. Consenti crossing soltanto quando è nel verso
        // dell'effettivo percorso sorgente.
        if(parallel<.53)continue;
        let score=candidate.distance+(penalty[seg.kind]||0)+(1-parallel)*5;
        if(previous) {
          const progress=metres(previous.point,candidate.point);
          const sourceStep=metres(previous.source,point);
          if(progress>sourceStep*1.85+5)continue; // non saltare carreggiata
          if(seg.wayId===previous.wayId)score-=2;
          else score+=Math.min(6,Math.abs(progress-sourceStep)*.8);
        }
        options.push({point:candidate.point,distance:candidate.distance,
          wayId:seg.wayId,kind:seg.kind,score,source:point});
      }
    }
    if(!options.length)return null;
    options.sort((a,b)=>a.score-b.score);
    // In caso di distanze identiche su due marciapiedi contrapposti,
    // non scegliere casualmente il lato per il primo punto.
    if(!previous && options.length>1 && options[0].wayId!==options[1].wayId
       && Math.abs(options[0].score-options[1].score)<.45
       && metres(options[0].point,options[1].point)>4) return null;
    return options[0];
  }
  function automaticPart(originalPart,waysOrIndex) {
    const original=cleanPart(originalPart);
    if(original.length<2)return null;
    const samples=samplePart(original);
    if(samples.length<2)return null;
    const index=waysOrIndex instanceof Map?waysOrIndex:indexWays(waysOrIndex||[]);
    if(!index.size)return null;
    const output=[], covered=[];
    let last=null, snapped=0, maxShift=0;
    for(let i=0;i<samples.length;i++){
      const tangentAt=(i===0)?[samples[0],samples[1]]
        :i===samples.length-1?[samples[i-1],samples[i]]
        :[samples[i-1],samples[i+1]];
      const a=xy(tangentAt[0],samples[i][0]),b=xy(tangentAt[1],samples[i][0]);
      const target=closest(samples[i],[b[0]-a[0],b[1]-a[1]],index,last);
      const p=target?target.point:samples[i];
      output.push(p);
      covered.push(!!target);
      if(target) {
        snapped++;
        maxShift=Math.max(maxShift,target.distance);
        last=target;
      }else last=null;
    }
    const fraction=snapped/samples.length;
    if(snapped<2||fraction<.50)return null;
    const moved=output.filter((p,i)=>metres(p,samples[i])>.6).length;
    if(moved<2)return null;
    // Mantieni tutti i pezzi originali, anche dove manca la mappatura OSM.
    // I campioni ravvicinati (6 m) evitano salti fra vertici distanti.
    // Nessun intervento diventa più corto per una corrispondenza parziale.
    const parts=[output];
    return {original, candidate:output,parts,moved,coverage:fraction,maxShift};
  }
  function boxFor(section) {
    const coords=section.stops.flatMap(s=>cloneParts(s).flat());
    if(!coords.length)return null;
    const south=Math.min(...coords.map(x=>x[0]))-.0015;
    const north=Math.max(...coords.map(x=>x[0]))+.0015;
    const west=Math.min(...coords.map(x=>x[1]))-.002;
    const east=Math.max(...coords.map(x=>x[1]))+.002;
    if(north-south>.075||east-west>.095)return null;
    return [south,west,north,east].map(n=>n.toFixed(6)).join(",");
  }
  async function fetchWays(bbox) {
    const query='[out:json][timeout:22];('+
      'way["highway"="footway"]('+bbox+');'+
      'way["highway"="pedestrian"]('+bbox+');'+
      'way["highway"="path"]["foot"!="no"]('+bbox+'););out geom;';
    let lastError;
    for(const endpoint of overpassURLs){
      const controller=new AbortController(),timeout=setTimeout(()=>controller.abort(),25000);
      try {
        const response=await fetch(endpoint,{
          method:"POST",
          headers:{"Content-Type":"application/x-www-form-urlencoded;charset=UTF-8"},
          body:new URLSearchParams({data:query}),signal:controller.signal
        });
        if(!response.ok)throw Error("OSM HTTP "+response.status);
        const json=await response.json();
        return (Array.isArray(json.elements)?json.elements:[]).slice(0,MAX_WAYS)
          .filter(e=>e.type==="way"&&Array.isArray(e.geometry))
          .map(e=>({id:e.id,tags:e.tags||{},points:e.geometry.map(p=>[Number(p.lat),Number(p.lon)])
            .filter(p=>p.every(Number.isFinite))})).filter(e=>e.points.length>=2);
      }catch(error){lastError=error;}
      finally{clearTimeout(timeout);}
    }
    throw lastError || Error("OSM non disponibile");
  }
  let running=false;
  async function run({db,admin,uid,onProgress,existing={},force=false}={}) {
    if(!admin||!db||running||!global.VARGA_SNOW_CM_CATALOG)return null;
    const key="varga-snow-sidewalk-auto-"+VERSION;
    if(!force)try{
      const last=Number(global.localStorage?.getItem(key)||0);
      if(Date.now()-last<24*3600*1000)return null;
    }catch(_){}
    running=true;
    const progress=message=>{if(typeof onProgress==="function")onProgress(message);};
    let saved=0,unchanged=0,unavailable=0,matched=0;
    try {
      for(const section of global.VARGA_SNOW_CM_CATALOG.sections){
        const bbox=boxFor(section);
        if(!bbox){unavailable+=section.stops.length;continue;}
        let ways;
        try{ways=await fetchWays(bbox);}
        catch(error){
          unavailable+=section.stops.length;
          progress("Riferimenti pedonali OSM non disponibili in "+section.name+": geometrie conservate.");
          continue;
        }
        const index=indexWays(ways);
        for(const stop of section.stops){
          const previous=existing[stop.key];
          const previousAuthor=String(previous?.updatedBy||"");
          // Le geometrie disegnate manualmente da amministratori prevalgono.
          if(previous?.parts?.length && !previousAuthor.startsWith("snow-sidewalk-auto/") &&
             !previousAuthor.startsWith("snow-pedestrian-auto/")){unchanged++;continue;}
          if(previousAuthor.includes(VERSION)){unchanged++;continue;}
          const original=cloneParts(stop);
          const candidate=[];
          let moved=0;
          for(const part of original){
            const aligned=automaticPart(part,index);
            if(aligned){
              moved+=aligned.moved;
              // Mantieni sia i tratti riconosciuti sia quelli ancora da
              // verificare, senza alterare la lista degli interventi.
              candidate.push(...aligned.parts);
            }else candidate.push(part);
          }
          if(!moved||candidate.some(p=>p.length<2)||
             candidate.reduce((sum,p)=>sum+p.length,0)>MAX_STOP_POINTS){
            unchanged++;continue;
          }
          matched++;
          try{
            const ref=db.collection("snowRouteGeometry").doc(stop.key);
            const updated=await db.runTransaction(async tx=>{
              const snap=await tx.get(ref);
              if(snap.exists && Array.isArray(snap.data()?.parts) && snap.data().parts.length){
                const editor=String(snap.data()?.updatedBy||"");
                if(!editor.startsWith("snow-sidewalk-auto/")&&!editor.startsWith("snow-pedestrian-auto/"))return false;
                if(editor.includes(VERSION))return false;
              }
              // Firestore rules consentono SOLO questi quattro campi.
              // L'originale rimane in snow-castel-maggiore-kml-real.js
              // o nel catalogo iniziale, senza campi extra rifiutati.
              tx.set(ref,{
                stopKey:stop.key,parts:candidate,
                updatedAt:global.firebase.firestore.FieldValue.serverTimestamp(),
                updatedBy:"snow-sidewalk-auto/"+VERSION+"/"+String(uid||"admin")
              },{merge:true});
              return true;
            });
            if(updated)saved++;else unchanged++;
          }catch(error){
            unavailable++;
            progress("Non riesco a salvare "+stop.key+" ("+(error?.code||error?.message||"errore")+").");
          }
        }
        progress("Marciapiedi: "+saved+" interventi corretti, "+unchanged+" invariati. Tracciati pedonali trovati: "+ways.length+".");
      }
      if(!unavailable)try{global.localStorage?.setItem(key,String(Date.now()));}catch(_){}
      progress("Allineamento automatico concluso: "+saved+" interventi aggiornati; "+unchanged+" invariati; "+unavailable+" non verificabili. Gli originali restano disponibili.");
      return {saved,unchanged,unavailable,matched};
    }finally{running=false;}
  }
  global.VargaSnowAutomaticSidewalks={run,automaticPart,metres,indexWays,VERSION};
})(window);
