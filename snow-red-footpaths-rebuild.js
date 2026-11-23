/* Varga Cantieri | Ricostruzione geometrica lungo i percorsi ROSSI OSM.
 * Geometrie pedonali conservate dal catalogo originale prima del KML.
 * Non è un offset: ogni segmento verificato viene tracciato sulla
 * geometria vettoriale OSM (vertici e curve originali).
 *
 * Dove manca la fonte pedonale il tratto My Maps resta INVARIATO e
 * visivamente tratteggiato. Non inventare passaggi o cambiare lato.
 */
(function installVargaSnowRedFootpaths(window) {
  "use strict";
  const VERSION = "20261123-red-footway-rebuild1";
  const C = window.VARGA_SNOW_CM_CATALOG;
  const refs = window.VARGA_SNOW_CM_PEDESTRIAN_REFERENCES;
  if (!C?.sections?.length || !refs) return;
  const EARTH = 6371000, R = Math.PI/180;
  const GRID_LAT = .00035, GRID_LNG = .00048;
  const MAX_SAME_STOP_METRES = 15, MAX_OTHER_STOP_METRES = 8;
  const SAMPLE_GAP_METRES = 4;
  const MAX_SAMPLES_PER_PART = 2600;
  const index=new Map(), routes=new Map();
  const keyOf=(a,b)=>Math.floor(a/GRID_LAT)+":"+Math.floor(b/GRID_LNG);
  const clonePoint=p=>[Number(p[0]),Number(p[1])];
  const valid=p=>Array.isArray(p)&&p.length>=2&&Number.isFinite(+p[0])&&Number.isFinite(+p[1]);
  const xy=(p,lat)=>[EARTH*R*p[1]*Math.cos(R*lat),EARTH*R*p[0]];
  function distance(a,b) {
    const p=xy(a,a[0]),q=xy(b,a[0]);
    return Math.hypot(p[0]-q[0],p[1]-q[1]);
  }
  function pointOnSegment(p,a,b) {
    const P=xy(p,p[0]),A=xy(a,p[0]),B=xy(b,p[0]);
    const dx=B[0]-A[0],dy=B[1]-A[1],l=dx*dx+dy*dy;
    const t=l?Math.max(0,Math.min(1,((P[0]-A[0])*dx+(P[1]-A[1])*dy)/l)):0;
    const point=[a[0]+t*(b[0]-a[0]),a[1]+t*(b[1]-a[1])];
    return {point,t,distance:distance(p,point)};
  }
  function addIndex(segment) {
    const {a,b}=segment;
    const minLat=Math.floor((Math.min(a[0],b[0])-.00025)/GRID_LAT);
    const maxLat=Math.floor((Math.max(a[0],b[0])+.00025)/GRID_LAT);
    const minLon=Math.floor((Math.min(a[1],b[1])-.00035)/GRID_LNG);
    const maxLon=Math.floor((Math.max(a[1],b[1])+.00035)/GRID_LNG);
    if ((maxLat-minLat+1)*(maxLon-minLon+1)>150) return;
    for(let row=minLat;row<=maxLat;row++)for(let col=minLon;col<=maxLon;col++){
      const key=row+":"+col;
      let items=index.get(key);if(!items){items=[];index.set(key,items);}
      items.push(segment);
    }
  }
  // SOLO le coordinate OSM originali: i tracciati stradali
  // centerline/DBTR non sono un marciapiede rosso OSM.
  for(const [stopKey,parts] of Object.entries(refs)){
    for(let id=0;id<parts.length;id++){
      const part=parts[id];
      if(part.source!=="osm")continue;
      const coords=(part.coords||[]).filter(valid).map(clonePoint);
      if(coords.length<2)continue;
      const routeId=stopKey+":"+id;
      const route={routeId,stopKey,coords};
      routes.set(routeId,route);
      for(let j=1;j<coords.length;j++){
        const a=coords[j-1],b=coords[j],length=distance(a,b);
        if(length<.05||length>220)continue;
        addIndex({routeId,stopKey,a,b,segmentIndex:j-1,length});
      }
    }
  }
  function getSegments(point) {
    const row=Math.floor(point[0]/GRID_LAT),col=Math.floor(point[1]/GRID_LNG);
    const found=[],seen=new Set();
    for(let a=-1;a<=1;a++)for(let b=-1;b<=1;b++){
      for(const segment of (index.get((row+a)+":"+(col+b))||[])){
        if(seen.has(segment))continue;
        seen.add(segment);found.push(segment);
      }
    }
    return found;
  }
  function sampleLine(points) {
    const samples=[clonePoint(points[0])];
    for(let i=1;i<points.length;i++){
      const a=points[i-1],b=points[i],d=distance(a,b);
      if(!Number.isFinite(d)||d>1500)return [];
      const steps=Math.max(1,Math.ceil(d/SAMPLE_GAP_METRES));
      if(samples.length+steps>MAX_SAMPLES_PER_PART)return [];
      for(let j=1;j<=steps;j++){
        const t=j/steps;
        samples.push([a[0]+t*(b[0]-a[0]),a[1]+t*(b[1]-a[1])]);
      }
    }
    return samples;
  }
  function candidateAt(p,tangent,stopKey,prev) {
    const d=T=>Math.hypot(T[0],T[1]);
    const norm=d(tangent), options=[];
    for(const segment of getSegments(p)){
      const cap=segment.stopKey===stopKey?MAX_SAME_STOP_METRES:MAX_OTHER_STOP_METRES;
      const projected=pointOnSegment(p,segment.a,segment.b);
      if(projected.distance>cap)continue;
      const A=xy(segment.a,p[0]),B=xy(segment.b,p[0]);
      const vx=B[0]-A[0],vy=B[1]-A[1],ln=Math.hypot(vx,vy);
      if(ln<.1)continue;
      const parallel=norm?Math.abs((vx*tangent[0]+vy*tangent[1])/(ln*norm)):1;
      if(parallel<.52)continue;
      let score=projected.distance+(1-parallel)*7;
      if(segment.stopKey!==stopKey)score+=1.8;
      if(prev){
        const travel=distance(prev.point,projected.point);
        const originalTravel=distance(prev.source,p);
        if(travel>originalTravel+10)continue;
        if(travel>originalTravel*2+6)continue;
        if(segment.routeId===prev.routeId)score-=3.5;
        else score+=3+Math.max(0,travel-originalTravel)*1.3;
      }
      options.push({point:projected.point,segmentIndex:segment.segmentIndex,
        t:projected.t,source:p,routeId:segment.routeId,
        distance:projected.distance,score});
    }
    if(!options.length)return null;
    options.sort((a,b)=>a.score-b.score);
    // Evita l'aggancio casuale all'altro lato se due candidati
    // risultano quasi equivalenti e molto distanti tra loro.
    if(!prev && options.length>1 && Math.abs(options[0].score-options[1].score)<.35 &&
        distance(options[0].point,options[1].point)>5)return null;
    return options[0];
  }
  function originalSamples(points) {
    const out=points.filter(valid).map(clonePoint);
    return out.length>=2?out:[];
  }
  function nativeSpan(a,b) {
    if(!a||!b)return null;
    const rawStep=distance(a.source,b.source);
    if(a.routeId!==b.routeId) {
      // Soltanto un collegamento esistente nello stesso nodo OSM:
      // mai saltare la carreggiata fra marciapiedi paralleli.
      if(distance(a.point,b.point)>Math.min(2.4,rawStep+1))return null;
      return [a.point,b.point];
    }
    const route=routes.get(a.routeId);if(!route)return null;
    let out=[a.point],n=route.coords.length;
    const from=a.segmentIndex,to=b.segmentIndex;
    if(Math.abs(from-to)>30)return null;
    if(from<to || (from===to&&a.t<=b.t)){
      for(let k=from+1;k<=to;k++)out.push(route.coords[k]);
    }else if(from>to || (from===to&&a.t>b.t)){
      for(let k=from;k>to;k--)out.push(route.coords[k]);
    }
    out.push(b.point);
    out=out.filter((p,i)=>i===0||distance(p,out[i-1])>.03);
    let length=0;
    for(let i=1;i<out.length;i++)length+=distance(out[i-1],out[i]);
    if(length>rawStep*2.3+8)return null;
    return out.length>=2?out:[a.point,b.point];
  }
  function matchPart(source,stopKey) {
    const pts=originalSamples(source),samples=pts.length?sampleLine(pts):[];
    if(samples.length<2)return null;
    const chosen=[];
    let previous=null;
    for(let i=0;i<samples.length;i++){
      const t0=samples[Math.max(0,i-1)],t1=samples[Math.min(samples.length-1,i+1)];
      const A=xy(t0,samples[i][0]),B=xy(t1,samples[i][0]);
      const tangent=[B[0]-A[0],B[1]-A[1]];
      const result=candidateAt(samples[i],tangent,stopKey,previous);
      chosen.push(result);
      previous=result;
    }
    const pieces=[];
    let seenDistance=0, matchedDistance=0;
    function append(points,matched) {
      if(points.length<2)return;
      const prior=pieces.at(-1);
      if(prior && prior.matched===matched && distance(prior.coords.at(-1),points[0])<.08)
        prior.coords.push(...points.slice(1));
      else pieces.push({coords:points,source:matched?"osm-red-footway":"user-mymaps-kml",matched});
    }
    for(let i=1;i<samples.length;i++){
      const p=samples[i-1],q=samples[i],segmentLength=distance(p,q);
      seenDistance+=segmentLength;
      const span=nativeSpan(chosen[i-1],chosen[i]);
      if(span){
        append(span,true);matchedDistance+=segmentLength;
      }else append([p,q],false);
    }
    return {parts:pieces,matchedDistance,seenDistance};
  }
  let matchingCount=0,fullyMatched=0;
  const results={version:VERSION,stops:{},matchedStops:[],partialStops:[],unmatchedStops:[]};
  for(const section of C.sections){
    for(const stop of section.stops){
      const parts=stop.parts||[],next=[],lenBefore=parts.length;
      let length=0,matched=0;
      for(const part of parts){
        const m=matchPart(part.coords||[],stop.key);
        if(!m){next.push(part);continue;}
        length+=m.seenDistance;matched+=m.matchedDistance;
        next.push(...m.parts);
      }
      const coverage=length?matched/length:0;
      const newMatched=next.filter(x=>x.source==="osm-red-footway"&&x.coords?.length>=2).length;
      const sufficient=coverage>=.35 && newMatched>=1;
      if(sufficient){
        stop.parts=next;
        stop.geometryStatus="osm-red-footway-rebuilt";
        stop.redFootwayRebuild={version:VERSION,coverage:Number(coverage.toFixed(4)),
          matchedParts:newMatched,originalParts:lenBefore,
          complete:coverage>=.985 && next.every(x=>x.matched===true)};
        matchingCount++;
        if(stop.redFootwayRebuild.complete)fullyMatched++;
        (stop.redFootwayRebuild.complete?results.matchedStops:results.partialStops).push(stop.key);
      }else results.unmatchedStops.push(stop.key);
      results.stops[stop.key]={coverage:+coverage.toFixed(4),applied:sufficient,
        matchedParts:newMatched,complete:!!stop.redFootwayRebuild?.complete};
    }
  }
  results.rebuilt=matchingCount;
  results.complete=fullyMatched;
  results.partial=results.partialStops.length;
  results.unmatched=results.unmatchedStops.length;
  results.referencePaths=routes.size;

  // Pubblica correzioni permanenti solo se TUTTO il tracciato è stato
  // riconosciuto sulla geometria OSM. Gli stati di lavoro restano separati.
  // Le modifiche manuali di un amministratore hanno sempre precedenza.
  let syncing=false;
  async function persist({db,admin,uid,existing={},onProgress}={}) {
    if(!admin||!db||syncing)return null;
    syncing=true;
    const feedback=msg=>{if(typeof onProgress==="function")onProgress(msg);};
    const summary={saved:0,unchanged:0,failed:0,partial:results.partial};
    try {
      for(const section of C.sections)for(const stop of section.stops){
        if(!stop.redFootwayRebuild?.complete)continue;
        const cached=existing[stop.key];
        const author=String(cached?.updatedBy||"");
        if(cached?.parts?.length && !/^(snow-sidewalk-auto|snow-pedestrian-auto|snow-red-rebuild)\//.test(author)){
          summary.unchanged++;continue;
        }
        if(author.includes(VERSION)){summary.unchanged++;continue;}
        try{
          const doc=db.collection("snowRouteGeometry").doc(stop.key);
          const saved=await db.runTransaction(async tx=>{
            const current=await tx.get(doc);
            if(current.exists && Array.isArray(current.data()?.parts) && current.data().parts.length){
              const writer=String(current.data()?.updatedBy||"");
              if(!/^(snow-sidewalk-auto|snow-pedestrian-auto|snow-red-rebuild)\//.test(writer)||writer.includes(VERSION))return false;
            }
            tx.set(doc,{
              stopKey:stop.key,
              parts:stop.parts.map(x=>x.coords),
              updatedAt:window.firebase.firestore.FieldValue.serverTimestamp(),
              updatedBy:"snow-red-rebuild/"+VERSION+"/"+String(uid||"admin")
            },{merge:true});
            return true;
          });
          if(saved)summary.saved++;else summary.unchanged++;
        }catch(err){
          summary.failed++;
          feedback("Percorso "+stop.key+" non salvato: "+(err.code||err.message||err));
        }
      }
      feedback("Ricostruzione marciapiedi: "+summary.saved+" interventi salvati; "+
        summary.partial+" parziali da controllare; "+results.unmatched+
        " senza riferimento OSM sufficiente. Nessun tracciato inventato.");
      return summary;
    }finally{syncing=false;}
  }
  window.VargaSnowRedFootpaths={results,persist,VERSION};
})(window);
