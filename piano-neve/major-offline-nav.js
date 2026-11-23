(function(){
  const data=window.MAJOR_OFFLINE||{roads:[],boundary:null};
  let graph=null,nodeList=null,voiceEnabled=true,lastSpeechKey='';
  const R=6371000;
  function distance(a,b){const p=Math.PI/180,dLat=(b[0]-a[0])*p,dLon=(b[1]-a[1])*p,q=Math.sin(dLat/2)**2+Math.cos(a[0]*p)*Math.cos(b[0]*p)*Math.sin(dLon/2)**2;return 2*R*Math.asin(Math.sqrt(q))}
  function bearing(a,b){const p=Math.PI/180,y=Math.sin((b[1]-a[1])*p)*Math.cos(b[0]*p),x=Math.cos(a[0]*p)*Math.sin(b[0]*p)-Math.sin(a[0]*p)*Math.cos(b[0]*p)*Math.cos((b[1]-a[1])*p);return(Math.atan2(y,x)*180/Math.PI+360)%360}
  function key(p){return Number(p[0]).toFixed(7)+','+Number(p[1]).toFixed(7)}
  function oneWayMode(v){v=String(v||'').toLowerCase();if(v==='-1')return-1;if(v==='yes'||v==='true'||v==='1')return 1;return 0}
  function buildGraph(){
    if(graph)return graph;
    const nodes=new Map(),adj=new Map();
    const addNode=p=>{const k=key(p);if(!nodes.has(k)){nodes.set(k,[+p[0],+p[1]]);adj.set(k,[])}return k};
    const addEdge=(a,b,w,name,roadId)=>adj.get(a).push({to:b,w,name:name||'',roadId});
    const sourceRoads=(data.routingRoads&&data.routingRoads.length)?data.routingRoads:(data.roads||[]);
    for(const road of sourceRoads){
      const access=String(road.access||'').toLowerCase(),motor=String(road.motor_vehicle||'').toLowerCase();
      if((access==='no'||access==='private'||motor==='no'||motor==='private')&&motor!=='yes'&&motor!=='designated')continue;
      const pts=road.coords||[],ow=oneWayMode(road.oneway);
      for(let i=1;i<pts.length;i++){
        const a=addNode(pts[i-1]),b=addNode(pts[i]),w=distance(pts[i-1],pts[i]);
        if(ow>=0)addEdge(a,b,w,road.name,road.id);
        if(ow<=0)addEdge(b,a,w,road.name,road.id);
      }
    }
    nodeList=[...nodes.entries()].map(([k,p])=>({k,p}));
    graph={nodes,adj};return graph
  }
  function nearestNode(p){
    buildGraph();let best=null,bestD=Infinity;
    for(const n of nodeList){const dd=distance(p,n.p);if(dd<bestD){bestD=dd;best=n}}
    return best?{...best,dist:bestD}:null
  }
  class Heap{
    constructor(){this.a=[]}
    push(x){this.a.push(x);let i=this.a.length-1;while(i){const p=(i-1)>>1;if(this.a[p].f<=x.f)break;this.a[i]=this.a[p];i=p}this.a[i]=x}
    pop(){if(!this.a.length)return null;const root=this.a[0],last=this.a.pop();if(this.a.length){let i=0;while(true){let l=i*2+1,r=l+1;if(l>=this.a.length)break;let c=r<this.a.length&&this.a[r].f<this.a[l].f?r:l;if(this.a[c].f>=last.f)break;this.a[i]=this.a[c];i=c}this.a[i]=last}return root}
    get length(){return this.a.length}
  }
  function angleDelta(a,b){return((b-a+540)%360)-180}
  function instructionFor(delta,name){
    const n=name?' su '+name:'';
    if(Math.abs(delta)<28)return{type:'straight',text:'Continua'+n};
    if(delta>145||delta<-145)return{type:'uturn',text:'Fai inversione'+n};
    if(delta>0)return{type:'right',text:'Gira a destra'+n};
    return{type:'left',text:'Gira a sinistra'+n}
  }
  function buildManeuvers(points,names){
    const out=[];if(points.length<3)return out;
    let lastName=names[0]||'';
    for(let i=1;i<points.length-1;i++){
      const n=names[i]||lastName,changed=n&&n!==lastName,b1=bearing(points[i-1],points[i]),b2=bearing(points[i],points[i+1]),delta=angleDelta(b1,b2);
      if(changed||Math.abs(delta)>=48){const ins=instructionFor(delta,n);out.push({index:i,point:points[i],name:n,...ins});lastName=n||lastName}
      else if(n)lastName=n;
    }
    out.push({index:points.length-1,point:points[points.length-1],type:'arrive',text:'Sei arrivato alla destinazione',name:lastName});
    return out
  }
  function route(from,to){
    const g=buildGraph();if(!data.roads?.length)return null;
    const s=nearestNode(from),t=nearestNode(to);if(!s||!t)return null;
    const open=new Heap(),gScore=new Map([[s.k,0]]),prev=new Map(),closed=new Set();
    open.push({k:s.k,f:distance(s.p,t.p)});
    let found=false,guard=0;
    while(open.length&&guard++<200000){
      const cur=open.pop();if(closed.has(cur.k))continue;if(cur.k===t.k){found=true;break}closed.add(cur.k);
      for(const e of g.adj.get(cur.k)||[]){
        if(closed.has(e.to))continue;const ng=(gScore.get(cur.k)||0)+e.w;
        if(ng<(gScore.get(e.to)??Infinity)){gScore.set(e.to,ng);prev.set(e.to,{from:cur.k,name:e.name});open.push({k:e.to,f:ng+distance(g.nodes.get(e.to),t.p)})}
      }
    }
    if(!found)return null;
    const keys=[t.k],names=[];let cur=t.k;
    while(cur!==s.k){const p=prev.get(cur);if(!p)return null;names.push(p.name||'');cur=p.from;keys.push(cur)}
    keys.reverse();names.reverse();
    const core=keys.map(k=>g.nodes.get(k)),points=[from,...core,to];
    const segNames=['',...names,names[names.length-1]||''];
    let total=0;for(let i=1;i<points.length;i++)total+=distance(points[i-1],points[i]);
    const maneuvers=buildManeuvers(points,segNames);
    return{points,names:segNames,distance:total,maneuvers,startSnap:s.dist,endSnap:t.dist}
  }
  function voices(){return typeof speechSynthesis!=='undefined'?speechSynthesis.getVoices():[]}
  function speak(text,keyValue='',force=false){
    if(!voiceEnabled||!text||typeof speechSynthesis==='undefined')return false;
    if(keyValue&&keyValue===lastSpeechKey&&!force)return false;
    lastSpeechKey=keyValue||text;
    const u=new SpeechSynthesisUtterance(text),vs=voices(),v=vs.find(x=>/^it(-|_)/i.test(x.lang)&&x.localService)||vs.find(x=>/^it(-|_)/i.test(x.lang));
    if(v)u.voice=v;u.lang='it-IT';u.rate=1;u.pitch=1;u.volume=1;
    if(force)speechSynthesis.cancel();speechSynthesis.speak(u);return true
  }
  function setVoiceEnabled(v){voiceEnabled=!!v;if(!voiceEnabled&&typeof speechSynthesis!=='undefined')speechSynthesis.cancel()}
  function isVoiceEnabled(){return voiceEnabled}
  function resetSpeech(){lastSpeechKey='';if(typeof speechSynthesis!=='undefined')speechSynthesis.cancel()}
  window.MajorOfflineNav={data,ready:()=>!!(data.boundary&&(data.roads||[]).length),distance,bearing,route,buildGraph,buildManeuvers,speak,setVoiceEnabled,isVoiceEnabled,resetSpeech};
})();