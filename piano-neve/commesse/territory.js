/* Varga Piano Neve — preparazione commesse: nessuna importazione o commessa creata. */
(function(root){
'use strict';
const norm=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().replace(/\b(via|viale|piazza|strada|vicolo)\b/g,'').replace(/[^a-z0-9]+/g,' ').trim();
const key=s=>String(s??'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
function pointInRing(point,ring){let inside=false;const x=point[0],y=point[1];for(let i=0,j=ring.length-1;i<ring.length;j=i++){const a=ring[i],b=ring[j];if((a[1]>y)!==(b[1]>y)&&x<(b[0]-a[0])*(y-a[1])/(b[1]-a[1])+a[0])inside=!inside;}return inside;}
function inGeometry(point,geometry){if(!geometry)return false;if(geometry.type==='Polygon')return pointInRing(point,geometry.coordinates[0])&&!geometry.coordinates.slice(1).some(r=>pointInRing(point,r));if(geometry.type==='MultiPolygon')return geometry.coordinates.some(poly=>inGeometry(point,{type:'Polygon',coordinates:poly}));return false;}
function asFeatures(data){return data?.type==='FeatureCollection'?data.features:data?.type==='Feature'?[data]:[];}
function territoryOf(feature){const p=feature.properties||{};return {municipality:String(p.municipality||p.comune||p.COMUNE||'').trim(),neighborhood:String(p.neighborhood||p.quartiere||p.QUARTIERE||'').trim(),id:String(feature.id||p.id||'')};}
function createResolver(boundaries){const features=asFeatures(boundaries);return {
 list(){return features.map(territoryOf).filter(t=>t.municipality);},
 locate(point){return features.filter(f=>inGeometry(point,f.geometry)).map(territoryOf);},
 resolveStreet(street,selection={}){const coords=street?.coordinates||street?.geometry?.coordinates||[];const points=street?.geometry?.type==='LineString'?coords:coords;const matches=new Map();for(let i=1;i<points.length;i++){const a=points[i-1],b=points[i],dx=b[0]-a[0],dy=b[1]-a[1],steps=Math.max(1,Math.ceil(Math.hypot(dx,dy)*10000));for(let j=0;j<=steps;j++){const t=j/steps;for(const zone of this.locate([a[0]+dx*t,a[1]+dy*t]))matches.set(key(zone.municipality)+'|'+key(zone.neighborhood),zone);}}const zones=[...matches.values()],municipalities=(selection.municipalities||[]).map(key),neighborhoods=(selection.neighborhoods||[]).map(key);const selected=zones.filter(z=>(!municipalities.length||municipalities.includes(key(z.municipality)))&&(!neighborhoods.length||neighborhoods.includes(key(z.neighborhood))));return {name:street.name||'',normalizedName:norm(street.name),zones,selected,ambiguous:zones.length>1,matched:selected.length>0,requiresReview:zones.length===0};}
};}
root.VargaTerritory={createResolver,normalizeStreetName:norm};
})(typeof window!=='undefined'?window:globalThis);
