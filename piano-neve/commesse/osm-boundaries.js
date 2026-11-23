/* On-demand OpenStreetMap administrative boundaries. No bulk scraping. */
(function(root){
'use strict';
const CACHE='varga-osm-territories-v1',MAX_AGE=30*86400000;
function cached(){try{const v=JSON.parse(localStorage.getItem(CACHE)||'null');return v&&Date.now()-v.saved<MAX_AGE?v.features:[]}catch{return []}}
function store(features){try{localStorage.setItem(CACHE,JSON.stringify({saved:Date.now(),features}))}catch{}}
function polygon(g){return g&&['Polygon','MultiPolygon'].includes(g.type)&&Array.isArray(g.coordinates)}
function validName(s){return typeof s==='string'&&s.length>1&&s.length<100&&!/[<>]/.test(s)}
async function loadMunicipality(name){
 if(!validName(name))throw Error('Comune non valido');
 const saved=cached(),found=saved.find(f=>f.properties?.municipality?.toLowerCase()===name.toLowerCase()&&f.properties?.source==='OpenStreetMap');if(found)return found;
 const u=new URL('https://nominatim.openstreetmap.org/search');u.searchParams.set('q',name+', Emilia-Romagna, Italia');u.searchParams.set('format','jsonv2');u.searchParams.set('polygon_geojson','1');u.searchParams.set('addressdetails','1');u.searchParams.set('limit','8');
 const response=await fetch(u,{headers:{Accept:'application/json'}});if(!response.ok)throw Error('OSM non disponibile: HTTP '+response.status);
 const rows=await response.json();const row=rows.find(x=>polygon(x.geojson)&&['administrative','city','town','village'].includes(x.type)&&String(x.display_name||'').toLowerCase().includes('emilia-romagna')&&(x.addresstype==='city'||x.addresstype==='town'||x.addresstype==='municipality'||x.type==='administrative'))||rows.find(x=>polygon(x.geojson)&&x.addresstype==='municipality');
 if(!row)throw Error('Confine comunale OSM non trovato: '+name);
 const feature={type:'Feature',properties:{municipality:name,neighborhood:'',source:'OpenStreetMap',osm_id:row.osm_id,osm_type:row.osm_type},geometry:row.geojson};
 const next=[...saved.filter(f=>f.properties?.municipality?.toLowerCase()!==name.toLowerCase()),feature];store(next);return feature;
}
function cachedFeatures(){return {type:'FeatureCollection',features:cached()}}
root.VargaOSMBoundaries={loadMunicipality,cachedFeatures};
})(typeof window!=='undefined'?window:globalThis);
