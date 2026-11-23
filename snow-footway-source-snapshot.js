/* Salva i riferimenti pedonali OSM PRIMA di applicare il KML My Maps.
 * La ricostruzione è separata dagli interventi e dagli stati di servizio.
 */
(function(window) {
  "use strict";
  const catalog=window.VARGA_SNOW_CM_CATALOG;
  if(!catalog?.sections) return;
  const refs={};
  for(const section of catalog.sections)for(const stop of section.stops){
    refs[stop.key]=(stop.parts||[]).filter(p=>p.source==="osm" || p.source==="dbtr-acp")
      .map(p=>({source:p.source,coords:(p.coords||[]).map(x=>[Number(x[0]),Number(x[1])])}))
      .filter(p=>p.coords.length>=2);
  }
  window.VARGA_SNOW_CM_PEDESTRIAN_REFERENCES=refs;
})(window);
