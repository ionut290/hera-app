/* Varga Cantieri · Precisione cartografica neve, solo editor admin.
 * Dati di riferimento: ACP DBTR (Regione Emilia-Romagna, CC BY 4.0) e
 * percorsi pedonali OSM (ODbL). Non cambia geometrie o stati senza il
 * salvataggio esplicito dell'amministratore in snow-plan-operations.js.
 */
(function installSnowPedestrianAssist(global) {
  "use strict";
  const WMS = "https://servizigis.regione.emilia-romagna.it/wms/dbtr";
  const OVERPASS = "https://overpass.private.coffee/api/interpreter";
  const SNAP_LIMIT_METERS = 6;
  let session = null;
  const cache = new Map();

  function setStatus(text) {
    if (session?.status) session.status.textContent = text;
  }

  function detach() {
    const old = session;
    session = null;
    if (!old) return;
    old.controller?.abort();
    if (old.timeout) clearTimeout(old.timeout);
    if (old.dbtr && old.map.hasLayer(old.dbtr)) old.map.removeLayer(old.dbtr);
    if (old.osmLayer && old.map.hasLayer(old.osmLayer)) old.map.removeLayer(old.osmLayer);
    old.toolbar?.remove();
  }

  // Calcola il punto della polilinea pedonale più vicino sul piano Leaflet,
  // ma applica un limite reale di sei metri (non uno snap illimitato).
  function nearest(point) {
    const s = session;
    if (!s?.snapEnabled || !s.lines.length || !s.map) return null;
    const origin = L.latLng(point[0], point[1]);
    const p = s.map.latLngToLayerPoint(origin);
    let result = null, best = SNAP_LIMIT_METERS;
    for (const coords of s.lines) {
      const a = s.map.latLngToLayerPoint(coords[0]);
      const b = s.map.latLngToLayerPoint(coords[1]);
      const dx = b.x - a.x, dy = b.y - a.y;
      const len2 = dx*dx + dy*dy;
      const t = len2 ? Math.max(0, Math.min(1, ((p.x-a.x)*dx+(p.y-a.y)*dy)/len2)) : 0;
      const ll = s.map.layerPointToLatLng(L.point(a.x+t*dx, a.y+t*dy));
      const metres = s.map.distance(origin, ll);
      if (metres < best) { best = metres; result = [ll.lat, ll.lng]; }
    }
    return result;
  }
  function snap(point) {
    const coords = [Number(point[0]), Number(point[1])];
    if (!Number.isFinite(coords[0]) || !Number.isFinite(coords[1])) return point;
    return nearest(coords) || coords;
  }

  function showWays(s, elements) {
    if (session !== s) return;
    s.osmLayer.clearLayers();
    s.lines = [];
    let count = 0, sidewalks = 0;
    for (const item of elements) {
      if (count >= 500 || s.lines.length >= 10000) break;
      if (item.type !== "way" || !Array.isArray(item.geometry)) continue;
      const path = item.geometry.map(p => [Number(p.lat), Number(p.lon)])
        .filter(p => Number.isFinite(p[0]) && Number.isFinite(p[1]));
      if (path.length < 2) continue;
      const footway = String(item.tags?.footway || "");
      const isSidewalk = footway === "sidewalk";
      const isCrossing = footway === "crossing";
      L.polyline(path, {
        color: isSidewalk ? "#0891b2" : isCrossing ? "#d97706" : "#64748b",
        weight: isSidewalk ? 3.5 : 2, opacity: .67,
        dashArray: isSidewalk ? null : "4 6", interactive: false
      }).addTo(s.osmLayer);
      for (let i = 1; i < path.length && s.lines.length < 10000; i++)
        s.lines.push([path[i-1], path[i]]);
      count++;
      if (isSidewalk) sidewalks++;
    }
    if (!s.map.hasLayer(s.osmLayer)) s.osmLayer.addTo(s.map);
    s.alignButton.disabled = !s.lines.length;
    if (!count) setStatus("Nessun percorso pedonale OSM separato: usa DBTR e satellite per correggere manualmente.");
    else setStatus(count+" percorsi pedonali OSM ("+sidewalks+" marcati come marciapiedi). Azzurro = marciapiedi, arancione = attraversamenti, grigio = altri percorsi. Controlla sempre il lato.");
  }

  async function loadOSM(s, stop, parts) {
    let points = (Array.isArray(parts) ? parts : []).flat()
      .filter(p => Array.isArray(p) && Number.isFinite(Number(p[0])) && Number.isFinite(Number(p[1])));
    if (!points.length && Array.isArray(stop.marker)) points = [stop.marker];
    if (!points.length) return setStatus("Geometria assente: impossibile caricare riferimenti pedonali.");
    const bounds = L.latLngBounds(points), center = bounds.getCenter();
    const box = [
      Math.max(bounds.getSouth()-.0015, center.lat-.012),
      Math.max(bounds.getWest()-.002, center.lng-.012),
      Math.min(bounds.getNorth()+.0015, center.lat+.012),
      Math.min(bounds.getEast()+.002, center.lng+.012)
    ].map(n => n.toFixed(6)).join(",");
    const key = stop.key + ":" + box;
    const cached = cache.get(key);
    if (cached && Date.now()-cached.at < 600000) {
      showWays(s,cached.elements);
      return;
    }
    const query = '[out:json][timeout:18];('+
      'way["highway"="footway"]('+box+');'+
      'way["highway"="pedestrian"]('+box+');'+
      'way["highway"="path"]["foot"!="no"]('+box+'););out geom;';
    s.controller = new AbortController();
    s.timeout = setTimeout(() => s.controller?.abort(), 18000);
    setStatus("Ricerca dei marciapiedi e percorsi pedonali su OpenStreetMap...");
    try {
      const resp = await fetch(OVERPASS, {
        method:"POST",
        headers:{"Content-Type":"application/x-www-form-urlencoded;charset=UTF-8"},
        body:new URLSearchParams({data:query}),
        signal:s.controller.signal
      });
      if (!resp.ok) throw new Error("HTTP "+resp.status);
      const json = await resp.json();
      if (session !== s) return;
      const ways = Array.isArray(json.elements) ? json.elements : [];
      cache.set(key,{at:Date.now(),elements:ways});
      if (cache.size > 15) cache.delete(cache.keys().next().value);
      showWays(s,ways);
    } catch (err) {
      if (session !== s) return;
      setStatus("Servizio OSM temporaneamente indisponibile. DBTR e disegno manuale restano utilizzabili.");
    } finally {
      if (s.timeout) clearTimeout(s.timeout);
      s.timeout = null;
      s.controller = null;
    }
  }

  // Applica una proposta geometrica SOLO all'anteprima in memoria.
  function alignPreview(parts) {
    const s = session;
    if (!s || !s.lines.length) return {parts,moved:0};
    let moved = 0;
    const candidate = parts.map(part => part.map(point => {
      const adjusted = snap(point);
      if (s.map.distance(point, adjusted) < .75) return point;
      moved++;
      return adjusted;
    }));
    return {parts:candidate,moved};
  }

  function attach(config) {
    detach();
    if (!config?.map || !config?.stop || !global.L || !global.fetch) return;
    const editor = document.getElementById("snow-plan-editor");
    if (!editor) return;
    const toolbar = document.createElement("div");
    toolbar.className = "snow-map-precision-tools";
    toolbar.innerHTML =
      '<strong>🎯 Precisione marciapiedi</strong>'+
      '<button type="button" data-precision="dbtr" aria-pressed="true">DBTR pedonale: ON</button>'+
      '<label><input type="checkbox" data-precision="snap" checked> Aggancio OSM (6 m)</label>'+
      '<button type="button" data-precision="align" disabled>Allinea tracce · anteprima</button>'+
      '<span data-precision="status" role="status" aria-live="polite">Avvio riferimenti...</span>';
    editor.appendChild(toolbar);
    const s = session = {
      map:config.map, stop:config.stop, toolbar, controller:null, timeout:null,
      lines:[], snapEnabled:true, dbtr:null, osmLayer:L.layerGroup(),
      status:toolbar.querySelector('[data-precision="status"]'),
      alignButton:toolbar.querySelector('[data-precision="align"]')
    };
    s.dbtr = L.tileLayer.wms(WMS, {
      layers:"ACP_Area_di_circolazione_pedonale",
      format:"image/png",transparent:true,version:"1.1.1",
      opacity:.58,attribution:"© Regione Emilia-Romagna · DBTR (CC BY 4.0)"
    }).addTo(s.map);
    const dbtrBtn = toolbar.querySelector('[data-precision="dbtr"]');
    dbtrBtn.addEventListener("click", () => {
      if (s.map.hasLayer(s.dbtr)) s.map.removeLayer(s.dbtr);
      else s.dbtr.addTo(s.map);
      const enabled = s.map.hasLayer(s.dbtr);
      dbtrBtn.textContent = "DBTR pedonale: "+(enabled?"ON":"OFF");
      dbtrBtn.setAttribute("aria-pressed",String(enabled));
    });
    toolbar.querySelector('[data-precision="snap"]').addEventListener("change",e=>{
      s.snapEnabled = !!e.target.checked;
    });
    s.alignButton.addEventListener("click",()=>{
      if (!s.lines.length) return setStatus("Nessun riferimento OSM disponibile per l'allineamento.");
      if (!global.confirm("Allineare SOLO IN ANTEPRIMA i vertici entro 6 m dalle linee pedonali OSM? Controlla lato e attraversamenti prima di salvare.")) return;
      const result = alignPreview(config.getParts());
      if (!result.moved) return setStatus("Nessun vertice da correggere entro 6 m dai percorsi OSM disponibili.");
      config.onPreview(result.parts);
      setStatus(result.moved+" vertici spostati in ANTEPRIMA. Controlla con DBTR e satellite; premi Salva mappa solo dopo la verifica.");
    });
    loadOSM(s, config.stop, config.parts);
  }

  global.VargaSnowPedestrianAssist = {attach,detach,snap,alignPreview};
})(window);
