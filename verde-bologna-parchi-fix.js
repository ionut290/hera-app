(() => {
  "use strict";

  const PAGE_ID = "verde-bologna-page";
  const CATEGORY_ID = "verde-bologna-operativo-category";
  const PARKS_ID = "carta-tecnica-comunale-toponimi-parchi-e-giardini";
  const QUARTERS_ID = "quartieri-di-bologna";
  const API_ROOT = "https://opendata.comune.bologna.it/api/explore/v2.1/catalog/datasets";
  const PANEL_ID = "vb-parks-fix-panel";
  const MAP_ID = "vb-parks-fix-map";
  const SHEET_ID = "vb-parks-fix-sheet";
  const STYLE_ID = "vb-parks-fix-style";
  const CACHE_TTL = 10 * 60 * 1000;
  const MOBILE = "(max-width: 760px)";
  const QUARTERS = [
    "Borgo Panigale - Reno",
    "Navile",
    "Porto - Saragozza",
    "San Donato - San Vitale",
    "Santo Stefano",
    "Savena"
  ];

  const state = {
    parks: [],
    quarters: [],
    filtered: [],
    loaded: false,
    loading: false,
    activeQuarter: "",
    user: null,
    locationAsked: false,
    map: null,
    markerLayer: null,
    boundaryLayer: null,
    selected: null,
    boundaryCache: new Map(),
    installed: false
  };

  const $ = (id) => document.getElementById(id);
  const esc = (value) => String(value ?? "").replace(/[&<>"']/g, (c) => ({"&":"&amp;","<":"&lt;",">":"&gt;",'"':"&quot;","'":"&#39;"}[c]));
  const norm = (value) => String(value ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLocaleLowerCase("it-IT").trim();
  const keyNorm = (value) => norm(value).replace(/[^a-z0-9]/g, "");

  function mobile() { return window.matchMedia(MOBILE).matches; }
  function active() {
    return mobile() && $(CATEGORY_ID)?.value === PARKS_ID && !$(PAGE_ID)?.classList.contains("hidden");
  }

  function field(record, names) {
    const map = new Map(Object.entries(record || {}).map(([k, v]) => [keyNorm(k), v]));
    for (const name of names) {
      const value = map.get(keyNorm(name));
      if (value !== undefined && value !== null && String(value).trim() !== "") return value;
    }
    return "";
  }

  function codeText(value) {
    const raw = String(value ?? "").trim();
    if (!raw || /^(null|undefined)$/i.test(raw)) return "";
    return raw.replace(/\.0+$/, "");
  }

  function usableCode(value) {
    const raw = codeText(value);
    if (!raw) return "";
    const digits = raw.replace(/[^0-9]/g, "");
    if (digits && /^0+$/.test(digits)) return "";
    return raw;
  }

  function codvia(record) { return codeText(field(record, ["codvia", "cod_via", "codice via"])); }
  function codOgg(record) { return codeText(field(record, ["cod_ogg", "codogg", "codice oggetto", "codice_oggetto"])); }
  function idTopon(record) { return codeText(field(record, ["idtopon", "id_topon"])); }

  function primaryCode(record) {
    const via = usableCode(codvia(record));
    if (via) return { value: via, type: "CODVIA" };
    const ogg = usableCode(codOgg(record));
    if (ogg) return { value: ogg, type: "COD_OGG" };
    const id = usableCode(idTopon(record));
    if (id) return { value: id, type: "IDTOPON" };
    return { value: "—", type: "CODICE" };
  }

  function parkName(record) {
    return String(field(record, ["nomevia", "nome_via", "completo", "porzione", "denominazione", "nome", "name", "toponimo"]) || "Parco / giardino").trim();
  }

  function parseGeometry(value) {
    if (!value) return null;
    if (typeof value === "string") {
      try { return parseGeometry(JSON.parse(value)); } catch (_) { return null; }
    }
    if (value.type === "Feature" && value.geometry) return parseGeometry(value.geometry);
    if (value.geometry) return parseGeometry(value.geometry);
    if (value.type && Array.isArray(value.coordinates)) return value;
    if (Number.isFinite(Number(value.lat)) && Number.isFinite(Number(value.lon ?? value.lng))) {
      return { type: "Point", coordinates: [Number(value.lon ?? value.lng), Number(value.lat)] };
    }
    return null;
  }

  function point(record) {
    const gp = field(record, ["geo_point_2d", "geopoint", "geo point"]);
    if (gp) {
      if (Array.isArray(gp) && gp.length >= 2) {
        const a = Number(gp[0]), b = Number(gp[1]);
        if (Number.isFinite(a) && Number.isFinite(b)) {
          if (Math.abs(a) <= 90 && Math.abs(b) <= 180) return { lat: a, lon: b };
          if (Math.abs(b) <= 90 && Math.abs(a) <= 180) return { lat: b, lon: a };
        }
      }
      if (typeof gp === "object") {
        const lat = Number(gp.lat), lon = Number(gp.lon ?? gp.lng);
        if (Number.isFinite(lat) && Number.isFinite(lon)) return { lat, lon };
      }
    }
    for (const [k, v] of Object.entries(record || {})) {
      const nk = keyNorm(k);
      if (!nk.includes("geoshape") && nk !== "geometry" && nk !== "geom") continue;
      const g = parseGeometry(v);
      if (!g) continue;
      const coords = [];
      const walk = (arr) => {
        if (!Array.isArray(arr)) return;
        if (arr.length >= 2 && Number.isFinite(Number(arr[0])) && Number.isFinite(Number(arr[1]))) {
          coords.push([Number(arr[0]), Number(arr[1])]); return;
        }
        arr.forEach(walk);
      };
      walk(g.coordinates);
      const valid = coords.filter(([lon, lat]) => Math.abs(lon) <= 180 && Math.abs(lat) <= 90);
      if (valid.length) {
        if (g.type === "Point") return { lat: valid[0][1], lon: valid[0][0] };
        const sum = valid.reduce((acc, [lon, lat]) => ({ lon: acc.lon + lon, lat: acc.lat + lat }), { lon: 0, lat: 0 });
        return { lat: sum.lat / valid.length, lon: sum.lon / valid.length };
      }
    }
    return null;
  }

  function geometry(record) {
    for (const [k, v] of Object.entries(record || {})) {
      const nk = keyNorm(k);
      if (!nk.includes("geoshape") && nk !== "geometry" && nk !== "geom") continue;
      const g = parseGeometry(v);
      if (g) return g;
    }
    return null;
  }

  function pointInRing(lon, lat, ring) {
    let inside = false;
    for (let i = 0, j = ring.length - 1; i < ring.length; j = i++) {
      const xi = Number(ring[i]?.[0]), yi = Number(ring[i]?.[1]);
      const xj = Number(ring[j]?.[0]), yj = Number(ring[j]?.[1]);
      if (![xi, yi, xj, yj].every(Number.isFinite)) continue;
      const hit = ((yi > lat) !== (yj > lat)) && lon < ((xj - xi) * (lat - yi)) / ((yj - yi) || Number.EPSILON) + xi;
      if (hit) inside = !inside;
    }
    return inside;
  }

  function pointInPolygon(lon, lat, polygon) {
    if (!Array.isArray(polygon) || !polygon.length || !pointInRing(lon, lat, polygon[0])) return false;
    for (let i = 1; i < polygon.length; i += 1) if (pointInRing(lon, lat, polygon[i])) return false;
    return true;
  }

  function contains(g, p) {
    if (!g || !p) return false;
    if (g.type === "Polygon") return pointInPolygon(p.lon, p.lat, g.coordinates);
    if (g.type === "MultiPolygon") return g.coordinates.some((poly) => pointInPolygon(p.lon, p.lat, poly));
    return false;
  }

  function quarterName(record) {
    for (const value of Object.values(record || {})) {
      if (typeof value !== "string") continue;
      const nv = norm(value);
      const known = QUARTERS.find((q) => norm(q) === nv);
      if (known) return known;
    }
    return String(field(record, ["quartiere", "nomequartiere", "denominazione", "nome"]) || "").trim();
  }

  function assignQuarter(record) {
    const p = point(record);
    if (!p) return "";
    for (const q of state.quarters) {
      if (contains(geometry(q), p)) return quarterName(q);
    }
    return "";
  }

  function distance(a, b) {
    if (!a || !b) return Infinity;
    const r = 6371000, rad = (d) => d * Math.PI / 180;
    const dLat = rad(b.lat - a.lat), dLon = rad(b.lon - a.lon);
    const h = Math.sin(dLat/2) ** 2 + Math.cos(rad(a.lat)) * Math.cos(rad(b.lat)) * Math.sin(dLon/2) ** 2;
    return 2 * r * Math.asin(Math.min(1, Math.sqrt(h)));
  }

  function recordDistance(record) { return distance(state.user, point(record)); }

  function cacheRead(key) {
    try {
      const raw = sessionStorage.getItem(key); if (!raw) return null;
      const parsed = JSON.parse(raw);
      if (!parsed || Date.now() - Number(parsed.savedAt || 0) > CACHE_TTL) return null;
      return parsed.data;
    } catch (_) { return null; }
  }
  function cacheWrite(key, data) { try { sessionStorage.setItem(key, JSON.stringify({ savedAt: Date.now(), data })); } catch (_) {} }

  async function fetchAll(datasetId) {
    const ck = `vb-parks-fix:${datasetId}`;
    const cached = cacheRead(ck); if (Array.isArray(cached)) return cached;
    const out = []; let offset = 0; let total = Infinity;
    while (offset < total && offset < 10000) {
      const params = new URLSearchParams({ limit: "100", offset: String(offset) });
      const res = await fetch(`${API_ROOT}/${encodeURIComponent(datasetId)}/records?${params}`, { headers: { Accept: "application/json" } });
      if (!res.ok) throw new Error(`Open Data Comune di Bologna non disponibili (${res.status}).`);
      const payload = await res.json();
      const rows = Array.isArray(payload?.results) ? payload.results : [];
      total = Number(payload?.total_count ?? rows.length);
      out.push(...rows);
      if (!rows.length) break;
      offset += rows.length;
    }
    cacheWrite(ck, out); return out;
  }

  function injectStyle() {
    if ($(STYLE_ID)) return;
    const style = document.createElement("style");
    style.id = STYLE_ID;
    style.textContent = `
      #${PANEL_ID},#${SHEET_ID}{display:none}
      @media ${MOBILE}{
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-search-form,
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-status,
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-map-card,
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-results,
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-load-more,
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-parchi-quartieri,
        #${PAGE_ID}.vb-parks-fix-active #verde-bologna-parchi-list{display:none!important}
        #${PAGE_ID}.vb-parks-fix-active #${PANEL_ID}{display:grid;gap:10px}
        .vbpf-search{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:8px}
        .vbpf-search input{min-width:0;min-height:48px;padding:10px 12px;border:1px solid #b8c8d8;border-radius:11px;background:#fff;font:inherit;font-size:1rem}
        .vbpf-search button{min-width:48px;min-height:48px}
        .vbpf-chips{display:flex;gap:6px;overflow-x:auto;padding:1px 0 4px;scrollbar-width:none}.vbpf-chips::-webkit-scrollbar{display:none}
        .vbpf-chip{flex:0 0 auto;min-height:36px;padding:7px 10px;border:1px solid #c7d5e2;border-radius:999px;background:#f5f8fb;color:#244766;font-size:.72rem;font-weight:900}
        .vbpf-chip.active{background:#12623a;border-color:#12623a;color:#fff}
        .vbpf-toolbar{display:grid;grid-template-columns:1fr 1fr;gap:7px}
        .vbpf-toolbar button{min-height:42px;font-size:.75rem;font-weight:900}
        .vbpf-status{margin:0;padding:8px 10px;border-radius:10px;background:#edf4fb;color:#355777;font-size:.76rem;line-height:1.35}
        #${MAP_ID}{height:52vh;min-height:360px;border-radius:13px;overflow:hidden;background:#e9eef4}
        .vbpf-marker-wrap{background:transparent!important;border:0!important}.vbpf-marker{display:flex;align-items:center;justify-content:center;min-width:44px;height:30px;padding:0 8px;border:2px solid #fff;border-radius:16px;background:#12623a;color:#fff;box-shadow:0 2px 7px rgba(0,0,0,.38);font-size:.73rem;font-weight:900;white-space:nowrap}.vbpf-marker.fallback{background:#455d76}
        .vbpf-list{display:grid;gap:7px}.vbpf-row{display:grid;grid-template-columns:minmax(78px,.32fr) minmax(0,1fr) auto;gap:9px;align-items:center;width:100%;padding:10px;border:1px solid #d7e2ec;border-radius:12px;background:#fff;text-align:left;box-shadow:0 3px 10px rgba(26,55,91,.06)}
        .vbpf-code{display:flex;align-items:center;justify-content:center;min-height:34px;padding:6px 8px;border-radius:10px;background:#12623a;color:#fff;font-size:.78rem;font-weight:900}.vbpf-name{min-width:0;color:#10264a;font-size:.9rem;font-weight:850;line-height:1.25}.vbpf-arrow{font-size:1.2rem;color:#55708f}
        #${SHEET_ID}.open{display:block;position:fixed;inset:0;z-index:13100;overflow:auto;background:#f1f6fb;padding:max(10px,env(safe-area-inset-top)) 10px max(16px,env(safe-area-inset-bottom))}
        .vbpf-sheet-head{position:sticky;top:0;z-index:2;display:flex;gap:8px;align-items:center;margin:-10px -10px 10px;padding:max(10px,env(safe-area-inset-top)) 10px 10px;background:rgba(255,255,255,.98);border-bottom:1px solid #d9e3ef}.vbpf-sheet-head h2{margin:0;min-width:0;flex:1;font-size:1.05rem;color:#10264a;white-space:nowrap;overflow:hidden;text-overflow:ellipsis}
        .vbpf-summary,.vbpf-field{padding:11px;border:1px solid #dce6ef;border-radius:12px;background:#fff}.vbpf-summary{display:grid;grid-template-columns:auto 1fr;gap:7px;margin-bottom:9px}.vbpf-summary strong{color:#12623a}.vbpf-summary span{color:#526b84}.vbpf-actions{display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:9px}.vbpf-actions .btn{min-height:44px}.vbpf-fields{display:grid;gap:7px}.vbpf-field span{display:block;margin-bottom:3px;color:#617990;font-size:.68rem;font-weight:900;text-transform:uppercase}.vbpf-field strong,.vbpf-field pre{display:block;margin:0;color:#203e59;font:inherit;font-size:.82rem;white-space:pre-wrap;overflow-wrap:anywhere}
      }`;
    document.head.appendChild(style);
  }

  function ensureUi() {
    const browser = $("verde-bologna-browser");
    if (!browser) return false;
    if (!$(PANEL_ID)) {
      const panel = document.createElement("section");
      panel.id = PANEL_ID;
      panel.innerHTML = `
        <div class="vbpf-search"><input id="vbpf-query" type="search" autocomplete="off" placeholder="Cerca codice o nome parco…"><button id="vbpf-clear" class="btn" type="button">✕</button></div>
        <div id="vbpf-chips" class="vbpf-chips"></div>
        <div class="vbpf-toolbar"><button id="vbpf-location" class="btn" type="button">⌖ LA MIA POSIZIONE</button><button id="vbpf-fit" class="btn" type="button">◎ MOSTRA RISULTATI</button></div>
        <p id="vbpf-status" class="vbpf-status">Seleziona Parchi e giardini.</p>
        <div id="${MAP_ID}"></div>
        <section id="vbpf-list" class="vbpf-list"></section>`;
      browser.prepend(panel);
    }
    if (!$(SHEET_ID)) {
      const sheet = document.createElement("section"); sheet.id = SHEET_ID; sheet.setAttribute("aria-hidden", "true"); document.body.appendChild(sheet);
    }
    return true;
  }

  function renderChips() {
    const node = $("vbpf-chips"); if (!node) return;
    const counts = new Map(); state.parks.forEach((r) => { const q = r.__quarter || ""; if (q) counts.set(q, (counts.get(q) || 0) + 1); });
    node.innerHTML = ["", ...QUARTERS].map((q) => `<button class="vbpf-chip${state.activeQuarter === q ? " active" : ""}" type="button" data-q="${esc(q)}">${esc(q || "TUTTI")}${q && counts.has(q) ? ` · ${counts.get(q)}` : ""}</button>`).join("");
    node.querySelectorAll("[data-q]").forEach((b) => b.addEventListener("click", () => { state.activeQuarter = b.getAttribute("data-q") || ""; renderChips(); applyFilter(); }));
  }

  function formatValue(value) {
    if (value === null || value === undefined || value === "") return "—";
    if (typeof value === "boolean") return value ? "Sì" : "No";
    if (typeof value === "object") { try { return JSON.stringify(value, null, 2); } catch (_) { return String(value); } }
    return String(value);
  }

  function queryMatches(record, query) {
    if (!query) return true;
    const nq = norm(query), digits = String(query).replace(/[^0-9]/g, "");
    const codes = [primaryCode(record).value, codvia(record), codOgg(record), idTopon(record)].map((v) => String(v || "").replace(/[^0-9a-z]/gi, "").toLowerCase()).filter(Boolean);
    if (digits && /^\d+$/.test(digits) && codes.some((c) => c.startsWith(digits))) return true;
    const name = norm(parkName(record));
    return name.includes(nq) || codes.some((c) => c.includes(nq.replace(/[^0-9a-z]/g, "")));
  }

  function applyFilter() {
    if (!active() || !state.loaded) return;
    const query = $("vbpf-query")?.value || "";
    let rows = state.parks.filter((r) => (!state.activeQuarter || r.__quarter === state.activeQuarter) && queryMatches(r, query));
    rows.sort((a, b) => {
      if (state.user) {
        const d = recordDistance(a) - recordDistance(b); if (Number.isFinite(d) && Math.abs(d) > .01) return d;
      }
      return parkName(a).localeCompare(parkName(b), "it", { sensitivity: "base", numeric: true });
    });
    state.filtered = rows;
    renderList(); renderMarkers(true);
    const status = $("vbpf-status");
    if (status) status.textContent = `${rows.length} risultati${state.user ? " · dal più vicino al più lontano" : " · posizione non disponibile"}. Cerca per codice o nome.`;
  }

  function renderList() {
    const node = $("vbpf-list"); if (!node) return;
    if (state.loading) { node.innerHTML = `<p class="vbpf-status">Carico parchi e quartieri ufficiali…</p>`; return; }
    if (!state.filtered.length) { node.innerHTML = `<p class="vbpf-status">Nessun risultato con questi filtri.</p>`; return; }
    node.innerHTML = state.filtered.map((r, i) => { const c = primaryCode(r); return `<button class="vbpf-row" type="button" data-i="${i}"><span class="vbpf-code">${esc(c.value)}</span><span class="vbpf-name">${esc(parkName(r))}</span><span class="vbpf-arrow">›</span></button>`; }).join("");
    node.querySelectorAll("[data-i]").forEach((b) => b.addEventListener("click", () => { const r = state.filtered[Number(b.getAttribute("data-i"))]; if (r) openSheet(r); }));
  }

  function ensureMap() {
    if (state.map || !active() || !window.L || !$(MAP_ID)) return;
    state.map = window.L.map(MAP_ID, { zoomControl: true, zoomAnimation: false, fadeAnimation: false, markerZoomAnimation: false }).setView([44.4949, 11.3426], 12);
    window.L.tileLayer("https://{s}.tile.openstreetmap.org/{z}/{x}/{y}.png", { maxZoom: 20, maxNativeZoom: 19, attribution: "&copy; OpenStreetMap contributors" }).addTo(state.map);
    state.markerLayer = window.L.layerGroup().addTo(state.map);
    state.boundaryLayer = window.L.layerGroup().addTo(state.map);
    setTimeout(() => state.map?.invalidateSize({ pan:false }), 80);
  }

  function renderMarkers(fit) {
    ensureMap(); if (!state.map || !state.markerLayer || !window.L) return;
    state.markerLayer.clearLayers();
    const bounds = window.L.latLngBounds([]);
    state.filtered.forEach((r) => {
      const p = point(r); if (!p) return;
      const c = primaryCode(r);
      const marker = window.L.marker([p.lat, p.lon], { icon: window.L.divIcon({ className:"vbpf-marker-wrap", html:`<span class="vbpf-marker${c.type === "CODVIA" ? "" : " fallback"}">${esc(c.value)}</span>`, iconSize:null, iconAnchor:[22,15] }), title:`${c.value} · ${parkName(r)}`, keyboard:true, riseOnHover:true }).addTo(state.markerLayer);
      marker.on("click", () => openSheet(r)); bounds.extend([p.lat,p.lon]);
    });
    if (fit && bounds.isValid()) state.map.fitBounds(bounds.pad(.08), { animate:false, maxZoom: state.filtered.length <= 3 ? 17 : 14 });
    setTimeout(() => state.map?.invalidateSize({ pan:false }), 40);
  }

  function requestLocation(force=false) {
    if (!active() || !navigator.geolocation || (state.locationAsked && !force)) return;
    state.locationAsked = true;
    const status = $("vbpf-status"); if (status) status.textContent = "Rilevo la tua posizione per ordinare i parchi…";
    navigator.geolocation.getCurrentPosition((pos) => {
      const lat = Number(pos.coords.latitude), lon = Number(pos.coords.longitude);
      if (Number.isFinite(lat) && Number.isFinite(lon)) { state.user = {lat,lon}; applyFilter(); }
    }, () => { if (status) status.textContent = "Posizione non disponibile: elenco ordinato per nome."; }, { enableHighAccuracy:true, timeout:12000, maximumAge:60000 });
  }

  function meaningfulWords(name) {
    const stop = new Set(["parco","giardino","giardini","area","verde","del","della","delle","dei","degli","di","da","il","lo","la","i","gli","le"]);
    return norm(name).replace(/[^a-z0-9 ]+/g," ").split(/\s+/).filter((w) => w.length > 2 && !stop.has(w));
  }

  async function showBoundaries(record) {
    ensureMap(); if (!state.map || !state.boundaryLayer || !window.L) return;
    const name = parkName(record), words = meaningfulWords(name), cacheKey = words.join(" ");
    const status = $("vbpf-status"); if (status) status.textContent = `Cerco i confini gestionali di ${name}…`;
    let rows = state.boundaryCache.get(cacheKey);
    if (!rows) {
      try {
        const term = words.slice(-3).join(" ") || name;
        const safe = String(term).replace(/\\/g,"\\\\").replace(/"/g,'\\"');
        const params = new URLSearchParams({ limit:"100", where:`search(\"${safe}\")` });
        const res = await fetch(`${API_ROOT}/un_gest/records?${params}`, { headers:{Accept:"application/json"} });
        const payload = res.ok ? await res.json() : {results:[]};
        rows = (Array.isArray(payload.results) ? payload.results : []).filter((r) => { const h = norm([r.nome,r.nome_ug,r.ubicazione].filter(Boolean).join(" ")); return words.length ? words.every((w) => h.includes(w)) : false; });
      } catch (_) { rows = []; }
      state.boundaryCache.set(cacheKey, rows);
    }
    state.boundaryLayer.clearLayers(); const bounds = window.L.latLngBounds([]);
    rows.forEach((r) => { const g = geometry(r); if (!g) return; const layer = window.L.geoJSON({type:"Feature",geometry:g,properties:{}},{style:{color:"#0b6b3a",weight:4,fillColor:"#45a96a",fillOpacity:.16}}).addTo(state.boundaryLayer); const b=layer.getBounds?.(); if(b?.isValid?.()) bounds.extend(b); });
    if (bounds.isValid()) { state.map.fitBounds(bounds.pad(.08), {animate:false,maxZoom:18}); if(status) status.textContent = `${rows.length} unità gestionali evidenziate per ${name}.`; }
    else if (status) status.textContent = `Nessun confine gestionale abbinato automaticamente a ${name}.`;
  }

  function closeSheet() { const s=$(SHEET_ID); s?.classList.remove("open"); s?.setAttribute("aria-hidden","true"); state.selected=null; }

  function openSheet(record) {
    const sheet = $(SHEET_ID); if (!sheet) return;
    state.selected = record;
    const c = primaryCode(record), p = point(record), dist = recordDistance(record);
    const nav = p ? `https://www.google.com/maps/dir/?api=1&destination=${encodeURIComponent(`${p.lat},${p.lon}`)}` : "";
    const fields = Object.entries(record).filter(([k]) => !String(k).startsWith("__"));
    sheet.innerHTML = `<header class="vbpf-sheet-head"><button class="btn" type="button" data-close>← INDIETRO</button><h2>${esc(parkName(record))}</h2></header><section class="vbpf-summary"><strong>CODICE</strong><span>${esc(c.value)} (${esc(c.type)})</span><strong>CODVIA</strong><span>${esc(codvia(record)||"0")}</span><strong>COD_OGG</strong><span>${esc(codOgg(record)||"—")}</span><strong>QUARTIERE</strong><span>${esc(record.__quarter||"Non determinato")}</span>${Number.isFinite(dist)?`<strong>DISTANZA</strong><span>${dist<1000?`${Math.round(dist)} m`:`${(dist/1000).toFixed(1)} km`}</span>`:""}</section><div class="vbpf-actions"><button class="btn" type="button" data-map>CONFINI</button>${p?`<a class="btn btn-primary" href="${esc(nav)}" target="_blank" rel="noopener">NAVIGA</a>`:`<button class="btn btn-primary" disabled>NAVIGA</button>`}</div><section class="vbpf-fields">${fields.map(([k,v])=>`<article class="vbpf-field"><span>${esc(k)}</span>${typeof v==="object"&&v!==null?`<pre>${esc(formatValue(v))}</pre>`:`<strong>${esc(formatValue(v))}</strong>`}</article>`).join("")}</section>`;
    sheet.classList.add("open"); sheet.setAttribute("aria-hidden","false");
    sheet.querySelector("[data-close]")?.addEventListener("click", closeSheet);
    sheet.querySelector("[data-map]")?.addEventListener("click", () => { closeSheet(); void showBoundaries(record); $(MAP_ID)?.scrollIntoView({behavior:"smooth",block:"center"}); });
  }

  async function loadData() {
    if (state.loaded || state.loading) return;
    state.loading = true; renderList();
    try {
      const [parks, quarters] = await Promise.all([fetchAll(PARKS_ID), fetchAll(QUARTERS_ID)]);
      state.quarters = quarters;
      state.parks = parks.map((r) => ({...r,__quarter:assignQuarter(r)}));
      state.loaded = true; state.loading = false; renderChips(); applyFilter(); requestLocation();
    } catch (e) { state.loading=false; const s=$("vbpf-status"); if(s)s.textContent=e?.message||"Impossibile caricare i parchi."; renderList(); }
  }

  function activate() {
    if (!active() || !ensureUi()) return;
    $(PAGE_ID)?.classList.add("vb-parks-fix-active");
    ensureMap(); void loadData(); renderChips(); if(state.loaded)applyFilter(); requestLocation();
  }
  function deactivate() { $(PAGE_ID)?.classList.remove("vb-parks-fix-active"); closeSheet(); state.boundaryLayer?.clearLayers?.(); }

  function wire() {
    if (state.installed) return true;
    const page=$(PAGE_ID), select=$(CATEGORY_ID); if(!page||!select||!ensureUi())return false;
    state.installed=true;
    $("vbpf-query")?.addEventListener("input", applyFilter);
    $("vbpf-clear")?.addEventListener("click",()=>{ const q=$("vbpf-query"); if(q)q.value=""; state.activeQuarter=""; renderChips(); applyFilter(); q?.focus(); });
    $("vbpf-location")?.addEventListener("click",()=>requestLocation(true));
    $("vbpf-fit")?.addEventListener("click",()=>renderMarkers(true));
    select.addEventListener("change",()=>setTimeout(()=>active()?activate():deactivate(),60));
    const observer=new MutationObserver(()=>active()?activate():deactivate()); observer.observe(page,{attributes:true,attributeFilter:["class","aria-hidden"]});
    document.addEventListener("keydown",(e)=>{if(e.key==="Escape"&&$(SHEET_ID)?.classList.contains("open")){e.stopPropagation();closeSheet();}},true);
    if(active())activate(); return true;
  }

  function install() {
    injectStyle(); ensureUi();
    if(wire())return;
    let n=0; const timer=setInterval(()=>{n+=1;if(wire()||n>120)clearInterval(timer);},100);
  }

  if(document.readyState==="loading")document.addEventListener("DOMContentLoaded",install,{once:true}); else install();
})();
