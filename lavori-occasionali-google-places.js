(() => {
  "use strict";

  if (window.HeraLavoriOccasionaliGooglePlaces?.installed) return;

  const API_SCRIPT_ID = "hera-google-maps-places-api";
  const COMMESSA_ID = "lavori-occasionali";
  const OPEN_BUTTON_ID = "lavoro-occasionale-google-map-open";
  const MODAL_ID = "lavoro-occasionale-google-map-modal";
  const SEARCH_HOST_ID = "lavoro-occasionale-google-map-search";
  const MAP_ID = "lavoro-occasionale-google-map-canvas";
  const CARD_ID = "lavoro-occasionale-google-map-card";
  let initPromise = null;
  let observer = null;
  let map = null;
  let marker = null;
  let selectedPlace = null;

  function isOccasionalSelected() {
    return document.getElementById("squadra-commessa")?.value === COMMESSA_ID;
  }

  function setEditorValue(id, value) {
    const editor = document.getElementById(id);
    if (!editor) return;
    const text = String(value || "").trim();
    if (editor.tagName === "INPUT" || editor.tagName === "TEXTAREA") editor.value = text;
    else editor.textContent = text;
    editor.dispatchEvent(new Event("input", { bubbles: true }));
    editor.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function getComponent(components, type, shortName = false) {
    const component = (components || []).find((item) => Array.isArray(item.types) && item.types.includes(type));
    if (!component) return "";
    return String(shortName
      ? (component.shortText || component.short_name || "")
      : (component.longText || component.long_name || "")
    ).trim();
  }

  function normalizePlace(place) {
    const components = place.addressComponents || place.address_components || [];
    const formattedAddress = String(place.formattedAddress || place.formatted_address || "").trim();
    const route = getComponent(components, "route");
    const streetNumber = getComponent(components, "street_number");
    const comune = getComponent(components, "locality")
      || getComponent(components, "administrative_area_level_3")
      || getComponent(components, "postal_town")
      || getComponent(components, "administrative_area_level_2");
    const provincia = getComponent(components, "administrative_area_level_2", true);
    const cap = getComponent(components, "postal_code");
    const lat = typeof place.location?.lat === "function" ? place.location.lat() : Number(place.location?.lat);
    const lng = typeof place.location?.lng === "function" ? place.location.lng() : Number(place.location?.lng);
    return {
      name: String(place.displayName || place.name || formattedAddress || "Cantiere").trim(),
      address: route ? `${route}${streetNumber ? ` ${streetNumber}` : ""}`.trim() : formattedAddress,
      formattedAddress,
      comune,
      provincia,
      cap,
      lat,
      lng
    };
  }

  function ensureGoogleLoader() {
    if (window.google?.maps?.importLibrary) return Promise.resolve();
    if (initPromise) return initPromise;

    initPromise = new Promise((resolve, reject) => {
      const apiKey = String(window.firebaseConfig?.apiKey || "").trim();
      if (!apiKey) return reject(new Error("Chiave Google Maps non configurata."));

      const waitReady = () => {
        const started = Date.now();
        const timer = window.setInterval(() => {
          if (window.google?.maps?.importLibrary) {
            window.clearInterval(timer);
            resolve();
          } else if (Date.now() - started > 15000) {
            window.clearInterval(timer);
            reject(new Error("Google Maps non disponibile."));
          }
        }, 100);
      };

      const existing = document.getElementById(API_SCRIPT_ID);
      if (existing) return waitReady();

      const script = document.createElement("script");
      script.id = API_SCRIPT_ID;
      script.async = true;
      script.defer = true;
      script.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&v=weekly&language=it&region=IT&loading=async`;
      script.onload = waitReady;
      script.onerror = () => reject(new Error("Impossibile caricare Google Maps."));
      document.head.appendChild(script);
    });

    return initPromise;
  }

  function ensureStyles() {
    if (document.querySelector("style[data-hera-occasional-google-map]")) return;
    const style = document.createElement("style");
    style.dataset.heraOccasionalGoogleMap = "1";
    style.textContent = `
      .lavoro-occasionale-google-map-open { width:100%; min-height:56px; border:0; border-radius:16px; padding:14px 18px; font:inherit; font-weight:800; font-size:17px; background:#0b57d0; color:#fff; display:flex; align-items:center; justify-content:center; gap:9px; }
      .lavoro-occasionale-google-map-help { display:block; margin-top:7px; opacity:.72; line-height:1.35; }
      .lavoro-occasionale-google-map-modal { position:fixed; inset:0; z-index:2147483000; background:#fff; display:none; flex-direction:column; }
      .lavoro-occasionale-google-map-modal.open { display:flex; }
      .lavoro-occasionale-google-map-top { position:absolute; top:max(14px, env(safe-area-inset-top)); left:14px; right:14px; z-index:20; display:grid; grid-template-columns:48px 1fr; gap:9px; align-items:start; }
      .lavoro-occasionale-google-map-back { width:48px; height:48px; border:0; border-radius:24px; background:#fff; box-shadow:0 2px 12px rgba(0,0,0,.22); font-size:27px; line-height:48px; }
      .lavoro-occasionale-google-map-search { min-height:48px; border-radius:24px; background:#fff; box-shadow:0 2px 12px rgba(0,0,0,.22); overflow:visible; }
      .lavoro-occasionale-google-map-search gmp-place-autocomplete { width:100%; min-height:48px; border-radius:24px; }
      .lavoro-occasionale-google-map-canvas { position:absolute; inset:0; }
      .lavoro-occasionale-google-map-card { position:absolute; left:12px; right:12px; bottom:max(12px, env(safe-area-inset-bottom)); z-index:20; background:#fff; border-radius:20px; box-shadow:0 4px 20px rgba(0,0,0,.28); padding:15px; display:none; }
      .lavoro-occasionale-google-map-card.show { display:block; }
      .lavoro-occasionale-google-map-card-title { font-weight:850; font-size:18px; margin-bottom:5px; }
      .lavoro-occasionale-google-map-card-address { font-size:14px; opacity:.72; line-height:1.35; margin-bottom:12px; }
      .lavoro-occasionale-google-map-use { width:100%; min-height:50px; border:0; border-radius:14px; font:inherit; font-weight:850; background:#0b57d0; color:#fff; }
      .lavoro-occasionale-google-map-status { position:absolute; top:82px; left:50%; transform:translateX(-50%); z-index:19; background:rgba(255,255,255,.94); border-radius:14px; padding:8px 12px; box-shadow:0 2px 10px rgba(0,0,0,.16); font-size:13px; max-width:82%; text-align:center; }
    `;
    document.head.appendChild(style);
  }

  function createModal() {
    let modal = document.getElementById(MODAL_ID);
    if (modal) return modal;
    modal = document.createElement("div");
    modal.id = MODAL_ID;
    modal.className = "lavoro-occasionale-google-map-modal";
    modal.innerHTML = `
      <div id="${MAP_ID}" class="lavoro-occasionale-google-map-canvas"></div>
      <div class="lavoro-occasionale-google-map-top">
        <button type="button" class="lavoro-occasionale-google-map-back" aria-label="Chiudi mappa">‹</button>
        <div id="${SEARCH_HOST_ID}" class="lavoro-occasionale-google-map-search"></div>
      </div>
      <div class="lavoro-occasionale-google-map-status">Caricamento Google Maps…</div>
      <div id="${CARD_ID}" class="lavoro-occasionale-google-map-card">
        <div class="lavoro-occasionale-google-map-card-title"></div>
        <div class="lavoro-occasionale-google-map-card-address"></div>
        <button type="button" class="lavoro-occasionale-google-map-use">USA QUESTO CANTIERE</button>
      </div>
    `;
    document.body.appendChild(modal);

    modal.querySelector(".lavoro-occasionale-google-map-back")?.addEventListener("click", closeMap);
    modal.querySelector(".lavoro-occasionale-google-map-use")?.addEventListener("click", useSelectedPlace);
    return modal;
  }

  function closeMap() {
    const modal = document.getElementById(MODAL_ID);
    modal?.classList.remove("open");
    document.documentElement.style.overflow = "";
    document.body.style.overflow = "";
  }

  function updateSelectionCard(data) {
    const card = document.getElementById(CARD_ID);
    if (!card) return;
    card.querySelector(".lavoro-occasionale-google-map-card-title").textContent = data?.name || "";
    card.querySelector(".lavoro-occasionale-google-map-card-address").textContent = data?.formattedAddress || data?.address || "";
    card.classList.toggle("show", Boolean(data));
  }

  function useSelectedPlace() {
    if (!selectedPlace) return;
    setEditorValue("lavoro-occasionale-nome", selectedPlace.name);
    setEditorValue("lavoro-occasionale-comune", selectedPlace.comune);
    setEditorValue("lavoro-occasionale-indirizzo", selectedPlace.address || selectedPlace.formattedAddress);
    if (Number.isFinite(selectedPlace.lat) && Number.isFinite(selectedPlace.lng)) {
      setEditorValue("lavoro-occasionale-coordinate", `${selectedPlace.lat.toFixed(6)}, ${selectedPlace.lng.toFixed(6)}`);
    }
    closeMap();
  }

  async function setPlace(place) {
    selectedPlace = normalizePlace(place);
    if (!Number.isFinite(selectedPlace.lat) || !Number.isFinite(selectedPlace.lng)) return;
    const position = { lat: selectedPlace.lat, lng: selectedPlace.lng };
    if (!marker) marker = new google.maps.Marker({ map, position });
    else {
      marker.setMap(map);
      marker.setPosition(position);
    }
    map.panTo(position);
    map.setZoom(17);
    updateSelectionCard(selectedPlace);
  }

  function centerOnUser() {
    if (!navigator.geolocation || !map) return;
    navigator.geolocation.getCurrentPosition(
      ({ coords }) => map.setCenter({ lat: coords.latitude, lng: coords.longitude }),
      () => {},
      { enableHighAccuracy: true, timeout: 5000, maximumAge: 120000 }
    );
  }

  async function openMap() {
    if (!isOccasionalSelected()) return;
    ensureStyles();
    const modal = createModal();
    modal.classList.add("open");
    document.documentElement.style.overflow = "hidden";
    document.body.style.overflow = "hidden";

    const status = modal.querySelector(".lavoro-occasionale-google-map-status");
    if (status) status.textContent = "Caricamento Google Maps…";

    try {
      await ensureGoogleLoader();
      await Promise.all([google.maps.importLibrary("maps"), google.maps.importLibrary("places")]);

      if (!map) {
        map = new google.maps.Map(document.getElementById(MAP_ID), {
          center: { lat: 44.4949, lng: 11.3426 },
          zoom: 13,
          mapTypeControl: true,
          streetViewControl: false,
          fullscreenControl: false,
          gestureHandling: "greedy"
        });
        centerOnUser();
      } else {
        google.maps.event.trigger(map, "resize");
      }

      const host = document.getElementById(SEARCH_HOST_ID);
      if (host && !host.querySelector("gmp-place-autocomplete")) {
        const { PlaceAutocompleteElement } = await google.maps.importLibrary("places");
        const autocomplete = new PlaceAutocompleteElement();
        autocomplete.placeholder = "Cerca qui";
        autocomplete.includedRegionCodes = ["it"];
        autocomplete.setAttribute("aria-label", "Cerca luogo su Google Maps");
        autocomplete.addEventListener("gmp-select", async ({ placePrediction }) => {
          const place = placePrediction.toPlace();
          await place.fetchFields({ fields: ["displayName", "formattedAddress", "location", "addressComponents"] });
          await setPlace(place);
        });
        host.replaceChildren(autocomplete);
      }

      if (status) status.textContent = "Cerca una strada, azienda o struttura e seleziona il risultato.";
      window.setTimeout(() => { if (status) status.style.display = "none"; }, 2200);
    } catch (error) {
      console.error("Google map lavori occasionali", error);
      if (status) {
        status.style.display = "block";
        status.textContent = "Google Maps non disponibile. Verifica Maps JavaScript API e Places API (New).";
      }
    }
  }

  function installLauncher() {
    const field = document.getElementById("lavoro-occasionale-field");
    const nameEditor = document.getElementById("lavoro-occasionale-nome");
    if (!field || !nameEditor || document.getElementById(OPEN_BUTTON_ID)) return;

    ensureStyles();
    const wrapper = document.createElement("div");
    wrapper.className = "lavoro-occasionale-google-map-launcher";
    wrapper.innerHTML = `
      <button type="button" id="${OPEN_BUTTON_ID}" class="lavoro-occasionale-google-map-open">🗺️ CERCA CANTIERE SU MAPPA GOOGLE</button>
      <small class="lavoro-occasionale-google-map-help">Cerca su Google Maps e scegli il luogo: nome, indirizzo, comune e coordinate verranno compilati automaticamente.</small>
    `;
    nameEditor.parentNode.insertBefore(wrapper, nameEditor);
    wrapper.querySelector(`#${OPEN_BUTTON_ID}`)?.addEventListener("click", openMap);
  }

  function removeOldInlineSearch() {
    document.querySelector(".lavoro-occasionale-google-places-wrap")?.remove();
  }

  function refresh() {
    removeOldInlineSearch();
    installLauncher();
  }

  document.addEventListener("change", (event) => {
    if (event.target?.id === "squadra-commessa") refresh();
  });

  observer = new MutationObserver(refresh);
  observer.observe(document.documentElement, { childList: true, subtree: true });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", refresh, { once: true });
  else refresh();

  window.HeraLavoriOccasionaliGooglePlaces = {
    installed: true,
    version: "2.0.0",
    refresh,
    openMap,
    closeMap,
    disconnect() {
      observer?.disconnect();
      observer = null;
    }
  };
})();
