(() => {
  "use strict";

  if (window.HeraLavoriOccasionaliGooglePlaces?.installed) return;

  const API_SCRIPT_ID = "hera-google-maps-places-api";
  const HOST_ID = "lavoro-occasionale-google-places-host";
  const STATUS_ID = "lavoro-occasionale-google-places-status";
  const COMMESSA_ID = "lavori-occasionali";
  let initPromise = null;
  let observer = null;

  function isOccasionalSelected() {
    return document.getElementById("squadra-commessa")?.value === COMMESSA_ID;
  }

  function setEditorValue(id, value) {
    const editor = document.getElementById(id);
    if (!editor) return;
    const text = String(value || "").trim();
    if ("value" in editor && editor.tagName === "INPUT") editor.value = text;
    else editor.textContent = text;
    editor.dispatchEvent(new Event("input", { bubbles: true }));
    editor.dispatchEvent(new Event("change", { bubbles: true }));
  }

  function setStatus(message, type = "info") {
    const status = document.getElementById(STATUS_ID);
    if (!status) return;
    status.textContent = message || "";
    status.dataset.type = type;
  }

  function getComponent(components, type, shortName = false) {
    const component = (components || []).find((item) => Array.isArray(item.types) && item.types.includes(type));
    if (!component) return "";
    return String(shortName ? (component.shortText || component.short_name || "") : (component.longText || component.long_name || "")).trim();
  }

  function buildStreetAddress(components, formattedAddress) {
    const route = getComponent(components, "route");
    const number = getComponent(components, "street_number");
    if (route) return `${route}${number ? ` ${number}` : ""}`.trim();
    return String(formattedAddress || "").split(",")[0].trim();
  }

  function buildComune(components) {
    return getComponent(components, "locality")
      || getComponent(components, "administrative_area_level_3")
      || getComponent(components, "postal_town")
      || getComponent(components, "administrative_area_level_2");
  }

  function fillFromPlace(place) {
    const lat = typeof place.location?.lat === "function" ? place.location.lat() : Number(place.location?.lat);
    const lng = typeof place.location?.lng === "function" ? place.location.lng() : Number(place.location?.lng);
    const components = place.addressComponents || place.address_components || [];
    const name = String(place.displayName || place.name || "").trim();
    const formattedAddress = String(place.formattedAddress || place.formatted_address || "").trim();
    const comune = buildComune(components);
    const indirizzo = buildStreetAddress(components, formattedAddress);

    if (name) setEditorValue("lavoro-occasionale-nome", name);
    if (comune) setEditorValue("lavoro-occasionale-comune", comune);
    if (indirizzo || formattedAddress) setEditorValue("lavoro-occasionale-indirizzo", indirizzo || formattedAddress);
    if (Number.isFinite(lat) && Number.isFinite(lng)) {
      setEditorValue("lavoro-occasionale-coordinate", `${lat.toFixed(6)}, ${lng.toFixed(6)}`);
    }

    const extra = [];
    const provincia = getComponent(components, "administrative_area_level_2", true);
    const cap = getComponent(components, "postal_code");
    if (provincia) extra.push(provincia);
    if (cap) extra.push(cap);
    setStatus(`✓ Selezionato da Google Maps${extra.length ? ` · ${extra.join(" · ")}` : ""}`, "success");
  }

  function ensureGoogleLoader() {
    if (window.google?.maps?.importLibrary) return Promise.resolve();
    if (initPromise) return initPromise;

    initPromise = new Promise((resolve, reject) => {
      const apiKey = String(window.firebaseConfig?.apiKey || "").trim();
      if (!apiKey) {
        reject(new Error("Chiave Google Maps non configurata."));
        return;
      }

      const existing = document.getElementById(API_SCRIPT_ID);
      if (existing) {
        const started = Date.now();
        const timer = window.setInterval(() => {
          if (window.google?.maps?.importLibrary) {
            window.clearInterval(timer);
            resolve();
          } else if (Date.now() - started > 12000) {
            window.clearInterval(timer);
            reject(new Error("Google Maps non disponibile."));
          }
        }, 100);
        return;
      }

      const script = document.createElement("script");
      script.id = API_SCRIPT_ID;
      script.async = true;
      script.defer = true;
      script.src = `https://maps.googleapis.com/maps/api/js?key=${encodeURIComponent(apiKey)}&v=weekly&language=it&region=IT&loading=async`;
      script.onload = () => resolve();
      script.onerror = () => reject(new Error("Impossibile caricare Google Maps."));
      document.head.appendChild(script);
    });

    return initPromise;
  }

  async function installAutocomplete(host) {
    if (!host || host.dataset.ready === "1") return;
    host.dataset.ready = "loading";
    setStatus("Caricamento ricerca Google Maps…");

    try {
      await ensureGoogleLoader();
      const { PlaceAutocompleteElement } = await google.maps.importLibrary("places");
      const autocomplete = new PlaceAutocompleteElement();
      autocomplete.id = "lavoro-occasionale-google-places";
      autocomplete.placeholder = "Cerca strada, azienda, struttura, luogo…";
      autocomplete.setAttribute("aria-label", "Cerca cantiere su Google Maps");
      autocomplete.includedRegionCodes = ["it"];
      autocomplete.style.width = "100%";
      autocomplete.style.display = "block";
      autocomplete.style.minHeight = "48px";

      autocomplete.addEventListener("gmp-select", async ({ placePrediction }) => {
        try {
          setStatus("Recupero dati del luogo…");
          const place = placePrediction.toPlace();
          await place.fetchFields({
            fields: ["displayName", "formattedAddress", "location", "addressComponents"]
          });
          fillFromPlace(place);
        } catch (error) {
          console.error("Google Places - selezione luogo", error);
          setStatus("Non riesco a leggere i dati del luogo selezionato.", "error");
        }
      });

      host.replaceChildren(autocomplete);
      host.dataset.ready = "1";
      setStatus("Cerca direttamente su Google Maps: strade, aziende, strutture e luoghi.");
    } catch (error) {
      console.error("Google Places - inizializzazione", error);
      host.dataset.ready = "error";
      setStatus("Ricerca Google non disponibile. Verifica che Maps JavaScript API e Places API (New) siano abilitate per la chiave del progetto.", "error");
    }
  }

  function installUi() {
    const field = document.getElementById("lavoro-occasionale-field");
    const nameEditor = document.getElementById("lavoro-occasionale-nome");
    if (!field || !nameEditor || document.getElementById(HOST_ID)) return;

    const wrapper = document.createElement("div");
    wrapper.className = "lavoro-occasionale-google-places-wrap";
    wrapper.innerHTML = `
      <span class="lavoro-occasionale-subtitle">🔎 Cerca cantiere su Google Maps</span>
      <div id="${HOST_ID}" class="lavoro-occasionale-google-places-host"></div>
      <small id="${STATUS_ID}" class="lavoro-occasionale-google-places-status">Cerca direttamente su Google Maps: strade, aziende, strutture e luoghi.</small>
    `;

    nameEditor.parentNode.insertBefore(wrapper, nameEditor);

    const style = document.createElement("style");
    style.dataset.heraOccasionalGooglePlaces = "1";
    style.textContent = `
      .lavoro-occasionale-google-places-wrap { margin: 0 0 14px; display: grid; gap: 7px; }
      .lavoro-occasionale-google-places-host { width: 100%; min-height: 48px; }
      .lavoro-occasionale-google-places-status { display: block; opacity: .78; line-height: 1.35; }
      .lavoro-occasionale-google-places-status[data-type="success"] { font-weight: 700; opacity: 1; }
      .lavoro-occasionale-google-places-status[data-type="error"] { font-weight: 700; opacity: 1; }
      gmp-place-autocomplete { width: 100%; box-sizing: border-box; }
    `;
    document.head.appendChild(style);

    if (isOccasionalSelected()) installAutocomplete(wrapper.querySelector(`#${HOST_ID}`));
  }

  function refresh() {
    installUi();
    const host = document.getElementById(HOST_ID);
    if (host && isOccasionalSelected() && host.dataset.ready !== "1" && host.dataset.ready !== "loading") {
      installAutocomplete(host);
    }
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
    version: "1.0.0",
    refresh,
    disconnect() {
      observer?.disconnect();
      observer = null;
    }
  };
})();
