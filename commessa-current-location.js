/* Pulsante "Usa la mia posizione" della Gestione commessa.
 * Implementazione isolata: non cambia viste, FATTO, WHAZZUP o Firestore.
 */
(function rebuildCommessaCurrentLocation(root) {
  "use strict";

  if (root.__commessaCurrentLocationRebuilt) return;
  root.__commessaCurrentLocationRebuilt = true;

  let requestSequence = 0;
  let reverseController = null;

  const BUTTON_ID = "commessa-mobile-current-location";
  const FORM_ID = "commessa-mobile-plant-form";
  const DEFAULT_LABEL = "📍 Usa la mia posizione";

  function getStatus(form) {
    return form && form.querySelector("#commessa-mobile-geocode-status");
  }

  function setStatus(form, message) {
    const status = getStatus(form);
    if (status) status.textContent = message;
  }

  function setButton(button, label, disabled) {
    if (!button) return;
    button.disabled = Boolean(disabled);
    button.textContent = label;
    button.setAttribute("aria-busy", disabled ? "true" : "false");
  }

  function field(form, name) {
    return form && form.querySelector(`[data-v2-field="${name}"]`);
  }

  function validCoordinate(value, min, max) {
    const number = Number(String(value ?? "").trim().replace(",", "."));
    return Number.isFinite(number) && number >= min && number <= max && number !== 0
      ? number
      : null;
  }

  function isWritableAddressField(input) {
    if (!input) return false;
    const value = String(input.value || "").trim();
    const generated = String(input.dataset.currentLocationGeneratedValue || "").trim();
    return !value || (generated && value === generated);
  }

  function writeGeneratedAddress(input, value) {
    const clean = String(value || "").trim();
    if (!clean || !isWritableAddressField(input)) return false;
    input.value = clean;
    input.dataset.currentLocationGeneratedValue = clean;
    input.dataset.mobileGeocodeValue = clean;
    input.dataset.mobileManualValue = "false";
    return true;
  }

  function geolocationErrorMessage(error) {
    if (error && error.code === 1) {
      return "Posizione non autorizzata. Consenti l’accesso alla posizione e riprova.";
    }
    if (error && error.code === 2) {
      return "Posizione non disponibile. Controlla che il GPS sia attivo e riprova.";
    }
    if (error && error.code === 3) {
      return "Il rilevamento GPS ha impiegato troppo tempo. Riprova.";
    }
    return "Non è stato possibile rilevare la posizione. Controlla il GPS e riprova.";
  }

  function extractAddress(data) {
    const address = data && data.address ? data.address : {};
    const comune = address.city
      || address.town
      || address.village
      || address.municipality
      || address.city_district
      || address.county
      || "";
    const road = address.road
      || address.pedestrian
      || address.residential
      || address.footway
      || address.path
      || address.neighbourhood
      || "";
    const number = address.house_number || "";
    const via = [road, number].filter(Boolean).join(" ").trim();
    return { comune: String(comune).trim(), via };
  }

  async function reverseGeocode(form, latitude, longitude, requestId) {
    if (typeof fetch !== "function") {
      setStatus(form, "Coordinate GPS acquisite. Inserisci Comune e Via manualmente.");
      return;
    }

    if (reverseController) reverseController.abort();
    reverseController = typeof AbortController === "function" ? new AbortController() : null;

    const url = new URL("https://nominatim.openstreetmap.org/reverse");
    url.searchParams.set("format", "jsonv2");
    url.searchParams.set("lat", String(latitude));
    url.searchParams.set("lon", String(longitude));
    url.searchParams.set("addressdetails", "1");
    url.searchParams.set("accept-language", "it");

    try {
      const response = await fetch(url.toString(), {
        method: "GET",
        headers: { Accept: "application/json" },
        signal: reverseController ? reverseController.signal : undefined
      });
      if (!response.ok) throw new Error(`HTTP ${response.status}`);
      const data = await response.json();
      if (requestId !== requestSequence || !form.isConnected) return;

      const result = extractAddress(data);
      const comuneAdded = writeGeneratedAddress(field(form, "comune"), result.comune);
      const viaAdded = writeGeneratedAddress(field(form, "indirizzo"), result.via);

      if (comuneAdded || viaAdded) {
        setStatus(form, "Posizione rilevata. Coordinate, Comune e Via compilati automaticamente; puoi modificarli.");
      } else if (result.comune || result.via) {
        setStatus(form, "Posizione rilevata. Comune e Via già compilati: non sono stati sovrascritti.");
      } else {
        setStatus(form, "Posizione rilevata. Coordinate compilate; Comune e Via non trovati automaticamente.");
      }
    } catch (error) {
      if (error && error.name === "AbortError") return;
      if (requestId !== requestSequence || !form.isConnected) return;
      console.warn("Reverse geocoding posizione commessa non riuscito:", error);
      setStatus(form, "Posizione rilevata. Coordinate compilate; Comune e Via possono essere inseriti manualmente.");
    }
  }

  function acquire(button) {
    const form = button && button.closest("form");
    if (!form || form.id !== FORM_ID) return;

    const geolocation = root.navigator && root.navigator.geolocation;
    if (!geolocation || typeof geolocation.getCurrentPosition !== "function") {
      setButton(button, DEFAULT_LABEL, false);
      setStatus(form, "La posizione automatica non è disponibile su questo dispositivo.");
      return;
    }

    const requestId = ++requestSequence;
    if (reverseController) reverseController.abort();
    setButton(button, "Rilevamento posizione…", true);
    setStatus(form, "Rilevamento GPS in corso…");

    geolocation.getCurrentPosition(
      function onSuccess(position) {
        if (requestId !== requestSequence || !form.isConnected) return;

        const latitude = validCoordinate(position && position.coords && position.coords.latitude, -90, 90);
        const longitude = validCoordinate(position && position.coords && position.coords.longitude, -180, 180);
        if (latitude == null || longitude == null) {
          setButton(button, DEFAULT_LABEL, false);
          setStatus(form, "Il telefono ha restituito coordinate non valide. Riprova.");
          return;
        }

        const latitudeInput = field(form, "latitudine");
        const longitudeInput = field(form, "longitudine");
        if (latitudeInput) latitudeInput.value = latitude.toFixed(6);
        if (longitudeInput) longitudeInput.value = longitude.toFixed(6);

        form.dataset.mobileGeocodeKey = "";
        setButton(button, "✓ Posizione acquisita", false);
        setStatus(form, "Posizione rilevata. Compilazione di Comune e Via in corso…");
        void reverseGeocode(form, latitude, longitude, requestId);
      },
      function onError(error) {
        if (requestId !== requestSequence || !form.isConnected) return;
        setButton(button, DEFAULT_LABEL, false);
        setStatus(form, geolocationErrorMessage(error));
      },
      {
        enableHighAccuracy: true,
        timeout: 15000,
        maximumAge: 0
      }
    );
  }

  /*
   * Capture phase: intercetta il pulsante prima del vecchio listener delegato.
   * Così il vecchio flusso non viene eseguito e la vista Gestione commessa
   * non viene ricreata o chiusa durante l'acquisizione GPS.
   */
  document.addEventListener("click", function onCurrentLocationClick(event) {
    const target = event.target && event.target.closest
      ? event.target.closest(`#${BUTTON_ID}`)
      : null;
    if (!target) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();
    acquire(target);
  }, true);
})(window);
