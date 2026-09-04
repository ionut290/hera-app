(() => {
  "use strict";

  function restoreMobilePlantEditor(form) {
    if (!form || !document.contains(form)) return false;
    const mobile = document.getElementById("commessa-mobile-management");
    const home = document.getElementById("commessa-mobile-management-home");
    const list = document.getElementById("commessa-mobile-plant-list");
    const editor = document.getElementById("commessa-mobile-plant-editor");
    if (!mobile || !editor || !editor.contains(form)) return false;

    mobile.classList.remove("hidden");
    mobile.setAttribute("aria-hidden", "false");
    home?.classList.add("hidden");
    home?.setAttribute("aria-hidden", "true");
    list?.classList.add("hidden");
    list?.setAttribute("aria-hidden", "true");
    editor.classList.remove("hidden");
    editor.setAttribute("aria-hidden", "false");
    return true;
  }

  function handleMobileCurrentLocation(event) {
    const button = event.target.closest?.("#commessa-mobile-current-location");
    if (!button) return;

    event.preventDefault();
    event.stopImmediatePropagation();

    const form = button.closest("#commessa-mobile-plant-form");
    if (!form) return;
    const status = form.querySelector("#commessa-mobile-geocode-status");
    const setStatus = (message) => { if (status) status.textContent = message; };
    restoreMobilePlantEditor(form);

    if (!navigator.geolocation) {
      setStatus("La posizione automatica non è disponibile su questo dispositivo.");
      return;
    }

    let requestFinished = false;
    const keepEditorOpen = () => {
      if (!requestFinished && !document.hidden) restoreMobilePlantEditor(form);
    };
    const cleanup = () => {
      requestFinished = true;
      window.removeEventListener("focus", keepEditorOpen);
      document.removeEventListener("visibilitychange", keepEditorOpen);
    };
    window.addEventListener("focus", keepEditorOpen);
    document.addEventListener("visibilitychange", keepEditorOpen);

    button.disabled = true;
    button.textContent = "Rilevamento posizione…";
    setStatus("Rilevamento GPS in corso…");

    navigator.geolocation.getCurrentPosition(position => {
      cleanup();
      if (!document.contains(button) || button.closest("#commessa-mobile-plant-form") !== form) return;
      restoreMobilePlantEditor(form);
      const latitude = Number(position?.coords?.latitude);
      const longitude = Number(position?.coords?.longitude);
      if (!Number.isFinite(latitude) || !Number.isFinite(longitude) || latitude < -90 || latitude > 90 || longitude < -180 || longitude > 180) {
        button.disabled = false;
        button.textContent = "📍 Usa la mia posizione";
        setStatus("Il telefono ha restituito coordinate non valide. Riprova.");
        return;
      }

      const latitudeInput = form.querySelector('[data-v2-field="latitudine"]');
      const longitudeInput = form.querySelector('[data-v2-field="longitudine"]');
      if (latitudeInput) latitudeInput.value = latitude.toFixed(6);
      if (longitudeInput) longitudeInput.value = longitude.toFixed(6);
      form.dataset.mobileGeocodeKey = "";
      latitudeInput?.dispatchEvent(new Event("input", { bubbles: true }));
      longitudeInput?.dispatchEvent(new Event("input", { bubbles: true }));

      restoreMobilePlantEditor(form);
      button.disabled = false;
      button.textContent = "✓ Posizione acquisita";
      setStatus("Posizione acquisita. Ricerca automatica di Comune e Via…");
    }, error => {
      cleanup();
      if (!document.contains(button) || button.closest("#commessa-mobile-plant-form") !== form) return;
      restoreMobilePlantEditor(form);
      button.disabled = false;
      button.textContent = "📍 Usa la mia posizione";
      if (error?.code === 1) setStatus("Non è stato possibile accedere alla posizione. Consenti la localizzazione, verifica che il GPS sia attivo e riprova.");
      else if (error?.code === 2) setStatus("Posizione non disponibile. Verifica che il GPS del telefono sia attivo e riprova.");
      else if (error?.code === 3) setStatus("Rilevamento della posizione scaduto. Spostati in un punto con più segnale e riprova.");
      else setStatus("Non è stato possibile rilevare la posizione. Verifica il GPS e riprova.");
    }, { enableHighAccuracy: true, timeout: 15000, maximumAge: 0 });
  }

  document.addEventListener("click", handleMobileCurrentLocation, true);
})();
