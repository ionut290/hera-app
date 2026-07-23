(function () {
  "use strict";

  window.fetchCompatibleFuelStations = async function (position, radiusKm, fuel) {
    if (!window.HeraFuelStationSearch?.search) throw new Error("Modulo ricerca distributori non caricato");
    fuelStationsAbortController = new AbortController();
    return window.HeraFuelStationSearch.search({
      position,
      radiusKm,
      fuel,
      distanceFn: haversine,
      signal: fuelStationsAbortController.signal,
      onProgress: (source) => {
        ui.fuelStationsList.innerHTML = `<p class="muted">Ricerca tramite ${escapeHTML(source)}...</p>`;
      }
    });
  };
})();
