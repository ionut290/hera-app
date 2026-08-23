(() => {
  'use strict';

  if (window.__heraMapTodoFilterInstalled) return;
  window.__heraMapTodoFilterInstalled = true;

  const TODO_BUTTON_ID = 'view-todo-btn';
  const DONE_BUTTON_ID = 'view-done-btn';
  let todoFilterActive = false;
  let markerObserver = null;
  let observedPane = null;

  function getCurrentImpianti() {
    try {
      if (typeof currentImpianti !== 'undefined' && Array.isArray(currentImpianti)) return currentImpianti;
    } catch (_) {}
    try {
      if (Array.isArray(window.currentImpianti)) return window.currentImpianti;
    } catch (_) {}
    return [];
  }

  function getMapInstance() {
    try {
      if (typeof map !== 'undefined' && map?.fitBounds) return map;
    } catch (_) {}
    try {
      if (window.map?.fitBounds) return window.map;
    } catch (_) {}
    return null;
  }

  function plantCoordinates(impianto) {
    const lat = Number(impianto?.gpsY ?? impianto?.lat ?? impianto?.latitude);
    const lng = Number(impianto?.gpsX ?? impianto?.lng ?? impianto?.lon ?? impianto?.longitude);
    return Number.isFinite(lat) && Number.isFinite(lng) ? [lat, lng] : null;
  }

  function setDoneMarkersVisible(visible) {
    document.querySelectorAll('#map .marker-pin.done, #map-fullscreen-view .marker-pin.done').forEach((pin) => {
      const marker = pin.closest('.leaflet-marker-icon');
      if (!marker) return;
      marker.style.display = visible ? '' : 'none';
      marker.setAttribute('aria-hidden', visible ? 'false' : 'true');
    });
  }

  function fitTodoBounds() {
    if (!todoFilterActive) return;
    const mapInstance = getMapInstance();
    if (!mapInstance || !window.L?.latLngBounds) return;

    const points = getCurrentImpianti()
      .filter((impianto) => !Boolean(impianto?.done))
      .map(plantCoordinates)
      .filter(Boolean);

    if (!points.length) return;
    try {
      const bounds = window.L.latLngBounds(points);
      if (bounds.isValid()) mapInstance.fitBounds(bounds, { padding: [24, 24], maxZoom: 17 });
    } catch (_) {}
  }

  function applyFilter({ fit = false } = {}) {
    setDoneMarkersVisible(!todoFilterActive);
    if (fit) fitTodoBounds();
  }

  function observeMarkers() {
    const pane = document.querySelector('#map .leaflet-marker-pane');
    if (!pane || pane === observedPane) return;
    markerObserver?.disconnect();
    observedPane = pane;
    markerObserver = new MutationObserver(() => {
      if (todoFilterActive) requestAnimationFrame(() => applyFilter({ fit: false }));
    });
    markerObserver.observe(pane, { childList: true, subtree: true });
  }

  function activateTodoFilter() {
    todoFilterActive = true;
    observeMarkers();
    requestAnimationFrame(() => applyFilter({ fit: true }));
    window.setTimeout(() => applyFilter({ fit: false }), 80);
  }

  function clearTodoFilter() {
    todoFilterActive = false;
    applyFilter({ fit: false });
  }

  function bindButtons() {
    const todoButton = document.getElementById(TODO_BUTTON_ID);
    const doneButton = document.getElementById(DONE_BUTTON_ID);

    if (todoButton && todoButton.dataset.mapTodoFilterBound !== '1') {
      todoButton.dataset.mapTodoFilterBound = '1';
      todoButton.addEventListener('click', activateTodoFilter);
    }

    if (doneButton && doneButton.dataset.mapTodoFilterBound !== '1') {
      doneButton.dataset.mapTodoFilterBound = '1';
      doneButton.addEventListener('click', clearTodoFilter);
    }
  }

  function init() {
    bindButtons();
    observeMarkers();
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
