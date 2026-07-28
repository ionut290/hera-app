/* Riepilogo economico live della commessa, basato sul calcolo condiviso della contabilità. */
(() => {
  "use strict";

  const widget = document.getElementById("commessa-produced-widget");
  const toggle = document.getElementById("commessa-produced-toggle");
  const value = document.getElementById("commessa-produced-value");
  const popover = document.getElementById("commessa-produced-popover");
  const money = new Intl.NumberFormat("it-IT", { style: "currency", currency: "EUR" });
  let unsubscribeWork = null;
  let unsubscribeLegacy = null;
  let unsubscribePrices = null;
  let unsubscribeCommessa = null;
  let workItems = [];
  let legacyPlants = [];
  let prices = [];
  let generalDiscount = 0;
  let selectedId = "";

  function close() {
    toggle.setAttribute("aria-expanded", "false");
    popover.setAttribute("aria-hidden", "true");
    popover.classList.remove("is-open");
  }

  function positionPopover() {
    const rect = toggle.getBoundingClientRect();
    const margin = 8;
    const width = Math.min(245, window.innerWidth - margin * 2);
    const left = Math.min(Math.max(margin, rect.right - width), window.innerWidth - width - margin);
    const below = rect.bottom + margin;
    const estimatedHeight = 82;
    const top = below + estimatedHeight <= window.innerHeight
      ? below
      : Math.max(margin, rect.top - estimatedHeight - margin);
    popover.style.width = `${width}px`;
    popover.style.left = `${left}px`;
    popover.style.top = `${top}px`;
  }

  function render() {
    const core = window.InreteWorkItemsV2;
    if (!core) return;
    const sourceItems = workItems.length
      ? workItems
      : legacyPlants.flatMap((plant) => core.adaptLegacyPlantToWorkItems(plant));
    const priceMap = core.buildPriceMap(prices);
    const items = sourceItems.map((item) => core.enrichWorkItem(item, priceMap, generalDiscount));
    value.textContent = money.format(core.calculateCompletedSubtotal(items));
  }

  function stop() {
    unsubscribeWork?.();
    unsubscribeLegacy?.();
    unsubscribePrices?.();
    unsubscribeCommessa?.();
    unsubscribeWork = null;
    unsubscribeLegacy = null;
    unsubscribePrices = null;
    unsubscribeCommessa = null;
    workItems = [];
    legacyPlants = [];
    prices = [];
    generalDiscount = 0;
    close();
  }

  function select(commessaId) {
    const nextId = String(commessaId || "").trim();
    if (nextId === selectedId && unsubscribeWork) return;
    stop();
    selectedId = nextId;
    widget.hidden = !selectedId;
    value.textContent = money.format(0);
    if (!selectedId) return;
    const ref = db.collection(getCommesseCollectionName()).doc(selectedId);
    unsubscribeWork = ref.collection("lavorazioni").onSnapshot((snapshot) => {
      workItems = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      render();
    }, (error) => console.error("Aggiornamento valore prodotto non riuscito:", error));
    unsubscribeLegacy = ref.collection("impianti").onSnapshot((snapshot) => {
      legacyPlants = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      render();
    }, (error) => console.error("Aggiornamento valore prodotto storico non riuscito:", error));
    unsubscribePrices = ref.collection("prezziario").onSnapshot((snapshot) => {
      prices = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      render();
    }, (error) => console.error("Aggiornamento prezzi del valore prodotto non riuscito:", error));
    unsubscribeCommessa = ref.onSnapshot((snapshot) => {
      generalDiscount = snapshot.data()?.percentualeRibassoGenerale ?? 0;
      render();
    }, (error) => console.error("Aggiornamento ribasso del valore prodotto non riuscito:", error));
  }

  toggle.addEventListener("click", (event) => {
    event.stopPropagation();
    const opening = toggle.getAttribute("aria-expanded") !== "true";
    if (!opening) return close();
    positionPopover();
    toggle.setAttribute("aria-expanded", "true");
    popover.setAttribute("aria-hidden", "false");
    requestAnimationFrame(() => popover.classList.add("is-open"));
  });
  popover.addEventListener("click", (event) => event.stopPropagation());
  document.addEventListener("click", close);
  document.addEventListener("keydown", (event) => { if (event.key === "Escape") close(); });
  window.addEventListener("resize", () => {
    if (toggle.getAttribute("aria-expanded") === "true") positionPopover();
  });

  window.CommessaProducedWidget = { select, stop };
})();
