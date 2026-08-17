(function installSquadreCommessaThemes(){
  "use strict";
  if (window.__heraSquadreCommessaThemesInstalled) return;
  window.__heraSquadreCommessaThemesInstalled = true;

  const THEME_CLASSES = [
    "commessa-theme-depurazione",
    "commessa-theme-inrete",
    "commessa-theme-wte",
    "commessa-theme-discarica",
    "commessa-theme-compostaggio",
    "commessa-theme-verde",
    "commessa-theme-acqua"
  ];

  function normalize(value){
    return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase();
  }

  function detectTheme(text){
    const value = normalize(text);
    if (/depur|idr|trattamento acque|reflue|fogn/.test(value)) return "commessa-theme-depurazione";
    if (/inrete|gas|cabina gas|metano|rete gas/.test(value)) return "commessa-theme-inrete";
    if (/wte|termoval|termoutil|incener|energy from waste|waste to energy/.test(value)) return "commessa-theme-wte";
    if (/discaric/.test(value)) return "commessa-theme-discarica";
    if (/compost|organico|biogas|digest/.test(value)) return "commessa-theme-compostaggio";
    if (/verde|giardin|sfalcio|potatur|parco|aree verdi/.test(value)) return "commessa-theme-verde";
    if (/acquedott|potabil|serbatoio acqua|centrale idrica/.test(value)) return "commessa-theme-acqua";
    return "";
  }

  function isSquadreArea(element){
    if (!element) return false;
    const marker = normalize(`${element.id || ""} ${element.className || ""} ${element.getAttribute?.("data-view") || ""} ${element.getAttribute?.("aria-label") || ""}`);
    if (/squadr/.test(marker)) return true;
    let node = element.parentElement;
    for (let i = 0; node && i < 8; i += 1, node = node.parentElement) {
      const ancestorMarker = normalize(`${node.id || ""} ${node.className || ""} ${node.getAttribute?.("data-view") || ""}`);
      if (/squadr/.test(ancestorMarker) || node.id === "today-squads-section") return true;
    }
    return false;
  }

  function candidateCards(root){
    const selectors = [
      "#today-squads-section .squadre-lista > *",
      ".squadre-commessa-card",
      ".squadra-commessa-card",
      "[data-squadre-commessa-card]",
      "[data-commessa-id]",
      ".commessa-card",
      ".card"
    ];
    const found = new Set();
    selectors.forEach((selector) => {
      root.querySelectorAll?.(selector).forEach((el) => {
        if (isSquadreArea(el)) found.add(el);
      });
    });
    return Array.from(found);
  }

  function applyTheme(card){
    const theme = detectTheme(card.textContent || "");
    THEME_CLASSES.forEach((className) => card.classList.remove(className));
    card.classList.remove("commessa-theme-card");
    if (!theme) return;
    card.classList.add("commessa-theme-card", theme);
  }

  function scan(root=document){
    candidateCards(root).forEach(applyTheme);
  }

  let scheduled = false;
  function scheduleScan(){
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      scan();
    });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", scheduleScan, {once:true});
  else scheduleScan();

  new MutationObserver(scheduleScan).observe(document.documentElement, {childList:true, subtree:true, characterData:true});
  window.HeraSquadreCommessaThemes = { refresh: scheduleScan, detectTheme };
})();
