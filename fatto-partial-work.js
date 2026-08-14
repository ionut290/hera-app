(() => {
  "use strict";

  const text = (value) => String(value ?? "").trim();
  const normalize = (value) => text(value).toLocaleLowerCase("it-IT");
  const isDone = (item) => ["fatto", "done", "completato"].includes(normalize(item?.stato || (item?.done ? "fatto" : "")));
  const typeOf = (item) => {
    const value = normalize(item?.tipoManutenzione || item?.tipologiaManutenzione || item?.tipologiaLavorazione || item?.tipologiaIntervento || item?.descrizione || "");
    if (/straordinar/.test(value)) return "straordinaria";
    if (/ordinar/.test(value)) return "ordinaria";
    return "";
  };

  function getOpenWorkItems(impianto) {
    const direct = Array.isArray(impianto?.lavorazioni) ? impianto.lavorazioni : [];
    return direct.filter((item) => !isDone(item));
  }

  function getOpenTypes(impianto) {
    return [...new Set(getOpenWorkItems(impianto).map(typeOf).filter(Boolean))];
  }

  function getSelectedWorkItems(impianto, selectedTypes) {
    const wanted = new Set(selectedTypes || []);
    return getOpenWorkItems(impianto).filter((item) => wanted.has(typeOf(item)));
  }

  function getSelectedSourceIds(impianto, selectedTypes) {
    const selectedItems = getSelectedWorkItems(impianto, selectedTypes);
    const ids = selectedItems.map((item) => text(item.id || item.sourceId || item.migrationSourceId)).filter(Boolean);
    if (ids.length) return [...new Set(ids)];
    const fallback = Array.isArray(impianto?.sourceIds) ? impianto.sourceIds.map(text).filter(Boolean) : [];
    return selectedTypes?.length === getOpenTypes(impianto).length ? fallback : [];
  }

  function buildSelectionMessage(selectedTypes) {
    const selected = new Set(selectedTypes || []);
    if (selected.size === 1 && selected.has("ordinaria")) {
      return "Hai scelto solo la manutenzione ordinaria. La manutenzione straordinaria resterà da fare e il puntino dell’impianto rimarrà visibile.";
    }
    if (selected.size === 1 && selected.has("straordinaria")) {
      return "Hai scelto solo la manutenzione straordinaria. La manutenzione ordinaria resterà da fare e il puntino dell’impianto rimarrà visibile.";
    }
    return "";
  }

  function buildWorkMessage(selectedTypes) {
    const selected = new Set(selectedTypes || []);
    if (selected.size === 1 && selected.has("ordinaria")) {
      return "🟢 INTERVENTO ESEGUITO\nIntervento eseguito: manutenzione ordinaria";
    }
    if (selected.size === 1 && selected.has("straordinaria")) {
      return "🟠 INTERVENTO STRAORDINARIO ESEGUITO\nIntervento eseguito: manutenzione straordinaria";
    }
    if (selected.has("ordinaria") && selected.has("straordinaria")) {
      return "🟢 INTERVENTI ESEGUITI\nInterventi eseguiti: manutenzione ordinaria e straordinaria";
    }
    return "";
  }

  function askSelection(impianto) {
    const types = getOpenTypes(impianto);
    if (!(types.includes("ordinaria") && types.includes("straordinaria"))) {
      return Promise.resolve(types.length ? types : null);
    }

    return new Promise((resolve) => {
      const overlay = document.createElement("div");
      overlay.className = "fatto-work-choice-overlay";
      overlay.innerHTML = `
        <section class="fatto-work-choice" role="dialog" aria-modal="true" aria-labelledby="fatto-work-choice-title">
          <h2 id="fatto-work-choice-title">Cosa hai eseguito?</h2>
          <label><input type="checkbox" value="ordinaria"> <span>Manutenzione ordinaria</span></label>
          <label><input type="checkbox" value="straordinaria"> <span>Manutenzione straordinaria</span></label>
          <p class="fatto-work-choice-feedback" role="status" aria-live="polite"></p>
          <div class="fatto-work-choice-actions">
            <button type="button" data-action="cancel">ANNULLA</button>
            <button type="button" data-action="continue">CONTINUA</button>
          </div>
        </section>`;
      document.body.appendChild(overlay);
      const finish = (value) => { overlay.remove(); resolve(value); };
      overlay.querySelector('[data-action="cancel"]').addEventListener("click", () => finish(null));
      overlay.querySelector('[data-action="continue"]').addEventListener("click", () => {
        const selected = [...overlay.querySelectorAll('input[type="checkbox"]:checked')].map((input) => input.value);
        const feedback = overlay.querySelector(".fatto-work-choice-feedback");
        if (!selected.length) {
          feedback.textContent = "Seleziona almeno un intervento.";
          return;
        }
        const warning = buildSelectionMessage(selected);
        if (warning && !window.confirm(`${warning}\n\nConfermi e vuoi continuare con l’invio?`)) return;
        finish(selected);
      });
    });
  }

  function selectionOptions(impianto, selectedTypes) {
    const allOpenTypes = getOpenTypes(impianto);
    const partial = Array.isArray(selectedTypes) && selectedTypes.length > 0 && selectedTypes.length < allOpenTypes.length;
    return {
      partialWorkCompletion: partial,
      selectedWorkTypes: [...selectedTypes],
      selectedSourceIds: getSelectedSourceIds(impianto, selectedTypes),
      remainingWorkTypes: allOpenTypes.filter((type) => !selectedTypes.includes(type)),
      whatsappWorkMessage: buildWorkMessage(selectedTypes)
    };
  }

  function install() {
    const original = window.handleImpiantoWhatsAppClick;
    if (typeof original !== "function" || original.__heraPartialWorkWrapped) return false;
    const wrapped = async function (impianto, ...args) {
      const selectedTypes = await askSelection(impianto);
      if (selectedTypes === null) return false;
      const options = selectionOptions(impianto, selectedTypes || []);
      const nextArgs = args.slice();
      if (nextArgs[0] && typeof nextArgs[0] === "object" && !Array.isArray(nextArgs[0])) nextArgs[0] = { ...nextArgs[0], ...options };
      else nextArgs.unshift(options);
      const decorated = { ...impianto, __fattoPartialWork: options };
      return original.call(this, decorated, ...nextArgs);
    };
    wrapped.__heraPartialWorkWrapped = true;
    wrapped.__original = original;
    window.handleImpiantoWhatsAppClick = wrapped;
    return true;
  }

  let attempts = 0;
  const timer = window.setInterval(() => {
    attempts += 1;
    if (install() || attempts >= 160) window.clearInterval(timer);
  }, 250);
  window.addEventListener("hera:data-ready", install);
  window.addEventListener("hera:auth-ready", install);
  window.HeraFattoPartialWork = Object.freeze({ getOpenTypes, getSelectedSourceIds, buildSelectionMessage, buildWorkMessage, selectionOptions, askSelection, install });
})();
