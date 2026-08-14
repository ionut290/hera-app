(() => {
  "use strict";

  if (window.__heraOrdinaryExtraordinaryFattoInstalled) return;
  window.__heraOrdinaryExtraordinaryFattoInstalled = true;

  const text = (value) => String(value ?? "").trim();
  const normalizeStatus = (value) => text(value || "DA FARE").toLocaleUpperCase("it-IT").replace(/_/g, " ");
  const isWorkItemDone = (item) => Boolean(item?.done) || ["FATTO", "DONE", "COMPLETATO"].includes(normalizeStatus(item?.stato));
  let dialogOpen = false;
  const processingPlants = new Set();

  function getCommesseCollection() {
    try {
      if (typeof window.getCommesseCollectionName === "function") return window.getCommesseCollectionName();
    } catch (_) {}
    return "commesse";
  }

  function getPlantIds(impianto) {
    const ids = [
      impianto?.physicalPlantId,
      impianto?.impiantoId,
      impianto?.migrationSourceId,
      impianto?.id,
      ...(Array.isArray(impianto?.sourceIds) ? impianto.sourceIds : [])
    ].map(text).filter(Boolean);
    return [...new Set(ids)];
  }

  function getWorkItemKind(item) {
    const raw = [item?.tipo, item?.categoria, item?.tipologia, item?.tipologiaLavorazione, item?.tipologiaIntervento, item?.descrizione]
      .map(text).join(" ").toLocaleUpperCase("it-IT");
    return raw.includes("STRAORD") ? "STRAORDINARIO" : "ORDINARIO";
  }

  function isInreteCommessaData(commessa) {
    return [commessa?.nome, commessa?.codice, commessa?.categoria, commessa?.tipo, commessa?.commessaPadre, commessa?.parentName]
      .map(text).join(" ").toLocaleUpperCase("it-IT").includes("INRETE");
  }

  async function loadContext(impianto) {
    const firestore = window.db;
    const commessaId = text(window.selectedCommessaId);
    const plantIds = getPlantIds(impianto);
    if (!firestore || !commessaId || !plantIds.length) return null;
    const commessaRef = firestore.collection(getCommesseCollection()).doc(commessaId);
    try {
      const commessaSnap = await commessaRef.get();
      const commessa = commessaSnap.exists ? { id: commessaSnap.id, ...commessaSnap.data() } : null;
      if (!commessa || !isInreteCommessaData(commessa)) return null;
      const snapshots = await Promise.all(plantIds.map((plantId) =>
        commessaRef.collection("lavorazioni").where("impiantoId", "==", plantId).get()
      ));
      const byId = new Map();
      snapshots.forEach((snapshot) => snapshot.docs.forEach((doc) => byId.set(doc.id, { id: doc.id, ...doc.data() })));
      const items = [...byId.values()];
      if (items.length < 2) return null;
      return { commessaRef, plantIds, items };
    } catch (error) {
      console.warn("[FATTO ordinario/straordinario] contesto non disponibile", error);
      return null;
    }
  }

  function createDialogCard(titleText, subtitleText = "") {
    const overlay = document.createElement("div");
    overlay.dataset.ordinaryExtraordinaryFatto = "true";
    Object.assign(overlay.style, {
      position: "fixed", inset: "0", zIndex: "2147483647", background: "rgba(0,0,0,.68)",
      display: "flex", alignItems: "center", justifyContent: "center", padding: "16px"
    });
    const card = document.createElement("section");
    Object.assign(card.style, {
      width: "min(520px,100%)", maxHeight: "86vh", overflow: "auto", background: "#fff", color: "#111",
      borderRadius: "18px", padding: "20px", boxShadow: "0 20px 60px rgba(0,0,0,.35)"
    });
    const title = document.createElement("h2");
    title.textContent = titleText;
    title.style.margin = "0 0 6px";
    card.appendChild(title);
    if (subtitleText) {
      const subtitle = document.createElement("p");
      subtitle.textContent = subtitleText;
      subtitle.style.margin = "0 0 16px";
      subtitle.style.opacity = ".72";
      card.appendChild(subtitle);
    }
    overlay.appendChild(card);
    document.body.appendChild(overlay);
    return { overlay, card };
  }

  function chooseKinds(items, impianto) {
    if (dialogOpen) return Promise.resolve(null);
    const unfinished = items.filter((item) => !isWorkItemDone(item));
    const availableKinds = [...new Set(unfinished.map(getWorkItemKind))];
    if (!availableKinds.includes("ORDINARIO") || !availableKinds.includes("STRAORDINARIO")) return Promise.resolve(null);
    dialogOpen = true;
    return new Promise((resolve) => {
      const { overlay, card } = createDialogCard(
        "Cosa hai eseguito?",
        text(impianto?.denominazione || impianto?.nome || impianto?.idSap || "Impianto")
      );
      [["ORDINARIO", "Manutenzione ordinaria"], ["STRAORDINARIO", "Manutenzione straordinaria"]].forEach(([kind, labelText]) => {
        const label = document.createElement("label");
        Object.assign(label.style, {
          display: "flex", alignItems: "center", gap: "12px", padding: "14px", margin: "0 0 10px",
          border: "1px solid #d7d7d7", borderRadius: "12px", background: "#f7f7f7", fontWeight: "700"
        });
        const input = document.createElement("input");
        input.type = "checkbox";
        input.value = kind;
        input.style.width = "20px";
        input.style.height = "20px";
        const span = document.createElement("span");
        span.textContent = labelText;
        label.append(input, span);
        card.appendChild(label);
      });
      const feedback = document.createElement("p");
      feedback.setAttribute("role", "status");
      feedback.style.minHeight = "22px";
      feedback.style.margin = "4px 0 10px";
      feedback.style.color = "#a61b1b";
      card.appendChild(feedback);
      const actions = document.createElement("div");
      Object.assign(actions.style, { display: "flex", gap: "10px", justifyContent: "flex-end", flexWrap: "wrap" });
      const cancel = document.createElement("button");
      cancel.type = "button";
      cancel.textContent = "ANNULLA";
      const proceed = document.createElement("button");
      proceed.type = "button";
      proceed.textContent = "CONTINUA";
      proceed.style.cssText = "font-weight:800;background:#f4c542;border:0;border-radius:10px;padding:12px 16px";
      actions.append(cancel, proceed);
      card.appendChild(actions);
      const finish = (value) => { dialogOpen = false; overlay.remove(); resolve(value); };
      cancel.addEventListener("click", () => finish(null));
      overlay.addEventListener("click", (event) => { if (event.target === overlay) finish(null); });
      proceed.addEventListener("click", () => {
        const selected = [...card.querySelectorAll('input[type="checkbox"]:checked')].map((input) => input.value);
        if (!selected.length) {
          feedback.textContent = "Seleziona almeno un intervento.";
          return;
        }
        finish(selected);
      });
    });
  }

  function confirmSingleSelection(selectedKinds) {
    if (selectedKinds.length !== 1) return Promise.resolve(true);
    const ordinaryOnly = selectedKinds[0] === "ORDINARIO";
    const message = ordinaryOnly
      ? "Hai scelto solo la manutenzione ordinaria. La manutenzione straordinaria resterà da fare e il puntino dell’impianto rimarrà visibile."
      : "Hai scelto solo la manutenzione straordinaria. La manutenzione ordinaria resterà da fare e il puntino dell’impianto rimarrà visibile.";
    dialogOpen = true;
    return new Promise((resolve) => {
      const { overlay, card } = createDialogCard("Conferma intervento");
      const paragraph = document.createElement("p");
      paragraph.textContent = message;
      paragraph.style.lineHeight = "1.5";
      card.appendChild(paragraph);
      const actions = document.createElement("div");
      Object.assign(actions.style, { display: "flex", gap: "10px", justifyContent: "flex-end", flexWrap: "wrap" });
      const back = document.createElement("button");
      back.type = "button";
      back.textContent = "TORNA INDIETRO";
      const confirm = document.createElement("button");
      confirm.type = "button";
      confirm.textContent = "CONFERMA E INVIA";
      confirm.style.cssText = "font-weight:800;background:#f4c542;border:0;border-radius:10px;padding:12px 16px";
      actions.append(back, confirm);
      card.appendChild(actions);
      const finish = (value) => { dialogOpen = false; overlay.remove(); resolve(value); };
      back.addEventListener("click", () => finish(false));
      confirm.addEventListener("click", () => finish(true));
    });
  }

  function buildPartialMessage(impianto, selectedKinds, saved) {
    const ordinary = selectedKinds.includes("ORDINARIO");
    const extraordinary = selectedKinds.includes("STRAORDINARIO");
    const heading = ordinary && extraordinary
      ? "🟢 INTERVENTI ESEGUITI"
      : ordinary ? "🟢 INTERVENTO ESEGUITO" : "🟠 INTERVENTO STRAORDINARIO ESEGUITO";
    const intervention = ordinary && extraordinary
      ? "Interventi eseguiti: manutenzione ordinaria e straordinaria"
      : ordinary ? "Intervento eseguito: manutenzione ordinaria" : "Intervento eseguito: manutenzione straordinaria";
    return [
      heading,
      intervention,
      `Impianto: ${text(impianto?.denominazione || impianto?.nome || impianto?.idSap || "—")}`,
      text(impianto?.comune) ? `Comune: ${text(impianto.comune)}` : "",
      saved?.doneBy ? `Operatore: ${saved.doneBy}` : "",
      Number.isFinite(saved?.remaining) ? `Lavorazioni ancora da fare: ${saved.remaining}` : ""
    ].filter(Boolean).join("\n");
  }

  function openPartialWhatsApp(message) {
    const url = `whatsapp://send?text=${encodeURIComponent(message)}`;
    if (typeof window.openWhatsApp === "function") {
      try { return window.openWhatsApp(url); } catch (_) {}
    }
    try { window.location.href = url; return true; } catch (_) { return false; }
  }

  async function saveSelectedWorkItems(context, selectedKinds, impianto) {
    const selectedItems = context.items.filter((entry) => !isWorkItemDone(entry) && selectedKinds.includes(getWorkItemKind(entry)));
    if (!selectedItems.length) throw new Error("Nessuna lavorazione selezionata da completare.");
    const now = new Date();
    const doneAt = now.toISOString();
    const user = window.auth?.currentUser;
    const doneBy = text(user?.displayName || user?.email || "Operatore");
    const pad = (value) => String(value).padStart(2, "0");
    const dataEsecuzione = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const oraEsecuzione = `${pad(now.getHours())}:${pad(now.getMinutes())}`;
    const selectedIds = new Set(selectedItems.map((item) => item.id));
    const nextItems = context.items.map((entry) => selectedIds.has(entry.id)
      ? { ...entry, stato: "FATTO", done: true, doneAt, dataEsecuzione, oraEsecuzione, operatoreNome: doneBy, operatoreUid: text(user?.uid) }
      : entry);
    const doneCount = nextItems.filter(isWorkItemDone).length;
    const allDone = doneCount === nextItems.length;
    const stato = allDone ? "FATTO" : (doneCount ? "PARZIALMENTE FATTO" : "DA FARE");
    const batch = window.db.batch();
    selectedItems.forEach((item) => {
      batch.set(context.commessaRef.collection("lavorazioni").doc(item.id), {
        stato: "FATTO",
        done: true,
        doneAt: window.firebase?.firestore?.Timestamp?.fromDate?.(now) || now,
        dataEsecuzione,
        oraEsecuzione,
        operatoreNome: doneBy,
        operatoreUid: text(user?.uid),
        operatoreEmail: text(user?.email)
      }, { merge: true });
    });
    const payload = {
      stato,
      statoGenerale: stato,
      done: allDone,
      numeroLavorazioni: nextItems.length,
      numeroLavorazioniFatte: doneCount,
      numeroLavorazioniDaFare: nextItems.length - doneCount,
      updatedAt: window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || now
    };
    context.plantIds.forEach((plantId) => {
      batch.set(context.commessaRef.collection("impiantiFisici").doc(plantId), payload, { merge: true });
      batch.set(context.commessaRef.collection("impianti").doc(plantId), payload, { merge: true });
    });
    await batch.commit();
    try {
      impianto.stato = stato;
      impianto.statoGenerale = stato;
      impianto.done = allDone;
      impianto.numeroLavorazioni = nextItems.length;
      impianto.numeroLavorazioniFatte = doneCount;
      impianto.numeroLavorazioniDaFare = nextItems.length - doneCount;
    } catch (_) {}
    try { if (typeof window.renderImpianti === "function") window.renderImpianti(); } catch (_) {}
    return { allDone, remaining: nextItems.length - doneCount, doneAt, doneBy };
  }

  async function handleSelection(impianto) {
    const plantKey = getPlantIds(impianto).join("|");
    if (processingPlants.has(plantKey)) return { handled: true, result: false };
    const context = await loadContext(impianto);
    if (!context) return { handled: false };
    const unfinished = context.items.filter((item) => !isWorkItemDone(item));
    const openKinds = [...new Set(unfinished.map(getWorkItemKind))];
    if (!(openKinds.includes("ORDINARIO") && openKinds.includes("STRAORDINARIO"))) return { handled: false };
    const selectedKinds = await chooseKinds(unfinished, impianto);
    if (!selectedKinds) return { handled: true, result: false };
    if (selectedKinds.length === 2) return { handled: false };
    const confirmed = await confirmSingleSelection(selectedKinds);
    if (!confirmed) return { handled: true, result: false };
    processingPlants.add(plantKey);
    try {
      const saved = await saveSelectedWorkItems(context, selectedKinds, impianto);
      if (saved.allDone) return { handled: false };
      openPartialWhatsApp(buildPartialMessage(impianto, selectedKinds, saved));
      return { handled: true, result: true };
    } finally {
      processingPlants.delete(plantKey);
    }
  }

  function install() {
    const original = window.handleImpiantoWhatsAppClick;
    if (typeof original !== "function" || !original.__heraQueueWrapped || original.__heraOrdinaryExtraordinaryWrapped) return false;
    const wrapped = async function (impianto, ...args) {
      try {
        const partial = await handleSelection(impianto);
        if (partial.handled) return partial.result;
      } catch (error) {
        console.error("[FATTO ordinario/straordinario] errore", error);
        if (typeof window.alert === "function") window.alert(`Impossibile completare l’intervento: ${text(error?.message || error)}`);
        return false;
      }
      return original.call(this, impianto, ...args);
    };
    wrapped.__heraOrdinaryExtraordinaryWrapped = true;
    wrapped.__original = original;
    window.handleImpiantoWhatsAppClick = wrapped;
    return true;
  }

  let attempts = 0;
  const timer = window.setInterval(() => {
    attempts += 1;
    if (install() || attempts >= 160) window.clearInterval(timer);
  }, 250);
  window.addEventListener("hera:auth-ready", install);
  window.addEventListener("hera:data-ready", install);
  window.addEventListener("pageshow", install);
  window.HeraOrdinaryExtraordinaryFatto = Object.freeze({ install, handleSelection, saveSelectedWorkItems, buildPartialMessage });
})();
