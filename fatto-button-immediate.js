(() => {
  "use strict";

  const DB_NAME = "hera-fatto-sync";
  const STORE = "operations";
  const MAX_ATTEMPTS = 12;
  const STALE_MS = 120000;
  const YELLOW = "#f4c542";
  const YELLOW_BORDER = "#c99700";
  let processing = false;
  let partialDialogOpen = false;

  const text = (value) => String(value ?? "").trim();
  const clone = (value) => JSON.parse(JSON.stringify(value || {}));
  const normalizeStatus = (value) => text(value || "DA FARE").toLocaleUpperCase("it-IT").replace(/_/g, " ");
  const isWorkItemDone = (item) => Boolean(item?.done) || ["FATTO", "DONE", "COMPLETATO"].includes(normalizeStatus(item?.stato));

  function openDb() {
    return new Promise((resolve, reject) => {
      if (!("indexedDB" in window)) return reject(new Error("IndexedDB non disponibile"));
      const request = indexedDB.open(DB_NAME, 1);
      request.onupgradeneeded = () => {
        if (!request.result.objectStoreNames.contains(STORE)) {
          request.result.createObjectStore(STORE, { keyPath: "operationId" });
        }
      };
      request.onsuccess = () => resolve(request.result);
      request.onerror = () => reject(request.error || new Error("Coda FATTO non disponibile"));
    });
  }

  async function transact(mode, callback) {
    const db = await openDb();
    try {
      return await new Promise((resolve, reject) => {
        const transaction = db.transaction(STORE, mode);
        const store = transaction.objectStore(STORE);
        const result = callback(store);
        transaction.oncomplete = () => resolve(result);
        transaction.onerror = () => reject(transaction.error || new Error("Errore coda FATTO"));
        transaction.onabort = () => reject(transaction.error || new Error("Coda FATTO annullata"));
      });
    } finally {
      db.close();
    }
  }

  const put = (operation) => transact("readwrite", (store) => store.put(operation));
  const remove = (operationId) => transact("readwrite", (store) => store.delete(operationId));

  async function list() {
    const db = await openDb();
    try {
      return await new Promise((resolve, reject) => {
        const request = db.transaction(STORE, "readonly").objectStore(STORE).getAll();
        request.onsuccess = () => resolve(Array.isArray(request.result) ? request.result : []);
        request.onerror = () => reject(request.error || new Error("Lettura coda FATTO fallita"));
      });
    } finally {
      db.close();
    }
  }

  function publishStatus(items) {
    const pending = items.filter((item) => item.status !== "COMPLETED");
    document.documentElement.dataset.fattoSyncPending = String(pending.length);
    window.dispatchEvent(new CustomEvent("hera:fatto-sync-status", {
      detail: {
        pending: pending.length,
        syncing: pending.some((item) => item.status === "SYNCING"),
        failed: pending.filter((item) => ["FAILED", "BLOCKED"].includes(item.status)).length
      }
    }));
  }

  async function refreshStatus() {
    try { publishStatus(await list()); } catch (_) {}
  }

  function createOperationId(impianto) {
    const plant = text(impianto?.id || impianto?.idSap || impianto?.sap || impianto?.nome || "impianto")
      .replace(/[^a-zA-Z0-9_-]+/g, "_");
    const user = text(window.auth?.currentUser?.uid || window.auth?.currentUser?.email || "utente")
      .replace(/[^a-zA-Z0-9_-]+/g, "_");
    return `fatto_${plant}_${user}_${Date.now()}`;
  }

  async function enqueue(impianto, metadata = {}) {
    const now = Date.now();
    const operation = {
      operationId: metadata.operationId || createOperationId(impianto),
      type: "IMPIANTO_FATTO",
      status: "PENDING",
      impianto: clone(impianto),
      commessaId: text(metadata.commessaId || window.selectedCommessaId),
      doneAt: metadata.doneAt || new Date(now).toISOString(),
      doneBy: text(metadata.doneBy || window.auth?.currentUser?.displayName || window.auth?.currentUser?.email),
      attempts: 0,
      createdAt: now,
      updatedAt: now,
      lastError: ""
    };
    await put(operation);
    await refreshStatus();
    return operation;
  }

  async function setStatus(operation, status, error = "") {
    const updated = {
      ...operation,
      status,
      updatedAt: Date.now(),
      lastError: text(error?.message || error)
    };
    await put(updated);
    return updated;
  }

  function formatDoneLabel(doneAt) {
    const date = new Date(doneAt);
    if (Number.isNaN(date.getTime())) return "";
    return new Intl.DateTimeFormat("it-IT", {
      day: "2-digit",
      month: "2-digit",
      year: "numeric",
      hour: "2-digit",
      minute: "2-digit"
    }).format(date);
  }

  function findPressedFattoButton() {
    const active = document.activeElement;
    if (!(active instanceof HTMLElement)) return null;
    const button = active.closest("button, [role='button'], input[type='button'], input[type='submit']");
    if (!(button instanceof HTMLElement)) return null;
    const label = text(button.textContent || button.getAttribute("value") || button.getAttribute("aria-label"));
    return /fatto|whazzup|whatsapp/i.test(label) ? button : null;
  }

  function applyPermanentYellowFeedback(button, doneAt) {
    if (!(button instanceof HTMLElement)) return;
    const label = formatDoneLabel(doneAt);
    button.dataset.fattoImmediate = "true";
    button.dataset.fattoDoneAt = doneAt;
    button.disabled = true;
    button.setAttribute("aria-disabled", "true");
    button.style.setProperty("background", YELLOW, "important");
    button.style.setProperty("background-color", YELLOW, "important");
    button.style.setProperty("border-color", YELLOW_BORDER, "important");
    button.style.setProperty("color", "#1d1d1d", "important");
    button.style.setProperty("opacity", "1", "important");
    button.style.setProperty("pointer-events", "none", "important");
    button.style.setProperty("cursor", "default", "important");

    let dateNode = button.previousElementSibling;
    if (!(dateNode instanceof HTMLElement) || dateNode.dataset.fattoImmediateDate !== "true") {
      dateNode = document.createElement("div");
      dateNode.dataset.fattoImmediateDate = "true";
      dateNode.style.fontWeight = "700";
      dateNode.style.fontSize = "0.82rem";
      dateNode.style.marginBottom = "4px";
      dateNode.style.textAlign = "center";
      button.parentNode?.insertBefore(dateNode, button);
    }
    dateNode.textContent = label;

    if ("value" in button && /^(INPUT|BUTTON)$/i.test(button.tagName)) {
      if (button.tagName === "INPUT") button.value = "FATTO";
      else button.textContent = "FATTO";
    } else {
      button.textContent = "FATTO";
    }
  }

  function getCommesseCollection() {
    try {
      if (typeof window.getCommesseCollectionName === "function") return window.getCommesseCollectionName();
      if (typeof getCommesseCollectionName === "function") return getCommesseCollectionName();
    } catch (_) {}
    return "commesse";
  }

  function getPlantId(impianto) {
    return text(impianto?.physicalPlantId || impianto?.impiantoId || impianto?.migrationSourceId || impianto?.id);
  }

  function getWorkItemTitle(item) {
    const code = text(item?.codiceVocePrezzo || item?.codicePrezzo || item?.codice);
    const description = text(item?.tipologiaLavorazione || item?.tipologiaIntervento || item?.descrizione || item?.nome);
    if (code && description) return `${code} · ${description}`;
    return description || code || "Lavorazione";
  }

  function getWorkItemKind(item) {
    const raw = [item?.tipo, item?.categoria, item?.tipologia, item?.tipologiaLavorazione, item?.tipologiaIntervento]
      .map(text).join(" ").toLocaleUpperCase("it-IT");
    return raw.includes("STRAORD") ? "STRAORDINARIO" : "ORDINARIO";
  }

  function isInreteCommessaData(commessa) {
    return [commessa?.nome, commessa?.codice, commessa?.categoria, commessa?.tipo, commessa?.commessaPadre, commessa?.parentName]
      .map(text).join(" ").toLocaleUpperCase("it-IT").includes("INRETE");
  }

  function chooseWorkItem(items, impianto) {
    if (partialDialogOpen) return Promise.resolve(null);
    partialDialogOpen = true;
    return new Promise((resolve) => {
      const overlay = document.createElement("div");
      overlay.dataset.inretePartialFatto = "true";
      Object.assign(overlay.style, {
        position: "fixed", inset: "0", zIndex: "2147483646", background: "rgba(0,0,0,.68)",
        display: "flex", alignItems: "center", justifyContent: "center", padding: "16px"
      });
      const card = document.createElement("div");
      Object.assign(card.style, {
        width: "min(520px,100%)", maxHeight: "86vh", overflow: "auto", background: "#fff", color: "#111",
        borderRadius: "18px", padding: "18px", boxShadow: "0 20px 60px rgba(0,0,0,.35)"
      });
      const title = document.createElement("h2");
      title.textContent = "Quale lavorazione hai completato?";
      title.style.margin = "0 0 6px";
      const subtitle = document.createElement("p");
      subtitle.textContent = text(impianto?.denominazione || impianto?.nome || impianto?.idSap || "Impianto");
      subtitle.style.margin = "0 0 14px";
      subtitle.style.opacity = ".72";
      card.append(title, subtitle);

      const finish = (value) => {
        partialDialogOpen = false;
        overlay.remove();
        resolve(value);
      };

      items.forEach((item) => {
        const button = document.createElement("button");
        button.type = "button";
        button.dataset.workItemId = item.id;
        Object.assign(button.style, {
          display: "block", width: "100%", textAlign: "left", border: "1px solid #d7d7d7", borderRadius: "12px",
          background: "#f7f7f7", color: "#111", padding: "13px", margin: "0 0 9px", font: "inherit", cursor: "pointer"
        });
        const kind = document.createElement("strong");
        kind.textContent = getWorkItemKind(item);
        kind.style.display = "block";
        kind.style.fontSize = ".78rem";
        kind.style.marginBottom = "3px";
        const label = document.createElement("span");
        label.textContent = getWorkItemTitle(item);
        button.append(kind, label);
        button.addEventListener("click", () => finish(item.id));
        card.appendChild(button);
      });

      const allButton = document.createElement("button");
      allButton.type = "button";
      allButton.textContent = "TUTTE LE LAVORAZIONI SONO FATTE";
      Object.assign(allButton.style, {
        width: "100%", border: "0", borderRadius: "12px", background: "#f4c542", color: "#111",
        fontWeight: "800", padding: "13px", marginTop: "4px", cursor: "pointer"
      });
      allButton.addEventListener("click", () => finish("__ALL__"));
      card.appendChild(allButton);

      const cancel = document.createElement("button");
      cancel.type = "button";
      cancel.textContent = "ANNULLA";
      Object.assign(cancel.style, {
        width: "100%", border: "0", background: "transparent", color: "#555", padding: "12px", marginTop: "3px", cursor: "pointer"
      });
      cancel.addEventListener("click", () => finish(null));
      card.appendChild(cancel);
      overlay.appendChild(card);
      overlay.addEventListener("click", (event) => { if (event.target === overlay) finish(null); });
      document.body.appendChild(overlay);
    });
  }

  function openPartialWhatsApp(impianto, item, doneAt, doneBy, remaining) {
    const when = formatDoneLabel(doneAt);
    const lines = [
      "🟡 LAVORAZIONE FATTA",
      `Impianto: ${text(impianto?.denominazione || impianto?.nome || impianto?.idSap || "—")}`,
      text(impianto?.comune) ? `Comune: ${text(impianto.comune)}` : "",
      `Tipo: ${getWorkItemKind(item)}`,
      `Lavorazione: ${getWorkItemTitle(item)}`,
      when ? `Data/Ora: ${when}` : "",
      doneBy ? `Operatore: ${doneBy}` : "",
      `Lavorazioni ancora da fare: ${remaining}`
    ].filter(Boolean);
    const url = `whatsapp://send?text=${encodeURIComponent(lines.join("\n"))}`;
    if (typeof window.openWhatsApp === "function") {
      try { return window.openWhatsApp(url); } catch (_) {}
    }
    try {
      window.location.href = url;
      return true;
    } catch (_) {
      return false;
    }
  }

  async function loadInreteWorkContext(impianto) {
    const firestore = window.db;
    const commessaId = text(window.selectedCommessaId);
    const plantId = getPlantId(impianto);
    if (!firestore || !commessaId || !plantId || !navigator.onLine) return null;
    try {
      const commessaRef = firestore.collection(getCommesseCollection()).doc(commessaId);
      const [commessaSnap, worksSnap] = await Promise.all([
        commessaRef.get(),
        commessaRef.collection("lavorazioni").where("impiantoId", "==", plantId).get()
      ]);
      const commessa = commessaSnap.exists ? { id: commessaSnap.id, ...commessaSnap.data() } : null;
      if (!commessa || !isInreteCommessaData(commessa) || worksSnap.size < 2) return null;
      const items = worksSnap.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      return { commessaRef, commessa, plantId, items };
    } catch (error) {
      console.warn("[FATTO parziale] lettura lavorazioni non riuscita; uso flusso FATTO classico", error);
      return null;
    }
  }

  async function savePartialWorkItem(context, workItemId, impianto) {
    const item = context.items.find((entry) => entry.id === workItemId);
    if (!item) throw new Error("Lavorazione non trovata");
    const now = new Date();
    const doneAt = now.toISOString();
    const user = window.auth?.currentUser;
    const doneBy = text(user?.displayName || user?.email || "Operatore");
    const pad = (value) => String(value).padStart(2, "0");
    const dataEsecuzione = `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
    const oraEsecuzione = `${pad(now.getHours())}:${pad(now.getMinutes())}`;
    const nextItems = context.items.map((entry) => entry.id === workItemId
      ? { ...entry, stato: "FATTO", done: true, doneAt, dataEsecuzione, oraEsecuzione, operatoreNome: doneBy, operatoreUid: text(user?.uid) }
      : entry);
    const doneCount = nextItems.filter(isWorkItemDone).length;
    const allDone = doneCount === nextItems.length;
    const stato = allDone ? "FATTO" : (doneCount ? "PARZIALMENTE FATTO" : "DA FARE");
    const workRef = context.commessaRef.collection("lavorazioni").doc(workItemId);
    const physicalRef = context.commessaRef.collection("impiantiFisici").doc(context.plantId);
    const operationalRef = context.commessaRef.collection("impianti").doc(context.plantId);
    const payload = {
      stato,
      statoGenerale: stato,
      done: allDone,
      numeroLavorazioni: nextItems.length,
      numeroLavorazioniFatte: doneCount,
      numeroLavorazioniDaFare: nextItems.length - doneCount,
      updatedAt: window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || now
    };
    const batch = window.db.batch();
    batch.set(workRef, {
      stato: "FATTO",
      done: true,
      doneAt: window.firebase?.firestore?.Timestamp?.fromDate?.(now) || now,
      dataEsecuzione,
      oraEsecuzione,
      operatoreNome: doneBy,
      operatoreUid: text(user?.uid),
      operatoreEmail: text(user?.email)
    }, { merge: true });
    batch.set(physicalRef, payload, { merge: true });
    batch.set(operationalRef, payload, { merge: true });
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
    try { window.dispatchEvent(new CustomEvent("hera:inrete-work-item-done", { detail: { commessaId: text(window.selectedCommessaId), impiantoId: context.plantId, workItemId, stato } })); } catch (_) {}
    return { item, allDone, remaining: nextItems.length - doneCount, doneAt, doneBy };
  }

  async function maybeHandlePartialInreteDone(impianto) {
    const context = await loadInreteWorkContext(impianto);
    if (!context) return { handled: false };
    const unfinished = context.items.filter((item) => !isWorkItemDone(item));
    if (!unfinished.length) return { handled: false };
    const selected = await chooseWorkItem(unfinished, impianto);
    if (!selected) return { handled: true, result: false };
    if (selected === "__ALL__") return { handled: false };

    const saved = await savePartialWorkItem(context, selected, impianto);
    if (saved.allDone) return { handled: false };
    openPartialWhatsApp(impianto, saved.item, saved.doneAt, saved.doneBy, saved.remaining);
    return { handled: true, result: true };
  }

  async function syncOperation(operation) {
    const options = {
      source: "resume-persistent-queue",
      operationId: operation.operationId,
      doneAt: operation.doneAt,
      doneBy: operation.doneBy,
      requireFirestoreConfirmation: false,
      reopenWhatsApp: false
    };
    if (typeof window.forceMoveImpiantoToFatti === "function") {
      return window.forceMoveImpiantoToFatti(operation.impianto, options);
    }
    if (typeof window.markImpiantoDone === "function") {
      return window.markImpiantoDone(operation.impianto, options);
    }
    throw new Error("Funzione FATTO non ancora pronta");
  }

  async function processQueue(reason = "manual") {
    if (processing || !navigator.onLine) return false;
    processing = true;
    document.documentElement.dataset.fattoSyncReason = reason;
    try {
      const items = (await list()).sort((a, b) => a.createdAt - b.createdAt);
      publishStatus(items);
      for (let operation of items) {
        if (operation.status === "SYNCING" && Date.now() - Number(operation.updatedAt || 0) < STALE_MS) continue;
        if (Number(operation.attempts || 0) >= MAX_ATTEMPTS) {
          await setStatus(operation, "BLOCKED", operation.lastError || "Troppi tentativi");
          continue;
        }
        operation = await setStatus({ ...operation, attempts: Number(operation.attempts || 0) + 1 }, "SYNCING");
        try {
          await syncOperation(operation);
          await remove(operation.operationId);
        } catch (error) {
          await setStatus(operation, "FAILED", error);
          if (/non ancora pronta/i.test(text(error?.message))) break;
        }
      }
      await refreshStatus();
      return true;
    } finally {
      processing = false;
    }
  }

  function isAndroidNative() {
    try { return window.Capacitor?.getPlatform?.() === "android"; } catch (_) { return false; }
  }

  function buildNativeWhatsAppUrl(args) {
    for (const value of args) {
      if (typeof value === "string" && /(whatsapp:|wa\.me|api\.whatsapp\.com)/i.test(value)) return value;
      if (value && typeof value === "object") {
        const candidate = value.url || value.href || value.link || value.whatsappUrl;
        if (typeof candidate === "string" && candidate) return candidate;
      }
    }

    const impianto = args[0];
    const options = args[1] && typeof args[1] === "object" ? args[1] : {};
    if (impianto && typeof impianto === "object" && typeof window.buildImpiantoWhatsAppPayload === "function") {
      try {
        const payload = window.buildImpiantoWhatsAppPayload(impianto, options);
        if (typeof payload?.appUrl === "string" && payload.appUrl) return payload.appUrl;
        if (typeof payload?.webUrl === "string" && payload.webUrl) return payload.webUrl;
        if (typeof payload?.message === "string" && text(payload.message)) {
          return `whatsapp://send?text=${encodeURIComponent(payload.message)}`;
        }
      } catch (error) {
        console.error("Errore preparazione messaggio WhatsApp nativo:", error);
      }
    }

    const message = args.find((value) => typeof value === "string" && text(value));
    return message ? `whatsapp://send?text=${encodeURIComponent(message)}` : "";
  }

  function installWhatsAppWrapper() {
    const original = window.openWhatsApp;
    if (typeof original !== "function" || original.__heraNativeWrapped) return;
    const wrapped = function (...args) {
      const plugin = window.Capacitor?.Plugins?.HeraWhatsApp;
      if (isAndroidNative() && plugin?.open) {
        const nativeUrl = buildNativeWhatsAppUrl(args);
        if (!nativeUrl) return original.apply(this, args);
        plugin.open({ url: nativeUrl }).catch((error) => {
          const message = text(error?.message || error || "WhatsApp non è installato sul dispositivo.");
          window.dispatchEvent(new CustomEvent("hera:whatsapp-error", { detail: { message } }));
          if (typeof window.alert === "function") window.alert(message);
        });
        return true;
      }
      return original.apply(this, args);
    };
    wrapped.__heraNativeWrapped = true;
    wrapped.__original = original;
    window.openWhatsApp = wrapped;
  }

  function installFattoWrapper() {
    const original = window.handleImpiantoWhatsAppClick;
    if (typeof original !== "function" || original.__heraQueueWrapped) return;
    const wrapped = async function (impianto, ...args) {
      try {
        const partial = await maybeHandlePartialInreteDone(impianto);
        if (partial.handled) return partial.result;
      } catch (error) {
        console.error("[FATTO parziale] errore; nessuna lavorazione è stata marcata senza conferma", error);
        if (typeof window.alert === "function") window.alert(`Impossibile completare la lavorazione: ${text(error?.message || error)}`);
        return false;
      }

      const pressedButton = findPressedFattoButton();
      const doneAt = new Date().toISOString();
      applyPermanentYellowFeedback(pressedButton, doneAt);

      const operation = await enqueue(impianto, {
        commessaId: window.selectedCommessaId,
        doneAt
      });
      try {
        const result = await original.call(this, impianto, ...args);
        if (result === true) await remove(operation.operationId);
        else await setStatus(operation, "FAILED", "Flusso FATTO non completato");
        await refreshStatus();
        return result;
      } catch (error) {
        await setStatus(operation, "FAILED", error);
        await refreshStatus();
        throw error;
      }
    };
    wrapped.__heraQueueWrapped = true;
    wrapped.__original = original;
    window.handleImpiantoWhatsAppClick = wrapped;
  }

  function resume(reason) {
    installWhatsAppWrapper();
    installFattoWrapper();
    window.setTimeout(() => processQueue(reason), 100);
  }

  window.addEventListener("online", () => resume("online"));
  window.addEventListener("pageshow", () => resume("pageshow"));
  window.addEventListener("focus", () => resume("focus"));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") resume("visible");
  });
  window.addEventListener("hera:auth-ready", () => resume("auth-ready"));
  window.addEventListener("hera:data-ready", () => resume("data-ready"));
  window.addEventListener("hera:native-resume", () => resume("native-resume"));

  let wrapperAttempts = 0;
  const wrapperTimer = window.setInterval(() => {
    wrapperAttempts += 1;
    installWhatsAppWrapper();
    installFattoWrapper();
    if (wrapperAttempts >= 120 || (window.openWhatsApp?.__heraNativeWrapped && window.handleImpiantoWhatsAppClick?.__heraQueueWrapped)) {
      window.clearInterval(wrapperTimer);
    }
  }, 250);

  window.HeraFattoSync = Object.freeze({ enqueue, processQueue, refreshStatus, list, maybeHandlePartialInreteDone });
  refreshStatus();
  resume("startup");
})();

(() => {
  "use strict";
  if (document.querySelector('script[data-password-access-manager="true"]')) return;
  const script = document.createElement("script");
  script.src = "password-access-manager.js?v=20260727a";
  script.defer = true;
  script.dataset.passwordAccessManager = "true";
  document.head.appendChild(script);
})();
