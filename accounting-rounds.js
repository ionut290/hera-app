/* Giri contabili: snapshot immutabile prima dello svuotamento della commessa. */
(() => {
  "use strict";

  const clean = value => String(value ?? "").trim();
  const upper = value => clean(value).toUpperCase();
  const server = () => firebase.firestore.FieldValue.serverTimestamp();
  const actorName = () => typeof getOperatorDisplayName === "function" ? getOperatorDisplayName() : (currentUser?.displayName || currentUser?.email || "Operatore");
  const collectionName = () => typeof getCommesseCollectionName === "function" ? getCommesseCollectionName() : "commesse";

  function currentCommessa() {
    const selectedId = typeof selectedCommessaId !== "undefined" ? clean(selectedCommessaId) : "";
    const managementId = typeof managementCommessaId !== "undefined" ? clean(managementCommessaId) : "";
    const ids = [selectedId, managementId].filter(Boolean);
    if (typeof commesseById !== "undefined" && commesseById?.get) {
      for (const id of ids) {
        const commessa = commesseById.get(id);
        if (commessa) return commessa;
      }
    }
    return null;
  }

  function plusDaysIso(days) {
    const d = new Date();
    d.setHours(9, 0, 0, 0);
    d.setDate(d.getDate() + days);
    return d.toISOString();
  }

  function sumDone(rows) {
    return rows.filter(row => upper(row.stato) === "FATTO").reduce((sum, row) => sum + (Number(row.totale) || 0), 0);
  }

  async function suggestedRoundNumber(ref) {
    const snap = await ref.collection("giriContabili").get();
    return snap.docs.reduce((max, doc) => Math.max(max, Number(doc.data()?.numeroGiro) || 0), 0) + 1;
  }

  async function askRealRoundNumber(ref, commessa) {
    const suggested = await suggestedRoundNumber(ref);
    while (true) {
      const value = prompt(
        `NUMERO REALE DEL GIRO\n\nCommessa: ${commessa.nome || commessa.name || ""}\n\nInserisci il numero reale del giro che stai chiudendo.\nPuò ripartire da 1 a inizio anno oppure continuare la numerazione reale.`,
        String(suggested)
      );
      if (value === null) return null;
      const numero = Number(String(value).trim());
      if (Number.isInteger(numero) && numero > 0 && numero <= 9999) return numero;
      alert("Inserisci un numero di giro valido, maggiore di 0.");
    }
  }

  async function copyRows(sourceRows, targetCollection) {
    for (let start = 0; start < sourceRows.length; start += 350) {
      const batch = db.batch();
      sourceRows.slice(start, start + 350).forEach(item => {
        const sourceId = clean(item.id) || targetCollection.doc().id;
        const target = targetCollection.doc(sourceId);
        const data = { ...(item.data || {}), sourceId, archivedAt: server() };
        batch.set(target, data);
      });
      await batch.commit();
    }
  }

  async function archiveCurrentRound(commessa, options = {}) {
    if (!commessa?.id) throw new Error("Commessa non valida.");
    if (typeof db === "undefined" || !db) throw new Error("Firestore non disponibile.");

    const ref = db.collection(collectionName()).doc(commessa.id);
    const [works, plants, evidence] = await Promise.all([
      ref.collection("lavorazioni").get(),
      ref.collection("impiantiFisici").get(),
      ref.collection("fattoVisualEvidence").get()
    ]);
    const workRows = works.docs.map(doc => ({ id: doc.id, data: doc.data() }));
    if (!workRows.length) throw new Error("La commessa non contiene lavorazioni da archiviare.");

    const doneRows = workRows.filter(item => upper(item.data?.stato) === "FATTO");
    if (!doneRows.length && options.allowWithoutDone !== true) {
      throw new Error("Non ci sono lavorazioni FATTO. Il giro non può essere chiuso come contabilità.");
    }

    const numeroGiro = Number(options.numeroGiro) || await askRealRoundNumber(ref, commessa);
    if (!numeroGiro) return { cancelled: true };
    const giroId = `giro_${String(numeroGiro).padStart(3, "0")}_${Date.now()}`;
    const giroRef = ref.collection("giriContabili").doc(giroId);
    const nowIso = new Date().toISOString();
    const accountingDueAt = plusDaysIso(1);

    await giroRef.set({
      giroId,
      numeroGiro,
      numeroGiroManuale: true,
      commessaId: commessa.id,
      commessaNome: commessa.nome || commessa.name || commessa.title || "Commessa",
      codiceCommessa: commessa.codice || commessa.code || commessa.codiceCommessa || "",
      stato: "CONTABILITA_DA_INVIARE",
      closedAt: server(),
      closedAtIso: nowIso,
      closedByUid: currentUser?.uid || "",
      closedByName: actorName(),
      totalRows: workRows.length,
      doneRows: doneRows.length,
      totalAmount: sumDone(workRows.map(item => item.data)),
      accountingDueAt,
      accountingSentAt: null,
      mapStatus: "NON_RICEVUTO",
      mapReceivedAt: null,
      mapReference: "",
      mapDocumentUrl: "",
      mapReminderDays: 7,
      archiveVersion: 1,
      immutableSnapshot: true
    });

    await copyRows(workRows, giroRef.collection("lavorazioni"));
    await copyRows(plants.docs.map(doc => ({ id: doc.id, data: doc.data() })), giroRef.collection("impiantiFisici"));
    await copyRows(evidence.docs.map(doc => ({ id: doc.id, data: doc.data() })), giroRef.collection("fattoVisualEvidence"));

    await ref.set({
      lastClosedGiroId: giroId,
      lastClosedGiroNumber: numeroGiro,
      lastClosedGiroAt: server(),
      currentGiroNumber: numeroGiro + 1,
      accountingRoundStatus: "CONTABILITA_DA_INVIARE"
    }, { merge: true });

    return { giroId, numeroGiro, doneRows: doneRows.length, totalRows: workRows.length, totalAmount: sumDone(workRows.map(item => item.data)), accountingDueAt };
  }

  function setRoundButtonsBusy(busy) {
    document.querySelectorAll('[data-commessa-round-action="close"]').forEach(button => {
      button.disabled = busy;
      if (button.dataset.roundDesktop === "1") {
        button.textContent = busy ? "Preparazione giro…" : "✓ Chiudi giro e archivia";
      } else {
        button.innerHTML = busy
          ? '<strong>Preparazione giro…</strong>'
          : '<span aria-hidden="true">✓</span><strong>Chiudi giro e archivia</strong><small>Scegli il numero reale e salva la contabilità</small>';
      }
    });
  }

  async function closeRoundAndOpenClear() {
    const commessa = currentCommessa();
    if (!commessa) return alert("Seleziona prima una commessa.");
    if (typeof canManageData === "function" && !canManageData()) return alert("Operazione riservata agli amministratori.");

    const ok = confirm(`CHIUDI GIRO CONTABILE\n\nCommessa: ${commessa.nome || commessa.name || ""}\n\nNel passaggio successivo scegli TU il numero reale del giro. Verrà poi creata una copia immutabile di lavorazioni FATTO, impianti e prove. Continuare?`);
    if (!ok) return;

    setRoundButtonsBusy(true);
    try {
      const result = await archiveCurrentRound(commessa);
      if (result?.cancelled) return;
      alert(`Giro ${result.numeroGiro} archiviato correttamente.\n${result.doneRows} lavorazioni FATTO su ${result.totalRows}.\n\nOra puoi svuotare la commessa. La contabilità resta conservata per Varga Gestionale.`);
      if (window.AccountingV2?.openClear) await window.AccountingV2.openClear(commessa);
    } catch (error) {
      console.error("Chiusura giro non riuscita", error);
      alert(`Il giro NON è stato chiuso e la commessa non verrà svuotata.\n\n${error.message || error}`);
    } finally {
      setRoundButtonsBusy(false);
    }
  }

  function installMobileButton() {
    const grid = document.querySelector("#commessa-mobile-management-home .commessa-mobile-action-grid");
    if (!grid || grid.querySelector('[data-commessa-round-action="close"]')) return;
    const button = document.createElement("button");
    button.type = "button";
    button.dataset.commessaRoundAction = "close";
    button.innerHTML = '<span aria-hidden="true">✓</span><strong>Chiudi giro e archivia</strong><small>Scegli il numero reale e salva la contabilità</small>';
    button.addEventListener("click", event => { event.preventDefault(); event.stopPropagation(); void closeRoundAndOpenClear(); });
    grid.appendChild(button);
  }

  function installDesktopButton() {
    const pricesButton = document.getElementById("open-prezziario-btn");
    const screen = document.getElementById("impianti-management-screen");
    if (!pricesButton || !screen || screen.querySelector('[data-round-desktop="1"]')) return;
    const button = document.createElement("button");
    button.type = "button";
    button.className = pricesButton.className || "btn";
    button.dataset.commessaRoundAction = "close";
    button.dataset.roundDesktop = "1";
    button.textContent = "✓ Chiudi giro e archivia";
    button.title = "Scegli il numero reale del giro e archivia la contabilità prima di svuotare la commessa";
    button.addEventListener("click", event => { event.preventDefault(); event.stopPropagation(); void closeRoundAndOpenClear(); });
    pricesButton.parentElement?.insertBefore(button, pricesButton);
  }

  function installButtons() {
    installMobileButton();
    installDesktopButton();
  }

  const observer = new MutationObserver(installButtons);
  observer.observe(document.documentElement, { childList: true, subtree: true });
  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", installButtons); else installButtons();

  window.VargaAccountingRounds = { archiveCurrentRound, closeRoundAndOpenClear };
})();
