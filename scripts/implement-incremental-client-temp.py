from pathlib import Path
from textwrap import dedent

path = Path("app.js")
source = path.read_text(encoding="utf-8")
start = source.index("function subscribeImpianti() {")
end = source.index("\nfunction stopImpiantiSubscription()", start)
helper_start = source.rfind("function getImpiantiIncrementalStateKey(commessaId) {", 0, start)
if helper_start != -1:
    start = helper_start

replacement = dedent(r'''
function getImpiantiIncrementalStateKey(commessaId) {
  return `heraImpiantiIncrementalStateV1:${String(commessaId || "").trim()}`;
}

function readImpiantiIncrementalState(commessaId) {
  try {
    const parsed = JSON.parse(localStorage.getItem(getImpiantiIncrementalStateKey(commessaId)) || "null");
    return parsed && parsed.seeded === true && parsed.lastChangedAtMs > 0 ? parsed : null;
  } catch (_) {
    return null;
  }
}

function saveImpiantiIncrementalState(commessaId, lastChangedAtMs) {
  if (!commessaId || !Number.isFinite(Number(lastChangedAtMs)) || Number(lastChangedAtMs) <= 0) return;
  try {
    localStorage.setItem(getImpiantiIncrementalStateKey(commessaId), JSON.stringify({
      seeded: true,
      lastChangedAtMs: Number(lastChangedAtMs),
      updatedAt: Date.now()
    }));
  } catch (_) {}
}

function renderImpiantiAfterRemoteSync(rawImpianti, previousDoneSignatureRef) {
  currentImpianti = applyPendingActionsToImpianti(combineImpiantiForView(rawImpianti), selectedCommessaId);
  refreshImpiantoWhatsAppTemplateCache(currentImpianti);
  impiantiByCommessaId.set(selectedCommessaId, currentImpianti);
  renderSquadre();
  renderHeaderActivitySummary();
  updateCommessaDashboard();
  renderImpianti();
  const commessaRoute = parseCommessaHash();
  if (commessaRoute.id === selectedCommessaId && commessaRoute.impianto) {
    const routedImpianto = findCurrentImpiantoByKey(commessaRoute.impianto);
    if (routedImpianto && highlightedImpiantoKey !== buildImpiantoKey(routedImpianto)) {
      window.requestAnimationFrame(() => focusSharedImpiantoFromRoute(commessaRoute.impianto));
    }
  }
  renderMap();
  runWhazzupPendingDoneSafetyCheck();
  preloadCommessaWeatherForVisibleImpianti();
  evaluateImpiantoProximityAlerts();
  autoCompletePassedSnowRoads().catch((error) => console.warn("Completamento automatico vie neve non riuscito:", error));
  if (!currentUserPos) fetchWeather();

  const currentDoneSignature = rawImpianti
    .filter((impianto) => Boolean(impianto.done))
    .map((impianto) => `${impianto.id}__${firestoreDateToMillis(impianto.doneAt)}`)
    .sort()
    .join("|");
  const previousDoneSignature = previousDoneSignatureRef.value;
  const doneStateChanged = previousDoneSignature !== null && currentDoneSignature !== previousDoneSignature;
  if (doneStateChanged && !hasRecentLocalSheetMutation(selectedCommessaId)) {
    scheduleCommessaSheetSync(selectedCommessaId, selectedCommessaName, 700);
  }
  previousDoneSignatureRef.value = currentDoneSignature;
}

function subscribeImpianti() {
  if (!selectedCommessaId) return;
  subscribeFattoVisualEvidence(selectedCommessaId);
  const requestedCommessaId = String(selectedCommessaId || "").trim();
  const previousDoneSignatureRef = { value: null };
  const impiantiRef = db.collection("commesse").doc(requestedCommessaId).collection("impianti");
  const changeIndexRef = db.collection("commesse").doc(requestedCommessaId).collection("impiantoChangeIndex");
  const cachedImpianti = Array.isArray(impiantiByCommessaId.get(requestedCommessaId))
    ? impiantiByCommessaId.get(requestedCommessaId)
    : [];
  const incrementalState = readImpiantiIncrementalState(requestedCommessaId);
  let markerSeedCaptured = false;

  const applyRaw = (rawImpianti) => {
    if (requestedCommessaId !== selectedCommessaId) return;
    renderImpiantiAfterRemoteSync(rawImpianti, previousDoneSignatureRef);
  };

  const captureLatestServerMarker = async () => {
    if (markerSeedCaptured) return;
    markerSeedCaptured = true;
    try {
      const markerSnapshot = await changeIndexRef.orderBy("changedAt", "desc").limit(1).get();
      const latestMarker = markerSnapshot.docs?.[0]?.data?.() || null;
      const latestMarkerMs = latestMarker ? firestoreDateToMillis(latestMarker.changedAt) : 0;
      if (latestMarkerMs > 0) saveImpiantiIncrementalState(requestedCommessaId, latestMarkerMs);
    } catch (error) {
      markerSeedCaptured = false;
      console.debug("Indice incrementale impianti non ancora disponibile; mantengo il listener completo.", {
        commessaId: requestedCommessaId,
        error: String(error?.message || error || "")
      });
    }
  };

  const startFullListener = () => {
    console.log("Query impianti completa avviata", { commessaId: requestedCommessaId });
    return impiantiRef.onSnapshot((snapshot) => {
      const rawImpianti = snapshot.docs.map((doc) => ({ id: doc.id, ...doc.data() }));
      applyRaw(rawImpianti);
      captureLatestServerMarker();
    }, (error) => {
      console.error("Errore Firestore caricamento impianti:", error);
      if (!cachedImpianti.length && ui.impiantiLista) {
        ui.impiantiLista.innerHTML = `<p class='muted'>Errore caricamento impianti: ${escapeHTML(getFirebaseErrorMessage(error) || "Firestore non disponibile.")}</p>`;
      }
    });
  };

  if (!incrementalState || !cachedImpianti.length) {
    unsubscribeImpianti = startFullListener();
    return;
  }

  currentImpianti = cachedImpianti.slice();
  renderImpianti();
  renderMap();

  const since = firebase.firestore.Timestamp.fromMillis(Number(incrementalState.lastChangedAtMs));
  const working = new Map(cachedImpianti.map((item) => [String(item.id || ""), item]).filter(([id]) => id));
  let incrementalSyncChain = Promise.resolve();
  console.log("Sync impianti incrementale avviata", { commessaId: requestedCommessaId, cached: working.size });
  unsubscribeImpianti = changeIndexRef
    .where("changedAt", ">", since)
    .orderBy("changedAt", "asc")
    .onSnapshot((snapshot) => {
      incrementalSyncChain = incrementalSyncChain.then(async () => {
        if (requestedCommessaId !== selectedCommessaId) return;
        let maxChangedAtMs = Number(readImpiantiIncrementalState(requestedCommessaId)?.lastChangedAtMs || incrementalState.lastChangedAtMs);
        for (const change of snapshot.docChanges()) {
          if (change.type === "removed") continue;
          const marker = change.doc.data() || {};
          const impiantoId = String(marker.impiantoId || change.doc.id || "").trim();
          const changedAtMs = firestoreDateToMillis(marker.changedAt);
          if (changedAtMs > maxChangedAtMs) maxChangedAtMs = changedAtMs;
          if (!impiantoId) continue;
          if (marker.deleted === true) {
            working.delete(impiantoId);
            continue;
          }
          const doc = await impiantiRef.doc(impiantoId).get();
          if (doc.exists) working.set(doc.id, { id: doc.id, ...doc.data() });
          else working.delete(impiantoId);
        }
        saveImpiantiIncrementalState(requestedCommessaId, maxChangedAtMs);
        applyRaw(Array.from(working.values()));
      }).catch((error) => {
        console.warn("Aggiornamento incrementale impianti fallito, ripristino listener completo:", error);
        try { unsubscribeImpianti?.(); } catch (_) {}
        unsubscribeImpianti = startFullListener();
      });
    }, (error) => {
      console.warn("Sync incrementale impianti non disponibile, ripristino listener completo:", error);
      try { unsubscribeImpianti?.(); } catch (_) {}
      unsubscribeImpianti = startFullListener();
    });
}
''').lstrip()

path.write_text(source[:start] + replacement + source[end:], encoding="utf-8")
