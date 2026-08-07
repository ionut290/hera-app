(() => {
  "use strict";

  const INSTALL_FLAG = "__heraFattoEmbeddedOptimizerInstalled";
  const DISABLE_KEY = "heraFattoEmbeddedDisabled";
  const MIGRATION_FIELD = "fattoEmbeddedMigrationV1";
  const EMBEDDED_FIELD = "fattoVisualEvidence";
  const SCHEMA_VERSION = 1;
  const WAIT_INTERVAL_MS = 100;
  const WAIT_MAX_ATTEMPTS = 80;

  if (window[INSTALL_FLAG]) return;
  window[INSTALL_FLAG] = true;

  function isDisabled() {
    try {
      return localStorage.getItem(DISABLE_KEY) === "1";
    } catch (_) {
      return false;
    }
  }

  function asMillis(value) {
    if (!value) return 0;
    if (typeof value.toMillis === "function") return Number(value.toMillis()) || 0;
    if (typeof value.toDate === "function") return value.toDate().getTime() || 0;
    const numeric = Number(value);
    if (Number.isFinite(numeric) && numeric > 0) return numeric;
    const parsed = Date.parse(String(value));
    return Number.isFinite(parsed) ? parsed : 0;
  }

  function normalizeEvidence(raw, fallbackId = "") {
    if (!raw || typeof raw !== "object") return null;
    const marked = raw.marked !== false;
    const markedAtMs = asMillis(raw.markedAtMs) || asMillis(raw.markedAt) || asMillis(raw.updatedAt);
    if (!marked && !markedAtMs) return null;
    return {
      ...raw,
      impiantoId: normalizeImpiantoId(raw.impiantoId || fallbackId),
      marked: true,
      markedAtMs,
      schemaVersion: Number(raw.schemaVersion) || SCHEMA_VERSION
    };
  }

  function pickNewestEvidence(primary, fallback) {
    if (!primary) return fallback || null;
    if (!fallback) return primary;
    return asMillis(primary.markedAtMs || primary.markedAt || primary.updatedAt)
      >= asMillis(fallback.markedAtMs || fallback.markedAt || fallback.updatedAt)
      ? primary
      : fallback;
  }

  function getImpiantoFromCurrentState(impiantoId) {
    const normalized = normalizeImpiantoId(impiantoId);
    if (!normalized || !Array.isArray(impianti)) return null;
    return impianti.find((item) => normalizeImpiantoId(item && item.id) === normalized) || null;
  }

  function waitForImpianti(commessaId) {
    return new Promise((resolve) => {
      let attempts = 0;
      const check = () => {
        if (String(fattoVisualEvidenceCommessaId || "") !== String(commessaId || "")) {
          resolve([]);
          return;
        }
        if (Array.isArray(impianti) && String(impiantiCommessaId || "") === String(commessaId || "")) {
          resolve(impianti);
          return;
        }
        attempts += 1;
        if (attempts >= WAIT_MAX_ATTEMPTS) {
          resolve(Array.isArray(impianti) ? impianti : []);
          return;
        }
        window.setTimeout(check, WAIT_INTERVAL_MS);
      };
      check();
    });
  }

  async function backfillLegacyEvidence(commessaId, legacyEntries) {
    if (!window.db || !legacyEntries.length) return false;
    const currentImpianti = await waitForImpianti(commessaId);
    if (String(fattoVisualEvidenceCommessaId || "") !== String(commessaId || "")) return false;

    const byId = new Map(
      currentImpianti.map((item) => [normalizeImpiantoId(item && item.id), item])
    );
    const pending = legacyEntries.filter(({ impiantoId, evidence }) => {
      const item = byId.get(impiantoId);
      if (!item) return false;
      const embedded = normalizeEvidence(item[EMBEDDED_FIELD], impiantoId);
      return !embedded || asMillis(evidence.markedAtMs) > asMillis(embedded.markedAtMs);
    });

    try {
      for (let offset = 0; offset < pending.length; offset += 400) {
        const batch = window.db.batch();
        pending.slice(offset, offset + 400).forEach(({ impiantoId, evidence }) => {
          const ref = window.db.collection("commesse").doc(commessaId).collection("impianti").doc(impiantoId);
          batch.set(ref, {
            [EMBEDDED_FIELD]: {
              ...evidence,
              schemaVersion: SCHEMA_VERSION,
              migratedFromLegacy: true,
              migratedAt: firebase.firestore.FieldValue.serverTimestamp()
            }
          }, { merge: true });
        });
        await batch.commit();
      }

      await window.db.collection("commesse").doc(commessaId).set({
        [MIGRATION_FIELD]: {
          complete: true,
          schemaVersion: SCHEMA_VERSION,
          migratedDocuments: legacyEntries.length,
          migratedAt: firebase.firestore.FieldValue.serverTimestamp()
        }
      }, { merge: true });
      console.info(`[FATTO embedded] Migrazione completata per ${commessaId}: ${legacyEntries.length} evidenze.`);
      return true;
    } catch (error) {
      console.warn("[FATTO embedded] Migrazione non completata; il fallback legacy resta disponibile.", error);
      return false;
    }
  }

  function buildEvidencePayload(context) {
    const currentUser = window.auth && window.auth.currentUser;
    const now = new Date();
    return {
      impiantoId: normalizeImpiantoId(context && context.impiantoId),
      commessaId: String(context && context.commessaId || commessaSelezionata || "").trim(),
      codice: String(context && context.codice || ""),
      indirizzo: String(context && context.indirizzo || ""),
      paese: String(context && context.paese || ""),
      marked: true,
      markedAt: now.toISOString(),
      markedAtMs: now.getTime(),
      markedByUid: String(currentUser && currentUser.uid || ""),
      markedByEmail: String(currentUser && currentUser.email || ""),
      markedByName: getFattoUserLabel(),
      source: String(context && context.source || ""),
      schemaVersion: SCHEMA_VERSION
    };
  }

  function installOverrides() {
    if (isDisabled()) {
      console.warn(`[FATTO embedded] Ottimizzazione disattivata tramite ${DISABLE_KEY}=1; uso sistema legacy.`);
      return true;
    }
    if (
      typeof subscribeFattoVisualEvidence !== "function"
      || typeof recordFattoVisualEvidence !== "function"
      || typeof deleteFattoVisualEvidence !== "function"
      || typeof getFattoVisualEvidence !== "function"
      || typeof clearFattoVisualEvidenceSubscription !== "function"
      || typeof refreshFattoVisualButtons !== "function"
      || !window.db
    ) {
      return false;
    }

    const legacyRecord = recordFattoVisualEvidence;
    const legacyDelete = deleteFattoVisualEvidence;

    getFattoVisualEvidence = function getFattoVisualEvidenceEmbedded(impiantoId) {
      const normalized = normalizeImpiantoId(impiantoId);
      if (!normalized) return null;
      const item = getImpiantoFromCurrentState(normalized);
      const embedded = normalizeEvidence(item && item[EMBEDDED_FIELD], normalized);
      const legacy = normalizeEvidence(fattoVisualEvidenceByImpianto.get(normalized), normalized);
      return pickNewestEvidence(embedded, legacy);
    };

    subscribeFattoVisualEvidence = function subscribeFattoVisualEvidenceEmbedded(commessaId) {
      const targetId = String(commessaId || "").trim();
      if (!window.db || !targetId) {
        clearFattoVisualEvidenceSubscription();
        return;
      }
      if (fattoVisualEvidenceCommessaId === targetId && fattoVisualEvidenceUnsubscribe) return;

      clearFattoVisualEvidenceSubscription();
      fattoVisualEvidenceCommessaId = targetId;
      let cancelled = false;
      fattoVisualEvidenceUnsubscribe = () => {
        cancelled = true;
      };

      (async () => {
        try {
          const commessaRef = window.db.collection("commesse").doc(targetId);
          const commessaSnapshot = await commessaRef.get();
          if (cancelled || fattoVisualEvidenceCommessaId !== targetId) return;
          const migration = commessaSnapshot.exists
            ? (commessaSnapshot.data() || {})[MIGRATION_FIELD]
            : null;
          if (migration && migration.complete === true && Number(migration.schemaVersion) >= SCHEMA_VERSION) {
            refreshFattoVisualButtons();
            return;
          }

          const snapshot = await commessaRef.collection("fattoVisualEvidence").get();
          if (cancelled || fattoVisualEvidenceCommessaId !== targetId) return;
          const legacyEntries = [];
          snapshot.forEach((doc) => {
            const evidence = normalizeEvidence({ id: doc.id, ...(doc.data() || {}) }, doc.id);
            const impiantoId = normalizeImpiantoId(evidence && evidence.impiantoId || doc.id);
            if (!evidence || !impiantoId) return;
            fattoVisualEvidenceByImpianto.set(impiantoId, evidence);
            legacyEntries.push({ impiantoId, evidence });
          });
          refreshFattoVisualButtons();
          if (activeTab === "squadre") renderFatto();
          else if (activeTab === "mappa") renderMappa();
          if (!cancelled) await backfillLegacyEvidence(targetId, legacyEntries);
        } catch (error) {
          console.warn("[FATTO embedded] Lettura fallback non riuscita; nessun dato legacy viene cancellato.", error);
        }
      })();
    };

    recordFattoVisualEvidence = async function recordFattoVisualEvidenceEmbedded(context) {
      const commessaId = String(context && context.commessaId || commessaSelezionata || "").trim();
      const impiantoId = normalizeImpiantoId(context && context.impiantoId);
      if (!window.db || !commessaId || !impiantoId) {
        return legacyRecord(context);
      }

      const payload = buildEvidencePayload({ ...context, commessaId, impiantoId });

      // Il salvataggio legacy resta il primo e obbligatorio: se il nuovo campo
      // non è autorizzato dalle regole, il comportamento usato in produzione
      // continua comunque a funzionare.
      await legacyRecord(context);
      fattoVisualEvidenceByImpianto.set(impiantoId, payload);
      refreshFattoVisualButtons();

      try {
        await window.db.collection("commesse").doc(commessaId).collection("impianti").doc(impiantoId).set({
          [EMBEDDED_FIELD]: {
            ...payload,
            updatedAt: firebase.firestore.FieldValue.serverTimestamp()
          }
        }, { merge: true });
      } catch (error) {
        console.warn("[FATTO embedded] Campo incorporato non salvato; il dato legacy è comunque al sicuro.", error);
      }
    };

    deleteFattoVisualEvidence = async function deleteFattoVisualEvidenceEmbedded(context) {
      const commessaId = String(context && context.commessaId || commessaSelezionata || "").trim();
      const impiantoId = normalizeImpiantoId(context && context.impiantoId);
      await legacyDelete(context);
      if (impiantoId) {
        fattoVisualEvidenceByImpianto.delete(impiantoId);
        refreshFattoVisualButtons();
      }
      if (!window.db || !commessaId || !impiantoId) return;
      try {
        await window.db.collection("commesse").doc(commessaId).collection("impianti").doc(impiantoId).update({
          [EMBEDDED_FIELD]: firebase.firestore.FieldValue.delete()
        });
      } catch (error) {
        console.warn("[FATTO embedded] Campo incorporato non rimosso; la rimozione legacy è stata completata.", error);
      }
    };

    window.HeraFattoEmbeddedOptimizer = Object.freeze({
      installed: true,
      schemaVersion: SCHEMA_VERSION,
      rollbackKey: DISABLE_KEY,
      disableAndReload() {
        try { localStorage.setItem(DISABLE_KEY, "1"); } catch (_) {}
        window.location.reload();
      },
      enableAndReload() {
        try { localStorage.removeItem(DISABLE_KEY); } catch (_) {}
        window.location.reload();
      }
    });

    console.info("[FATTO embedded] Ottimizzazione installata: dual-write, fallback una tantum e rollback immediato.");
    return true;
  }

  let attempts = 0;
  const installWhenReady = () => {
    if (installOverrides()) return;
    attempts += 1;
    if (attempts < 300) window.setTimeout(installWhenReady, 25);
    else console.error("[FATTO embedded] Installazione non riuscita: app.js non pronto. Sistema legacy invariato.");
  };

  installWhenReady();
})();
