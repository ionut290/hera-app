(() => {
  "use strict";

  const GLOBAL = "HeraSharedStaticViews";
  const COLLECTION = "sharedStaticViews";
  const CACHE_PREFIX = "hera-shared-static-view:";
  const MAX_PAYLOAD_BYTES = 700000;
  const subscriptions = new Map();
  const memory = new Map();
  const stats = {
    reads: 0,
    cacheHits: 0,
    snapshotsReceived: 0,
    publishes: 0,
    publishSkippedUnchanged: 0,
    publishErrors: 0
  };

  if (window[GLOBAL]?.installed) return;

  const pad = (value) => String(value).padStart(2, "0");
  const todayKey = () => {
    const now = new Date();
    return `${now.getFullYear()}-${pad(now.getMonth() + 1)}-${pad(now.getDate())}`;
  };
  const monthKey = (value = todayKey()) => String(value || "").slice(0, 7);
  const normalizeDateKey = (value) => {
    const text = String(value || "").trim();
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return match ? `${match[1]}-${match[2]}-${match[3]}` : todayKey();
  };
  const docId = (type, key) => `${type}__${String(key).replace(/[^0-9A-Za-z_-]/g, "-")}`;
  const cacheKey = (type, key) => `${CACHE_PREFIX}${docId(type, key)}`;
  const stableStringify = (value) => JSON.stringify(value, Object.keys(value || {}).sort());

  function db() {
    return window.firebase?.firestore?.() || null;
  }

  function canPublish() {
    try {
      return typeof window.canManageData === "function" && window.canManageData();
    } catch (_) {
      return false;
    }
  }

  function safeClone(value) {
    return JSON.parse(JSON.stringify(value, (_key, item) => {
      if (item?.toDate instanceof Function) return item.toDate().toISOString();
      if (item instanceof Map) return Object.fromEntries(item);
      return item;
    }));
  }

  function trimSquadraRow(row) {
    if (!row || typeof row !== "object") return null;
    const result = {};
    ["id", "nome", "name", "commessa", "commessaId", "impianto", "impiantoId", "operatori", "personale", "mezzi", "mezzo", "note", "stato", "oraInizio", "oraFine"].forEach((key) => {
      if (row[key] != null && row[key] !== "") result[key] = safeClone(row[key]);
    });
    return Object.keys(result).length ? result : safeClone(row);
  }

  function collectSquadre(date = normalizeDateKey()) {
    const key = normalizeDateKey(date);
    let source = null;
    try {
      if (window.squadreHistoryByDate instanceof Map) source = window.squadreHistoryByDate.get(key);
    } catch (_) {}
    if (!source && typeof window.getActiveSquadreDateKey === "function") {
      try {
        const activeKey = normalizeDateKey(window.getActiveSquadreDateKey());
        if (activeKey === key && window.squadreHistoryByDate instanceof Map) source = window.squadreHistoryByDate.get(activeKey);
      } catch (_) {}
    }

    const groups = [];
    const append = (value, commessaId = "") => {
      const rows = Array.isArray(value?.squadre) ? value.squadre : (Array.isArray(value) ? value : []);
      rows.forEach((row) => {
        const compact = trimSquadraRow(row);
        if (compact) groups.push(commessaId && !compact.commessaId ? { ...compact, commessaId } : compact);
      });
    };
    if (source instanceof Map) source.forEach((value, keyValue) => append(value, String(keyValue || "")));
    else if (source && typeof source === "object") Object.entries(source).forEach(([keyValue, value]) => append(value, keyValue));
    else append(source);

    return { date: key, squadre: groups };
  }

  function collectCalendar(month = monthKey()) {
    const key = monthKey(month);
    const reports = [];
    const seen = new Set();
    const append = (items) => {
      if (!Array.isArray(items)) return;
      items.forEach((item) => {
        if (!item || String(item.status || "").toLowerCase() === "rejected") return;
        const date = normalizeDateKey(item.date || item.data || item.giorno || "");
        if (!date.startsWith(key)) return;
        const id = String(item.id || `${date}:${reports.length}`);
        if (seen.has(id)) return;
        seen.add(id);
        reports.push({
          id,
          date,
          status: item.status || "",
          commessaId: item.commessaId || item.commessa || "",
          entries: safeClone(Array.isArray(item.entries) ? item.entries : [])
        });
      });
    };
    try { append(window.allHoursReports); } catch (_) {}
    try { append(window.allHoursApprovalRequests); } catch (_) {}
    return { month: key, reports };
  }

  function writeLocal(type, key, value) {
    memory.set(`${type}:${key}`, value);
    try { localStorage.setItem(cacheKey(type, key), JSON.stringify(value)); } catch (_) {}
  }

  function readLocal(type, key) {
    const id = `${type}:${key}`;
    if (memory.has(id)) return memory.get(id);
    try {
      const value = JSON.parse(localStorage.getItem(cacheKey(type, key)) || "null");
      if (value) {
        memory.set(id, value);
        stats.cacheHits += 1;
        return value;
      }
    } catch (_) {}
    return null;
  }

  async function publish(type, key, payload) {
    if (!canPublish()) throw new Error("Permessi amministratore richiesti per pubblicare la vista condivisa.");
    const firestore = db();
    if (!firestore) throw new Error("Firestore non disponibile.");

    const clean = safeClone(payload);
    const serialized = JSON.stringify(clean);
    if (new Blob([serialized]).size > MAX_PAYLOAD_BYTES) throw new Error("Vista condivisa troppo grande; pubblicazione bloccata.");

    const id = docId(type, key);
    const ref = firestore.collection(COLLECTION).doc(id);
    const previous = readLocal(type, key);
    const contentHash = stableStringify(clean);
    if (previous?.contentHash === contentHash) {
      stats.publishSkippedUnchanged += 1;
      return previous;
    }

    const value = {
      type,
      key,
      version: Number(previous?.version || 0) + 1,
      updatedAt: window.firebase.firestore.FieldValue.serverTimestamp(),
      updatedAtClient: new Date().toISOString(),
      updatedBy: window.firebase?.auth?.()?.currentUser?.email || "",
      contentHash,
      payload: clean
    };

    try {
      await ref.set(value, { merge: false });
      const localValue = { ...value, updatedAt: null };
      writeLocal(type, key, localValue);
      stats.publishes += 1;
      window.dispatchEvent(new CustomEvent("hera-shared-static-view-published", { detail: { type, key, value: localValue } }));
      return localValue;
    } catch (error) {
      stats.publishErrors += 1;
      throw error;
    }
  }

  function subscribe(type, key, callback) {
    const firestore = db();
    if (!firestore) return () => {};
    const subscriptionKey = `${type}:${key}`;
    if (subscriptions.has(subscriptionKey)) return subscriptions.get(subscriptionKey).unsubscribe;

    const cached = readLocal(type, key);
    if (cached) callback?.(cached, { source: "local" });
    stats.reads += 1;
    const unsubscribeFirestore = firestore.collection(COLLECTION).doc(docId(type, key)).onSnapshot((snapshot) => {
      if (!snapshot.exists) return;
      const value = { id: snapshot.id, ...(snapshot.data() || {}) };
      writeLocal(type, key, value);
      stats.snapshotsReceived += 1;
      callback?.(value, { source: snapshot.metadata?.fromCache ? "firestore-cache" : "firestore" });
      window.dispatchEvent(new CustomEvent("hera-shared-static-view-updated", { detail: { type, key, value } }));
    }, (error) => {
      console.warn(`Vista condivisa ${type}/${key} non disponibile`, error);
    });

    const unsubscribe = () => {
      unsubscribeFirestore?.();
      subscriptions.delete(subscriptionKey);
    };
    subscriptions.set(subscriptionKey, { unsubscribe });
    return unsubscribe;
  }

  const api = {
    installed: true,
    collection: COLLECTION,
    getState: () => ({ stats: { ...stats }, subscriptions: subscriptions.size, cachedViews: memory.size }),
    getCached: readLocal,
    subscribe,
    collectSquadre,
    collectCalendar,
    publishSquadre: (date) => {
      const payload = collectSquadre(date);
      return publish("squadre", payload.date, payload);
    },
    publishCalendar: (month) => {
      const payload = collectCalendar(month);
      return publish("calendario", payload.month, payload);
    }
  };

  window[GLOBAL] = api;

  function start() {
    const date = normalizeDateKey(typeof window.getActiveSquadreDateKey === "function" ? window.getActiveSquadreDateKey() : todayKey());
    subscribe("squadre", date);
    subscribe("calendario", monthKey(date));
  }

  window.addEventListener("hera-squadre-saved", (event) => {
    api.publishSquadre(event?.detail?.date).catch((error) => console.warn("Vista squadre non pubblicata", error));
  });
  window.addEventListener("hera-hours-saved", (event) => {
    api.publishCalendar(event?.detail?.month || event?.detail?.date).catch((error) => console.warn("Vista calendario non pubblicata", error));
  });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
