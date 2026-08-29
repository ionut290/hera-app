(() => {
  "use strict";

  const GLOBAL = "HeraSharedStaticViews";
  const COLLECTION = "sharedStaticViews";
  const CACHE_PREFIX = "hera-shared-static-view:";
  const MAX_PAYLOAD_BYTES = 700000;
  const CALENDAR_SCHEMA_VERSION = 2;
  const subscriptions = new Map();
  const memory = new Map();
  const stats = {
    reads: 0,
    cacheHits: 0,
    snapshotsReceived: 0,
    publishes: 0,
    publishSkippedUnchanged: 0,
    publishErrors: 0,
    invalidCalendarCacheDropped: 0,
    invalidCalendarSnapshotsIgnored: 0
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

  function isCompleteCalendarView(value) {
    return Boolean(
      value &&
      value.schemaVersion === CALENDAR_SCHEMA_VERSION &&
      value.completeRecords === true &&
      value.payload &&
      value.payload.schemaVersion === CALENDAR_SCHEMA_VERSION &&
      value.payload.completeRecords === true &&
      Array.isArray(value.payload.reports)
    );
  }

  function dropLocal(type, key) {
    memory.delete(`${type}:${key}`);
    try { localStorage.removeItem(cacheKey(type, key)); } catch (_) {}
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
    const append = (items, sourceCollection) => {
      if (!Array.isArray(items)) return;
      items.forEach((item) => {
        if (!item || String(item.status || "").toLowerCase() === "rejected") return;
        const date = normalizeDateKey(item.date || item.data || item.giorno || "");
        if (!date.startsWith(key)) return;
        const id = String(item.id || `${sourceCollection}:${date}:${reports.length}`);
        const sourceKey = `${sourceCollection}/${id}`;
        if (seen.has(sourceKey)) return;
        seen.add(sourceKey);
        reports.push({
          ...safeClone(item),
          id,
          date,
          status: item.status || "",
          sourceCollection,
          sourceKey
        });
      });
    };

    // allHoursReports/allHoursApprovalRequests sono global lexical bindings
    // dichiarati con `let` in app.js: non sono proprietà di window.
    try { append(typeof allHoursReports !== "undefined" ? allHoursReports : [], "oreReports"); } catch (_) {}
    try { append(typeof allHoursApprovalRequests !== "undefined" ? allHoursApprovalRequests : [], "oreApprovalRequests"); } catch (_) {}

    reports.sort((a, b) => `${a.date}|${a.operatore || a.operatorName || ""}|${a.sourceKey}`
      .localeCompare(`${b.date}|${b.operatore || b.operatorName || ""}|${b.sourceKey}`, "it"));
    const existing = readLocal("calendario", key);

    return {
      month: key,
      schemaVersion: CALENDAR_SCHEMA_VERSION,
      completeRecords: true,
      reports,
      activities: Array.isArray(existing?.payload?.activities) ? existing.payload.activities : []
    };
  }

  function writeLocal(type, key, value) {
    if (type === "calendario" && !isCompleteCalendarView(value)) {
      dropLocal(type, key);
      stats.invalidCalendarCacheDropped += 1;
      return false;
    }
    memory.set(`${type}:${key}`, value);
    try { localStorage.setItem(cacheKey(type, key), JSON.stringify(value)); } catch (_) {}
    return true;
  }

  function readLocal(type, key) {
    const id = `${type}:${key}`;
    if (memory.has(id)) {
      const value = memory.get(id);
      if (type !== "calendario" || isCompleteCalendarView(value)) return value;
      dropLocal(type, key);
      stats.invalidCalendarCacheDropped += 1;
      return null;
    }
    try {
      const value = JSON.parse(localStorage.getItem(cacheKey(type, key)) || "null");
      if (value) {
        if (type === "calendario" && !isCompleteCalendarView(value)) {
          dropLocal(type, key);
          stats.invalidCalendarCacheDropped += 1;
          console.warn("[SHARED VIEWS] cache calendario legacy eliminata", { key });
          return null;
        }
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

    const calendarMetadata = type === "calendario"
      ? {
          schemaVersion: CALENDAR_SCHEMA_VERSION,
          completeRecords: clean.schemaVersion === CALENDAR_SCHEMA_VERSION && clean.completeRecords === true
        }
      : {};
    const value = {
      type,
      key,
      version: Number(previous?.version || 0) + 1,
      ...calendarMetadata,
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
    const existing = subscriptions.get(subscriptionKey);
    if (existing) {
      if (typeof callback === "function") existing.callbacks.add(callback);
      const cachedExisting = readLocal(type, key);
      if (cachedExisting) callback?.(cachedExisting, { source: "memory" });
      return () => {
        existing.callbacks.delete(callback);
        if (existing.callbacks.size) return;
        existing.unsubscribeFirestore?.();
        subscriptions.delete(subscriptionKey);
      };
    }

    const cached = readLocal(type, key);
    if (cached) callback?.(cached, { source: "local" });
    const callbacks = new Set(typeof callback === "function" ? [callback] : []);
    stats.reads += 1;
    console.debug("[SHARED VIEWS] listener", { type, key, reads: stats.reads });
    const unsubscribeFirestore = firestore.collection(COLLECTION).doc(docId(type, key)).onSnapshot((snapshot) => {
      if (!snapshot.exists) return;
      const value = { id: snapshot.id, ...(snapshot.data() || {}) };
      if (type === "calendario" && !isCompleteCalendarView(value)) {
        dropLocal(type, key);
        stats.invalidCalendarSnapshotsIgnored += 1;
        console.warn("[SHARED VIEWS] snapshot calendario incompleto ignorato; attendo la Cloud Function", {
          key,
          schemaVersion: value.schemaVersion,
          completeRecords: value.completeRecords
        });
        return;
      }
      writeLocal(type, key, value);
      stats.snapshotsReceived += 1;
      const metadata = { source: snapshot.metadata?.fromCache ? "firestore-cache" : "firestore" };
      callbacks.forEach((handler) => handler(value, metadata));
      window.dispatchEvent(new CustomEvent("hera-shared-static-view-updated", { detail: { type, key, value } }));
    }, (error) => {
      console.warn(`Vista condivisa ${type}/${key} non disponibile`, error);
    });

    const unsubscribe = () => {
      callbacks.delete(callback);
      if (callbacks.size) return;
      unsubscribeFirestore?.();
      subscriptions.delete(subscriptionKey);
    };
    subscriptions.set(subscriptionKey, { unsubscribe, unsubscribeFirestore, callbacks });
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
    const today = todayKey();
    const tomorrowDate = new Date(`${today}T12:00:00`);
    tomorrowDate.setDate(tomorrowDate.getDate() + 1);
    const tomorrow = normalizeDateKey(tomorrowDate.toISOString());
    const selected = normalizeDateKey(typeof window.getActiveSquadreDateKey === "function" ? window.getActiveSquadreDateKey() : today);
    let nextWorkday = "";
    try {
      nextWorkday = window.getNextWorkdayCandidateDateKeys?.(today)?.[0] || "";
    } catch (_) {}
    [...new Set([selected, today, tomorrow, nextWorkday].filter(Boolean))]
      .forEach((date) => subscribe("squadre", date));
    subscribe("calendario", monthKey(selected));
  }

  window.addEventListener("hera-squadre-saved", (event) => {
    api.publishSquadre(event?.detail?.date).catch((error) => console.warn("Vista squadre non pubblicata", error));
  });

  // Il calendario è aggiornato dalla Cloud Function transazionale su ogni write
  // a oreReports/oreApprovalRequests. Il browser non deve sovrascrivere quella
  // vista con uno snapshot locale potenzialmente non ancora aggiornato.
  window.addEventListener("hera-hours-saved", () => {
    console.debug("[SHARED VIEWS] aggiornamento calendario affidato alla Cloud Function");
  });

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
