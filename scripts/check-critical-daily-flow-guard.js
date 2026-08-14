"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.resolve(__dirname, "..");
const guardSource = fs.readFileSync(path.join(root, "critical-daily-flow-guard.js"), "utf8");
const configSource = fs.readFileSync(path.join(root, "firebase-config.js"), "utf8");
const packageSource = JSON.parse(fs.readFileSync(path.join(root, "package.json"), "utf8"));

for (const expected of [
  "heraPendingOfflineMutations",
  "server-note-newer-than-offline-edit",
  "syncInFlight",
  "stale-auth-event",
  "snow-management-context",
  "navigateToImpiantoGuarded"
]) {
  assert.ok(guardSource.includes(expected), `Protezione critica mancante: ${expected}`);
}
assert.ok(configSource.includes("critical-daily-flow-guard.js?v="), "Il guard critico non viene caricato da firebase-config.js.");
assert.ok(packageSource.scripts["check:critical-daily-flow"], "Manca il comando di test del guard critico.");
for (const protectedFlow of ["android-whazzup-photo-order", "sharePendingPhotosToWhatsApp", "markImpiantoDone"]) {
  assert.equal(guardSource.includes(protectedFlow), false, `Il guard tocca il flusso protetto: ${protectedFlow}`);
}

function deferred() {
  let resolve;
  let reject;
  const promise = new Promise((res, rej) => { resolve = res; reject = rej; });
  return { promise, resolve, reject };
}

function createStorage() {
  const values = new Map();
  return {
    getItem: (key) => values.has(String(key)) ? values.get(String(key)) : null,
    setItem: (key, value) => values.set(String(key), String(value)),
    removeItem: (key) => values.delete(String(key)),
    clear: () => values.clear()
  };
}

async function runBehaviorChecks() {
  const localStorage = createStorage();
  const sessionStorage = createStorage();
  const classNames = new Set();
  const auth = {
    currentUser: { uid: "user-a", email: "a@example.com", displayName: "Operatore A" },
    onAuthStateChanged(callback) { this.callback = callback; callback(this.currentUser); }
  };
  let serverUpdatedAtMs = 0;
  const firestore = {
    collection() {
      return {
        doc() {
          return {
            collection() {
              return {
                doc() {
                  return {
                    async get() {
                      return {
                        exists: serverUpdatedAtMs > 0,
                        data: () => ({ updatedAt: { toMillis: () => serverUpdatedAtMs } })
                      };
                    }
                  };
                }
              };
            }
          };
        }
      };
    }
  };

  let originalSyncCalls = 0;
  let syncGate = deferred();
  const navigationWrites = [];
  const navigationChats = [];
  const navigationNotifications = [];
  let navigationGate = deferred();
  let approvalGate = deferred();

  const context = {
    console,
    Promise,
    Date,
    Math,
    URL,
    WeakSet,
    Set,
    Map,
    JSON,
    localStorage,
    sessionStorage,
    location: { href: "https://example.test/", hash: "", replace() {} },
    document: {
      readyState: "complete",
      body: {
        classList: {
          contains: (name) => classNames.has(name),
          add: (name) => classNames.add(name),
          remove: (name) => classNames.delete(name)
        }
      },
      addEventListener() {}
    },
    MutationObserver: class { observe() {} },
    addEventListener() {},
    dispatchEvent() {},
    CustomEvent: class { constructor(type, options) { this.type = type; this.detail = options?.detail; } },
    setTimeout(fn, ms) { if (ms === 0) fn(); return 1; },
    clearTimeout() {},
    authStateResolved: true,
    selectedCommessaId: "commessa-a",
    selectedCommessaName: "Commessa A",
    currentUser: { uid: "user-a", email: "a@example.com", displayName: "Operatore A" },
    firebase: {
      auth: () => auth,
      firestore: () => firestore
    },
    enqueueOfflineMutation(type, payload) {
      const queue = JSON.parse(localStorage.getItem("heraPendingOfflineMutations") || "[]");
      const item = {
        id: `${type}:${queue.length + 1}`,
        type,
        payload,
        userId: context.currentUser.uid,
        createdAt: new Date().toISOString(),
        status: "pending"
      };
      queue.push(item);
      localStorage.setItem("heraPendingOfflineMutations", JSON.stringify(queue));
      return item;
    },
    async syncPendingOfflineMutations() {
      originalSyncCalls += 1;
      return syncGate.promise;
    },
    openCommessaNoteForm(note) { return note; },
    async setImpiantoNavigated(commessaId, ids, date, operatorName) {
      navigationWrites.push({ commessaId, ids, date, operatorName });
      return true;
    },
    async sendChatMessage(payload) { navigationChats.push(payload); },
    async publishGlobalNotificationEvent(type, payload) { navigationNotifications.push({ type, payload }); },
    async navigateToImpianto(impianto) {
      await navigationGate.promise;
      await context.setImpiantoNavigated(context.selectedCommessaId, [impianto.id], new Date(), context.currentUser.displayName);
      await context.sendChatMessage({
        text: "testo non protetto",
        metadata: { type: "impianto_navigate", commessaId: context.selectedCommessaId }
      });
      await context.publishGlobalNotificationEvent("impianto-navigate", { commessaId: context.selectedCommessaId });
    },
    HeraAccessApproval: { verify: () => approvalGate.promise }
  };
  context.window = context;
  vm.createContext(context);
  vm.runInContext(guardSource, context, { filename: "critical-daily-flow-guard.js" });

  const hoursPayload = { date: "2026-08-14", entries: [{ commessaId: "commessa-a", rows: [{ ore: 8 }] }] };
  const first = context.enqueueOfflineMutation("hoursReport", hoursPayload);
  const second = context.enqueueOfflineMutation("hoursReport", { ...hoursPayload });
  assert.equal(first.id, second.id, "Due tocchi identici creano ancora due operazioni offline.");
  assert.equal(JSON.parse(localStorage.getItem("heraPendingOfflineMutations")).length, 1);

  const syncOne = context.syncPendingOfflineMutations();
  const syncTwo = context.syncPendingOfflineMutations();
  assert.equal(syncOne, syncTwo, "La sincronizzazione offline non è single-flight.");
  await new Promise((resolve) => setImmediate(resolve));
  assert.equal(originalSyncCalls, 1, "La sincronizzazione offline parte più volte in parallelo.");
  syncGate.resolve(true);
  await syncOne;

  const baseMs = Date.now() - 10_000;
  context.openCommessaNoteForm({ id: "nota-1", updatedAt: { toMillis: () => baseMs } });
  const noteItem = context.enqueueOfflineMutation("commessaNote", {
    noteId: "nota-1",
    commessaId: "commessa-a",
    title: "Nota offline",
    text: "Test"
  });
  assert.equal(noteItem.guard.baseUpdatedAtMs, baseMs, "La nota offline non conserva la versione di partenza.");
  serverUpdatedAtMs = baseMs + 5000;
  syncGate = deferred();
  const conflictSync = context.syncPendingOfflineMutations();
  await Promise.resolve();
  syncGate.resolve(true);
  await conflictSync;
  const queueAfterConflict = JSON.parse(localStorage.getItem("heraPendingOfflineMutations") || "[]");
  assert.equal(queueAfterConflict.some((item) => item.id === noteItem.id), false, "Una nota obsoleta resta pronta a sovrascrivere il server.");
  assert.equal(context.HeraCriticalDailyFlowGuard.getConflicts().length, 1, "Il conflitto non viene archiviato.");

  navigationGate = deferred();
  const navigation = context.navigateToImpianto({ id: "impianto-1", denominazione: "Impianto Uno", comune: "Bologna" });
  context.selectedCommessaId = "commessa-b";
  context.selectedCommessaName = "Commessa B";
  navigationGate.resolve();
  await navigation;
  assert.equal(navigationWrites[0].commessaId, "commessa-a", "NAVIGA usa la commessa selezionata dopo l'attesa.");
  assert.equal(navigationWrites[0].operatorName, "Operatore A", "NAVIGA perde l'operatore iniziale.");
  assert.equal(navigationChats[0].metadata.commessaId, "commessa-a");
  assert.match(navigationChats[0].text, /^🧭 Operatore A naviga verso Impianto Uno\./);
  assert.equal(navigationNotifications[0].payload.commessaId, "commessa-a");

  context.selectedCommessaId = "commessa-a";
  context.selectedCommessaName = "Commessa A";
  context.currentUser = { uid: "user-a", email: "a@example.com", displayName: "Operatore A" };
  auth.currentUser = context.currentUser;
  navigationGate = deferred();
  const staleNavigation = context.navigateToImpianto({ id: "impianto-2", denominazione: "Impianto Due", comune: "Modena" });
  const writesBeforeSwitch = navigationWrites.length;
  context.currentUser = { uid: "user-b", email: "b@example.com", displayName: "Operatore B" };
  auth.currentUser = context.currentUser;
  navigationGate.resolve();
  await staleNavigation;
  assert.equal(navigationWrites.length, writesBeforeSwitch, "NAVIGA salva dopo un cambio account.");

  context.currentUser = { uid: "user-a", email: "a@example.com", displayName: "Operatore A" };
  auth.currentUser = context.currentUser;
  approvalGate = deferred();
  const approval = context.HeraAccessApproval.verify(context.currentUser);
  context.currentUser = { uid: "user-b", email: "b@example.com", displayName: "Operatore B" };
  auth.currentUser = context.currentUser;
  approvalGate.resolve({ allowed: true, status: "attivo" });
  const staleApproval = await approval;
  assert.equal(staleApproval.allowed, false, "Una verifica autorizzazione vecchia abilita il nuovo account.");
  assert.equal(staleApproval.stale, true);

  const guardState = context.HeraCriticalDailyFlowGuard.getState();
  assert.ok(guardState.deduplicated >= 1);
  assert.ok(guardState.conflicts >= 1);
  assert.ok(guardState.staleActionsBlocked >= 1);
}

runBehaviorChecks().then(() => {
  console.log("Critical daily flow guard check passed.");
}).catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
