#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

class FakeStorage {
  constructor(limit, entries = []) {
    this.limit = limit;
    this.values = new Map(entries);
  }

  get length() { return this.values.size; }
  key(index) { return Array.from(this.values.keys())[index] || null; }
  getItem(key) { return this.values.has(String(key)) ? this.values.get(String(key)) : null; }
  removeItem(key) { this.values.delete(String(key)); }
  setItem(key, value) {
    const normalizedKey = String(key);
    const normalizedValue = String(value);
    const next = new Map(this.values);
    next.set(normalizedKey, normalizedValue);
    const bytes = Array.from(next.entries()).reduce((sum, [itemKey, itemValue]) => sum + itemKey.length + itemValue.length, 0);
    if (bytes > this.limit) {
      const error = new Error("Storage quota exceeded");
      error.name = "QuotaExceededError";
      throw error;
    }
    this.values = next;
  }
}

const protectedKey = "heraOfflineMutationQueue";
const diagnosticKey = "varga_fs_diag_v4_2026-08-29";
const localStorage = new FakeStorage(760, [
  [protectedKey, "P".repeat(260)],
  [diagnosticKey, "D".repeat(260)],
  ["heraCommesseCache", "C".repeat(140)]
]);

global.window = { Storage: FakeStorage, localStorage };
require(path.resolve(__dirname, "..", "storage-quota-guard.js"));

const firestoreKey = "firestore_clients_firestore/[DEFAULT]/hera-app-6cd2b/client";
localStorage.setItem(firestoreKey, "F".repeat(180));

assert.equal(localStorage.getItem(protectedKey), "P".repeat(260), "La coda offline deve essere preservata.");
assert.equal(localStorage.getItem(diagnosticKey), null, "I diagnostici ricostruibili devono essere rimossi per primi.");
assert.equal(localStorage.getItem(firestoreKey), "F".repeat(180), "La scrittura interna Firestore deve riuscire dopo il recupero quota.");

const state = window.HeraStorageQuotaGuard.getState();
assert.equal(state.quotaRecoveries, 1);
assert.ok(state.disposableKeysRemoved >= 1);
assert.ok(state.disposableBytesRemoved > 0);

console.log("✅ Quota localStorage recuperata eliminando soltanto cache ricostruibili.");
console.log("✅ Login, code offline e dati operativi protetti restano conservati.");
