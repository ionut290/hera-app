#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const authFixPath = path.resolve(__dirname, "..", "auth-login-fix.js");
const authFix = fs.readFileSync(authFixPath, "utf8");

assert.match(authFix, /profileBootstrapInFlight\s*=\s*new Map\(\)/,
  "Il bootstrap profilo deve condividere la stessa Promise tra listener concorrenti");
assert.match(authFix, /INTERNAL ASSERTION FAILED:\\s\*Unexpected state/,
  "Deve riconoscere l'errore interno Unexpected state di Firestore");
assert.match(authFix, /get\(\{\s*source:\s*"server"\s*\}\)/,
  "Il singolo tentativo di recupero deve forzare una lettura server");
assert.match(authFix, /platform-profile-sync-deferred/,
  "Il recupero rinviato deve essere osservabile senza bloccare il login");
assert.doesNotMatch(authFix, /clearPersistence\s*\(/,
  "Il recupero auth non deve cancellare la persistenza Firestore");
assert.doesNotMatch(authFix, /location\.reload\s*\(/,
  "Il recupero auth non deve creare ricaricamenti automatici o loop");

function createClassList() {
  const values = new Set();
  return {
    add(value) { values.add(value); },
    remove(value) { values.delete(value); },
    contains(value) { return values.has(value); },
    toggle(value, force) {
      if (force === true) values.add(value);
      else if (force === false) values.delete(value);
      else if (values.has(value)) values.delete(value);
      else values.add(value);
    }
  };
}

function createStyle() {
  const values = new Map();
  return {
    setProperty(name, value) { values.set(name, value); },
    removeProperty(name) { values.delete(name); },
    getPropertyValue(name) { return values.get(name) || ""; }
  };
}

function createHarness(getImplementation) {
  const nativeAuthCallbacks = [];
  const deferredEvents = [];
  const warnings = [];
  const errors = [];
  let getCalls = 0;
  let setCalls = 0;
  const getOptions = [];

  const currentRef = {
    get(options) {
      getCalls += 1;
      getOptions.push(options || null);
      return getImplementation({ call: getCalls, options });
    },
    async set() {
      setCalls += 1;
    }
  };

  const database = {
    collection(name) {
      assert.equal(name, "platformUsers");
      return {
        doc() {
          return currentRef;
        },
        where() {
          return {
            limit() {
              return {
                async get() {
                  return { empty: true, docs: [] };
                }
              };
            }
          };
        }
      };
    }
  };

  const authInstance = {
    currentUser: null,
    async setPersistence() {},
    onIdTokenChanged() {
      return () => {};
    },
    onAuthStateChanged(nextOrObserver) {
      const next = typeof nextOrObserver === "function"
        ? nextOrObserver
        : nextOrObserver && nextOrObserver.next;
      nativeAuthCallbacks.push(next);
      return () => {};
    },
    async signInWithPopup() {}
  };

  function auth() {
    return authInstance;
  }
  auth.Auth = { Persistence: { LOCAL: "local" } };
  auth.GoogleAuthProvider = class GoogleAuthProvider {
    addScope() {}
    setCustomParameters() {}
    static credential() {
      return {};
    }
  };

  function firestore() {
    return database;
  }
  firestore.FieldValue = {
    serverTimestamp() {
      return { __serverTimestamp: true };
    }
  };

  const firebase = {
    apps: [{}],
    auth,
    firestore
  };

  const gate = {
    hidden: true,
    classList: createClassList(),
    style: createStyle(),
    setAttribute() {},
    removeAttribute() {}
  };

  const document = {
    readyState: "complete",
    hidden: false,
    documentElement: { classList: createClassList() },
    getElementById(id) {
      return id === "auth-gate" ? gate : null;
    },
    querySelector() {
      return null;
    },
    addEventListener() {},
    createElement() {
      return {
        style: {},
        classList: createClassList(),
        setAttribute() {},
        addEventListener() {},
        appendChild() {}
      };
    }
  };

  const windowObject = {
    firebase,
    Capacitor: null,
    setTimeout(callback, milliseconds) {
      if (Number(milliseconds) <= 500) return setTimeout(callback, 0);
      return 0;
    },
    clearTimeout,
    addEventListener() {},
    dispatchEvent(event) {
      deferredEvents.push(event);
      return true;
    }
  };

  class MutationObserver {
    constructor(callback) {
      this.callback = callback;
    }
    observe() {}
    disconnect() {}
  }

  class CustomEvent {
    constructor(type, init = {}) {
      this.type = type;
      this.detail = init.detail;
    }
  }

  const context = {
    window: windowObject,
    document,
    firebase,
    MutationObserver,
    CustomEvent,
    queueMicrotask,
    setTimeout,
    clearTimeout,
    Promise,
    Map,
    Set,
    Date,
    RegExp,
    String,
    Boolean,
    Number,
    Array,
    Object,
    console: {
      info() {},
      log() {},
      warn(...args) { warnings.push(args); },
      error(...args) { errors.push(args); }
    },
    alert() {}
  };
  context.globalThis = context;
  windowObject.window = windowObject;

  vm.runInNewContext(authFix, context, { filename: "auth-login-fix.js" });

  return {
    authInstance,
    nativeAuthCallbacks,
    deferredEvents,
    warnings,
    errors,
    windowObject,
    getCalls: () => getCalls,
    setCalls: () => setCalls,
    getOptions
  };
}

async function registerTwoListenersAndDeliver(harness, user) {
  const delivered = [];
  harness.authInstance.onAuthStateChanged((receivedUser) => {
    delivered.push(["first", receivedUser]);
  });
  harness.authInstance.onAuthStateChanged((receivedUser) => {
    delivered.push(["second", receivedUser]);
  });

  assert.equal(harness.nativeAuthCallbacks.length, 2,
    "Entrambi i listener logici devono essere registrati");

  await Promise.all(harness.nativeAuthCallbacks.map((callback) => callback(user)));
  return delivered;
}

async function checkSingleFlightRecovery() {
  const unexpected = new Error("FIRESTORE (8.10.1) INTERNAL ASSERTION FAILED: Unexpected state");
  const harness = createHarness(({ call }) => {
    if (call === 1) return Promise.reject(unexpected);
    return Promise.resolve({
      exists: true,
      data() {
        return { role: "user" };
      }
    });
  });

  const user = {
    uid: "user-single-flight",
    email: "utente@example.com",
    emailVerified: true,
    displayName: "Utente",
    providerData: []
  };

  const delivered = await registerTwoListenersAndDeliver(harness, user);

  assert.equal(harness.getCalls(), 2,
    "Due listener concorrenti devono produrre una sola lettura iniziale e un solo retry");
  assert.equal(harness.getOptions[0], null);
  assert.equal(harness.getOptions[1].source, "server");
  assert.equal(delivered.length, 2);
  assert.ok(delivered.every((entry) => entry[1] === user),
    "Il login deve proseguire per entrambi i listener");
  assert.equal(harness.windowObject.HeraPlatformProfileBootstrap.exists, true);
  assert.equal(harness.windowObject.HeraPlatformProfileBootstrap.deferred, false);

  await harness.nativeAuthCallbacks[0](user);
  assert.equal(harness.getCalls(), 2,
    "Dopo il bootstrap completato non deve rileggere il profilo nella stessa sessione");
}

async function checkNonBlockingDeferredRecovery() {
  const unexpected = new Error("FIRESTORE (8.10.1) INTERNAL ASSERTION FAILED: Unexpected state");
  const harness = createHarness(() => Promise.reject(unexpected));

  const user = {
    uid: "user-deferred",
    email: "utente2@example.com",
    emailVerified: true,
    displayName: "Utente Due",
    providerData: []
  };

  const delivered = await registerTwoListenersAndDeliver(harness, user);

  assert.equal(harness.getCalls(), 2,
    "Anche in caso di errore persistente sono ammessi soltanto lettura iniziale e retry server");
  assert.equal(harness.setCalls(), 0,
    "Non deve creare o sovrascrivere il profilo quando la lettura è in stato interno incerto");
  assert.equal(delivered.length, 2,
    "L'errore interno Firestore non deve bloccare i callback di autenticazione");
  assert.equal(harness.windowObject.__heraPlatformProfileSyncDeferred, true);
  assert.equal(harness.windowObject.HeraPlatformProfileBootstrap.deferred, true);
  assert.ok(harness.deferredEvents.some((event) => event.type === "platform-profile-sync-deferred"));
  assert.equal(harness.errors.length, 0,
    "Il recupero gestito non deve aggiungere un secondo console.error applicativo");
}

(async () => {
  await checkSingleFlightRecovery();
  await checkNonBlockingDeferredRecovery();
  console.log("Auth profile recovery checks passed: single-flight, retry server unico e login non bloccante.");
})().catch((error) => {
  console.error(error);
  process.exitCode = 1;
});
