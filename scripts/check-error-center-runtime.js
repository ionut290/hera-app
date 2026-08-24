"use strict";
const fs = require("node:fs");
const vm = require("node:vm");

function makeSandbox() {
  const listeners = new Map();
  const local = new Map();
  const document = {
    readyState: "complete",
    visibilityState: "visible",
    body: { dataset: {}, appendChild() {} },
    baseURI: "https://example.test/",
    addEventListener(type, fn) { listeners.set(`document:${type}`, fn); },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    getElementById() { return null; },
    createElement(tag) {
      return {
        tagName: String(tag).toUpperCase(),
        dataset: {},
        className: "",
        innerHTML: "",
        appendChild() {},
        addEventListener() {},
        querySelector() { return null; },
        querySelectorAll() { return []; },
        setAttribute() {},
        removeAttribute() {},
        closest() { return null; }
      };
    }
  };
  const sandbox = {
    console,
    document,
    navigator: {
      onLine: true,
      userAgent: "UnitTest",
      platform: "test",
      language: "it-IT",
      serviceWorker: { addEventListener() {} }
    },
    location: { pathname: "/", hash: "", search: "", origin: "https://example.test" },
    performance: { now: () => 0 },
    crypto: { randomUUID: () => "test-id" },
    localStorage: {
      getItem(key) { return local.get(key) ?? null; },
      setItem(key, value) { local.set(key, String(value)); },
      removeItem(key) { local.delete(key); }
    },
    URL,
    URLSearchParams,
    Map,
    Set,
    Object,
    Array,
    String,
    Number,
    Math,
    Date,
    RegExp,
    Error,
    JSON,
    Promise,
    CSS: { escape: (value) => String(value) },
    getComputedStyle: () => ({ display: "block", visibility: "visible" }),
    requestAnimationFrame: (fn) => { fn(); return 1; },
    setTimeout: () => 1,
    clearTimeout() {},
    addEventListener(type, fn) { listeners.set(`window:${type}`, fn); },
    dispatchEvent() {},
    alert() {},
    confirm: () => true,
    prompt: () => ""
  };
  sandbox.window = sandbox;
  sandbox.globalThis = sandbox;
  return sandbox;
}

const clientSandbox = makeSandbox();
vm.runInNewContext(fs.readFileSync("client-error-reporter.js", "utf8"), clientSandbox, { filename: "client-error-reporter.js" });
if (!clientSandbox.HeraClientErrorReporter?.installed) throw new Error("Reporter non installato");

const adminSandbox = makeSandbox();
adminSandbox.canManageData = () => false;
vm.runInNewContext(fs.readFileSync("admin-error-center.js", "utf8"), adminSandbox, { filename: "admin-error-center.js" });
if (!adminSandbox.HeraAdminErrorCenter?.installed) throw new Error("Centro amministratore non esportato");

console.log("Smoke runtime client error center passed.");
