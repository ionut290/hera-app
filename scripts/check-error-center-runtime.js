"use strict";
const fs = require("node:fs");
const vm = require("node:vm");

function makeNode(tag) {
  return {
    tagName: String(tag || "div").toUpperCase(),
    dataset: {},
    style: {},
    className: "",
    innerHTML: "",
    hidden: false,
    open: false,
    classList: { add() {}, remove() {}, contains() { return false; } },
    appendChild() {},
    append() {},
    remove() {},
    addEventListener() {},
    querySelector() { return null; },
    querySelectorAll() { return []; },
    setAttribute() {},
    removeAttribute() {},
    getAttribute() { return null; },
    closest() { return null; },
    showModal() { this.open = true; },
    close() { this.open = false; },
    select() {},
    focus() {}
  };
}

function makeSandbox() {
  const listeners = new Map();
  const local = new Map();
  const head = makeNode("head");
  const body = makeNode("body");
  body.dataset = {};
  const document = {
    readyState: "complete",
    visibilityState: "visible",
    body,
    head,
    baseURI: "https://example.test/",
    addEventListener(type, fn) { listeners.set(`document:${type}`, fn); },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    getElementById() { return null; },
    createElement: makeNode,
    execCommand() { return true; }
  };
  const sandbox = {
    console,
    document,
    navigator: {
      onLine: true,
      userAgent: "UnitTest",
      platform: "test",
      language: "it-IT",
      serviceWorker: { addEventListener() {} },
      clipboard: { writeText: async () => undefined }
    },
    location: { pathname: "/", hash: "", search: "", origin: "https://example.test" },
    history: { state: null, replaceState() {} },
    performance: { now: () => 0 },
    crypto: { randomUUID: () => "test-id" },
    localStorage: {
      getItem(key) { return local.get(key) ?? null; },
      setItem(key, value) { local.set(key, String(value)); },
      removeItem(key) { local.delete(key); }
    },
    URL,
    URLSearchParams,
    FormData: class FormData { get() { return ""; } },
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
if (!clientSandbox.HeraClientErrorReporter?.installed) throw new Error("Reporter email non installato");

const monitorSandbox = makeSandbox();
vm.runInNewContext(fs.readFileSync("app-error-monitor.js", "utf8"), monitorSandbox, { filename: "app-error-monitor.js" });
if (!monitorSandbox.HeraAppErrorMonitor?.installed) throw new Error("Monitor globale non esportato");
if (typeof monitorSandbox.HeraAppErrorMonitor.reportManual !== "function") throw new Error("Segnalazione manuale non disponibile");

const adminSandbox = makeSandbox();
adminSandbox.canManageData = () => false;
vm.runInNewContext(fs.readFileSync("admin-error-center.js", "utf8"), adminSandbox, { filename: "admin-error-center.js" });
if (!adminSandbox.HeraAdminErrorCenter?.installed) throw new Error("Centro amministratore non esportato");
if (typeof adminSandbox.HeraAdminErrorCenter.openBugReport !== "function") throw new Error("Modulo segnalazione bug non esportato");

console.log("Smoke runtime Centro errori superato.");
