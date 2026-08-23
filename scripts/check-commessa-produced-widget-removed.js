#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

function createElement() {
  const attributes = new Map();
  return {
    hidden: false,
    disabled: false,
    dataset: {},
    classList: { remove() {} },
    setAttribute(name, value) { attributes.set(name, String(value)); },
    getAttribute(name) { return attributes.get(name) || null; }
  };
}

const elements = {
  "commessa-produced-widget": createElement(),
  "commessa-produced-toggle": createElement(),
  "commessa-produced-popover": createElement()
};

global.document = {
  documentElement: {},
  readyState: "loading",
  head: { appendChild() {} },
  body: { appendChild() {} },
  createElement() { return createElement(); },
  getElementById(id) { return elements[id] || null; },
  querySelector() { return null; },
  querySelectorAll() { return []; },
  addEventListener() {}
};
global.window = {
  addEventListener() {},
  setTimeout,
  clearTimeout,
  setInterval,
  clearInterval
};
global.MutationObserver = class MutationObserver {
  observe() {}
  disconnect() {}
};

let firestoreTouched = false;
global.db = new Proxy({}, {
  get() {
    firestoreTouched = true;
    throw new Error("Il widget rimosso non deve accedere a Firestore");
  }
});

require(path.resolve(__dirname, "..", "commessa-produced-widget.js"));

assert.equal(window.CommessaProducedWidget.removed, true);
assert.equal(elements["commessa-produced-widget"].hidden, true);
assert.equal(elements["commessa-produced-toggle"].disabled, true);
assert.equal(elements["commessa-produced-toggle"].getAttribute("aria-expanded"), "false");
assert.equal(elements["commessa-produced-popover"].getAttribute("aria-hidden"), "true");

window.CommessaProducedWidget.select("commessa-a");
window.CommessaProducedWidget.stop();

assert.equal(firestoreTouched, false, "PRODOTTO rimosso non deve aprire alcun listener o lettura Firestore");
assert.equal(elements["commessa-produced-widget"].hidden, true);

console.log("✅ PRODOTTO è nascosto e disabilitato.");
console.log("✅ Nessuna lettura o listener Firestore viene avviato.");
