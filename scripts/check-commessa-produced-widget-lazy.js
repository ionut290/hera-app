#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const path = require("node:path");

function createElement() {
  const attributes = new Map();
  const listeners = new Map();
  return {
    hidden: false,
    textContent: "",
    style: {},
    classList: { add() {}, remove() {} },
    setAttribute(name, value) { attributes.set(name, String(value)); },
    getAttribute(name) { return attributes.get(name) || null; },
    addEventListener(type, listener) { listeners.set(type, listener); },
    emit(type, event = { stopPropagation() {} }) { return listeners.get(type)?.(event); },
    getBoundingClientRect() { return { right: 200, bottom: 40, top: 10 }; }
  };
}

const elements = {
  "commessa-produced-widget": createElement(),
  "commessa-produced-toggle": createElement(),
  "commessa-produced-value": createElement(),
  "commessa-produced-popover": createElement()
};

global.document = {
  getElementById(id) { return elements[id]; },
  addEventListener() {}
};
global.requestAnimationFrame = (callback) => callback();
global.window = {
  innerWidth: 390,
  innerHeight: 844,
  addEventListener() {},
  InreteWorkItemsV2: {
    adaptLegacyPlantToWorkItems() { return []; },
    buildPriceMap() { return new Map(); },
    enrichWorkItem(item) { return item; },
    calculateCompletedSubtotal() { return 0; }
  }
};

let subscriptions = 0;
let unsubscriptions = 0;
function subscribe() {
  subscriptions += 1;
  return () => { unsubscriptions += 1; };
}

const commessaRef = {
  collection() { return { onSnapshot() { return subscribe(); } }; },
  onSnapshot() { return subscribe(); }
};

global.db = {
  collection() {
    return { doc() { return commessaRef; } };
  }
};
global.getCommesseCollectionName = () => "commesse";

require(path.resolve(__dirname, "..", "commessa-produced-widget.js"));

window.CommessaProducedWidget.select("commessa-a");
assert.equal(subscriptions, 0, "Aprire una commessa non deve avviare i listener PRODOTTO");
assert.equal(elements["commessa-produced-widget"].hidden, false);
assert.equal(elements["commessa-produced-value"].textContent, "—");

elements["commessa-produced-toggle"].emit("click");
assert.equal(subscriptions, 4, "Il primo tocco su PRODOTTO deve avviare i quattro listener necessari");

elements["commessa-produced-toggle"].emit("click");
elements["commessa-produced-toggle"].emit("click");
assert.equal(subscriptions, 4, "Chiudere e riaprire PRODOTTO non deve duplicare i listener");

window.CommessaProducedWidget.select("commessa-b");
assert.equal(unsubscriptions, 4, "Cambiare commessa deve chiudere tutti i listener precedenti");
assert.equal(subscriptions, 4, "La nuova commessa resta lazy finché PRODOTTO non viene aperto");

elements["commessa-produced-toggle"].emit("click");
assert.equal(subscriptions, 8, "PRODOTTO deve avviare i listener della nuova commessa al primo tocco");

window.CommessaProducedWidget.stop();
assert.equal(unsubscriptions, 8, "Stop deve chiudere tutti i listener attivi");
assert.equal(elements["commessa-produced-widget"].hidden, true);

console.log("✅ Il widget PRODOTTO non legge Firestore finché non viene aperto.");
console.log("✅ Riaperture e cambi commessa non duplicano i listener.");
