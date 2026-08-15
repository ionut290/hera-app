#!/usr/bin/env node
"use strict";

const fs = require("node:fs");

function replaceOnce(source, before, after, label) {
  const first = source.indexOf(before);
  if (first < 0) throw new Error(`Blocco non trovato: ${label}`);
  if (source.indexOf(before, first + before.length) >= 0) throw new Error(`Blocco ambiguo: ${label}`);
  return source.slice(0, first) + after + source.slice(first + before.length);
}

let optimizer = fs.readFileSync("firestore-safe-optimizer.js", "utf8");
optimizer = replaceOnce(
  optimizer,
  `    const canonical = canonicalQuery(query);\n    // Una CollectionReference semplice ha \`path\`. Per una Query filtrata,\n    // invece, pretendiamo sempre un identificatore canonico: così query con\n    // filtri o limiti diversi non possono mai essere unite per errore.\n    const isPlainCollection = Boolean(query?.path && path === collection);\n    if (!canonical && !isPlainCollection) return null;`,
  `    const canonical = canonicalQuery(query);\n    // Una CollectionReference semplice espone il proprio \`path\` pubblico.\n    // È sicuro usare il percorso completo anche per sottocollezioni come\n    // commesse/{id}/impianti: due reference con lo stesso path rappresentano\n    // esattamente la stessa raccolta. Le Query filtrate/ordinate/limitate, che\n    // non espongono una CollectionReference semplice, continuano invece a\n    // richiedere un identificatore canonico e non vengono mai unite a intuito.\n    const publicPath = canonicalPath(query?.path).replace(/^\\/+|\\/+$/g, \"\");\n    const isPlainCollectionReference = Boolean(publicPath && publicPath === path);\n    if (!canonical && !isPlainCollectionReference) return null;`,
  "queryInfo nested collection"
);
fs.writeFileSync("firestore-safe-optimizer.js", optimizer);

let test = fs.readFileSync("scripts/check-firestore-safe-optimizer.js", "utf8");
test = replaceOnce(
  test,
  `class FakeQuery {\n  constructor(collection, canonical) {\n    this.path = collection;\n    this._query = {\n      path: { canonicalString: () => collection },\n      canonicalId: () => canonical\n    };\n  }`,
  `class FakeQuery {\n  constructor(collection, canonical, { exposePath = true } = {}) {\n    if (exposePath) this.path = collection;\n    this._query = {\n      path: { canonicalString: () => collection },\n      canonicalId: () => canonical\n    };\n  }`,
  "FakeQuery options"
);

test = replaceOnce(
  test,
  `  const chatA = new FakeQuery(\"chatMessages\", \"chatMessages|limit:500\");\n  const chatB = new FakeQuery(\"chatMessages\", \"chatMessages|limit:500\");\n  const unsubscribeChatA = chatA.onSnapshot(() => {});\n  const unsubscribeChatB = chatB.onSnapshot(() => {});\n  assert.equal(physicalStarts, 5, \"Le raccolte non autorizzate devono mantenere il comportamento Firestore originale\");`,
  `  const nestedA = new FakeQuery(\"commesse/alpha/impianti\", \"\");\n  const nestedB = new FakeQuery(\"commesse/alpha/impianti\", \"\");\n  const nestedReceivedA = [];\n  const nestedReceivedB = [];\n  const unsubscribeNestedA = nestedA.onSnapshot((snapshot) => nestedReceivedA.push(snapshot));\n  const unsubscribeNestedB = nestedB.onSnapshot((snapshot) => nestedReceivedB.push(snapshot));\n  assert.equal(physicalStarts, 4, \"Due CollectionReference annidate identiche devono usare un solo listener fisico\");\n  const nestedSnapshot = { marker: \"impianti-alpha\" };\n  emit(3, nestedSnapshot);\n  assert.deepEqual(nestedReceivedA, [nestedSnapshot]);\n  assert.deepEqual(nestedReceivedB, [nestedSnapshot]);\n\n  const nestedOther = new FakeQuery(\"commesse/beta/impianti\", \"\");\n  const unsubscribeNestedOther = nestedOther.onSnapshot(() => {});\n  assert.equal(physicalStarts, 5, \"Sottocollezioni di commesse diverse devono restare indipendenti\");\n\n  const unresolvedA = new FakeQuery(\"commesse/alpha/impianti\", \"\", { exposePath: false });\n  const unresolvedB = new FakeQuery(\"commesse/alpha/impianti\", \"\", { exposePath: false });\n  const unsubscribeUnresolvedA = unresolvedA.onSnapshot(() => {});\n  const unsubscribeUnresolvedB = unresolvedB.onSnapshot(() => {});\n  assert.equal(physicalStarts, 7, \"Query senza path pubblico né canonical id non devono essere unite\");\n\n  const chatA = new FakeQuery(\"chatMessages\", \"chatMessages|limit:500\");\n  const chatB = new FakeQuery(\"chatMessages\", \"chatMessages|limit:500\");\n  const unsubscribeChatA = chatA.onSnapshot(() => {});\n  const unsubscribeChatB = chatB.onSnapshot(() => {});\n  assert.equal(physicalStarts, 9, \"Le raccolte non autorizzate devono mantenere il comportamento Firestore originale\");`,
  "nested tests insertion"
);

test = replaceOnce(
  test,
  `  unsubscribeAlertA();\n  unsubscribeAlertB();\n  unsubscribeChatA();\n  unsubscribeChatB();`,
  `  unsubscribeAlertA();\n  unsubscribeAlertB();\n  unsubscribeNestedA();\n  unsubscribeNestedB();\n  unsubscribeNestedOther();\n  unsubscribeUnresolvedA();\n  unsubscribeUnresolvedB();\n  unsubscribeChatA();\n  unsubscribeChatB();`,
  "nested unsubscribes"
);

test = replaceOnce(
  test,
  `    3,\n    \"Devono essere evitati il secondo commesse, la riapertura commesse e il secondo userAlerts\"\n  );\n  assert.equal(state.stats.gracePeriodReuses, 1);\n  assert.equal(state.stats.physicalListenersStarted, 3);\n  assert.equal(physicalCloses, 5, \"Tutti i listener fisici devono poter essere chiusi normalmente\");\n\n  console.log(\"✅ Commesse, squadreStorico e userAlerts condividono solo query identiche.\");`,
  `    4,\n    \"Devono essere evitati il secondo commesse, la riapertura commesse, il secondo userAlerts e il secondo listener impianti annidato\"\n  );\n  assert.equal(state.stats.gracePeriodReuses, 1);\n  assert.equal(state.stats.physicalListenersStarted, 5);\n  assert.equal(physicalCloses, 9, \"Tutti i listener fisici devono poter essere chiusi normalmente\");\n\n  console.log(\"✅ Commesse, squadreStorico e userAlerts condividono solo query identiche.\");\n  console.log(\"✅ Le CollectionReference annidate identiche condividono un solo listener fisico.\");`,
  "expected counters"
);
fs.writeFileSync("scripts/check-firestore-safe-optimizer.js", test);

console.log("Ottimizzatore esteso alle CollectionReference annidate con test dedicati.");
