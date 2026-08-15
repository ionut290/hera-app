#!/usr/bin/env node
"use strict";

const assert = require("node:assert/strict");
const fs = require("node:fs");

const code = fs.readFileSync("commessa-navigation-repair.js", "utf8");
const marker = "HeraImpiantiListenerLifecycleGuard";
const idx = code.indexOf(marker);
assert.ok(idx >= 0, "Impianti listener lifecycle guard missing");

const sectionStart = code.lastIndexOf("// Evita il ciclo stop -> subscribe", idx);
const section = code.slice(sectionStart >= 0 ? sectionStart : idx);
const executable = section
  .replace(/\/\*[\s\S]*?\*\//g, "")
  .replace(/^\s*\/\/.*$/gm, "");

assert.match(section, /const DEFER_MS = 80;/);
assert.match(section, /window\.stopImpiantiSubscription = function stopImpiantiSubscriptionWithLifecycleGuard/);
assert.match(section, /window\.subscribeImpianti = function subscribeImpiantiWithLifecycleGuard/);
assert.match(section, /pendingStop\.commessaId === commessaId/);
assert.match(section, /activeCommessaId === commessaId/);
assert.match(section, /cancelledSameCommessaRestarts \+= 1/);
assert.match(section, /runPendingStop\(\)/);
assert.doesNotMatch(executable, /\.collection\(|\.doc\(|\.get\(|\.set\(|\.update\(|\.delete\(|\.onSnapshot\(/, "Lifecycle guard must not call Firestore directly");
assert.doesNotMatch(executable, /fattoVisualEvidence|subscribeFattoVisualEvidence|stopFattoVisualEvidence/i, "FATTO visual evidence must stay untouched");
assert.doesNotMatch(executable, /WhatsApp|WHAZZUP|whazzup/i, "WhatsApp/WHAZZUP must stay untouched");

console.log("Impianti listener lifecycle safety checks passed");
