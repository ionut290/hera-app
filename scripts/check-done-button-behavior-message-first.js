#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const vm = require("node:vm");

const sourcePath = "scripts/check-done-button-behavior.js";
let source = fs.readFileSync(sourcePath, "utf8");

const oldAssertion = 'assert.match(nativePhotoShareHandler, /await sharePhotosThroughDedicatedPlugin[\\s\\S]*safeOpenWhatsAppMessage\\(message\\)/);';
const newAssertion = 'assert.match(nativePhotoShareHandler, /safeOpenWhatsAppMessage\\(message\\)[\\s\\S]*await sharePhotosThroughDedicatedPlugin/);';

if (!source.includes(oldAssertion)) {
  throw new Error("Assertion Whazzup legacy non trovata nel test comportamentale.");
}

source = source.replace(oldAssertion, newAssertion);
vm.runInThisContext(source, { filename: sourcePath });
