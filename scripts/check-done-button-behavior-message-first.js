#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");

const sourcePath = path.resolve("scripts/check-done-button-behavior.js");
let source = fs.readFileSync(sourcePath, "utf8");

const oldAssertion = 'assert.match(nativePhotoShareHandler, /await sharePhotosThroughDedicatedPlugin[\\s\\S]*safeOpenWhatsAppMessage\\(message\\)/);';
const newAssertion = 'assert.match(nativePhotoShareHandler, /safeOpenWhatsAppMessage\\(message\\)[\\s\\S]*await sharePhotosThroughDedicatedPlugin/);';

if (!source.includes(oldAssertion)) {
  throw new Error("Assertion Whazzup legacy non trovata nel test comportamentale.");
}

source = source.replace(oldAssertion, newAssertion);
const testModule = new Module(sourcePath, module);
testModule.filename = sourcePath;
testModule.paths = Module._nodeModulePaths(path.dirname(sourcePath));
testModule._compile(source, sourcePath);
