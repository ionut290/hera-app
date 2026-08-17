#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const path = require("node:path");
const Module = require("node:module");

const sourcePath = path.resolve("scripts/check-done-button-behavior.js");
let source = fs.readFileSync(sourcePath, "utf8");

const oldClickAssertion = 'assert.doesNotMatch(immediateSource, /addEventListener\\(\"click\"/);';
const newClickAssertion = 'assert.doesNotMatch(immediateSource, /(?:window|document)\\.addEventListener\\(\"click\"/);';
if (!source.includes(oldClickAssertion)) throw new Error("Assertion click globale non trovata nel test comportamentale.");
source = source.replace(oldClickAssertion, newClickAssertion);

const testModule = new Module(sourcePath, module);
testModule.filename = sourcePath;
testModule.paths = Module._nodeModulePaths(path.dirname(sourcePath));
testModule._compile(source, sourcePath);
