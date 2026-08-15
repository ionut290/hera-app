#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const acorn = require("acorn");

const source = fs.readFileSync("app.js", "utf8");
const ast = acorn.parse(source, { ecmaVersion: "latest", sourceType: "script", ranges: true });

const nameRe = /(ferie|disponibilit|assenza|assenze|permesso|malattia)/i;
const excluded = new Set([
  "applyCalendarAbsenceWarningsToSquadraRows",
  "validateSquadraOperatorAvailability"
]);

const rows = [];
for (const node of ast.body) {
  if (node.type !== "FunctionDeclaration" || !node.id?.name) continue;
  const name = node.id.name;
  const body = source.slice(node.start, node.end);
  if (!nameRe.test(name) && !nameRe.test(body.slice(0, 500))) continue;
  if (excluded.has(name)) continue;
  rows.push({
    name,
    bytes: Buffer.byteLength(body),
    async: !!node.async,
    dom: /document\.|querySelector|getElementById|classList/.test(body),
    firebase: /\bdb\.|firestore|collection\(|onSnapshot|get\(|set\(|update\(|delete\(/.test(body)
  });
}

console.log(JSON.stringify({ count: rows.length, bytes: rows.reduce((n, r) => n + r.bytes, 0), functions: rows }, null, 2));
