#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const source = fs.readFileSync("app.js", "utf8");
const lines = source.split(/\r?\n/);

const hits = [];
for (let i = 0; i < lines.length; i += 1) {
  const line = lines[i];
  if (!/impianti/.test(line)) continue;
  const start = Math.max(0, i - 14);
  const end = Math.min(lines.length, i + 18);
  const block = lines.slice(start, end).join("\n");
  if (!/onSnapshot\s*\(/.test(block)) continue;
  if (!/(collection\s*\(\s*["']impianti["']\s*\)|\.collection\s*\(\s*["']impianti["']\s*\))/.test(block)) continue;
  hits.push({ line: i + 1, block: lines.slice(start, end).map((text, index) => `${start + index + 1}: ${text}`).join("\n") });
}

console.log(`AUDIT_IMPIANTI_LISTENER_HITS=${hits.length}`);
hits.forEach((hit, index) => {
  console.log(`\n===== HIT ${index + 1} @ line ${hit.line} =====`);
  console.log(hit.block);
});

const generic = [];
for (let i = 0; i < lines.length; i += 1) {
  if (!/onSnapshot\s*\(/.test(lines[i])) continue;
  const start = Math.max(0, i - 10);
  const end = Math.min(lines.length, i + 12);
  const block = lines.slice(start, end).join("\n");
  if (/impianti/.test(block) && /commess/.test(block)) {
    generic.push({ line: i + 1, block: lines.slice(start, end).map((text, index) => `${start + index + 1}: ${text}`).join("\n") });
  }
}
console.log(`\nAUDIT_GENERIC_COMMESSA_IMPIANTI_ONSNAPSHOT=${generic.length}`);
generic.forEach((hit, index) => {
  console.log(`\n----- GENERIC ${index + 1} @ line ${hit.line} -----`);
  console.log(hit.block);
});
