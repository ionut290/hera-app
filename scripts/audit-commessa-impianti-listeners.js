#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const source = fs.readFileSync("app.js", "utf8");
const lines = source.split(/\r?\n/);

function showMatches(label, pattern, context = 8) {
  const indexes = [];
  lines.forEach((line, index) => { if (pattern.test(line)) indexes.push(index); });
  console.log(`\n${label}=${indexes.length}`);
  indexes.forEach((index, n) => {
    const start = Math.max(0, index - context);
    const end = Math.min(lines.length, index + context + 1);
    console.log(`\n===== ${label} ${n + 1} @ line ${index + 1} =====`);
    console.log(lines.slice(start, end).map((text, offset) => `${start + offset + 1}: ${text}`).join("\n"));
  });
}

showMatches("SUBSCRIBE_STATS_CALLS", /subscribeStatsForCommesse\s*\(/, 14);
showMatches("COMMESSA_STATS_MAP_USES", /commessaStatsById/, 8);
showMatches("IMPIANTI_BY_COMMESSA_USES", /impiantiByCommessaId/, 8);
showMatches("WORK_SUMMARY_USES", /commessaWorkSummariesById/, 8);
showMatches("HOURS_BY_COMMESSA_USES", /commessaHoursById/, 8);
showMatches("CALCULATE_STATS", /calculateImpiantiStats\s*\(/, 12);
showMatches("PARENT_OVERVIEW_CALLS", /renderParentCommessaOverview\s*\(/, 10);
showMatches("MANAGEMENT_LIST_CALLS", /renderCommesseManagementList\s*\(/, 10);
