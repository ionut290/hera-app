#!/usr/bin/env node
"use strict";

const fs = require("node:fs");
const source = fs.readFileSync("app.js", "utf8");
const lines = source.split(/\r?\n/);

function showMatches(label, pattern, context = 12) {
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

showMatches("GET_COMMESSA_STATS_CALLS", /getCommessaStats\s*\(/, 14);
showMatches("GET_WORK_SUMMARY_CALLS", /getCommessaWorkSummary\s*\(/, 14);
showMatches("HOME_BUTTON", /function renderCommessaHomeButton\s*\(/, 45);
showMatches("PARENT_OVERVIEW", /function renderParentCommessaOverview\s*\(/, 55);
showMatches("MANAGEMENT_LIST", /function renderCommesseManagementList\s*\(/, 45);
showMatches("DASHBOARD", /function updateCommessaDashboard\s*\(/, 45);
showMatches("REFRESH_DEPENDENT_UI", /function refreshCommesseDependentUI\s*\(/, 25);
