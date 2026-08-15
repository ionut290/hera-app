#!/usr/bin/env node
'use strict';

const fs = require('fs');
const path = require('path');

const ROOT = path.resolve(__dirname, '..');
const EXCLUDED_DIRS = new Set([
  '.git', 'node_modules', 'www', 'dist', 'build', 'coverage', '.netlify',
  'android', 'android-resources'
]);
const EXCLUDED_FILES = new Set([
  'scripts/audit-firestore-reads.js'
]);

const READ_PATTERNS = [
  { kind: 'listener', re: /\.onSnapshot\s*\(/g },
  { kind: 'get', re: /\.get\s*\(/g },
  { kind: 'getDoc', re: /\bgetDoc\s*\(/g },
  { kind: 'getDocs', re: /\bgetDocs\s*\(/g },
  { kind: 'collectionGroup', re: /\bcollectionGroup\s*\(/g },
  { kind: 'countFromServer', re: /\bgetCountFromServer\s*\(/g }
];

const FIRESTORE_HINT = /(firebase\.firestore|\bdb\.collection\s*\(|\.collection\s*\(|\bgetDoc\s*\(|\bgetDocs\s*\(|\bonSnapshot\s*\()/;
const STARTUP_SENSITIVE = new Set([
  'app.js',
  'approval-access.js',
  'auth-login-fix.js',
  'login-retry-fix.js',
  'notification-center.js',
  'active-commesse-first-boot-guard.js',
  'firestore-startup-cost-optimizer.js',
  'firestore-safe-optimizer.js',
  'shared-static-views.js',
  'shared-static-views-client-core.js',
  'squadra-current-save-sync.js'
]);
const SHOULD_BE_ON_DEMAND = [
  /^documents\.js$/,
  /^private-documents-v2\.js$/,
  /^accounting-v2\.js$/,
  /^preventivi-.*\.js$/,
  /^app-worklimate\.js$/,
  /^app-atex\.js$/,
  /^hours-export-range\.js$/,
  /^operational-import-repair\.js$/,
  /^global-archive-sync\.js$/,
  /^inrete-work-items-v2\.js$/,
  /^operator-profile-feature\.js$/,
  /^identity-card-feature\.js$/
];

function walk(dir, out = []) {
  for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
    if (entry.name.startsWith('.') && entry.name !== '.github') {
      if (entry.isDirectory()) continue;
    }
    const full = path.join(dir, entry.name);
    const rel = path.relative(ROOT, full).replace(/\\/g, '/');
    if (entry.isDirectory()) {
      if (!EXCLUDED_DIRS.has(entry.name)) walk(full, out);
      continue;
    }
    if (!entry.isFile() || !/\.(?:js|mjs|cjs)$/.test(entry.name)) continue;
    if (EXCLUDED_FILES.has(rel)) continue;
    out.push({ full, rel });
  }
  return out;
}

function lineForIndex(text, index) {
  return text.slice(0, index).split('\n').length;
}

function nearby(text, index, radius = 120) {
  return text.slice(Math.max(0, index - radius), Math.min(text.length, index + radius))
    .replace(/\s+/g, ' ')
    .trim();
}

function looksTopLevel(text, index) {
  const before = text.slice(0, index);
  let depth = 0;
  let inSingle = false, inDouble = false, inTemplate = false, escape = false;
  for (let i = 0; i < before.length; i += 1) {
    const ch = before[i];
    if (escape) { escape = false; continue; }
    if (ch === '\\') { escape = true; continue; }
    if (!inDouble && !inTemplate && ch === "'") { inSingle = !inSingle; continue; }
    if (!inSingle && !inTemplate && ch === '"') { inDouble = !inDouble; continue; }
    if (!inSingle && !inDouble && ch === '`') { inTemplate = !inTemplate; continue; }
    if (inSingle || inDouble || inTemplate) continue;
    if (ch === '{') depth += 1;
    else if (ch === '}') depth = Math.max(0, depth - 1);
  }
  return depth <= 1;
}

const files = walk(ROOT);
const findings = [];

for (const file of files) {
  const text = fs.readFileSync(file.full, 'utf8');
  if (!FIRESTORE_HINT.test(text)) continue;
  for (const pattern of READ_PATTERNS) {
    pattern.re.lastIndex = 0;
    let match;
    while ((match = pattern.re.exec(text))) {
      const base = path.basename(file.rel);
      const onDemand = SHOULD_BE_ON_DEMAND.some((re) => re.test(base));
      const topLevel = looksTopLevel(text, match.index);
      const startupSensitive = STARTUP_SENSITIVE.has(base);
      let risk = 'low';
      if (pattern.kind === 'listener') risk = 'medium';
      if (startupSensitive && pattern.kind === 'listener') risk = 'high';
      if (onDemand && topLevel) risk = 'high';
      findings.push({
        file: file.rel,
        line: lineForIndex(text, match.index),
        kind: pattern.kind,
        risk,
        startupSensitive,
        onDemandExpected: onDemand,
        possibleTopLevel: topLevel,
        context: nearby(text, match.index)
      });
    }
  }
}

const byFile = new Map();
for (const item of findings) {
  const row = byFile.get(item.file) || { total: 0, listeners: 0, high: 0, medium: 0, low: 0 };
  row.total += 1;
  if (item.kind === 'listener') row.listeners += 1;
  row[item.risk] += 1;
  byFile.set(item.file, row);
}

const summary = [...byFile.entries()]
  .map(([file, counts]) => ({ file, ...counts }))
  .sort((a, b) => b.high - a.high || b.listeners - a.listeners || b.total - a.total || a.file.localeCompare(b.file));

const report = {
  generatedAt: new Date().toISOString(),
  filesScanned: files.length,
  filesWithFirestoreReads: summary.length,
  totalReadSites: findings.length,
  highRiskSites: findings.filter((x) => x.risk === 'high').length,
  listenerSites: findings.filter((x) => x.kind === 'listener').length,
  summary,
  findings
};

const outDir = path.join(ROOT, 'audit-output');
fs.mkdirSync(outDir, { recursive: true });
fs.writeFileSync(path.join(outDir, 'firestore-read-audit.json'), JSON.stringify(report, null, 2) + '\n');

const md = [];
md.push('# Firestore read audit');
md.push('');
md.push(`Generated: ${report.generatedAt}`);
md.push(`Files scanned: ${report.filesScanned}`);
md.push(`Files with read sites: ${report.filesWithFirestoreReads}`);
md.push(`Read sites: ${report.totalReadSites}`);
md.push(`Listener sites: ${report.listenerSites}`);
md.push(`High-risk sites: ${report.highRiskSites}`);
md.push('');
md.push('| File | Total | Listener | High | Medium | Low |');
md.push('|---|---:|---:|---:|---:|---:|');
for (const row of summary) md.push(`| ${row.file} | ${row.total} | ${row.listeners} | ${row.high} | ${row.medium} | ${row.low} |`);
md.push('');
md.push('## High-risk sites');
md.push('');
for (const item of findings.filter((x) => x.risk === 'high')) {
  md.push(`- \`${item.file}:${item.line}\` — ${item.kind}${item.onDemandExpected ? ' — modulo atteso on-demand' : ''}${item.possibleTopLevel ? ' — possibile esecuzione top-level' : ''}`);
}
fs.writeFileSync(path.join(outDir, 'firestore-read-audit.md'), md.join('\n') + '\n');

console.log(`Firestore audit: ${report.totalReadSites} read sites in ${report.filesWithFirestoreReads} files; ${report.highRiskSites} high-risk; ${report.listenerSites} listener sites.`);
console.log('Reports: audit-output/firestore-read-audit.json and audit-output/firestore-read-audit.md');

if (process.argv.includes('--fail-on-high') && report.highRiskSites > 0) process.exit(2);
