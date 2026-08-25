const fs = require('fs');
const path = require('path');

const file = path.join(__dirname, '..', 'admin-error-center.js');
let source = fs.readFileSync(file, 'utf8');

function replaceOnce(label, from, to) {
  if (!source.includes(from)) throw new Error(`Patch ${label}: anchor not found`);
  source = source.replace(from, to);
}

replaceOnce('version', 'const VERSION = "1.3.0";', 'const VERSION = "1.4.0";');

replaceOnce(
  'github repair helper',
  '  function solutionHtml(item) {',
  `  function githubRepairIssueUrl(item) {\n    const impact = impactCategory(item);\n    const diagnostic = {\n      errorCenterId: item?.id || "",\n      title: item?.title || "",\n      category: item?.category || "",\n      impact: impact?.label || "",\n      severity: item?.severity || "",\n      feature: item?.feature || "",\n      page: item?.lastPage || item?.lastActiveView || "",\n      platform: item?.lastPlatform || "",\n      appVersion: item?.lastAppVersion || "",\n      message: item?.lastMessage || "",\n      stack: String(item?.lastStack || "").slice(0, 5000),\n      commessaId: item?.commessaId || "",\n      commessaName: item?.commessaName || "",\n      impiantoId: item?.impiantoId || ""\n    };\n    const title = '[AUTO-REPAIR] ' + String(item?.title || item?.category || 'Errore app').slice(0, 160);\n    const advice = repairAdvice(item);\n    const body = [\n      '## Richiesta automatica dal Centro errori',\n      '',\n      '**Obiettivo:** correggere questo errore con la modifica minima e sicura, senza alterare i flussi non coinvolti.',\n      '',\n      '**Impatto:** ' + (impact?.icon || '') + ' ' + (impact?.label || ''),\n      '**Gravità:** ' + (item?.severity || ''),\n      '',\n      '### Diagnostica',\n      '~~~json',\n      JSON.stringify(diagnostic, null, 2),\n      '~~~',\n      '',\n      '### Indicazione del Centro errori',\n      advice.summary || '',\n      '',\n      '### Vincoli di sicurezza',\n      '- Non modificare né indebolire la protezione FATTO / Whazzup.',\n      '- Non modificare automaticamente dati di commesse o impianti.',\n      '- Non eseguire merge automatico.',\n      '- Creare una PR in bozza solo se i controlli critici passano.'\n    ].join('\\n');\n    return 'https://github.com/ionut290/hera-app/issues/new?title=' + encodeURIComponent(title) + '&body=' + encodeURIComponent(body);\n  }\n\n  function startGithubRepair() {\n    const item = selectedItem();\n    if (!item) return;\n    const root = document.querySelector('[data-error-repair-result]');\n    const url = githubRepairIssueUrl(item);\n    if (root) root.innerHTML = '<div class="hera-error-solution-note">🚀 Sto aprendo GitHub con la diagnostica già compilata. Dopo aver creato la richiesta, il workflow Codex prepara una correzione su branch separato, esegue i controlli critici e crea una PR in bozza solo se i test passano.</div>';\n    const opened = window.open(url, '_blank', 'noopener,noreferrer');\n    if (!opened) location.href = url;\n  }\n\n  function solutionHtml(item) {`
);

replaceOnce(
  'github repair button',
  '<div class="hera-error-repair-actions"><button class="btn hera-error-repair-btn" type="button" data-error-find-solution>🧠 TROVA SOLUZIONE</button><button class="btn hera-error-repair-btn ${repairAdvice(item).safeAutoRepair ? "is-safe" : ""}" type="button" data-error-repair>🛠️ RIPARA ERRORE</button></div>',
  '<div class="hera-error-repair-actions"><button class="btn hera-error-repair-btn" type="button" data-error-find-solution>🧠 TROVA SOLUZIONE</button><button class="btn hera-error-repair-btn ${repairAdvice(item).safeAutoRepair ? "is-safe" : ""}" type="button" data-error-repair>🛠️ RIPARA ERRORE</button><button class="btn hera-error-repair-btn" type="button" data-error-github-repair>🚀 RIPARA SU GITHUB</button></div>'
);

replaceOnce(
  'github repair handler',
  '    if (event.target.closest?.("[data-error-find-solution]")) return showSelectedSolution();\n    if (event.target.closest?.("[data-error-repair]")) return void repairSelectedError();',
  '    if (event.target.closest?.("[data-error-find-solution]")) return showSelectedSolution();\n    if (event.target.closest?.("[data-error-repair]")) return void repairSelectedError();\n    if (event.target.closest?.("[data-error-github-repair]")) return startGithubRepair();'
);

fs.writeFileSync(file, source, 'utf8');
console.log('Applied GitHub repair UI patch.');
