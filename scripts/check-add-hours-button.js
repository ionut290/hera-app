const assert = require("node:assert/strict");
const fs = require("node:fs");
const path = require("node:path");
const vm = require("node:vm");

const root = path.join(__dirname, "..");
const app = fs.readFileSync(path.join(root, "app.js"), "utf8");
const sharedClient = fs.readFileSync(path.join(root, "shared-static-views-client-core.js"), "utf8");

const start = app.indexOf("function getQuickHoursOperatorEntries");
const end = app.indexOf("function createAddHoursButton", start);
assert.ok(start >= 0 && end > start, "Nuovo motore +ORE non trovato");
const engine = app.slice(start, end);

const normalizeName = (value) => String(value || "").toLocaleLowerCase("it-IT").replace(/\s+/g, " ").trim();
const parse = (value) => String(value || "").split(/[;,\n|]+/).map((part) => part.trim()).filter(Boolean);
const operatorIds = new Map([
  ["varga ionel", "p1"],
  ["benito pietro", "p2"],
  ["mario rossi", "p3"]
]);

function runScenario({ admin = false, user = "VARGA IONEL", rows = [], reports = [], loaded = true, date = "2026-08-28" } = {}) {
  const context = {
    Set,
    Map,
    Array,
    String,
    Number,
    currentUser: { uid: "u1", displayName: user, email: "ionel@example.test" },
    allHoursReports: reports,
    hoursReportsLoaded: loaded,
    canManageData: () => admin,
    getTodayDateKey: () => "2026-08-28",
    getActiveSquadreDateKey: () => date,
    getSquadraDataForCommessaDate: () => ({ squadre: rows }),
    getLegacySquadreRows: () => [],
    parseMultiEntryValue: parse,
    normalizeHoursOperatorName: normalizeName,
    normalizeEmail: (value) => String(value || "").trim().toLowerCase(),
    resolveHoursOperatorId: (value) => operatorIds.get(normalizeName(value)) || normalizeName(value),
    getCurrentUserSquadraIdentity: () => ({ user: normalizeName(user) }),
    getSquadraRowMembers: (row) => [row.personale, row.operatori, row.caposquadra].flatMap(parse),
    doesSquadraMemberMatchCurrentUser: (member, identity) => normalizeName(member) === identity.user
  };
  vm.runInNewContext(`${engine}\nresult = { visible: getQuickHoursContextForCommessa("c1", ${JSON.stringify(date)}), state: getQuickTeamHoursState("c1", ${JSON.stringify(date)}) };`, context);
  return context.result;
}

const twoTeams = [
  { personale: "VARGA IONEL, BENITO PIETRO", mezzi: "MEZZO 1" },
  { personale: "MARIO ROSSI", mezzi: "MEZZO 2" }
];

{
  const result = runScenario({ rows: twoTeams });
  assert.ok(result.visible, "L'operatore assegnato deve vedere +ORE");
  assert.deepEqual(Array.from(result.state.requiredParticipants, (item) => item.operatore), ["VARGA IONEL", "BENITO PIETRO"], "L'operatore deve compilare solo la propria squadra");
}

assert.equal(runScenario({ user: "UTENTE NON ASSEGNATO", rows: twoTeams }).visible, null, "Un operatore non assegnato non deve vedere +ORE");

{
  const result = runScenario({ admin: true, rows: twoTeams });
  assert.ok(result.visible, "L'admin deve vedere +ORE sulle squadre con ore mancanti");
  assert.deepEqual(Array.from(result.state.requiredParticipants, (item) => item.operatore), ["VARGA IONEL", "BENITO PIETRO", "MARIO ROSSI"], "L'admin deve poter completare tutte le squadre");
}

assert.ok(runScenario({ rows: twoTeams, loaded: false }).visible, "+ORE non deve sparire mentre i report ore stanno caricando");
assert.equal(runScenario({ rows: [{ mezzi: "MEZZO 1", note: "Nessun operatore" }] }).visible, null, "+ORE non deve comparire su una riga senza operatori");
assert.equal(runScenario({ rows: twoTeams, date: "2026-08-27" }).visible, null, "+ORE non deve comparire sulle date passate");

const teamOneCompleteByName = [{
  date: "2026-08-28",
  entries: [{
    commessaId: "c1",
    rows: [
      { operatore: "VARGA IONEL", ore: 8 },
      { operatore: "BENITO PIETRO", ore: 8 }
    ]
  }]
}];
assert.equal(runScenario({ rows: twoTeams, reports: teamOneCompleteByName }).visible, null, "+ORE deve sparire quando la squadra dell'operatore e completa");

{
  const result = runScenario({ rows: twoTeams, reports: [{
    date: "2026-08-28",
    entries: [{ commessaId: "c1", rows: [{ participantId: "utente:p1", operatore: "Nome storico", ore: 8 }] }]
  }] });
  assert.ok(result.visible, "+ORE deve restare visibile se manca un componente della squadra");
  assert.deepEqual(Array.from(result.state.missingParticipants, (item) => item.operatore), ["BENITO PIETRO"], "Il participantId storico deve riconoscere l'operatore gia compilato");
}

{
  const result = runScenario({ admin: true, rows: twoTeams, reports: teamOneCompleteByName });
  assert.ok(result.visible, "L'admin deve continuare a vedere +ORE se un'altra squadra e incompleta");
  assert.deepEqual(Array.from(result.state.missingParticipants, (item) => item.operatore), ["MARIO ROSSI"]);
}

assert.doesNotMatch(engine, /allHoursApprovalRequests|hoursApprovalsLoaded/, "Il nuovo +ORE non deve dipendere dalle approvazioni legacy");
assert.doesNotMatch(engine, /\.collection\(/, "Il render di +ORE non deve eseguire letture Firestore");
assert.match(app, /refreshQuickHoursReportsForDate[\s\S]*?\.where\("date", "==", dateKey\)/, "Al clic deve essere eseguita una sola verifica mirata per data");
assert.match(app, /appendAddHoursButtonIfAllowed\(head, commessa, dateKey\)/, "+ORE deve usare la data reale della scheda squadra");
assert.match(sharedClient, /render pulsante \+ORE da ore statiche/, "+ORE deve aggiornarsi quando arriva la vista ore condivisa");

const identityStart = app.indexOf("function getCurrentUserSquadraIdentity");
const identityEnd = app.indexOf("function getSquadraMemberIdentifiers", identityStart);
const identityBlock = app.slice(identityStart, identityEnd);
assert.match(identityBlock, /person\.linkedUserId/, "+ORE deve riconoscere il collegamento account-operatore");
assert.match(identityBlock, /person\.emailAccessoApp/, "+ORE deve riconoscere l'email di accesso app");
assert.match(identityBlock, /person\.linkedUserEmail/, "+ORE deve riconoscere l'email dell'utente collegato");

console.log("✅ Nuovo motore +ORE verificato con 10 scenari operatore/admin");
