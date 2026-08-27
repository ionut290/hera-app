const fs = require("fs");
const path = require("path");

const app = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const start = app.indexOf("function getQuickHoursContextForCommessa");
const end = app.indexOf("function createAddHoursButton", start);
if (start < 0 || end < 0) throw new Error("Logica +ORE non trovata");
const block = app.slice(start, end);

const checks = [
  [block.includes("if (!hoursReportsLoaded) return null;"), "+ORE attende i report ore"],
  [!block.includes("!hoursReportsLoaded || !hoursApprovalsLoaded"), "+ORE non dipende dalla raccolta approvazioni legacy"],
  [block.includes("getCurrentUserSquadraAssignment(commessaId, dateKey)"), "+ORE verifica l'assegnazione dell'utente"],
  [block.includes("if (!assignment) return null;"), "+ORE resta nascosto agli utenti non assegnati"],
  [block.includes("areAllHoursParticipantsCompleteForCommessaDate(commessaId, dateKey)"), "+ORE scompare quando le ore squadra sono complete"]
];

for (const [ok, message] of checks) {
  if (!ok) throw new Error(message);
}

console.log("✅ Logica pulsante +ORE verificata");
