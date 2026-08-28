const fs = require("fs");
const path = require("path");

const app = fs.readFileSync(path.join(__dirname, "..", "app.js"), "utf8");
const start = app.indexOf("function getQuickHoursContextForCommessa");
const end = app.indexOf("function createAddHoursButton", start);
if (start < 0 || end < 0) throw new Error("Logica +ORE non trovata");
const block = app.slice(start, end);

const checks = [
  [block.includes("if (!hoursReportsLoaded) return null;"), "+ORE attende i report ore"],
  [block.includes("dateKey < getTodayDateKey()"), "+ORE non compare per date passate"],
  [!block.includes("!hoursReportsLoaded || !hoursApprovalsLoaded"), "+ORE non dipende dalla raccolta approvazioni legacy"],
  [block.includes("getCurrentUserSquadraAssignment(commessaId, dateKey)"), "+ORE verifica l'assegnazione dell'utente"],
  [block.includes("if (!assignment && !canManageData()) return null;"), "+ORE resta nascosto agli operatori non assegnati ma visibile agli admin"],
  [block.includes("areAllHoursParticipantsCompleteForCommessaDate(commessaId, dateKey)"), "+ORE scompare quando le ore squadra sono complete"]
];

const identityStart = app.indexOf("function getCurrentUserSquadraIdentity");
const identityEnd = app.indexOf("function getSquadraMemberIdentifiers", identityStart);
if (identityStart < 0 || identityEnd < 0) throw new Error("Identità squadra non trovata");
const identityBlock = app.slice(identityStart, identityEnd);
checks.push(
  [identityBlock.includes("person.linkedUserId"), "+ORE riconosce il collegamento account-operatore"],
  [identityBlock.includes("person.emailAccessoApp"), "+ORE riconosce l'email di accesso app"],
  [identityBlock.includes("person.linkedUserEmail"), "+ORE riconosce l'email dell'utente collegato"]
);

const subscriptionStart = app.indexOf("function subscribeSquadre");
const subscriptionEnd = app.indexOf("function stopPersonaleSubscription", subscriptionStart);
const subscriptionBlock = app.slice(subscriptionStart, subscriptionEnd);
checks.push(
  [subscriptionBlock.includes("selectedDateKey, todayDateKey, tomorrowDateKey"), "Il listener squadre carica anche i cantieri di domani"]
);

const renderStart = app.indexOf("function renderSquadre()");
const renderEnd = app.indexOf("function renderSquadraImpiantiButtons", renderStart);
const renderBlock = app.slice(renderStart, renderEnd);
checks.push(
  [renderBlock.includes("canManageData() && selectedDateKey === todayDateKey"), "L'admin vede domani restando sulla data odierna"],
  [renderBlock.includes("appendAddHoursButtonIfAllowed(head, commessa, dateKey)"), "+ORE usa la data reale della squadra visualizzata"]
);

for (const [ok, message] of checks) {
  if (!ok) throw new Error(message);
}

console.log("✅ Logica pulsante +ORE verificata");
