const fs = require("fs");

const source = fs.readFileSync("app.js", "utf8");
const start = source.indexOf("function subscribeGlobalNotifications()");
const end = source.indexOf("function stopGlobalNotificationsSubscription()", start);

if (start < 0 || end < 0) {
  throw new Error("Listener notifiche globali non trovato.");
}

const listener = source.slice(start, end);

for (const expected of [
  "firebase.firestore.Timestamp.now()",
  '.where("createdAt", ">", listenFrom)',
  '.orderBy("createdAt", "desc")',
  ".limit(10)",
  'change.type === "added"'
]) {
  if (!listener.includes(expected)) {
    throw new Error(`Protezione letture notifiche incompleta: ${expected}`);
  }
}

if (listener.includes("if (!globalNotificationsInitialized)")) {
  throw new Error("Il listener scarica ancora notifiche storiche soltanto per scartarle.");
}

console.log("✅ Il listener globale legge solo notifiche successive all’avvio.");
