const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");
const app = read("app.js");
const html = read("index.html");
const rules = read("firestore.rules");
const css = read("calendar-feature.css");
const androidWorkflow = read(".github/workflows/build-android-aab.yml");

const checks = [
  ["La home apre il calendario", html.includes('id="home-calendar-btn"') && !html.includes('id="home-private-docs-btn"')],
  ["La view calendario esiste", html.includes('id="calendar-page"') && html.includes('id="calendar-grid"')],
  ["Il form include i tipi richiesti", ["ferie", "permesso", "malattia", "intervento"].every((type) => html.includes(`value="${type}"`))],
  ["Il calendario usa Firestore in tempo reale", app.includes('db.collection("calendarEvents").onSnapshot')],
  ["Il giorno mostra il conteggio eventi", app.includes("calendar-event-count") && app.includes("getCalendarEventsForDate")],
  ["La cancellazione verifica autore o admin", app.includes("canModifyCalendarEvent") && app.includes('db.collection("calendarEvents").doc(eventId).delete()')],
  ["Le regole consentono lettura condivisa", rules.includes("match /calendarEvents/{eventId}") && rules.includes("allow read: if signedIn();")],
  ["Le regole proteggono la cancellazione", rules.includes("resource.data.createdByUid == request.auth.uid")],
  ["Lo stile mobile del calendario esiste", css.includes("@media (max-width: 600px)") && css.includes(".calendar-grid")],
  ["Android include lo stile del calendario", androidWorkflow.includes("cp calendar-feature.css www/") && androidWorkflow.includes("test -s www/calendar-feature.css")]
];

const failed = checks.filter(([, ok]) => !ok);
checks.forEach(([label, ok]) => console.log(`${ok ? "✓" : "✗"} ${label}`));
if (failed.length) process.exit(1);
