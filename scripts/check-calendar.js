const fs = require("fs");
const path = require("path");

const root = path.resolve(__dirname, "..");
const read = (file) => fs.readFileSync(path.join(root, file), "utf8");
const appCore = read("app.js");
const calendarModule = fs.existsSync(path.join(root, "app-calendar.js")) ? read("app-calendar.js") : "";
// I controlli funzionali devono osservare l'intero runtime classico, non soltanto
// app.js: le funzioni calendario estratte mantengono gli stessi nomi globali e
// vengono caricate prima del core.
const app = `${appCore}\n${calendarModule}`;
const html = read("index.html");
const rules = read("firestore.rules");
const css = read("calendar-feature.css");
const androidWorkflow = read(".github/workflows/build-android-aab.yml");
const capacitorBundle = read("scripts/prepare-capacitor-web.js");
const notifications = read("functions/user-notifications.js");
const deployWorkflow = read(".github/workflows/deploy-firebase-functions.yml");
const reminderRunner = read("functions/run-calendar-admin-reminders.js");
const reminderWorkflow = read(".github/workflows/calendar-admin-reminders.yml");

const checks = [
  ["La home apre il calendario", html.includes('id="home-calendar-btn"') && !html.includes('id="home-private-docs-btn"')],
  ["La view calendario esiste", html.includes('id="calendar-page"') && html.includes('id="calendar-grid"')],
  ["La scelta iniziale separa ore personali e calendario condiviso", html.includes('id="calendar-choice-hours-btn"') && html.includes('id="calendar-choice-shared-btn"')],
  ["Le modalità restano accessibili come schede", html.includes('id="calendar-hours-tab"') && html.includes('id="calendar-shared-tab"') && app.includes('setCalendarMode("hours")')],
  ["Il calendario personale filtra le righe per utente", app.includes("getPersonalHoursRowsForDate") && app.includes("doesSquadraMemberMatchCurrentUser")],
  ["Le ore personali sono formattate in ore e minuti", app.includes("function formatPersonalHours") && app.includes('padStart(2, "0")')],
  ["Il calendario personale si aggiorna con i report ore", app.includes('calendarMode === "hours"') && app.includes("hoursReportsLoaded")],
  ["Il form include i tipi richiesti", ["ferie", "permesso", "malattia", "intervento"].every((type) => html.includes(`value="${type}"`))],
  ["Il calendario usa Firestore in tempo reale", app.includes('db.collection("calendarEvents")') && app.includes("onSnapshot(applySnapshot")],
  ["Il giorno mostra il conteggio eventi", app.includes("calendar-event-count") && app.includes("getCalendarEventsForDate")],
  ["La cancellazione verifica autore o admin", app.includes("canModifyCalendarEvent") && app.includes('db.collection("calendarEvents").doc(eventId).delete()')],
  ["Le regole consentono lettura condivisa e proteggono le scadenze private", rules.includes("match /calendarEvents/{eventId}") && rules.includes('resource.data.type != "SCADENZA_DOCUMENTO"') && rules.includes("resource.data.ownerUserId == request.auth.uid")],
  ["Le regole proteggono la cancellazione", rules.includes("resource.data.createdByUid == request.auth.uid")],
  ["Lo stile mobile del calendario esiste", css.includes("@media (max-width: 600px)") && css.includes(".calendar-grid")],
  ["Android include lo stile del calendario", androidWorkflow.includes("npm run android:aab:prepare") && capacitorBundle.includes('"calendar-feature.css"')],
  ["Il partecipante corrente viene proposto automaticamente", app.includes("getCurrentUserResolvedName") && app.includes("addCalendarParticipant")],
  ["I partecipanti vengono scelti dal personale o scritti liberamente", app.includes("calendarSelectedParticipants") && app.includes("data-calendar-free-person")],
  ["Commessa e impianto usano gli elenchi dell'app", app.includes("populateCalendarCommesse") && app.includes("populateCalendarImpianti")],
  ["Il luogo viene compilato dall'impianto", app.includes("getCalendarImpiantoLocation") && app.includes("handleCalendarImpiantoChange")],
  ["Il titolo può essere generato automaticamente", app.includes("titleWasGenerated") && !/id="calendar-event-title"[^>]*required/.test(html)],
  ["Le squadre controllano gli operatori assenti", app.includes("validateSquadraOperatorAvailability") && app.includes("avvisoAutomaticoAssenze")],
  ["La funzione crea avvisi squadra dalle assenze", notifications.includes("syncCalendarAbsenceSquadraAlerts") && notifications.includes("calendar-absence")],
  ["L'admin riceve il promemoria un giorno prima", reminderRunner.includes("getTomorrowRomeDateKey") && reminderRunner.includes("calendar-admin-reminder")],
  ["Le assenze notificano solo operatori assegnati a una squadra", reminderRunner.includes("absenceEventHasSquadraAssignment") && reminderRunner.includes('db.collection("squadreStorico").where("dateKey", "==", tomorrowKey)')],
  ["Il promemoria parte alle 07:05 italiane", reminderWorkflow.includes('cron: "5 5 * * *"') && reminderWorkflow.includes('cron: "5 6 * * *"') && reminderRunner.includes('romeParts.hour !== "07"')],
  ["Le funzioni Cloud restano nel deploy senza lo scheduler IAM", deployWorkflow.includes("functions:syncCalendarAbsenceSquadraAlerts") && !deployWorkflow.includes("functions:notifyAdminsCalendarEventsTomorrow")]
];

const failed = checks.filter(([, ok]) => !ok);
checks.forEach(([label, ok]) => console.log(`${ok ? "✓" : "✗"} ${label}`));
if (failed.length) process.exit(1);
