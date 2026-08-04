#!/usr/bin/env node
"use strict";

const admin = require("firebase-admin");

const date = String(process.argv[2] || "2026-08-04");
const month = String(process.argv[3] || "2026-08");

if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) {
  throw new Error(`Input date non valido: ${date}. Formato richiesto YYYY-MM-DD.`);
}
if (!/^\d{4}-\d{2}$/.test(month)) {
  throw new Error(`Input month non valido: ${month}. Formato richiesto YYYY-MM.`);
}
if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  throw new Error("GOOGLE_APPLICATION_CREDENTIALS non configurato.");
}

admin.initializeApp({
  credential: admin.credential.applicationDefault(),
  projectId: process.env.GCLOUD_PROJECT || "hera-app-6cd2b"
});

const { rebuildRegistryView, rebuildSquadreDate, rebuildCalendarMonth } =
  require("../shared-operational-views").__server;

async function run() {
  const completed = [];
  try {
    const registri = await rebuildRegistryView();
    completed.push(registri.id);
    console.log("Registri ricostruiti", registri);

    const squadre = await rebuildSquadreDate(date);
    if (!squadre) throw new Error(`Ricostruzione squadre non riuscita per ${date}.`);
    completed.push(squadre.id);
    console.log("Squadre ricostruite", squadre);

    const calendario = await rebuildCalendarMonth(month);
    if (!calendario) throw new Error(`Ricostruzione calendario non riuscita per ${month}.`);
    completed.push(calendario.id);
    console.log("Calendario ricostruito", calendario);

    const result = {
      personale: registri.personale,
      mezzi: registri.mezzi,
      squadre: squadre.count,
      righeCalendario: calendario.count,
      documentiScritti: completed,
      completedAt: new Date().toISOString()
    };
    console.log("BACKFILL_SHARED_STATIC_VIEWS_RESULT=" + JSON.stringify(result));
  } catch (error) {
    console.error("Backfill fallito.", {
      message: error?.message || String(error),
      documentiGiaScritti: completed,
      failedAt: new Date().toISOString()
    });
    process.exitCode = 1;
  } finally {
    await admin.app().delete();
  }
}

run();
