#!/usr/bin/env node
"use strict";

const admin = require("firebase-admin");

if (!process.env.GOOGLE_APPLICATION_CREDENTIALS) {
  throw new Error("GOOGLE_APPLICATION_CREDENTIALS non configurato.");
}

admin.initializeApp({
  credential: admin.credential.applicationDefault(),
  projectId: process.env.GCLOUD_PROJECT || "hera-app-6cd2b"
});

const { rebuildRegistryView } = require("../shared-operational-views").__server;

async function run() {
  try {
    const registri = await rebuildRegistryView();
    const result = {
      personale: registri.personale,
      mezzi: registri.mezzi,
      payloadBytes: registri.bytes,
      documentoScritto: registri.id,
      completedAt: new Date().toISOString()
    };
    console.log("INITIALIZE_SHARED_REGISTRIES_RESULT=" + JSON.stringify(result));
  } catch (error) {
    console.error("Inizializzazione registri condivisi fallita.", {
      message: error?.message || String(error),
      failedAt: new Date().toISOString()
    });
    process.exitCode = 1;
  } finally {
    await admin.app().delete();
  }
}

run();
