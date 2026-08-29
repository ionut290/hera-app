"use strict";

const ROOT_PACKAGE = require("../../package.json");
const ROOT_LOCK = require("../../package-lock.json");
const FUNCTIONS_PACKAGE = require("../../functions/package.json");
const FUNCTIONS_LOCK = require("../../functions/package-lock.json");

const PACKAGE_GROUPS = Object.freeze([
  { title: "Web/PWA", manifest: ROOT_PACKAGE, names: ["firebase"] },
  {
    title: "Android/Capacitor",
    manifest: ROOT_PACKAGE,
    names: [
      "@capacitor/core", "@capacitor/android", "@capacitor/cli",
      "@capacitor-firebase/authentication", "@capacitor/filesystem",
      "@capacitor/geolocation", "@capacitor/push-notifications", "@capacitor/share"
    ]
  },
  { title: "Backend/Functions", manifest: FUNCTIONS_PACKAGE, names: ["firebase-admin", "firebase-functions", "googleapis"] }
]);

const severityRank = Object.freeze({ info: 0, low: 1, moderate: 2, high: 3, critical: 4 });

function cleanVersion(value) {
  const match = String(value || "").match(/\d+(?:\.\d+){0,2}/);
  return match ? match[0] : String(value || "Non rilevata");
}

function major(version) {
  return Number.parseInt(cleanVersion(version).split(".")[0], 10) || 0;
}

function configuredVersion(manifest, name) {
  return manifest.dependencies?.[name] || manifest.devDependencies?.[name] || "";
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: { Accept: "application/json", ...(options.headers || {}) },
    signal: AbortSignal.timeout(8000)
  });
  if (!response.ok) throw new Error(`Servizio non disponibile (${response.status})`);
  return response.json();
}

async function latestVersion(packageName) {
  const payload = await fetchJson(`https://registry.npmjs.org/${encodeURIComponent(packageName)}/latest`);
  return String(payload.version || "");
}

function installedVersions(lock) {
  const versions = {};
  Object.entries(lock.packages || {}).forEach(([packagePath, data]) => {
    if (!packagePath.startsWith("node_modules/") || !data?.version) return;
    versions[packagePath.slice("node_modules/".length)] = String(data.version);
  });
  return versions;
}

async function auditLock(lock, label) {
  const body = Object.fromEntries(Object.entries(installedVersions(lock)).map(([name, version]) => [name, [version]]));
  try {
    const advisories = await fetchJson("https://registry.npmjs.org/-/npm/v1/security/advisories/bulk", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body)
    });
    const unique = new Map();
    Object.values(advisories || {}).flat().forEach((advisory) => {
      const key = advisory.url || `${advisory.name}:${advisory.title}`;
      if (!unique.has(key)) unique.set(key, advisory);
    });
    const totals = { info: 0, low: 0, moderate: 0, high: 0, critical: 0 };
    unique.forEach((advisory) => {
      const severity = String(advisory.severity || "moderate").toLowerCase();
      totals[severity] = (totals[severity] || 0) + 1;
    });
    const highest = Object.keys(totals).reduce((best, level) =>
      totals[level] && severityRank[level] > severityRank[best] ? level : best, "info");
    const total = Object.values(totals).reduce((sum, count) => sum + count, 0);
    return {
      category: "Sicurezza",
      name: `npm audit · ${label}`,
      current: total ? `${total} ${total === 1 ? "avviso" : "avvisi"}` : "0 avvisi",
      latest: total ? `Critiche ${totals.critical} · Alte ${totals.high} · Moderate ${totals.moderate} · Basse ${totals.low}` : "Nessuna vulnerabilità nota",
      status: totals.critical || totals.high ? "urgent" : total ? "planned" : "ok",
      deadline: totals.critical || totals.high ? "Appena possibile" : "Nessuna",
      message: total ? `Priorità massima: ${highest}. Verificare npm audit e applicare soltanto correzioni compatibili dopo i test.` : "Nessun avviso di sicurezza noto nel registro npm."
    };
  } catch (error) {
    return {
      category: "Sicurezza", name: `npm audit · ${label}`, current: "Controllo non riuscito",
      latest: "Registro npm non disponibile", status: "warning", deadline: "Ricontrollare",
      message: error.message || "Impossibile eseguire il controllo sicurezza."
    };
  }
}

async function dependencyItems() {
  const wanted = PACKAGE_GROUPS.flatMap((group) => group.names);
  const latestEntries = await Promise.all(wanted.map(async (name) => {
    try { return [name, await latestVersion(name), null]; }
    catch (error) { return [name, "", error]; }
  }));
  const latest = Object.fromEntries(latestEntries.map(([name, version]) => [name, version]));
  const failed = new Set(latestEntries.filter((entry) => entry[2]).map(([name]) => name));

  return PACKAGE_GROUPS.flatMap((group) => group.names.map((name) => {
    const current = cleanVersion(configuredVersion(group.manifest, name));
    const available = latest[name];
    if (failed.has(name)) return {
      category: group.title, name, current, latest: "Controllo non disponibile",
      status: "warning", deadline: "Ricontrollare",
      message: "Il registro npm non ha risposto; la dipendenza non viene considerata aggiornata automaticamente."
    };
    const changed = available && available !== current;
    const breaking = changed && major(available) > major(current);
    return {
      category: group.title, name, current, latest: available,
      status: changed ? "planned" : "ok", deadline: "Nessuna",
      message: breaking
        ? "Nuova versione principale: pianificare migrazione, build Android e test completi prima della pubblicazione."
        : changed
          ? "Aggiornamento compatibile disponibile: eseguire test prima della pubblicazione."
          : "Versione allineata al registro npm."
    };
  }));
}

function platformItems() {
  const buildId = String(process.env.COMMIT_REF || process.env.DEPLOY_ID || "locale").slice(0, 12);
  return [
    {
      category: "Runtime e pubblicazione", name: "Node.js runtime",
      current: cleanVersion(ROOT_PACKAGE.engines?.node || FUNCTIONS_PACKAGE.engines?.node), latest: "22 LTS",
      status: "ok", deadline: "Nessuna", message: "Runtime Node.js 22 configurato per build Netlify e Firebase Functions."
    },
    {
      category: "Runtime e pubblicazione", name: "Versione PWA / Service Worker",
      current: buildId, latest: "Distribuzione corrente", status: "ok", deadline: "Nessuna",
      message: "La cache PWA viene versionata automaticamente con il commit o con l’identificativo della pubblicazione."
    },
    {
      category: "Backend/Functions", name: "Firebase functions.config()",
      current: "Migrato", latest: "Secret RUNTIME_CONFIG", status: "ok", deadline: "Nessuna",
      message: "Migrazione completata: le Cloud Functions usano Firebase Secret Manager tramite RUNTIME_CONFIG."
    },
    {
      category: "Android/Capacitor", name: "Configurazione Android nativa",
      current: "Capacitor 7 · Java 17", latest: "Coerenza con Capacitor", status: "ok", deadline: "Nessuna",
      message: "Quando Capacitor passa a una nuova versione principale devono essere verificati insieme Gradle, SDK Android, Java e tutti i plugin nativi."
    }
  ];
}

exports.handler = async () => {
  try {
    const [dependencies, rootAudit, functionsAudit] = await Promise.all([
      dependencyItems(), auditLock(ROOT_LOCK, "Web/Android"), auditLock(FUNCTIONS_LOCK, "Functions")
    ]);
    const allItems = [...platformItems(), ...dependencies, rootAudit, functionsAudit];
    const categoryOrder = ["Runtime e pubblicazione", "Web/PWA", "Android/Capacitor", "Backend/Functions", "Sicurezza"];
    const items = categoryOrder.flatMap((category) => allItems.filter((item) => item.category === category));
    return {
      statusCode: 200,
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      body: JSON.stringify({ checkedAt: new Date().toISOString(), items })
    };
  } catch (error) {
    return {
      statusCode: 503,
      headers: { "Content-Type": "application/json", "Cache-Control": "no-store" },
      body: JSON.stringify({ error: error.message || "Controllo aggiornamenti non disponibile" })
    };
  }
};
