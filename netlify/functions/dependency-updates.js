const INSTALLED = Object.freeze({
  "firebase-admin": "13.10.0",
  "firebase-functions": "7.3.2",
  googleapis: "176.0.0"
});

function major(version) {
  return Number.parseInt(String(version || "0").replace(/^[^0-9]*/, "").split(".")[0], 10) || 0;
}

async function latestVersion(packageName) {
  const response = await fetch(`https://registry.npmjs.org/${encodeURIComponent(packageName)}/latest`, {
    headers: { Accept: "application/json" },
    signal: AbortSignal.timeout(6000)
  });
  if (!response.ok) throw new Error(`Registro npm non disponibile per ${packageName}`);
  const payload = await response.json();
  return String(payload.version || "");
}

exports.handler = async () => {
  try {
    const latestEntries = await Promise.all(Object.keys(INSTALLED).map(async (name) => [name, await latestVersion(name)]));
    const latest = Object.fromEntries(latestEntries);
    const dependencyItems = Object.entries(INSTALLED).map(([name, current]) => {
      const available = latest[name];
      const changed = available && available !== current;
      const breaking = changed && major(available) > major(current);
      return {
        name, current, latest: available, status: changed ? "planned" : "ok", deadline: "Nessuna",
        message: breaking ? "Nuova versione principale: aggiornare soltanto con migrazione e test completi." : changed ? "Aggiornamento disponibile: verificare compatibilità e test prima della pubblicazione." : "Versione allineata al registro npm."
      };
    });
    const items = [
      { name: "Node.js runtime", current: "22", latest: "22 LTS", status: "ok", deadline: "Nessuna", message: "Runtime Node.js 22 configurato per build, funzioni e controlli automatici." },
      ...dependencyItems,
      { name: "Firebase functions.config()", current: "Ancora utilizzato", latest: "Secrets e parametri", status: "planned", deadline: "Marzo 2027", message: "Sostituire functions.config() con Firebase Secrets e parametri in una migrazione separata." }
    ];
    return { statusCode: 200, headers: { "Content-Type": "application/json", "Cache-Control": "public, max-age=900" }, body: JSON.stringify({ checkedAt: new Date().toISOString(), items }) };
  } catch (error) {
    return { statusCode: 503, headers: { "Content-Type": "application/json", "Cache-Control": "no-store" }, body: JSON.stringify({ error: error.message || "Controllo aggiornamenti non disponibile" }) };
  }
};
