/* Timbratura: impostazioni locali e sveglie native, isolate dal gestionale. */
(() => {
  "use strict";
  const KEY = "varga-timbratura-sveglie-v1";
  const DEFAULTS = Object.freeze({
    latitude: "", longitude: "", radius: "150",
    arrivalFrom: "06:30", arrivalTo: "07:15", arrivalAlarm: "06:40",
    returnFrom: "15:15", returnTo: "16:15", returnAlarm: "15:35"
  });
  const $ = (id) => document.getElementById(id);
  const FIELDS = Object.keys(DEFAULTS);
  const modal = $("timbratura-reminder-panel");
  if (!modal || window.__vargaTimbraturaReminders) return;
  window.__vargaTimbraturaReminders = true;
  const feedback = $("timbratura-reminder-feedback");
  const status = $("timbratura-reminder-status");
  let working = false;

  function say(message, error = false) {
    if (!feedback) return;
    feedback.textContent = message;
    feedback.dataset.error = error ? "1" : "0";
  }
  function saved() {
    try {
      const data = JSON.parse(localStorage.getItem(KEY) || "{}");
      return { ...DEFAULTS, ...(data && typeof data === "object" ? data : {}) };
    } catch (_) { return { ...DEFAULTS }; }
  }
  function fill(data) {
    for (const name of FIELDS) {
      const element = $("timbratura-" + name);
      if (element) element.value = data[name] ?? DEFAULTS[name];
    }
  }
  function read() {
    const data = {};
    for (const name of FIELDS) data[name] = ($("timbratura-" + name)?.value || "").trim();
    return data;
  }
  function minutes(value) {
    if (!/^([01]\\d|2[0-3]):[0-5]\\d$/.test(value)) return null;
    const [h, m] = value.split(":").map(Number);
    return h * 60 + m;
  }
  function validate(data) {
    for (const type of ["arrival", "return"]) {
      const start = minutes(data[type + "From"]);
      const stop = minutes(data[type + "To"]);
      const alarm = minutes(data[type + "Alarm"]);
      if (start === null || stop === null || alarm === null || start >= stop || alarm < start || alarm > stop) {
        throw new Error("Controlla la fascia oraria " + (type === "arrival" ? "di entrata" : "di rientro") + ": la sveglia deve essere dentro la fascia.");
      }
    }
    const latEntered = data.latitude !== "";
    const lonEntered = data.longitude !== "";
    if (latEntered !== lonEntered) throw new Error("Inserisci entrambe le coordinate GPS, latitudine e longitudine.");
    if (latEntered) {
      const latitude = Number(data.latitude.replace(",", "."));
      const longitude = Number(data.longitude.replace(",", "."));
      if (!Number.isFinite(latitude) || latitude < -90 || latitude > 90 ||
          !Number.isFinite(longitude) || longitude < -180 || longitude > 180) {
        throw new Error("Coordinate GPS non valide.");
      }
    }
    const radius = Number(data.radius);
    if (!Number.isFinite(radius) || radius < 30 || radius > 2000) {
      throw new Error("Il raggio GPS deve essere fra 30 e 2000 metri.");
    }
    return data;
  }
  function save() {
    const data = validate(read());
    localStorage.setItem(KEY, JSON.stringify(data));
    status.textContent = "Impostazioni salvate su questo dispositivo (senza inviare la posizione al server).";
    return data;
  }
  function platform() {
    const cap = window.Capacitor;
    return cap?.isNativePlatform?.() ? cap.getPlatform?.() : "web";
  }
  function plugin() {
    const cap = window.Capacitor;
    if (!cap?.isNativePlatform?.() || !["ios", "android"].includes(platform())) {
      throw new Error("Le sveglie sonore richiedono VARGA CANTIERI installata come app nativa iOS o Android. Nel browser puoi solo salvare le impostazioni.");
    }
    if (!cap?.isPluginAvailable?.("CapgoAlarm")) {
      throw new Error("Questa versione dell'app non include ancora le sveglie native: installa la prossima build aggiornata.");
    }
    return cap.Plugins?.CapgoAlarm || cap.registerPlugin?.("CapgoAlarm");
  }
  async function arm(type, settings) {
    const alarm = plugin();
    const nativePlatform = platform();
    const time = settings[type + "Alarm"];
    const [hour, minute] = time.split(":").map(Number);
    const label = type === "arrival" ? "VARGA CANTIERI · Timbratura ENTRATA magazzino" : "VARGA CANTIERI · Timbratura USCITA / RIENTRO magazzino";
    if (nativePlatform === "ios") {
      const version = Number((navigator.userAgent.match(/OS (\\d+)_/) || [])[1]);
      if (Number.isFinite(version) && version < 26) {
        throw new Error("Le sveglie che suonano fino allo stop su iPhone richiedono iOS 26 o successivo (AlarmKit).");
      }
      const p = await alarm.requestPermissions();
      if (!p?.granted) throw new Error("Permesso AlarmKit negato: autorizza le sveglie nelle impostazioni dell'iPhone.");
    }
    const result = await alarm.createAlarm({
      hour, minute, label, vibrate: true,
      // Android: show the Clock app so the user verifies actual alarm creation.
      skipUi: false
    });
    if (!result?.success) throw new Error(result?.message || "Il sistema operativo non ha confermato la sveglia.");
    return time;
  }
  async function activate(type) {
    if (working) return;
    working = true;
    try {
      const data = save();
      if (type === "both") {
        say("Sto configurando la sveglia di entrata...");
        const first = await arm("arrival", data);
        say("Entrata " + first + " configurata. Ora la sveglia del rientro...");
        const second = await arm("return", data);
        say("Sveglie native richieste: entrata " + first + " e rientro " + second + ". Controllale sul dispositivo. Non equivalgono alla timbratura effettiva.");
      } else {
        say("Configurazione della sveglia in corso...");
        const time = await arm(type, data);
        say("Sveglia " + (type === "arrival" ? "di entrata" : "di rientro") + " alle " + time + " richiesta al sistema. Controlla l'orario sul dispositivo.");
      }
      if (platform() === "ios") {
        status.textContent = "iPhone: questa versione crea la prossima sveglia singola con AlarmKit. Non è ancora ricorrente: per i giorni successivi occorre riattivarla.";
      } else if (platform() === "android") {
        status.textContent = "Android: la sveglia viene creata nell'app Orologio. Controlla lì ripetizione, suoneria e cancellazione; l'app non può cancellarla automaticamente.";
      }
    } catch (e) { say(e?.message || "Impossibile attivare la sveglia.", true); }
    finally { working = false; }
  }
  function open() {
    fill(saved());
    modal.hidden = false;
    modal.setAttribute("aria-hidden", "false");
    const menuClose = $("menu-close-btn");
    if (document.body.classList.contains("menu-fullscreen-open")) menuClose?.click();
    say("");
    status.textContent = "Le sveglie ti ricordano di timbrare: non effettuano timbrature automatiche e non registrano presenze.";
    $("timbratura-close-btn")?.focus();
  }
  function close() { modal.hidden = true; modal.setAttribute("aria-hidden", "true"); }
  function distanceMeters(lat1, lon1, lat2, lon2) {
    const r = Math.PI / 180;
    const a = Math.sin(((lat2 - lat1) * r) / 2) ** 2
      + Math.cos(lat1 * r) * Math.cos(lat2 * r) * Math.sin(((lon2 - lon1) * r) / 2) ** 2;
    return 6371000 * 2 * Math.atan2(Math.sqrt(a), Math.sqrt(1 - a));
  }
  function getPosition() {
    return new Promise((resolve, reject) => {
      if (!navigator.geolocation) return reject(new Error("GPS non disponibile."));
      navigator.geolocation.getCurrentPosition(resolve, (error) => reject(
        new Error("Posizione GPS non disponibile o permesso non concesso (" + (error?.code || "?") + ").")
      ), { enableHighAccuracy: true, maximumAge: 0, timeout: 15000 });
    });
  }
  function busyAction(action) {
    if (working) return;
    working = true;
    Promise.resolve().then(action).catch(e => say(e?.message || "Operazione GPS non riuscita.", true))
      .finally(() => { working = false; });
  }
  $("open-timbratura-reminder-btn")?.addEventListener("click", open);
  $("timbratura-close-btn")?.addEventListener("click", close);
  modal.addEventListener("click", e => { if (e.target === modal) close(); });
  $("timbratura-save-btn")?.addEventListener("click", () => {
    try { save(); say("Impostazioni salvate. Per far suonare una sveglia premi «Attiva» qui sotto."); }
    catch (e) { say(e.message, true); }
  });
  $("timbratura-gps-save-btn")?.addEventListener("click", () => busyAction(async () => {
    say("Leggo la posizione GPS del magazzino...");
    const pos = await getPosition();
    $("timbratura-latitude").value = pos.coords.latitude.toFixed(7);
    $("timbratura-longitude").value = pos.coords.longitude.toFixed(7);
    say("Coordinate GPS inserite. Premi Salva mentre sei effettivamente al magazzino.");
  }));
  $("timbratura-gps-check-btn")?.addEventListener("click", () => busyAction(async () => {
    const config = validate(read());
    if (!config.latitude) throw new Error("Prima inserisci le coordinate GPS del magazzino.");
    say("Verifico la distanza dal magazzino...");
    const pos = await getPosition();
    const d = distanceMeters(Number(config.latitude.replace(",", ".")), Number(config.longitude.replace(",", ".")), pos.coords.latitude, pos.coords.longitude);
    const within = d <= Number(config.radius);
    say((within ? "Sei nel raggio configurato" : "Sei fuori dal raggio configurato") +
      " (distanza approssimativa " + Math.round(d) + " m; precisione GPS ±" + Math.round(pos.coords.accuracy) + " m). Il GPS non certifica una timbratura.", !within);
  }));
  $("timbratura-arm-arrival-btn")?.addEventListener("click", () => activate("arrival"));
  $("timbratura-arm-return-btn")?.addEventListener("click", () => activate("return"));
  $("timbratura-arm-both-btn")?.addEventListener("click", () => activate("both"));
  $("timbratura-open-clock-btn")?.addEventListener("click", async () => {
    try {
      const alarm = plugin();
      const result = await alarm.openAlarms();
      if (!result?.success) throw new Error(result?.message || "Apri Orologio manualmente dal dispositivo.");
    } catch (e) { say(e?.message || "Impossibile aprire Orologio.", true); }
  });
})();
