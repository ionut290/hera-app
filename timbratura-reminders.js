/* Promemoria di timbratura: GPS + fascia oraria, SOLO sul dispositivo. */
(() => {
  "use strict";
  if (window.__vargaTimbraturaReminders) return;
  window.__vargaTimbraturaReminders = true;

  const KEY = "varga-timbratura-promemoria-v2";
  const DEFAULTS = {
    enabled: false, lat: "", lon: "", radius: 150,
    arrivalFrom: "06:30", arrivalTo: "07:15",
    returnFrom: "15:15", returnTo: "16:15",
    days: [1, 2, 3, 4, 5], tone: "pulse", snooze: 5, iosScheduled: false, iosGeofence: false
  };
  const $ = (id) => document.getElementById(id);
  let config = load(), watchId = null, windowTimer = null, soundTimer = null;
  let audio = null, ringing = "", waitingUntil = {}, done = {}, nativePending = false;
  try { done = JSON.parse(localStorage.getItem(KEY + "-ack") || "{}") || {}; } catch (_) {}
  function saveAck() {
    try { localStorage.setItem(KEY + "-ack", JSON.stringify(done)); } catch (_) {}
  }

  function load() {
    try {
      const value = JSON.parse(localStorage.getItem(KEY) || "{}");
      if (!value || typeof value !== "object") return { ...DEFAULTS };
      return { ...DEFAULTS, ...value };
    } catch (_) { return { ...DEFAULTS }; }
  }
  const two = n => String(n).padStart(2,"0");
  function dayKey() { const d = new Date(); return [d.getFullYear(), two(d.getMonth()+1), two(d.getDate())].join("-"); }
  function clockMinutes(s) {
    if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(s)) return null;
    const [h,m] = s.split(":").map(Number); return h*60+m;
  }
  function inWindow(which, date = new Date()) {
    if (!config.days.includes(date.getDay())) return false;
    const minutes = date.getHours()*60 + date.getMinutes();
    return minutes >= clockMinutes(config[which+"From"]) && minutes <= clockMinutes(config[which+"To"]);
  }
  function distance(a, b, c, d) {
    const rad = Math.PI / 180;
    const delta = Math.sin((c-a)*rad/2)**2 + Math.cos(a*rad)*Math.cos(c*rad)*Math.sin((d-b)*rad/2)**2;
    return 6371000 * 2 * Math.atan2(Math.sqrt(delta), Math.sqrt(1-delta));
  }
  function valid() {
    if (!Number.isFinite(Number(config.lat)) || Number(config.lat)<-90 || Number(config.lat)>90 ||
        !Number.isFinite(Number(config.lon)) || Number(config.lon)<-180 || Number(config.lon)>180 ||
        config.lat === "" || config.lon === "") return false;
    return config.radius >= 30 && config.radius <= 2000;
  }
  function status(message, error = false) {
    const el = $("timbratura-feedback"); if (!el) return;
    el.textContent = message; el.dataset.error = error ? "1":"0";
  }
  function makeUi() {
    if (!$("side-menu") || $("timbratura-reminder-panel")) return;
    const section = $("menu-personale-title")?.closest(".menu-section");
    if (!section) return;
    const button = document.createElement("button");
    button.type="button"; button.id="open-timbratura-reminder-btn"; button.className="btn menu-title-btn";
    button.innerHTML = '<span class="menu-item-icon" aria-hidden="true">🔔</span>Ricorda di timbrare';
    section.appendChild(button);
    const node = document.createElement("div");
    node.innerHTML = `
      <section id="timbratura-reminder-panel" class="timbratura-overlay" role="dialog" aria-modal="true" aria-label="Promemoria timbratura" hidden>
        <div class="timbratura-shell">
          <header class="timbratura-header">
            <div><h2>🔔 Ricorda di timbrare</h2><p>Suona quando sei al magazzino durante la fascia oraria.</p></div>
            <button type="button" class="btn" id="timbratura-close">Chiudi</button>
          </header>
          <label class="timbratura-enable"><input id="timbratura-enabled" type="checkbox"> Attiva promemoria GPS quando l'app è aperta</label>
          <label class="timbratura-enable"><input id="timbratura-ios-geofence" type="checkbox"> 📍🔔 Controlla GPS anche con app chiusa: avvisa solo quando entro nel magazzino durante le fasce impostate</label>
          <label class="timbratura-enable"><input id="timbratura-ios-scheduled" type="checkbox"> 🕒 Avvisi a orario fisso senza controllo GPS (alternativa)</label>
          <p class="timbratura-note">Il GPS con app chiusa richiede il permesso Posizione «Sempre» e notifiche autorizzate. iOS rileva l'ingresso nel raggio e controlla giorno e fascia oraria. Non è GPS continuo: può esserci ritardo e se rimani già dentro senza uscire e rientrare non arriva un nuovo evento. La seconda opzione invia avvisi a orario anche lontano dal magazzino.</p>
          <div class="timbratura-grid">
            <div class="timbratura-card timbratura-wide">
              <h3>📍 Posizione magazzino</h3>
              <div class="timbratura-fields">
                <label>Latitudine<input id="timbratura-lat" inputmode="decimal" placeholder="44.00000"></label>
                <label>Longitudine<input id="timbratura-lon" inputmode="decimal" placeholder="11.00000"></label>
                <label class="timbratura-wide">Raggio di rilevamento (metri)<input id="timbratura-radius" type="number" min="30" max="2000"></label>
              </div>
              <div class="timbratura-actions">
                <button type="button" class="btn" id="timbratura-here">📍 Usa posizione attuale</button>
                <button type="button" class="btn" id="timbratura-check">Verifica se sono al magazzino</button>
              </div>
            </div>
            <div class="timbratura-card">
              <h3>🟢 Entrata</h3>
              <div class="timbratura-fields">
                <label>Da<input id="timbratura-arrivalFrom" type="time"></label>
                <label>A<input id="timbratura-arrivalTo" type="time"></label>
              </div>
            </div>
            <div class="timbratura-card">
              <h3>🔵 Rientro / uscita</h3>
              <div class="timbratura-fields">
                <label>Da<input id="timbratura-returnFrom" type="time"></label>
                <label>A<input id="timbratura-returnTo" type="time"></label>
              </div>
            </div>
            <div class="timbratura-card timbratura-wide">
              <h3>Giorni della settimana</h3>
              <div id="timbratura-days" class="timbratura-days">
                ${["Dom","Lun","Mar","Mer","Gio","Ven","Sab"].map((s,i) =>
                  `<label><input type="checkbox" value="${i}"> ${s}</label>`).join("")}
              </div>
              <p class="timbratura-note">Ogni giornata prevede un solo avviso per l'entrata e uno per il rientro. Dopo «Ferma» lo stesso avviso non si ripete.</p>
            </div>
          </div>
          <div class="timbratura-actions">
            <button type="button" class="btn timbratura-primary" id="timbratura-save">Salva e attiva</button>
            <button type="button" class="btn" id="timbratura-test">🔊 Prova suono</button>
          </div>
          <p id="timbratura-feedback" role="status"></p>
          <p class="timbratura-note">Con la nuova modalità GPS iPhone, iOS controlla gli ingressi nel raggio anche quando Varga Cantieri è in background: le notifiche arrivano solo nei giorni e nelle fasce scelte. Serve «Posizione Sempre». Se scegli il precedente controllo GPS con l'app aperta, l'avviso può usare AlarmKit oppure il suono nell'app. Nessuna funzione registra la timbratura reale.</p>
        </div>
      </section>
      <section id="timbratura-ringing" class="timbratura-overlay timbratura-alarm" role="alertdialog" aria-modal="true" aria-label="Promemoria timbratura" hidden>
        <div class="timbratura-shell">
          <div class="timbratura-alarm-icon">🔔</div>
          <h2 id="timbratura-ringing-title">Ricordati di timbrare!</h2>
          <p>Sei nel magazzino durante la fascia prevista. Ricordati di timbrare.</p>
          <div class="timbratura-actions">
            <button type="button" class="btn timbratura-primary" id="timbratura-stop">✓ Ferma avviso</button>
            <button type="button" class="btn" id="timbratura-snooze">Rimanda 5 minuti</button>
          </div>
          <p class="timbratura-note">Premere Ferma non equivale a timbrare: usa il sistema presenze aziendale.</p>
        </div>
      </section>`;
    document.body.append(...Array.from(node.children));
    const easyBoxDialog = document.createElement("section");
    easyBoxDialog.id = "timbratura-open-easybox";
    easyBoxDialog.className = "timbratura-overlay";
    easyBoxDialog.setAttribute("role", "dialog");
    easyBoxDialog.setAttribute("aria-modal", "true");
    easyBoxDialog.setAttribute("aria-label", "Apri Easy-Box per timbrare");
    easyBoxDialog.hidden = true;
    easyBoxDialog.innerHTML = `
      <div class="timbratura-shell">
        <div class="timbratura-alarm-icon">🕒</div>
        <h2 id="timbratura-open-title">È ora di timbrare!</h2>
        <p>Apri l'app aziendale Easy-Box ed effettua la timbratura di persona.</p>
        <div class="timbratura-actions">
          <button class="btn timbratura-primary" type="button" id="timbratura-shortcut">⚡ Apri Easy-Box con Comandi</button>
          <button class="btn" type="button" id="timbratura-appstore">📲 Easy-Box su App Store</button>
          <button class="btn" type="button" id="timbratura-open-close">Chiudi</button>
        </div>
        <p class="timbratura-note">Per l'apertura diretta, crea nell'app Apple «Comandi Rapidi» un comando chiamato <strong>Apri Easy-Box</strong> con azione «Apri app» → «Easy-Box». In alternativa, la pagina App Store offre «Apri» se Easy-Box è installata. Non equivale a timbrare.</p>
      </div>`;
    document.body.append(easyBoxDialog);
    button.addEventListener("click", () => {
      $("menu-close-btn")?.click();
      render(); $("timbratura-reminder-panel").hidden = false;
    });
    $("timbratura-close").addEventListener("click", () => { $("timbratura-reminder-panel").hidden = true; });
    $("timbratura-save").addEventListener("click", save);
    $("timbratura-here").addEventListener("click", () => locate(true));
    $("timbratura-check").addEventListener("click", () => locate(false));
    $("timbratura-test").addEventListener("click", () => ring("test"));
    $("timbratura-stop").addEventListener("click", stop);
    $("timbratura-snooze").addEventListener("click", snooze);
    $("timbratura-shortcut").addEventListener("click", () => window.VargaTimbraturaIOS?.openEasyBoxShortcut());
    $("timbratura-appstore").addEventListener("click", () => window.VargaTimbraturaIOS?.openEasyBoxStore());
    $("timbratura-open-close").addEventListener("click", () => { $("timbratura-open-easybox").hidden = true; });
    function showEasyBoxReminder(which) {
      $("timbratura-open-title").textContent = which === "return" ? "È ora di timbrare l'USCITA!" : "È ora di timbrare l'ENTRATA!";
      $("timbratura-open-easybox").hidden = false;
    }
    window.addEventListener("varga:apri-timbratura", e => showEasyBoxReminder(e.detail?.which));
    void window.VargaTimbraturaIOS?.init().then(() => {
      const pending = window.VargaTimbraturaIOS?.pending;
      if (pending) { showEasyBoxReminder(pending); window.VargaTimbraturaIOS.clearPending(); }
    });
    if (config.iosScheduled) void window.VargaTimbraturaIOS?.sync(config, { prompt: false });
    // Il processo iOS può essere stato riavviato dal GPS prima della pagina web.
    void window.VargaTimbraturaGeofence?.status({ consumePending: true });
    window.addEventListener("pagehide", () => { silence(); stopWatching(); });
    document.addEventListener("visibilitychange", () => {
      if (document.hidden) { stopWatching(); silence(); if (ringing) { ringing=""; $("timbratura-ringing").hidden=true; } }
      else if (config.enabled && !(config.iosGeofence && window.VargaTimbraturaGeofence?.isAvailable())) startWatching();
    });
    if (config.enabled && !(config.iosGeofence && window.VargaTimbraturaGeofence?.isAvailable())) startWatching();
  }
  function render() {
    $("timbratura-enabled").checked = !!config.enabled;
    $("timbratura-ios-scheduled").checked = !!config.iosScheduled;
    $("timbratura-ios-geofence").checked = !!config.iosGeofence;
    for (const key of ["lat","lon","radius","arrivalFrom","arrivalTo","returnFrom","returnTo"]) {
      $("timbratura-"+key).value = config[key] ?? "";
    }
    $("timbratura-days").querySelectorAll("input").forEach(el => { el.checked = config.days.includes(Number(el.value)); });
    status("Impostazioni memorizzate solo su questo dispositivo.");
  }
  function save() {
    const next={...config};
    next.enabled=$("timbratura-enabled").checked;
    next.iosScheduled=$("timbratura-ios-scheduled").checked;
    next.iosGeofence=$("timbratura-ios-geofence").checked;
    // Il GPS native mode non deve mai pianificare avvisi soltanto a orario.
    if (next.iosGeofence) {
      next.iosScheduled=false;
      $("timbratura-ios-scheduled").checked=false;
    }
    next.lat=$("timbratura-lat").value.trim().replace(",",".");
    next.lon=$("timbratura-lon").value.trim().replace(",",".");
    next.radius=Number($("timbratura-radius").value);
    for(const field of ["arrivalFrom","arrivalTo","returnFrom","returnTo"]) next[field]=$("timbratura-"+field).value;
    next.days=Array.from($("timbratura-days").querySelectorAll("input:checked")).map(el=>Number(el.value));
    for(const which of ["arrival","return"]) {
      const from=clockMinutes(next[which+"From"]),to=clockMinutes(next[which+"To"]);
      if (from===null || to===null || from>=to) return status("Orari non validi: l'ora finale deve essere successiva a quella iniziale.",true);
    }
    if (!next.days.length) return status("Seleziona almeno un giorno.",true);
    const old=config; config=next;
    if((next.enabled || next.iosGeofence) && !valid()) { config=old;return status("Inserisci coordinate valide e un raggio fra 30 e 2000 metri.",true); }
    if (next.iosGeofence && next.radius < 100) { config=old;return status("Per gli avvisi GPS con app chiusa imposta almeno 100 metri (consigliati 150–300).",true); }
    try { localStorage.setItem(KEY, JSON.stringify(config)); }
    catch (_) { config=old; return status("Non riesco a salvare le impostazioni sul telefono.",true); }
    stopWatching();
    if (config.enabled && !(config.iosGeofence && window.VargaTimbraturaGeofence?.isAvailable())) {
      // Consenti l'audio con il gesto esplicito usato per salvare.
      try {
        const AudioCtor = window.AudioContext || window.webkitAudioContext;
        if (AudioCtor) {
          if (!audio) audio = new AudioCtor();
          void audio.resume().catch(() => {});
        }
      } catch (_) {}
      startWatching();
      void prepareAlarmPermission();
    }
    status(config.enabled ? "Promemoria GPS attivo mentre l'app è aperta." : "Impostazioni salvate.");
    if (window.VargaTimbraturaIOS?.isNativeIOS()) {
      void window.VargaTimbraturaIOS.sync(config, { prompt: !!config.iosScheduled }).then(result => {
        if (result.busy) return;
        if (result.error) return status("Notifiche iPhone non pianificate: " + result.error, true);
        if (result.permissionDenied) return status("Abilita le notifiche di Varga Cantieri nelle Impostazioni iPhone.", true);
        if (result.scheduled) return status("🔔 Programmati " + result.scheduled + " avvisi settimanali. Funzionano anche con l'app chiusa.");
        if (!config.enabled) status("Promemoria disattivati; notifiche future cancellate.");
      });
    } else if (config.iosScheduled) {
      status("Per gli avvisi a app chiusa installa la nuova versione iPhone di Varga Cantieri.", true);
    }
    if (window.VargaTimbraturaGeofence) {
      void window.VargaTimbraturaGeofence.sync(config).then(result => {
        if (result.error) return status("GPS background non attivato: " + result.error, true);
        if (config.iosGeofence && result.supported === false)
          return status(result.reason || "Geofencing non disponibile. Aggiorna Varga Cantieri iPhone.", true);
        if (config.iosGeofence && result.permission !== "always")
          return status("Per il GPS con app chiusa vai in Impostazioni iPhone → Varga Cantieri → Posizione → Sempre. Poi riapri l'app.", true);
        if (config.iosGeofence && !result.monitoring)
          return status("GPS in attivazione: controlla permessi, localizzazione e raggio.", true);
        if (config.iosGeofence)
          return status("📍 GPS magazzino attivo anche con app chiusa: avvisa solo entrando nel raggio, nei giorni e orari scelti.");
      });
    }
  }
  function locate(use) {
    if(!navigator.geolocation) return status("GPS non disponibile su questo dispositivo.",true);
    status("Verifico la posizione...");
    navigator.geolocation.getCurrentPosition(pos=>{
      if(use) {
        $("timbratura-lat").value=pos.coords.latitude.toFixed(7);
        $("timbratura-lon").value=pos.coords.longitude.toFixed(7);
        status("Coordinate inserite. Salvale soltanto se ti trovi effettivamente al magazzino.");
      } else {
        const lat=Number($("timbratura-lat").value.replace(",",".")),lon=Number($("timbratura-lon").value.replace(",","."));
        if(!$("timbratura-lat").value || !$("timbratura-lon").value) return status("Imposta prima il GPS del magazzino.",true);
        const d=distance(lat,lon,pos.coords.latitude,pos.coords.longitude);
        status((d<=Number($("timbratura-radius").value)?"Sei nel raggio":"Sei fuori dal raggio")+
          " · circa "+Math.round(d)+" m (precisione ±"+Math.round(pos.coords.accuracy)+" m).");
      }
    },()=>status("Non riesco a leggere il GPS: verifica i permessi.",true),
    {enableHighAccuracy:true,maximumAge:0,timeout:15000});
  }
  function startWatching() {
    if(!config.enabled || !valid() || watchId!==null || document.hidden || (config.iosGeofence && window.VargaTimbraturaGeofence?.isAvailable())) return;
    if(!navigator.geolocation) return;
    // GPS non interrogato continuamente: solo nelle due fasce e con pagina visibile.
    tick();
    windowTimer=setInterval(tick,30000);
  }
  function stopWatching() {
    if(windowTimer!==null) {clearInterval(windowTimer);windowTimer=null;}
    if(watchId!==null){navigator.geolocation.clearWatch(watchId);watchId=null;}
  }
  function tick() {
    if(document.hidden || !config.enabled) return;
    if(!inWindow("arrival") && !inWindow("return")) {
      if(watchId!==null){navigator.geolocation.clearWatch(watchId);watchId=null;}
      return;
    }
    if(watchId!==null) return;
    watchId=navigator.geolocation.watchPosition(pos=>{
      if(pos.coords.accuracy>Number(config.radius)) return; // posizione troppo imprecisa
      const meters=distance(Number(config.lat),Number(config.lon),pos.coords.latitude,pos.coords.longitude);
      if(meters>Number(config.radius)) return;
      for(const which of ["arrival","return"]) {
        const k=dayKey()+":"+which;
        if(inWindow(which) && !done[k] && Date.now()>=(waitingUntil[k]||0) && !ringing) void notify(which);
      }
    },()=>{ if(watchId!==null) navigator.geolocation.clearWatch(watchId);watchId=null;},
    {enableHighAccuracy:true,maximumAge:0,timeout:15000});
  }
  async function prepareAlarmPermission() {
    try {
      const c=window.Capacitor;
      if (c?.isNativePlatform?.() && c.getPlatform?.() === "ios" && c.isPluginAvailable?.("CapgoAlarm")) {
        const alarm=c.Plugins?.CapgoAlarm || c.registerPlugin?.("CapgoAlarm");
        const result=await alarm.requestPermissions();
        if (!result?.granted) status("Consenso suoneria iPhone negato: resta l'avviso nell'app aperta.",true);
      }
    } catch (_) {}
  }
  async function tryNativeAlarm(which) {
    const c = window.Capacitor;
    if (!c?.isNativePlatform?.() || !c?.isPluginAvailable?.("CapgoAlarm")) return false;
    const alarm = c.Plugins?.CapgoAlarm || c.registerPlugin?.("CapgoAlarm");
    if (!alarm?.getOSInfo || !alarm?.createAlarm) return false;
    const capabilities = await alarm.getOSInfo();
    if (!capabilities?.supportsNativeAlarms) return false;
    if (c.getPlatform?.() === "ios") {
      const permission = await alarm.requestPermissions();
      if (!permission?.granted) return false;
    }
    // Non è una sveglia giornaliera a ora fissa: viene creata SOLTANTO dopo il
    // rilevamento GPS, dentro la fascia. Il plugin nativo accetta ore/minuti.
    // Programmare il minuto successivo evita di perdere l'allarme mentre si salva.
    const when = new Date(Date.now() + 65000);
    const title = which === "arrival"
      ? "VARGA CANTIERI - sei in magazzino: timbra ENTRATA"
      : "VARGA CANTIERI - sei in magazzino: timbra USCITA";
    const response = await alarm.createAlarm({
      hour: when.getHours(), minute: when.getMinutes(), label: title,
      vibrate: true, skipUi: true
    });
    if (!response?.success) return false;
    status(c.getPlatform?.() === "ios"
      ? "Avviso sonoro nativo richiesto dopo la verifica GPS; fermalo tramite i comandi di iOS."
      : "Avviso richiesto ad Android: verifica in Orologio la creazione. Alcuni telefoni richiedono conferma.");
    return true;
  }
  async function notify(which) {
    if (ringing || nativePending || !inWindow(which)) return;
    const key = dayKey() + ":" + which;
    if (done[key] || Date.now() < (waitingUntil[key] || 0)) return;
    nativePending = true;
    try {
      if (await tryNativeAlarm(which)) {
        done[key] = true;
        saveAck();
        return;
      }
    } catch (error) {
      console.warn("Avviso sonoro nativo non disponibile; uso l'avviso in-app.", error);
    } finally {
      nativePending = false;
    }
    // Fallback: suono a ripetizione e stop in app solo quando è visibile.
    if (!document.hidden && !done[key]) ring(which);
  }
  function playBeep() {
    if(!ringing || document.hidden) return;
    try {
      if(!audio) audio=new (window.AudioContext||window.webkitAudioContext)();
      if(audio.state==="suspended") audio.resume().catch(()=>{});
      const o=audio.createOscillator(), g=audio.createGain();
      o.type="sine";o.frequency.setValueAtTime(880,audio.currentTime);
      g.gain.setValueAtTime(.001,audio.currentTime);
      g.gain.exponentialRampToValueAtTime(.23,audio.currentTime+.03);
      g.gain.exponentialRampToValueAtTime(.001,audio.currentTime+.42);
      o.connect(g);g.connect(audio.destination);o.start();o.stop(audio.currentTime+.45);
    } catch(_) {}
  }
  function ring(which) {
    if(ringing) return;
    ringing=which;
    $("timbratura-ringing-title").textContent=which==="arrival"?"🟢 Ricordati di timbrare l'ENTRATA!":
      which==="return"?"🔵 Ricordati di timbrare l'USCITA!":"🔔 Prova avviso acustico";
    $("timbratura-ringing").hidden=false;
    $("timbratura-stop").focus();
    playBeep();soundTimer=setInterval(playBeep,900);
    if(navigator.vibrate) navigator.vibrate([350,180,350]);
  }
  function silence() {
    if(soundTimer!==null){clearInterval(soundTimer);soundTimer=null;}
    if(navigator.vibrate)navigator.vibrate(0);
  }
  function stop() {
    const which=ringing;
    silence();ringing="";$("timbratura-ringing").hidden=true;
    if(which && which!=="test") { done[dayKey()+":"+which]=true; saveAck(); }
  }
  function snooze() {
    const which=ringing;
    silence();ringing="";$("timbratura-ringing").hidden=true;
    if(which && which!=="test") waitingUntil[dayKey()+":"+which]=Date.now()+5*60000;
  }
  if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",makeUi,{once:true});else makeUi();
})();
