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
    days: [1, 2, 3, 4, 5], tone: "pulse", snooze: 5
  };
  const $ = (id) => document.getElementById(id);
  let config = load(), watchId = null, windowTimer = null, soundTimer = null;
  let audio = null, ringing = "", waitingUntil = {}, done = {}, lastFix = null;
  let checking = false, muted = false;

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
    if (!/^([01]\\d|2[0-3]):[0-5]\\d$/.test(s)) return null;
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
          <label class="timbratura-enable"><input id="timbratura-enabled" type="checkbox"> Attiva i promemoria GPS</label>
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
          <p class="timbratura-note">Versione di prova: suono continuo e tasti Ferma/Rimanda quando l'app è aperta. Su iOS e Android, con app chiusa o schermo bloccato, il GPS e l'audio non sono garantiti finché non viene installato e verificato il modulo nativo. Non si tratta di una timbratura reale.</p>
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
    window.addEventListener("pagehide", () => { silence(); stopWatching(); });
    document.addEventListener("visibilitychange", () => {
      if (document.hidden) { stopWatching(); silence(); }
      else if (config.enabled) startWatching();
    });
    if (config.enabled) startWatching();
  }
  function render() {
    $("timbratura-enabled").checked = !!config.enabled;
    for (const key of ["lat","lon","radius","arrivalFrom","arrivalTo","returnFrom","returnTo"]) {
      $("timbratura-"+key).value = config[key] ?? "";
    }
    $("timbratura-days").querySelectorAll("input").forEach(el => { el.checked = config.days.includes(Number(el.value)); });
    status("Impostazioni memorizzate solo su questo dispositivo.");
  }
  function save() {
    const next={...config};
    next.enabled=$("timbratura-enabled").checked;
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
    if(next.enabled && !valid()) { config=old;return status("Inserisci coordinate valide e un raggio fra 30 e 2000 metri.",true); }
    try { localStorage.setItem(KEY, JSON.stringify(config)); }
    catch (_) { config=old; return status("Non riesco a salvare le impostazioni sul telefono.",true); }
    stopWatching();
    if(config.enabled) startWatching();
    status(config.enabled ? "Promemoria GPS attivo mentre l'app è aperta." : "Promemoria disattivato.");
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
    if(!config.enabled || !valid() || watchId!==null || document.hidden) return;
    if(!navigator.geolocation) return;
    // GPS non interrogato continuamente: solo nelle due fasce e con pagina visibile.
    tick();
    windowTimer=setInterval(tick,30000);
  }
  function stopWatching() {
    if(windowTimer!==null) {clearInterval(windowTimer);windowTimer=null;}
    if(watchId!==null){navigator.geolocation.clearWatch(watchId);watchId=null;}
    checking=false;
  }
  function tick() {
    if(document.hidden || !config.enabled) return;
    if(!inWindow("arrival") && !inWindow("return")) {
      if(watchId!==null){navigator.geolocation.clearWatch(watchId);watchId=null;}
      return;
    }
    if(watchId!==null) return;
    watchId=navigator.geolocation.watchPosition(pos=>{
      lastFix=pos;
      if(pos.coords.accuracy>Number(config.radius)) return; // posizione troppo imprecisa
      const meters=distance(Number(config.lat),Number(config.lon),pos.coords.latitude,pos.coords.longitude);
      if(meters>Number(config.radius)) return;
      for(const which of ["arrival","return"]) {
        const k=dayKey()+":"+which;
        if(inWindow(which) && !done[k] && Date.now()>=(waitingUntil[k]||0) && !ringing) ring(which);
      }
    },()=>{ if(watchId!==null) navigator.geolocation.clearWatch(watchId);watchId=null;},
    {enableHighAccuracy:true,maximumAge:0,timeout:15000});
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
    ringing=which; muted=false;
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
    if(which && which!=="test") done[dayKey()+":"+which]=true;
  }
  function snooze() {
    const which=ringing;
    silence();ringing="";$("timbratura-ringing").hidden=true;
    if(which && which!=="test") waitingUntil[dayKey()+":"+which]=Date.now()+5*60000;
  }
  if(document.readyState==="loading") document.addEventListener("DOMContentLoaded",makeUi,{once:true});else makeUi();
})();
