(() => {
  "use strict";

  const STREAM_URL = "https://radio-romanii-italia.duckdns.org/listen/radio-romanii-italia/radio.mp3";
  const BUTTON_ID = "update-app-btn";

  const mount = () => {
    const button = document.getElementById(BUTTON_ID);
    if (!button || button.dataset.rdiPlayerReady === "true") return;

    button.dataset.rdiPlayerReady = "true";
    button.classList.add("rdi-radio-toggle");
    button.type = "button";
    button.setAttribute("aria-label", "Ascolta R.D.I.");
    button.setAttribute("title", "Ascolta R.D.I.");
    button.setAttribute("aria-pressed", "false");
    button.innerHTML = `
      <svg class="rdi-radio-crest" viewBox="0 0 36 36" aria-hidden="true" focusable="false">
        <path d="M18 2.5 31 7v9.4c0 8.1-5.3 13.7-13 17.1C10.3 30.1 5 24.5 5 16.4V7z" fill="#15243a" stroke="#e8edf4" stroke-width="1.5"/>
        <path d="M9 9h6v13H9z" fill="#244aa5"/>
        <path d="M15 9h6v13h-6z" fill="#f5d242"/>
        <path d="M21 9h6v13h-6z" fill="#d8444c"/>
        <path d="M9 22h18v4H9z" fill="#15243a"/>
        <text x="18" y="25.1" text-anchor="middle" fill="#fff" font-size="5.7" font-weight="800" font-family="Arial,sans-serif">R.D.I.</text>
      </svg>
      <i class="rdi-radio-live" aria-hidden="true"></i>
    `;

    const style = document.createElement("style");
    style.textContent = `
      #update-app-btn.rdi-radio-toggle {
        position: relative;
        box-sizing: border-box;
        display: inline-flex;
        flex-direction: column;
        align-items: center;
        justify-content: center;
        padding: 4px;
        overflow: visible;
        -webkit-tap-highlight-color: transparent;
      }
      #update-app-btn.rdi-radio-toggle .rdi-radio-crest {
        display: block;
        width: 40px;
        height: 40px;
      }
      #update-app-btn.rdi-radio-toggle .rdi-radio-live {
        display: block;
        position: absolute;
        font-style: normal;
        width: 9px;
        height: 9px;
        right: 5px;
        bottom: 5px;
        border: 1.5px solid #fff;
        border-radius: 50%;
        background: #94a3b8;
      }
      #update-app-btn.rdi-radio-toggle[aria-pressed="true"] {
        border-color: #e43d48;
        box-shadow: 0 0 0 2px rgba(228, 61, 72, .16), var(--shadow, 0 1px 4px rgba(0,0,0,.12));
      }
      #update-app-btn.rdi-radio-toggle[aria-pressed="true"] .rdi-radio-live {
        background: #e43d48;
        box-shadow: 0 0 7px rgba(228, 61, 72, .7);
      }
      #update-app-btn.rdi-radio-toggle:focus-visible {
        outline: 3px solid rgba(37, 99, 235, .35);
        outline-offset: 2px;
      }
    `;
    document.head.append(style);

    const referenceButton = document.getElementById("refresh-app-btn")
      || document.getElementById("menu-toggle-btn");
    const crest = button.querySelector(".rdi-radio-crest");
    const matchHeaderButtonSize = () => {
      if (!referenceButton || !crest) return;
      const rect = referenceButton.getBoundingClientRect();
      if (rect.width <= 0 || rect.height <= 0) return;
      button.style.width = `${rect.width}px`;
      button.style.minWidth = `${rect.width}px`;
      button.style.height = `${rect.height}px`;
      const crestSize = Math.max(22, Math.min(42, rect.height - 14));
      crest.style.width = `${crestSize}px`;
      crest.style.height = `${crestSize}px`;
    };
    matchHeaderButtonSize();
    window.addEventListener("resize", matchHeaderButtonSize, { passive: true });
    if ("ResizeObserver" in window && referenceButton) {
      new ResizeObserver(matchHeaderButtonSize).observe(referenceButton);
    }

    // Dopo una telefonata riparti dal flusso live, non dal vecchio buffer.
    const audio = document.createElement("audio");
    audio.preload = "none";
    audio.playsInline = true;
    audio.setAttribute("aria-hidden", "true");
    audio.style.display = "none";
    document.body.append(audio);

    const RETRY_DELAYS_MS = [1200, 2500, 4500, 8000, 14000, 23000];
    let wantsRadio = false;
    let needsUserGesture = false;
    let connecting = false;
    let connectionAttempts = 0;
    let reconnectTimer = null;
    let startupTimer = null;
    let stalledTimer = null;
    let mustRefreshStream = false;

    const foreground = () => document.visibilityState !== "hidden";
    const online = () => navigator.onLine !== false;
    const clearReconnectTimer = () => {
      if (reconnectTimer !== null) clearTimeout(reconnectTimer);
      reconnectTimer = null;
    };
    const clearStartupTimer = () => {
      if (startupTimer !== null) clearTimeout(startupTimer);
      startupTimer = null;
    };
    const clearStalledTimer = () => {
      if (stalledTimer !== null) clearTimeout(stalledTimer);
      stalledTimer = null;
    };
    const clearTimers = () => {
      clearReconnectTimer();
      clearStartupTimer();
      clearStalledTimer();
    };

    const showStatus = (state) => {
      const titles = {
        idle: "Ascolta R.D.I.",
        connecting: "Connessione alla diretta R.D.I.",
        buffering: "R.D.I. sta caricando la diretta",
        retrying: "R.D.I. si sta riconnettendo",
        interrupted: "R.D.I. interrotta: ripresa al ritorno nell'app",
        tap: "Tocca per riprendere R.D.I.",
        playing: "Ferma R.D.I."
      };
      const title = titles[state] || titles.idle;
      button.setAttribute("aria-label", title);
      button.title = title;
      button.setAttribute("aria-pressed", String(wantsRadio && !needsUserGesture));
      button.dataset.rdiRadioStatus = state;
      button.classList.toggle("rdi-radio-buffering", wantsRadio && state !== "playing");
    };

    // Un nuovo load scarta completamente i frammenti audio della connessione interrotta.
    const connectFresh = () => {
      if (!wantsRadio || needsUserGesture) return;
      if (!foreground() || !online()) {
        showStatus("interrupted");
        return;
      }
      clearTimers();
      connecting = true;
      mustRefreshStream = false;
      showStatus("connecting");
      try {
        audio.pause();
        audio.removeAttribute("src");
        audio.load();
        audio.src = STREAM_URL;
        audio.load();
        startupTimer = setTimeout(() => {
          if (!wantsRadio || !connecting) return;
          connecting = false;
          mustRefreshStream = true;
          scheduleRetry();
        }, 12500);
        Promise.resolve(audio.play()).catch((error) => {
          if (!wantsRadio) return;
          connecting = false;
          if (error && error.name === "NotAllowedError") {
            // Alcuni browser richiedono un altro tocco dopo il cambio di audio focus.
            needsUserGesture = true;
            clearTimers();
            showStatus("tap");
            return;
          }
          mustRefreshStream = true;
          scheduleRetry();
        });
      } catch (_) {
        connecting = false;
        mustRefreshStream = true;
        scheduleRetry();
      }
    };

    const queueReconnect = (delayMs, state = "retrying") => {
      if (!wantsRadio || needsUserGesture) return;
      clearReconnectTimer();
      if (!foreground() || !online()) {
        showStatus("interrupted");
        return;
      }
      showStatus(state);
      reconnectTimer = setTimeout(() => {
        reconnectTimer = null;
        if (wantsRadio && !needsUserGesture) connectFresh();
      }, delayMs);
    };

    function scheduleRetry() {
      clearStartupTimer();
      clearStalledTimer();
      connecting = false;
      if (!wantsRadio || needsUserGesture) return;
      const delay = RETRY_DELAYS_MS[Math.min(connectionAttempts, RETRY_DELAYS_MS.length - 1)];
      connectionAttempts += 1;
      queueReconnect(delay);
    }

    const stopRadio = () => {
      wantsRadio = false;
      needsUserGesture = false;
      connecting = false;
      mustRefreshStream = false;
      connectionAttempts = 0;
      clearTimers();
      audio.pause();
      audio.removeAttribute("src");
      audio.load();
      showStatus("idle");
    };

    const startRadio = () => {
      wantsRadio = true;
      needsUserGesture = false;
      connectionAttempts = 0;
      // play() parte direttamente dal tocco, per rispettare le regole autoplay.
      connectFresh();
    };

    button.addEventListener("click", (event) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (wantsRadio && !needsUserGesture) stopRadio();
      else startRadio();
    }, true);

    audio.addEventListener("playing", () => {
      if (!wantsRadio) {
        audio.pause();
        return;
      }
      connecting = false;
      connectionAttempts = 0;
      clearTimers();
      showStatus("playing");
    });

    audio.addEventListener("pause", () => {
      if (!wantsRadio || !audio.paused || connecting) return;
      mustRefreshStream = true;
      // Non competere con il focus audio della chiamata mentre l'app è nascosta.
      if (!foreground()) showStatus("interrupted");
      else queueReconnect(1800, "interrupted");
    });

    const detectBuffering = () => {
      if (!wantsRadio || !foreground()) return;
      showStatus("buffering");
      clearStalledTimer();
      stalledTimer = setTimeout(() => {
        stalledTimer = null;
        if (!wantsRadio || !foreground() || audio.readyState >= 3) return;
        mustRefreshStream = true;
        scheduleRetry();
      }, 8500);
    };
    audio.addEventListener("waiting", detectBuffering);
    audio.addEventListener("stalled", detectBuffering);
    audio.addEventListener("canplay", clearStalledTimer);
    audio.addEventListener("timeupdate", () => {
      if (wantsRadio && !audio.paused && audio.readyState >= 3) clearStalledTimer();
    });

    const streamFailed = () => {
      if (!wantsRadio) return;
      mustRefreshStream = true;
      scheduleRetry();
    };
    audio.addEventListener("error", streamFailed);
    audio.addEventListener("ended", streamFailed);

    const recoverAfterInterruption = () => {
      if (!wantsRadio || needsUserGesture || !foreground()) return;
      if (!online()) {
        showStatus("interrupted");
        return;
      }
      if (mustRefreshStream || audio.paused || audio.error || (!connecting && audio.readyState < 3)) {
        queueReconnect(650, "connecting");
      }
    };
    document.addEventListener("visibilitychange", () => {
      if (!foreground()) {
        // Lascia suonare la radio se Android continua a riprodurla in background.
        clearReconnectTimer();
        clearStalledTimer();
      } else {
        recoverAfterInterruption();
      }
    });
    window.addEventListener("pageshow", recoverAfterInterruption);
    window.addEventListener("online", () => {
      if (wantsRadio && !needsUserGesture) queueReconnect(500, "connecting");
    });
    window.addEventListener("offline", () => {
      clearReconnectTimer();
      clearStartupTimer();
      if (wantsRadio) showStatus("interrupted");
    });
    showStatus("idle");

  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  } else {
    mount();
  }
})();