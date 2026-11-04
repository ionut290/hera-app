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
        padding: 0;
        overflow: visible;
        -webkit-tap-highlight-color: transparent;
      }
      #update-app-btn.rdi-radio-toggle .rdi-radio-crest {
        display: block;
        width: 24px;
        height: 24px;
      }
      #update-app-btn.rdi-radio-toggle .rdi-radio-live {
        display: block;
        position: absolute;
        font-style: normal;
        width: 7px;
        height: 7px;
        right: 1px;
        bottom: 1px;
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

    const audio = document.createElement("audio");
    audio.preload = "none";
    audio.playsInline = true;
    audio.src = STREAM_URL;
    audio.setAttribute("aria-hidden", "true");
    audio.style.display = "none";
    document.body.append(audio);

    const setPlaying = (playing) => {
      button.setAttribute("aria-pressed", String(playing));
      button.setAttribute("aria-label", playing ? "Ferma R.D.I." : "Ascolta R.D.I.");
      button.title = playing ? "Ferma R.D.I." : "Ascolta R.D.I.";
    };

    button.addEventListener("click", async (event) => {
      event.preventDefault();
      event.stopImmediatePropagation();
      if (!audio.paused) {
        audio.pause();
        return;
      }
      if (audio.error) audio.load();
      button.setAttribute("aria-label", "Connessione a R.D.I. in corso");
      button.title = "Connessione a R.D.I. in corso";
      try {
        await audio.play();
      } catch (_) {
        setPlaying(false);
        button.setAttribute("aria-label", "Riprova ad ascoltare R.D.I.");
        button.title = "Riprova ad ascoltare R.D.I.";
      }
    }, true);

    audio.addEventListener("playing", () => setPlaying(true));
    audio.addEventListener("pause", () => setPlaying(false));
    audio.addEventListener("ended", () => setPlaying(false));
    audio.addEventListener("error", () => {
      setPlaying(false);
      button.setAttribute("aria-label", "Radio R.D.I. non raggiungibile");
      button.title = "Radio R.D.I. non raggiungibile";
    });
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  } else {
    mount();
  }
})();