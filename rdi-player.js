(() => {
  "use strict";

  const STREAM_URL = "https://radio-romanii-italia.duckdns.org/listen/radio-romanii-italia/radio.mp3";
  const HOST_ID = "varga-rdi-mini-player";

  const mount = () => {
    if (!document.body || document.getElementById(HOST_ID)) return;

    const host = document.createElement("div");
    host.id = HOST_ID;
    host.setAttribute("aria-label", "Player Radio Românii din Italia");
    const root = host.attachShadow({ mode: "open" });

    root.innerHTML = `
      <style>
        :host {
          position: fixed;
          z-index: 999;
          right: max(88px, env(safe-area-inset-right, 0px) + 88px);
          bottom: max(16px, env(safe-area-inset-bottom, 0px) + 16px);
          display: block;
          width: min(280px, calc(100vw - 112px));
          color: #f8fafc;
          font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
          -webkit-tap-highlight-color: transparent;
        }

        * { box-sizing: border-box; }

        .player {
          display: grid;
          grid-template-columns: minmax(0, 1fr) 44px 32px;
          align-items: center;
          gap: 8px;
          min-height: 64px;
          padding: 8px 9px;
          border: 1px solid rgba(255, 255, 255, .16);
          border-radius: 17px;
          background: linear-gradient(135deg, #17212d, #0b1119 72%);
          box-shadow: 0 10px 30px rgba(0, 0, 0, .3);
        }

        .brand {
          display: flex;
          min-width: 0;
          align-items: center;
          gap: 9px;
        }

        .mark {
          display: grid;
          flex: 0 0 38px;
          width: 38px;
          height: 38px;
          place-items: center;
          border: 2px solid #f04b55;
          border-radius: 50%;
          color: #fff;
          font-size: 10px;
          font-weight: 900;
          letter-spacing: .04em;
          box-shadow: inset 0 0 0 2px rgba(255, 255, 255, .1);
        }

        .copy { display: grid; min-width: 0; gap: 3px; }
        .station {
          overflow: hidden;
          color: #fff;
          font-size: 11px;
          font-weight: 800;
          line-height: 1.2;
          text-overflow: ellipsis;
          white-space: nowrap;
        }
        .state {
          display: flex;
          align-items: center;
          gap: 5px;
          color: #bdc8d5;
          font-size: 10px;
          line-height: 1.2;
        }
        .dot {
          width: 7px;
          height: 7px;
          border-radius: 50%;
          background: #95a3b5;
        }
        .player.is-playing .dot { background: #ef4444; box-shadow: 0 0 8px rgba(239, 68, 68, .8); }
        .player.has-error .state { color: #fecaca; }
        .player.has-error .dot { background: #ef4444; }

        button {
          display: grid;
          place-items: center;
          border: 0;
          cursor: pointer;
          font: inherit;
        }
        button:focus-visible { outline: 3px solid #fca5a5; outline-offset: 2px; }
        .toggle {
          width: 44px;
          height: 44px;
          border-radius: 50%;
          background: #e43d48;
          color: #fff;
          font-size: 16px;
          font-weight: 900;
          box-shadow: 0 4px 12px rgba(228, 61, 72, .34);
        }
        .toggle:hover { background: #c92d39; }
        .collapse {
          width: 32px;
          height: 36px;
          border-radius: 10px;
          background: transparent;
          color: #cbd5e1;
          font-size: 18px;
        }
        .collapse:hover { background: rgba(255,255,255,.1); color: #fff; }

        :host([data-collapsed="true"]) { width: auto; }
        :host([data-collapsed="true"]) .player {
          grid-template-columns: 38px 44px;
          min-height: 58px;
          padding: 7px;
        }
        :host([data-collapsed="true"]) .copy,
        :host([data-collapsed="true"]) .collapse { display: none; }

        audio { display: none; }

        @media (max-width: 560px) {
          :host {
            right: max(12px, env(safe-area-inset-right, 0px) + 12px);
            bottom: max(88px, env(safe-area-inset-bottom, 0px) + 82px);
            width: min(280px, calc(100vw - 24px));
          }
        }

        @media (prefers-reduced-motion: no-preference) {
          .toggle { transition: background-color .16s ease, transform .16s ease; }
          .toggle:active { transform: scale(.96); }
        }
      </style>
      <section class="player" aria-label="R.D.I. Radio Românii din Italia">
        <div class="brand">
          <span class="mark" aria-hidden="true">R.D.I.</span>
          <span class="copy">
            <span class="station">Radio Românii din Italia</span>
            <span class="state" role="status" aria-live="polite"><i class="dot" aria-hidden="true"></i><span data-player-state>Pronta all’ascolto</span></span>
          </span>
        </div>
        <button class="toggle" type="button" data-player-toggle aria-label="Ascolta R.D.I." aria-pressed="false" title="Ascolta R.D.I.">▶</button>
        <button class="collapse" type="button" data-player-collapse aria-label="Riduci player" title="Riduci player">⌄</button>
        <audio preload="none" playsinline></audio>
      </section>
    `;

    const player = root.querySelector(".player");
    const audio = root.querySelector("audio");
    const toggle = root.querySelector("[data-player-toggle]");
    const collapse = root.querySelector("[data-player-collapse]");
    const state = root.querySelector("[data-player-state]");

    audio.src = STREAM_URL;

    const setPlaying = (playing) => {
      player.classList.toggle("is-playing", playing);
      toggle.textContent = playing ? "Ⅱ" : "▶";
      toggle.setAttribute("aria-pressed", String(playing));
      toggle.setAttribute("aria-label", playing ? "Metti in pausa R.D.I." : "Ascolta R.D.I.");
      toggle.title = playing ? "Metti in pausa" : "Ascolta R.D.I.";
    };

    toggle.addEventListener("click", async () => {
      player.classList.remove("has-error");
      if (!audio.paused) {
        audio.pause();
        return;
      }
      if (audio.error) audio.load();
      state.textContent = "Connessione alla radio…";
      try {
        await audio.play();
      } catch (_) {
        setPlaying(false);
        player.classList.add("has-error");
        state.textContent = "Tocca per riprovare";
      }
    });

    audio.addEventListener("playing", () => {
      player.classList.remove("has-error");
      state.textContent = "In diretta";
      setPlaying(true);
    });
    audio.addEventListener("waiting", () => { state.textContent = "Caricamento audio…"; });
    audio.addEventListener("pause", () => {
      setPlaying(false);
      if (!player.classList.contains("has-error")) state.textContent = "In pausa";
    });
    audio.addEventListener("error", () => {
      setPlaying(false);
      player.classList.add("has-error");
      state.textContent = "Radio non raggiungibile";
    });

    collapse.addEventListener("click", () => {
      const collapsed = host.dataset.collapsed !== "true";
      host.dataset.collapsed = String(collapsed);
      collapse.setAttribute("aria-label", collapsed ? "Espandi player" : "Riduci player");
      collapse.title = collapsed ? "Espandi player" : "Riduci player";
      collapse.textContent = collapsed ? "⌃" : "⌄";
    });

    document.body.append(host);
  };

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", mount, { once: true });
  } else {
    mount();
  }
})();
