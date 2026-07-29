(() => {
  "use strict";

  const LOCK_TTL_MS = 30000;
  const locks = new Map();
  const pendingCards = new Map();
  const queueKey = "hera_fatto_ui_queue_v1";

  const normalize = (value) => String(value ?? "").trim();
  const normalizeKey = (value) => normalize(value).toLocaleLowerCase("it-IT").replace(/\s+/g, " ");
  const nowIso = () => new Date().toISOString();

  function readQueue() {
    try {
      const parsed = JSON.parse(localStorage.getItem(queueKey) || "[]");
      return Array.isArray(parsed) ? parsed : [];
    } catch (_) {
      return [];
    }
  }

  function writeQueue(items) {
    try {
      localStorage.setItem(queueKey, JSON.stringify(items.slice(-300)));
    } catch (_) {
      // La cache UI non deve mai bloccare il lavoro dell'operatore.
    }
  }

  function getImpiantoCard(button) {
    return button?.closest?.(
      "[data-impianto-id],[data-impianto-key],.impianto-card,.impianto-item,.impianto-row,.resource-item,.simple-list-item,article,li"
    ) || null;
  }

  function getButtonKey(button) {
    const card = getImpiantoCard(button);
    return normalizeKey(
      button?.dataset?.impiantoId
      || button?.dataset?.impiantoKey
      || button?.dataset?.id
      || card?.dataset?.impiantoId
      || card?.dataset?.impiantoKey
      || card?.dataset?.id
      || card?.querySelector?.("[data-impianto-id]")?.dataset?.impiantoId
      || card?.querySelector?.("[data-impianto-key]")?.dataset?.impiantoKey
      || card?.querySelector?.("h3,h4,strong,b")?.textContent
    );
  }

  function getImpiantoObjectKeys(impianto) {
    return [
      impianto?.id,
      impianto?.key,
      impianto?.sap,
      impianto?.idSap,
      impianto?.ID_SAP,
      impianto?.denominazione,
      impianto?.nome,
      impianto?.name
    ].map(normalizeKey).filter(Boolean);
  }

  function resolvePendingKey(impianto) {
    const objectKeys = getImpiantoObjectKeys(impianto);
    for (const key of objectKeys) {
      if (pendingCards.has(key)) return key;
    }
    for (const [pendingKey, pending] of pendingCards) {
      const text = normalizeKey(pending.card?.textContent);
      if (objectKeys.some((key) => text.includes(key) || key.includes(text))) return pendingKey;
    }
    return pendingCards.size === 1 ? pendingCards.keys().next().value : "";
  }

  function isFattoButton(button) {
    if (!(button instanceof HTMLElement)) return false;
    const action = normalize(button.dataset?.actionKey || button.dataset?.action || "").toLowerCase();
    const label = normalize(button.getAttribute("aria-label") || button.title || button.textContent).toLowerCase();
    if (button.classList.contains("is-completed-done")) return false;
    if (action === "whatsapp") return false;
    return action === "done" || action === "fatto" || /(^|\s)fatto(\s|$)/i.test(label);
  }

  function formatShortDate(date = new Date()) {
    return new Intl.DateTimeFormat("it-IT", { day: "2-digit", month: "2-digit" }).format(date);
  }

  function acquireLock(key, button) {
    const lockKey = key || `button:${Date.now()}`;
    const existing = locks.get(lockKey);
    if (existing && Date.now() - existing.startedAt < LOCK_TTL_MS) return false;
    locks.set(lockKey, { startedAt: Date.now(), button });
    window.setTimeout(() => releaseLock(lockKey), LOCK_TTL_MS);
    return lockKey;
  }

  function releaseLock(key) {
    const item = locks.get(key);
    if (item?.button?.isConnected) {
      item.button.disabled = false;
      item.button.removeAttribute("aria-busy");
      item.button.classList.remove("fatto-operation-pending");
    }
    locks.delete(key);
  }

  function hidePreparingOverlay() {
    const overlay = document.getElementById("whazzup-preparing-feedback");
    if (!overlay) return;
    overlay.classList.add("hidden");
    overlay.setAttribute("aria-hidden", "true");
  }

  function findFattiContainer() {
    const headings = Array.from(document.querySelectorAll("h2,h3,h4,.section-title,.list-title,.pill"));
    const heading = headings.find((node) => /^\s*fatti\s*$/i.test(normalize(node.textContent)) || /impianti\s+fatti/i.test(normalize(node.textContent)));
    if (!heading) return null;
    const section = heading.closest("section,.card,[data-list-type],.impianti-section") || heading.parentElement;
    if (!section) return null;
    return section.querySelector(".impianti-list,.simple-list,.resource-list,[data-list],ul,ol") || section;
  }

  function escapeHtml(value) {
    return String(value).replace(/[&<>'"]/g, (char) => ({
      "&": "&amp;", "<": "&lt;", ">": "&gt;", "'": "&#39;", '"': "&quot;"
    })[char]);
  }

  function createPendingGhost(card, key) {
    const target = findFattiContainer();
    if (!target || !card || target.contains(card)) return null;
    const title = normalize(card.querySelector("h3,h4,strong,b")?.textContent) || "Impianto";
    const ghost = document.createElement("div");
    ghost.className = "fatto-pending-ghost";
    ghost.dataset.fattoPendingKey = key;
    ghost.setAttribute("role", "status");
    ghost.innerHTML = `<strong>${escapeHtml(title)}</strong><small>🟡 FATTO ${formatShortDate()} · sincronizzazione in corso</small>`;
    target.prepend(ghost);
    return ghost;
  }

  function applyOptimisticUi(button, key) {
    const card = getImpiantoCard(button);
    const original = {
      html: button.innerHTML,
      disabled: button.disabled,
      cardOpacity: card?.style?.opacity || "",
      cardPointerEvents: card?.style?.pointerEvents || ""
    };

    button.setAttribute("aria-busy", "true");
    button.classList.add("fatto-operation-pending");
    button.dataset.doneLabel = `FATTO ${formatShortDate()}`;
    button.setAttribute("aria-label", `${button.dataset.doneLabel}: salvataggio in corso`);
    button.innerHTML = `<span aria-hidden="true">⚠️</span><span>${button.dataset.doneLabel}</span>`;

    // Il blocco effettivo avviene dopo la propagazione del click corrente:
    // così il listener FATTO già presente in app.js riceve sempre il primo click.
    window.setTimeout(() => {
      if (button.isConnected) button.disabled = true;
      if (card?.isConnected) {
        card.classList.add("fatto-card-pending");
        card.dataset.fattoPending = "true";
        card.style.opacity = "0.64";
        card.style.pointerEvents = "none";
      }
    }, 0);

    const ghost = createPendingGhost(card, key);
    pendingCards.set(key, { button, card, ghost, original });

    const queue = readQueue().filter((item) => item.key !== key);
    queue.push({ key, createdAt: nowIso(), status: navigator.onLine ? "saving" : "offline" });
    writeQueue(queue);
  }

  function confirmOptimisticUi(key) {
    const pending = pendingCards.get(key);
    if (pending?.ghost?.isConnected) {
      const small = pending.ghost.querySelector("small");
      if (small) small.textContent = "🟢 Salvato e sincronizzato";
      window.setTimeout(() => pending.ghost?.remove(), 900);
    }
    if (pending?.card?.isConnected) {
      pending.card.classList.remove("fatto-card-pending");
      pending.card.removeAttribute("data-fatto-pending");
      pending.card.style.opacity = pending.original.cardOpacity;
      pending.card.style.pointerEvents = pending.original.cardPointerEvents;
    }
    pendingCards.delete(key);
    writeQueue(readQueue().filter((item) => item.key !== key));
    releaseLock(key);
  }

  function rollbackOptimisticUi(key, message) {
    const pending = pendingCards.get(key);
    if (!pending) return;
    if (pending.ghost?.isConnected) pending.ghost.remove();
    if (pending.button?.isConnected) {
      pending.button.innerHTML = pending.original.html;
      pending.button.disabled = pending.original.disabled;
      pending.button.removeAttribute("aria-busy");
      pending.button.classList.remove("fatto-operation-pending");
    }
    if (pending.card?.isConnected) {
      pending.card.classList.remove("fatto-card-pending");
      pending.card.removeAttribute("data-fatto-pending");
      pending.card.style.opacity = pending.original.cardOpacity;
      pending.card.style.pointerEvents = pending.original.cardPointerEvents;
    }
    pendingCards.delete(key);
    writeQueue(readQueue().filter((item) => item.key !== key));
    releaseLock(key);
    if (message) console.warn("FATTO ripristinato:", message);
  }

  function installFunctionWrappers() {
    const originalMarkDone = window.markImpiantoDone;
    if (typeof originalMarkDone === "function" && !originalMarkDone.__fluidFattoWrapped) {
      const wrapped = function fluidMarkImpiantoDone(impianto) {
        const pendingKey = resolvePendingKey(impianto);
        let result;
        try {
          result = originalMarkDone.apply(this, arguments);
        } catch (error) {
          if (pendingKey) rollbackOptimisticUi(pendingKey, error);
          throw error;
        }
        if (result && typeof result.then === "function") {
          return result.then((value) => {
            if (pendingKey) confirmOptimisticUi(pendingKey);
            return value;
          }).catch((error) => {
            if (pendingKey) rollbackOptimisticUi(pendingKey, error);
            throw error;
          });
        }
        if (pendingKey) confirmOptimisticUi(pendingKey);
        return result;
      };
      wrapped.__fluidFattoWrapped = true;
      wrapped.__original = originalMarkDone;
      window.markImpiantoDone = wrapped;
    }

    const originalOpenWhatsApp = window.openWhatsApp;
    if (typeof originalOpenWhatsApp === "function" && !originalOpenWhatsApp.__fluidFattoWrapped) {
      const wrapped = function fluidOpenWhatsApp() {
        hidePreparingOverlay();
        return originalOpenWhatsApp.apply(this, arguments);
      };
      wrapped.__fluidFattoWrapped = true;
      wrapped.__original = originalOpenWhatsApp;
      window.openWhatsApp = wrapped;
    }
  }

  document.addEventListener("click", (event) => {
    const button = event.target?.closest?.("button");
    if (!isFattoButton(button)) return;

    const key = getButtonKey(button) || `fatto:${Date.now()}`;
    const lockKey = acquireLock(key, button);
    if (!lockKey) {
      event.preventDefault();
      event.stopImmediatePropagation();
      return;
    }

    applyOptimisticUi(button, lockKey);
    hidePreparingOverlay();
  }, true);

  window.addEventListener("online", () => {
    document.documentElement.dataset.fattoNetwork = "online";
  });
  window.addEventListener("offline", () => {
    document.documentElement.dataset.fattoNetwork = "offline";
  });

  const style = document.createElement("style");
  style.textContent = `
    .fatto-operation-pending{background:#f5c518!important;color:#171717!important;cursor:wait!important}
    .fatto-card-pending{transition:opacity .16s ease}
    .fatto-pending-ghost{display:flex;flex-direction:column;gap:3px;margin:7px 0;padding:10px 12px;border:1px solid #d4a800;border-radius:10px;background:#fff8d1;color:#27210b}
    .fatto-pending-ghost small{font-weight:700;color:#7a5d00}
    html[data-fatto-network="offline"] .fatto-pending-ghost small::after{content:" · salvato sul dispositivo"}
    @media(prefers-reduced-motion:reduce){.fatto-card-pending{transition:none}}
  `;
  document.head.appendChild(style);

  installFunctionWrappers();
  window.setTimeout(installFunctionWrappers, 0);
  window.setTimeout(installFunctionWrappers, 1000);

  const observer = new MutationObserver(() => {
    installFunctionWrappers();
    for (const [key, pending] of pendingCards) {
      if (!pending.button?.isConnected && !pending.card?.isConnected) confirmOptimisticUi(key);
    }
  });
  observer.observe(document.documentElement, { childList: true, subtree: true });

  window.HeraFattoFluid = Object.freeze({
    confirm: confirmOptimisticUi,
    rollback: rollbackOptimisticUi,
    pending: () => Array.from(pendingCards.keys()),
    queue: readQueue
  });
})();

(() => {
  "use strict";
  if (document.querySelector('script[data-password-access-manager="true"]')) return;
  const script = document.createElement("script");
  script.src = "password-access-manager.js?v=20260727a";
  script.defer = true;
  script.dataset.passwordAccessManager = "true";
  document.head.appendChild(script);
})();