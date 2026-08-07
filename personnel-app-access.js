(function installPersonnelAppAccess() {
  "use strict";

  if (window.__heraPersonnelAppAccessInstalled) return;
  window.__heraPersonnelAppAccessInstalled = true;

  let currentQuery = "";
  let scheduled = false;

  const normalize = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase()
    .trim();

  const escapeHtml = (value) => String(value ?? "").replace(/[&<>'"]/g, (char) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "'": "&#39;",
    '"': "&quot;"
  }[char]));

  function firstValue(person, keys) {
    for (const key of keys) {
      const value = String(person?.[key] || "").trim();
      if (value) return value;
    }
    return "";
  }

  function displayName(person) {
    if (typeof window.getPersonaleDisplayName === "function") {
      const value = String(window.getPersonaleDisplayName(person) || "").trim();
      if (value) return value;
    }
    return [person?.cognome, person?.nome].filter(Boolean).join(" ").trim()
      || [person?.nome, person?.cognome].filter(Boolean).join(" ").trim()
      || "Operatore";
  }

  function slugFromName(value) {
    return normalize(value)
      .replace(/[^a-z0-9]+/g, ".")
      .replace(/^\.+|\.+$/g, "")
      .replace(/\.{2,}/g, ".");
  }

  function accessInfo(person) {
    const email = firstValue(person, [
      "linkedUserEmail",
      "LINKED_USER_EMAIL",
      "emailAccessoApp",
      "EMAIL_ACCESSO_APP",
      "email"
    ]).toLowerCase();

    let username = firstValue(person, ["loginUsername", "LOGIN_USERNAME"]);
    if (!username && email.endsWith("@operatori.vargacantieri.app")) {
      username = email.slice(0, -"@operatori.vargacantieri.app".length);
    }

    const displayUsername = username || (email ? "" : slugFromName(displayName(person)));
    const primary = displayUsername || email || "Non configurato";
    const showEmail = Boolean(email && normalize(email) !== normalize(primary));

    return {
      username: displayUsername,
      email,
      primary,
      showEmail,
      searchText: normalize([displayUsername, email, primary].filter(Boolean).join(" "))
    };
  }

  function findPerson(id) {
    const records = Array.isArray(window.personaleRecords) ? window.personaleRecords : [];
    return records.find((person) => String(person?.id) === String(id)) || null;
  }

  function decorateCards(root) {
    root.querySelectorAll("[data-person]").forEach((card) => {
      const person = findPerson(card.dataset.person);
      if (!person) return;
      const info = accessInfo(person);
      let block = card.querySelector("[data-person-app-access]");
      if (!block) {
        block = document.createElement("div");
        block.setAttribute("data-person-app-access", "1");
        block.className = "person-app-access";
        const enabledButton = card.querySelector("[data-enabled]");
        if (enabledButton) card.insertBefore(block, enabledButton);
        else card.appendChild(block);
      }
      block.innerHTML = `<p><strong>🔐 Accesso app:</strong> ${escapeHtml(info.primary)}</p>${info.showEmail ? `<p><strong>📧 Email:</strong> ${escapeHtml(info.email)}</p>` : ""}`;
      card.dataset.appAccessSearch = info.searchText;
    });
  }

  function applySearch(root) {
    const query = normalize(currentQuery);
    let visible = 0;
    root.querySelectorAll("[data-person]").forEach((card) => {
      const haystack = normalize(`${card.textContent || ""} ${card.dataset.appAccessSearch || ""}`);
      card.hidden = Boolean(query && !haystack.includes(query));
      if (!card.hidden) visible += 1;
    });
    const resultLabel = root.querySelector(".registry-head small");
    if (resultLabel) resultLabel.textContent = `${visible} risultati`;
  }

  function enhanceSearch(root) {
    const original = root.querySelector("input[data-search]");
    if (!original) return;
    if (original.dataset.appAccessEnhanced === "1") {
      original.value = currentQuery;
      return;
    }

    if (!currentQuery) currentQuery = original.value || "";
    const replacement = original.cloneNode(true);
    replacement.dataset.appAccessEnhanced = "1";
    replacement.placeholder = "Cerca nome, codice, email, accesso app, ruolo…";
    replacement.value = currentQuery;
    replacement.addEventListener("input", () => {
      currentQuery = replacement.value;
      applySearch(root);
    });
    original.replaceWith(replacement);
  }

  function enhance() {
    const root = document.getElementById("personnel-v2");
    if (!root) return;
    decorateCards(root);
    enhanceSearch(root);
    applySearch(root);
  }

  function scheduleEnhance() {
    if (scheduled) return;
    scheduled = true;
    requestAnimationFrame(() => {
      scheduled = false;
      enhance();
    });
  }

  const observer = new MutationObserver(scheduleEnhance);
  observer.observe(document.documentElement, { childList: true, subtree: true });

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", scheduleEnhance, { once: true });
  } else {
    scheduleEnhance();
  }

  window.HeraPersonnelAppAccess = {
    installed: true,
    refresh: scheduleEnhance,
    accessInfo
  };
})();
