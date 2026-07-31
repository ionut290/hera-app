"use strict";

(() => {
  if (window.__heraSquadOperatorProfileInstalled) return;
  window.__heraSquadOperatorProfileInstalled = true;

  const ROOT_SELECTORS = ["#squadre-lista", "#snow-squadre-lista"];
  const TEAM_LINE_PATTERN = /^(.*?\bSquadra\s+\d+\s*:)\s*(.+)$/i;
  const EMPTY_MEMBERS_PATTERN = /^(?:-|—|nessun(?:o| operatore)?|non assegnat[oi])$/i;
  let enhanceScheduled = false;

  const escapeMarkup = (value) => String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");

  const normalize = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLocaleLowerCase("it-IT")
    .replace(/[^a-z0-9@+]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();

  const normalizeEmail = (value) => String(value || "").trim().toLocaleLowerCase("it-IT");
  const compact = (value) => String(value || "").replace(/\s+/g, " ").trim();
  const firstValue = (...values) => values.find((value) => compact(value)) || "";
  const asArray = (value) => Array.isArray(value) ? value : [];

  function installStyles() {
    if (document.getElementById("squad-operator-profile-style")) return;
    const style = document.createElement("style");
    style.id = "squad-operator-profile-style";
    style.textContent = `
      .squad-operator-link{appearance:none;border:0;background:transparent;padding:0 1px;margin:0;color:#075985;font:inherit;font-weight:800;line-height:inherit;text-decoration:underline;text-decoration-thickness:1px;text-underline-offset:3px;cursor:pointer;border-radius:5px;vertical-align:baseline}
      .squad-operator-link:hover{color:#0369a1;background:#e0f2fe}
      .squad-operator-link:focus-visible{outline:3px solid rgba(14,165,233,.28);outline-offset:2px;background:#e0f2fe}
      .squad-operator-profile-modal{position:fixed;inset:0;z-index:10050;display:grid;place-items:center;padding:16px;background:rgba(15,23,42,.58);backdrop-filter:blur(4px)}
      .squad-operator-profile-modal.hidden{display:none!important}
      .squad-operator-profile-card{width:min(620px,100%);max-height:min(88vh,820px);overflow:auto;border:1px solid #d7e3ea;border-radius:24px;background:#fff;box-shadow:0 28px 80px rgba(15,23,42,.28)}
      .squad-operator-profile-head{position:sticky;top:0;z-index:2;display:flex;justify-content:space-between;align-items:center;gap:12px;padding:14px 16px;border-bottom:1px solid #e2e8f0;background:rgba(255,255,255,.96);backdrop-filter:blur(10px)}
      .squad-operator-profile-head h2{margin:0;font-size:1.08rem;color:#0f172a}
      .squad-operator-profile-close{width:42px;height:42px;display:grid;place-items:center;border:1px solid #cbd5e1;border-radius:999px;background:#fff;color:#0f172a;font-size:1.45rem;cursor:pointer}
      .squad-operator-profile-body{padding:18px}
      .squad-operator-hero{display:grid;grid-template-columns:82px minmax(0,1fr);gap:16px;align-items:center;margin-bottom:18px}
      .squad-operator-photo,.squad-operator-avatar{width:82px;height:82px;border-radius:22px;object-fit:cover;border:1px solid #cbd5e1;background:#eff6ff}
      .squad-operator-avatar{display:grid;place-items:center;color:#075985;font-size:1.55rem;font-weight:900}
      .squad-operator-hero h3{margin:0 0 4px;color:#0f172a;font-size:1.35rem;line-height:1.15}
      .squad-operator-hero p{margin:3px 0;color:#475569}
      .squad-operator-status{display:inline-flex;align-items:center;min-height:28px;margin-top:7px;padding:0 10px;border:1px solid #bbf7d0;border-radius:999px;background:#f0fdf4;color:#166534;font-size:.8rem;font-weight:850}
      .squad-operator-source-list{display:flex;flex-wrap:wrap;gap:7px;margin:0 0 16px}
      .squad-operator-source{display:inline-flex;align-items:center;min-height:28px;padding:0 9px;border:1px solid #bae6fd;border-radius:999px;background:#f0f9ff;color:#075985;font-size:.76rem;font-weight:800}
      .squad-operator-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:10px}
      .squad-operator-field{display:flex;flex-direction:column;gap:3px;min-width:0;padding:12px;border:1px solid #e2e8f0;border-radius:15px;background:#f8fafc;color:#0f172a;text-decoration:none}
      .squad-operator-field.clickable{border-color:#bae6fd;background:#f0f9ff;cursor:pointer}
      .squad-operator-field.clickable:hover{background:#e0f2fe}
      .squad-operator-field small{color:#64748b;font-size:.72rem;font-weight:850;text-transform:uppercase;letter-spacing:.035em}
      .squad-operator-field strong{overflow-wrap:anywhere;font-size:.94rem}
      .squad-operator-note{margin:14px 0 0;padding:11px 12px;border-left:4px solid #0ea5e9;border-radius:10px;background:#f0f9ff;color:#334155;font-size:.86rem;line-height:1.4}
      .squad-operator-empty{padding:18px;border:1px dashed #cbd5e1;border-radius:16px;background:#f8fafc;color:#475569;line-height:1.45}
      @media(max-width:560px){.squad-operator-profile-modal{padding:8px;align-items:end}.squad-operator-profile-card{max-height:92vh;border-radius:22px 22px 12px 12px}.squad-operator-profile-body{padding:14px}.squad-operator-hero{grid-template-columns:68px minmax(0,1fr)}.squad-operator-photo,.squad-operator-avatar{width:68px;height:68px;border-radius:18px}.squad-operator-grid{grid-template-columns:1fr}}
    `;
    document.head.appendChild(style);
  }

  function getPersonnelRecords() {
    return typeof personaleRecords !== "undefined" && Array.isArray(personaleRecords) ? personaleRecords : [];
  }

  function getPlatformUsers() {
    return typeof platformUsers !== "undefined" && Array.isArray(platformUsers) ? platformUsers : [];
  }

  function personDisplayName(person) {
    if (!person) return "";
    try {
      if (typeof getPersonaleDisplayName === "function") {
        const result = compact(getPersonaleDisplayName(person));
        if (result) return result;
      }
    } catch (_) {
      // Usa i campi compatibili qui sotto.
    }
    return firstValue(
      person.nomeCompleto,
      person.displayName,
      [person.cognome, person.nome].filter(Boolean).join(" "),
      [person.nome, person.cognome].filter(Boolean).join(" "),
      person.googleDisplayName
    );
  }

  function personNameCandidates(person) {
    return [
      personDisplayName(person),
      person.nomeCompleto,
      person.displayName,
      [person.cognome, person.nome].filter(Boolean).join(" "),
      [person.nome, person.cognome].filter(Boolean).join(" "),
      person.googleDisplayName
    ].map(compact).filter(Boolean);
  }

  function tokenKey(value) {
    return normalize(value).split(" ").filter(Boolean).sort().join("|");
  }

  function matchNameScore(targetName, candidateName) {
    const target = normalize(targetName);
    const candidate = normalize(candidateName);
    if (!target || !candidate) return 0;
    if (target === candidate) return 100;
    if (tokenKey(target) === tokenKey(candidate)) return 96;
    const targetTokens = new Set(target.split(" "));
    const candidateTokens = new Set(candidate.split(" "));
    const common = [...targetTokens].filter((token) => candidateTokens.has(token)).length;
    const maxSize = Math.max(targetTokens.size, candidateTokens.size);
    if (maxSize >= 2 && common === maxSize) return 92;
    if (target.length >= 6 && (candidate.includes(target) || target.includes(candidate))) return 82;
    return 0;
  }

  function findPersonnel(name) {
    const scored = getPersonnelRecords().map((person) => ({
      person,
      score: Math.max(0, ...personNameCandidates(person).map((candidate) => matchNameScore(name, candidate)))
    })).filter((item) => item.score >= 92).sort((a, b) => b.score - a.score);
    if (!scored.length) return null;
    if (scored.length > 1 && scored[0].score === scored[1].score) return null;
    return scored[0].person;
  }

  function normalizedUser(user) {
    if (!user) return null;
    try {
      if (window.HeraManagementCore?.normalizeAppUser) return window.HeraManagementCore.normalizeAppUser(user);
    } catch (_) {
      // Usa il profilo compatibile qui sotto.
    }
    const googleProvider = asArray(user.providerData).find((provider) => provider?.providerId === "google.com") || {};
    return {
      uid: compact(user.uid || user.id),
      displayName: firstValue(user.displayName, googleProvider.displayName),
      email: normalizeEmail(firstValue(user.email, googleProvider.email)),
      photoURL: firstValue(user.photoURL, googleProvider.photoURL),
      phoneNumber: firstValue(user.phoneNumber, googleProvider.phoneNumber),
      providerId: firstValue(user.providerId, googleProvider.providerId, user.linkedAuthProvider),
      emailVerified: user.emailVerified === true,
      role: firstValue(user.role, user.ruolo),
      lastLoginAt: user.lastLoginAt || user.lastSeenAt || null
    };
  }

  function findLinkedAccount(person, clickedName) {
    const users = getPlatformUsers();
    const linkedUid = compact(person?.linkedUserId);
    if (linkedUid) {
      const byUid = users.find((user) => compact(user.uid || user.id) === linkedUid);
      if (byUid) return normalizedUser(byUid);
    }

    const emails = [person?.email, person?.emailAccessoApp, person?.linkedUserEmail]
      .map(normalizeEmail).filter(Boolean);
    if (emails.length) {
      const byEmail = users.find((user) => emails.includes(normalizeEmail(user.email)));
      if (byEmail) return normalizedUser(byEmail);
    }

    const byName = users.map(normalizedUser).filter(Boolean)
      .map((user) => ({ user, score: matchNameScore(clickedName, user.displayName) }))
      .filter((item) => item.score >= 96)
      .sort((a, b) => b.score - a.score);
    return byName.length === 1 ? byName[0].user : null;
  }

  function currentAuthProfile(person, clickedName) {
    if (typeof currentUser === "undefined" || !currentUser) return null;
    const profile = normalizedUser(currentUser);
    if (!profile) return null;
    const linkedUid = compact(person?.linkedUserId);
    const emails = [person?.email, person?.emailAccessoApp, person?.linkedUserEmail].map(normalizeEmail).filter(Boolean);
    const isMatch = (linkedUid && linkedUid === profile.uid)
      || (profile.email && emails.includes(profile.email))
      || matchNameScore(clickedName, profile.displayName) >= 96;
    return isMatch ? profile : null;
  }

  function mergeAccountProfiles(primary, fallback) {
    if (!primary) return fallback;
    if (!fallback) return primary;
    return {
      ...fallback,
      ...primary,
      displayName: firstValue(primary.displayName, fallback.displayName),
      email: firstValue(primary.email, fallback.email),
      photoURL: firstValue(primary.photoURL, fallback.photoURL),
      phoneNumber: firstValue(primary.phoneNumber, fallback.phoneNumber),
      providerId: firstValue(primary.providerId, fallback.providerId),
      role: firstValue(primary.role, fallback.role),
      lastLoginAt: primary.lastLoginAt || fallback.lastLoginAt || null,
      emailVerified: primary.emailVerified === true || fallback.emailVerified === true
    };
  }

  function isManager() {
    try {
      return typeof canManageData === "function" && canManageData();
    } catch (_) {
      return false;
    }
  }

  function isCurrentOperator(person, account) {
    if (typeof currentUser === "undefined" || !currentUser) return false;
    const current = normalizedUser(currentUser);
    if (!current) return false;
    const ids = [person?.linkedUserId, account?.uid].map(compact).filter(Boolean);
    const emails = [person?.email, person?.emailAccessoApp, person?.linkedUserEmail, account?.email]
      .map(normalizeEmail).filter(Boolean);
    return (current.uid && ids.includes(current.uid)) || (current.email && emails.includes(current.email));
  }

  function formatDate(value) {
    if (!value) return "";
    if (typeof value?.toDate === "function") value = value.toDate();
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value.toLocaleDateString("it-IT");
    const text = compact(value);
    const match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
    return match ? `${match[3]}/${match[2]}/${match[1]}` : text;
  }

  function formatDateTime(value) {
    if (!value) return "";
    if (typeof value?.toDate === "function") value = value.toDate();
    const date = value instanceof Date ? value : new Date(value);
    if (!Number.isNaN(date.getTime())) return date.toLocaleString("it-IT", { dateStyle: "short", timeStyle: "short" });
    return compact(value);
  }

  function initials(value) {
    return compact(value).split(" ").filter(Boolean).slice(0, 2).map((part) => part[0]).join("").toUpperCase() || "OP";
  }

  function safeHref(type, value) {
    const clean = compact(value);
    if (!clean) return "";
    if (type === "phone") return `tel:${clean.replace(/[^+\d]/g, "")}`;
    if (type === "email") return `mailto:${encodeURIComponent(clean)}`;
    if (type === "address") return `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(clean)}`;
    return "";
  }

  function fieldMarkup(label, value, options = {}) {
    const clean = compact(value);
    if (!clean) return "";
    const href = options.type ? safeHref(options.type, clean) : "";
    const content = `<small>${escapeMarkup(label)}</small><strong>${escapeMarkup(clean)}</strong>`;
    return href
      ? `<a class="squad-operator-field clickable" href="${escapeMarkup(href)}"${options.external ? ' target="_blank" rel="noopener noreferrer"' : ""}>${content}</a>`
      : `<div class="squad-operator-field">${content}</div>`;
  }

  function ensureModal() {
    let modal = document.getElementById("squad-operator-profile-modal");
    if (modal) return modal;
    modal = document.createElement("section");
    modal.id = "squad-operator-profile-modal";
    modal.className = "squad-operator-profile-modal hidden";
    modal.setAttribute("aria-hidden", "true");
    modal.innerHTML = `
      <article class="squad-operator-profile-card" role="dialog" aria-modal="true" aria-labelledby="squad-operator-profile-title">
        <header class="squad-operator-profile-head">
          <h2 id="squad-operator-profile-title">Scheda operatore</h2>
          <button class="squad-operator-profile-close" type="button" aria-label="Chiudi scheda operatore">×</button>
        </header>
        <div class="squad-operator-profile-body"></div>
      </article>`;
    document.body.appendChild(modal);
    modal.querySelector(".squad-operator-profile-close").addEventListener("click", closeModal);
    modal.addEventListener("click", (event) => {
      if (event.target === modal) closeModal();
    });
    return modal;
  }

  function closeModal() {
    const modal = document.getElementById("squad-operator-profile-modal");
    if (!modal) return;
    modal.classList.add("hidden");
    modal.setAttribute("aria-hidden", "true");
  }

  function renderProfile(clickedName) {
    const modal = ensureModal();
    const body = modal.querySelector(".squad-operator-profile-body");
    const person = findPersonnel(clickedName);
    const linkedAccount = findLinkedAccount(person, clickedName);
    const authAccount = currentAuthProfile(person, clickedName);
    const account = mergeAccountProfiles(linkedAccount, authAccount);

    if (!person && !account) {
      body.innerHTML = `<div class="squad-operator-empty"><strong>${escapeMarkup(clickedName)}</strong><br>La scheda non è ancora collegata a un operatore dell’anagrafica o a un account dell’app. Un amministratore può completare il collegamento dalla sezione <strong>Personale</strong>.</div>`;
      return;
    }

    const displayName = firstValue(personDisplayName(person), account?.displayName, clickedName);
    const photoURL = firstValue(person?.photoURL, person?.fotoUrl, person?.fotoProfilo, account?.photoURL);
    const email = firstValue(person?.email, person?.emailAccessoApp, person?.linkedUserEmail, account?.email);
    const phone = firstValue(person?.telefono, person?.phone, person?.cellulare, account?.phoneNumber);
    const role = firstValue(person?.mansione, person?.ruolo, account?.role, "Operatore");
    const company = firstValue(person?.azienda, person?.company, person?.ditta);
    const status = firstValue(person?.stato, person?.status, "Attivo");
    const sources = [];
    if (person) sources.push("Anagrafica personale");
    if (account?.providerId === "google.com" || account?.photoURL || account?.email) sources.push("Account Google");
    if (linkedAccount || authAccount) sources.push("Account app");
    const showPrivate = isManager() || isCurrentOperator(person, account);

    const publicFields = [
      fieldMarkup("Telefono", phone, { type: "phone" }),
      fieldMarkup("Email", email, { type: "email" }),
      fieldMarkup("Mansione", role),
      fieldMarkup("Azienda", company),
      fieldMarkup("Squadra", firstValue(person?.squadra, person?.squadraAssegnata)),
      fieldMarkup("Codice operatore", firstValue(person?.codiceOperatore, person?.matricola, person?.idOperatore)),
      fieldMarkup("Account Google", account?.emailVerified ? "Email verificata" : account?.providerId === "google.com" ? "Collegato" : ""),
      fieldMarkup("Ultimo accesso", formatDateTime(account?.lastLoginAt))
    ].filter(Boolean).join("");

    const privateFields = showPrivate ? [
      fieldMarkup("Data di nascita", formatDate(firstValue(person?.dataNascita, person?.data_nascita, person?.birthday))),
      fieldMarkup("Indirizzo", firstValue(person?.indirizzo, person?.residenza, person?.domicilio, person?.address), { type: "address", external: true }),
      fieldMarkup("Contatto emergenza", firstValue(person?.contattoEmergenza, person?.emergencyContact)),
      fieldMarkup("Telefono emergenza", firstValue(person?.telefonoEmergenza, person?.emergencyPhone), { type: "phone" }),
      fieldMarkup("Patente", firstValue(person?.patente, person?.licenzaGuida)),
      fieldMarkup("Scadenza visita medica", formatDate(firstValue(person?.dataScadenzaVisitaMedica, person?.scadenzaVisitaMedica)))
    ].filter(Boolean).join("") : "";

    body.innerHTML = `
      <div class="squad-operator-hero">
        ${photoURL ? `<img class="squad-operator-photo" src="${escapeMarkup(photoURL)}" alt="Foto di ${escapeMarkup(displayName)}" referrerpolicy="no-referrer">` : `<div class="squad-operator-avatar" aria-hidden="true">${escapeMarkup(initials(displayName))}</div>`}
        <div>
          <h3>${escapeMarkup(displayName)}</h3>
          <p>${escapeMarkup(role)}${company ? ` • ${escapeMarkup(company)}` : ""}</p>
          <span class="squad-operator-status">${escapeMarkup(status)}</span>
        </div>
      </div>
      ${sources.length ? `<div class="squad-operator-source-list">${[...new Set(sources)].map((source) => `<span class="squad-operator-source">${escapeMarkup(source)}</span>`).join("")}</div>` : ""}
      <div class="squad-operator-grid">${publicFields}${privateFields}</div>
      <p class="squad-operator-note">Nome, foto ed email possono essere completati dall’account Google collegato. Telefono, indirizzo e compleanno vengono mostrati solo quando sono presenti nell’anagrafica o in una fonte autorizzata. I dati personali riservati sono visibili soltanto all’amministratore o all’operatore interessato.</p>`;
  }

  function openProfile(name) {
    installStyles();
    const modal = ensureModal();
    renderProfile(name);
    modal.classList.remove("hidden");
    modal.setAttribute("aria-hidden", "false");
    modal.querySelector(".squad-operator-profile-close")?.focus();
  }

  function splitOperatorNames(value) {
    return compact(value)
      .split(/\s*,\s*/)
      .map((name) => name.replace(/[.;]+$/, "").trim())
      .filter((name) => name && !EMPTY_MEMBERS_PATTERN.test(name));
  }

  function makeOperatorButton(name) {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "squad-operator-link";
    button.dataset.operatorName = name;
    button.textContent = name;
    button.setAttribute("aria-label", `Apri scheda operatore di ${name}`);
    return button;
  }

  function replaceMembersAfterLabel(element, labelNode, names) {
    let node = labelNode.nextSibling;
    while (node) {
      const next = node.nextSibling;
      node.remove();
      node = next;
    }
    labelNode.after(document.createTextNode(" "));
    names.forEach((name, index) => {
      if (index) element.appendChild(document.createTextNode(", "));
      element.appendChild(makeOperatorButton(name));
    });
  }

  function enhanceTeamLine(element) {
    if (!element || element.dataset.operatorLinksReady === "1") return;
    if (element.closest("#squad-operator-profile-modal")) return;
    const text = compact(element.textContent);
    const match = text.match(TEAM_LINE_PATTERN);
    if (!match || text.length > 220 || /\b(?:Giorno|Mezzi|senza pausa|Media impianti)\b/i.test(text)) return;
    const names = splitOperatorNames(match[2]);
    if (!names.length) return;

    const labelNode = [...element.querySelectorAll("strong,b")]
      .find((candidate) => /\bSquadra\s+\d+\s*:/i.test(compact(candidate.textContent)));

    if (labelNode) {
      replaceMembersAfterLabel(element, labelNode, names);
    } else {
      element.textContent = `${match[1]} `;
      names.forEach((name, index) => {
        if (index) element.appendChild(document.createTextNode(", "));
        element.appendChild(makeOperatorButton(name));
      });
    }
    element.dataset.operatorLinksReady = "1";
  }

  function enhanceRoot(root) {
    if (!root) return;
    root.querySelectorAll("p, li, div").forEach((element) => {
      if (element.querySelector("p,li,section,article")) return;
      enhanceTeamLine(element);
    });
  }

  function enhanceAll() {
    enhanceScheduled = false;
    ROOT_SELECTORS.forEach((selector) => enhanceRoot(document.querySelector(selector)));
  }

  function scheduleEnhance() {
    if (enhanceScheduled) return;
    enhanceScheduled = true;
    window.requestAnimationFrame(enhanceAll);
  }

  function start() {
    installStyles();
    ensureModal();
    enhanceAll();

    document.addEventListener("click", (event) => {
      const button = event.target.closest(".squad-operator-link");
      if (!button) return;
      event.preventDefault();
      event.stopPropagation();
      openProfile(button.dataset.operatorName || button.textContent);
    }, true);

    document.addEventListener("keydown", (event) => {
      if (event.key === "Escape") closeModal();
    });

    const observer = new MutationObserver((mutations) => {
      if (mutations.some((mutation) => [...mutation.addedNodes].some((node) => {
        if (!(node instanceof Element)) return false;
        return ROOT_SELECTORS.some((selector) => node.matches(selector) || node.closest(selector) || node.querySelector(selector));
      }))) scheduleEnhance();
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
