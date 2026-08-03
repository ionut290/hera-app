(function () {
  "use strict";

  const COURSE_PRESETS = [
    "Primo Soccorso",
    "Antincendio",
    "Preposto",
    "ATEX"
  ];

  let activePersonId = "";
  let enhancementScheduled = false;

  const escapeHTML = (value) => String(value ?? "").replace(/[&<>'"]/g, (character) => ({
    "&": "&amp;",
    "<": "&lt;",
    ">": "&gt;",
    "'": "&#39;",
    '"': "&quot;"
  }[character]));

  const courseKey = (value) => String(value || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .replace(/\s+/g, " ")
    .toLocaleUpperCase("it-IT");

  function normalizeCourses(value) {
    let values = [];
    if (Array.isArray(value)) {
      values = value.flatMap(normalizeCourses);
    } else if (value && typeof value === "object") {
      values = Object.entries(value).flatMap(([key, item]) => {
        if (item === true || item === 1 || courseKey(item) === "SI" || courseKey(item) === "SÌ") return [key];
        if (typeof item === "string" && item.trim() && item !== key) return normalizeCourses(item);
        return [];
      });
    } else {
      values = String(value || "")
        .split(/[\n;,|]+/)
        .map((item) => item.trim())
        .filter(Boolean);
    }

    const unique = new Map();
    values.forEach((item) => {
      const clean = String(item || "").trim().replace(/\s+/g, " ");
      const key = courseKey(clean);
      if (key && !unique.has(key)) unique.set(key, clean);
    });
    return [...unique.values()];
  }

  function serializeCourses(values) {
    return normalizeCourses(values).join("; ");
  }

  function getPersonnelRecords() {
    return typeof personaleRecords !== "undefined" && Array.isArray(personaleRecords)
      ? personaleRecords
      : [];
  }

  function findPerson(id) {
    const target = String(id || "").trim();
    return target ? getPersonnelRecords().find((person) => String(person?.id || "") === target) || null : null;
  }

  function getPersonCourses(person) {
    if (!person) return [];
    return normalizeCourses(
      person.abilitazioni
      ?? person.corsi
      ?? person.formazione
      ?? person.corsiAbilitazioni
      ?? []
    );
  }

  function resolveEditorPerson(form) {
    const selected = findPerson(activePersonId);
    if (selected) return selected;

    const code = String(form.querySelector('[name="codiceOperatore"]')?.value || "").trim();
    const email = String(form.querySelector('[name="email"]')?.value || "").trim().toLowerCase();
    if (!code && !email) return null;

    return getPersonnelRecords().find((person) =>
      (code && String(person?.codiceOperatore || "").trim() === code)
      || (email && String(person?.email || "").trim().toLowerCase() === email)
    ) || null;
  }

  function installStyles() {
    if (document.getElementById("personnel-training-manager-style")) return;
    const style = document.createElement("style");
    style.id = "personnel-training-manager-style";
    style.textContent = `
      .personnel-training-section{grid-column:1/-1;border:1px solid #d8e2ec;border-radius:16px;padding:14px;background:#f8fbfd}
      .personnel-training-section h3{margin:0 0 4px;color:#123a31;font-size:1rem}
      .personnel-training-section>p{margin:0 0 12px;color:#52616d;font-size:.86rem}
      .personnel-training-presets{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:8px;margin-bottom:12px}
      .personnel-training-option{display:flex;align-items:center;gap:8px;min-height:44px;padding:8px 10px;border:1px solid #cdd9e3;border-radius:12px;background:#fff;font-weight:700;color:#183b34}
      .personnel-training-option input{width:20px;height:20px;flex:0 0 auto}
      .personnel-training-custom{display:grid;gap:6px;font-weight:700;color:#183b34}
      .personnel-training-custom textarea{width:100%;min-height:82px;resize:vertical;font:inherit;font-weight:400}
      .personnel-training-help{display:block;margin-top:6px;color:#64748b;font-size:.78rem;font-weight:400}
      .personnel-course-link{margin-top:4px}
      .personnel-course-link.is-missing{color:#a15c00}
      @media(max-width:560px){.personnel-training-presets{grid-template-columns:1fr}.personnel-training-section{padding:12px}}
    `;
    document.head.appendChild(style);
  }

  function enhancePersonnelCards() {
    document.querySelectorAll("#personnel-v2 [data-person]").forEach((card) => {
      if (card.dataset.trainingEnhanced === "1") return;
      card.dataset.trainingEnhanced = "1";

      const person = findPerson(card.dataset.person);
      const courses = getPersonCourses(person);
      const button = document.createElement("button");
      button.type = "button";
      button.className = `registry-link personnel-course-link${courses.length ? "" : " is-missing"}`;
      button.dataset.courseManage = card.dataset.person || "";
      button.textContent = courses.length
        ? `Corsi e abilitazioni: ${courses.length}`
        : "Corsi e abilitazioni: nessuno";
      button.title = courses.length ? courses.join(" • ") : "Apri la scheda per aggiungere i corsi";

      const enabledButton = card.querySelector("[data-enabled]");
      if (enabledButton) enabledButton.insertAdjacentElement("afterend", button);
      else card.appendChild(button);
    });
  }

  function enhancePersonnelEditor() {
    const form = document.getElementById("registry-editor");
    if (!form || form.dataset.trainingEnhanced === "1") return;
    if (!form.querySelector('[name="codiceOperatore"]')) return;
    form.dataset.trainingEnhanced = "1";

    const person = resolveEditorPerson(form);
    const currentCourses = getPersonCourses(person);
    const currentByKey = new Map(currentCourses.map((course) => [courseKey(course), course]));
    const presetKeys = new Set(COURSE_PRESETS.map(courseKey));
    const customCourses = currentCourses.filter((course) => !presetKeys.has(courseKey(course)));

    const section = document.createElement("section");
    section.className = "personnel-training-section";
    section.setAttribute("aria-labelledby", "personnel-training-title");
    section.innerHTML = `
      <h3 id="personnel-training-title">🎓 Corsi e abilitazioni</h3>
      <p>Seleziona i requisiti posseduti dal collega. Il controllo squadra userà questi dati insieme alle commesse abilitate.</p>
      <div class="personnel-training-presets">
        ${COURSE_PRESETS.map((course) => `
          <label class="personnel-training-option">
            <input type="checkbox" data-course-preset value="${escapeHTML(course)}" ${currentByKey.has(courseKey(course)) ? "checked" : ""}>
            <span>${escapeHTML(course)}</span>
          </label>`).join("")}
      </div>
      <label class="personnel-training-custom">
        Altri corsi
        <textarea data-custom-courses placeholder="Un corso per riga, per esempio: Lavori in quota">${escapeHTML(customCourses.join("\n"))}</textarea>
        <small class="personnel-training-help">Puoi inserire più corsi, uno per riga. I corsi già presenti non vengono eliminati automaticamente.</small>
      </label>
      <input type="hidden" name="abilitazioni" value="${escapeHTML(serializeCourses(currentCourses))}">
    `;

    const noteField = form.querySelector('textarea[name="note"]');
    if (noteField) noteField.insertAdjacentElement("beforebegin", section);
    else form.querySelector(".registry-dialog-actions")?.insertAdjacentElement("beforebegin", section);

    const hidden = section.querySelector('input[name="abilitazioni"]');
    const custom = section.querySelector("[data-custom-courses]");
    const sync = () => {
      const selected = [...section.querySelectorAll("[data-course-preset]:checked")]
        .map((input) => input.value);
      hidden.value = serializeCourses([...selected, ...normalizeCourses(custom.value)]);
    };
    section.querySelectorAll("[data-course-preset]").forEach((input) => input.addEventListener("change", sync));
    custom.addEventListener("input", sync);
    sync();
  }

  function runEnhancements() {
    enhancementScheduled = false;
    installStyles();
    enhancePersonnelCards();
    enhancePersonnelEditor();
  }

  function scheduleEnhancements() {
    if (enhancementScheduled) return;
    enhancementScheduled = true;
    window.requestAnimationFrame(runEnhancements);
  }

  document.addEventListener("click", (event) => {
    const personCard = event.target.closest("[data-person]");
    if (personCard) activePersonId = personCard.dataset.person || "";
    if (event.target.closest('[data-new="personale"]')) activePersonId = "";
  }, true);

  const observer = new MutationObserver(scheduleEnhancements);
  const start = () => {
    installStyles();
    observer.observe(document.body, { childList: true, subtree: true });
    scheduleEnhancements();
  };

  if (document.readyState === "loading") document.addEventListener("DOMContentLoaded", start, { once: true });
  else start();
})();
