(() => {
  "use strict";

  const loadedMonths = new Set();
  const loadingMonths = new Map();

  function getMonthKey() {
    if (typeof calendarVisibleMonth === "undefined" || !(calendarVisibleMonth instanceof Date)) return "";
    return `${calendarVisibleMonth.getFullYear()}-${String(calendarVisibleMonth.getMonth() + 1).padStart(2, "0")}`;
  }

  function getMonthRange() {
    if (typeof calendarVisibleMonth === "undefined" || !(calendarVisibleMonth instanceof Date)) return null;
    const year = calendarVisibleMonth.getFullYear();
    const month = calendarVisibleMonth.getMonth();
    const first = `${year}-${String(month + 1).padStart(2, "0")}-01`;
    const lastDay = new Date(year, month + 1, 0).getDate();
    const last = `${year}-${String(month + 1).padStart(2, "0")}-${String(lastDay).padStart(2, "0")}`;
    return { first, last };
  }

  function mergeReports(docs) {
    if (typeof allHoursReports === "undefined" || !Array.isArray(allHoursReports)) return;
    const merged = new Map(allHoursReports.map((report) => [String(report?.id || ""), report]));
    docs.forEach((doc) => merged.set(String(doc.id), { id: doc.id, ...doc.data() }));
    allHoursReports.splice(0, allHoursReports.length, ...merged.values());
    if (typeof hoursReportsLoaded !== "undefined") hoursReportsLoaded = true;
  }

  async function loadVisibleMonth() {
    if (typeof calendarMode === "undefined" || calendarMode !== "hours") return;
    if (typeof currentUser === "undefined" || !currentUser) return;
    if (typeof db === "undefined" || !db) return;

    const monthKey = getMonthKey();
    const range = getMonthRange();
    if (!monthKey || !range || loadedMonths.has(monthKey)) return;
    if (loadingMonths.has(monthKey)) return loadingMonths.get(monthKey);

    const task = (async () => {
      try {
        if (typeof ui !== "undefined" && ui.calendarFeedback) {
          ui.calendarFeedback.textContent = "Caricamento delle ore personali del mese...";
        }

        const query = db.collection("oreReports")
          .where("date", ">=", range.first)
          .where("date", "<=", range.last);
        const snapshot = typeof runFirestoreGetWithRetry === "function"
          ? await runFirestoreGetWithRetry(query, {
              label: `CALENDARIO ORE PERSONALI ${monthKey}`,
              timeoutMs: 10000,
              retries: 2
            })
          : await query.get();

        mergeReports(snapshot.docs || []);
        loadedMonths.add(monthKey);
        if (typeof renderCalendar === "function" && calendarMode === "hours") renderCalendar();
      } catch (error) {
        console.warn("Caricamento ore del calendario personale non riuscito", error);
        if (typeof ui !== "undefined" && ui.calendarFeedback) {
          ui.calendarFeedback.textContent = "Impossibile caricare le ore personali. Riprova.";
        }
      } finally {
        loadingMonths.delete(monthKey);
      }
    })();

    loadingMonths.set(monthKey, task);
    return task;
  }

  function scheduleLoad() {
    window.setTimeout(() => void loadVisibleMonth(), 0);
  }

  document.addEventListener("click", (event) => {
    const target = event.target instanceof Element ? event.target.closest("button") : null;
    if (!target) return;
    if (["calendar-choice-hours-btn", "calendar-hours-tab", "calendar-prev-btn", "calendar-next-btn", "calendar-today-btn"].includes(target.id)) {
      window.setTimeout(scheduleLoad, 0);
    }
  }, true);

  const observer = new MutationObserver(() => {
    const page = document.getElementById("calendar-page");
    if (page && !page.classList.contains("hidden") && typeof calendarMode !== "undefined" && calendarMode === "hours") {
      scheduleLoad();
    }
  });

  observer.observe(document.documentElement, { subtree: true, attributes: true, attributeFilter: ["class"] });
  window.addEventListener("online", scheduleLoad);
})();
