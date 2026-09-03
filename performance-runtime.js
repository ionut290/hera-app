(() => {
  "use strict";

  if (window.HeraPerformanceRuntime?.installed) return;

  const state = {
    installedAt: new Date().toISOString(),
    longTasks: 0,
    longTaskTotalMs: 0,
    longTaskMaxMs: 0,
    lazyImagesAdjusted: 0,
    idleJobsRun: 0,
    visibilityChanges: 0
  };

  function runIdle(job, timeout = 1800) {
    if (typeof job !== "function") return;
    const wrapped = () => {
      state.idleJobsRun += 1;
      try { job(); } catch (error) { console.warn("Job prestazioni non completato:", error); }
    };
    if ("requestIdleCallback" in window) {
      window.requestIdleCallback(wrapped, { timeout });
    } else {
      window.setTimeout(wrapped, Math.min(timeout, 700));
    }
  }

  function optimizeImages() {
    document.querySelectorAll("img").forEach((img) => {
      if (img.closest("#app-startup-loading")) return;
      if (img.classList.contains("leaflet-tile")) return;
      if (!img.hasAttribute("loading")) img.loading = "lazy";
      if (!img.hasAttribute("decoding")) img.decoding = "async";
      state.lazyImagesAdjusted += 1;
    });
  }

  function installLongTaskObserver() {
    if (!("PerformanceObserver" in window)) return;
    try {
      const observer = new PerformanceObserver((list) => {
        list.getEntries().forEach((entry) => {
          const duration = Number(entry.duration || 0);
          state.longTasks += 1;
          state.longTaskTotalMs += duration;
          state.longTaskMaxMs = Math.max(state.longTaskMaxMs, duration);
        });
      });
      observer.observe({ type: "longtask", buffered: true });
    } catch (_) {}
  }

  function installVisibilityState() {
    const sync = () => {
      state.visibilityChanges += 1;
      document.documentElement.classList.toggle("hera-app-background", document.hidden);
      window.__heraAppBackground = document.hidden;
    };
    document.addEventListener("visibilitychange", sync, { passive: true });
    sync();
  }

  function schedulePostStartupOptimization() {
    const run = () => {
      runIdle(optimizeImages, 1400);
      runIdle(() => {
        document.querySelectorAll("section.hidden, aside.hidden, [role='dialog'].hidden").forEach((node) => {
          if (!(node instanceof HTMLElement)) return;
          node.style.setProperty("content-visibility", "auto");
          node.style.setProperty("contain-intrinsic-size", "1px 800px");
        });
      }, 2200);
    };
    if (document.readyState === "complete") run();
    else window.addEventListener("load", run, { once: true });
  }

  installLongTaskObserver();
  installVisibilityState();
  schedulePostStartupOptimization();

  window.HeraPerformanceRuntime = {
    installed: true,
    version: "1.0.0",
    runIdle,
    optimizeImages,
    getState: () => ({
      ...state,
      longTaskAverageMs: state.longTasks ? state.longTaskTotalMs / state.longTasks : 0,
      background: Boolean(document.hidden)
    })
  };
})();
