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
    visibilityChanges: 0,
    optimizedIntervals: 0,
    skippedHomeTickerTicks: 0,
    skippedHiddenVisualTicks: 0,
    foregroundRefreshes: 0
  };

  const originalSetInterval = window.setInterval.bind(window);
  const visualIntervals = new Set();
  let homeTickerCallback = null;

  function callbackSource(callback) {
    try { return typeof callback === "function" ? Function.prototype.toString.call(callback) : ""; }
    catch (_) { return ""; }
  }

  function isHomeMinuteTicker(callback, delay) {
    if (Number(delay) !== 60 * 1000) return false;
    const source = callbackSource(callback);
    return source.includes("renderCommesseHomeList") && !source.includes("renderCalendar");
  }

  function isVisualAnimation(callback, delay) {
    const source = callbackSource(callback);
    if (Number(delay) <= 2500 && (source.includes("showRadarFrame") || source.includes("radarFrame"))) return true;
    return false;
  }

  function shouldRunHomeBoundaryTick() {
    const now = new Date();
    const hour = now.getHours();
    const minute = now.getMinutes();
    // +SQ cambia disponibilità alle 12:00 e alle 19:00. Il rendering normale
    // di Home continua a coprire aperture, navigazione e modifiche dati.
    return minute <= 1 && (hour === 12 || hour === 19);
  }

  window.setInterval = function heraOptimizedSetInterval(callback, delay, ...args) {
    if (typeof callback !== "function") {
      return originalSetInterval(callback, delay, ...args);
    }

    if (isHomeMinuteTicker(callback, delay)) {
      state.optimizedIntervals += 1;
      homeTickerCallback = () => callback(...args);
      return originalSetInterval(() => {
        if (document.hidden) {
          state.skippedHiddenVisualTicks += 1;
          return;
        }
        if (!shouldRunHomeBoundaryTick()) {
          state.skippedHomeTickerTicks += 1;
          return;
        }
        callback(...args);
      }, delay);
    }

    if (isVisualAnimation(callback, delay)) {
      state.optimizedIntervals += 1;
      const wrapped = () => {
        if (document.hidden) {
          state.skippedHiddenVisualTicks += 1;
          return;
        }
        callback(...args);
      };
      const timer = originalSetInterval(wrapped, delay);
      visualIntervals.add(timer);
      return timer;
    }

    return originalSetInterval(callback, delay, ...args);
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
      if (!document.hidden && homeTickerCallback) {
        const list = document.getElementById("commesse-lista");
        const isVisible = list && !list.closest(".hidden, [hidden], [aria-hidden='true']");
        if (isVisible) {
          state.foregroundRefreshes += 1;
          window.requestAnimationFrame(() => {
            try { homeTickerCallback(); } catch (_) {}
          });
        }
      }
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
    version: "1.1.0",
    runIdle,
    optimizeImages,
    getState: () => ({
      ...state,
      longTaskAverageMs: state.longTasks ? state.longTaskTotalMs / state.longTasks : 0,
      background: Boolean(document.hidden),
      trackedVisualIntervals: visualIntervals.size
    })
  };
})();
