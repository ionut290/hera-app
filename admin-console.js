(() => {
  'use strict';

  const REPORT_VERSION = '1.0.0';
  const SAMPLE_DURATION_MS = 1600;
  const EVENT_LOOP_SAMPLE_MS = 50;
  const EVENT_LOOP_SAMPLE_COUNT = 16;
  let latestReport = null;
  let running = false;

  const round = (value, digits = 1) => {
    if (!Number.isFinite(value)) return null;
    const factor = 10 ** digits;
    return Math.round(value * factor) / factor;
  };

  const bytesToMiB = (bytes) => round((Number(bytes) || 0) / 1024 / 1024, 2);

  const safePerformanceEntries = (type) => {
    try {
      return performance.getEntriesByType(type) || [];
    } catch (_) {
      return [];
    }
  };

  const makeElement = (tag, options = {}) => {
    const element = document.createElement(tag);
    if (options.className) element.className = options.className;
    if (options.id) element.id = options.id;
    if (options.text) element.textContent = options.text;
    if (options.type) element.type = options.type;
    return element;
  };

  function installDiagnosticCard() {
    if (document.getElementById('run-performance-diagnostic-btn')) return;

    const existingButton = document.getElementById('run-control-check-btn');
    const existingCard = existingButton ? existingButton.closest('.control-card') : null;
    const container = existingCard ? existingCard.parentElement : document.querySelector('.control-center-grid');
    if (!container) return;

    const card = makeElement('div', { className: 'control-card performance-diagnostic-card' });
    const title = makeElement('h3');
    const icon = makeElement('i', { className: 'fas fa-gauge-high' });
    title.append(icon, document.createTextNode(' Diagnostica prestazioni e fluidità'));

    const description = makeElement('p', {
      text: 'Analizza caricamento, risorse, memoria, DOM, blocchi del thread principale e stabilità grafica. Il controllo parte solo quando premi il pulsante.'
    });

    const actions = makeElement('div', { className: 'performance-diagnostic-actions' });
    actions.style.display = 'flex';
    actions.style.flexWrap = 'wrap';
    actions.style.gap = '8px';

    const runButton = makeElement('button', {
      id: 'run-performance-diagnostic-btn',
      className: 'billing-report-btn',
      text: 'Avvia diagnostica',
      type: 'button'
    });

    const downloadButton = makeElement('button', {
      id: 'download-performance-report-btn',
      className: 'billing-report-btn',
      text: 'Scarica report JSON',
      type: 'button'
    });
    downloadButton.disabled = true;
    downloadButton.setAttribute('aria-disabled', 'true');

    const results = makeElement('div', {
      id: 'performance-diagnostic-results',
      className: 'control-center-results',
      text: 'Nessuna diagnostica eseguita.'
    });
    results.style.marginTop = '10px';

    actions.append(runButton, downloadButton);
    card.append(title, description, actions, results);

    if (existingCard) existingCard.insertAdjacentElement('afterend', card);
    else container.append(card);

    runButton.addEventListener('click', () => runDiagnostic(runButton, downloadButton, results));
    downloadButton.addEventListener('click', downloadLatestReport);
  }

  function collectNavigationMetrics() {
    const nav = safePerformanceEntries('navigation')[0];
    if (!nav) return { supported: false };

    return {
      supported: true,
      type: nav.type || null,
      dnsMs: round(nav.domainLookupEnd - nav.domainLookupStart),
      connectionMs: round(nav.connectEnd - nav.connectStart),
      tlsMs: nav.secureConnectionStart > 0 ? round(nav.connectEnd - nav.secureConnectionStart) : null,
      requestMs: round(nav.responseStart - nav.requestStart),
      ttfbMs: round(nav.responseStart - nav.startTime),
      responseDownloadMs: round(nav.responseEnd - nav.responseStart),
      domInteractiveMs: round(nav.domInteractive - nav.startTime),
      domContentLoadedMs: round(nav.domContentLoadedEventEnd - nav.startTime),
      loadCompleteMs: round(nav.loadEventEnd - nav.startTime),
      transferSizeBytes: nav.transferSize || 0,
      encodedBodySizeBytes: nav.encodedBodySize || 0,
      decodedBodySizeBytes: nav.decodedBodySize || 0
    };
  }

  function collectPaintMetrics() {
    const paints = {};
    for (const entry of safePerformanceEntries('paint')) paints[entry.name] = round(entry.startTime);
    return {
      firstPaintMs: paints['first-paint'] ?? null,
      firstContentfulPaintMs: paints['first-contentful-paint'] ?? null
    };
  }

  function collectResourceMetrics() {
    const resources = safePerformanceEntries('resource');
    const byType = {};
    let transferSize = 0;
    let encodedBodySize = 0;
    let decodedBodySize = 0;
    let slowest = null;

    for (const resource of resources) {
      const type = resource.initiatorType || 'altro';
      byType[type] = (byType[type] || 0) + 1;
      transferSize += resource.transferSize || 0;
      encodedBodySize += resource.encodedBodySize || 0;
      decodedBodySize += resource.decodedBodySize || 0;
      if (!slowest || resource.duration > slowest.durationMs) {
        slowest = {
          name: String(resource.name || '').split('?')[0],
          type,
          durationMs: round(resource.duration),
          transferSizeBytes: resource.transferSize || 0
        };
      }
    }

    return {
      count: resources.length,
      byType,
      transferSizeBytes: transferSize,
      transferSizeMiB: bytesToMiB(transferSize),
      encodedBodySizeBytes: encodedBodySize,
      decodedBodySizeBytes: decodedBodySize,
      slowest
    };
  }

  function collectDomMetrics() {
    return {
      totalNodes: document.getElementsByTagName('*').length,
      images: document.images.length,
      scripts: document.scripts.length,
      stylesheets: document.styleSheets.length,
      buttons: document.querySelectorAll('button').length,
      forms: document.forms.length,
      iframes: document.querySelectorAll('iframe').length
    };
  }

  function collectDeviceMetrics() {
    const connection = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    const memory = performance.memory;
    return {
      online: navigator.onLine,
      userAgent: navigator.userAgent,
      language: navigator.language,
      hardwareConcurrency: navigator.hardwareConcurrency || null,
      deviceMemoryGiB: navigator.deviceMemory || null,
      viewport: { width: window.innerWidth, height: window.innerHeight, pixelRatio: window.devicePixelRatio || 1 },
      connection: connection ? {
        effectiveType: connection.effectiveType || null,
        downlinkMbps: connection.downlink || null,
        rttMs: connection.rtt || null,
        saveData: Boolean(connection.saveData)
      } : { supported: false },
      memory: memory ? {
        usedJsHeapMiB: bytesToMiB(memory.usedJSHeapSize),
        totalJsHeapMiB: bytesToMiB(memory.totalJSHeapSize),
        heapLimitMiB: bytesToMiB(memory.jsHeapSizeLimit),
        usagePercent: round((memory.usedJSHeapSize / memory.jsHeapSizeLimit) * 100, 2)
      } : { supported: false }
    };
  }

  function sampleEventLoop() {
    return new Promise((resolve) => {
      const delays = [];
      let samples = 0;
      let expected = performance.now() + EVENT_LOOP_SAMPLE_MS;

      const tick = () => {
        const now = performance.now();
        delays.push(Math.max(0, now - expected));
        samples += 1;
        if (samples >= EVENT_LOOP_SAMPLE_COUNT) {
          const total = delays.reduce((sum, value) => sum + value, 0);
          resolve({
            sampleCount: delays.length,
            averageDelayMs: round(total / delays.length, 2),
            maximumDelayMs: round(Math.max(...delays), 2)
          });
          return;
        }
        expected = now + EVENT_LOOP_SAMPLE_MS;
        window.setTimeout(tick, EVENT_LOOP_SAMPLE_MS);
      };

      window.setTimeout(tick, EVENT_LOOP_SAMPLE_MS);
    });
  }

  function observeRuntimeSignals(durationMs) {
    return new Promise((resolve) => {
      const result = {
        durationMs,
        longTasks: { supported: false, count: 0, totalDurationMs: 0, maximumDurationMs: 0 },
        layoutShift: { supported: false, cumulativeScore: 0, count: 0 }
      };
      const observers = [];

      try {
        if (window.PerformanceObserver && PerformanceObserver.supportedEntryTypes?.includes('longtask')) {
          result.longTasks.supported = true;
          const observer = new PerformanceObserver((list) => {
            for (const entry of list.getEntries()) {
              result.longTasks.count += 1;
              result.longTasks.totalDurationMs += entry.duration;
              result.longTasks.maximumDurationMs = Math.max(result.longTasks.maximumDurationMs, entry.duration);
            }
          });
          observer.observe({ type: 'longtask', buffered: true });
          observers.push(observer);
        }
      } catch (_) {
        result.longTasks.supported = false;
      }

      try {
        if (window.PerformanceObserver && PerformanceObserver.supportedEntryTypes?.includes('layout-shift')) {
          result.layoutShift.supported = true;
          const observer = new PerformanceObserver((list) => {
            for (const entry of list.getEntries()) {
              if (!entry.hadRecentInput) {
                result.layoutShift.cumulativeScore += entry.value;
                result.layoutShift.count += 1;
              }
            }
          });
          observer.observe({ type: 'layout-shift', buffered: true });
          observers.push(observer);
        }
      } catch (_) {
        result.layoutShift.supported = false;
      }

      window.setTimeout(() => {
        for (const observer of observers) observer.disconnect();
        result.longTasks.totalDurationMs = round(result.longTasks.totalDurationMs, 2);
        result.longTasks.maximumDurationMs = round(result.longTasks.maximumDurationMs, 2);
        result.layoutShift.cumulativeScore = round(result.layoutShift.cumulativeScore, 4);
        resolve(result);
      }, durationMs);
    });
  }

  function evaluateReport(report) {
    let score = 100;
    const findings = [];
    const recommendations = [];
    const nav = report.navigation;
    const resource = report.resources;
    const dom = report.dom;
    const runtime = report.runtime;
    const loop = report.eventLoop;

    const addFinding = (severity, area, message, penalty, recommendation) => {
      findings.push({ severity, area, message });
      score -= penalty;
      if (recommendation) recommendations.push(recommendation);
    };

    if (nav.supported && nav.loadCompleteMs != null) {
      if (nav.loadCompleteMs > 6000) addFinding('alta', 'caricamento', `Caricamento completo lento: ${nav.loadCompleteMs} ms.`, 18, 'Ridurre il lavoro iniziale e caricare in modo differito i moduli non essenziali.');
      else if (nav.loadCompleteMs > 3500) addFinding('media', 'caricamento', `Caricamento completo migliorabile: ${nav.loadCompleteMs} ms.`, 9, 'Controllare gli script caricati all’avvio e rinviare quelli non necessari.');
    }

    if (nav.supported && nav.ttfbMs != null) {
      if (nav.ttfbMs > 1200) addFinding('alta', 'rete', `TTFB elevato: ${nav.ttfbMs} ms.`, 12, 'Verificare hosting, cache e dimensione del documento iniziale.');
      else if (nav.ttfbMs > 600) addFinding('media', 'rete', `TTFB migliorabile: ${nav.ttfbMs} ms.`, 6, 'Controllare cache e tempi di risposta dell’hosting.');
    }

    if (resource.count > 180) addFinding('alta', 'risorse', `Sono state caricate ${resource.count} risorse.`, 14, 'Ridurre richieste duplicate e caricare i moduli solo quando la relativa sezione viene aperta.');
    else if (resource.count > 100) addFinding('media', 'risorse', `Sono state caricate ${resource.count} risorse.`, 7, 'Valutare caricamento lazy per funzioni secondarie.');

    if (resource.transferSizeMiB > 8) addFinding('alta', 'peso', `Trasferimento iniziale elevato: ${resource.transferSizeMiB} MiB.`, 14, 'Comprimere immagini e risorse e rimuovere asset non necessari dal caricamento iniziale.');
    else if (resource.transferSizeMiB > 4) addFinding('media', 'peso', `Trasferimento iniziale migliorabile: ${resource.transferSizeMiB} MiB.`, 7, 'Ottimizzare immagini, cache e suddivisione dei bundle.');

    if (dom.totalNodes > 5000) addFinding('alta', 'DOM', `DOM molto grande: ${dom.totalNodes} elementi.`, 14, 'Renderizzare solo le sezioni visibili e rimuovere nodi non più utilizzati.');
    else if (dom.totalNodes > 2500) addFinding('media', 'DOM', `DOM grande: ${dom.totalNodes} elementi.`, 7, 'Valutare rendering progressivo e liste virtualizzate.');

    if (runtime.longTasks.supported && runtime.longTasks.count > 0) {
      const penalty = runtime.longTasks.totalDurationMs > 500 ? 14 : 7;
      addFinding(runtime.longTasks.totalDurationMs > 500 ? 'alta' : 'media', 'thread principale', `${runtime.longTasks.count} blocchi lunghi, per ${runtime.longTasks.totalDurationMs} ms complessivi.`, penalty, 'Suddividere le operazioni pesanti e rinviare calcoli o rendering non urgenti.');
    }

    if (loop.maximumDelayMs > 150) addFinding('alta', 'fluidità', `Ritardo massimo event loop: ${loop.maximumDelayMs} ms.`, 12, 'Individuare funzioni sincrone pesanti, cicli estesi e rendering ripetuti.');
    else if (loop.maximumDelayMs > 60) addFinding('media', 'fluidità', `Ritardo massimo event loop: ${loop.maximumDelayMs} ms.`, 6, 'Ridurre il lavoro sincrono durante le interazioni.');

    if (runtime.layoutShift.supported && runtime.layoutShift.cumulativeScore > 0.25) addFinding('alta', 'stabilità grafica', `Layout shift elevato: ${runtime.layoutShift.cumulativeScore}.`, 10, 'Riservare spazio per immagini, pannelli e contenuti caricati in ritardo.');
    else if (runtime.layoutShift.supported && runtime.layoutShift.cumulativeScore > 0.1) addFinding('media', 'stabilità grafica', `Layout shift migliorabile: ${runtime.layoutShift.cumulativeScore}.`, 5, 'Definire dimensioni stabili per contenuti dinamici.');

    const memory = report.device.memory;
    if (memory.supported !== false && memory.usagePercent > 70) addFinding('alta', 'memoria', `Uso heap JavaScript elevato: ${memory.usagePercent}%.`, 10, 'Controllare oggetti, mappe, immagini e listener che restano in memoria.');

    score = Math.max(0, Math.min(100, Math.round(score)));
    if (!findings.length) findings.push({ severity: 'ok', area: 'generale', message: 'Non sono emerse criticità evidenti durante questo controllo.' });
    if (!recommendations.length) recommendations.push('Ripetere la diagnostica nelle schermate più pesanti e su telefono per confrontare i risultati.');

    return {
      score,
      rating: score >= 85 ? 'ottima' : score >= 70 ? 'buona' : score >= 50 ? 'da migliorare' : 'critica',
      findings,
      recommendations: [...new Set(recommendations)]
    };
  }

  async function buildReport() {
    const [eventLoop, runtime] = await Promise.all([
      sampleEventLoop(),
      observeRuntimeSignals(SAMPLE_DURATION_MS)
    ]);

    const report = {
      reportType: 'HERA_APP_PERFORMANCE_DIAGNOSTIC',
      version: REPORT_VERSION,
      generatedAt: new Date().toISOString(),
      page: { url: location.href, title: document.title, visibilityState: document.visibilityState },
      navigation: collectNavigationMetrics(),
      paint: collectPaintMetrics(),
      resources: collectResourceMetrics(),
      dom: collectDomMetrics(),
      device: collectDeviceMetrics(),
      eventLoop,
      runtime,
      limitations: [
        'I valori dipendono dal dispositivo, dalla rete, dalla schermata aperta e dal supporto del browser.',
        'La diagnostica non legge né scrive dati Firestore.',
        'Il controllo non modifica FATTO, WhatsApp/WHAZZUP o la gestione degli impianti.'
      ]
    };
    report.assessment = evaluateReport(report);
    return report;
  }

  function renderReport(report, container) {
    container.replaceChildren();

    const heading = makeElement('div', { text: `Punteggio: ${report.assessment.score}/100 — ${report.assessment.rating}` });
    heading.style.fontWeight = '700';
    heading.style.marginBottom = '8px';
    container.append(heading);

    const summary = makeElement('div', {
      text: `Caricamento: ${report.navigation.loadCompleteMs ?? 'n/d'} ms · Risorse: ${report.resources.count} (${report.resources.transferSizeMiB} MiB) · DOM: ${report.dom.totalNodes} elementi · Ritardo massimo: ${report.eventLoop.maximumDelayMs} ms`
    });
    summary.style.marginBottom = '8px';
    container.append(summary);

    const findingsTitle = makeElement('strong', { text: 'Risultati principali' });
    const findingsList = makeElement('ul');
    for (const finding of report.assessment.findings) {
      findingsList.append(makeElement('li', { text: `${finding.severity.toUpperCase()} — ${finding.area}: ${finding.message}` }));
    }

    const recommendationsTitle = makeElement('strong', { text: 'Azioni consigliate' });
    const recommendationsList = makeElement('ul');
    for (const recommendation of report.assessment.recommendations) {
      recommendationsList.append(makeElement('li', { text: recommendation }));
    }

    const note = makeElement('small', { text: 'La misurazione è locale, parte solo su richiesta e non usa Firestore.' });
    container.append(findingsTitle, findingsList, recommendationsTitle, recommendationsList, note);
  }

  async function runDiagnostic(runButton, downloadButton, results) {
    if (running) return;
    running = true;
    runButton.disabled = true;
    runButton.textContent = 'Analisi in corso…';
    results.textContent = 'Misurazione di caricamento, fluidità e stabilità grafica in corso…';

    try {
      latestReport = await buildReport();
      renderReport(latestReport, results);
      downloadButton.disabled = false;
      downloadButton.setAttribute('aria-disabled', 'false');
    } catch (error) {
      latestReport = null;
      downloadButton.disabled = true;
      downloadButton.setAttribute('aria-disabled', 'true');
      results.textContent = `Diagnostica non completata: ${error?.message || 'errore sconosciuto'}`;
      console.error('[Performance diagnostics]', error);
    } finally {
      running = false;
      runButton.disabled = false;
      runButton.textContent = 'Ripeti diagnostica';
    }
  }

  function downloadLatestReport() {
    if (!latestReport) return;
    const json = JSON.stringify(latestReport, null, 2);
    const blob = new Blob([json], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const stamp = new Date().toISOString().replace(/[:.]/g, '-');
    link.href = url;
    link.download = `hera-diagnostica-prestazioni-${stamp}.json`;
    document.body.append(link);
    link.click();
    link.remove();
    window.setTimeout(() => URL.revokeObjectURL(url), 0);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', installDiagnosticCard, { once: true });
  else installDiagnosticCard();
})();
