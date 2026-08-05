(() => {
  'use strict';
  if (window.__vargaFirestoreDiagnosticsDashboardV4) return;
  window.__vargaFirestoreDiagnosticsDashboardV4 = true;

  const VERSION = '4.0.0';
  const PRICE = Object.freeze({
    currency: 'USD',
    readsPer100k: 0.03,
    writesPer100k: 0.09,
    deletesPer100k: 0.01,
    freeReadsPerDay: 50000,
    freeWritesPerDay: 20000,
    freeDeletesPerDay: 20000,
    verifiedAt: '2026-08-05',
    source: 'Google Cloud Firestore pricing',
    note: 'Stima teorica fuori quota gratuita. La console Google Cloud resta la fonte definitiva.'
  });

  const txt = (v) => String(v ?? '').trim();
  const num = (v) => Math.max(0, Number(v) || 0);
  const clamp = (v, min, max) => Math.min(max, Math.max(min, Number(v) || 0));
  const fmt = (v) => new Intl.NumberFormat('it-IT', { maximumFractionDigits: 0 }).format(num(v));
  const money = (v) => new Intl.NumberFormat('it-IT', {
    style: 'currency', currency: PRICE.currency, minimumFractionDigits: 4, maximumFractionDigits: 6
  }).format(num(v));
  const esc = (v) => txt(v).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  const today = () => new Date().toLocaleDateString('en-CA', { timeZone: 'Europe/Rome' });
  const nowIso = () => new Date().toISOString();

  function baseReport() {
    try {
      if (typeof window.VargaFirestoreOptimizerDiagnostics?.read === 'function') return window.VargaFirestoreOptimizerDiagnostics.read() || {};
      if (typeof window.VargaFirestoreDiagnostics?.read === 'function') return window.VargaFirestoreDiagnostics.read() || {};
    } catch (_) {}
    return {};
  }

  function optimizer(report) {
    const a = report.registryOptimizer || {};
    const b = report.inflightReadCoalescer || {};
    const s = a.stats || {};
    const c = b.stats || {};
    return {
      available: Boolean(a.available || b.available),
      networkGets: num(s.networkGets),
      reusedDeviceCache: num(s.reusedDeviceCache),
      reusedRecent: num(s.reusedRecent),
      reusedInFlight: num(s.reusedInFlight),
      coalesced: num(c.duplicateCallsShared),
      cacheWrites: num(s.deviceCacheWrites),
      invalidations: num(s.invalidations),
      listenerSnapshots: num(s.listenerSnapshots),
      profileWritesPassed: num(s.profileWritesPassed),
      profileWritesSkipped: num(s.profileWritesSkipped),
      avoided: num(s.reusedDeviceCache) + num(s.reusedRecent) + num(s.reusedInFlight) + num(c.duplicateCallsShared)
    };
  }

  function env() {
    const nav = performance.getEntriesByType?.('navigation')?.[0];
    const conn = navigator.connection || navigator.mozConnection || navigator.webkitConnection;
    const mem = performance.memory;
    return {
      capturedAt: nowIso(),
      url: location.href,
      screen: location.hash || location.pathname,
      online: navigator.onLine,
      visibilityState: document.visibilityState,
      standalone: Boolean(window.matchMedia?.('(display-mode: standalone)')?.matches || navigator.standalone),
      language: navigator.language,
      timezone: Intl.DateTimeFormat().resolvedOptions().timeZone,
      hardwareConcurrency: navigator.hardwareConcurrency || null,
      deviceMemoryGb: navigator.deviceMemory || null,
      connection: conn ? { effectiveType: conn.effectiveType || null, downlinkMbps: conn.downlink || null, rttMs: conn.rtt || null, saveData: Boolean(conn.saveData) } : null,
      memory: mem ? { usedJsHeapBytes: num(mem.usedJSHeapSize), totalJsHeapBytes: num(mem.totalJSHeapSize), heapLimitBytes: num(mem.jsHeapSizeLimit) } : null,
      navigation: nav ? {
        type: nav.type,
        durationMs: Math.round(num(nav.duration)),
        domInteractiveMs: Math.round(num(nav.domInteractive)),
        domContentLoadedMs: Math.round(num(nav.domContentLoadedEventEnd)),
        loadEventEndMs: Math.round(num(nav.loadEventEnd)),
        transferSizeBytes: num(nav.transferSize),
        decodedBodySizeBytes: num(nav.decodedBodySize)
      } : null,
      firebase: {
        projectId: window.firebase?.apps?.[0]?.options?.projectId || window.firebaseConfig?.projectId || '',
        sdkVersion: window.firebase?.SDK_VERSION || ''
      }
    };
  }

  function activeGroups(report) {
    const groups = {};
    Object.values(report.listenerInstances || {}).filter((x) => x?.active).forEach((x) => {
      const key = `${x.path || 'sconosciuto'}|${x.functionName || 'non attribuita'}|${x.screen || 'sconosciuta'}`;
      const g = groups[key] || { key, path: x.path || '', functionName: x.functionName || '', screen: x.screen || '', count: 0, documents: 0, deliveries: 0, ids: [] };
      g.count += 1;
      g.documents += num(x.documents);
      g.deliveries += num(x.deliveries);
      g.ids.push(x.id);
      groups[key] = g;
    });
    return Object.values(groups).sort((a, b) => b.count - a.count || b.documents - a.documents);
  }

  function hourly(report) {
    const out = Object.fromEntries(Array.from({ length: 24 }, (_, i) => [String(i).padStart(2, '0'), { documents: 0, operations: 0, listenerDocuments: 0, oneShotDocuments: 0 }]));
    (report.details || []).forEach((d) => {
      const date = new Date(d.at);
      if (Number.isNaN(date.getTime())) return;
      const hour = date.toLocaleTimeString('it-IT', { timeZone: 'Europe/Rome', hour: '2-digit', hour12: false });
      const bucket = out[hour] || (out[hour] = { documents: 0, operations: 0, listenerDocuments: 0, oneShotDocuments: 0 });
      bucket.operations += 1;
      bucket.documents += num(d.amount);
      if (d.type === 'listener-delivery') bucket.listenerDocuments += num(d.amount);
      if (d.type === 'read') bucket.oneShotDocuments += num(d.amount);
    });
    return out;
  }

  function sorted(map, metric = 'documents', limit = 10) {
    return Object.entries(map || {})
      .map(([name, value]) => [name, typeof value === 'object' ? value : { [metric]: num(value) }])
      .sort((a, b) => num(b[1]?.[metric]) - num(a[1]?.[metric]))
      .slice(0, limit);
  }

  function bounds(report) {
    let lower = 0;
    let upper = num(report.readDocuments ?? report.reads);
    let initialListener = 0;
    let updateSnapshotObserved = 0;
    let oneShot = 0;
    (report.details || []).forEach((d) => {
      if (d.type === 'read') oneShot += num(d.amount);
      if (d.type === 'listener-delivery' && d.initialDelivery) initialListener += Math.max(1, num(d.amount));
      if (d.type === 'listener-delivery' && !d.initialDelivery) updateSnapshotObserved += num(d.amount);
    });
    lower = oneShot + initialListener;
    if (!report.details?.length) lower = Math.max(0, upper - num(report.listenerDocuments));
    return {
      lowerBoundReads: lower,
      upperBoundReads: upper,
      oneShotReads: oneShot,
      initialListenerReads: initialListener,
      updateSnapshotObserved,
      explanation: 'Il limite inferiore include get() e snapshot iniziali. Il limite superiore usa tutti i documenti osservati V3. Gli aggiornamenti successivi possono essere sovrastimati perché V3 conta la dimensione completa dello snapshot.'
    };
  }

  function compute(report) {
    const opt = optimizer(report);
    const groups = activeGroups(report);
    const duplicates = groups.filter((g) => g.count > 1);
    const range = bounds(report);
    const observed = num(report.readDocuments ?? report.reads);
    const listenerDocs = num(report.listenerDocuments);
    const oneShotDocs = Math.max(0, observed - listenerDocs);
    const active = num(report.activeListeners);
    const unattributedPct = observed ? (num(report.unattributedReads) / observed) * 100 : 0;
    const avgLatency = num(report.readOperations) ? num(report.readLatencyMsTotal) / num(report.readOperations) : 0;
    const durationMinutes = Math.max(1 / 60, (Date.now() - new Date(report.startedAt || Date.now()).getTime()) / 60000);
    const readsPerMinute = observed / durationMinutes;
    const listenerShare = observed ? (listenerDocs / observed) * 100 : 0;
    const cacheUsePct = (opt.reusedDeviceCache + opt.reusedRecent + opt.networkGets)
      ? ((opt.reusedDeviceCache + opt.reusedRecent) / (opt.reusedDeviceCache + opt.reusedRecent + opt.networkGets)) * 100
      : 0;

    let score = 100;
    if (duplicates.length) score -= Math.min(30, duplicates.reduce((s, g) => s + g.count - 1, 0) * 8);
    if (active > 30) score -= 20; else if (active > 20) score -= 12; else if (active > 12) score -= 5;
    if (unattributedPct > 20) score -= 15; else if (unattributedPct > 10) score -= 7;
    if (num(report.readLatencyMsMax) > 1500) score -= 12; else if (num(report.readLatencyMsMax) > 700) score -= 6;
    if (readsPerMinute > 1000) score -= 20; else if (readsPerMinute > 300) score -= 12; else if (readsPerMinute > 100) score -= 6;
    if (num(report.detailsDropped) > 0) score -= 4;
    if (cacheUsePct >= 70) score += 4;
    score = clamp(Math.round(score), 0, 100);
    const health = score >= 90 ? { label: 'OTTIMA', icon: '🟢', tone: 'green' } : score >= 70 ? { label: 'BUONA, DA CONTROLLARE', icon: '🟡', tone: 'yellow' } : { label: 'ATTENZIONE', icon: '🔴', tone: 'red' };

    const findings = [];
    findings.push(duplicates.length
      ? { severity: 'danger', title: `${duplicates.length} gruppo/i di listener duplicati`, detail: 'Più listener attivi usano lo stesso percorso, funzione e schermata.' }
      : { severity: 'success', title: 'Nessun listener duplicato evidente', detail: 'Nel campione corrente non risultano gruppi attivi identici.' });
    findings.push(active > 20
      ? { severity: 'warning', title: `${active} listener contemporaneamente attivi`, detail: 'Verificare chiusura listener al cambio schermata e dopo reinizializzazioni.' }
      : { severity: 'success', title: `${active} listener attivi`, detail: 'Numero sotto la soglia di attenzione impostata a 20.' });
    findings.push(unattributedPct > 10
      ? { severity: 'warning', title: `${Math.round(unattributedPct)}% letture non attribuite`, detail: 'Alcuni chiamanti non vengono riconosciuti dal monitor.' }
      : { severity: 'success', title: `${Math.round(100 - unattributedPct)}% letture attribuite`, detail: 'Buona identificazione di funzioni e schermate.' });
    if (listenerShare > 70) findings.push({ severity: 'warning', title: `${Math.round(listenerShare)}% delle letture deriva dai listener`, detail: 'Controllare listener costosi e snapshot completi ripetuti.' });
    if (opt.avoided > 0) findings.push({ severity: 'success', title: `${fmt(opt.avoided)} richieste evitate`, detail: 'Cache, richieste recenti e coalescing stanno riducendo le chiamate.' });
    if (num(report.readLatencyMsMax) > 700) findings.push({ severity: 'warning', title: `Picco latenza ${fmt(report.readLatencyMsMax)} ms`, detail: 'Possibile rete lenta, cold start o query pesante.' });
    if (range.updateSnapshotObserved > 0) findings.push({ severity: 'info', title: 'V3 può sovrastimare gli aggiornamenti listener', detail: `${fmt(range.updateSnapshotObserved)} documenti appartengono a snapshot successivi al primo.` });

    const lowerCost = (range.lowerBoundReads / 100000) * PRICE.readsPer100k;
    const upperCost = (range.upperBoundReads / 100000) * PRICE.readsPer100k;
    const writeCost = (num(report.writes) / 100000) * PRICE.writesPer100k;
    const deleteCost = (num(report.deletes) / 100000) * PRICE.deletesPer100k;

    return {
      generatedAt: nowIso(), score, health, observed, listenerDocs, oneShotDocs, listenerShare,
      activeListeners: active, peakListeners: num(report.peakActiveListeners), duplicateGroups: duplicates, allListenerGroups: groups,
      unattributedPct, attributionPct: 100 - unattributedPct, avgLatencyMs: Math.round(avgLatency), maxLatencyMs: num(report.readLatencyMsMax),
      durationMinutes, readsPerMinute, cacheUsePct, optimizer: opt, bounds: range, findings,
      costs: {
        lowerBoundReadCostUsd: lowerCost,
        upperBoundReadCostUsd: upperCost,
        writeCostUsd: writeCost,
        deleteCostUsd: deleteCost,
        lowerBoundTotalUsd: lowerCost + writeCost + deleteCost,
        upperBoundTotalUsd: upperCost + writeCost + deleteCost,
        localObservedReadsVsDailyFreeQuotaPct: (observed / PRICE.freeReadsPerDay) * 100,
        note: PRICE.note
      },
      topScreens: sorted(report.screens, 'documents', 10),
      topFunctions: sorted(report.functions, 'documents', 10),
      topQueries: sorted(report.queries, 'documents', 15),
      topCollections: Object.entries(report.collections || {})
        .filter(([k]) => k.endsWith(':listener-delivery') || k.endsWith(':read'))
        .reduce((acc, [k, v]) => {
          const name = k.replace(/:(listener-delivery|read)$/, '');
          acc[name] = { documents: num(acc[name]?.documents) + num(v) };
          return acc;
        }, {}),
      hourly: hourly(report),
      sourceOfTruth: 'Questa diagnostica misura il client. Google Cloud/Firestore resta la fonte definitiva per fatturazione, letture degli indici e Security Rules.'
    };
  }

  async function storageInfo() {
    try {
      const e = await navigator.storage?.estimate?.();
      if (!e) return null;
      return { usageBytes: num(e.usage), quotaBytes: num(e.quota), usagePct: e.quota ? (e.usage / e.quota) * 100 : null };
    } catch (_) { return null; }
  }

  async function fullReport() {
    const base = baseReport();
    const summary = compute(base);
    return {
      reportVersion: 4,
      scriptVersion: VERSION,
      generatedAt: nowIso(),
      pricingAssumptions: PRICE,
      environment: env(),
      storage: await storageInfo(),
      summary,
      baseDiagnostics: base,
      diagnosticNotes: [
        'V3 conta la dimensione completa dello snapshot a ogni consegna listener e può sovrastimare gli aggiornamenti successivi.',
        'Il report V4 presenta un intervallo: limite inferiore = get() + snapshot iniziali; limite superiore = tutti i documenti osservati V3.',
        'Cache e richieste evitate derivano dai contatori degli ottimizzatori quando disponibili.',
        'Il monitor non apre query, non apre listener e non scrive su Firestore.'
      ]
    };
  }

  function css() {
    if (document.getElementById('fsd4-css')) return;
    const s = document.createElement('style');
    s.id = 'fsd4-css';
    s.textContent = `
      #firestore-diagnostics-dashboard-v4{overflow:hidden;color:#f8fafc;background:linear-gradient(145deg,#0f172a,#1e293b);border:1px solid rgba(148,163,184,.25);box-shadow:0 22px 60px rgba(15,23,42,.26)}
      #firestore-diagnostics-dashboard-v4 *{box-sizing:border-box}.fs4-head{display:flex;justify-content:space-between;gap:12px;flex-wrap:wrap}.fs4-title{margin:0;font-size:clamp(1.25rem,4vw,1.8rem)}.fs4-sub{color:#cbd5e1;margin:5px 0 0}.fs4-actions{display:flex;gap:8px;flex-wrap:wrap}.fs4-btn{border:1px solid rgba(255,255,255,.18);background:rgba(255,255,255,.08);color:#fff;border-radius:11px;padding:9px 12px;font-weight:800}
      .fs4-hero{display:grid;grid-template-columns:minmax(180px,.8fr) minmax(0,2.2fr);gap:12px;margin-top:14px}.fs4-score{display:flex;align-items:center;gap:15px;padding:17px;border-radius:20px;background:rgba(34,197,94,.15);border:1px solid rgba(34,197,94,.34)}.fs4-score.yellow{background:rgba(245,158,11,.16);border-color:rgba(245,158,11,.36)}.fs4-score.red{background:rgba(239,68,68,.17);border-color:rgba(239,68,68,.4)}.fs4-num{font-size:clamp(3rem,11vw,5rem);font-weight:950;line-height:.9;letter-spacing:-.06em}.fs4-status{font-size:1.05rem;font-weight:900}.fs4-note{font-size:.78rem;color:#cbd5e1;margin-top:4px}
      .fs4-kpis{display:grid;grid-template-columns:repeat(3,minmax(0,1fr));gap:9px}.fs4-kpi{padding:13px;border-radius:16px;background:rgba(255,255,255,.055);border:1px solid rgba(148,163,184,.2)}.fs4-kl{font-size:.72rem;color:#cbd5e1;font-weight:800;text-transform:uppercase}.fs4-kv{font-size:clamp(1.2rem,5vw,1.9rem);font-weight:950;margin-top:7px;line-height:1}.fs4-kn{font-size:.72rem;color:#94a3b8;margin-top:6px}
      .fs4-grid{display:grid;grid-template-columns:repeat(2,minmax(0,1fr));gap:12px;margin-top:12px}.fs4-panel{padding:14px;border-radius:17px;background:rgba(255,255,255,.045);border:1px solid rgba(148,163,184,.18);min-width:0}.fs4-panel h3{margin:0 0 11px;font-size:1rem}.fs4-muted{font-size:.76rem;color:#94a3b8}.fs4-donuts{display:grid;grid-template-columns:1fr 1fr;gap:10px;text-align:center}.fs4-donut{width:105px;height:105px;border-radius:50%;margin:auto;display:grid;place-items:center;position:relative}.fs4-donut:after{content:"";position:absolute;inset:13px;background:#182236;border-radius:50%}.fs4-donut strong{position:relative;z-index:1}.fs4-dlabel{font-size:.74rem;color:#cbd5e1;margin-top:6px}
      .fs4-bar{display:grid;grid-template-columns:minmax(80px,1fr) minmax(100px,2fr) auto;gap:8px;align-items:center;margin:8px 0}.fs4-bn{font-size:.75rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.fs4-track{height:9px;background:rgba(148,163,184,.18);border-radius:99px;overflow:hidden}.fs4-fill{height:100%;background:linear-gradient(90deg,#38bdf8,#818cf8);border-radius:99px;min-width:2px}.fs4-bv{font-size:.72rem;font-weight:900}.fs4-findings{display:grid;gap:7px}.fs4-find{padding:10px 11px;border-radius:12px;background:rgba(2,6,23,.35);border-left:4px solid #64748b}.fs4-find.success{border-left-color:#22c55e}.fs4-find.warning{border-left-color:#f59e0b}.fs4-find.danger{border-left-color:#ef4444}.fs4-find.info{border-left-color:#38bdf8}.fs4-find strong,.fs4-find span{display:block}.fs4-find strong{font-size:.82rem}.fs4-find span{font-size:.72rem;color:#cbd5e1;margin-top:3px}
      .fs4-listrow{display:grid;grid-template-columns:minmax(0,1fr) auto;gap:10px;padding:8px 0;border-bottom:1px solid rgba(148,163,184,.13)}.fs4-listrow:last-child{border-bottom:0}.fs4-ln{font-size:.74rem;overflow:hidden;text-overflow:ellipsis;white-space:nowrap}.fs4-lv{font-size:.72rem;font-weight:900}.fs4-timeline{height:130px;display:flex;align-items:flex-end;gap:5px;overflow-x:auto}.fs4-hour{height:100%;min-width:32px;display:grid;grid-template-rows:1fr auto;gap:5px;align-items:end}.fs4-hbar{width:100%;min-height:2px;border-radius:6px 6px 2px 2px;background:linear-gradient(#f59e0b,#ef4444)}.fs4-hl{font-size:.62rem;color:#94a3b8;text-align:center}.fs4-foot{margin-top:12px;padding:11px 13px;border-radius:13px;background:rgba(2,6,23,.4);font-size:.73rem;color:#cbd5e1;line-height:1.45}.fs4-badge{display:inline-block;padding:4px 8px;border-radius:99px;background:rgba(255,255,255,.08);margin:2px;font-weight:800}
      @media(max-width:760px){.fs4-hero,.fs4-grid{grid-template-columns:1fr}.fs4-kpis{grid-template-columns:repeat(2,minmax(0,1fr))}}
    `;
    document.head.appendChild(s);
  }

  function ensureCard() {
    css();
    let card = document.getElementById('firestore-diagnostics-dashboard-v4');
    if (card) return card;
    const root = document.getElementById('control-center-content') || document.getElementById('control-center-page');
    if (!root) return null;
    card = document.createElement('section');
    card.id = 'firestore-diagnostics-dashboard-v4';
    card.className = 'card';
    const old = document.getElementById('firestore-operation-diagnostics-card');
    old?.parentNode === root ? root.insertBefore(card, old) : root.prepend(card);
    card.addEventListener('click', async (e) => {
      const b = e.target.closest('button');
      if (!b) return;
      if (b.dataset.refresh != null) render();
      if (b.dataset.export != null) {
        const report = await fullReport();
        const url = URL.createObjectURL(new Blob([JSON.stringify(report, null, 2)], { type: 'application/json' }));
        const a = document.createElement('a');
        a.href = url; a.download = `diagnostica-firestore-${today()}-v4.json`; a.click();
        setTimeout(() => URL.revokeObjectURL(url), 1000);
      }
      if (b.dataset.reset != null && confirm('Azzerare tutta la diagnostica locale e ricaricare l’app?')) {
        try { window.VargaFirestoreDiagnostics?.reset?.(); } catch (_) {}
        location.reload();
      }
    });
    return card;
  }

  function bars(items, metric = 'documents') {
    if (!items.length) return '<p class="fs4-muted">Nessun dato.</p>';
    const max = Math.max(1, ...items.map(([, s]) => num(s?.[metric])));
    return items.map(([name, s]) => {
      const value = num(s?.[metric]);
      return `<div class="fs4-bar" title="${esc(name)}"><div class="fs4-bn">${esc(name)}</div><div class="fs4-track"><div class="fs4-fill" style="width:${clamp(value / max * 100, value ? 2 : 0, 100)}%"></div></div><div class="fs4-bv">${fmt(value)}</div></div>`;
    }).join('');
  }

  function list(items, empty) {
    if (!items.length) return `<p class="fs4-muted">${esc(empty)}</p>`;
    return items.map((x) => `<div class="fs4-listrow"><div class="fs4-ln" title="${esc(x.name)}">${esc(x.name)}</div><div class="fs4-lv">${esc(x.value)}</div></div>`).join('');
  }

  function timeline(data) {
    const hours = Array.from({ length: 24 }, (_, i) => String(i).padStart(2, '0'));
    const max = Math.max(1, ...hours.map((h) => num(data[h]?.documents)));
    return hours.map((h) => {
      const v = num(data[h]?.documents);
      return `<div class="fs4-hour" title="${h}:00 — ${fmt(v)} documenti osservati"><div class="fs4-hbar" style="height:${clamp(v / max * 100, v ? 3 : 1, 100)}%"></div><div class="fs4-hl">${h}</div></div>`;
    }).join('');
  }

  function render() {
    const card = ensureCard();
    if (!card) return;
    const report = baseReport();
    const s = compute(report);
    const collectionItems = sorted(s.topCollections, 'documents', 10);
    const duplicates = s.duplicateGroups.slice(0, 10).map((g) => ({ name: `${g.functionName} · ${g.path}`, value: `${g.count} attivi · ${fmt(g.documents)} doc` }));
    const costly = Object.values(report.listenerInstances || {}).sort((a, b) => num(b.documents) - num(a.documents)).slice(0, 10).map((x) => ({ name: `${x.functionName || 'non attribuita'} · ${x.path || 'sconosciuto'}`, value: `${fmt(x.documents)} doc · ${fmt(x.deliveries)} snapshot` }));
    const cachePct = clamp(s.cacheUsePct, 0, 100);
    const attrPct = clamp(s.attributionPct, 0, 100);
    card.innerHTML = `
      <div class="fs4-head"><div><h2 class="fs4-title">📊 Diagnostica Firestore V4</h2><p class="fs4-sub">Panoramica immediata di letture, listener, cache, costi teorici, prestazioni e punti da ottimizzare.</p></div><div class="fs4-actions"><button class="fs4-btn" data-refresh>AGGIORNA</button><button class="fs4-btn" data-export>SCARICA REPORT V4</button><button class="fs4-btn" data-reset>AZZERA</button></div></div>
      <div class="fs4-hero"><div class="fs4-score ${s.health.tone}"><div class="fs4-num">${s.score}</div><div><div class="fs4-status">${s.health.icon} ${esc(s.health.label)}</div><div class="fs4-note">Indice 0–100<br>${Math.round(s.durationMinutes * 10) / 10} minuti osservati</div></div></div><div class="fs4-kpis">
        <div class="fs4-kpi"><div class="fs4-kl">Documenti osservati</div><div class="fs4-kv">${fmt(s.observed)}</div><div class="fs4-kn">Totale V3</div></div>
        <div class="fs4-kpi"><div class="fs4-kl">Intervallo letture</div><div class="fs4-kv">${fmt(s.bounds.lowerBoundReads)}–${fmt(s.bounds.upperBoundReads)}</div><div class="fs4-kn">Stima prudente client</div></div>
        <div class="fs4-kpi"><div class="fs4-kl">Listener attivi</div><div class="fs4-kv">${fmt(s.activeListeners)}</div><div class="fs4-kn">Picco ${fmt(s.peakListeners)}</div></div>
        <div class="fs4-kpi"><div class="fs4-kl">Da listener</div><div class="fs4-kv">${Math.round(s.listenerShare)}%</div><div class="fs4-kn">${fmt(s.listenerDocs)} documenti</div></div>
        <div class="fs4-kpi"><div class="fs4-kl">Latenza massima</div><div class="fs4-kv">${fmt(s.maxLatencyMs)} ms</div><div class="fs4-kn">Media ${fmt(s.avgLatencyMs)} ms</div></div>
        <div class="fs4-kpi"><div class="fs4-kl">Costo teorico</div><div class="fs4-kv">${money(s.costs.lowerBoundTotalUsd)}–${money(s.costs.upperBoundTotalUsd)}</div><div class="fs4-kn">Fuori quota gratuita</div></div>
      </div></div>
      <div class="fs4-grid">
        <div class="fs4-panel"><h3>Cache e attribuzione</h3><div class="fs4-donuts"><div><div class="fs4-donut" style="background:conic-gradient(#22c55e ${cachePct}%,rgba(148,163,184,.22) 0)"><strong>${Math.round(cachePct)}%</strong></div><div class="fs4-dlabel">Riusi cache rilevati</div></div><div><div class="fs4-donut" style="background:conic-gradient(#38bdf8 ${attrPct}%,rgba(148,163,184,.22) 0)"><strong>${Math.round(attrPct)}%</strong></div><div class="fs4-dlabel">Letture attribuite</div></div></div><p class="fs4-muted">Richieste evitate: ${fmt(s.optimizer.avoided)} · get di rete: ${fmt(s.optimizer.networkGets)}.</p></div>
        <div class="fs4-panel"><h3>Analisi automatica</h3><div class="fs4-findings">${s.findings.slice(0, 8).map((f) => `<div class="fs4-find ${f.severity}"><strong>${esc(f.title)}</strong><span>${esc(f.detail)}</span></div>`).join('')}</div></div>
        <div class="fs4-panel"><h3>Schermate più costose</h3>${bars(s.topScreens)}</div>
        <div class="fs4-panel"><h3>Collection più lette</h3>${bars(collectionItems)}</div>
        <div class="fs4-panel"><h3>Funzioni più costose</h3>${bars(s.topFunctions)}</div>
        <div class="fs4-panel"><h3>Andamento orario</h3><div class="fs4-timeline">${timeline(s.hourly)}</div><p class="fs4-muted">Documenti osservati dal monitor per ora locale.</p></div>
        <div class="fs4-panel"><h3>Listener duplicati attivi</h3>${list(duplicates, 'Nessun duplicato evidente.')}</div>
        <div class="fs4-panel"><h3>Listener più costosi</h3>${list(costly, 'Nessun listener osservato.')}</div>
      </div>
      <div class="fs4-foot"><span class="fs4-badge">${navigator.onLine ? '🌐 Online' : '📴 Offline'}</span><span class="fs4-badge">⚡ ${fmt(s.readsPerMinute)} doc/min</span><span class="fs4-badge">🚫 ${fmt(s.optimizer.avoided)} evitate</span><span class="fs4-badge">✍️ ${fmt(report.writes)} scritture</span><span class="fs4-badge">🗑️ ${fmt(report.deletes)} eliminazioni</span><br>${esc(s.sourceOfTruth)} Prezzo letture usato: ${PRICE.currency} ${PRICE.readsPer100k}/100.000, verificato il ${PRICE.verifiedAt}.</div>`;
  }

  window.VargaFirestoreDiagnosticsDashboardV4 = { installed: true, version: VERSION, compute: () => compute(baseReport()), report: fullReport, render, pricing: PRICE };

  function init() {
    render();
    new MutationObserver(() => { if (!document.getElementById('firestore-diagnostics-dashboard-v4')) setTimeout(render, 0); }).observe(document.documentElement, { childList: true, subtree: true });
    setInterval(() => {
      const card = document.getElementById('firestore-diagnostics-dashboard-v4');
      if (card && card.offsetParent !== null) render();
    }, 2000);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
