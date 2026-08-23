(() => {
  'use strict';
  if (window.HeraEquipmentAdvisor?.installed) return;

  const VERSION = '1.0.0';
  const PROFILE_KEY = 'heraEquipmentPerformanceProfilesV1';
  const ACTIVE_KEY = 'heraEquipmentPerformanceActiveV1';
  const text = value => String(value ?? '').trim();
  const upper = value => text(value).toLocaleUpperCase('it-IT');
  const escapeHtml = value => String(value ?? '').replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));

  function readJson(key, fallback) {
    try { const raw = localStorage.getItem(key); return raw ? JSON.parse(raw) : fallback; }
    catch (_) { return fallback; }
  }
  function writeJson(key, value) {
    try { localStorage.setItem(key, JSON.stringify(value)); } catch (_) {}
  }
  function todayKey() {
    const d = new Date();
    return `${d.getFullYear()}-${String(d.getMonth()+1).padStart(2,'0')}-${String(d.getDate()).padStart(2,'0')}`;
  }
  function selectedCommessaId() {
    try { return text(selectedCommessaId || window.selectedCommessaId); }
    catch (_) { return text(window.selectedCommessaId); }
  }
  function rows(source, commessaId) {
    if (Array.isArray(source)) return source;
    if (source instanceof Map) {
      const direct = source.get(commessaId);
      if (Array.isArray(direct)) return direct;
      if (direct) return [direct];
      return [...source.values()].flatMap(v => Array.isArray(v) ? v : [v]);
    }
    if (source && typeof source === 'object') {
      const direct = source[commessaId];
      if (Array.isArray(direct)) return direct;
      if (direct) return [direct];
      return Object.values(source).flatMap(v => Array.isArray(v) ? v : [v]);
    }
    return [];
  }
  function collectStrings(value, out = [], depth = 0) {
    if (depth > 5 || value == null) return out;
    if (typeof value === 'string' || typeof value === 'number') { out.push(String(value)); return out; }
    if (Array.isArray(value)) { value.forEach(v => collectStrings(v, out, depth + 1)); return out; }
    if (typeof value === 'object') Object.entries(value).forEach(([k,v]) => {
      if (/mez|attrezz|macchin|tratt|trinc|decesp|strument|daily|iveco|soffiat|tagliasiep|motoseg|spazzatr|piattaform|mezzo/i.test(k)) collectStrings(v, out, depth + 1);
    });
    return out;
  }
  function countOperators(squad) {
    for (const value of [squad?.operatori,squad?.operators,squad?.membri,squad?.members,squad?.persone,squad?.componenti,squad?.utenti]) {
      if (Array.isArray(value) && value.filter(Boolean).length) return value.filter(Boolean).length;
    }
    for (const value of [squad?.numeroOperatori,squad?.operatorCount,squad?.teamSize,squad?.numeroPersone]) {
      const n = Number(value); if (Number.isFinite(n) && n > 0) return Math.round(n);
    }
    return 2;
  }
  function detectTeam() {
    const commessaId = selectedCommessaId();
    const today = todayKey();
    const sources = [];
    try { if (typeof currentSquadre !== 'undefined') sources.push(currentSquadre); } catch (_) {}
    sources.push(window.currentSquadre, window.squadre, window.squadreOggi, window.todaySquads, window.squadreByCommessa);
    let best = null;
    for (const source of sources) {
      for (const squad of rows(source, commessaId)) {
        if (!squad || typeof squad !== 'object') continue;
        const cid = text(squad.commessaId ?? squad.idCommessa ?? squad.commessa);
        const date = text(squad.data ?? squad.date ?? squad.giorno).slice(0,10);
        if (cid && commessaId && cid !== commessaId) continue;
        if (date && date !== today) continue;
        const equipmentText = upper(collectStrings(squad).join(' '));
        const bigMatch = equipmentText.match(/(?:^|\s)T\s*[-_.]?\s*\d{1,5}(?:\s|$)/);
        const smallMatch = equipmentText.match(/(?:^|\s)R\s*[-_.]?\s*\d{1,5}(?:\s|$)/);
        const dailyMatch = equipmentText.match(/(?:^|\s)A\s*[-_.]?\s*\d{1,5}(?:\s|$)/);
        const info = {
          teamSize: countOperators(squad), equipmentText,
          hasBigTrincia: Boolean(bigMatch), hasSmallTrincia: Boolean(smallMatch),
          hasTrincia: Boolean(bigMatch || smallMatch || /TRINCIA|TRINCIATR/.test(equipmentText)),
          hasDaily: Boolean(dailyMatch || /DAILY|IVECO/.test(equipmentText)),
          hasDecespugliatore: /DECESPUGLI/.test(equipmentText),
          hasSoffiatore: /SOFFIAT/.test(equipmentText),
          hasTagliasiepe: /TAGLIASIEP/.test(equipmentText),
          hasMotosega: /MOTOSEG/.test(equipmentText),
          hasSpazzatrice: /SPAZZATR/.test(equipmentText),
          hasPiattaforma: /PIATTAFORM/.test(equipmentText),
          bigCode: bigMatch ? bigMatch[0].trim().replace(/\s+/g,'') : '',
          smallCode: smallMatch ? smallMatch[0].trim().replace(/\s+/g,'') : '',
          dailyCode: dailyMatch ? dailyMatch[0].trim().replace(/\s+/g,'') : ''
        };
        best = info;
      }
    }
    return best || {teamSize:2,equipmentText:'',hasBigTrincia:false,hasSmallTrincia:false,hasTrincia:false,hasDaily:false,hasDecespugliatore:false,hasSoffiatore:false,hasTagliasiepe:false,hasMotosega:false,hasSpazzatrice:false,hasPiattaforma:false,bigCode:'',smallCode:'',dailyCode:''};
  }

  function machineClass(team) {
    if (team.hasBigTrincia) return 'T';
    if (team.hasSmallTrincia) return 'R';
    return team.hasTrincia ? 'TRINCIA' : 'MANUALE';
  }
  function machineLabel(team) {
    if (team.hasBigTrincia) return `🚜 ${team.bigCode || 'T'} · trattore grande con trincia`;
    if (team.hasSmallTrincia) return `🚜 ${team.smallCode || 'R'} · trattore piccolo con trincia`;
    if (team.hasTrincia) return '🚜 Trincia presente';
    if (team.hasDecespugliatore) return '🧑‍🌾 Decespugliatore presente';
    return '🧑‍🌾 Nessuna trincia rilevata';
  }

  function plantId(item) { return text(item?.physicalPlantId || item?.impiantoId || item?.migrationSourceId || item?.id || item?.idSap || item?.idSAP || item?.sap || item?.denominazione || item?.nome); }
  function currentPlants() {
    try { if (typeof currentImpianti !== 'undefined' && Array.isArray(currentImpianti)) return currentImpianti; } catch (_) {}
    return Array.isArray(window.currentImpianti) ? window.currentImpianti : [];
  }
  function workText(item) {
    const parts = [item?.tipologiaIntervento,item?.tipologiaLavorazione,item?.lavorazione,item?.attivitaRichiesta,item?.descrizioneLavoro,item?.descrizione,item?.noteLavoro,item?.note,item?.noteImpianto,item?.modalitaEsecuzione,item?.accesso,item?.difficolta,item?.atex,item?.ATEX];
    for (const key of ['lavorazioni','lavorazioniRichieste','lavorazioniStraordinarie','extra','extras','lavoriExtra','interventi','attivita','workItems']) {
      const value = item?.[key];
      if (Array.isArray(value)) parts.push(...value.map(v => typeof v === 'string' ? v : JSON.stringify(v)));
      else if (value) parts.push(typeof value === 'string' ? value : JSON.stringify(value));
    }
    return upper(parts.filter(Boolean).join(' '));
  }
  function areaMq(item) {
    for (const value of [item?.areaMq,item?.mq,item?.superficieMq,item?.metriQuadri,item?.superficie,item?.quantitaMq,item?.quantita]) {
      const n = Number(String(value ?? '').replace(',','.').replace(/[^0-9.]/g,'')); if (Number.isFinite(n) && n > 0) return n;
    }
    const match = workText(item).match(/(\d+(?:[.,]\d+)?)\s*(?:MQ|M2|M²)/); return match ? Number(match[1].replace(',','.')) : 0;
  }

  function recommendations(item, team) {
    const t = workText(item), mq = areaMq(item), rec = [];
    const add = (key,label,have,reason) => { if (!rec.some(r => r.key === key)) rec.push({key,label,have:Boolean(have),reason}); };
    if (/SFALC|TRINCI|ERBA|VERDE/.test(t) || mq > 0) {
      if (mq >= 2500 && !/STRETT|ACCESSO DIFFIC|RECINZ|FOSSO|SCARPAT/.test(t)) add('T','🚜 Trattore grande con trincia (T###)',team.hasBigTrincia,'superficie ampia');
      else add('R','🚜 Trattore piccolo con trincia (R###)',team.hasSmallTrincia,'area piccola/media o accesso più stretto');
      add('DECESP','🌿 Decespugliatore',team.hasDecespugliatore,'rifiniture, bordi e ostacoli');
      add('SOFF','💨 Soffiatore',team.hasSoffiatore,'pulizia finale');
    }
    if (/SIEP|TAGLIASIEP|POTAT.*SIEP/.test(t)) { add('TAGLIAS','✂️ Tagliasiepe',team.hasTagliasiepe,'potatura siepe'); add('SOFF','💨 Soffiatore',team.hasSoffiatore,'pulizia residui'); }
    if (/ALBER|ARBOR|MOTOSEG|ABBATT|POTAT.*PIANT/.test(t)) { add('MOTO','🪚 Motosega / attrezzatura potatura',team.hasMotosega,'potatura o abbattimento'); if (/ALTO|ALTA|PIATTAFORM|CHIOMA/.test(t)) add('PIATT','🛗 Piattaforma',team.hasPiattaforma,'lavoro in quota'); }
    if (/SPAZZATR/.test(t)) add('SPAZZ','🧹 Spazzatrice',team.hasSpazzatrice,'pulizia meccanica');
    if (/ATEX|AREA GAS|CABINA GAS|REMI|GRF/.test(t)) add('ATEX','🦺 Kit/DPI ATEX',/ATEX/.test(team.equipmentText),'area con requisiti ATEX');
    if (/RACCOLT|ASPORT|SMALT|INRETE/.test(t) || mq > 0) add('DAILY','🚛 Daily (A###)',team.hasDaily,'raccolta e trasporto materiale');
    return rec;
  }

  function findPlantForCard(card) {
    const id = text(card?.dataset?.recommendedId);
    const plants = currentPlants();
    if (id) {
      const found = plants.find(p => [plantId(p),text(p?.id),text(p?.idSap),text(p?.idSAP),text(p?.sap)].includes(id));
      if (found) return found;
    }
    const body = upper(card?.textContent || '');
    return plants.find(p => [p?.denominazione,p?.nome,p?.idSap].map(text).filter(Boolean).some(v => body.includes(upper(v)))) || null;
  }

  function applyMachineCorrection(state, team) {
    if (!state?.lastPlan?.length || !team.hasTrincia) return;
    const cls = machineClass(team);
    const profiles = readJson(PROFILE_KEY, {});
    const learned = profiles[`${cls}|team=${team.teamSize}`];
    const learnedFactor = learned?.factor && Number.isFinite(learned.factor) ? Math.max(0.65,Math.min(1.35,learned.factor)) : null;
    const initialFactor = cls === 'T' ? 0.55 : cls === 'R' ? 0.68 : 0.60;
    const factor = learnedFactor ?? initialFactor;
    let cumulative = 0;
    for (const entry of state.lastPlan) {
      if (!entry || entry.__equipmentAdjusted) { cumulative = Math.max(cumulative, Number(entry?.cumulativeMinutes)||0); continue; }
      const labels = entry.work?.breakdown || [];
      const shouldAdjust = labels.some(x => /SFALCIO SENZA TRINCIA|LAVORAZIONE MANUALE|DECESPUGLI/i.test(x?.label || ''));
      if (shouldAdjust) {
        entry.workMinutes = Math.max(5, Math.round((entry.workMinutes || entry.work?.minutes || 0) * factor));
        if (entry.work) {
          entry.work.minutes = entry.workMinutes;
          entry.work.breakdown = labels.map(x => /SFALCIO SENZA TRINCIA|LAVORAZIONE MANUALE/i.test(x?.label || '') ? {...x,label:cls==='T'?'Sfalcio con trattore grande T':'Sfalcio con trattore piccolo R',minutes:Math.max(5,Math.round((x.minutes||0)*factor))} : x);
        }
      }
      cumulative += (entry.driveMinutes||0) + (entry.workMinutes||0) + (entry.unload?.driveMinutes||0) + (entry.unload?.unloadMinutes||0);
      entry.cumulativeMinutes = cumulative;
      entry.fitsDay = cumulative <= 8*60;
      entry.__equipmentAdjusted = true;
    }
  }

  function decorate() {
    const panel = document.getElementById('recommended-plants-panel');
    if (!panel || panel.classList.contains('hidden')) return;
    const team = detectTeam();
    const state = window.HeraRecommendedPlants?.getState?.();
    applyMachineCorrection(state, team);
    const summarySpans = panel.querySelectorAll('.recommended-summary span');
    if (summarySpans[2]) summarySpans[2].textContent = machineLabel(team);

    panel.querySelectorAll('.recommended-item').forEach(card => {
      const item = findPlantForCard(card); if (!item) return;
      const recs = recommendations(item, team);
      let block = card.querySelector('[data-equipment-advice]');
      if (!block) { block = document.createElement('div'); block.dataset.equipmentAdvice = '1'; block.style.cssText = 'grid-column:2 / -1;margin-top:6px;padding:8px 10px;border-radius:10px;background:rgba(37,99,235,.06);font-size:.82rem;line-height:1.35'; const main = card.querySelector('.recommended-main'); main?.appendChild(block); }
      const missing = recs.filter(r => !r.have);
      block.innerHTML = `<strong>🧰 Attrezzature consigliate</strong><br>${recs.length ? recs.map(r => `${r.have?'✅':'⚠️'} ${escapeHtml(r.label)}`).join(' · ') : 'Nessuna attrezzatura speciale rilevata'}${missing.length ? `<br><strong>Da verificare prima di partire:</strong> ${missing.map(r=>escapeHtml(r.label.replace(/^\S+\s*/,''))).join(', ')}` : '<br><strong>✅ Attrezzatura completa</strong>'}`;
    });
  }

  function startPerformanceSession(item) {
    if (!item) return;
    const team = detectTeam();
    writeJson(ACTIVE_KEY,{date:todayKey(),startedAt:Date.now(),plantId:plantId(item),machine:machineClass(team),teamSize:team.teamSize,mq:areaMq(item)});
  }
  function finishPerformanceSession(item) {
    const active = readJson(ACTIVE_KEY,null); if (!active || active.date !== todayKey()) return;
    if (item && plantId(item) && active.plantId && plantId(item) !== active.plantId) return;
    const elapsed = (Date.now()-active.startedAt)/60000;
    try { localStorage.removeItem(ACTIVE_KEY); } catch (_) {}
    if (elapsed < 8 || elapsed > 720 || !active.mq || !['T','R'].includes(active.machine)) return;
    const observedPer100 = elapsed / Math.max(1,active.mq/100);
    const basePer100 = active.machine === 'T' ? 10 : 13;
    const observedFactor = Math.max(0.5,Math.min(1.8,observedPer100/basePer100));
    const profiles = readJson(PROFILE_KEY,{}), key = `${active.machine}|team=${active.teamSize}`;
    const p = profiles[key] || {count:0,factor:1};
    const alpha = Math.min(0.28,0.12 + (p.count||0)*0.01);
    profiles[key] = {count:(p.count||0)+1,factor:(p.count||0)?p.factor*(1-alpha)+observedFactor*alpha:observedFactor,lastObservedPer100:observedPer100,lastUpdated:new Date().toISOString()};
    writeJson(PROFILE_KEY,profiles);
  }

  document.addEventListener('click', event => {
    const node = event.target?.closest?.('a,button,[role="button"],input[type="button"],input[type="submit"]'); if (!node) return;
    const label = upper(node.textContent || node.getAttribute('aria-label') || node.value || '');
    const card = node.closest?.('[data-recommended-id],.impianto-card,.impianto-item,.plant-card,article,li,.card');
    const item = card ? findPlantForCard(card) : null;
    if (/\bNAVIGA\b/.test(label) && item) startPerformanceSession(item);
    if (/(^|\s)FATTO($|\s)/.test(label)) window.setTimeout(()=>finishPerformanceSession(item),900);
  }, true);

  let timer = 0;
  const observer = new MutationObserver(() => { clearTimeout(timer); timer = window.setTimeout(decorate,100); });
  function install() { observer.observe(document.body,{childList:true,subtree:true,attributes:true,attributeFilter:['class']}); decorate(); }
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded',install,{once:true}); else install();

  window.HeraEquipmentAdvisor = {installed:true,version:VERSION,getTeamInfo:detectTeam,recommendations,getProfiles:()=>readJson(PROFILE_KEY,{}),refresh:decorate};
})();