(() => {
  'use strict';

  const db = () => window.firebase?.firestore?.() || null;
  const text = (value) => String(value ?? '').trim();
  const clean = (source) => Object.fromEntries(Object.entries(source || {}).filter(([, value]) => value !== undefined && value !== null && text(value) !== ''));
  const timestamp = () => window.firebase?.firestore?.FieldValue?.serverTimestamp?.() || new Date().toISOString();

  function visibleCommessaRef(commessaId) {
    return db()?.collection('globalCommesse').doc(text(commessaId));
  }

  function visiblePlantsRef(commessaId) {
    return visibleCommessaRef(commessaId)?.collection('impianti');
  }

  function visibleCommessaData(raw = {}, commessaId = '') {
    return clean({
      nome: raw.nome || raw.name || raw.denominazione || raw.titolo || raw.commessa || 'Commessa',
      codice: raw.codice || raw.codiceCommessa || raw.code || raw.numeroCommessa || '',
      cliente: raw.cliente || raw.committente || raw.clientName || raw.ragioneSociale || '',
      numeroContratto: raw.numeroContratto || raw.contratto || raw.contractNumber || '',
      sourceCommessaId: commessaId,
      archivioPermanente: true,
      sincronizzataDaArchivio: true,
      updatedAt: timestamp()
    });
  }

  function visiblePlantData(raw = {}, commessaId = '', plantId = '') {
    return clean({
      sourceCommessaId: commessaId,
      sourceImpiantoId: plantId || raw.sourceImpiantoId || '',
      globalImpiantoId: raw.globalImpiantoId || '',
      idSap: raw.idSap || raw.idSAP || raw['ID SAP'] || raw.sap || raw.codiceSap || raw.codiceHera || '',
      idSAP: raw.idSap || raw.idSAP || raw['ID SAP'] || raw.sap || raw.codiceSap || raw.codiceHera || '',
      denominazione: raw.denominazione || raw['Denominazione Impianto'] || raw.nome || raw.name || raw.impianto || '',
      comune: raw.comune || raw.Comune || raw.city || raw.localita || '',
      indirizzo: raw.indirizzo || raw['Descrizione via'] || raw.descrizioneVia || raw.via || raw.address || raw.ubicazione || '',
      descrizioneVia: raw.descrizioneVia || raw['Descrizione via'] || raw.indirizzo || raw.via || '',
      tipologiaImpianto: raw.tipologiaImpianto || raw.tipologia || raw['Tipologia impianto'] || raw.tipo || raw.type || '',
      area: raw.area || raw.AREA || raw.competenza || raw['Area/Competenza'] || '',
      competenza: raw.competenza || raw.area || raw.AREA || '',
      coordinate: raw.coordinate || raw.coordinates || raw.coordinateGps || '',
      gpsX: raw.gpsX ?? raw.longitude ?? raw.longitudine ?? null,
      gpsY: raw.gpsY ?? raw.latitude ?? raw.latitudine ?? null,
      dittaEsecutrice: raw.dittaEsecutrice || raw['Ditta esecutrice'] || '',
      archivioPermanente: true,
      sincronizzatoDaArchivio: true,
      updatedAt: timestamp()
    });
  }

  async function mirrorCommessa(commessaId, raw = {}) {
    const ref = visibleCommessaRef(commessaId);
    if (!ref || !text(commessaId)) return;
    await ref.set(visibleCommessaData(raw, commessaId), { merge: true });
  }

  async function mirrorPlant(commessaId, plantId, raw = {}) {
    const ref = visiblePlantsRef(commessaId);
    if (!ref || !text(commessaId)) return;
    const archiveApi = window.HeraGlobalArchive;
    const key = archiveApi?.plantKey?.(raw, plantId)
      || `source-${text(plantId || raw.sourceImpiantoId || raw.idSap || raw.idSAP || raw.denominazione).replace(/[^a-zA-Z0-9_-]+/g, '-').toLowerCase()}`;
    await mirrorCommessa(commessaId, raw);
    await ref.doc(key).set(visiblePlantData(raw, commessaId, plantId), { merge: true });
  }

  async function mirrorAll() {
    const firestore = db();
    if (!firestore || !window.firebase?.auth?.()?.currentUser) return;
    const archive = await firestore.collection('globalArchive').doc('commesse').collection('items').get();
    for (const commessaDoc of archive.docs) {
      const commessaId = commessaDoc.id;
      await mirrorCommessa(commessaId, commessaDoc.data() || {});
      const plants = await commessaDoc.ref.collection('impianti').get();
      for (const plantDoc of plants.docs) await mirrorPlant(commessaId, plantDoc.id, plantDoc.data() || {});
    }
  }

  function patchArchiveApi() {
    const api = window.HeraGlobalArchive;
    if (!api || api.__visibilityPatched) return false;
    const originalArchiveCommessa = api.archiveCommessa;
    const originalArchivePlant = api.archivePlant;
    const patched = {
      ...api,
      __visibilityPatched: true,
      archiveCommessa: async (commessaId, raw = {}) => {
        await originalArchiveCommessa(commessaId, raw);
        await mirrorCommessa(commessaId, raw);
      },
      archivePlant: async (commessaId, plantId, raw = {}) => {
        await originalArchivePlant(commessaId, plantId, raw);
        await mirrorPlant(commessaId, plantId, raw);
      },
      mirrorAll,
      mirrorCommessa,
      mirrorPlant
    };
    window.HeraGlobalArchive = Object.freeze(patched);
    return true;
  }

  function start() {
    const install = () => {
      if (!patchArchiveApi()) return false;
      window.firebase?.auth?.().onAuthStateChanged((user) => {
        if (user) mirrorAll().catch((error) => console.warn('Global archive: commesse non rese visibili.', error));
      });
      if (window.firebase?.auth?.()?.currentUser) mirrorAll().catch((error) => console.warn('Global archive: commesse non rese visibili.', error));
      return true;
    };
    if (install()) return;
    let attempts = 0;
    const timer = setInterval(() => {
      attempts += 1;
      if (install() || attempts >= 40) clearInterval(timer);
    }, 250);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
