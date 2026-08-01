(() => {
  'use strict';

  const PV = window.HeraPreventivi;
  const M = window.HeraPreventiviModels;
  if (!PV || !M) return;

  const SCOPE = 'https://www.googleapis.com/auth/drive.file';
  const SETTINGS_KEY = 'hera_preventivi_drive_v1';
  const text = (value) => String(value ?? '').trim();
  const escape = (value) => PV.escapeHtml ? PV.escapeHtml(value) : text(value);
  const state = {
    token: '',
    tokenClient: null,
    pendingUpload: null,
    settings: loadSettings()
  };

  function loadSettings() {
    try {
      const parsed = JSON.parse(localStorage.getItem(SETTINGS_KEY) || '{}');
      return parsed && typeof parsed === 'object' ? parsed : {};
    } catch (_) {
      return {};
    }
  }

  function saveSettings() {
    localStorage.setItem(SETTINGS_KEY, JSON.stringify(state.settings));
    PV.state.settings = { ...(PV.state.settings || {}), drive: { ...state.settings } };
    PV.persistLocal?.();
  }

  function clientId() {
    return text(window.HERA_GOOGLE_DRIVE_CLIENT_ID || window.HERA_GOOGLE_CLIENT_ID || document.querySelector('meta[name="google-drive-client-id"]')?.content);
  }

  function folderIdFromInput(value) {
    const raw = text(value);
    const match = raw.match(/\/folders\/([a-zA-Z0-9_-]+)/) || raw.match(/[?&]id=([a-zA-Z0-9_-]+)/);
    return match?.[1] || (/^[a-zA-Z0-9_-]{10,}$/.test(raw) ? raw : '');
  }

  function loadGoogleIdentity() {
    if (window.google?.accounts?.oauth2) return Promise.resolve();
    return new Promise((resolve, reject) => {
      const existing = document.querySelector('script[data-google-identity-services]');
      if (existing) {
        existing.addEventListener('load', resolve, { once: true });
        existing.addEventListener('error', reject, { once: true });
        return;
      }
      const script = document.createElement('script');
      script.src = 'https://accounts.google.com/gsi/client';
      script.async = true;
      script.defer = true;
      script.dataset.googleIdentityServices = '1';
      script.addEventListener('load', resolve, { once: true });
      script.addEventListener('error', reject, { once: true });
      document.head.appendChild(script);
    });
  }

  async function authorize() {
    const id = clientId();
    if (!id) throw new Error('Google Drive non configurato: manca il Client ID OAuth dell’app.');
    await loadGoogleIdentity();
    if (!state.tokenClient) {
      state.tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: id,
        scope: SCOPE,
        callback: (response) => {
          if (response?.error) {
            state.pendingUpload?.reject(new Error(response.error_description || response.error));
          } else {
            state.token = response.access_token || '';
            state.settings.connectedEmail = window.firebase?.auth?.()?.currentUser?.email || '';
            state.settings.connectedAt = new Date().toISOString();
            saveSettings();
            state.pendingUpload?.resolve(state.token);
          }
          state.pendingUpload = null;
          renderDriveView();
        }
      });
    }
    if (state.token) return state.token;
    return new Promise((resolve, reject) => {
      state.pendingUpload = { resolve, reject };
      state.tokenClient.requestAccessToken({ prompt: state.settings.connectedAt ? '' : 'consent' });
    });
  }

  function ensureTab() {
    const page = PV.page?.() || document.getElementById('preventivi-page');
    const nav = page?.querySelector('.pv-nav');
    if (!nav || nav.querySelector('[data-pv-view="drive"]')) return;
    nav.insertAdjacentHTML('beforeend', '<button type="button" class="pv-tab" data-pv-view="drive">Google Drive</button>');
  }

  function renderDriveView() {
    const page = PV.page?.() || document.getElementById('preventivi-page');
    if (!page || PV.state.view !== 'drive') return;
    const content = page.querySelector('[data-pv-content]') || page.querySelector('.pv-content');
    if (!content) return;
    const folderValue = state.settings.folderUrl || state.settings.folderId || '';
    const configured = Boolean(clientId());
    content.innerHTML = `
      <div class="pv-shell">
        <div class="pv-section-head"><div><h2>Google Drive</h2><p class="pv-muted">Imposta la cartella dove conservare preventivi e consuntivi.</p></div></div>
        <section class="pv-form-card">
          <div class="pv-form-grid">
            <label class="pv-label pv-span-2"><span>Cartella Google Drive</span><input data-drive-folder value="${escape(folderValue)}" placeholder="Incolla il link della cartella o il suo ID"></label>
            <label class="pv-check-card pv-span-2"><input type="checkbox" data-drive-auto ${state.settings.autoSave !== false ? 'checked' : ''}><span><strong>Salva automaticamente su Drive</strong><small>Carica una copia quando scarichi il modello compilato o il PDF.</small></span></label>
          </div>
          <div class="pv-form-actions">
            <button type="button" class="pv-btn pv-btn-primary" data-drive-save-settings>Salva percorso</button>
            <button type="button" class="pv-btn pv-btn-secondary" data-drive-connect>${state.token ? 'Drive collegato' : 'Collega Google Drive'}</button>
            ${state.settings.folderId ? '<button type="button" class="pv-btn pv-btn-light" data-drive-open>Apri cartella</button>' : ''}
            <button type="button" class="pv-btn pv-btn-light" data-drive-disconnect>Disconnetti</button>
          </div>
          <p class="pv-feedback" data-pv-feedback role="status"></p>
        </section>
        <section class="pv-form-card">
          <h3>Stato</h3>
          <p>${configured ? 'Configurazione OAuth presente.' : '<b>Configurazione tecnica mancante:</b> inserire il Client ID OAuth in <code>window.HERA_GOOGLE_DRIVE_CLIENT_ID</code> o nel meta tag <code>google-drive-client-id</code>.'}</p>
          <p>Account: ${escape(state.settings.connectedEmail || 'non collegato')}</p>
          <p>Cartella: ${escape(state.settings.folderId || 'non impostata')}</p>
        </section>
      </div>`;
  }

  async function uploadBlob(blob, filename) {
    if (!blob || !state.settings.folderId) throw new Error('Imposta prima la cartella Google Drive.');
    const token = await authorize();
    const metadata = {
      name: filename || `Documento-${Date.now()}`,
      parents: [state.settings.folderId]
    };
    const boundary = `hera_${Date.now()}_${Math.random().toString(16).slice(2)}`;
    const body = new Blob([
      `--${boundary}\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n${JSON.stringify(metadata)}\r\n`,
      `--${boundary}\r\nContent-Type: ${blob.type || 'application/octet-stream'}\r\n\r\n`,
      blob,
      `\r\n--${boundary}--`
    ]);
    const response = await fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id,name,webViewLink', {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': `multipart/related; boundary=${boundary}`
      },
      body
    });
    if (!response.ok) {
      if (response.status === 401) state.token = '';
      const detail = await response.text();
      throw new Error(`Caricamento Drive non riuscito (${response.status}): ${detail.slice(0, 180)}`);
    }
    const file = await response.json();
    markCurrentDocumentOnDrive(file, filename);
    return file;
  }

  function currentDocument() {
    const id = PV.state?.editingQuoteId || PV.state?.editingConsuntivoId;
    if (id && id !== 'new') return PV.getQuote?.(id) || PV.getConsuntivo?.(id) || null;
    return null;
  }

  function markCurrentDocumentOnDrive(file, filename) {
    const doc = currentDocument();
    if (!doc) return;
    doc.driveSaved = true;
    doc.driveFileId = file.id || '';
    doc.driveUrl = file.webViewLink || (file.id ? `https://drive.google.com/file/d/${file.id}/view` : '');
    doc.driveFileName = file.name || filename || '';
    doc.driveSavedAt = new Date().toISOString();
    doc.driveFolderId = state.settings.folderId;
    doc.syncPending = true;
    PV.persistLocal?.();
    PV.scheduleSync?.();
  }

  const originalDownload = M.download?.bind(M);
  if (originalDownload && !M.__driveDownloadPatched) {
    M.__driveDownloadPatched = true;
    M.download = (blob, filename) => {
      const result = originalDownload(blob, filename);
      if (state.settings.autoSave !== false && state.settings.folderId) {
        uploadBlob(blob, filename)
          .then(() => PV.setFeedback?.('Documento scaricato e salvato su Google Drive.', 'success'))
          .catch((error) => {
            console.warn('Google Drive:', error);
            PV.setFeedback?.(`Documento scaricato. Copia Drive non salvata: ${error.message}`, 'error');
          });
      }
      return result;
    };
  }

  document.addEventListener('click', async (event) => {
    const tab = event.target.closest('[data-pv-view="drive"]');
    if (tab) {
      event.preventDefault();
      PV.state.view = 'drive';
      PV.state.editingQuoteId = '';
      PV.state.editingConsuntivoId = '';
      renderDriveView();
      document.querySelectorAll('[data-pv-view]').forEach((button) => button.classList.toggle('is-active', button.dataset.pvView === 'drive'));
      return;
    }
    if (event.target.closest('[data-drive-save-settings]')) {
      const input = document.querySelector('[data-drive-folder]');
      const folderId = folderIdFromInput(input?.value);
      if (!folderId) return PV.setFeedback?.('Inserisci un link o un ID valido di una cartella Drive.', 'error');
      state.settings.folderUrl = text(input?.value);
      state.settings.folderId = folderId;
      state.settings.autoSave = document.querySelector('[data-drive-auto]')?.checked !== false;
      saveSettings();
      PV.setFeedback?.('Percorso Google Drive salvato.', 'success');
      renderDriveView();
      return;
    }
    if (event.target.closest('[data-drive-connect]')) {
      try {
        await authorize();
        PV.setFeedback?.('Google Drive collegato.', 'success');
      } catch (error) {
        PV.setFeedback?.(error.message, 'error');
      }
      return;
    }
    if (event.target.closest('[data-drive-open]') && state.settings.folderId) {
      window.open(`https://drive.google.com/drive/folders/${state.settings.folderId}`, '_blank', 'noopener');
      return;
    }
    if (event.target.closest('[data-drive-disconnect]')) {
      if (state.token && window.google?.accounts?.oauth2?.revoke) google.accounts.oauth2.revoke(state.token, () => {});
      state.token = '';
      delete state.settings.connectedAt;
      delete state.settings.connectedEmail;
      saveSettings();
      renderDriveView();
    }
  }, true);

  const originalRender = PV.renderCurrentView?.bind(PV);
  if (originalRender && !PV.__driveRenderPatched) {
    PV.__driveRenderPatched = true;
    PV.renderCurrentView = (...args) => {
      ensureTab();
      if (PV.state.view === 'drive') return renderDriveView();
      const result = originalRender(...args);
      queueMicrotask(ensureTab);
      return result;
    };
  }

  const start = () => {
    ensureTab();
    window.HeraPreventiviDrive = { authorize, uploadBlob, settings: state.settings };
  };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', start, { once: true });
  else start();
})();
