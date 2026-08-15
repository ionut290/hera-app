(() => {
  'use strict';

  if (window.__heraMapDetailCloseFixInstalled) return;
  window.__heraMapDetailCloseFixInstalled = true;

  const PANEL_ID = 'map-impianto-detail-panel';
  const BODY_ID = 'map-impianto-detail-body';
  const STYLE_ID = 'hera-map-detail-close-fix-style';

  function closePanel() {
    const panel = document.getElementById(PANEL_ID);
    if (!panel) return;
    panel.classList.add('hidden');
    panel.setAttribute('aria-hidden', 'true');
  }

  function installStyle() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${PANEL_ID} .map-popup-header {
        position: sticky !important;
        top: 0 !important;
        z-index: 50 !important;
        min-height: 60px;
        padding-right: 68px !important;
      }
      #${PANEL_ID} .map-popup-header > .hera-map-detail-close {
        position: absolute !important;
        top: 50% !important;
        right: 12px !important;
        left: auto !important;
        transform: translateY(-50%) !important;
        z-index: 100 !important;
        width: 42px !important;
        height: 42px !important;
        display: grid !important;
        place-items: center !important;
        opacity: 1 !important;
        visibility: visible !important;
        pointer-events: auto !important;
        touch-action: manipulation;
      }
      #${PANEL_ID} .map-popup-header > .hera-map-detail-close:active {
        transform: translateY(-50%) scale(.96) !important;
      }
      @media (max-width: 720px) {
        #${PANEL_ID} .map-popup-header {
          min-height: 62px;
          padding-right: 66px !important;
        }
        #${PANEL_ID} .map-popup-header > .hera-map-detail-close {
          right: 10px !important;
          width: 42px !important;
          height: 42px !important;
        }
      }
    `;
    document.head.appendChild(style);
  }

  function ensureCloseButton() {
    const panel = document.getElementById(PANEL_ID);
    const body = document.getElementById(BODY_ID);
    if (!panel || !body) return;

    const header = body.querySelector('.map-popup-header');
    if (!header) return;

    let close = panel.querySelector('.hera-map-detail-close');
    if (!close) {
      close = document.createElement('button');
      close.type = 'button';
      close.className = 'hera-map-detail-close';
      close.setAttribute('aria-label', 'Chiudi dettaglio impianto');
      close.setAttribute('title', 'Chiudi');
      close.textContent = '×';
      close.addEventListener('click', (event) => {
        event.preventDefault();
        event.stopPropagation();
        closePanel();
      });
    }

    if (close.parentElement !== header) header.appendChild(close);
  }

  function init() {
    installStyle();
    ensureCloseButton();

    const panel = document.getElementById(PANEL_ID);
    if (panel) {
      const observer = new MutationObserver(() => ensureCloseButton());
      observer.observe(panel, { childList: true, subtree: true });
    }

    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') closePanel();
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
