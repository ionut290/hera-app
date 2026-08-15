(() => {
  'use strict';

  if (window.__heraMapDetailCloseFixInstalled) return;
  window.__heraMapDetailCloseFixInstalled = true;

  const PANEL_ID = 'map-impianto-detail-panel';
  const BODY_ID = 'map-impianto-detail-body';
  const PAGE_ID = 'map-fullscreen-page';
  const BACK_ID = 'map-fullscreen-back-btn';
  const SEARCH_FORM_ID = 'map-fullscreen-number-search-form';
  const SEARCH_INPUT_ID = 'map-fullscreen-number-search-input';
  const FEEDBACK_ID = 'map-fullscreen-feedback-banner';
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
        min-height: 58px;
        padding-right: 64px !important;
        background: rgba(255,255,255,.98) !important;
      }
      #${PANEL_ID} .map-popup-header > .hera-map-detail-close {
        position: absolute !important;
        top: 50% !important;
        right: 10px !important;
        left: auto !important;
        transform: translateY(-50%) !important;
        z-index: 100 !important;
        width: 40px !important;
        height: 40px !important;
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
        #${PAGE_ID},
        #${PAGE_ID} .map-fullscreen-page-card {
          position: relative !important;
          width: 100% !important;
          height: 100dvh !important;
          min-height: 100dvh !important;
          overflow: hidden !important;
          padding: 0 !important;
          margin: 0 !important;
          background: transparent !important;
        }

        #${PAGE_ID} .map-fullscreen-map-wrap {
          position: absolute !important;
          inset: 0 !important;
          width: 100% !important;
          height: 100dvh !important;
          min-height: 100dvh !important;
          overflow: hidden !important;
        }

        #${PAGE_ID} #map-fullscreen-view {
          width: 100% !important;
          height: 100% !important;
          min-height: 100% !important;
        }

        #${PAGE_ID} .map-fullscreen-toolbar {
          position: absolute !important;
          inset: 0 !important;
          z-index: 1450 !important;
          height: 0 !important;
          min-height: 0 !important;
          padding: 0 !important;
          margin: 0 !important;
          border: 0 !important;
          background: transparent !important;
          box-shadow: none !important;
          pointer-events: none !important;
        }

        #${PAGE_ID} .map-fullscreen-toolbar > * {
          pointer-events: auto;
        }

        #${PAGE_ID} #${BACK_ID} {
          position: fixed !important;
          top: calc(env(safe-area-inset-top, 0px) + 14px) !important;
          left: 20px !important;
          z-index: 1600 !important;
          width: 52px !important;
          height: 52px !important;
          min-width: 52px !important;
          min-height: 52px !important;
          padding: 0 !important;
          display: grid !important;
          place-items: center !important;
          border: 1px solid rgba(203,213,225,.9) !important;
          border-radius: 50% !important;
          background: rgba(255,255,255,.96) !important;
          color: #334155 !important;
          font-size: 29px !important;
          font-weight: 500 !important;
          line-height: 1 !important;
          box-shadow: 0 8px 22px rgba(15,23,42,.16) !important;
          backdrop-filter: blur(12px);
          -webkit-backdrop-filter: blur(12px);
        }

        #${PAGE_ID} #${SEARCH_FORM_ID} {
          position: fixed !important;
          top: calc(env(safe-area-inset-top, 0px) + 90px) !important;
          left: 20px !important;
          right: 20px !important;
          z-index: 1580 !important;
          width: auto !important;
          min-width: 0 !important;
          max-width: none !important;
          margin: 0 !important;
          padding: 0 !important;
          background: transparent !important;
        }

        #${PAGE_ID} #${SEARCH_FORM_ID} .hera-map-search-shell {
          width: 100% !important;
        }

        #${PAGE_ID} #${SEARCH_INPUT_ID} {
          height: 58px !important;
          min-height: 58px !important;
          width: 100% !important;
          padding-left: 56px !important;
          padding-right: 48px !important;
          border-radius: 22px !important;
          border: 1px solid rgba(203,213,225,.9) !important;
          background: rgba(255,255,255,.98) !important;
          box-shadow: 0 10px 28px rgba(15,23,42,.18) !important;
          font-size: 16px !important;
          font-weight: 650 !important;
        }

        #${PAGE_ID} #${SEARCH_FORM_ID} .hera-map-search-icon {
          left: 19px !important;
          font-size: 22px !important;
          opacity: .72 !important;
        }

        #${PANEL_ID} {
          position: fixed !important;
          top: calc(env(safe-area-inset-top, 0px) + 166px) !important;
          right: 20px !important;
          bottom: calc(env(safe-area-inset-bottom, 0px) + 92px) !important;
          left: 20px !important;
          z-index: 1500 !important;
          width: auto !important;
          max-width: none !important;
          max-height: none !important;
          overflow: hidden !important;
          border: 1px solid rgba(203,213,225,.92) !important;
          border-radius: 22px !important;
          background: rgba(255,255,255,.99) !important;
          box-shadow: 0 18px 42px rgba(15,23,42,.22) !important;
        }

        #${BODY_ID} {
          width: 100% !important;
          height: 100% !important;
          max-height: 100% !important;
          overflow-y: auto !important;
          overscroll-behavior: contain !important;
          -webkit-overflow-scrolling: touch !important;
          border-radius: 22px !important;
        }

        #${PANEL_ID} .map-popup-card {
          min-height: 100% !important;
          max-height: none !important;
          border-radius: 22px !important;
          display: flex !important;
          flex-direction: column !important;
          background: #fff !important;
        }

        #${PANEL_ID} .map-popup-header {
          min-height: 54px !important;
          padding: 14px 58px 11px 14px !important;
          flex: 0 0 auto !important;
        }

        #${PANEL_ID} .map-popup-header h3 {
          font-size: 1.02rem !important;
          line-height: 1.2 !important;
        }

        #${PANEL_ID} .map-popup-details {
          padding: 4px 14px 6px !important;
          flex: 1 1 auto !important;
        }

        #${PANEL_ID} .map-popup-details dt {
          margin-top: 9px !important;
          font-size: .72rem !important;
          letter-spacing: .05em !important;
        }

        #${PANEL_ID} .map-popup-details dd {
          margin-top: 2px !important;
          font-size: .94rem !important;
          line-height: 1.28 !important;
        }

        #${PANEL_ID} .map-popup-actions {
          position: sticky !important;
          bottom: 0 !important;
          z-index: 30 !important;
          grid-template-columns: 1fr !important;
          gap: 7px !important;
          padding: 10px 14px 12px !important;
          margin-top: auto !important;
          border-top: 1px solid rgba(226,232,240,.96) !important;
          background: rgba(255,255,255,.99) !important;
          flex: 0 0 auto !important;
        }

        #${PANEL_ID} .map-popup-actions .btn {
          width: 100% !important;
          min-height: 46px !important;
          border-radius: 14px !important;
          font-size: .9rem !important;
          font-weight: 800 !important;
        }

        #${PANEL_ID} .map-popup-header > .hera-map-detail-close {
          right: 8px !important;
          width: 38px !important;
          height: 38px !important;
        }

        #${FEEDBACK_ID} {
          position: fixed !important;
          left: 20px !important;
          right: 20px !important;
          bottom: calc(env(safe-area-inset-bottom, 0px) + 12px) !important;
          z-index: 1520 !important;
          min-height: 58px !important;
          border-radius: 20px !important;
          box-shadow: 0 10px 26px rgba(15,23,42,.14) !important;
        }
      }
    `;
    document.head.appendChild(style);
  }

  function normalizeBackButton() {
    const back = document.getElementById(BACK_ID);
    if (!back) return;
    if (back.dataset.arrowOnly === '1') return;
    back.dataset.arrowOnly = '1';
    back.textContent = '←';
    back.setAttribute('aria-label', 'Indietro');
    back.setAttribute('title', 'Indietro');
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
    normalizeBackButton();
    ensureCloseButton();

    const page = document.getElementById(PAGE_ID);
    if (page) {
      const pageObserver = new MutationObserver(() => {
        normalizeBackButton();
        ensureCloseButton();
      });
      pageObserver.observe(page, { childList: true, subtree: true });
    }

    document.addEventListener('keydown', (event) => {
      if (event.key === 'Escape') closePanel();
    });
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true });
  else init();
})();
