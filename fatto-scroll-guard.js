(() => {
  'use strict';

  const MOVE_THRESHOLD_PX = 8;
  const BLOCK_AFTER_SCROLL_MS = 450;
  let lastTouchScrollAt = 0;
  let touchStartX = 0;
  let touchStartY = 0;
  let trackingTouch = false;
  let moved = false;

  function cleanText(value) {
    return String(value || '').replace(/\s+/g, ' ').trim().toUpperCase();
  }

  function findFattoButton(target) {
    const button = target?.closest?.('button, [role="button"], input[type="button"], input[type="submit"]');
    if (!button) return null;
    const label = cleanText(button.value || button.getAttribute('aria-label') || button.textContent);
    return /(^|\s)FATTO($|\s)/.test(label) ? button : null;
  }

  document.addEventListener('touchstart', (event) => {
    const touch = event.touches?.[0];
    if (!touch) return;
    trackingTouch = true;
    moved = false;
    touchStartX = touch.clientX;
    touchStartY = touch.clientY;
  }, { capture: true, passive: true });

  document.addEventListener('touchmove', (event) => {
    if (!trackingTouch) return;
    const touch = event.touches?.[0];
    if (!touch) return;
    const dx = touch.clientX - touchStartX;
    const dy = touch.clientY - touchStartY;
    if (!moved && Math.hypot(dx, dy) >= MOVE_THRESHOLD_PX) moved = true;
    if (moved) lastTouchScrollAt = Date.now();
  }, { capture: true, passive: true });

  document.addEventListener('touchend', () => {
    if (moved) lastTouchScrollAt = Date.now();
    trackingTouch = false;
    moved = false;
  }, { capture: true, passive: true });

  document.addEventListener('touchcancel', () => {
    if (moved) lastTouchScrollAt = Date.now();
    trackingTouch = false;
    moved = false;
  }, { capture: true, passive: true });

  document.addEventListener('click', (event) => {
    const button = findFattoButton(event.target);
    if (!button) return;
    if ((Date.now() - lastTouchScrollAt) > BLOCK_AFTER_SCROLL_MS) return;

    event.preventDefault();
    event.stopPropagation();
    event.stopImmediatePropagation();

    button.classList.remove('fatto-scroll-guard-blocked');
    void button.offsetWidth;
    button.classList.add('fatto-scroll-guard-blocked');
    window.setTimeout(() => button.classList.remove('fatto-scroll-guard-blocked'), 260);
  }, true);

  function loadScriptOnce(selector, src, datasetKey, onerror) {
    if (document.querySelector(selector)) return;
    const script = document.createElement('script');
    script.src = src;
    script.async = true;
    script.dataset[datasetKey] = 'true';
    script.onerror = onerror;
    document.head.appendChild(script);
  }

  function loadRecommendedHelpers() {
    if (!window.HeraEquipmentAdvisor?.installed) {
      loadScriptOnce(
        'script[data-equipment-recommendations]',
        './equipment-recommendations.js?v=20260823a',
        'equipmentRecommendations',
        () => console.warn('Consigli attrezzature non caricati.')
      );
    }
    if (!window.HeraAdaptiveWorkLearning?.installed) {
      loadScriptOnce(
        'script[data-adaptive-work-learning]',
        './adaptive-work-learning.js?v=20260823a',
        'adaptiveWorkLearning',
        () => console.warn('Apprendimento adattivo non caricato.')
      );
    }
    if (!window.HeraRecommendedTrafficWeather?.installed) {
      loadScriptOnce(
        'script[data-recommended-traffic-weather]',
        './recommended-traffic-weather.js?v=20260823a',
        'recommendedTrafficWeather',
        () => console.warn('Traffico/meteo Impianti consigliati non caricato.')
      );
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', loadRecommendedHelpers, { once: true });
  } else {
    loadRecommendedHelpers();
  }
})();