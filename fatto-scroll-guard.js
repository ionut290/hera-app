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
})();
