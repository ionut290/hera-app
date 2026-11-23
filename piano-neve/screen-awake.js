// Screen wake lock belongs to navigation, not GPS tracking or map browsing.
(function () {
 'use strict';
 let wanted = false, sentinel = null, pending = false, suspended = false;
 let retryTimer = null, retries = 0;
 const visible = () => !suspended && document.visibilityState === 'visible';
 function cancelRetry() {
  if (retryTimer !== null) clearTimeout(retryTimer);
  retryTimer = null;
 }
 async function release() {
  const old = sentinel;
  sentinel = null;
  if (old && !old.released) {
   try { await old.release(); } catch (error) { console.warn('Screen wake lock release failed', error); }
  }
 }
 async function sync() {
  if (!wanted || !visible()) { cancelRetry(); await release(); return; }
  if (pending || (sentinel && !sentinel.released) || !navigator.wakeLock?.request) return;
  pending = true;
  try {
   const lock = await navigator.wakeLock.request('screen');
   // Stop/background may occur while the browser is granting the lock.
   if (!wanted || !visible()) { await lock.release(); return; }
   sentinel = lock;
   lock.addEventListener('release', () => {
    if (sentinel !== lock) return;
    sentinel = null;
    // Bounded retries: respect OS power-saving policy without a request loop.
    if (wanted && visible() && retries < 2) {
     retries++;
     retryTimer = setTimeout(() => { retryTimer = null; void sync(); }, 1000 * retries);
    }
   });
  } catch (error) {
   console.warn('Screen wake lock unavailable', error);
  } finally { pending = false; }
 }
 function restore() { suspended = false; retries = 0; cancelRetry(); void sync(); }
 window.VargaScreenAwake = {
  setActive(active) {
   const next = !!active;
   if (next === wanted) return;
   wanted = next; retries = 0; cancelRetry(); void sync();
  }
 };
 document.addEventListener('visibilitychange', () => {
  if (document.visibilityState === 'visible') restore(); else void sync();
 });
 // A user gesture permits retry after an earlier browser/OS refusal.
 document.addEventListener('pointerdown', () => { if (wanted && !sentinel && !pending) restore(); }, {passive:true});
 window.addEventListener('focus', restore);
 window.addEventListener('pageshow', restore);
 window.addEventListener('pagehide', () => { suspended = true; cancelRetry(); void release(); });
})();
