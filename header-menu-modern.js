(() => {
  function setLoginPhoto() {
    const button = document.getElementById('user-toggle-btn');
    if (!button) return;

    const icon = button.querySelector('.header-action-icon');
    const image = button.querySelector('.header-login-photo');
    const user = window.firebase?.auth?.()?.currentUser || null;
    const photoURL = user?.photoURL || '';

    if (image && photoURL) {
      image.src = photoURL;
      image.alt = user?.displayName ? `Foto profilo di ${user.displayName}` : 'Foto profilo Google';
      button.classList.add('has-photo');
    } else {
      button.classList.remove('has-photo');
      if (image) image.removeAttribute('src');
      if (icon) icon.textContent = '👤';
    }
  }

  function syncMenuViewState() {
    const menu = document.getElementById('side-menu');
    const isOpen = Boolean(menu && !menu.classList.contains('hidden'));
    document.body.classList.toggle('menu-fullscreen-open', isOpen);
  }

  function initMenuViewObserver() {
    const menu = document.getElementById('side-menu');
    if (!menu) return;

    syncMenuViewState();

    const observer = new MutationObserver(syncMenuViewState);
    observer.observe(menu, {
      attributes: true,
      attributeFilter: ['class', 'aria-hidden']
    });

    const closeButton = document.getElementById('menu-close-btn');
    if (closeButton) {
      closeButton.addEventListener('click', () => {
        window.setTimeout(syncMenuViewState, 0);
      });
    }
  }

  function init() {
    setLoginPhoto();
    initMenuViewObserver();

    if (window.firebase?.auth) {
      try {
        window.firebase.auth().onAuthStateChanged(() => setLoginPhoto());
      } catch (error) {
        console.warn('Impossibile aggiornare la foto profilo Google:', error);
      }
    }
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init, { once: true });
  } else {
    init();
  }
})();
