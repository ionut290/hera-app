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

  function init() {
    setLoginPhoto();

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
