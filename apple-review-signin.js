(() => {
  "use strict";
  const form = document.getElementById("review-login-form");
  const feedback = document.getElementById("review-login-message");
  const submit = document.getElementById("review-enter-btn");
  if (!form || !feedback || !submit) return;

  let reviewerAuth;
  function auth() {
    if (!reviewerAuth) {
      const app = firebase.initializeApp(window.VargaReviewFirebase, "varga-apple-review-only");
      reviewerAuth = app.auth();
    }
    return reviewerAuth;
  }
  form.addEventListener("submit", async (event) => {
    event.preventDefault();
    const email = String(document.getElementById("review-email").value || "").trim();
    const password = String(document.getElementById("review-password").value || "");
    if (!email || !password) {
      feedback.textContent = "Inserisci email e password.";
      return;
    }
    submit.disabled = true;
    feedback.textContent = "Verifica dell'account Apple Review...";
    try {
      const reviewAuth = auth();
      // No persisted reviewer session or access to normal app Auth state.
      await reviewAuth.setPersistence(firebase.auth.Auth.Persistence.NONE);
      const response = await reviewAuth.signInWithEmailAndPassword(email, password);
      const claims = (await response.user.getIdTokenResult(true)).claims || {};
      if (claims.apple_review !== true || claims.admin === true || claims.isAdmin === true) {
        await reviewAuth.signOut();
        feedback.textContent = "Questo account non è abilitato alla revisione Apple.";
        return;
      }
      await reviewAuth.signOut();
      document.getElementById("review-password").value = "";
      window.location.replace("./apple-review-demo.html");
    } catch (error) {
      feedback.textContent = "Accesso non riuscito. Verifica le credenziali di revisione.";
    } finally {
      submit.disabled = false;
    }
  });
})();