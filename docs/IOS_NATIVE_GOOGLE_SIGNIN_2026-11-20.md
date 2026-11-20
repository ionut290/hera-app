# VARGA CANTIERI — Google login iPhone (fix staged, 20 Nov 2026)

## Root cause in source

Prior `auth-login-fix.js` showed **Login con Google** on native iOS but its
click handler called Firebase Auth `signInWithPopup()`, a web-browser popup
flow not reliable in a Capacitor iOS WKWebView. A native sign-in helper existed
but was unused by both the button and the exported login entrypoint.

## Corrected paths

- **iPhone**: `FirebaseAuthentication.signInWithGoogle({skipNativeAuth:true})`
  returns a Google ID token that is exchanged with Firebase JS Auth via
  `GoogleAuthProvider.credential()` and `signInWithCredential()`.
  The app continues using the same Firebase Auth session and existing
  `platformUsers` checks. There is **no fallback popup on iPhone**.
- **Web**: retains `signInWithPopup()`; for `auth/unauthorized-domain` the
  user receives an actionable message to authorize the site in Firebase.
- **Android**: unchanged email/password-only policy; Google remains hidden.
- **Builds**: before iOS `cap sync`, generated Podfile receives the
  `CapacitorFirebaseAuthentication/Google` subspec as required by the plugin
  maintainer. iOS release workflows also keep `configure-ios-firebase.py`
  to validate bundle, client ID, reversed URL scheme. This uses the
  independently stored registered Firebase plist and **does not add keys to git**.
- **Tests**: synthetic click/credential flow (iOS, browser, Android),
  rejected/missing plugin, missing token, cancelled login, rapid double tap,
  generated CocoaPods Podfile patch, versioned web/PWA cache.

## Rollout and remaining evidence

Changes are code only until CI, PR merge and **a newer TestFlight build than
20012**. No Apple App Store submission.

On real iPhone, test with a dedicated **nonproduction Google identity** that
has no access to employee/enterprise data:
1. Install the new internal TestFlight build.
2. Sign out; tap **Login con Google** and select the demo Google account.
3. Verify the native Google account chooser returns to VARGA CANTIERI and
   the same Firebase identity appears after app restart.
4. Confirm Google login failure/cancel returns to login with the ability
   to retry, without user/session data corruption.
5. Verify Apple sign in and email/password are still functional.
6. Record only app build/iOS version, test outcome and **redacted** screenshots.
7. If login reports Google developer configuration error, check registered
   Firebase iOS OAuth client and reverse URL scheme rather than re-enabling
   WKWebView popup.

Source: https://github.com/capawesome-team/capacitor-firebase/blob/main/packages/authentication/docs/setup-google.md
