# Apple Review account — restricted Firebase identity (v1, candidate)

## Status
**Only a pull request. No automatic App Store upload or App Review submission.**
The feature requires server review, protected CI, manual testing, a real account creation step and an updated iOS build. Do not expose a reviewer identity until the server-side guards are deployed.

## Flow
1. Owner signs into Varga Cantieri with the existing administrator account.
2. Owner opens `apple-review-admin.html` on the **same app origin**. The dedicated page uses the existing signed-in Firebase session; the server checks the verified administrator email against `appConfig/adminUsers` or the fixed current owner email.
3. The administrator enters an email address reserved for Apple Review. A v2 callable `createAppleReviewAccount` rejects duplicate accounts, generates a unique long password server-side, creates a Firebase Auth user, assigns the server-signed custom claim `apple_review: true`, and creates a `platformUsers` profile with role `apple_review`. Credentials are returned **one time to the administrator only**, not written to GitHub, client storage, console logs or Firestore.
4. The reviewer opens the iPhone app and taps **ACCESSO REVISORI APPLE** on the existing login screen.
5. `apple-review-signin.html` uses a *secondary Firebase Auth app with persistence NONE*; after sign-in it forces an ID token refresh and checks `claims.apple_review === true`. If false, it signs out and rejects the user. If true, it signs out and opens `apple-review-demo.html`.
6. The reviewer sees synthetic commesse/impianti/squadre/ore and can simulate interventions locally without modifying anything real. The normal user login remains unchanged.

## Security coverage
- Firestore and Storage rules merged in PR #1586 exclude signed `apple_review` identities from all operational records and file paths. Only `appleReviewDemo` is readable, and never writable, by reviewers; the existing offline demo does not require that collection.
- `functions/apple-review-guard.js` checks the signed Firebase claim inside every directly callable app backend endpoint from `functions/index.js`, central-notifications, shared-calendar-view, shared-operational-views, whazzup-pdf-drive, varga-session-transfer, error-center, error-reporting, password-recovery-code, varga-gestionale-sync and user-access-approval.
- Firestore scheduled triggers, change listeners and public non-sensitive fuel-station endpoints do not grant extra client mutation rights.
- Existing anonymous recovery and registration functions can still be invoked without a Firebase session (pre-existing behavior). The reviewer cannot use its authenticated token to bypass the new guard. Do not mistake this restriction for a full application penetration test.
- Server-signing keys, passwords and other secrets are never embedded in native or web assets.
- The new administrator creation callable must be invoked only **after** deployment of the backend guards. **No credentials exist yet.**

## Validation
- `node scripts/check-apple-review-account.js`
- `node scripts/check-fatto-whazzup-immutable.js`
- `node scripts/check-fatto-whazzup-isolation.js`
- `node scripts/check-auth-stability-contract.js`
- Browser test reviewer valid/invalid claims on staging with real Firebase Auth; check no normal auth session is disturbed; verify that all callable endpoints respond PERMISSION_DENIED for a reviewer.
- Test real production Firebase access **with test-only account** after safe backend deploy, and inspect logs for denied operations.
- Test the new iOS build on an iPhone using these credentials; build 5 predates the new login and cannot be used.
- Remember: Apple App Review guideline 2.1(a) asks for a real demo account or a fully featured demo. A reviewer who can only see/simulate a subset of functionality may be insufficient and Apple can request more access. Explain the demonstration-only restrictions in App Review Notes.

## App Store Connect after deployment and test
Go to `Varga Cantieri → iOS App Version 1.0 → App Review Information`. Enable **Sign-in required** and paste the username and generated password from the administrator page into the fields. Add the following notes, adjusting based on actual device testing:

> VARGA CANTIERI is an operational management application. For privacy and safety, Apple receives a dedicated restricted account. On the first sign-in screen, tap "ACCESSO REVISORI APPLE" and enter the supplied credentials. Successful authentication opens an isolated example workspace with synthetic projects, facilities, teams, hours and documents. Sample interventions can be simulated; real business operations, staff records and modifications are unavailable. Contact us if additional access is required.

Keep **manual release** and check the privacy statement, screenshots, age rating and export compliance. Do not hit Add for Review until the login works in the current uploaded build.

## Safety gates
1. Do not merge this PR without green checks and a review of server changes.
2. Because some other workflows deploy Firebase automatically on a push to `main`, merging this PR is a **real production backend change**. Confirm readiness of every callable guard and compatibility with existing functions first.
3. The admin creation tool is accessible only after the backend deployment succeeds; do not create account before that.
4. Rotate/revoke the reviewer account after App Review if access is no longer required.
