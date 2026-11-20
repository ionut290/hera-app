# Apple sign-in and account deletion UI — synthetic regression test

**This is test-only code and does not change the app UI or enable deletion.**

The workflow `check-apple-account-privacy-ui.yml` runs
`scripts/check-apple-account-privacy-ui-runtime.cjs` without a network
connection, Firebase Admin SDK or real Apple credentials. The test executes
the actual browser entrypoint `account-privacy-controls.js` with synthetic
DOM, Firebase Auth, callable and Capacitor Apple plugin adapters.

It verifies Apple sign-in wiring on iOS, failure if the native Apple plugin is
missing, default-hidden deletion actions, independent flags and a validated
1–90-day SLA, request/status controls for normal/pending/menu/login gate,
confirmation payload not containing any UID/email, rapid double-tap guard,
signed-out behavior, already-completed state and mutual exclusion between
staged request mode and immediate destructive demo mode.

**Not yet verified**: working Apple login or token revocation on a real
iPhone, live Cloud Functions request/status deployment, real deletion
fulfillment and legally approved retention, Apple reviewer access and physical
device screen recording. TestFlight 1.0 (20012) is available for internal
testing, not approved for public App Store distribution.
