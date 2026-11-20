# Integration verification: staged privacy-review workflow (local emulators)

This regression test exercises the **actual unmodified source** of both
`functions/account-deletion-request.js` and
`functions/account-deletion-review-staged.js` using real **local** Firebase
Auth and Firestore Emulator instances, with a wholly synthetic Firebase project
`demo-deletion-review-e2e`. All user identities are generated at runtime and
deleted from the emulators at completion.

**Evidence tested:**
- Real local Firebase Auth user creation and current custom claims; a request
  with a forged or stale reviewer token is rejected.
- Revoking the privileged review claim or disabling the operator in the
  Auth Emulator immediately blocks the callable even with an old valid token.
- An authenticated user requests their own deletion through the genuine
  request-intake body, and a separate reviewer sees only minimal review
  metadata from actual Firestore queries.
- Two simultaneous reviewer transitions against one pending request are
  serialized by true Firestore transactions; only one may take ownership.
- Transitions `pending → in_review → rejected` with fixed allowed reasons.
  Invalid `completed` action, arbitrary text, or wrong state fails closed.
- After rejection the owner can re-request, with timestamps and reason
  history preserved. A second user's record stays pending.
- Original Firebase Authentication accounts, representative business
  profiles and private documents remain intact throughout.

The dedicated CI workflow is **NO DEPLOY** and read-only with respect
to production. It fails if there are real Google credentials, a Firebase
project other than the demo ID, or non-loopback emulator endpoints.

This is **NOT** a demonstration that real-account deletion works or is
lawful. Both modules remain staged, are not exported into the normal Firebase
production functions entrypoint, and no review operator is granted new
privileges. Apple Guideline 5.1.1(v) still requires validated operational
fulfillment, retention/GDPR approvals, notification of completion and a real
iPhone test/video. The App Store release readiness gate stays blocked.
