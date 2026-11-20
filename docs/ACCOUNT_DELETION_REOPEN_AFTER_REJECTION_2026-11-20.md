# Account deletion request: retry after rejection (STAGED ONLY)

20 November 2026 — The authenticated user must be able to ask again if an
earlier privacy review rejected or closed a request. Previously the staged
`requestOwnVargaAccountDeletion` API returned `rejected` forever without
opening a new review cycle; the app could not submit a renewed request.

The staged, **non-destructive** callable now:
- requires the SAME fresh Firebase ID token (10-minute `auth_time` limit),
  explicit `RICHIEDO ELIMINAZIONE` confirmation and current caller UID;
- when the own record's status is `rejected`, atomically changes only the
  request status to `pending`, sets new server timestamps, increments
  `reopenCount`, and preserves earlier reviewer notes/metadata for audit;
- remains idempotent: subsequent calls while pending/in_review/completed
  do not write again; corrupt state/history fails closed;
- does not delete Auth accounts, enterprise documents, Storage, staff data
  or other users' records;
- is covered by both the isolated synthetic runtime and a real local
  Firebase Firestore Emulator transaction integration.

**NOT DEPLOYED TO PRODUCTION.** The feature flag remains disabled and the
callable is excluded from the default Firebase Functions deployment. The
protected `deploy-account-deletion-request-protected.yml` requires the
documented privacy approval, verified manual fulfillment and explicit
reviewer approval. This patch neither satisfies full Apple Guideline
5.1.1(v) compliance on its own nor authorizes a new App Store submission.

According to [Apple's official account deletion guidance](https://developer.apple.com/help/app-review/guideline-reference/5-1-1-account-deletion),
manual fulfillment can be acceptable if it takes a reasonable amount of
time, users know when to expect completion, and receive a final
confirmation. A request button alone with no actual fulfillment is
insufficient. Real processing/retention signoff and a physical device
screen recording remain required.
