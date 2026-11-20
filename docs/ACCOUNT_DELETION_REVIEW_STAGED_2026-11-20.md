# STAGED privacy-request operator queue — 20 November 2026

This repository contains a separate **NOT DEPLOYED** operator backend:
`functions/account-deletion-review-staged.js`. It is NOT imported in
`functions/main.js`, is NOT on the standard Firebase deployment allowlist,
has no browser interface, and does not modify `firebase.json`, production
Firestore, user identities, claims or App Store distribution.

## Purpose

A manual account-erasure process needs an operator to see received requests
and to record review decisions. The current request intake supports
`pending`, `in_review`, `rejected`, `completed`. This module is limited to:
- `listVargaAccountDeletionReviewQueue`: see at most 25 `pending`/
  `in_review` requests and minimal UID/status/reopened fields;
- `reviewVargaAccountDeletionRequest`: atomically change `pending` to
  `in_review`, or `in_review` to `rejected` using one of three **fixed
  reason codes**. It preserves the rest of the document. No arbitrary
  free-text personal data can be submitted.

These callables require **both** `admin` and
`privacy_deletion_reviewer` custom claims in the caller's freshly issued
Firebase Auth ID token **and** the currently active Firebase Auth user record.
No account is provisioned and **no such claim is granted** by this PR.

## Intentionally impossible operations

There is **no** callable to:
- set `status=completed`, which could falsely certify erasure;
- perform bulk account deletion or grant account privileges;
- read/write documents, GPS, timecards, businesses, Drive, or Storage;
- alter real user accounts or change existing retention policies.

A separate protected and reviewed **fulfillment** process must verify
which private data can be erased and what must lawfully be retained, then
delete the account, revoke Apple credentials if applicable, produce an
auditable completion record, and send a confirmation. This module does not
claim to fulfill that process. A rejected request remains reopenable by
its authenticated owner through PR #1658.

## Verification before any possible deploy

Tests: `node scripts/check-staged-deletion-review-queue.cjs` and the
no-deploy GitHub Actions workflow. These use only synthetic in-memory
documents and Firebase stubs, and fail if the module leaks into the
routine functions entrypoint.

**Release gates remain BLOCKED** pending the data-controller's
authorizations, privacy impact assessment, verified deletion operation,
realistic service time/SLA, exact roles and scoped secret management,
separate protected workflow, sandbox tests and real-device Apple
evidence. Even after code review this queue is NOT ready to deploy.

Apple's guidance: https://developer.apple.com/help/app-review/guideline-reference/5-1-1-account-deletion/
