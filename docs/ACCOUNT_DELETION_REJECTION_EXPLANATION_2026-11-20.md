# Account-deletion rejection explanation — staged and non-destructive

The staged privacy request/status service and the ordinary iPhone UI now
support a **safe, human-readable rejection explanation** without exposing
notes from other staff, enterprise records or free-text operator content.

### Status response and privacy boundaries
- Only the signed-in account owner can ask about their own request; its
  Firestore document path comes exclusively from their verified Auth UID.
- The backend returns `reasonCode` **only when status is `rejected`** and
  its value is one of three fixed codes: `identity_not_verified`,
  `retention_review_required`, `request_needs_clarification`.
- Unrecognized values, previous reviewer notes, request metadata,
  employee names and all other database fields never leave the callable.
- When the owner resubmits, status becomes `pending`; old rejection metadata
  remains in Firestore for audit but is **not disclosed** during the pending
  review cycle.
- The client maps only the three approved codes to clear Italian messages,
  explicitly states **no account was deleted**, and explains that a user can
  sign in freshly and file a new request. Unknown or missing codes use a
  neutral fallback, **never raw backend strings**.

### Verification
The existing staged backend mock test, real local Firestore Emulator
integration and real UI synthetic runtime now include denial of other-user
access, known/unknown reason codes, absence of free-text leaks and
reopening of an existing rejected request.

### Release status
This is a preparatory source code change only. The staged functions remain
excluded from `functions/main.js` and routine production deployment.
The UI feature flag remains **off**, pending a legally approved privacy
fulfillment procedure and SLA. No account/employee data is changed;
no credentials or privileges are granted and no TestFlight/App Store upload
is performed. These tests do not prove full Guideline 5.1.1(v) compliance.
