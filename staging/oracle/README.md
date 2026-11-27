# VARGA CANTIERI 2.0 — Oracle backend candidate

Prototype on `varga-cantieri-v2-oracle`, isolated from V1, V2 staging, and radio infrastructure. No production migration or deployment. PostgreSQL hosted on OCI is the intended target, not Oracle Database SQL.

## Implemented
- Node HTTP API with Firebase Admin verification of V2 staging ID tokens, revoked-token checks and verified email.
- Company creation with idempotent request IDs, membership list, company-code joins, admin acceptance/rejection.
- Commesse and impianti: list/create; composite tenant keys, row-level tenant scope and server-side quotas (1 commessa, 10 impianti, 1 admin owner, 3 accepted operators). Tenant row locks serialize quota checks. Only admins create operational records in this initial API. Super Admin plan changes use only verified token claims, lock the company, enforce current usage and record an append-only audit. The administrator count remains fixed to one until role-management APIs are added. No client-supplied role accepted.
- Synthetic PostgreSQL integration test. Test authentication is injected only in the test process; server entry always uses Firebase Admin.

- Authenticated API client in client.mjs, separate from the existing launcher. No Firestore fallback; session changes invalidate pending responses.

## Interface
`GET /health`; bearer token required on `/v1/*`.
`GET/POST /v1/companies` (POST: name, requestId UUID).
`POST /v1/join` (code).
`GET /v1/companies/:id/requests`; `POST /v1/companies/:id/requests/:uid` (decision accepted/rejected).
`GET/POST /v1/companies/:id/jobs` (name).
`GET/POST /v1/companies/:id/jobs/:jobId/plants` (name, optional paired lat/lng).

GET /v1/admin/companies and POST /v1/admin/companies/:id/plan require the verified superAdmin claim. Full-plan requests specify commesse, impianti, operatori; trial restores 1/10/3 only if current usage permits it. The list is limited to 100 records pending pagination.

## Before deployment
Use a dedicated OCI instance/network, separate from radio; verify free resource availability, backups and cost first. PostgreSQL database varga_v2_oracle; apply both SQL migrations in order; migration owner separate from LOGIN role varga_app (no superuser, no BYPASSRLS, no membership of migration owner). Keep DB private, API behind TLS reverse proxy with rate limits (especially company-code lookup), monitoring, audit and secret storage. Configure minimally privileged Firebase Admin credentials only for V2 staging token verification; do not copy credentials from V1. Do not expose port 8080 publicly. Lock dependency versions and review advisories before live deployment. No credentials in Git.

## Remaining work
Integrate a client provider with existing launcher; migrate operational modules without pretending the REST API is a full Firestore SDK. Wire plan controls to the interface and add role management, updates/deletes and additional resources, realtime events/offline queue, Object Storage with signed URLs and tenant checks, full migration parity and live device authentication. Use an isolated staging copy with synthetic data first. No production data migration or Firestore shutdown until end-to-end verification and rollback plan.
