# Staging integration gates

Do not deploy or run the legacy application against production.

1. Run `node --test staging/tenant-context.test.mjs`.
2. Import the existing staging portal from Cloud Shell, without credentials or local environment files.
3. Use a separate browser origin and staging Firebase web configuration. Verify the project ID before initializing Firebase.
4. Keep legacy index.html quarantined until all direct endpoints, service workers, native plugins, cached data and deployment workflows are audited.
5. Inventory legacy collections and queries. Migrate one module at a time to tenants/{tenantId}/... with explicit Firestore rules and emulator tests. Never fall back to root production collections.
6. Start with synthetic Avola-style green maintenance and worksite records. No real employee, customer, Drive or Sheets data.
7. Preserve legacy UI and workflows; do not enable writes until tenant isolation tests pass.
8. Integrate snow service only after its source and dependencies are verified.
9. Keep Spark billing disabled. Do not deploy functions or storage or replace existing staging rules with legacy rules.
10. Native builds require separate bundle IDs, signing configuration and Firebase files; no TestFlight upload until explicitly reviewed.

The frontend tenant context is not a security boundary. Server-side rules are mandatory.
