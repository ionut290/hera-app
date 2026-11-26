# Full original runtime cutover — staging only

The owner has chosen the complete original root application over rebuilding isolated preview screens. The root source is already in this branch. Freeze new reconstructed preview features; preserve existing tests as regression fixtures.

1. Inventory original runtime entrypoints, direct/dynamic Firestore calls, storage, cache and network integrations using `node staging/audit-original-runtime.mjs`. The report is a candidate list, not proof of isolation.
2. Use `createOriginalRuntimeSession` as the candidate authenticated compat entry boundary; it verifies active membership on the server, closes the old scope when switching, and rejects production. It does **not** install itself into app.js, and does not revoke SDK references already issued: server rules and listener teardown remain necessary.
3. Validate all root accesses, nested paths, identity/profile reads, storage, Functions, native plugins, offline/cache, and tenant-specific integration settings. Keep production-only hardcoded content and private data out of new companies. The original protected FATTO/WHAZZUP functions and personnel IDs must not be changed.
4. Integrate the original runtime only after those gates are verified, then run full browser/mobile and tenant isolation tests. Keep the Netlify deployment block and native submission block. No paid services or production config changes.

No live data migration, Firebase deployment or TestFlight upload is authorized by this source change. The demo portal is not the operational app.
