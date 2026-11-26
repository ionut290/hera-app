# Full original runtime cutover — staging only

The owner has chosen the complete original root application over rebuilding isolated preview screens. The root source is already in this branch. Freeze new reconstructed preview features; preserve existing tests as regression fixtures.

1. Inventory original runtime entrypoints, direct/dynamic Firestore calls, storage, cache and network integrations using `node staging/audit-original-runtime.mjs`. The report is a candidate list, not proof of isolation.
2. Use `createOriginalRuntimeSession` as the candidate authenticated compat entry boundary; it verifies active membership on the server, closes the old scope when switching, and rejects production. It does **not** install itself into app.js, and does not revoke SDK references already issued: server rules and listener teardown remain necessary.
3. Validate all root accesses, nested paths, identity/profile reads, storage, Functions, native plugins, offline/cache, and tenant-specific integration settings. Keep production-only hardcoded content and private data out of new companies. The original protected FATTO/WHAZZUP functions and personnel IDs must not be changed.
4. Integrate the original runtime only after those gates are verified, then run full browser/mobile and tenant isolation tests. Keep the Netlify deployment block and native submission block. No paid services or production config changes.

No live data migration, Firebase deployment or TestFlight upload is authorized by this source change. The demo portal is not the operational app.

## Candidate source integration

`original-runtime-source.mjs` prepares the complete original `app.js` in memory by replacing only the initial SDK acquisition statements. The generated candidate is syntax-checked and reversibility is tested. The startup auth, legacy approval/profile collections, local/native caches, direct SDK access in other modules, messaging, Functions and external integrations still require review. Do not load the generated candidate in the portal/native app or remove quarantine. Existing listeners and references are closed by the session scope; in-flight writes cannot be recalled and remain governed by server rules.

## Startup authorization and cache candidate

`original-runtime-entry.mjs` combines the tested session scope with a fresh server membership check and a compatibility profile. Legacy built-in email privileges and globally saved profiles are not inputs. Its two bounded storage objects are disposable, in memory, and never touch the device's existing cache or offline outbox. These objects are NOT installed into `window` or the original app yet. Native Preferences, IndexedDB, service workers, direct global SDK calls, legacy approval UI and protected offline queues still require explicit integration and regression testing. Do not discard existing production offline records. No real runtime activation or deployment.

## Owner priority — fastest verified path to TestFlight

Use the complete original runtime; stop adding reconstructed screens. Prefer direct integration and automated concrete checks over further independent prototypes. At each session report verified changes, CI status, remaining blockers and TestFlight readiness. Preserve original production and protected FATTO/WhatsApp behavior, no paid services. Do not report untested features as complete.

The new integrated candidate connects the original startup gate, company role and in-memory cache identifiers in `app.js` itself (in-memory source transformation only). Tests execute the actual transformed startup gate, not a separately recreated login screen. Other original modules' direct Firebase, storage and native accesses remain blocked from live startup until isolation is validated. This is not a native build or a live application launch.

## Original shell browser smoke

`staging/portal/test-full-original-runtime-browser.cjs` serves the actual root HTML and local original scripts in a headless browser with the integrated app.js candidate. Network requests cannot reach external hosts; no Firebase token or live database is provided. It checks original shell elements and the actual candidate SDK bootstrap, and logs remaining browser exceptions as diagnostic blockers, not proof of operational readiness. The root CSP is not changed on disk. This test-only fixture cannot be deployed.

## Shared legacy SDK namespace

`legacy-firebase-namespace.mjs` gives original scripts calling `firebase.firestore()` or `firebase.app().firestore()` the same selected-company facade. The browser smoke now exercises this namespace; external Firebase services and creation of alternate apps are refused. Auth profile views are limited; the original app candidate still has its trusted entry auth reference. Do not treat this client namespace as a substitute for server-side rules, and do not enable live operation yet. Direct SDK static types and module-specific permissions must be reviewed in actual runtime tests.
