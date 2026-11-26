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

## Operational validation checkpoint

The emulator integration runs original handler code in the Firebase SDK's realm to avoid cross-realm plain-object validation errors. The isolated browser smoke now delivers an authenticated fixture user to the original auth callback and requires completion of the transformed company approval gate. It does not represent a live login or successful completion of every startup subscription. External services stay blocked. Continue autonomously on staging with verified commits and explicit readiness checkpoints, not estimated progress alone.

## Legacy authentication guard integration

The scoped auth facade exposes the existing original module's already-installed guard marker, and its observer verifies the actual selected-company membership before delivering a user. It does not create `platformUsers` or overwrite the original module. The browser fixture returns empty read-only responses for the two startup indexes; these fixture responses are not real permissions. External services and unsupported archive access remain denied. Continue verification with the real emulator and preserve the production guards.

## Snow creation checkpoint

The emulator candidate enables the original `createCommessa` handler for `neve_commesse`, with parent validation in the same snow archive, and the original `app-snow.js` road-list batch handler. Both are gated by explicit demo capability flags and admin membership. The test checks two companies, no green/snow leakage, and denied completion writes. Route edits, automatic completion, snow-specific price lists, existing local catalog data, and all production/native activation remain outside this change.

## Startup archive alignment

The company runtime router now matches the candidate read-only rules for oreLocks, normal and snow personnel/vehicles, and the three exact original appConfig startup documents. The appConfig collection itself, arbitrary settings, credentials, writes and cross-company reads remain denied by emulator rules. The integrated emulator verifies the original startup query shapes and repeated document IDs across two companies. No rules deployed; no production records touched.

## Integrated delivery block — local emulator package

The new `staging/portal/build-full-runtime.mjs` builds a single original application package: local Firebase compat login, real verified membership selection, original HTML and original source through the existing scoped runtime, local Leaflet and an isolated test frame. The only configured Firebase endpoints are loopback Auth/Firestore emulators. No production API key, native secrets, deployment or signing. An artifact is produced only for repository-internal local collaudo. Original independently booting snow pages remain quarantined pending separate isolation; the original root snow module is included. The separate manual iOS simulator compile uses a different bundle ID and no credentials; it is not a TestFlight build. Full authenticated browser acceptance with real emulator accounts, storage/Functions, protected completion and native capabilities remain release gates. This block replaces incremental preview work; never publish this package.

## Local package progress

Real emulator Auth, verified company A, original iframe and tenant-specific document read passed; iframe-origin commessa write and backend count also passed before the test raced the asynchronous company B switch. The switch wait now tolerates the expected closed-session interval, without weakening access rules. iOS isolated compile is still unsigned, not a device/TestFlight build. Keep live staging rules and production unchanged until acceptance passes.

## Delivery checkpoint: iOS compile and backend transition

GitHub run 45753044631 successfully compiled the unsigned iOS simulator app with Xcode 26.6 and separate bundle ID it.vargacantieri.staging. Real integrated package run 45752360354 passed browser Auth, isolated company data reads/writes, switch and logout. The launcher now delegates Firebase initialization to `firebase-runtime-provider.mjs` (local emulator only); this is a first separation for the owner's agreed future Oracle migration, not an Oracle implementation. The Super Admin candidate rules are claim-based and permit metadata only, strictly non-destructive with atomic membership role/status mirror updates, and no broad operational bypass. Both CI workflows assemble the same candidate rules through a single generator. No live Firebase rules or production changes made.

## Staging rules assembly correction and connection preparation

Literal callback substitution fixes JavaScript's special `$'` replacement syntax when inserting the unchanged COBO regex. A unit test checks every fragment is byte-for-byte included. `connectStagingFirebase` is an unwired, explicitly native-only connection provider with fixed project ID, app ID and auth domain, separate from the already verified local emulator. It does not activate the original frame against live data. Real staging rules, staging account configuration, operational isolation and native permission regression are still release gates. No deployment or billing change.

## Separate build modes (no release activation)

The build now explicitly accepts `local-emulator` (default, existing acceptance test) or `private-staging` (separate output directory, requires validated staging web config). The latter can authenticate and inspect the user's own demo memberships but cannot open the operational original frame; it is not ready for TestFlight. Compile-only CI uses a synthetic key and never connects to Firebase. The local frame receives a scoped auth facade rather than raw auth; compatibility must be retested. Production sources, Firebase rules and native plugins remain untouched. Before any live operational beta, require verified rules, staging test identities, protected FATTO/WhatsApp regression, review of persisted/native data isolation, and Apple signing.

## Native device compile gate

The separate native workflow now chooses private-staging login-only mode if the staging-specific public client config is supplied, otherwise stays local. It checks both simulator and unsigned physical-device builds, with the independent bundle ID and no production native Firebase config. The private staging launcher verifies the selected membership fresh through the same company entry service, then closes it without opening operational UI. No signing or publication. The new rules are not deployed.

## Release review artifact

After the entire Firestore emulator suite succeeds, CI now exports the exact byte-for-byte tested candidate rules, source revision and SHA-256 manifest. The bundle is review-only; it intentionally contains no deployment config or automatic publish command. It preserves the distinction between an historical selected rules version and current active rules. Do not overwrite an unpublished console draft. Protected FATTO/WhatsApp execution and full live operational activation remain gated by the project instructions. Native build may succeed independently; it is not TestFlight readiness.

## Actual device compile checkpoint

Run 45753988405 passed both iPhone simulator and unsigned iPhone device architecture builds, but its environment still lacked the staging web key and therefore compiled the local-emulator package. Run 45753988347 passed the full original emulator browser acceptance on retry. The native workflow now requires the staging-specific public client configuration, fails closed without it, and validates the private-staging marker in the native device bundle. This candidate remains login/membership only; no signing, operational enablement, Apple upload or production changes. The console audit could not conclusively recover the active rules and detected an unpublished deny-all draft; do not overwrite or publish it.

## Integrated Super Admin panel

The full application launcher now opens the Super Admin panel directly when the verified server token contains the `superAdmin` claim. Local emulator mode supports creation/rename/suspension of demo companies and atomic role/status updates for existing demo memberships; server rules deny destructive operations and cross-company changes. Private live staging mode is read-only until the actual ruleset is established. The integrated browser test exercises real emulator Auth custom claims, company creation and mirrored membership role changes, checking preserved historical IDs and unaffected second company. No protected FATTO/WhatsApp source, production or real personnel records changed. This is not a live release.
