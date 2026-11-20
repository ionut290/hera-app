# VARGA CANTIERI — Firebase Hosting release boundary (20 Nov 2026)

**Status: PREPARATION ONLY. No Hosting deploy, no rollback, no Firebase backend changes.**

## Why this matters

`firebase.json` uses `hosting.public: "."`. Firebase Hosting will consider files
below that root except the configured `ignore` globs. The initial config excluded
`functions/**`, `node_modules/**`, dotfiles and a few files, but NOT the internal
`docs/**`, `scripts/**`, native build directories, archives, local PDFs and
other sensitive file types. Blindly running `firebase deploy --only hosting` is
therefore unacceptable.

The repository now adds a targeted exclusion list to its Hosting config.
An offline test checks for accidental removal of that list and presence of essential
runtime entrypoints. The normal web app still uses root-level `index.html`,
`app.js`, `sw.js` and `account-privacy-controls.js`.

**A denylist is not a full allowlist.** It cannot prove that other root-level files
are safe to publish or that the currently live site contains the same assets.
There is also no single-file Firebase Hosting production deploy: deploying a
directory can replace the entire live Hosting release.

## Release checklist — not verified yet

1. Inventory every file Firebase CLI would upload and audit against a strict
   allowlist of web runtime assets. Validate that the app, PWA, icons, fonts,
   styles and privacy page still work when packaged.
2. Compare the active Hosting release and custom domain/site with the intended
   app entrypoint. Read-only inspection on Nov 20 found main Firebase Hosting
   `hera-app-6cd2b` at its Apr 8 release; the public deletion page could NOT
   be verified on the Firebase domain. Do not claim it is published.
3. Prefer a dedicated generated web-only artifact directory for future Hosting,
   with `firebase.json` pointing to it only after a fully verified migration.
4. Obtain independent approval of public content, privacy text, data-retention
   and actual business/process timing, review preview deploy separately.
5. Only then use a protected manual deployment; afterward verify publicly
   served `/delete-account`, its provider-neutral wording and real URLs.
6. Do not treat an informational contact page as fulfilling Apple 5.1.1(v).
   In-app authenticated account deletion and a functional completion process
   remain independently required.

Official Firebase reference: https://firebase.google.com/docs/hosting/full-config
