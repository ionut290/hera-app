# Netlify: web-only publish directory (20 Nov 2026)

**Production site:** https://creative-syrniki-dddbae.netlify.app/

On 20 November the live Netlify deploy was checked using its connected deployment
API. The production deploy `6b005e6b178451b673910e88` is READY from
GitHub SHA `ddf37da96792de845eefd6bbe9827c0a6cd54b73` (PR #1651).

## Previous exposure risk

`netlify.toml` set `publish = "."`, the repository root. Netlify could
therefore publish internal `docs/`, `scripts/`, `package.json`, CI details
and other technical artifacts as ordinary static files. Firebase Hosting's
separate ignore list does not protect Netlify.

## Safer build

The Netlify build now:
1. stamps the PWA build version;
2. generates the tested native/web asset shell in `www/`;
3. removes all root JSON and non-web file types and copies safe root image/font
   assets without recursing through directories;
4. verifies critical PWA/script assets and rejects internal repo file paths;
5. publishes **only** `www/`, not the repository root. Netlify serverless
   functions continue to build separately from `netlify/functions`.

This isolates technical files from the site. It is a deployment safety change,
not a modification of Firebase Auth, Firestore/Storage rules or user data.

## Verification before treating as released

- GitHub CI checks and Netlify deploy-preview build must succeed.
- Inspect the preview homepage, sign-in UI, `/delete-account.html` and
  key JavaScript/CSS assets without authenticating with real employees.
- Confirm previously directly accessible technical URLs no longer serve raw
  repo data (SPA fallback may still return `index.html` with HTTP 200).
- Confirm production deploy actually includes this main commit, then check
  live website. Do not equate the preview or GitHub merge with a deploy.
- No claim is made about Apple account deletion compliance or iOS TestFlight.

Sources: https://docs.netlify.com/build/configure-builds/overview/
