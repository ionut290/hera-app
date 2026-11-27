# Reuse the original Varga Cantieri commessa and impianti experience

## Non-negotiable product requirement
Do not build a second commessa/cantiere interface in the React tenant portal. Every company must use the same original Varga Cantieri screens, forms, validation, mobile management hub, map, and existing actions. The only company-specific difference is the authenticated tenant data scope and the permissions assigned to that member.

## Source-confirmed original UI integration points
- `commessa-impianti-menu.js` creates the existing mobile commessa management panel and calls `window.AccountingV2.openMobileHub(commessa)`. It exposes the original actions: Aggiungi impianto, Modifica impianti, Importa impianti, Esporta impianti, Prezziario, Vista completa.
- `accounting-v2.js` reads the selected commessa through `db.collection(getCommesseCollectionName()).doc(commessaId)`; the existing accounting screen loads the original `impiantiFisici`, `lavorazioni`, `prezziario` and operational `impianti` subcollections.
- `commessa-navigation-repair.js` preserves the original `selectCommessa` navigation and the selected commessa context.
- Root `index.html` loads the existing original CSS and feature scripts, including `management-v2.css`, `accounting-v2.css`, `documentazione-cantiere.css`, `cobo-mowing-work-orders.css`, and the map/runtime scripts.
- Root `firestore.rules` currently defines the legacy top-level `/commesse/{commessaId}` parent; its nested subcollections are readable/writable to signed-in accounts. Do not assume this already supplies tenant isolation.

## Required integration approach
1. Preserve the original root Varga Cantieri app shell and its original commessa/impianto UI. Do not substitute a simplified portal or create new forms with different fields.
2. Introduce a tenant-scoped data-context adapter at the boundary used by `getCommesseCollectionName()`, so the existing `db.collection(getCommesseCollectionName()).doc(id)` and all child collection calls resolve to `tenants/{tenantId}/commesse/{commessaId}/...`.
3. Ensure the tenant is selected and validated from the authenticated membership before the original app initializes any commesse subscriptions or reads. Never accept tenantId from a freely editable client field.
4. Preserve the original commessa form, original impianto form, map and autocomplete, original validations, existing import/export and management flows, and the original `AccountingV2.openMobileHub` entry point.
5. Port and test all required Firestore rules for the nested tenant tree. A parent commessa can only be created/edited by the tenant roles authorized by the product; all child data must remain inside that tenant. A user's membership in tenant A must not grant access to tenant B.
6. Keep the FATTO workflow and its WhatsApp/movement actions byte-for-byte unchanged unless a separate reviewed change is required. Tenant scoping must be applied beneath the data-access layer without altering its business actions.
7. Do not seed any commesse or cantieri automatically for new companies. Each company starts empty and uses the same original Nuova Commessa / Aggiungi impianto flow to create its own records.

## Important distinction
In the original UI the mobile hub action is labelled `Aggiungi impianto` and the accounting subcollections include `impiantiFisici`, `lavorazioni`, `prezziario`, and `impianti`. The user calls the work locations cantieri; retain the original labels and data model wherever the existing app uses impianti.

## Work status
The current `staging/portal/src/App.jsx` is a simplified tenant landing page and is not the desired final user interface. The generic `staging/commesse-store.mjs` is a staging data-layer prototype, not a replacement for the original Varga UI. Do not mark the commessa/impianto experience complete until the original app shell and workflows are running against the selected tenant path and pass emulator/browser checks.
