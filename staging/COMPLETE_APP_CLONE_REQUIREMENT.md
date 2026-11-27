# Owner directive — literal full Varga Cantieri application reuse for every company

This document records the clarified requirement: the deliverable is NOT a commesse-only feature, a reduced tenant portal, or a new app that resembles Varga Cantieri. It is the complete existing Varga Cantieri application, copied/reused from this repository and made multi-company.

## Required result

Every company user must receive the same complete application that exists in the original Varga Cantieri codebase: same application shell, screens, navigation, visual design, forms, dialogs, menus, map, workflows, business rules, and feature set. Use GitHub source as the canonical implementation. Do not rebuild the features from descriptions and do not replace the application with a simplified React dashboard.

The only intended company-specific differences are:
- authenticated tenant context and tenant-scoped data paths;
- membership/role permissions;
- company-owned settings and integration credentials;
- optional modules/features enabled by authorized company configuration.

A new company starts with an empty company-owned operational archive, but with the full application and all original functions available according to permissions. No Avola Coop operational records, employees, customer data, Drive credentials, Sheets, WhatsApp recipients, or other private company configuration may be copied into another tenant.

## Full feature coverage — not only commesse

The implementation must include and verify the complete original application modules, including (where present in the current source):
- authentication, account approval, password recovery, profile and privacy/account controls;
- home, main menu, navigation, notifications, banners, information and help;
- commesse, subcommesse, impianti/cantieri, management hub, map, search, import/export, prices, accounting, work items and documentation;
- operational execution: Da fare/Fatti, FATTO, notes, photos, date/time/operator, WhatsApp/WHAZZUP, navigation and offline/sync;
- teams, personnel, users, membership and role management, vehicles and equipment;
- work scheduling, calendars, working hours, attendance/timbrature, leave/absence and approvals;
- occasional work, mowing/COBO work orders, tree work/pruning, tree registry, green areas, urban furniture, wastewater plants/infrastructure and related specialist modules;
- documents, POS/safety documents, private documents, reports, exports, Google Drive/Sheets integrations where separately configured;
- snow service, road salting, routes and other existing modules;
- administrative, diagnostics, backup/recovery and support tools, restricted to authorized roles.

This list is a coverage checklist, not permission to expose every admin function to every employee. Preserve original role restrictions and enforce tenant isolation server-side.

## Canonical source and integration strategy

- Canonical source: repository root application and its actual source/assets/scripts/styles, especially root `index.html`, `app.js`, `management-v2.js`, `accounting-v2.js`, `commessa-impianti-menu.js`, and all scripts loaded by the original `index.html`.
- The current `staging/portal/src/App.jsx` is only a tenant administration/login prototype. It is not the final operational application and must not be presented as complete.
- Integrate the full original application shell into the multi-tenant staging experience, preserving original modules and their dependencies. A tenant selector/login may be the new entry gate, then the user enters the original complete application.
- Implement tenant-aware data access centrally and comprehensively. Audit every Firestore collection path, query, listener, local cache key, storage path, import/export, Drive/Sheets/Functions integration and offline queue. Do not merely change `getCommesseCollectionName()`: all tenant-owned modules and data must use the authenticated tenant context.
- Tenant ID must be derived from verified authentication/membership and not trusted from an editable browser field or local storage.
- Maintain compatibility with the original Avola Coop production application and existing data. Do not point staging to production or alter production configuration.
- Keep the original user experience unchanged except the necessary company selection and tenant-specific labels/settings.
- Do not invent a substitute UI for any existing feature.

## Protected original behavior and safety

Follow root `AGENTS.md` and the staging requirements. In particular, preserve the protected FATTO and WhatsApp/WHAZZUP code and behavior; preserve personnel/users and their historical IDs and relationships; do not perform destructive migration, bulk deletion or automatic copying of private company data. Where the multi-tenant product requires tenant-specific FATTO behavior, follow the already recorded approved policy and keep Avola's protected vault unchanged.

## Required delivery and acceptance

The work is complete only when:
1. the entire original application, not just the commesse screen, is accessible after tenant selection;
2. all original modules and their actual actions have been inventoried against root `index.html` and tested;
3. a newly created tenant sees the complete app with empty company data;
4. creating a commessa and inserting cantieri/impianti uses the original forms and flow;
5. the same tenant-scoped behavior is verified for every other module, including nested data, files, cache, offline operations and integrations;
6. automated Firestore emulator tests prove that tenant A cannot read or write tenant B data;
7. original Avola data, staff, users, links, protected FATTO/WhatsApp behavior and production remain intact;
8. staging build, browser tests, mobile layout and module regression tests pass.

Do not declare the work finished after creating the tenant shell, a preview, a mock, a partial module, or a commesse-only integration. Do not deploy or submit a native build without a separate explicit instruction.
