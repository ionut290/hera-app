# VARGA CANTIERI 2.0 — shared application

Product requirement: every new company receives the same complete application, navigation, visual design and operational capabilities as the existing Avola Coop application, with an initially empty company-owned archive. No reduced edition or separate fork per company.

- One shared codebase; tenant-specific data and settings.
- New companies start with zero operational records. Synthetic fixtures are for explicit staging previews only, never automatic onboarding.
- Existing Avola data stays untouched until an explicitly approved, validated migration; no production changes in staging.
- All operational modules are available by default in the intended final product, subject to role permissions and company-specific configuration. Integrations and credentials must be configured separately, never inherited from Avola.
- Preserve original UI and workflows, migrate each module to tenant-scoped access, and verify isolation before enabling real reads or writes.
- Staging only, demo tenants only, no paid services, no deployment implied by this specification.

Acceptance: a newly created company can open the same modules and sees empty operational lists; creating data in one company never populates another. This is a target requirement, not a claim of completed implementation.

## Mandatory implementation rule — owner directive

Reuse the current Avola application UI, navigation, business logic, workflows, and systems for every tenant. Do not replace existing operational screens with newly designed screens or simplified substitutes. Change only tenant data routing, authorization, tenant-specific settings and necessary isolation infrastructure. New tenants begin with empty operational archives, ready for manual entry. Compare each migrated module against the original before acceptance. Temporary staging previews are not final product interfaces. The super-admin and company selection are new administrative entry points, not replacements for the existing operational application.

## Owner-approved FATTO workflow exception (2026-11-25)

Avola Coop retains the original irrevocable FATTO/WHAZZUP vault, including its immutable functions, files and isolation checks. Do not modify hashes, protected code, checks or intercept protected functions. Other tenants retain the original operational UI but use a separately implemented, tenant-scoped configurable FATTO workflow. Company administrators can select supported actions and their valid execution order; employees execute the configured workflow but cannot edit it. SuperAdmin controls available capabilities. Never inherit Avola recipients, credentials or operational data. This concerns the actions performed by FATTO, not permission to edit completed records.

Implementation gates: identify Avola using a trusted immutable tenant identity, not display name; fail closed for unknown tenant policy; enforce role and tenant scope server-side; validate action dependencies and reject arbitrary executable code; record configuration versions and execution outcomes; prevent duplicate effects on retry. External integrations require explicit tenant setup. Preserve existing production and vault checks. This specification is not implementation or deployment.
