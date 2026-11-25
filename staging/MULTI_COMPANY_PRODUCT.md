# VARGA CANTIERI 2.0 — shared application

Product requirement: every new company receives the same complete application, navigation, visual design and operational capabilities as the existing Avola Coop application, with an initially empty company-owned archive. No reduced edition or separate fork per company.

- One shared codebase; tenant-specific data and settings.
- New companies start with zero operational records. Synthetic fixtures are for explicit staging previews only, never automatic onboarding.
- Existing Avola data stays untouched until an explicitly approved, validated migration; no production changes in staging.
- All operational modules are available by default in the intended final product, subject to role permissions and company-specific configuration. Integrations and credentials must be configured separately, never inherited from Avola.
- Preserve original UI and workflows, migrate each module to tenant-scoped access, and verify isolation before enabling real reads or writes.
- Staging only, demo tenants only, no paid services, no deployment implied by this specification.

Acceptance: a newly created company can open the same modules and sees empty operational lists; creating data in one company never populates another. This is a target requirement, not a claim of completed implementation.
