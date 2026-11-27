# FATTO execution boundary — staging only

No Firebase billing activation, Cloud Functions deployment, service account keys in clients, or changes to the original vault.

## Implementation decision
Use Firestore transactions and restrictive rules for internal completion records. A client can request a completion, but cannot assert that an external effect succeeded. Until a trusted executor exists, the journal is read-only. A trusted executor on existing infrastructure is an option, not an authorized deployment; it must verify Firebase ID tokens, paired active memberships, immutable tenant eligibility, approved policy version, site identity and operation ownership. Admin SDK bypasses rules, so all checks must be enforced in that executor. Do not reuse Avola credentials.

## Activation gates
- Recover and compare complete deployed staging rules; test the merged rules.
- Provision explicit configurable demo tenant capability through trusted administration; unknown modes fail closed.
- Establish tenant-specific site schema and atomic completion transaction.
- Bind operation identity to tenant, site and completion cycle; reject concurrent duplicates.
- Persist intent before external effect; reconcile uncertain results without automatic retry.
- WhatsApp opening is not evidence of message delivery. Sheets success requires provider confirmation.
- Verify original FATTO vault and isolation checks without modifying them.
- Browser test original UI against synthetic tenant data before any activation.

The current preview, policy, planner and journal are not an operational replacement. No external effects or live journal writes are enabled.
