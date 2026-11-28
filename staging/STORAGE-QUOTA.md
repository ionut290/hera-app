# Private file storage candidate (V2 only)

The trial quota is 100,000,000 bytes; Starter 1,000,000,000; Professional 5,000,000,000; Business 20,000,000,000. Individual PDF/JPEG/PNG/WebP files are limited to 15 MiB. Signature checks are basic format checks, not malware scanning.

A verified Firebase ID token, active tenant membership and mirrored membership are required. The server verifies the tenant entitlement and reserves bytes in a Firestore transaction before creating the object at a tenant-scoped immutable path. Concurrent uploads cannot oversubscribe the recorded quota. Repeating an already completed file ID with the identical content returns the recorded result without charging twice. A failed or ambiguous upload remains charged and requires reconciliation; never release capacity based solely on a client error.

Direct client access is denied by the candidate Storage rules. Download goes through the authenticated service, with hash and generation verification. Only the uploader and tenant administrators can read a file in this initial model; other sharing permissions must be designed explicitly. Expired subscriptions permit existing authorized downloads but deny new uploads. Suspended members cannot download. No public links are issued.

## Not ready to deploy

- No deployed callable/HTTP endpoint, original photo/document UI integration, file list, or sharing UI.
- No deletion, reservation cleanup, migration of legacy data, malware scanning, rate limiting, or verified retention process. Reconcile reservations and inventory before changing any counter.
- Verify a dedicated private V2 bucket, IAM, object generation preconditions, no unintended public ACL/token links, versioning/soft-delete costs and service identity. Do not reuse the original app bucket or its production Storage rules.
- Gate real deployment behind project-specific configuration, verified tests, cost review and explicit release procedure. The current workflow uses local Firebase emulators only.
