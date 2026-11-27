# Sfalci COBO integration contract
Source: original cobo-mowing-work-orders.js, preserved unchanged.
- Original commessa ID: sfalcio-cobo.
- Original plant ID: cobo-parco-<identifier>.
- Resolve all database references under tenants/{tenantId}/commesse/sfalcio-cobo/impianti.
- Original saveWorkOrder reads commessa and plant, initializes missing commessa only with canManageData(), and uses batch.set merge for changes.
- Preserve original payload, modal and navigation. Never use production globals or root collections in staging.
- Before enabling writes, enforce role, paired membership, active demo tenant, allowed document fields, creation versus update, and prevent completion state changes through this module.
- Preserve existing completion fields on merge; FATTO is a separate protected workflow.
- New tenant archives remain empty; do not import Avola records or initialize commesse without an authorized action.
- Original module is not yet connected or enabled. Read-only adapter is insufficient for saveWorkOrder.
