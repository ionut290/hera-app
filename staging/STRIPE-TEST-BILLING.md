# Stripe test billing — Varga Cantieri 2.0

This candidate runs only with the localhost Firestore emulator and a synthetic Stripe signing secret. It uses the official Stripe SDK to verify raw signed snapshot events, with the pinned 2024-10-28.acacia payload contract. It never makes a Stripe network request, creates a Checkout session, registers an endpoint, charges a card, or sends a real invoice.

## Commercial contract

Existing catalog: Starter EUR 19.90, Professional EUR 39.90, Business EUR 79.90 for 30 days. The fixture accepts one fixed-price invoice line per subscription, one quantity, EUR, 30-day recurrence, no proration, discounts, extra lines, or out-of-band payment. A trusted server-side binding associates the exact subscription/customer with an approved tenant and tier. Only paid invoices advance the entitlement; event, invoice and billing-period IDs are recorded atomically and globally deduplicated. A failed payment does not extend access. Cancellation stops new grants and retains the already paid period. Out-of-order or incompatible invoices require reconciliation rather than silently granting time. Manual payment confirmation cannot override a Stripe-bound test tenant. Test receipts are marked as simulations, not tax invoices.

## Before actual sandbox and production

- Confirm the business sells directly and exclusively to organizations for employees under Apple App Review 3.1.3(c), or implement the applicable in-app purchase / regional rules. Do not assume the enterprise exception applies to individual self-service sales. No Stripe purchasing button or external payment link has been added to iOS.
- Configure a separate Stripe sandbox with test-only credentials in secret storage; never place credentials in the repository or client. Add authenticated organization-owner checkout, server-created customer/subscription bindings, provider price ID validation, and explicit terms/cancellation UX. Prevent concurrent checkout and manual payment overlap.
- Pin and verify the actual Stripe API version and retrieve the authoritative invoice/payment intent/subscription before granting access; signed local fixtures are not a real sandbox end-to-end payment. Handle asynchronous payment completion, retry/reconciliation, invoice gaps, refunds, disputes, credit notes, tax, proration and plan changes. Avoid automatic destructive data removal on expiry.
- Confirm VAT-inclusive/exclusive pricing, billing details and Italian electronic invoicing/SDI obligations with a qualified tax professional. Stripe receipts and test invoice records are not a substitute for statutory invoices. No invoice emails or fiscal documents are generated.
- Add an authenticated server-only webhook endpoint, verified account separation, durable event queue, monitoring, rate limits and secrets management; test actual Stripe sandbox before any real charge or deployment. Keep original V1 and production Firebase untouched.

References: https://developer.apple.com/app-store/review/guidelines/ (3.1.1, 3.1.3(c)); https://docs.stripe.com/billing/subscriptions/webhooks ; https://docs.stripe.com/webhooks .
