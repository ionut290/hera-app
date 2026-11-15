# VARGA CANTIERI — Apple Guideline 2.5.19 remediation

## Root cause
The rejected app contained a reviewer-specific Firebase sign-in that validated the apple_review claim
and redirected reviewers into an offline fictional demo instead of the actual operational application.

## Completed source changes
- Removed the two reviewer-only links from the normal login screen.
- Removed all nine client-side apple-review-* HTML/JS/CSS files.
- Removed obsolete fictional demo tests; replaced relevant CI steps with package audits.
- Native web packaging excludes reviewer-only asset names and fails when reviewer-specific markers appear.
- Standard Firebase user sign-in and operational application code remain unchanged.
- Existing server-side restrictions for old apple_review identities remain for data protection; they can
  be retired only after all such legacy accounts are identified and revoked safely.

## Tasks before submitting another build
1. Wait for all pull request checks, including the iOS simulator build, to pass.
2. Set up a real, fully functional test account through the normal login in a privacy-safe
   test environment with synthetic data. Do NOT mark it as apple_review; never expose
   confidential personnel or customer data to reviewers.
3. Verify normal functions with that account on an iPhone: sign-in, commesse, impianti,
   timekeeping, photos, and any other functionality needed for a meaningful review.
4. Merge the approved change, generate and sign a NEW build with an incremented build number,
   and confirm the built IPA has no reviewer-only client components.
5. Upload the build and supply actual test credentials in App Store Connect > App Review Information.
6. Only then reply to Apple explaining the removed reviewer-only code and confirming the tested account.

## Suggested reply to Apple after the new build and account are verified
Thank you for identifying the concern under Guideline 2.5.19. We removed the reviewer-specific
authentication flow, interface and isolated demonstration from the app in this new build.
App Review now uses the same normal sign-in and user-facing functionality as other users.
We supplied valid test credentials in App Review Information. We added a build-time check
that prevents reviewer-only code from entering subsequent native releases.

Do not send this statement until every claim is true. Source changes alone cannot guarantee approval.
