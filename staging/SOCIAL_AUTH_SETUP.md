# Google and Apple authentication — Varga Cantieri 2.0 staging

The launcher has three social entry points (company creation, company-code join, existing login). The provider credential is exchanged via Firebase JS Auth, preserving the existing verified-user and membership checks. Apple requires the raw nonce; no credentials are stored in localStorage. Only the user's onboarding intent is stored by UID.

## Release gate

The social controls are hidden unless the private staging build has VARGA_V2_SOCIAL_AUTH_READY=true and the iOS FirebaseAuthentication native plugin is available. The current native build does not yet include that plugin. Do not enable the flag or publish a social-enabled build before the native configuration is verified.

1. In the independent staging Firebase project, enable Google and Apple Authentication providers. Check one-account-per-email behavior and approved support email.
2. Register/verify the exact V2 iOS bundle ID in Firebase and download its own GoogleService-Info.plist. Never reuse the production Varga Cantieri file. Confirm the iOS OAuth client ID and reversed URL scheme.
3. Install a Capacitor 7-compatible pinned version of @capacitor-firebase/authentication in the staging native build; configure skipNativeAuth=true and providers google.com and apple.com. Add the Google CocoaPods subspec, URL scheme and plist to the Xcode target, using the provider's version-specific setup instructions.
4. Enable Sign in with Apple on the V2 App ID in Apple Developer, add the capability/entitlement to the native target and regenerate the matching distribution provisioning profile. Configure any Apple provider settings required by Firebase, including private email relay if relevant.
5. On a private signed build only, set VARGA_V2_SOCIAL_AUTH_READY=true at web bundle build time after verifying the native integration. No production config changes.
6. Verify real-device Google and Apple signup for each flow, existing account collision handling, cancellation, private relay email, verification state, duplicate company prevention, pending join requests and tenant isolation. Verify logout and subsequent sign-in. Test email/password regression. Do not silently link identities.

References:
- https://github.com/capawesome-team/capacitor-firebase/blob/main/packages/authentication/docs/firebase-js-sdk.md
- https://github.com/capawesome-team/capacitor-firebase/blob/main/packages/authentication/docs/setup-google.md
- https://github.com/capawesome-team/capacitor-firebase/blob/main/packages/authentication/docs/setup-apple.md

No actual Firebase provider settings, Apple capabilities or provisioning profiles were changed in this branch.
