# CocoaPods GTMSessionFetcher conflict discovered by first Google-native TestFlight run

On 20 Nov 2026, internal TestFlight workflow #13 (run 44773557475,
commit `f91e332`) FAILED during `npx cap sync ios`.
Apple archive/upload did NOT run; **build 1.0 (20013) was not delivered**.

Root cause from runner error: `cap add ios` initially generated
`Podfile.lock` pinning `GTMSessionFetcher/Core 4.5.0`. After Google pod
was added, `GoogleSignIn 7.1.0` required `GTMSessionFetcher/Core ~> 3.3`;
FirebaseAuth 11.7 permits `>=3.4,<5`, thus a 3.4+ release is compatible.
Normal `pod install` does not change the pinned 4.5.0 and fails.

Fix: in each signed iOS workflow, after adding the Google subspec and
**before** `npx cap sync ios`, run exactly
`(cd ios/App && pod update GTMSessionFetcher/Core)`.
This selectively updates the generated CocoaPods lock, not the repository,
not Firebase user data and not all unrelated libraries.

The new macOS CI `check-ios-google-cocoapods.yml` replicates the native
Podfile generation and dependency resolution without signing credentials or
deploying. It must finish green before another TestFlight upload is attempted.
If a new private workflow is dispatched after merging, it gets a **new**
build number (the cancelled #13 build number is not reused).

The Google login on a real iPhone still needs the separate physical
test tracked in issue #1656. No App Store submission is authorized.
