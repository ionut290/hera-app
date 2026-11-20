# Android emulator smoke: diagnostic and strict recovery (20 Nov 2026)

## Confirmed failing run

The post-merge CI for commit `48fcf2821957b205ff5dcb5c837f7e5676452533` **failed** in [run 44729200766](https://github.com/ionut290/hera-app/actions/runs/44729200766). The Android APK was assembled successfully, the emulator booted, and installation/launch commands completed. The final `uiautomator dump` failed with `null root node returned by UiTestAutomationBridge` followed by `window.xml: No such file or directory`. The emulator log also reports `bad color buffer handle 135`. This is **not** proof of a clean Android UI test.

## Repair

- The workflow now selects Android Emulator `-gpu software` instead of the deprecated indirect SwiftShader graphics mode. The Android documentation lists `software` as supported and the indirect graphics mode as deprecated since Emulator 36.4.9: https://developer.android.com/studio/run/emulator-acceleration
- The smoke test runs via `scripts/android-emulator-smoke.sh`, waits for `sys.boot_completed`, installs and launches the built APK, and retries the Android UI dump up to six times. It requires the resulting hierarchy to mention `it.vargacantieri.hera`, a live app process and a nonempty screenshot.
- **Fail closed:** repeated missing UI hierarchies, the wrong app, crashed process or empty screenshot still produce failure. Diagnostics and screenshots are saved when possible, never used to manufacture a pass.
- An isolated job exercises the **actual script** with fake ADB responses for valid UI, absent hierarchy, wrong-package hierarchy and dead process. No Google/Firebase credentials or emulators are used in this unit check.

## Release meaning

A green workflow proves only the configured headless Android emulator smoke test, not real-device behavior, Google Play approval or iOS App Store readiness. The real emulator workflow runs on `push main`, after merge, so the **post-merge** execution must still pass. Android failures must be reported rather than hidden from the Apple readiness checklist.
