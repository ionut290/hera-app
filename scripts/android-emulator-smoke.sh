#!/usr/bin/env bash
# Strict Android emulator smoke test, headless CI. No Firebase writes.
# Graphics/UI automation can lag behind Android's boot_completed property;
# retry the UI hierarchy dump, but never accept a missing or wrong screen.
set -euo pipefail

PACKAGE="it.vargacantieri.hera"
APK="android/app/build/outputs/apk/debug/app-debug.apk"
test -s "$APK" || { echo "::error::Missing compiled Android debug APK"; exit 1; }
adb wait-for-device
for attempt in $(seq 1 20); do
  if [[ "$(adb shell getprop sys.boot_completed | tr -d '\r')" == "1" ]]; then
    break
  fi
  if [[ "$attempt" == "20" ]]; then
    echo "::error::Emulator never finished booting."
    exit 1
  fi
  sleep 3
done

adb install -r "$APK"
adb shell am force-stop "$PACKAGE"
adb shell monkey -p "$PACKAGE" -c android.intent.category.LAUNCHER 1
sleep 12

# Collect evidence even when the UI automation bridge is slow or fails.
# The UI hierarchy is required for success: screenshot/process alone do NOT pass.
screen_ready=0
for attempt in $(seq 1 6); do
  echo "Inspecting Android UI hierarchy (attempt $attempt/6)"
  rm -f window.xml
  if adb shell uiautomator dump /sdcard/window.xml &&
     adb pull /sdcard/window.xml window.xml &&
     test -s window.xml &&
     grep -Fq "$PACKAGE" window.xml; then
    screen_ready=1
    break
  fi
  sleep 5
done

# Preserve useful diagnostics for unsuccessful device runs.
adb exec-out screencap -p > android-emulator.png || true
if [[ "$screen_ready" != "1" ]]; then
  echo "::error::Emulator UI hierarchy unavailable or does not show VARGA CANTIERI after retries."
  adb shell dumpsys window windows | tail -n 70 || true
  adb shell dumpsys activity activities | tail -n 70 || true
  exit 1
fi

# Verify the application is alive, not just an XML view from before a crash.
if ! adb shell pidof "$PACKAGE" >/dev/null; then
  echo "::error::VARGA CANTIERI process exited during the smoke test."
  exit 1
fi
if grep -Eqi "has stopped|keeps stopping|isn't responding" window.xml; then
  echo "::error::Android shows an application failure dialog."
  exit 1
fi
test -s android-emulator.png || {
  echo "::error::Unable to capture a nonempty Android screenshot."
  exit 1
}
echo "PASS: installed, launched and inspected VARGA CANTIERI on the Android emulator."
