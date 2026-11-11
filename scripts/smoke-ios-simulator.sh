#!/usr/bin/env bash
# Varga Cantieri: run only the unsigned iOS Simulator app on an ephemeral CI simulator.
# Never configures signing, production access, app distribution or real credentials.
set -euo pipefail

if [[ "$#" -ne 2 ]]; then
  echo "Uso: bash scripts/smoke-ios-simulator.sh /percorso/App.app /percorso/screenshot.png" >&2
  exit 2
fi

APP_PATH="$1"
SCREENSHOT_PATH="$2"
BUNDLE_ID="it.vargacantieri.hera"

if [[ ! -d "$APP_PATH" ]] || [[ ! -s "$APP_PATH/Info.plist" ]]; then
  echo "App iOS per il simulatore mancante: $APP_PATH" >&2
  exit 1
fi

ACTUAL_BUNDLE_ID=$(/usr/libexec/PlistBuddy -c "Print :CFBundleIdentifier" "$APP_PATH/Info.plist")
if [[ "$ACTUAL_BUNDLE_ID" != "$BUNDLE_ID" ]]; then
  echo "Bundle ID inatteso: $ACTUAL_BUNDLE_ID" >&2
  exit 1
fi

# An existing available iPhone simulator is supplied by the macOS hosted runner.
SIMULATOR_UDID="$(xcrun simctl list devices available -j | python3 -c '
import json, sys
data = json.load(sys.stdin)
for devices in data["devices"].values():
    for device in devices:
        if device.get("name", "").startswith("iPhone") and device.get("isAvailable", True):
            print(device["udid"])
            sys.exit(0)
sys.exit("Nessun simulatore iPhone disponibile")
')"

if [[ -z "$SIMULATOR_UDID" ]]; then
  echo "Nessun simulatore iPhone trovato" >&2
  exit 1
fi

echo "Simulatore selezionato: $SIMULATOR_UDID"
if ! xcrun simctl list devices booted | grep -Fq "$SIMULATOR_UDID"; then
  xcrun simctl boot "$SIMULATOR_UDID"
fi
xcrun simctl bootstatus "$SIMULATOR_UDID" -b
xcrun simctl install "$SIMULATOR_UDID" "$APP_PATH"
mkdir -p "$(dirname "$SCREENSHOT_PATH")"
BEFORE_SCREENSHOT="${SCREENSHOT_PATH%.png}-before.png"
xcrun simctl io "$SIMULATOR_UDID" screenshot "$BEFORE_SCREENSHOT"

LAUNCH_OUTPUT="$(xcrun simctl launch "$SIMULATOR_UDID" "$BUNDLE_ID")"
echo "Avvio: $LAUNCH_OUTPUT"
if [[ "$LAUNCH_OUTPUT" != "$BUNDLE_ID:"* ]]; then
  echo "L'app non è stata avviata correttamente dal simulatore" >&2
  exit 1
fi

# Do not enter demo/production accounts. Capture a non-authenticated launch only.
sleep 8
xcrun simctl io "$SIMULATOR_UDID" screenshot "$SCREENSHOT_PATH"
test -s "$SCREENSHOT_PATH"
sips -g pixelWidth -g pixelHeight "$SCREENSHOT_PATH"

if ! swift scripts/check-ios-simulator-foreground.swift "$BEFORE_SCREENSHOT" "$SCREENSHOT_PATH"; then
  # Allow slow WebView/Firebase initialization without treating an 8-second blank as final.
  cp "$SCREENSHOT_PATH" "${SCREENSHOT_PATH%.png}-after8s.png"
  echo "Avvio iOS non ancora visibile: nuova verifica dopo ulteriori 15 secondi."
  sleep 15
  xcrun simctl io "$SIMULATOR_UDID" screenshot "$SCREENSHOT_PATH"
  if swift scripts/check-ios-simulator-foreground.swift "$BEFORE_SCREENSHOT" "$SCREENSHOT_PATH"; then
    echo "Schermata iniziale caricata alla seconda verifica (23 secondi)."
    exit 0
  fi
  DIAGNOSTIC_DIR="$(dirname "$SCREENSHOT_PATH")/varga-ios-simulator-diagnostics"
  mkdir -p "$DIAGNOSTIC_DIR"
  xcrun simctl spawn "$SIMULATOR_UDID" log show --last 1m --style compact \
    --predicate 'process == "App" OR eventMessage CONTAINS[c] "vargacantieri"' \
    2>&1 | tail -n 250 > "$DIAGNOSTIC_DIR/app-launch-log.txt" || true
  echo "App iOS non risulta in primo piano: screenshot e diagnostica salvati." >&2
  exit 1
fi

echo "App installata, lanciata e schermo modificato; screenshot per revisione: $SCREENSHOT_PATH"
echo "ATTENZIONE: il test NON prova login, WebView interattiva o funzioni operative."
