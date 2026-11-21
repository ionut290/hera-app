#!/usr/bin/env bash
# Fail-closed numeric build number validation. Apple confirmed TestFlight
# VARGA CANTIERI 1.0 (20016) on 20 Nov 2026. This is a HISTORICAL floor, NOT
# a live App Store Connect lookup; operator must still check the latest build.
set -euo pipefail
number="${1:-}"
last_confirmed=20016
if [[ ! "$number" =~ ^[1-9][0-9]{0,5}$ ]] || (( 10#$number <= last_confirmed )); then
  echo "::error::Use a new iOS build number above 20016 and check App Store Connect for any newer upload." >&2
  exit 1
fi
echo "iOS build number verified against last confirmed TestFlight 20016. Recheck App Store Connect before upload."
