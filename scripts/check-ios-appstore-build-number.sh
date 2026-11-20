#!/usr/bin/env bash
# Fail-closed numeric build number validation. Apple confirmed TestFlight
# VARGA CANTIERI 1.0 (20011) on 19 Nov 2026. This is a HISTORICAL floor, NOT
# a live App Store Connect lookup; operator must still check the latest build.
set -euo pipefail
number="${1:-}"
last_confirmed=20011
if [[ ! "$number" =~ ^[1-9][0-9]{0,5}$ ]] || (( 10#$number <= last_confirmed )); then
  echo "::error::Use a new iOS build number above 20011 and check App Store Connect for any newer upload." >&2
  exit 1
fi
echo "iOS build number verified against last confirmed TestFlight 20011. Recheck App Store Connect before upload."
