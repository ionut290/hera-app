#!/usr/bin/env python3
"""Create only LOCAL Xcode export options for the separate V2 internal candidate."""
import os
import plistlib
import sys
from pathlib import Path
from validate_ios_profile import BUNDLE, TEAM, PROFILE

def options():
    return {
        "method": "app-store-connect",
        "destination": "export",
        "signingStyle": "manual",
        "signingCertificate": "Apple Distribution",
        "provisioningProfiles": {BUNDLE: PROFILE},
        "teamID": TEAM,
        "testFlightInternalTestingOnly": True,
        "manageAppVersionAndBuildNumber": False,
        "uploadSymbols": False,
    }

if __name__ == "__main__":
    root = Path(os.environ["RUNNER_TEMP"]).resolve()
    destination = Path(sys.argv[1]).resolve()
    if destination.parent != root or destination.name != "VargaV2ExportOptions.plist":
        sys.exit("Only the expected private temporary export path is allowed.")
    with destination.open("xb") as file:
        plistlib.dump(options(), file)
    print("V2 local export options validated. No upload destination.")
