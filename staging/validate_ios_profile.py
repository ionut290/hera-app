#!/usr/bin/env python3
"""Validate a decoded provisioning profile for the independent V2 app only."""
import datetime
import os
import plistlib
import re
import sys
from pathlib import Path

BUNDLE = "it.vargacantieri.staging"
TEAM = "NJ64236Q54"
PROFILE = "Varga Cantieri 2 Staging TestFlight"

def validate(profile, leaf, now=None):
    now = now or datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None)
    ent = profile.get("Entitlements", {})
    if profile.get("Name") != PROFILE or TEAM not in profile.get("TeamIdentifier", []):
        raise ValueError("Wrong profile name or Apple team")
    if ent.get("application-identifier") != TEAM + "." + BUNDLE:
        raise ValueError("Wrong bundle: only V2 staging accepted")
    if ent.get("com.apple.developer.team-identifier", TEAM) != TEAM:
        raise ValueError("Wrong entitlement team")
    if ent.get("get-task-allow") is not False or "ProvisionedDevices" in profile or profile.get("ProvisionsAllDevices"):
        raise ValueError("Not an App Store distribution profile")
    expires = profile.get("ExpirationDate")
    if not isinstance(expires, datetime.datetime) or expires <= now:
        raise ValueError("Profile expired")
    if leaf not in profile.get("DeveloperCertificates", []) or not leaf:
        raise ValueError("Distribution certificate mismatch")
    uuid = profile.get("UUID", "")
    if not re.fullmatch(r"[0-9a-fA-F-]{36}", uuid):
        raise ValueError("Invalid profile UUID")
    return uuid

if __name__ == "__main__":
    try:
        uuid = validate(plistlib.loads(Path(sys.argv[1]).read_bytes()), Path(sys.argv[2]).read_bytes())
        with open(os.environ["GITHUB_ENV"], "a", encoding="utf-8") as f:
            f.write("V2_PROFILE_UUID=" + uuid + "\n")
        print("Verified exact V2 profile, team, certificate and distribution eligibility.")
    except (ValueError, KeyError, OSError, TypeError, IndexError):
        sys.exit("V2 provisioning validation failed. No signing permitted.")
