#!/usr/bin/env python3
"""Validate an existing Firebase Apple registration and configure URL scheme.

No Firebase project is created. No login code or existing user data is changed.
The GoogleService-Info.plist must come from the *registered iOS app* with the
same bundle ID; never synthesize it from the web configuration.
"""
from pathlib import Path
import plistlib
import sys

ROOT = Path(__file__).resolve().parent.parent
SOURCE = ROOT / "ios" / "App" / "App" / "GoogleService-Info.plist"
INFO = ROOT / "ios" / "App" / "App" / "Info.plist"
EXPECTED_BUNDLE_ID = "it.vargacantieri.hera"
EXPECTED_PROJECT_ID = "hera-app-6cd2b"


def main():
    if not SOURCE.is_file():
        raise SystemExit("Missing registered iOS GoogleService-Info.plist")

    with SOURCE.open("rb") as file:
        firebase = plistlib.load(file)

    if not isinstance(firebase, dict):
        raise SystemExit("Invalid Firebase iOS configuration")
    if firebase.get("BUNDLE_ID") != EXPECTED_BUNDLE_ID:
        raise SystemExit("Firebase iOS bundle ID mismatch")
    if firebase.get("PROJECT_ID") != EXPECTED_PROJECT_ID:
        raise SystemExit("Firebase project ID mismatch")
    if ":ios:" not in str(firebase.get("GOOGLE_APP_ID", "")):
        raise SystemExit("Firebase Apple app ID missing or invalid")

    scheme = firebase.get("REVERSED_CLIENT_ID")
    if not isinstance(scheme, str) or not scheme.startswith("com.googleusercontent.apps."):
        raise SystemExit("Reversed Google OAuth client ID missing")
    if not firebase.get("CLIENT_ID"):
        raise SystemExit("iOS Google OAuth client ID missing")

    with INFO.open("rb") as file:
        info = plistlib.load(file)
    if not isinstance(info, dict):
        raise SystemExit("Invalid Info.plist")
    url_types = list(info.get("CFBundleURLTypes", []))
    if not any(scheme in item.get("CFBundleURLSchemes", []) for item in url_types):
        url_types.append({"CFBundleURLSchemes": [scheme]})
        info["CFBundleURLTypes"] = url_types
        with INFO.open("wb") as file:
            plistlib.dump(info, file, sort_keys=False)

    print("Firebase iOS bundle/project identity and Google URL scheme verified.")


if __name__ == "__main__":
    main()
