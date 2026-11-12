#!/usr/bin/env python3
"""Apply the safe-area marker to the GENERATED iOS WebView HTML only.

Run after `npx cap sync ios`. This never changes index.html, www/, Android,
the deployed website, authentication, or application business logic.
"""
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
IOS_PUBLIC = ROOT / "ios" / "App" / "App" / "public"
HTML_FILE = IOS_PUBLIC / "index.html"
CSS_FILE = IOS_PUBLIC / "style.css"
ORIGINAL_TAG = '<html lang="it">'
NATIVE_TAG = '<html lang="it" class="varga-ios-native-safe-area">'
EXPECTED_CSS_SELECTOR = "html.varga-ios-native-safe-area main.container"


def main() -> None:
    if not HTML_FILE.is_file() or not CSS_FILE.is_file():
        raise SystemExit("Generated iOS assets missing: run npx cap sync ios first.")

    css = CSS_FILE.read_text(encoding="utf-8")
    if EXPECTED_CSS_SELECTOR not in css:
        raise SystemExit("iOS native safe-area CSS fallback is missing.")

    html = HTML_FILE.read_text(encoding="utf-8")
    if NATIVE_TAG in html:
        print("Generated iOS page already has the native safe-area marker.")
        return
    if html.count(ORIGINAL_TAG) != 1:
        raise SystemExit("Unexpected iOS HTML root tag: refusing unsafe modification.")

    patched = html.replace(ORIGINAL_TAG, NATIVE_TAG, 1)
    HTML_FILE.write_text(patched, encoding="utf-8")

    if NATIVE_TAG not in HTML_FILE.read_text(encoding="utf-8"):
        raise SystemExit("Generated iOS safe-area marker could not be verified.")
    print("iOS-only HTML safe-area marker applied; web and Android unchanged.")


if __name__ == "__main__":
    main()
