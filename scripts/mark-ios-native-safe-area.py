#!/usr/bin/env python3
"""Apply a reliable status-bar clearance ONLY to the generated iOS WKWebView.

Run after `npx cap sync ios`. Neither source files, web/PWA, Android,
authentication nor operational application logic are modified.
"""

from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
IOS_PUBLIC = ROOT / "ios" / "App" / "App" / "public"
HTML_FILE = IOS_PUBLIC / "index.html"
CSS_FILE = IOS_PUBLIC / "style.css"
ORIGINAL_TAG = '<html lang="it">'
NATIVE_TAG = '<html lang="it" class="varga-ios-native-safe-area">'
EXPECTED_CSS_SELECTOR = "html.varga-ios-native-safe-area main.container"
STYLE_ID = "varga-ios-native-header-safe-area"

# An inline iOS-only rule is more reliable than a class-dependent rule in a
# separately cached stylesheet. Only the native generated HTML receives it.
# 82 CSS px clears the iPhone Dynamic Island/status bar in portrait even when
# WKWebView reports env(safe-area-inset-top) as 0; larger insets still win.
NATIVE_STYLE = """<style id="varga-ios-native-header-safe-area">
@media (max-width: 560px) and (orientation: portrait) {
  html.varga-ios-native-safe-area body > main.container {
    padding-top: max(82px, calc(14px + env(safe-area-inset-top, 0px))) !important;
  }
}
</style>"""


def main() -> None:
    if not HTML_FILE.is_file() or not CSS_FILE.is_file():
        raise SystemExit("Generated iOS assets missing: run npx cap sync ios first.")

    css = CSS_FILE.read_text(encoding="utf-8")
    if EXPECTED_CSS_SELECTOR not in css:
        raise SystemExit("The existing iOS safe-area CSS selector is missing.")

    html = HTML_FILE.read_text(encoding="utf-8")
    if html.count('<main class="container">') != 1:
        raise SystemExit("Unexpected iOS app root container: refusing unsafe modification.")
    if 'viewport-fit=cover' not in html:
        raise SystemExit("Missing iOS viewport-fit=cover meta tag.")
    if html.count("</head>") != 1:
        raise SystemExit("Unexpected iOS HTML head: refusing unsafe modification.")

    if NATIVE_TAG not in html:
        if html.count(ORIGINAL_TAG) != 1:
            raise SystemExit("Unexpected iOS HTML root tag: refusing unsafe modification.")
        html = html.replace(ORIGINAL_TAG, NATIVE_TAG, 1)
    elif html.count(NATIVE_TAG) != 1:
        raise SystemExit("Duplicate iOS native safe-area marker.")

    if STYLE_ID in html:
        if html.count(NATIVE_STYLE) != 1:
            raise SystemExit("Conflicting or duplicate iOS safe-area style.")
    else:
        html = html.replace("</head>", NATIVE_STYLE + "\n</head>", 1)

    HTML_FILE.write_text(html, encoding="utf-8")
    verify = HTML_FILE.read_text(encoding="utf-8")
    if verify.count(NATIVE_TAG) != 1 or verify.count(NATIVE_STYLE) != 1:
        raise SystemExit("iOS header safe-area patch verification failed.")
    print("iOS-only header clearance installed and verified (82px portrait floor).")


if __name__ == "__main__":
    main()
