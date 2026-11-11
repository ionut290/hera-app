// Checks that launching the simulator app visibly changed the iPhone screen.
// Not a UI regression test: it only rejects a simulator that stays on Home.
// Uses macOS AppKit shipped with the GitHub-hosted Xcode runner.
import AppKit
import Foundation

guard CommandLine.arguments.count == 3 else {
    fputs("Usage: swift check-ios-simulator-foreground.swift before.png after.png\n", stderr)
    exit(2)
}

func image(_ path: String) -> NSBitmapImageRep? {
    guard let data = try? Data(contentsOf: URL(fileURLWithPath: path)) else { return nil }
    return NSBitmapImageRep(data: data)
}

guard let before = image(CommandLine.arguments[1]),
      let after = image(CommandLine.arguments[2]),
      before.pixelsWide == after.pixelsWide,
      before.pixelsHigh == after.pixelsHigh else {
    fputs("Invalid or mismatched screenshots\n", stderr)
    exit(2)
}

let width = before.pixelsWide
let height = before.pixelsHigh
var samples = 0
var changed = 0
var nearlyWhite = 0

// Exclude status bar and dock; count substantial color changes across the page.
for y in stride(from: height / 6, to: height * 5 / 6, by: 16) {
    for x in stride(from: width / 10, to: width * 9 / 10, by: 16) {
        guard let a = before.colorAt(x: x, y: y)?.usingColorSpace(.deviceRGB),
              let b = after.colorAt(x: x, y: y)?.usingColorSpace(.deviceRGB) else { continue }
        samples += 1
        if b.redComponent > 0.975 && b.greenComponent > 0.975 && b.blueComponent > 0.975 {
            nearlyWhite += 1
        }
        let delta = abs(a.redComponent - b.redComponent)
            + abs(a.greenComponent - b.greenComponent)
            + abs(a.blueComponent - b.blueComponent)
        if delta > 0.18 { changed += 1 }
    }
}

let ratio = samples == 0 ? 0.0 : Double(changed) / Double(samples)
print(String(format: "Screen changed samples: %.2f%% (%d/%d)", ratio * 100, changed, samples))
if ratio < 0.03 {
    fputs("FAIL: App launch did not visibly leave the iPhone Home screen. Inspect crash diagnostics.\n", stderr)
    exit(1)
}
let whiteRatio = samples == 0 ? 0.0 : Double(nearlyWhite) / Double(samples)
print(String(format: "Nearly white central screen: %.2f%% (%d/%d)", whiteRatio * 100, nearlyWhite, samples))
if whiteRatio > 0.998 {
    fputs("FAIL: iPhone app opened but displays a blank white screen. Inspect WebView startup diagnostics.\n", stderr)
    exit(1)
}
print("PASS: Simulator left Home and screen is not uniformly white; inspect UI manually.")
