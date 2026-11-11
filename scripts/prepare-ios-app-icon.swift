// Prepare branded, opaque iOS AppIcon assets from the existing Varga icon.
// Runs only on the ephemeral macOS CI runner; never edits production web assets.
import AppKit
import Foundation

guard CommandLine.arguments.count == 3 else {
    fputs("Usage: swift prepare-ios-app-icon.swift icons/varga-cantieri-512.png ios/App/App/Assets.xcassets/AppIcon.appiconset\n", stderr)
    exit(2)
}
let sourceURL = URL(fileURLWithPath: CommandLine.arguments[1])
let iconSet = URL(fileURLWithPath: CommandLine.arguments[2], isDirectory: true)
let contentsURL = iconSet.appendingPathComponent("Contents.json")
guard let source = NSImage(contentsOf: sourceURL),
      let data = try? Data(contentsOf: contentsURL),
      let manifest = try? JSONSerialization.jsonObject(with: data) as? [String: Any],
      let images = manifest["images"] as? [[String: Any]] else {
    fputs("Missing source image or generated AppIcon Contents.json\n", stderr)
    exit(1)
}

func pixelSize(_ item: [String: Any]) -> (Int, Int)? {
    guard let sizeString = item["size"] as? String else { return nil }
    let parts = sizeString.split(separator: "x")
    guard parts.count == 2, let width = Double(parts[0]), let height = Double(parts[1]) else { return nil }
    let scaleString = (item["scale"] as? String ?? "1x").replacingOccurrences(of: "x", with: "")
    guard let scale = Double(scaleString) else { return nil }
    let w = Int((width * scale).rounded()), h = Int((height * scale).rounded())
    return w > 0 && h > 0 && w <= 2048 && h <= 2048 ? (w, h) : nil
}

var count = 0
for item in images {
    guard let filename = item["filename"] as? String, !filename.isEmpty else { continue }
    guard filename == URL(fileURLWithPath: filename).lastPathComponent,
          let (width, height) = pixelSize(item),
          let bitmap = NSBitmapImageRep(
            bitmapDataPlanes: nil, pixelsWide: width, pixelsHigh: height,
            bitsPerSample: 8, samplesPerPixel: 3, hasAlpha: false, isPlanar: false,
            colorSpaceName: .deviceRGB, bytesPerRow: 0, bitsPerPixel: 0),
          let context = NSGraphicsContext(bitmapImageRep: bitmap) else {
        fputs("Invalid AppIcon asset specification\n", stderr)
        exit(1)
    }
    NSGraphicsContext.saveGraphicsState()
    NSGraphicsContext.current = context
    context.imageInterpolation = .high
    NSColor.white.setFill()
    NSRect(x: 0, y: 0, width: width, height: height).fill()
    let bounds = NSRect(x: 0, y: 0, width: width, height: height)
    source.draw(in: bounds, from: NSRect(origin: .zero, size: source.size),
                operation: .sourceOver, fraction: 1.0)
    context.flushGraphics()
    NSGraphicsContext.restoreGraphicsState()
    guard let pngData = bitmap.representation(using: .png, properties: [:]) else {
        fputs("Cannot encode AppIcon PNG\n", stderr)
        exit(1)
    }
    try pngData.write(to: iconSet.appendingPathComponent(filename), options: .atomic)
    count += 1
}
guard count > 0 else {
    fputs("Generated AppIcon asset catalog has no named image slots\n", stderr)
    exit(1)
}
print("Prepared \(count) branded, opaque iOS AppIcon assets (no alpha).")
