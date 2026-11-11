// Prepare opaque, branded iOS AppIcon assets from the existing Varga Cantieri logo.
// Only generated iOS native assets are touched, never production web/Android files.
import AppKit
import CoreGraphics
import Foundation

guard CommandLine.arguments.count == 3 else {
    fputs("Usage: swift prepare-ios-app-icon.swift logo.png AppIcon.appiconset\n", stderr)
    exit(2)
}
let logoURL = URL(fileURLWithPath: CommandLine.arguments[1])
let assetSet = URL(fileURLWithPath: CommandLine.arguments[2], isDirectory: true)
let catalogURL = assetSet.appendingPathComponent("Contents.json")
guard let logo = NSImage(contentsOf: logoURL),
      let sourceImage = logo.cgImage(forProposedRect: nil, context: nil, hints: nil),
      let catalogData = try? Data(contentsOf: catalogURL),
      let catalog = try? JSONSerialization.jsonObject(with: catalogData) as? [String: Any],
      let images = catalog["images"] as? [[String: Any]] else {
    fputs("Missing logo or generated Xcode AppIcon manifest\n", stderr)
    exit(1)
}

func pixelDimensions(_ item: [String: Any]) -> (Int, Int)? {
    guard let declaredSize = item["size"] as? String else { return nil }
    let dimensions = declaredSize.split(separator: "x")
    guard dimensions.count == 2,
          let w = Double(dimensions[0]), let h = Double(dimensions[1]) else { return nil }
    let scaleText = (item["scale"] as? String ?? "1x").replacingOccurrences(of: "x", with: "")
    guard let scale = Double(scaleText) else { return nil }
    let width = Int((w * scale).rounded())
    let height = Int((h * scale).rounded())
    guard width > 0, width <= 4096, height > 0, height <= 4096 else { return nil }
    return (width, height)
}

var written = 0
for image in images {
    guard let filename = image["filename"] as? String, !filename.isEmpty else { continue }
    guard URL(fileURLWithPath: filename).lastPathComponent == filename,
          let size = pixelDimensions(image) else {
        fputs("Unsupported AppIcon manifest entry: \(image)\n", stderr)
        exit(1)
    }
    let (width, height) = size
    guard let context = CGContext(
        data: nil, width: width, height: height, bitsPerComponent: 8,
        bytesPerRow: 0, space: CGColorSpaceCreateDeviceRGB(),
        bitmapInfo: CGImageAlphaInfo.noneSkipLast.rawValue) else {
        fputs("Cannot create opaque AppIcon bitmap (\(width)x\(height))\n", stderr)
        exit(1)
    }
    let rect = CGRect(x: 0, y: 0, width: width, height: height)
    context.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    context.fill(rect)
    context.interpolationQuality = .high
    context.draw(sourceImage, in: rect)

    guard let output = context.makeImage(),
          let data = NSBitmapImageRep(cgImage: output).representation(using: .png, properties: [:]) else {
        fputs("Cannot encode native icon PNG\n", stderr)
        exit(1)
    }
    try data.write(to: assetSet.appendingPathComponent(filename), options: .atomic)
    print("Prepared opaque iOS icon: \(filename) \(width)x\(height)")
    written += 1
}
guard written > 0 else {
    fputs("No named AppIcon slots in generated Xcode asset manifest\n", stderr)
    exit(1)
}
