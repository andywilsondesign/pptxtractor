// pdfrender - render PDF pages to PNG or JPEG at an exact long-edge size.
//
// Uses CoreGraphics directly: each page is drawn into a bitmap at the target
// scale, so text and vector art are rasterised AT that size rather than being
// enlarged into it. Annotations (speaker notes) are not page content and are
// never drawn.
import Foundation
import CoreGraphics
import ImageIO
import UniformTypeIdentifiers

func fail(_ msg: String) -> Never {
    FileHandle.standardError.write((msg + "\n").data(using: .utf8)!)
    exit(2)
}

let args = CommandLine.arguments
guard args.count >= 5, let target = Double(args[3]) else {
    fail("usage: pdfrender <in.pdf> <outDir> <longEdgePx> <prefix> [format] [page]\n"
       + "       format: png (default) | jpeg | jpg")
}
let src = URL(fileURLWithPath: args[1])
let outDir = URL(fileURLWithPath: args[2])
let prefix = args[4]
let format = args.count >= 6 ? args[5].lowercased() : "png"
let onlyPage = args.count >= 7 ? Int(args[6]) : nil

let isJPEG = (format == "jpeg" || format == "jpg")
guard isJPEG || format == "png" else { fail("unknown format: \(format)") }
let utType = isJPEG ? UTType.jpeg.identifier : UTType.png.identifier
let ext = isJPEG ? "jpg" : "png"

guard let doc = CGPDFDocument(src as CFURL) else { fail("cannot open pdf: \(args[1])") }
try? FileManager.default.createDirectory(at: outDir, withIntermediateDirectories: true)

let n = doc.numberOfPages
let pad = max(2, String(n).count)
var written = 0

for i in 1...max(n, 1) {
    if let only = onlyPage, i != only { continue }
    guard i <= n, let page = doc.page(at: i) else { continue }
    let box = page.getBoxRect(.cropBox)
    guard box.width > 0, box.height > 0 else { continue }
    let scale = target / max(box.width, box.height)
    let w = Int((box.width * scale).rounded()), h = Int((box.height * scale).rounded())
    guard w > 0, h > 0 else { continue }

    // JPEG has no alpha; PNG keeps one. Both get an opaque white ground so a
    // slide with a transparent background does not come out black.
    let alpha = isJPEG ? CGImageAlphaInfo.noneSkipFirst : CGImageAlphaInfo.premultipliedFirst
    guard let ctx = CGContext(data: nil, width: w, height: h, bitsPerComponent: 8, bytesPerRow: 0,
                              space: CGColorSpace(name: CGColorSpace.sRGB)!,
                              bitmapInfo: alpha.rawValue) else { continue }
    ctx.interpolationQuality = .high
    ctx.setAllowsAntialiasing(true); ctx.setShouldAntialias(true)
    ctx.setShouldSmoothFonts(true); ctx.setAllowsFontSmoothing(true)
    ctx.setFillColor(CGColor(red: 1, green: 1, blue: 1, alpha: 1))
    ctx.fill(CGRect(x: 0, y: 0, width: w, height: h))
    ctx.scaleBy(x: scale, y: scale)
    ctx.translateBy(x: -box.origin.x, y: -box.origin.y)
    ctx.drawPDFPage(page)

    guard let img = ctx.makeImage() else { continue }
    let name = onlyPage != nil
        ? "\(prefix).\(ext)"
        : String(format: "\(prefix)-%0\(pad)d.\(ext)", i)
    let url = outDir.appendingPathComponent(name)
    guard let dest = CGImageDestinationCreateWithURL(url as CFURL, utType as CFString, 1, nil) else { continue }
    var props: [CFString: Any] = [kCGImagePropertyDPIWidth: 144, kCGImagePropertyDPIHeight: 144]
    if isJPEG { props[kCGImageDestinationLossyCompressionQuality] = 0.92 }
    CGImageDestinationAddImage(dest, img, props as CFDictionary)
    if CGImageDestinationFinalize(dest) {
        written += 1
        print("\(name) \(w)x\(h)")
    }
}
if written == 0 { fail("no pages rendered") }
