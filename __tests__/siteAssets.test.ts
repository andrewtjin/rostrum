// Guards the shipped weight of the comparison-page screenshots (assets/ is copied
// verbatim into dist/ by CopyWebpackPlugin and served to site/comparison.html — a
// page students hit on tournament Wi-Fi). The originals shipped as raw screen
// captures (81KB + 85KB = 45% of all non-JS shipped bytes); a lossless zopflipng
// pass cut them to ~32KB each. This suite makes re-bloating them (e.g. pasting a
// fresh unoptimized capture) a test failure instead of a silent shipped-bytes
// regression.
//
// When a screenshot legitimately changes: re-run a LOSSLESS optimizer (zopflipng /
// optipng -o7) on the new capture, then update the expected dimensions below.
import * as fs from "fs";
import * as path from "path";

// Per-file byte budget. Current optimized sizes are ~32KB; an unoptimized re-export
// of comparable content lands at ~80KB+, so 48KB allows modest legitimate growth
// while still catching any capture that skipped the lossless recompression step.
const BUDGET_BYTES = 48 * 1024;

// Expected pixel dimensions, pinned so the byte budget cannot be satisfied by
// lossily DOWNSCALING the screenshot instead of recompressing it (losslessness is
// the product's core promise — the marketing assets should hold the same bar).
const SCREENSHOTS: ReadonlyArray<{ file: string; width: number; height: number }> = [
  { file: "hidden-rostrum.png", width: 1705, height: 607 },
  { file: "hidden-verbatim.png", width: 1753, height: 880 },
];

// PNG files open with a fixed 8-byte signature, immediately followed by the IHDR
// chunk whose first two fields are big-endian uint32 width (offset 16) and height
// (offset 20) — enough structure to validate without a PNG decoder dependency.
const PNG_SIGNATURE = Buffer.from([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]);

describe("comparison-page screenshots stay optimized", () => {
  for (const { file, width, height } of SCREENSHOTS) {
    const fullPath = path.join(__dirname, "..", "assets", file);

    it(`${file} is a valid PNG at the expected ${width}x${height}`, () => {
      const bytes = fs.readFileSync(fullPath);
      // Signature check catches truncation/corruption from a bad optimizer run.
      expect(bytes.subarray(0, 8).equals(PNG_SIGNATURE)).toBe(true);
      expect(bytes.readUInt32BE(16)).toBe(width);
      expect(bytes.readUInt32BE(20)).toBe(height);
    });

    it(`${file} fits the ${BUDGET_BYTES / 1024}KB shipped-bytes budget`, () => {
      expect(fs.statSync(fullPath).size).toBeLessThanOrEqual(BUDGET_BYTES);
    });
  }
});
