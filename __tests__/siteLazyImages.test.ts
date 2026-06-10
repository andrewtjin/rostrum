// Guards the lazy-loading contract of the comparison-page screenshot <img> tags.
//
// site/comparison.html shows each hidden-page PNG twice: a visible thumbnail and
// a same-URL copy inside a `display: none` CSS-only lightbox. display:none
// suppresses CSS background-image fetches but NOT <img> element fetches, and
// browsers dedupe same-URL images per document — so a single eager copy anywhere
// in the page forces both PNGs (the heaviest assets on the page) to download at
// parse time, silently defeating every loading="lazy" on the thumbnails. This
// suite fails if any screenshot <img> — today's four tags or a future addition —
// drops the attribute. Lazy is safe on the lightbox copies: the thumbnails fetch
// the same URLs on scroll, so the zoom is served from cache before any click.
//
// Companion to siteAssets.test.ts, which guards the BYTE WEIGHT of these PNGs;
// this suite guards WHEN those bytes are requested.
import * as fs from "fs";
import * as path from "path";

// CopyWebpackPlugin copies site/ into dist/ verbatim (webpack.config.js), so
// asserting on the source file is asserting on the served page.
const PAGE = path.join(__dirname, "..", "site", "comparison.html");

describe("comparison-page screenshot <img> tags stay lazy", () => {
  const html = fs.readFileSync(PAGE, "utf8");
  // `[^>]` matches newlines inside a character class, so the multi-line
  // formatted thumbnail tags are captured whole alongside the one-line ones.
  const imgTags = html.match(/<img\b[^>]*>/g) || [];
  const screenshotTags = imgTags.filter((tag) => /assets\/hidden-[^"']*\.png/.test(tag));

  it("finds the screenshot tags (regex stays in sync with the markup)", () => {
    // Two thumbnails + two lightbox copies today. If a markup restructure ever
    // hid them from the regex, the lazy assertion below would pass vacuously —
    // pinning the floor here keeps the guard genuine.
    expect(screenshotTags.length).toBeGreaterThanOrEqual(4);
  });

  it('every screenshot <img> carries loading="lazy"', () => {
    // Filter-then-compare (instead of asserting inside a loop) so a failure
    // prints the exact offending tag(s).
    const eager = screenshotTags.filter((tag) => !tag.includes('loading="lazy"'));
    expect(eager).toEqual([]);
  });
});
