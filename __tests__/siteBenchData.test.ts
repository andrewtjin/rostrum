// Guards the comparison-page interactive benchmark charts (bench-data.js +
// bench-chart.js, swapped in over the static SVG <img> fallbacks).
//
// The load-bearing assertion is PRIVACY: the benchmark corpus is real tournament
// documents whose filenames carry student and school names, so the shipped
// dataset must be PURE NUMBERS — a regenerated bench-data.js that ever grew a
// filename/category/text column must fail CI, not ship. The rest pins the
// progressive-enhancement contract: the page stays complete without JS (static
// imgs + lightboxes remain in the markup), and the scripts load deferred with
// the data ahead of the renderer.
import * as fs from "fs";
import * as path from "path";

const SITE = path.join(__dirname, "..", "site");
const dataJs = fs.readFileSync(path.join(SITE, "bench-data.js"), "utf8");
const chartJs = fs.readFileSync(path.join(SITE, "bench-chart.js"), "utf8");
const page = fs.readFileSync(path.join(SITE, "comparison.html"), "utf8");

describe("bench-data.js stays a numbers-only dataset", () => {
  // everything between `rows: [` and the closing `]` of the array — the part
  // regenerated from the private corpus and therefore the part that could leak
  const rowsMatch = dataJs.match(/rows:\s*(\[\[[\s\S]*?\]\])/);

  it("declares the expected global and row payload", () => {
    expect(dataJs).toContain("window.ROSTRUM_BENCH");
    expect(rowsMatch).not.toBeNull();
  });

  it("row payload contains ONLY numbers, brackets, and null — no text", () => {
    // null marks a tool that never finished a doc; after removing that one
    // token, any remaining letter means a string column crept in.
    const stripped = (rowsMatch as RegExpMatchArray)[1].replace(/null/g, "");
    expect(stripped).toMatch(/^[\d[\],.\s-]+$/);
  });

  it("carries the full corpus (one row per document)", () => {
    const rowCount = ((rowsMatch as RegExpMatchArray)[1].match(/\[/g) || []).length - 1;
    expect(rowCount).toBeGreaterThanOrEqual(800);
  });

  it("fits a 64KB shipped-bytes budget", () => {
    // ~32KB today; doubling would signal an accidental extra column or
    // precision blow-up rather than legitimate corpus growth.
    expect(Buffer.byteLength(dataJs, "utf8")).toBeLessThanOrEqual(64 * 1024);
  });
});

describe("comparison.html keeps the enhancement progressive", () => {
  it("loads both scripts deferred, data before renderer", () => {
    const dataIdx = page.indexOf('<script src="./bench-data.js" defer>');
    const chartIdx = page.indexOf('<script src="./bench-chart.js" defer>');
    expect(dataIdx).toBeGreaterThan(-1);
    expect(chartIdx).toBeGreaterThan(dataIdx);
  });

  it("marks both chart figures for the script to find", () => {
    expect(page).toContain('data-bench="words"');
    expect(page).toContain('data-bench="cards"');
  });

  it("keeps the no-JS fallback: static chart imgs wrapped in lightbox zooms", () => {
    // every content image on the site must be zoomable (CSS-only :target
    // lightbox) so the page degrades to static-with-zoom, not static-flat
    for (const kind of ["words", "cards"]) {
      expect(page).toContain(`href="#zoom-bench-${kind}"`);
      expect(page).toContain(`id="zoom-bench-${kind}"`);
    }
  });

  it("bench-chart.js renders from the global the data file sets", () => {
    expect(chartJs).toContain("window.ROSTRUM_BENCH");
  });
});
