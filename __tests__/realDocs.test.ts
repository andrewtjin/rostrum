// AUTOMATED REAL-DOCUMENT TESTS — the headless automation of the manual "wet test".
//
// Runs engine checks over the ACTUAL `.docx` files in `rostrum/samples/` (see
// realDocs.ts): outline-level resolution fidelity, the Stage 4.2g relationship-type
// coverage guard, and the NBSP space-bridge regressions on the real dds2 card.
//
// The heavyweight hide/showAll round-trips (full engine + faked adapter per sample)
// live in the realDocsShard*.test.ts family instead — Jest parallelizes per test FILE,
// and keeping every sample's CPU-bound round-trip here made this one file the whole
// suite's wall (see runHideReverseShard in realDocs.ts). This file also guards the
// shard family's wiring so a deleted/drifted shard can't silently drop sample coverage.
//
// The outline-level resolution suite below always runs (no samples needed).

import * as fs from "fs";
import * as path from "path";
import { classifyParagraph } from "../src/core/invisibility";
import { readRuns } from "../src/core/ooxml";
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import {
  discoverSamples,
  paragraphsFromDocumentXml,
  readDocxParts,
  readDocumentRelTypes,
  SampleRef
} from "./realDocs";

const all = discoverSamples();

// ---------------------------------------------------------------------------
// Outline-level resolution fidelity (basedOn cascade). Always runs — proves the
// harness classifies a paragraph's outline level the way the live host's
// `Paragraph.outlineLevel` does, INCLUDING style inheritance (Analytics → Heading4).
// The structural basedOn cascade is preferred; a heading-name map is the LAST-RESORT fallback for
// when styles.xml can't resolve a level (adversarial-review C-1), so both paths are covered here.
// ---------------------------------------------------------------------------
describe("outline-level resolution (basedOn cascade → Paragraph.outlineLevel)", () => {
  const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
  const styles =
    `<w:styles xmlns:w="${W}">` +
    `<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Heading4"><w:name w:val="heading 4"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Analytics"><w:name w:val="Analytics"/><w:basedOn w:val="Heading4"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Analytic"><w:name w:val="Analytic"/><w:basedOn w:val="Analytics"/></w:style>` +
    `</w:styles>`;
  const body = (inner: string): string => `<w:document xmlns:w="${W}"><w:body>${inner}</w:body></w:document>`;
  const para = (style: string | null, text: string, inlineOl?: number): string => {
    const ppr =
      (style ? `<w:pStyle w:val="${style}"/>` : "") +
      (inlineOl !== undefined ? `<w:outlineLvl w:val="${inlineOl}"/>` : "");
    return `<w:p>${ppr ? `<w:pPr>${ppr}</w:pPr>` : ""}<w:r><w:t>${text}</w:t></w:r></w:p>`;
  };

  it("resolves outline level through basedOn (Analytics & Analytic → Heading4 → level 3)", () => {
    const doc = body(
      para("Heading1", "pocket") +
        para("Heading4", "tag") +
        para("Analytics", "analytic") +
        para("Analytic", "typo dup") +
        para("Normal", "card body") +
        para(null, "no style")
    );
    // OOXML 0-based outlineLvl → Word 1-based: H1(0)→1, H4(3)→4, Analytics/Analytic→4, body→10.
    expect(paragraphsFromDocumentXml(doc, styles).map((p) => p.outlineNumber)).toEqual([1, 4, 4, 4, 10, 10]);
  });

  it("a direct paragraph <w:outlineLvl> overrides the style", () => {
    const doc = body(para("Normal", "promoted", 0) + para("Heading4", "demoted", 9));
    expect(paragraphsFromDocumentXml(doc, styles).map((p) => p.outlineNumber)).toEqual([1, 10]);
  });

  it("name-fallback keeps a built-in heading with no styles.xml; a non-heading custom style stays body", () => {
    // Adversarial-review C-1: when styles.xml can't resolve a level (here it's absent), a HEADING-named
    // pStyle still reports as a heading via the last-resort name map (the level Word carries latently),
    // so the heading isn't hidden. A non-heading custom style has no signal at all → body.
    expect(paragraphsFromDocumentXml(body(para("Heading4", "x")), null).map((p) => p.outlineNumber)).toEqual([4]);
    expect(paragraphsFromDocumentXml(body(para("CardBody", "y")), null).map((p) => p.outlineNumber)).toEqual([10]);
  });
});

// ---------------------------------------------------------------------------
// Shard-family integrity guard. The hide/showAll round-trips run round-robin across the
// realDocsShard*.test.ts files (sample i runs in shard i % N — runHideReverseShard in
// realDocs.ts). The failure mode this guards is SILENT: delete a shard file, rename it
// off the `realDocs` prefix, or let a file's (shard, of) arguments drift from the family
// size, and some samples simply stop being round-tripped — nothing else would fail.
// This pins the family to its contract, and needs no samples, so it also runs on CI.
// ---------------------------------------------------------------------------
describe("real-doc shard family integrity (realDocsShard*.test.ts)", () => {
  const SHARD_RE = /^realDocsShard(\d+)\.test\.ts$/;
  const shardFiles = fs.readdirSync(__dirname).filter((f) => SHARD_RE.test(f));

  it("shards exist, are contiguous 1..N, and each wires (k-1, N) into the shared runner", () => {
    const n = shardFiles.length;
    // At least 2: one shard would re-serialize the round-trips into a single worker,
    // recreating the very wall the family exists to remove.
    expect(n).toBeGreaterThanOrEqual(2);

    // Contiguous 1..N — a gap means a deleted/renamed shard left a hole in coverage.
    const indices = shardFiles.map((f) => Number(SHARD_RE.exec(f)![1])).sort((a, b) => a - b);
    expect(indices).toEqual(Array.from({ length: n }, (_, i) => i + 1));

    // Each shard k must call runHideReverseShard(k-1, N) with N = the FAMILY size, so
    // the modulo partition is exact: every active sample lands in exactly one shard.
    for (const f of shardFiles) {
      const k = Number(SHARD_RE.exec(f)![1]);
      const src = fs.readFileSync(path.join(__dirname, f), "utf8");
      expect({ file: f, wired: new RegExp(`runHideReverseShard\\(\\s*${k - 1}\\s*,\\s*${n}\\s*\\)`).test(src) })
        .toEqual({ file: f, wired: true });
    }
  });
});

// ---------------------------------------------------------------------------
// Stage 4.2g coverage guard. A per-paragraph `insertOoxml` commit renders the inserted
// fragment against ONLY the parts `commitXml` bundles, so a body paragraph that inherits
// formatting from an UNBUNDLED part silently collapses to document defaults (the wet-test
// "underline/box/18pt → plain 11pt" bug). This asserts EVERY sample doc's
// document-relationship types are CLASSIFIED — so a future doc carrying a NEW part type
// fails loudly HERE (headless) instead of silently losing formatting on the live host.
// Reads only the tiny .rels, so it runs on ALL tiers (incl. the deferred heavy ones, which
// carry the most exotic parts — footnotes/endnotes/image).
// ---------------------------------------------------------------------------
describe("commitXml part coverage across real docs (Stage 4.2g formatting guard)", () => {
  // Bundled by commitXml — MUST mirror AUX_REL_TYPES in ooxmlPackage.ts. These are the parts a
  // body paragraph inherits run/paragraph FORMATTING from (underline, border/box, size, list).
  const BUNDLED = new Set(["styles", "numbering", "theme"]);
  // Render-relevant but DELIBERATELY not bundled — proven non-load-bearing for a single inserted
  // paragraph on the live host (ndca rendered perfectly without them): fontTable (fonts come from
  // the OS; bundling only inflates the commit), settings/webSettings (document-level — excluded so
  // an inserted fragment can't re-impose track-changes/protection; they don't drive run formatting).
  const SAFE_TO_SKIP = new Set(["fontTable", "settings", "webSettings"]);
  // Not formatting-bearing for the kept/hidden paragraph: hyperlink/image relationships are carried
  // PER-PARAGRAPH by commitXml (referencedRelIds), and footnotes/endnotes/comments/headers/footers/
  // customXml/glossary carry CONTENT, not a body run's formatting. (An inline image's binary or a
  // footnote's body is a separate, non-formatting concern — out of scope for THIS guard.)
  const KNOWN_NON_FORMATTING = new Set([
    "hyperlink", "image", "footnotes", "endnotes", "comments",
    "header", "footer", "customXml", "glossaryDocument"
  ]);

  if (!all.length) {
    it("no .docx samples to guard — drop files into rostrum/samples", () => {
      expect(all).toHaveLength(0);
    });
    // No fixtures (e.g. CI — samples/ is gitignored): skip the data-driven guard. Calling
    // `it.each([])` is a hard Jest error ("empty Array of table data"), so we must `return`
    // BEFORE it, mirroring the empty-shard guard in runHideReverseShard (realDocs.ts).
    return;
  }

  it.each(all)(
    "every document-relationship type is classified (no silent formatting gap) — $file [$tier]",
    async (s: SampleRef) => {
      const types = await readDocumentRelTypes(s.fullPath);
      const unclassified = types.filter(
        (t) => !BUNDLED.has(t) && !SAFE_TO_SKIP.has(t) && !KNOWN_NON_FORMATTING.has(t)
      );
      // The alarm. A new, unclassified relationship type means: DECIDE whether a body paragraph
      // inherits formatting from it. If YES → add it to AUX_REL_TYPES + AUX_PART_CONTENT_TYPES in
      // ooxmlPackage.ts (and to BUNDLED here). If it's render-relevant but proven not to matter →
      // SAFE_TO_SKIP. If it carries content, not formatting → KNOWN_NON_FORMATTING with a why.
      expect({ file: s.file, unclassified }).toEqual({ file: s.file, unclassified: [] });
    }
  );
});

// ---------------------------------------------------------------------------
// NBSP space-bridge regression on the REAL card (dds2 bunzel-18). The card's inter-word
// separators are non-breaking spaces (U+00A0); before the predicate fix the bridge didn't
// see them as spaces, so hidden NBSP gaps fused kept words ("rising oil revenues" →
// "risingrevenues", "gives Russia" → "givesRussia"). This asserts the fix end-to-end on the
// actual file, not just synthetic runs.
// ---------------------------------------------------------------------------
describe("NBSP bridge on a real card (dds2 bunzel-18 regression)", () => {
  const dds2 = all.find((s) => /dds2/i.test(s.file));
  const KEEP = new Set([
    "yellow", "green", "cyan", "magenta", "blue", "red", "darkblue", "darkcyan",
    "darkgreen", "darkmagenta", "darkred", "darkyellow", "darkgray", "lightgray", "black", "white"
  ]);
  const maybe = dds2 ? it : it.skip;

  maybe("condenses 'rising oil revenues' WITHOUT fusing the kept words", async () => {
    const { documentXml } = await readDocxParts(dds2!.fullPath);
    const wb = new WholeBodyPackage(documentXml);
    let checked = 0;
    for (let i = 0; i < wb.count; i++) {
      // The card's separators are NBSP, so the finder must be whitespace-tolerant (JS \s
      // matches U+00A0) — a literal-space match finds nothing, which is the bug itself.
      if (!/rising\s+oil\s+revenues/.test(wb.paragraphText(i))) continue;
      checked++;
      const ooxml = wb.paragraphXml(i);
      const plan = classifyParagraph({ index: i, headingLevel: null, inTable: false, ooxml }, { keepColors: KEEP });
      const visible = readRuns(plan.ooxml).filter((r) => !r.hidden).map((r) => r.text).join("");
      // The hidden NBSP separators are preserved as separators → kept words stay apart.
      expect(visible).not.toMatch(/risingrevenues/);
      expect(visible).not.toMatch(/givesRussia/);
      // ...and a separator survives between the kept words (the exposed NBSP renders as a space).
      expect(visible).toMatch(/rising\s+revenues/);
    }
    expect(checked).toBeGreaterThan(0);
  });

  maybe("condenses the palmeri 'capital—and' / 'it—than' card WITHOUT fusing across em dashes", async () => {
    const { documentXml } = await readDocxParts(dds2!.fullPath);
    const wb = new WholeBodyPackage(documentXml);
    let checked = 0;
    for (let i = 0; i < wb.count; i++) {
      const text = wb.paragraphText(i);
      if (!/knows more about/i.test(text) || !/capital/i.test(text)) continue;
      checked++;
      const ooxml = wb.paragraphXml(i);
      const plan = classifyParagraph({ index: i, headingLevel: null, inTable: false, ooxml }, { keepColors: KEEP });
      const visible = readRuns(plan.ooxml).filter((r) => !r.hidden).map((r) => r.text).join("");
      // The hidden em dashes (U+2014) are preserved as separators → kept words no longer fuse.
      expect(visible).not.toMatch(/cand Trump/i); // "capital—and" was fusing to "cand"
      expect(visible).not.toMatch(/itthan/i); // "it—than" was fusing to "itthan"
      expect(visible).toMatch(/c—and/); // the em dash is exposed as the separator
    }
    expect(checked).toBeGreaterThan(0);
  });
});
