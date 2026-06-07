// CITE REPAIR — the leak-critical suite.
//
// Two halves:
//   1. A REAL regression against `[small] 2ac---ndca---semis.docx` — the whole point: a
//      cite ("Valcke et al. 20") that lost its cite character style is mis-classified as
//      body and hidden. We prove it is NOT kept before repair, that `planCiteRepairs`
//      targets its bold author run, and that after `applyCiteStyleToParagraphXml` it has
//      the cite rStyle and IS kept. The cite paragraph is found by TEXT, never by index.
//   2. SYNTHETIC unit cases for `planCiteRepairs` (positive + false-positive leak guards)
//      and `applyCiteStyleToParagraphXml` (the four rPr shapes + idempotency + byte-
//      preservation of everything except the injected rStyle).
//
// Leak-safety framing: a FALSE POSITIVE keeps body text that should be hidden (a real
// invisibility leak), so the false-positive guards are as important as the positive case.

import * as fs from "fs";
import * as path from "path";
import {
  applyCiteStyleToParagraphXml,
  CiteRepairParagraph,
  planCiteRepairs
} from "../src/core/citeRepair";
import { CITE_STYLE_ID } from "../src/core/styles";
import { readRuns } from "../src/core/ooxml";
import { paragraphHasCiteRun } from "../src/core/keepers";
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { readDocxParts, SAMPLES_DIR } from "./realDocs";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

// ---------------------------------------------------------------------------
// Synthetic OOXML builders (mirror the real run shapes seen in the ndca sample)
// ---------------------------------------------------------------------------

/** A bold run (`<w:b/>`) with text — the cite author shape (no rStyle, just bold). */
function boldRun(text: string): string {
  return `<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

/** A plain (non-bold, no style) run with text — the cite's descriptor tail / body prose. */
function plainRun(text: string): string {
  return `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

/** A run that already carries the cite character style (a proper cite). */
function citeRun(text: string): string {
  return `<w:r><w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/><w:b/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;
}

/** Wrap inner run markup in a standalone `<w:p>`; `pPr` injects pPr markup (numbering, etc.). */
function p(inner: string, pPr = ""): string {
  return `<w:p xmlns:w="${W}">${pPr ? `<w:pPr>${pPr}</w:pPr>` : ""}${inner}</w:p>`;
}

/** A heading/tag paragraph view (headingLevel set). */
function tag(inner: string, level = 3): CiteRepairParagraph {
  return { xml: p(inner), headingLevel: level };
}

/** A body paragraph view (headingLevel null). */
function body(inner: string, pPr = ""): CiteRepairParagraph {
  return { xml: p(inner, pPr), headingLevel: null };
}

/** A `<w:numPr>` pPr fragment marking a list/bullet paragraph. */
const NUM_PR = `<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>`;

// ---------------------------------------------------------------------------
// 1. REAL regression — the ndca "Valcke et al. 20" cite
// ---------------------------------------------------------------------------
// The ndca sample lives in the gitignored samples/ dir — present locally, ABSENT in CI. Skip
// this real-doc regression when the fixture is missing rather than failing the build (the
// synthetic suites below are the engine coverage CI can actually run). describe.skip keeps the
// skip VISIBLE in the report instead of silently passing.
const NDCA_SAMPLE = path.join(SAMPLES_DIR, "[small] 2ac---ndca---semis.docx");
const describeNdca = fs.existsSync(NDCA_SAMPLE) ? describe : describe.skip;

describeNdca("real regression: ndca-semis 'Valcke et al. 20' mis-styled cite", () => {
  it("is not kept before repair, is planned, and is kept after the rStyle injection", async () => {
    const { documentXml } = await readDocxParts(NDCA_SAMPLE);
    const pkg = new WholeBodyPackage(documentXml);

    // Find the cite paragraph by TEXT (never a hardcoded index).
    let citeIdx = -1;
    for (let i = 0; i < pkg.count; i++) {
      if (pkg.paragraphText(i).includes("Valcke et al. 20")) {
        citeIdx = i;
        break;
      }
    }
    expect(citeIdx).toBeGreaterThanOrEqual(0);

    const paras: CiteRepairParagraph[] = [];
    for (let i = 0; i < pkg.count; i++) {
      paras.push({ xml: pkg.paragraphXml(i), headingLevel: pkg.headingLevel(i) });
    }
    const citeXml = paras[citeIdx].xml;

    // (i) BEFORE: no cite-styled run, and the keeper rule does NOT keep it.
    expect(citeXml).not.toContain(`<w:rStyle w:val="${CITE_STYLE_ID}"/>`);
    expect(paragraphHasCiteRun(readRuns(citeXml))).toBe(false);

    // (ii) planCiteRepairs returns a repair targeting this paragraph's bold author run.
    const repairs = planCiteRepairs(paras, CITE_STYLE_ID);
    const repair = repairs.find((r) => r.paragraphIndex === citeIdx);
    expect(repair).toBeDefined();
    expect(repair!.runIndices.length).toBeGreaterThanOrEqual(1);
    // The first targeted run is the bold "Valcke et al. 20" author run (index 0).
    expect(repair!.runIndices).toContain(0);

    // (iii) AFTER: the paragraph has a cite rStyle run AND is kept.
    const repaired = applyCiteStyleToParagraphXml(citeXml, repair!.runIndices, CITE_STYLE_ID);
    expect(repaired).toContain(`<w:rStyle w:val="${CITE_STYLE_ID}"/>`);
    expect(paragraphHasCiteRun(readRuns(repaired))).toBe(true);
  });
});

// ---------------------------------------------------------------------------
// 2a. planCiteRepairs — positive cases
// ---------------------------------------------------------------------------
describe("planCiteRepairs — qualifying cites", () => {
  it("all-bold author+date ('Valcke et al. 20')", () => {
    const paras = [tag("Tag text"), body(boldRun("Valcke et al. 20") + plainRun(" [descriptor]"))];
    const repairs = planCiteRepairs(paras);
    expect(repairs).toEqual([{ paragraphIndex: 1, runIndices: [0] }]);
  });

  it("first-name-last-name with only the last name bold ('Barbara **Valcke** 20')", () => {
    // "Barbara " is a non-bold prefix (name-like, ≤30, no digit/punct); "Valcke" and "20"
    // are bold → both repaired (every bold run up to and including the year run).
    const inner = plainRun("Barbara ") + boldRun("Valcke") + boldRun(" 20") + plainRun(" [descriptor]");
    const repairs = planCiteRepairs([tag("Tag"), body(inner)]);
    expect(repairs).toEqual([{ paragraphIndex: 1, runIndices: [1, 2] }]);
  });

  it("skips bullet (numPr) paragraphs under the tag, repairs the cite below them", () => {
    const paras = [
      tag("Tag"),
      body(plainRun("first bullet"), NUM_PR),
      body(plainRun("second bullet"), NUM_PR),
      body(boldRun("Smith 21") + plainRun(" [descriptor]"))
    ];
    const repairs = planCiteRepairs(paras);
    expect(repairs).toEqual([{ paragraphIndex: 3, runIndices: [0] }]);
  });

  it("skips empty paragraphs under the tag, repairs the cite below them", () => {
    const paras = [
      tag("Tag"),
      body(""), // empty <w:p>
      body(plainRun("   ")), // whitespace-only → empty after trim
      body(boldRun("Jones 19") + plainRun(" [descriptor]"))
    ];
    const repairs = planCiteRepairs(paras);
    expect(repairs).toEqual([{ paragraphIndex: 3, runIndices: [0] }]);
  });

  it("restarts the window at a deeper subtag (tag → subtag → cite)", () => {
    const paras = [
      tag("Block", 2),
      tag("Tag", 3),
      body(boldRun("Doe 20") + plainRun(" [descriptor]"))
    ];
    const repairs = planCiteRepairs(paras);
    // The cite is the first real paragraph after the SUBTAG; planned once.
    expect(repairs).toEqual([{ paragraphIndex: 2, runIndices: [0] }]);
  });

  it("accepts apostrophe-year shorthand (Author '20)", () => {
    const repairs = planCiteRepairs([tag("Tag"), body(boldRun("Doe ’20") + plainRun(" [d]"))]);
    expect(repairs).toEqual([{ paragraphIndex: 1, runIndices: [0] }]);
  });
});

// ---------------------------------------------------------------------------
// 2b. planCiteRepairs — must NOT repair (negatives + leak guards)
// ---------------------------------------------------------------------------
describe("planCiteRepairs — non-repairs (leak prevention)", () => {
  it("an already cite-styled paragraph is not repaired", () => {
    const paras = [tag("Tag"), body(citeRun("Smith 20") + plainRun(" [descriptor]"))];
    expect(planCiteRepairs(paras)).toEqual([]);
  });

  it("a paragraph NOT after a heading is not repaired", () => {
    // No tag anywhere → no window opens, even though the paragraph looks cite-like.
    const paras = [body(plainRun("intro")), body(boldRun("Smith 20") + plainRun(" [d]"))];
    expect(planCiteRepairs(paras)).toEqual([]);
  });

  it("body prose 'In 2008, the crisis worsened and **experts** said' is NOT repaired", () => {
    // Year is early, but the first bold run is DEEP and the prefix has a digit + comma.
    const inner =
      plainRun("In 2008, the crisis worsened and ") + boldRun("experts") + plainRun(" said the worst.");
    expect(planCiteRepairs([tag("Tag"), body(inner)])).toEqual([]);
  });

  it("a non-bold paragraph containing a year is NOT repaired", () => {
    expect(planCiteRepairs([tag("Tag"), body(plainRun("Smith 2020 says something"))])).toEqual([]);
  });

  it("'Smith 2020' deep in body prose (not the first post-tag candidate) is NOT repaired", () => {
    const paras = [
      tag("Tag"),
      body(boldRun("Author 20") + plainRun(" [proper cite descriptor]")), // the real candidate
      body(plainRun("Body text later cites ") + boldRun("Smith 2020") + plainRun(" in passing.")) // deeper
    ];
    const repairs = planCiteRepairs(paras);
    // Only the FIRST post-tag candidate is considered; the deep mention is never reached.
    expect(repairs).toEqual([{ paragraphIndex: 1, runIndices: [0] }]);
  });

  it("a candidate with no year near the front is NOT repaired", () => {
    const inner = boldRun("Smith") + plainRun(" argues at length without any date in the first part");
    expect(planCiteRepairs([tag("Tag"), body(inner)])).toEqual([]);
  });

  it("a candidate whose first bold word is past char 30 is NOT repaired", () => {
    // Year present early, but the bold word starts well past offset 30 and prefix is prose.
    const prefix = "This sentence runs on for quite a while before "; // > 30 chars
    const inner = plainRun(`${prefix}`) + boldRun("word") + plainRun(" 20 more text");
    expect(planCiteRepairs([tag("Tag"), body(inner)])).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 3. applyCiteStyleToParagraphXml — the four rPr shapes + idempotency + fidelity
// ---------------------------------------------------------------------------
describe("applyCiteStyleToParagraphXml", () => {
  it("creates an rPr (as the run's first child) when the run has none", () => {
    const xml = p(`<w:r><w:t>x</w:t></w:r>`);
    const out = applyCiteStyleToParagraphXml(xml, [0], CITE_STYLE_ID);
    expect(out).toContain(`<w:r><w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/></w:rPr><w:t>x</w:t></w:r>`);
  });

  it("inserts rStyle as the FIRST child of an existing rPr (no rStyle yet)", () => {
    const xml = p(`<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t>x</w:t></w:r>`);
    const out = applyCiteStyleToParagraphXml(xml, [0], CITE_STYLE_ID);
    expect(out).toContain(`<w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/><w:b/><w:bCs/></w:rPr>`);
  });

  it("replaces an existing rStyle's w:val", () => {
    const xml = p(`<w:r><w:rPr><w:rStyle w:val="SomethingElse"/><w:b/></w:rPr><w:t>x</w:t></w:r>`);
    const out = applyCiteStyleToParagraphXml(xml, [0], CITE_STYLE_ID);
    expect(out).toContain(`<w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/><w:b/></w:rPr>`);
    expect(out).not.toContain("SomethingElse");
  });

  it("expands a self-closing <w:rPr/> into a paired rPr carrying the rStyle", () => {
    const xml = p(`<w:r><w:rPr/><w:t>x</w:t></w:r>`);
    const out = applyCiteStyleToParagraphXml(xml, [0], CITE_STYLE_ID);
    expect(out).toContain(`<w:r><w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/></w:rPr><w:t>x</w:t></w:r>`);
  });

  it("is idempotent (applying twice equals applying once)", () => {
    const xml = p(`<w:r><w:rPr><w:b/></w:rPr><w:t>x</w:t></w:r>`);
    const once = applyCiteStyleToParagraphXml(xml, [0], CITE_STYLE_ID);
    const twice = applyCiteStyleToParagraphXml(once, [0], CITE_STYLE_ID);
    expect(twice).toBe(once);
  });

  it("leaves NON-targeted runs and all other markup byte-identical (only rStyle added)", () => {
    const r0 = `<w:r><w:rPr><w:b/><w:bCs/><w:sz w:val="26"/></w:rPr><w:t>Valcke et al. 20</w:t></w:r>`;
    const r1 = `<w:r><w:t xml:space="preserve"> [Barbara Valcke &amp; al.]</w:t></w:r>`;
    const xml = p(r0 + r1);
    const out = applyCiteStyleToParagraphXml(xml, [0], CITE_STYLE_ID);
    // Run 1 (non-targeted) is untouched.
    expect(out).toContain(r1);
    // The ONLY difference is the inserted rStyle as the first child of run 0's rPr.
    const expectedR0 = `<w:r><w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/><w:b/><w:bCs/><w:sz w:val="26"/></w:rPr><w:t>Valcke et al. 20</w:t></w:r>`;
    expect(out).toBe(p(expectedR0 + r1));
  });

  it("returns the input unchanged when no run indices are given", () => {
    const xml = p(`<w:r><w:t>x</w:t></w:r>`);
    expect(applyCiteStyleToParagraphXml(xml, [], CITE_STYLE_ID)).toBe(xml);
  });

  it("repairs only the targeted run when multiple runs exist", () => {
    const xml = p(boldRun("A 20") + boldRun("B"));
    const out = applyCiteStyleToParagraphXml(xml, [1], CITE_STYLE_ID);
    // Run 0 unchanged (still no rStyle); run 1 gets it.
    expect(out).toContain(`<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">A 20</w:t></w:r>`);
    expect(out).toContain(`<w:r><w:rPr><w:rStyle w:val="${CITE_STYLE_ID}"/><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">B</w:t></w:r>`);
  });
});
