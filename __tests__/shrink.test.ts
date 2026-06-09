// Unit tests for the pure Shrink engine (src/core/shrink.ts): the size ladder, the keep predicate,
// the omission scan, heading refusal, mixed-size normalization, and Unshrink.
import {
  keepFullSize,
  nextShrinkSize,
  omissionRunIndices,
  shrinkFragment,
  unshrinkFragment
} from "../src/core/shrink";
import { readFragmentParagraphs, resolveStyleEmphasis } from "../src/core/ooxmlCondense";
import { OmissionPattern, ShrinkOptions } from "../src/core/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

interface RunOpts {
  u?: boolean;
  sz?: number;
  highlight?: string;
  cite?: boolean;
}
function run(text: string, o: RunOpts = {}): string {
  const rPr: string[] = [];
  if (o.cite) rPr.push(`<w:rStyle w:val="Style13ptBold"/>`);
  if (o.sz) rPr.push(`<w:sz w:val="${o.sz}"/>`);
  if (o.highlight) rPr.push(`<w:highlight w:val="${o.highlight}"/>`);
  if (o.u) rPr.push(`<w:u w:val="single"/>`);
  const rPrXml = rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";
  return `<w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}
const bareP = (inner: string): string => `<w:p xmlns:w="${W_NS}">${inner}</w:p>`;

/** A run carrying a character STYLE (rStyle) instead of direct formatting — how real docs encode the cut. */
const runStyled = (text: string, styleId: string): string =>
  `<w:r><w:rPr><w:rStyle w:val="${styleId}"/></w:rPr><w:t xml:space="preserve">${text}</w:t></w:r>`;

/** Character-style defs mirroring a real doc: StyleUnderline (u), Emphasis (u + box), Plain (bold only). */
const STYLE_DEFS =
  `<w:style w:type="character" w:styleId="StyleUnderline"><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:u w:val="single"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="Emphasis"><w:basedOn w:val="DefaultParagraphFont"/><w:rPr><w:u w:val="single"/><w:bdr w:val="single"/></w:rPr></w:style>` +
  `<w:style w:type="character" w:styleId="Plain"><w:rPr><w:b/></w:rPr></w:style>`;

/** A fragment with a styles part + a one-paragraph body, so style-resolved emphasis applies (like a range read). */
const styledP = (inner: string): string =>
  `<pkg xmlns:w="${W_NS}"><w:styles>${STYLE_DEFS}</w:styles><w:body><w:p>${inner}</w:p></w:body></pkg>`;

/** Sizes (half-points) per run in the (single) paragraph after a transform. */
const sizes = (xml: string): (number | null)[] => readFragmentParagraphs(xml)[0].map((r) => r.sizeHalfPts);

const opts = (over: Partial<ShrinkOptions> = {}): ShrinkOptions => ({
  normalHalfPts: 22, // 11pt
  outlineLevels: [null],
  omissionPatterns: [],
  shrinkParagraphMarks: false,
  ...over
});

describe("nextShrinkSize ladder (half-points)", () => {
  it.each([
    [22, 16], // 11pt → 8pt
    [24, 16], // 12pt → 8pt
    [16, 14], // 8 → 7
    [14, 12], // 7 → 6
    [12, 10], // 6 → 5
    [10, 8], //  5 → 4
    [8, null], // 4 → Normal (clear)
    [6, null], // 3pt (below the 4pt floor) → clear to Normal, NEVER grow back up
    [17, 16], // 8.5pt → 8pt (largest rung strictly below; no rounding skip)
    [15, 14] //  7.5pt → 7pt
  ])("from %ihp → %s", (from, expected) => {
    expect(nextShrinkSize(from)).toBe(expected);
  });

  it("never grows: a sub-floor size clears (does not jump up to 8pt)", () => {
    expect(nextShrinkSize(6)).toBeNull();
    expect(nextShrinkSize(8)).toBeNull();
  });
});

describe("keepFullSize predicate", () => {
  const view = (inner: string) => readFragmentParagraphs(bareP(inner))[0][0];
  it("keeps underlined, highlighted, cite, and structural runs full-size", () => {
    expect(keepFullSize(view(run("plain")))).toBe(false);
    expect(keepFullSize(view(run("cut", { u: true })))).toBe(true);
    expect(keepFullSize(view(run("hl", { highlight: "yellow" })))).toBe(true);
    expect(keepFullSize(view(run("cite", { cite: true })))).toBe(true);
    expect(keepFullSize(view(`<w:r><w:fldChar w:fldCharType="begin"/></w:r>`))).toBe(true); // ineligible
  });
});

describe("keepFullSize — style-resolved underline/box (real-doc encoding, bug-1)", () => {
  const first = (frag: string): ReturnType<typeof readFragmentParagraphs>[0][0] => readFragmentParagraphs(frag)[0][0];
  it("keeps a run underlined via a character STYLE (StyleUnderline), not a direct <w:u>", () => {
    const r = first(styledP(runStyled("the cut", "StyleUnderline")));
    expect(r.underline).toBe(true);
    expect(keepFullSize(r)).toBe(true);
  });
  it("keeps a BOXED run via a character style (Emphasis = underline + box)", () => {
    const r = first(styledP(runStyled("emph", "Emphasis")));
    expect(r.boxed).toBe(true);
    expect(keepFullSize(r)).toBe(true);
  });
  it("still shrinks a styled run with no underline/box (Plain = bold only)", () => {
    expect(keepFullSize(first(styledP(runStyled("loud", "Plain"))))).toBe(false);
  });
  it('a direct <w:u w:val="none"/> overrides a style that underlines', () => {
    const r = first(styledP(`<w:r><w:rPr><w:rStyle w:val="StyleUnderline"/><w:u w:val="none"/></w:rPr><w:t>x</w:t></w:r>`));
    expect(r.underline).toBe(false);
    expect(keepFullSize(r)).toBe(false);
  });
});

describe("shrinkFragment — keeps style-underlined/boxed runs full-size (bug-1 end-to-end)", () => {
  it("shrinks only the plain run; StyleUnderline + Emphasis stay full-size", () => {
    const xml = styledP(runStyled("plain ", "Plain") + runStyled("cut", "StyleUnderline") + runStyled("emph", "Emphasis"));
    const out = shrinkFragment(xml, opts());
    expect(out.changed).toBe(true);
    expect(sizes(out.xml)).toEqual([16, null, null]); // plain → 8pt; styled keepers untouched
  });
});

describe("resolveStyleEmphasis (basedOn cascade)", () => {
  const styles =
    `<w:styles xmlns:w="${W_NS}">` +
    `<w:style w:type="character" w:styleId="U"><w:rPr><w:u w:val="single"/></w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="Box"><w:rPr><w:bdr w:val="single"/></w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="ChildOfU"><w:basedOn w:val="U"/><w:rPr><w:b/></w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="NoneOverU"><w:basedOn w:val="U"/><w:rPr><w:u w:val="none"/></w:rPr></w:style>` +
    `<w:style w:type="character" w:styleId="Plain"><w:rPr><w:b/></w:rPr></w:style>` +
    `</w:styles>`;
  const map = resolveStyleEmphasis(styles);
  it("reads own underline/box", () => {
    expect(map.get("U")).toEqual({ underline: true, boxed: false });
    expect(map.get("Box")).toEqual({ underline: false, boxed: true });
  });
  it("inherits underline through basedOn", () => {
    expect(map.get("ChildOfU")?.underline).toBe(true);
  });
  it("a child overrides inherited underline to none", () => {
    expect(map.get("NoneOverU")?.underline).toBe(false);
  });
  it("a non-emphasis style is neither underlined nor boxed", () => {
    expect(map.get("Plain")).toEqual({ underline: false, boxed: false });
  });
});

describe("shrinkFragment", () => {
  it("shrinks non-kept body text to 8pt, keeping underline/highlight/cite full-size", () => {
    const xml = bareP(run("body ") + run("cut", { u: true }) + run(" mid ") + run("hl", { highlight: "yellow" }));
    const out = shrinkFragment(xml, opts());
    expect(out.changed).toBe(true);
    expect(out.appliedSizeHalfPts).toBe(16);
    expect(sizes(out.xml)).toEqual([16, null, 16, null]); // kept runs untouched
  });

  it("normalizes mixed sizes to one target (reads the first non-kept run's size)", () => {
    const xml = bareP(run("aaa", { sz: 16 }) + run("bbb", { sz: 28 }));
    const out = shrinkFragment(xml, opts());
    // representative = first non-kept (16hp = 8pt) → next rung 7pt (14hp); applied to BOTH.
    expect(sizes(out.xml)).toEqual([14, 14]);
  });

  it("refuses to shrink a single-paragraph heading (Verbatim parity)", () => {
    const xml = bareP(run("Heading text"));
    const out = shrinkFragment(xml, opts({ outlineLevels: [0] }));
    expect(out.refusedHeading).toBe(true);
    expect(out.changed).toBe(false);
  });

  it("restores omission spans to Normal (clears their size)", () => {
    const xml = bareP(run("before ") + run("[Omitted: text]") + run(" after"));
    const patterns: OmissionPattern[] = [{ open: "[", close: "]", keyword: "Omitted" }];
    const out = shrinkFragment(xml, opts({ omissionPatterns: patterns }));
    // "before " and " after" shrink to 8pt; the omission run is restored to Normal (null).
    expect(sizes(out.xml)).toEqual([16, null, 16]);
  });

  it("no-ops when every run is a keeper (nothing to shrink)", () => {
    const xml = bareP(run("cut", { u: true }) + run("hl", { highlight: "yellow" }));
    const out = shrinkFragment(xml, opts());
    expect(out.changed).toBe(false);
    expect(out.appliedSizeHalfPts).toBeUndefined();
  });

  it("shrinks the paragraph mark to 6pt when shrinkParagraphMarks is on", () => {
    const xml = bareP(run("body"));
    const out = shrinkFragment(xml, opts({ shrinkParagraphMarks: true }));
    expect(out.xml).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:sz w:val="12"\/>/);
  });

  it("does not GROW text when the Normal size is already below the floor (adversarial H1)", () => {
    // Normal = 6hp (3pt). A plain body run inherits it; Shrink must NOT bump it to 8pt.
    const xml = bareP(run("tiny"));
    const out = shrinkFragment(xml, opts({ normalHalfPts: 6 }));
    expect(out.changed).toBe(false); // clearing a run with no explicit size is a no-op (not a grow)
    expect(out.xml).not.toContain("<w:sz "); // no explicit size was injected
  });

  it("shrinks all-punctuation sized runs via the fallback representative (adversarial L1)", () => {
    const xml = bareP(run("...", { sz: 28 }) + run("!!!", { sz: 28 }));
    const out = shrinkFragment(xml, opts());
    expect(out.changed).toBe(true);
    // 14pt → next rung 8pt; applied to both punctuation runs.
    expect(sizes(out.xml)).toEqual([16, 16]);
  });

  it("cycles down rung by rung across repeated presses", () => {
    let xml = bareP(run("body"));
    const seen: (number | null | undefined)[] = [];
    for (let i = 0; i < 6; i++) {
      const out = shrinkFragment(xml, opts());
      xml = out.xml;
      seen.push(out.appliedSizeHalfPts);
    }
    // 11pt → 8 → 7 → 6 → 5 → 4 → Normal(null)
    expect(seen).toEqual([16, 14, 12, 10, 8, null]);
  });
});

describe("unshrinkFragment", () => {
  it("clears non-kept run sizes, leaving keepers untouched", () => {
    const xml = bareP(run("body", { sz: 16 }) + run("cut", { u: true, sz: 28 }));
    const out = unshrinkFragment(xml, [null]);
    expect(out.changed).toBe(true);
    expect(sizes(out.xml)).toEqual([null, 28]); // body cleared; underlined keeper's size kept
  });
});

describe("omissionRunIndices", () => {
  it("marks runs overlapping a keyword-bearing bracketed span", () => {
    const runs = readFragmentParagraphs(bareP(run("a ") + run("[Omitted x]") + run(" b")))[0];
    const idx = omissionRunIndices(runs, [{ open: "[", close: "]", keyword: "Omitted" }]);
    expect([...idx]).toEqual([1]);
  });
  it("ignores a bracketed span WITHOUT the keyword (ordinary [sic])", () => {
    const runs = readFragmentParagraphs(bareP(run("a ") + run("[sic]") + run(" b")))[0];
    const idx = omissionRunIndices(runs, [{ open: "[", close: "]", keyword: "Omitted" }]);
    expect(idx.size).toBe(0);
  });
});
