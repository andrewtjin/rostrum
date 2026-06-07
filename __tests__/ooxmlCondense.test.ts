// Unit tests for the Condense & Shrink OOXML editor (src/core/ooxmlCondense.ts): the fragment reads,
// the shrink-size apply, the whitespace collapse, the paragraph merge + marker model, and Uncondense.
import {
  applyFragmentShrink,
  condenseFragmentOoxml,
  readFragmentParagraphs,
  resolveNormalSizeHalfPts,
  uncondenseFragmentOoxml
} from "../src/core/ooxmlCondense";
import { CONDENSE_MARK_STYLE } from "../src/core/styles";

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

const bareP = (inner: string, pPr = ""): string => `<w:p xmlns:w="${W_NS}">${pPr}${inner}</w:p>`;
const body = (...ps: string[]): string => `<w:body xmlns:w="${W_NS}">${ps.join("")}</w:body>`;
const p = (inner: string, pPr = ""): string => `<w:p>${pPr}${inner}</w:p>`;

/** Concatenated text per paragraph in a fragment. */
const paraTexts = (xml: string): string[] =>
  readFragmentParagraphs(xml).map((runs) => runs.map((r) => r.text).join(""));

describe("readFragmentParagraphs", () => {
  it("reads underline, explicit size, cite, and break-marker flags per run", () => {
    const xml = bareP(
      run("plain") +
        run("cut", { u: true }) +
        run("small", { sz: 16 }) +
        run("cited", { cite: true })
    );
    const [runs] = readFragmentParagraphs(xml);
    expect(runs).toHaveLength(4);
    expect(runs[0]).toMatchObject({ text: "plain", underline: false, sizeHalfPts: null, breakMarker: false });
    expect(runs[1]).toMatchObject({ text: "cut", underline: true });
    expect(runs[2]).toMatchObject({ text: "small", sizeHalfPts: 16 });
    expect(runs[3]).toMatchObject({ citeStyled: true });
  });

  it("treats underline none/0/false as not underlined", () => {
    const xml = bareP(`<w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>a</w:t></w:r>` + run("b"));
    const [runs] = readFragmentParagraphs(xml);
    expect(runs[0].underline).toBe(false);
    expect(runs[1].underline).toBe(false);
  });

  it("walks EVERY paragraph in a multi-paragraph fragment (not just the first)", () => {
    const xml = body(p(run("one")), p(run("two")), p(run("three")));
    const paras = readFragmentParagraphs(xml);
    expect(paras).toHaveLength(3);
    expect(paraTexts(xml)).toEqual(["one", "two", "three"]);
  });

  it("flags a condense break-marker run", () => {
    const marker = `<w:r><w:rPr><w:rStyle w:val="${CONDENSE_MARK_STYLE}"/></w:rPr><w:t> </w:t></w:r>`;
    const [runs] = readFragmentParagraphs(bareP(run("a") + marker + run("b")));
    expect(runs.map((r) => r.breakMarker)).toEqual([false, true, false]);
  });
});

describe("resolveNormalSizeHalfPts", () => {
  it("reads docDefaults rPrDefault size", () => {
    const styles = `<w:styles><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults></w:styles>`;
    expect(resolveNormalSizeHalfPts(styles)).toBe(22);
  });
  it("falls back to the Normal style size", () => {
    const styles = `<w:styles><w:style w:type="paragraph" w:styleId="Normal"><w:rPr><w:sz w:val="20"/></w:rPr></w:style></w:styles>`;
    expect(resolveNormalSizeHalfPts(styles)).toBe(20);
  });
  it("returns null when no styles part is present", () => {
    expect(resolveNormalSizeHalfPts(bareP(run("x")))).toBeNull();
  });
});

describe("applyFragmentShrink", () => {
  it("sets <w:sz> AND <w:szCs> on the targeted run and reports a change", () => {
    const xml = bareP(run("a") + run("b"));
    const { xml: out, changed } = applyFragmentShrink(xml, [{ runSizes: [16, undefined] }]);
    expect(changed).toBe(true);
    expect(out).toContain('<w:sz w:val="16"/>');
    expect(out).toContain('<w:szCs w:val="16"/>');
    const [runs] = readFragmentParagraphs(out);
    expect(runs[0].sizeHalfPts).toBe(16);
    expect(runs[1].sizeHalfPts).toBeNull(); // undefined left it untouched
  });

  it("clears an existing size when given null", () => {
    const xml = bareP(run("a", { sz: 16 }));
    const { xml: out } = applyFragmentShrink(xml, [{ runSizes: [null] }]);
    expect(readFragmentParagraphs(out)[0][0].sizeHalfPts).toBeNull();
  });

  it("inserts <w:sz> in schema order before <w:highlight>", () => {
    const xml = bareP(run("a", { highlight: "yellow" }));
    const { xml: out } = applyFragmentShrink(xml, [{ runSizes: [14] }]);
    // sz must precede highlight in the rPr (CT_RPr order), or the host can reject the OOXML.
    expect(out.indexOf("<w:sz ")).toBeLessThan(out.indexOf("<w:highlight"));
  });

  it("sets the paragraph-mark size when asked (Shrink ¶)", () => {
    const xml = bareP(run("a"));
    const { xml: out } = applyFragmentShrink(xml, [{ runSizes: [undefined], markSizeHalfPts: 12 }]);
    expect(out).toMatch(/<w:pPr>[\s\S]*<w:rPr>[\s\S]*<w:sz w:val="12"\/>/);
  });

  it("is a no-op (changed=false) when every run is left undefined", () => {
    const xml = bareP(run("a") + run("b"));
    expect(applyFragmentShrink(xml, [{ runSizes: [undefined, undefined] }]).changed).toBe(false);
  });
});

describe("condenseFragmentOoxml — whitespace collapse", () => {
  it("collapses double spaces, tabs, and breaks to one space", () => {
    const xml = body(p(`<w:r><w:t xml:space="preserve">a  b</w:t><w:tab/><w:t>c</w:t><w:br/><w:t>d</w:t></w:r>`));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(paraTexts(out.xml)).toEqual(["a b c d"]);
  });

  it("collapses a space that spans two runs", () => {
    const xml = body(p(run("a ") + run(" b")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(paraTexts(out.xml)).toEqual(["a b"]);
  });
});

describe("condenseFragmentOoxml — merge + markers", () => {
  it("merges N paragraphs into one with N-1 break markers (space glyph), changed=true", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(out.changed).toBe(true);
    expect(out.boundariesMarked).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA BBB CCC"]); // one paragraph, space markers between
    // The markers carry the break style.
    const [runs] = readFragmentParagraphs(out.xml);
    expect(runs.filter((r) => r.breakMarker)).toHaveLength(2);
  });

  it("uses a visible 6pt pilcrow glyph in pilcrow mode", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: true, retainParagraphs: false, reversal: "marker" });
    expect(paraTexts(out.xml)).toEqual(["AAA¶BBB"]);
    expect(out.xml).toContain('<w:sz w:val="12"/>'); // 6pt pilcrow
  });

  it("destructive merge (reversal none, pilcrows off) leaves a plain space and NO marker style", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "none" });
    expect(paraTexts(out.xml)).toEqual(["AAA BBB"]);
    expect(out.xml).not.toContain(CONDENSE_MARK_STYLE); // no reversible marker
  });
});

describe("uncondenseFragmentOoxml", () => {
  it("splits a merged paragraph back at its markers", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    const out = uncondenseFragmentOoxml(condensed.xml);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB", "CCC"]);
  });

  it("is a no-op on a fragment with no markers", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    expect(uncondenseFragmentOoxml(xml).changed).toBe(false);
  });
});

describe("retain-paragraphs mode", () => {
  it("lossless: hides a blank paragraph's mark with the break style (kept for reversal)", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(out.boundariesMarked).toBe(1); // the one blank paragraph
    expect(readFragmentParagraphs(out.xml)).toHaveLength(3); // structure retained
    expect(out.xml).toContain(CONDENSE_MARK_STYLE);
    // Uncondense un-hides it.
    const restored = uncondenseFragmentOoxml(out.xml);
    expect(restored.xml).not.toContain(CONDENSE_MARK_STYLE);
  });

  it("destructive: actually removes blank paragraphs", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "none" });
    expect(readFragmentParagraphs(out.xml)).toHaveLength(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB"]);
  });

  it("losslessly condenses a blank paragraph whose mark has a foreign style, restoring it on uncondense", () => {
    // A blank paragraph carrying a non-Rostrum mark style used to be SKIPPED (the original style collides
    // with our break style in the single rStyle slot). Now we park the pristine mark rPr in a hidden
    // payload and swap in our break style, so the blank IS condensed and the user's style round-trips.
    const styledBlank = `<w:p><w:pPr><w:rPr><w:rStyle w:val="SomeOtherStyle"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">  </w:t></w:r></w:p>`;
    const xml = body(p(run("AAA")), styledBlank, p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(out.boundariesMarked).toBe(1); // the foreign-styled blank IS now condensed
    expect(out.xml).toContain("<w:vanish/>"); // hidden
    expect(out.xml).toContain("SomeOtherStyle"); // original mark style preserved (parked in the payload)
    expect(readFragmentParagraphs(out.xml)).toHaveLength(3); // structure retained

    const restored = uncondenseFragmentOoxml(out.xml);
    expect(restored.xml).not.toContain(CONDENSE_MARK_STYLE); // our break style + payload gone
    expect(restored.xml).toContain(`<w:rStyle w:val="SomeOtherStyle"/>`); // user's mark style restored exactly
    expect(restored.xml).not.toContain("<w:vanish/>"); // un-hidden
    expect(readFragmentParagraphs(restored.xml)).toHaveLength(3);
  });

  it("condenses an underlined-but-empty newline whose mark is styled via a char style (the reported bug)", () => {
    // Repro: a newline whose paragraph mark is underlined via a character style, with NO text. It must
    // collapse under retain-paragraphs mode — it didn't before, because the foreign mark style made us
    // skip it. The underline char style must come back on Uncondense.
    const underlinedNewline = `<w:p><w:pPr><w:rPr><w:rStyle w:val="StyleUnderline"/></w:rPr></w:pPr></w:p>`;
    const xml = body(p(run("AAA")), underlinedNewline, p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(out.boundariesMarked).toBe(1); // the underlined empty newline is condensed
    expect(out.xml).toContain("<w:vanish/>");

    const restored = uncondenseFragmentOoxml(out.xml);
    expect(restored.xml).toContain(`<w:rStyle w:val="StyleUnderline"/>`);
    expect(restored.xml).not.toContain(CONDENSE_MARK_STYLE);
  });
});
