// Unit tests for the Condense & Shrink OOXML editor (src/core/ooxmlCondense.ts): the fragment reads,
// the shrink-size apply, the whitespace collapse, the paragraph merge + marker model, and Uncondense.
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
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

// ---------------------------------------------------------------------------
// Word's insertOoxml import normalization (proven via a real-Word COM round-trip, 2026-06-10):
// adjacent identically-formatted runs are COALESCED into one, and proofing marks can appear between
// a boundary glyph and its payload run. Uncondense must survive both, or breaks/pPr silently vanish
// on the live host while xmldom-only tests stay green.
// ---------------------------------------------------------------------------
describe("uncondense under Word's import normalization", () => {
  /** Simulate Word's run coalescing: adjacent sibling `<w:r>`s with identical rPr merge into one. */
  const mergeAdjacentIdenticalRuns = (xml: string): string => {
    /* eslint-disable @typescript-eslint/no-explicit-any */
    const doc = new DOMParser().parseFromString(xml, "text/xml") as any;
    const ser = new XMLSerializer();
    const paras = doc.getElementsByTagName("w:p");
    for (let i = 0; i < paras.length; i++) {
      const para = paras.item(i);
      const kids: any[] = [];
      for (let k = 0; k < para.childNodes.length; k++) kids.push(para.childNodes.item(k));
      let prev: { el: any; rPr: string } | null = null;
      for (const k of kids) {
        if (k.nodeType !== 1) continue;
        if (k.nodeName !== "w:r") {
          prev = null; // any non-run element breaks adjacency
          continue;
        }
        const rPrEl = k.getElementsByTagName("w:rPr").item(0);
        const rPr = rPrEl ? ser.serializeToString(rPrEl) : "";
        const prevT = prev ? prev.el.getElementsByTagName("w:t").item(0) : null;
        const curT = k.getElementsByTagName("w:t").item(0);
        if (prev && prev.rPr === rPr && prevT && curT) {
          prevT.appendChild(doc.createTextNode(curT.textContent || ""));
          para.removeChild(k);
        } else {
          prev = { el: k, rPr };
        }
      }
    }
    return ser.serializeToString(doc);
  };

  it("restores one break per glyph CHARACTER when Word coalesced adjacent space markers", () => {
    // A real blank paragraph between cards has no runs, so its two surrounding boundary glyphs end up
    // adjacent and Word merges them into ONE run with text "  " — which must still mean TWO breaks.
    const xml = body(p(run("AAA")), p(""), p(run("BBB")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(condensed.boundariesMarked).toBe(2);
    const wordized = mergeAdjacentIdenticalRuns(condensed.xml);
    expect(wordized).not.toBe(condensed.xml); // the simulation really coalesced the glyph pair
    const out = uncondenseFragmentOoxml(wordized);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "", "BBB"]); // the blank comes back
  });

  it("restores one break per pilcrow CHARACTER when Word coalesced adjacent ¶ markers", () => {
    const xml = body(p(run("AAA")), p(""), p(run("BBB")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: true, retainParagraphs: false, reversal: "marker" });
    const wordized = mergeAdjacentIdenticalRuns(condensed.xml);
    expect(wordized).not.toBe(condensed.xml);
    const out = uncondenseFragmentOoxml(wordized);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "", "BBB"]);
  });

  it("consumes a pPr payload separated from its glyph by Word proofing noise", () => {
    const xml = body(p(run("AAA")), p(run("BBB"), `<w:pPr><w:jc w:val="center"/></w:pPr>`));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    // Inject a proofing mark between the boundary glyph and its payload run (Word does this on import).
    const payloadStart = `<w:r><w:rPr><w:rStyle w:val="${CONDENSE_MARK_STYLE}"/><w:vanish/></w:rPr>`;
    expect(condensed.xml).toContain(payloadStart);
    const noisy = condensed.xml.replace(payloadStart, `<w:proofErr w:type="spellStart"/>${payloadStart}`);
    const out = uncondenseFragmentOoxml(noisy);
    expect(out.breaksRestored).toBe(1);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB"]);
    expect(out.xml).toContain(`<w:jc w:val="center"/>`); // the stored pPr was consumed, not dropped
    expect(out.xml).not.toContain("&lt;w:pPr"); // no leaked payload text
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

  // -------------------------------------------------------------------------
  // Marker-style durability (the 2026-06-09 live bug: Word strips `<w:rStyle>` references whose
  // style is not DEFINED in the package, so every condense was irreversible on the host — "no
  // Rostrum condense markers found" — while xmldom-based tests stayed green. The fix defines the
  // style in the fragment's own styles part whenever styled markers are written.)
  // -------------------------------------------------------------------------
  describe("marker style definition (survives Word's dangling-style strip)", () => {
    const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
    /** A flat-OPC fragment shaped like `range.getOoxml()` output: document part + styles part. */
    const pkgFragment = (bodyXml: string, stylesInner = ""): string =>
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}">${bodyXml}</w:document>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData>` +
      `<w:styles xmlns:w="${W_NS}">${stylesInner}</w:styles>` +
      `</pkg:xmlData></pkg:part>` +
      `</pkg:package>`;
    /** Count of style DEFINITIONS (w:styleId, not the markers' w:val references). */
    const defCount = (xml: string): number =>
      (xml.match(new RegExp(`w:styleId="${CONDENSE_MARK_STYLE}"`, "g")) || []).length;

    it("merge mode defines the marker style as a hidden character style in the styles part", () => {
      const xml = pkgFragment(body(p(run("AAA")), p(run("BBB"))));
      const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
      // Defined exactly once, INSIDE the styles part, with what Word needs to import + keep it.
      expect(defCount(out.xml)).toBe(1);
      const stylesPart = /<pkg:part pkg:name="\/word\/styles\.xml">[\s\S]*?<\/pkg:part>/.exec(out.xml);
      expect(stylesPart).not.toBeNull();
      expect(stylesPart![0]).toContain(`w:styleId="${CONDENSE_MARK_STYLE}"`);
      expect(stylesPart![0]).toContain(`w:type="character"`);
      expect(stylesPart![0]).toContain(`w:customStyle="1"`);
      expect(stylesPart![0]).toContain("<w:semiHidden/>");
      // The round-trip still restores.
      expect(uncondenseFragmentOoxml(out.xml).breaksRestored).toBe(1);
    });

    it("retain mode (marker) defines it too — dropped blanks key on the same style", () => {
      const xml = pkgFragment(body(p(run("AAA")), p(run("   ")), p(run("BBB"))));
      const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
      expect(out.boundariesMarked).toBe(1);
      expect(defCount(out.xml)).toBe(1);
    });

    it("is idempotent — a doc that already carries the definition gains no duplicate", () => {
      const already =
        `<w:style w:type="character" w:customStyle="1" w:styleId="${CONDENSE_MARK_STYLE}">` +
        `<w:name w:val="Rostrum Condense Break"/><w:semiHidden/></w:style>`;
      const xml = pkgFragment(body(p(run("AAA")), p(run("BBB"))), already);
      const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
      expect(defCount(out.xml)).toBe(1);
    });

    it("destructive mode writes no styled markers, so no definition is injected", () => {
      const xml = pkgFragment(body(p(run("AAA")), p(run("   ")), p(run("BBB"))));
      const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "none" });
      expect(defCount(out.xml)).toBe(0);
    });

    it("a bare fragment with no styles part is left structurally unchanged (live packages always bundle one)", () => {
      const out = condenseFragmentOoxml(body(p(run("AAA")), p(run("BBB"))), {
        usePilcrows: false,
        retainParagraphs: false,
        reversal: "marker"
      });
      expect(out.xml).not.toContain("<w:styles");
      expect(uncondenseFragmentOoxml(out.xml).breaksRestored).toBe(1);
    });
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
