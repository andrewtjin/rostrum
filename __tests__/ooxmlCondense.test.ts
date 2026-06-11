// Unit tests for the Condense & Shrink OOXML editor (src/core/ooxmlCondense.ts): the fragment reads,
// the shrink-size apply, the whitespace collapse, the paragraph merge + marker model, and Uncondense.
//
// Marker identity is the intrinsic zero-width TEXT signature `MARK_SIGNATURE` (U+200B ZWSP +
// U+2060 WORD JOINER), NOT a custom character STYLE. The retired style approach depended on a net-new
// `<w:rStyle>` surviving `insertOoxml` into a populated doc — which it does NOT (Word drops a reference
// whose styleId is not resident), erasing the only signal Uncondense keyed on while xmldom/COM gates
// stayed green. These tests model that host normalization directly (see "Word's insertOoxml import
// normalization" below).
//
// The WRITTEN form moved twice on 2026-06-10: U+2063 alone proved FONT-DEPENDENT in live Word (a
// visible comma-like mark — the boundary glyph run is deliberately NOT vanished), and the interim
// single U+200B was REFUTED by an adversarial repro: organic ZWSPs are endemic in web-pasted text, so
// bare containment + the followed-by-{space,¶} guard fabricated breaks (eating a space) and misflagged
// organic-ZWSP runs as `breakMarker` (Shrink-exempt, collapse-skipped). Hence the two-code-unit pair —
// see the "organic ZWSP" describe below. READ paths accept `MARK_SIGNATURES` = [pair, legacy U+2063]
// ONLY (single U+200B never shipped, so it has no back-compat value); the "legacy signature" tests pin
// that window.
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import {
  applyFragmentShrink,
  condenseFragmentOoxml,
  readFragmentParagraphs,
  resolveNormalSizeHalfPts,
  uncondenseFragmentOoxml
} from "../src/core/ooxmlCondense";
import { LEGACY_MARK_SIGNATURE, MARK_SIGNATURE } from "../src/core/styles";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const S = MARK_SIGNATURE; // the zero-width boundary signature (ZWSP+WJ), for terse fixtures/assertions
const L = LEGACY_MARK_SIGNATURE; // legacy U+2063 (read-only back-compat; never written anymore)
const ZWSP = String.fromCodePoint(0x200b); // a LONE organic zero-width space — real user content, NOT a marker

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

/** Concatenated text per paragraph in a fragment (signatures are filtered from the view by design). */
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

  it("flags a condense marker run by its intrinsic text signature (and hides the signature from the view)", () => {
    const marker = `<w:r><w:t xml:space="preserve">${S} </w:t></w:r>`;
    const [runs] = readFragmentParagraphs(bareP(run("a") + marker + run("b")));
    expect(runs.map((r) => r.breakMarker)).toEqual([false, true, false]);
    expect(runs[1].text).toBe(" "); // signature filtered out of the visible text
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
  it("merges N paragraphs into one with N-1 signature markers (space glyph), changed=true", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(out.changed).toBe(true);
    expect(out.boundariesMarked).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA BBB CCC"]); // one paragraph, space markers between
    // The markers are signature-tagged runs — and carry NO custom character style (the retired signal).
    const [runs] = readFragmentParagraphs(out.xml);
    expect(runs.filter((r) => r.breakMarker)).toHaveLength(2);
    expect(out.xml).not.toContain("w:rStyle"); // no style reference anywhere
    expect(out.xml).toContain(S); // the intrinsic signal is present
    expect(out.xml).not.toContain(L); // WRITE paths emit only the current signature, never the legacy one
  });

  it("uses a visible 6pt pilcrow glyph in pilcrow mode", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: true, retainParagraphs: false, reversal: "marker" });
    expect(paraTexts(out.xml)).toEqual(["AAA¶BBB"]);
    expect(out.xml).toContain('<w:sz w:val="12"/>'); // 6pt pilcrow (direct formatting, host-durable)
  });

  it("destructive merge (reversal none, pilcrows off) leaves a plain space and NO signature", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "none" });
    expect(paraTexts(out.xml)).toEqual(["AAA BBB"]);
    expect(out.xml).not.toContain(S); // no reversible marker
  });

  it("preserves a divergent following-paragraph pPr in a hidden payload and restores it", () => {
    const xml = body(p(run("AAA")), p(run("BBB"), `<w:pPr><w:jc w:val="center"/></w:pPr>`));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(condensed.xml).toContain("<w:vanish/>"); // payload run is hidden
    const restored = uncondenseFragmentOoxml(condensed.xml);
    expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB"]);
    expect(restored.xml).toContain(`<w:jc w:val="center"/>`); // divergent pPr restored
    expect(restored.xml).not.toContain(S); // no leaked signal
  });
});

describe("uncondenseFragmentOoxml", () => {
  it("splits a merged paragraph back at its markers", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    const out = uncondenseFragmentOoxml(condensed.xml);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB", "CCC"]);
    expect(out.xml).not.toContain(S); // every signature consumed
  });

  it("is a no-op on a fragment with no markers", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    expect(uncondenseFragmentOoxml(xml).changed).toBe(false);
  });

  it("round-trips full-merge, pilcrow, and retain modes back to the original text", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    for (const opts of [
      { usePilcrows: false, retainParagraphs: false, reversal: "marker" as const },
      { usePilcrows: true, retainParagraphs: false, reversal: "marker" as const }
    ]) {
      const condensed = condenseFragmentOoxml(xml, opts);
      const restored = uncondenseFragmentOoxml(condensed.xml);
      expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB", "CCC"]);
      expect(restored.xml).not.toContain(S);
    }
  });
});

// ---------------------------------------------------------------------------
// Word's insertOoxml import-normalization SIMULATOR (host behavior proven via a real-Word COM
// round-trip, 2026-06-10):
//   (a) adjacent identically-formatted runs are COALESCED into one;
//   (b) a `<w:rStyle>` whose styleId is NOT resident in the destination is DROPPED on import;
//   (c) proofing marks can appear between a boundary glyph and its payload run.
// Module-scoped because BOTH the R2 gate and the legacy back-compat suite need the true post-host
// shapes — xmldom alone never coalesces, so a fixture that skips this simulator silently tests a
// pre-host shape that the live pipeline never produces.
// ---------------------------------------------------------------------------
/* eslint-disable @typescript-eslint/no-explicit-any */
/** Simulate Word's run coalescing: adjacent sibling `<w:r>`s with identical rPr merge into one. */
const mergeAdjacentIdenticalRuns = (xml: string): string => {
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

/** Simulate Word dropping any `<w:rStyle>` whose styleId is not resident in the destination styles. */
const dropNonResidentRStyle = (xml: string, resident: Set<string> = new Set()): string => {
  const doc = new DOMParser().parseFromString(xml, "text/xml") as any;
  const refs = doc.getElementsByTagName("w:rStyle");
  const toRemove: any[] = [];
  for (let i = 0; i < refs.length; i++) {
    const el = refs.item(i);
    if (!resident.has(el.getAttribute("w:val") || "")) toRemove.push(el);
  }
  for (const el of toRemove) if (el.parentNode) el.parentNode.removeChild(el);
  return new XMLSerializer().serializeToString(doc);
};

/** The faithful host pipeline: drop non-resident style refs, THEN coalesce identical runs. */
const insertOoxmlFaithful = (xml: string, resident?: Set<string>): string =>
  mergeAdjacentIdenticalRuns(dropNonResidentRStyle(xml, resident));
/* eslint-enable @typescript-eslint/no-explicit-any */

// This block is the headless REGRESSION GATE (R2) the diagnosis calls for: it encodes host behavior
// (b) that xmldom/COM over-preserved — the exact gap that let the 2026-06-09 bug ship. (R2 cannot
// model whether the signature chars themselves survive `insertOoxml`; the COM harness proved the
// ZWSP+WJ pair and legacy U+2063 both do, and only a live wet-test (R1) settles the WebView host.)
describe("uncondense under Word's insertOoxml import normalization (R2)", () => {
  it("CONTROL: the simulator really strips a non-resident <w:rStyle> (the retired signal would vanish)", () => {
    // A hand-built OLD-style marker (keyed on `rStyle`, no signature) with no resident style definition.
    const oldStyleMarker = body(
      p(run("AAA") + `<w:r><w:rPr><w:rStyle w:val="RostrumCondenseBreak"/></w:rPr><w:t xml:space="preserve"> </w:t></w:r>` + run("BBB"))
    );
    const wordized = insertOoxmlFaithful(oldStyleMarker);
    expect(wordized).not.toContain("RostrumCondenseBreak"); // the only old signal is gone
    // Uncondense (which now keys on the signature) finds NOTHING — exactly the live 0-marker failure.
    expect(uncondenseFragmentOoxml(wordized).breaksRestored).toBe(0);
  });

  it("FIX A survives the same pipeline: full-merge restores every break", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    const wordized = insertOoxmlFaithful(condensed.xml); // no resident styles at all
    const out = uncondenseFragmentOoxml(wordized);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB", "CCC"]);
  });

  it("restores a break per SIGNATURE occurrence even when Word coalesced body INTO the marker run", () => {
    // No-rPr body + no-rPr space markers all coalesce into ONE run: "AAA<sig> <sig> BBB". The
    // per-occurrence tokenizer must recover AAA / (blank) / BBB rather than dropping the coalesced run.
    const xml = body(p(run("AAA")), p(""), p(run("BBB")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(condensed.boundariesMarked).toBe(2);
    const wordized = insertOoxmlFaithful(condensed.xml);
    expect(wordized).not.toBe(condensed.xml); // coalescing really happened
    const out = uncondenseFragmentOoxml(wordized);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "", "BBB"]); // the blank between cards comes back
    expect(out.xml).not.toContain(S);
  });

  it("restores one break per pilcrow when Word coalesced adjacent ¶ markers", () => {
    const xml = body(p(run("AAA")), p(""), p(run("BBB")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: true, retainParagraphs: false, reversal: "marker" });
    const wordized = insertOoxmlFaithful(condensed.xml);
    const out = uncondenseFragmentOoxml(wordized);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "", "BBB"]);
  });

  it("consumes a pPr payload separated from its glyph by Word proofing noise", () => {
    const xml = body(p(run("AAA")), p(run("BBB"), `<w:pPr><w:jc w:val="center"/></w:pPr>`));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    // The payload run is vanished and self-describing; inject a proofing mark before it (Word does this).
    const payloadStart = `<w:r><w:rPr><w:vanish/></w:rPr>`;
    expect(condensed.xml).toContain(payloadStart);
    const noisy = condensed.xml.replace(payloadStart, `<w:proofErr w:type="spellStart"/>${payloadStart}`);
    const out = uncondenseFragmentOoxml(noisy);
    expect(out.breaksRestored).toBe(1);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB"]);
    expect(out.xml).toContain(`<w:jc w:val="center"/>`); // the stored pPr was consumed, not dropped
    expect(out.xml).not.toContain("&lt;w:pPr"); // no leaked payload text
  });
});

// ---------------------------------------------------------------------------
// LEGACY signature back-compat (single U+2063, deployed 2026-06-10 only). Docs condensed by that
// build carry U+2063 markers; the READ paths must honor them forever or those docs become
// irreversible. Fixtures are built exactly the way such a doc came to exist: condense fresh (writes
// the ZWSP+WJ pair), then string-replace the signature with U+2063 — byte-identical to the deployed
// build's output. (A bare single U+200B is NOT in the read set: it never shipped, and accepting it
// would re-open the organic-ZWSP false positives — see the "organic ZWSP" suite below.)
// ---------------------------------------------------------------------------
describe("legacy U+2063 signature back-compat", () => {
  /** A condensed fragment as the DEPLOYED 2026-06-10 build would have produced it. */
  const legacyize = (xml: string): string => xml.split(S).join(L);

  it("uncondenses a legacy-condensed doc IDENTICALLY to a current one, in all three modes", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB"), `<w:pPr><w:jc w:val="center"/></w:pPr>`));
    for (const opts of [
      { usePilcrows: false, retainParagraphs: false, reversal: "marker" as const }, // merge
      { usePilcrows: true, retainParagraphs: false, reversal: "marker" as const }, // pilcrow
      { usePilcrows: false, retainParagraphs: true, reversal: "marker" as const } // retain
    ]) {
      const condensed = condenseFragmentOoxml(xml, opts);
      const legacy = legacyize(condensed.xml);
      expect(legacy).not.toBe(condensed.xml); // the swap really happened (signatures were present)
      const fromLegacy = uncondenseFragmentOoxml(legacy);
      const fromCurrent = uncondenseFragmentOoxml(condensed.xml);
      // Identical restoration: same break count, byte-identical XML (boundaries, pPr, mark rPr — and
      // every character of body text — all back; the signature kind is fully consumed either way).
      expect(fromLegacy.breaksRestored).toBe(fromCurrent.breaksRestored);
      expect(fromLegacy.xml).toBe(fromCurrent.xml);
      expect(fromLegacy.xml).not.toContain(L);
      expect(fromLegacy.xml).not.toContain(S);
    }
  });

  it("restores breaks from a legacy doc even after Word coalesced adjacent legacy markers into ONE run", () => {
    // The per-occurrence split (two boundary signatures in ONE run) must work for the legacy char too —
    // and the fixture must BE one run: xmldom never coalesces, so pipe Word's serialization shape
    // through the faithful host simulator instead of hand-waving "they would coalesce".
    const xml = body(p(run("AAA")), p(""), p(run("BBB")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    const wordized = insertOoxmlFaithful(legacyize(condensed.xml)); // legacy doc, then host coalescing
    expect(wordized).toContain(`AAA${L} ${L} BBB`); // PROOF: all four runs truly merged into one text
    const out = uncondenseFragmentOoxml(wordized);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "", "BBB"]);
  });

  it("splits a MIXED run carrying BOTH signature kinds (re-condensed legacy doc, then Word-coalesced)", () => {
    // A legacy boundary (U+2063) and a new boundary (ZWSP+WJ) coalesced into ONE run — the shape a
    // re-condense of a partially-uncondensed legacy doc can produce. Both must split.
    const mixed = body(p(`<w:r><w:t xml:space="preserve">AAA${L} BBB${S} CCC</w:t></w:r>`));
    const out = uncondenseFragmentOoxml(mixed);
    expect(out.breaksRestored).toBe(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB", "CCC"]);
    expect(out.xml).not.toContain(L);
    expect(out.xml).not.toContain(S);
  });

  it("keeps a stray signature of EITHER kind verbatim when not followed by a boundary glyph", () => {
    // The no-glyph guard must be lossless per kind: a stray legacy char stays U+2063 (never normalized
    // to the current pair) and a stray current pair stays ZWSP+WJ — no character lost or rewritten.
    const mixed = body(p(`<w:r><w:t xml:space="preserve">AAA${L}X${S} BBB${S}Y</w:t></w:r>`));
    const out = uncondenseFragmentOoxml(mixed);
    expect(out.breaksRestored).toBe(1); // only the one real boundary (signature + space)
    // Strays survive in run TEXT (the split may re-emit them in their own run); assert each kept its
    // ORIGINAL kind in the raw XML — the view (paraTexts) filters signatures, proving no visible loss.
    expect(out.xml).toContain(`${L}X`); // legacy stray kept verbatim as U+2063
    expect(out.xml).toContain(`${S}Y`); // current stray kept verbatim as the ZWSP+WJ pair
    expect(paraTexts(out.xml)).toEqual(["AAAX", "BBBY"]); // every visible character intact
  });
});

// ---------------------------------------------------------------------------
// ORGANIC ZWSP safety (the executed adversarial repro that killed the single-U+200B design,
// 2026-06-10). Web-pasted debate text routinely carries lone U+200B characters; none of them may be
// mistaken for a marker. These tests are the repro turned regression gate: marker semantics must key
// on the FULL ZWSP+WJ pair, so untouched user text round-trips identically, organic-ZWSP runs stay
// shrinkable/collapsible, and a ZWSP-only paragraph is real content (NOT blank).
// ---------------------------------------------------------------------------
describe("organic U+200B in user text (never a marker)", () => {
  it("uncondense is a NO-OP on a pristine selection containing organic ZWSP+space", () => {
    // The killer shape: "word<ZWSP> word" — under single-char detection this fabricated a paragraph
    // break AND ate the space (one button press corrupts an untouched selection).
    const xml = body(p(run(`Russia strikes first${ZWSP} and escalates`)));
    const out = uncondenseFragmentOoxml(xml);
    expect(out.changed).toBe(false); // nothing to do — no marker signature anywhere
    expect(out.breaksRestored).toBe(0);
    expect(paraTexts(out.xml)).toEqual([`Russia strikes first${ZWSP} and escalates`]); // ZWSP intact
  });

  it("condense -> uncondense round-trips text identically when a card contains organic ZWSP+space", () => {
    const xml = body(p(run(`Deterrence fails${ZWSP} when signaling collapses`)), p(run("BBB second card")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(condensed.boundariesMarked).toBe(1);
    // Through the faithful host pipeline too — organic ZWSP can coalesce into a marker run's text.
    const out = uncondenseFragmentOoxml(insertOoxmlFaithful(condensed.xml));
    expect(out.breaksRestored).toBe(1); // exactly the one real boundary, no fabricated break
    expect(paraTexts(out.xml)).toEqual([`Deterrence fails${ZWSP} when signaling collapses`, "BBB second card"]);
  });

  it("a run with an organic ZWSP is NOT a breakMarker — it still shrinks and whitespace-collapses", () => {
    const xml = body(p(run(`pasted${ZWSP} evidence   with   gaps`)));
    // Not a marker: Shrink's keepFullSize(breakMarker) exemption must not trigger...
    const [runs] = readFragmentParagraphs(xml);
    expect(runs[0].breakMarker).toBe(false);
    // ...so the run is sizeable like any other text run...
    const { xml: shrunk } = applyFragmentShrink(xml, [{ runSizes: [16] }]);
    expect(readFragmentParagraphs(shrunk)[0][0].sizeHalfPts).toBe(16);
    // ...and the collapse pass no longer skips it (markers are exempt from collapse by design).
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(paraTexts(out.xml)).toEqual([`pasted${ZWSP} evidence with gaps`]); // gaps collapsed, ZWSP kept
  });

  it("a ZWSP-only paragraph is real content, NOT a blank to drop in retain mode", () => {
    // Under single-char detection the view stripped the lone ZWSP, so the paragraph read as blank and
    // destructive retain DELETED user content. The signature view keeps a lone ZWSP.
    const xml = body(p(run("AAA")), p(run(ZWSP)), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "none" });
    expect(out.boundariesMarked).toBe(0); // nothing dropped
    expect(readFragmentParagraphs(out.xml)).toHaveLength(3); // all three paragraphs survive
    expect(paraTexts(out.xml)).toEqual(["AAA", ZWSP, "BBB"]);
  });
});

// ---------------------------------------------------------------------------
// Glyph-marker fail-safe: the split rebuilds a marker run from `runTextRaw`, which concatenates the
// `<w:t>`s ACROSS intervening elements and drops them. A run whose signature only exists in that
// concatenation (e.g. ZWSP at the end of one <w:t>, WJ at the start of the next, a <w:tab/> between)
// must therefore FAIL glyph-marker detection and be left verbatim — losslessness outranks marker
// detection. (Pre-existing class: the same concat-and-drop corrupted on the deployed U+2063 build.)
// ---------------------------------------------------------------------------
describe("glyph-marker detection requires a text-only run (fail-safe)", () => {
  it("leaves a run verbatim when its signature spans <w:t>s around a tab — no fabricated break, tab kept", () => {
    // The executed reviewer fixture: raw-text concat is "AAA" + ZWSP + WJ + " BBB" (a full signature
    // followed by a space), but the tab BETWEEN the <w:t>s would be dropped by the rebuild.
    const WJ = S.slice(1); // U+2060 WORD JOINER (second half of the pair)
    const xml = body(
      p(`<w:r><w:t xml:space="preserve">AAA${ZWSP}</w:t><w:tab/><w:t xml:space="preserve">${WJ} BBB</w:t></w:r>`)
    );
    const out = uncondenseFragmentOoxml(xml);
    expect(out.changed).toBe(false); // complete no-op — the run is not a glyph marker
    expect(out.breaksRestored).toBe(0); // no fabricated break
    expect(out.xml).toContain("<w:tab/>"); // the tab survives
    expect(out.xml).toContain(`AAA${ZWSP}`); // every text byte still in place, in its original <w:t>
    expect(out.xml).toContain(`${WJ} BBB`);
  });

  it("still detects every GENUINE glyph marker (condense only writes rPr+t runs) — no false negative", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    for (const opts of [
      { usePilcrows: false, retainParagraphs: false, reversal: "marker" as const }, // space markers, no rPr
      { usePilcrows: true, retainParagraphs: false, reversal: "marker" as const } // 6pt ¶ markers, WITH rPr
    ]) {
      const condensed = condenseFragmentOoxml(xml, opts);
      expect(condensed.boundariesMarked).toBe(2);
      const restored = uncondenseFragmentOoxml(condensed.xml);
      expect(restored.breaksRestored).toBe(2); // both markers detected and split
      expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB", "CCC"]);
    }
  });
});

describe("portability + idempotence", () => {
  it("a condensed block with NO styles part still uncondenses (copy/paste portability)", () => {
    // The signal travels in run TEXT, not a document-local style table — so a condensed block pasted
    // into a fresh doc (here: a bare fragment, no styles part at all) still reverses exactly.
    const condensed = condenseFragmentOoxml(body(p(run("AAA")), p(run("BBB"))), {
      usePilcrows: false,
      retainParagraphs: false,
      reversal: "marker"
    });
    expect(condensed.xml).not.toContain("<w:styles"); // nothing injected into a style table
    const restored = uncondenseFragmentOoxml(condensed.xml);
    expect(restored.breaksRestored).toBe(1);
    expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB"]);
  });

  it("the signature never leaks into restored visible text (byte-exact bodies)", () => {
    const xml = body(p(run("Alpha")), p(run("Beta")), p(run("Gamma")));
    const condensed = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    const restored = uncondenseFragmentOoxml(condensed.xml);
    expect(restored.xml).not.toContain(S);
    expect(restored.xml).not.toContain(L); // nor the legacy char
    expect(restored.xml).not.toContain("\\u200B");
    expect(restored.xml).not.toContain("\\u2063");
    expect(paraTexts(restored.xml)).toEqual(["Alpha", "Beta", "Gamma"]);
  });

  it("re-condense is idempotent — condensing an already-condensed fragment adds no new break", () => {
    const xml = body(p(run("AAA")), p(run("BBB")), p(run("CCC")));
    const once = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    const twice = condenseFragmentOoxml(once.xml, { usePilcrows: false, retainParagraphs: false, reversal: "marker" });
    expect(twice.boundariesMarked).toBe(0); // already one paragraph — nothing to merge
    const restored = uncondenseFragmentOoxml(twice.xml);
    expect(restored.breaksRestored).toBe(2);
    expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB", "CCC"]);
  });
});

describe("retain-paragraphs mode", () => {
  it("lossless: collapses a blank paragraph (signature-tagged, hidden) and structure is retained", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(out.boundariesMarked).toBe(1); // the one blank paragraph
    expect(readFragmentParagraphs(out.xml)).toHaveLength(3); // structure retained
    expect(out.xml).toContain(S);
    expect(out.xml).toContain("<w:vanish/>"); // the blank is hidden (collapsed)
    expect(out.xml).not.toContain("w:rStyle"); // no custom style reference
    // Uncondense restores it and removes every trace of the signal.
    const restored = uncondenseFragmentOoxml(out.xml);
    expect(restored.xml).not.toContain(S);
    expect(restored.xml).not.toContain("<w:vanish/>");
    expect(readFragmentParagraphs(restored.xml)).toHaveLength(3);
  });

  it("re-condense is idempotent in RETAIN mode (an already-dropped blank is not re-marked)", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB")));
    const opts = { usePilcrows: false, retainParagraphs: true, reversal: "marker" as const };
    const once = condenseFragmentOoxml(xml, opts);
    expect(once.boundariesMarked).toBe(1);
    const twice = condenseFragmentOoxml(once.xml, opts);
    expect(twice.boundariesMarked).toBe(0); // already dropped — nothing new to mark
    // Still exactly one parked payload (no double-park), and it round-trips.
    expect((twice.xml.match(/<w:vanish\/>/g) || []).length).toBe((once.xml.match(/<w:vanish\/>/g) || []).length);
    const restored = uncondenseFragmentOoxml(twice.xml);
    expect(restored.xml).not.toContain(S);
    expect(readFragmentParagraphs(restored.xml)).toHaveLength(3);
  });

  it("destructive: actually removes blank paragraphs", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "none" });
    expect(readFragmentParagraphs(out.xml)).toHaveLength(2);
    expect(paraTexts(out.xml)).toEqual(["AAA", "BBB"]);
  });

  it("destructive: never removes the final remaining paragraph (all-blank selection)", () => {
    // Pins the keep-at-least-one guard (a range must end with a paragraph) — the guard now uses a
    // live count instead of re-scanning the document per blank, and this is its only direct test.
    const xml = body(p(run("   ")), p(run(" ")), p(run("  ")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "none" });
    expect(readFragmentParagraphs(out.xml)).toHaveLength(1);
  });

  it("survives the faithful host pipeline (R2): a dropped blank still restores", () => {
    const xml = body(p(run("AAA")), p(run("   ")), p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    const doc = new DOMParser().parseFromString(out.xml, "text/xml") as any;
    // (no rStyle to drop; just exercise reversibility post-condense)
    const restored = uncondenseFragmentOoxml(new XMLSerializer().serializeToString(doc));
    expect(restored.xml).not.toContain(S);
    expect(readFragmentParagraphs(restored.xml)).toHaveLength(3);
  });

  it("losslessly condenses a blank whose mark has a FOREIGN char style, restoring it on uncondense", () => {
    // A blank paragraph carrying a non-Rostrum mark style: the original mark rPr is parked verbatim in a
    // hidden payload, the live mark is vanished to collapse, and uncondense restores the user's mark.
    const styledBlank = `<w:p><w:pPr><w:rPr><w:rStyle w:val="SomeOtherStyle"/></w:rPr></w:pPr><w:r><w:t xml:space="preserve">  </w:t></w:r></w:p>`;
    const xml = body(p(run("AAA")), styledBlank, p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(out.boundariesMarked).toBe(1);
    expect(out.xml).toContain("<w:vanish/>"); // hidden
    expect(out.xml).toContain("SomeOtherStyle"); // original mark style preserved (parked in the payload)
    expect(readFragmentParagraphs(out.xml)).toHaveLength(3);

    const restored = uncondenseFragmentOoxml(out.xml);
    expect(restored.xml).not.toContain(S); // signature + payload gone
    expect(restored.xml).toContain(`<w:rStyle w:val="SomeOtherStyle"/>`); // user's mark style restored exactly
    expect(restored.xml).not.toContain("<w:vanish/>"); // un-hidden
    expect(readFragmentParagraphs(restored.xml)).toHaveLength(3);
  });

  it("condenses an underlined-but-empty newline whose mark is styled via a char style (the reported bug)", () => {
    const underlinedNewline = `<w:p><w:pPr><w:rPr><w:rStyle w:val="StyleUnderline"/></w:rPr></w:pPr></w:p>`;
    const xml = body(p(run("AAA")), underlinedNewline, p(run("BBB")));
    const out = condenseFragmentOoxml(xml, { usePilcrows: false, retainParagraphs: true, reversal: "marker" });
    expect(out.boundariesMarked).toBe(1); // the underlined empty newline is condensed
    expect(out.xml).toContain("<w:vanish/>");

    const restored = uncondenseFragmentOoxml(out.xml);
    expect(restored.xml).toContain(`<w:rStyle w:val="StyleUnderline"/>`);
    expect(restored.xml).not.toContain(S);
  });
});
