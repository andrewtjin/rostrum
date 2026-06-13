// Loop 002 B1 — the node layer proof (ParsedParagraph.fromNode + applyVisibilityInPlace).
//
// THREE gates, all in this file (the new primitives' unit proofs; wiring is a LATER step):
//   1. A3 SIX-FIELD (+hasInternalPart) DIFFERENTIAL (002-S3). For every ALL_FIXTURES paragraph,
//      `fromNode`'s RunView[] must be IDENTICAL — all seven fields — to the legacy per-scan reader
//      (`legacyRunViewsForTest`, the standalone runText walk + seven getElementsByTagName scans +
//      individual rPr reads + internal-part probe). Proves the FUSED single traversal is faithful.
//   2. `applyVisibilityInPlace` LOSSLESSNESS (002-S1 / 002-F1). Serialize a representative paragraph
//      before/after a node-mode hide and assert the semantic-diff oracle accepts the delta (only
//      <w:vanish> toggles + whitelisted bridge-split insertions); assert IDEMPOTENCE (re-apply =
//      no change); cover fully-hidden, partial-hide-with-boundary-bridge, and interior-split.
//   3. parseCount===0 (002-S1 sub-gate, tests-4). With the SAME DOMParser spy parseCount.test.ts
//      uses, prove the node path (`fromNode` + `applyVisibilityInPlace`) triggers ZERO
//      `parseFromString` calls — the falsifiable "the win is real" check.
//
// The spy mock MUST be installed before importing anything that touches @xmldom/xmldom, exactly as
// parseCount.test.ts does (jest hoists jest.mock above imports, but we keep the same shape for clarity).

/* eslint-disable @typescript-eslint/no-explicit-any */

let parseCount = 0;
jest.mock("@xmldom/xmldom", () => {
  const actual = jest.requireActual("@xmldom/xmldom");
  class CountingDOMParser {
    private readonly inner: any;
    constructor(opts?: any) {
      this.inner = new actual.DOMParser(opts);
    }
    parseFromString(source: string, mimeType: string): any {
      parseCount++;
      return this.inner.parseFromString(source, mimeType);
    }
  }
  return { ...actual, DOMParser: CountingDOMParser };
});

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { ParsedParagraph, legacyRunViewsForTest, VisibilityPlan } from "../src/core/ooxml";
import { computeRunKeepFlags, planCrossGapSeparators } from "../src/core/keepers";
import { ALL_FIXTURES } from "./fixtures/engine";
import { assertVanishBridgeOnlyDeltaPara } from "./semanticDiff";
import { RunView } from "../src/core/types";

/** Parse a package string WITH the fatal-only policy the engine uses (so the spy counts these parses). */
function parsePkg(xml: string): any {
  return new DOMParser({
    onError: (level: string, message: string) => {
      if (level === "fatalError") throw new Error(message);
    }
  }).parseFromString(xml, "text/xml");
}

function serialize(doc: any): string {
  return new XMLSerializer().serializeToString(doc);
}

/** The body scope = first <w:body> (or the doc for a bare fragment) — matches ooxml.ts firstParagraph. */
function bodyScope(doc: any): any {
  const bodies = doc.getElementsByTagName("w:body");
  return bodies && bodies.length > 0 ? bodies.item(0) : doc;
}

/** True when a node lives inside a textbox (excluded story) — matches the oracle's isInTextbox. */
function isInTextbox(node: any): boolean {
  let n = node ? node.parentNode : null;
  while (n) {
    if (n.nodeName === "w:txbxContent") return true;
    n = n.parentNode;
  }
  return false;
}

/** Every story <w:p> in document order (textbox-nested excluded) — the paragraphs fromNode operates on. */
function storyParagraphs(doc: any): any[] {
  const live = bodyScope(doc).getElementsByTagName("w:p");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) {
    const p = live.item(i);
    if (p && !isInTextbox(p)) out.push(p);
  }
  return out;
}

/** Serialize the FIRST story <w:p> of a doc as a standalone fragment for the per-paragraph oracle. */
function firstParaXml(doc: any): string {
  const p = storyParagraphs(doc)[0];
  return serialize(p);
}

describe("ParsedParagraph.fromNode — A3 six-field (+hasInternalPart) differential (002-S3)", () => {
  // Over EVERY fixture, the fused node-mode read must equal the legacy per-scan read, field for field.
  for (const [name, pkg] of Object.entries(ALL_FIXTURES)) {
    it(`fromNode RunView[] equals the legacy per-scan reader on '${name}'`, () => {
      const doc = parsePkg(pkg);
      const paras = storyParagraphs(doc);
      expect(paras.length).toBeGreaterThan(0);
      for (const pEl of paras) {
        const fused = ParsedParagraph.fromNode(doc, pEl).runs;
        // Legacy reference reads the SAME paragraph as a standalone fragment (independent parse/scan).
        const legacy = legacyRunViewsForTest(serialize(pEl));
        expect(fused).toEqual(legacy);
        // Belt-and-suspenders: every one of the seven fields present and of the right kind.
        for (const rv of fused) {
          expect(typeof rv.index).toBe("number");
          expect(typeof rv.text).toBe("string");
          expect(rv.highlight === null || typeof rv.highlight === "string").toBe(true);
          expect(typeof rv.citeStyled).toBe("boolean");
          expect(typeof rv.underline).toBe("boolean");
          expect(typeof rv.hidden).toBe("boolean");
          expect(typeof rv.eligible).toBe("boolean");
          expect(typeof rv.hasInternalPart).toBe("boolean");
        }
      }
    });
  }

  it("classifies eligibility + internal parts correctly on the named hard shapes", () => {
    // mediaDrawing: run[0] has a <w:drawing> → ineligible AND hasInternalPart; run[1] is plain caption.
    {
      const doc = parsePkg(ALL_FIXTURES.mediaDrawing);
      const runs = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]).runs;
      expect(runs[0].eligible).toBe(false);
      expect(runs[0].hasInternalPart).toBe(true);
      expect(runs[1].eligible).toBe(true);
      expect(runs[1].hasInternalPart).toBe(false);
      expect(runs[1].text).toBe("figure caption text");
    }
    // mcAlternateContent: drawing/pict nested under mc:AlternateContent → still detected.
    {
      const doc = parsePkg(ALL_FIXTURES.mcAlternateContent);
      const runs = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]).runs;
      expect(runs[0].eligible).toBe(false);
      expect(runs[0].hasInternalPart).toBe(true);
      expect(runs[1].text).toBe("diagram caption");
    }
    // fldSimple: the field-result run is ineligible (fldSimple ANCESTOR) but carries NO internal part.
    {
      const doc = parsePkg(ALL_FIXTURES.fldSimple);
      const runs = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]).runs;
      // run[0] "page " is plain/eligible; run[1] "1" is inside <w:fldSimple> → ineligible, no part.
      expect(runs[0].eligible).toBe(true);
      expect(runs[1].eligible).toBe(false);
      expect(runs[1].hasInternalPart).toBe(false);
    }
    // hyperlinkDense: hyperlink-wrapped runs are ordinary prose — eligible, no internal part.
    {
      const doc = parsePkg(ALL_FIXTURES.hyperlinkDense);
      const runs = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]).runs;
      expect(runs.every((r) => r.hasInternalPart === false)).toBe(true);
      expect(runs.every((r) => r.eligible === true)).toBe(true);
      expect(runs.map((r) => r.text).join("")).toBe("first source and second source and third source");
    }
    // astralText: the grinning face survives as a code point, and the lone surrogate run reads verbatim.
    {
      const doc = parsePkg(ALL_FIXTURES.astralText);
      const runs = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]).runs;
      expect(runs[0].text).toBe("grin \u{1F600} end");
      expect(runs[1].text).toBe("lone\uD800surrogate");
    }
  });
});

describe("applyVisibilityInPlace — losslessness, idempotence, and mutation correctness (002-S1/F1)", () => {
  /** Build the keep/hide plan a real hide would produce for a single-paragraph fixture. */
  function planFor(runs: readonly RunView[], keepColors: Set<string>): VisibilityPlan {
    const keep = computeRunKeepFlags(runs, keepColors);
    const { extraKeep, splits } = planCrossGapSeparators(runs, keep);
    for (const i of extraKeep) keep[i] = true;
    const hideFlags = keep.map((k) => !k);
    // Hide the paragraph mark only when the whole paragraph is hidden (mirrors the engine's condensed view).
    const hideParaMark = hideFlags.every((h) => h);
    return { hideFlags, hideParaMark, splits };
  }

  it("fully-hidden body paragraph: oracle-clean delta + idempotent re-apply (rPrLessRuns)", () => {
    const doc = parsePkg(ALL_FIXTURES.rPrLessRuns);
    const before = firstParaXml(doc);
    const pp = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]);
    const plan = planFor(pp.runs, new Set()); // no keep colors → hide everything
    const r1 = pp.applyVisibilityInPlace(plan);
    expect(r1.changed).toBe(true);
    const after = firstParaXml(doc);
    // The committed paragraph differs from the input ONLY by added <w:vanish> (+ permitted bridges).
    assertVanishBridgeOnlyDeltaPara(before, after);
    // Re-apply the same plan on the now-mutated tree → no further change (idempotent / convergent).
    const r2 = pp.applyVisibilityInPlace(plan);
    expect(r2.changed).toBe(false);
    expect(firstParaXml(doc)).toBe(after);
  });

  it("partial hide with a BOUNDARY bridge split keeps two visible words apart, oracle-clean (bridgeGap)", () => {
    // alpha | (hidden) beta gamma delta | omega — fuse-prevention moves ONE space into a visible run.
    // We force the two anchors visible by treating them as kept (highlight) so the middle run hides
    // and the planner emits a boundary split. We synthesize that by building the plan from keep flags.
    const doc = parsePkg(ALL_FIXTURES.bridgeGap);
    const before = firstParaXml(doc);
    const pp = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]);
    // Keep run 0 ("alpha") and run 2 ("omega"); hide run 1 (the middle). The planner sees the gap.
    const keep = [true, false, true];
    const { extraKeep, splits } = planCrossGapSeparators(pp.runs, keep);
    for (const i of extraKeep) keep[i] = true;
    const plan: VisibilityPlan = { hideFlags: keep.map((k) => !k), hideParaMark: false, splits };
    expect(splits.length).toBeGreaterThan(0); // a real bridge split was planned for this shape
    const res = pp.applyVisibilityInPlace(plan);
    expect(res.changed).toBe(true);
    const after = firstParaXml(doc);
    assertVanishBridgeOnlyDeltaPara(before, after); // moved space + vanish only — nothing else
    // Idempotent: re-deriving the plan on the mutated tree and re-applying changes nothing further.
    const pp2 = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]);
    const keep2 = pp2.runs.map((r) => !r.hidden); // visible runs stay kept
    const { extraKeep: ek2, splits: sp2 } = planCrossGapSeparators(pp2.runs, keep2);
    for (const i of ek2) keep2[i] = true;
    const r2 = pp2.applyVisibilityInPlace({ hideFlags: keep2.map((k) => !k), hideParaMark: false, splits: sp2 });
    expect(r2.changed).toBe(false);
    expect(firstParaXml(doc)).toBe(after);
  });

  it("interior bridge split (3-way) is oracle-clean — source rPr cloned verbatim on the 'after' run", () => {
    // Build a paragraph where two kept chunks are separated by a hidden run whose ONLY space is interior
    // (", and so " has no boundary space if we trim — use a run with interior-only spacing). We craft a
    // minimal bare <w:p> so the interior-split path (exposeInteriorSpace) is exercised on a live node.
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const xml =
      `<w:document xmlns:w="${W}"><w:body><w:p>` +
      `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">reduc</w:t></w:r>` +
      `<w:r><w:t xml:space="preserve">e e</w:t></w:r>` +
      `<w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t xml:space="preserve">x</w:t></w:r>` +
      `</w:p></w:body></w:document>`;
    const doc = parsePkg(xml);
    const before = firstParaXml(doc);
    const pp = ParsedParagraph.fromNode(doc, storyParagraphs(doc)[0]);
    const keep = computeRunKeepFlags(pp.runs, new Set(["yellow"]));
    const { extraKeep, splits } = planCrossGapSeparators(pp.runs, keep);
    for (const i of extraKeep) keep[i] = true;
    expect(splits.some((s) => s.side === "interior")).toBe(true); // the interior 3-way split path
    const res = pp.applyVisibilityInPlace({ hideFlags: keep.map((k) => !k), hideParaMark: false, splits });
    expect(res.changed).toBe(true);
    const after = firstParaXml(doc);
    assertVanishBridgeOnlyDeltaPara(before, after); // 3-way split (before/space/after) accepted
  });

  it("a node-less fragment (no <w:p>) is a safe no-op", () => {
    const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
    const doc = parsePkg(`<w:document xmlns:w="${W}"><w:body><w:sectPr/></w:body></w:document>`);
    const pp = ParsedParagraph.fromNode(doc, null);
    expect(pp.runs).toEqual([]);
    expect(pp.applyVisibilityInPlace({ hideFlags: [true], hideParaMark: true }).changed).toBe(false);
  });
});

describe("the node path triggers ZERO DOMParser.parseFromString calls (002-S1 sub-gate, tests-4)", () => {
  it("fromNode + applyVisibilityInPlace parse NOTHING (parseCount === 0 on the node path)", () => {
    // Parse the package ONCE up front (this parse is expected and NOT counted against the node path).
    const doc = parsePkg(ALL_FIXTURES.multiHeading);
    const paras = storyParagraphs(doc);

    // From here on, the node path must not parse. Reset the counter AFTER the setup parse.
    parseCount = 0;
    for (const pEl of paras) {
      const pp = ParsedParagraph.fromNode(doc, pEl); // zero parse
      const keep = computeRunKeepFlags(pp.runs, new Set(["yellow"]));
      const { extraKeep, splits } = planCrossGapSeparators(pp.runs, keep);
      for (const i of extraKeep) keep[i] = true;
      pp.applyVisibilityInPlace({ hideFlags: keep.map((k) => !k), hideParaMark: false, splits }); // zero parse
    }
    expect(parseCount).toBe(0);
  });

  it("six-field-identical to the string path AND zero-parse, on the same paragraph", () => {
    const doc = parsePkg(ALL_FIXTURES.hyperlinkDense);
    const pEl = storyParagraphs(doc)[0];
    // String-path reference reads (this parses — happens before the reset).
    const stringView = legacyRunViewsForTest(serialize(pEl));

    parseCount = 0;
    const nodeView = ParsedParagraph.fromNode(doc, pEl).runs;
    expect(parseCount).toBe(0); // the node read parsed nothing
    expect(nodeView).toEqual(stringView); // and it is byte-identical to the string-path read
  });
});
