// SELF-TEST for the semantic diff oracle (PLAN.md §8 loss-1 / tests-1).
//
// An oracle is worthless if it green-passes bugs. These tests PROVE the oracle both
// (a) ACCEPTS every legitimate engine delta — identical in/out, a pure `<w:vanish>`
// toggle, and a real bridge-split (one space moved into a new `xml:space="preserve"`
// run) — and (b) THROWS on every violation class: a changed text character, an added
// non-vanish rPr child (`<w:b/>`), a reordered run, a deleted run, and a DUPLICATED
// (not moved) bridge space.
//
// Where possible the "legitimate" outputs are produced by the REAL engine
// (`ParsedParagraph.applyVisibility`, the exact code the hide path runs), so the oracle
// is validated against genuine engine output — not just hand-written XML that happens to
// match my mental model of it. The violation outputs are hand-crafted minimal mutations.

import { ParsedParagraph } from "../src/core/ooxml";
import { classifyParagraph } from "../src/core/invisibility";
import { BridgeSplit, RawParagraph } from "../src/core/types";
import { settings } from "./fakeWord";
import {
  assertVanishBridgeOnlyDelta,
  assertVanishBridgeOnlyDeltaPara
} from "./semanticDiff";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/** A bare `<w:p>` fragment with the given inner run XML (what ooxml.ts operates on). */
const para = (inner: string): string => `<w:document xmlns:w="${W}"><w:body><w:p>${inner}</w:p></w:body></w:document>`;
const run = (text: string, rPr = ""): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;

// ===========================================================================
// (a) LEGITIMATE deltas — the oracle must PASS these.
// ===========================================================================

describe("semantic diff oracle — accepts legitimate engine deltas", () => {
  it("identical in/out → passes (no-op hide)", () => {
    const x = para(run("alpha") + run("beta"));
    expect(() => assertVanishBridgeOnlyDeltaPara(x, x)).not.toThrow();
  });

  it("a pure <w:vanish> toggle (added) → passes", () => {
    // Drive the REAL engine: hide every run, no bridge splits. Output differs from input
    // ONLY by added <w:vanish> elements — exactly permitted delta (a).
    const input = para(run("hidden body one") + run("hidden body two"));
    const { xml: output, changed } = new ParsedParagraph(input).applyVisibility([true, true], false);
    expect(changed).toBe(true);
    expect(output).toContain("<w:vanish/>");
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).not.toThrow();
  });

  it("a pure <w:vanish> toggle (removed, Show All) → passes", () => {
    // Start from a hidden paragraph; makeAllVisible clears vanish. Output differs from the
    // hidden input only by REMOVED <w:vanish> — permitted delta (a) in the other direction.
    const hidden = para(run("x", "<w:vanish/>") + run("y", "<w:vanish/>"));
    const { xml: shown, changed } = new ParsedParagraph(hidden).makeAllVisible();
    expect(changed).toBe(true);
    expect(shown).not.toContain("<w:vanish/>");
    expect(() => assertVanishBridgeOnlyDeltaPara(hidden, shown)).not.toThrow();
  });

  it("hiding the paragraph mark (<w:pPr><w:rPr><w:vanish/>) → passes", () => {
    const input = para(run("body"));
    const { xml: output } = new ParsedParagraph(input).applyVisibility([true], true);
    expect(output).toContain("<w:pPr>");
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).not.toThrow();
  });

  it("a legitimate bridge split (boundary space MOVED into a new preserve run) → passes", () => {
    // "alpha" | hidden " beta " | "omega": hide the middle run AND expose its leading space
    // as a visible bridge run. The engine MOVES one space (drops it from the source <w:t>,
    // re-adds it in a new <w:r xml:space="preserve">). Net paragraph text is unchanged.
    const input = para(run("alpha") + run(" beta ") + run("omega"));
    const splits: BridgeSplit[] = [{ index: 1, side: "lead" }];
    const { xml: output, changed } = new ParsedParagraph(input).applyVisibility(
      [false, true, false],
      false,
      splits
    );
    expect(changed).toBe(true);
    // Sanity: a new visible space run was inserted with preserve.
    expect(output).toContain('xml:space="preserve"');
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).not.toThrow();
  });

  it("a legitimate INTERIOR bridge split (3-way, space MOVED) → passes", () => {
    // One hidden run "left right" whose only space is interior; the engine splits it into
    // [left](hidden) [space](visible) [right](hidden). Text total unchanged.
    const input = para(run("a") + run("left right") + run("b"));
    const splits: BridgeSplit[] = [{ index: 1, side: "interior", offset: 4 }];
    const { xml: output, changed } = new ParsedParagraph(input).applyVisibility(
      [false, true, false],
      false,
      splits
    );
    expect(changed).toBe(true);
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).not.toThrow();
  });

  it("whole-package multi-paragraph vanish toggle → passes", () => {
    const pkg = (paras: string): string =>
      `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W}"><w:body>${paras}<w:sectPr/></w:body></w:document>` +
      `</pkg:xmlData></pkg:part></pkg:package>`;
    const p = (inner: string): string => `<w:p>${inner}</w:p>`;
    const input = pkg(p(run("keep heading")) + p(run("hide this body")));
    // Hand-build the output: second paragraph's run gains a <w:vanish/>; first untouched.
    const output = pkg(
      p(run("keep heading")) + p(`<w:r><w:rPr><w:vanish/></w:rPr><w:t xml:space="preserve">hide this body</w:t></w:r>`)
    );
    expect(() => assertVanishBridgeOnlyDelta(input, output)).not.toThrow();
  });
});

// ===========================================================================
// (b) VIOLATIONS — the oracle must THROW on each.
// ===========================================================================

describe("semantic diff oracle — throws on every violation class", () => {
  it("a CHANGED text character → throws (text-preservation, 002-F1)", () => {
    const input = para(run("alpha") + run("beta"));
    const output = para(run("alpha") + run("beXa")); // 't' -> 'X'
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow(/visible run text changed|text/i);
  });

  it("an ADDED non-vanish rPr child (<w:b/>) → throws", () => {
    const input = para(run("alpha"));
    const output = para(run("alpha", "<w:b/>")); // bold injected — NOT a permitted delta
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow(/attributes differ|child count differs|<w:b\/>|structure|permitted/i);
  });

  it("a non-vanish rPr child added ALONGSIDE a vanish → still throws (vanish doesn't license <w:b/>)", () => {
    const input = para(run("alpha"));
    const output = para(run("alpha", "<w:vanish/><w:b/>")); // vanish ok, bold is not
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow();
  });

  it("a REORDERED run → throws", () => {
    const input = para(run("alpha") + run("beta"));
    const output = para(run("beta") + run("alpha")); // same runs, swapped order
    // Reordering changes the concatenated text ("alphabeta" -> "betaalpha"), so #1 fires;
    // either way it must throw.
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow();
  });

  it("a reordered run that PRESERVES total text → still throws (structure)", () => {
    // Swap two runs whose concatenation is identical forwards/backwards is impossible for
    // distinct text, so use rPr-distinguished runs with the SAME text to force a structural
    // (not text) catch: run A has bold, run B plain, both "xy". Swapping changes which run
    // is bold at which position → structural diff must catch it even though text is "xyxy"
    // both ways.
    const input = para(run("xy", "<w:b/>") + run("xy"));
    const output = para(run("xy") + run("xy", "<w:b/>"));
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow();
  });

  it("a DELETED run → throws", () => {
    const input = para(run("alpha") + run("beta"));
    const output = para(run("alpha")); // 'beta' run removed entirely
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow(/missing|deleted|text/i);
  });

  it("a DUPLICATED (not moved) bridge space → throws", () => {
    // The legitimate bridge split MOVES a space: source loses one space, a new run gains it,
    // net text unchanged. Here we INSERT a visible space run but DO NOT remove the space from
    // the source — so the paragraph now has an EXTRA space → concatenated text grew. #1 fires.
    const input = para(run("alpha") + run(" beta ", "<w:vanish/>") + run("omega"));
    // Output: source keeps " beta " intact AND a duplicate visible space run is inserted.
    const output = para(
      run("alpha") +
        `<w:r><w:t xml:space="preserve"> </w:t></w:r>` + // duplicated space (not moved)
        run(" beta ", "<w:vanish/>") +
        run("omega")
    );
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow(/text|visible run text changed/i);
  });

  it("an inserted NON-space run (smuggled visible word) → throws", () => {
    // An inserted run that is NOT a bare bridge space must be rejected even though it carries
    // preserve — it changes the visible text (#1) and is not a whitelisted insertion.
    const input = para(run("alpha") + run("omega", "<w:vanish/>"));
    const output = para(
      run("alpha") + `<w:r><w:t xml:space="preserve">INJECTED</w:t></w:r>` + run("omega", "<w:vanish/>")
    );
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow();
  });

  it("a changed rPr ATTRIBUTE (highlight recolored) → throws", () => {
    const input = para(run("alpha", '<w:highlight w:val="yellow"/>'));
    const output = para(run("alpha", '<w:highlight w:val="green"/>')); // recolor — not permitted
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow(/attributes differ|differ/i);
  });

  it("a changed paragraph COUNT (whole package) → throws", () => {
    const pkg = (paras: string): string =>
      `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W}"><w:body>${paras}<w:sectPr/></w:body></w:document>` +
      `</pkg:xmlData></pkg:part></pkg:package>`;
    const p = (inner: string): string => `<w:p>${inner}</w:p>`;
    const input = pkg(p(run("one")) + p(run("two")));
    const output = pkg(p(run("one"))); // a paragraph vanished
    expect(() => assertVanishBridgeOnlyDelta(input, output)).toThrow(/paragraph count changed/i);
  });

  it("a changed pStyle (paragraph structure) → throws", () => {
    const input = para(`<w:pPr><w:pStyle w:val="Normal"/></w:pPr>` + run("body"));
    const output = para(`<w:pPr><w:pStyle w:val="Heading1"/></w:pPr>` + run("body"));
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow(/attributes differ|structure|differ/i);
  });

  it("an empty/absent paragraph side → throws (no vacuous pass)", () => {
    const input = para(run("alpha"));
    const empty = `<w:document xmlns:w="${W}"><w:body></w:body></w:document>`;
    expect(() => assertVanishBridgeOnlyDeltaPara(input, empty)).toThrow(/no <w:p>|nothing to compare/i);
  });

  it("an interior-split 'after' run that SMUGGLES a format change → throws (clone must be faithful)", () => {
    // A legit interior split's after-run clones the SOURCE rPr verbatim. Here the inserted
    // after-run carries a <w:b/> the source never had — the matcher must reject it (it is
    // neither the next input run's image nor a faithful clone of the just-consumed source).
    const input = para(run("a") + run("left right", "<w:i/>") + run("b"));
    const output = para(
      run("a") +
        `<w:r><w:rPr><w:i/><w:vanish/></w:rPr><w:t xml:space="preserve">left</w:t></w:r>` + // before (faithful)
        `<w:r><w:t xml:space="preserve"> </w:t></w:r>` + // visible space
        `<w:r><w:rPr><w:i/><w:b/><w:vanish/></w:rPr><w:t xml:space="preserve">right</w:t></w:r>` + // after + SMUGGLED bold
        run("b")
    );
    expect(() => assertVanishBridgeOnlyDeltaPara(input, output)).toThrow();
  });
});

// ===========================================================================
// END-TO-END — drive the REAL engine pipeline (classifyParagraph) over body
// paragraphs and prove the oracle accepts genuine whole-engine output, plus
// Re-hide idempotence (002-S1 AMENDED: Hide(Hide(x)) run-set == Hide(x)).
// ===========================================================================

describe("semantic diff oracle — accepts genuine classifyParagraph output (end to end)", () => {
  const KEEP = settings(["yellow"]);
  const hl = (color: string): string => `<w:highlight w:val="${color}"/>`;

  it("a fully-hidden body paragraph (classifyParagraph) → oracle passes", () => {
    const input = para(run("an entirely unhighlighted body card sentence"));
    const plan = classifyParagraph(
      { index: 0, headingLevel: null, inTable: false, ooxml: input },
      KEEP
    );
    expect(plan.action).toBe("hideWhole");
    expect(() => assertVanishBridgeOnlyDeltaPara(input, plan.ooxml)).not.toThrow();
  });

  it("a partial-hide with a real cross-gap bridge (classifyParagraph) → oracle passes", () => {
    // Two highlighted anchors with hidden prose between them: the engine hides the middle
    // and MOVES a space so the anchors don't fuse — the exact loss-1 (b) shape, produced by
    // the real keepers + ooxml pipeline (not hand-written).
    const input = para(
      run("rising", hl("yellow")) + run(" oil prices fall and ") + run("revenues", hl("yellow"))
    );
    const plan = classifyParagraph(
      { index: 0, headingLevel: null, inTable: false, ooxml: input },
      KEEP
    );
    expect(plan.action).toBe("hidePartial");
    expect(() => assertVanishBridgeOnlyDeltaPara(input, plan.ooxml)).not.toThrow();
  });

  it("Re-hide idempotence: Hide(Hide(x)) differs from Hide(x) only by permitted deltas", () => {
    // 002-S1 AMENDED: re-hiding an already-hidden paragraph must not change text/structure
    // beyond permitted deltas. We run classify twice and assert the oracle accepts the
    // second pass's output vs the first pass's output (convergent).
    const input = para(
      run("keep", hl("yellow")) + run(" hidden middle clause ") + run("end", hl("yellow"))
    );
    const p: RawParagraph = { index: 0, headingLevel: null, inTable: false, ooxml: input };
    const once = classifyParagraph(p, KEEP).ooxml;
    const twice = classifyParagraph({ ...p, ooxml: once }, KEEP).ooxml;
    // Re-hide vs first hide: permitted-delta only (the run set converges).
    expect(() => assertVanishBridgeOnlyDeltaPara(once, twice)).not.toThrow();
    // And the first hide vs the original input is itself permitted-delta only.
    expect(() => assertVanishBridgeOnlyDeltaPara(input, once)).not.toThrow();
  });
});
