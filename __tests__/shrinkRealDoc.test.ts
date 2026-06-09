// REAL-DOCUMENT validation of the two Shrink fixes — the headless proof they hold on the ACTUAL
// wet-test doc, not just the synthetic fixtures in shrink.test.ts.
//
// BUG 1 — style-resolved keep. In real debate docs the cut/emphasis is applied through CHARACTER
// STYLES (StyleUnderline → <w:u>; Emphasis → <w:u>+<w:bdr>), NOT direct run formatting. The fix
// (resolveStyleEmphasis + the `boxed` signal + keepFullSize resolving direct-THEN-style in readRun) must
// keep those runs at full size and shrink only the genuinely un-emphasized prose, on the real styles.xml
// Word emitted. Layers A/B/C below run the real engine over the real OOXML; a NEGATIVE CONTROL (strip the
// styles part → the style-only cut shrinks) proves the style resolution is load-bearing — the in-CI,
// permanent equivalent of reverting the fix, with teeth.
//   TEETH (asserts that FAIL on the pre-fix code): this doc carries NO direct <w:bdr> anywhere and the
//   target card paragraph (P6) has NO direct <w:u>/<w:bdr> at all — every box/underline signal there is
//   style-derived. So `boxed.length>0` (Layer B) and `some(underline)`/`some(boxed)` + the exact shrink
//   rung (Layer C) can ONLY hold once readRun resolves a run's character style.
//
// BUG 2 — selection persistence after a write-back. replaceActiveRangeOoxml re-select()s the inserted
// range so a repeated Shrink keeps the whole block live. That is HOST behavior (insertOoxml / Range.select)
// and is NOT reproducible in a headless doc test; it is already covered by the adapter fake in
// rangePort.test.ts and remains a live wet-test confirmation. It is documented here as an `it.todo` so the
// coverage boundary is VISIBLE, never silently absent (this plan adds no false headless test for it).
//
// RESILIENCE: rostrum/samples is gitignored, so the whole suite SKIPS (never fails) when the sample —
// or its styles.xml — is absent, mirroring realDocs.test.ts. CI has no samples; it is green-by-skip there.

import {
  discoverSamples,
  readDocxParts,
  buildRangePackage,
  singleParagraphDocumentXml,
  SampleRef,
  W_NS
} from "./realDocs";
import { readFragmentParagraphs, resolveStyleEmphasis } from "../src/core/ooxmlCondense";
import {
  keepFullSize,
  shrinkFragment,
  resolveNormalHalfPts,
  nextShrinkSize,
  omissionRunIndices
} from "../src/core/shrink";
import { DEFAULT_OMISSION_PATTERNS } from "../src/core/settings";
import { ShrinkOptions } from "../src/core/types";

/** A valid-but-empty styles part: the negative control where NO character style can resolve emphasis. */
const EMPTY_STYLES = `<w:styles xmlns:w="${W_NS}"></w:styles>`;
/** The cut is encoded through this character style in real docs; also the Layer-C paragraph selector. */
const STYLE_UNDERLINE_RE = /w:val="StyleUnderline"/;

// Locate the wet-test doc by filename substring; the whole suite skips cleanly when it is absent.
const sample: SampleRef | undefined = discoverSamples().find((s) => /shrink test/i.test(s.file));

describe("Shrink on the real wet-test doc (samples/)", () => {
  if (!sample) {
    // eslint-disable-next-line no-console
    console.info("[shrinkRealDoc] '[…shrink test…]' sample absent (samples/ is gitignored) — skipping.");
    it.skip("real-doc Shrink validation (drop the sample into rostrum/samples to enable)", () => undefined);
    return;
  }
  const present: SampleRef = sample;

  // Read the doc ONCE for the whole suite (a .docx is a zip; no need to re-open per test).
  let documentXml = "";
  let stylesXml: string | null = null;
  beforeAll(async () => {
    const parts = await readDocxParts(present.fullPath);
    documentXml = parts.documentXml;
    stylesXml = parts.stylesXml;
  });

  /**
   * True when the doc has no styles part — a degenerate shape a real .docx never has. Per the plan's
   * "never fail for a missing corpus" mandate we log + bail out of the test body rather than assert; the
   * name says so at the call site (`if (skipIfNoStyles()) return;`). The whole-suite-absent case is a true
   * `it.skip` above; this finer guard only fires for a corrupt sample and is defense-in-depth.
   */
  const skipIfNoStyles = (): boolean => {
    if (stylesXml) return false;
    // eslint-disable-next-line no-console
    console.info("[shrinkRealDoc] sample carries no styles.xml — skipping (cannot validate style-resolved keep).");
    return true;
  };

  /** The first style-underlined card paragraph, as a Shrink-ready range package (document + styles parts). */
  const cardParagraphPkg = (styles: string = stylesXml ?? ""): { oneP: string; pkg: string } => {
    const oneP = singleParagraphDocumentXml(documentXml, (x) => STYLE_UNDERLINE_RE.test(x));
    if (oneP === null) throw new Error("no StyleUnderline card paragraph found in the sample");
    return { oneP, pkg: buildRangePackage(oneP, styles) };
  };

  /** Shrink options for a single body card paragraph: outlineLevels [null] ⇒ no heading refusal. */
  const cardOpts = (pkg: string): ShrinkOptions => ({
    normalHalfPts: resolveNormalHalfPts(pkg),
    outlineLevels: [null],
    omissionPatterns: [...DEFAULT_OMISSION_PATTERNS],
    shrinkParagraphMarks: false
  });

  // bug-2 is host-only — see the file header + rangePort.test.ts + the live wet-test. No headless test.
  it.todo("bug-2: reselect after write-back is host-only (rangePort.test.ts + live wet-test cover it)");

  // --- Layer A: the resolver over the REAL styles.xml (no body, no outline). ---
  it("Layer A — resolveStyleEmphasis reads the doc's real character styles", () => {
    if (skipIfNoStyles()) return;
    const map = resolveStyleEmphasis(stylesXml as string);
    // The cut is StyleUnderline (underline, no box); the boxed emphasis is Emphasis (underline + box).
    expect(map.get("StyleUnderline")).toEqual({ underline: true, boxed: false });
    expect(map.get("Emphasis")).toEqual({ underline: true, boxed: true });
    // The cite style EXISTS but carries an explicit <w:u w:val="none"/> in this doc, so it resolves to
    // NO emphasis — cite runs are kept via citeStyled, not via the emphasis map. This also exercises the
    // tri-state explicit-"none" path of the resolver on real OOXML.
    expect(map.has("Style13ptBold")).toBe(true);
    expect(map.get("Style13ptBold")).toEqual({ underline: false, boxed: false });
  });

  // --- Layer B: run-keep over the REAL whole body (range package, no shrink, no outline). ---
  it("Layer B — every emphasized run on the real body is kept; plain prose is shrinkable", () => {
    if (skipIfNoStyles()) return;
    const pkg = buildRangePackage(documentXml, stylesXml as string);
    const runs = readFragmentParagraphs(pkg).flat();
    expect(runs.length).toBeGreaterThan(0); // parse smoke: the engine reads the real ~99-paragraph body

    const underlined = runs.filter((r) => r.underline);
    const boxed = runs.filter((r) => r.boxed);
    const highlighted = runs.filter((r) => r.highlight !== null);
    const cited = runs.filter((r) => r.citeStyled);

    // TEETH: the doc carries NO direct <w:bdr> anywhere — every boxed run is boxed THROUGH the Emphasis
    // character style, so a non-empty boxed set can ONLY hold once style resolution works (bug-1). On the
    // pre-fix reader (direct rPr only) this set is empty and the assertion fails.
    expect(boxed.length).toBeGreaterThan(0);
    expect(underlined.length).toBeGreaterThan(0);
    expect(highlighted.length).toBeGreaterThan(0); // the kept highlight (cyan)
    expect(cited.length).toBeGreaterThan(0); // the cite runs (Style13ptBold)

    // Every emphasized run is kept full-size — the keep contract holds across the WHOLE real body, not
    // just a fixture. (keepFullSize is a superset of the emphasis signals, so this guards a future
    // regression that NARROWS keepFullSize away from underline/box/highlight/cite.)
    for (const r of runs) {
      if (r.underline || r.boxed || r.highlight !== null || r.citeStyled) {
        expect(keepFullSize(r)).toBe(true);
      }
    }

    // Non-vacuous: at least one PLAIN, eligible, lettered run exists that WOULD shrink — otherwise the
    // doc has nothing to shrink and the whole exercise would be meaningless.
    const plainShrinkable = runs.filter((r) => !keepFullSize(r) && /\p{L}/u.test(r.text));
    expect(plainShrinkable.length).toBeGreaterThan(0);
  });

  // --- Layer C: end-to-end Shrink on a single real card paragraph (the strongest proof). ---
  it("Layer C — shrinks plain prose while keeping the style-underlined/boxed cut full-size", () => {
    if (skipIfNoStyles()) return;
    const { oneP, pkg } = cardParagraphPkg();
    const opts = cardOpts(pkg);
    const before = readFragmentParagraphs(pkg)[0];
    expect(before.length).toBeGreaterThan(0); // parse smoke on the reconstructed single-paragraph package

    // TEETH for bug-1: this card paragraph has NO direct <w:u>/<w:bdr> — its underline AND box come ONLY
    // from the StyleUnderline / Emphasis character styles, so these are true ONLY when readRun resolves
    // the run's character style (the fix). On the pre-fix code both are empty here.
    expect(before.some((r) => r.underline)).toBe(true);
    expect(before.some((r) => r.boxed)).toBe(true);

    const out = shrinkFragment(pkg, opts);
    const after = readFragmentParagraphs(out.xml)[0];
    expect(out.changed).toBe(true);
    // The doc's plain card prose is explicit 8pt (sz=16); one press steps it to 7pt (sz=14). This exact
    // rung ALSO distinguishes the fix: pre-fix the style-cut runs are NOT kept, so the first non-kept run is
    // the underlined opener with no explicit size → its effective size is the inherited Normal (12pt, sz=24
    // per this doc's docDefaults) → nextShrinkSize(24)=16, not 14. (Verified by the teeth probe.)
    expect(out.appliedSizeHalfPts).toBe(14);

    // This card carries no "[…Omitted…]" spans, so the omission branch below never fires here — omission
    // RESTORE on real OOXML is out of scope for this suite (the synthetic shrink.test.ts covers it). The
    // branch stays so the run-by-run check is exhaustive and correct if a future sample does carry one.
    const omitted = omissionRunIndices(before, opts.omissionPatterns);
    let kept = 0;
    let shrunk = 0;
    before.forEach((r, i) => {
      if (keepFullSize(r)) {
        expect(after[i].sizeHalfPts).toBe(r.sizeHalfPts); // kept: size UNCHANGED (the cut stays full size)
        kept++;
      } else if (omitted.has(i)) {
        expect(after[i].sizeHalfPts).toBeNull(); // omission spans restore to Normal
      } else {
        expect(after[i].sizeHalfPts).toBe(out.appliedSizeHalfPts); // plain prose → the next rung
        shrunk++;
      }
    });
    // Non-vacuous: this single press BOTH kept something and shrank something.
    expect(kept).toBeGreaterThan(0);
    expect(shrunk).toBeGreaterThan(0);

    // NEGATIVE CONTROL — proves the fix is load-bearing (the regression nub, in-CI). Rebuild the SAME
    // paragraph with the styles part STRIPPED: now no character style resolves, so the style-only cut runs
    // are NOT kept. At least one run flips kept→shrinkable, and it shrinks — the exact pre-fix bug.
    const stylelessPkg = buildRangePackage(oneP, EMPTY_STYLES);
    const beforeNoStyles = readFragmentParagraphs(stylelessPkg)[0];
    const flipped = before
      .map((r, i) => ({ r, i }))
      .filter(({ r, i }) => keepFullSize(r) && !keepFullSize(beforeNoStyles[i]));
    expect(flipped.length).toBeGreaterThan(0);
    for (const { i } of flipped) {
      // A flipped run is a style-only cut: in this corpus those carry NO explicit size (the cut inherits
      // the full 11pt), and once styles are gone they show no direct underline/box. If a future sample
      // hand-sized a style-cut run, this asserts LOUDLY (the intended fail-loud, not a silent pass).
      expect(before[i].sizeHalfPts).toBeNull();
      expect(beforeNoStyles[i].underline).toBe(false);
      expect(beforeNoStyles[i].boxed).toBe(false);
    }
    const outNoStyles = shrinkFragment(stylelessPkg, {
      ...opts,
      normalHalfPts: resolveNormalHalfPts(stylelessPkg)
    });
    const afterNoStyles = readFragmentParagraphs(outNoStyles.xml)[0];
    const probe = flipped[0].i;
    expect(after[probe].sizeHalfPts).toBe(before[probe].sizeHalfPts); // kept full-size WITH real styles
    expect(afterNoStyles[probe].sizeHalfPts).toBe(outNoStyles.appliedSizeHalfPts); // shrinks WITHOUT them
  });

  // --- Layer C+: repeated presses (the path bug-2's reselect makes usable on the host). ---
  it("Layer C+ — a second press steps plain prose further down while the cut stays full size", () => {
    if (skipIfNoStyles()) return;
    const { pkg } = cardParagraphPkg();
    const opts = cardOpts(pkg);

    const press1 = shrinkFragment(pkg, opts);
    expect(press1.appliedSizeHalfPts).toBe(14); // 8pt → 7pt
    expect(press1.appliedSizeHalfPts).toBe(nextShrinkSize(16)); // self-consistent with the ladder

    const press2 = shrinkFragment(press1.xml, opts);
    expect(press2.appliedSizeHalfPts).toBe(12); // 7pt → 6pt
    expect(press2.appliedSizeHalfPts).toBe(nextShrinkSize(14));

    const before = readFragmentParagraphs(pkg)[0];
    const after2 = readFragmentParagraphs(press2.xml)[0];
    const omitted = omissionRunIndices(before, opts.omissionPatterns);
    before.forEach((r, i) => {
      if (keepFullSize(r)) {
        expect(after2[i].sizeHalfPts).toBe(r.sizeHalfPts); // the cut is STILL full size after two presses
      } else if (!omitted.has(i)) {
        expect(after2[i].sizeHalfPts).toBe(12); // plain prose now at 6pt
      }
    });
  });
});
