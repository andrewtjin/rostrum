// The pure Shrink engine — Rostrum's answer to Verbatim's `Shrink.bas`.
//
// Verbatim's Shrink cycles NON-underlined card text down a font size at a time (11→8→7→6→5→4→Normal),
// keeping the underlined cut readable. Rostrum widens the keep-signal (underline OR highlight OR cite,
// plus headings/structural runs are never shrunk) and adds restore-omissions and an optional shrink-
// the-paragraph-mark pass — all over the active range only.
//
// This module is PURE policy: it reads the fragment's runs (via ooxmlCondense), decides the size for
// each run with the ladder + keep predicate + omission scan, and hands a size plan back to
// ooxmlCondense to apply. No Office.js, no DOM here — the whole ladder is unit-tested in Node. It is
// the Condense & Shrink analogue of keepers.ts (policy) over ooxml.ts (editing).

import { FragmentRunView, OmissionPattern, ShrinkOptions, ShrinkResult } from "./types";
import { KEEP_OUTLINE_MAX } from "./styles";
import {
  ParagraphShrinkPlan,
  applyFragmentShrink,
  readFragmentParagraphs,
  resolveNormalSizeHalfPts
} from "./ooxmlCondense";

/** The default Normal size (half-points) when a fragment carries no styles part: 11pt (modern Word). */
export const DEFAULT_NORMAL_HALF_PTS = 22;
/** Half-points for a 6pt paragraph mark (Verbatim's "Shrink ¶"). */
const SIX_PT_HALF = 12;

/**
 * A run is kept at FULL size (never shrunk) when it carries Shrink's keep-signal: it is underlined (the
 * cut) or character-boxed — BOTH resolved through the run's character style too, since real briefs apply
 * the cut via StyleUnderline / Emphasis rather than a direct rPr — or highlighted, cite-styled, a
 * structural/ineligible run (fields, footnote refs, drawings), or a condense break marker. Mirrors
 * decision #2 ("keep underlined OR highlighted runs full-size, plus boxed/cites/headings/structural
 * always"). Headings are handled per-paragraph by the caller, not here.
 */
export function keepFullSize(run: FragmentRunView): boolean {
  return (
    run.breakMarker ||
    !run.eligible ||
    run.underline ||
    run.boxed ||
    run.highlight !== null ||
    run.citeStyled
  );
}

/** The effective current size of a run in half-points: its explicit `<w:sz>`, else the Normal size. */
export function effectiveRunSizeHalfPts(run: FragmentRunView, normalHalfPts: number): number {
  return run.sizeHalfPts ?? normalHalfPts;
}

/** The Verbatim shrink rungs in half-points: 8pt, 7pt, 6pt, 5pt, 4pt (4pt is the floor). */
const SHRINK_RUNGS = [16, 14, 12, 10, 8] as const;

/**
 * The Verbatim size ladder, one rung per press, reading the EFFECTIVE current size in half-points:
 * step to the LARGEST rung strictly below the current size (>8pt → 8pt, 8→7, 7→6, 6→5, 5→4), and once
 * at/below the 4pt floor, return `null` to clear the explicit size — so it resets to the inherited
 * Normal (the cycle-back / un-shrink rung). The cycle then resumes from Normal on the next press.
 *
 * Using "largest rung strictly below current" (rather than rounding to whole points then table-mapping)
 * fixes two adversarial defects: it never GROWS text when Normal is already small (a sub-4pt size has no
 * rung below it, so it clears to Normal instead of jumping up to 8pt), and it handles non-whole-point
 * sizes correctly (8.5pt → 8pt, 7.5pt → 7pt) without skipping a rung.
 */
export function nextShrinkSize(currentHalfPts: number): number | null {
  for (const rung of SHRINK_RUNGS) {
    if (rung < currentHalfPts) return rung;
  }
  return null; // at/below the 4pt floor → clear to Normal (reset)
}

/** A paragraph's canonical outline level marks it a heading (kept whole) when in [0, KEEP_OUTLINE_MAX]. */
function isHeading(level: number | null): boolean {
  return level !== null && level >= 0 && level <= KEEP_OUTLINE_MAX;
}

/**
 * The run indices in ONE paragraph whose text falls inside an omission span — a region from a pattern's
 * `open` delimiter to its next `close`, whose inner text contains the pattern's `keyword` (case-
 * insensitive; empty keyword matches any). Those runs are restored to Normal size so an "[…Omitted…]"
 * indicator stays readable. A run counts when its character range overlaps any omission span.
 */
export function omissionRunIndices(
  runs: readonly FragmentRunView[],
  patterns: readonly OmissionPattern[]
): Set<number> {
  const omitted = new Set<number>();
  if (patterns.length === 0) return omitted;
  const text = runs.map((r) => r.text).join("");
  // Character range [start, end) per run, in concatenation order.
  const ranges: Array<[number, number]> = [];
  let pos = 0;
  for (const r of runs) {
    ranges.push([pos, pos + r.text.length]);
    pos += r.text.length;
  }
  for (const pat of patterns) {
    if (!pat.open || !pat.close) continue;
    let from = 0;
    // eslint-disable-next-line no-constant-condition
    while (true) {
      const open = text.indexOf(pat.open, from);
      if (open < 0) break;
      const close = text.indexOf(pat.close, open + pat.open.length);
      if (close < 0) break;
      const spanStart = open;
      const spanEnd = close + pat.close.length;
      const inner = text.slice(open + pat.open.length, close);
      const matches = pat.keyword === "" || inner.toLowerCase().includes(pat.keyword.toLowerCase());
      if (matches) {
        for (let i = 0; i < runs.length; i++) {
          const [s, e] = ranges[i];
          if (s < spanEnd && e > spanStart) omitted.add(i); // ranges overlap
        }
      }
      from = spanEnd;
    }
  }
  return omitted;
}

/** Does any run in the fragment carry an alphanumeric character (used to pick a representative size)? */
function hasAlphanumeric(text: string): boolean {
  return /[\p{L}\p{N}]/u.test(text);
}

/**
 * Shrink one press over the active-range fragment. Finds a representative current size from the first
 * non-kept alphanumeric run (Verbatim ignores punctuation so fixed-size markers don't skew it), computes
 * the next rung, and applies it to EVERY non-kept run across the non-heading paragraphs (normalizing
 * mixed sizes — Verbatim parity), restoring omission spans to Normal. Optionally shrinks each
 * non-heading paragraph mark to 6pt. A collapsed/single-paragraph heading refuses (Verbatim parity).
 */
export function shrinkFragment(fragmentXml: string, opts: ShrinkOptions): ShrinkResult & { xml: string } {
  const paras = readFragmentParagraphs(fragmentXml);
  const normalHalfPts = opts.normalHalfPts;

  // Heading refusal: a single-paragraph (collapsed or one-paragraph) selection on a heading is a no-op.
  if (paras.length === 1 && isHeading(opts.outlineLevels[0] ?? null)) {
    return { xml: fragmentXml, changed: false, paragraphsScanned: paras.length, refusedHeading: true };
  }

  // Representative current size: the first non-kept ALPHANUMERIC run in a non-heading paragraph
  // (Verbatim ignores punctuation so fixed-size markers don't skew the reading). If a paragraph is all
  // punctuation but still has a sizeable non-kept run, fall back to the first non-kept run's size so the
  // text isn't left un-shrinkable (the adversarial "all-punctuation never shrinks" gap).
  let representative: number | null = null;
  let fallback: number | null = null;
  for (let p = 0; p < paras.length && representative === null; p++) {
    if (isHeading(opts.outlineLevels[p] ?? null)) continue;
    for (const run of paras[p]) {
      if (keepFullSize(run)) continue;
      if (fallback === null) fallback = effectiveRunSizeHalfPts(run, normalHalfPts);
      if (hasAlphanumeric(run.text)) {
        representative = effectiveRunSizeHalfPts(run, normalHalfPts);
        break;
      }
    }
  }
  representative = representative ?? fallback;
  if (representative === null) {
    // Nothing eligible to shrink (everything is a keeper / structural / heading).
    return { xml: fragmentXml, changed: false, paragraphsScanned: paras.length, appliedSizeHalfPts: undefined };
  }

  const target = nextShrinkSize(representative); // number (half-points) or null (→ Normal)

  const plans: ParagraphShrinkPlan[] = paras.map((runs, p) => {
    if (isHeading(opts.outlineLevels[p] ?? null)) {
      return { runSizes: runs.map(() => undefined) };
    }
    const omitted = omissionRunIndices(runs, opts.omissionPatterns);
    const runSizes = runs.map((run, i) => {
      if (keepFullSize(run)) return undefined; // keepers stay full size
      if (omitted.has(i)) return null; // restore omissions to Normal
      return target; // non-kept body text → the next rung
    });
    const plan: ParagraphShrinkPlan = { runSizes };
    if (opts.shrinkParagraphMarks) plan.markSizeHalfPts = SIX_PT_HALF;
    return plan;
  });

  const applied = applyFragmentShrink(fragmentXml, plans);
  return {
    xml: applied.xml,
    changed: applied.changed,
    paragraphsScanned: paras.length,
    appliedSizeHalfPts: target
  };
}

/**
 * Unshrink: clear the explicit size on every non-kept run across the non-heading paragraphs (reverting
 * them to the inherited Normal size — the debate convention), and clear any shrunk paragraph mark. This
 * is "reset to Normal," not a per-run size restore (a future lossless add-on could reuse a sidecar).
 * Heading paragraphs and keeper runs are left untouched.
 */
export function unshrinkFragment(
  fragmentXml: string,
  outlineLevels: (number | null)[]
): ShrinkResult & { xml: string } {
  const paras = readFragmentParagraphs(fragmentXml);
  const plans: ParagraphShrinkPlan[] = paras.map((runs, p) => {
    if (isHeading(outlineLevels[p] ?? null)) {
      return { runSizes: runs.map(() => undefined) };
    }
    const runSizes = runs.map((run) => (keepFullSize(run) ? undefined : null));
    // Clear any shrunk paragraph mark too (markSize null = remove the explicit mark size).
    return { runSizes, markSizeHalfPts: null };
  });
  const applied = applyFragmentShrink(fragmentXml, plans);
  return {
    xml: applied.xml,
    changed: applied.changed,
    paragraphsScanned: paras.length,
    appliedSizeHalfPts: null
  };
}

/**
 * Resolve the Normal size (half-points) to feed the ladder: the fragment's `styles.xml` default, else
 * the modern-Word 11pt fallback. Exposed so the controller can pass it through `ShrinkOptions`.
 */
export function resolveNormalHalfPts(fragmentXml: string): number {
  return resolveNormalSizeHalfPts(fragmentXml) ?? DEFAULT_NORMAL_HALF_PTS;
}
