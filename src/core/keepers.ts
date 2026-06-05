// Pure keeper rules — the heart of the policy (decisions #6, #7, #19).
//
// A run/paragraph is KEPT iff:
//   (a) its paragraph's resolved outline level is 0–3  (heading rule, #7), OR
//   (b) the paragraph contains a cite-styled run        (kept whole, #6b), OR
//   (c) the run is highlighted in a keep-color           (#6c, #11).
// We show ONLY highlighted text: a highlight covering part of a word does NOT drag the
// unhighlighted remainder visible (decision #19 reversed per wet-test). Everything else
// in the body is hidden. No Office.js here — these are total,
// deterministic functions over the views ooxml.ts produces, so the entire policy
// is unit-tested against the rule matrix.

import { RunView, BridgeSplit } from "./types";
import { KEEP_OUTLINE_MAX } from "./styles";

/**
 * Heading rule (#7): keep when the canonical 0-based outline level is in [0, 3].
 * `headingLevel` is null for body text. Keying on the resolved level (not a style
 * allowlist) auto-covers Heading 1–4, the template's navy Analytics style
 * (level 3), and any future derived heading style.
 */
export function isHeadingKept(headingLevel: number | null): boolean {
  return headingLevel !== null && headingLevel >= 0 && headingLevel <= KEEP_OUTLINE_MAX;
}

/** Cite rule (#6b): a paragraph containing any cite-styled run is kept whole. */
export function paragraphHasCiteRun(runs: readonly RunView[]): boolean {
  return runs.some((r) => r.citeStyled);
}

/**
 * Run separators for the fusion-prevention logic — every character that RENDERS AS A
 * VISIBLE SPACE, so exposing one keeps two visible chunks from gluing. This includes the
 * ASCII whitespace AND the no-break / typographic spaces that pasted debate text is full of:
 * NBSP (U+00A0), figure space (U+2007), thin space (U+2009), narrow NBSP (U+202F). NBSP was
 * ORIGINALLY excluded ("so it joins words, e.g. Dr. Smith") back when a highlight extended to
 * the whole word it touched; that whole-word extension was removed (decision #19 reversed), and
 * the exclusion then caused a real wet-test bug: a hidden NBSP between two kept words fused them
 * ("rising oil revenues" → "risingrevenues", "gives Russia" → "givesRussia") because the bridge
 * and the whitespace-only rescue never saw the NBSP as a space. Zero-width characters (U+200B,
 * U+2060) are deliberately NOT here — exposing one wouldn't separate anything visually.
 *
 * EM/EN DASHES (U+2014, U+2013) are ALSO included — not because they render as a space, but because
 * in prose they SEPARATE words with no surrounding space ("capital—and", "it—than"). When the dash
 * and its neighbouring word are hidden, the two kept words fuse ("cand", "itthan"); exposing the
 * existing dash yields "c—and"/"it—than" instead (dds2 palmeri wet-test). The regular hyphen-minus
 * (U+002D) is EXCLUDED — word-INTERNAL ("term-locked"), not a separator. So despite the name, this
 * predicate is "characters that separate words and can be exposed to prevent fusion," NOT literal
 * whitespace. We always expose the EXISTING character (never insert/rewrite); reversibility lossless.
 */
// The word-separator predicate (the name is historical \u2014 see above; it is NOT literal whitespace).
// A char matches when it can SEPARATE two words and be EXPOSED in a hidden gap to keep them apart.
// THREE groups, finalized by the Stage B corpus audit (`node scripts/auditSeparators.mjs`):
//   1. `\p{Zs}` \u2014 EVERY Unicode Space_Separator: ASCII space, NBSP (U+00A0), and the figure/thin/
//      narrow/four-per-em/etc. spaces (U+2000\u2013200A, U+202F, U+205F, U+3000\u2026). This GENERALIZES the
//      old hand-listed subset (U+00A0/2007/2009/202F): the audit found a four-per-em space (U+2005,
//      "25 kg") the explicit list missed, so `\p{Zs}` closes the whole space class \u2014 no future space
//      variant can slip through and refuse the bridge.
//   2. ASCII control whitespace `\t\n\v\f\r` \u2014 Unicode category Cc, not Zs, so listed explicitly.
//   3. Non-space PROSE separators: EN/EM dash (U+2013/U+2014 \u2014 "capital\u2014and" fusion, dds2 palmeri
//      wet-test) and SLASH (U+002F \u2014 "and/or", "Good/Bad"; the audit's only unambiguous prose
//      separator among uncovered punctuation). We always expose the EXISTING char (lossless).
// DELIBERATELY EXCLUDED (Stage B audit \u2014 all word-INTERNAL; exposing/keeping them would mis-split or
// do nothing): the hyphen family \u2014 hyphen-minus U+002D ("E-Commerce"), Unicode hyphen U+2010
// ("non\u2010US"), non-breaking hyphen U+2011 ("Secretary\u2011General"), soft hyphen U+00AD (invisible,
// mid-word); underscore U+005F (identifiers/filenames); ampersand U+0026 ("R&D", "AT&T"); and the
// math/unit/URL symbols (= # $ \u00B0 + \u00D7 ^ \u2212 \u2026). Those JOIN tokens; they don't separate words.
const WHITESPACE = /[\t\n\v\f\r\u2013\u2014/]|\p{Zs}/u;

function isWhitespace(ch: string): boolean {
  return WHITESPACE.test(ch);
}

/**
 * Punctuation that hugs adjacent text, so the cross-gap bridge must NOT insert a space
 * at that junction. HUGS_LEFT chars attach to the text on their LEFT (closing marks:
 * ". People are" reads "myopic. People are", not "myopic . People are"); HUGS_RIGHT
 * chars attach to the text on their RIGHT (opening marks: "(" before "in house").
 */
const HUGS_LEFT = /[.,;:!?)\]}%”’]/;
const HUGS_RIGHT = /[(\[{“‘]/;

/**
 * Decide, per run, whether it is kept (true) or hidden (false) within a
 * non-heading, non-cite paragraph. We show ONLY highlighted text.
 *
 *   * Keep = the run is highlighted in a keep-color (#6c, #11). A highlight that
 *     covers only PART of a word no longer drags the unhighlighted remainder
 *     visible (decision #19 reversed per wet-test): "reduc"[hl]+"e" shows just
 *     "reduc"; an unhighlighted " society" glued to a highlighted ". People are"
 *     stays hidden. Highlight boundaries are already run boundaries, so each
 *     unhighlighted fragment is its own run and simply isn't kept — no sub-run
 *     splitting needed. Inter-chunk spacing is restored separately by the
 *     whitespace-only bridge below and `planCrossGapSeparators`.
 *   * Ineligible runs (fields, footnote refs, drawings, objects — decision #16)
 *     are always kept; we never hide structural content.
 */
export function computeRunKeepFlags(
  runs: readonly RunView[],
  keepColors: ReadonlySet<string>
): boolean[] {
  const keep = runs.map((r) => r.highlight !== null && keepColors.has(r.highlight));

  // Structural runs are never hidden.
  for (let i = 0; i < runs.length; i++) if (!runs[i].eligible) keep[i] = true;

  // Preserve inter-word spacing between kept runs. Because runs split on formatting
  // boundaries, the space between two highlighted words is often its OWN unhighlighted
  // run; hiding it collapses "reasons for" into "reasonsfor" in the condensed view. A
  // run that is ENTIRELY whitespace is kept when the nearest non-whitespace-only run on
  // BOTH sides is kept — i.e. it bridges two visible words. Whitespace adjacent to
  // hidden content (only one side kept, or neither) stays hidden, so no spurious
  // leading/trailing spaces leak into the visible text.
  const wsOnly = runs.map((r) => r.text.length > 0 && [...r.text].every(isWhitespace));
  for (let i = 0; i < runs.length; i++) {
    if (keep[i] || !wsOnly[i]) continue;
    let left = i - 1;
    while (left >= 0 && wsOnly[left]) left--;
    let right = i + 1;
    while (right < runs.length && wsOnly[right]) right++;
    if (left >= 0 && keep[left] && right < runs.length && keep[right]) keep[i] = true;
  }

  return keep;
}

/** A run made only of whitespace characters (a pure separator, not a word). */
function isWhitespaceOnly(text: string): boolean {
  return text.length > 0 && [...text].every(isWhitespace);
}

/** True when the run carries at least one non-whitespace (word) character. */
function hasWordChar(text: string): boolean {
  return [...text].some((c) => !isWhitespace(c));
}

/**
 * Bridge separators ACROSS hidden gaps (wet-test bug 1).
 *
 * `computeRunKeepFlags` already rescues a whitespace-ONLY run that sits directly
 * between two kept words (lesson #14). But when highlighted chunks are separated
 * by NON-highlighted prose, the separating space lives INSIDE that larger hidden
 * run (e.g. the run " are a constant threat, " between "ultraviolet radiation" and
 * "would") — there is no whitespace-only run to keep, so hiding the run fuses the
 * visible words into "radiationwould". This planner restores exactly one space per
 * fused gap.
 *
 * Reversibility is intrinsic: we NEVER insert characters. For each fused gap we
 * either keep an existing whitespace-only run (`extraKeep`) or emit a `BridgeSplit`
 * telling ooxml.ts to MOVE one existing space out of the hidden run into a visible
 * sibling (`splits`). The paragraph's concatenated text is therefore unchanged, so
 * Show All / native Font-dialog reversal restore the exact original, and Re-hide
 * re-derives identically — the moved space is then a whitespace-only run this very
 * function rescues via `extraKeep` (no re-splitting; convergent).
 *
 * @param runs the paragraph's runs in document order (readRuns order)
 * @param keep per-run keep flags from `computeRunKeepFlags`
 */
export function planCrossGapSeparators(
  runs: readonly RunView[],
  keep: readonly boolean[],
): { extraKeep: number[]; splits: BridgeSplit[] } {
  const extraKeep: number[] = [];
  const splits: BridgeSplit[] = [];

  // Anchors = kept runs that carry a visible word. A kept whitespace-only run is a
  // separator, not an anchor, so it does not start/close a gap.
  const anchors: number[] = [];
  for (let i = 0; i < runs.length; i++) {
    if (keep[i] && hasWordChar(runs[i].text)) anchors.push(i);
  }

  // Examine the gap between each pair of consecutive anchors.
  for (let p = 0; p + 1 < anchors.length; p++) {
    const a = anchors[p];
    const b = anchors[p + 1];
    if (b === a + 1) continue; // adjacent anchors: any needed space lives inside a or b

    // Bridge only a junction that would actually fuse. Skip when the gap is already
    // visually separated (a side ends/starts with whitespace) OR when punctuation hugs
    // the junction: a chunk that STARTS with a closing mark attaches to the left
    // ("myopic" + ". People are" → "myopic. People are", not "myopic . People are"),
    // and one that ENDS with an opening mark attaches to the right. Inserting a space
    // there would be wrong, so leave the punctuation flush.
    const aText = runs[a].text;
    const bText = runs[b].text;
    const lc = aText[aText.length - 1];
    const rc = bText[0];
    if (isWhitespace(lc) || isWhitespace(rc)) continue;
    if (HUGS_LEFT.test(rc) || HUGS_RIGHT.test(lc)) continue;

    // Already bridged by a kept whitespace-only run inside the gap? Leave it.
    let bridged = false;
    for (let g = a + 1; g < b; g++) {
      if (keep[g] && isWhitespaceOnly(runs[g].text)) {
        bridged = true;
        break;
      }
    }
    if (bridged) continue;

    // Priority 1: an unkept whitespace-only run exists → keep the first one (no split).
    let chosen = -1;
    for (let g = a + 1; g < b; g++) {
      if (isWhitespaceOnly(runs[g].text)) {
        chosen = g;
        break;
      }
    }
    if (chosen >= 0) {
      extraKeep.push(chosen);
      continue;
    }

    // Priority 2: split the first gap run that has a leading space.
    let done = false;
    for (let g = a + 1; g < b && !done; g++) {
      if (runs[g].text.startsWith(" ")) {
        splits.push({ index: g, side: "lead" });
        done = true;
      }
    }
    if (done) continue;

    // Priority 3: split the last gap run that has a trailing space.
    for (let g = b - 1; g > a && !done; g--) {
      if (runs[g].text.endsWith(" ")) {
        splits.push({ index: g, side: "trail" });
        done = true;
      }
    }
    if (done) continue;

    // Priority 4: no boundary space, but a gap run has an INTERIOR space. Expose the
    // first one via a 3-way split. With show-only-highlighted both fragments are
    // unhighlighted (hidden) wherever we split, and on Re-hide the exposed space is a
    // whitespace-only run that priority 1 rescues — so ANY interior space is
    // convergence-safe. Restricted to pure-text runs so ooxml's single-<w:t> split is
    // structurally safe. (Fixes "reduce extinction" -> "reduc x", not "reducx", when
    // only "reduc" and "x" are highlighted and the gap is one run "e e".)
    for (let g = a + 1; g < b && !done; g++) {
      const t = runs[g].text;
      if (t.includes("\t") || t.includes("\n")) continue;
      const k = t.indexOf(" ");
      if (k > 0 && k < t.length - 1) {
        splits.push({ index: g, side: "interior", offset: k });
        done = true;
      }
    }
    // else: no exposable space anywhere in the gap → leave the words fused, faithful
    // to a source that genuinely had no space there (e.g. "A"+hidden"/"+"B").
  }

  return { extraKeep, splits };
}
