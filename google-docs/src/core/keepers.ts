// Pure keeper rules — the heart of the policy (plan rows 3-5, A8/A9/A10), the
// Docs sibling of the Word engine's src/core/keepers.ts.
//
// A paragraph/element is KEPT iff, evaluated IN THIS ORDER:
//   (a) its paragraph is inside a table                      (row 1 — untouched), OR
//   (b) its paragraph is a kept heading style                (row 3), OR
//   (c) its paragraph is a cite — by SIGNATURE or STRUCTURE  (rows 4 + A10), OR
//   (d) the element is non-text ("other")                    (whitelist, A9), OR
//   (e) the element is highlighted in a keeper color         (row 5 + A8), OR
//   (f) the element is whitespace-only bridging two keeps    (Word lesson #14).
// Everything else hides. (a)-(c) keep the WHOLE paragraph including its
// trailing newline; (d)-(f) are per-element, and an element's keep flag covers
// its full range — the trailing newline living inside a paragraph's final text
// element simply follows that element's fate (plan D6; the planner owns the
// one API-hard exception, the segment-final clamp).
//
// PURITY CONTRACT: total, deterministic functions over the GIVEN views. No
// Apps Script, no I/O, no awareness of sentinels or rstm ranges — callers
// (Hide's reconcile, plan A1) pass a RESTORED view whose sizes have already
// been reconciled, so the cite-size trap (a hidden 1pt cite failing the
// >= citeMinPt test) cannot reach this module by construction.

import { NEAR_WHITE_MIN_CHANNEL } from "./constants";
import { GDoc, GdocsSettings, GElement, GNamedStyleType, GParagraph, GRange } from "./types";

// ---------------------------------------------------------------------------
// Public contract
// ---------------------------------------------------------------------------

/** The keeper verdict for one paragraph, consumed by the planner. */
export interface ParagraphKeep {
  /** True = the whole paragraph (trailing newline included) is kept; the
   * per-element flags are then trivially all-true. */
  keepWhole: boolean;
  /** Parallel to GParagraph.elements: true = that element's range is kept. */
  elementKeep: boolean[];
  /** True only when the CITE rule (signature or structural) is what kept the
   * paragraph — receipts and diagnostics report cite counts separately. */
  citeDetected: boolean;
}

// ---------------------------------------------------------------------------
// Heading rule (row 3)
// ---------------------------------------------------------------------------

/**
 * Named styles kept whole. Word keeps outline levels 0-3 (Heading 1-4); Docs
 * has no outline levels, so we key on the named style. TITLE and SUBTITLE are
 * a DELIBERATE over-keep divergence from Word's outline-0-3: they sit above
 * H1 in Docs' style ladder and are visually structural (a doc title hidden to
 * 1pt would look like data loss), so conservatism keeps them. HEADING_5/6 are
 * BODY — debate templates use H1-4 (pocket/hat/block/tag) and Word's level-4+
 * styles were likewise hidden.
 */
const KEEP_WHOLE_STYLES: ReadonlySet<GNamedStyleType> = new Set([
  "HEADING_1",
  "HEADING_2",
  "HEADING_3",
  "HEADING_4",
  "TITLE",
  "SUBTITLE"
]);

/** Heading rule: keep the paragraph whole when its named style is structural. */
function isHeadingKeptStyle(style: GNamedStyleType): boolean {
  return KEEP_WHOLE_STYLES.has(style);
}

// ---------------------------------------------------------------------------
// Highlight rule (row 5 + plan A8)
// ---------------------------------------------------------------------------

/**
 * Min RGB channel (0-255) of a "#rrggbb" hex, or null when the string is not
 * in the contract shape. parse.ts guarantees lower-case "#rrggbb" | null, so
 * the null branch is purely defensive — but the predicate below must still
 * decide something for it, and the decision is policy, so it lives here.
 */
function minChannel(hex: string): number | null {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/.exec(hex);
  if (m === null) return null;
  return Math.min(parseInt(m[1], 16), parseInt(m[2], 16), parseInt(m[3], 16));
}

/**
 * Highlight keeper (plan A8). Docs has no separate highlight channel —
 * highlight IS backgroundColor — so the default mode is a CLOSED set
 * ("set": exact membership in settings.keepColors, Word keep-colors mental
 * model; membership is deliberately literal, with NO near-white filter, so a
 * user who explicitly checks an odd color gets exactly that color).
 * "anyHighlight" is the explicit master toggle: ANY background keeps EXCEPT
 * near-white — web-pasted evidence carries shading like #f8f9fa that must
 * HIDE, or the open predicate would defeat Hide on every pasted card.
 * An unparseable non-null hex in "anyHighlight" mode KEEPS: we cannot prove
 * it near-white, and the house failure direction is over-keep (never hide a
 * highlight the user may have made — Word decision #16's conservatism).
 */
export function isHighlightKept(backgroundHex: string | null, settings: GdocsSettings): boolean {
  if (backgroundHex === null) return false;
  if (settings.keepMode === "set") return settings.keepColors.has(backgroundHex);
  const min = minChannel(backgroundHex);
  if (min === null) return true; // defensive over-keep (see contract above)
  return min < NEAR_WHITE_MIN_CHANNEL;
}

// ---------------------------------------------------------------------------
// Whitespace bridging (Word lesson #14, deliberately simplified)
// ---------------------------------------------------------------------------

/**
 * Whitespace for the bridging rescue: the Unicode space class (\p{Zs} — ASCII
 * space, NBSP, thin/figure/narrow spaces pasted text is full of) plus the
 * ASCII control whitespace (category Cc, so listed explicitly). The Word
 * engine's separator predicate ALSO exposed em/en dashes and slashes and ran
 * a whole cross-gap bridge planner — NONE of that is ported, on purpose:
 * Word's hide removed glyphs from the visible flow, so a hidden separator
 * fused its neighbors ("risingrevenues"). Under SHRINK the 1pt glyph still
 * occupies the line and physically separates kept words (plan D7), so fusion
 * is not fatal here. We keep only the cheap whitespace-only rescue because it
 * costs nothing and makes the kept view read naturally at full size.
 */
const WHITESPACE = /[\t\n\v\f\r]|\p{Zs}/u;

/** A text element made only of whitespace (a pure separator, not a word). */
function isWhitespaceOnly(text: string): boolean {
  return text.length > 0 && [...text].every((c) => WHITESPACE.test(c));
}

// ---------------------------------------------------------------------------
// Cite signature rule (row 4)
// ---------------------------------------------------------------------------

/**
 * Docs' factory Normal size — the last-resort fallback when an element
 * inherits its size AND the read carried no namedStyles entry for its style
 * (a defensive double-inherit; real reads state Normal's size). 11 because
 * that is what every Docs account ships with.
 */
const DOCS_DEFAULT_BODY_PT = 11;

/**
 * The size the cite predicate sees for an element: its explicit size, else
 * the paragraph's named-style size, else the Docs default. Resolution is
 * REQUIRED (not a nicety): Docs-native cites routinely inherit their size, so
 * testing fontSizePt alone would silently fail the >= citeMinPt comparison
 * for every inherited run.
 */
function resolvedSizePt(element: GElement, paragraph: GParagraph, doc: GDoc): number {
  return element.fontSizePt ?? doc.namedStyleSizesPt[paragraph.namedStyleType] ?? DOCS_DEFAULT_BODY_PT;
}

/**
 * Signature cite (row 4): any TEXT element bold AND resolved >= citeMinPt
 * keeps the paragraph whole. Applies to EVERY paragraph not already kept
 * whole — not just NORMAL_TEXT (plan A10): a cite living in an H5/H6
 * paragraph (body styles) must still be kept. Only kind "text" participates
 * (whitelist, A9); "other" elements are kept structurally regardless.
 */
function paragraphHasCiteSignature(paragraph: GParagraph, doc: GDoc, settings: GdocsSettings): boolean {
  return paragraph.elements.some(
    (e) => e.kind === "text" && e.bold && resolvedSizePt(e, paragraph, doc) >= settings.citeMinPt
  );
}

// ---------------------------------------------------------------------------
// Structural cite rule (plan A10 — the ported citeRepair heuristic)
//
// WHY THIS EXISTS (the loop-001 plan-review BLOCKER): Docs-native debate docs
// never ran Apply-styles or Mark-cite — their cite lines are bold at the 11pt
// default, so bold AND >= 13 HIDES them. That is the exact "leak in the wrong
// direction" the Word product engineered citeRepair against. The heuristic is
// ported with its conservatism intact: a FALSE POSITIVE keeps body prose that
// should hide, so every gate errs toward "not a cite".
// ---------------------------------------------------------------------------

/** Leading chars of the paragraph text searched for a year — a year buried
 * deep in prose is NOT a cite signal (ported from citeRepair). */
const YEAR_SEARCH_WINDOW = 80;

/** Max start offset of the FIRST bold element (cites LEAD with the author). */
const FIRST_BOLD_MAX_OFFSET = 30;

/** Max length of the non-bold prefix before the first bold element. */
const NAME_PREFIX_MAX_LEN = 30;

/**
 * A YEAR token: optional apostrophe (straight or curly — "'20"/"’20" debate
 * shorthand) then 2-4 digits not followed by another digit. A candidate whose
 * digits are PRECEDED by a digit is rejected manually in firstYearIndex (the
 * tail of a longer number — page counts, DOIs), replacing a lookbehind for
 * ES2019 portability. Ported verbatim from the Word heuristic.
 */
const YEAR_RE_SOURCE = "['’]?[0-9]{2,4}(?![0-9])";

/**
 * Char index of the first YEAR token within the search window, or -1. A fresh
 * regex per call keeps the global lastIndex from bleeding between callers.
 */
function firstYearIndex(text: string): number {
  const head = text.slice(0, YEAR_SEARCH_WINDOW);
  const re = new RegExp(YEAR_RE_SOURCE, "g");
  let m: RegExpExecArray | null;
  while ((m = re.exec(head)) !== null) {
    // First DIGIT of the token (skip the optional apostrophe).
    const digitStart = m[0][0] === "'" || m[0][0] === "’" ? m.index + 1 : m.index;
    const before = digitStart > 0 ? head[digitStart - 1] : "";
    if (!/[0-9]/.test(before)) return m.index; // not the tail of a longer number
    re.lastIndex = m.index + 1; // advance past the false match
  }
  return -1;
}

/** Sentence/clause punctuation that disqualifies a "name-like" prefix. */
const PREFIX_PUNCT_RE = /[,.;:!?]/;

/**
 * Is the text BEFORE the first bold element "name-like"? A real cite's lead
 * prefix is a bare author-name fragment ("Barbara ") — short, digit-free,
 * free of clause punctuation. Rejects prose like "In 2008, the crisis ..."
 * whose first bold word sits mid-sentence. Empty prefix trivially qualifies.
 */
function prefixIsNameLike(prefix: string): boolean {
  if (prefix.length === 0) return true;
  if (prefix.length > NAME_PREFIX_MAX_LEN) return false;
  if (/[0-9]/.test(prefix)) return false;
  if (PREFIX_PUNCT_RE.test(prefix)) return false;
  return true;
}

/** Per-text-element facts for one candidate paragraph: bold flag, visible
 * text, and the element's start offset within the concatenated paragraph text
 * (offsets locate the first bold element and the year). "other" elements are
 * skipped — they contribute no visible text and can never be a cite lead. */
interface LeadFacts {
  element: GElement;
  bold: boolean;
  text: string;
  startOffset: number;
}

/** Compute LeadFacts for a paragraph's text elements, in document order. */
function leadFactsOf(paragraph: GParagraph): LeadFacts[] {
  const facts: LeadFacts[] = [];
  let offset = 0;
  for (const e of paragraph.elements) {
    if (e.kind !== "text") continue;
    facts.push({ element: e, bold: e.bold, text: e.text, startOffset: offset });
    offset += e.text.length;
  }
  return facts;
}

/** The concatenated visible text of a paragraph (text elements only). */
function paragraphVisibleText(paragraph: GParagraph): string {
  let out = "";
  for (const e of paragraph.elements) if (e.kind === "text") out += e.text;
  return out;
}

/**
 * Evaluate ONE candidate paragraph against the ported gates (a)-(e); returns
 * its lead elements when it qualifies, else null. ALL gates required —
 * conservative, a miss means "not a cite":
 *   (a) no element already matches the cite SIGNATURE (already kept/at the
 *       convention → nothing to detect or repair);
 *   (b) at least one bold text element exists (cites bold the author);
 *   (c) a YEAR token appears within the first 80 chars;
 *   (d) the first bold element starts within the first 30 chars;
 *   (e) the prefix before the first bold element is name-like.
 * The lead = every bold element up to and including the one containing the
 * year (captures both "Valcke et al. 20" in one element and split
 * author/year elements); fallback to the single first bold element so a
 * qualifying candidate always yields >= 1 lead.
 */
function evaluateCandidate(
  paragraph: GParagraph,
  doc: GDoc,
  settings: GdocsSettings
): GElement[] | null {
  // (a) already a signature cite — the keeper keeps it without us, and
  // Apply-styles has nothing to repair.
  if (paragraphHasCiteSignature(paragraph, doc, settings)) return null;
  const facts = leadFactsOf(paragraph);
  // (b) needs a bold element.
  const firstBold = facts.findIndex((f) => f.bold);
  if (firstBold < 0) return null;
  const fullText = facts.map((f) => f.text).join("");
  // (c) a year near the front.
  const yearIndex = firstYearIndex(fullText);
  if (yearIndex < 0) return null;
  // (d) the first bold element leads.
  const firstBoldOffset = facts[firstBold].startOffset;
  if (firstBoldOffset > FIRST_BOLD_MAX_OFFSET) return null;
  // (e) the prefix before it is name-like.
  if (!prefixIsNameLike(fullText.slice(0, firstBoldOffset))) return null;

  // The element whose text span contains the year token (default: last).
  let yearElement = facts.length - 1;
  for (let i = 0; i < facts.length; i++) {
    const start = facts[i].startOffset;
    if (yearIndex >= start && yearIndex < start + facts[i].text.length) {
      yearElement = i;
      break;
    }
  }
  const leads: GElement[] = [];
  for (let i = 0; i <= yearElement && i < facts.length; i++) {
    if (facts[i].bold) leads.push(facts[i].element);
  }
  // Defensive fallback (mirrors the Word port): a qualifying candidate must
  // yield at least its first bold element.
  if (leads.length === 0) leads.push(facts[firstBold].element);
  return leads;
}

/** One detected cite lead: the candidate paragraph and its lead elements. */
interface CiteLeadCandidate {
  paragraph: GParagraph;
  paragraphIndex: number;
  leadElements: GElement[];
}

/**
 * The shared structural detection both planKeeps and detectCiteLeads consume
 * (single source so the keeper and the Apply-styles repair can never disagree
 * about what a cite lead is). Ported scan-and-pick from citeRepair:
 *
 *   Each KEPT heading opens a window (Word opened one from ANY heading; we
 *   narrow to kept headings per A10 — H5/6 are body here and never anchor a
 *   cite slot). Scanning forward: empty paragraphs are SKIPPED; any
 *   non-NORMAL_TEXT paragraph BREAKS the window (a deeper tag supersedes — the
 *   outer loop opens its own window from it if kept); a table BREAKS the
 *   window (tables are untouched territory, never scanned through). The first
 *   non-empty NORMAL_TEXT paragraph is THE candidate, then the window closes —
 *   a cite "a few lines down" is deliberately not handled, so a "Smith 2020"
 *   inside body prose is never mistaken for the cite.
 *
 * PORT GAP, documented: Word also skipped LIST paragraphs (bullets under the
 * tag, <w:numPr>). GParagraph carries no list/bullet flag, so a bulleted line
 * here becomes the candidate and (lacking a bold Author-YEAR lead) fails the
 * gates — conservative in the false-positive direction, but a cite BELOW
 * bullets goes undetected. Flagged for the parse contract's v2.
 */
function findCiteLeadCandidates(doc: GDoc, settings: GdocsSettings): CiteLeadCandidate[] {
  const out: CiteLeadCandidate[] = [];
  // Two adjacent kept headings can't reach the same candidate (the first
  // window breaks at the second heading), but the dedupe set stays as a
  // defensive rail mirroring the Word port.
  const planned = new Set<number>();
  const paras = doc.paragraphs;
  for (let t = 0; t < paras.length; t++) {
    // Only a kept heading OUTSIDE a table opens a cite window.
    if (paras[t].inTable || !isHeadingKeptStyle(paras[t].namedStyleType)) continue;
    let candidate = -1;
    for (let j = t + 1; j < paras.length; j++) {
      const p = paras[j];
      if (p.inTable) break; // never scan through a table
      if (p.namedStyleType !== "NORMAL_TEXT") break; // superseded by the next structural paragraph
      if (paragraphVisibleText(p).trim().length === 0) continue; // skip empties
      candidate = j;
      break;
    }
    if (candidate < 0 || planned.has(candidate)) continue;
    const leadElements = evaluateCandidate(paras[candidate], doc, settings);
    if (leadElements !== null) {
      out.push({ paragraph: paras[candidate], paragraphIndex: candidate, leadElements });
      planned.add(candidate);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Public: planKeeps
// ---------------------------------------------------------------------------

/** All-true element flags for a whole-kept paragraph (explicit so the planner
 * can treat elementKeep as total without special-casing keepWhole). */
function keepAll(paragraph: GParagraph, citeDetected: boolean): ParagraphKeep {
  return { keepWhole: true, elementKeep: paragraph.elements.map(() => true), citeDetected };
}

/**
 * Per-element verdicts for a paragraph that is NOT kept whole:
 *   * kind "other" keeps (whitelist, A9 — chips/breaks/objects are never
 *     style-targeted, so they are structurally kept);
 *   * a keeper-highlighted element keeps (row 5);
 *   * a whitespace-only element keeps iff the nearest non-whitespace-only
 *     neighbor on BOTH sides keeps (lesson #14: the space between two
 *     highlighted words is often its own unhighlighted element; rescuing it
 *     keeps the visible view reading naturally). Whitespace adjacent to
 *     hidden content stays hidden — no leading/trailing space floats at full
 *     size next to shrunken text.
 * Everything else hides. The paragraph's trailing newline lives inside its
 * final text element and simply follows that element's flag (plan D6).
 */
function perElementKeep(paragraph: GParagraph, settings: GdocsSettings): ParagraphKeep {
  const els = paragraph.elements;
  const keep = els.map((e) => e.kind === "other" || isHighlightKept(e.backgroundHex, settings));

  // Whitespace bridge: scan past consecutive whitespace-only elements so a
  // multi-element gap (" " + "\t") between two keeps is rescued as a unit.
  const wsOnly = els.map((e) => e.kind === "text" && isWhitespaceOnly(e.text));
  for (let i = 0; i < els.length; i++) {
    if (keep[i] || !wsOnly[i]) continue;
    let left = i - 1;
    while (left >= 0 && wsOnly[left]) left--;
    let right = i + 1;
    while (right < els.length && wsOnly[right]) right++;
    if (left >= 0 && keep[left] && right < els.length && keep[right]) keep[i] = true;
  }

  return { keepWhole: false, elementKeep: keep, citeDetected: false };
}

/**
 * The keeper policy, whole-doc: one ParagraphKeep per paragraph, in order.
 * PURE over the GIVEN views — callers pass a RESTORED view (plan A1), so
 * every size this module compares is an original size, never a sentinel.
 *
 * Rule order is load-bearing: table and heading keeps win before the cite
 * predicates run (their citeDetected stays false — a bold 26pt pocket heading
 * is a heading, not a cite), and the structural rule only matters for
 * paragraphs the signature missed.
 */
export function planKeeps(doc: GDoc, settings: GdocsSettings): ParagraphKeep[] {
  // Structural detection is doc-global (it needs the heading→candidate
  // geometry), so compute the kept set once, not per paragraph. The flag
  // gates the KEEPER rule only — detectCiteLeads below stays ungated.
  const structurallyKept = new Set<number>();
  if (settings.structuralCite) {
    for (const c of findCiteLeadCandidates(doc, settings)) structurallyKept.add(c.paragraphIndex);
  }

  return doc.paragraphs.map((p, i) => {
    if (p.inTable) return keepAll(p, false); // row 1: tables untouched
    if (isHeadingKeptStyle(p.namedStyleType)) return keepAll(p, false); // row 3
    if (paragraphHasCiteSignature(p, doc, settings)) return keepAll(p, true); // row 4
    if (structurallyKept.has(i)) return keepAll(p, true); // A10 structural
    return perElementKeep(p, settings); // rows 2/5 + lesson #14
  });
}

// ---------------------------------------------------------------------------
// Public: detectCiteLeads (for styles.ts — plan A10(b))
// ---------------------------------------------------------------------------

/**
 * The exact ranges Apply-styles' bulk cite-repair writes the convention onto
 * (bold CITE_PT) — the SAME structural detection the keeper uses, one GRange
 * per lead element in document order. Deliberately NOT gated on
 * settings.structuralCite: the repair pass is Word's Apply-Styles parity and
 * must work even when the keeper fallback flag is off (repairing the lead to
 * the convention is precisely what makes the SIGNATURE rule keep it
 * afterwards). Gate (a) inside the detection already skips leads that are
 * at/above the signature threshold, so repair stays minimal.
 *
 * Ranges follow element boundaries — including a paragraph-final newline
 * riding inside a fully-bold lead element (sizing a paragraph mark is benign
 * and matches what Mark-cite on a selected line does) — EXCEPT the
 * segment-final newline, which the API refuses to style and is clamped off
 * here so styles.ts can emit these ranges verbatim.
 */
export function detectCiteLeads(doc: GDoc, settings: GdocsSettings): GRange[] {
  const ranges: GRange[] = [];
  for (const c of findCiteLeadCandidates(doc, settings)) {
    const lastElement = c.paragraph.elements[c.paragraph.elements.length - 1];
    for (const el of c.leadElements) {
      const clamp = c.paragraph.isLastInSegment && el === lastElement;
      const endIndex = clamp ? el.endIndex - 1 : el.endIndex;
      // A clamp can only empty a range if the lead were a lone newline, which
      // the gates exclude — but a zero-length range would 400, so guard it.
      if (endIndex > el.startIndex) ranges.push({ startIndex: el.startIndex, endIndex });
    }
  }
  return ranges;
}
