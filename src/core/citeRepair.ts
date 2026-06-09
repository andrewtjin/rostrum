// Pure, host-free CITE REPAIR — detect debate cites that LOST their cite character
// style and re-apply it, so the keeper rule (`paragraphHasCiteRun`) keeps the card.
//
// WHY THIS EXISTS. Rostrum keeps a paragraph as a cite iff some run carries the cite
// character style (`<w:rStyle w:val="Style13ptBold"/>`). Real docs contain cites
// that were typed/pasted with the author bolded but WITHOUT that character style — e.g.
// the ndca sample's "Valcke et al. 20" cite, whose author run is `<w:b/>` bold but has
// no rStyle. Such a cite is mis-classified as body and HIDDEN with the card body, which
// is an invisibility LEAK in the wrong direction (it hides content that should stay).
// This module finds those mis-styled cites and plans the minimal rStyle injection that
// makes them kept again.
//
// LEAK-SAFETY IS THE WHOLE POINT (see `planCiteRepairs`). A FALSE POSITIVE — applying
// the cite style to a paragraph that is actually body prose — would KEEP body text that
// should be hidden, i.e. a real invisibility leak. So every gate here is CONSERVATIVE:
// when in doubt, do NOT repair. We only ever look at the SINGLE first real paragraph
// after a tag/heading (the structural slot a cite occupies), and only repair it when it
// looks unmistakably like a "Author YEAR" cite lead.
//
// NO OFFICE.JS, NO SHARED READ PATH. We operate DIRECTLY on each paragraph's OOXML
// string with regexes — the same string-scan approach `outline.ts` uses on paragraph
// XML — rather than the `RunView`/`readRuns` path, so this is a self-contained pure
// transform that runs identically in Node tests and the task-pane browser. The only
// mutation (`applyCiteStyleToParagraphXml`) adds/replaces an `<w:rStyle>` and touches
// nothing else, preserving the reversibility guarantee the rest of the engine relies on.

import { CITE_STYLE_ID } from "./styles";

// ---------------------------------------------------------------------------
// Public interfaces
// ---------------------------------------------------------------------------

/**
 * A lightweight paragraph view for cite-repair planning: the paragraph's raw OOXML
 * plus its resolved 0-based outline level (0 = Heading 1 … 8 = Heading 9, null = body)
 * — exactly what `WholeBodyPackage.headingLevel(i)` returns. We carry the raw `xml`
 * (not parsed runs) because the planner inspects bold/text/year structure with regexes.
 */
export interface CiteRepairParagraph {
  xml: string;
  headingLevel: number | null;
}

/**
 * One planned repair: the indices (0-based, in document order among that paragraph's
 * `<w:r>` elements) whose rPr should get the cite rStyle so the paragraph is kept.
 */
export interface CiteRepair {
  paragraphIndex: number;
  runIndices: number[];
}

// ---------------------------------------------------------------------------
// Run parsing (regex-based, document order — mirrors outline.ts's string scan)
// ---------------------------------------------------------------------------

/**
 * One `<w:r>` element of a paragraph, captured as raw XML in document order. `inner`
 * is the run's content between its start/end tags ("" for a self-closing `<w:r/>`),
 * from which we derive the run's own rPr, bold flag, and visible text.
 */
interface RawRun {
  /** The full `<w:r …>…</w:r>` (or `<w:r/>`) substring. */
  full: string;
  /** The content between the run's tags ("" for self-closing). */
  inner: string;
}

/**
 * Match every `<w:r>` in document order: either a self-closing `<w:r/>`/`<w:r …/>`, or
 * a paired `<w:r …>…</w:r>`. Runs never nest runs, so a non-greedy body capture is
 * unambiguous; hyperlink-wrapped runs (`<w:hyperlink><w:r>…</w:r></w:hyperlink>`) are
 * still matched in order, matching `readRuns`'s `getElementsByTagName("w:r")` ordering.
 * The `[^>]*` tolerates run attributes (`w:rsidRPr`, etc.) before the tag closes.
 */
const RUN_RE = /<w:r(?:\s[^>]*)?(?:\/>|>([\s\S]*?)<\/w:r>)/g;

/** Parse a paragraph's runs (document order). Self-closing runs get inner === "". */
function parseRuns(paragraphXml: string): RawRun[] {
  const runs: RawRun[] = [];
  let m: RegExpExecArray | null;
  // Reset lastIndex implicitly by constructing a fresh exec loop on a global regex; we
  // use a local copy of the source so concurrent callers can't clobber lastIndex.
  const re = new RegExp(RUN_RE.source, "g");
  while ((m = re.exec(paragraphXml)) !== null) {
    runs.push({ full: m[0], inner: m[1] ?? "" });
  }
  return runs;
}

/**
 * The run's OWN `<w:rPr>…</w:rPr>` block, or "" when it has none. A run's rPr is the
 * FIRST element of its content (OOXML CT_R ordering), so the first `<w:rPr>` in `inner`
 * is unambiguously the run's. Also tolerates a self-closing `<w:rPr/>` (no properties).
 */
function runRPr(inner: string): string {
  const paired = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(inner);
  if (paired) return paired[0];
  const selfClosing = /<w:rPr\b[^>]*\/>/.exec(inner);
  return selfClosing ? selfClosing[0] : "";
}

/**
 * True when a run's rPr marks it BOLD. Bold = `<w:b/>` or `<w:b w:val="true|1|on"/>`.
 * Explicitly NOT bold: `<w:b w:val="false|0|none"/>` (an override that turns bold off).
 * `<w:bCs/>` (complex-script bold) is deliberately ignored — it doesn't bold Latin
 * author names, and treating it as bold would mis-detect non-bold prose. Matching the
 * cite's real shape (`<w:rPr><w:b/><w:bCs/>…`), `<w:b/>` self-closing is the common case.
 */
function rPrIsBold(rPr: string): boolean {
  // `<w:b>` but NOT `<w:bCs>` — a negative lookahead on the next char rejects `bCs`.
  const b = /<w:b(?![A-Za-z])[^>]*>/.exec(rPr);
  if (!b) return false;
  const valMatch = /\bw:val="([^"]*)"/.exec(b[0]);
  if (!valMatch) return true; // bare <w:b/> ⇒ bold on
  const v = valMatch[1].toLowerCase();
  return !(v === "false" || v === "0" || v === "off" || v === "none");
}

/** True when a run's rPr already carries the cite character style. */
function rPrHasCiteStyle(rPr: string, citeStyleId: string): boolean {
  const m = /<w:rStyle\b[^>]*\bw:val="([^"]*)"/.exec(rPr);
  return m !== null && m[1] === citeStyleId;
}

/** Decode the XML entities that appear in `<w:t>` content so text scans see real chars. */
function decodeEntities(s: string): string {
  return s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
    .replace(/&amp;/g, "&");
}

/**
 * The visible text of a run, in document order: `<w:t>` content (entity-decoded),
 * `<w:tab>`→"\t", `<w:br>`/`<w:cr>`→"\n" — mirroring `ooxml.ts`'s `runText`. Used to
 * concatenate the paragraph's text for the year/prefix heuristics.
 */
function runText(inner: string): string {
  let out = "";
  // Walk the run content left-to-right, emitting text for each meaningful element. A
  // single scan over the relevant tags keeps this dependency-free (no DOM) and order-
  // preserving, which the prefix/year offsets depend on.
  const re = /<w:t\b[^>]*>([\s\S]*?)<\/w:t>|<w:t\b[^>]*\/>|<w:tab\b[^>]*\/?>|<w:br\b[^>]*\/?>|<w:cr\b[^>]*\/?>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(inner)) !== null) {
    const tag = m[0];
    if (tag.startsWith("<w:t") && m[1] !== undefined) out += decodeEntities(m[1]);
    else if (tag.startsWith("<w:tab")) out += "\t";
    else if (tag.startsWith("<w:br") || tag.startsWith("<w:cr")) out += "\n";
    // a self-closing <w:t/> contributes nothing
  }
  return out;
}

// ---------------------------------------------------------------------------
// Cite-detection heuristics (the leak-safety surface)
// ---------------------------------------------------------------------------

/** A paragraph that is a list item (`<w:numPr>` in its pPr) — a bullet under the tag. */
function isListParagraph(paragraphXml: string): boolean {
  return /<w:numPr\b/.test(paragraphXml);
}

/**
 * The number of leading characters of the concatenated paragraph text to search for a
 * year. Cites lead with "Author YEAR"; a year buried deep in prose is NOT a cite signal.
 */
const YEAR_SEARCH_WINDOW = 80;

/** The maximum start offset of the FIRST bold run for a cite (cites lead with the author). */
const FIRST_BOLD_MAX_OFFSET = 30;

/** The maximum length of the (non-bold) prefix before the first bold run for a name-like lead. */
const NAME_PREFIX_MAX_LEN = 30;

/**
 * A YEAR token: an optional leading apostrophe (straight `'` or curly `’`) then 2–4
 * digits NOT followed by another digit (so it isn't the head of a longer number). We
 * also reject a token PRECEDED by a digit (checked manually below, not via a lookbehind,
 * to stay portable at the ES2019 target) so "20" in "Valcke et al. 20" matches but the
 * "12696" / "10.1111" digits inside a DOI do not, and "1990s"/"2020-2021" boundaries
 * behave. The apostrophe forms ("'20", "’20") are common debate shorthand for a year.
 * Global so we can advance past a digit-preceded false match to the next candidate.
 */
const YEAR_RE = /['’]?[0-9]{2,4}(?![0-9])/g;

/**
 * The character index of the first YEAR token in `text` within the first
 * `YEAR_SEARCH_WINDOW` chars, or -1 when none. We scan only the window so a year deep in
 * the card text (a body sentence "… in 2020 …") can't qualify a non-cite paragraph. A
 * candidate whose digit run is PRECEDED by a digit is skipped (it's the tail of a longer
 * number) — the manual preceding-char check replaces an ES2018 lookbehind for portability.
 */
function firstYearIndex(text: string): number {
  const head = text.slice(0, YEAR_SEARCH_WINDOW);
  const re = new RegExp(YEAR_RE.source, "g");
  let m: RegExpExecArray | null;
  while ((m = re.exec(head)) !== null) {
    // The first DIGIT of the token (after an optional apostrophe).
    const digitStart = m[0][0] === "'" || m[0][0] === "’" ? m.index + 1 : m.index;
    const before = digitStart > 0 ? head[digitStart - 1] : "";
    if (!/[0-9]/.test(before)) return m.index; // not the tail of a longer number
    // Otherwise advance one char past this match's start to find the next candidate.
    re.lastIndex = m.index + 1;
  }
  return -1;
}

/** Sentence/clause punctuation that disqualifies a "name-like" prefix (a real cite leads with a bare name). */
const PREFIX_PUNCT_RE = /[,.;:!?]/;

/**
 * Is the text BEFORE the first bold run "name-like"? (gate (e)). A real cite's lead is a
 * bare author name fragment ("Barbara "), so the prefix must be short, digit-free, and
 * free of sentence/clause punctuation — which rejects body prose like
 * "In 2008, the crisis worsened and " whose prefix has a digit and a comma. An EMPTY
 * prefix (the bold run is first) trivially qualifies.
 */
function prefixIsNameLike(prefix: string): boolean {
  if (prefix.length === 0) return true;
  if (prefix.length > NAME_PREFIX_MAX_LEN) return false;
  if (/[0-9]/.test(prefix)) return false;
  if (PREFIX_PUNCT_RE.test(prefix)) return false;
  return true;
}

/**
 * Per-run facts the planner needs, computed once per candidate paragraph: whether the
 * run is bold, whether it already has the cite style, its visible text, and the running
 * START OFFSET of the run within the paragraph's concatenated text (so we can locate the
 * first bold run and the year relative to text positions).
 */
interface RunFacts {
  bold: boolean;
  cite: boolean;
  text: string;
  startOffset: number;
}

/** Compute `RunFacts` for every run of a paragraph (document order, with text offsets). */
function runFactsOf(paragraphXml: string, citeStyleId: string): RunFacts[] {
  const runs = parseRuns(paragraphXml);
  const facts: RunFacts[] = [];
  let offset = 0;
  for (const r of runs) {
    const rPr = runRPr(r.inner);
    const text = runText(r.inner);
    facts.push({
      bold: rPrIsBold(rPr),
      cite: rPrHasCiteStyle(rPr, citeStyleId),
      text,
      startOffset: offset
    });
    offset += text.length;
  }
  return facts;
}

/**
 * Plan the run repair for a paragraph that ALREADY passed the candidate gates: the runs
 * to restyle are every BOLD run whose index is ≤ the index of the run CONTAINING the
 * first year token. This captures "Valcke et al. 20" (one bold run holding the year) and
 * "Barbara <b>Valcke</b> … <b>20</b>" (the non-bold "Barbara" is skipped; both bold runs
 * up to and including the year run are styled). If that set is empty for any reason but
 * the candidate qualified, fall back to the single first bold run so the paragraph is
 * guaranteed ≥1 cite-styled run (and is therefore kept). Returns the run indices.
 */
function planRunIndices(facts: RunFacts[], yearIndex: number): number[] {
  // The index of the run whose text span contains the first year token.
  let yearRun = facts.length - 1;
  for (let i = 0; i < facts.length; i++) {
    const start = facts[i].startOffset;
    const end = start + facts[i].text.length;
    if (yearIndex >= start && yearIndex < end) {
      yearRun = i;
      break;
    }
  }
  const indices: number[] = [];
  for (let i = 0; i < facts.length && i <= yearRun; i++) {
    if (facts[i].bold) indices.push(i);
  }
  if (indices.length === 0) {
    // Guarantee ≥1 styled run for a qualifying candidate: the first bold run anywhere.
    const firstBold = facts.findIndex((f) => f.bold);
    if (firstBold >= 0) indices.push(firstBold);
  }
  return indices;
}

/**
 * Evaluate ONE candidate paragraph against gates (a)–(e) and, if it qualifies, return
 * the run indices to repair; otherwise null. Split out so `planCiteRepairs` reads as the
 * scan-and-pick structure while this holds the leak-safety predicate in one place.
 *
 * Gates (ALL required; conservative — a miss means "do not repair"):
 *   (a) NO run already has the cite style (else it's a proper cite → nothing to do).
 *   (b) at least one BOLD run exists (cites bold the author).
 *   (c) a YEAR token appears within the first 80 chars of the concatenated text.
 *   (d) the FIRST bold run starts within the first 30 chars (cites lead with the author).
 *   (e) the text before the first bold run is "name-like" (short, no digit, no clause
 *       punctuation) — rejects body prose whose first bold word is deep in a sentence.
 */
function evaluateCandidate(paragraphXml: string, citeStyleId: string): number[] | null {
  const facts = runFactsOf(paragraphXml, citeStyleId);
  // (a) already a proper cite — skip.
  if (facts.some((f) => f.cite)) return null;
  // (b) needs a bold run.
  const firstBold = facts.findIndex((f) => f.bold);
  if (firstBold < 0) return null;
  // Concatenated visible text for the year/prefix gates.
  const fullText = facts.map((f) => f.text).join("");
  // (c) a year near the front.
  const yearIndex = firstYearIndex(fullText);
  if (yearIndex < 0) return null;
  // (d) the first bold run leads.
  const firstBoldOffset = facts[firstBold].startOffset;
  if (firstBoldOffset > FIRST_BOLD_MAX_OFFSET) return null;
  // (e) the prefix before the first bold run is name-like.
  if (!prefixIsNameLike(fullText.slice(0, firstBoldOffset))) return null;
  const indices = planRunIndices(facts, yearIndex);
  return indices.length > 0 ? indices : null;
}

// ---------------------------------------------------------------------------
// Public: plan
// ---------------------------------------------------------------------------

/**
 * PURE. Scan a whole document's paragraphs and return one `CiteRepair` per mis-styled
 * cite found. Algorithm (CONSERVATIVE by design — a false positive leaks body text):
 *
 *   For each TAG/heading paragraph (`headingLevel != null`), scan FORWARD to find the
 *   single cite CANDIDATE — the FIRST subsequent paragraph that is NOT a list item
 *   (`<w:numPr>` — "bullet points under the tag"), NOT empty (has visible text), and NOT
 *   itself a heading. While scanning we SKIP list/empty paragraphs; if we hit ANOTHER
 *   heading we RESTART the window from it (a tag→subtag chain). The first qualifying
 *   non-list/non-empty/non-heading paragraph is the candidate; we evaluate it and then
 *   STOP scanning for this tag — cites "a few lines down" are deliberately NOT handled,
 *   so a "Smith 2020" inside body prose is never mistaken for the cite.
 *
 * A candidate that passes `evaluateCandidate`'s gates (a)–(e) yields a `CiteRepair`
 * targeting its bold author/year runs. Results are in document order; at most one repair
 * per distinct candidate paragraph (a candidate already repaired via one tag is the same
 * object, and because each tag stops at its first candidate, the same body paragraph is
 * never planned twice from different tags — its preceding tag is the only one that can
 * reach it as a first candidate).
 */
export function planCiteRepairs(
  paragraphs: readonly CiteRepairParagraph[],
  citeStyleId: string = CITE_STYLE_ID
): CiteRepair[] {
  const repairs: CiteRepair[] = [];
  // Track candidate paragraphs already planned so two adjacent tags pointing at the same
  // first body paragraph can't emit a duplicate repair (defensive; the stop-at-first-
  // candidate rule already makes this rare).
  const planned = new Set<number>();

  for (let t = 0; t < paragraphs.length; t++) {
    // Only a tag/heading opens a cite-search window.
    if (paragraphs[t].headingLevel === null) continue;

    // Scan forward for the FIRST real (non-list, non-empty, non-heading) paragraph.
    let candidate = -1;
    for (let j = t + 1; j < paragraphs.length; j++) {
      const p = paragraphs[j];
      // Another heading: this tag's window is superseded by a deeper tag — the outer
      // loop will open a window from THAT heading, so stop scanning for the current one.
      if (p.headingLevel !== null) break;
      // Skip bullets under the tag.
      if (isListParagraph(p.xml)) continue;
      // Skip empty paragraphs (no visible text after trimming).
      const text = paragraphTextOf(p.xml);
      if (text.trim().length === 0) continue;
      // First real paragraph → the candidate; stop (do not look deeper).
      candidate = j;
      break;
    }
    if (candidate < 0 || planned.has(candidate)) continue;

    const runIndices = evaluateCandidate(paragraphs[candidate].xml, citeStyleId);
    if (runIndices) {
      repairs.push({ paragraphIndex: candidate, runIndices });
      planned.add(candidate);
    }
  }
  return repairs;
}

/** The full concatenated visible text of a paragraph (all runs, document order). */
function paragraphTextOf(paragraphXml: string): string {
  return parseRuns(paragraphXml)
    .map((r) => runText(r.inner))
    .join("");
}

// ---------------------------------------------------------------------------
// Public: apply
// ---------------------------------------------------------------------------

/**
 * PURE. Inject `<w:rStyle w:val="{citeStyleId}"/>` into the rPr of the given run indices
 * (0-based in document order among the paragraph's `<w:r>` elements), returning the new
 * paragraph XML. This is the reversibility-critical mutation: it ONLY adds/replaces the
 * rStyle and preserves every other byte, so Show All / native reversal still restore the
 * exact original run formatting (the rStyle is a character-style reference, removable like
 * any other run property).
 *
 * Per OOXML CT_RPr ordering, `<w:rStyle>` is the FIRST child of `<w:rPr>`, so:
 *   * run with no `<w:rPr>`        → create one as the run's FIRST child:
 *                                    `<w:r><w:rPr><w:rStyle …/></w:rPr>…`
 *   * run with `<w:rPr>` but no rStyle → insert `<w:rStyle>` as the rPr's FIRST child.
 *   * run that already has `<w:rStyle>` → REPLACE its `w:val` (idempotent).
 * Self-closing `<w:r/>` runs carry no text and are never targeted, but are handled by
 * being skipped (the regex only rewrites paired runs).
 */
export function applyCiteStyleToParagraphXml(
  paragraphXml: string,
  runIndices: readonly number[],
  citeStyleId: string = CITE_STYLE_ID
): string {
  if (runIndices.length === 0) return paragraphXml;
  const targets = new Set(runIndices);
  const rStyleEl = `<w:rStyle w:val="${escapeAttr(citeStyleId)}"/>`;

  // Rewrite each run in place by walking the same RUN_RE matches in document order, so a
  // run's index here matches the index `planCiteRepairs` produced. We rebuild the string
  // from non-run segments + (possibly rewritten) run substrings; preserving the exact
  // inter-run bytes keeps everything outside the targeted rPr byte-identical.
  const re = new RegExp(RUN_RE.source, "g");
  let out = "";
  let last = 0;
  let runIndex = 0;
  let m: RegExpExecArray | null;
  while ((m = re.exec(paragraphXml)) !== null) {
    out += paragraphXml.slice(last, m.index); // bytes between runs, verbatim
    const full = m[0];
    const isSelfClosing = m[1] === undefined; // <w:r/> has no captured body
    if (targets.has(runIndex) && !isSelfClosing) {
      out += injectRStyleIntoRun(full, rStyleEl);
    } else {
      out += full; // untargeted (or self-closing) run: verbatim
    }
    last = m.index + full.length;
    runIndex++;
  }
  out += paragraphXml.slice(last); // trailing bytes after the last run, verbatim
  return out;
}

/**
 * Inject/replace the cite rStyle inside ONE paired `<w:r>…</w:r>` substring, returning
 * the rewritten run. Handles the three rPr shapes (none / present-without-rStyle /
 * present-with-rStyle) per the contract in `applyCiteStyleToParagraphXml`. Only the rPr
 * is touched; the run's content (`<w:t>`, etc.) is preserved verbatim.
 */
function injectRStyleIntoRun(runXml: string, rStyleEl: string): string {
  // Locate the run's own rPr (paired form first, then self-closing).
  const pairedRPr = /<w:rPr\b[^>]*>[\s\S]*?<\/w:rPr>/.exec(runXml);
  const selfClosingRPr = pairedRPr ? null : /<w:rPr\b([^>]*)\/>/.exec(runXml);

  if (pairedRPr) {
    const block = pairedRPr[0];
    const existing = /<w:rStyle\b[^>]*\/>|<w:rStyle\b[^>]*>[\s\S]*?<\/w:rStyle>/.exec(block);
    let newBlock: string;
    if (existing) {
      // REPLACE: swap the existing rStyle element for one with the right w:val (idempotent).
      newBlock = block.slice(0, existing.index) + rStyleEl + block.slice(existing.index + existing[0].length);
    } else {
      // INSERT rStyle as the FIRST child of <w:rPr> (CT_RPr ordering).
      newBlock = block.replace(/^(<w:rPr\b[^>]*>)/, `$1${rStyleEl}`);
    }
    return runXml.slice(0, pairedRPr.index) + newBlock + runXml.slice(pairedRPr.index + block.length);
  }

  if (selfClosingRPr) {
    // Expand a self-closing `<w:rPr …/>` into a paired block carrying the rStyle, keeping
    // any attributes the empty rPr had.
    const attrs = selfClosingRPr[1] ?? "";
    const expanded = `<w:rPr${attrs}>${rStyleEl}</w:rPr>`;
    return (
      runXml.slice(0, selfClosingRPr.index) +
      expanded +
      runXml.slice(selfClosingRPr.index + selfClosingRPr[0].length)
    );
  }

  // No rPr at all: create one as the FIRST child of the run, right after the `<w:r …>`
  // start tag (before the run content), per CT_R ordering.
  return runXml.replace(/^(<w:r(?:\s[^>]*)?>)/, `$1<w:rPr>${rStyleEl}</w:rPr>`);
}

/** Escape a value for use inside a double-quoted XML attribute (mirrors ooxmlPackage.ts). */
function escapeAttr(v: string): string {
  return v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}
