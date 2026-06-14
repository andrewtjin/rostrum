// Delete-analytics — the engine's SOLE content-deleter (plan 003-F1, the
// destructive verb). Isolated in its own module ON PURPOSE: the standing
// invariant is that NOTHING in the engine inserts, deletes, or reorders content
// (case 001-F1), and Delete-analytics is the single, audited exception. Keeping
// it alone means the 003-F1 audit ("what can emit deleteContentRange?") is one
// grep — every OTHER planner can stay provably delete-free.
//
// WHAT IT DELETES: exactly the analytics text the analytic-ify tool wrote — the
// off-palette navy runs keepers.isAnalytics recognizes (the SINGLE shared
// predicate, so keeper / delete / confirm-count can never disagree about what
// "analytics" is). Two granularities, decided per paragraph:
//
//   * WHOLLY-ANALYTICS paragraph (>= 1 analytics text run, NO non-analytics
//     text run, NO kind:"other" element, NOT inTable) -> delete the WHOLE
//     paragraph range [startIndex, endIndex) so the trailing newline goes too
//     and the now-empty line COLLAPSES (Docs merges it into the next). The one
//     exception is the segment-final paragraph: the API refuses to delete the
//     doc-final newline, so the clamp leaves [startIndex, endIndex - 1) and an
//     empty line remains (documented, unavoidable).
//   * EVERY OTHER paragraph (partial analytics, OR any table paragraph) ->
//     delete only the analytics RUN ranges, never the whole line. A partial
//     line keeps its non-analytics text; a table line keeps its structure
//     (tables ARE processed for the privacy purpose, but ONLY the partial path
//     — a whole-paragraph/newline collapse would damage the table grid).
//
// ORDER IS LOAD-BEARING (the reason this planner is a plan-review correctness
// BLOCKER zone): see planDeleteAnalytics. Unlike the four style/reconcile verbs
// — where descending emission is a non-load-bearing convention because no
// emitted request type mutates indexes — deleteContentRange SHIFTS every
// downstream index, and a Docs batchUpdate applies its requests SEQUENTIALLY.

import { isAnalytics } from "./keepers";
import { DeleteContentRangeRequest, GDoc, GParagraph, GRange, RequestGroup } from "./types";

// ---------------------------------------------------------------------------
// Public contract
// ---------------------------------------------------------------------------

/**
 * planDeleteAnalytics output. `groups` is the cross-module contract (the
 * controller applies them via the same chunk/revision machinery as every other
 * verb — chunkGroups preserves the descending order across chunk boundaries);
 * `result` feeds the receipt.
 */
export interface DeleteAnalyticsPlan {
  groups: RequestGroup[];
  result: {
    /** Distinct paragraphs that contributed >= 1 delete range (receipt unit —
     * paragraphs, Word parity; "runs" is banned receipt lexicon). */
    paragraphsAffected: number;
    /** Number of delete ranges emitted (post-clamp, post-coalesce). Surfaced
     * for diagnostics, never the receipt's headline unit. */
    runsDeleted: number;
  };
}

// ---------------------------------------------------------------------------
// Range geometry (private mechanics)
// ---------------------------------------------------------------------------

/**
 * One delete candidate before clamping/coalescing: the [start, end) range plus
 * the ordinal of the paragraph it came from, so paragraphsAffected stays a
 * DISTINCT-paragraph count even after two whole-analytics lines coalesce into a
 * single range (coalescing must not undercount the paragraphs the user sees
 * disappear).
 */
interface Candidate {
  startIndex: number;
  endIndex: number;
  paragraphIndex: number;
}

/**
 * Clamp a candidate's END off the segment-final newline — the SAME
 * isLastInSegment predicate the styles/restore/cite emitters all apply (planner
 * styleEnd, styles.clampedParagraphRange / headingTextRanges, restore
 * sentinelSpans, keepers.detectCiteLeads): the API rejects styling OR deleting
 * the doc-final newline (plan D6), so any range reaching the final paragraph's
 * endIndex stops one short. styleEnd itself is private to planner.ts and shaped
 * around its Atom type, so this is a deliberate LOCAL mirror of the one-line
 * predicate rather than an import (the spec sanctions a local mirror when
 * extraction across the parallel-edited planner is risky). Returns null when
 * the clamp empties the range (a lone segment-final newline) — a zero-length
 * deleteContentRange would 400, and there is nothing left to remove anyway.
 */
function clampFinalNewline(p: GParagraph, range: GRange): GRange | null {
  const endIndex = p.isLastInSegment && range.endIndex === p.endIndex ? range.endIndex - 1 : range.endIndex;
  return endIndex > range.startIndex ? { startIndex: range.startIndex, endIndex } : null;
}

/**
 * Is this paragraph WHOLLY analytics? (plan §3 correctness classification.)
 * Requires ALL of: at least one analytics text run, no non-analytics text run,
 * no kind:"other" element (a chip/break/object can never be deleted as
 * analytics — A9 whitelist), and NOT a table paragraph (tables never collapse a
 * whole line). Only such a paragraph may have its WHOLE range (incl. newline)
 * deleted; everything else takes the run-only partial path. Computed in one
 * pass so the three text-run facts (has-analytics, has-non-analytics) and the
 * has-other fact are read off the same element walk.
 */
function isWhollyAnalytics(p: GParagraph): boolean {
  if (p.inTable) return false; // tables: partial path only (structure untouched)
  let hasAnalytics = false;
  for (const el of p.elements) {
    if (el.kind === "other") return false; // a non-text element breaks "wholly"
    // kind:"text" from here on.
    if (isAnalytics(el)) hasAnalytics = true;
    else return false; // a non-analytics text run breaks "wholly"
  }
  return hasAnalytics;
}

/**
 * Collect every delete candidate for one paragraph, classification decided
 * here so the whole/partial choice and the range it produces live in one place:
 *   * wholly-analytics -> the WHOLE paragraph range [startIndex, endIndex)
 *     (the clamp later trims a segment-final newline; a non-final line keeps
 *     its newline so the line collapses);
 *   * else (partial, OR a table paragraph) -> one candidate per ANALYTICS run,
 *     each the run element's own [startIndex, endIndex), but with the
 *     PARAGRAPH-MARK NEWLINE clamped off (see below).
 *
 * PARTIAL-PATH NEWLINE CLAMP (exec-review BLOCKER fix): a paragraph's trailing
 * "\n" is bundled INTO its last text run (parse attaches it to the final
 * element). On the partial path the paragraph SURVIVES, so a run that reaches
 * p.endIndex must NOT carry the paragraph mark into its delete range — deleting
 * it would (a) collapse a partial body line, merging the surviving
 * non-analytics text into the next paragraph, and (b) destroy a table cell's
 * structural newline (damaging the grid / forcing a torn mid-batch delete the
 * API rejects). So a partial-path run whose end reaches p.endIndex is trimmed to
 * p.endIndex - 1 REGARDLESS of isLastInSegment — the only place a trailing
 * newline is ever removed is the deliberate wholly-analytics line collapse
 * above. A run that is ONLY the trailing newline (nothing left after the trim)
 * is dropped (a zero-length deleteContentRange would 400, and there is no text
 * to remove). A paragraph with no analytics at all yields nothing.
 */
function paragraphCandidates(p: GParagraph): Candidate[] {
  if (isWhollyAnalytics(p)) {
    return [{ startIndex: p.startIndex, endIndex: p.endIndex, paragraphIndex: p.index }];
  }
  const out: Candidate[] = [];
  for (const el of p.elements) {
    if (!isAnalytics(el)) continue;
    // Trim the paragraph mark off a partial-path run that reaches the line end.
    const endIndex = el.endIndex === p.endIndex ? p.endIndex - 1 : el.endIndex;
    if (endIndex > el.startIndex) {
      out.push({ startIndex: el.startIndex, endIndex, paragraphIndex: p.index });
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Public: planDeleteAnalytics
// ---------------------------------------------------------------------------

/** A deleteContentRange request over one range — the SINGLE emission shape this
 * planner produces (003-F1 audit: Delete-analytics emits ONLY this type). The
 * range is a fresh object so no two emitted requests alias one another. */
function deleteRequest(range: GRange): DeleteContentRangeRequest {
  return { deleteContentRange: { range: { startIndex: range.startIndex, endIndex: range.endIndex } } };
}

/**
 * Plan the full Delete-analytics pass over a gated view (the controller runs
 * assertSingleTab + assertNoSuggestions + assertNotHidden first; this module
 * assumes a single-tab, suggestion-free, unarmed view and never re-checks).
 *
 * THE ORDER IS LOAD-BEARING — the three phases run in exactly this sequence:
 *
 *   1. COLLECT + CLAMP each candidate independently (isLastInSegment-only): the
 *      segment-final newline is unremovable, so a range that reaches it is
 *      trimmed BEFORE anything else looks at adjacency. Clamping first is what
 *      makes phase 2 correct — a whole-analytics FINAL paragraph (clamped to
 *      stop one short of its newline) must NOT merge with an interior line that
 *      still owns the newline between them, because that newline is no longer in
 *      any range. Coalescing before clamping would fuse across the gap and
 *      delete the wrong span.
 *   2. COALESCE only ranges STILL TOUCHING after the clamp (range[i].end ===
 *      range[i+1].start in ascending order): two adjacent WHOLE-analytics
 *      paragraphs share a boundary (one's end is the next's start) and fold into
 *      a single delete; a clamped final line leaves a one-index gap and stays
 *      separate. Fewer, larger ranges are equivalent to many touching ones but
 *      shift fewer indexes and read cleaner in a wire dump.
 *   3. SORT DESCENDING by startIndex and emit one RequestGroup per range. This
 *      descending order is LOAD-BEARING, in deliberate contrast to the four
 *      convention-order verbs: a Docs batchUpdate applies its requests
 *      SEQUENTIALLY, and deleteContentRange shifts every downstream index, so a
 *      lower-index delete done first would invalidate every higher range after
 *      it. Highest-index-first means each delete only moves indexes BELOW ranges
 *      already applied. chunkGroups preserves this order across revision-chained
 *      chunks (it never reorders), so the guarantee holds end to end.
 *
 * One RequestGroup PER range (the chunker never splits a group): each delete
 * stands alone, so a torn multi-chunk apply removes a clean PREFIX of the
 * descending sequence (the highest ranges) and leaves a coherent doc the user
 * can re-run Delete-analytics over to finish (the partial-apply error copy says
 * so). paragraphsAffected counts DISTINCT contributing paragraphs (the receipt
 * unit); runsDeleted counts the emitted ranges.
 */
export function planDeleteAnalytics(doc: GDoc): DeleteAnalyticsPlan {
  // --- Phase 1: collect every candidate, then clamp each independently. ---
  // Ascending document order falls out of iterating paragraphs/elements in
  // order, which phase 2's adjacency merge relies on (it walks left to right).
  const clamped: Candidate[] = [];
  for (const p of doc.paragraphs) {
    for (const c of paragraphCandidates(p)) {
      const range = clampFinalNewline(p, { startIndex: c.startIndex, endIndex: c.endIndex });
      // A clamp that empties the range (a lone segment-final newline) drops the
      // candidate entirely — nothing to delete, and a zero-length range 400s.
      if (range !== null) {
        clamped.push({ startIndex: range.startIndex, endIndex: range.endIndex, paragraphIndex: c.paragraphIndex });
      }
    }
  }

  // --- Phase 2: coalesce only ranges STILL TOUCHING after the clamp. ---
  // Candidates are already in ascending start order (document order), so a
  // single forward sweep merges runs whose end meets the next start. The
  // contributing-paragraph set rides along so a merged range still counts every
  // paragraph it swallowed (two coalesced whole-analytics lines = 2 affected).
  const merged: { startIndex: number; endIndex: number; paragraphs: Set<number> }[] = [];
  for (const c of clamped) {
    const last = merged[merged.length - 1];
    if (last !== undefined && last.endIndex === c.startIndex) {
      last.endIndex = c.endIndex;
      last.paragraphs.add(c.paragraphIndex);
    } else {
      merged.push({ startIndex: c.startIndex, endIndex: c.endIndex, paragraphs: new Set([c.paragraphIndex]) });
    }
  }

  // --- Phase 3: sort DESCENDING by startIndex and emit one group per range. ---
  // The merged ranges are disjoint and non-touching, so sorting by startIndex
  // descending fully orders them highest-first — the load-bearing emission order.
  const ordered = [...merged].sort((a, b) => b.startIndex - a.startIndex);
  const groups: RequestGroup[] = ordered.map((m) => ({ requests: [deleteRequest(m)] }));

  // paragraphsAffected: distinct paragraphs across ALL ranges (a paragraph
  // counts once even if it contributed several partial runs).
  const affected = new Set<number>();
  for (const m of merged) for (const idx of m.paragraphs) affected.add(idx);

  return {
    groups,
    result: { paragraphsAffected: affected.size, runsDeleted: groups.length }
  };
}
