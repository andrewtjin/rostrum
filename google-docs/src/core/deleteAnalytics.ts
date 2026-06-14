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
//     text run, NO kind:"other" element, NOT inTable) -> delete the paragraph
//     range. Its trailing newline (the paragraph mark) is removed ONLY when the
//     line can safely COLLAPSE into an immediately-adjacent following body
//     paragraph; otherwise the newline is PRESERVED and only the text is removed
//     (an empty line remains). Deleting a paragraph mark is legal solely as a
//     paragraph MERGE: the Docs API returns 400 "cannot delete the requested
//     range" for deleting "the last newline of a Body/TableCell" and "the
//     newline before a Table/TableOfContents/SectionBreak" (wet-confirmed
//     2026-06-13, requests[8] — a wholly-analytics line sitting just before a
//     table). canCollapseInto() detects ALL of those structural cases with ONE
//     uniform test (the next paragraph is absent or at an index GAP, or is a
//     table cell), so the planner never has to decode WHICH element follows. The
//     analytics text is fully removed either way (the privacy goal); only the
//     blank-line cleanup differs.
//   * EVERY OTHER paragraph (partial analytics, OR any table paragraph) ->
//     delete only the analytics RUN ranges, never the whole line, each trimmed
//     off its paragraph mark. A partial line keeps its non-analytics text; a
//     table line keeps its structure (tables ARE processed for the privacy
//     purpose, but ONLY this partial path — deleting "content within a table
//     cell" is the one table edit the API permits; a whole-line collapse would
//     hit the "newline before / last newline of a cell" 400s).
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
 * May a WHOLLY-analytics paragraph delete its trailing newline (collapsing the
 * now-empty line into the next), or must it preserve it (leaving an empty line)?
 *
 * Deleting a paragraph mark is legal ONLY as a paragraph MERGE into an
 * immediately-adjacent following paragraph. The Docs API returns a 400 ("cannot
 * delete the requested range") for deleting "the last newline of a
 * Body/TableCell" and "the newline before a Table/TableOfContents/SectionBreak"
 * (the documented invalid-delete list; wet-confirmed 2026-06-13, requests[8]).
 * Every one of those cases — and the doc end — shares a single OBSERVABLE
 * signature in the parsed view: the next paragraph is either ABSENT (segment
 * end) or begins at an INDEX GAP after this paragraph's endIndex (a table start,
 * section break, table-of-contents or equation consumes index space that the
 * flattened paragraph list skips over), or is itself a table cell. So one
 * uniform test covers every structural case WITHOUT the planner decoding what
 * follows: collapse iff a next paragraph exists, begins EXACTLY at this
 * paragraph's endIndex (no gap), and is a body paragraph (not in a table). The
 * `!next.inTable` guard is what makes the index-flat test harness exercise this
 * too — in real Docs a table always sits behind a structural gap, but the flat
 * model has none, so the explicit table check keeps both honest. When the test
 * fails the caller preserves the newline; the analytics TEXT is still fully
 * removed (the privacy goal), only an empty line is left where a collapse would
 * have closed up.
 */
function canCollapseInto(p: GParagraph, next: GParagraph | undefined): boolean {
  return next !== undefined && next.startIndex === p.endIndex && !next.inTable;
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
 * p.endIndex - 1 — the partial path NEVER removes a paragraph mark. A run that
 * is ONLY the trailing newline (nothing left after the trim) is dropped (a
 * zero-length deleteContentRange would 400, and there is no text to remove). A
 * paragraph with no analytics at all yields nothing.
 *
 * WHOLLY path (003-F8 fix): the paragraph mark is removed (collapsing the line)
 * ONLY when canCollapseInto(p, next) — a safe merge into an adjacent body
 * paragraph. Otherwise the mark is preserved exactly like the partial path,
 * because deleting a newline before a table / section break, or the
 * segment-final newline, is a 400 (the wet finding). `next` is the following
 * paragraph in document order (undefined past the last).
 */
function paragraphCandidates(p: GParagraph, next: GParagraph | undefined): Candidate[] {
  if (isWhollyAnalytics(p)) {
    // Collapse (include the paragraph mark) only when structurally safe; else
    // preserve the newline and remove just the run text (an empty line remains).
    const endIndex = canCollapseInto(p, next) ? p.endIndex : p.endIndex - 1;
    // A lone-newline analytics paragraph trims to empty — drop it (a zero-length
    // deleteContentRange would 400, and there is no text left to remove).
    return endIndex > p.startIndex ? [{ startIndex: p.startIndex, endIndex, paragraphIndex: p.index }] : [];
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
 *   1. COLLECT each candidate, deciding its newline's fate AS IT IS BUILT
 *      (paragraphCandidates + canCollapseInto): a wholly-analytics line that
 *      cannot safely collapse (segment end, or a table / section break next)
 *      stops one short of its paragraph mark, exactly like a partial run. Fixing
 *      each range's end here is what makes phase 2 correct — a preserved-newline
 *      FINAL line must NOT merge with an interior line that still owns the
 *      newline between them, because that newline is no longer in any range.
 *      Deciding adjacency before coalescing would fuse across the gap and delete
 *      the wrong span.
 *   2. COALESCE only ranges STILL TOUCHING (range[i].end === range[i+1].start in
 *      ascending order): two adjacent WHOLE-analytics paragraphs that both
 *      collapse share a boundary (one's end is the next's start) and fold into a
 *      single delete; a preserved-newline final line leaves a one-index gap and
 *      stays separate. Fewer, larger ranges are equivalent to many touching ones
 *      but shift fewer indexes and read cleaner in a wire dump.
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
  // --- Phase 1: collect every candidate, each already shaped (newline kept or
  // dropped) by paragraphCandidates/canCollapseInto. ---
  // Ascending document order falls out of iterating paragraphs/elements in
  // order, which phase 2's adjacency merge relies on (it walks left to right).
  // The NEXT paragraph (doc.paragraphs[i + 1], undefined past the last) is
  // handed in so the wholly-analytics collapse decision can see its neighbor.
  const candidates: Candidate[] = [];
  for (let i = 0; i < doc.paragraphs.length; i++) {
    for (const c of paragraphCandidates(doc.paragraphs[i], doc.paragraphs[i + 1])) candidates.push(c);
  }

  // --- Phase 2: coalesce only ranges STILL TOUCHING. ---
  // Candidates are already in ascending start order (document order), so a
  // single forward sweep merges runs whose end meets the next start. The
  // contributing-paragraph set rides along so a merged range still counts every
  // paragraph it swallowed (two coalesced whole-analytics lines = 2 affected).
  const merged: { startIndex: number; endIndex: number; paragraphs: Set<number> }[] = [];
  for (const c of candidates) {
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
