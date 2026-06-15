// Show All + convergence sweep (plan D3/A2/A6/A7/A14; edge rows 9-11, 16).
//
// Show All is the engine's recovery verb, so it is deliberately the LEAST
// gated, most forgiving pass in the system: it never calls the suggestion
// gate (plan A7 — a teammate's pending suggestion must never lock a doc in
// the hidden state), it restores whatever the in-document manifests still
// describe exactly, normalizes whatever they no longer describe, and then
// SWEEPS any sentinel-size text left with no record at all, so a doc
// converges to all-visible no matter how badly the rstm state decayed
// (cut/paste destroys named ranges — WAI per Google — and "Make a copy"
// drops them entirely). The one boundary it will not cross silently: when
// the doc carries ZERO rstm ranges, tiny text might simply be the user's own
// formatting, so the controller asks first (pureSweepConsentNeeded, plan
// A14) and re-invokes planShowAll with the answer.
//
// Everything here is pure planning: GDoc view in, RequestGroups + counted
// result out. The only request types ever emitted are updateTextStyle
// (fontSize + foregroundColor — reveal un-shrinks AND clears the hide white,
// HIDE_FIELDS), updateParagraphStyle (spacing fields only) and deleteNamedRange
// — content is never inserted, deleted or reordered (case 001-F1). That hard
// invariant is also why emission ORDER is a mere convention here, never a
// correctness requirement: no request can shift the indexes a later request
// targets.

import { HIDE_FIELDS, SENTINELS } from "./constants";
import { decodeRangeName, isRstmName } from "./rangeNames";
import {
  DecodedRangeName,
  DocsParagraphStyle,
  DocsRequest,
  GDoc,
  GNamedRange,
  GParagraph,
  GRange,
  GShowAllResult,
  RequestGroup,
  RleEntry,
  UpdateParagraphStyleRequest,
  UpdateTextStyleRequest
} from "./types";

/** The one field mask spacing restores ever write. Both channels always ride
 * together because the spacing record always carries both values: a null
 * entry is restored by CLEARING (named in fields, absent from the style) —
 * the documented clear-to-inherit semantics (plan D13). */
const SPACING_FIELDS = "spaceAbove,spaceBelow";

/**
 * The ONE spacing-restore request shape (plan A12): recorded channels
 * materialize as explicit PT values; null (inherited) channels clear via the
 * field mask. Exported because planner.ts's rework path emits the SAME
 * restore when a hidden paragraph surfaces on re-hide (its spacing record
 * dies with its sizes anchor) — sharing the builder keeps the two emitters
 * from ever drifting on shape or field mask.
 */
export function spacingRestoreRequest(
  range: GRange,
  spaceAbovePt: number | null,
  spaceBelowPt: number | null
): UpdateParagraphStyleRequest {
  const paragraphStyle: DocsParagraphStyle = {};
  if (spaceAbovePt !== null) {
    paragraphStyle.spaceAbove = { magnitude: spaceAbovePt, unit: "PT" };
  }
  if (spaceBelowPt !== null) {
    paragraphStyle.spaceBelow = { magnitude: spaceBelowPt, unit: "PT" };
  }
  return { updateParagraphStyle: { range, paragraphStyle, fields: SPACING_FIELDS } };
}

/**
 * planShowAll's product. `groups` + `result` are the cross-module contract
 * (controller applies the groups via the chunker and renders the result
 * through strings.showAllReceipt). Edge row 16 — ranges written by a NEWER
 * engine version, left entirely untouched because deleting or "restoring"
 * grammar this build cannot decode could corrupt that version's manifest —
 * rides INSIDE the result (rangesSkippedNewerVersion) so the receipt can
 * warn the user without a side channel.
 */
export interface ShowAllPlan {
  groups: RequestGroup[];
  result: GShowAllResult;
}

/**
 * The pure-sweep consent boundary (plan A14). Returns the number of
 * sentinel-size text PASSAGES (maximal merged sub-spans — the unit
 * strings.consentPrompt renders) when the doc carries ZERO rstm-family
 * ranges but still shows sentinel-size text: a copied doc whose ranges were
 * dropped, or the user's own deliberately tiny text. Returns null when no
 * consent is needed — either rstm state exists (orphans are ours; Show All
 * sweeps without asking) or there is nothing tiny to ask about.
 *
 * The range test is the version-AGNOSTIC family check on purpose: a doc
 * armed by a newer engine version is still an armed doc, not a consent case
 * — decode strictness must never turn it into a prompt.
 */
export function pureSweepConsentNeeded(doc: GDoc): number | null {
  if (doc.namedRanges.some((nr) => isRstmName(nr.name))) return null;
  const passages = sentinelSpans(doc).length;
  return passages === 0 ? null : passages;
}

/**
 * Plan the complete Show All pass: per-range restore, anchor deletion, and
 * the convergence sweep. Pure — emits requests, applies nothing.
 *
 * `sweepUnrecorded` is the consent answer for the PURE-sweep path (no rstm
 * ranges at all): false leaves unrecorded tiny text exactly as the user made
 * it; true adopts it into the sweep. When ANY rstm-family range exists the
 * flag is irrelevant — the sweep always runs, because orphan sentinel text
 * next to our own state is decayed hide state by overwhelming likelihood
 * (plan A14: "orphans are ours").
 *
 * Show All NEVER gates on suggestions (plan A7) and never asks about tabs —
 * those gates belong to the verbs that CREATE state; the controller owns
 * wiring them. Idempotence falls out of the algorithm: a second pass sees no
 * rstm ranges and no sentinel text, so it emits zero groups.
 */
export function planShowAll(doc: GDoc, sweepUnrecorded: boolean): ShowAllPlan {
  // Single geometry source for both restore and sweep (see sentinelSpans) so
  // the two passes can never disagree about what counts as hidden text.
  const inventory = sentinelSpans(doc);

  const result: GShowAllResult = {
    segmentsRestoredExact: 0,
    segmentsNormalized: 0,
    sweptOrphans: 0,
    rangesDeleted: 0,
    rangesSkippedNewerVersion: 0
  };

  // Classify every rstm-family range exactly once. Foreign (non-rstm) ranges
  // are invisible to this verb: other add-ons' state is not ours to delete.
  const known: { nr: GNamedRange; decoded: DecodedRangeName }[] = [];
  // Index ranges the sweep must NOT clear. Two sources, one reason each:
  // known "sizes" segments are already owned by restore (exact or normalize),
  // and unknown-version segments are hands-off entirely (edge row 16).
  // Known SPACING segments are deliberately NOT holes: a spacing record says
  // nothing about TEXT sizes, and when its parallel sizes range was destroyed
  // the sweep is the only thing left that can un-shrink that text.
  const sweepHoles: GRange[] = [];

  for (const nr of doc.namedRanges) {
    if (!isRstmName(nr.name)) continue;
    const decoded = decodeRangeName(nr.name);
    if (decoded === null) {
      // rstm-family but undecodable: a newer version's state, or corruption
      // indistinguishable from it. Count it (the receipt's row-16 amber line),
      // shield its text, touch nothing.
      result.rangesSkippedNewerVersion++;
      pushNonEmpty(sweepHoles, nr.segments);
      continue;
    }
    known.push({ nr, decoded });
    if (decoded.kind === "sizes") pushNonEmpty(sweepHoles, nr.segments);
  }

  const groups: RequestGroup[] = [];

  // Descending anchor order across ranges is a CONVENTION, not load-bearing
  // (plan A16): no emitted request type mutates indexes, so any order applies
  // identically. We keep it because the planner emits descending and one
  // shared convention keeps wire dumps comparable across verbs.
  const ordered = [...known].sort((a, b) => anchorStart(b.nr) - anchorStart(a.nr));

  for (const { nr, decoded } of ordered) {
    const requests: DocsRequest[] = [];
    // A zero-length segment describes NOTHING — no text to restore, no
    // passage the user can see. Dropping it without counting keeps the
    // amber receipt truthful (counting would report a reset that never
    // happened); the anchor delete below still retires the record.
    const liveSegments = nr.segments.filter((seg) => seg.endIndex > seg.startIndex);
    if (decoded.kind === "sizes") {
      // Sizes ranges restore RANGE-at-once, not segment-at-once: a pure
      // split's segments only restore exactly when the RLE is walked
      // CONTINUOUSLY across them (see restoreSizesRange).
      restoreSizesRange(liveSegments, decoded.entries, inventory, requests, result);
    } else {
      for (const seg of [...liveSegments].sort((a, b) => b.startIndex - a.startIndex)) {
        restoreSpacingSegment(seg, decoded, doc.paragraphs, requests, result);
      }
    }
    // The anchor delete rides LAST in the SAME atomic group as its restores
    // (guards.chunkGroups never splits a group): if the restores land, the
    // record must die with them, or a re-run would treat already-restored
    // text as still recorded and write stale sizes over fresh edits.
    requests.push({ deleteNamedRange: { namedRangeId: nr.id } });
    result.rangesDeleted++;
    groups.push({ requests });
  }

  // CONVERGENCE SWEEP (plan D3): clear-to-inherit any sentinel-size sub-span
  // no surviving record accounts for. Gated only at the pure-sweep boundary.
  const hadRstmState = known.length > 0 || result.rangesSkippedNewerVersion > 0;
  if (hadRstmState || sweepUnrecorded) {
    const orphans = subtractSegments(inventory, sweepHoles);
    // Descending again — same non-load-bearing convention as above.
    for (const orphan of orphans.sort((a, b) => b.startIndex - a.startIndex)) {
      // One group per orphan: a sweep clear stands alone (no anchor to pair
      // with), so the chunker may break anywhere between them.
      groups.push({ requests: [clearSizeRequest(orphan)] });
      result.sweptOrphans++;
    }
  }

  return { groups, result };
}

// ---------------------------------------------------------------------------
// Per-segment restore (private mechanics)
// ---------------------------------------------------------------------------

/**
 * Restore all live segments of one "sizes" range. The exactness rule (plan
 * A2, sharpened for splits): when the segments' summed length, in DOCUMENT
 * order, equals the RLE's total char count, the range was merely SPLIT —
 * Docs removed indexes between the pieces but every recorded character is
 * still at its recorded offset within the concatenation, so the RLE is
 * walked CONTINUOUSLY across the segments and every entry's size is written
 * back over its exact sub-spans (no amber: a pure split restores byte-exact,
 * and counting it normalized would misreport work that was in fact perfect).
 * An intact single segment is just the one-piece case of the same walk.
 *
 * When the sum does NOT match, the correspondence is broken somewhere, but
 * not necessarily everywhere: each segment is then judged on its OWN length
 * (a surviving segment that alone matches the total is the original intact
 * piece and restores exactly; every other remnant normalizes — clear to the
 * style's inherited size). Guessing which chars a drifted record still
 * describes could restore wrong sizes over wrong spans, which is strictly
 * worse than the visible, amber-counted normalize.
 *
 * ALL paths write only over sub-spans CURRENTLY at a sentinel size (the
 * intersection with the inventory): text the user pasted or resized inside a
 * hidden region is theirs and keeps its size (plan A6). This also makes the
 * segment-final-newline clamp structural — the inventory never contains that
 * index, so a decayed segment that grew over it still cannot target it.
 */
function restoreSizesRange(
  segments: readonly GRange[],
  entries: readonly RleEntry[],
  inventory: readonly GRange[],
  requests: DocsRequest[],
  result: GShowAllResult
): void {
  // Document order is the walk order — the RLE records sizes left to right.
  const segs = [...segments].sort((a, b) => a.startIndex - b.startIndex);
  const totalLength = segs.reduce((sum, s) => sum + (s.endIndex - s.startIndex), 0);
  const rleLength = entries.reduce((sum, e) => sum + e.count, 0);

  if (segs.length > 0 && totalLength === rleLength) {
    // Pure split (or intact): walk entries continuously across segments. An
    // entry may straddle a split point, so each entry consumes characters
    // from as many segments as it needs; every consumed piece restores at
    // the entry's recorded size, scoped by the sentinel inventory (A6).
    let segIdx = 0;
    let pos = segs[0].startIndex;
    for (const entry of entries) {
      let remaining = entry.count;
      while (remaining > 0) {
        const seg = segs[segIdx];
        const take = Math.min(remaining, seg.endIndex - pos);
        for (const piece of intersectWithSpans(pos, pos + take, inventory)) {
          requests.push(entry.sizePt === null ? clearSizeRequest(piece) : setSizeRequest(piece, entry.sizePt));
        }
        pos += take;
        remaining -= take;
        // Hop to the next segment exactly at its boundary; the length match
        // guarantees the final entry ends on the final segment's end.
        if (pos === seg.endIndex && segIdx + 1 < segs.length) {
          segIdx++;
          pos = segs[segIdx].startIndex;
        }
      }
    }
    // Each segment is a passage the user sees come back — count them all.
    result.segmentsRestoredExact += segs.length;
    return;
  }

  // Mismatched range: judge each segment on its own length (see contract).
  for (const seg of segs) {
    if (seg.endIndex - seg.startIndex === rleLength) {
      // The original intact piece — restore it via the one-segment walk
      // above (its sum trivially matches, so this recursion is depth 1).
      restoreSizesRange([seg], entries, inventory, requests, result);
    } else {
      for (const piece of intersectWithSpans(seg.startIndex, seg.endIndex, inventory)) {
        requests.push(clearSizeRequest(piece));
      }
      result.segmentsNormalized++;
    }
  }
}

/**
 * Restore one segment of a spacing range (plan A12). "Aligned" means the
 * segment still tiles one or more COMPLETE, non-table paragraphs — only then
 * do we know the recorded values land on the paragraphs they were taken
 * from, and one updateParagraphStyle over the segment restores them (nulls
 * clear to inherit via the field mask).
 *
 * A misaligned segment emits NOTHING and counts normalized. Unlike sizes,
 * spacing has no sentinel marker to scope a safe write by — and
 * updateParagraphStyle hits every paragraph its range merely touches — so a
 * drifted record could stamp recorded spacing onto neighbors the user
 * formatted themselves. Leaving the collapsed (zeroed) spacing in place is
 * visible, benign, and surfaced by the amber count; overwriting a neighbor
 * is silent damage.
 */
function restoreSpacingSegment(
  seg: GRange,
  record: Extract<DecodedRangeName, { kind: "spacing" }>,
  paragraphs: readonly GParagraph[],
  requests: DocsRequest[],
  result: GShowAllResult
): void {
  const writeEnd = alignedSpacingEnd(seg, paragraphs);
  if (writeEnd === null) {
    result.segmentsNormalized++;
    return;
  }
  // Clamp the write off the segment-final newline (writeEnd) — uniform with the
  // hide-side collapse write (planner.ts) and every other paragraph-style
  // emitter, so restore never bets that updateParagraphStyle tolerates the
  // final-newline index that updateTextStyle rejects. updateParagraphStyle still
  // hits the whole paragraph because it targets every paragraph its range
  // overlaps, so the clamp changes nothing the user sees.
  requests.push(
    spacingRestoreRequest({ ...seg, endIndex: writeEnd }, record.spaceAbovePt, record.spaceBelowPt)
  );
  // An aligned spacing restore IS a segment restored exactly from its record
  // — it shares the exact counter with sizes segments so the receipt's
  // "Restored N passages" stays truthful about work done (strings.ts already
  // folds normalized + swept into one amber bucket the same way).
  result.segmentsRestoredExact++;
}

/**
 * When the segment exactly tiles consecutive whole paragraphs, return the END
 * INDEX to write the spacing restore over — clamped off the unstylable final
 * newline on the segment-final paragraph, so every paragraph-style write the
 * engine emits obeys the same no-style-past-the-final-newline rule. Returns
 * null when the segment is misaligned (caller counts it normalized).
 *
 * Two tolerances, both deliberate: table paragraphs always fail (the planner
 * never collapses table spacing — edge row 1 — so an aligned-looking table hit
 * is decay, not ours); and the segment may close at either the planner's
 * unclamped anchor ([start, p.endIndex]) or the clamped form (p.endIndex - 1) —
 * both describe the same whole paragraph, and either way we emit the clamped
 * end.
 */
function alignedSpacingEnd(seg: GRange, paragraphs: readonly GParagraph[]): number | null {
  let i = paragraphs.findIndex((p) => p.startIndex === seg.startIndex);
  if (i === -1) return null;
  let cursor = seg.startIndex;
  for (; i < paragraphs.length; i++) {
    const p = paragraphs[i];
    // Body paragraphs are contiguous; a gap means the segment start matched
    // inside structure we do not model the same way (table overhead indexes).
    if (p.startIndex !== cursor) return null;
    if (p.inTable) return null;
    // Segment closes at this paragraph (unclamped or clamped anchor form):
    // emit the clamped write end so a last-in-segment paragraph stops one short
    // of its final newline, matching the hide-side collapse write.
    if (seg.endIndex === p.endIndex || (p.isLastInSegment && seg.endIndex === p.endIndex - 1)) {
      return p.isLastInSegment ? p.endIndex - 1 : p.endIndex;
    }
    if (seg.endIndex < p.endIndex) return null; // ends mid-paragraph
    cursor = p.endIndex;
  }
  return null; // ran past the last paragraph without closing the segment
}

// ---------------------------------------------------------------------------
// Sentinel-span geometry (private mechanics)
// ---------------------------------------------------------------------------

/**
 * Inventory every style-targetable sub-span CURRENTLY at a sentinel size, as
 * maximal merged [start, end) ranges in ascending document order. This is
 * the single geometry source for restore, sweep, and the consent count, so
 * no two passes can ever disagree about what looks hidden.
 *
 * Three exclusions are load-bearing:
 *   * kind "other" never enters — the A9 whitelist means only textRuns are
 *     ever style-targeted, so only textRuns can need un-hiding;
 *   * table paragraphs never enter — the planner never hides table content
 *     (edge row 1), so sentinel text in a table is the user's own and
 *     neither restore nor sweep may touch it;
 *   * the segment-final newline is clamped off (the API rejects styling it —
 *     plan D6), which also keeps a sentinel-sized empty final paragraph
 *     correctly invisible to the sweep AND the consent count.
 *
 * Membership is EXACT equality against the append-only SENTINELS set: 1.5pt
 * text is not ours while the sentinel is 1 — we only ever reclaim sizes we
 * ourselves wrote (plan A16), everything else is user formatting.
 */
function sentinelSpans(doc: GDoc): GRange[] {
  const spans: GRange[] = [];
  for (const p of doc.paragraphs) {
    if (p.inTable) continue;
    for (const el of p.elements) {
      if (el.kind !== "text") continue;
      if (el.fontSizePt === null || !SENTINELS.includes(el.fontSizePt)) continue;
      let end = el.endIndex;
      // Clamp the unstylable segment-final newline (it lives inside the final
      // paragraph's last text element — builder/parse invariant).
      if (p.isLastInSegment && end === p.endIndex) end -= 1;
      if (end <= el.startIndex) continue;
      const prev = spans[spans.length - 1];
      // Body indexes are contiguous across paragraphs, so adjacency merges
      // hidden regions that span paragraph boundaries into one passage; a
      // chip ("other") or foreign-size element in between leaves a gap and
      // correctly splits the passage.
      if (prev !== undefined && prev.endIndex === el.startIndex) {
        prev.endIndex = end;
      } else {
        spans.push({ startIndex: el.startIndex, endIndex: end });
      }
    }
  }
  return spans;
}

/** Pieces of [startIndex, endIndex) that are inside the inventory — the A6
 * scope filter. Inventory is ascending and disjoint, so output is too. */
function intersectWithSpans(startIndex: number, endIndex: number, spans: readonly GRange[]): GRange[] {
  const out: GRange[] = [];
  for (const span of spans) {
    const s = Math.max(startIndex, span.startIndex);
    const e = Math.min(endIndex, span.endIndex);
    if (s < e) out.push({ startIndex: s, endIndex: e });
  }
  return out;
}

/**
 * Pieces of `spans` NOT covered by any hole — the orphan finder. Holes may
 * arrive unsorted and overlapping (segments of many ranges); a cursor walk
 * over the sorted holes handles both without merging them first.
 */
function subtractSegments(spans: readonly GRange[], holes: readonly GRange[]): GRange[] {
  const sortedHoles = [...holes].sort((a, b) => a.startIndex - b.startIndex);
  const out: GRange[] = [];
  for (const span of spans) {
    let cursor = span.startIndex;
    for (const hole of sortedHoles) {
      if (hole.endIndex <= cursor) continue;
      if (hole.startIndex >= span.endIndex) break;
      if (hole.startIndex > cursor) {
        out.push({ startIndex: cursor, endIndex: hole.startIndex });
      }
      cursor = Math.max(cursor, hole.endIndex);
      if (cursor >= span.endIndex) break;
    }
    if (cursor < span.endIndex) out.push({ startIndex: cursor, endIndex: span.endIndex });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Small shared builders (private mechanics)
// ---------------------------------------------------------------------------

/** Append only segments that cover at least one character (zero-length decay
 * must not split sweep spans into cosmetic extra pieces). */
function pushNonEmpty(target: GRange[], segments: readonly GRange[]): void {
  for (const seg of segments) {
    if (seg.endIndex > seg.startIndex) target.push(seg);
  }
}

/** Where a range anchors for the ordering convention: its lowest live segment
 * start. A fully-decayed range (no segments) anchors nowhere; -1 sorts it
 * after everything real, and it still gets its delete. */
function anchorStart(nr: GNamedRange): number {
  return nr.segments.length === 0 ? -1 : Math.min(...nr.segments.map((s) => s.startIndex));
}

/** Clear fontSize to inherited AND clear the hide white to inherited — the
 * "i"-entry restore, normalize, and the sweep all share this exact shape (empty
 * style + BOTH fields named, HIDE_FIELDS). Reveal sheds the invisibility color
 * with the size because size is the source of truth, not color: any text we
 * un-shrink here is text we (or a decayed copy of us) painted white, so the white
 * must go too — that is what keeps a copied / range-destroyed doc fully readable
 * after Show All, not just un-shrunk. (Foreign-size text inside a hidden region
 * is never in scope — the inventory filter only ever feeds us sentinel-size
 * sub-spans, plan A6 — so its color is never cleared.) */
function clearSizeRequest(range: GRange): UpdateTextStyleRequest {
  return { updateTextStyle: { range, textStyle: {}, fields: HIDE_FIELDS } };
}

/** Write an explicit recorded size back AND clear the hide white to inherited
 * (foregroundColor named in HIDE_FIELDS but absent from the style). Same reason
 * as clearSizeRequest: a recorded-size restore reveals text we painted white, so
 * it sheds the white in the same write. */
function setSizeRequest(range: GRange, sizePt: number): UpdateTextStyleRequest {
  return {
    updateTextStyle: {
      range,
      textStyle: { fontSize: { magnitude: sizePt, unit: "PT" } },
      fields: HIDE_FIELDS
    }
  };
}
