// Hide planner — Hide as a RECONCILE, never an append (plan A1/A2/A12/D6/D7;
// edge rows 6-8 and 12).
//
// The BLOCKER the reconcile design fixes (plan A1): on an armed doc, naive
// re-hide would evaluate keeper predicates against SENTINEL sizes — a hidden
// 1pt cite fails "bold AND >= citeMinPt" and a user's new highlight inside a
// hidden region would never resurface. So every Hide:
//   1. decodes the rstm "sizes" ranges and reconstructs the RESTORED VIEW —
//      what the doc would look like after a perfect Show All;
//   2. runs the keeper policy (keepers.planKeeps) against THAT view;
//   3. diffs the verdict against the CURRENT hidden state; and
//   4. emits the minimal request set: shrink what newly hides, restore what
//      newly keeps, leave everything already-correct untouched (idempotence —
//      a re-hide with nothing changed emits ZERO groups, edge row 7).
//
// Emission shape (plan A2): one contiguous hidden REGION = one RequestGroup of
// [createNamedRange anchor(s) whose NAME carries the original sizes as RLE,
// then ONE updateTextStyle shrinking the region to the sentinel]. Anchors
// always precede style writes inside a group: if an apply tears between them
// (it cannot — groups are atomic per guards.chunkGroups — but belt and
// suspenders), an anchor over normal-size text is a harmless no-op while
// sentinel text without an anchor is the amber-normalize path. Groups are
// ordered by DESCENDING region start as a convention only — none of the four
// request kinds we emit mutates content indexes (case 001-F1), so order is
// not load-bearing (plan A16).
//
// Invariants this module owns:
//   * only updateTextStyle(fontSize) / createNamedRange / deleteNamedRange
//     (+ updateParagraphStyle for the A12 spacing class) are ever emitted;
//   * rstm "sizes" ranges NEVER overlap, and regions are never merged across
//     re-hides — adjacent regions from different passes keep separate anchors;
//   * only kind:"text" sub-spans are ever style-targeted (whitelist, plan A9);
//   * the segment-final newline is never style-targeted (isLastInSegment
//     clamp, plan D6 — the API rejects styling it);
//   * pre-existing sentinel-size text with no rstm record is counted and left
//     alone (edge rows 8/12 — Show All's sweep owns adopting it, with consent
//     on the pure-sweep path);
//   * foreign-size text absorbed into a hidden region (pasted while hidden) is
//     never touched (plan A6) — it falls OUT of anchor coverage when its
//     region is reworked, and is otherwise simply left visible.

import { SENTINEL_PT, SENTINELS } from "./constants";
import { planKeeps } from "./keepers";
import { decodeRangeName, encodeSizeEntries, encodeSpacingName, isRstmName } from "./rangeNames";
// The ONE spacing-restore shape is owned by restore.ts (Show All is the
// primary emitter); the rework path shares it so the two can never drift.
import { spacingRestoreRequest } from "./restore";
import {
  CreateNamedRangeRequest,
  DocsRequest,
  GDoc,
  GdocsSettings,
  GElement,
  GHideResult,
  GParagraph,
  RequestGroup,
  RleEntry,
  UpdateParagraphStyleRequest,
  UpdateTextStyleRequest
} from "./types";

// ---------------------------------------------------------------------------
// Coverage — decoding the document's existing hidden state
// ---------------------------------------------------------------------------

/** A half-open [start, end) span — the internal lingua franca (GRange uses
 * startIndex/endIndex; normalizing here keeps one binary search reusable). */
interface Span {
  start: number;
  end: number;
}

/** One known "sizes" range with its decoded restore slices and, once atoms are
 * built, the atoms it covers (filled by buildAtoms — rework walks them). */
interface SizesCoverage {
  /** NamedRange id — what deleteNamedRange targets when the range is reworked. */
  id: string;
  /** Segments as read. Rework deletes the WHOLE range by id, so all segments
   * live or die together even when edits split the range. */
  segments: Span[];
  slices: CoverSlice[];
  /** Atoms under this range, in document order (text atoms only matter). */
  atoms: Atom[];
}

/** One restore slice: [start, end) restores to sizePt (null = clear to
 * inherit). Per-RLE-entry for a length-matching segment; ONE whole-segment
 * null slice for a mismatched one — the normalize rule: an interior-edited
 * segment's record no longer maps 1:1 onto its text, and guessing which chars
 * each entry covers risks restoring wrong sizes over wrong spans, so the
 * whole segment reconciles to inherit (plan A2's amber path; GHideResult has
 * no normalized counter — Show All owns the amber receipt). */
interface CoverSlice extends Span {
  sizePt: number | null;
  owner: SizesCoverage;
}

/** One known "spacing" record (plan A12): its anchor id plus the recorded
 * channels. The rework path needs all three — when a hidden paragraph
 * surfaces, its spacing record must die WITH its sizes anchor and the
 * recorded values must be written back (mirroring Show All's restore). */
interface SpacingCoverage {
  id: string;
  spaceAbovePt: number | null;
  spaceBelowPt: number | null;
}

/** A spacing record's extent with its owner — Span-shaped so the same spanAt
 * lookup serves both the double-record check and the rework restore. */
interface SpacingExtent extends Span {
  owner: SpacingCoverage;
}

/** Everything decodeCoverage learns about the doc's existing rstm state. */
interface Coverage {
  sizes: SizesCoverage[];
  /** All slices across all sizes ranges, sorted by start (lookup index). */
  slices: CoverSlice[];
  /** rstm-FAMILY ranges this build cannot decode (future versions, mangled
   * names). Their text is hidden state, just not ours to rewrite — left
   * untouched and excluded from fresh hiding (edge row 16; Show All warns). */
  unknownExtents: Span[];
  /** Known "spacing" record extents — re-hides must not double-record a
   * paragraph's spacing, and rework restores through them (plan A12). */
  spacingExtents: SpacingExtent[];
}

/** Binary search: the span containing pos, or null. Spans must be sorted by
 * start and non-overlapping — our own emissions guarantee that; a hand-edited
 * doc violating it degrades to "first match wins", which is safe (worst case
 * a slice restores as its sibling would have). */
function spanAt<T extends Span>(spans: T[], pos: number): T | null {
  let lo = 0;
  let hi = spans.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (spans[mid].start <= pos) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (found < 0) return null;
  return pos < spans[found].end ? spans[found] : null;
}

/** Decode the doc's named ranges into the planner's coverage model. Foreign
 * (non-rstm) ranges are invisible to us — other add-ons own them. */
function decodeCoverage(doc: GDoc): Coverage {
  const sizes: SizesCoverage[] = [];
  const slices: CoverSlice[] = [];
  const unknownExtents: Span[] = [];
  const spacingExtents: SpacingExtent[] = [];

  for (const nr of doc.namedRanges) {
    if (!isRstmName(nr.name)) continue;
    const segments: Span[] = nr.segments
      .filter((s) => s.endIndex > s.startIndex) // zero-width segments anchor nothing
      .map((s) => ({ start: s.startIndex, end: s.endIndex }));
    const decoded = decodeRangeName(nr.name);
    if (decoded === null) {
      unknownExtents.push(...segments); // ours-but-newer/corrupt: untouchable
      continue;
    }
    if (decoded.kind === "spacing") {
      // One owner per range even when edits split it: all extents share the
      // id, so the rework path deletes the record exactly once.
      const owner: SpacingCoverage = {
        id: nr.id,
        spaceAbovePt: decoded.spaceAbovePt,
        spaceBelowPt: decoded.spaceBelowPt
      };
      spacingExtents.push(...segments.map((s): SpacingExtent => ({ ...s, owner })));
      continue;
    }
    const cov: SizesCoverage = { id: nr.id, segments, slices: [], atoms: [] };
    const rleLen = decoded.entries.reduce((n, e) => n + e.count, 0);
    for (const seg of segments) {
      if (seg.end - seg.start === rleLen) {
        // Intact segment: each RLE entry maps to an exact sub-span.
        let cursor = seg.start;
        for (const e of decoded.entries) {
          cov.slices.push({ start: cursor, end: cursor + e.count, sizePt: e.sizePt, owner: cov });
          cursor += e.count;
        }
      } else {
        // Interior-edited (or split) segment: normalize to inherit (see
        // CoverSlice contract). A range split into N segments mismatches in
        // every segment by construction — exactly plan A2's degradation rule.
        cov.slices.push({ start: seg.start, end: seg.end, sizePt: null, owner: cov });
      }
    }
    sizes.push(cov);
    slices.push(...cov.slices);
  }
  slices.sort((a, b) => a.start - b.start);
  unknownExtents.sort((a, b) => a.start - b.start);
  spacingExtents.sort((a, b) => a.start - b.start);
  return { sizes, slices, unknownExtents, spacingExtents };
}

// ---------------------------------------------------------------------------
// Atoms — the diff's unit of account
// ---------------------------------------------------------------------------

/**
 * An atom is a maximal sub-span of one element that is UNIFORM in every fact
 * the reconcile cares about: same element (so same current size/bold/
 * highlight), same covering slice (so same restored size), same unknown-rstm
 * status. Elements are split at slice and unknown-extent boundaries — this is
 * where the plan's "sub-element spans arise at RLE entry boundaries" comes
 * from. Atoms double as the RESTORED view's elements, so keeper verdicts map
 * back 1:1 with no index juggling.
 */
interface Atom {
  start: number;
  end: number;
  kind: "text" | "other";
  /** The element-text slice this atom covers ("" for kind "other"). */
  text: string;
  paragraph: GParagraph;
  element: GElement;
  currentSizePt: number | null;
  /** What the restored view shows here: the covering slice's recorded size
   * over CURRENTLY-SENTINEL text, else the current size unchanged. The
   * "currently sentinel" condition is plan A6: foreign-size text pasted into
   * a hidden region keeps its own size — restore never flattens it. */
  restoredSizePt: number | null;
  /** The known sizes range covering this atom, when any. */
  cover: SizesCoverage | null;
  underUnknownRstm: boolean;
  /** Text currently AT a sentinel size — the "hidden now" predicate. The
   * SENTINELS set (not just SENTINEL_PT) so text hidden by any shipped
   * version is recognized (plan A16, append-only set). */
  currentSentinel: boolean;
  /** Keeper verdict against the RESTORED view (filled after planKeeps). */
  kept: boolean;
}

/** Build the per-paragraph atom lists (see Atom contract). Also fills each
 * SizesCoverage's `atoms` so rework can walk its own territory directly. */
function buildAtoms(doc: GDoc, coverage: Coverage): Atom[][] {
  // Every boundary that can change an atom's facts mid-element, sorted once.
  // (Not deduped here — the per-element bounds walk drops adjacent repeats.)
  const cuts: number[] = [];
  for (const s of coverage.slices) cuts.push(s.start, s.end);
  for (const s of coverage.unknownExtents) cuts.push(s.start, s.end);
  cuts.sort((a, b) => a - b);

  // Monotonic cursor into the sorted `cuts`. Paragraphs and their elements are
  // visited in ascending, non-overlapping index order, so each cut is interior
  // to at most one element and the cursor only advances — making bounds
  // collection O(elements + cuts), not O(elements × cuts). This matters on the
  // ARMED reconcile path (plan A1), where `cuts` grows with the hidden span
  // count and a per-element full scan would go quadratic on a long doc.
  let ci = 0;
  return doc.paragraphs.map((p) =>
    p.elements.flatMap((el) => {
      // Whitelist (plan A9): non-text elements are one indivisible atom —
      // never style-targeted, never covered, always a region breaker.
      if (el.kind !== "text") {
        return [
          {
            start: el.startIndex,
            end: el.endIndex,
            kind: "other" as const,
            text: "",
            paragraph: p,
            element: el,
            currentSizePt: el.fontSizePt,
            restoredSizePt: el.fontSizePt,
            cover: null,
            underUnknownRstm: false,
            currentSentinel: false,
            kept: false
          }
        ];
      }
      // Advance the cursor past cuts at or before this element's start (they
      // belong to earlier elements, or are shared boundaries that are never
      // interior cuts), then collect only the cuts strictly inside this element.
      while (ci < cuts.length && cuts[ci] <= el.startIndex) ci++;
      // Bounds: element edges plus every cut strictly inside it.
      const bounds: number[] = [el.startIndex];
      for (let j = ci; j < cuts.length && cuts[j] < el.endIndex; j++) {
        if (cuts[j] !== bounds[bounds.length - 1]) bounds.push(cuts[j]);
      }
      bounds.push(el.endIndex);

      const atoms: Atom[] = [];
      for (let i = 0; i + 1 < bounds.length; i++) {
        const start = bounds[i];
        const end = bounds[i + 1];
        const slice = spanAt(coverage.slices, start);
        const currentSentinel = el.fontSizePt !== null && SENTINELS.includes(el.fontSizePt);
        const atom: Atom = {
          start,
          end,
          kind: "text",
          text: el.text.slice(start - el.startIndex, end - el.startIndex),
          paragraph: p,
          element: el,
          currentSizePt: el.fontSizePt,
          restoredSizePt: slice !== null && currentSentinel ? slice.sizePt : el.fontSizePt,
          cover: slice !== null ? slice.owner : null,
          underUnknownRstm: spanAt(coverage.unknownExtents, start) !== null,
          currentSentinel,
          kept: false
        };
        if (atom.cover !== null) atom.cover.atoms.push(atom);
        atoms.push(atom);
      }
      return atoms;
    })
  );
}

/**
 * The RESTORED view: the same GDoc with elements replaced by atoms carrying
 * restored sizes. Splitting elements is keeper-safe: highlight/bold/cite
 * predicates are per-element-uniform facts that splitting preserves, and the
 * whitespace bridge scans PAST consecutive whitespace-only pieces, so a split
 * gap still bridges as a unit. (One benign drift: a split can expose a
 * whitespace-only piece of a mixed element to the bridge — strictly the
 * over-keep direction, which is the house failure direction.)
 */
function restoredView(doc: GDoc, atomsPerParagraph: Atom[][]): GDoc {
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p, i) => ({
      ...p,
      elements: atomsPerParagraph[i].map(
        (a): GElement => ({
          startIndex: a.start,
          endIndex: a.end,
          kind: a.kind,
          text: a.text,
          fontSizePt: a.restoredSizePt,
          bold: a.element.bold,
          backgroundHex: a.element.backgroundHex,
          // Foreground rides through unchanged: the reconcile only ever
          // reasons about SIZE, so an atom's color is its source element's
          // color (atoms split an element on size/coverage boundaries, never
          // on color). The analytics keeper (keepers.isAnalytics) runs against
          // THIS restored view, so omitting foreground here would show it
          // `undefined`, mis-hide a body-sized analytics run on a fresh Hide
          // (003-S2) and fail to break an adjacent hidden region (Loop 003).
          foregroundHex: a.element.foregroundHex
        })
      )
    }))
  };
}

// ---------------------------------------------------------------------------
// Request builders (shared by fresh-hide and rework emission — DRY)
// ---------------------------------------------------------------------------

/** The style-targetable end of an atom: the segment-final newline can never
 * be style-targeted (plan D6 clamp), and it is by definition the last char of
 * an isLastInSegment paragraph. Ranges (anchors) are NOT clamped — they only
 * bookkeep and may legitimately cover that newline. */
function styleEnd(a: Atom): number {
  return a.paragraph.isLastInSegment ? Math.min(a.end, a.paragraph.endIndex - 1) : a.end;
}

/** Anchor createNamedRange(s) for one contiguous run starting at `start` with
 * the given RLE. encodeSizeEntries splits at entry boundaries when the name
 * would overflow NAME_MAX; each piece reports its char count, so every piece
 * pairs with its exact sub-range and restore stays exact across splits. */
function anchorsForRun(start: number, entries: RleEntry[]): CreateNamedRangeRequest[] {
  const out: CreateNamedRangeRequest[] = [];
  let cursor = start;
  for (const piece of encodeSizeEntries(entries)) {
    out.push({
      createNamedRange: { name: piece.name, range: { startIndex: cursor, endIndex: cursor + piece.charCount } }
    });
    cursor += piece.charCount;
  }
  return out;
}

/** The ONE shrink per region: fontSize-only field mask so bold/underline/
 * highlight/links survive untouched (plan D1). */
function sentinelShrink(start: number, end: number): UpdateTextStyleRequest {
  return {
    updateTextStyle: {
      range: { startIndex: start, endIndex: end },
      textStyle: { fontSize: { magnitude: SENTINEL_PT, unit: "PT" } },
      fields: "fontSize"
    }
  };
}

/** One restore write: a recorded size is materialized; a null (inherit)
 * record is restored by CLEARING — empty textStyle with fontSize in the field
 * mask (the documented clear-to-inherit semantics, plan D13). */
function restoreWrite(start: number, end: number, sizePt: number | null): UpdateTextStyleRequest {
  return {
    updateTextStyle: {
      range: { startIndex: start, endIndex: end },
      textStyle: sizePt === null ? {} : { fontSize: { magnitude: sizePt, unit: "PT" } },
      fields: "fontSize"
    }
  };
}

// ---------------------------------------------------------------------------
// planHide
// ---------------------------------------------------------------------------

/** A region of NEW hiding under construction: contiguous target-hidden atoms
 * not under any rstm coverage. Entries record the RESTORED sizes (== current
 * sizes here, since the atoms are uncovered and non-sentinel). */
interface NewRegion {
  start: number;
  end: number;
  entries: RleEntry[];
  /** Paragraph ordinals contributing text — drives paragraphsChanged. */
  paragraphOrdinals: Set<number>;
  /** Fully-hidden paragraphs whose direct spacing collapses with this region
   * (plan A12) — their requests ride in the SAME group as the region's. */
  spacingParas: GParagraph[];
}

/** A group with its sort key (region start) — sorted descending before
 * return; the key never leaves this module. */
interface KeyedGroup {
  sortKey: number;
  requests: DocsRequest[];
}

/** Merge sorted spans, treating TOUCHING (end === start) as one — the receipt
 * counts contiguous hidden PASSAGES, and two abutting anchors read as one
 * passage to the user even though their anchors are never merged. */
function countMergedSpans(spans: Span[]): number {
  const sorted = [...spans].sort((a, b) => a.start - b.start);
  let count = 0;
  let reach = -1;
  for (const s of sorted) {
    if (s.start > reach) count++;
    reach = Math.max(reach, s.end);
  }
  return count;
}

/**
 * Plan one Hide pass: the reconcile described in the module header. Pure —
 * no I/O, no host globals; the controller owns gates (suggestions, tabs) and
 * the apply loop. Returns ZERO groups when there is nothing to do (edge row
 * 12 — "Nothing to hide here").
 */
export function planHide(doc: GDoc, settings: GdocsSettings): { groups: RequestGroup[]; result: GHideResult } {
  // ---- 1+2: coverage, atoms, restored view, keeper verdicts ----------------
  const coverage = decodeCoverage(doc);
  const atomsPerParagraph = buildAtoms(doc, coverage);
  const keeps = planKeeps(restoredView(doc, atomsPerParagraph), settings);
  atomsPerParagraph.forEach((atoms, i) => {
    atoms.forEach((a, j) => {
      a.kept = keeps[i].keepWhole || keeps[i].elementKeep[j];
    });
  });

  const groups: KeyedGroup[] = [];

  // ---- 3+4a: NEW hidden regions --------------------------------------------
  // Target-hidden = not-kept text that is not under ANY rstm coverage (known
  // sizes, unknown version — neither is ours to re-anchor here) and not
  // already at a sentinel size (uncovered sentinel text is pre-existing tiny,
  // edge rows 8/12: counted below, never touched).
  const isFreshHide = (a: Atom): boolean =>
    a.kind === "text" && !a.kept && a.cover === null && !a.underUnknownRstm && !a.currentSentinel;

  const regions: NewRegion[] = [];
  // Pre-existing tiny runs are merged across atom boundaries the same way
  // regions are — the count is a passage count, not an element count.
  const tinySpans: Span[] = [];
  // Paragraph ordinal -> its region, for the spacing class; and the set of
  // paragraphs that contain anything that does NOT hide this pass.
  const regionOfParagraph = new Map<number, NewRegion>();
  const paragraphHasNonHidden = new Set<number>();

  let open: NewRegion | null = null;
  let openTiny: Span | null = null;
  for (const atoms of atomsPerParagraph) {
    for (const a of atoms) {
      // Pre-existing tiny tracking (independent of region tracking).
      const isTiny = a.currentSentinel && a.cover === null && !a.underUnknownRstm;
      if (isTiny) {
        if (openTiny !== null && openTiny.end === a.start) openTiny.end = a.end;
        else tinySpans.push((openTiny = { start: a.start, end: a.end }));
      } else {
        openTiny = null;
      }

      const effEnd = styleEnd(a);
      if (!isFreshHide(a) || effEnd <= a.start) {
        // A clamped-to-nothing final newline still counts as hiding for the
        // "fully hidden paragraph" test — it is unhideable by API rule, not
        // by keeper verdict — so only NON-fresh atoms mark the paragraph.
        if (!isFreshHide(a)) paragraphHasNonHidden.add(a.paragraph.index);
        open = null;
        continue;
      }
      if (open !== null && open.end === a.start) {
        open.end = effEnd;
        open.entries.push({ count: effEnd - a.start, sizePt: a.restoredSizePt });
      } else {
        open = {
          start: a.start,
          end: effEnd,
          entries: [{ count: effEnd - a.start, sizePt: a.restoredSizePt }],
          paragraphOrdinals: new Set(),
          spacingParas: []
        };
        regions.push(open);
      }
      open.paragraphOrdinals.add(a.paragraph.index);
      regionOfParagraph.set(a.paragraph.index, open);
    }
    // NO reset at paragraph boundaries: a region legitimately spans adjacent
    // fully-hidden paragraphs (their newlines are part of it — plan D6). Real
    // index gaps (table structural overhead) fail the `end === start`
    // contiguity check above and break the region on their own.
  }

  // ---- 4b: spacing class (plan A12, default OFF) ---------------------------
  // Only paragraphs that are FULLY hidden by a region emitted THIS pass: a
  // partially-kept paragraph keeps visible spacing context, and untouched
  // still-hidden regions emit no group to ride in (their spacing was handled
  // when they were hidden, or the toggle was off then — flipping the toggle
  // retroactively requires Show All + re-Hide, documented limitation).
  if (settings.collapseSpacing) {
    for (let i = 0; i < doc.paragraphs.length; i++) {
      const p = doc.paragraphs[i];
      const region = regionOfParagraph.get(p.index);
      if (region === undefined || paragraphHasNonHidden.has(p.index)) continue;
      // DIRECT non-zero spacing only: null inherits (the named style owns it,
      // and Apply-styles zeroes Normal's spacing doc-wide), and an existing
      // zero has nothing to collapse.
      if ((p.spaceAbovePt ?? 0) <= 0 && (p.spaceBelowPt ?? 0) <= 0) continue;
      // Never double-record: a paragraph already carrying a spacing anchor
      // (re-hide after a hand-edit) keeps its ORIGINAL record — two records
      // would fight at restore time.
      if (spanAt(coverage.spacingExtents, p.startIndex) !== null) continue;
      region.spacingParas.push(p);
    }
  }

  for (const region of regions) {
    const requests: DocsRequest[] = [];
    // All anchors precede all style writes (see module header for why).
    requests.push(...anchorsForRun(region.start, region.entries));
    for (const p of region.spacingParas) {
      // The spacing anchor covers the PARAGRAPH (the unit being restored) —
      // it may overlap the region's sizes anchor; the never-overlap invariant
      // is per record family (a sizes range never overlaps a sizes range).
      requests.push({
        createNamedRange: {
          name: encodeSpacingName(p.spaceAbovePt, p.spaceBelowPt),
          range: { startIndex: p.startIndex, endIndex: p.endIndex }
        }
      });
    }
    requests.push(sentinelShrink(region.start, region.end));
    for (const p of region.spacingParas) {
      // updateParagraphStyle targets paragraphs OVERLAPPING the range, so the
      // segment-final clamp (which can only trim the very last char) still
      // hits the paragraph; clamping keeps the no-style-past-the-final-newline
      // rule uniform across every style-write we emit.
      const end = p.isLastInSegment ? p.endIndex - 1 : p.endIndex;
      const req: UpdateParagraphStyleRequest = {
        updateParagraphStyle: {
          range: { startIndex: p.startIndex, endIndex: end },
          paragraphStyle: {
            spaceAbove: { magnitude: 0, unit: "PT" },
            spaceBelow: { magnitude: 0, unit: "PT" }
          },
          fields: "spaceAbove,spaceBelow"
        }
      };
      requests.push(req);
    }
    groups.push({ sortKey: region.start, requests });
  }

  // ---- 4c: rework of regions where something became KEPT (plan A1) ---------
  // Trigger: any currently-sentinel sub-span of a known range that the
  // restored-view keeper now keeps (new highlight, keep-color toggle,
  // citeMinPt change, style change). The group restores those sub-spans to
  // their recorded sizes, drops the stale anchor, and re-anchors ONLY the
  // still-hidden remainder with a recomputed RLE. Foreign-size (non-sentinel)
  // text inside the range is neither restored nor re-anchored — plan A6.
  // Every restored sub-span across all reworks, for the receipt: the user
  // sees PASSAGES come back, so adjacent sub-spans (split at RLE entry
  // boundaries into several writes) must COUNT as one (countMergedSpans).
  const restoredSpans: Span[] = [];
  const untouchedHiddenSpans: Span[] = [];
  // Spacing records already retired this pass — global across reworks so a
  // (decay-split) record shared by two sizes ranges dies exactly once.
  const reworkedSpacingIds = new Set<string>();
  for (const cov of coverage.sizes) {
    const keptSentinel = cov.atoms.filter((a) => a.currentSentinel && a.kept);
    if (keptSentinel.length === 0) {
      // Untouched. It reads as an already-hidden passage only if something is
      // actually still hidden under it; an anchor whose text was all resized
      // away by the user is dead weight Show All will clean up.
      if (cov.atoms.some((a) => a.currentSentinel && !a.kept)) {
        untouchedHiddenSpans.push(...cov.segments);
      }
      continue;
    }

    const requests: DocsRequest[] = [{ deleteNamedRange: { namedRangeId: cov.id } }];

    // Spacing restore for surfacing paragraphs (plan A12's reverse): a
    // paragraph with newly-kept text is visible again, so the zeroed spacing
    // its record describes must come back IN THE SAME atomic group — a torn
    // pair would leave a visible paragraph squashed with no record, or a
    // live record over restored spacing that a later Show All re-fights.
    // Only a record that exactly tiles the paragraph is ours to honor (the
    // pinned invariant: spacing ranges tile whole paragraphs; this planner
    // emits one per paragraph) — a drifted extent stays untouched and decays
    // into Show All's amber normalize, never a guessed write here.
    const spacingRestores: DocsRequest[] = [];
    const surfacedParagraphs = new Map<number, GParagraph>();
    for (const a of keptSentinel) surfacedParagraphs.set(a.paragraph.index, a.paragraph);
    for (const p of surfacedParagraphs.values()) {
      const ext = spanAt(coverage.spacingExtents, p.startIndex);
      if (ext === null || reworkedSpacingIds.has(ext.owner.id)) continue;
      const tilesParagraph =
        ext.start === p.startIndex &&
        // Tolerate the planner's own clamp variant on the segment-final
        // paragraph (restore.ts extends the same tolerance).
        (ext.end === p.endIndex || (p.isLastInSegment && ext.end === p.endIndex - 1));
      if (!tilesParagraph) continue;
      reworkedSpacingIds.add(ext.owner.id);
      requests.push({ deleteNamedRange: { namedRangeId: ext.owner.id } });
      // The restore write clamps off the segment-final newline — the uniform
      // no-style-past-it rule every planner style write follows.
      const end = p.isLastInSegment ? p.endIndex - 1 : p.endIndex;
      if (end <= p.startIndex) continue; // empty final paragraph: nothing stylable
      spacingRestores.push(spacingRestoreRequest({ startIndex: p.startIndex, endIndex: end }, ext.owner.spaceAbovePt, ext.owner.spaceBelowPt));
    }

    // Re-anchor the still-hidden remainder: contiguous runs of sentinel,
    // not-kept atoms, RLE'd from their RESTORED sizes (exactness preserved
    // for intact records; normalized records re-anchor as inherit).
    let run: { start: number; entries: RleEntry[] } | null = null;
    let runEnd = -1;
    const flushRun = (): void => {
      if (run !== null) requests.push(...anchorsForRun(run.start, run.entries));
      run = null;
    };
    for (const a of cov.atoms) {
      if (!(a.currentSentinel && !a.kept)) {
        flushRun();
        continue;
      }
      if (run !== null && runEnd === a.start) {
        run.entries.push({ count: a.end - a.start, sizePt: a.restoredSizePt });
      } else {
        flushRun();
        run = { start: a.start, entries: [{ count: a.end - a.start, sizePt: a.restoredSizePt }] };
      }
      runEnd = a.end;
    }
    flushRun();

    // Restore writes, one per newly-kept sub-span at a uniform recorded size
    // (sub-spans split at RLE entry boundaries by atom construction; adjacent
    // same-size atoms merge so the write count is minimal). The COUNT is a
    // different unit on purpose: a kept span crossing a size boundary needs
    // two writes but reads as ONE surfaced passage, so each flushed write
    // joins restoredSpans and newlyKeptRestored merges them at the end.
    let span: { start: number; end: number; sizePt: number | null } | null = null;
    const flushSpan = (): void => {
      if (span === null) return;
      // The same segment-final clamp as everywhere: a covered final newline
      // (possible only on decayed manifests) must not be style-targeted.
      if (span.end > span.start) {
        requests.push(restoreWrite(span.start, span.end, span.sizePt));
        restoredSpans.push({ start: span.start, end: span.end });
      }
      span = null;
    };
    for (const a of keptSentinel) {
      const effEnd = styleEnd(a);
      if (span !== null && span.end === a.start && span.sizePt === a.restoredSizePt) {
        span.end = effEnd;
      } else {
        flushSpan();
        span = { start: a.start, end: effEnd, sizePt: a.restoredSizePt };
      }
    }
    flushSpan();
    // Spacing restores ride last: they are style writes, and the group's
    // anchors-precede-style-writes convention is per group, not per kind.
    requests.push(...spacingRestores);

    groups.push({ sortKey: Math.min(...cov.segments.map((s) => s.start)), requests });
  }

  // ---- 5: already-hidden passages (untouched known + unknown-version) ------
  // Unknown rstm extents count when they cover sentinel text: they ARE hidden
  // state (guards.hasRstmState agrees), we just refuse to rewrite it. Extents
  // covering nothing sentinel are invisible here — nothing reads as hidden.
  const sentinelUnknownSpans = coverage.unknownExtents.filter((ext) =>
    atomsPerParagraph.some((atoms) =>
      atoms.some((a) => a.underUnknownRstm && a.currentSentinel && a.start >= ext.start && a.start < ext.end)
    )
  );

  // ---- 6: result ------------------------------------------------------------
  const changed = new Set<number>();
  for (const region of regions) for (const ord of region.paragraphOrdinals) changed.add(ord);

  const result: GHideResult = {
    paragraphsScanned: doc.paragraphs.length,
    paragraphsChanged: changed.size,
    regionsHidden: regions.length,
    regionsAlreadyHidden: countMergedSpans([...untouchedHiddenSpans, ...sentinelUnknownSpans]),
    // Merged at the END so adjacent sub-spans restored by separate writes —
    // even across two abutting reworked regions — read as one passage.
    newlyKeptRestored: countMergedSpans(restoredSpans),
    preexistingTinyCount: tinySpans.length
  };

  // Descending region start across groups — convention only (see header).
  groups.sort((a, b) => b.sortKey - a.sortKey);
  return { groups: groups.map((g): RequestGroup => ({ requests: g.requests })), result };
}
