// Show All + convergence sweep suite for google-docs/src/core/restore.ts (plan S8;
// D3/A2/A6/A14; edge rows 9-11, 16). Fixtures are built AS THE PLANNER WOULD
// LEAVE THEM — sentinel-size text anchored by rstm ranges whose names encode
// the original sizes — using the shared builders plus the real rangeNames
// codec, so restore is tested against the same grammar Hide writes. Failure
// paths (decayed ranges, foreign edits, consent declined, future versions)
// get the same weight as the happy round-trip, because Show All's whole job
// is converging from damage.

import { planShowAll, pureSweepConsentNeeded, ShowAllPlan } from "../google-docs/src/core/restore";
import { SENTINEL_PT } from "../google-docs/src/core/constants";
import { encodeSizeEntries, isKnownRstmName } from "../google-docs/src/core/rangeNames";
import { DocsRequest, GDoc, GRange, GShowAllResult } from "../google-docs/src/core/types";
import { buildDoc, para, r, range } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** Expected-result builder: all-zero baseline so each test states ONLY the
 * counters it expects to move (a stray increment then fails loudly). */
function res(over: Partial<GShowAllResult> = {}): GShowAllResult {
  return {
    segmentsRestoredExact: 0,
    segmentsNormalized: 0,
    sweptOrphans: 0,
    rangesDeleted: 0,
    rangesSkippedNewerVersion: 0,
    ...over
  };
}

/** All requests across groups, flattened in emission order. */
function flat(plan: ShowAllPlan): DocsRequest[] {
  return plan.groups.flatMap((g) => g.requests);
}

/** A text-size write as the tests reason about it: sizePt null = clear. */
interface SizeWrite {
  start: number;
  end: number;
  sizePt: number | null;
}

function sizeWrites(plan: ShowAllPlan): SizeWrite[] {
  const out: SizeWrite[] = [];
  for (const req of flat(plan)) {
    if (!("updateTextStyle" in req)) continue;
    const u = req.updateTextStyle;
    out.push({
      start: u.range.startIndex,
      end: u.range.endIndex,
      sizePt: u.textStyle.fontSize ? u.textStyle.fontSize.magnitude : null
    });
  }
  return out;
}

/** Order-insensitive size-write comparison (emission order is a convention
 * pinned separately; geometry + values are the correctness contract). */
function expectSizeWrites(plan: ShowAllPlan, expected: SizeWrite[]): void {
  const byStart = (a: SizeWrite, b: SizeWrite) => a.start - b.start;
  expect([...sizeWrites(plan)].sort(byStart)).toEqual([...expected].sort(byStart));
}

/** A spacing write as the tests reason about it: null channel = cleared. */
interface SpacingWrite {
  start: number;
  end: number;
  abovePt: number | null;
  belowPt: number | null;
}

function spacingWrites(plan: ShowAllPlan): SpacingWrite[] {
  const out: SpacingWrite[] = [];
  for (const req of flat(plan)) {
    if (!("updateParagraphStyle" in req)) continue;
    const u = req.updateParagraphStyle;
    out.push({
      start: u.range.startIndex,
      end: u.range.endIndex,
      abovePt: u.paragraphStyle.spaceAbove ? u.paragraphStyle.spaceAbove.magnitude : null,
      belowPt: u.paragraphStyle.spaceBelow ? u.paragraphStyle.spaceBelow.magnitude : null
    });
  }
  return out;
}

function deletedIds(plan: ShowAllPlan): string[] {
  return flat(plan).flatMap((req) => ("deleteNamedRange" in req ? [req.deleteNamedRange.namedRangeId] : []));
}

/** Every range any emitted request would style (text or paragraph). */
function styledRanges(plan: ShowAllPlan): GRange[] {
  return flat(plan).flatMap((req) => {
    if ("updateTextStyle" in req) return [req.updateTextStyle.range];
    if ("updateParagraphStyle" in req) return [req.updateParagraphStyle.range];
    return [];
  });
}

/** Assert NO emitted request overlaps [start, end) — the hands-off check for
 * unknown-version ranges, foreign-size text, and table content. */
function expectNothingTouches(plan: ShowAllPlan, start: number, end: number): void {
  for (const rg of styledRanges(plan)) {
    const overlap = Math.min(rg.endIndex, end) - Math.max(rg.startIndex, start);
    expect(overlap).toBeLessThanOrEqual(0);
  }
}

/**
 * The Show All invariants (plan A11.iii / 001-F1), assertable on ANY plan:
 *   * only updateTextStyle / updateParagraphStyle / deleteNamedRange — never
 *     createNamedRange, never any content-mutating request;
 *   * updateTextStyle carries fontSize + foregroundColor (reveal clears the
 *     foreground to inherit alongside the size); updateParagraphStyle touches
 *     ONLY the spacing fields; sizes written are positive PT;
 *   * no text-style range reaches the segment-final newline (plan D6 clamp);
 *   * deletes target every KNOWN rstm range exactly once, nothing else, and
 *     each delete CLOSES its group (rides with its restores — A11.viii).
 */
function expectShowAllInvariants(doc: GDoc, plan: ShowAllPlan): void {
  const lastPara = doc.paragraphs[doc.paragraphs.length - 1];
  const finalNewlineIndex = lastPara.endIndex - 1;
  const deleted: string[] = [];
  for (const group of plan.groups) {
    expect(group.requests.length).toBeGreaterThan(0);
    group.requests.forEach((req, i) => {
      const keys = Object.keys(req);
      expect(keys).toHaveLength(1);
      expect(["updateTextStyle", "updateParagraphStyle", "deleteNamedRange"]).toContain(keys[0]);
      if ("updateTextStyle" in req) {
        const u = req.updateTextStyle;
        expect(u.fields).toBe("fontSize,foregroundColor");
        expect(u.range.endIndex).toBeGreaterThan(u.range.startIndex);
        expect(u.range.endIndex).toBeLessThanOrEqual(finalNewlineIndex);
        if (u.textStyle.fontSize) {
          expect(u.textStyle.fontSize.unit).toBe("PT");
          expect(u.textStyle.fontSize.magnitude).toBeGreaterThan(0);
        }
      }
      if ("updateParagraphStyle" in req) {
        const u = req.updateParagraphStyle;
        expect(u.fields).toBe("spaceAbove,spaceBelow");
        // A spacing restore obeys the same no-style-past-the-final-newline clamp
        // the text-style writes do: over the segment-final paragraph it must stop
        // at or before the final newline (regression guard — the planner stores
        // the anchor UNCLAMPED, so restore is the one that has to clamp the write).
        expect(u.range.endIndex).toBeGreaterThan(u.range.startIndex);
        expect(u.range.endIndex).toBeLessThanOrEqual(finalNewlineIndex);
      }
      if ("deleteNamedRange" in req) {
        deleted.push(req.deleteNamedRange.namedRangeId);
        expect(i).toBe(group.requests.length - 1);
      }
    });
  }
  const knownIds = doc.namedRanges.filter((nr) => isKnownRstmName(nr.name)).map((nr) => nr.id);
  expect([...deleted].sort()).toEqual([...knownIds].sort());
  expect(plan.result.rangesDeleted).toBe(knownIds.length);
}

// ---------------------------------------------------------------------------
// Exact round-trip restore (plan A2; case 001-S4)
// ---------------------------------------------------------------------------

describe("planShowAll — exact restore", () => {
  // As the planner leaves a hidden paragraph: both elements shrunk to the
  // sentinel, original sizes (11pt explicit + inherited) recorded in the name.
  // Indexes: "Keep me\n" [1,9) | "Hello " [9,15) "world\n" [15,21) | "After\n" [21,27).
  const doc = buildDoc(
    [
      para("Keep me"),
      { elements: [{ text: "Hello ", size: SENTINEL_PT }, { text: "world", size: SENTINEL_PT }] },
      para("After")
    ],
    { namedRanges: [range("nr1", "rstm:v1:6x11,6xi", 9, 21)] }
  );

  it("restores each RLE entry over its exact sub-range — explicit size set, 'i' cleared", () => {
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 9, end: 15, sizePt: 11 },
      { start: 15, end: 21, sizePt: null }
    ]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("rides the restores and the anchor delete in ONE atomic group, delete last", () => {
    const plan = planShowAll(doc, false);
    expect(plan.groups).toHaveLength(1);
    const requests = plan.groups[0].requests;
    expect(requests).toHaveLength(3);
    expect("updateTextStyle" in requests[0]).toBe(true);
    expect("updateTextStyle" in requests[1]).toBe(true);
    expect(requests[2]).toEqual({ deleteNamedRange: { namedRangeId: "nr1" } });
  });

  it("counts a segment exact even when nothing is left at the sentinel (user already revealed it)", () => {
    // Select-all + font-size-11 by hand leaves the record matching but the
    // text foreign-size: restore writes nothing, the stale anchor still dies.
    // "Keep me\n" [1,9) | "Visible\n" [9,17) at explicit 11pt | "End\n" [17,21).
    const revealed = buildDoc(
      [para("Keep me"), { elements: [{ text: "Visible", size: 11 }] }, para("End")],
      { namedRanges: [range("nr1", "rstm:v1:8x14", 9, 17)] }
    );
    const plan = planShowAll(revealed, false);
    expect(sizeWrites(plan)).toEqual([]);
    expect(deletedIds(plan)).toEqual(["nr1"]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
  });
});

// ---------------------------------------------------------------------------
// Split-region restore: a NAME_MAX overflow splits one region across several
// ranges at entry boundaries — restore must stay EXACT across all pieces
// (plan A2 / case 001-S4). Built with the REAL codec so the planner's
// emission grammar and this restore can never drift apart.
// ---------------------------------------------------------------------------

describe("planShowAll — region split across ranges by NAME_MAX (exact, via the shared codec)", () => {
  it("restores every original size across all pieces of an overflow-split region", () => {
    // 50 one-char elements with alternating original sizes: alternation
    // defeats RLE coalescing, forcing one entry per element and a name long
    // enough to split. The doc holds them all AT the sentinel.
    const COUNT = 50;
    const originalSizes = Array.from({ length: COUNT }, (_, i) => (i % 2 === 0 ? 10 : 11));
    const elements = originalSizes.map(() => ({ text: "x", size: SENTINEL_PT }));
    // Build once unanchored to learn the element geometry (the builder gives
    // the final element the trailing newline, so its entry covers 2 chars).
    const layout = buildDoc([{ elements }, para("After")]);
    const built = layout.paragraphs[0].elements;
    const entries = built.map((el, i) => ({ count: el.endIndex - el.startIndex, sizePt: originalSizes[i] }));

    const pieces = encodeSizeEntries(entries);
    expect(pieces.length).toBe(2); // pin: this fixture really exercises the split
    let cursor = built[0].startIndex;
    const namedRanges = pieces.map((p, i) => {
      const nr = range(`nr${i}`, p.name, cursor, cursor + p.charCount);
      cursor += p.charCount;
      return nr;
    });
    expect(cursor).toBe(layout.paragraphs[0].endIndex); // pieces tile the region exactly

    const doc = buildDoc([{ elements }, para("After")], { namedRanges });
    const plan = planShowAll(doc, false);
    expectSizeWrites(
      plan,
      built.map((el, i) => ({ start: el.startIndex, end: el.endIndex, sizePt: originalSizes[i] }))
    );
    expect(plan.result).toEqual(res({ segmentsRestoredExact: pieces.length, rangesDeleted: pieces.length }));
    expectShowAllInvariants(doc, plan);
  });
});

// ---------------------------------------------------------------------------
// Per-segment exactness + normalize (plan A2; edge rows 9, 10)
// ---------------------------------------------------------------------------

describe("planShowAll — segment mismatch normalizes (edge rows 9, 10)", () => {
  it("clears a mismatched segment to inherit instead of guessing sizes", () => {
    // Interior edit shrank the segment: record says 20 chars, doc has 12.
    // "Keep me\n" [1,9) | "Hello world\n" [9,21) sentinel | "After\n" [21,27).
    const doc = buildDoc(
      [para("Keep me"), { elements: [{ text: "Hello world", size: SENTINEL_PT }] }, para("After")],
      { namedRanges: [range("nr1", "rstm:v1:20x11", 9, 21)] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [{ start: 9, end: 21, sizePt: null }]);
    expect(plan.result).toEqual(res({ segmentsNormalized: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("judges each segment of a Docs-split range on its own length (exact + normalized in one range)", () => {
    // An edit split the range AND changed its content: the segments sum to 17
    // chars against a 12-char record, so this is NOT a pure split (no
    // cross-segment walk). The surviving first segment still matches the RLE
    // total on its own; the second is a remnant that does not.
    // "Hello world\n" [1,13) sentinel | "Keep\n" [13,18) | "tiny\n" [18,23) sentinel | "End\n" [23,27).
    const doc = buildDoc(
      [
        { elements: [{ text: "Hello world", size: SENTINEL_PT }] },
        para("Keep"),
        { elements: [{ text: "tiny", size: SENTINEL_PT }] },
        para("End")
      ],
      { namedRanges: [{ id: "nr1", name: "rstm:v1:12x14", segments: [r(1, 13), r(18, 23)] }] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 1, end: 13, sizePt: 14 }, // matching segment: exact restore
      { start: 18, end: 23, sizePt: null } // remnant: normalized
    ]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, segmentsNormalized: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("skips a zero-length decayed segment without counting it (no passage, no reset)", () => {
    // "Keep me\n" [1,9) | "Visible\n" [9,17) | "End\n" [17,21).
    const doc = buildDoc([para("Keep me"), para("Visible"), para("End")], {
      namedRanges: [{ id: "nrz", name: "rstm:v1:1x11", segments: [r(9, 9)] }]
    });
    const plan = planShowAll(doc, false);
    expect(styledRanges(plan)).toEqual([]);
    expect(deletedIds(plan)).toEqual(["nrz"]);
    expect(plan.result).toEqual(res({ rangesDeleted: 1 }));
  });

  it("still deletes a range whose segments are all gone (fully decayed anchor)", () => {
    const doc = buildDoc([para("Keep me"), para("End")], {
      namedRanges: [{ id: "nr0", name: "rstm:v1:1x11", segments: [] }]
    });
    const plan = planShowAll(doc, false);
    expect(plan.groups).toHaveLength(1);
    expect(plan.groups[0].requests).toEqual([{ deleteNamedRange: { namedRangeId: "nr0" } }]);
    expect(plan.result).toEqual(res({ rangesDeleted: 1 }));
  });
});

// ---------------------------------------------------------------------------
// PURE splits restore exactly (plan A2 sharpened): when a range's segments,
// in document order, sum to the RLE's total char count, the split merely
// removed indexes BETWEEN the pieces — the RLE walks continuously across
// them and every character restores byte-exact, with NO amber count.
// ---------------------------------------------------------------------------

describe("planShowAll — pure splits walk the RLE across segments (exact, no amber)", () => {
  // "Keep\n" [1,6) | "abcde\n" [6,12) sentinel | "Mid\n" [12,16) kept |
  // "fgh\n" [16,20) sentinel | "End\n" [20,24). The record "7x12,3x9" sums to
  // 10 == 6 + 4, and its first entry STRADDLES the split point — the walk
  // must carry it across the gap.
  const splitParas = [
    para("Keep"),
    { elements: [{ text: "abcde", size: SENTINEL_PT }] },
    para("Mid"),
    { elements: [{ text: "fgh", size: SENTINEL_PT }] },
    para("End")
  ];

  it("restores every recorded size across both fragments — entry straddling the split point included", () => {
    const doc = buildDoc(splitParas, {
      namedRanges: [{ id: "nr1", name: "rstm:v1:7x12,3x9", segments: [r(6, 12), r(16, 20)] }]
    });
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 6, end: 12, sizePt: 12 }, // entry 1's first 6 chars
      { start: 16, end: 17, sizePt: 12 }, // entry 1's 7th char, after the hop
      { start: 17, end: 20, sizePt: 9 } // entry 2
    ]);
    // BOTH segments count exact — the user sees two passages come back, and
    // neither took the amber path.
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 2, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("sorts segments into document order before walking (Docs reports them unordered)", () => {
    const doc = buildDoc(splitParas, {
      namedRanges: [{ id: "nr1", name: "rstm:v1:7x12,3x9", segments: [r(16, 20), r(6, 12)] }]
    });
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 6, end: 12, sizePt: 12 },
      { start: 16, end: 17, sizePt: 12 },
      { start: 17, end: 20, sizePt: 9 }
    ]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 2, rangesDeleted: 1 }));
  });

  it("still scopes writes to sentinel text inside a pure split (A6 holds across the walk)", () => {
    // The user already revealed the second fragment by hand (14pt): the
    // record still matches, so the range is exact, but no write may flatten
    // the user's size — the walk emits nothing over the revealed piece.
    const doc = buildDoc(
      [
        para("Keep"),
        { elements: [{ text: "abcde", size: SENTINEL_PT }] },
        para("Mid"),
        { elements: [{ text: "fgh", size: 14 }] },
        para("End")
      ],
      { namedRanges: [{ id: "nr1", name: "rstm:v1:7x12,3x9", segments: [r(6, 12), r(16, 20)] }] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [{ start: 6, end: 12, sizePt: 12 }]);
    expectNothingTouches(plan, 16, 20);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 2, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });
});

// ---------------------------------------------------------------------------
// Foreign-size text inside a hidden region is NEVER flattened (plan A6)
// ---------------------------------------------------------------------------

describe("planShowAll — foreign-size text absorbed into a range is untouched (plan A6)", () => {
  it("exact path: restores around an in-place 14pt resize, never over it", () => {
    // User selected part of the hidden region and set it to 14pt (no length
    // change, so the segment still matches its record).
    // "Keep me\n" [1,9) | "abc" [9,12) 1pt, "def" [12,15) 14pt, "gh\n" [15,18) 1pt | "After\n" [18,24).
    const doc = buildDoc(
      [
        para("Keep me"),
        {
          elements: [
            { text: "abc", size: SENTINEL_PT },
            { text: "def", size: 14 },
            { text: "gh", size: SENTINEL_PT }
          ]
        },
        para("After")
      ],
      { namedRanges: [range("nr1", "rstm:v1:9x11", 9, 18)] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 9, end: 12, sizePt: 11 },
      { start: 15, end: 18, sizePt: 11 }
    ]);
    expectNothingTouches(plan, 12, 15);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("normalize path: a 14pt paste grew the segment — sentinel text clears, the paste keeps its size", () => {
    // "Keep me\n" [1,9) | "abc" [9,12) 1pt, "PASTED" [12,18) 14pt, "gh\n" [18,21) 1pt | "After\n" [21,27).
    const doc = buildDoc(
      [
        para("Keep me"),
        {
          elements: [
            { text: "abc", size: SENTINEL_PT },
            { text: "PASTED", size: 14 },
            { text: "gh", size: SENTINEL_PT }
          ]
        },
        para("After")
      ],
      { namedRanges: [range("nr1", "rstm:v1:6x11", 9, 21)] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 9, end: 12, sizePt: null },
      { start: 18, end: 21, sizePt: null }
    ]);
    expectNothingTouches(plan, 12, 18);
    expect(plan.result).toEqual(res({ segmentsNormalized: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });
});

// ---------------------------------------------------------------------------
// Convergence sweep (plan D3; edge rows 10, 11)
// ---------------------------------------------------------------------------

describe("planShowAll — convergence sweep", () => {
  it("sweeps a destroyed-range orphan WITHOUT consent while any rstm state exists", () => {
    // p2's anchor was destroyed by cut/paste; p1's survived. Orphans next to
    // our own state are ours — no consent gate (plan A14 "normal Show All").
    // "Keep\n" [1,6) | "hid\n" [6,10) anchored | "orphan tiny\n" [10,22) orphaned | "End\n" [22,26).
    const doc = buildDoc(
      [
        para("Keep"),
        { elements: [{ text: "hid", size: SENTINEL_PT }] },
        { elements: [{ text: "orphan tiny", size: SENTINEL_PT }] },
        para("End")
      ],
      { namedRanges: [range("nr1", "rstm:v1:4x11", 6, 10)] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 6, end: 10, sizePt: 11 }, // anchored: exact restore
      { start: 10, end: 22, sizePt: null } // orphan: swept to inherit
    ]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, sweptOrphans: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("sweeps an orphan sitting BEFORE the surviving anchor inside one merged passage", () => {
    // Mirror of the test above: the destroyed-range text precedes the
    // anchored text, so the merged sentinel span is holed at its END and the
    // sweep must emit the piece in front of the hole.
    // "Keep\n" [1,6) | "orphan tiny\n" [6,18) orphaned | "hid\n" [18,22) anchored | "End\n" [22,26).
    const doc = buildDoc(
      [
        para("Keep"),
        { elements: [{ text: "orphan tiny", size: SENTINEL_PT }] },
        { elements: [{ text: "hid", size: SENTINEL_PT }] },
        para("End")
      ],
      { namedRanges: [range("nr1", "rstm:v1:4x11", 18, 22)] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [
      { start: 6, end: 18, sizePt: null }, // orphan: swept
      { start: 18, end: 22, sizePt: 11 } // anchored: exact restore
    ]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, sweptOrphans: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("does NOT sweep text at a non-sentinel tiny size (1.5pt is the user's, not ours)", () => {
    // Exact-membership rule: only sizes in SENTINELS are reclaimable.
    // "Normal\n" [1,8) | "tiny\n" [8,13) 1pt | "Mid\n" [13,17) | "small\n" [17,23) 1.5pt | "End\n" [23,27).
    const doc = buildDoc([
      para("Normal"),
      { elements: [{ text: "tiny", size: SENTINEL_PT }] },
      para("Mid"),
      { elements: [{ text: "small", size: SENTINEL_PT + 0.5 }] },
      para("End")
    ]);
    expect(pureSweepConsentNeeded(doc)).toBe(1); // only the true-sentinel passage
    const plan = planShowAll(doc, true);
    expectSizeWrites(plan, [{ start: 8, end: 13, sizePt: null }]);
    expectNothingTouches(plan, 17, 23);
    expect(plan.result).toEqual(res({ sweptOrphans: 1 }));
  });

  it("sweeps text under a SURVIVING spacing record whose sizes anchor was destroyed", () => {
    // Spacing segments must not shield text from the sweep: a spacing record
    // says nothing about text sizes, and here it is the only thing left.
    // "Keep\n" [1,6) | "was hidden\n" [6,17) sentinel + spacing anchor | "End\n" [17,21).
    const doc = buildDoc(
      [para("Keep"), { elements: [{ text: "was hidden", size: SENTINEL_PT }] }, para("End")],
      { namedRanges: [range("p1", "rstm:v1:p:12x6", 6, 17)] }
    );
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [{ start: 6, end: 17, sizePt: null }]); // swept
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 17, abovePt: 12, belowPt: 6 }]); // restored
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, sweptOrphans: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });
});

// ---------------------------------------------------------------------------
// Pure-sweep consent boundary (plan A14)
// ---------------------------------------------------------------------------

describe("pureSweepConsentNeeded + the consent-gated sweep (plan A14)", () => {
  // Zero rstm ranges, two separate tiny passages (a copied doc's leftovers,
  // or the user's own formatting — we cannot tell, hence the consent).
  // "Normal text\n" [1,13) | "tiny one\n" [13,22) | "More normal\n" [22,34) | "tiny two\n" [34,43) | "End\n" [43,47).
  const unrecorded = buildDoc([
    para("Normal text"),
    { elements: [{ text: "tiny one", size: SENTINEL_PT }] },
    para("More normal"),
    { elements: [{ text: "tiny two", size: SENTINEL_PT }] },
    para("End")
  ]);

  it("counts unrecorded tiny passages when no rstm range exists", () => {
    expect(pureSweepConsentNeeded(unrecorded)).toBe(2);
  });

  it("is null while ANY rstm range exists — even one this build cannot decode", () => {
    const armed = buildDoc([{ elements: [{ text: "hid", size: SENTINEL_PT }] }, para("End")], {
      namedRanges: [range("nr1", "rstm:v1:3x11", 1, 4)]
    });
    expect(pureSweepConsentNeeded(armed)).toBeNull();
    const armedByNewer = buildDoc([{ elements: [{ text: "hid", size: SENTINEL_PT }] }, para("End")], {
      namedRanges: [range("u1", "rstm:v2:future-grammar", 1, 4)]
    });
    expect(pureSweepConsentNeeded(armedByNewer)).toBeNull();
  });

  it("is null when there is nothing tiny to ask about", () => {
    expect(pureSweepConsentNeeded(buildDoc([para("Just"), para("text")]))).toBeNull();
  });

  it("declined consent: zero groups, zero counts, the tiny text stays the user's", () => {
    const plan = planShowAll(unrecorded, false);
    expect(plan.groups).toEqual([]);
    expect(plan.result).toEqual(res());
  });

  it("granted consent: sweeps every counted passage — sweep count equals the consent count", () => {
    const plan = planShowAll(unrecorded, true);
    expectSizeWrites(plan, [
      { start: 13, end: 22, sizePt: null },
      { start: 34, end: 43, sizePt: null }
    ]);
    expect(plan.result).toEqual(res({ sweptOrphans: 2 }));
    expect(plan.result.sweptOrphans).toBe(pureSweepConsentNeeded(unrecorded));
    expectShowAllInvariants(unrecorded, plan);
  });
});

// ---------------------------------------------------------------------------
// Unknown rstm versions (edge row 16)
// ---------------------------------------------------------------------------

describe("planShowAll — unknown rstm versions are counted, never touched (edge row 16)", () => {
  // "Keep\n" [1,6) | "newer\n" [6,12) under rstm:v2 | "orphan\n" [12,19) | "End\n" [19,23).
  const doc = buildDoc(
    [
      para("Keep"),
      { elements: [{ text: "newer", size: SENTINEL_PT }] },
      { elements: [{ text: "orphan", size: SENTINEL_PT }] },
      para("End")
    ],
    { namedRanges: [range("u1", "rstm:v2:future-grammar", 6, 12)] }
  );

  it("leaves the unknown range and its text alone, counts it in the RESULT, and still sweeps true orphans", () => {
    const plan = planShowAll(doc, false);
    expect(deletedIds(plan)).toEqual([]); // the v2 anchor is NOT ours to delete
    expectNothingTouches(plan, 6, 12); // nor is its text ours to restore or sweep
    // The doc IS armed (family-level), so the orphan next door sweeps freely.
    expectSizeWrites(plan, [{ start: 12, end: 19, sizePt: null }]);
    // The skipped count rides in the result so the receipt's row-16 amber
    // line renders from the same contract as every other counter.
    expect(plan.result).toEqual(res({ sweptOrphans: 1, rangesSkippedNewerVersion: 1 }));
    expectShowAllInvariants(doc, plan);
  });
});

// ---------------------------------------------------------------------------
// Spacing restore (plan A12)
// ---------------------------------------------------------------------------

describe("planShowAll — spacing ranges (plan A12)", () => {
  // Spacing fixtures use explicit 11pt text so the sweep stays quiet and the
  // asserts isolate the paragraph-style channel.
  // "Keep\n" [1,6) | "was hidden\n" [6,17) | "End\n" [17,21).
  const base = [para("Keep"), { elements: [{ text: "was hidden", size: 11 }] }, para("End")];

  it("restores both recorded values over an aligned paragraph", () => {
    const doc = buildDoc(base, { namedRanges: [range("p1", "rstm:v1:p:12x6", 6, 17)] });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 17, abovePt: 12, belowPt: 6 }]);
    // Counted with the exact restores: the record matched and was honored.
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("clears an 'i' channel via the field mask (absent from the style, named in fields)", () => {
    const doc = buildDoc(base, { namedRanges: [range("p1", "rstm:v1:p:ix4.5", 6, 17)] });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 17, abovePt: null, belowPt: 4.5 }]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
  });

  it("clears BOTH channels for an ixi record (paragraph inherited all its spacing)", () => {
    const doc = buildDoc(base, { namedRanges: [range("p1", "rstm:v1:p:ixi", 6, 17)] });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 17, abovePt: null, belowPt: null }]);
    const req = flat(plan)[0];
    if (!("updateParagraphStyle" in req)) throw new Error("expected updateParagraphStyle first");
    expect(req.updateParagraphStyle.paragraphStyle).toEqual({});
    expect(req.updateParagraphStyle.fields).toBe("spaceAbove,spaceBelow");
  });

  it("restores once over a segment tiling SEVERAL whole paragraphs", () => {
    // "Keep\n" [1,6) | "was hidden\n" [6,17) + "also\n" [17,22) | "End\n" [22,26).
    const doc = buildDoc(
      [para("Keep"), { elements: [{ text: "was hidden", size: 11 }] }, { elements: [{ text: "also", size: 11 }] }, para("End")],
      { namedRanges: [range("p1", "rstm:v1:p:12x6", 6, 22)] }
    );
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 22, abovePt: 12, belowPt: 6 }]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
  });

  it("accepts the clamped emission on the segment-final paragraph (range stops short of the last newline)", () => {
    // "Keep\n" [1,6) | "hid\n" [6,10) is the LAST paragraph; the planner may
    // have anchored [6,9) to keep the unstylable final newline out of ranges.
    const doc = buildDoc([para("Keep"), { elements: [{ text: "hid", size: 11 }] }], {
      namedRanges: [range("p1", "rstm:v1:p:3x0", 6, 9)]
    });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 9, abovePt: 3, belowPt: 0 }]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
  });

  it("CLAMPS the write on the segment-final paragraph when the planner anchored the FULL unclamped range", () => {
    // The planner stores the spacing anchor UNCLAMPED at [p.startIndex, p.endIndex]
    // — the createNamedRange legitimately covers the paragraph's final newline.
    // On Show All the restore must still clamp its updateParagraphStyle write off
    // that unstylable newline, like every other style emitter (regression: this
    // path used to write the raw [6,10) anchor over the final newline).
    // "Keep\n" [1,6) | "hid\n" [6,10) is the LAST paragraph; anchor = full [6,10).
    const doc = buildDoc([para("Keep"), { elements: [{ text: "hid", size: 11 }] }], {
      namedRanges: [range("p1", "rstm:v1:p:3x0", 6, 10)]
    });
    const plan = planShowAll(doc, false);
    // Written over [6,9) — clamped one short of the final newline — NOT [6,10).
    expect(spacingWrites(plan)).toEqual([{ start: 6, end: 9, abovePt: 3, belowPt: 0 }]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan);
  });

  it("counts a misaligned segment normalized and writes NOTHING (no sentinel to scope a safe write by)", () => {
    // Drifted to start mid-paragraph: the recorded values can no longer be
    // proven to belong to these paragraphs — leave the zeroed spacing visible.
    const doc = buildDoc(base, { namedRanges: [range("p1", "rstm:v1:p:12x6", 7, 17)] });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([]);
    expect(deletedIds(plan)).toEqual(["p1"]); // the decayed record still dies
    expect(plan.result).toEqual(res({ segmentsNormalized: 1, rangesDeleted: 1 }));
  });

  it("treats a segment over a TABLE paragraph as misaligned (tables untouched, edge row 1)", () => {
    // "Keep\n" [1,6) | "in table\n" [6,15) inTable | "End\n" [15,19).
    const doc = buildDoc(
      [para("Keep"), { inTable: true, elements: [{ text: "in table", size: 11 }] }, para("End")],
      { namedRanges: [range("p1", "rstm:v1:p:12x6", 6, 15)] }
    );
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([]);
    expect(plan.result).toEqual(res({ segmentsNormalized: 1, rangesDeleted: 1 }));
  });

  it("treats a segment that runs PAST the last paragraph as misaligned (decayed end)", () => {
    // Starts on a real paragraph boundary but claims indexes beyond the doc.
    const doc = buildDoc(base, { namedRanges: [range("p1", "rstm:v1:p:12x6", 6, 99)] });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([]);
    expect(plan.result).toEqual(res({ segmentsNormalized: 1, rangesDeleted: 1 }));
  });

  it("treats a segment ending MID-paragraph as misaligned (decayed interior)", () => {
    // Starts on the boundary but stops inside [6,17) — no longer whole.
    const doc = buildDoc(base, { namedRanges: [range("p1", "rstm:v1:p:12x6", 6, 10)] });
    const plan = planShowAll(doc, false);
    expect(spacingWrites(plan)).toEqual([]);
    expect(plan.result).toEqual(res({ segmentsNormalized: 1, rangesDeleted: 1 }));
  });
});

// ---------------------------------------------------------------------------
// Tables and the whitelist (edge rows 1, 2)
// ---------------------------------------------------------------------------

describe("planShowAll — table content is invisible to restore, sweep, and consent", () => {
  it("neither counts nor sweeps sentinel-size text inside a table", () => {
    // "Keep\n" [1,6) | "tiny table\n" [6,17) inTable 1pt | "orphan\n" [17,24) body 1pt | "End\n" [24,28).
    const doc = buildDoc([
      para("Keep"),
      { inTable: true, elements: [{ text: "tiny table", size: SENTINEL_PT }] },
      { elements: [{ text: "orphan", size: SENTINEL_PT }] },
      para("End")
    ]);
    expect(pureSweepConsentNeeded(doc)).toBe(1); // body orphan only
    const plan = planShowAll(doc, true);
    expectSizeWrites(plan, [{ start: 17, end: 24, sizePt: null }]);
    expectNothingTouches(plan, 6, 17);
    expect(plan.result).toEqual(res({ sweptOrphans: 1 }));
  });

  it("emits no text writes for a decayed record sitting over a table", () => {
    // Even an exact-length match writes nothing into a table: the inventory
    // excludes table paragraphs, so the scope filter holds structurally.
    const doc = buildDoc(
      [para("Keep"), { inTable: true, elements: [{ text: "tiny table", size: SENTINEL_PT }] }, para("End")],
      { namedRanges: [range("nr1", "rstm:v1:11x13", 6, 17)] }
    );
    const plan = planShowAll(doc, false);
    expect(sizeWrites(plan)).toEqual([]);
    expect(deletedIds(plan)).toEqual(["nr1"]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
  });
});

// ---------------------------------------------------------------------------
// Segment-final newline clamp (plan D6)
// ---------------------------------------------------------------------------

describe("planShowAll — the segment-final newline is never style-targeted (plan D6)", () => {
  it("restores a hidden FINAL paragraph up to, never including, the last newline", () => {
    // Lone paragraph "Hello world\n" [1,13), all sentinel; the planner's
    // clamped region was [1,12) with an 11-char record.
    const doc = buildDoc([{ elements: [{ text: "Hello world", size: SENTINEL_PT }] }], {
      namedRanges: [range("nr1", "rstm:v1:11x12", 1, 12)]
    });
    const plan = planShowAll(doc, false);
    expectSizeWrites(plan, [{ start: 1, end: 12, sizePt: 12 }]);
    expect(plan.result).toEqual(res({ segmentsRestoredExact: 1, rangesDeleted: 1 }));
    expectShowAllInvariants(doc, plan); // includes the <= finalNewline bound
  });

  it("sweeps a sentinel final paragraph with the same clamp", () => {
    const doc = buildDoc([{ elements: [{ text: "Hello world", size: SENTINEL_PT }] }]);
    expect(pureSweepConsentNeeded(doc)).toBe(1);
    const plan = planShowAll(doc, true);
    expectSizeWrites(plan, [{ start: 1, end: 12, sizePt: null }]);
    expect(plan.result).toEqual(res({ sweptOrphans: 1 }));
  });

  it("treats a sentinel-sized EMPTY final paragraph as nothing to do (clamp empties it)", () => {
    // "x\n" [1,3) | "\n" [3,4) at the sentinel — only the unstylable newline.
    const doc = buildDoc([para("x"), { elements: [{ text: "", size: SENTINEL_PT }] }]);
    expect(pureSweepConsentNeeded(doc)).toBeNull();
    const plan = planShowAll(doc, true);
    expect(plan.groups).toEqual([]);
    expect(plan.result).toEqual(res());
  });
});

// ---------------------------------------------------------------------------
// Idempotence: Show All converges in one pass (plan D3)
// ---------------------------------------------------------------------------

describe("planShowAll — second pass emits zero groups (idempotence)", () => {
  it("a restored doc (sizes back, anchors gone) plans nothing, with or without consent", () => {
    // The exact post-state of the round-trip fixture's plan: 11pt explicit,
    // inherited tail, no ranges left.
    const restored = buildDoc([
      para("Keep me"),
      { elements: [{ text: "Hello ", size: 11 }, { text: "world" }] },
      para("After")
    ]);
    expect(pureSweepConsentNeeded(restored)).toBeNull();
    for (const consent of [false, true]) {
      const plan = planShowAll(restored, consent);
      expect(plan.groups).toEqual([]);
      expect(plan.result).toEqual(res());
    }
  });

  it("a swept doc (everything cleared to inherit) plans nothing", () => {
    const swept = buildDoc([para("Normal text"), { elements: [{ text: "tiny one" }] }, para("End")]);
    expect(pureSweepConsentNeeded(swept)).toBeNull();
    expect(planShowAll(swept, true).groups).toEqual([]);
  });

  it("an empty/all-kept doc plans nothing (edge row 12's Show All face)", () => {
    const doc = buildDoc([para("Just a doc")]);
    expect(planShowAll(doc, false).groups).toEqual([]);
    expect(planShowAll(doc, false).result).toEqual(res());
  });
});

// ---------------------------------------------------------------------------
// Composite: every path at once + the full invariant battery (plan A11.iii)
// ---------------------------------------------------------------------------

describe("planShowAll — composite doc, invariants and emission conventions", () => {
  // "Keep me\n"[1,9) | A:"abcd\n"[9,14) exact | "Mid\n"[14,18) | B:"efgh\n"[18,23)
  // mismatch | "Mid2\n"[23,28) | U:"tiny\n"[28,33) rstm:v2 | "orph\n"[33,38)
  // orphan | P:"sp\n"[38,41) spacing | "End\n"[41,45).
  const doc = buildDoc(
    [
      para("Keep me"),
      { elements: [{ text: "abcd", size: SENTINEL_PT }] },
      para("Mid"),
      { elements: [{ text: "efgh", size: SENTINEL_PT }] },
      para("Mid2"),
      { elements: [{ text: "tiny", size: SENTINEL_PT }] },
      { elements: [{ text: "orph", size: SENTINEL_PT }] },
      { elements: [{ text: "sp", size: 11 }] },
      para("End")
    ],
    {
      namedRanges: [
        range("A", "rstm:v1:5x12", 9, 14),
        range("B", "rstm:v1:9x10", 18, 23),
        range("U", "rstm:v2:future", 28, 33),
        range("P", "rstm:v1:p:8xi", 38, 41)
      ]
    }
  );
  const plan = planShowAll(doc, false);

  it("moves every counter exactly once per path", () => {
    expect(plan.result).toEqual(
      res({
        segmentsRestoredExact: 2,
        segmentsNormalized: 1,
        sweptOrphans: 1,
        rangesDeleted: 3,
        rangesSkippedNewerVersion: 1
      })
    );
  });

  it("passes the full invariant battery", () => {
    expectShowAllInvariants(doc, plan);
    expectNothingTouches(plan, 28, 33); // the v2 range's text
  });

  it("emits range groups in descending anchor order, then sweep groups (non-load-bearing convention)", () => {
    // Pinned as a CONVENTION shared with the planner — correctness never
    // depends on it because no emitted request mutates indexes (001-F1).
    expect(plan.groups.map((g) => g.requests.length)).toEqual([2, 2, 2, 1]);
    expect(deletedIds(plan)).toEqual(["P", "B", "A"]); // descending anchors 38, 18, 9
    const sweepGroup = plan.groups[3].requests[0];
    if (!("updateTextStyle" in sweepGroup)) throw new Error("expected the sweep clear last");
    expect(sweepGroup.updateTextStyle.range).toEqual(r(33, 38));
  });
});
