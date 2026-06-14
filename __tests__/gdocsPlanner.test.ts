// Planner suite for google-docs/src/core/planner.ts (plan S7 + A1/A2/A12; edge rows
// 6-8 and 12). Hide is a RECONCILE: the suite's flagship cases are therefore
// round trips — plan a hide, APPLY it to the view with a request simulator,
// mutate the "document" the way a user would (highlight inside a hidden
// region, narrow the keep set, change the cite threshold), and assert the
// NEXT plan surfaces exactly the right text and nothing else. Every emission
// also runs through one invariant gauntlet (assertEmissionInvariants) so the
// load-bearing rules — request-type whitelist, non-overlapping anchors,
// names-always-decode, anchors-before-style-writes, the segment-final-newline
// clamp — are asserted on EVERY plan, not just the test that aims at them.

import { planHide } from "../google-docs/src/core/planner";
import { decodeRangeName } from "../google-docs/src/core/rangeNames";
import { ANALYTICS_FG_HEX, ANALYTICS_PT, DEFAULT_CITE_MIN_PT, SENTINEL_PT } from "../google-docs/src/core/constants";
import { DEFAULT_KEEP_HEXES } from "../google-docs/src/core/settings";
import {
  DocsRequest,
  GDoc,
  GdocsSettings,
  GElement,
  RequestGroup,
  RleEntry
} from "../google-docs/src/core/types";
import { buildDoc, GeSpec, para, range as namedRange } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** Resolved settings with the engine's real defaults; override per test (the
 * literal keeps this suite independent of settings.ts parsing — the planner
 * only ever sees a RESOLVED GdocsSettings). */
function cfg(over: Partial<GdocsSettings> = {}): GdocsSettings {
  return {
    keepMode: "set",
    keepColors: DEFAULT_KEEP_HEXES,
    citeMinPt: DEFAULT_CITE_MIN_PT,
    structuralCite: true,
    collapseSpacing: false,
    ...over
  };
}

/** Deep-copy a GDoc view (plain data throughout, so spreads suffice). */
function cloneDoc(doc: GDoc): GDoc {
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p) => ({ ...p, elements: p.elements.map((el) => ({ ...el })) })),
    namedRanges: doc.namedRanges.map((nr) => ({ ...nr, segments: nr.segments.map((s) => ({ ...s })) }))
  };
}

/** Slice one text element to [s, e) keeping every other fact. */
function sliceElement(el: GElement, s: number, e: number): GElement {
  return { ...el, startIndex: s, endIndex: e, text: el.text.slice(s - el.startIndex, e - el.startIndex) };
}

/** Rewrite the text elements overlapping [start, end): split at the span
 * edges and map the covered middle piece — the view-level analog of how the
 * Docs API applies a ranged update (shared by the simulator and the
 * highlight-while-hidden mutation, so both split identically). */
function mapSpan(doc: GDoc, start: number, end: number, f: (el: GElement) => GElement): GDoc {
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p) => ({
      ...p,
      elements: p.elements.flatMap((el) => {
        if (el.kind !== "text" || el.endIndex <= start || el.startIndex >= end) return [el];
        const a = Math.max(el.startIndex, start);
        const b = Math.min(el.endIndex, end);
        const pieces: GElement[] = [];
        if (a > el.startIndex) pieces.push(sliceElement(el, el.startIndex, a));
        pieces.push(f(sliceElement(el, a, b)));
        if (b < el.endIndex) pieces.push(sliceElement(el, b, el.endIndex));
        return pieces;
      })
    }))
  };
}

/** The user action the A1 reconcile exists for: highlighting inside a hidden
 * region (background set on a sub-span of sentinel-size text). */
function setBackground(doc: GDoc, start: number, end: number, hex: string): GDoc {
  return mapSpan(doc, start, end, (el) => ({ ...el, backgroundHex: hex }));
}

/** Apply a plan's requests to the view — a faithful-enough simulator for the
 * four request kinds the planner may emit, enabling round-trip tests
 * (hide -> apply -> mutate -> re-hide) without a host. */
function applyPlan(doc: GDoc, groups: RequestGroup[]): GDoc {
  let next = cloneDoc(doc);
  let counter = next.namedRanges.length; // unique ids even on pre-armed docs
  for (const g of groups) {
    for (const req of g.requests) {
      if ("updateTextStyle" in req) {
        const { range, textStyle } = req.updateTextStyle;
        // Empty textStyle with fontSize in the mask = clear to inherit (null).
        const size = textStyle.fontSize?.magnitude ?? null;
        next = mapSpan(next, range.startIndex, range.endIndex, (el) => ({ ...el, fontSizePt: size }));
      } else if ("createNamedRange" in req) {
        next.namedRanges = [
          ...next.namedRanges,
          { id: `nr-${counter++}`, name: req.createNamedRange.name, segments: [{ ...req.createNamedRange.range }] }
        ];
      } else if ("deleteNamedRange" in req) {
        next.namedRanges = next.namedRanges.filter((nr) => nr.id !== req.deleteNamedRange.namedRangeId);
      } else if ("updateParagraphStyle" in req) {
        // Honor the request's actual values: an absent channel with the field
        // named in the mask clears to inherit (null) — exactly the API rule —
        // so the simulator replays collapses (0/0) AND rework restores.
        const { range, paragraphStyle } = req.updateParagraphStyle;
        const above = paragraphStyle.spaceAbove?.magnitude ?? null;
        const below = paragraphStyle.spaceBelow?.magnitude ?? null;
        next = {
          ...next,
          paragraphs: next.paragraphs.map((p) =>
            p.startIndex < range.endIndex && p.endIndex > range.startIndex
              ? { ...p, spaceAbovePt: above, spaceBelowPt: below }
              : p
          )
        };
      }
    }
  }
  return next;
}

/** Every request kind the planner is allowed to emit (case 001-F1: nothing
 * that inserts, deletes, or reorders content — indexes must never move). */
const ALLOWED_KEYS = ["updateTextStyle", "createNamedRange", "deleteNamedRange", "updateParagraphStyle"];

/** The min content index a group touches — the planner's descending-order key. */
function groupKey(g: RequestGroup): number {
  let min = Infinity;
  for (const r of g.requests) {
    if ("updateTextStyle" in r) min = Math.min(min, r.updateTextStyle.range.startIndex);
    else if ("createNamedRange" in r) min = Math.min(min, r.createNamedRange.range.startIndex);
    else if ("updateParagraphStyle" in r) min = Math.min(min, r.updateParagraphStyle.range.startIndex);
  }
  return min;
}

/**
 * The invariant gauntlet, run against EVERY emission in this suite:
 *   * only the whitelisted request kinds; updateParagraphStyle only when the
 *     spacing class is on;
 *   * updateTextStyle is fontSize-only, non-empty, never reaches the
 *     segment-final newline, never overlaps a kind:"other" element (A9);
 *   * every created name decodes; a sizes name's RLE total equals its range
 *     length (restore-exactness precondition);
 *   * created sizes anchors never overlap each other NOR any surviving
 *     pre-existing sizes anchor (the rstm never-overlap invariant);
 *   * within a group every createNamedRange precedes every style write;
 *   * groups are emitted in descending start order (convention — asserted so
 *     it cannot silently drift, though it is documented non-load-bearing).
 */
function assertEmissionInvariants(doc: GDoc, settings: GdocsSettings, groups: RequestGroup[]): void {
  const lastPara = doc.paragraphs.find((p) => p.isLastInSegment);
  const styleCeiling = lastPara !== undefined ? lastPara.endIndex - 1 : Number.MAX_SAFE_INTEGER;
  const otherSpans = doc.paragraphs.flatMap((p) =>
    p.elements.filter((el) => el.kind === "other").map((el) => ({ start: el.startIndex, end: el.endIndex }))
  );

  const deletedIds = new Set<string>();
  const createdSizes: { start: number; end: number }[] = [];

  for (const g of groups) {
    let sawStyleWrite = false;
    // Rework groups (and only rework groups) carry deletes — the marker that
    // distinguishes a spacing RESTORE from the collapse class's zero-write.
    const isReworkGroup = g.requests.some((r) => "deleteNamedRange" in r);
    for (const req of g.requests) {
      const keys = Object.keys(req);
      expect(keys).toHaveLength(1);
      expect(ALLOWED_KEYS).toContain(keys[0]);

      if ("createNamedRange" in req) {
        expect(sawStyleWrite).toBe(false); // anchors precede style writes
        const { name, range } = req.createNamedRange;
        const decoded = decodeRangeName(name);
        if (decoded === null) throw new Error(`emitted name does not decode: ${name}`);
        if (decoded.kind === "sizes") {
          const total = decoded.entries.reduce((n, e) => n + e.count, 0);
          expect(total).toBe(range.endIndex - range.startIndex);
          createdSizes.push({ start: range.startIndex, end: range.endIndex });
        }
      } else if ("updateTextStyle" in req) {
        sawStyleWrite = true;
        const { range, textStyle, fields } = req.updateTextStyle;
        expect(fields).toBe("fontSize");
        expect(range.endIndex).toBeGreaterThan(range.startIndex);
        expect(range.endIndex).toBeLessThanOrEqual(styleCeiling); // newline clamp
        expect(textStyle.bold).toBeUndefined(); // fontSize is the ONLY channel
        if (textStyle.fontSize !== undefined) {
          expect(textStyle.fontSize.unit).toBe("PT");
          expect(textStyle.fontSize.magnitude).toBeGreaterThan(0);
        }
        for (const o of otherSpans) {
          // Whitelist (A9): chips/breaks/objects are never style-targeted.
          expect(range.endIndex <= o.start || range.startIndex >= o.end).toBe(true);
        }
      } else if ("updateParagraphStyle" in req) {
        sawStyleWrite = true;
        const { paragraphStyle, fields } = req.updateParagraphStyle;
        expect(fields).toBe("spaceAbove,spaceBelow");
        if (isReworkGroup) {
          // Rework spacing RESTORE: recorded values come back (absent channel
          // = clear to inherit); any present channel is a real PT value.
          for (const channel of [paragraphStyle.spaceAbove, paragraphStyle.spaceBelow]) {
            if (channel !== undefined) {
              expect(channel.unit).toBe("PT");
              expect(channel.magnitude).toBeGreaterThanOrEqual(0);
            }
          }
        } else {
          // Collapse class: only when the setting is on, and always a full
          // explicit zero on both channels.
          expect(settings.collapseSpacing).toBe(true);
          expect(paragraphStyle.spaceAbove).toEqual({ magnitude: 0, unit: "PT" });
          expect(paragraphStyle.spaceBelow).toEqual({ magnitude: 0, unit: "PT" });
        }
      } else if ("deleteNamedRange" in req) {
        deletedIds.add(req.deleteNamedRange.namedRangeId);
      }
    }
  }

  // rstm sizes anchors never overlap: created vs created AND created vs the
  // pre-existing sizes ranges that survive this plan's deletes.
  const surviving = doc.namedRanges
    .filter((nr) => !deletedIds.has(nr.id) && decodeRangeName(nr.name)?.kind === "sizes")
    .flatMap((nr) => nr.segments.map((s) => ({ start: s.startIndex, end: s.endIndex })));
  const all = [...createdSizes, ...surviving].sort((a, b) => a.start - b.start);
  for (let i = 1; i < all.length; i++) {
    expect(all[i].start).toBeGreaterThanOrEqual(all[i - 1].end);
  }

  // Descending group order.
  const order = groups.map(groupKey).filter((k) => k !== Infinity);
  for (let i = 1; i < order.length; i++) {
    expect(order[i]).toBeLessThan(order[i - 1]);
  }
}

/** Plan + gauntlet in one call — the default way this suite invokes planHide. */
function plan(doc: GDoc, settings: GdocsSettings = cfg()): ReturnType<typeof planHide> {
  const out = planHide(doc, settings);
  assertEmissionInvariants(doc, settings, out.groups);
  return out;
}

/** All requests of a plan, flattened, filtered to one kind via an "in" key
 * (DocsRequest is a union, so keyof would be never — a literal union does). */
function requestsOf(
  groups: RequestGroup[],
  key: "updateTextStyle" | "createNamedRange" | "deleteNamedRange" | "updateParagraphStyle"
): DocsRequest[] {
  return groups.flatMap((g) => g.requests.filter((r) => key in r));
}

// ---------------------------------------------------------------------------
// The shared mixed doc (heading / cite / highlight / body) and its geometry.
// Indexes (body starts at 1; every paragraph ends in its newline):
//   P0 [1,10)   HEADING_4 "Tag line\n"                      kept (heading)
//   P1 [10,30)  bold-14 "Smith 24" + " says stuff\n"        kept (signature cite)
//   P2 [30,50)  "The quick " + yellow "brown" + " fox\n"    mixed
//   P3 [50,66)  "All hidden here\n"                          hidden, LAST (clamp)
// Expected fresh hide: region [30,40) and region [45,65) (the " fox\n" tail
// runs into P3 across the paragraph boundary; P3's final newline is clamped).
// ---------------------------------------------------------------------------

function mixedDoc(): GDoc {
  return buildDoc([
    para("Tag line", "HEADING_4"),
    { elements: [{ text: "Smith 24", bold: true, size: 14 }, { text: " says stuff" }] },
    { elements: [{ text: "The quick " }, { text: "brown", bg: "#ffff00" }, { text: " fox" }] },
    para("All hidden here")
  ]);
}

// ---------------------------------------------------------------------------
// 1. Fresh hide
// ---------------------------------------------------------------------------

describe("planHide — fresh hide", () => {
  it("hides the mixed doc as two regions with correct RLE, newline inclusion and final clamp", () => {
    const { groups, result } = plan(mixedDoc());

    // Derived expectation: 2 regions x (1 anchor + 1 shrink) = 4 requests.
    expect(groups).toHaveLength(2);
    expect(groups.flatMap((g) => g.requests)).toHaveLength(4);

    // Descending: the cross-paragraph tail region [45,65) precedes [30,40).
    // Its RLE covers " fox\n" (5) + "All hidden here" (15, final \n clamped).
    expect(groups[0].requests).toEqual([
      { createNamedRange: { name: "rstm:v1:20xi", range: { startIndex: 45, endIndex: 65 } } },
      {
        updateTextStyle: {
          range: { startIndex: 45, endIndex: 65 },
          textStyle: { fontSize: { magnitude: SENTINEL_PT, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    expect(groups[1].requests).toEqual([
      { createNamedRange: { name: "rstm:v1:10xi", range: { startIndex: 30, endIndex: 40 } } },
      {
        updateTextStyle: {
          range: { startIndex: 30, endIndex: 40 },
          textStyle: { fontSize: { magnitude: SENTINEL_PT, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);

    expect(result).toEqual({
      paragraphsScanned: 4,
      paragraphsChanged: 2, // P2 and P3 — kept P0/P1 untouched
      regionsHidden: 2,
      regionsAlreadyHidden: 0,
      newlyKeptRestored: 0,
      preexistingTinyCount: 0
    });
  });

  it("canary: exact emission for a one-paragraph doc (whole-para hide, last-paragraph clamp)", () => {
    // "xyz\n" occupies [1,5); the segment-final newline at 4 is unstylable,
    // so both the anchor and the shrink stop at 4 and the RLE counts 3 chars.
    const { groups, result } = plan(buildDoc([para("xyz")]));
    expect(groups).toEqual([
      {
        requests: [
          { createNamedRange: { name: "rstm:v1:3xi", range: { startIndex: 1, endIndex: 4 } } },
          {
            updateTextStyle: {
              range: { startIndex: 1, endIndex: 4 },
              textStyle: { fontSize: { magnitude: 1, unit: "PT" } },
              fields: "fontSize"
            }
          }
        ]
      }
    ]);
    expect(result).toEqual({
      paragraphsScanned: 1,
      paragraphsChanged: 1,
      regionsHidden: 1,
      regionsAlreadyHidden: 0,
      newlyKeptRestored: 0,
      preexistingTinyCount: 0
    });
  });

  it("a mid-doc paragraph's trailing newline IS included in its region (D6)", () => {
    // Alternating 8/11pt body -> a SINGLE region with a multi-entry RLE; the
    // newline rides in the final 8pt element, so the last entry counts it.
    const doc = buildDoc([
      { elements: [{ text: "ab", size: 8 }, { text: "cd", size: 11 }, { text: "ef", size: 8 }] },
      para("Hat", "HEADING_1")
    ]);
    const { groups, result } = plan(doc);
    expect(groups).toHaveLength(1);
    expect(groups[0].requests).toEqual([
      { createNamedRange: { name: "rstm:v1:2x8,2x11,3x8", range: { startIndex: 1, endIndex: 8 } } },
      {
        updateTextStyle: {
          range: { startIndex: 1, endIndex: 8 },
          textStyle: { fontSize: { magnitude: 1, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    expect(result.regionsHidden).toBe(1);
    expect(result.paragraphsChanged).toBe(1);
  });

  it("indexes are UTF-16 code units: a surrogate pair counts 2 in the RLE (row 14)", () => {
    // "🙂ab\n" = 2 + 2 + 1 code units at [1,6); clamp drops the final newline.
    const { groups } = plan(buildDoc([para("🙂ab")]));
    expect(requestsOf(groups, "createNamedRange")).toEqual([
      { createNamedRange: { name: "rstm:v1:4xi", range: { startIndex: 1, endIndex: 5 } } }
    ]);
  });

  it("splits an overflowing RLE name at entry boundaries (NAME_MAX) with exact piece ranges", () => {
    // 60 alternating-size elements -> 60 non-coalescible entries whose name
    // would far exceed NAME_MAX, so the region carries several anchors but
    // still exactly ONE shrink.
    const els: GeSpec[] = [];
    for (let i = 0; i < 60; i++) els.push({ text: "ab", size: i % 2 === 0 ? 7 : 9 });
    const doc = buildDoc([{ elements: els }, para("Hat", "HEADING_1")]);
    const { groups, result } = plan(doc);
    expect(result.regionsHidden).toBe(1);
    expect(groups).toHaveLength(1);

    const creates = requestsOf(groups, "createNamedRange");
    expect(creates.length).toBeGreaterThanOrEqual(2);
    expect(requestsOf(groups, "updateTextStyle")).toHaveLength(1);

    // The pieces partition the region contiguously and their decoded entries
    // recombine into EXACTLY the original size sequence (restore stays exact
    // across the split — case 001-S4).
    const regionEnd = 1 + 60 * 2 + 1; // 120 chars + the paragraph newline
    let cursor = 1;
    const recombined: RleEntry[] = [];
    for (const c of creates) {
      if (!("createNamedRange" in c)) throw new Error("filtered wrong kind");
      expect(c.createNamedRange.range.startIndex).toBe(cursor);
      cursor = c.createNamedRange.range.endIndex;
      const decoded = decodeRangeName(c.createNamedRange.name);
      if (decoded === null || decoded.kind !== "sizes") throw new Error("piece must decode as sizes");
      recombined.push(...decoded.entries);
    }
    expect(cursor).toBe(regionEnd);
    const expected: RleEntry[] = els.map((e, i) => ({ count: i === 59 ? 3 : 2, sizePt: e.size ?? null }));
    expect(recombined).toEqual(expected);
  });

  it("a chip (kind 'other') splits the region and is never touched (A9, row 2)", () => {
    const doc = buildDoc([
      { elements: [{ text: "aa" }, { kind: "other", text: "@@@@@" }, { text: "bb" }] },
      para("Hat", "HEADING_1")
    ]);
    const { groups, result } = plan(doc);
    expect(result.regionsHidden).toBe(2); // [1,3) and [8,11) around the chip [3,8)
    const creates = requestsOf(groups, "createNamedRange");
    for (const c of creates) {
      if (!("createNamedRange" in c)) throw new Error("filtered wrong kind");
      const r = c.createNamedRange.range;
      expect(r.endIndex <= 3 || r.startIndex >= 8).toBe(true); // anchors avoid the chip too
    }
  });

  it("tables and headings emit nothing (rows 1+3) — the all-kept no-op", () => {
    const doc = buildDoc([
      para("Pocket", "HEADING_1"),
      { inTable: true, elements: [{ text: "cell prose" }] },
      para("Tag", "HEADING_4")
    ]);
    const { groups, result } = plan(doc);
    expect(groups).toEqual([]);
    expect(result).toEqual({
      paragraphsScanned: 3,
      paragraphsChanged: 0,
      regionsHidden: 0,
      regionsAlreadyHidden: 0,
      newlyKeptRestored: 0,
      preexistingTinyCount: 0
    });
  });

  it("an empty doc and a lone empty paragraph both no-op (row 12)", () => {
    expect(plan(buildDoc([])).groups).toEqual([]);
    // A lone empty paragraph is just the segment-final newline — unstylable,
    // so there is nothing to hide.
    const { groups, result } = plan(buildDoc([{ elements: [{ text: "" }] }]));
    expect(groups).toEqual([]);
    expect(result.paragraphsScanned).toBe(1);
    expect(result.paragraphsChanged).toBe(0);
  });

  it("a foreign add-on's named range is invisible: its text hides like any other (ownership)", () => {
    // Another add-on's range over P3 must not read as rstm coverage — the
    // text under it is fresh hide territory and the range itself untouched.
    const doc = buildDoc(
      [para("Tag line", "HEADING_4"), para("All hidden here")],
      { namedRanges: [namedRange("f1", "docs-internal-abc123", 10, 26)] }
    );
    const { groups, result } = plan(doc);
    expect(requestsOf(groups, "deleteNamedRange")).toEqual([]);
    expect(requestsOf(groups, "createNamedRange")).toEqual([
      { createNamedRange: { name: "rstm:v1:15xi", range: { startIndex: 10, endIndex: 25 } } }
    ]);
    expect(result.regionsHidden).toBe(1);
    expect(result.regionsAlreadyHidden).toBe(0); // foreign coverage is not hidden state
  });

  it("pre-existing sentinel-size text with no record is counted and NEVER touched (rows 8/12)", () => {
    const doc = buildDoc([
      // TWO adjacent tiny elements — one contiguous passage, counted once.
      { elements: [{ text: "normal " }, { text: "ti", size: 1 }, { text: "ny", size: 1 }, { text: " more" }] },
      para("Hat", "HEADING_1")
    ]);
    const { groups, result } = plan(doc);
    // The tiny span [8,12) breaks the region in two and appears in no request.
    expect(result.regionsHidden).toBe(2);
    expect(result.preexistingTinyCount).toBe(1);
    for (const g of groups) {
      for (const r of g.requests) {
        if ("updateTextStyle" in r) {
          const rr = r.updateTextStyle.range;
          expect(rr.endIndex <= 8 || rr.startIndex >= 12).toBe(true);
        }
        if ("createNamedRange" in r) {
          const rr = r.createNamedRange.range;
          expect(rr.endIndex <= 8 || rr.startIndex >= 12).toBe(true);
        }
      }
    }
  });
});

// ---------------------------------------------------------------------------
// 2. Reconcile round trips (plan A1 — the BLOCKER fix)
// ---------------------------------------------------------------------------

describe("planHide — reconcile on an armed doc", () => {
  it("re-hide is idempotent: zero groups, already-hidden passages counted (row 7)", () => {
    const first = plan(mixedDoc());
    const armed = applyPlan(mixedDoc(), first.groups);

    const second = plan(armed);
    expect(second.groups).toEqual([]);
    expect(second.result).toEqual({
      paragraphsScanned: 4,
      paragraphsChanged: 0,
      regionsHidden: 0,
      regionsAlreadyHidden: 2, // [30,40) and [45,65) — separated by kept text
      newlyKeptRestored: 0,
      preexistingTinyCount: 0
    });
  });

  it("THE A1 case: a highlight added inside a hidden region surfaces via rework", () => {
    const armed = applyPlan(mixedDoc(), plan(mixedDoc()).groups);
    // User highlights "hidden" ([54,60)) inside the hidden tail region [45,65).
    const marked = setBackground(armed, 54, 60, "#00ffff");

    const { groups, result } = plan(marked);
    expect(groups).toHaveLength(1);
    // ONE atomic group: drop the stale anchor, re-anchor the two still-hidden
    // flanks with recomputed RLE, restore the kept sub-span. The record was
    // "20xi" (inherit), so the restore is a CLEAR — empty textStyle.
    expect(groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "nr-0" } },
      { createNamedRange: { name: "rstm:v1:9xi", range: { startIndex: 45, endIndex: 54 } } },
      { createNamedRange: { name: "rstm:v1:5xi", range: { startIndex: 60, endIndex: 65 } } },
      {
        updateTextStyle: {
          range: { startIndex: 54, endIndex: 60 },
          textStyle: {},
          fields: "fontSize"
        }
      }
    ]);
    expect(result).toEqual({
      paragraphsScanned: 4,
      paragraphsChanged: 0, // nothing newly hidden
      regionsHidden: 0,
      regionsAlreadyHidden: 1, // the untouched [30,40) region
      newlyKeptRestored: 1,
      preexistingTinyCount: 0
    });

    // And the reconcile CONVERGES: applying the rework and re-hiding again
    // emits nothing (the highlight stays surfaced, the flanks stay hidden).
    // Passages: [30,40) plus the two flanks around the surfaced highlight.
    const settled = plan(applyPlan(marked, groups));
    expect(settled.groups).toEqual([]);
    expect(settled.result.regionsAlreadyHidden).toBe(3);
  });

  it("restores an EXPLICIT recorded size (not inherit) when the RLE says so", () => {
    // Hand-built armed doc: original 8pt text hidden under an intact "9x8"
    // record; the user then highlighted "de" ([13,15)).
    const doc = buildDoc(
      [
        para("Tag line", "HEADING_4"),
        {
          elements: [
            { text: "abc", size: 1 },
            { text: "de", size: 1, bg: "#ffff00" },
            { text: "fgh", size: 1 }
          ]
        },
        para("Hat", "HEADING_1")
      ],
      { namedRanges: [namedRange("x1", "rstm:v1:9x8", 10, 19)] }
    );
    const { groups, result } = plan(doc);
    expect(groups).toHaveLength(1);
    expect(groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "x1" } },
      { createNamedRange: { name: "rstm:v1:3x8", range: { startIndex: 10, endIndex: 13 } } },
      { createNamedRange: { name: "rstm:v1:4x8", range: { startIndex: 15, endIndex: 19 } } },
      {
        updateTextStyle: {
          range: { startIndex: 13, endIndex: 15 },
          textStyle: { fontSize: { magnitude: 8, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    expect(result.newlyKeptRestored).toBe(1);
  });

  it("splits ONE element at an interior RLE entry boundary (sub-element restore spans)", () => {
    // A single sentinel element [10,19) covered by a two-entry record
    // "4x8,5x9"; the keeper keeps the WHOLE paragraph (heading style was
    // applied while hidden), so both entries restore — at their own sizes.
    const doc = buildDoc(
      [
        para("Tag line", "HEADING_4"),
        { style: "HEADING_2", elements: [{ text: "abcdefgh", size: 1 }] },
        para("Hat", "HEADING_1")
      ],
      { namedRanges: [namedRange("x1", "rstm:v1:4x8,5x9", 10, 19)] }
    );
    const { groups, result } = plan(doc);
    expect(groups).toHaveLength(1);
    expect(groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "x1" } },
      {
        updateTextStyle: {
          range: { startIndex: 10, endIndex: 14 },
          textStyle: { fontSize: { magnitude: 8, unit: "PT" } },
          fields: "fontSize"
        }
      },
      {
        updateTextStyle: {
          range: { startIndex: 14, endIndex: 19 },
          textStyle: { fontSize: { magnitude: 9, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    // TWO writes (one per recorded size) but ONE passage: the receipt counts
    // what the user sees come back, and the two adjacent sub-spans read as a
    // single surfaced passage (visual unit, not write unit).
    expect(result.newlyKeptRestored).toBe(1);
    expect(result.regionsAlreadyHidden).toBe(0);
  });

  it("merges adjacent same-size kept sub-spans into ONE restore write (and counts them as one)", () => {
    // The highlight crosses an element boundary inside the hidden region:
    // two adjacent kept atoms restore to the same recorded 8pt, so they merge
    // into a single updateTextStyle and a single newly-kept passage.
    const doc = buildDoc(
      [
        para("Tag line", "HEADING_4"),
        {
          elements: [
            { text: "ab", size: 1 },
            { text: "cd", size: 1, bg: "#ffff00" },
            { text: "ef", size: 1, bg: "#ffff00" },
            { text: "gh", size: 1 }
          ]
        },
        para("Hat", "HEADING_1")
      ],
      { namedRanges: [namedRange("x1", "rstm:v1:9x8", 10, 19)] }
    );
    const { groups, result } = plan(doc);
    expect(groups).toHaveLength(1);
    expect(groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "x1" } },
      { createNamedRange: { name: "rstm:v1:2x8", range: { startIndex: 10, endIndex: 12 } } },
      { createNamedRange: { name: "rstm:v1:3x8", range: { startIndex: 16, endIndex: 19 } } },
      {
        updateTextStyle: {
          range: { startIndex: 12, endIndex: 16 }, // ONE write across both elements
          textStyle: { fontSize: { magnitude: 8, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    expect(result.newlyKeptRestored).toBe(1);
  });

  it("narrowing the keep set hides the formerly-kept color WITHOUT merging regions (A1)", () => {
    const doc = buildDoc([
      para("Tag line", "HEADING_4"),
      {
        elements: [
          { text: "aa " },
          { text: "yellow", bg: "#ffff00" },
          { text: " bb " },
          { text: "cyan", bg: "#00ffff" },
          { text: " cc" }
        ]
      }
    ]);
    const armed = applyPlan(doc, plan(doc).groups); // both colors kept: 3 regions

    const narrowed = cfg({ keepColors: new Set(["#ffff00"]) });
    const second = planHide(armed, narrowed);
    assertEmissionInvariants(armed, narrowed, second.groups);

    // Exactly the cyan span [23,27) is newly hidden, as its OWN region —
    // adjacent to the existing anchors at [19,23) and [27,30), never merged.
    expect(second.groups).toHaveLength(1);
    expect(second.groups[0].requests).toEqual([
      { createNamedRange: { name: "rstm:v1:4xi", range: { startIndex: 23, endIndex: 27 } } },
      {
        updateTextStyle: {
          range: { startIndex: 23, endIndex: 27 },
          textStyle: { fontSize: { magnitude: 1, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    expect(second.result.regionsHidden).toBe(1);
    expect(second.result.paragraphsChanged).toBe(1);
    expect(second.result.regionsAlreadyHidden).toBe(3); // untouched anchors stay distinct
    expect(second.result.newlyKeptRestored).toBe(0);
  });

  it("lowering citeMinPt surfaces a hidden cite — the cite-size trap the restored view fixes", () => {
    // Docs-shaped trap: bold 13pt cite hidden under citeMinPt 14. Against
    // CURRENT sizes the cite is 1pt and would stay buried forever; the
    // restored view re-evaluates it at its recorded 13pt.
    const doc = buildDoc([
      para("Tag line", "HEADING_4"),
      { elements: [{ text: "Smith argues stuff", bold: true, size: 13 }] },
      para("body text")
    ]);
    const strict = cfg({ citeMinPt: 14, structuralCite: false });
    const first = planHide(doc, strict);
    assertEmissionInvariants(doc, strict, first.groups);
    // One region spans both paragraphs: "19x13" (cite + its newline) then 9
    // inherit chars of body (final newline clamped).
    expect(requestsOf(first.groups, "createNamedRange")).toEqual([
      { createNamedRange: { name: "rstm:v1:19x13,9xi", range: { startIndex: 10, endIndex: 38 } } }
    ]);

    const armed = applyPlan(doc, first.groups);
    const relaxed = cfg({ citeMinPt: 13, structuralCite: false });
    const second = planHide(armed, relaxed);
    assertEmissionInvariants(armed, relaxed, second.groups);

    expect(second.groups).toHaveLength(1);
    expect(second.groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "nr-0" } },
      { createNamedRange: { name: "rstm:v1:9xi", range: { startIndex: 29, endIndex: 38 } } },
      {
        updateTextStyle: {
          range: { startIndex: 10, endIndex: 29 },
          textStyle: { fontSize: { magnitude: 13, unit: "PT" } },
          fields: "fontSize"
        }
      }
    ]);
    expect(second.result.newlyKeptRestored).toBe(1);
    expect(second.result.regionsHidden).toBe(0);
  });

  it("a mismatched RLE record reconciles as inherit: untouched when still hidden, CLEARED when kept", () => {
    // The record claims 5 chars but the segment holds 9 (interior edit).
    const armedSpec = (withHighlight: boolean) =>
      buildDoc(
        [
          para("Tag line", "HEADING_4"),
          {
            elements: withHighlight
              ? [
                  { text: "ab", size: 1 },
                  { text: "cd", size: 1, bg: "#ffff00" },
                  { text: "efgh", size: 1 }
                ]
              : [{ text: "abcdefgh", size: 1 }]
          },
          para("Hat", "HEADING_1")
        ],
        { namedRanges: [namedRange("x1", "rstm:v1:5x8", 10, 19)] }
      );

    // Still hidden -> nothing to do; the decayed record is Show All's amber
    // problem, not Hide's.
    const quiet = plan(armedSpec(false));
    expect(quiet.groups).toEqual([]);
    expect(quiet.result.regionsAlreadyHidden).toBe(1);

    // A keeper inside it -> rework; the restore must be a CLEAR (inherit),
    // never the no-longer-trustworthy 8pt, and the remainder re-anchors as
    // inherit too.
    const reworked = plan(armedSpec(true));
    expect(reworked.groups).toHaveLength(1);
    expect(reworked.groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "x1" } },
      { createNamedRange: { name: "rstm:v1:2xi", range: { startIndex: 10, endIndex: 12 } } },
      { createNamedRange: { name: "rstm:v1:5xi", range: { startIndex: 14, endIndex: 19 } } },
      {
        updateTextStyle: {
          range: { startIndex: 12, endIndex: 14 },
          textStyle: {},
          fields: "fontSize"
        }
      }
    ]);
    expect(reworked.result.newlyKeptRestored).toBe(1);
  });

  it("foreign-size text pasted into a hidden region is untouched (A6, hide side)", () => {
    // "xy" at 14pt sits inside an intact inherit record covering [10,17):
    // not hidden (not sentinel), not a keeper — and absolutely not ours to
    // shrink or restore. The region reads as one already-hidden passage.
    const doc = buildDoc(
      [
        para("Tag line", "HEADING_4"),
        {
          elements: [
            { text: "ab", size: 1 },
            { text: "xy", size: 14 },
            { text: "cd", size: 1 }
          ]
        },
        para("Hat", "HEADING_1")
      ],
      { namedRanges: [namedRange("x1", "rstm:v1:7xi", 10, 17)] }
    );
    const { groups, result } = plan(doc);
    expect(groups).toEqual([]);
    expect(result.regionsAlreadyHidden).toBe(1);
    expect(result.preexistingTinyCount).toBe(0);
  });

  it("an unknown rstm version is left untouched but still reads as hidden state (row 16)", () => {
    const doc = buildDoc(
      [para("Tag line", "HEADING_4"), { elements: [{ text: "abcdef", size: 1 }] }, para("Hat", "HEADING_1")],
      { namedRanges: [namedRange("u1", "rstm:v2:future", 10, 17)] }
    );
    const { groups, result } = plan(doc);
    expect(groups).toEqual([]); // never rewrite a future version's state
    expect(result.regionsAlreadyHidden).toBe(1); // it IS hidden, truthfully
    expect(result.preexistingTinyCount).toBe(0); // and NOT orphan tiny text
  });
});

// ---------------------------------------------------------------------------
// 3. Spacing-collapse class (plan A12, default OFF)
// ---------------------------------------------------------------------------

describe("planHide — collapseSpacing", () => {
  /** P1 carries direct 12/6 spacing and hides whole; P0/P2 pin the window. */
  function spacedDoc(extra: Partial<{ namedRanges: GDoc["namedRanges"] }> = {}): GDoc {
    return buildDoc(
      [
        para("Tag line", "HEADING_4"),
        { elements: [{ text: "hide me" }], spaceAbovePt: 12, spaceBelowPt: 6 },
        para("Hat", "HEADING_1")
      ],
      extra
    );
  }

  it("a fully-hidden paragraph's spacing pair rides in the SAME group as its region", () => {
    const settings = cfg({ collapseSpacing: true });
    const { groups, result } = planHide(spacedDoc(), settings);
    assertEmissionInvariants(spacedDoc(), settings, groups);

    expect(groups).toHaveLength(1);
    expect(groups[0].requests).toEqual([
      { createNamedRange: { name: "rstm:v1:8xi", range: { startIndex: 10, endIndex: 18 } } },
      { createNamedRange: { name: "rstm:v1:p:12x6", range: { startIndex: 10, endIndex: 18 } } },
      {
        updateTextStyle: {
          range: { startIndex: 10, endIndex: 18 },
          textStyle: { fontSize: { magnitude: 1, unit: "PT" } },
          fields: "fontSize"
        }
      },
      {
        updateParagraphStyle: {
          range: { startIndex: 10, endIndex: 18 },
          paragraphStyle: { spaceAbove: { magnitude: 0, unit: "PT" }, spaceBelow: { magnitude: 0, unit: "PT" } },
          fields: "spaceAbove,spaceBelow"
        }
      }
    ]);
    expect(result.regionsHidden).toBe(1);
  });

  it("OFF by default: the same doc emits no paragraph-style requests at all", () => {
    const { groups } = plan(spacedDoc()); // cfg() has collapseSpacing false
    expect(requestsOf(groups, "updateParagraphStyle")).toEqual([]);
    expect(groups[0].requests).toHaveLength(2); // anchor + shrink only
  });

  it("a PARTIALLY hidden paragraph never collapses its spacing", () => {
    const doc = buildDoc([
      {
        elements: [{ text: "keep", bg: "#ffff00" }, { text: " hide" }],
        spaceAbovePt: 12,
        spaceBelowPt: 6
      },
      para("Hat", "HEADING_1")
    ]);
    const settings = cfg({ collapseSpacing: true });
    const { groups } = planHide(doc, settings);
    assertEmissionInvariants(doc, settings, groups);
    expect(requestsOf(groups, "updateParagraphStyle")).toEqual([]);
    expect(groups[0].requests).toHaveLength(2);
  });

  it("inherited (null) and already-zero spacing have nothing to collapse", () => {
    const doc = buildDoc([
      { elements: [{ text: "inherits" }] },
      { elements: [{ text: "zeroed" }], spaceAbovePt: 0, spaceBelowPt: 0 },
      para("Hat", "HEADING_1")
    ]);
    const settings = cfg({ collapseSpacing: true });
    const { groups } = planHide(doc, settings);
    assertEmissionInvariants(doc, settings, groups);
    expect(requestsOf(groups, "updateParagraphStyle")).toEqual([]);
  });

  it("a paragraph already carrying a spacing record is never double-recorded", () => {
    const doc = spacedDoc({ namedRanges: [namedRange("s1", "rstm:v1:p:5x5", 10, 18)] });
    const settings = cfg({ collapseSpacing: true });
    const { groups } = planHide(doc, settings);
    assertEmissionInvariants(doc, settings, groups);
    expect(requestsOf(groups, "updateParagraphStyle")).toEqual([]);
    expect(groups[0].requests).toHaveLength(2); // sizes anchor + shrink only
  });

  it("clamps the paragraph-style range at the segment-final newline; the anchor keeps the full paragraph", () => {
    const doc = buildDoc([
      para("Tag line", "HEADING_4"),
      { elements: [{ text: "hide me" }], spaceAbovePt: 12, spaceBelowPt: null } // LAST paragraph
    ]);
    const settings = cfg({ collapseSpacing: true });
    const { groups } = planHide(doc, settings);
    assertEmissionInvariants(doc, settings, groups);
    expect(groups[0].requests).toEqual([
      { createNamedRange: { name: "rstm:v1:7xi", range: { startIndex: 10, endIndex: 17 } } },
      // The spacing record covers the whole PARAGRAPH (the unit restore will
      // re-space); a null side encodes "i" and restores by clearing.
      { createNamedRange: { name: "rstm:v1:p:12xi", range: { startIndex: 10, endIndex: 18 } } },
      {
        updateTextStyle: {
          range: { startIndex: 10, endIndex: 17 },
          textStyle: { fontSize: { magnitude: 1, unit: "PT" } },
          fields: "fontSize"
        }
      },
      {
        updateParagraphStyle: {
          range: { startIndex: 10, endIndex: 17 }, // clamped — never style the final newline
          paragraphStyle: { spaceAbove: { magnitude: 0, unit: "PT" }, spaceBelow: { magnitude: 0, unit: "PT" } },
          fields: "spaceAbove,spaceBelow"
        }
      }
    ]);
  });

  it("round trip: an applied spacing collapse is idempotent on re-hide", () => {
    const settings = cfg({ collapseSpacing: true });
    const first = planHide(spacedDoc(), settings);
    const armed = applyPlan(spacedDoc(), first.groups);
    const second = planHide(armed, settings);
    assertEmissionInvariants(armed, settings, second.groups);
    expect(second.groups).toEqual([]);
    expect(second.result.regionsAlreadyHidden).toBe(1);
  });

  it("rework restores a surfacing paragraph's RECORDED spacing in the same group; still-hidden paragraphs keep zero", () => {
    // Two spaced paragraphs hide as one region with per-paragraph spacing
    // records; the user then highlights inside P1. The reconcile must bring
    // P1's text AND its 12/6 spacing back atomically (a torn pair would leave
    // visible text squashed, or a live record fighting a later Show All),
    // while P2 — still fully hidden — keeps its zeroed spacing and record.
    const settings = cfg({ collapseSpacing: true });
    // "Tag line\n" [1,10) H4 | "hide me one\n" [10,22) 12/6 | "hide me two\n"
    // [22,34) 12/6 | "Hat\n" [34,38) H1.
    const doc = buildDoc([
      para("Tag line", "HEADING_4"),
      { elements: [{ text: "hide me one" }], spaceAbovePt: 12, spaceBelowPt: 6 },
      { elements: [{ text: "hide me two" }], spaceAbovePt: 12, spaceBelowPt: 6 },
      para("Hat", "HEADING_1")
    ]);
    const first = planHide(doc, settings);
    assertEmissionInvariants(doc, settings, first.groups);
    const armed = applyPlan(doc, first.groups);
    // Sanity: one region [10,34) + a spacing record per collapsed paragraph.
    expect(armed.namedRanges.map((nr) => nr.name).sort()).toEqual([
      "rstm:v1:24xi",
      "rstm:v1:p:12x6",
      "rstm:v1:p:12x6"
    ]);

    // The user highlights [13,18) inside hidden P1, then re-hides.
    const marked = setBackground(armed, 13, 18, "#ffff00");
    const second = planHide(marked, settings);
    assertEmissionInvariants(marked, settings, second.groups);

    expect(second.groups).toHaveLength(1);
    expect(second.groups[0].requests).toEqual([
      { deleteNamedRange: { namedRangeId: "nr-0" } }, // the stale sizes anchor
      { deleteNamedRange: { namedRangeId: "nr-1" } }, // P1's spacing record dies with it
      { createNamedRange: { name: "rstm:v1:3xi", range: { startIndex: 10, endIndex: 13 } } },
      { createNamedRange: { name: "rstm:v1:16xi", range: { startIndex: 18, endIndex: 34 } } },
      { updateTextStyle: { range: { startIndex: 13, endIndex: 18 }, textStyle: {}, fields: "fontSize" } },
      {
        updateParagraphStyle: {
          range: { startIndex: 10, endIndex: 22 },
          paragraphStyle: {
            spaceAbove: { magnitude: 12, unit: "PT" },
            spaceBelow: { magnitude: 6, unit: "PT" }
          },
          fields: "spaceAbove,spaceBelow"
        }
      }
    ]);
    expect(second.result.newlyKeptRestored).toBe(1);

    // Apply the rework: P1 is re-spaced, P2 stays zeroed with its record.
    const settled = applyPlan(marked, second.groups);
    expect([settled.paragraphs[1].spaceAbovePt, settled.paragraphs[1].spaceBelowPt]).toEqual([12, 6]);
    expect([settled.paragraphs[2].spaceAbovePt, settled.paragraphs[2].spaceBelowPt]).toEqual([0, 0]);
    expect(settled.namedRanges.filter((nr) => nr.name.startsWith("rstm:v1:p:"))).toHaveLength(1);

    // And the reconcile converges: a third pass plans nothing.
    const third = planHide(settled, settings);
    assertEmissionInvariants(settled, settings, third.groups);
    expect(third.groups).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 4. Analytics foreground carry (Loop 003, case 003-S2 regression)
//
// The BLOCKER this guards: restoredView() rebuilds each element from its atom,
// and Hide's keeper runs against THAT view. If the rebuild drops foregroundHex
// the analytics keeper (keepers.isAnalytics) sees `undefined`, so a navy-14pt
// analytics run — which must be KEPT, always-on (goal) — would be shrunk to the
// sentinel on a FRESH Hide, and would fail to break a hidden region it abuts.
// These cases assert the keeper VERDICT through the planner's emission: no
// sentinel shrink ever covers an analytics span. Color is body-sized on
// purpose (detection is color-only, size-independent), so the run reads as
// ordinary body to every size-based rule and ONLY the foreground carry saves it.
// ---------------------------------------------------------------------------

describe("planHide — analytics foreground carry (003-S2)", () => {
  /** Every sentinel-shrink range the plan emits — the spans Hide would hide.
   * (A shrink is the fontSize→SENTINEL_PT updateTextStyle; restore/clear writes
   * in this suite's fresh-hide cases never appear, but filtering on the
   * sentinel magnitude keeps the helper honest if one ever did.) */
  function shrinkRanges(groups: RequestGroup[]): { startIndex: number; endIndex: number }[] {
    return requestsOf(groups, "updateTextStyle")
      .filter((r) => "updateTextStyle" in r && r.updateTextStyle.textStyle.fontSize?.magnitude === SENTINEL_PT)
      .map((r) => ("updateTextStyle" in r ? r.updateTextStyle.range : { startIndex: 0, endIndex: 0 }));
  }

  /** Assert no emitted shrink overlaps [start, end) — the analytics span. */
  function expectNoShrinkOver(groups: RequestGroup[], start: number, end: number): void {
    for (const rng of shrinkRanges(groups)) {
      expect(rng.endIndex <= start || rng.startIndex >= end).toBe(true);
    }
  }

  it("a body-sized analytics run is KEPT on a FRESH Hide (no shrink over it)", () => {
    // P0 [1,16) is one analytics run at the BODY size (14 == ANALYTICS_PT, but
    // even were it 11 the rule is color-only) — it would hide like any body
    // text if the keeper could not see its navy foreground on the restored
    // view. The trailing heading pins a kept neighbor so the doc is not a
    // degenerate all-kept no-op for the wrong reason.
    const doc = buildDoc([
      { elements: [{ text: "analytics text", fg: ANALYTICS_FG_HEX, size: ANALYTICS_PT }] },
      para("Hat", "HEADING_1")
    ]);
    const { groups, result } = plan(doc);

    // Nothing hides: the analytics run is kept, the heading is kept.
    expect(groups).toEqual([]);
    expect(result.regionsHidden).toBe(0);
    expect(result.paragraphsChanged).toBe(0);
    // Belt and suspenders: even if some unrelated region appeared, none of it
    // may cover the analytics span [1,16).
    expectNoShrinkOver(groups, 1, 16);
  });

  it("an analytics run BREAKS an adjacent hidden region (flanks hide, run untouched)", () => {
    // One paragraph: hidden body "aaaa" [1,5), analytics "MID" [5,8) (navy, but
    // body-sized so size rules treat it as ordinary), hidden body "bbbb\n"
    // [8,13); trailing heading [13,17) keeps P0 from being the final segment so
    // its own newline is an ordinary hidden char, not a clamp artifact.
    const doc = buildDoc([
      {
        elements: [
          { text: "aaaa" },
          { text: "MID", fg: ANALYTICS_FG_HEX, size: ANALYTICS_PT },
          { text: "bbbb" }
        ]
      },
      para("Hat", "HEADING_1")
    ]);
    const { groups, result } = plan(doc);

    // The kept analytics run splits the body into TWO hidden regions.
    expect(result.regionsHidden).toBe(2);
    expect(requestsOf(groups, "updateTextStyle")).toHaveLength(2);
    expect(shrinkRanges(groups).sort((a, b) => a.startIndex - b.startIndex)).toEqual([
      { startIndex: 1, endIndex: 5 },
      { startIndex: 8, endIndex: 13 }
    ]);
    // The analytics span [5,8) is never style-targeted.
    expectNoShrinkOver(groups, 5, 8);
  });
});
