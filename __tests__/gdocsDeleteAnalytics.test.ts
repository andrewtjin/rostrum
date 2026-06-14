// Delete-analytics planner suite (Loop 003, plan §5) — the engine's SOLE
// content-deleter (003-F1) and a plan-review correctness BLOCKER zone. The
// planner classifies each paragraph (wholly-analytics vs partial vs table),
// clamps each delete range off the unremovable segment-final newline, coalesces
// only ranges STILL touching after that clamp, and emits one deleteContentRange
// per range in strictly DESCENDING start order — load-bearing, because a Docs
// batchUpdate applies sequentially and every delete shifts downstream indexes.
//
// The suite proves two layers:
//   * PURE PLANNER (planDeleteAnalytics over buildDoc views): the classification
//     matrix, the clamp-then-coalesce ORDER, descending emission, the
//     emit-only-deleteContentRange audit, and that non-analytics is never in any
//     range. buildDoc is used here because it (unlike FakeDocs) can model the
//     inTable paragraph the partial-path table case needs.
//   * ROUND TRIP (the plan applied through the Wave-A FakeDocs delete model):
//     the doc ends with EXACTLY the analytics removed — the planner's ranges,
//     fed to the real sequential splice, leave precisely the non-analytics text.
// No jest snapshots (house gdocs culture); request shapes are hand-pinned.

import { parseDocument } from "../google-docs/src/core/parse";
import { ANALYTICS_FG_HEX, ANALYTICS_PT } from "../google-docs/src/core/constants";
import { planDeleteAnalytics } from "../google-docs/src/core/deleteAnalytics";
import { DocsRequest, GDoc, RequestGroup } from "../google-docs/src/core/types";
import { FakeDocs } from "./fakeDocs";
import { buildDoc, GeSpec, GpSpec, para } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** An analytics text element spec (the off-palette navy + 14pt the
 * analytic-ify tool writes). Size is irrelevant to detection (color-only), but
 * set to ANALYTICS_PT so fixtures model a real analytic-ify'd run. */
function navy(text: string, extra: Partial<GeSpec> = {}): GeSpec {
  return { text, fg: ANALYTICS_FG_HEX, size: ANALYTICS_PT, ...extra };
}

/** Every deleteContentRange range across all groups, in EMISSION order (the
 * planner's descending order is observable here). */
function deleteRanges(groups: RequestGroup[]): { startIndex: number; endIndex: number }[] {
  const out: { startIndex: number; endIndex: number }[] = [];
  for (const g of groups) for (const q of g.requests) if ("deleteContentRange" in q) out.push(q.deleteContentRange.range);
  return out;
}

/** The request-kind tag for the emit-only-deleteContentRange audit (003-F1). */
function reqKind(q: DocsRequest): string {
  if ("deleteContentRange" in q) return "delete";
  if ("updateTextStyle" in q) return "text";
  if ("updateParagraphStyle" in q) return "paragraph";
  if ("updateNamedStyle" in q) return "named";
  if ("createNamedRange" in q) return "createRange";
  if ("deleteNamedRange" in q) return "deleteRange";
  return "OTHER";
}

/** The production read path: fetch the fake and decode with parse.ts. */
async function view(fake: FakeDocs): Promise<GDoc> {
  return parseDocument(await fake.fetchDocument());
}

/** The body as the user reads it (newlines shown "¶", chips "␂"), one string
 * per paragraph, joined by "|". The projection a delete bug shows up in. */
function bodyText(doc: GDoc): string {
  return doc.paragraphs
    .map((p) => p.elements.map((el) => (el.kind === "other" ? "␂" : el.text)).join("").replace(/\n/g, "¶"))
    .join("|");
}

/**
 * Round-trip a delete plan through the FakeDocs sequential splice model: build
 * the same fixture as a fake, plan against the parsed view, apply the plan's
 * groups (flattened, IN EMISSION ORDER — descending) as one atomic batch, and
 * return the post-delete parsed view. Proves the planner's ranges removed
 * exactly the analytics text against the real index-shifting model.
 */
async function roundTrip(specs: GpSpec[]): Promise<GDoc> {
  const fake = new FakeDocs(specs);
  const before = await view(fake);
  const { groups } = planDeleteAnalytics(before);
  const requests = groups.flatMap((g) => g.requests);
  if (requests.length > 0) await fake.applyBatch(requests, before.revisionId);
  return view(fake);
}

// ---------------------------------------------------------------------------
// 1. Classification + range geometry (pure planner)
// ---------------------------------------------------------------------------

describe("planDeleteAnalytics classifies paragraphs and shapes ranges (plan §3)", () => {
  it("WHOLLY-analytics NON-final paragraph deletes the whole range INCL. the newline", () => {
    // p0 "Analytics\n" = [1,11) is wholly navy; a following body line keeps it
    // off the segment end, so its newline (index 10) IS deleted -> line collapses.
    const doc = buildDoc([{ elements: [navy("Analytics")] }, para("body")]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(deleteRanges(groups)).toEqual([{ startIndex: 1, endIndex: 11 }]);
    expect(result.paragraphsAffected).toBe(1);
    expect(result.runsDeleted).toBe(1);
  });

  it("WHOLLY-analytics FINAL-segment paragraph clamps the newline (empty line remains)", () => {
    // The ONLY paragraph is segment-final: "Analytics\n" = [1,11); the doc-final
    // newline (index 10) is UNREMOVABLE, so the whole-range delete clamps to
    // [1,10) and an empty line is left behind (documented).
    const doc = buildDoc([{ elements: [navy("Analytics")] }]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(deleteRanges(groups)).toEqual([{ startIndex: 1, endIndex: 10 }]);
    expect(result.paragraphsAffected).toBe(1);
  });

  it("PARTIAL paragraph deletes ONLY the analytics runs, never the surrounding text", () => {
    // "keep " [1,6) + navy "DROP" [6,10) + " keep\n" [10,16). Only [6,10) is
    // analytics; the two "keep" runs (and the newline) stay untouched.
    const doc = buildDoc([
      { elements: [{ text: "keep " }, navy("DROP"), { text: " keep" }] },
      para("after")
    ]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(deleteRanges(groups)).toEqual([{ startIndex: 6, endIndex: 10 }]);
    expect(result.paragraphsAffected).toBe(1);
    expect(result.runsDeleted).toBe(1);
  });

  it("a wholly-NAVY-TEXT paragraph WITH a chip takes the PARTIAL path (chip span + newline untouched)", () => {
    // Every TEXT run is navy, but a kind:"other" chip means the paragraph is NOT
    // wholly-analytics (a chip can never be deleted as analytics — A9). So the
    // partial path runs: only the navy text runs are deleted, the chip's index
    // space [4,9) survives, and the line does NOT collapse.
    // Layout: navy "Big"[1,4) + chip[4,9) + navy "Deal\n"[9,14) (the "\n" rides
    // inside the trailing "Deal" run, so that run spans [9,14)).
    const doc = buildDoc([
      { elements: [navy("Big"), { text: "@chip", kind: "other" }, navy("Deal")] },
      para("after")
    ]);
    const { groups, result } = planDeleteAnalytics(doc);
    // Two separate analytics runs, broken by the chip -> two ranges (the chip
    // gap means they do NOT coalesce). Emitted DESCENDING. The PARTIAL-PATH
    // NEWLINE CLAMP applies: the trailing navy run reaches p.endIndex (14), so
    // its delete STOPS one short at 13 — the paragraph mark survives and the
    // line is NOT collapsed (the post-fix behavior; the pre-fix range was the
    // data-loss [9,14) that swallowed the newline). The leading "Big" run is
    // interior (ends at 4, before p.endIndex), so it is untrimmed at [1,4).
    expect(deleteRanges(groups)).toEqual([
      { startIndex: 9, endIndex: 13 },
      { startIndex: 1, endIndex: 4 }
    ]);
    expect(result.paragraphsAffected).toBe(1); // one paragraph, two runs
    expect(result.runsDeleted).toBe(2);
  });

  it("a TABLE paragraph that is all-navy takes the PARTIAL path AND keeps its structural newline", () => {
    // Tables ARE processed (privacy purpose) but ONLY the partial path: the navy
    // run is deleted, the table paragraph's newline is NOT (structure untouched).
    // p0 table "cell\n" [1,6) all navy; the trailing "\n" rides inside the navy
    // run, so the single navy element spans [1,6) and REACHES p.endIndex (6).
    const doc = buildDoc([
      { inTable: true, elements: [navy("cell")] },
      para("body")
    ]);
    const { groups, result } = planDeleteAnalytics(doc);
    // PARTIAL-PATH NEWLINE CLAMP: a table paragraph is never wholly-analytics
    // (inTable forces the partial path), and the partial path trims the run that
    // reaches p.endIndex to p.endIndex - 1. So the cell's structural newline at
    // index 5 SURVIVES and only the "cell" text [1,5) is deleted — the cell's
    // paragraph boundary (and thus the table grid) is left intact. The pre-fix
    // range was the data-loss [1,6) that tore out the cell's newline.
    expect(deleteRanges(groups)).toEqual([{ startIndex: 1, endIndex: 5 }]);
    expect(result.paragraphsAffected).toBe(1);
  });

  it("a TABLE paragraph with BOTH navy and plain text deletes ONLY the navy run", () => {
    // The discriminating table case: a wholly-analytics classification would
    // delete the whole line; the partial path deletes only the navy run. Here
    // "plain " [1,7) + navy "X" [7,8) + " plain\n" [8,15): only [7,8) goes.
    const doc = buildDoc([
      { inTable: true, elements: [{ text: "plain " }, navy("X"), { text: " plain" }] },
      para("body")
    ]);
    const { groups } = planDeleteAnalytics(doc);
    expect(deleteRanges(groups)).toEqual([{ startIndex: 7, endIndex: 8 }]);
  });

  it("emits ONLY deleteContentRange — never a style/named/range write (003-F1 audit)", () => {
    // The audit that pins Delete-analytics as the sole, isolated content-deleter:
    // a mix of whole and partial paragraphs still emits exclusively delete.
    const doc = buildDoc([
      { elements: [navy("whole")] },
      { elements: [{ text: "keep" }, navy("part")] },
      para("tail")
    ]);
    const { groups } = planDeleteAnalytics(doc);
    const kinds = groups.flatMap((g) => g.requests.map(reqKind));
    expect(kinds.length).toBeGreaterThan(0); // not vacuous
    for (const k of kinds) expect(k).toBe("delete");
  });
});

// ---------------------------------------------------------------------------
// 2. Clamp-then-coalesce ORDER + descending emission (the load-bearing core)
// ---------------------------------------------------------------------------

describe("planDeleteAnalytics: clamp THEN coalesce THEN sort descending (plan §3 order)", () => {
  it("flattens a multi-paragraph doc into a strictly DESCENDING range stream (partial runs keep their line)", async () => {
    // Three PARTIAL paragraphs, each "<plain> <navy>" where the navy run is the
    // paragraph's LAST element and so carries the trailing "\n", plus a trailing
    // body paragraph (keeps p2 off the segment end). The PARTIAL-PATH NEWLINE
    // CLAMP (exec-review BLOCKER fix) trims each navy run off its paragraph mark,
    // so the surviving plain text keeps its OWN line (no collapse/merge), and the
    // emitted ranges are strictly descending by startIndex regardless of order.
    //   p0 ["a " [1,3), navy "X\n" [3,5)] -> analytics run [3,5) CLAMPED to [3,4).
    //   p1 ["b " [5,7), navy "Y\n" [7,9)] -> [7,9) CLAMPED to [7,8).
    //   p2 ["c " [9,11), navy "Z\n" [11,13)] -> [11,13) CLAMPED to [11,12).
    //   p3 "tail\n" [13,18) (segment-final), no analytics.
    const specs: GpSpec[] = [
      { elements: [{ text: "a " }, navy("X")] },
      { elements: [{ text: "b " }, navy("Y")] },
      { elements: [{ text: "c " }, navy("Z")] },
      para("tail")
    ];
    const ranges = deleteRanges(planDeleteAnalytics(buildDoc(specs)).groups);
    expect(ranges).toEqual([
      { startIndex: 11, endIndex: 12 },
      { startIndex: 7, endIndex: 8 },
      { startIndex: 3, endIndex: 4 }
    ]);
    // Strictly descending start indexes (the load-bearing emission order).
    for (let i = 1; i < ranges.length; i++) expect(ranges[i].startIndex).toBeLessThan(ranges[i - 1].startIndex);
    // BEHAVIORAL proof through the FakeDocs splice model (not just ranges): the
    // navy text is gone but EVERY partial line keeps its boundary — 4 paragraphs
    // survive, none merges into the next. (Pre-fix this read "a Xb Yc Ztail" as
    // the deleted newlines collapsed the lines together.)
    expect(bodyText(await roundTrip(specs))).toBe("a ¶|b ¶|c ¶|tail¶");
  });

  it("COALESCES two adjacent whole-analytics paragraphs into ONE range (both newlines go)", () => {
    // p0 "AA\n" [1,4) wholly navy, p1 "BB\n" [4,7) wholly navy, p2 body keeps both
    // NON-final so BOTH newlines (indexes 3 and 6) are deletable. The two whole
    // ranges [1,4) and [4,7) TOUCH at 4 -> coalesce into [1,7). Still counts as
    // 2 affected paragraphs (the user sees two lines disappear).
    const doc = buildDoc([{ elements: [navy("AA")] }, { elements: [navy("BB")] }, para("body")]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(deleteRanges(groups)).toEqual([{ startIndex: 1, endIndex: 7 }]);
    expect(result.runsDeleted).toBe(1); // one coalesced range
    expect(result.paragraphsAffected).toBe(2); // but two paragraphs swallowed
  });

  it("a whole-analytics FINAL paragraph does NOT merge ACROSS the clamped newline", () => {
    // p0 interior whole-analytics "AA\n" [1,4) -> range [1,4) (newline 3 deleted).
    // p1 FINAL whole-analytics "BB\n" [4,7) -> whole range [4,7) CLAMPED to [4,6)
    // (the doc-final newline at 6 is unremovable). Because the clamp runs BEFORE
    // coalescing: [1,4) and [4,6) still TOUCH at 4, so they DO coalesce into
    // [1,6) — but the merged range STOPS at the clamp (index 6, the final
    // newline, is NOT in any range). The "does not merge across the clamped
    // newline" guarantee: nothing reaches over index 6. (Clamp-AFTER-coalesce
    // would instead have produced [1,7) then clamped to [1,6) — same end here,
    // but the order matters for the next case where it diverges.)
    const doc = buildDoc([{ elements: [navy("AA")] }, { elements: [navy("BB")] }]);
    const { groups } = planDeleteAnalytics(doc);
    const ranges = deleteRanges(groups);
    expect(ranges).toEqual([{ startIndex: 1, endIndex: 6 }]);
    // The clamped final newline (index 6) is excluded from every range.
    for (const rg of ranges) expect(rg.endIndex).toBeLessThanOrEqual(6);
  });

  it("clamp-BEFORE-coalesce keeps a clamped final line SEPARATE from a non-adjacent interior", () => {
    // The order-divergence case: p0 interior whole-analytics, p1 a NON-analytics
    // interior body line (a gap with no analytics), p2 FINAL whole-analytics.
    //   p0 "AA\n" [1,4) -> range [1,4).
    //   p1 "body\n" [4,9) -> NO analytics, contributes NOTHING (so no range
    //      touches across it).
    //   p2 FINAL "BB\n" [9,12) -> clamped [9,11).
    // The interior delete [1,4) and the final delete [9,11) are NOT adjacent
    // (the untouched body line sits between), so they stay TWO ranges, emitted
    // descending. Non-analytics is never swept into a range (003-F3).
    const doc = buildDoc([{ elements: [navy("AA")] }, para("body"), { elements: [navy("BB")] }]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(deleteRanges(groups)).toEqual([
      { startIndex: 9, endIndex: 11 },
      { startIndex: 1, endIndex: 4 }
    ]);
    expect(result.paragraphsAffected).toBe(2);
    expect(result.runsDeleted).toBe(2);
  });
});

// ---------------------------------------------------------------------------
// 3. Negative / boundary cases
// ---------------------------------------------------------------------------

describe("planDeleteAnalytics never touches non-analytics and no-ops cleanly", () => {
  it("a doc with ZERO analytics yields zero groups and a zero-count result", () => {
    const doc = buildDoc([para("plain body"), { style: "HEADING_1", elements: [{ text: "Heading" }] }]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(groups).toEqual([]);
    expect(result).toEqual({ paragraphsAffected: 0, runsDeleted: 0 });
  });

  it("the near-miss 'dark blue 2' (#0b5394) is NOT analytics and is never deleted (003-F3)", () => {
    // Detection is EXACT: the genuine palette navy must survive untouched, or a
    // user picking real dark blue 2 would lose their text.
    const doc = buildDoc([
      { elements: [{ text: "near", fg: "#0b5394", size: ANALYTICS_PT }] },
      para("body")
    ]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(groups).toEqual([]);
    expect(result.paragraphsAffected).toBe(0);
  });

  it("an empty segment-final analytics paragraph (lone newline) yields no range", () => {
    // An empty paragraph has only its "\n"; if it were somehow classified navy,
    // the whole-range delete [start, start+1) clamps to empty on the final
    // segment -> dropped. (Here the lone "\n" run is not navy, so it is not even
    // a candidate; the assertion pins the no-op either way.)
    const doc = buildDoc([para("")]);
    const { groups, result } = planDeleteAnalytics(doc);
    expect(groups).toEqual([]);
    expect(result.paragraphsAffected).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// 4. Round trip through the FakeDocs sequential delete model
// ---------------------------------------------------------------------------

describe("planDeleteAnalytics round trip: the doc ends with EXACTLY analytics removed", () => {
  it("removes a partial paragraph's navy run, leaving the surrounding text intact", async () => {
    const after = await roundTrip([
      { elements: [{ text: "keep " }, navy("DROP"), { text: " keep" }] },
      para("after")
    ]);
    expect(bodyText(after)).toBe("keep  keep¶|after¶");
  });

  it("collapses a wholly-analytics interior line and slides the tail up", async () => {
    const after = await roundTrip([para("first"), { elements: [navy("GONE")] }, para("last")]);
    // The navy line collapses entirely (newline and all); only the two body
    // lines remain, contiguous.
    expect(after.paragraphs).toHaveLength(2);
    expect(bodyText(after)).toBe("first¶|last¶");
  });

  it("leaves an EMPTY line when the wholly-analytics paragraph is segment-final", async () => {
    const after = await roundTrip([para("keep"), { elements: [navy("ANALYTICS")] }]);
    // The final navy line's text is removed but its newline is unremovable -> an
    // empty paragraph remains where the analytics was.
    expect(after.paragraphs).toHaveLength(2);
    expect(bodyText(after)).toBe("keep¶|¶");
  });

  it("removes MANY analytics spans across paragraphs in one descending batch, exactly", async () => {
    // A mixed doc: partial p0, whole interior p1, partial p2, plain p3. Applying
    // the descending plan as one atomic batch must remove every navy span and
    // nothing else — the sequential splice only works because the ranges are
    // emitted highest-index-first. Every partial navy run is INTERIOR to its
    // line (a plain run follows it), so no partial delete touches a newline and
    // every surviving line keeps its own boundary.
    const after = await roundTrip([
      { elements: [{ text: "A " }, navy("x"), { text: " B" }] },
      { elements: [navy("WHOLE")] }, // interior whole-analytics -> collapses
      { elements: [{ text: "C " }, navy("y"), { text: " D" }] },
      para("plain")
    ]);
    // p0 keeps "A  B"; p1 collapses entirely; p2 keeps "C  D"; p3 unchanged.
    expect(after.paragraphs).toHaveLength(3);
    expect(bodyText(after)).toBe("A  B¶|C  D¶|plain¶");
  });

  it("two adjacent whole-analytics lines both collapse (the coalesced batch applies cleanly)", async () => {
    const after = await roundTrip([para("head"), { elements: [navy("AA")] }, { elements: [navy("BB")] }, para("tail")]);
    expect(after.paragraphs).toHaveLength(2);
    expect(bodyText(after)).toBe("head¶|tail¶");
  });

  // The PARTIAL-PATH NEWLINE CLAMP, proven by BEHAVIOR (exec-review data-loss
  // BLOCKER). The existing partial round trips above all put a PLAIN run AFTER
  // the navy one, so the analytics run never reached the paragraph mark and the
  // clamp was never exercised. These three pin the boundary contract on the
  // exact shapes that regressed: a TRAILING-analytics body line and a table
  // cell. Each is falsifiable against the pre-fix code, which deleted the run's
  // full [start, p.endIndex) range and tore out the paragraph mark.

  it("(a) a partial body line whose TRAILING run is analytics keeps its boundary — the line does NOT collapse", async () => {
    // p0 = "keep " [1,6) + navy "DROP\n" [6,11): the navy run OWNS the trailing
    // newline (it is p0's last text run). Pre-fix, the delete was [6,11) and the
    // newline went with "DROP", collapsing p0 into p1 -> ONE paragraph
    // "keep after¶". Post-fix the clamp trims the delete to [6,10), so the "\n"
    // at index 10 survives: "keep " stays on its OWN line and the paragraph
    // COUNT is unchanged (2). Asserting two paragraphs + the surviving boundary
    // FAILS against the pre-fix [6,11) range — the behavioral proof of the fix.
    const after = await roundTrip([{ elements: [{ text: "keep " }, navy("DROP")] }, para("after")]);
    expect(after.paragraphs).toHaveLength(2); // boundary preserved (pre-fix: 1)
    expect(bodyText(after)).toBe("keep ¶|after¶"); // pre-fix: "keep after¶"
  });

  it("(b) a table cell's analytics is removed but its STRUCTURAL newline survives — grid intact", async () => {
    // p0 is a table cell "secret\n" [1,8), wholly navy; p1 is a following body
    // line (so p0 is non-final — proving the survival is the partial-path clamp,
    // not the segment-final clamp). A table paragraph is never wholly-analytics,
    // so the partial path runs and trims the run reaching p.endIndex (8) to
    // [1,7): the cell's structural "\n" at index 7 survives, the cell paragraph
    // remains, and the count is unchanged (2). Pre-fix deleted [1,8), collapsing
    // the cell paragraph (damaging the grid) -> ONE paragraph "body¶". This test
    // FAILS pre-fix.
    const after = await roundTrip([{ inTable: true, elements: [navy("secret")] }, para("body")]);
    expect(after.paragraphs).toHaveLength(2); // cell newline kept (pre-fix: 1)
    expect(after.paragraphs[0].inTable).toBe(true); // still a table paragraph
    expect(bodyText(after)).toBe("¶|body¶"); // emptied cell + body (pre-fix: "body¶")
  });

  it("(c) a WHOLLY-analytics non-final paragraph STILL collapses — the intended newline removal is untouched by the clamp", async () => {
    // The contrast case that proves the clamp did NOT over-correct: a whole-line
    // navy paragraph between two body lines is NOT on the partial path, so its
    // newline is STILL deleted and the empty line collapses (Docs merges it up).
    // The fix narrowed ONLY the partial/table path; this deliberate collapse is
    // unchanged. (Same behavior as the interior-collapse test above, asserted
    // here beside (a)/(b) to pin the boundary between "collapse" and "preserve".)
    const after = await roundTrip([para("before"), { elements: [navy("GONE")] }, para("after")]);
    expect(after.paragraphs).toHaveLength(2); // the navy line is gone
    expect(bodyText(after)).toBe("before¶|after¶"); // contiguous, no empty line
  });

  it("the ASCENDING form of the same plan would corrupt — descending is why the round trip is exact", async () => {
    // Belt-and-suspenders on the load-bearing order: take the planner's
    // descending ranges, REVERSE them, and apply ascending. The sequential
    // splice then deletes the wrong chars, so the result differs from the
    // descending (correct) round trip — proving the planner's order is load-
    // bearing, not cosmetic (mirrors the fakeDocs differential, here driven by
    // the REAL planner output).
    const specs: GpSpec[] = [
      { elements: [{ text: "A " }, navy("x"), { text: " B" }] },
      { elements: [{ text: "C " }, navy("y"), { text: " D" }] },
      para("tail")
    ];
    const correct = await roundTrip(specs);

    const fake = new FakeDocs(specs);
    const before = await view(fake);
    const ascending = planDeleteAnalytics(before).groups.flatMap((g) => g.requests).reverse();
    await fake.applyBatch(ascending, before.revisionId);
    const corrupted = await view(fake);

    expect(bodyText(corrupted)).not.toBe(bodyText(correct));
  });
});
