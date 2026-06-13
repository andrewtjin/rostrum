// Gate + chunker suite for gdocs/src/core/guards.ts (plan S6; edge rows 13,
// 16, 18; chunker invariant A11.viii). Gates are the engine's refusal surface,
// so every test here is a failure-path test by design; the chunker tests
// assert by REQUEST IDENTITY (===) because the invariant is "the planner's
// exact objects, untorn and unreordered", not "equivalent-looking JSON".

import {
  assertNoSuggestions,
  assertNotHidden,
  assertSingleTab,
  chunkGroups,
  hasRstmState
} from "../gdocs/src/core/guards";
import { CHUNK_MAX, SENTINEL_PT } from "../gdocs/src/core/constants";
import {
  CreateNamedRangeRequest,
  DocsRequest,
  HiddenStateError,
  MultiTabError,
  RequestGroup,
  SuggestionsActiveError,
  UpdateTextStyleRequest
} from "../gdocs/src/core/types";
import { buildDoc, para, range } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** Run fn expecting a throw; return the error so shape asserts can follow.
 * (jest-circus has no global fail(), so we re-throw a plain miss instead.) */
function captureThrow(fn: () => void): unknown {
  try {
    fn();
  } catch (e) {
    return e;
  }
  throw new Error("expected fn to throw, but it returned");
}

/** A realistic shrink request — the shape the planner emits per region. */
function shrink(start: number, end: number): UpdateTextStyleRequest {
  return {
    updateTextStyle: {
      range: { startIndex: start, endIndex: end },
      textStyle: { fontSize: { magnitude: SENTINEL_PT, unit: "PT" } },
      fields: "fontSize"
    }
  };
}

/** A realistic region anchor request (name content is irrelevant to chunking). */
function anchor(start: number, end: number): CreateNamedRangeRequest {
  return {
    createNamedRange: {
      name: `rstm:v1:${end - start}x11`,
      range: { startIndex: start, endIndex: end }
    }
  };
}

/** A region's atomic pair: anchor BEFORE shrink (the emission-order pin). */
function regionGroup(start: number, end: number): RequestGroup {
  return { requests: [anchor(start, end), shrink(start, end)] };
}

/** A group of n distinct single-shrink requests (sized filler for boundaries). */
function groupOf(n: number, base = 1): RequestGroup {
  return { requests: Array.from({ length: n }, (_, i) => shrink(base + i, base + i + 1)) };
}

/** Assert chunking preserved the exact request sequence — no reorder, no
 * duplication, no drop — by object identity. */
function expectSequencePreserved(groups: RequestGroup[], chunks: DocsRequest[][]): void {
  const flatIn = groups.flatMap((g) => g.requests);
  const flatOut = chunks.flat();
  expect(flatOut).toHaveLength(flatIn.length);
  flatOut.forEach((req, i) => expect(req).toBe(flatIn[i]));
}

/** Assert no group was split across a chunk boundary: every request of a group
 * lives in the SAME chunk as its first request (identity containment). */
function expectGroupsWhole(groups: RequestGroup[], chunks: DocsRequest[][]): void {
  for (const g of groups) {
    if (g.requests.length === 0) continue;
    const home = chunks.find((c) => c.includes(g.requests[0]));
    expect(home).toBeDefined();
    for (const req of g.requests) {
      expect(home).toContain(req);
    }
  }
}

// ---------------------------------------------------------------------------
// Tab gate (plan A3, edge row 18)
// ---------------------------------------------------------------------------

describe("assertSingleTab (plan A3 — multi-tab refusal)", () => {
  it("passes a single-tab doc through", () => {
    expect(() => assertSingleTab(buildDoc([para("hello")]))).not.toThrow();
  });

  it.each([2, 3, 7])("refuses a %s-tab doc with MultiTabError", (tabCount) => {
    const doc = buildDoc([para("hello")], { tabCount });
    expect(() => assertSingleTab(doc)).toThrow(MultiTabError);
  });

  it("carries the tab count on the error (feeds the refusal copy)", () => {
    const err = captureThrow(() => assertSingleTab(buildDoc([para("x")], { tabCount: 4 })));
    expect(err).toBeInstanceOf(MultiTabError);
    expect((err as MultiTabError).tabCount).toBe(4);
  });
});

// ---------------------------------------------------------------------------
// Suggestions gate (plan D5/A7, edge row 13). Caller contract — Hide and
// Apply-styles ONLY; Show All never calls this (the reverse path is always
// available, Word parity) — is enforced at the controller, exercised by the
// flagship suite. Here we prove the gate itself fires on the flag and only
// on the flag.
// ---------------------------------------------------------------------------

describe("assertNoSuggestions (plan D5/A7 — suggestion refusal)", () => {
  it("passes a suggestion-free doc through", () => {
    expect(() => assertNoSuggestions(buildDoc([para("clean")]))).not.toThrow();
  });

  it("throws SuggestionsActiveError when the read carried any suggested* key", () => {
    const doc = buildDoc([para("pending")], { suggestionsPresent: true });
    expect(() => assertNoSuggestions(doc)).toThrow(SuggestionsActiveError);
  });
});

// ---------------------------------------------------------------------------
// rstm state detection + Apply-styles gate (plan A5, edge row 16)
// ---------------------------------------------------------------------------

describe("hasRstmState (armed-state detection, version-agnostic)", () => {
  it("is false with no named ranges at all", () => {
    expect(hasRstmState(buildDoc([para("plain")]))).toBe(false);
  });

  it("is false when only foreign ranges exist", () => {
    const doc = buildDoc([para("plain")], {
      namedRanges: [range("nr1", "kix.abc123", 1, 3), range("nr2", "bookmark-7", 3, 5)]
    });
    expect(hasRstmState(doc)).toBe(false);
  });

  it("detects a v1 sizes range", () => {
    const doc = buildDoc([para("hidden text")], {
      namedRanges: [range("nr1", "rstm:v1:42x11,7xi", 1, 8)]
    });
    expect(hasRstmState(doc)).toBe(true);
  });

  it("detects a v1 spacing range (the A12 parallel grammar)", () => {
    const doc = buildDoc([para("hidden text")], {
      namedRanges: [range("nr1", "rstm:v1:p:12x6", 1, 8)]
    });
    expect(hasRstmState(doc)).toBe(true);
  });

  it("detects an UNKNOWN-version rstm name (edge row 16: future state still counts)", () => {
    const doc = buildDoc([para("hidden by a newer build")], {
      namedRanges: [range("nr1", "rstm:v9:some-future-grammar", 1, 8)]
    });
    expect(hasRstmState(doc)).toBe(true);
  });

  it("detects rstm state mixed in among foreign ranges", () => {
    const doc = buildDoc([para("mixed")], {
      namedRanges: [range("nr1", "kix.zzz", 1, 2), range("nr2", "rstm:v1:3x11", 2, 5)]
    });
    expect(hasRstmState(doc)).toBe(true);
  });

  it("ignores names that merely CONTAIN the prefix mid-string (ownership is a prefix)", () => {
    const doc = buildDoc([para("plain")], {
      namedRanges: [range("nr1", "backup rstm:v1:1x11", 1, 3)]
    });
    expect(hasRstmState(doc)).toBe(false);
  });

  it("requires the colon — a look-alike 'rstmx:' name is foreign", () => {
    const doc = buildDoc([para("plain")], {
      namedRanges: [range("nr1", "rstmx:v1:1x11", 1, 3)]
    });
    expect(hasRstmState(doc)).toBe(false);
  });

  it("is case-sensitive — we never write upper-case, so 'RSTM:' is not ours", () => {
    const doc = buildDoc([para("plain")], {
      namedRanges: [range("nr1", "RSTM:v1:1x11", 1, 3)]
    });
    expect(hasRstmState(doc)).toBe(false);
  });
});

describe("assertNotHidden (plan A5 — Apply-styles refuses on an armed doc)", () => {
  it("passes a doc with no rstm state", () => {
    const doc = buildDoc([para("plain")], {
      namedRanges: [range("nr1", "kix.abc", 1, 3)] // foreign ranges do not arm
    });
    expect(() => assertNotHidden(doc)).not.toThrow();
  });

  it("throws HiddenStateError while any rstm range exists", () => {
    const doc = buildDoc([para("hidden")], {
      namedRanges: [range("nr1", "rstm:v1:5x11", 1, 6)]
    });
    expect(() => assertNotHidden(doc)).toThrow(HiddenStateError);
  });

  it("throws even for unknown-version state (never restyle over foreign-version hides)", () => {
    const doc = buildDoc([para("hidden")], {
      namedRanges: [range("nr1", "rstm:v2:whatever", 1, 6)]
    });
    expect(() => assertNotHidden(doc)).toThrow(HiddenStateError);
  });
});

// ---------------------------------------------------------------------------
// Chunker (plan A11.viii). Boundary tests use a tiny max so the arithmetic is
// readable; the default-cap test proves CHUNK_MAX is actually wired.
// ---------------------------------------------------------------------------

describe("chunkGroups (plan A11.viii — whole-group packing)", () => {
  it("emits no chunks for zero groups (empty/all-kept doc → zero requests)", () => {
    expect(chunkGroups([], 4)).toEqual([]);
  });

  it("packs groups totalling exactly max into ONE chunk", () => {
    const groups = [groupOf(2, 1), groupOf(2, 10)];
    const chunks = chunkGroups(groups, 4);
    expect(chunks).toHaveLength(1);
    expect(chunks[0]).toHaveLength(4);
    expectSequencePreserved(groups, chunks);
  });

  it("starts a second chunk at max+1 — the boundary falls BETWEEN groups", () => {
    const groups = [groupOf(2, 1), groupOf(2, 10), groupOf(1, 20)];
    const chunks = chunkGroups(groups, 4);
    expect(chunks.map((c) => c.length)).toEqual([4, 1]);
    // The straggler is the 5th request itself, not a copy.
    expect(chunks[1][0]).toBe(groups[2].requests[0]);
    expectSequencePreserved(groups, chunks);
    expectGroupsWhole(groups, chunks);
  });

  it("keeps a single group of max+3 WHOLE as one oversized chunk (atomicity beats the soft cap)", () => {
    const big = groupOf(7, 1); // max 4 + 3
    const chunks = chunkGroups([big], 4);
    expect(chunks).toHaveLength(1);
    expect(chunks[0]).toHaveLength(7);
    expectSequencePreserved([big], chunks);
  });

  it("never lets an oversized group absorb its neighbors", () => {
    const groups = [groupOf(2, 1), groupOf(7, 10), groupOf(2, 20)];
    const chunks = chunkGroups(groups, 4);
    // The oversized middle group flushes the chunk before it and closes itself
    // off from the one after — three chunks, each one group.
    expect(chunks.map((c) => c.length)).toEqual([2, 7, 2]);
    expectSequencePreserved(groups, chunks);
    expectGroupsWhole(groups, chunks);
  });

  it("never splits a region's createNamedRange from its updateTextStyle (identity)", () => {
    // Three region pairs under max 3: no two pairs fit together, and no pair
    // may be torn to fill the gap — each chunk is exactly one whole pair.
    const groups = [regionGroup(50, 60), regionGroup(30, 40), regionGroup(10, 20)];
    const chunks = chunkGroups(groups, 3);
    expect(chunks.map((c) => c.length)).toEqual([2, 2, 2]);
    expectGroupsWhole(groups, chunks);
    // Within each pair the anchor precedes its style write (emission-order pin).
    chunks.forEach((chunk, i) => {
      expect(chunk[0]).toBe(groups[i].requests[0]); // createNamedRange first
      expect(chunk[1]).toBe(groups[i].requests[1]); // its updateTextStyle second
    });
  });

  it("preserves group order and intra-group order verbatim (descending-index convention passes through)", () => {
    // Planner emits regions descending by start index; the chunker must not be
    // where that order silently changes (the convention is non-load-bearing,
    // but order stability IS part of this function's contract).
    const groups = [regionGroup(90, 95), groupOf(3, 50), regionGroup(10, 15)];
    const chunks = chunkGroups(groups, 5);
    expectSequencePreserved(groups, chunks);
    expectGroupsWhole(groups, chunks);
  });

  it("defaults the cap to CHUNK_MAX", () => {
    // CHUNK_MAX single-request groups fill chunk 1 exactly; one more spills.
    const groups = Array.from({ length: CHUNK_MAX + 1 }, (_, i) => groupOf(1, i * 2 + 1));
    const chunks = chunkGroups(groups);
    expect(chunks.map((c) => c.length)).toEqual([CHUNK_MAX, 1]);
  });

  it("skips empty groups without flushing early or emitting empty chunks", () => {
    const empty: RequestGroup = { requests: [] };
    const groups = [groupOf(2, 1), empty, groupOf(2, 10)];
    const chunks = chunkGroups(groups, 4);
    expect(chunks.map((c) => c.length)).toEqual([4]);
    expectSequencePreserved(groups, chunks);
  });

  it("returns nothing for all-empty groups (the API rejects an empty batch)", () => {
    expect(chunkGroups([{ requests: [] }, { requests: [] }], 4)).toEqual([]);
  });

  it("degrades a sub-group max to one chunk per group — never drops or splits", () => {
    // A max smaller than every group is degenerate config, not data loss: each
    // group still ships whole, one per chunk.
    const groups = [regionGroup(30, 40), regionGroup(10, 20)];
    const chunks = chunkGroups(groups, 1);
    expect(chunks.map((c) => c.length)).toEqual([2, 2]);
    expectSequencePreserved(groups, chunks);
    expectGroupsWhole(groups, chunks);
  });
});
