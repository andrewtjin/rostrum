// THE FLAGSHIP SUITE (plan S9): the controller verbs driven end to end over
// the in-memory DocsPort fake (fakeDocs.ts) — fetch through parse through
// plan through chunked, revision-chained apply, with every read decoded by
// the PRODUCTION parser (the fixture-realism rail, plan A11.ii). Each named
// case from the loop charter has a test here:
//
//   (a) hide -> showAll round trip, view byte-equal (001-S4)
//   (b) hide twice converged, counts say so (row 7)
//   (c) hide -> destroyRange -> showAll converges, no sentinel text (001-F5/S7)
//   (d) hide -> splitRange -> showAll restores both fragments (row 9)
//   (e) foreign edit before the first chunk -> silent retry (2 fetches)
//   (f) foreign edit BETWEEN chunks -> PartialApplyError; showAll unwinds
//   (g) suggestions: Hide refused untouched; SHOW ALL SUCCEEDS (A7, 001-F11)
//   (h) retries exhausted -> RevisionConflictError, doc untouched
//   (i) highlight inside a hidden region -> re-hide surfaces it; showAll exact (A1)
//   (j) 14pt pasted into a hidden region -> showAll preserves it (A6)
//   (k) pure-sweep consent: ask / consent sweeps / decline untouched (A14)
//   (l) applyStyles refused on hidden (001-F10) and suggestion docs
//   (m) applyStyles named-style 400 -> degraded path, retro still lands (D13)
//   (n) showAll interrupted between chunks -> second showAll converges (A11.iv)
//   (o) multi-tab doc -> every verb refuses (001-F9)
//   (p) chip doc hides cleanly around chips (001-F12)
//
// Equality discipline: "byte-equal" is asserted on a PER-CHARACTER projection
// of the parsed view (charView below), not on element arrays — applyBatch
// legitimately fragments text runs at write boundaries, and what 001-S4
// promises the user is that every character reads back with its original
// size/emphasis/highlight, not that Google's internal run layout survived.

import type { SelectionPick } from "../google-docs/src/core/adapterPure";
import {
  ANALYTICS_FG_HEX,
  ANALYTICS_PT,
  MAX_REPLAN_ATTEMPTS,
  SENTINEL_PT,
  SENTINELS
} from "../google-docs/src/core/constants";
import {
  analyticify,
  applyStyles,
  deleteAnalytics,
  hide,
  markCite,
  showAll
} from "../google-docs/src/core/controller";
import { parseDocument } from "../google-docs/src/core/parse";
import { isRstmName } from "../google-docs/src/core/rangeNames";
import { errorMessage, STRINGS } from "../google-docs/src/core/strings";
import {
  DocsApiError,
  GDoc,
  HiddenStateError,
  MultiTabError,
  PartialApplyError,
  RevisionConflictError,
  RevisionMismatchError,
  SuggestionsActiveError
} from "../google-docs/src/core/types";
import { FakeDocs, FakeDocsOptions } from "./fakeDocs";
import { GpSpec } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** The production read path: fetch the fake and decode with parse.ts. NOTE:
 * this counts as a fetch — tests asserting fetch counts take deltas around
 * the verb call alone. */
async function view(fake: FakeDocs): Promise<GDoc> {
  return parseDocument(await fake.fetchDocument());
}

/**
 * Per-character projection of everything Hide may touch and Show All must
 * restore: each char's size/bold/background plus each paragraph's style and
 * direct spacing. Two projections are equal iff the user would see identical
 * documents — run fragmentation is invisible here on purpose (see header).
 */
function charView(doc: GDoc): string {
  return doc.paragraphs
    .map((p) => {
      const head = `${p.namedStyleType}(${p.spaceAbovePt ?? "i"}/${p.spaceBelowPt ?? "i"}):`;
      const chars = p.elements
        .map((el) => {
          if (el.kind === "other") return `<other:${el.endIndex - el.startIndex}>`;
          let out = "";
          for (let i = 0; i < el.text.length; i++) {
            const ch = el.text[i] === "\n" ? "¶" : el.text[i];
            out += `${ch}@${el.fontSizePt ?? "i"}${el.bold ? "b" : ""}${el.backgroundHex ?? ""};`;
          }
          return out;
        })
        .join("");
      return head + chars;
    })
    .join("\n");
}

/** Count of text chars currently AT a sentinel size — the "still hidden"
 * probe; 0 together with rstmCount 0 is the converged all-visible state. */
function sentinelChars(doc: GDoc): number {
  let n = 0;
  for (const p of doc.paragraphs) {
    for (const el of p.elements) {
      if (el.kind === "text" && el.fontSizePt !== null && SENTINELS.includes(el.fontSizePt)) {
        n += el.text.length;
      }
    }
  }
  return n;
}

/** rstm-family ranges still present (armed-state probe). */
function rstmCount(doc: GDoc): number {
  return doc.namedRanges.filter((nr) => isRstmName(nr.name)).length;
}

/** Await a rejection of EXACTLY the given class and hand the typed error back
 * so field assertions (verb, chunk counts) read naturally. */
async function expectRejects<T extends Error>(
  p: Promise<unknown>,
  cls: new (...args: never[]) => T
): Promise<T> {
  try {
    await p;
  } catch (e) {
    if (e instanceof cls) return e;
    throw new Error(`expected ${cls.name}, got: ${String(e)}`);
  }
  throw new Error(`expected ${cls.name}, but the promise resolved`);
}

// ---------------------------------------------------------------------------
// Shared docs
// ---------------------------------------------------------------------------

/**
 * The mixed doc (same geometry the planner suite pins). Body indexes:
 *   P0 [1,10)   HEADING_4 "Tag line\n"                  kept (heading)
 *   P1 [10,30)  bold-14 "Smith 24" + " says stuff\n"    kept (signature cite)
 *   P2 [30,50)  "The quick " + yellow "brown" + " fox\n" mixed
 *   P3 [50,66)  "All hidden here\n"                      hidden, LAST (clamp)
 * Fresh hide: regions [30,40) and [45,65) — 30 sentinel chars, 2 anchors.
 */
function mixedFake(opts: FakeDocsOptions = {}): FakeDocs {
  return new FakeDocs(
    [
      { style: "HEADING_4", elements: [{ text: "Tag line" }] },
      { elements: [{ text: "Smith 24", bold: true, size: 14 }, { text: " says stuff" }] },
      { elements: [{ text: "The quick " }, { text: "brown", bg: "#ffff00" }, { text: " fox" }] },
      { elements: [{ text: "All hidden here" }] }
    ],
    opts
  );
}

/** The fresh-hide result the mixed doc always produces (pinned once, DRY). */
const MIXED_HIDE_RESULT = {
  paragraphsScanned: 4,
  paragraphsChanged: 2,
  regionsHidden: 2,
  regionsAlreadyHidden: 0,
  newlyKeptRestored: 0,
  preexistingTinyCount: 0
};

/**
 * A doc big enough to force MULTIPLE batchUpdate chunks (CHUNK_MAX 5000):
 * `bodies` hidden paragraphs separated by kept headings -> `bodies` regions
 * of 2 requests each. 2501 bodies = 5002 requests = chunks of [5000, 2] —
 * the smallest shape that exercises mid-sequence interruption (cases f/n).
 */
function bigSpecs(bodies: number): GpSpec[] {
  const out: GpSpec[] = [];
  for (let i = 0; i < bodies; i++) {
    out.push({ elements: [{ text: `evidence ${i}` }] });
    out.push({ style: "HEADING_4", elements: [{ text: `tag ${i}` }] });
  }
  return out;
}
const BIG_BODIES = 2501;

// ---------------------------------------------------------------------------
// (a) + (b): round trip and convergence
// ---------------------------------------------------------------------------

describe("hide -> showAll round trip", () => {
  it("(a) restores a per-character view byte-equal to the original (001-S4)", async () => {
    const fake = mixedFake();
    const original = charView(await view(fake));

    const r = await hide(fake);
    expect(r).toEqual(MIXED_HIDE_RESULT);
    expect(fake.appliedBatches).toHaveLength(1); // 4 requests fit one chunk

    const hidden = await view(fake);
    expect(sentinelChars(hidden)).toBe(30); // [30,40) + [45,65): final \n clamped
    expect(rstmCount(hidden)).toBe(2);
    // Kept text untouched in the hidden state: the cite still reads 14pt bold.
    expect(hidden.paragraphs[1].elements[0].fontSizePt).toBe(14);

    const out = await showAll(fake);
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 2,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 2,
        rangesSkippedNewerVersion: 0
      }
    });

    const after = await view(fake);
    expect(charView(after)).toBe(original);
    expect(rstmCount(after)).toBe(0);
    expect(sentinelChars(after)).toBe(0);
  });

  it("(b) a second hide is a converged no-op and its counts say so (row 7)", async () => {
    const fake = mixedFake();
    await hide(fake);
    const batchesAfterFirst = fake.appliedBatches.length;

    const second = await hide(fake);
    expect(second).toEqual({
      paragraphsScanned: 4,
      paragraphsChanged: 0,
      regionsHidden: 0,
      regionsAlreadyHidden: 2,
      newlyKeptRestored: 0,
      preexistingTinyCount: 0
    });
    // Converged means converged: the second pass WROTE nothing.
    expect(fake.appliedBatches).toHaveLength(batchesAfterFirst);
  });
});

// ---------------------------------------------------------------------------
// (c) + (d): manifest decay — destroyed and split ranges
// ---------------------------------------------------------------------------

describe("showAll under manifest decay", () => {
  it("(c) a destroyed range's text is swept clean — convergence, no sentinel text (001-F5/S7)", async () => {
    const fake = mixedFake();
    const original = charView(await view(fake));
    await hide(fake);

    const armed = await view(fake);
    const tail = armed.namedRanges.find((nr) => nr.name === "rstm:v1:20xi");
    if (tail === undefined) throw new Error("expected the [45,65) anchor");
    fake.destroyRange(tail.id); // cut/paste killed the record; the tiny text remains

    const out = await showAll(fake);
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 1,
        segmentsNormalized: 0,
        sweptOrphans: 1,
        rangesDeleted: 1,
        rangesSkippedNewerVersion: 0
      }
    });

    const after = await view(fake);
    expect(sentinelChars(after)).toBe(0);
    expect(rstmCount(after)).toBe(0);
    // The hidden text was inherit-size, so even the sweep's normalize IS the
    // exact restore here — the view comes back byte-equal.
    expect(charView(after)).toBe(original);

    // And convergence holds: a second showAll finds nothing at all.
    const again = await showAll(fake);
    expect(again).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 0,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 0,
        rangesSkippedNewerVersion: 0
      }
    });
  });

  it("(d) a split range restores both fragments to the original view (row 9)", async () => {
    const fake = mixedFake();
    const original = charView(await view(fake));
    await hide(fake);

    const armed = await view(fake);
    const tail = armed.namedRanges.find((nr) => nr.name === "rstm:v1:20xi");
    if (tail === undefined) throw new Error("expected the [45,65) anchor");
    fake.splitRange(tail.id, 54); // an edit split the record into [45,54)+[54,65)

    const out = await showAll(fake);
    // A PURE split: the two fragments sum to the record's 20 chars, so the
    // RLE walks continuously across them and BOTH restore exactly — no amber
    // (counting a byte-exact restore as normalized would misreport perfect
    // work). The 001-S4 criterion asserted here is the byte-equal VIEW.
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 3,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 2,
        rangesSkippedNewerVersion: 0
      }
    });

    const after = await view(fake);
    expect(charView(after)).toBe(original);
    expect(sentinelChars(after)).toBe(0);
    expect(rstmCount(after)).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// (e), (h): first-chunk races — silent retry and exhaustion
// ---------------------------------------------------------------------------

describe("revision races on the first chunk", () => {
  it("(e) a foreign edit before the first chunk retries silently from a fresh fetch (2 fetches)", async () => {
    const fake = mixedFake();
    // The teammate edits between our fetch and our first write — once.
    fake.beforeApply = (call) => {
      if (call === 0) fake.injectForeignEdit();
    };

    const fetchesBefore = fake.fetchCount;
    const r = await hide(fake);
    expect(fake.fetchCount - fetchesBefore).toBe(2); // initial + ONE re-plan
    expect(r).toEqual(MIXED_HIDE_RESULT);
    expect(fake.appliedBatches).toHaveLength(1); // the retry's chunk landed
    expect(sentinelChars(await view(fake))).toBe(30);
  });

  it("(h) retries exhausted -> RevisionConflictError(hide), doc untouched", async () => {
    const fake = mixedFake();
    const original = charView(await view(fake));
    fake.beforeApply = () => fake.injectForeignEdit(); // every attempt loses the race

    const fetchesBefore = fake.fetchCount;
    const err = await expectRejects(hide(fake), RevisionConflictError);
    expect(err.verb).toBe("hide");
    expect(fake.fetchCount - fetchesBefore).toBe(1 + MAX_REPLAN_ATTEMPTS);
    expect(fake.appliedBatches).toHaveLength(0);

    fake.beforeApply = null;
    const after = await view(fake);
    expect(charView(after)).toBe(original);
    expect(rstmCount(after)).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Mark cite — the same silent reconcile, so concurrent marks don't reject
// ---------------------------------------------------------------------------

describe("markCite concurrent reconcile", () => {
  // A plain single-tab, suggestion-free, unarmed doc: every Mark-cite gate passes.
  const markFake = (): FakeDocs => new FakeDocs([{ elements: [{ text: "Smith 2020 says the thing" }] }]);
  const wholeParagraph = (ordinal: number): SelectionPick => ({
    paragraphOrdinal: ordinal,
    startOffset: 0,
    endOffset: null
  });

  it("marks a cite on a clean doc (happy path, one fetch)", async () => {
    const fake = markFake();
    const before = fake.fetchCount;
    const cited = await markCite(fake, [wholeParagraph(0)]);
    expect(cited).toBe(1);
    expect(fake.fetchCount - before).toBe(1);
    expect(fake.appliedBatches.length).toBeGreaterThanOrEqual(1);
  });

  it("nothing markable (empty selection) applies nothing and teaches a zero count", async () => {
    const fake = markFake();
    const cited = await markCite(fake, []);
    expect(cited).toBe(0);
    expect(fake.appliedBatches).toHaveLength(0);
  });

  it("a concurrent edit before the write RECONCILES silently — the cite still lands, no throw (2 fetches)", async () => {
    const fake = markFake();
    // A second Mark cite (or a teammate) commits between our fetch and our write, once.
    fake.beforeApply = (call) => {
      if (call === 0) fake.injectForeignEdit();
    };
    const before = fake.fetchCount;
    const cited = await markCite(fake, [wholeParagraph(0)]);
    expect(cited).toBe(1); // reconciled, not the revision-conflict dialog
    expect(fake.fetchCount - before).toBe(2); // initial + ONE re-plan from a fresh fetch
    expect(fake.appliedBatches).toHaveLength(1); // the retry's chunk landed
  });

  it("retries exhausted -> RevisionMismatchError (the honest refusal), nothing applied", async () => {
    const fake = markFake();
    fake.beforeApply = () => fake.injectForeignEdit(); // every attempt loses the race
    const before = fake.fetchCount;
    await expectRejects(markCite(fake, [wholeParagraph(0)]), RevisionMismatchError);
    expect(fake.fetchCount - before).toBe(1 + MAX_REPLAN_ATTEMPTS);
    expect(fake.appliedBatches).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// (f), (n): mid-sequence interruption on multi-chunk applies
// ---------------------------------------------------------------------------

describe("multi-chunk interruption (PartialApplyError)", () => {
  it("(f) a foreign edit between hide chunks -> PartialApplyError; showAll fully unwinds", async () => {
    const fake = new FakeDocs(bigSpecs(BIG_BODIES));
    const original = charView(await view(fake));
    // Chunk 0 (call 0) lands; the edit arrives before chunk 1 (call 1).
    fake.beforeApply = (call) => {
      if (call === 1) fake.injectForeignEdit();
    };

    const err = await expectRejects(hide(fake), PartialApplyError);
    expect(err.verb).toBe("hide");
    expect(err.appliedChunks).toBe(1);
    expect(err.totalChunks).toBe(2);

    // The torn state is real: most regions armed, the last chunk's missing.
    fake.beforeApply = null;
    const torn = await view(fake);
    expect(rstmCount(torn)).toBe(BIG_BODIES - 1);
    expect(sentinelChars(torn)).toBeGreaterThan(0);

    // The recovery route the failure copy promises: Show All unwinds it all.
    const out = await showAll(fake);
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: BIG_BODIES - 1,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: BIG_BODIES - 1,
        rangesSkippedNewerVersion: 0
      }
    });
    const after = await view(fake);
    expect(charView(after)).toBe(original);
    expect(sentinelChars(after)).toBe(0);
    expect(rstmCount(after)).toBe(0);
  });

  it("(n) showAll interrupted between restore chunks -> PartialApplyError(showAll); a second showAll converges (A11.iv)", async () => {
    const fake = new FakeDocs(bigSpecs(BIG_BODIES));
    const original = charView(await view(fake));
    await hide(fake); // clean 2-chunk hide: applyBatch calls 0 and 1

    // Interrupt the SECOND restore chunk (call 3 — calls 2,3 are showAll's).
    fake.beforeApply = (call) => {
      if (call === 3) fake.injectForeignEdit();
    };
    const err = await expectRejects(showAll(fake), PartialApplyError);
    expect(err.verb).toBe("showAll");
    expect(err.appliedChunks).toBe(1);
    expect(err.totalChunks).toBe(2);

    // "Use Show All again to finish" — and it does, exactly.
    fake.beforeApply = null;
    const out = await showAll(fake);
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 1,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 1,
        rangesSkippedNewerVersion: 0
      }
    });
    const after = await view(fake);
    expect(charView(after)).toBe(original);
    expect(sentinelChars(after)).toBe(0);
    expect(rstmCount(after)).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// (g): the suggestion gate asymmetry (plan A7)
// ---------------------------------------------------------------------------

describe("suggestions", () => {
  it("(g) hide refuses a suggestion-laden doc untouched; Show All proceeds (001-F11)", async () => {
    const suggested = mixedFake({ suggestions: true });
    await expectRejects(hide(suggested), SuggestionsActiveError);
    expect(suggested.appliedBatches).toHaveLength(0);

    // The reverse path must stay available: arm a clean doc, then let a
    // teammate's suggestion appear — Show All still restores everything.
    const fake = mixedFake();
    const original = charView(await view(fake));
    await hide(fake);
    fake.setSuggestions(true);

    const out = await showAll(fake);
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 2,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 2,
        rangesSkippedNewerVersion: 0
      }
    });
    const after = await view(fake);
    expect(sentinelChars(after)).toBe(0);
    expect(rstmCount(after)).toBe(0);
    expect(charView(after)).toBe(original);
  });
});

// ---------------------------------------------------------------------------
// (i), (j): the reconcile and the foreign-text guarantee
// ---------------------------------------------------------------------------

describe("edits inside hidden regions", () => {
  it("(i) a highlight added inside a hidden region surfaces on re-hide; showAll stays exact (A1)", async () => {
    const fake = mixedFake();
    await hide(fake);
    // The user highlights "hidden" ([54,60)) inside the hidden [45,65) region.
    fake.injectForeignEdit((f) => f.editSetBackground(54, 60, "#00ffff"));

    const second = await hide(fake);
    expect(second).toEqual({
      paragraphsScanned: 4,
      paragraphsChanged: 0,
      regionsHidden: 0,
      regionsAlreadyHidden: 1, // the untouched [30,40) region
      newlyKeptRestored: 1,
      preexistingTinyCount: 0
    });
    // The highlighted span is back at full size while its flanks stay hidden.
    expect(sentinelChars(await view(fake))).toBe(24); // 30 - the 6 surfaced chars

    const out = await showAll(fake);
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 3,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 3,
        rangesSkippedNewerVersion: 0
      }
    });

    // Expected end state: the pristine doc with the SAME user highlight.
    const twin = mixedFake();
    twin.injectForeignEdit((f) => f.editSetBackground(54, 60, "#00ffff"));
    expect(charView(await view(fake))).toBe(charView(await view(twin)));
  });

  it("(j) 14pt text pasted into a hidden region keeps its 14pt through showAll (A6)", async () => {
    const fake = mixedFake();
    await hide(fake);
    // Same-length paste-over at 14pt inside the hidden [45,65) region.
    fake.injectForeignEdit((f) => f.editSetFontSize(50, 54, 14));

    const out = await showAll(fake);
    // The record still matches its segment length, so the restore is EXACT —
    // and it writes only over sub-spans still at the sentinel, leaving the
    // user's 14pt untouched (never flattened).
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 2,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 2,
        rangesSkippedNewerVersion: 0
      }
    });

    const twin = mixedFake();
    twin.injectForeignEdit((f) => f.editSetFontSize(50, 54, 14));
    const after = await view(fake);
    expect(charView(after)).toBe(charView(await view(twin)));
    expect(sentinelChars(after)).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// (k): the pure-sweep consent boundary (plan A14)
// ---------------------------------------------------------------------------

describe("pure-sweep consent", () => {
  /** A copied doc: sentinel-size text survives, the rstm ranges did not. */
  function copiedFake(): FakeDocs {
    return new FakeDocs([
      { style: "HEADING_4", elements: [{ text: "Tag line" }] },
      { elements: [{ text: "tiny one", size: 1 }] },
      { elements: [{ text: "normal body" }] }
    ]);
  }

  it("(k) asks first, sweeps on consent, leaves the doc untouched on decline", async () => {
    // Ask: no answer supplied -> the question, and NOTHING is written.
    const fake = copiedFake();
    const ask = await showAll(fake);
    expect(ask).toEqual({ kind: "needsConsent", unrecordedTinyCount: 1 });
    expect(fake.appliedBatches).toHaveLength(0);

    // Consent: the orphan passage is adopted and cleared.
    const consented = await showAll(fake, { sweepUnrecorded: true });
    expect(consented).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 0,
        segmentsNormalized: 0,
        sweptOrphans: 1,
        rangesDeleted: 0,
        rangesSkippedNewerVersion: 0
      }
    });
    expect(sentinelChars(await view(fake))).toBe(0);

    // Decline (fresh copy): a clean zero-write no-op — the tiny text is the
    // user's own formatting until they say otherwise.
    const declined = copiedFake();
    const before = charView(await view(declined));
    const out = await showAll(declined, { sweepUnrecorded: false });
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 0,
        segmentsNormalized: 0,
        sweptOrphans: 0,
        rangesDeleted: 0,
        rangesSkippedNewerVersion: 0
      }
    });
    expect(declined.appliedBatches).toHaveLength(0);
    expect(charView(await view(declined))).toBe(before);
  });

  it("an ARMED doc never asks: orphans next to our own state sweep without consent", async () => {
    const fake = mixedFake();
    await hide(fake);
    const armed = await view(fake);
    const anchor = armed.namedRanges.find((nr) => nr.name === "rstm:v1:10xi");
    if (anchor === undefined) throw new Error("expected the [30,40) anchor");
    fake.destroyRange(anchor.id); // one orphan, one surviving record

    const out = await showAll(fake); // no consent answer needed
    expect(out).toEqual({
      kind: "done",
      result: {
        segmentsRestoredExact: 1,
        segmentsNormalized: 0,
        sweptOrphans: 1,
        rangesDeleted: 1,
        rangesSkippedNewerVersion: 0
      }
    });
    expect(sentinelChars(await view(fake))).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// (l), (m): the styles verb — gates and the degraded path
// ---------------------------------------------------------------------------

describe("applyStyles", () => {
  /** H1 + spaced body + H4: one pocket, one spacing clear, one tag. */
  function stylesFake(): FakeDocs {
    return new FakeDocs([
      { style: "HEADING_1", elements: [{ text: "Pocket head" }] },
      { elements: [{ text: "card body" }], spaceAbovePt: 12 },
      { style: "HEADING_4", elements: [{ text: "Tag line" }] }
    ]);
  }

  it("(l) refuses a hidden doc (001-F10) and a suggestion doc, untouched", async () => {
    const armed = mixedFake();
    await hide(armed);
    const writesAfterHide = armed.appliedBatches.length;
    await expectRejects(applyStyles(armed), HiddenStateError);
    expect(armed.appliedBatches).toHaveLength(writesAfterHide);

    const suggested = mixedFake({ suggestions: true });
    await expectRejects(applyStyles(suggested), SuggestionsActiveError);
    expect(suggested.appliedBatches).toHaveLength(0);
  });

  it("applies both batches on the happy path: named styles land AND existing paragraphs restyle", async () => {
    const fake = stylesFake();
    const r = await applyStyles(fake);
    expect(r).toEqual({
      namedStylesApplied: true,
      restyled: { pocket: 1, hat: 0, block: 0, tag: 1 },
      spacingCleared: 1,
      citesRepaired: 0
    });
    expect(fake.appliedBatches).toHaveLength(2); // named batch, then retro batch

    // Named-style writes visible to later reads (future typing styles itself).
    expect(fake.namedStyleSize("NORMAL_TEXT")).toBe(11);
    expect(fake.namedStyleSize("HEADING_1")).toBe(26);
    expect(fake.namedStyleSize("HEADING_2")).toBe(22);
    expect(fake.namedStyleSize("HEADING_3")).toBe(16);
    expect(fake.namedStyleSize("HEADING_4")).toBe(14);

    // Retro writes on the existing paragraphs: pocket boxed + 26 bold, tag 14
    // bold, the import's direct spacing cleared to an explicit zero.
    expect(fake.isBoxed(0)).toBe(true);
    const after = await view(fake);
    expect(after.paragraphs[0].elements[0].fontSizePt).toBe(26);
    expect(after.paragraphs[0].elements[0].bold).toBe(true);
    expect(after.paragraphs[1].spaceAbovePt).toBe(0);
    expect(after.paragraphs[1].spaceBelowPt).toBe(0);
    expect(after.paragraphs[2].elements[0].fontSizePt).toBe(14);
    expect(after.paragraphs[2].elements[0].bold).toBe(true);
  });

  it("(m) a named-style 400 takes the degraded path: retro still lands, namedStylesApplied false (A5/D13)", async () => {
    const fake = stylesFake();
    fake.rejectRequestType("updateNamedStyle"); // the consumer-account 400

    const r = await applyStyles(fake);
    expect(r).toEqual({
      namedStylesApplied: false,
      restyled: { pocket: 1, hat: 0, block: 0, tag: 1 },
      spacingCleared: 1,
      citesRepaired: 0
    });

    // Batch 1 was rejected whole: no named-style write reached the doc…
    expect(fake.namedStyleWrites).toHaveLength(0);
    expect(fake.namedStyleSize("HEADING_1")).toBeUndefined();
    expect(fake.appliedBatches).toHaveLength(1); // …but the retro batch landed:
    expect(fake.isBoxed(0)).toBe(true);
    const after = await view(fake);
    expect(after.paragraphs[0].elements[0].fontSizePt).toBe(26);
    expect(after.paragraphs[1].spaceAbovePt).toBe(0);
    expect(after.paragraphs[2].elements[0].fontSizePt).toBe(14);
  });
});

// ---------------------------------------------------------------------------
// (o): the universal tab gate (plan A3)
// ---------------------------------------------------------------------------

describe("multi-tab docs", () => {
  it("(o) every verb refuses a multi-tab doc outright, nothing written (001-F9)", async () => {
    const fake = mixedFake({ tabCount: 2 });
    const hideErr = await expectRejects(hide(fake), MultiTabError);
    expect(hideErr.tabCount).toBe(2);
    await expectRejects(showAll(fake), MultiTabError);
    await expectRejects(applyStyles(fake), MultiTabError);
    expect(fake.appliedBatches).toHaveLength(0);
  });
});

// ---------------------------------------------------------------------------
// (p): chips (the A9 whitelist end to end)
// ---------------------------------------------------------------------------

describe("chips", () => {
  it("(p) hides cleanly AROUND a chip and restores byte-equal (001-F12)", async () => {
    // "aa" [1,3) + chip [3,8) + "bb\n" [8,11), then a kept H1 [11,15).
    const fake = new FakeDocs([
      { elements: [{ text: "aa" }, { kind: "other", text: "@@@@@" }, { text: "bb" }] },
      { style: "HEADING_1", elements: [{ text: "Hat" }] }
    ]);
    const original = charView(await view(fake));

    const r = await hide(fake);
    expect(r.regionsHidden).toBe(2); // [1,3) and [8,11) — split by the chip

    // NO applied request — style write or anchor — overlaps the chip [3,8).
    for (const batch of fake.appliedBatches) {
      for (const req of batch) {
        const range =
          "updateTextStyle" in req
            ? req.updateTextStyle.range
            : "createNamedRange" in req
              ? req.createNamedRange.range
              : null;
        if (range !== null) {
          expect(range.endIndex <= 3 || range.startIndex >= 8).toBe(true);
        }
      }
    }

    // The chip survives the hidden state intact…
    const hidden = await view(fake);
    const chip = hidden.paragraphs[0].elements.find((el) => el.kind === "other");
    expect(chip).toEqual(expect.objectContaining({ startIndex: 3, endIndex: 8 }));

    // …and the round trip comes back byte-equal, chip included.
    const out = await showAll(fake);
    expect(out.kind).toBe("done");
    expect(charView(await view(fake))).toBe(original);
  });
});

// ---------------------------------------------------------------------------
// Failure paths and wiring beyond the named cases
// ---------------------------------------------------------------------------

describe("failure paths and wiring", () => {
  it("an unmapped 400 on a hide batch propagates as DocsApiError with NOTHING applied (atomic batch)", async () => {
    const fake = mixedFake();
    const original = charView(await view(fake));
    fake.rejectRequestType("updateTextStyle");

    await expectRejects(hide(fake), DocsApiError);
    expect(fake.appliedBatches).toHaveLength(0);
    const after = await view(fake);
    expect(charView(after)).toBe(original);
    expect(rstmCount(after)).toBe(0); // the anchor never landed either
  });

  it("hide resolves settings from the port: an explicit empty keep-set hides the highlight too", async () => {
    // Doc-properties tier pins keepColors [] — the explicit "keep nothing by
    // color" choice (Word decision #11 parity) — so "brown" now hides and
    // P2+P3 fuse into ONE region.
    const fake = mixedFake({ settingsJson: JSON.stringify({ keepColors: [] }) });
    const r = await hide(fake);
    expect(r.regionsHidden).toBe(1);
    expect(sentinelChars(await view(fake))).toBe(35); // [30,65), final \n clamped
  });

  it("revision chaining is real: a clean multi-chunk hide lands every chunk against the PREVIOUS response's revision", async () => {
    const fake = new FakeDocs(bigSpecs(BIG_BODIES));
    const r = await hide(fake);
    // The fake validates requiredRevisionId on every call — two applied
    // batches prove the controller chained the response id, not the fetch id.
    expect(fake.appliedBatches).toHaveLength(2);
    expect(r.regionsHidden).toBe(BIG_BODIES);
    expect(rstmCount(await view(fake))).toBe(BIG_BODIES);
  });
});

// ---------------------------------------------------------------------------
// Loop 003 — the analytics verbs (analytic-ify + Delete analytics), driven end
// to end over the same DocsPort fake through the PRODUCTION parser. The delete
// path runs through the Wave-A spliceDelete model, so a delete actually removes
// chars and SHIFTS downstream indexes — the only way these tests are
// non-vacuous. Cases mirror plan §5:
//   * both verbs gate single-tab / suggestions / not-hidden (each mapped error)
//   * analytic-ify emits ONLY updateTextStyle (003-F2 audit) + idempotence
//   * Delete analytics emits ONLY deleteContentRange (003-F1 audit)
//   * a torn multi-chunk delete -> PartialApplyError("deleteAnalytics"),
//     mapped to the delete-specific copy (NOT errors.unknown)
//   * the TOCTOU zero-after-confirm path -> a graceful no-op result
// ---------------------------------------------------------------------------

describe("analytic-ify (controller)", () => {
  /** A plain, single-tab, suggestion-free, unarmed doc the gates all pass:
   *  P0 plain body, P1 plain body. Ordinals select which paragraphs to style. */
  const plainFake = (): FakeDocs =>
    new FakeDocs([{ elements: [{ text: "first line" }] }, { elements: [{ text: "second line" }] }]);

  it("styles the touched paragraph navy 14pt and counts it; one batch, untouched line clean", async () => {
    const fake = plainFake();
    const before = fake.fetchCount;
    const r = await analyticify(fake, new Set([0]));
    expect(r).toEqual({ paragraphsStyled: 1 });
    expect(fake.fetchCount - before).toBe(1); // one documents.get per verb (A13)
    expect(fake.appliedBatches).toHaveLength(1); // one small group fits one chunk

    // The touched line reads back BOTH analytics attributes through the parsed
    // FakeDocs view: the analytics size AND the off-palette navy foreground. The
    // Wave-1 fake now models foregroundColor writes (updateTextStyle's mask
    // honors foregroundColor, serialize() emits it, parse decodes it), so this is
    // a true write round-trip — replacing the old "deferred to the request-level
    // 003-F2 audit / foregroundHex toBeNull" workaround. Asserting the EXACT hex
    // is falsifiable: a wrong color (or a dropped foreground write) fails here,
    // and the keeper that protects analytics through Hide keys off this exact hex
    // (keepers.isAnalytics), so the read-back IS the contract Hide depends on.
    const after = await view(fake);
    expect(after.paragraphs[0].elements[0].fontSizePt).toBe(ANALYTICS_PT);
    expect(after.paragraphs[0].elements[0].foregroundHex).toBe(ANALYTICS_FG_HEX);
    // The untouched line is fully clean: no size, no foreground was written.
    expect(after.paragraphs[1].elements[0].fontSizePt).toBeNull();
    expect(after.paragraphs[1].elements[0].foregroundHex).toBeNull();
  });

  it("(003-F2 audit) emits ONLY updateTextStyle — never a paragraph/named-style/delete write", async () => {
    const fake = plainFake();
    await analyticify(fake, new Set([0, 1]));
    expect(fake.appliedBatches.length).toBeGreaterThan(0);
    for (const batch of fake.appliedBatches) {
      for (const req of batch) {
        expect(Object.keys(req)).toEqual(["updateTextStyle"]);
        // The mask is exactly the two analytics attributes, nothing more.
        expect("updateTextStyle" in req && req.updateTextStyle.fields).toBe("foregroundColor,fontSize");
      }
    }
  });

  it("is idempotent (003-S5): a second analytic-ify re-applies the SAME write as a server no-op", async () => {
    const fake = plainFake();
    await analyticify(fake, new Set([0]));
    const firstView = charView(await view(fake));

    const second = await analyticify(fake, new Set([0]));
    expect(second).toEqual({ paragraphsStyled: 1 }); // still reports the styled line
    // Byte-equal: re-styling already-navy text changes nothing the user sees.
    expect(charView(await view(fake))).toBe(firstView);
  });

  it("an empty ordinal set is a graceful no-op: zero paragraphs, nothing written", async () => {
    const fake = plainFake();
    const r = await analyticify(fake, new Set());
    expect(r).toEqual({ paragraphsStyled: 0 });
    expect(fake.appliedBatches).toHaveLength(0);
  });

  it("gates: refuses a multi-tab / suggestion / hidden doc, each its mapped error, nothing written", async () => {
    // Multi-tab (universal tab gate, A3).
    const multi = new FakeDocs([{ elements: [{ text: "x" }] }], { tabCount: 2 });
    await expectRejects(analyticify(multi, new Set([0])), MultiTabError);
    expect(multi.appliedBatches).toHaveLength(0);

    // Suggestions present (indexes untrustworthy — styles-lane gate).
    const suggested = new FakeDocs([{ elements: [{ text: "x" }] }], { suggestions: true });
    await expectRejects(analyticify(suggested, new Set([0])), SuggestionsActiveError);
    expect(suggested.appliedBatches).toHaveLength(0);

    // Hidden state present: analytic-ify is run-level size, refuses an armed doc.
    const armed = mixedFake();
    await hide(armed);
    const writesAfterHide = armed.appliedBatches.length;
    await expectRejects(analyticify(armed, new Set([0])), HiddenStateError);
    expect(armed.appliedBatches).toHaveLength(writesAfterHide);
  });

  it("a torn multi-chunk analytic-ify -> PartialApplyError('analyticify') mapped to the SAFE-TO-REPEAT copy", async () => {
    // CHUNK_MAX 5000 groups (one per touched paragraph) forces two chunks.
    const specs: GpSpec[] = [];
    for (let i = 0; i < 5001; i++) specs.push({ elements: [{ text: `line ${i}` }] });
    const fake = new FakeDocs(specs);
    fake.beforeApply = (call) => {
      if (call === 1) fake.injectForeignEdit(); // the edit lands before chunk 1
    };
    const err = await expectRejects(
      analyticify(fake, new Set(specs.map((_s, i) => i))),
      PartialApplyError
    );
    expect(err.verb).toBe("analyticify");
    expect(err.appliedChunks).toBe(1);
    expect(err.totalChunks).toBe(2);
    // The copy is analytic-ify's own (idempotent, safe to repeat), never unknown.
    expect(errorMessage(err)).toEqual(errorMessage(new PartialApplyError("analyticify", 1, 2)));
    expect(errorMessage(err).body).toBe(STRINGS.errors.partialApply.analyticify.body);
  });

  it("analytic-ify -> Hide round trip: the freshly-navy run is KEPT at full size, never shrunk (003-S1+S2)", async () => {
    // The core Loop-003 user promise, end to end and write-driven: a plain body
    // line the user analytic-ifies must survive the very next Hide at FULL size
    // (the analytics navy is an explicit "keep this" mark, like a highlight), not
    // shrink to the 1pt sentinel. This is only non-vacuous because the Wave-1
    // fake now models the foreground WRITE: analytic-ify's navy lands in the
    // model, reads back through parse, and keepers.isAnalytics (color-only,
    // size-independent) keeps it — so Hide leaves it visible.
    //
    // P0 plain body, P1 plain body. We analytic-ify ONLY P0. The control is P1:
    // an identical plain body line with NO analytics MUST hide to the sentinel on
    // the same pass, proving it is the analytics mark — not some blanket "Hide
    // skips this doc" — that keeps P0 at full size. (A plain NORMAL_TEXT line is
    // neither a heading nor a cite, so it hides by default; verified below.)
    const fake = plainFake();
    const styled = await analyticify(fake, new Set([0]));
    expect(styled).toEqual({ paragraphsStyled: 1 });

    // After analytic-ify, P0 is navy 14pt; P1 is untouched (no size, no navy).
    const armedView = await view(fake);
    expect(armedView.paragraphs[0].elements[0].fontSizePt).toBe(ANALYTICS_PT);
    expect(armedView.paragraphs[0].elements[0].foregroundHex).toBe(ANALYTICS_FG_HEX);
    expect(armedView.paragraphs[1].elements[0].foregroundHex).toBeNull();

    // Hide runs the reconcile over the analytic-ified doc.
    const r = await hide(fake);
    // Exactly one paragraph changed (P1 hidden); P0's analytics line is kept and
    // emits no hide region, so regionsHidden counts only P1's passage.
    expect(r.paragraphsScanned).toBe(2);
    expect(r.paragraphsChanged).toBe(1);
    expect(r.regionsHidden).toBe(1);

    const hidden = await view(fake);
    // THE PROMISE: every char of the navy P0 line stays at its written 14pt —
    // NOT one is at the sentinel. (P0 is non-final, so its trailing newline is
    // part of the kept run too; nothing in P0 may carry a sentinel size.)
    const navyChars = hidden.paragraphs[0].elements.flatMap((el) =>
      el.kind === "text" ? [...el.text].map(() => el.fontSizePt) : []
    );
    expect(navyChars.every((size) => size === ANALYTICS_PT)).toBe(true);
    expect(navyChars).not.toContain(SENTINEL_PT);
    // The navy survived in color too — Hide's fontSize-only shrink never touches
    // foreground, and the keeper kept the whole run.
    expect(hidden.paragraphs[0].elements[0].foregroundHex).toBe(ANALYTICS_FG_HEX);

    // The CONTROL: the un-analytic-ified P1 line DID hide — its body text now
    // reads at the sentinel. "second line" is 11 chars; P1 is the LAST paragraph
    // so its trailing newline is the segment-final newline, which the clamp
    // leaves un-shrunk, so the 11 body chars (not the newline) sit at 1pt.
    expect(sentinelChars(hidden)).toBe(11);
    expect(rstmCount(hidden)).toBe(1); // one rstm anchor for P1's hidden region
  });
});

describe("Delete analytics (controller)", () => {
  const navy = (text: string): { text: string; fg: string } => ({ text, fg: ANALYTICS_FG_HEX });

  /** A mixed analytics doc the gates all pass:
   *   P0 [1,11)  wholly-navy "analytics\n"      -> whole-line delete (collapses)
   *   P1 [11,30) plain "keep " + navy "navy" + plain " tail\n"  -> partial path
   *   P2 [30,?)  plain "plain only\n"            -> no analytics, no range
   * P0 is non-final so its newline goes too; the line collapses. */
  const mixedAnalyticsFake = (): FakeDocs =>
    new FakeDocs([
      { elements: [navy("analytics")] },
      { elements: [{ text: "keep " }, navy("navy"), { text: " tail" }] },
      { elements: [{ text: "plain only" }] }
    ]);

  it("removes exactly the analytics text, collapses the wholly-navy line, keeps everything else", async () => {
    const fake = mixedAnalyticsFake();
    const r = await deleteAnalytics(fake);
    // Two affected paragraphs (P0 whole-line + P1 partial); P2 untouched.
    expect(r.paragraphsAffected).toBe(2);
    expect(r.runsDeleted).toBe(2);

    const after = await view(fake);
    // P0 collapsed away: the doc now opens with the former P1 (partial) line,
    // its navy run gone, plain text fused.
    expect(after.paragraphs).toHaveLength(2);
    const flat = after.paragraphs.map((p) => p.elements.map((e) => e.text).join(""));
    expect(flat[0]).toBe("keep  tail\n"); // "navy" excised, neighbors retained
    expect(flat[1]).toBe("plain only\n");
    // No analytics survives anywhere.
    for (const p of after.paragraphs) {
      for (const e of p.elements) expect(e.foregroundHex).not.toBe(ANALYTICS_FG_HEX);
    }
  });

  it("(003-F1 audit) emits ONLY deleteContentRange — no style/named-range write", async () => {
    const fake = mixedAnalyticsFake();
    await deleteAnalytics(fake);
    expect(fake.appliedBatches.length).toBeGreaterThan(0);
    for (const batch of fake.appliedBatches) {
      for (const req of batch) expect(Object.keys(req)).toEqual(["deleteContentRange"]);
    }
  });

  it("zero analytics is a graceful no-op: nothing planned, nothing written", async () => {
    const fake = new FakeDocs([{ elements: [{ text: "no analytics here" }] }]);
    const r = await deleteAnalytics(fake);
    expect(r).toEqual({ paragraphsAffected: 0, runsDeleted: 0 });
    expect(fake.appliedBatches).toHaveLength(0);
  });

  it("a final-segment wholly-navy line clamps off its newline (empty line remains, not collapsed)", async () => {
    // Single paragraph, so it IS the final segment: the API refuses to delete
    // the doc-final newline, so the clamp leaves [start, end-1) and an empty
    // line survives (documented, unavoidable).
    const fake = new FakeDocs([{ elements: [navy("solo")] }]);
    const r = await deleteAnalytics(fake);
    expect(r).toEqual({ paragraphsAffected: 1, runsDeleted: 1 });

    const after = await view(fake);
    expect(after.paragraphs).toHaveLength(1); // line NOT collapsed
    // Only the newline is left — the navy text is gone.
    const remaining = after.paragraphs[0].elements.map((e) => e.text).join("");
    expect(remaining).toBe("\n");
  });

  it("gates: refuses a multi-tab / suggestion / hidden doc, each its mapped error, nothing written", async () => {
    const multi = new FakeDocs([{ elements: [navy("a")] }], { tabCount: 2 });
    await expectRejects(deleteAnalytics(multi), MultiTabError);
    expect(multi.appliedBatches).toHaveLength(0);

    const suggested = new FakeDocs([{ elements: [navy("a")] }], { suggestions: true });
    await expectRejects(deleteAnalytics(suggested), SuggestionsActiveError);
    expect(suggested.appliedBatches).toHaveLength(0);

    // Hidden state: deleting analytics out of an armed doc would desync the RLE
    // restore records — refuse, exactly like analytic-ify / Apply-styles.
    const armed = mixedFake();
    await hide(armed);
    const writesAfterHide = armed.appliedBatches.length;
    await expectRejects(deleteAnalytics(armed), HiddenStateError);
    expect(armed.appliedBatches).toHaveLength(writesAfterHide);
  });

  it("descending emission is real: the multi-paragraph delete removes exactly analytics through the shifting model", async () => {
    // Four NON-adjacent wholly-navy lines (plain spacers between them keep the
    // ranges from coalescing), so the planner emits separate descending deletes.
    // If the controller/planner emitted ASCENDING, the shifting model would
    // delete the wrong chars — this asserts the surviving text is exact. A
    // trailing plain spacer is the final segment, so all four navy lines are
    // INTERIOR and collapse cleanly (newline included) — no clamped empty line
    // muddies the assertion.
    const fake = new FakeDocs([
      { elements: [navy("aaa")] },
      { elements: [{ text: "spacer one" }] },
      { elements: [navy("bbb")] },
      { elements: [{ text: "spacer two" }] },
      { elements: [navy("ccc")] },
      { elements: [{ text: "tail spacer" }] }
    ]);
    const r = await deleteAnalytics(fake);
    expect(r.paragraphsAffected).toBe(3);

    const after = await view(fake);
    const flat = after.paragraphs.map((p) => p.elements.map((e) => e.text).join(""));
    // The three navy lines collapsed; only the spacers remain, in order.
    expect(flat).toEqual(["spacer one\n", "spacer two\n", "tail spacer\n"]);
  });

  it("a torn multi-chunk delete -> PartialApplyError('deleteAnalytics') mapped to the DELETE-SPECIFIC copy (not unknown)", async () => {
    // 5001 partial-analytics paragraphs (navy run flanked by plain text so the
    // ranges never coalesce) -> 5001 delete groups -> two chunks [5000, 1].
    const specs: GpSpec[] = [];
    for (let i = 0; i < 5001; i++) specs.push({ elements: [{ text: `k${i} ` }, navy("x"), { text: " t" }] });
    const fake = new FakeDocs(specs);
    fake.beforeApply = (call) => {
      if (call === 1) fake.injectForeignEdit(); // the edit lands before chunk 1
    };

    const err = await expectRejects(deleteAnalytics(fake), PartialApplyError);
    expect(err.verb).toBe("deleteAnalytics");
    expect(err.appliedChunks).toBe(1);
    expect(err.totalChunks).toBe(2);

    // The mapped copy is the delete-specific partial body — it must NOT fall
    // through to errors.unknown (an unmapped delete failure would lie about
    // whether content was removed) and must differ from the analyticify copy.
    const mapped = errorMessage(err);
    expect(mapped.body).toBe(STRINGS.errors.partialApply.deleteAnalytics.body);
    expect(mapped.body).not.toBe(STRINGS.errors.unknown.body);
    expect(mapped.body).not.toBe(STRINGS.errors.partialApply.analyticify.body);

    // The torn state is real: the highest-index chunk's deletes landed, so SOME
    // analytics is already gone — exactly the "Show All will not bring it back"
    // scenario the copy warns about.
    fake.beforeApply = null;
    const torn = await view(fake);
    let navyRuns = 0;
    for (const p of torn.paragraphs) for (const e of p.elements) if (e.foregroundHex === ANALYTICS_FG_HEX) navyRuns++;
    expect(navyRuns).toBeLessThan(5001);
    expect(navyRuns).toBeGreaterThan(0);
  });

  it("a NON-REVISION DocsApiError on a non-first delete chunk -> PartialApplyError, NOT errors.docsApi's 'nothing applied' lie", async () => {
    // The exec-review data-loss BLOCKER (distinct from the revision-race tear
    // above): a transient/non-revision Docs rejection (500/429, or a generic
    // 400) on chunk 2+ must NOT surface "nothing was applied" while chunk 1's
    // deletes ALREADY destroyed content. runVerb now wraps ANY applied>0
    // mid-sequence failure as PartialApplyError, so the copy stays truthful.
    // Same 5001-paragraph two-chunk shape; the SECOND chunk is rejected with a
    // DocsApiError (not a foreign edit).
    const specs: GpSpec[] = [];
    for (let i = 0; i < 5001; i++) specs.push({ elements: [{ text: `k${i} ` }, navy("x"), { text: " t" }] });
    const fake = new FakeDocs(specs);
    fake.beforeApply = (call) => {
      // Arm the rejection only for the 2nd chunk; chunk 1 (the highest indexes,
      // descending order) lands first and really removes ~5000 runs.
      if (call === 1) fake.rejectRequestType("deleteContentRange");
    };

    const err = await expectRejects(deleteAnalytics(fake), PartialApplyError);
    expect(err.verb).toBe("deleteAnalytics");
    expect(err.appliedChunks).toBe(1);
    expect(err.totalChunks).toBe(2);

    // Truthful delete-specific copy — must NEVER be errors.docsApi ("Docs
    // rejected the change — nothing was applied"), which would be a data-loss
    // lie now that chunk 1's deletes committed.
    const mapped = errorMessage(err);
    expect(mapped.body).toBe(STRINGS.errors.partialApply.deleteAnalytics.body);
    expect(mapped.body).not.toBe(STRINGS.errors.docsApi.body);
    expect(mapped.body).not.toBe(STRINGS.errors.unknown.body);

    // Content really was removed (the "nothing applied" lie would have denied it).
    fake.beforeApply = null;
    fake.rejectRequestType(null);
    const torn = await view(fake);
    let navyRuns = 0;
    for (const p of torn.paragraphs) for (const e of p.elements) if (e.foregroundHex === ANALYTICS_FG_HEX) navyRuns++;
    expect(navyRuns).toBeLessThan(5001);
    expect(navyRuns).toBeGreaterThan(0);
  });

  it("TOCTOU: analytics vanished between the confirm-count read and the verb fetch -> graceful no-op", async () => {
    // The adapter counts analytics paragraphs and fires the confirm against ONE
    // snapshot; the verb then takes its OWN fetch (plan §3 docsAdapter TOCTOU).
    // If the analytics vanished in between — the user hand-deleted/recolored it,
    // or a teammate did — the verb's fetched view carries zero analytics, so it
    // plans zero ranges, applies nothing, and returns the no-op result the
    // receipt renders gracefully (never a crash, never a phantom count). The
    // race is modeled at the only seam the verb can observe: the doc it fetches.
    const confirmCount = await countNavyParagraphs(mixedAnalyticsFake());
    expect(confirmCount).toBeGreaterThan(0); // the confirm genuinely fired

    // The doc the verb actually fetches no longer has any analytics.
    const fake = new FakeDocs([
      { elements: [{ text: "analytics" }] },
      { elements: [{ text: "keep  tail" }] },
      { elements: [{ text: "plain only" }] }
    ]);
    const r = await deleteAnalytics(fake);
    expect(r).toEqual({ paragraphsAffected: 0, runsDeleted: 0 });
    expect(fake.appliedBatches).toHaveLength(0);
  });
});

/** Count paragraphs carrying >= 1 analytics run in a fake's CURRENT view — the
 * test-side analog of the adapter's confirm-count, used only to prove the
 * TOCTOU snapshot genuinely saw analytics before the race removed it. */
async function countNavyParagraphs(fake: FakeDocs): Promise<number> {
  const doc = await view(fake);
  return doc.paragraphs.filter((p) => p.elements.some((e) => e.foregroundHex === ANALYTICS_FG_HEX)).length;
}
