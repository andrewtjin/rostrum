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

import { MAX_REPLAN_ATTEMPTS, SENTINELS } from "../gdocs/src/core/constants";
import { applyStyles, hide, showAll } from "../gdocs/src/core/controller";
import { parseDocument } from "../gdocs/src/core/parse";
import { isRstmName } from "../gdocs/src/core/rangeNames";
import {
  DocsApiError,
  GDoc,
  HiddenStateError,
  MultiTabError,
  PartialApplyError,
  RevisionConflictError,
  SuggestionsActiveError
} from "../gdocs/src/core/types";
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
