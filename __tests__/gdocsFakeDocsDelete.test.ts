// Low-level harness suite for the fakeDocs.ts deleteContentRange model (plan
// 003 Wave A, the BLOCKING harness deliverable). Delete-analytics is the
// engine's SOLE content-deleter, and its planner emits deleteContentRange
// ranges in strictly DESCENDING start order because a real batchUpdate applies
// its requests SEQUENTIALLY — a lower-index delete done first shifts every
// higher index after it. The delete planner / controller tests are only as
// trustworthy as the fake they run against, so this suite proves the fake's
// delete model directly, against HAND-BUILT requests (no planner in the loop):
//
//   * a single deleteContentRange splices exactly its chars out and shifts
//     every downstream run + paragraph index (indexes are derived, so the
//     post-delete read renumbers with no gaps);
//   * deleting a paragraph's whole range (incl. its trailing newline) collapses
//     the line — the paragraph disappears and the next one slides up;
//   * THE DIFFERENTIAL: the SAME two-range plan applied DESCENDING removes
//     exactly the targeted chars, while applied ASCENDING corrupts the doc
//     (the first low delete shifts the indexes out from under the second) —
//     proving descending order is LOAD-BEARING, not cosmetic;
//   * validate() rejects the unremovable segment-final newline (mirrors the
//     styleCeiling rule), out-of-range, zero-length, and start >= end ranges,
//     leaving the model byte-identical (the atomic-batch contract).
//
// Reads go through the PRODUCTION parser (parse.ts), the same fixture-realism
// rail the flagship uses: a delete-model bug surfaces as a wrong PARSED view,
// not a wrong internal field.

import { encodeRgbColor } from "../google-docs/src/core/color";
import { ANALYTICS_FG_HEX } from "../google-docs/src/core/constants";
import { parseDocument } from "../google-docs/src/core/parse";
import {
  DeleteContentRangeRequest,
  DocsApiError,
  GDoc,
  RevisionMismatchError,
  UpdateTextStyleRequest
} from "../google-docs/src/core/types";
import { FakeDocs } from "./fakeDocs";
import { para } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** The production read path: fetch the fake and decode with parse.ts. */
async function view(fake: FakeDocs): Promise<GDoc> {
  return parseDocument(await fake.fetchDocument());
}

/** A deleteContentRange request over [start, end). */
function del(start: number, end: number): DeleteContentRangeRequest {
  return { deleteContentRange: { range: { startIndex: start, endIndex: end } } };
}

/**
 * The full body text as the parser sees it (newlines shown as "¶"), with each
 * paragraph's [start, end) range — the projection a splice bug shows up in.
 * Two docs are equal here iff a user would read identical text at identical
 * indexes, which is exactly what deleteContentRange must guarantee.
 */
function bodyView(doc: GDoc): string {
  return doc.paragraphs
    .map((p) => {
      const text = p.elements.map((el) => (el.kind === "other" ? `␂` : el.text)).join("");
      return `[${p.startIndex},${p.endIndex})${text.replace(/\n/g, "¶")}`;
    })
    .join("|");
}

/** A paragraph's full text, concatenated across its runs. deleteContentRange
 * (like updateTextStyle) legitimately FRAGMENTS runs at the cut boundaries, so
 * the user-visible promise is the concatenated text, never a single run's
 * payload — asserting per-element would fail on harmless fragmentation. */
function paraText(doc: GDoc, index: number): string {
  return doc.paragraphs[index].elements.map((e) => e.text).join("");
}

/** Apply one batch at the fake's current revision (the only mutation path).
 * Each successful batch bumps the revision, so chained batches must re-read. */
async function applyAt(fake: FakeDocs, requests: DeleteContentRangeRequest[]): Promise<void> {
  const rev = (await view(fake)).revisionId;
  await fake.applyBatch(requests, rev);
}

// ---------------------------------------------------------------------------
// (1) Single delete: splice + downstream index shift
// ---------------------------------------------------------------------------

describe("deleteContentRange splices chars and shifts downstream indexes", () => {
  // One paragraph, "ABCDEF\n" — body indexes [1,8): A=1 .. F=6, newline=7.
  it("removes exactly the targeted middle chars and renumbers the tail", async () => {
    const fake = new FakeDocs([para("ABCDEF")]);
    // Delete "CD" at [3,5). Survivors: "AB" + "EF\n" => "ABEF\n" over [1,6).
    await applyAt(fake, [del(3, 5)]);

    const doc = await view(fake);
    expect(bodyView(doc)).toBe("[1,6)ABEF¶");
    expect(paraText(doc, 0)).toBe("ABEF\n");
    // The whole tail shifted left by the 2 deleted chars: the paragraph that
    // ended at 8 now ends at 6, and the (possibly fragmented) runs tile the
    // range CONTIGUOUSLY with no gap — the index shift the planner depends on.
    const els = doc.paragraphs[0].elements;
    expect(els[0].startIndex).toBe(1);
    expect(els[els.length - 1].endIndex).toBe(6);
    for (let i = 1; i < els.length; i++) expect(els[i].startIndex).toBe(els[i - 1].endIndex);
  });

  // Two paragraphs: "ABC\n" [1,5) then "DEF\n" [5,9). Deleting inside P0 must
  // slide P1 (and its indexes) left by the deleted width — the cross-paragraph
  // shift the planner's descending order exists to keep coherent.
  it("shifts a LATER paragraph's indexes when an earlier paragraph shrinks", async () => {
    const fake = new FakeDocs([para("ABC"), para("DEF")]);
    expect(bodyView(await view(fake))).toBe("[1,5)ABC¶|[5,9)DEF¶");

    await applyAt(fake, [del(2, 3)]); // delete "B" from P0
    // P0 -> "AC\n" [1,4); P1 unchanged content but slid to [4,8).
    expect(bodyView(await view(fake))).toBe("[1,4)AC¶|[4,8)DEF¶");
  });

  // Deleting a whole paragraph's range INCLUDING its newline collapses the
  // line: the paragraph vanishes and the next one slides up to take its start.
  it("collapses a line when its entire range (incl. newline) is deleted", async () => {
    const fake = new FakeDocs([para("keep"), para("gone"), para("tail")]);
    // P1 "gone\n" occupies [6,11). Delete it whole.
    await applyAt(fake, [del(6, 11)]);

    const doc = await view(fake);
    expect(doc.paragraphs).toHaveLength(2);
    expect(bodyView(doc)).toBe("[1,6)keep¶|[6,11)tail¶");
  });
});

// ---------------------------------------------------------------------------
// (2) THE DIFFERENTIAL: descending removes exactly; ascending corrupts
// ---------------------------------------------------------------------------

describe("delete ORDER is load-bearing within a batch (descending vs ascending)", () => {
  // One paragraph "0123456789\n" over [1,12): digit d sits at index d+1, the
  // newline at 11. We target two disjoint single-char ranges that a planner
  // would have produced as "analytics" runs: "2" at [3,4) and "7" at [8,9).
  // The CORRECT result removes only "2" and "7" -> "01345689\n".
  const SOURCE = "0123456789";
  const CORRECT = "01345689";
  const lowRange = del(3, 4); // the "2"
  const highRange = del(8, 9); // the "7"

  it("DESCENDING (high then low) removes exactly the two targeted chars", async () => {
    const fake = new FakeDocs([para(SOURCE)]);
    // The order the planner emits: strictly descending by startIndex.
    await applyAt(fake, [highRange, lowRange]);

    const doc = await view(fake);
    expect(paraText(doc, 0)).toBe(`${CORRECT}\n`);
    expect(bodyView(doc)).toBe(`[1,${CORRECT.length + 2})${CORRECT}¶`);
  });

  it("ASCENDING (low then high) corrupts — the wrong second char is removed", async () => {
    const fake = new FakeDocs([para(SOURCE)]);
    // Same two ranges, WRONG order. After the low delete removes "2", every
    // index >= 4 slides left by one, so the still-original [8,9) range now
    // straddles a shifted char: the second delete removes "8" (the digit that
    // slid into index 8), not "7". Result diverges from CORRECT.
    await applyAt(fake, [lowRange, highRange]);

    const corrupted = paraText(await view(fake), 0);
    // It removed "2" (correct) and then "8" (WRONG — after "2" left, the chars
    // slid down one, so the still-original index 8 now points at "8" rather
    // than the intended "7"). "7" wrongly survives, "8" is wrongly gone.
    expect(corrupted).toBe("01345679\n");
    // And it is DIFFERENT from the correct descending result — the whole point.
    expect(corrupted).not.toBe(`${CORRECT}\n`);
  });

  // The cross-paragraph form of the same hazard: two whole-line collapses in
  // one batch. Descending collapses both targeted lines; the surviving lines
  // are exactly the untargeted ones.
  it("DESCENDING multi-paragraph collapse removes exactly the targeted lines", async () => {
    // P0 "a\n"[1,3) P1 "b\n"[3,5) P2 "c\n"[5,7) P3 "d\n"[7,9). Collapse P0 and
    // P2 (both NON-final; P3's newline is the unremovable segment-final one).
    const fake = new FakeDocs([para("a"), para("b"), para("c"), para("d")]);
    await applyAt(fake, [del(5, 7), del(1, 3)]); // descending: P2 then P0

    const doc = await view(fake);
    expect(doc.paragraphs.map((p) => p.elements.map((e) => e.text).join(""))).toEqual(["b\n", "d\n"]);
  });
});

// ---------------------------------------------------------------------------
// (3) validate() failure paths — every rejection leaves the model untouched
// ---------------------------------------------------------------------------

describe("deleteContentRange validation rejects illegal ranges (atomic, untouched)", () => {
  // "ABC\n" over [1,5): A=1 B=2 C=3, the segment-final newline at 4. docEnd=5,
  // styleCeiling=4, so deleting [4,5) would remove the unremovable final
  // newline and must be refused (mirrors the styleCeiling/clamp rule).
  function singleParaFake(): FakeDocs {
    return new FakeDocs([para("ABC")]);
  }

  it("refuses deleting the segment-final newline and leaves the doc byte-identical", async () => {
    const fake = singleParaFake();
    const before = bodyView(await view(fake));
    const rev = (await view(fake)).revisionId;

    await expect(fake.applyBatch([del(4, 5)], rev)).rejects.toThrow(DocsApiError);
    // Atomicity: nothing mutated, so the very same revision still applies a
    // legal delete (a stale revision would prove a phantom mutation bumped rev).
    expect(bodyView(await view(fake))).toBe(before);
    await expect(fake.applyBatch([del(2, 3)], rev)).resolves.toBeDefined();
  });

  it("refuses an out-of-range endIndex beyond docEnd", async () => {
    const fake = singleParaFake();
    const rev = (await view(fake)).revisionId;
    // endIndex 6 > docEnd 5.
    await expect(fake.applyBatch([del(3, 6)], rev)).rejects.toThrow(/invalid range/);
  });

  it("refuses a startIndex below the body floor of 1", async () => {
    const fake = singleParaFake();
    const rev = (await view(fake)).revisionId;
    await expect(fake.applyBatch([del(0, 2)], rev)).rejects.toThrow(/invalid range/);
  });

  it("refuses a zero-length range (start === end)", async () => {
    const fake = singleParaFake();
    const rev = (await view(fake)).revisionId;
    await expect(fake.applyBatch([del(2, 2)], rev)).rejects.toThrow(/invalid range/);
  });

  it("refuses an inverted range (start > end)", async () => {
    const fake = singleParaFake();
    const rev = (await view(fake)).revisionId;
    await expect(fake.applyBatch([del(3, 2)], rev)).rejects.toThrow(/invalid range/);
  });

  // A rejected batch validates EVERYTHING before mutating, so a legal delete
  // bundled with an illegal one must not partially apply.
  it("rejects the whole batch when ANY range is illegal (no partial splice)", async () => {
    const fake = singleParaFake();
    const before = bodyView(await view(fake));
    const rev = (await view(fake)).revisionId;
    // First range legal, second illegal (final newline) — the batch must abort
    // with NOTHING applied.
    await expect(fake.applyBatch([del(2, 3), del(4, 5)], rev)).rejects.toThrow(DocsApiError);
    expect(bodyView(await view(fake))).toBe(before);
    // The original revision is still current — proof nothing mutated.
    await expect(fake.applyBatch([del(2, 3)], rev)).resolves.toBeDefined();
  });

  // A stale revision refuses BEFORE any range validation, the writeControl
  // ordering the controller's revision chain relies on (here exercised for the
  // delete request type specifically).
  it("refuses a stale revision before touching the ranges", async () => {
    const fake = singleParaFake();
    await expect(fake.applyBatch([del(2, 3)], "r999")).rejects.toThrow(RevisionMismatchError);
  });
});

// ---------------------------------------------------------------------------
// (4) Table cells: a buildDoc spec with inTable parses to inTable paragraphs
//     AND deleteContentRange splices WITHIN a cell, shifting the tail like body
// ---------------------------------------------------------------------------

describe("deleteContentRange splices within a table cell and shifts the tail", () => {
  // A single-cell table holding "CELL" followed by a body paragraph "tail".
  // The harness folds the inTable paragraph into one table structural element,
  // so the PRODUCTION parser must (a) recurse into the cell and mark the
  // paragraph inTable, and (b) keep its index contiguous with the body that
  // follows it — the flat-index promise a cell splice depends on.
  function tableThenBody(): FakeDocs {
    return new FakeDocs([{ inTable: true, elements: [{ text: "CELL" }] }, para("tail")]);
  }

  it("parses a cell paragraph as inTable with body-contiguous indexes", async () => {
    const doc = await view(tableThenBody());
    // "CELL\n" [1,6) inside the cell; "tail\n" [6,11) in the body right after.
    // Asserting the EXACT index layout (not just the flag) is what makes the
    // contiguity claim falsifiable — a structural-slot model would shift these.
    expect(doc.paragraphs.map((p) => [p.inTable, p.startIndex, p.endIndex])).toEqual([
      [true, 1, 6],
      [false, 6, 11]
    ]);
    expect(bodyView(doc)).toBe("[1,6)CELL¶|[6,11)tail¶");
  });

  it("splices chars out of the cell and slides the body paragraph left", async () => {
    const fake = tableThenBody();
    // "CELL\n": C=1 E=2 L=3 L=4 newline=5. Delete the "LL" at [3,5).
    await applyAt(fake, [del(3, 5)]);

    const doc = await view(fake);
    // The cell paragraph keeps its structural newline -> "CE\n" over [1,4); it
    // is STILL inTable (the splice did not dissolve the table), and the body
    // "tail\n" slid left by the 2 deleted chars to [4,9) — the cross-structure
    // index shift the planner's descending order keeps coherent.
    expect(doc.paragraphs.map((p) => [p.inTable, p.startIndex, p.endIndex])).toEqual([
      [true, 1, 4],
      [false, 4, 9]
    ]);
    expect(paraText(doc, 0)).toBe("CE\n");
    expect(paraText(doc, 1)).toBe("tail\n");
  });
});

// ---------------------------------------------------------------------------
// (5) updateTextStyle foregroundColor: an analytic-ify navy write reads back as
//     analytics through the production parser (otherwise it was silently dropped)
// ---------------------------------------------------------------------------

describe("updateTextStyle honors the foregroundColor field mask", () => {
  /** An updateTextStyle request writing the analytics navy over [start, end)
   * with the EXACT wire shape styles.ts emits (color.encodeRgbColor + the
   * "foregroundColor" mask) — so this exercises the same encode/decode pair the
   * analytic-ify verb relies on, not a test-only color literal. */
  function navyWrite(start: number, end: number): UpdateTextStyleRequest {
    return {
      updateTextStyle: {
        range: { startIndex: start, endIndex: end },
        textStyle: { foregroundColor: encodeRgbColor(ANALYTICS_FG_HEX) },
        fields: "foregroundColor"
      }
    };
  }

  it("makes the navy foreground visible to a later fetch+parse", async () => {
    const fake = new FakeDocs([para("ABCDEF")]);
    // Precondition the assertion can actually fail against: the run starts with
    // NO foreground (a dropped write would leave it here, passing a weaker test).
    expect((await view(fake)).paragraphs[0].elements.every((el) => el.foregroundHex === null)).toBe(true);

    // Write navy over "BCDE" at [2,6); "A" before and "F\n" after stay inherited.
    const rev = (await view(fake)).revisionId;
    await fake.applyBatch([navyWrite(2, 6)], rev);

    const doc = await view(fake);
    // Reconstruct the per-character foreground from the (possibly fragmented)
    // runs: only the covered span carries the navy, proving the mask was honored
    // at the right boundaries and the OptionalColor byte round-tripped exactly.
    const fgByChar: (string | null)[] = [];
    for (const el of doc.paragraphs[0].elements) for (const _ of el.text) fgByChar.push(el.foregroundHex);
    const N = ANALYTICS_FG_HEX;
    expect(fgByChar).toEqual([null, N, N, N, N, null, null]); // A B C D E F \n
  });

  it("CLEARS the foreground to inherited when the mask names it but the style omits it", async () => {
    // Seed an already-navy run (fg set via the spec), then write an EMPTY
    // foreground style with the field still in the mask — the documented
    // clear-to-inherit contract, identical to fontSize/bold. A fake that only
    // SET colors (never cleared) would wrongly leave the navy here.
    const fake = new FakeDocs([para("X")]);
    // First paint it navy so there is something to clear.
    let rev = (await view(fake)).revisionId;
    await fake.applyBatch([navyWrite(1, 2)], rev);
    expect((await view(fake)).paragraphs[0].elements[0].foregroundHex).toBe(ANALYTICS_FG_HEX);

    // Now clear: mask names foregroundColor, textStyle omits it -> null.
    rev = (await view(fake)).revisionId;
    await fake.applyBatch(
      [{ updateTextStyle: { range: { startIndex: 1, endIndex: 2 }, textStyle: {}, fields: "foregroundColor" } }],
      rev
    );
    expect((await view(fake)).paragraphs[0].elements[0].foregroundHex).toBeNull();
  });
});
