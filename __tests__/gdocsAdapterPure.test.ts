// gdocs adapter-pure suite (plan A4/A11.i, step S11). adapterPure.ts is the
// adapter's entire decision surface — error classification, dialog routing,
// the call map, sidebar state, Mark-cite planning from selection picks, and
// Diagnostics interpretation — extracted to core precisely so it sits under
// root tsc AND the gdocs coverage floor. The host adapter keeps only Google
// calls, so what passes here is everything the wet round does not have to
// re-prove about decisions.
//
// Raw inputs are WIRE-SHAPED documents.get JSON (not pre-parsed GDocs):
// adapterPure's public surface consumes raw reads end to end, so the tests
// exercise the same parseDocument seam the live adapter does.

import {
  BG_HISTOGRAM_TOP_N,
  buildDiagnostics,
  CALL_MAP,
  classifyBatchError,
  countDebateHeadings,
  diagnosticsReadFacts,
  DiagnosticsModel,
  FONT_FLOOR_TRY_SIZES_PT,
  hideDialog,
  interpretFontFloor,
  markCiteDialog,
  planMarkCiteFromPicks,
  probeReadbackSizePt,
  renderDiagnosticsText,
  routeError,
  SelectionPick,
  SETTINGS_PROPERTY_KEY,
  showAllDialog,
  sidebarState,
  SMALL_DOC_DUMP_MAX_BYTES,
  STATE_FIELDS_MASK,
  stylesDialog,
  sumApiUnits,
  textPickOffsets
} from "../gdocs/src/core/adapterPure";
import { CITE_PT, GDOCS_VERSION } from "../gdocs/src/core/constants";
import { errorMessage, markCiteReceipt, STRINGS } from "../gdocs/src/core/strings";
import {
  DocsApiError,
  GHideResult,
  GShowAllResult,
  GStylesResult,
  HiddenStateError,
  MultiTabError,
  PartialApplyError,
  RevisionMismatchError,
  SuggestionsActiveError,
  UpdateTextStyleRequest
} from "../gdocs/src/core/types";

// ---------------------------------------------------------------------------
// Wire-shaped raw builders (the documents.get JSON parse.ts decodes). Index
// rules mirror __tests__/gdocsBuilders.ts: body starts at 1; the trailing
// newline lives inside the final text run (or as its own run after a chip).
// ---------------------------------------------------------------------------

interface RawElSpec {
  text: string;
  sizePt?: number;
  bold?: boolean;
  /** Lower-case "#rrggbb" — encoded as the API's rgbColor floats (zero
   * channels omitted, the wire's omitted-means-zero reality). */
  bgHex?: string;
  /** "other" models a chip/object: occupies text.length index units. */
  kind?: "text" | "other";
  /** Plant a suggested* key on this run (trips the suggestion gate). */
  suggested?: boolean;
}

interface RawParaSpec {
  style?: string;
  /** Wrapped in its own single-cell table (parse flags it inTable). */
  inTable?: boolean;
  elements: RawElSpec[];
}

interface RawDocOptions {
  revisionId?: string;
  /** Wire-shaped namedRanges map (name -> group). */
  namedRanges?: Record<string, unknown>;
  tabs?: unknown;
  namedStyles?: unknown;
}

/** rgbColor floats for a hex, zero channels omitted (wire reality). */
function rgbColorOf(hex: string): Record<string, number> {
  const out: Record<string, number> = {};
  const channels: Array<[string, number]> = [
    ["red", parseInt(hex.slice(1, 3), 16)],
    ["green", parseInt(hex.slice(3, 5), 16)],
    ["blue", parseInt(hex.slice(5, 7), 16)]
  ];
  for (const [name, byte] of channels) {
    if (byte !== 0) out[name] = byte / 255;
  }
  return out;
}

/** Build an index-consistent raw documents.get response. */
function rawDoc(paras: RawParaSpec[], opts: RawDocOptions = {}): unknown {
  let cursor = 1; // Docs body content starts at index 1
  const content: unknown[] = [];
  paras.forEach((p, i) => {
    const startIndex = cursor;
    const specs = [...p.elements];
    const last = specs[specs.length - 1];
    // The trailing newline rides inside the final TEXT run; a chip-final (or
    // empty) paragraph gets the newline as its own run — builder parity.
    if (last !== undefined && (last.kind ?? "text") === "text") {
      specs[specs.length - 1] = { ...last, text: last.text + "\n" };
    } else {
      specs.push({ text: "\n" });
    }
    const elements = specs.map((e) => {
      const el: Record<string, unknown> = { startIndex: cursor, endIndex: cursor + e.text.length };
      if ((e.kind ?? "text") === "other") {
        el.inlineObjectElement = {}; // any non-textRun shape = whitelist "other"
      } else {
        const textStyle: Record<string, unknown> = {};
        if (e.sizePt !== undefined) textStyle.fontSize = { magnitude: e.sizePt, unit: "PT" };
        if (e.bold === true) textStyle.bold = true;
        if (e.bgHex !== undefined) textStyle.backgroundColor = { color: { rgbColor: rgbColorOf(e.bgHex) } };
        const textRun: Record<string, unknown> = { content: e.text, textStyle };
        if (e.suggested === true) textRun.suggestedInsertionIds = ["sg.1"];
        el.textRun = textRun;
      }
      cursor += e.text.length;
      return el;
    });
    const paragraphItem = {
      startIndex,
      endIndex: cursor,
      paragraph: { elements, paragraphStyle: { namedStyleType: p.style ?? "NORMAL_TEXT" } }
    };
    // inTable paragraphs ride inside their own one-cell table (collectParagraphs
    // recurses and flags them); index bookkeeping stays monotonic either way.
    content.push(
      p.inTable === true
        ? { startIndex, endIndex: cursor, table: { tableRows: [{ tableCells: [{ content: [paragraphItem] }] }] } }
        : paragraphItem
    );
    void i;
  });
  const doc: Record<string, unknown> = {
    revisionId: opts.revisionId ?? "rev-1",
    body: { content }
  };
  if (opts.namedRanges !== undefined) doc.namedRanges = opts.namedRanges;
  if (opts.tabs !== undefined) doc.tabs = opts.tabs;
  if (opts.namedStyles !== undefined) doc.namedStyles = opts.namedStyles;
  return doc;
}

/** Wire-shaped namedRanges map entry: one range, one segment. */
function rawRange(name: string, startIndex: number, endIndex: number, id = `id-${name}`): Record<string, unknown> {
  return { [name]: { namedRanges: [{ namedRangeId: id, name, ranges: [{ startIndex, endIndex }] }] } };
}

/** Merge several rawRange maps into one namedRanges object. */
function rawRanges(...maps: Record<string, unknown>[]): Record<string, unknown> {
  return Object.assign({}, ...maps) as Record<string, unknown>;
}

// Zeroed result builders (the strings suite's convention) so dialog tests
// state only the field under test.
function hideResult(over: Partial<GHideResult> = {}): GHideResult {
  return {
    paragraphsScanned: 0,
    paragraphsChanged: 0,
    regionsHidden: 0,
    regionsAlreadyHidden: 0,
    newlyKeptRestored: 0,
    preexistingTinyCount: 0,
    ...over
  };
}
function showResult(over: Partial<GShowAllResult> = {}): GShowAllResult {
  return {
    segmentsRestoredExact: 0,
    segmentsNormalized: 0,
    sweptOrphans: 0,
    rangesDeleted: 0,
    rangesSkippedNewerVersion: 0,
    ...over
  };
}
function stylesResult(over: Partial<GStylesResult> = {}): GStylesResult {
  return {
    namedStylesApplied: true,
    restyled: { pocket: 0, hat: 0, block: 0, tag: 0 },
    spacingCleared: 0,
    citesRepaired: 0,
    ...over
  };
}

// ---------------------------------------------------------------------------
// The call map — the build contract's other half (gdocsBuild.test.ts owns the
// shim side; this owns content and menu-order invariants).
// ---------------------------------------------------------------------------

describe("CALL_MAP (plan S12)", () => {
  it("includes onOpen — no onOpen, no menu, no product", () => {
    expect(CALL_MAP.map((e) => e.fn)).toContain("onOpen");
  });

  it("carries unique, identifier-shaped function names (none shadowing the bundle global)", () => {
    const names = CALL_MAP.map((e) => e.fn);
    expect(new Set(names).size).toBe(names.length);
    for (const n of names) {
      expect(n).toMatch(/^[A-Za-z_$][A-Za-z0-9_$]*$/);
      expect(n).not.toBe("Rostrum");
    }
  });

  it("lists the seven menu items with STRINGS labels, in the deck's menu order", () => {
    const labeled = CALL_MAP.filter((e) => e.label !== null);
    expect(labeled.map((e) => e.label)).toEqual([
      STRINGS.menu.hide,
      STRINGS.menu.showAll,
      STRINGS.menu.applyStyles,
      STRINGS.menu.markCite,
      STRINGS.menu.openPanel,
      STRINGS.menu.helpShortcuts,
      STRINGS.menu.diagnostics
    ]);
  });

  it("groups the menu with separators before the tools and panel sections (frontendDraft Step 3)", () => {
    const separated = CALL_MAP.filter((e) => e.separatorBefore).map((e) => e.fn);
    expect(separated).toEqual(["rostrumApplyStyles", "rostrumOpenPanel"]);
  });

  it("keeps onOpen and the google.script.run targets out of the menu (label null)", () => {
    const unlabeled = CALL_MAP.filter((e) => e.label === null).map((e) => e.fn);
    expect(unlabeled).toEqual([
      "onOpen",
      "rostrumSidebarState",
      "rostrumHideFromSidebar",
      "rostrumShowAllFromSidebar",
      "rostrumSaveSettings"
    ]);
  });
});

describe("pinned adapter constants", () => {
  it("pins the settings property key (persisted into users' documents)", () => {
    expect(SETTINGS_PROPERTY_KEY).toBe("rostrum.settings");
  });

  it("keeps the sidebar state mask to namedRanges + revisionId only (plan A13)", () => {
    expect(STATE_FIELDS_MASK).toBe("revisionId,namedRanges");
  });

  it("pins the diagnostics probe sizes and the small-doc embed cap", () => {
    expect(FONT_FLOOR_TRY_SIZES_PT).toEqual([0.25, 0.5, 0.75]);
    expect(SMALL_DOC_DUMP_MAX_BYTES).toBe(20000);
  });
});

// ---------------------------------------------------------------------------
// classifyBatchError — the matched 400 shape (documented in adapterPure).
// ---------------------------------------------------------------------------

describe("classifyBatchError", () => {
  it("classifies the revision-mismatch 400 by its message text", () => {
    const e = classifyBatchError(
      new Error(
        "API call to docs.documents.batchUpdate failed with error: The revision specified in the write control is not the most recent revision of the document."
      )
    );
    expect(e).toBeInstanceOf(RevisionMismatchError);
  });

  it("matches case-insensitively (the shape is message text, not a code)", () => {
    expect(classifyBatchError(new Error("REVISION conflict"))).toBeInstanceOf(RevisionMismatchError);
  });

  it("maps every other rejection to DocsApiError, preserving the message for logs", () => {
    const e = classifyBatchError(new Error("Invalid requests[0].updateTextStyle: bad field mask"));
    expect(e).toBeInstanceOf(DocsApiError);
    expect(e.message).toContain("bad field mask");
  });

  it("never throws on non-Error values (stringifies them)", () => {
    expect(classifyBatchError("boom")).toBeInstanceOf(DocsApiError);
    expect(classifyBatchError(null)).toBeInstanceOf(DocsApiError);
    // ...and still classifies a revision-shaped string correctly.
    expect(classifyBatchError("required revision mismatch")).toBeInstanceOf(RevisionMismatchError);
  });
});

// ---------------------------------------------------------------------------
// Dialog routing — error and receipt shaping the adapter renders verbatim.
// ---------------------------------------------------------------------------

describe("routeError", () => {
  it("routes every mapped error class to a red refusal with errorMessage's exact copy", () => {
    const cases: unknown[] = [
      new SuggestionsActiveError(),
      new MultiTabError(3),
      new HiddenStateError(),
      new RevisionMismatchError(),
      new PartialApplyError("hide", 1, 3),
      new DocsApiError("x")
    ];
    for (const e of cases) {
      const routed = routeError(e);
      const expected = errorMessage(e);
      expect(routed).toEqual({ dialog: "refusal", title: expected.title, body: expected.body, severity: "red" });
    }
  });

  it("routes unknown throws to the unknown refusal (the fail-visible promise)", () => {
    const routed = routeError("boom");
    expect(routed.dialog).toBe("refusal");
    expect(routed.body).toBe(STRINGS.errors.unknown.body);
  });
});

describe("receipt dialogs", () => {
  it("hideDialog is always a plain receipt under the shared receipt title", () => {
    const d = hideDialog(hideResult({ paragraphsScanned: 41, paragraphsChanged: 12, regionsHidden: 12 }));
    expect(d).toEqual({
      dialog: "receipt",
      title: STRINGS.dialogs.receiptTitle,
      body: "Hid text in 12 of 41 paragraphs.",
      severity: "plain"
    });
  });

  it("showAllDialog goes amber on each degraded bucket (normalize / sweep / newer-version skip)", () => {
    expect(showAllDialog(showResult({ segmentsRestoredExact: 9, segmentsNormalized: 1 })).severity).toBe("amber");
    expect(showAllDialog(showResult({ segmentsRestoredExact: 9, sweptOrphans: 1 })).severity).toBe("amber");
    expect(showAllDialog(showResult({ segmentsRestoredExact: 9, rangesSkippedNewerVersion: 1 })).severity).toBe("amber");
  });

  it("showAllDialog stays plain for exact-only restores and the no-op", () => {
    expect(showAllDialog(showResult({ segmentsRestoredExact: 9, rangesDeleted: 3 })).severity).toBe("plain");
    expect(showAllDialog(showResult()).severity).toBe("plain");
  });

  it("stylesDialog goes amber only on the degraded named-style path (plan A5/D13)", () => {
    expect(stylesDialog(stylesResult()).severity).toBe("plain");
    const degraded = stylesDialog(stylesResult({ namedStylesApplied: false }));
    expect(degraded.severity).toBe("amber");
    expect(degraded.body).toContain(STRINGS.receipts.stylesDegraded);
  });

  it("markCiteDialog renders the counted receipt and the teach-don't-scold zero", () => {
    expect(markCiteDialog(2).body).toBe(markCiteReceipt(2));
    expect(markCiteDialog(0).body).toBe(STRINGS.receipts.markCiteNoop);
    expect(markCiteDialog(0).severity).toBe("plain");
  });
});

// ---------------------------------------------------------------------------
// Sidebar state from the masked read.
// ---------------------------------------------------------------------------

describe("sidebarState", () => {
  const para = (text: string): RawParaSpec => ({ elements: [{ text }] });

  it("counts sizes records and undecodable rstm-family records; spacing records never double count", () => {
    const raw = rawDoc([para("abcdefghij")], {
      namedRanges: rawRanges(
        rawRange("rstm:v1:4x11", 1, 5), // a sizes record: one hidden passage
        rawRange("rstm:v1:p:6x4", 1, 11), // its parallel spacing record: NOT a second passage
        rawRange("rstm:v2:future-grammar", 6, 8), // newer version: hidden state, counted
        rawRange("docs-internal-guid", 2, 3) // foreign add-on: invisible
      )
    });
    expect(sidebarState(raw)).toEqual({ hiddenRegionCount: 2, armed: true });
  });

  it("reports a spacing-only remnant as armed with zero passages (decayed state is still armed)", () => {
    const raw = rawDoc([para("abc")], { namedRanges: rawRanges(rawRange("rstm:v1:p:6x4", 1, 4)) });
    expect(sidebarState(raw)).toEqual({ hiddenRegionCount: 0, armed: true });
  });

  it("reports a clean doc — and a foreign-ranges-only doc — as unarmed", () => {
    expect(sidebarState(rawDoc([para("abc")]))).toEqual({ hiddenRegionCount: 0, armed: false });
    const foreign = rawDoc([para("abc")], { namedRanges: rawRanges(rawRange("docs-internal-x", 1, 2)) });
    expect(sidebarState(foreign)).toEqual({ hiddenRegionCount: 0, armed: false });
  });

  it("never crashes the sidebar on junk reads (parse is defensive end to end)", () => {
    expect(sidebarState(42)).toEqual({ hiddenRegionCount: 0, armed: false });
    expect(sidebarState(null)).toEqual({ hiddenRegionCount: 0, armed: false });
  });
});

// ---------------------------------------------------------------------------
// Apply-styles confirm support.
// ---------------------------------------------------------------------------

describe("countDebateHeadings", () => {
  it("counts only H1-4 outside tables (the exact retro-pass blast radius)", () => {
    const raw = rawDoc([
      { style: "HEADING_1", elements: [{ text: "Pocket" }] },
      { style: "HEADING_2", elements: [{ text: "Hat" }] },
      { style: "HEADING_4", elements: [{ text: "Tag" }] },
      { style: "HEADING_5", elements: [{ text: "body-style heading" }] }, // H5 is body — not restyled
      { style: "TITLE", elements: [{ text: "Title" }] }, // kept by Hide but never restyled
      { style: "HEADING_3", inTable: true, elements: [{ text: "tabled" }] }, // tables untouched (row 1)
      { elements: [{ text: "prose" }] }
    ]);
    expect(countDebateHeadings(raw)).toBe(3);
  });

  it("returns zero for empty and junk reads (the zero-friction first run)", () => {
    expect(countDebateHeadings(rawDoc([]))).toBe(0);
    expect(countDebateHeadings("junk")).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Mark cite from selection picks.
// ---------------------------------------------------------------------------

describe("planMarkCiteFromPicks", () => {
  // H1 "Tag\n" [1,5) ; NORMAL "Smith " [5,11) + chip [11,12) + "2020 says\n"
  // [12,22). The NORMAL paragraph is segment-final, so index 21 (the final
  // newline) is unstylable.
  const markableDoc = (): unknown =>
    rawDoc([
      { style: "HEADING_1", elements: [{ text: "Tag" }] },
      {
        elements: [{ text: "Smith ", bold: true }, { text: "X", kind: "other" }, { text: "2020 says" }]
      }
    ]);
  const wholeParagraph = (ordinal: number): SelectionPick => ({ paragraphOrdinal: ordinal, startOffset: 0, endOffset: null });

  /** Narrow helper: every Mark-cite emission is an updateTextStyle. */
  function asTextStyleRequests(plan: { requests: unknown[] }): UpdateTextStyleRequest[] {
    return plan.requests as UpdateTextStyleRequest[];
  }

  it("refuses multi-tab docs outright (plan A3)", () => {
    const raw = rawDoc([], { tabs: [{ tabProperties: { tabId: "t1" } }, { tabProperties: { tabId: "t2" } }] });
    expect(() => planMarkCiteFromPicks(raw, [])).toThrow(MultiTabError);
  });

  it("refuses while suggestions are pending — indexes are untrustworthy (plan D5)", () => {
    const raw = rawDoc([{ elements: [{ text: "abc", suggested: true }] }]);
    expect(() => planMarkCiteFromPicks(raw, [])).toThrow(SuggestionsActiveError);
  });

  it("refuses on an armed doc — CITE_PT over a hidden span would fight the restore records (plan A5)", () => {
    const raw = rawDoc([{ elements: [{ text: "abcdef" }] }], { namedRanges: rawRanges(rawRange("rstm:v1:3x11", 1, 4)) });
    expect(() => planMarkCiteFromPicks(raw, [wholeParagraph(0)])).toThrow(HiddenStateError);
  });

  it("writes the cite convention over text spans only, splitting around the chip (whitelist, plan A9)", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), [wholeParagraph(1)]);
    const requests = asTextStyleRequests(plan);
    // Two ranges: before and after the chip; the final newline (21) clamped off.
    expect(requests.map((r) => r.updateTextStyle.range)).toEqual([
      { startIndex: 5, endIndex: 11 },
      { startIndex: 12, endIndex: 21 }
    ]);
    // The ONE emission shape shared with Apply-styles' repair (CITE_PT bold).
    for (const r of requests) {
      expect(r.updateTextStyle.textStyle).toEqual({ bold: true, fontSize: { magnitude: CITE_PT, unit: "PT" } });
      expect(r.updateTextStyle.fields).toBe("bold,fontSize");
    }
    expect(plan.citedParagraphs).toBe(1);
    expect(plan.revisionId).toBe("rev-1");
  });

  it("respects partial offsets (API-unit offsets relative to the paragraph start)", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), [{ paragraphOrdinal: 1, startOffset: 0, endOffset: 6 }]);
    expect(asTextStyleRequests(plan).map((r) => r.updateTextStyle.range)).toEqual([{ startIndex: 5, endIndex: 11 }]);
  });

  it("clamps a negative start offset to the paragraph start (defensive host math)", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), [{ paragraphOrdinal: 1, startOffset: -4, endOffset: 6 }]);
    expect(asTextStyleRequests(plan)[0].updateTextStyle.range).toEqual({ startIndex: 5, endIndex: 11 });
  });

  it("coalesces adjacent text elements into one write (fewer requests, same effect)", () => {
    const raw = rawDoc([{ elements: [{ text: "Valcke ", bold: true }, { text: "et al. 20" }] }]);
    const plan = planMarkCiteFromPicks(raw, [wholeParagraph(0)]);
    // 1..17 text + newline at 17 clamped (segment-final paragraph).
    expect(asTextStyleRequests(plan).map((r) => r.updateTextStyle.range)).toEqual([{ startIndex: 1, endIndex: 17 }]);
    expect(plan.citedParagraphs).toBe(1);
  });

  it("keeps a NON-final paragraph's trailing newline markable (sizing a paragraph mark is benign)", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), [wholeParagraph(0)]);
    // H1 paragraph [1,5) is not segment-final: the newline at 4 stays in range.
    expect(asTextStyleRequests(plan)[0].updateTextStyle.range).toEqual({ startIndex: 1, endIndex: 5 });
  });

  it("counts distinct paragraphs, not picks (two picks in one line = one cite)", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), [
      { paragraphOrdinal: 1, startOffset: 0, endOffset: 6 },
      { paragraphOrdinal: 1, startOffset: 7, endOffset: null }
    ]);
    expect(plan.citedParagraphs).toBe(1);
    expect(plan.requests.length).toBe(2);
  });

  it("drops out-of-range ordinals instead of writing at guessed indexes", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), [wholeParagraph(99)]);
    expect(plan.requests).toEqual([]);
    expect(plan.citedParagraphs).toBe(0);
  });

  it("plans nothing for an empty selection (the adapter then shows the markCiteNoop receipt)", () => {
    const plan = planMarkCiteFromPicks(markableDoc(), []);
    expect(plan.requests).toEqual([]);
    expect(plan.citedParagraphs).toBe(0);
  });

  it("yields nothing for a chip-only pick (whitelist: never markable)", () => {
    const raw = rawDoc([{ elements: [{ text: "ab" }, { text: "XX", kind: "other" }, { text: "cd" }] }]);
    // Pick covering ONLY the chip's two index units: text [1,3) chip [3,5) text [5,8).
    const plan = planMarkCiteFromPicks(raw, [{ paragraphOrdinal: 0, startOffset: 2, endOffset: 4 }]);
    expect(plan.requests).toEqual([]);
    expect(plan.citedParagraphs).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// Diagnostics interpretation.
// ---------------------------------------------------------------------------

describe("sumApiUnits + textPickOffsets — the host-selection lowering arithmetic", () => {
  // The DocumentApp tree-walk that produces these inputs is host-only and
  // untestable without a live Doc; this pins the index math it feeds, which is
  // the one off-by-one that would silently mis-span a Mark-cite (review gap).
  describe("sumApiUnits", () => {
    it("is 0 when nothing precedes the element", () => {
      expect(sumApiUnits([])).toBe(0);
    });

    it("sums preceding TEXT siblings by character length", () => {
      expect(sumApiUnits([{ isText: true, textLength: 5 }, { isText: true, textLength: 3 }])).toBe(8);
    });

    it("counts every non-TEXT sibling (image, chip, break) as exactly one unit", () => {
      expect(sumApiUnits([{ isText: false, textLength: 0 }, { isText: false, textLength: 999 }])).toBe(2);
    });

    it("mixes text-length and one-per-object correctly", () => {
      // "ab" (2) + image (1) + "cde" (3) = 6 units before the target.
      expect(
        sumApiUnits([
          { isText: true, textLength: 2 },
          { isText: false, textLength: 0 },
          { isText: true, textLength: 3 }
        ])
      ).toBe(6);
    });
  });

  describe("textPickOffsets", () => {
    it("spans the whole run from `before` when the element is wholly selected", () => {
      expect(textPickOffsets(4, { partial: false, textLength: 7 })).toEqual({ startOffset: 4, endOffset: 11 });
    });

    it("converts a partial range's INCLUSIVE end to an EXCLUSIVE offset (the +1)", () => {
      // chars [2..5] inclusive within the run, no preceding units -> [2, 6).
      expect(textPickOffsets(0, { partial: true, startOffsetInText: 2, endOffsetInclusiveInText: 5 })).toEqual({
        startOffset: 2,
        endOffset: 6
      });
    });

    it("adds the preceding-unit offset to both ends of a partial range", () => {
      // 10 units precede; a single char [0..0] inclusive -> [10, 11).
      expect(textPickOffsets(10, { partial: true, startOffsetInText: 0, endOffsetInclusiveInText: 0 })).toEqual({
        startOffset: 10,
        endOffset: 11
      });
    });

    it("matches the whole-element path for a full-run partial selection", () => {
      // A partial range covering [0..len-1] inclusive equals the whole-run span.
      const before = 3;
      const len = 7;
      expect(
        textPickOffsets(before, { partial: true, startOffsetInText: 0, endOffsetInclusiveInText: len - 1 })
      ).toEqual(textPickOffsets(before, { partial: false, textLength: len }));
    });
  });
});

describe("interpretFontFloor", () => {
  it("interprets the three verdicts", () => {
    expect(interpretFontFloor({ triedPt: 0.25, appliedOk: false, readBackPt: null })).toEqual({
      triedPt: 0.25,
      readBackPt: null,
      verdict: "rejected"
    });
    expect(interpretFontFloor({ triedPt: 0.5, appliedOk: true, readBackPt: 0.5 })).toEqual({
      triedPt: 0.5,
      readBackPt: 0.5,
      verdict: "accepted exact"
    });
    expect(interpretFontFloor({ triedPt: 0.5, appliedOk: true, readBackPt: 1 })).toEqual({
      triedPt: 0.5,
      readBackPt: 1,
      verdict: "clamped"
    });
  });

  it("treats an applied-but-unreadable size as clamped, never as exact", () => {
    expect(interpretFontFloor({ triedPt: 0.75, appliedOk: true, readBackPt: null }).verdict).toBe("clamped");
  });
});

describe("diagnosticsReadFacts", () => {
  it("locates the probe slot just before the body's final newline", () => {
    const raw = rawDoc([{ elements: [{ text: "Hi" }] }]); // "Hi\n" = [1,4)
    expect(diagnosticsReadFacts(raw)).toEqual({ revisionId: "rev-1", probeIndex: 3, headingSixSizePt: null });
  });

  it("reads HEADING_6's stated size for the no-visible-change probe", () => {
    const raw = rawDoc([{ elements: [{ text: "Hi" }] }], {
      namedStyles: { styles: [{ namedStyleType: "HEADING_6", textStyle: { fontSize: { magnitude: 12, unit: "PT" } } }] }
    });
    expect(diagnosticsReadFacts(raw).headingSixSizePt).toBe(12);
  });

  it("degrades junk reads to index 1 so the insert fails loudly server-side", () => {
    expect(diagnosticsReadFacts(null)).toEqual({ revisionId: "", probeIndex: 1, headingSixSizePt: null });
  });
});

describe("probeReadbackSizePt", () => {
  const raw = rawDoc([{ elements: [{ text: "ab", sizePt: 0.5 }, { text: "XX", kind: "other" }, { text: "cd" }] }]);
  // Layout: text [1,3) at 0.5pt; chip [3,5); text [5,8) inheriting.

  it("returns the stated size of the text element containing the index", () => {
    expect(probeReadbackSizePt(raw, 1)).toBe(0.5);
    expect(probeReadbackSizePt(raw, 2)).toBe(0.5);
  });

  it("returns null for inherited sizes, chip indexes, and out-of-document indexes", () => {
    expect(probeReadbackSizePt(raw, 5)).toBeNull(); // text but no stated size
    expect(probeReadbackSizePt(raw, 3)).toBeNull(); // inside the chip
    expect(probeReadbackSizePt(raw, 99)).toBeNull(); // nowhere
  });
});

describe("buildDiagnostics", () => {
  const baseInput = {
    fetchLatencyMs: 88,
    applyLatenciesMs: [12, 34],
    fontFloor: [
      { triedPt: 0.25, appliedOk: false, readBackPt: null },
      { triedPt: 0.5, appliedOk: true, readBackPt: 1 },
      { triedPt: 0.75, appliedOk: true, readBackPt: 0.75 }
    ],
    namedStyleProbe: "ok" as const,
    inheritClear: "ok" as const,
    cleanupOk: true
  };

  it("measures the payload in true UTF-8 bytes (multibyte and astral chars included)", () => {
    const raw = rawDoc([{ elements: [{ text: "café — 💡" }] }]);
    const model = buildDiagnostics({ raw, ...baseInput });
    expect(model.payloadBytes).toBe(Buffer.byteLength(JSON.stringify(raw), "utf8"));
  });

  it("inventories rstm-family ranges of EVERY kind (sizes, spacing, unknown) and ignores foreign ones", () => {
    const raw = rawDoc([{ elements: [{ text: "abcdefghij" }] }], {
      namedRanges: rawRanges(
        rawRange("rstm:v1:4x11", 1, 5),
        rawRange("rstm:v1:p:6x4", 1, 11),
        rawRange("rstm:v9:???", 6, 8),
        rawRange("docs-internal-x", 2, 3)
      )
    });
    expect(buildDiagnostics({ raw, ...baseInput }).rstmRangeCount).toBe(3);
  });

  it("builds the character-weighted background histogram, capped at the top N", () => {
    // 10 distinct hexes with strictly descending character counts. The
    // builder folds each paragraph's trailing newline into its final text
    // element, so each count is text length + 1 (11..2).
    const paras: RawParaSpec[] = [];
    for (let i = 0; i < 10; i++) {
      const hex = `#0${i}0000`;
      paras.push({ elements: [{ text: "x".repeat(10 - i), bgHex: hex }] });
    }
    const model = buildDiagnostics({ raw: rawDoc(paras), ...baseInput });
    expect(model.bgHistogram).toHaveLength(BG_HISTOGRAM_TOP_N);
    expect(model.bgHistogram[0]).toEqual({ hex: "#000000", count: 11 });
    const counts = model.bgHistogram.map((r) => r.count);
    expect([...counts].sort((a, b) => b - a)).toEqual(counts); // descending
  });

  it("breaks histogram count ties by hex so the report is deterministic", () => {
    const raw = rawDoc([
      { elements: [{ text: "abc", bgHex: "#ff0000" }] },
      { elements: [{ text: "xyz", bgHex: "#00ff00" }] }
    ]);
    const model = buildDiagnostics({ raw, ...baseInput });
    expect(model.bgHistogram.map((r) => r.hex)).toEqual(["#00ff00", "#ff0000"]);
  });

  it("embeds whole-doc JSON only at or under the cap (plan A11.ii)", () => {
    const small = buildDiagnostics({ raw: rawDoc([{ elements: [{ text: "tiny" }] }]), ...baseInput });
    expect(small.smallDocJson).toBe(JSON.stringify(rawDoc([{ elements: [{ text: "tiny" }] }])));
    const big = buildDiagnostics({ raw: rawDoc([{ elements: [{ text: "x".repeat(SMALL_DOC_DUMP_MAX_BYTES) }] }]), ...baseInput });
    expect(big.smallDocJson).toBeNull();
  });

  it("interprets the font-floor readings and passes facts through", () => {
    const model = buildDiagnostics({ raw: rawDoc([{ elements: [{ text: "hi" }] }]), ...baseInput });
    expect(model.version).toBe(GDOCS_VERSION);
    expect(model.tabCount).toBe(1);
    expect(model.fontFloor.map((r) => r.verdict)).toEqual(["rejected", "clamped", "accepted exact"]);
    expect(model.fetchLatencyMs).toBe(88);
    expect(model.applyLatenciesMs).toEqual([12, 34]);
  });
});

describe("renderDiagnosticsText", () => {
  /** A complete model so each case overrides only what it renders. */
  function model(over: Partial<DiagnosticsModel> = {}): DiagnosticsModel {
    return {
      version: GDOCS_VERSION,
      tabCount: 1,
      rstmRangeCount: 2,
      payloadBytes: 1234,
      fetchLatencyMs: 88,
      applyLatenciesMs: [12, 34],
      fontFloor: [
        { triedPt: 0.25, readBackPt: null, verdict: "rejected" },
        { triedPt: 0.5, readBackPt: 1, verdict: "clamped" },
        { triedPt: 0.75, readBackPt: 0.75, verdict: "accepted exact" }
      ],
      namedStyleProbe: "ok",
      inheritClear: "ok",
      cleanupOk: true,
      bgHistogram: [{ hex: "#ffff00", count: 120 }],
      smallDocJson: '{"tiny":true}',
      ...over
    };
  }

  it("renders the headline facts (version, tabs, records, payload, latencies)", () => {
    const text = renderDiagnosticsText(model());
    expect(text).toContain(`Rostrum for Google Docs diagnostics (v${GDOCS_VERSION})`);
    expect(text).toContain("tabs: 1");
    expect(text).toContain("hidden-text records: 2");
    expect(text).toContain("read payload: 1234 bytes (fetched in 88 ms)");
    expect(text).toContain("write latencies (ms): 12, 34");
  });

  it("renders one line per font-floor verdict, with the read-back size on clamps", () => {
    const text = renderDiagnosticsText(model());
    expect(text).toContain("0.25pt -> rejected by the API");
    expect(text).toContain("0.5pt -> clamped (reads back as 1pt)");
    expect(text).toContain("0.75pt -> accepted exact");
    // A clamp whose read-back stated nothing says so rather than inventing a number.
    expect(
      renderDiagnosticsText(model({ fontFloor: [{ triedPt: 0.5, readBackPt: null, verdict: "clamped" }] }))
    ).toContain("clamped (reads back as no stated size)");
  });

  it("renders the empty/skipped branches honestly", () => {
    const text = renderDiagnosticsText(
      model({ applyLatenciesMs: [], fontFloor: [], namedStyleProbe: "skipped", bgHistogram: [], smallDocJson: null })
    );
    expect(text).toContain("write latencies (ms): none");
    expect(text).toContain("not attempted");
    expect(text).toContain("skipped (the read stated no heading 6 size)");
    expect(text).toContain("none found");
    expect(text).toContain(`omitted - 1234 bytes exceeds the ${SMALL_DOC_DUMP_MAX_BYTES}-byte cap`);
  });

  it("tells the user how to fix a failed probe cleanup by hand", () => {
    expect(renderDiagnosticsText(model({ cleanupOk: false }))).toContain("a stray '.' probe character may remain");
    expect(renderDiagnosticsText(model())).toContain("probe cleanup: ok");
  });

  it("embeds the small-doc JSON and the histogram rows", () => {
    const text = renderDiagnosticsText(model());
    expect(text).toContain('{"tiny":true}');
    expect(text).toContain("#ffff00 - 120 chars");
  });
});
