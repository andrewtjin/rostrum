// Keeper-policy suite for gdocs/src/core/keepers.ts (plan S5; edge rows 3-5;
// A8 closed keep-set, A9 whitelist, A10 structural cite). The policy is the
// heart of invisibility, so the suite leans on the two leak directions:
//   * keeping too little hides a cite/highlight the debater needs (the
//     plan-review BLOCKER — Docs-native bold-11pt cites), and
//   * keeping too much leaves body prose visible (a structural-rule false
//     positive, or near-white web shading treated as a highlight).
// Every negative here is therefore a failure-path test by design, ported from
// the Word citeRepair suite's leak-prevention cases where one exists.

import { detectCiteLeads, isHighlightKept, planKeeps } from "../gdocs/src/core/keepers";
import { DEFAULT_CITE_MIN_PT, NEAR_WHITE_MIN_CHANNEL } from "../gdocs/src/core/constants";
import { DEFAULT_KEEP_HEXES } from "../gdocs/src/core/settings";
import { GdocsSettings } from "../gdocs/src/core/types";
import { buildDoc, GpSpec, para } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** Resolved settings with the engine's real defaults; override per test. The
 * literal (not resolveSettings) keeps this suite independent of the settings
 * module's parsing — keepers only ever sees a RESOLVED GdocsSettings. */
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

/** A kept tag heading — the structural-cite window opener. */
const TAG: GpSpec = para("Impact is fast", "HEADING_4");

/**
 * The Docs-NATIVE cite shape the structural rule exists for (plan A10 / the
 * plan-review BLOCKER): author+year bolded at the inherited 11pt default —
 * bold AND >= 13 misses it, so only the structural rule can keep it.
 */
function nativeCite(): GpSpec {
  return {
    elements: [
      { text: "Valcke et al. 20", bold: true },
      { text: " The card body argues the impact." }
    ]
  };
}

// ---------------------------------------------------------------------------
// 1. isHighlightKept (row 5 + plan A8)
// ---------------------------------------------------------------------------

describe("isHighlightKept (closed default set + anyHighlight master toggle, A8)", () => {
  it("a null background never keeps, in either mode", () => {
    expect(isHighlightKept(null, cfg())).toBe(false);
    expect(isHighlightKept(null, cfg({ keepMode: "anyHighlight" }))).toBe(false);
  });

  it("the default closed set keeps classic yellow and the Word-import highlight hexes", () => {
    expect(isHighlightKept("#ffff00", cfg())).toBe(true); // classic yellow
    expect(isHighlightKept("#8b0000", cfg())).toBe(true); // Word darkRed materialized by import
  });

  it("white and near-white web shading are ignored in BOTH modes (#f8f9fa must hide)", () => {
    for (const hex of ["#ffffff", "#f8f9fa"]) {
      expect(isHighlightKept(hex, cfg())).toBe(false); // not in the closed set
      expect(isHighlightKept(hex, cfg({ keepMode: "anyHighlight" }))).toBe(false); // near-white floor
    }
  });

  it("a custom closed set keeps exactly its members and excludes everything else", () => {
    const custom = cfg({ keepColors: new Set(["#00ffff"]) });
    expect(isHighlightKept("#00ffff", custom)).toBe(true);
    expect(isHighlightKept("#ffff00", custom)).toBe(false); // even classic yellow is out
  });

  it("an empty keep-set keeps nothing by color (an honored explicit choice)", () => {
    expect(isHighlightKept("#ffff00", cfg({ keepColors: new Set() }))).toBe(false);
  });

  it('"anyHighlight" keeps an off-palette pink but not near-white', () => {
    const any = cfg({ keepMode: "anyHighlight" });
    expect(isHighlightKept("#ff69b4", any)).toBe(true); // a real (if odd) highlight
    expect(isHighlightKept("#f8f9fa", any)).toBe(false); // web-paste residue
  });

  it('"anyHighlight" boundary sits exactly at NEAR_WHITE_MIN_CHANNEL (plan A8: min >= 0.95)', () => {
    // Plan A8 draws near-white at min(r,g,b) >= 0.95; 0.95 * 255 = 242.25, so the
    // floor is 243 and 0xf2 (242 ≈ 0.949) is just below the line and stays kept.
    expect(NEAR_WHITE_MIN_CHANNEL).toBe(243);
    expect(parseInt("f2", 16)).toBe(NEAR_WHITE_MIN_CHANNEL - 1);
    const any = cfg({ keepMode: "anyHighlight" });
    expect(isHighlightKept("#f3f3f3", any)).toBe(false); // min channel == floor -> hide
    expect(isHighlightKept("#f2f2f2", any)).toBe(true); // one below the floor -> keep
    expect(isHighlightKept("#f1f2f3", any)).toBe(true); // below the floor -> keep
  });

  it('"anyHighlight" over-keeps an unparseable hex (defensive: cannot prove near-white)', () => {
    expect(isHighlightKept("yellow", cfg({ keepMode: "anyHighlight" }))).toBe(true);
    // In set mode the same junk simply fails membership — no over-keep needed.
    expect(isHighlightKept("yellow", cfg())).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// 2. planKeeps — whole-paragraph structural rules (rows 1 + 3)
// ---------------------------------------------------------------------------

describe("planKeeps — table and heading keeps (rows 1 + 3)", () => {
  it.each(["HEADING_1", "HEADING_2", "HEADING_3", "HEADING_4", "TITLE", "SUBTITLE"] as const)(
    "%s keeps the whole paragraph (TITLE/SUBTITLE = deliberate over-keep divergence)",
    (style) => {
      const k = planKeeps(buildDoc([para("Structural line", style)]), cfg())[0];
      expect(k).toEqual({ keepWhole: true, elementKeep: [true], citeDetected: false });
    }
  );

  it.each(["HEADING_5", "HEADING_6", "NORMAL_TEXT"] as const)("%s is body (hidden when plain)", (style) => {
    const k = planKeeps(buildDoc([para("plain body line", style)]), cfg())[0];
    expect(k).toEqual({ keepWhole: false, elementKeep: [false], citeDetected: false });
  });

  it("a table paragraph is kept whole (row 1 — tables untouched)", () => {
    const k = planKeeps(buildDoc([{ inTable: true, elements: [{ text: "cell text" }] }]), cfg())[0];
    expect(k).toEqual({ keepWhole: true, elementKeep: [true], citeDetected: false });
  });

  it("rule order: table and heading keeps win BEFORE the cite predicates (citeDetected stays false)", () => {
    // A bold-26pt pocket heading and a bold-14pt table cell both satisfy the
    // signature shape — but they are kept as structure, not counted as cites.
    const doc = buildDoc([
      { style: "HEADING_1", elements: [{ text: "Pocket", bold: true, size: 26 }] },
      { inTable: true, elements: [{ text: "Smith 20", bold: true, size: 14 }] }
    ]);
    const keeps = planKeeps(doc, cfg());
    expect(keeps[0]).toEqual({ keepWhole: true, elementKeep: [true], citeDetected: false });
    expect(keeps[1]).toEqual({ keepWhole: true, elementKeep: [true], citeDetected: false });
  });

  it("an empty doc plans no keeps", () => {
    expect(planKeeps(buildDoc([]), cfg())).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 3. planKeeps — cite SIGNATURE rule (row 4)
// ---------------------------------------------------------------------------

describe("planKeeps — cite signature (bold AND >= citeMinPt, row 4)", () => {
  it("bold + 13pt (the default threshold, inclusive) keeps the paragraph whole as a cite", () => {
    const doc = buildDoc([{ elements: [{ text: "Smith 20", bold: true, size: 13 }, { text: " descriptor" }] }]);
    const k = planKeeps(doc, cfg())[0];
    expect(k).toEqual({ keepWhole: true, elementKeep: [true, true], citeDetected: true });
  });

  it("14pt but NOT bold is not a cite (both signature halves are required)", () => {
    const doc = buildDoc([{ elements: [{ text: "Smith 20", size: 14 }, { text: " descriptor" }] }]);
    const k = planKeeps(doc, cfg())[0];
    expect(k.keepWhole).toBe(false);
    expect(k.citeDetected).toBe(false);
  });

  it("a signature cite inside a HEADING_5 paragraph is kept via cite (not just NORMAL_TEXT, A10)", () => {
    // H5/H6 are body styles here, so the cite predicate must still reach them.
    const doc = buildDoc([
      { style: "HEADING_5", elements: [{ text: "Smith 20", bold: true, size: 14 }, { text: " analysis" }] }
    ]);
    const k = planKeeps(doc, cfg())[0];
    expect(k.keepWhole).toBe(true);
    expect(k.citeDetected).toBe(true);
  });

  it("resolves an INHERITED element size through namedStyleSizesPt (Docs-native cites inherit)", () => {
    const citeLine: GpSpec = { elements: [{ text: "Smith 20", bold: true }, { text: " descriptor" }] };
    // Normal resolved at 14pt: bold + inherited-14 IS a cite...
    const kept = planKeeps(buildDoc([citeLine], { namedStyleSizesPt: { NORMAL_TEXT: 14 } }), cfg())[0];
    expect(kept.keepWhole).toBe(true);
    expect(kept.citeDetected).toBe(true);
    // ...and at 11pt the same paragraph is not (no heading: structural can't fire either).
    const hidden = planKeeps(buildDoc([citeLine], { namedStyleSizesPt: { NORMAL_TEXT: 11 } }), cfg())[0];
    expect(hidden.keepWhole).toBe(false);
  });

  it("a double-inherit (no namedStyles entry at all) falls back to the 11pt Docs default", () => {
    const doc = buildDoc([{ elements: [{ text: "Smith 20", bold: true }] }], { namedStyleSizesPt: {} });
    expect(planKeeps(doc, cfg())[0].keepWhole).toBe(false); // 11 < 13: not a cite
  });

  it("honors a custom citeMinPt", () => {
    const doc = buildDoc([{ elements: [{ text: "Smith 20", bold: true, size: 12 }] }]);
    expect(planKeeps(doc, cfg())[0].keepWhole).toBe(false); // 12 < default 13
    const k = planKeeps(doc, cfg({ citeMinPt: 12 }))[0];
    expect(k.keepWhole).toBe(true);
    expect(k.citeDetected).toBe(true);
  });

  it('a bold "other" element never triggers the signature (text-only whitelist, A9)', () => {
    // A chip carrying bold/size metadata is structure, not a cite run.
    const doc = buildDoc([
      { elements: [{ text: "@DocChip", kind: "other", bold: true, size: 14 }, { text: " body text" }] }
    ]);
    const k = planKeeps(doc, cfg())[0];
    expect(k.keepWhole).toBe(false);
    expect(k.citeDetected).toBe(false);
    expect(k.elementKeep).toEqual([true, false]); // chip kept structurally, text hidden
  });
});

// ---------------------------------------------------------------------------
// 4. planKeeps — STRUCTURAL cite rule (plan A10, the ported citeRepair heuristic)
// ---------------------------------------------------------------------------

describe("planKeeps — structural cite (bold Author-YEAR lead after a kept heading, A10)", () => {
  it("bold-only 11pt is NOT a cite by signature, but IS kept structurally after a kept heading", () => {
    // Alone (no heading) the Docs-native cite hides — the signature misses it...
    const alone = planKeeps(buildDoc([nativeCite()]), cfg())[0];
    expect(alone.keepWhole).toBe(false);
    expect(alone.citeDetected).toBe(false);
    // ...but in its real structural slot (first body paragraph under a tag) it is kept whole.
    const keeps = planKeeps(buildDoc([TAG, nativeCite()]), cfg());
    expect(keeps[1]).toEqual({ keepWhole: true, elementKeep: [true, true], citeDetected: true });
  });

  it("is gated on settings.structuralCite (off -> the same paragraph hides)", () => {
    const keeps = planKeeps(buildDoc([TAG, nativeCite()]), cfg({ structuralCite: false }));
    expect(keeps[1].keepWhole).toBe(false);
    expect(keeps[1].citeDetected).toBe(false);
  });

  it("first-name last-name with only the last name bold ('Barbara **Valcke** 20')", () => {
    // The non-bold "Barbara " prefix is name-like (short, digit-free, no
    // clause punctuation), so the lead still qualifies — the Word port's
    // split-name positive.
    const doc = buildDoc([
      TAG,
      { elements: [{ text: "Barbara " }, { text: "Valcke 20", bold: true }, { text: " descriptor" }] }
    ]);
    expect(planKeeps(doc, cfg())[1]).toEqual({
      keepWhole: true,
      elementKeep: [true, true, true],
      citeDetected: true
    });
  });

  it("accepts apostrophe-year shorthand (Author '18)", () => {
    const doc = buildDoc([TAG, { elements: [{ text: "Bunzel '18", bold: true }, { text: " descriptor text" }] }]);
    expect(planKeeps(doc, cfg())[1].citeDetected).toBe(true);
  });

  it("skips empty paragraphs between the tag and the cite", () => {
    const keeps = planKeeps(buildDoc([TAG, { elements: [{ text: "" }] }, nativeCite()]), cfg());
    expect(keeps[2].keepWhole).toBe(true);
    expect(keeps[2].citeDetected).toBe(true);
  });

  it("a deeper subtag supersedes the outer window (hat -> tag -> cite)", () => {
    const keeps = planKeeps(buildDoc([para("Warming hat", "HEADING_2"), TAG, nativeCite()]), cfg());
    expect(keeps[2].keepWhole).toBe(true);
    expect(keeps[2].citeDetected).toBe(true);
  });

  it("TITLE opens a cite window too (it is a kept heading here)", () => {
    const keeps = planKeeps(buildDoc([para("Aff Case", "TITLE"), nativeCite()]), cfg());
    expect(keeps[1].citeDetected).toBe(true);
  });

  // --- leak prevention: ported negative cases from the Word citeRepair suite ---

  it("a cite-shaped paragraph NOT after a heading is not kept (no window opens)", () => {
    const keeps = planKeeps(buildDoc([para("intro prose"), nativeCite()]), cfg());
    expect(keeps[1].keepWhole).toBe(false);
  });

  it("body prose 'In 2008, the crisis worsened and BOLD said' is NOT a cite (prefix gate)", () => {
    // Year is early, but the first bold element is deep and the prefix carries
    // a digit + comma — the exact false positive the heuristic must refuse.
    const doc = buildDoc([
      TAG,
      {
        elements: [
          { text: "In 2008, the crisis worsened and " },
          { text: "experts", bold: true },
          { text: " said the worst." }
        ]
      }
    ]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("a SHORT prefix carrying a digit is not name-like ('In 2008, ' + bold)", () => {
    // The first bold element is within 30 chars here, so the earlier offset
    // gate passes and the prefix gate itself must do the rejecting.
    const doc = buildDoc([
      TAG,
      { elements: [{ text: "In 2008, " }, { text: "experts", bold: true }, { text: " warned of 20 risks" }] }
    ]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("a SHORT prefix with clause punctuation is not name-like ('He said, ' + bold)", () => {
    const doc = buildDoc([
      TAG,
      { elements: [{ text: "He said, " }, { text: "Smith 20", bold: true }, { text: " matters" }] }
    ]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("a non-bold paragraph containing a year is NOT a cite", () => {
    const doc = buildDoc([TAG, { elements: [{ text: "Smith 2020 says something" }] }]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("a bold Author-YEAR MID-CARD (not the first candidate) is never reached — only the slot counts", () => {
    // The first body paragraph is the only candidate; bold emphasis deeper in
    // the card must not be mistaken for a cite even when it matches the shape.
    const doc = buildDoc([
      TAG,
      nativeCite(), // the real candidate, kept
      {
        elements: [
          { text: "Body text later cites " },
          { text: "Smith 2020", bold: true },
          { text: " in passing." }
        ]
      }
    ]);
    const keeps = planKeeps(doc, cfg());
    expect(keeps[1].citeDetected).toBe(true);
    expect(keeps[2].keepWhole).toBe(false); // the deep mention hides with the card
    expect(keeps[2].citeDetected).toBe(false);
  });

  it("a candidate with no year near the front is NOT a cite", () => {
    const doc = buildDoc([
      TAG,
      { elements: [{ text: "Smith", bold: true }, { text: " argues at length without any date in the first part" }] }
    ]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("a candidate whose first bold word is past char 30 is NOT a cite", () => {
    const doc = buildDoc([
      TAG,
      {
        elements: [
          { text: "This sentence runs on for quite a while before " }, // 47-char prose prefix
          { text: "word", bold: true },
          { text: " 20 more text" }
        ]
      }
    ]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("digits that are the tail of a longer number are not a year (page counts, DOIs)", () => {
    const doc = buildDoc([TAG, { elements: [{ text: "Valcke", bold: true }, { text: " 12696 attendees were counted" }] }]);
    expect(planKeeps(doc, cfg())[1].keepWhole).toBe(false);
  });

  it("the window BREAKS at a non-NORMAL_TEXT paragraph (an H5 line shields the cite below it)", () => {
    // Conservative in the false-positive direction: a structural paragraph
    // between tag and cite means the slot is not the canonical tag->cite shape.
    const doc = buildDoc([TAG, para("analysis line", "HEADING_5"), nativeCite()]);
    expect(planKeeps(doc, cfg())[2].keepWhole).toBe(false);
  });

  it("the window BREAKS at a table (tables are never scanned through)", () => {
    const doc = buildDoc([TAG, { inTable: true, elements: [{ text: "row" }] }, nativeCite()]);
    expect(planKeeps(doc, cfg())[2].keepWhole).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// 5. planKeeps — per-element verdicts (rows 2 + 5, whitelist A9, lesson #14)
// ---------------------------------------------------------------------------

describe("planKeeps — per-element verdicts in a hidden paragraph", () => {
  it('keeps "other" elements (chips/breaks/objects) while hiding the text around them (A9)', () => {
    const doc = buildDoc([{ elements: [{ text: "see " }, { text: "@chip", kind: "other" }, { text: " here" }] }]);
    const k = planKeeps(doc, cfg())[0];
    expect(k.keepWhole).toBe(false);
    expect(k.elementKeep).toEqual([false, true, false]);
  });

  it("keeps a keeper-highlighted element and hides its unhighlighted neighbors", () => {
    const doc = buildDoc([{ elements: [{ text: "key claim", bg: "#ffff00" }, { text: " filler prose" }] }]);
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([true, false]);
  });

  it('"anyHighlight" through planKeeps: off-palette pink keeps, near-white shading hides', () => {
    const doc = buildDoc([
      { elements: [{ text: "pink", bg: "#ff69b4" }, { text: " and " }, { text: "shaded", bg: "#f8f9fa" }] }
    ]);
    expect(planKeeps(doc, cfg({ keepMode: "anyHighlight" }))[0].elementKeep).toEqual([true, false, false]);
    // In the default closed set the pink is off-palette and hides too.
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([false, false, false]);
  });

  it("keeps a whitespace-only element bridging two kept words (lesson #14)", () => {
    const doc = buildDoc([
      { elements: [{ text: "reasons", bg: "#ffff00" }, { text: " " }, { text: "for", bg: "#ffff00" }] }
    ]);
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([true, true, true]);
  });

  it("bridges multiple consecutive whitespace-only elements as a unit", () => {
    const doc = buildDoc([
      { elements: [{ text: "a", bg: "#ffff00" }, { text: " " }, { text: "\t" }, { text: "b", bg: "#ffff00" }] }
    ]);
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([true, true, true, true]);
  });

  it("bridges typographic spaces too (NBSP is in the Zs class pasted text carries)", () => {
    const doc = buildDoc([
      { elements: [{ text: "gives", bg: "#ffff00" }, { text: " " }, { text: "Russia", bg: "#ffff00" }] }
    ]);
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([true, true, true]);
  });

  it('a kept "other" element anchors a bridge like any kept element', () => {
    const doc = buildDoc([
      { elements: [{ text: "@chip", kind: "other" }, { text: " " }, { text: "kept", bg: "#ffff00" }] }
    ]);
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([true, true, true]);
  });

  it("hides whitespace adjacent to hidden content (no full-size space floats beside shrunken text)", () => {
    const oneSide = buildDoc([{ elements: [{ text: "keep", bg: "#ffff00" }, { text: " " }, { text: "hiddenbody" }] }]);
    expect(planKeeps(oneSide, cfg())[0].elementKeep).toEqual([true, false, false]);
    const noSide = buildDoc([{ elements: [{ text: "foo" }, { text: " " }, { text: "bar" }] }]);
    expect(planKeeps(noSide, cfg())[0].elementKeep).toEqual([false, false, false]);
  });

  it("the trailing newline lives inside the final element and follows its fate (plan D6)", () => {
    // Builder reality check first: the final text element's range covers the
    // paragraph end (the newline has no node of its own in the Docs API).
    const kept = buildDoc([{ elements: [{ text: "kept line", bg: "#ffff00" }] }]);
    const p = kept.paragraphs[0];
    expect(p.elements[p.elements.length - 1].endIndex).toBe(p.endIndex);
    // Kept element -> its newline is kept with it; hidden element -> hidden with it.
    expect(planKeeps(kept, cfg())[0].elementKeep).toEqual([true]);
    expect(planKeeps(buildDoc([para("hidden line")]), cfg())[0].elementKeep).toEqual([false]);
  });

  it("an empty paragraph (lone newline element) hides, so empty lines condense", () => {
    const doc = buildDoc([{ elements: [{ text: "" }] }]);
    expect(planKeeps(doc, cfg())[0].elementKeep).toEqual([false]);
  });
});

// ---------------------------------------------------------------------------
// 6. detectCiteLeads (plan A10(b) — the ranges styles.ts repairs to bold CITE_PT)
// ---------------------------------------------------------------------------

describe("detectCiteLeads — exact repair ranges for Apply-styles", () => {
  it("returns the exact range of a single bold lead element", () => {
    // "Tag\n" occupies [1,5); the bold lead "Valcke et al. 20" is 16 chars at [5,21).
    const doc = buildDoc([para("Tag", "HEADING_4"), nativeCite()]);
    expect(detectCiteLeads(doc, cfg())).toEqual([{ startIndex: 5, endIndex: 21 }]);
  });

  it("returns one exact range per bold lead element when author and year are split", () => {
    const doc = buildDoc([
      para("Tag", "HEADING_4"), // [1,5)
      {
        elements: [
          { text: "Valcke", bold: true }, // [5,11)  lead
          { text: " et al. " }, //           [11,19) not bold -> excluded
          { text: "20", bold: true }, //     [19,21) lead (contains the year)
          { text: " card body." } //         [21,33) after the year -> excluded
        ]
      }
    ]);
    expect(detectCiteLeads(doc, cfg())).toEqual([
      { startIndex: 5, endIndex: 11 },
      { startIndex: 19, endIndex: 21 }
    ]);
  });

  it("clamps off the segment-final newline (the API refuses to style it)", () => {
    // The cite is the doc's LAST paragraph and fully bold: "Valcke 20\n" is
    // [5,15), but the segment-final newline at 14 can never be style-targeted.
    const doc = buildDoc([para("Tag", "HEADING_4"), { elements: [{ text: "Valcke 20", bold: true }] }]);
    expect(detectCiteLeads(doc, cfg())).toEqual([{ startIndex: 5, endIndex: 14 }]);
  });

  it("keeps a NON-segment-final trailing newline inside a fully-bold lead (benign, documented)", () => {
    const doc = buildDoc([
      para("Tag", "HEADING_4"),
      { elements: [{ text: "Valcke 20", bold: true }] }, // "Valcke 20\n" = [5,15)
      para("More body.") // pushes the cite off the segment end
    ]);
    expect(detectCiteLeads(doc, cfg())).toEqual([{ startIndex: 5, endIndex: 15 }]);
  });

  it("is deliberately NOT gated on structuralCite (the repair pass is Apply-styles parity)", () => {
    const doc = buildDoc([para("Tag", "HEADING_4"), nativeCite()]);
    expect(detectCiteLeads(doc, cfg({ structuralCite: false }))).toEqual([{ startIndex: 5, endIndex: 21 }]);
  });

  it("skips a lead already at/above the signature threshold (nothing to repair)", () => {
    const doc = buildDoc([
      para("Tag", "HEADING_4"),
      { elements: [{ text: "Valcke 20", bold: true, size: 14 }, { text: " body" }] }
    ]);
    expect(detectCiteLeads(doc, cfg())).toEqual([]);
  });

  it("detects nothing without a kept heading (no window ever opens)", () => {
    expect(detectCiteLeads(buildDoc([nativeCite()]), cfg())).toEqual([]);
  });

  it("never targets a deep mid-card bold mention — only the first candidate's lead", () => {
    const doc = buildDoc([
      para("Tag", "HEADING_4"), // [1,5)
      nativeCite(), // lead at [5,21)
      { elements: [{ text: "Body text later cites " }, { text: "Smith 2020", bold: true }, { text: " in passing." }] }
    ]);
    expect(detectCiteLeads(doc, cfg())).toEqual([{ startIndex: 5, endIndex: 21 }]);
  });

  it("ignores chips inside the candidate (A9 x A10): the lead range lands on the bold text only", () => {
    // A chip contributes NO visible text to the heuristic's offsets but DOES
    // occupy index space — the emitted range must use real doc indexes.
    const doc = buildDoc([
      para("Tag", "HEADING_4"), // [1,5)
      {
        elements: [
          { text: "@chip", kind: "other" }, // [5,10) — skipped by the heuristic
          { text: "Valcke 20", bold: true }, // [10,19) — the lead
          { text: " body" } // [19,25) with the appended newline
        ]
      }
    ]);
    expect(detectCiteLeads(doc, cfg())).toEqual([{ startIndex: 10, endIndex: 19 }]);
    expect(planKeeps(doc, cfg())[1].citeDetected).toBe(true);
  });

  it("agrees with planKeeps about WHICH paragraph is the structural cite (single-source detection)", () => {
    const doc = buildDoc([para("Tag", "HEADING_4"), nativeCite()]);
    const leads = detectCiteLeads(doc, cfg());
    const keeps = planKeeps(doc, cfg());
    // The lead range falls inside paragraph 1, and paragraph 1 is the one kept as a cite.
    expect(leads).toHaveLength(1);
    expect(leads[0].startIndex).toBeGreaterThanOrEqual(doc.paragraphs[1].startIndex);
    expect(leads[0].endIndex).toBeLessThanOrEqual(doc.paragraphs[1].endIndex);
    expect(keeps[1].citeDetected).toBe(true);
  });
});
