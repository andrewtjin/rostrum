// Styles-lane suite for google-docs/src/core/styles.ts (plan S10; A5/A10/A12/D13).
// The lane's two leak directions drive the suite:
//   * a TORN or PARTIAL emission (incomplete border object, named-style and
//     retro writes in one batch, a half-styled paragraph across a chunk
//     boundary) — pinned by completeness + batch-separation + grouping tests;
//   * a run-level fontSize write OUTSIDE heading paragraphs / cite leads,
//     which would collide with Hide's restore records (plan A5's invariant) —
//     pinned structurally over a mixed doc.
// No jest snapshots (house gdocs culture — a snapshot would bless a wrong
// emission); the canary is a hand-pinned numeric summary instead.

// Analytics-specific assertions (planAnalyticify, the off-palette color, encodeRgbColor)
// live in gdocsAnalytics.test.ts — this suite stays scoped to the debate-styles lane.
import {
  CITE_PT,
  DEFAULT_CITE_MIN_PT,
  STYLE_SIZES_PT
} from "../google-docs/src/core/constants";
import { detectCiteLeads } from "../google-docs/src/core/keepers";
import { DEFAULT_KEEP_HEXES } from "../google-docs/src/core/settings";
import { headingTextRanges, planApplyStyles, planMarkCite, pt } from "../google-docs/src/core/styles";
import {
  DocsRequest,
  GDoc,
  GdocsSettings,
  RequestGroup,
  UpdateNamedStyleRequest,
  UpdateParagraphStyleRequest,
  UpdateTextStyleRequest
} from "../google-docs/src/core/types";
import { buildDoc, GpSpec, para, r } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** Resolved settings with the engine's real defaults; override per test (the
 * literal, not resolveSettings, keeps this suite independent of settings
 * parsing — styles only ever sees a RESOLVED GdocsSettings). */
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

/** The Docs-NATIVE cite shape (bold author+year at the inherited 11pt) whose
 * repair to the convention is the whole point of plan A10(b). */
function nativeCite(): GpSpec {
  return {
    elements: [
      { text: "Valcke et al. 20", bold: true },
      { text: " The card body argues the impact." }
    ]
  };
}

/** Narrowed inner payloads of every request of one kind across groups. */
function textUpdates(groups: RequestGroup[]): UpdateTextStyleRequest["updateTextStyle"][] {
  const out: UpdateTextStyleRequest["updateTextStyle"][] = [];
  for (const g of groups) for (const q of g.requests) if ("updateTextStyle" in q) out.push(q.updateTextStyle);
  return out;
}
function paraUpdates(groups: RequestGroup[]): UpdateParagraphStyleRequest["updateParagraphStyle"][] {
  const out: UpdateParagraphStyleRequest["updateParagraphStyle"][] = [];
  for (const g of groups)
    for (const q of g.requests) if ("updateParagraphStyle" in q) out.push(q.updateParagraphStyle);
  return out;
}
function namedUpdates(groups: RequestGroup[]): UpdateNamedStyleRequest["updateNamedStyle"][] {
  const out: UpdateNamedStyleRequest["updateNamedStyle"][] = [];
  for (const g of groups) for (const q of g.requests) if ("updateNamedStyle" in q) out.push(q.updateNamedStyle);
  return out;
}

/** Compact request signature for group-shape assertions: spacing and border
 * paragraph writes are distinguished by their (exact, pinned) field masks. */
function reqKind(q: DocsRequest): string {
  if ("updateTextStyle" in q) return "text";
  if ("updateParagraphStyle" in q)
    return q.updateParagraphStyle.fields === "spaceAbove,spaceBelow" ? "spacing" : "border";
  if ("updateNamedStyle" in q) return "named";
  return "OTHER"; // create/deleteNamedRange would be a contract violation here
}

/** Smallest range start inside a group — the descending-emission anchor. */
function groupMinStart(g: RequestGroup): number {
  let min = Number.POSITIVE_INFINITY;
  for (const q of g.requests) {
    if ("updateTextStyle" in q) min = Math.min(min, q.updateTextStyle.range.startIndex);
    if ("updateParagraphStyle" in q) min = Math.min(min, q.updateParagraphStyle.range.startIndex);
  }
  return min;
}

/** One complete pocket-box side, exactly as the API demands it (0.5pt black
 * solid, 1pt padding — Word parity; black is the empty rgbColor, the omitted-
 * zero-channels wire shape the fixture-lint pins for reads). */
const BORDER_SIDE = {
  color: { color: { rgbColor: {} } },
  width: { magnitude: 0.5, unit: "PT" },
  padding: { magnitude: 1, unit: "PT" },
  dashStyle: "SOLID"
};
const SIDES = ["borderTop", "borderBottom", "borderLeft", "borderRight"] as const;

/**
 * The mixed doc most of the suite operates on: every debate heading once, a
 * TITLE and an H5 (outside the debate ladder, untouched), a structural-slot
 * Docs-native cite under the tag, two direct-spacing carriers, a table H1
 * (untouched territory), and a plain segment-final body line.
 */
function mixedDoc(): GDoc {
  return buildDoc([
    para("Aff Case", "TITLE"), // p0: not a debate style
    { style: "HEADING_1", elements: [{ text: "Pocket A" }] }, // p1: border + text
    { style: "HEADING_2", elements: [{ text: "Hat" }], spaceAbovePt: 10 }, // p2: text + spacing
    para("Block line", "HEADING_3"), // p3: text
    para("Tag line", "HEADING_4"), // p4: text (opens the cite window)
    nativeCite(), // p5: cite repair
    { elements: [{ text: "card body prose" }], spaceBelowPt: 6 }, // p6: spacing only
    { style: "HEADING_1", inTable: true, elements: [{ text: "boxed cell" }] }, // p7: skipped
    para("analysis", "HEADING_5"), // p8: body style, untouched
    para("closing body") // p9: plain + segment-final
  ]);
}

// ---------------------------------------------------------------------------
// 1. Batch 1 — the named-style redefinitions
// ---------------------------------------------------------------------------

describe("planApplyStyles — named-style batch (plan A5/D13)", () => {
  const plan = planApplyStyles(mixedDoc(), cfg());
  const named = namedUpdates(plan.namedStyleGroups);

  it("emits exactly five single-request groups, Normal first then H1-4", () => {
    expect(plan.namedStyleGroups).toHaveLength(5);
    for (const g of plan.namedStyleGroups) expect(g.requests.map(reqKind)).toEqual(["named"]);
    expect(named.map((u) => u.namedStyle.namedStyleType)).toEqual([
      "NORMAL_TEXT",
      "HEADING_1",
      "HEADING_2",
      "HEADING_3",
      "HEADING_4"
    ]);
  });

  it("NORMAL_TEXT = 11pt + zeroed spacing (the condensation dependency, A12)", () => {
    expect(named[0]).toEqual({
      namedStyle: {
        namedStyleType: "NORMAL_TEXT",
        textStyle: { fontSize: { magnitude: 11, unit: "PT" } },
        paragraphStyle: {
          spaceAbove: { magnitude: 0, unit: "PT" },
          spaceBelow: { magnitude: 0, unit: "PT" }
        }
      },
      fields: "textStyle.fontSize,paragraphStyle.spaceAbove,paragraphStyle.spaceBelow"
    });
  });

  it("HEADING_1 = pocket 26 bold + the COMPLETE four-side box, all sides in the mask", () => {
    const h1 = named[1];
    expect(h1.namedStyle.textStyle).toEqual({ fontSize: { magnitude: 26, unit: "PT" }, bold: true });
    for (const side of SIDES) expect(h1.namedStyle.paragraphStyle?.[side]).toEqual(BORDER_SIDE);
    expect(h1.fields).toBe(
      "textStyle.fontSize,textStyle.bold," +
        "paragraphStyle.borderTop,paragraphStyle.borderBottom," +
        "paragraphStyle.borderLeft,paragraphStyle.borderRight"
    );
  });

  it("HEADING_2/3 = size only (no bold key — the author's emphasis survives)", () => {
    expect(named[2].namedStyle.textStyle).toEqual({ fontSize: { magnitude: 22, unit: "PT" } });
    expect(named[2].fields).toBe("textStyle.fontSize");
    expect(named[3].namedStyle.textStyle).toEqual({ fontSize: { magnitude: 16, unit: "PT" } });
    expect(named[3].fields).toBe("textStyle.fontSize");
  });

  it("HEADING_4 = tag 14 bold", () => {
    expect(named[4].namedStyle.textStyle).toEqual({ fontSize: { magnitude: 14, unit: "PT" }, bold: true });
    expect(named[4].fields).toBe("textStyle.fontSize,textStyle.bold");
  });

  it("sizes are single-sourced from STYLE_SIZES_PT (the D-table, A11.iii)", () => {
    expect(named.map((u) => u.namedStyle.textStyle?.fontSize?.magnitude)).toEqual([
      STYLE_SIZES_PT.normal,
      STYLE_SIZES_PT.pocket,
      STYLE_SIZES_PT.hat,
      STYLE_SIZES_PT.block,
      STYLE_SIZES_PT.tag
    ]);
  });

  it("the batch is doc-independent: a zero-heading doc gets the identical five writes", () => {
    const bare = planApplyStyles(buildDoc([para("just prose")]), cfg());
    expect(bare.namedStyleGroups).toEqual(plan.namedStyleGroups);
  });
});

// ---------------------------------------------------------------------------
// 2. Batch separation (a consumer-account 400 on batch 1 must not kill batch 2)
// ---------------------------------------------------------------------------

describe("planApplyStyles — batch separation (plan A5/D13)", () => {
  it("batch 1 carries ONLY updateNamedStyle; batch 2 carries NONE", () => {
    const plan = planApplyStyles(mixedDoc(), cfg());
    for (const g of plan.namedStyleGroups)
      for (const q of g.requests) expect("updateNamedStyle" in q).toBe(true);
    for (const g of plan.retroGroups)
      for (const q of g.requests) expect("updateNamedStyle" in q).toBe(false);
  });

  it("only the three styles-lane request types are ever emitted (case 001-F1)", () => {
    const plan = planApplyStyles(mixedDoc(), cfg());
    for (const g of [...plan.namedStyleGroups, ...plan.retroGroups])
      for (const q of g.requests) expect(["named", "text", "spacing", "border"]).toContain(reqKind(q));
  });
});

// ---------------------------------------------------------------------------
// 3. Batch 2 — retro restyle of existing headings
// ---------------------------------------------------------------------------

describe("planApplyStyles — retro heading pass (the import MAINLINE, plan A5)", () => {
  it("restyles one paragraph per debate style in the mixed doc, counted per style", () => {
    const plan = planApplyStyles(mixedDoc(), cfg());
    expect(plan.counts.restyled).toEqual({ pocket: 1, hat: 1, block: 1, tag: 1 });
  });

  it("H1 retro = complete border + bold 26pt text over the exact paragraph range", () => {
    // "Pocket\n" = [1,8); a body paragraph after keeps it off the segment end.
    const plan = planApplyStyles(buildDoc([para("Pocket", "HEADING_1"), para("body")]), cfg());
    expect(plan.retroGroups).toHaveLength(1);
    expect(plan.retroGroups[0].requests).toEqual([
      {
        updateParagraphStyle: {
          range: r(1, 8),
          paragraphStyle: {
            borderTop: BORDER_SIDE,
            borderBottom: BORDER_SIDE,
            borderLeft: BORDER_SIDE,
            borderRight: BORDER_SIDE
          },
          fields: "borderTop,borderBottom,borderLeft,borderRight"
        }
      },
      {
        updateTextStyle: {
          range: r(1, 8),
          textStyle: { bold: true, fontSize: { magnitude: 26, unit: "PT" } },
          fields: "bold,fontSize"
        }
      }
    ]);
    expect(plan.counts.restyled.pocket).toBe(1);
  });

  it("H2 retro = size only, bold absent from style AND mask (emphasis survives)", () => {
    const plan = planApplyStyles(buildDoc([para("Hat", "HEADING_2"), para("body")]), cfg());
    expect(textUpdates(plan.retroGroups)).toEqual([
      { range: r(1, 5), textStyle: { fontSize: { magnitude: 22, unit: "PT" } }, fields: "fontSize" }
    ]);
  });

  it("H4 retro = bold 14pt (tag is a bold debate style)", () => {
    const plan = planApplyStyles(buildDoc([para("Tagline", "HEADING_4"), para("body")]), cfg());
    expect(textUpdates(plan.retroGroups)).toEqual([
      {
        range: r(1, 9),
        textStyle: { bold: true, fontSize: { magnitude: 14, unit: "PT" } },
        fields: "bold,fontSize"
      }
    ]);
  });

  it("clamps the segment-final newline off BOTH the text and the border range", () => {
    // "Pocket\n" = [1,8) but the doc-final newline can never be style-targeted.
    const plan = planApplyStyles(buildDoc([para("Pocket", "HEADING_1")]), cfg());
    expect(paraUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 7)]);
    expect(textUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 7)]);
  });

  it("an EMPTY segment-final heading is skipped entirely (clamp leaves nothing to target)", () => {
    const plan = planApplyStyles(buildDoc([para("", "HEADING_1")]), cfg());
    expect(plan.retroGroups).toEqual([]);
    expect(plan.counts.restyled.pocket).toBe(0);
  });

  it("text spans BREAK around a chip inside a heading (whitelist A9); the border spans the paragraph", () => {
    // "Big " [1,5) + chip [5,10) + "Deal\n" [10,15); border covers [1,15).
    const plan = planApplyStyles(
      buildDoc([
        { style: "HEADING_1", elements: [{ text: "Big " }, { text: "@chip", kind: "other" }, { text: "Deal" }] },
        para("after")
      ]),
      cfg()
    );
    expect(paraUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 15)]);
    expect(textUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 5), r(10, 15)]);
    expect(plan.counts.restyled.pocket).toBe(1); // still ONE paragraph in the receipt
  });

  it("ADJACENT text elements in a heading coalesce into ONE write (a bolded word splits runs)", () => {
    // "Big " [1,5) + "Deal\n" [5,10) are contiguous text — one range, not two
    // (fewer requests, and the API would accept either; coalescing is the
    // maximal-span counterpart of breaking around chips).
    const plan = planApplyStyles(
      buildDoc([
        { style: "HEADING_1", elements: [{ text: "Big " }, { text: "Deal", bold: true }] },
        para("after")
      ]),
      cfg()
    );
    expect(textUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 10)]);
    expect(paraUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 10)]);
  });

  it("a segment-final heading ENDING IN A CHIP drops the clamped lone-newline span, keeps the rest", () => {
    // "Big" [1,4) + chip [4,9) + the newline as its own text run [9,10) — the
    // doc-final newline clamps that last run to nothing, so only the real
    // text gets restyled; the border still wraps the clamped paragraph.
    const plan = planApplyStyles(
      buildDoc([{ style: "HEADING_1", elements: [{ text: "Big" }, { text: "@chip", kind: "other" }] }]),
      cfg()
    );
    expect(textUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 4)]);
    expect(paraUpdates(plan.retroGroups).map((u) => u.range)).toEqual([r(1, 9)]);
    expect(plan.counts.restyled.pocket).toBe(1);
  });

  it("TITLE / SUBTITLE / H5 / H6 are outside the debate ladder — never retro-restyled", () => {
    const plan = planApplyStyles(
      buildDoc([
        para("Aff Case", "TITLE"),
        para("v3", "SUBTITLE"),
        para("analysis", "HEADING_5"),
        para("note", "HEADING_6")
      ]),
      cfg()
    );
    expect(plan.retroGroups).toEqual([]);
    expect(plan.counts.restyled).toEqual({ pocket: 0, hat: 0, block: 0, tag: 0 });
  });

  it("a heading inside a table is untouched (row 1 parity)", () => {
    const plan = planApplyStyles(
      buildDoc([{ style: "HEADING_1", inTable: true, elements: [{ text: "cell" }] }, para("x")]),
      cfg()
    );
    expect(plan.retroGroups).toEqual([]);
    expect(plan.counts.restyled.pocket).toBe(0);
  });
});

// ---------------------------------------------------------------------------
// 4. Batch 2 — direct-spacing clear
// ---------------------------------------------------------------------------

describe("planApplyStyles — direct-spacing clear (imports carry direct spacing, plan A5)", () => {
  it("targets ONLY paragraphs carrying a direct value — explicit zero included, inherit skipped", () => {
    const plan = planApplyStyles(
      buildDoc([
        para("no direct"), // null/null -> untouched
        { elements: [{ text: "above ten" }], spaceAbovePt: 10 }, // [11,21)
        // An explicit direct 0 still SHADOWS the named style (parse.ts keeps
        // the 0-vs-null distinction for exactly this), so it is re-pinned.
        { elements: [{ text: "explicit zero" }], spaceBelowPt: 0 }, // [21,35)
        para("tail") // null/null + segment-final -> untouched
      ]),
      cfg()
    );
    expect(plan.counts.spacingCleared).toBe(2);
    // Descending-start emission: the later paragraph's clear comes first.
    expect(paraUpdates(plan.retroGroups)).toEqual([
      {
        range: r(21, 35),
        paragraphStyle: {
          spaceAbove: { magnitude: 0, unit: "PT" },
          spaceBelow: { magnitude: 0, unit: "PT" }
        },
        fields: "spaceAbove,spaceBelow"
      },
      {
        range: r(11, 21),
        paragraphStyle: {
          spaceAbove: { magnitude: 0, unit: "PT" },
          spaceBelow: { magnitude: 0, unit: "PT" }
        },
        fields: "spaceAbove,spaceBelow"
      }
    ]);
  });

  it("never touches a table paragraph's spacing", () => {
    const plan = planApplyStyles(
      buildDoc([{ inTable: true, spaceAbovePt: 12, elements: [{ text: "cell" }] }, para("x")]),
      cfg()
    );
    expect(plan.counts.spacingCleared).toBe(0);
    expect(plan.retroGroups).toEqual([]);
  });

  it("a heading carrying direct spacing gets restyle AND clear in ONE atomic group", () => {
    const plan = planApplyStyles(
      buildDoc([{ style: "HEADING_2", elements: [{ text: "Hat" }], spaceAbovePt: 10 }, para("body")]),
      cfg()
    );
    // One group per touched paragraph: a torn chunk can never half-style it.
    expect(plan.retroGroups).toHaveLength(1);
    expect(plan.retroGroups[0].requests.map(reqKind)).toEqual(["text", "spacing"]);
    expect(plan.counts.restyled.hat).toBe(1);
    expect(plan.counts.spacingCleared).toBe(1);
  });
});

// ---------------------------------------------------------------------------
// 5. Batch 2 — cite repair (plan A10(b))
// ---------------------------------------------------------------------------

describe("planApplyStyles — cite repair flows from detectCiteLeads", () => {
  it("writes the convention (default-black, bold, CITE_PT, fields 'foregroundColor,bold,fontSize') onto the exact lead ranges", () => {
    const doc = buildDoc([para("Tag", "HEADING_4"), nativeCite(), para("after")]);
    const leads = detectCiteLeads(doc, cfg());
    expect(leads).toEqual([r(5, 21)]); // precondition: the detection itself
    const plan = planApplyStyles(doc, cfg());
    const citeWrites = textUpdates(plan.retroGroups).filter((u) => u.range.startIndex === 5);
    expect(citeWrites).toEqual([
      {
        range: r(5, 21),
        // 003-F9: the cite convention forces foreground to default-black (proto3
        // empty rgbColor {}) so marking analytics text as a cite CLEARS the navy
        // signature — the cite renders black and survives Delete analytics.
        textStyle: { foregroundColor: { color: { rgbColor: {} } }, bold: true, fontSize: { magnitude: CITE_PT, unit: "PT" } },
        fields: "foregroundColor,bold,fontSize"
      }
    ]);
    expect(plan.counts.citesRepaired).toBe(1);
  });

  it("a split author/year lead = TWO writes in ONE group, counted as ONE cite (receipt unit)", () => {
    const doc = buildDoc([
      para("Tag", "HEADING_4"), // [1,5)
      {
        elements: [
          { text: "Valcke", bold: true }, // [5,11) lead
          { text: " et al. " }, // [11,19) not bold
          { text: "20", bold: true }, // [19,21) lead (the year)
          { text: " card body." }
        ]
      }
    ]);
    const plan = planApplyStyles(doc, cfg());
    expect(plan.counts.citesRepaired).toBe(1);
    // Descending anchors: the cite group (anchor 5) precedes the tag group (anchor 1).
    expect(plan.retroGroups.map((g) => g.requests.map(reqKind))).toEqual([["text", "text"], ["text"]]);
    expect(textUpdates([plan.retroGroups[0]]).map((u) => u.range)).toEqual([r(5, 11), r(19, 21)]);
  });

  it("repair is NOT gated on structuralCite (Apply-styles parity — the keeper flag is separate)", () => {
    const doc = buildDoc([para("Tag", "HEADING_4"), nativeCite(), para("after")]);
    const plan = planApplyStyles(doc, cfg({ structuralCite: false }));
    expect(plan.counts.citesRepaired).toBe(1);
  });

  it("a lead already at the signature threshold is NOT re-repaired (gate (a) — repair stays minimal)", () => {
    const doc = buildDoc([
      para("Tag", "HEADING_4"),
      { elements: [{ text: "Valcke 20", bold: true, size: 14 }, { text: " body" }] }
    ]);
    const plan = planApplyStyles(doc, cfg());
    expect(plan.counts.citesRepaired).toBe(0);
    // The only retro write left is the tag heading's own restyle.
    expect(plan.retroGroups.map((g) => g.requests.map(reqKind))).toEqual([["text"]]);
  });

  it("no kept heading, no cite window — a cite-shaped line in plain prose is never touched", () => {
    const plan = planApplyStyles(buildDoc([para("intro prose"), nativeCite()]), cfg());
    expect(plan.counts.citesRepaired).toBe(0);
    expect(plan.retroGroups).toEqual([]);
  });
});

// ---------------------------------------------------------------------------
// 6. Whole-plan invariants over the mixed doc
// ---------------------------------------------------------------------------

describe("planApplyStyles — emission invariants (plan A5/A11.iii)", () => {
  const doc = mixedDoc();
  const settings = cfg();
  const plan = planApplyStyles(doc, settings);

  it("INVARIANT: run-level fontSize lands ONLY inside H1-4 paragraphs or cite-lead ranges", () => {
    const leads = detectCiteLeads(doc, settings);
    const headingParas = doc.paragraphs.filter(
      (p) => !p.inTable && ["HEADING_1", "HEADING_2", "HEADING_3", "HEADING_4"].includes(p.namedStyleType)
    );
    const writes = textUpdates(plan.retroGroups);
    expect(writes.length).toBeGreaterThan(0); // the walk must not be vacuous
    for (const u of writes) {
      expect(u.fields).toContain("fontSize"); // every text write here sizes
      const inHeading = headingParas.some(
        (p) => u.range.startIndex >= p.startIndex && u.range.endIndex <= p.endIndex
      );
      const isLead = leads.some(
        (l) => l.startIndex === u.range.startIndex && l.endIndex === u.range.endIndex
      );
      expect(inHeading || isLead).toBe(true);
    }
  });

  it("every emitted size is a D-table size or CITE_PT — nothing else can reach a run", () => {
    const allowed = new Set<number>([
      STYLE_SIZES_PT.pocket,
      STYLE_SIZES_PT.hat,
      STYLE_SIZES_PT.block,
      STYLE_SIZES_PT.tag,
      CITE_PT
    ]);
    for (const u of textUpdates(plan.retroGroups)) {
      expect(u.textStyle.fontSize).toBeDefined();
      expect(allowed.has(u.textStyle.fontSize!.magnitude)).toBe(true);
    }
  });

  it("EVERY border emission (named-style AND retro) carries all four complete sides", () => {
    const borderCarriers = [
      ...namedUpdates(plan.namedStyleGroups)
        .filter((u) => u.fields.includes("border"))
        .map((u) => u.namedStyle.paragraphStyle),
      ...paraUpdates(plan.retroGroups)
        .filter((u) => u.fields.includes("border"))
        .map((u) => u.paragraphStyle)
    ];
    // Exactly two in the mixed doc: the H1 named style + p1's retro box (the
    // table H1 is skipped). A partial border update would be API-rejected.
    expect(borderCarriers).toHaveLength(2);
    for (const ps of borderCarriers) for (const side of SIDES) expect(ps?.[side]).toEqual(BORDER_SIDE);
    // Fresh objects per emission — no aliasing between batches.
    expect(borderCarriers[0]?.borderTop).not.toBe(borderCarriers[1]?.borderTop);
  });

  it("groups are per-paragraph/per-cite atomic units, emitted in descending start order", () => {
    // p6 spacing > p5 cite > p4 tag > p3 block > p2 hat(+spacing) > p1 pocket(boxed).
    expect(plan.retroGroups.map((g) => g.requests.map(reqKind))).toEqual([
      ["spacing"],
      ["text"],
      ["text"],
      ["text"],
      ["text", "spacing"],
      ["border", "text"]
    ]);
    const anchors = plan.retroGroups.map(groupMinStart);
    for (let i = 1; i < anchors.length; i++) expect(anchors[i]).toBeLessThan(anchors[i - 1]);
  });

  it("no group is ever empty (the chunker contract expects substance per group)", () => {
    for (const g of [...plan.namedStyleGroups, ...plan.retroGroups])
      expect(g.requests.length).toBeGreaterThan(0);
  });

  it("NUMERIC CANARY: hand-pinned emission totals for the mixed doc", () => {
    // Deliberately literal — any change to what Apply-styles emits must come
    // here and justify itself (the no-snapshot culture's tripwire).
    const summary = {
      namedStyleGroups: plan.namedStyleGroups.length,
      retroGroups: plan.retroGroups.length,
      retroRequests: plan.retroGroups.reduce((n, g) => n + g.requests.length, 0),
      retroTextWrites: textUpdates(plan.retroGroups).length,
      retroParagraphWrites: paraUpdates(plan.retroGroups).length,
      textSizesSorted: textUpdates(plan.retroGroups)
        .map((u) => u.textStyle.fontSize!.magnitude)
        .sort((a, b) => a - b),
      counts: plan.counts
    };
    expect(summary).toEqual({
      namedStyleGroups: 5,
      retroGroups: 6,
      retroRequests: 8,
      retroTextWrites: 5, // 4 heading restyles + 1 cite repair
      retroParagraphWrites: 3, // 1 pocket box + 2 spacing clears
      textSizesSorted: [14, 14, 16, 22, 26], // tag + cite, block, hat, pocket
      counts: {
        restyled: { pocket: 1, hat: 1, block: 1, tag: 1 },
        spacingCleared: 2,
        citesRepaired: 1
      }
    });
  });
});

// ---------------------------------------------------------------------------
// 7. Zero-work and idempotence
// ---------------------------------------------------------------------------

describe("planApplyStyles — zero-heading docs and idempotent re-apply", () => {
  it("a zero-heading doc plans the named batch only, with all-zero counts", () => {
    const plan = planApplyStyles(buildDoc([para("just prose"), para("more prose")]), cfg());
    expect(plan.namedStyleGroups).toHaveLength(5);
    expect(plan.retroGroups).toEqual([]);
    expect(plan.counts).toEqual({
      restyled: { pocket: 0, hat: 0, block: 0, tag: 0 },
      spacingCleared: 0,
      citesRepaired: 0
    });
  });

  it("an EMPTY doc behaves the same (no paragraphs at all)", () => {
    const plan = planApplyStyles(buildDoc([]), cfg());
    expect(plan.namedStyleGroups).toHaveLength(5);
    expect(plan.retroGroups).toEqual([]);
    expect(plan.counts.citesRepaired).toBe(0);
  });

  /**
   * The mixed doc AS APPLY-STYLES LEAVES IT: headings carry direct size/bold,
   * cleared paragraphs carry direct 0/0 (the clear WRITES zeros — parse reads
   * them back as 0, not null), and the cite lead is at the convention.
   * Identical text, so identical indexes.
   */
  function styledMixedDoc(): GDoc {
    return buildDoc([
      para("Aff Case", "TITLE"),
      { style: "HEADING_1", elements: [{ text: "Pocket A", size: 26, bold: true }] },
      { style: "HEADING_2", elements: [{ text: "Hat", size: 22 }], spaceAbovePt: 0, spaceBelowPt: 0 },
      { style: "HEADING_3", elements: [{ text: "Block line", size: 16 }] },
      { style: "HEADING_4", elements: [{ text: "Tag line", size: 14, bold: true }] },
      {
        elements: [
          { text: "Valcke et al. 20", bold: true, size: CITE_PT }, // repaired lead
          { text: " The card body argues the impact." }
        ]
      },
      { elements: [{ text: "card body prose" }], spaceAbovePt: 0, spaceBelowPt: 0 },
      { style: "HEADING_1", inTable: true, elements: [{ text: "boxed cell" }] },
      para("analysis", "HEADING_5"),
      para("closing body")
    ]);
  }

  it("re-apply emits the SAME heading/spacing writes (safe: absolute-value no-ops) and ZERO cite writes (converged)", () => {
    // WHY same-writes is safe: every retro write is an absolute value (26pt,
    // bold, 0pt spacing, a complete border) keyed off facts the writes
    // themselves preserve — namedStyleType never changes, and a cleared
    // paragraph still CARRIES direct spacing (as 0). Re-applying each write
    // is a server-side no-op, and no emission inserts/deletes content
    // (001-F1), so repetition can never shift an index. The cite pass instead
    // CONVERGES: the repaired lead now meets the signature, so gate (a) skips
    // it and the second run emits nothing for it.
    const first = planApplyStyles(mixedDoc(), cfg());
    const second = planApplyStyles(styledMixedDoc(), cfg());

    expect(second.namedStyleGroups).toEqual(first.namedStyleGroups);

    // The cite group is the one whose write targets the detected lead range.
    const lead = detectCiteLeads(mixedDoc(), cfg())[0];
    const isCiteGroup = (g: RequestGroup): boolean =>
      g.requests.some(
        (q) =>
          "updateTextStyle" in q &&
          q.updateTextStyle.range.startIndex === lead.startIndex &&
          q.updateTextStyle.range.endIndex === lead.endIndex
      );
    expect(first.retroGroups.filter(isCiteGroup)).toHaveLength(1); // precondition
    expect(second.retroGroups).toEqual(first.retroGroups.filter((g) => !isCiteGroup(g)));

    expect(second.counts).toEqual({
      restyled: { pocket: 1, hat: 1, block: 1, tag: 1 },
      spacingCleared: 2,
      citesRepaired: 0
    });
  });
});

// ---------------------------------------------------------------------------
// 8. planMarkCite
// ---------------------------------------------------------------------------

describe("planMarkCite — the cite convention over a selection", () => {
  it("emits exactly one default-black+bold+CITE_PT write over the given range", () => {
    expect(planMarkCite(r(5, 20))).toEqual({
      requests: [
        {
          updateTextStyle: {
            range: { startIndex: 5, endIndex: 20 },
            // 003-F9: foreground default-black clears any analytics navy so a
            // cited selection turns black and is no longer swept by Delete analytics.
            textStyle: { foregroundColor: { color: { rgbColor: {} } }, bold: true, fontSize: { magnitude: CITE_PT, unit: "PT" } },
            fields: "foregroundColor,bold,fontSize"
          }
        }
      ]
    });
  });

  it("copies the selection — the caller's range object is never aliased into the request", () => {
    const selection = r(7, 12);
    const group = planMarkCite(selection);
    const q = group.requests[0];
    if (!("updateTextStyle" in q)) throw new Error("expected an updateTextStyle request");
    expect(q.updateTextStyle.range).toEqual(selection);
    expect(q.updateTextStyle.range).not.toBe(selection);
  });

  it("matches the bulk cite-repair emission shape exactly (single-source convention)", () => {
    // Mark-cite on the lead range and Apply-styles' repair of the same lead
    // must be byte-identical requests — drift here was a named plan-review
    // finding (the Mark-cite/keeper split).
    const doc = buildDoc([para("Tag", "HEADING_4"), nativeCite(), para("after")]);
    const lead = detectCiteLeads(doc, cfg())[0];
    const repair = textUpdates(planApplyStyles(doc, cfg()).retroGroups).find(
      (u) => u.range.startIndex === lead.startIndex
    );
    const marked = planMarkCite(lead).requests[0];
    if (!("updateTextStyle" in marked)) throw new Error("expected an updateTextStyle request");
    expect(marked.updateTextStyle).toEqual(repair);
  });
});

// ---------------------------------------------------------------------------
// 9. Newly-exported span primitives (Loop 003) — planAnalyticify reuses these
//    rather than re-deriving them, so the exports are part of the contract.
// ---------------------------------------------------------------------------

describe("exported pt / headingTextRanges (the reuse contract, plan A5/A9)", () => {
  it("pt builds the Docs dimension literal analyticify and the size writes share", () => {
    expect(pt(14)).toEqual({ magnitude: 14, unit: "PT" });
  });

  it("headingTextRanges over a multi-run line coalesces adjacent text into ONE span", () => {
    // "Big " [1,5) + "Deal\n" [5,10) are contiguous text — the maximal span is
    // [1,10), the same geometry the heading retro pass and analyticify consume.
    const doc = buildDoc([
      { elements: [{ text: "Big " }, { text: "Deal", bold: true }] },
      para("after")
    ]);
    expect(headingTextRanges(doc.paragraphs[0])).toEqual([r(1, 10)]);
  });

  it("headingTextRanges BREAKS the span around a chip (A9 whitelist) and clamps the final newline", () => {
    // "Big" [1,4) + chip [4,9) + "Deal\n" [9,14) on the segment-final paragraph:
    // the chip breaks the run, and the doc-final newline clamps the tail to [9,13).
    const doc = buildDoc([
      { elements: [{ text: "Big" }, { text: "@chip", kind: "other" }, { text: "Deal" }] }
    ]);
    expect(headingTextRanges(doc.paragraphs[0])).toEqual([r(1, 4), r(9, 13)]);
  });
});
