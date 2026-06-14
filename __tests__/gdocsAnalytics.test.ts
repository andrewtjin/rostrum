// Analytics-lane suite (Loop 003): the OFF-PALETTE color guarantee (003-S4) and
// the analytic-ify planner (core/styles.ts planAnalyticify). The two halves
// share one fixture vocabulary (the gdocsBuilders `fg` field models an
// analytic-ify'd run) but pin two distinct contracts:
//   * the color is a reliable SIGNATURE — it matches NO one-click Docs swatch,
//     so analytic-ify is the only producer of it (the keeper/delete predicate
//     can trust an exact match);
//   * analytic-ify is PURE character formatting applied WHOLE-PARAGRAPH — it
//     emits only updateTextStyle (003-F2), reuses the proven A9/clamp span
//     geometry, skips tables, and is idempotent (003-S5).
// No jest snapshots (house gdocs culture); request shapes are hand-pinned.

import { encodeRgbColor } from "../google-docs/src/core/color";
import { ANALYTICS_FG_HEX, ANALYTICS_PT } from "../google-docs/src/core/constants";
import { DEFAULT_KEEP_HEXES } from "../google-docs/src/core/settings";
import { planAnalyticify } from "../google-docs/src/core/styles";
import { DocsRequest, RequestGroup } from "../google-docs/src/core/types";
import { buildDoc, GpSpec, para, r } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Local helpers
// ---------------------------------------------------------------------------

/** The narrowed updateTextStyle payloads across all groups, in emission order. */
function textWrites(groups: RequestGroup[]): { range: { startIndex: number; endIndex: number }; textStyle: unknown; fields: string }[] {
  const out: { range: { startIndex: number; endIndex: number }; textStyle: unknown; fields: string }[] = [];
  for (const g of groups) for (const q of g.requests) if ("updateTextStyle" in q) out.push(q.updateTextStyle);
  return out;
}

/** The request-kind tag for the emit-only-updateTextStyle audit (003-F2). */
function reqKind(q: DocsRequest): string {
  if ("updateTextStyle" in q) return "text";
  if ("updateParagraphStyle" in q) return "paragraph";
  if ("updateNamedStyle" in q) return "named";
  if ("deleteContentRange" in q) return "delete";
  if ("createNamedRange" in q) return "createRange";
  if ("deleteNamedRange" in q) return "deleteRange";
  return "OTHER";
}

/** The exact textStyle every analytic-ify write must carry: the off-palette
 * navy (through the shared encoder) + 14pt. Pinned ONCE so each assertion
 * proves the same single-sourced value. */
const ANALYTICS_TEXT_STYLE = {
  foregroundColor: encodeRgbColor(ANALYTICS_FG_HEX),
  fontSize: { magnitude: ANALYTICS_PT, unit: "PT" }
};

/** Ordinal set sugar: the paragraphs the adapter "touched". */
function touch(...ordinals: number[]): ReadonlySet<number> {
  return new Set(ordinals);
}

// ---------------------------------------------------------------------------
// 1. Off-palette color guarantee (003-S4)
// ---------------------------------------------------------------------------

describe("ANALYTICS_FG_HEX is off-palette (003-S4)", () => {
  // The off-palette claim is an EMPIRICAL one (R6-confirmed on the live color
  // picker), not a mathematical invariant: a user who manually TYPES #0b5396
  // into the custom-color box WOULD be treated as analytics (an accepted,
  // documented edge). These assertions pin the one thing code CAN guarantee —
  // the navy is not in the source-of-truth keep set the engine ships with, and
  // its genuine near-miss IS — so the signature can't silently start matching a
  // one-click swatch if the palette is ever edited.

  it("the navy is NOT in DEFAULT_KEEP_HEXES (the shipped source-of-truth set)", () => {
    // Asserted against the EXISTING constant, never a re-typed palette fixture:
    // if someone adds #0b5396 to the palette the signature collapses, and this
    // is the tripwire.
    expect(DEFAULT_KEEP_HEXES.has(ANALYTICS_FG_HEX)).toBe(false);
  });

  it("its genuine near-miss 'dark blue 2' (#0b5394) IS in DEFAULT_KEEP_HEXES", () => {
    // The real palette swatch the navy is nudged off of (blue 148 -> 150). A
    // user picking the real dark blue 2 must NOT be read as analytics, which is
    // only true if #0b5394 is a kept palette color and #0b5396 is not — the two
    // assertions are a matched pair.
    expect(DEFAULT_KEEP_HEXES.has("#0b5394")).toBe(true);
    expect(ANALYTICS_FG_HEX).not.toBe("#0b5394");
  });

  it("the navy is not any of the 10 grayscale swatches (the picker's gray column)", () => {
    // The Docs picker's grayscale column is deliberately NOT in the CHROMATIC
    // palette (web-paste shading lives there and must hide). We assert only that
    // the navy collides with none of the 10 grays — NOT that every gray is
    // un-kept: black (#000000) is a real Word highlight choice and IS in
    // DEFAULT_KEEP_HEXES, so asserting the column's absence from the keep set
    // would be false. The signature claim is just: the navy isn't a gray swatch.
    const GRAYSCALE_SWATCHES = [
      "#000000", "#434343", "#666666", "#999999", "#b7b7b7",
      "#cccccc", "#d9d9d9", "#efefef", "#f3f3f3", "#ffffff"
    ];
    expect(GRAYSCALE_SWATCHES).not.toContain(ANALYTICS_FG_HEX);
  });
});

// ---------------------------------------------------------------------------
// 2. planAnalyticify — the whole-paragraph navy-14pt planner
// ---------------------------------------------------------------------------

describe("planAnalyticify — whole-paragraph character formatting (plan §3 styles.ts)", () => {
  it("styles a touched paragraph's text WHOLE, coalescing multi-run spans into one write", () => {
    // "Made" [1,5) + " these claims\n" [5,19) are contiguous text -> ONE span
    // [1,19) (" these claims" is 13 chars + the appended newline); a following
    // paragraph keeps p0 off the segment end (no clamp).
    const doc = buildDoc([
      { elements: [{ text: "Made" }, { text: " these claims", bold: true }] },
      para("untouched body")
    ]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch(0));
    expect(paragraphsStyled).toBe(1);
    expect(groups).toHaveLength(1); // one RequestGroup per paragraph
    expect(groups[0].requests).toEqual([
      { updateTextStyle: { range: r(1, 19), textStyle: ANALYTICS_TEXT_STYLE, fields: "foregroundColor,fontSize" } }
    ]);
  });

  it("BREAKS the span around an A9 chip — the chip never gets a foreground write", () => {
    // "Big" [1,4) + chip [4,9) + "Deal\n" [9,14): the chip breaks the run into
    // [1,4) and [9,14), and neither write covers the chip's index space.
    const doc = buildDoc([
      { elements: [{ text: "Big" }, { text: "@chip", kind: "other" }, { text: "Deal" }] },
      para("after")
    ]);
    const { groups } = planAnalyticify(doc, touch(0));
    expect(textWrites(groups).map((u) => u.range)).toEqual([r(1, 4), r(9, 14)]);
    // Both writes carry the analytics style and only it.
    for (const u of textWrites(groups)) {
      expect(u.textStyle).toEqual(ANALYTICS_TEXT_STYLE);
      expect(u.fields).toBe("foregroundColor,fontSize");
    }
  });

  it("clamps the segment-final newline off the tail span (the API refuses it)", () => {
    // The ONLY paragraph is segment-final: "Analytics\n" = [1,11), clamped to
    // [1,10) so the doc-final newline is never style-targeted.
    const doc = buildDoc([para("Analytics")]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch(0));
    expect(paragraphsStyled).toBe(1);
    expect(textWrites(groups).map((u) => u.range)).toEqual([r(1, 10)]);
  });

  it("a touched EMPTY segment-final paragraph yields NO write and is NOT counted", () => {
    // The lone "\n" clamps to nothing -> no targetable span -> no group, and
    // paragraphsStyled must not overstate (the receipt can't claim a styled
    // line that got no write).
    const doc = buildDoc([para("")]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch(0));
    expect(groups).toEqual([]);
    expect(paragraphsStyled).toBe(0);
  });

  it("SKIPS a table paragraph even when its ordinal is in the set (row 1 parity)", () => {
    // A touched table paragraph + a touched body paragraph: only the body line
    // is styled; table structure is never character-formatted by analytics.
    const doc = buildDoc([
      { inTable: true, elements: [{ text: "cell" }] }, // p0: touched but skipped
      para("body line") // p1: touched, styled
    ]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch(0, 1));
    expect(paragraphsStyled).toBe(1);
    expect(groups).toHaveLength(1);
    // The one write is over p1 ("body line\n" = [6,16), segment-final clamp -> [6,15)).
    expect(textWrites(groups).map((u) => u.range)).toEqual([r(6, 15)]);
  });

  it("styles ONLY the paragraphs whose ordinal is in the set, in document order", () => {
    // p0 and p2 touched, p1 NOT: two groups, emitted in ascending doc position
    // (non-load-bearing order — updateTextStyle never shifts indexes).
    const doc = buildDoc([
      para("first"), // p0: "first\n" [1,7)
      para("middle"), // p1: skipped
      para("third") // p2: "third\n" [14,20), segment-final -> [14,19)
    ]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch(0, 2));
    expect(paragraphsStyled).toBe(2);
    expect(textWrites(groups).map((u) => u.range)).toEqual([r(1, 7), r(14, 19)]);
  });

  it("emits ONLY updateTextStyle — never a paragraph-style or named-style write (003-F2)", () => {
    // The audit that pins "analytics is pure CHARACTER formatting": no
    // namedStyleType change (which would restyle every same-style paragraph),
    // no paragraph-style write, no content delete.
    const doc = buildDoc([
      { style: "HEADING_1", elements: [{ text: "Heading line" }] }, // even a heading: char-only
      para("body line")
    ]);
    const { groups } = planAnalyticify(doc, touch(0, 1));
    const kinds = groups.flatMap((g) => g.requests.map(reqKind));
    expect(kinds.length).toBeGreaterThan(0); // not vacuous
    for (const k of kinds) expect(k).toBe("text");
    // And every write names exactly the two analytics fields, nothing more.
    for (const u of textWrites(groups)) expect(u.fields).toBe("foregroundColor,fontSize");
  });

  it("is IDEMPOTENT: re-planning over the same ordinals emits byte-identical requests (003-S5)", () => {
    // Every write is an absolute value (the same navy, the same 14pt) keyed off
    // ordinals + geometry the write never mutates, so a second run is identical
    // -> a server-side no-op. (The fixture even pre-stains the runs navy to model
    // an already-analytic'd doc; the plan is still the same writes.)
    const fresh = buildDoc([
      { elements: [{ text: "Claim A" }] },
      { elements: [{ text: "Claim B" }] }
    ]);
    const already: GpSpec[] = [
      { elements: [{ text: "Claim A", fg: ANALYTICS_FG_HEX, size: ANALYTICS_PT }] },
      { elements: [{ text: "Claim B", fg: ANALYTICS_FG_HEX, size: ANALYTICS_PT }] }
    ];
    const first = planAnalyticify(fresh, touch(0, 1));
    const second = planAnalyticify(buildDoc(already), touch(0, 1));
    expect(second.groups).toEqual(first.groups);
    expect(second.paragraphsStyled).toBe(first.paragraphsStyled);
  });

  it("an EMPTY ordinals set yields zero groups and a zero count (the no-touch case)", () => {
    const doc = buildDoc([para("nothing touched"), para("still nothing")]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch());
    expect(groups).toEqual([]);
    expect(paragraphsStyled).toBe(0);
  });

  it("an ordinal NOT present in the doc is harmlessly ignored (TOCTOU safety)", () => {
    // The adapter may lower a stale selection to an ordinal that no longer
    // exists; planAnalyticify must not throw or fabricate a write.
    const doc = buildDoc([para("only paragraph")]);
    const { groups, paragraphsStyled } = planAnalyticify(doc, touch(7));
    expect(groups).toEqual([]);
    expect(paragraphsStyled).toBe(0);
  });
});
