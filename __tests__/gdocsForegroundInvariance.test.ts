// 003-F6 foreground two-parse invariance (exec-review MINOR — spec-mandated).
//
// THE SPEC INVARIANT: foregroundHex is read by EXACTLY ONE policy path — the
// analytics keeper (keepers.isAnalytics, the off-palette navy ANALYTICS_FG_HEX
// that analytic-ify writes). Every other verdict the engine forms must be
// BLIND to text color. This suite is the standing tripwire that keeps it that
// way: if a future edit ever makes planApplyStyles read foreground at all, or
// makes planKeeps/planHide read it for any reason OTHER than that one exact
// analytics match (e.g. a fuzzy/approximate navy match, a "keep colored text"
// feature), one of these assertions FAILS.
//
// WHY THE NAIVE "strip everything → identical" CLAIM IS WRONG, and what we
// assert instead. The foreground.json fixture is built with THREE foreground
// states (parse suite pins all three): P0 carries the EXACT analytics navy
// (#0b5396), P1 the genuine one-byte-off "dark blue 2" near-miss (#0b5394),
// P2 inherits (null). P0's navy is SUPPOSED to flip its keeper verdict — that
// is the whole point of the analytics keep (003-S2) — so deep-equality between
// the present doc and an ALL-foreground-stripped twin is provably FALSE for
// planKeeps/planHide (verified: P0 keeps with navy, hides without it). The real
// invariant therefore has two honest halves:
//   (1) planApplyStyles is FULLY foreground-blind — it never reads color — so
//       stripping ALL foreground leaves its plan byte-identical.
//   (2) planKeeps / planHide are foreground-invariant on the INERT foreground —
//       the near-miss and the null — i.e. stripping color off every NON-analytics
//       element changes no verdict. Only the documented analytics navy matters.
// A NON-VACUITY guard then proves the fixture actually exercises foreground:
// stripping the analytics navy too DOES change planKeeps. Without that guard a
// fixture with no live foreground would pass (1)+(2) trivially and the tripwire
// would be asleep.

import * as fs from "fs";
import * as path from "path";
import { parseDocument } from "../google-docs/src/core/parse";
import { isAnalytics, planKeeps } from "../google-docs/src/core/keepers";
import { planApplyStyles } from "../google-docs/src/core/styles";
import { planHide } from "../google-docs/src/core/planner";
import { resolveSettings } from "../google-docs/src/core/settings";
import { GDoc, GElement, GdocsSettings } from "../google-docs/src/core/types";

// The exact loader the parse/schema suites use — one fixture, one parse, real
// documents.get-shaped JSON (not a hand-built GDoc), so this guards the WHOLE
// pipeline from wire bytes through the planners.
const FIXTURE = path.join(__dirname, "fixtures", "gdocs", "foreground.json");

/** Parse the committed foreground fixture exactly as the engine would. */
function parseForegroundFixture(): GDoc {
  return parseDocument(JSON.parse(fs.readFileSync(FIXTURE, "utf8")) as unknown);
}

/** Deep-copy a parsed GDoc, mapping every element's foregroundHex through
 * `next` — a plain-data structure, so spreads suffice (matches the planner
 * suite's cloneDoc). The other facts (size/bold/background/text/indexes) ride
 * through untouched, so the ONLY axis these twins differ on is text color. */
function mapForeground(doc: GDoc, next: (el: GElement) => string | null): GDoc {
  return {
    ...doc,
    paragraphs: doc.paragraphs.map((p) => ({
      ...p,
      elements: p.elements.map((el): GElement => ({ ...el, foregroundHex: next(el) }))
    })),
    namedRanges: doc.namedRanges.map((nr) => ({ ...nr, segments: nr.segments.map((s) => ({ ...s })) }))
  };
}

// The engine's real defaults (analytics keep is always-on, ungated; structural
// cite ON; collapseSpacing OFF) — the same resolveSettings(null, null) the
// realDocs scale suite drives the planners with.
const SETTINGS: GdocsSettings = resolveSettings(null, null);

describe("003-F6: foreground is invisible to every policy path EXCEPT the analytics keeper", () => {
  // Parsed ONCE (the "two-parse" guard is about the two foreground variants of
  // the SAME parse, not two reads); the planners are pure, so reuse is safe.
  const present = parseForegroundFixture();

  // Twin A: ALL foreground gone — the doc as if analytic-ify had never run.
  const strippedAll = mapForeground(present, () => null);
  // Twin B: foreground gone on every element EXCEPT the analytics navy — the
  // near-miss and the inherited-null lose their (inert) color, the one real
  // analytics run keeps its navy. isAnalytics is the SINGLE-SOURCE predicate
  // the keeper itself consults, so "inert" here means precisely "what the
  // policy is allowed to ignore".
  const strippedInert = mapForeground(present, (el) => (isAnalytics(el) ? el.foregroundHex : null));

  it("the fixture really exercises foreground: exactly one analytics run, the near-miss is NOT it", () => {
    // Guards the premise of every assertion below. If the fixture drifted to
    // carry zero analytics foreground, the invariance tests would pass
    // vacuously — this makes that drift fail LOUDLY instead.
    const analyticsEls = present.paragraphs.flatMap((p) => p.elements).filter(isAnalytics);
    expect(analyticsEls).toHaveLength(1);
    // P0 is the analytics paragraph; P1's near-miss (#0b5394) is a genuine
    // palette swatch that must NOT read as analytics; P2 inherits.
    expect(present.paragraphs[0].elements.some(isAnalytics)).toBe(true);
    expect(present.paragraphs[1].elements.some(isAnalytics)).toBe(false);
    expect(present.paragraphs[1].elements[0].foregroundHex).not.toBeNull(); // it DOES carry color
    expect(present.paragraphs[2].elements[0].foregroundHex).toBeNull();
  });

  it("planApplyStyles is FULLY foreground-blind: stripping ALL color changes no styles verdict", () => {
    // The strongest of the three — styles must ignore foreground entirely
    // (003-F2: analytic-ify is the only color writer; the styles lane never
    // reads color). Comparing against the all-stripped twin is the real
    // tripwire: if styles ever branched on foreground, even the analytics navy,
    // this deep-equality breaks.
    expect(planApplyStyles(present, SETTINGS)).toEqual(planApplyStyles(strippedAll, SETTINGS));
  });

  it("planKeeps ignores INERT foreground (the near-miss and the null) — verdicts unchanged", () => {
    // The near-miss #0b5394 and the inherited null are color the policy MUST
    // treat as nothing. Removing them must not move a single keeper flag.
    expect(planKeeps(present, SETTINGS)).toEqual(planKeeps(strippedInert, SETTINGS));
  });

  it("planHide ignores INERT foreground: identical groups AND receipt", () => {
    // Hide runs planKeeps over its restored view, so foreground reaches it only
    // through that same keeper. Stripping the inert color must yield a
    // byte-identical plan — groups and the counted receipt alike.
    const a = planHide(present, SETTINGS);
    const b = planHide(strippedInert, SETTINGS);
    expect(a.groups).toEqual(b.groups);
    expect(a.result).toEqual(b.result);
  });

  it("NON-VACUITY: the analytics navy DOES matter — stripping it flips a keeper verdict", () => {
    // Proves the invariance above is not passing over a foreground-free fixture:
    // the ONE allowed color axis genuinely changes the outcome. P0's analytics
    // run keeps with its navy and would hide without it (003-S2), so the
    // all-stripped keeps must DIFFER from present — and differ precisely on P0.
    const keptPresent = planKeeps(present, SETTINGS);
    const keptStripped = planKeeps(strippedAll, SETTINGS);
    expect(keptStripped).not.toEqual(keptPresent);
    // Pin WHERE they differ: P0's lone element flips kept -> hidden; P1/P2
    // (already inert) are unchanged, confirming the divergence is exactly the
    // analytics keep and nothing else.
    expect(keptPresent[0].elementKeep).toEqual([true]);
    expect(keptStripped[0].elementKeep).toEqual([false]);
    expect(keptStripped[1]).toEqual(keptPresent[1]);
    expect(keptStripped[2]).toEqual(keptPresent[2]);
  });
});
