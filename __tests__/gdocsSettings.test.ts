// Settings resolution suite (plan S4 + A8 + A16): the closed default keep
// set, the per-field precedence triangle, corruption tolerance on every tier
// (never throws), the explicit-empty-set rule (Word decision #11 parity),
// citeMinPt clamps, hex normalization + invalid-entry drops, and the
// serialize/resolve round-trip invariant.

import {
  DEFAULT_KEEP_HEXES,
  resolveSettings,
  serializeSettings
} from "../gdocs/src/core/settings";
import {
  CITE_PT,
  DEFAULT_CITE_MIN_PT,
  NEAR_WHITE_MIN_CHANNEL
} from "../gdocs/src/core/constants";
import { GdocsSettings } from "../gdocs/src/core/types";

/** Tiers arrive as raw JSON strings (DocsPort.readSettingsJson) — this is the
 * shorthand for "a tier that validly states exactly these fields". */
const tier = (o: unknown): string => JSON.stringify(o);

/** Sorted-array view of a keep set, for order-independent equality asserts. */
const colors = (s: GdocsSettings): string[] => [...s.keepColors].sort();

describe("DEFAULT_KEEP_HEXES (plan A8 closed default keep set)", () => {
  it("is 70 chromatic palette swatches + 15 Word hexes, 6 shared = 79", () => {
    expect(DEFAULT_KEEP_HEXES.size).toBe(79);
  });

  it.each([
    // Docs base row (styles.json verified palette fact)
    "#ffff00", // yellow
    "#00ff00", // green
    "#00ffff", // cyan
    "#ff00ff", // magenta
    "#ff0000", // red
    "#ff9900", // orange
    "#4a86e8", // cornflower blue
    "#0000ff", // blue
    "#9900ff", // purple
    "#980000", // red berry
    // light/dark variants named in styles.json
    "#ffd966", // light yellow 1
    "#93c47d", // light green 1
    "#76a5af", // light cyan 1
    // Word w:highlight classic hexes absent from the Docs palette
    "#00008b", // darkBlue
    "#008b8b", // darkCyan
    "#006400", // darkGreen
    "#8b008b", // darkMagenta
    "#8b0000", // darkRed
    "#808000", // darkYellow
    "#808080", // darkGray
    "#c0c0c0", // lightGray
    "#000000" // black (deliberate over-keep — Word highlight choice)
  ])("contains %s", (hex) => {
    expect(DEFAULT_KEEP_HEXES.has(hex)).toBe(true);
  });

  it.each([
    "#ffffff", // white: indistinguishable from unhighlighted — must NOT keep
    "#f8f9fa", // web-paste shading (the A8 named example) — must hide
    "#f3f3f3", // Docs palette light gray 3 — grayscale column excluded
    "#efefef", // light gray 2
    "#d9d9d9", // light gray 1
    "#cccccc", // gray
    "#b7b7b7", // dark gray 1
    "#999999", // dark gray 2
    "#666666", // dark gray 3
    "#434343" // dark gray 4
  ])("excludes %s (white / grayscale / web-paste shading)", (hex) => {
    expect(DEFAULT_KEEP_HEXES.has(hex)).toBe(false);
  });

  it("every member is already normalized lower-case #rrggbb", () => {
    for (const hex of DEFAULT_KEEP_HEXES) {
      expect(hex).toMatch(/^#[0-9a-f]{6}$/);
    }
  });

  it("every member is a highlight under the anyHighlight near-white floor", () => {
    // Consistency invariant between the two keep modes: a color kept by the
    // closed set must never be one the open predicate calls "not a highlight".
    for (const hex of DEFAULT_KEEP_HEXES) {
      const channels = [1, 3, 5].map((i) => parseInt(hex.slice(i, i + 2), 16));
      expect(Math.min(...channels)).toBeLessThan(NEAR_WHITE_MIN_CHANNEL);
    }
  });
});

describe("built-ins (both tiers absent)", () => {
  it("resolves the documented defaults", () => {
    const s = resolveSettings(null, null);
    expect(s.keepMode).toBe("set");
    expect(s.keepColors).toBe(DEFAULT_KEEP_HEXES); // shared by reference
    expect(s.citeMinPt).toBe(DEFAULT_CITE_MIN_PT);
    expect(s.structuralCite).toBe(true);
    expect(s.collapseSpacing).toBe(false);
  });
});

describe("precedence triangle — each field independently overridable", () => {
  it("keepMode: doc > device > built-in", () => {
    expect(
      resolveSettings(tier({ keepMode: "anyHighlight" }), tier({ keepMode: "set" })).keepMode
    ).toBe("anyHighlight");
    expect(resolveSettings(null, tier({ keepMode: "anyHighlight" })).keepMode).toBe(
      "anyHighlight"
    );
    expect(resolveSettings(null, null).keepMode).toBe("set");
  });

  it("keepColors: doc > device > built-in", () => {
    const doc = tier({ keepColors: ["#111111"] });
    const device = tier({ keepColors: ["#222222"] });
    expect(colors(resolveSettings(doc, device))).toEqual(["#111111"]);
    expect(colors(resolveSettings(null, device))).toEqual(["#222222"]);
    expect(resolveSettings(null, null).keepColors).toBe(DEFAULT_KEEP_HEXES);
  });

  it("citeMinPt: doc > device > built-in", () => {
    expect(resolveSettings(tier({ citeMinPt: 12 }), tier({ citeMinPt: 9 })).citeMinPt).toBe(12);
    expect(resolveSettings(null, tier({ citeMinPt: 9 })).citeMinPt).toBe(9);
    expect(resolveSettings(null, null).citeMinPt).toBe(DEFAULT_CITE_MIN_PT);
  });

  it("structuralCite: an explicit doc FALSE beats device true (?? not ||)", () => {
    expect(
      resolveSettings(tier({ structuralCite: false }), tier({ structuralCite: true }))
        .structuralCite
    ).toBe(false);
    expect(resolveSettings(null, tier({ structuralCite: false })).structuralCite).toBe(false);
    expect(resolveSettings(null, null).structuralCite).toBe(true);
  });

  it("collapseSpacing: doc > device > built-in", () => {
    expect(
      resolveSettings(tier({ collapseSpacing: false }), tier({ collapseSpacing: true }))
        .collapseSpacing
    ).toBe(false);
    expect(resolveSettings(null, tier({ collapseSpacing: true })).collapseSpacing).toBe(true);
    expect(resolveSettings(null, null).collapseSpacing).toBe(false);
  });

  it("fields merge ACROSS tiers (a doc pinning one field inherits the rest)", () => {
    const s = resolveSettings(
      tier({ citeMinPt: 12 }),
      tier({ keepMode: "anyHighlight", collapseSpacing: true })
    );
    expect(s.citeMinPt).toBe(12); // from doc
    expect(s.keepMode).toBe("anyHighlight"); // from device
    expect(s.collapseSpacing).toBe(true); // from device
    expect(s.keepColors).toBe(DEFAULT_KEEP_HEXES); // built-in
    expect(s.structuralCite).toBe(true); // built-in
  });
});

describe("corruption tolerance — a bad tier contributes nothing, never throws", () => {
  it("truncated doc-props JSON: device tier still wins", () => {
    const s = resolveSettings('{"keepMode":"anyHighl', tier({ keepMode: "anyHighlight" }));
    expect(s.keepMode).toBe("anyHighlight");
  });

  it("truncated device JSON: built-ins fill in", () => {
    const s = resolveSettings(null, '{"citeMinPt": 1');
    expect(s.citeMinPt).toBe(DEFAULT_CITE_MIN_PT);
  });

  it("BOTH tiers corrupt: full built-ins, no throw", () => {
    let s: GdocsSettings | undefined;
    expect(() => {
      s = resolveSettings("not json at all", "{{{{");
    }).not.toThrow();
    expect(s).toEqual(resolveSettings(null, null));
  });

  it.each([
    ["a bare number", "42"],
    ["a bare string", '"set"'],
    ["JSON null", "null"],
    ["JSON true", "true"],
    ["an empty array", "[]"],
    ["an array of colors (not an object)", '["#ffff00"]'],
    ["an empty string", ""]
  ])("non-object root (%s) is inert", (_label, raw) => {
    expect(resolveSettings(raw, null)).toEqual(resolveSettings(null, null));
  });

  it("keepColors as a NUMBER (wrong type) is silent — device colors win", () => {
    const s = resolveSettings(tier({ keepColors: 7 }), tier({ keepColors: ["#222222"] }));
    expect(colors(s)).toEqual(["#222222"]);
  });

  it("keepColors as a bare STRING (wrong type) is silent", () => {
    const s = resolveSettings(tier({ keepColors: "#ffff00" }), null);
    expect(s.keepColors).toBe(DEFAULT_KEEP_HEXES);
  });

  it.each([
    // [field under attack, garbled doc tier, expected default]
    ["citeMinPt as string", { citeMinPt: "12" }, DEFAULT_CITE_MIN_PT],
    ["citeMinPt as null", { citeMinPt: null }, DEFAULT_CITE_MIN_PT]
  ])("%s falls back to the built-in", (_label, doc, expected) => {
    expect(resolveSettings(tier(doc), null).citeMinPt).toBe(expected);
  });

  it("an unknown keepMode (future version?) is not guessed at", () => {
    expect(resolveSettings(tier({ keepMode: "open" }), null).keepMode).toBe("set");
  });

  it("non-boolean toggles are silent", () => {
    const s = resolveSettings(tier({ structuralCite: 1, collapseSpacing: "yes" }), null);
    expect(s.structuralCite).toBe(true);
    expect(s.collapseSpacing).toBe(false);
  });

  it("ONE garbled field does not poison its valid siblings", () => {
    const s = resolveSettings(
      tier({ keepColors: 7, citeMinPt: 12 }),
      tier({ keepColors: ["#222222"] })
    );
    expect(s.citeMinPt).toBe(12); // the valid doc field still applies
    expect(colors(s)).toEqual(["#222222"]); // the garbled one falls through
  });

  it("an empty settings object contributes nothing", () => {
    expect(resolveSettings(tier({}), tier({}))).toEqual(resolveSettings(null, null));
  });
});

describe("explicit empty keep-set honored (Word decision #11 parity)", () => {
  it("doc keepColors: [] beats a populated device set", () => {
    const s = resolveSettings(tier({ keepColors: [] }), tier({ keepColors: ["#222222"] }));
    expect(s.keepColors.size).toBe(0);
  });

  it("an array whose every entry is invalid is STILL an explicit (empty) choice", () => {
    // The array is the user's statement; bad entries are dropped, not the array.
    const s = resolveSettings(
      tier({ keepColors: ["bogus", 42] }),
      tier({ keepColors: ["#222222"] })
    );
    expect(s.keepColors.size).toBe(0);
  });

  it("device keepColors: [] is honored when the doc is silent", () => {
    expect(resolveSettings(null, tier({ keepColors: [] })).keepColors.size).toBe(0);
  });
});

describe("citeMinPt clamps (>= 6, <= CITE_PT — plan A16)", () => {
  it.each([
    [99, CITE_PT],
    [15, CITE_PT],
    [CITE_PT, CITE_PT],
    [13.5, 13.5], // fractional in-range survives (Docs sizes are doubles)
    [6, 6],
    [5, 6],
    [0, 6],
    [-3, 6]
  ])("citeMinPt %p resolves to %p", (input, expected) => {
    expect(resolveSettings(tier({ citeMinPt: input }), null).citeMinPt).toBe(expected);
  });

  it("a device-tier value is clamped identically (clamp lives after the merge)", () => {
    expect(resolveSettings(null, tier({ citeMinPt: 99 })).citeMinPt).toBe(CITE_PT);
  });
});

describe("hex normalization + invalid-entry drop", () => {
  it("lower-cases, trims, and accepts a missing '#'", () => {
    const s = resolveSettings(
      tier({ keepColors: ["#FFFF00", " #00ff00 ", "4a86e8"] }),
      null
    );
    expect(colors(s)).toEqual(["#00ff00", "#4a86e8", "#ffff00"]);
  });

  it.each([
    ["non-hex digits", "#ggg000"],
    ["too short", "#ffff0"],
    ["too long", "#ffff000"],
    ["3-digit shorthand", "#ff0"],
    ["a Word color NAME (gdocs settings are hex-only)", "yellow"],
    ["empty string", ""]
  ])("drops %s (%p) but keeps valid siblings", (_label, bad) => {
    const s = resolveSettings(tier({ keepColors: [bad, "#ffff00"] }), null);
    expect(colors(s)).toEqual(["#ffff00"]);
  });

  it("drops non-string entries mixed into the array", () => {
    const s = resolveSettings(
      tier({ keepColors: [42, null, {}, true, ["#ffff00"], "#00ff00"] }),
      null
    );
    expect(colors(s)).toEqual(["#00ff00"]);
  });

  it("deduplicates entries that normalize to the same hex", () => {
    const s = resolveSettings(tier({ keepColors: ["#FFFF00", "ffff00", "#ffff00"] }), null);
    expect(colors(s)).toEqual(["#ffff00"]);
  });
});

describe("serialize/resolve round-trip", () => {
  it("the built-in defaults survive a round-trip exactly", () => {
    const s = resolveSettings(null, null);
    const back = resolveSettings(serializeSettings(s), null);
    expect(back.keepMode).toBe(s.keepMode);
    expect(colors(back)).toEqual(colors(s));
    expect(back.citeMinPt).toBe(s.citeMinPt);
    expect(back.structuralCite).toBe(s.structuralCite);
    expect(back.collapseSpacing).toBe(s.collapseSpacing);
  });

  it("a fully customized doc tier survives a round-trip exactly", () => {
    const s = resolveSettings(
      tier({
        keepMode: "anyHighlight",
        keepColors: ["#ffd966", "#111111"],
        citeMinPt: 12.5,
        structuralCite: false,
        collapseSpacing: true
      }),
      null
    );
    const back = resolveSettings(serializeSettings(s), null);
    expect(back).toEqual({ ...s, keepColors: back.keepColors });
    expect(colors(back)).toEqual(["#111111", "#ffd966"]);
  });

  it("an explicit EMPTY keep-set survives a round-trip (decision #11)", () => {
    const s = resolveSettings(tier({ keepColors: [] }), null);
    expect(resolveSettings(serializeSettings(s), null).keepColors.size).toBe(0);
  });

  it("serializing normalizes a hand-built settings object (Word save* parity)", () => {
    const handBuilt: GdocsSettings = {
      keepMode: "set",
      keepColors: new Set(["FFFF00", "#FFFF00", "bogus"]), // un-normalized + junk
      citeMinPt: 99, // out of range
      structuralCite: true,
      collapseSpacing: false
    };
    const back = resolveSettings(serializeSettings(handBuilt), null);
    expect(colors(back)).toEqual(["#ffff00"]); // canonicalized + deduped + junk dropped
    expect(back.citeMinPt).toBe(CITE_PT); // clamped at write time
  });

  it("a non-finite citeMinPt serializes as the default, never as JSON null", () => {
    const handBuilt: GdocsSettings = {
      keepMode: "set",
      keepColors: new Set<string>(),
      citeMinPt: Number.NaN,
      structuralCite: true,
      collapseSpacing: false
    };
    const blob = JSON.parse(serializeSettings(handBuilt)) as { citeMinPt: unknown };
    expect(blob.citeMinPt).toBe(DEFAULT_CITE_MIN_PT);
  });

  it("serialization is deterministic regardless of set insertion order", () => {
    const a: GdocsSettings = {
      keepMode: "set",
      keepColors: new Set(["#ffff00", "#00ff00"]),
      citeMinPt: 13,
      structuralCite: true,
      collapseSpacing: false
    };
    const b: GdocsSettings = { ...a, keepColors: new Set(["#00ff00", "#ffff00"]) };
    expect(serializeSettings(a)).toBe(serializeSettings(b));
  });

  it("the persisted blob is a plain object with exactly the five fields", () => {
    // Pins the storage shape: a new field in GdocsSettings must show up here
    // (and get its own tier validation) or this fails — drift protection.
    const blob = JSON.parse(serializeSettings(resolveSettings(null, null))) as Record<
      string,
      unknown
    >;
    expect(Object.keys(blob).sort()).toEqual([
      "citeMinPt",
      "collapseSpacing",
      "keepColors",
      "keepMode",
      "structuralCite"
    ]);
  });
});
