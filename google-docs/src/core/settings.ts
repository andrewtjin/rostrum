// 3-tier settings resolution for the gdocs engine (plan S4 + A8 + A16).
//
// Precedence: document properties > device defaults > built-ins — the Word
// engine's resolution order (src/core/settings.ts, decision #15) ported to
// Docs. Both upper tiers arrive as RAW JSON strings because the adapter reads
// them verbatim from DocumentProperties / UserProperties (DocsPort
// readSettingsJson): all parsing and validation live HERE so the adapter
// stays a dumb pipe and every corruption path is unit-testable without a
// host.
//
// Corruption-tolerance philosophy (mirrors the Word reference): a settings
// blob can be truncated, hand-edited, or written by a future version — the
// verbs must NEVER be bricked by it. So resolution never throws; a malformed
// tier (or a wrong-typed field within an otherwise-valid tier) simply
// contributes nothing and the next tier down fills the gap, field by field.
// The ONE deliberate exception to "invalid contributes nothing": an explicit
// keepColors ARRAY is always honored even when its (normalized) result is
// empty — "keep nothing by color" is a real user choice (Word decision #11
// parity), not corruption.

import { GdocsSettings } from "./types";
import { CITE_PT, DEFAULT_CITE_MIN_PT } from "./constants";

// ---------------------------------------------------------------------------
// DEFAULT_KEEP_HEXES — the CLOSED default keep set (plan A8).
//
// Docs has no separate highlight channel: highlight IS
// textStyle.backgroundColor, so "keep any background" would also keep the
// near-white shading web pastes carry (#f8f9fa and friends), which must HIDE.
// Word parity instead: a closed set of the colors a debater can actually
// CHOOSE as a highlight — the Docs palette's chromatic swatches plus the
// classic Word highlight hexes that .docx imports materialize. The sidebar
// keep-color checkboxes render exactly this set; "keep any highlight color"
// is a separate explicit master toggle (keepMode "anyHighlight").
// ---------------------------------------------------------------------------

/**
 * The Docs default color picker's 70 CHROMATIC swatches: 10 hue columns
 * (red berry, red, orange, yellow, green, cyan, cornflower blue, blue,
 * purple, magenta — the grid's left-to-right order) x 7 rows (base +
 * light 1-3 + dark 1-3). Provenance: loop-001 research styles.json keyFact,
 * cross-verified across 4 independent palette dumps (verified confidence).
 *
 * The grid's 10-swatch GRAYSCALE column (white, light/dark grays, black) is
 * deliberately ABSENT: those are exactly the shades web-paste evidence
 * carries as incidental background, and plan A8's hard rule is that
 * near-white shading must hide. (Black re-enters via the Word group below —
 * it is a real Word highlight choice; the light grays never do.)
 */
const DOCS_CHROMATIC_PALETTE: readonly string[] = [
  // base row
  "#980000", "#ff0000", "#ff9900", "#ffff00", "#00ff00",
  "#00ffff", "#4a86e8", "#0000ff", "#9900ff", "#ff00ff",
  // light 3 (palest tints)
  "#e6b8af", "#f4cccc", "#fce5cd", "#fff2cc", "#d9ead3",
  "#d0e0e3", "#c9daf8", "#cfe2f3", "#d9d2e9", "#ead1dc",
  // light 2
  "#dd7e6b", "#ea9999", "#f9cb9c", "#ffe599", "#b6d7a8",
  "#a2c4c9", "#a4c2f4", "#9fc5e8", "#b4a7d6", "#d5a6bd",
  // light 1 (the row styles.json names: light yellow 1 #ffd966, etc.)
  "#cc4125", "#e06666", "#f6b26b", "#ffd966", "#93c47d",
  "#76a5af", "#6d9eeb", "#6fa8dc", "#8e7cc3", "#c27ba0",
  // dark 1
  "#a61c00", "#cc0000", "#e69138", "#f1c232", "#6aa84f",
  "#45818e", "#3c78d8", "#3d85c6", "#674ea7", "#a64d79",
  // dark 2
  "#85200c", "#990000", "#b45f06", "#bf9000", "#38761d",
  "#134f5c", "#1155cc", "#0b5394", "#351c75", "#741b47",
  // dark 3 (deepest shades)
  "#5b0f00", "#660000", "#783f04", "#7f6000", "#274e13",
  "#0c343d", "#1c4587", "#073763", "#20124d", "#4c1130"
];

/**
 * The Word `w:highlight` enum's classic RGB values — what a .docx import
 * materializes as backgroundColor in Docs, so Word-highlighted evidence stays
 * kept by default after import. Provenance: ECMA-376 ST_HighlightColor; the
 * names mirror the Word engine's HIGHLIGHT_COLORS (src/core/styles.ts), which
 * keeps ALL of them by default. ONE exception: "white" (#ffffff) is dropped —
 * in Docs a white background is indistinguishable from unhighlighted /
 * web-pasted text, and keeping it would defeat the hide pass wholesale.
 * Black stays (conservative over-keep: a deliberate black highlight beats
 * the rare dark-web-paste false keep).
 */
const WORD_HIGHLIGHT_HEXES: readonly string[] = [
  "#ffff00", // yellow      (also Docs base row)
  "#00ff00", // green       (also Docs base row)
  "#00ffff", // cyan        (also Docs base row)
  "#ff00ff", // magenta     (also Docs base row)
  "#0000ff", // blue        (also Docs base row)
  "#ff0000", // red         (also Docs base row)
  "#00008b", // darkBlue
  "#008b8b", // darkCyan
  "#006400", // darkGreen
  "#8b008b", // darkMagenta
  "#8b0000", // darkRed
  "#808000", // darkYellow
  "#808080", // darkGray
  "#c0c0c0", // lightGray
  "#000000" // black
];

/** The closed default keep set: 70 chromatic palette swatches + 15 Word
 * highlight hexes, deduplicated (the 6 shared base hues collapse) = 79. */
export const DEFAULT_KEEP_HEXES: ReadonlySet<string> = new Set([
  ...DOCS_CHROMATIC_PALETTE,
  ...WORD_HIGHLIGHT_HEXES
]);

// ---------------------------------------------------------------------------
// Field validators — each tier is reduced to a TierContribution of ONLY the
// fields it states with valid types, so the merge in resolveSettings is a
// uniform per-field `??` cascade (a `||` would eat explicit false / empty).
// ---------------------------------------------------------------------------

/** What one tier validly contributes; absent field = "this tier is silent". */
interface TierContribution {
  keepMode?: GdocsSettings["keepMode"];
  keepColors?: ReadonlySet<string>;
  citeMinPt?: number;
  structuralCite?: boolean;
  collapseSpacing?: boolean;
}

/** Canonical hex shape after normalization: exactly "#rrggbb" lower-case.
 * The optional leading "#" is the ONE liberality (a hand-edited blob that
 * drops it is unambiguous); 3-digit shorthand and color NAMES are rejected —
 * Docs only ever reports 6-digit RGB, so anything else is corruption. */
const HEX_SHAPE = /^#?[0-9a-f]{6}$/;

/** Normalize one keep-color entry to "#rrggbb" lower-case; null = invalid
 * (wrong type / wrong shape), to be DROPPED rather than guessed at. */
function normalizeHex(value: unknown): string | null {
  if (typeof value !== "string") return null;
  const v = value.trim().toLowerCase();
  if (!HEX_SHAPE.test(v)) return null;
  return v.startsWith("#") ? v : `#${v}`;
}

/**
 * Validate a keepColors field. Non-array = the field contributes nothing
 * (null). An ARRAY — even one whose entries all fail normalization — yields a
 * (possibly empty) set: the array itself is the explicit user choice, and an
 * empty keep-set is honored (decision #11 parity, see module header).
 */
function normalizeColorList(value: unknown): ReadonlySet<string> | null {
  if (!Array.isArray(value)) return null;
  const out = new Set<string>();
  for (const entry of value) {
    const hex = normalizeHex(entry);
    if (hex !== null) out.add(hex); // Set membership dedupes "#FFFF00"/"ffff00"
  }
  return out;
}

/**
 * Docs' UI font-size floor is 6pt (styles.json: "UI font size 6-400pt"), so a
 * cite threshold below it could never match a real run; above CITE_PT the
 * threshold would miss the very convention Mark-cite writes. Out-of-range
 * values are CLAMPED, not dropped — a user who typed "20" meant "strict",
 * and the nearest legal value preserves that intent.
 */
const CITE_MIN_FLOOR_PT = 6;

/** Clamp citeMinPt into [CITE_MIN_FLOOR_PT, CITE_PT] (plan A16 minor). */
function clampCiteMinPt(value: number): number {
  return Math.min(CITE_PT, Math.max(CITE_MIN_FLOOR_PT, value));
}

/**
 * Parse one tier's raw JSON into its valid contribution. NEVER throws:
 * malformed JSON or a non-object root means the whole tier is silently inert;
 * within a valid object each field is validated INDEPENDENTLY (one garbled
 * field must not poison its siblings — the Word loadCondenseSettings
 * philosophy), so a tier can win on some fields and be silent on others.
 */
function parseTier(json: string | null): TierContribution {
  if (json === null || json === "") return {};
  let parsed: unknown;
  try {
    parsed = JSON.parse(json);
  } catch {
    return {}; // truncated / hand-mangled JSON: tier contributes nothing
  }
  // Arrays and primitives parse fine but are not settings objects.
  if (typeof parsed !== "object" || parsed === null || Array.isArray(parsed)) return {};
  const o = parsed as Record<string, unknown>;
  const out: TierContribution = {};
  // keepMode: only the two known modes; anything else (incl. a future mode
  // from a newer version) falls through rather than guessing.
  if (o.keepMode === "set" || o.keepMode === "anyHighlight") out.keepMode = o.keepMode;
  const keepColors = normalizeColorList(o.keepColors);
  if (keepColors !== null) out.keepColors = keepColors;
  // citeMinPt: finite numbers only (JSON cannot carry NaN/Infinity, but a
  // null written by a buggy serializer must not become a threshold). The
  // clamp happens once, after the merge, so all tiers share it.
  if (typeof o.citeMinPt === "number" && Number.isFinite(o.citeMinPt)) {
    out.citeMinPt = o.citeMinPt;
  }
  if (typeof o.structuralCite === "boolean") out.structuralCite = o.structuralCite;
  if (typeof o.collapseSpacing === "boolean") out.collapseSpacing = o.collapseSpacing;
  return out;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Resolve the settings the engine runs with. Each field independently honors
 * doc properties > device defaults > built-ins (the precedence triangle), so
 * e.g. a doc that pins only citeMinPt still inherits the device keep-set.
 * Built-ins: keepMode "set" with DEFAULT_KEEP_HEXES (plan A8), citeMinPt
 * DEFAULT_CITE_MIN_PT, structuralCite ON (plan A10), collapseSpacing OFF
 * (plan A12). Never throws — see parseTier. The default keep set is returned
 * by REFERENCE (it is a ReadonlySet; copying 79 hexes per resolve buys
 * nothing).
 */
export function resolveSettings(
  docPropsJson: string | null,
  deviceDefaultsJson: string | null
): GdocsSettings {
  const doc = parseTier(docPropsJson);
  const device = parseTier(deviceDefaultsJson);
  return {
    // `??` (never `||`) is load-bearing: an explicit empty keep-set, an
    // explicit `false`, and a to-be-clamped 0 must all win their tier.
    keepMode: doc.keepMode ?? device.keepMode ?? "set",
    keepColors: doc.keepColors ?? device.keepColors ?? DEFAULT_KEEP_HEXES,
    citeMinPt: clampCiteMinPt(doc.citeMinPt ?? device.citeMinPt ?? DEFAULT_CITE_MIN_PT),
    structuralCite: doc.structuralCite ?? device.structuralCite ?? true,
    collapseSpacing: doc.collapseSpacing ?? device.collapseSpacing ?? false
  };
}

/**
 * Serialize settings to the JSON blob the adapter persists (either tier).
 * Normalizes on the way OUT as well as in (Word's save* parity): hexes are
 * canonicalized + deduped + SORTED (deterministic output regardless of set
 * insertion order — stable for tests and for change detection), citeMinPt is
 * clamped (a non-finite value — possible in a hand-built object since NaN is
 * a `number` — falls back to the default rather than serializing as JSON
 * null). Invariant: a persisted blob always round-trips through
 * resolveSettings unchanged, so storage never accumulates drift.
 */
export function serializeSettings(s: GdocsSettings): string {
  const normalized = new Set<string>();
  for (const entry of s.keepColors) {
    const hex = normalizeHex(entry);
    if (hex !== null) normalized.add(hex);
  }
  return JSON.stringify({
    keepMode: s.keepMode,
    keepColors: [...normalized].sort(),
    citeMinPt: Number.isFinite(s.citeMinPt) ? clampCiteMinPt(s.citeMinPt) : DEFAULT_CITE_MIN_PT,
    structuralCite: s.structuralCite,
    collapseSpacing: s.collapseSpacing
  });
}
