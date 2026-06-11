// Style identifiers, the canonical STYLE_MAP, and the pocket-box OOXML helper.
//
// Two hard-won facts from the planning review are baked in here (see LESSONS.md):
//   * Match the cite character style on its stable styleId "Style13ptBold"
//     (alias "Cite"), NEVER the friendly name "cite" — the literal name matches
//     nothing in the real Debate.dotm (decision #8).
//   * "Style13ptBold" is mis-named: it is 14pt (sz=28 half-points), not 13
//     (decision #9).

/**
 * Cite character style id, as it appears in OOXML run properties:
 * `<w:rPr><w:rStyle w:val="Style13ptBold"/></w:rPr>`. Friendly name "Cite" is a
 * UI label only.
 */
export const CITE_STYLE_ID = "Style13ptBold";

/**
 * LEGACY condense break-marker character-style id. RETIRED as the marker signal — it depended on a
 * custom `<w:rStyle>` reference surviving `insertOoxml` into a populated doc, which it does NOT: Word
 * drops an `<w:rStyle>` whose styleId is not resident in the destination's style table, erasing the
 * only signal Uncondense keyed on (the 2026-06-09 live irreversibility bug; xmldom + COM both
 * over-preserved the style, so every headless/COM gate stayed green while the live host stripped it).
 * Replaced by {@link MARK_SENTINEL} — an intrinsic, directly-formatted, self-describing TEXT signal
 * that survives the round-trip. Kept only as documentation of the retired approach; nothing keys on it.
 */
export const CONDENSE_MARK_STYLE = "RostrumCondenseBreak";

/**
 * Zero-width TEXT sentinel that marks a **condense boundary / payload** run. This is the condense
 * marker signal — carried in the run's own `<w:t>` text, NOT in a style reference — because per
 * Microsoft's OOXML guidance only the provided STYLE definition is applied at insert time; a net-new
 * custom `<w:rStyle>` is not persisted into a populated doc's style table and is stripped on the
 * `getOoxml` round-trip (see {@link CONDENSE_MARK_STYLE}). Run TEXT and direct run formatting DO
 * survive `insertOoxml`, so Uncondense keys on the sentinel (`runTextRaw().includes(MARK_SENTINEL)`).
 *
 * U+2063 INVISIBLE SEPARATOR: Unicode category Cf (format), zero-width, and — critically — NOT a
 * member of `\p{Zs}`/`\s`, so the whitespace-collapse pass never folds it and `isBlankParagraph`
 * never misreads it. It is purpose-built to be invisible, so a glyph marker stays visually identical.
 *
 * IT IS A SINGLE NAMED CONSTANT ON PURPOSE: the one thing only a live wet-test can confirm is whether
 * Word preserves this character through `insertOoxml`→`getOoxml`. If a wet-test ever shows it stripped,
 * swap this constant (e.g. to a different invisible/private token) — no other code changes.
 */
export const MARK_SENTINEL = String.fromCodePoint(0x2063);

/**
 * A paragraph is kept by the heading rule iff its canonical 0-based outline level
 * is in [0, 3] — i.e. Heading 1–4 and any derived style at those levels, which is
 * how the template's navy "Analytics" style (outlineLvl 3) survives Hide
 * (decision #7). Keying on the resolved level, not a hardcoded style list, makes
 * this auto-cover future derived heading styles.
 */
export const KEEP_OUTLINE_MAX = 3;

/**
 * Valid `w:highlight` values, minus "none". OOXML highlight is a *named* enum
 * (not a hex color), which is why keep-colors are stored as these names. Default
 * keep-set is ALL of them; the user narrows the set in settings (decision #11).
 */
export const HIGHLIGHT_COLORS: readonly string[] = [
  "yellow",
  "green",
  "cyan",
  "magenta",
  "blue",
  "red",
  "darkBlue",
  "darkCyan",
  "darkGreen",
  "darkMagenta",
  "darkRed",
  "darkYellow",
  "darkGray",
  "lightGray",
  "black",
  "white"
];

/** Set form (lower-cased) for O(1) validation when parsing manifests / settings. */
export const HIGHLIGHT_COLOR_SET: ReadonlySet<string> = new Set(
  HIGHLIGHT_COLORS.map((c) => c.toLowerCase())
);

/** Friendly-name -> style spec for the gated "Apply Rostrum styles" action (Stage 2). */
export interface StyleSpec {
  /** Built-in style name for paragraph styles (Word.BuiltInStyleName form). */
  builtIn?: string;
  /** Concrete styleId for character styles. */
  styleId?: string;
  /** Point size (decision #9 sizes). */
  sizePt: number;
  /** Character styles are applied to runs, not paragraphs. */
  isCharacterStyle?: boolean;
  /** Whether this style gets the boxed border (pocket only). */
  boxed?: boolean;
}

/**
 * Canonical sizes from the real template (decision #9): pocket=26, hat=22,
 * block=16, tag=14, cite=14. Headings map to built-in Heading1–4 so their
 * outline levels (0–3) drive the keep rule for free.
 */
export const STYLE_MAP: Readonly<Record<string, StyleSpec>> = {
  pocket: { builtIn: "Heading1", sizePt: 26, boxed: true },
  hat: { builtIn: "Heading2", sizePt: 22 },
  block: { builtIn: "Heading3", sizePt: 16 },
  tag: { builtIn: "Heading4", sizePt: 14 },
  cite: { styleId: CITE_STYLE_ID, sizePt: 14, isCharacterStyle: true }
};

/** OOXML uses half-points for font size; 14pt -> 28. */
export function ptToHalfPoints(pt: number): number {
  return Math.round(pt * 2);
}

/** Pocket box appearance — single thin border on all four sides. */
const POCKET_BORDER = {
  /** OOXML border style token. */
  val: "single",
  /** Eighths of a point: 4 = 0.5pt. */
  sizeEighths: 4,
  /** Space between border and text, in points. */
  spacePt: 1,
  /** Border color (hex, no leading #). */
  color: "000000"
} as const;

/**
 * The `<w:pBdr>` fragment for the boxed pocket. Used as the OOXML fallback when
 * the host lacks `Style.borders` (decision #17 — there is no public OOXML-on-
 * style API, so Stage 2 applies this per-paragraph on pocket paragraphs rather
 * than to the style definition). Pure + testable.
 */
export function buildPocketBorderOoxml(): string {
  const b = POCKET_BORDER;
  const side = (name: string): string =>
    `<w:${name} w:val="${b.val}" w:sz="${b.sizeEighths}" w:space="${b.spacePt}" w:color="${b.color}"/>`;
  return (
    "<w:pBdr>" +
    side("top") +
    side("left") +
    side("bottom") +
    side("right") +
    "</w:pBdr>"
  );
}
