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
 * Replaced by {@link MARK_SIGNATURE} — an intrinsic, directly-formatted, self-describing TEXT signal
 * that survives the round-trip. Kept only as documentation of the retired approach; nothing keys on it.
 */
export const CONDENSE_MARK_STYLE = "RostrumCondenseBreak";

/**
 * Zero-width TEXT signature that marks a **condense boundary / payload** run. This is the condense
 * marker signal — carried in the run's own `<w:t>` text, NOT in a style reference — because per
 * Microsoft's OOXML guidance only the provided STYLE definition is applied at insert time; a net-new
 * custom `<w:rStyle>` is not persisted into a populated doc's style table and is stripped on the
 * `getOoxml` round-trip (see {@link CONDENSE_MARK_STYLE}). Run TEXT and direct run formatting DO
 * survive `insertOoxml`, so Uncondense keys on the signature set (`MARK_SIGNATURES`).
 *
 * WHY A TWO-CODE-UNIT PAIR (U+200B ZERO WIDTH SPACE + U+2060 WORD JOINER), not one char:
 *   * The original single char, U+2063 INVISIBLE SEPARATOR, proved FONT-DEPENDENT in live Word —
 *     a visible comma-like mark in the user's fonts — because the boundary glyph run is
 *     deliberately NOT vanished (its space/¶ IS the visible separator). That was the 2026-06-10
 *     render regression. U+200B and U+2060 both render zero-width in Word regardless of font.
 *   * A single U+200B is NOT usable either: organic ZWSPs are ENDEMIC in web-pasted text (debate
 *     evidence!), and marker detection keys on bare containment with only a followed-by-{space,¶}
 *     guard. Organic "word​ word" would fabricate a paragraph break AND eat the space on
 *     Uncondense (breaking round-trip identity on untouched user text), and ANY run carrying an
 *     organic ZWSP would be misflagged `breakMarker` — exempting it from Shrink and skipping it in
 *     whitespace collapse (adversarial repro, 2026-06-10). The PAIR closes that hole: ZWSP (allows
 *     a break) immediately followed by WORD JOINER (forbids one) serve OPPOSITE purposes, so their
 *     organic adjacency is essentially nonexistent — and a lone organic ZWSP matches nothing.
 *   * Both chars are category Cf and NOT in `\p{Zs}`/`\s`, so the whitespace-collapse pass never
 *     folds the signature, `isBlankParagraph` never misreads it, and keepers' word-separator
 *     predicate ignores it (same class behavior as the legacy U+2063).
 *
 * Survival of the pair through Word's real import/export pipeline was verified via the COM harness
 * (Range.InsertXML into a populated doc → Range.WordOpenXML) for all three condense modes.
 *
 * This is the WRITTEN form only. READ paths must accept every member of {@link MARK_SIGNATURES}.
 */
export const MARK_SIGNATURE = String.fromCodePoint(0x200b) + String.fromCodePoint(0x2060);

/**
 * LEGACY signature: U+2063 INVISIBLE SEPARATOR alone — the original written form, shipped in the
 * deployed build of 2026-06-10 ONLY, retired the same day because some fonts render it as a visible
 * comma-like glyph in Word's normal view (see {@link MARK_SIGNATURE}). Never written anymore; kept
 * in the read set so docs condensed by that build still uncondense losslessly. (Organic U+2063 in
 * user text is astronomically rare — unlike ZWSP — so the single-char READ is tolerable for the
 * legacy window even though it would be unacceptable as the written form going forward.)
 */
export const LEGACY_MARK_SIGNATURE = String.fromCodePoint(0x2063);

/**
 * EVERY signature the READ/STRIP paths must honor, LONGEST FIRST (the regex alternation built from
 * this set must try longer signatures before shorter ones so none can shadow another). Exactly two
 * members: the current written pair and the legacy U+2063 window (2026-06-10 deployed build). A
 * bare single U+200B is DELIBERATELY ABSENT — no shipped build ever wrote one (only an unpushed
 * interim commit did, which never reached a user doc), and accepting it would re-open the whole
 * organic-ZWSP false-positive class above. A single run can carry BOTH kinds — re-condensing a doc
 * that still holds legacy markers writes new-signature boundaries beside the old ones, and Word
 * coalesces adjacent identically-formatted runs — so detection/split/strip must treat the whole SET
 * as one signal class, never just `MARK_SIGNATURE`.
 */
export const MARK_SIGNATURES: readonly string[] = [MARK_SIGNATURE, LEGACY_MARK_SIGNATURE];

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

/**
 * Pocket box appearance — single 3pt border on all four sides. This is the SINGLE source of
 * truth for the pocket box: both the live `Style.borders` path (officeStyles.ts) and the OOXML
 * fragment below read it, so the two representations can never drift (they previously diverged
 * to 1pt vs 0.5pt). 3pt matches Verbatim's pocket (Heading 1) style, whose `w:pBdr` is
 * `w:sz="24"` on every side — and OOXML border `w:sz` is in EIGHTHS of a point, so 24/8 = 3.0pt.
 */
export const POCKET_BORDER = {
  /** OOXML border style token. */
  val: "single",
  /** Eighths of a point: 24 = 3.0pt (Verbatim's `w:sz="24"`). */
  sizeEighths: 24,
  /** The equivalent `Word.BorderWidth` enum token for the live `Style.borders` path. The API
   *  takes a string token, NOT a number, and "Pt300" is the 3.0pt member — same width as
   *  `sizeEighths: 24`, kept here so both paths derive the pocket box from one constant. */
  borderWidthToken: "Pt300",
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
