// Engine constants — single-sourced so policy numbers can never drift between
// modules (the Mark-cite/keeper split was a named plan-review finding).

/**
 * The sentinel size hidden text is shrunk to, in points. 1pt is the verified
 * Docs UI floor; whether the API accepts sub-1pt is a Diagnostics probe — if a
 * future version lowers this, the OLD value must be APPENDED to SENTINELS so
 * Show All's sweep still recognizes text hidden by earlier versions.
 */
export const SENTINEL_PT = 1;

/** Every sentinel any shipped version ever wrote (append-only — plan A16). */
export const SENTINELS: readonly number[] = [SENTINEL_PT];

/**
 * The foreground color hidden text is painted, lower-case "#rrggbb". Hiding
 * shrinks text to SENTINEL_PT AND paints it white in the SAME updateTextStyle, so
 * a hidden passage is truly invisible on the (universal-in-debate) white page
 * rather than a faint 1pt smear. Two deliberate properties:
 *   * UNCONDITIONAL — every hidden run is painted white regardless of its
 *     original color; the engine never RECORDS the original (reveal clears
 *     foreground to inherit, not to a saved value). This keeps Hide's plan
 *     independent of input color, preserving the foreground-invariance invariant
 *     (gdocsForegroundInvariance): only the analytics keeper reads foreground,
 *     and analytics is always KEPT, so it is never hidden/painted. The accepted
 *     trade-off is that a manually-recolored, non-kept passage returns in the
 *     INHERITED color after a hide→show cycle.
 *   * RIDES WITH THE SIZE — font size stays the SOLE "is-hidden" marker
 *     (SENTINELS); white is a cosmetic companion set wherever the sentinel size
 *     is set and cleared wherever it is cleared (Show All restore/normalize/
 *     sweep + the re-hide rework). White is only invisible on a white page — a
 *     non-white page background is a documented out-of-scope limit.
 */
export const HIDDEN_FG_HEX = "#ffffff";

/**
 * The field mask every Hide / Show All TEXT-style write carries: fontSize (the
 * hide marker) and foregroundColor (the invisibility color) ALWAYS move together
 * — set as a pair on hide, cleared as a pair on reveal — so the two channels can
 * never drift apart. Single-sourced here exactly like restore.ts's SPACING_FIELDS.
 * NOTE: paragraph SPACING reveals (updateParagraphStyle) use their own mask and
 * must NOT include foregroundColor.
 */
export const HIDE_FIELDS = "fontSize,foregroundColor";

/**
 * The cite convention size (Word decision #9 parity: the template cite is
 * 14pt bold). Mark-cite writes this; settings clamp citeMinPt <= CITE_PT.
 */
export const CITE_PT = 14;

/** Default cite-signature threshold (bold AND >= this keeps the paragraph). */
export const DEFAULT_CITE_MIN_PT = 13;

/**
 * Max NamedRange name length we will emit. Docs documents 0-256 by convention;
 * we stay comfortably under so a region whose RLE overflows splits into
 * multiple ranges at entry boundaries (restore stays exact — plan A2).
 */
export const NAME_MAX = 200;

/** Max requests per batchUpdate chunk. Chunk boundaries fall only between
 * atomic RequestGroups (plan A11.viii). */
export const CHUNK_MAX = 5000;

/** Bounded silent re-plan attempts when the FIRST chunk hits a revision
 * mismatch (nothing applied yet). Retries are immediate — no sleeps in core. */
export const MAX_REPLAN_ATTEMPTS = 3;

/**
 * Near-white floor for the "anyHighlight" keep mode: a background whose every
 * channel is >= this is treated as NOT a highlight (web-paste shading like
 * #f8f9fa must hide — plan A8). Channels are 0-255 here (hex-decoded). Plan A8
 * draws the line at min(r,g,b) >= 0.95; 0.95 * 255 = 242.25, so the smallest
 * integer channel that counts as near-white is 243 — #f2f2f2 (242 ≈ 0.949) sits
 * below the line and stays a kept highlight.
 */
export const NEAR_WHITE_MIN_CHANNEL = 243;

/** gdocs artifact version — deliberately NOT package.json's version, which
 * drives the WORD manifest (plan D14). Surfaced in the sidebar footer.
 * 0.2.0: the analytics style + analytic-ify + Delete analytics (Loop 003).
 * 0.2.1: the "Make a copy" template install path + the static install-page
 *        version pointer in Help (Loop 004). PATCH — no new in-product verb;
 *        the only Code.gs change is that one informational help line.
 * 0.2.2: Hide paints hidden text white (HIDDEN_FG_HEX) on top of the 1pt shrink
 *        for true invisibility; Show All clears it. PATCH — no new verb, an
 *        in-place refinement of Hide / Show All's existing updateTextStyle writes;
 *        backward-compatible (Show All still reveals; ≤0.2.1 1pt-black docs still
 *        restore). */
export const GDOCS_VERSION = "0.2.2";

/**
 * Debate style sizes (Word STYLE_MAP parity, decision #9): pocket 26 / hat 22 /
 * block 16 / tag 14, pocket boxed. Normal = 11 with zeroed spacing (the
 * condensation dependency — plan R1/A12).
 */
export const STYLE_SIZES_PT = {
  normal: 11,
  pocket: 26,
  hat: 22,
  block: 16,
  tag: 14
} as const;

/**
 * The analytics style foreground — a deliberately OFF-PALETTE navy. It is the
 * Docs "dark blue 2" (#0b5394) nudged by two on the blue channel (148 -> 150),
 * so it is visually identical to the template's navy Analytics style yet
 * matches NO one-click Google Docs swatch (proven against settings.ts's
 * DEFAULT_KEEP_HEXES by gdocsKeepers/strings tests). That makes the color a
 * reliable analytics SIGNATURE: Rostrum's "analytic-ify" tool is the only
 * producer, so a user who picks the real "dark blue 2" from the palette is NOT
 * treated as analytics. Detection is EXACT and color-only (size-independent),
 * so resizing analytics text never breaks the keeper/delete match.
 *
 * The off-palette safety is an EMPIRICAL claim pinned by the palette test +
 * confirmed on the live color picker (R6) — not a mathematical invariant; a
 * user who manually types this hex into the custom-color box WOULD be treated
 * as analytics (an accepted, documented edge).
 */
export const ANALYTICS_FG_HEX = "#0b5396";

/** Analytics size — parity with the tag style (decision #9: tag is 14pt). The
 * goal pins analytics to "the same size as the tag style", single-sourced here
 * so the two can never drift. */
export const ANALYTICS_PT = STYLE_SIZES_PT.tag;
