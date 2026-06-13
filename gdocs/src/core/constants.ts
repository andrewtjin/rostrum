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
 * drives the WORD manifest (plan D14). Surfaced in the sidebar footer. */
export const GDOCS_VERSION = "0.1.0";

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
