// Domain model for Rostrum's Google Docs invisibility engine.
//
// Everything here is host-agnostic: no Apps Script globals, no DOM. The engine
// reasons about a parsed view of a `documents.get` response and emits plain
// Docs API batchUpdate request JSON; the only thing that ever touches Google is
// a `DocsPort` implementation (gdocs/src/adapter/). This mirrors the Word
// engine's WordPort seam (src/core/types.ts) but is a PARALLEL port shaped for
// Docs' JSON reality, never a contortion of WordPort (repo lesson #48).
//
// Two platform facts shape the whole model (see the loop-001 research record):
//   * Docs has NO hidden-text attribute. "Hide" = shrink to a sentinel font
//     size (constants.SENTINEL_PT), touching ONLY fontSize via field mask, so
//     highlights/bold/links survive untouched and failure is always VISIBLE.
//   * Reversibility lives in the document itself (repo lesson #4): each hidden
//     region is anchored by a NamedRange whose NAME carries the original sizes
//     as RLE (rangeNames.ts). No external store is ever required to restore.

// ---------------------------------------------------------------------------
// Parsed document views (parse.ts output)
// ---------------------------------------------------------------------------

/** Docs named paragraph styles. There are no user-defined or character styles. */
export type GNamedStyleType =
  | "NORMAL_TEXT"
  | "TITLE"
  | "SUBTITLE"
  | "HEADING_1"
  | "HEADING_2"
  | "HEADING_3"
  | "HEADING_4"
  | "HEADING_5"
  | "HEADING_6";

/** A half-open [startIndex, endIndex) UTF-16 index range, as the Docs API uses. */
export interface GRange {
  startIndex: number;
  endIndex: number;
}

/**
 * One paragraph element. ELIGIBILITY IS A WHITELIST (plan A9): only `kind:
 * "text"` (a textRun) may ever be style-targeted; EVERYTHING else — smart
 * chips, page/column breaks, autoText, footnote refs, inline objects,
 * equations, horizontal rules, and any future element type — is `kind:
 * "other"`, is always kept, and breaks hidden regions around it. This is
 * Word's conservative over-keep (decision #16) made structural.
 */
export interface GElement {
  startIndex: number;
  endIndex: number;
  kind: "text" | "other";
  /** Visible text ("" for kind "other"). Length === endIndex - startIndex for text. */
  text: string;
  /** Explicit font size in points, or null when inheriting from the named style. */
  fontSizePt: number | null;
  bold: boolean;
  /**
   * Lower-case "#rrggbb" background color, or null when absent. Docs has no
   * separate highlight channel — highlight IS backgroundColor, which is why
   * the keeper default is a CLOSED palette set, not "any non-white" (plan A8):
   * web-pasted evidence carries near-white shading that must NOT keep.
   */
  backgroundHex: string | null;
}

/** One body paragraph in document order. */
export interface GParagraph {
  /** Ordinal among body paragraphs (receipts count paragraphs, Word parity). */
  index: number;
  /** Paragraph range INCLUDING the trailing newline at endIndex - 1. */
  startIndex: number;
  endIndex: number;
  namedStyleType: GNamedStyleType;
  /** Table content is never touched (Word decision #16 parity). */
  inTable: boolean;
  /**
   * True for the segment's final paragraph: the API rejects styling the final
   * newline, so whole-paragraph hide spans clamp to endIndex - 1 here (plan D6).
   */
  isLastInSegment: boolean;
  /** DIRECT paragraph spacing in points (null = inherits). Drives the optional
   * spacing-collapse class (plan A12) and the styles retro pass (plan A5). */
  spaceAbovePt: number | null;
  spaceBelowPt: number | null;
  elements: GElement[];
}

/** A NamedRange as read from the document (may be split into segments by edits). */
export interface GNamedRange {
  id: string;
  name: string;
  segments: GRange[];
}

/** The parsed document the engine operates on. */
export interface GDoc {
  revisionId: string;
  /**
   * Total tabs in the document. v1 supports exactly one — the guard refuses
   * multi-tab docs OUTRIGHT rather than silently operating on tab 1 (plan A3).
   */
  tabCount: number;
  /** Body-segment paragraphs of the single tab, in document order. */
  paragraphs: GParagraph[];
  /** ALL named ranges (engine filters for rstm: ownership via rangeNames.ts). */
  namedRanges: GNamedRange[];
  /**
   * Per named style: its effective font size in points, when stated. Used to
   * resolve a null (inherited) run size for the cite predicate and receipts.
   */
  namedStyleSizesPt: Partial<Record<GNamedStyleType, number>>;
  /** True when ANY suggested* key appears anywhere in the read (plan A16).
   * Hide and Apply-styles gate on this; Show All deliberately does NOT (A7). */
  suggestionsPresent: boolean;
}

// ---------------------------------------------------------------------------
// Emitted requests — typed mirrors of the exact Docs API batchUpdate JSON.
// The engine only ever emits these shapes; the adapter passes them through
// verbatim. Only fontSize is ever written to runs (case 001-F1); paragraph
// and named-style writes exist solely for the styles/spacing features.
// ---------------------------------------------------------------------------

export interface DocsTextStyle {
  fontSize?: { magnitude: number; unit: "PT" };
  bold?: boolean;
}

export interface UpdateTextStyleRequest {
  updateTextStyle: {
    range: GRange;
    /** Empty textStyle + "fontSize" in fields = CLEAR to inherited (restore of
     * an "i" RLE entry). Never materializes direct formatting (plan v1). */
    textStyle: DocsTextStyle;
    fields: string;
  };
}

export interface CreateNamedRangeRequest {
  createNamedRange: { name: string; range: GRange };
}

export interface DeleteNamedRangeRequest {
  deleteNamedRange: { namedRangeId: string };
}

export interface DocsParagraphStyle {
  namedStyleType?: GNamedStyleType;
  spaceAbove?: { magnitude: number; unit: "PT" };
  spaceBelow?: { magnitude: number; unit: "PT" };
  borderTop?: DocsParagraphBorder;
  borderBottom?: DocsParagraphBorder;
  borderLeft?: DocsParagraphBorder;
  borderRight?: DocsParagraphBorder;
}

/** Paragraph borders must always be specified IN ENTIRETY (API rejects partial). */
export interface DocsParagraphBorder {
  color: { color: { rgbColor: { red?: number; green?: number; blue?: number } } };
  width: { magnitude: number; unit: "PT" };
  padding: { magnitude: number; unit: "PT" };
  dashStyle: "SOLID";
}

export interface UpdateParagraphStyleRequest {
  updateParagraphStyle: {
    range: GRange;
    paragraphStyle: DocsParagraphStyle;
    fields: string;
  };
}

export interface UpdateNamedStyleRequest {
  updateNamedStyle: {
    namedStyle: {
      namedStyleType: GNamedStyleType;
      textStyle?: DocsTextStyle & { fontSize?: { magnitude: number; unit: "PT" } };
      paragraphStyle?: DocsParagraphStyle;
    };
    fields: string;
  };
}

export type DocsRequest =
  | UpdateTextStyleRequest
  | CreateNamedRangeRequest
  | DeleteNamedRangeRequest
  | UpdateParagraphStyleRequest
  | UpdateNamedStyleRequest;

/**
 * An ATOMIC group of requests the chunker may never split (plan A11.viii):
 * e.g. a region's createNamedRange + its updateTextStyle. A torn pair would
 * degrade a partially-applied Hide from exactly-restorable to amber-normalized,
 * so chunk boundaries fall only BETWEEN groups (guards.chunkGroups).
 */
export interface RequestGroup {
  requests: DocsRequest[];
}

// ---------------------------------------------------------------------------
// RLE manifest model (rangeNames.ts)
// ---------------------------------------------------------------------------

/** One run-length entry of a hidden region: `count` chars at `sizePt`
 * (null = the run inherited its size — encoded "i", restored by CLEARING). */
export interface RleEntry {
  count: number;
  sizePt: number | null;
}

/** A contiguous hidden region the planner produced: its doc range + the
 * original sizes as maximal-uniform RLE. One NamedRange anchors each region;
 * names over NAME_MAX chars split the region at entry boundaries (plan A2). */
export interface HiddenRegion {
  start: number;
  end: number;
  entries: RleEntry[];
}

/** Decoded rstm range name. `kind: "sizes"` = rstm:v1:<rle>; `kind: "spacing"`
 * = rstm:v1:p:<above>x<below> (plan A12). Unknown/foreign names decode to null. */
export type DecodedRangeName =
  | { kind: "sizes"; entries: RleEntry[] }
  | { kind: "spacing"; spaceAbovePt: number | null; spaceBelowPt: number | null };

// ---------------------------------------------------------------------------
// Settings (settings.ts) — 3-tier precedence ported from the Word engine:
// doc properties > device default (embedded by adapter) > built-ins.
// ---------------------------------------------------------------------------

export interface GdocsSettings {
  /**
   * "set": keep exactly `keepColors` (closed set of lower-case hexes — the
   * default, Word mental-model parity, plan A8). "anyHighlight": keep any
   * background that is not near-white (the explicit master toggle).
   */
  keepMode: "set" | "anyHighlight";
  keepColors: ReadonlySet<string>;
  /** Cite signature threshold: bold AND >= citeMinPt keeps the paragraph whole.
   * Clamped to <= CITE_PT at resolution (plan A16 minor). */
  citeMinPt: number;
  /** Structural cite rule (plan A10, default ON): first NORMAL_TEXT paragraph
   * after a kept heading with a bold Author-YEAR lead is kept whole. */
  structuralCite: boolean;
  /** Spacing-collapse class (plan A12, default OFF): zero direct spacing on
   * fully-hidden paragraphs, recorded in rstm:v1:p: ranges. */
  collapseSpacing: boolean;
}

// ---------------------------------------------------------------------------
// Results (controller.ts) — receipt-shaped, counted, in the Word taskpane
// voice. Receipts count PARAGRAPHS (Word parity; "runs"/"cards" are banned
// lexicon — plan A11.v).
// ---------------------------------------------------------------------------

export interface GHideResult {
  paragraphsScanned: number;
  /** Paragraphs in which anything was hidden this pass. */
  paragraphsChanged: number;
  /** Hidden regions newly created (anchors written). */
  regionsHidden: number;
  /** Regions that were already hidden and untouched (idempotent re-hide). */
  regionsAlreadyHidden: number;
  /** Sub-spans RESTORED because they became keepers since the last Hide
   * (reconcile semantics, plan A1: new highlight / keep-color / cite change). */
  newlyKeptRestored: number;
  /** Pre-existing sentinel-size text found (counted, never touched by Hide). */
  preexistingTinyCount: number;
}

export interface GShowAllResult {
  /** Range segments restored exactly from their RLE record. */
  segmentsRestoredExact: number;
  /** Segments normalized to inherited size (interior-edited / RLE mismatch). */
  segmentsNormalized: number;
  /** Orphan sentinel-size sub-spans cleared by the convergence sweep. */
  sweptOrphans: number;
  rangesDeleted: number;
  /** rstm-FAMILY ranges this build could not decode — a NEWER engine version's
   * state (edge row 16). They are left entirely untouched (deleting or
   * "restoring" grammar we cannot read could corrupt that version's records),
   * which means hidden text may REMAIN after Show All — so the receipt must
   * warn (strings.showAllReceipt appends an amber line when > 0). */
  rangesSkippedNewerVersion: number;
}

/** Show All may need consent on the PURE-sweep path (no rstm ranges at all —
 * a copied doc or the user's own tiny text; plan A14). The adapter shows the
 * confirm and re-invokes with `sweepUnrecorded: true|false`. */
export type GShowAllOutcome =
  | { kind: "done"; result: GShowAllResult }
  | { kind: "needsConsent"; unrecordedTinyCount: number };

export interface GStylesResult {
  /** False = updateNamedStyle was rejected (consumer-account 400 risk) and only
   * the per-paragraph retro pass ran — the documented degraded path (plan A5/D13).
   * Surfaced by Diagnostics; testable because it lives HERE, not in the adapter. */
  namedStylesApplied: boolean;
  /** Existing paragraphs restyled, per debate style. */
  restyled: { pocket: number; hat: number; block: number; tag: number };
  /** Paragraphs whose direct spacing was cleared (imports carry direct spacing
   * that named-style updates cannot pierce — plan A5). */
  spacingCleared: number;
  /** Docs-native cite leads repaired to the convention (plan A10). */
  citesRepaired: number;
}

// ---------------------------------------------------------------------------
// Errors — every named failure mode maps to exactly one STRINGS entry
// (core/strings.ts); anything else is a DocsApiError (generic refusal copy).
// ---------------------------------------------------------------------------

/** Hide/Apply-styles refuse while suggestions exist. Show All NEVER throws
 * this — the reverse path is always available (Word parity, plan A7). */
export class SuggestionsActiveError extends Error {
  constructor() {
    super("unresolved suggestions present");
    this.name = "SuggestionsActiveError";
  }
}

/** v1 refuses multi-tab docs outright (plan A3). */
export class MultiTabError extends Error {
  constructor(public readonly tabCount: number) {
    super(`document has ${tabCount} tabs`);
    this.name = "MultiTabError";
  }
}

/** Apply-styles refuses while hidden state exists (plan A5). */
export class HiddenStateError extends Error {
  constructor() {
    super("document has hidden text — Show All first");
    this.name = "HiddenStateError";
  }
}

/** Adapter-thrown: batchUpdate 400 with revision-mismatch semantics. */
export class RevisionMismatchError extends Error {
  constructor() {
    super("required revision did not match");
    this.name = "RevisionMismatchError";
  }
}

/** A multi-chunk apply was interrupted after >= 1 chunk landed. The verb's
 * STRINGS entry must be truthful about what already applied (plan A11.iv). */
export class PartialApplyError extends Error {
  constructor(
    public readonly verb: "hide" | "showAll" | "applyStyles",
    public readonly appliedChunks: number,
    public readonly totalChunks: number
  ) {
    super(`${verb} interrupted after ${appliedChunks}/${totalChunks} chunks`);
    this.name = "PartialApplyError";
  }
}

/** All re-plan retries exhausted with nothing applied — doc untouched. */
export class RevisionConflictError extends Error {
  constructor(public readonly verb: "hide" | "showAll" | "applyStyles") {
    super(`${verb}: revision conflict after retries`);
    this.name = "RevisionConflictError";
  }
}

/** Unmapped Docs API rejection. The atomic batch means nothing was applied. */
export class DocsApiError extends Error {
  constructor(message: string) {
    super(message);
    this.name = "DocsApiError";
  }
}

// ---------------------------------------------------------------------------
// The port — the single seam to Google (adapter implements; tests fake).
// ---------------------------------------------------------------------------

export interface DocsPort {
  /**
   * Raw, fields-masked `documents.get(includeTabsContent: true)` JSON (plan
   * A13). Returned as `unknown` ON PURPOSE: parse.ts owns all interpretation,
   * so the in-memory fake must produce REAL response shapes and the parser is
   * the one place reality is decoded (fixture-realism rail, plan A11.ii).
   */
  fetchDocument(): Promise<unknown>;
  /**
   * One atomic batchUpdate guarded by writeControl.requiredRevisionId. Returns
   * the post-apply revision id FROM THE RESPONSE (never a fresh get — A13).
   * Throws RevisionMismatchError on revision 400s; DocsApiError otherwise.
   */
  applyBatch(requests: DocsRequest[], requiredRevisionId: string): Promise<{ revisionId: string }>;
  /** Document-properties settings JSON (settings ONLY — never restore state). */
  readSettingsJson(): Promise<string | null>;
  writeSettingsJson(json: string): Promise<void>;
}
