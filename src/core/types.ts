// Domain model for Rostrum's invisibility engine.
//
// Everything here is host-agnostic: no Office.js, no React. The engine reasons
// about paragraphs as (normalized heading level + OOXML string); the only thing
// that ever touches Word is a `WordPort` implementation (see officeWordPort.ts,
// Stage 2). Keeping the contract narrow is what makes the whole engine unit
// testable against a fake port.

// Type-only: the cooperative pacing/cancellation contract (core/cancel.ts) is as
// host-agnostic as this file — no Office.js behind it.
import type { Pacer } from "./cancel";

/**
 * Track-changes state, mirroring `Word.ChangeTrackingMode`. We gate Hide when
 * this is anything other than "Off" (decision #14) so tracked-revision noise
 * can't strand a document mid-hide.
 */
export type TrackChangesMode = "Off" | "TrackAll" | "TrackMineOnly";

/**
 * A single body paragraph as handed to the pure engine by a WordPort.
 *
 * `headingLevel` is the engine's canonical, 0-based outline level: 0 = Heading 1
 * ... 8 = Heading 9, and `null` for body text / no outline level. The adapter is
 * responsible for converting Word's `Paragraph.outlineLevel` into this canonical
 * form (the raw Office.js value's base is host-version dependent — see
 * officeWordPort.ts), so the keeper rule here stays unambiguous.
 *
 * `ooxml` is the OOXML for *this paragraph only* — what `paragraph.getRange()
 * .getOoxml()` returns (a flat-OPC package wrapping one `<w:p>`). The engine
 * never assumes more than one `<w:p>` of interest in the string.
 */
export interface RawParagraph {
  /** Stable position of this paragraph in the body, used to write changes back. */
  index: number;
  /** Canonical 0-based outline level (0 = H1 … 8 = H9), or null for body text. */
  headingLevel: number | null;
  /** True if the paragraph lives inside a table cell (kept untouched, decision #16). */
  inTable: boolean;
  /** OOXML for this paragraph (flat-OPC package or bare `<w:p>` fragment). */
  ooxml: string;
  /**
   * Opaque handle for the P1 node-direct path (Loop 002): a paragraph already parsed into
   * a live, attached DOM node so the hide pass can read + mutate it WITHOUT the
   * per-paragraph serialize→parse the string `ooxml` forces. Optional and ADDITIVE — the
   * compat shim is `p.parsed ?? new ParsedParagraph(p.ooxml)`, so every existing caller
   * (which never sets it) keeps the exact string-parse behavior and `parseCount.test.ts`
   * keeps its meaning. Typed `unknown` to avoid a premature circular import
   * (`ParsedParagraph` lives in ooxml.ts, which imports this module); the P1 implementer
   * narrows it at the one site that consumes it. No behavior depends on it until then.
   */
  parsed?: unknown;
}

/**
 * What the classifier decided to do with a paragraph. The adapter uses this to pick
 * the cheapest faithful write (Stage 4 perf): a whole-paragraph NATIVE `font.hidden`
 * toggle for `keepWhole`/`hideWhole` (no OOXML parse or reflow), and an OOXML replace
 * only for `hidePartial`, which needs sub-run targeting Office.js can't do natively.
 */
export type ParagraphAction = "keepWhole" | "hidePartial" | "hideWhole";

/**
 * A paragraph the engine wants Word to update. Always carries the new `ooxml` (used by
 * the whole-body perf path and the `hidePartial` case). `action` is OPTIONAL: when set,
 * the safe per-paragraph adapter path may apply a native `font.hidden` toggle for whole-
 * paragraph cases instead of an expensive `insertOoxml`; when absent it falls back to the
 * OOXML replace (so every pre-Stage-4 caller keeps its exact behavior).
 */
export interface ParagraphUpdate {
  index: number;
  action?: ParagraphAction;
  ooxml: string;
}

/**
 * A lightweight, read-only view of one OOXML run (`<w:r>`), produced by ooxml.ts.
 * The engine never manipulates runs directly — it decides keep/hide per run and
 * hands the flags back to ooxml.ts, which edits the XML.
 */
export interface RunView {
  /** Index among the paragraph's runs, in document order (matches apply order). */
  index: number;
  /** Visible text (w:t joined; w:tab -> "\t"; w:br/w:cr -> "\n"). */
  text: string;
  /** Lower-cased `w:highlight` value, or null when absent / "none". */
  highlight: string | null;
  /** True when the run carries the cite character style (`w:rStyle` == CITE_STYLE_ID). */
  citeStyled: boolean;
  /**
   * True when the run is underlined (`<w:u>` present with a value other than none/0/false). This is
   * Shrink's primary keep-signal: Verbatim cycles only NON-underlined text down a size, keeping the
   * underlined cut readable. Read here so the pure shrink engine can decide keep-full-size per run.
   */
  underline: boolean;
  /** True when the run currently has `<w:vanish/>` (already hidden). */
  hidden: boolean;
  /**
   * False for runs we must never hide (fields, footnote/endnote refs, drawings,
   * embedded objects). Conservative over-keep per decision #16.
   */
  eligible: boolean;
  /**
   * True when the run carries an embedded INTERNAL PART — a `<w:drawing>`, `<w:object>`, or
   * `<w:pict>` anywhere in its subtree (an inline image, OLE object, or VML picture). This is a
   * STRICT SUBSET of the ineligibility signal (`eligible` is additionally false for fields and
   * footnote/endnote refs, which carry no part), surfaced separately so the node-direct hide path
   * can recognize "this paragraph references a media part" WITHOUT the per-paragraph serialize the
   * string path used (Loop 002 B1 rider): an internal-part paragraph must never be re-serialized,
   * or it silently re-pays the cost P1 deletes. Read by the same fused traversal that fills the
   * other fields, so it costs nothing extra; the legacy string path computes it identically.
   */
  hasInternalPart: boolean;
}

/**
 * An instruction to expose ONE already-present space inside an otherwise-hidden
 * run, so two kept chunks separated by hidden prose don't fuse in the condensed
 * view (e.g. "radiation" + hidden " are a constant threat, " + "would" must read
 * "radiation would", not "radiationwould" — wet-test bug 1).
 *
 * The space is MOVED (not inserted) so the paragraph's concatenated text stays
 * byte-identical, keeping native Font-dialog reversal (and Show All) perfectly
 * lossless: a recipient without the add-in gets the exact original back, with no
 * fabricated spaces. "lead"/"trail" move the run's first/last space into a visible
 * sibling. "interior" handles a hidden run whose only spaces are interior (e.g.
 * ", such as scorpions." — starts "," and ends ".") by splitting it in three —
 * [before](hidden) [space](visible) [after](hidden) — at `offset`, the chosen
 * space's character index (wet-test bug 2 follow-up). With show-only-highlighted both
 * fragments stay hidden wherever we split, and Re-hide rescues the exposed space as a
 * whitespace-only run, so the split is idempotent.
 */
export interface BridgeSplit {
  /** Index (in readRuns order) of the hidden run to split. */
  index: number;
  /** Which space to expose: the run's leading, trailing, or an interior one. */
  side: "lead" | "trail" | "interior";
  /** For `side: "interior"`: character index (in the run's text) of the space to expose. */
  offset?: number;
}

/** Manifest persisted as a document-level custom XML part (decisions #10, #13, #15). */
export interface RostrumManifest {
  /** Whether invisibility is currently ON for this document. */
  active: boolean;
  /** Highlight colors that count as "keep". Stored so any machine re-derives identically. */
  keepColors: string[];
  /** Manifest schema version, for forward migration. */
  schemaVersion: number;
}

/** Fully-resolved settings the engine actually runs with. */
export interface ResolvedSettings {
  /** Highlight color names (lower-cased) that mean "keep this run". */
  keepColors: ReadonlySet<string>;
}

/** Per-device defaults cached in localStorage (labeled per-device, NOT roaming — decision #15). */
export interface DeviceDefaults {
  keepColors: string[];
}

/** Result of a hide / re-hide pass, surfaced to the task pane and asserted in tests. */
export interface HideResult {
  paragraphsScanned: number;
  paragraphsChanged: number;
  /** Paragraphs left unchanged because their OOXML failed to parse (the UI warns). */
  paragraphsSkipped: number;
  /** True when Track Changes was auto-toggled off for the operation and restored. */
  trackChangesToggled: boolean;
}

/** Result of a show-all pass. */
export interface ShowAllResult {
  paragraphsScanned: number;
  paragraphsChanged: number;
  /** Paragraphs left unchanged because their OOXML failed to parse. */
  paragraphsSkipped: number;
}

/** Options shared by the mutating orchestrators. */
export interface HideOptions {
  /**
   * When Track Changes is ON: if true, toggle it Off for the operation and
   * restore afterward; if false (default), throw TrackChangesActiveError so the
   * UI can prompt first (decision #14).
   */
  autoToggleTrackChanges?: boolean;
  /**
   * Paces the classify loop: `tick()` is awaited once per paragraph. It throws
   * `CancelledError` to abort BEFORE anything is written (updates are buffered in
   * the adapter and never flushed; the Track-Changes gate's `finally` restores the
   * prior mode), and it yields a macrotask when its time budget elapses so a live
   * task pane stays paintable and cancellable through this otherwise-synchronous
   * pure-JS phase. Optional — headless/engine-direct callers (tests, benchmarks)
   * omit it and pay zero overhead.
   */
  pacing?: Pacer;
}

/** Which optional capabilities the current host advertises (decision #18). */
export interface FeatureSupport {
  /** font.hidden / <w:vanish/> — WordApiDesktop 1.2. Hard requirement. */
  canHide: boolean;
  /** customXmlParts — WordApi 1.4. Hard requirement (manifest store). */
  canCustomXml: boolean;
  /** changeTrackingMode read/write — WordApi 1.4. */
  canChangeTracking: boolean;
  /** Style.borders — WordApiDesktop 1.1 (pocket box, Stage 2). */
  canStyleBorders: boolean;
  /** Style.font / Style.paragraphFormat — WordApi 1.5 (Apply-styles sizes, Stage 2). */
  canStyleFormat: boolean;
  /**
   * document.getStyles() — WordApi 1.5. Required to enumerate and edit the built-in
   * heading + cite styles for "Apply Rostrum styles" (Stage 2). This is the async
   * getStyles() METHOD (WordApi 1.5), NOT the synchronous Document.styles property
   * (WordApiDesktop 1.4); Rostrum calls the former, so the floor is WordApi 1.5 —
   * the same set as canStyleFormat, but kept a distinct flag for clarity.
   */
  canGetStyles: boolean;
}

/**
 * The single seam between the pure engine and Word. A real implementation
 * (officeWordPort.ts) wraps `Word.run`; tests supply an in-memory fake.
 */
export interface WordPort {
  /** Read the document's current change-tracking mode. */
  getChangeTrackingMode(): Promise<TrackChangesMode>;
  /** Set the document's change-tracking mode. */
  setChangeTrackingMode(mode: TrackChangesMode): Promise<void>;
  /** Enumerate body-story paragraphs (footnotes/headers/etc. excluded — decision #16). */
  readParagraphs(): Promise<RawParagraph[]>;
  /** Replace the OOXML of the given paragraphs (batched + synced by the adapter). */
  writeParagraphs(updates: ParagraphUpdate[]): Promise<void>;
  /**
   * Reveal ALL hidden text in the body story by clearing `font.hidden` natively — the
   * fast Show All path (Stage 4). NO OOXML read or write: one or two host round-trips
   * instead of thousands of `insertOoxml` reflows. Behavior-identical to the old
   * per-paragraph makeAllVisible pass (it reveals the same set, including any run the
   * user hid manually — decision #10) and convergent (safe to re-run from any state).
   * Returns how many body paragraphs were in scope and how many it changed.
   */
  clearHidden(): Promise<{ paragraphsScanned: number; paragraphsChanged: number }>;
  /** Return the Rostrum manifest XML, or null when the document has none. */
  readManifest(): Promise<string | null>;
  /** Create or overwrite the Rostrum manifest custom XML part. */
  writeManifest(xml: string): Promise<void>;
  /** Remove the Rostrum manifest custom XML part (no-op when absent). */
  clearManifest(): Promise<void>;
}

// ===========================================================================
// Condense & Shrink — domain model (Rostrum's lossless answer to Verbatim's
// Shrink + Condense). These run over a RANGE FRAGMENT (the active selection, or
// the current paragraph when collapsed), not the whole-body paragraph array, so
// they live behind the separate `RangeScopedPort` below — the invisibility
// engine's one-`<w:p>`-per-write invariant is deliberately untouched.
// ===========================================================================

/**
 * A richer per-run view than `RunView`, used by the fragment transforms: it adds the run's explicit
 * font size and whether it is a condense break marker. `readFragmentParagraphs` produces these so the
 * pure Shrink engine can compute the size ladder and skip markers, and Uncondense can find boundaries.
 */
export interface FragmentRunView extends RunView {
  /** Explicit `<w:sz>` value in half-points on this run, or null when it inherits from the style. */
  sizeHalfPts: number | null;
  /** True when this run is a condense marker — identified by an intrinsic text signature in `MARK_SIGNATURES` (current ZWSP+WJ pair or legacy U+2063). */
  breakMarker: boolean;
  /**
   * True when the run is character-boxed (`<w:bdr>` with a real border) — directly OR through its
   * character style. Shrink keeps boxed runs full-size alongside underlined ones (debate emphasis).
   * NOTE: for this fragment view, BOTH `underline` and `boxed` are STYLE-RESOLVED — in real docs the
   * cut/emphasis is applied through a character style (StyleUnderline / Emphasis), not a direct rPr — so
   * resolving only the run's own `<w:u>`/`<w:bdr>` (as the invisibility reader does) misses the signal.
   */
  boxed: boolean;
}

/**
 * How a Condense operation handles reversibility — the single seam that flips lossless ↔ destructive.
 *   * `"marker"` (default): each former paragraph boundary becomes a signature-tagged marker run
 *     (`MARK_SIGNATURE`), so Uncondense is an exact inverse. Lossless.
 *   * `"none"`: boundaries collapse to a plain space with no marker — Verbatim-style destructive,
 *     allowed ONLY when pilcrows are off (a visible pilcrow IS a marker). Faster, not reversible.
 */
export type ReversalStrategy = "marker" | "none";

/** The fully-resolved knobs a single Condense run executes with (resolved from settings + the mode). */
export interface CondenseOptions {
  /** Render each boundary marker as a visible `¶` at 6pt (Verbatim parity) vs an invisible hidden run. */
  usePilcrows: boolean;
  /** Keep paragraph structure and only drop blank/whitespace-only paragraphs, instead of merging all. */
  retainParagraphs: boolean;
  /** Lossless marker boundaries vs destructive plain-space (`"none"` requires `usePilcrows: false`). */
  reversal: ReversalStrategy;
}

/** Result of a Condense pass over the active range. */
export interface CondenseResult {
  /** Paragraphs the range held before the operation. */
  paragraphsScanned: number;
  /** Boundary markers inserted (merge mode) or blank paragraphs removed (retain mode). */
  boundariesMarked: number;
  /** True when the OOXML actually changed (a no-op range reports false). */
  changed: boolean;
}

/** Result of an Uncondense pass over the active range. */
export interface UncondenseResult {
  /** Boundary markers turned back into paragraph breaks. */
  breaksRestored: number;
  changed: boolean;
}

/**
 * A user-configurable omission marker: a bracketed region whose text Shrink restores to full size so
 * an "[…Omitted…]" indicator stays readable in a shrunk card. A span runs from `open` to the next
 * `close`; it counts as an omission only when the text between them contains `keyword` (case-
 * insensitive) — so ordinary bracketed prose ("[sic]", "[their]") is NOT un-shrunk. Defaults mirror
 * Verbatim's set ([…Omitted…], [[…Omitted…]], <…Omitted…>).
 */
export interface OmissionPattern {
  open: string;
  close: string;
  keyword: string;
}

/** Per-device Condense & Shrink settings (localStorage, like the keep-color defaults). */
export interface CondenseSettings {
  /** Default Condense marker style: visible `¶` (true) vs invisible hidden run (false). */
  usePilcrows: boolean;
  /** Default Condense structure handling: drop-blank-paragraphs (true) vs merge-all (false). */
  retainParagraphs: boolean;
  /** Default reversibility strategy. `"none"` is honored only when `usePilcrows` is false. */
  reversal: ReversalStrategy;
  /** Also shrink each paragraph mark to 6pt when running Shrink (Verbatim's "Shrink ¶" extra). */
  shrinkParagraphMarks: boolean;
  /** Omission markers Shrink restores to full size. */
  omissionPatterns: OmissionPattern[];
}

/** Options the pure Shrink engine runs with for one press of Shrink. */
export interface ShrinkOptions {
  /** Resolved Normal/default font size in half-points (from the package styles), used by the ladder. */
  normalHalfPts: number;
  /** Canonical 0-based outline level per paragraph in the fragment (null = body), for heading refusal. */
  outlineLevels: (number | null)[];
  /** Omission markers whose spans are restored to Normal size. */
  omissionPatterns: OmissionPattern[];
  /** Also shrink each paragraph mark to 6pt (Verbatim's "Shrink ¶"). */
  shrinkParagraphMarks: boolean;
}

/** Result of a Shrink / Unshrink pass over the active range. */
export interface ShrinkResult {
  paragraphsScanned: number;
  /** True when the OOXML actually changed. */
  changed: boolean;
  /**
   * The half-point size Shrink applied to non-kept runs this press, or null when it cleared sizes
   * (the "→ Normal" rung, and every Unshrink). The pane shows this as the current shrink size.
   * Undefined when nothing was eligible to size (e.g. a heading-only refusal).
   */
  appliedSizeHalfPts?: number | null;
  /** Set when a collapsed selection on a heading refused to shrink (Verbatim parity). */
  refusedHeading?: boolean;
}

/**
 * What a `readActiveRangeOoxml` returns: the active range's OOXML plus the metadata the engines need.
 * `outlineLevels` is canonical 0-based (0 = Heading 1 … 8 = Heading 9, null = body), aligned to the
 * range's paragraphs in document order. `collapsed` is true when the selection was an insertion point
 * (so the operation targets just the current paragraph).
 */
export interface RangeRead {
  ooxml: string;
  collapsed: boolean;
  outlineLevels: (number | null)[];
}

/**
 * The SECOND document-access seam (beside `WordPort`), for the range-scoped Condense & Shrink ops.
 * It reads the active selection (or current paragraph) as ONE OOXML fragment and replaces it whole —
 * which is exactly what Condense's paragraph merge needs and what `WordPort`'s single-`<w:p>` write
 * contract forbids by design. Track-Changes is gated through the same two methods Hide uses, so the
 * shared gate keeps a partial Undo from stranding the document.
 */
export interface RangeScopedPort {
  /** Read the document's current change-tracking mode (for the shared TC gate). */
  getChangeTrackingMode(): Promise<TrackChangesMode>;
  /** Set the document's change-tracking mode (for the shared TC gate). */
  setChangeTrackingMode(mode: TrackChangesMode): Promise<void>;
  /** Read the active selection — or the current paragraph when collapsed — as one OOXML fragment. */
  readActiveRangeOoxml(): Promise<RangeRead>;
  /** Replace the active range with new OOXML (`insertOoxml(…, "Replace")`). */
  replaceActiveRangeOoxml(ooxml: string): Promise<void>;
}
