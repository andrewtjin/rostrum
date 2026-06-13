// Pure adapter-side decision logic (plan A4/A11.i, step S11).
//
// WHY THIS MODULE EXISTS: the host adapter (gdocs/src/adapter/docsAdapter.ts)
// is excluded from unit coverage because it cannot run without the Apps Script
// host — so EVERY decision the adapter would otherwise make lives HERE, in
// core, where root tsc and the gdocs coverage floor both see it. The adapter
// keeps only Google-global calls plus try/catch dispatch; this module owns:
//
//   * the public entry-point CALL MAP (one constant feeding BOTH the onOpen
//     menu and the build's top-level function shims — plan S12, the contract
//     tools/build-gdocs.mjs extracts via its /call.?map/i export finder);
//   * batchUpdate error classification (revision mismatch vs everything else);
//   * error -> dialog routing (so the adapter renders, never decides);
//   * receipt -> dialog shaping for every verb (severity included);
//   * the sidebar state-line model from the fields-masked state read;
//   * Mark-cite planning from host-extracted selection picks (gates + the A9
//     whitelist pre-split planMarkCite cannot apply itself);
//   * Diagnostics probe interpretation + the copy-pasteable report text.
//
// PURITY CONTRACT: no Apps Script globals, no I/O, no Date/random — total
// functions over given values, exactly like the rest of core/.

import { GDOCS_VERSION } from "./constants";
import { assertNoSuggestions, assertNotHidden, assertSingleTab } from "./guards";
import { parseDocument } from "./parse";
import { decodeRangeName, isRstmName } from "./rangeNames";
import {
  errorMessage,
  hideReceipt,
  markCiteReceipt,
  showAllReceipt,
  STRINGS,
  stylesReceipt
} from "./strings";
import { planMarkCite } from "./styles";
import {
  DocsApiError,
  DocsRequest,
  GDoc,
  GHideResult,
  GNamedStyleType,
  GRange,
  GShowAllResult,
  GStylesResult,
  RevisionMismatchError
} from "./types";

// ---------------------------------------------------------------------------
// The call map (plan S12) — the single source of the product's entry surface.
// ---------------------------------------------------------------------------

/**
 * One public entry point. `fn` is the TOP-LEVEL global function name the build
 * shims and the menu both reference; `label` is the STRINGS.menu copy for menu
 * items and null for entries Apps Script reaches another way (the onOpen
 * simple trigger; google.script.run targets the sidebar calls).
 */
export interface CallMapEntry {
  fn: string;
  label: string | null;
  /** Menu grouping: render a separator BEFORE this item (frontendDraft Step 3
   * groups: verbs / tools / panel+help+diagnostics). Ignored for label:null. */
  separatorBefore: boolean;
}

/**
 * Every public entry point, in menu order for the labeled ones. The build
 * (tools/build-gdocs.mjs) generates one global function shim per entry and the
 * adapter builds the onOpen menu from the same constant, so the menu and the
 * entry surface can never drift (plan S12). Labels come ONLY from STRINGS.menu
 * — the deck is the single legal source of user copy.
 */
export const CALL_MAP: readonly CallMapEntry[] = [
  // The simple trigger — not a menu item; Apps Script calls it on open.
  { fn: "onOpen", label: null, separatorBefore: false },
  // Verbs first, Hide/Show All always adjacent (mid-round muscle memory).
  { fn: "rostrumHide", label: STRINGS.menu.hide, separatorBefore: false },
  { fn: "rostrumShowAll", label: STRINGS.menu.showAll, separatorBefore: false },
  // Tools group.
  { fn: "rostrumApplyStyles", label: STRINGS.menu.applyStyles, separatorBefore: true },
  { fn: "rostrumMarkCite", label: STRINGS.menu.markCite, separatorBefore: false },
  // Panel / help / diagnostics group.
  { fn: "rostrumOpenPanel", label: STRINGS.menu.openPanel, separatorBefore: true },
  { fn: "rostrumHelp", label: STRINGS.menu.helpShortcuts, separatorBefore: false },
  { fn: "rostrumDiagnostics", label: STRINGS.menu.diagnostics, separatorBefore: false },
  // google.script.run targets (sidebar) — top-level by Apps Script rule, so
  // they need shims too, but they never appear in the menu.
  { fn: "rostrumSidebarState", label: null, separatorBefore: false },
  { fn: "rostrumHideFromSidebar", label: null, separatorBefore: false },
  { fn: "rostrumShowAllFromSidebar", label: null, separatorBefore: false },
  { fn: "rostrumSaveSettings", label: null, separatorBefore: false }
];

// ---------------------------------------------------------------------------
// Adapter constants — defined here (not in the adapter) so they are visible to
// tests and pinned against drift.
// ---------------------------------------------------------------------------

/**
 * The one Properties key both tiers use (DocumentProperties = this doc's
 * settings; UserProperties = the device default "Save as my default" writes).
 * Persisted into users' documents — changing it would orphan saved settings,
 * so it is pinned by test like rangeNames.RSTM_PREFIX.
 */
export const SETTINGS_PROPERTY_KEY = "rostrum.settings";

/**
 * The sidebar state line's fields mask (plan A13): namedRanges + revisionId
 * ONLY — the state read must stay tiny because the sidebar refreshes it after
 * every verb. No includeTabsContent: the legacy top-level namedRanges field is
 * where single-tab docs (the only docs the verbs accept — plan A3) report
 * them, and parseDocument's legacy path reads exactly that shape.
 */
export const STATE_FIELDS_MASK = "revisionId,namedRanges";

// ---------------------------------------------------------------------------
// batchUpdate error classification
// ---------------------------------------------------------------------------

/**
 * MATCHED SHAPE, documented (plan D5): the Docs API rejects a stale
 * writeControl.requiredRevisionId with a 400 whose message names the revision
 * — Apps Script's advanced service surfaces it as an exception reading like
 * "API call to docs.documents.batchUpdate failed with error: The revision
 * specified in the write control is not the most recent revision of the
 * document." There is no structured code on the thrown value, only message
 * text, so the classifier keys on the word "revision" (case-insensitive).
 * False positives are structurally impossible from OUR traffic: no request
 * shape the engine emits can produce any OTHER 400 that mentions revisions.
 */
const REVISION_MISMATCH_MESSAGE = /revision/i;

/**
 * Classify a raw throw from Docs.Documents.batchUpdate into the two error
 * classes the controller's retry protocol distinguishes: RevisionMismatchError
 * (re-plan / partial-apply handling) vs DocsApiError (generic refusal copy).
 * Non-Error throws stringify — the classifier must never itself throw.
 */
export function classifyBatchError(raw: unknown): RevisionMismatchError | DocsApiError {
  const message = raw instanceof Error ? raw.message : String(raw);
  if (REVISION_MISMATCH_MESSAGE.test(message)) return new RevisionMismatchError();
  return new DocsApiError(message);
}

// ---------------------------------------------------------------------------
// Dialog routing — the adapter renders these, never composes them.
// ---------------------------------------------------------------------------

/** "plain" = a neutral receipt; "amber" = degraded-but-healthy; "red" is
 * reserved for nothing-was-applied refusals (frontendDraft Step 2). */
export type DialogSeverity = "plain" | "amber" | "red";

/** What the adapter shows: a modal receipt or refusal (menu path) — the same
 * object renders inline in the sidebar's receipt region (single source, DRY). */
export interface RoutedDialog {
  dialog: "receipt" | "refusal";
  title: string;
  body: string;
  severity: DialogSeverity;
}

/**
 * Map ANY thrown value to its dialog. strings.errorMessage owns the copy and
 * the severity; this only adds the dialog kind — red means a refusal (nothing
 * healthy to receipt), anything milder would be receipt-shaped. Today every
 * mapped error is red, so the receipt branch is future-proofing, not policy.
 */
export function routeError(e: unknown): RoutedDialog {
  const m = errorMessage(e);
  return { dialog: m.severity === "red" ? "refusal" : "receipt", title: m.title, body: m.body, severity: m.severity };
}

/** Shared receipt shell — every receipt carries the one receipt title. */
function receipt(body: string, severity: DialogSeverity): RoutedDialog {
  return { dialog: "receipt", title: STRINGS.dialogs.receiptTitle, body, severity };
}

/** Hide's receipt is always plain: a reconcile that surfaced keepers is
 * information, not degradation (the doc is exactly as the keeper rules say). */
export function hideDialog(r: GHideResult): RoutedDialog {
  return receipt(hideReceipt(r), "plain");
}

/** Show All goes amber whenever anything came back at OTHER than its saved
 * size — normalized segments, swept orphans, or newer-version skips (plan D3 /
 * edge row 16 — the same buckets the receipt copy itself reports). */
export function showAllDialog(r: GShowAllResult): RoutedDialog {
  const amber = r.segmentsNormalized + r.sweptOrphans + r.rangesSkippedNewerVersion > 0;
  return receipt(showAllReceipt(r), amber ? "amber" : "plain");
}

/** Apply-styles goes amber on the documented degraded path (plan A5/D13):
 * existing paragraphs styled, future typing not guaranteed. */
export function stylesDialog(r: GStylesResult): RoutedDialog {
  return receipt(stylesReceipt(r), r.namedStylesApplied ? "plain" : "amber");
}

/** Mark-cite receipt: counted, or the teach-don't-scold empty-selection line. */
export function markCiteDialog(citedParagraphs: number): RoutedDialog {
  return receipt(markCiteReceipt(citedParagraphs), "plain");
}

// ---------------------------------------------------------------------------
// Sidebar state (plan A13 — the namedRanges+revisionId-only read)
// ---------------------------------------------------------------------------

export interface SidebarStateModel {
  /** Hidden passages for the state line: one per "sizes" record, PLUS every
   * rstm-family range this build cannot decode (a newer version's state is
   * still hidden text the user should see counted). SPACING records are
   * excluded — each parallels a sizes record for the same passage, and double
   * counting would inflate the line. */
  hiddenRegionCount: number;
  /** ANY rstm-family range present (the guards.hasRstmState predicate over
   * the state read): drives the sidebar's armed/clean rendering. */
  armed: boolean;
}

/** Interpret the raw state read. parseDocument is defensive end to end, so
 * junk in = a clean zero-state out, never a sidebar crash. */
export function sidebarState(rawMaskedGet: unknown): SidebarStateModel {
  const doc = parseDocument(rawMaskedGet);
  let hiddenRegionCount = 0;
  let armed = false;
  for (const nr of doc.namedRanges) {
    if (!isRstmName(nr.name)) continue; // foreign add-ons' ranges are invisible
    armed = true;
    const decoded = decodeRangeName(nr.name);
    if (decoded === null || decoded.kind === "sizes") hiddenRegionCount++;
  }
  return { hiddenRegionCount, armed };
}

// ---------------------------------------------------------------------------
// Apply-styles first-run confirm support
// ---------------------------------------------------------------------------

/** The H1-4 ladder Apply-styles retro-restyles (mirror of styles.ts's private
 * DEBATE_HEADINGS keys — pinned by test so the two can never drift). */
const DEBATE_HEADING_STYLES: ReadonlySet<GNamedStyleType> = new Set<GNamedStyleType>([
  "HEADING_1",
  "HEADING_2",
  "HEADING_3",
  "HEADING_4"
]);

/**
 * How many existing headings Apply-styles would restyle — the blast radius the
 * first-run confirm states (frontendDraft Step 5; zero skips the dialog).
 * Table headings are excluded because planApplyStyles never touches tables
 * (edge row 1), so counting them would overstate the radius.
 */
export function countDebateHeadings(raw: unknown): number {
  const doc = parseDocument(raw);
  let count = 0;
  for (const p of doc.paragraphs) {
    if (!p.inTable && DEBATE_HEADING_STYLES.has(p.namedStyleType)) count++;
  }
  return count;
}

// ---------------------------------------------------------------------------
// Mark cite (plan A10(c) / step S11) — picks in, gated request plan out.
// ---------------------------------------------------------------------------

/**
 * One piece of the user's selection, as the adapter extracts it from
 * DocumentApp (the host API exposes paragraph-relative facts only — absolute
 * indexes exist solely in the parsed view, so the join happens here).
 * Offsets are API-INDEX units from the paragraph's startIndex (the adapter
 * already charges 1 unit per non-text sibling, matching Docs index rules).
 */
export interface SelectionPick {
  /** Ordinal in GDoc.paragraphs (the parse's flattened document order). */
  paragraphOrdinal: number;
  /** Inclusive start offset within the paragraph. */
  startOffset: number;
  /** EXCLUSIVE end offset; null = through the end of the paragraph. */
  endOffset: number | null;
}

/**
 * API-unit arithmetic for lowering a host text selection to a SelectionPick.
 * The DocumentApp tree-walk that produces these inputs is host-only and lives
 * in docsAdapter.ts (untestable without a live Doc); the index math it feeds is
 * isolated here so the unit gate exercises it directly. A wrong sibling sum or
 * a missed inclusive-to-exclusive conversion is the one off-by-one that would
 * silently mis-span a cite, so both are pinned by tests.
 */

/** Docs-API index units occupied by the siblings PRECEDING a target element:
 * each preceding TEXT sibling counts its character length; every other sibling
 * (inline image, chip, break) occupies exactly one index unit. */
export function sumApiUnits(precedingSiblings: ReadonlyArray<{ isText: boolean; textLength: number }>): number {
  let units = 0;
  for (const sibling of precedingSiblings) units += sibling.isText ? sibling.textLength : 1;
  return units;
}

/** The shape of a host RangeElement within one paragraph: a PARTIAL range
 * carries INCLUSIVE char offsets inside the Text run; a whole-element pick
 * spans the entire run. */
export type TextRangeShape =
  | { partial: true; startOffsetInText: number; endOffsetInclusiveInText: number }
  | { partial: false; textLength: number };

/** Convert a host text RangeElement to paragraph-relative API-unit offsets.
 * `before` is sumApiUnits(the element's preceding siblings). RangeElement end
 * offsets are INCLUSIVE while the SelectionPick model is half-open, hence the
 * +1 on the partial path. */
export function textPickOffsets(
  before: number,
  shape: TextRangeShape
): { startOffset: number; endOffset: number } {
  if (shape.partial) {
    return {
      startOffset: before + shape.startOffsetInText,
      endOffset: before + shape.endOffsetInclusiveInText + 1
    };
  }
  return { startOffset: before, endOffset: before + shape.textLength };
}

export interface MarkCitePlan {
  /** The revision the one batchUpdate must be guarded with (plan D4). */
  revisionId: string;
  /** Zero requests = nothing markable (empty/chip-only selection): the
   * adapter applies nothing and shows the markCiteNoop receipt. */
  requests: DocsRequest[];
  /** Distinct paragraphs that received a write — the receipt's unit ("Marked
   * 1 cite." — a cite is a line, not a request). */
  citedParagraphs: number;
}

/**
 * The text-only ranges of one pick — THE WHITELIST PRE-SPLIT (plan A9):
 * planMarkCite writes verbatim over whatever range it is handed, so chips,
 * breaks, and objects must be cut out HERE. Adjacent text ranges coalesce
 * (fewer requests, same effect); the segment-final newline is clamped off
 * exactly as detectCiteLeads does (the API refuses to style it). A paragraph's
 * NON-final trailing newline stays markable — sizing a paragraph mark is
 * benign and matches what the Docs toolbar itself would do (keepers.ts's
 * documented stance). Table paragraphs are allowed: row 1's "tables untouched"
 * protects the BULK passes; an explicit user selection wins (styles.ts's
 * documented Mark-cite stance).
 */
function citeRangesForPick(doc: GDoc, pick: SelectionPick): GRange[] {
  const p = doc.paragraphs[pick.paragraphOrdinal];
  // Out-of-range ordinal (the host tree and the parse disagreed — possible on
  // exotic structure): skip the pick rather than write at a guessed index.
  if (p === undefined) return [];
  const absStart = p.startIndex + Math.max(0, pick.startOffset);
  const pickEnd = pick.endOffset === null ? p.endIndex : Math.min(p.endIndex, p.startIndex + pick.endOffset);
  // The unstylable segment-final newline is the paragraph's very last index.
  const styleCeiling = p.isLastInSegment ? p.endIndex - 1 : p.endIndex;
  const absEnd = Math.min(pickEnd, styleCeiling);

  const ranges: GRange[] = [];
  for (const el of p.elements) {
    if (el.kind !== "text") continue; // the whitelist: chips/objects never styled
    const s = Math.max(absStart, el.startIndex);
    const e = Math.min(absEnd, el.endIndex);
    if (e <= s) continue;
    const last = ranges[ranges.length - 1];
    // Coalesce only true adjacency — a chip between two text elements leaves
    // an index gap that correctly keeps the ranges separate.
    if (last !== undefined && last.endIndex === s) last.endIndex = e;
    else ranges.push({ startIndex: s, endIndex: e });
  }
  return ranges;
}

/**
 * Plan the Mark-cite verb from the raw verb read + the host's selection picks.
 * Gates mirror the styles lane (styles.ts's module contract): single tab
 * (plan A3), suggestion-free (indexes — plan D5), and NOT hidden (writing
 * CITE_PT over a hidden span would fight the restore records — plan A5).
 * Each throws its named error class for strings.errorMessage to map. The
 * returned requests are planMarkCite's exact emission per range — ONE shared
 * builder with Apply-styles' repair pass, so the convention can never fork.
 */
export function planMarkCiteFromPicks(raw: unknown, picks: readonly SelectionPick[]): MarkCitePlan {
  const doc = parseDocument(raw);
  assertSingleTab(doc);
  assertNoSuggestions(doc);
  assertNotHidden(doc);

  const requests: DocsRequest[] = [];
  const paragraphsTouched = new Set<number>();
  for (const pick of picks) {
    const ranges = citeRangesForPick(doc, pick);
    if (ranges.length === 0) continue;
    paragraphsTouched.add(pick.paragraphOrdinal);
    for (const range of ranges) requests.push(...planMarkCite(range).requests);
  }
  return { revisionId: doc.revisionId, requests, citedParagraphs: paragraphsTouched.size };
}

// ---------------------------------------------------------------------------
// Diagnostics (plan D12/A11.ii/A13) — probe interpretation + the report.
// ---------------------------------------------------------------------------

/** The sub-1pt sizes the font-floor probe tries (plan D13: SENTINEL_PT stays
 * 1 in v1; these readings inform whether v2 can go lower). */
export const FONT_FLOOR_TRY_SIZES_PT: readonly number[] = [0.25, 0.5, 0.75];

/** Reports at or under this embed whole-doc JSON (plan A11.ii's "copy
 * small-doc JSON" probe — the wet round captures one real dump for fixtures).
 * 20k keeps the textarea copy-pasteable through chat without truncation. */
export const SMALL_DOC_DUMP_MAX_BYTES = 20000;

/** Histogram width: enough hexes to spot a wrong keep-default (plan A8's
 * stated purpose) without drowning the report in tint noise. */
export const BG_HISTOGRAM_TOP_N = 8;

/** What the adapter observed for one font-floor try (raw, uninterpreted). */
export interface FontFloorReading {
  triedPt: number;
  /** False = the batchUpdate itself was rejected. */
  appliedOk: boolean;
  /** The size read back from the probe char after the write (null = the
   * read-back could not state a size). Meaningless when appliedOk is false. */
  readBackPt: number | null;
}

export type FontFloorVerdict = "accepted exact" | "clamped" | "rejected";

/** One interpreted font-floor row for the report model. */
export interface FontFloorResult {
  triedPt: number;
  readBackPt: number | null;
  verdict: FontFloorVerdict;
}

/**
 * Interpret one reading: rejected (the API refused the write), accepted exact
 * (read-back equals the tried size), else clamped — the write landed but the
 * document reports a different (or unreadable) size, i.e. the API silently
 * snapped it to its floor.
 */
export function interpretFontFloor(reading: FontFloorReading): FontFloorResult {
  if (!reading.appliedOk) return { triedPt: reading.triedPt, readBackPt: null, verdict: "rejected" };
  if (reading.readBackPt !== null && reading.readBackPt === reading.triedPt) {
    return { triedPt: reading.triedPt, readBackPt: reading.readBackPt, verdict: "accepted exact" };
  }
  return { triedPt: reading.triedPt, readBackPt: reading.readBackPt, verdict: "clamped" };
}

/** "skipped" = the probe never ran (e.g. the read stated no HEADING_6 size,
 * so a redefine could have visibly changed the doc — exactly what the probe
 * promises not to do). */
export type NamedStyleProbeVerdict = "ok" | "rejected" | "skipped";
export type InheritClearVerdict = "ok" | "failed" | "skipped";

/** Everything the report needs from the verb-style document read, in one pure
 * call so the adapter never touches parse itself. */
export interface DiagnosticsReadFacts {
  revisionId: string;
  /** Where the probe char goes: just before the body's final newline (the
   * very end of visible content — self-cleaning insert+delete, plan D12). */
  probeIndex: number;
  /** HEADING_6's CURRENT stated size, or null when the read omitted it. The
   * namedStyle probe rewrites this exact value so nothing visibly changes. */
  headingSixSizePt: number | null;
}

export function diagnosticsReadFacts(raw: unknown): DiagnosticsReadFacts {
  const doc = parseDocument(raw);
  const last = doc.paragraphs[doc.paragraphs.length - 1];
  // Empty/junk reads degrade to index 1 (the start of any real body) — the
  // insert then fails loudly server-side instead of corrupting math here.
  const probeIndex = last !== undefined ? Math.max(1, last.endIndex - 1) : 1;
  return {
    revisionId: doc.revisionId,
    probeIndex,
    headingSixSizePt: doc.namedStyleSizesPt.HEADING_6 ?? null
  };
}

/** The probe char's size as a fresh read reports it: the text element whose
 * range contains `index`, or null (gone, or not a text element). */
export function probeReadbackSizePt(raw: unknown, index: number): number | null {
  const doc = parseDocument(raw);
  for (const p of doc.paragraphs) {
    for (const el of p.elements) {
      if (el.kind === "text" && el.startIndex <= index && index < el.endIndex) return el.fontSizePt;
    }
  }
  return null;
}

/** One histogram row: a background hex and how many characters carry it. */
export interface BgHistogramEntry {
  hex: string;
  count: number;
}

/**
 * Character-weighted background-color histogram (plan A8's wrong-default
 * detector): counts UTF-16 units per hex over text elements. Character
 * weighting (not element counting) because what matters is how much EVIDENCE
 * each color covers. Sorted by count descending, then hex ascending so the
 * report is deterministic.
 */
function backgroundHistogram(doc: GDoc, topN: number): BgHistogramEntry[] {
  const counts = new Map<string, number>();
  for (const p of doc.paragraphs) {
    for (const el of p.elements) {
      if (el.kind !== "text" || el.backgroundHex === null) continue;
      counts.set(el.backgroundHex, (counts.get(el.backgroundHex) ?? 0) + (el.endIndex - el.startIndex));
    }
  }
  return [...counts.entries()]
    .map(([hex, count]): BgHistogramEntry => ({ hex, count }))
    .sort((a, b) => b.count - a.count || (a.hex < b.hex ? -1 : 1))
    .slice(0, topN);
}

/**
 * UTF-8 byte length without Buffer/TextEncoder — neither exists on the Apps
 * Script V8 runtime, and the count must match what the wire actually carried.
 * Counts per code point: 1 (ASCII), 2 (<= U+07FF), 3 (BMP), 4 (astral —
 * detected via the surrogate-pair high half; the low half is skipped).
 */
function utf8ByteLength(s: string): number {
  let bytes = 0;
  for (let i = 0; i < s.length; i++) {
    const code = s.charCodeAt(i);
    if (code <= 0x7f) bytes += 1;
    else if (code <= 0x7ff) bytes += 2;
    else if (code >= 0xd800 && code <= 0xdbff) {
      bytes += 4;
      i++; // the trailing low surrogate is part of this code point
    } else bytes += 3;
  }
  return bytes;
}

/** The adapter's raw observations, handed over for interpretation. */
export interface DiagnosticsInput {
  /** The verb-style masked read the probes ran against. */
  raw: unknown;
  fetchLatencyMs: number;
  /** One entry per batchUpdate the probe sequence issued, in order. */
  applyLatenciesMs: readonly number[];
  fontFloor: readonly FontFloorReading[];
  namedStyleProbe: NamedStyleProbeVerdict;
  inheritClear: InheritClearVerdict;
  /** False = the probe char could not be deleted (the report says so and how
   * to fix it by hand). */
  cleanupOk: boolean;
}

/** The interpreted report model renderDiagnosticsText renders. */
export interface DiagnosticsModel {
  version: string;
  tabCount: number;
  /** ALL rstm-family ranges, spacing and unknown versions included — this is
   * a state inventory, not the sidebar's passage count (plan A16). */
  rstmRangeCount: number;
  payloadBytes: number;
  fetchLatencyMs: number;
  applyLatenciesMs: readonly number[];
  fontFloor: FontFloorResult[];
  namedStyleProbe: NamedStyleProbeVerdict;
  inheritClear: InheritClearVerdict;
  cleanupOk: boolean;
  bgHistogram: BgHistogramEntry[];
  /** Whole-doc JSON for small docs (plan A11.ii), else null. */
  smallDocJson: string | null;
}

/** Build the report model from raw observations — every derived number is
 * computed here so the renderer below is pure string assembly. */
export function buildDiagnostics(input: DiagnosticsInput): DiagnosticsModel {
  const doc = parseDocument(input.raw);
  const json = JSON.stringify(input.raw) ?? "null"; // stringify(undefined) === undefined
  const payloadBytes = utf8ByteLength(json);
  let rstmRangeCount = 0;
  for (const nr of doc.namedRanges) {
    if (isRstmName(nr.name)) rstmRangeCount++;
  }
  return {
    version: GDOCS_VERSION,
    tabCount: doc.tabCount,
    rstmRangeCount,
    payloadBytes,
    fetchLatencyMs: input.fetchLatencyMs,
    applyLatenciesMs: input.applyLatenciesMs,
    fontFloor: input.fontFloor.map(interpretFontFloor),
    namedStyleProbe: input.namedStyleProbe,
    inheritClear: input.inheritClear,
    cleanupOk: input.cleanupOk,
    bgHistogram: backgroundHistogram(doc, BG_HISTOGRAM_TOP_N),
    smallDocJson: payloadBytes <= SMALL_DOC_DUMP_MAX_BYTES ? json : null
  };
}

/** Report wording for one font-floor row. */
function fontFloorLine(r: FontFloorResult): string {
  if (r.verdict === "rejected") return `  ${r.triedPt}pt -> rejected by the API`;
  if (r.verdict === "accepted exact") return `  ${r.triedPt}pt -> accepted exact`;
  return `  ${r.triedPt}pt -> clamped (reads back as ${r.readBackPt === null ? "no stated size" : `${r.readBackPt}pt`})`;
}

/**
 * The copy-pasteable plain-text report (plan D12: the wet round pastes this
 * back into chat). Deliberately technical — it is the diagnostic instrument,
 * not product copy (STRINGS carries the dialog chrome around it) — but it
 * still keeps to plain language a debater can skim.
 */
export function renderDiagnosticsText(m: DiagnosticsModel): string {
  const lines: string[] = [
    `Rostrum for Google Docs diagnostics (v${m.version})`,
    `tabs: ${m.tabCount}`,
    `hidden-text records: ${m.rstmRangeCount}`,
    `read payload: ${m.payloadBytes} bytes (fetched in ${m.fetchLatencyMs} ms)`,
    `write latencies (ms): ${m.applyLatenciesMs.length === 0 ? "none" : m.applyLatenciesMs.join(", ")}`,
    "font floor probe:"
  ];
  if (m.fontFloor.length === 0) lines.push("  not attempted");
  else for (const r of m.fontFloor) lines.push(fontFloorLine(r));
  lines.push(
    `heading 6 redefine probe: ${
      m.namedStyleProbe === "skipped" ? "skipped (the read stated no heading 6 size)" : m.namedStyleProbe
    }`,
    `clear-to-inherit probe: ${m.inheritClear}`,
    m.cleanupOk
      ? "probe cleanup: ok"
      : "probe cleanup: failed - a stray '.' probe character may remain at the very end of the doc; delete it",
    `background colors (top ${BG_HISTOGRAM_TOP_N} by characters):`
  );
  if (m.bgHistogram.length === 0) lines.push("  none found");
  else for (const row of m.bgHistogram) lines.push(`  ${row.hex} - ${row.count} chars`);
  lines.push("document json (embedded for small docs only):");
  lines.push(
    m.smallDocJson === null ? `  omitted - ${m.payloadBytes} bytes exceeds the ${SMALL_DOC_DUMP_MAX_BYTES}-byte cap` : m.smallDocJson
  );
  return lines.join("\n");
}
