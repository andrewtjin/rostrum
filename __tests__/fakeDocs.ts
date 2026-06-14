// Shared test harness: an in-memory DocsPort whose backing model is mutated
// ONLY by applyBatch — the fakeWord.ts sync-semantics parity (reads serialize
// the model; the controller's writes are the one mutation path; explicit test
// hooks model the USER's out-of-band edits). Not a `*.test.ts`, so jest never
// collects it as a suite; it is plain shared code for the flagship
// gdocsInvisibility suite.
//
// FIDELITY RULES (the reason this fake earns the flagship's trust):
//   * fetchDocument() returns REAL `documents.get` wire shapes and the suite
//     decodes them with the PRODUCTION parser (parse.ts) — the fixture-realism
//     rail (plan A11.ii). That means: body content starts at index 1 behind a
//     masked-out sectionBreak stub; every paragraph's trailing newline lives
//     INSIDE its final textRun's content; rgb channels at 0.0 are OMITTED
//     (proto3 wire reality); false booleans and absent styles are omitted too.
//   * applyBatch validates writeControl.requiredRevisionId first (mismatch ->
//     RevisionMismatchError, nothing applied) and bumps the revision once per
//     successful batch — exactly the chaining contract the controller relies
//     on (plan A13/D5).
//   * Batches are ATOMIC: every request is validated before anything mutates,
//     so a rejected batch (unknown request type via rejectRequestType, bad
//     range, unknown range id) leaves the model byte-identical — the property
//     the "nothing was applied" refusal copy depends on.
//   * updateTextStyle SPLITS text runs at range boundaries the way the real
//     service does, so round trips exercise the engine against fragmented
//     run layouts it did not itself create.
//   * Styling the segment-final newline throws (the API rule the
//     isLastInSegment clamp exists for, plan D6) — a clamp regression fails
//     LOUDLY here instead of silently passing a fake that is laxer than
//     Google.

import { decodeOptionalColor } from "../google-docs/src/core/color";
import {
  DocsApiError,
  DocsPort,
  DocsRequest,
  GNamedStyleType,
  RevisionMismatchError,
  UpdateNamedStyleRequest
} from "../google-docs/src/core/types";
import { GeSpec, GpSpec } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// Backing model
// ---------------------------------------------------------------------------

/** One run of the backing model. Text runs carry their content (including a
 * paragraph's trailing newline, in the last text run); "other" runs (chips,
 * objects) carry only an index width — their bodies are irrelevant to the
 * engine by the A9 whitelist. */
interface ModelRun {
  kind: "text" | "other";
  text: string;
  /** Index width: === text.length for text runs; arbitrary for "other". */
  width: number;
  fontSizePt: number | null;
  bold: boolean;
  /** Lower-case "#rrggbb" | null — same contract as GElement. */
  backgroundHex: string | null;
  /** Lower-case "#rrggbb" | null FOREGROUND (text) color — the exact analog of
   * backgroundHex (Loop 003). Modeled so a fixture can carry analytic-ify'd navy
   * runs through the production parser, which is what makes a PLANNER-driven
   * Delete-analytics round trip non-vacuous (the planner only deletes runs the
   * parsed view reports as analytics). */
  foregroundHex: string | null;
}

interface ModelParagraph {
  namedStyleType: GNamedStyleType;
  spaceAbovePt: number | null;
  spaceBelowPt: number | null;
  /** Bookkeeping: true once a border write landed (the retro pocket box).
   * Borders are not serialized — the fields mask never selects them — so the
   * flagship asserts them through this flag instead of the parse view. */
  boxed: boolean;
  /** True when this paragraph lives inside a table cell (a buildDoc spec with
   * inTable:true). serialize() wraps a run of consecutive inTable paragraphs in
   * a single-cell `table` structural element so the PRODUCTION parser
   * (collectParagraphs) re-derives inTable from the wire shape — proving the
   * keeper/planner's "table paragraphs are kept, never style-targeted" rule
   * against a real round trip rather than a hand-set flag. The index walk stays
   * FLAT (the table wrapper allocates no extra index width in this minimal
   * model): docEnd/spliceDelete/splitAndMutate all sum run widths, so a cell's
   * paragraphs occupy contiguous body indexes and a deleteContentRange splices
   * within a cell exactly as it does in the body. Faithful enough for the one
   * thing the harness must support — splice + parse-marks-inTable — without
   * modeling the real API's per-row/per-cell structural index slots, which no
   * verb in the engine ever targets. */
  inTable: boolean;
  runs: ModelRun[];
}

interface ModelRange {
  id: string;
  name: string;
  segments: { start: number; end: number }[];
}

export interface FakeDocsOptions {
  /** Leaf tabs to report. >1 makes every verb refuse (plan A3). */
  tabCount?: number;
  /** Serialize a live `suggested*` key (trips the Hide/Apply-styles gate). */
  suggestions?: boolean;
  /** What readSettingsJson returns (doc-properties tier; null = built-ins). */
  settingsJson?: string | null;
}

/** Copy one slice of a run, preserving every style fact (the split primitive). */
function sliceRun(run: ModelRun, from: number, to: number): ModelRun {
  return { ...run, text: run.text.slice(from, to), width: to - from };
}

/** "#rrggbb" -> proto3-shaped rgbColor with ZERO channels OMITTED — the wire
 * reality the fixture lint pins (plan A11.ii); parse.ts re-materializes them. */
function rgbColorOf(hex: string): Record<string, number> {
  const out: Record<string, number> = {};
  const channel = (offset: number): number => parseInt(hex.slice(offset, offset + 2), 16) / 255;
  const r = channel(1);
  const g = channel(3);
  const b = channel(5);
  if (r > 0) out.red = r;
  if (g > 0) out.green = g;
  if (b > 0) out.blue = b;
  return out;
}

/** Direct spacing dimension: an explicit ZERO omits its magnitude (the
 * omitted-means-zero rule parse.ts decodes by presence). */
function spacingDim(pt: number): Record<string, unknown> {
  return pt > 0 ? { magnitude: pt, unit: "PT" } : { unit: "PT" };
}

// ---------------------------------------------------------------------------
// The fake port
// ---------------------------------------------------------------------------

export class FakeDocs implements DocsPort {
  private readonly paragraphs: ModelParagraph[];
  private namedRanges: ModelRange[] = [];
  /** Per named style: its stated font size. updateNamedStyle writes land here
   * so a successful named-style batch is visible to later fetches. */
  private readonly namedStyleSizes: Partial<Record<GNamedStyleType, number>> = { NORMAL_TEXT: 11 };

  private rev = 1;
  private idCounter = 0;
  private suggestionsOn: boolean;
  private rejected: string | null = null;
  private readonly tabCount: number;
  private settingsJson: string | null;

  // ---- observability (read by the suite, never by the engine) -------------
  /** documents.get calls — the A13 one-fetch-per-verb / retry-count probe. */
  fetchCount = 0;
  /** Batches that APPLIED (validated + mutated + revision bumped). */
  readonly appliedBatches: DocsRequest[][] = [];
  /** Every applyBatch invocation, successful or not (0-based ordinals). */
  applyAttempts = 0;
  /** Named-style writes that landed (degraded-path assertions). */
  readonly namedStyleWrites: UpdateNamedStyleRequest[] = [];
  /**
   * Invoked at the TOP of every applyBatch, before revision validation, with
   * the 0-based call ordinal — the seam tests use to model a teammate editing
   * between chunks (pair it with injectForeignEdit).
   */
  beforeApply: ((call: number) => void) | null = null;

  /**
   * Build the backing model from the same paragraph specs the view builders
   * use (gdocsBuilders GpSpec — one spec vocabulary across the suites, DRY).
   * The newline rule mirrors buildDoc exactly: the trailing "\n" is appended
   * to the last TEXT element, or added as its own 1-char text run when the
   * paragraph ends in an "other" element / is empty (the API never puts the
   * newline inside a non-text element).
   */
  constructor(paras: GpSpec[], opts: FakeDocsOptions = {}) {
    this.tabCount = opts.tabCount ?? 1;
    this.suggestionsOn = opts.suggestions ?? false;
    this.settingsJson = opts.settingsJson ?? null;
    this.paragraphs = paras.map((p) => {
      const specs: GeSpec[] = [...p.elements];
      const last = specs[specs.length - 1];
      if (last !== undefined && (last.kind ?? "text") === "text") {
        specs[specs.length - 1] = { ...last, text: last.text + "\n" };
      } else {
        specs.push({ text: "\n" });
      }
      return {
        namedStyleType: p.style ?? "NORMAL_TEXT",
        spaceAbovePt: p.spaceAbovePt ?? null,
        spaceBelowPt: p.spaceBelowPt ?? null,
        boxed: false,
        // Carry inTable through to serialize(), which folds a contiguous run of
        // cell paragraphs into one single-cell table element (see ModelParagraph).
        inTable: p.inTable ?? false,
        runs: specs.map(
          (e): ModelRun => ({
            kind: e.kind ?? "text",
            text: (e.kind ?? "text") === "text" ? e.text : "",
            width: e.text.length,
            fontSizePt: e.size ?? null,
            bold: e.bold ?? false,
            backgroundHex: e.bg ?? null,
            foregroundHex: e.fg ?? null
          })
        )
      };
    });
  }

  // -------------------------------------------------------------------------
  // DocsPort
  // -------------------------------------------------------------------------

  /** Serialize the model to documents.get JSON (see fidelity rules above).
   * The JSON round trip both deep-copies (callers can never reach the model
   * through the response) and drops `undefined` keys exactly like real wire
   * omission would. */
  async fetchDocument(): Promise<unknown> {
    this.fetchCount++;
    return JSON.parse(JSON.stringify(this.serialize()));
  }

  async applyBatch(requests: DocsRequest[], requiredRevisionId: string): Promise<{ revisionId: string }> {
    const call = this.applyAttempts++;
    if (this.beforeApply !== null) this.beforeApply(call);
    // writeControl check FIRST — a stale revision must refuse before any
    // request-level validation, exactly as the guarded service does.
    if (requiredRevisionId !== this.revisionString()) throw new RevisionMismatchError();
    const rejected = this.rejected;
    if (rejected !== null && requests.some((r) => rejected in r)) {
      throw new DocsApiError(`request type rejected by fake: ${rejected}`);
    }
    // Atomicity: validate EVERYTHING, then mutate — a mid-batch failure may
    // never leave the model half-written (the real batch is transactional).
    this.validate(requests);
    for (const req of requests) this.applyOne(req);
    this.rev++;
    this.appliedBatches.push(requests);
    return { revisionId: this.revisionString() };
  }

  async readSettingsJson(): Promise<string | null> {
    return this.settingsJson;
  }

  async writeSettingsJson(json: string): Promise<void> {
    this.settingsJson = json;
  }

  // -------------------------------------------------------------------------
  // Test hooks — the sanctioned out-of-band mutations, each modeling one real
  // user/world action. Edits that a real user makes bump the revision, so the
  // controller's revision chain sees them exactly as live races.
  // -------------------------------------------------------------------------

  /** A foreign (teammate) edit: bumps the revision, optionally mutating the
   * model through the edit* primitives below (which deliberately do NOT bump
   * on their own — one user action, one revision bump, however many spans it
   * touched). */
  injectForeignEdit(edit?: (fake: FakeDocs) => void): void {
    this.rev++;
    if (edit !== undefined) edit(this);
  }

  /** User edit primitive: set a background over [start, end) — the A1 case
   * (highlighting inside a hidden region). Meaningful inside injectForeignEdit. */
  editSetBackground(start: number, end: number, hex: string): void {
    this.splitAndMutate(start, end, (run) => {
      run.backgroundHex = hex;
    });
  }

  /** User edit primitive: retype/paste-over at an explicit size — the A6 case
   * (14pt pasted into a hidden region; same-length replace, so indexes hold). */
  editSetFontSize(start: number, end: number, sizePt: number | null): void {
    this.splitAndMutate(start, end, (run) => {
      run.fontSizePt = sizePt;
    });
  }

  /** Cut/paste destroying a NamedRange (WAI per Google): the record vanishes,
   * the shrunken text stays — the convergence-sweep scenario (case 001-F5). */
  destroyRange(id: string): void {
    const before = this.namedRanges.length;
    this.namedRanges = this.namedRanges.filter((nr) => nr.id !== id);
    if (this.namedRanges.length === before) throw new Error(`destroyRange: unknown id ${id}`);
    this.rev++;
  }

  /** An edit splitting a range's segment at `at` into two segments — the
   * decayed-manifest shape Show All must still converge from (edge row 9). */
  splitRange(id: string, at: number): void {
    const nr = this.namedRanges.find((r) => r.id === id);
    if (nr === undefined) throw new Error(`splitRange: unknown id ${id}`);
    const seg = nr.segments.find((s) => s.start < at && at < s.end);
    if (seg === undefined) throw new Error(`splitRange: ${at} is not strictly inside a segment`);
    nr.segments = nr.segments.flatMap((s) =>
      s === seg
        ? [
            { start: s.start, end: at },
            { start: at, end: s.end }
          ]
        : [s]
    );
    this.rev++;
  }

  /** Make the NEXT batches containing this request type 400 (DocsApiError) —
   * the named-style consumer-account risk (plan D13) and generic-400 paths.
   * Pass null to clear. */
  rejectRequestType(type: string | null): void {
    this.rejected = type;
  }

  /** Toggle a live suggestion in the serialized read (key presence is the
   * signal — plan A16). No revision bump: the gate reads presence, not age. */
  setSuggestions(on: boolean): void {
    this.suggestionsOn = on;
  }

  // -------------------------------------------------------------------------
  // Inspection (facts the fields-masked parse view cannot carry)
  // -------------------------------------------------------------------------

  /** True once the paragraph received a border write (the pocket box). */
  isBoxed(paragraphIndex: number): boolean {
    return this.paragraphs[paragraphIndex].boxed;
  }

  /** The named style's stated size after any updateNamedStyle writes. */
  namedStyleSize(type: GNamedStyleType): number | undefined {
    return this.namedStyleSizes[type];
  }

  // -------------------------------------------------------------------------
  // Private mechanics
  // -------------------------------------------------------------------------

  private revisionString(): string {
    return `r${this.rev}`;
  }

  /** Index one past the last paragraph's end (body starts at 1). */
  private docEnd(): number {
    let cursor = 1;
    for (const p of this.paragraphs) for (const run of p.runs) cursor += run.width;
    return cursor;
  }

  /**
   * The split-and-mutate primitive shared by updateTextStyle and the user
   * edit hooks (one splitter, so engine writes and user edits fragment runs
   * identically). Text runs overlapping [start, end) split at the boundaries;
   * the covered middle piece is mutated. "Other" runs are skipped: the engine
   * never targets them (A9) and the user edits modeled here are text edits.
   */
  private splitAndMutate(start: number, end: number, mutate: (run: ModelRun) => void): void {
    let cursor = 1;
    for (const p of this.paragraphs) {
      const next: ModelRun[] = [];
      for (const run of p.runs) {
        const rs = cursor;
        const re = cursor + run.width;
        cursor = re;
        if (run.kind !== "text" || re <= start || rs >= end) {
          next.push(run);
          continue;
        }
        const a = Math.max(rs, start);
        const b = Math.min(re, end);
        if (a > rs) next.push(sliceRun(run, 0, a - rs));
        const mid = sliceRun(run, a - rs, b - rs);
        mutate(mid);
        next.push(mid);
        if (b < re) next.push(sliceRun(run, b - rs, re - rs));
      }
      p.runs = next;
    }
  }

  /**
   * Splice the half-open [start, end) index range OUT of the model — the
   * deleteContentRange primitive. Walks the same derived cursor splitAndMutate
   * does, but instead of mutating the covered slice it DROPS it: a run keeps
   * only the kept-before and kept-after pieces, so its width (and thus every
   * downstream index) shrinks by the overlap. Both kinds are spliced — sliceRun
   * shrinks an "other" run's width while leaving its empty text alone — because
   * a delete by index removes whatever occupies that index; the analytics
   * planner only ever targets text-run ranges, so the "other" path is purely
   * defensive index-fidelity. Runs trimmed to zero width are dropped, and a
   * paragraph left with NO runs (its trailing newline was deleted too) is
   * removed entirely — the line collapse the whole-analytics delete relies on.
   * Indexes are NOT stored, so nothing else needs renumbering: the next
   * docEnd()/serialize() re-derives them from the shrunken run widths.
   */
  private spliceDelete(start: number, end: number): void {
    let cursor = 1;
    for (const p of this.paragraphs) {
      const next: ModelRun[] = [];
      for (const run of p.runs) {
        const rs = cursor;
        const re = cursor + run.width;
        cursor = re;
        // Untouched runs (entirely before/after the cut) pass through whole.
        if (re <= start || rs >= end) {
          next.push(run);
          continue;
        }
        const a = Math.max(rs, start); // first deleted index within this run
        const b = Math.min(re, end); // one past the last deleted index
        // Keep the surviving prefix [rs, a) and suffix [b, re); drop [a, b).
        if (a > rs) next.push(sliceRun(run, 0, a - rs));
        if (b < re) next.push(sliceRun(run, b - rs, re - rs));
      }
      p.runs = next;
    }
    // A paragraph whose every run (including the trailing newline) was deleted
    // collapses — Docs merges the now-newline-less line into the next. Drop it
    // in place so the readonly array reference holds (mutation, not reassign).
    for (let i = this.paragraphs.length - 1; i >= 0; i--) {
      if (this.paragraphs[i].runs.length === 0) this.paragraphs.splice(i, 1);
    }
  }

  /** Pre-mutation validation pass (see applyBatch atomicity note). */
  private validate(requests: DocsRequest[]): void {
    const end = this.docEnd();
    // The segment-final newline (the last body index) is unstylable — the
    // real-API rule the planner's isLastInSegment clamp exists for (plan D6).
    const styleCeiling = end - 1;
    const ids = new Set(this.namedRanges.map((nr) => nr.id));
    for (const req of requests) {
      if ("updateTextStyle" in req) {
        const r = req.updateTextStyle.range;
        if (!(r.startIndex >= 1 && r.endIndex > r.startIndex && r.endIndex <= end)) {
          throw new DocsApiError(`updateTextStyle: invalid range [${r.startIndex}, ${r.endIndex})`);
        }
        if (r.endIndex > styleCeiling) {
          throw new DocsApiError("updateTextStyle: cannot style the segment-final newline");
        }
      } else if ("createNamedRange" in req) {
        const { name, range } = req.createNamedRange;
        if (name.length < 1 || name.length > 256) {
          throw new DocsApiError("createNamedRange: name must be 1-256 chars");
        }
        if (!(range.startIndex >= 1 && range.endIndex > range.startIndex && range.endIndex <= end)) {
          throw new DocsApiError(`createNamedRange: invalid range [${range.startIndex}, ${range.endIndex})`);
        }
      } else if ("deleteNamedRange" in req) {
        if (!ids.has(req.deleteNamedRange.namedRangeId)) {
          throw new DocsApiError(`deleteNamedRange: unknown id ${req.deleteNamedRange.namedRangeId}`);
        }
      } else if ("updateParagraphStyle" in req) {
        const r = req.updateParagraphStyle.range;
        if (!(r.startIndex >= 1 && r.endIndex > r.startIndex && r.endIndex <= end)) {
          throw new DocsApiError(`updateParagraphStyle: invalid range [${r.startIndex}, ${r.endIndex})`);
        }
      } else if ("deleteContentRange" in req) {
        // Delete-analytics is the sole content-deleter (plan 003-F1). The same
        // range rules as the style verbs (in-range, non-empty, start < end)
        // PLUS the segment-final newline is UNREMOVABLE — Docs rejects deleting
        // the last body index, which is exactly why the planner clamps a
        // final-segment whole-analytics delete to [start, end-1) and leaves an
        // empty line. The ceiling is reused verbatim from the style rule so a
        // clamp regression in the planner fails LOUDLY here, never silently
        // against a fake laxer than Google.
        const r = req.deleteContentRange.range;
        if (!(r.startIndex >= 1 && r.endIndex > r.startIndex && r.endIndex <= end)) {
          throw new DocsApiError(`deleteContentRange: invalid range [${r.startIndex}, ${r.endIndex})`);
        }
        if (r.endIndex > styleCeiling) {
          throw new DocsApiError("deleteContentRange: cannot delete the segment-final newline");
        }
      }
      // updateNamedStyle: shape is compile-time guaranteed; nothing to check.
    }
  }

  /** Apply one validated request to the model. */
  private applyOne(req: DocsRequest): void {
    if ("updateTextStyle" in req) {
      const { range, textStyle, fields } = req.updateTextStyle;
      const mask = new Set(fields.split(","));
      this.splitAndMutate(range.startIndex, range.endIndex, (run) => {
        // A field named in the mask but absent from the style CLEARS — the
        // documented clear-to-inherit semantics the restore path leans on.
        if (mask.has("fontSize")) run.fontSizePt = textStyle.fontSize?.magnitude ?? null;
        if (mask.has("bold")) run.bold = textStyle.bold === true;
        // Foreground honors the SAME field-mask contract, decoded by the
        // PRODUCTION codec (color.decodeOptionalColor) — the exact decoder
        // parse.ts uses on the read side — so analytic-ify's navy write lands as
        // the byte it round-trips to, and is then visible to a later
        // fetch+parse (otherwise the write was silently dropped and the navy
        // could never be verified). Mask-present-but-style-absent CLEARS to
        // null (inherit), mirroring fontSize/bold above.
        if (mask.has("foregroundColor")) run.foregroundHex = decodeOptionalColor(textStyle.foregroundColor);
      });
    } else if ("createNamedRange" in req) {
      const { name, range } = req.createNamedRange;
      this.namedRanges.push({
        id: `fnr-${++this.idCounter}`,
        name,
        segments: [{ start: range.startIndex, end: range.endIndex }]
      });
    } else if ("deleteNamedRange" in req) {
      this.namedRanges = this.namedRanges.filter((nr) => nr.id !== req.deleteNamedRange.namedRangeId);
    } else if ("deleteContentRange" in req) {
      // Delete-analytics's one mutation. Splice the chars out of the live
      // model; because every index in this fake is DERIVED from run widths
      // (docEnd/serialize re-walk the runs), removing chars automatically
      // shifts every downstream run + paragraph index. applyBatch calls
      // applyOne sequentially, so a LATER delete in the same batch already
      // sees the post-delete model — which is precisely why the planner must
      // emit ranges in DESCENDING start order (a real batchUpdate is
      // sequential too). The differential test pins this: the same ranges
      // applied ascending corrupt the doc here.
      this.spliceDelete(req.deleteContentRange.range.startIndex, req.deleteContentRange.range.endIndex);
    } else if ("updateParagraphStyle" in req) {
      const { range, paragraphStyle, fields } = req.updateParagraphStyle;
      // Real semantics: the style applies to every paragraph the range
      // OVERLAPS, not only those it contains.
      let cursor = 1;
      for (const p of this.paragraphs) {
        const pStart = cursor;
        for (const run of p.runs) cursor += run.width;
        const pEnd = cursor;
        if (pStart >= range.endIndex || pEnd <= range.startIndex) continue;
        for (const field of fields.split(",")) {
          if (field === "spaceAbove") {
            p.spaceAbovePt = paragraphStyle.spaceAbove !== undefined ? paragraphStyle.spaceAbove.magnitude : null;
          } else if (field === "spaceBelow") {
            p.spaceBelowPt = paragraphStyle.spaceBelow !== undefined ? paragraphStyle.spaceBelow.magnitude : null;
          } else if (field.startsWith("border")) {
            p.boxed = true;
          }
        }
      }
    } else if ("updateNamedStyle" in req) {
      const { namedStyle, fields } = req.updateNamedStyle;
      if (fields.split(",").includes("textStyle.fontSize") && namedStyle.textStyle?.fontSize !== undefined) {
        this.namedStyleSizes[namedStyle.namedStyleType] = namedStyle.textStyle.fontSize.magnitude;
      }
      this.namedStyleWrites.push(req);
    }
  }

  /** Build the wire-shaped documents.get response (fidelity rules above). */
  private serialize(): unknown {
    const content: unknown[] = [{ endIndex: 1 }]; // masked sectionBreak stub
    let cursor = 1;
    // Plant exactly one suggested* key on the first text run when toggled on
    // (any key anywhere trips the generic walk — plan A16).
    let suggestionPlanted = false;
    // Serialize ONE model paragraph to its `{ startIndex, endIndex, paragraph }`
    // structural element, advancing the shared cursor. Hoisted to a closure so
    // the result can be emitted either directly into body content OR into a
    // table cell's content (table grouping below) from a SINGLE serializer —
    // the index walk is identical in both, which is what keeps a cell paragraph's
    // indexes contiguous with the body (the flat-index simplification).
    const serializeParagraph = (p: ModelParagraph): Record<string, unknown> => {
      const pStart = cursor;
      const elements = p.runs.map((run) => {
        const el: Record<string, unknown> = { startIndex: cursor, endIndex: cursor + run.width };
        if (run.kind === "text") {
          const textStyle: Record<string, unknown> = {};
          if (run.fontSizePt !== null) textStyle.fontSize = { magnitude: run.fontSizePt, unit: "PT" };
          if (run.bold) textStyle.bold = true; // false is omitted on the wire
          if (run.backgroundHex !== null) {
            textStyle.backgroundColor = { color: { rgbColor: rgbColorOf(run.backgroundHex) } };
          }
          // Foreground rides the same proto3-shaped OptionalColor as background
          // (zero channels omitted; parse.ts re-materializes them) so an
          // analytic-ify'd navy run reads back as analytics through the parser.
          if (run.foregroundHex !== null) {
            textStyle.foregroundColor = { color: { rgbColor: rgbColorOf(run.foregroundHex) } };
          }
          const textRun: Record<string, unknown> = { content: run.text, textStyle };
          if (this.suggestionsOn && !suggestionPlanted) {
            textRun.suggestedInsertionIds = ["kix.suggestion1"];
            suggestionPlanted = true;
          }
          el.textRun = textRun;
        } else {
          // Any non-textRun element type reads as kind "other" (A9 whitelist
          // keys off textRun ABSENCE); a person chip is the representative.
          el.person = { personId: "kix.person1", personProperties: { name: "Someone" } };
        }
        cursor += run.width;
        return el;
      });
      const paragraphStyle: Record<string, unknown> = { namedStyleType: p.namedStyleType };
      if (p.spaceAbovePt !== null) paragraphStyle.spaceAbove = spacingDim(p.spaceAbovePt);
      if (p.spaceBelowPt !== null) paragraphStyle.spaceBelow = spacingDim(p.spaceBelowPt);
      return { startIndex: pStart, endIndex: cursor, paragraph: { elements, paragraphStyle } };
    };

    // Walk the model, folding each maximal run of consecutive inTable paragraphs
    // into ONE single-cell `table` structural element (one row, one cell whose
    // `content` holds those paragraphs). parse.collectParagraphs recurses into
    // table.tableRows[].tableCells[].content with inTable=true, so the cell's
    // paragraphs read back marked inTable — the round trip the keeper/planner
    // table rule is verified against. Body paragraphs serialize straight into
    // content. The cursor is shared, so cell paragraphs occupy contiguous body
    // indexes (no extra structural slots — the minimal-table simplification).
    for (let i = 0; i < this.paragraphs.length; ) {
      if (!this.paragraphs[i].inTable) {
        content.push(serializeParagraph(this.paragraphs[i]));
        i++;
        continue;
      }
      const cellContent: unknown[] = [];
      while (i < this.paragraphs.length && this.paragraphs[i].inTable) {
        cellContent.push(serializeParagraph(this.paragraphs[i]));
        i++;
      }
      content.push({ table: { tableRows: [{ tableCells: [{ content: cellContent }] }] } });
    }

    // namedRanges: the doubly-nested name-keyed map — same-name siblings (two
    // regions with identical RLE) group under ONE key, as the real API does.
    const namedRangesMap: Record<string, { name: string; namedRanges: unknown[] }> = {};
    for (const nr of this.namedRanges) {
      const group = (namedRangesMap[nr.name] ??= { name: nr.name, namedRanges: [] });
      group.namedRanges.push({
        namedRangeId: nr.id,
        name: nr.name,
        ranges: nr.segments.map((s) => ({ startIndex: s.start, endIndex: s.end }))
      });
    }

    const namedStyles = {
      styles: Object.entries(this.namedStyleSizes).map(([type, sizePt]) => ({
        namedStyleType: type,
        textStyle: { fontSize: { magnitude: sizePt, unit: "PT" } }
      }))
    };

    const documentTab = { body: { content }, namedRanges: namedRangesMap, namedStyles };
    // includeTabsContent:true reality: every doc has >= 1 tab and the content
    // rides in tabs[0].documentTab; extra LEAF tab stubs make the doc
    // multi-tab (parse counts leaves; the controller's tab gate refuses).
    const tabs: unknown[] = [{ tabProperties: { tabId: "t.0", index: 0 }, documentTab }];
    for (let i = 1; i < this.tabCount; i++) {
      tabs.push({ tabProperties: { tabId: `t.${i}`, index: i } });
    }
    return { revisionId: this.revisionString(), tabs };
  }
}
