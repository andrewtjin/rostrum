// Apply-debate-styles + Mark-cite request builders (plan A5/A10/A12, step
// S10) — the Docs sibling of the Word styles lane (src/core/styles.ts
// STYLE_MAP + officeStyles.ensureRostrumStyles + the controller's cite-repair
// phase), collapsed into ONE pure planner because Docs takes plain batchUpdate
// JSON where Word needed live host API calls.
//
// GATING CONTRACT (callers gate BEFORE planning): the controller runs
// guards.assertSingleTab + assertNoSuggestions + assertNotHidden and only then
// calls planApplyStyles — restyling an armed doc would write run-level sizes
// that collide with the rstm restore records (plan A5), and suggestion-laden
// indexes are untrustworthy (plan D5). This module assumes a single-tab,
// suggestion-free, UNARMED view and never re-checks; planMarkCite likewise
// assumes the adapter already refused an empty selection (markCiteNoop copy)
// and handed a body range that excludes the segment-final newline.
//
// TWO BATCHES BY DESIGN (plan A5/D13): updateNamedStyle is the one request
// shape whose acceptance is UNVERIFIED on consumer accounts (the documented
// 400 risk). The named-style writes and the retro writes are therefore
// SEPARATE RequestGroup arrays — the controller applies them as separate
// batchUpdates, so a 400 on batch 1 lands on the documented degraded path
// (named styles unchanged, future typing unstyled, Diagnostics reports it)
// while batch 2 still restyles every EXISTING paragraph. The retro pass is
// the MAINLINE for imports, not a fallback: .docx imports materialize direct
// sizes/spacing that named-style updates cannot pierce.
//
// EMISSION INVARIANTS (cross-module contract):
//   * Only updateNamedStyle / updateParagraphStyle / updateTextStyle are
//     emitted — never an insert/delete/reorder of content (case 001-F1).
//   * Run-level fontSize is written ONLY over H1-4 paragraph text and the
//     detectCiteLeads ranges (plan A5's invariant) — structurally guaranteed
//     here because those are the only two places updateTextStyle is built.
//   * Only kind:"text" spans are style-targeted (whitelist, plan A9): heading
//     restyles break around chips/breaks/objects instead of spanning them.
//   * Border objects are always COMPLETE — all four sides, all fields — in
//     BOTH the named-style attempt and the retro pass (the API rejects
//     partial border updates; plan D13).

import { CITE_PT, STYLE_SIZES_PT } from "./constants";
import { detectCiteLeads } from "./keepers";
import {
  DocsParagraphBorder,
  DocsParagraphStyle,
  DocsRequest,
  DocsTextStyle,
  GDoc,
  GdocsSettings,
  GNamedStyleType,
  GParagraph,
  GRange,
  RequestGroup,
  UpdateNamedStyleRequest,
  UpdateParagraphStyleRequest,
  UpdateTextStyleRequest
} from "./types";

// ---------------------------------------------------------------------------
// Public contract
// ---------------------------------------------------------------------------

/** The debate style names receipts count by (Word STYLE_MAP parity). */
export type DebateStyleName = "pocket" | "hat" | "block" | "tag";

/**
 * planApplyStyles output. `counts` feeds GStylesResult verbatim — the
 * controller adds `namedStylesApplied` after it learns whether batch 1 landed
 * (that fact is an APPLY outcome, unknowable at plan time).
 */
export interface StylesPlan {
  /** BATCH 1: the five updateNamedStyle writes (Normal + H1-4). */
  namedStyleGroups: RequestGroup[];
  /** BATCH 2: per-paragraph retro restyle + direct-spacing clear + cite repair. */
  retroGroups: RequestGroup[];
  counts: {
    /** Existing paragraphs restyled, per debate style (receipt unit). */
    restyled: Record<DebateStyleName, number>;
    /** Paragraphs whose direct spacing was cleared to 0. */
    spacingCleared: number;
    /** Cite LINES repaired (paragraphs, not lead elements — a split
     * author/year lead is still one cite in the receipt). */
    citesRepaired: number;
  };
}

// ---------------------------------------------------------------------------
// Style table (Word STYLE_MAP parity, decision #9 sizes via STYLE_SIZES_PT)
// ---------------------------------------------------------------------------

/** What one debate heading style writes: size always; bold only for pocket
 * and tag (hat/block keep the author's own emphasis — Word parity, where only
 * Heading1/Heading4 are bold in the template); the box only for pocket. */
interface DebateStyleSpec {
  count: DebateStyleName;
  sizePt: number;
  bold: boolean;
  boxed: boolean;
}

/**
 * H1-4 -> debate spec. TITLE/SUBTITLE/H5/H6 are DELIBERATELY absent: the
 * debate ladder is pocket/hat/block/tag on H1-4 (plan goal), and restyling
 * styles outside it would surprise users whose docs use them for prose.
 * Single-sourced from STYLE_SIZES_PT so the policy numbers cannot drift from
 * the keeper/cite constants (the Mark-cite/keeper split finding).
 */
const DEBATE_HEADINGS: Partial<Record<GNamedStyleType, DebateStyleSpec>> = {
  HEADING_1: { count: "pocket", sizePt: STYLE_SIZES_PT.pocket, bold: true, boxed: true },
  HEADING_2: { count: "hat", sizePt: STYLE_SIZES_PT.hat, bold: false, boxed: false },
  HEADING_3: { count: "block", sizePt: STYLE_SIZES_PT.block, bold: false, boxed: false },
  HEADING_4: { count: "tag", sizePt: STYLE_SIZES_PT.tag, bold: true, boxed: false }
};

/** The named-style batch is emitted in this fixed order, Normal FIRST: its
 * spacing zeroing is the condensation dependency (plan R1/A12). The order is
 * non-load-bearing (one atomic batch) but fixed for deterministic output. */
const NAMED_STYLE_ORDER: readonly GNamedStyleType[] = [
  "NORMAL_TEXT",
  "HEADING_1",
  "HEADING_2",
  "HEADING_3",
  "HEADING_4"
];

// ---------------------------------------------------------------------------
// Request shapes (private builders — fresh objects per call, so no two
// emitted requests ever alias one another's JSON)
// ---------------------------------------------------------------------------

/** The Docs dimension literal for a point value. */
function pt(magnitude: number): { magnitude: number; unit: "PT" } {
  return { magnitude, unit: "PT" };
}

/**
 * One side of the pocket box — Word parity with officeStyles' border spec:
 * 0.5pt black single/solid with 1pt padding. Black is the EMPTY rgbColor {}:
 * proto3 omits zero channels, the same wire shape the fixture-lint pins for
 * reads, so emissions match what the API itself would echo back.
 */
function pocketBorderSide(): DocsParagraphBorder {
  return {
    color: { color: { rgbColor: {} } },
    width: pt(0.5),
    padding: pt(1),
    dashStyle: "SOLID"
  };
}

/** The four border field names, in the fixed emission order tests pin. */
const BORDER_FIELD_NAMES = ["borderTop", "borderBottom", "borderLeft", "borderRight"] as const;

/** Field mask for a retro per-paragraph border write. */
const BORDER_FIELDS = BORDER_FIELD_NAMES.join(",");

/**
 * The COMPLETE four-side border set (plan D13: partial border updates are
 * rejected, so every border-carrying paragraphStyle always specifies all four
 * sides in entirety — both in the named-style attempt and the retro pass).
 */
function pocketBorders(): DocsParagraphStyle {
  return {
    borderTop: pocketBorderSide(),
    borderBottom: pocketBorderSide(),
    borderLeft: pocketBorderSide(),
    borderRight: pocketBorderSide()
  };
}

/**
 * One updateNamedStyle request. Field paths are relative to `namedStyle`
 * ("textStyle.fontSize", "paragraphStyle.borderTop", ...) — this request
 * shape is the engine's one UNVERIFIED emission (plan D13), so the mask
 * convention is a documented choice the wet round's Diagnostics confirms; a
 * rejection lands on the degraded path, never on data loss.
 */
function namedStyleRequest(style: GNamedStyleType, spec: DebateStyleSpec): UpdateNamedStyleRequest {
  const textStyle: DocsTextStyle = { fontSize: pt(spec.sizePt) };
  const fields: string[] = ["textStyle.fontSize"];
  if (spec.bold) {
    textStyle.bold = true;
    fields.push("textStyle.bold");
  }
  let paragraphStyle: DocsParagraphStyle | undefined;
  if (spec.boxed) {
    paragraphStyle = pocketBorders();
    for (const side of BORDER_FIELD_NAMES) fields.push(`paragraphStyle.${side}`);
  }
  return {
    updateNamedStyle: {
      namedStyle: {
        namedStyleType: style,
        textStyle,
        ...(paragraphStyle !== undefined ? { paragraphStyle } : {})
      },
      fields: fields.join(",")
    }
  };
}

/**
 * The NORMAL_TEXT named-style write: 11pt AND spaceAbove/Below 0 — the
 * condensation dependency (plan R1/A12; hidden 1pt paragraphs still carry
 * their style's spacing, so a non-zero Normal spacing caps the page ratio).
 * The Apply-styles confirm copy owns telling the user this is doc-wide and
 * not undone by Show All.
 */
function normalNamedStyleRequest(): UpdateNamedStyleRequest {
  return {
    updateNamedStyle: {
      namedStyle: {
        namedStyleType: "NORMAL_TEXT",
        textStyle: { fontSize: pt(STYLE_SIZES_PT.normal) },
        paragraphStyle: { spaceAbove: pt(0), spaceBelow: pt(0) }
      },
      fields: "textStyle.fontSize,paragraphStyle.spaceAbove,paragraphStyle.spaceBelow"
    }
  };
}

/** Retro border write for one existing H1 paragraph (complete set, D13). */
function borderRequest(range: GRange): UpdateParagraphStyleRequest {
  return {
    updateParagraphStyle: {
      range,
      paragraphStyle: pocketBorders(),
      fields: BORDER_FIELDS
    }
  };
}

/**
 * Retro run-level restyle for one coalesced text span of an H1-4 paragraph.
 * The mask carries bold only for the bold debate styles, so hat/block runs
 * keep whatever emphasis the author wrote (Word parity: sizes change,
 * emphasis survives).
 */
function headingTextRequest(range: GRange, spec: DebateStyleSpec): UpdateTextStyleRequest {
  const textStyle: DocsTextStyle = spec.bold
    ? { bold: true, fontSize: pt(spec.sizePt) }
    : { fontSize: pt(spec.sizePt) };
  return {
    updateTextStyle: {
      range,
      textStyle,
      fields: spec.bold ? "bold,fontSize" : "fontSize"
    }
  };
}

/**
 * Direct-spacing clear: write 0 to BOTH spacing fields. Writing 0 (not
 * clearing to inherit) is deliberate — plan A5's wording — and parse.ts
 * guarantees an explicit zero reads back as 0 (non-null), so a re-apply sees
 * the same "carries direct spacing" fact and emits the same write (idempotent
 * by repetition, see planApplyStyles).
 */
function spacingClearRequest(range: GRange): UpdateParagraphStyleRequest {
  return {
    updateParagraphStyle: {
      range,
      paragraphStyle: { spaceAbove: pt(0), spaceBelow: pt(0) },
      fields: "spaceAbove,spaceBelow"
    }
  };
}

/**
 * The cite convention (bold CITE_PT) over one range — the SINGLE emission
 * shape shared by Apply-styles' bulk repair and Mark-cite, so the two paths
 * can never drift (the Mark-cite/keeper split was a named plan-review
 * finding). The range is copied, never aliased, so callers' objects stay
 * theirs.
 */
function citeConventionRequest(range: GRange): UpdateTextStyleRequest {
  return {
    updateTextStyle: {
      range: { startIndex: range.startIndex, endIndex: range.endIndex },
      textStyle: { bold: true, fontSize: pt(CITE_PT) },
      fields: "bold,fontSize"
    }
  };
}

// ---------------------------------------------------------------------------
// Range geometry (private)
// ---------------------------------------------------------------------------

/**
 * A paragraph's stylable range: its full range including the trailing newline
 * (paragraph-mark size drives line height, plan D6, and updateParagraphStyle
 * needs only to OVERLAP the paragraph), clamped off the segment-final newline
 * the API refuses to touch (the cross-module isLastInSegment pin — applied to
 * paragraph-style ranges too, conservatively, rather than betting that
 * updateParagraphStyle tolerates what updateTextStyle rejects). Null when the
 * clamp empties the range: an EMPTY segment-final paragraph has nothing left
 * to target, and a zero-length range would 400.
 */
function clampedParagraphRange(p: GParagraph): GRange | null {
  const endIndex = p.isLastInSegment ? p.endIndex - 1 : p.endIndex;
  return endIndex > p.startIndex ? { startIndex: p.startIndex, endIndex } : null;
}

/**
 * The text-only spans of a heading paragraph, maximal-coalesced: adjacent
 * kind:"text" elements merge into one range; every kind:"other" element
 * breaks the run (whitelist, plan A9 — a chip inside a heading is never
 * style-targeted, exactly as hide spans break around it). The segment-final
 * newline is clamped off the tail range; a tail emptied by the clamp (an
 * empty segment-final heading whose only text is the newline) is dropped.
 */
function headingTextRanges(p: GParagraph): GRange[] {
  const ranges: GRange[] = [];
  for (const e of p.elements) {
    if (e.kind !== "text") continue;
    const last = ranges[ranges.length - 1];
    // Adjacency check (not blind merging): an "other" element between two
    // text elements occupies index space, so the gap breaks the run.
    if (last !== undefined && last.endIndex === e.startIndex) last.endIndex = e.endIndex;
    else ranges.push({ startIndex: e.startIndex, endIndex: e.endIndex });
  }
  if (p.isLastInSegment && ranges.length > 0) {
    const tail = ranges[ranges.length - 1];
    // The newline lives inside the final text element (parse/builder
    // guarantee), so only a tail that reaches the paragraph end needs the clamp.
    if (tail.endIndex === p.endIndex) tail.endIndex -= 1;
    if (tail.endIndex <= tail.startIndex) ranges.pop();
  }
  return ranges;
}

// ---------------------------------------------------------------------------
// Public: planApplyStyles
// ---------------------------------------------------------------------------

/** A retro group plus the doc position it sorts by (see emission order). */
interface AnchoredGroup {
  anchor: number;
  group: RequestGroup;
}

/**
 * Plan the full Apply-debate-styles pass over a gated view.
 *
 * Batch 1 (namedStyleGroups) — five updateNamedStyle writes redefining the
 * per-doc named styles, so Docs' NATIVE Ctrl+Alt+1..4 chords produce debate
 * styles from then on (the whole access-path premise: add-ons cannot register
 * shortcuts).
 *
 * Batch 2 (retroGroups) — three concerns over EXISTING paragraphs:
 *   * H1-4 restyle: run-level size (+bold for pocket/tag) over each heading's
 *     text spans, plus the complete pocket box per H1 paragraph;
 *   * direct-spacing clear: 0pt over every paragraph carrying ANY direct
 *     spaceAbove/Below — including an explicit direct 0, which still shadows
 *     the named style (parse.ts preserves the distinction for exactly this);
 *   * cite repair (plan A10(b)): the convention (bold CITE_PT) onto every
 *     detectCiteLeads range, so Docs-native bold-11pt cites meet the SIGNATURE
 *     keeper rule afterwards. Settings reach this planner ONLY through
 *     detectCiteLeads (its already-at-convention skip uses citeMinPt).
 *
 * Tables are untouched territory (row 1 parity): a heading or direct spacing
 * inside a table is skipped entirely.
 *
 * GROUPING: one RequestGroup per touched paragraph (borders + text + spacing
 * together) and one per cite line — the chunker never splits a group, so an
 * interrupted apply can never leave a paragraph half-styled (boxed but not
 * resized). Groups are ordered by DESCENDING start index — the same
 * non-load-bearing convention the hide planner uses (plan A16; nothing here
 * mutates indexes, case 001-F1).
 *
 * IDEMPOTENT BY REPETITION: every write is an absolute value (26pt, bold,
 * 0pt, a complete border) keyed off facts the writes themselves preserve
 * (namedStyleType; direct spacing stays direct as 0), so re-running on an
 * already-styled doc emits the SAME heading/spacing writes — each a server
 * no-op. The cite pass instead CONVERGES: a repaired lead now meets the
 * signature threshold, so detection skips it and the second run emits zero
 * cite writes. Safe either way because no emission moves content.
 */
export function planApplyStyles(doc: GDoc, settings: GdocsSettings): StylesPlan {
  // ---- Batch 1: the named-style redefinitions (fixed order, Normal first).
  const namedStyleGroups: RequestGroup[] = NAMED_STYLE_ORDER.map((style) => {
    const spec = DEBATE_HEADINGS[style];
    const request = spec === undefined ? normalNamedStyleRequest() : namedStyleRequest(style, spec);
    // One request per group: the five writes have no atomicity coupling
    // beyond the batch itself (batchUpdate is atomic), so the chunker is free
    // to pack them however its cap dictates.
    return { requests: [request] };
  });

  // ---- Batch 2: retro pass, per paragraph.
  const anchored: AnchoredGroup[] = [];
  const restyled: Record<DebateStyleName, number> = { pocket: 0, hat: 0, block: 0, tag: 0 };
  let spacingCleared = 0;

  for (const p of doc.paragraphs) {
    if (p.inTable) continue; // tables untouched (row 1 parity)
    const requests: DocsRequest[] = [];
    const range = clampedParagraphRange(p);
    if (range === null) continue; // empty segment-final paragraph: nothing targetable

    const spec = DEBATE_HEADINGS[p.namedStyleType];
    if (spec !== undefined) {
      // Paragraph-level write first, then run-level — cosmetic order, the
      // group is atomic either way.
      if (spec.boxed) requests.push(borderRequest(range));
      for (const textRange of headingTextRanges(p)) requests.push(headingTextRequest(textRange, spec));
      // Count only paragraphs that actually received a restyle write (a
      // degenerate heading that yielded none stays out of the receipt).
      if (requests.length > 0) restyled[spec.count]++;
    }

    // Direct spacing: PRESENCE of a direct value is the trigger (null =
    // inherits, untouched), zero included — see spacingClearRequest.
    if (p.spaceAbovePt !== null || p.spaceBelowPt !== null) {
      requests.push(spacingClearRequest(range));
      spacingCleared++;
    }

    if (requests.length > 0) anchored.push({ anchor: p.startIndex, group: { requests } });
  }

  // ---- Cite repair: one atomic group per cite LINE. detectCiteLeads emits
  // lead-element ranges in document order, so consecutive ranges inside the
  // same paragraph (a split author/year lead) fold into one group via a
  // single forward sweep — and the group count IS citesRepaired (the receipt
  // counts cites, not lead elements).
  let citesRepaired = 0;
  let sweep = 0; // paragraph pointer, advanced monotonically with the ranges
  let currentParagraph = -1;
  let citeRequests: DocsRequest[] = [];
  let citeAnchor = 0;
  const flushCite = (): void => {
    if (citeRequests.length === 0) return;
    anchored.push({ anchor: citeAnchor, group: { requests: citeRequests } });
    citesRepaired++;
    citeRequests = [];
  };
  for (const lead of detectCiteLeads(doc, settings)) {
    while (sweep < doc.paragraphs.length && doc.paragraphs[sweep].endIndex <= lead.startIndex) sweep++;
    if (sweep !== currentParagraph) {
      flushCite();
      currentParagraph = sweep;
      citeAnchor = lead.startIndex;
    }
    citeRequests.push(citeConventionRequest(lead));
  }
  flushCite();

  // Descending-start-index emission across groups — convention only (plan
  // A16): no emitted request type mutates indexes, so order cannot corrupt;
  // it is fixed purely so output is deterministic and diffs read top-down.
  // Array.prototype.sort is stable, so equal anchors keep build order.
  anchored.sort((a, b) => b.anchor - a.anchor);

  return {
    namedStyleGroups,
    retroGroups: anchored.map((a) => a.group),
    counts: { restyled, spacingCleared, citesRepaired }
  };
}

// ---------------------------------------------------------------------------
// Public: planMarkCite
// ---------------------------------------------------------------------------

/**
 * Mark-cite: the cite convention (bold CITE_PT) over the user's selection —
 * one RequestGroup so the controller routes it through the same gate/chunk/
 * revision machinery as every other verb.
 *
 * The selection arrives from the adapter as-is: there is no doc view here to
 * filter against, so the A9 text-only whitelist is NOT applied (flagged
 * deviation, by design) — a selection containing a chip gets the same
 * run-level bold/size write the Docs toolbar itself would make, and the
 * whitelist exists to protect HIDE's restore exactness, which an explicit
 * user-selection write cannot break (no rstm state exists over it — the
 * controller's assertNotHidden gate ran first). Empty selections are the
 * adapter's refusal (markCiteNoop), never planned here.
 */
export function planMarkCite(selection: GRange): RequestGroup {
  return { requests: [citeConventionRequest(selection)] };
}
