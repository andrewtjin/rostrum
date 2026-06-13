// The documents.get interpreter (plan S3 + A3/A9/A11.ii/A13) — the ONE place
// raw Docs API JSON is decoded into the engine's GDoc view. The adapter hands
// the response over as `unknown` ON PURPOSE (DocsPort.fetchDocument): every
// assumption about Google's wire shapes lives here, behind tests pinned to the
// official discovery schema, so the rest of the engine reasons about a typed
// view it can trust.
//
// Two non-negotiables shape every function in this file:
//   * DEFENSIVE, NEVER THROWING. A malformed piece (missing textStyle, junk
//     element type, absent index) decodes to a safe default instead of
//     aborting the verb — the guards and planner downstream decide what is
//     actionable; the parser's job is only to never lie about what it saw.
//     Safe defaults are always the CONSERVATIVE reading (inherit / not bold /
//     no background / kind "other"), because conservative readings keep text
//     rather than hide it.
//   * OMITTED MEANS ZERO on the wire. Real payloads drop zero-valued fields:
//     rgb channels at 0.0, indexes at 0, dimension magnitudes at 0. Decoding
//     must re-materialize those zeros (plan A11.ii fixture lint exists to keep
//     our fixtures honest about this exact trap).

import { GDoc, GElement, GNamedRange, GNamedStyleType, GParagraph, GRange } from "./types";

// ---------------------------------------------------------------------------
// DOC_FIELDS_MASK — payload discipline (plan A13)
// ---------------------------------------------------------------------------

/** TextRun is the hot path (thousands per doc), so its selection is tight:
 * exactly the three style channels the engine reads, plus the suggestion keys
 * the Hide gate depends on (presence of any `suggested*` key = refusal). */
const TEXT_RUN_FIELDS =
  "content,suggestedInsertionIds,suggestedDeletionIds,suggestedTextStyleChanges," +
  "textStyle(fontSize,bold,backgroundColor)";

/**
 * Non-text ParagraphElement types are selected WHOLE on purpose: they are rare
 * (chips, breaks, objects), their bodies are small, and each carries its own
 * suggested* keys — sub-masking them would blind the suggestion gate to e.g. a
 * suggested chip insertion while saving almost no bytes. The whitelist (plan
 * A9) keys off textRun ABSENCE, so these stay `kind: "other"` regardless of
 * what their bodies contain.
 */
const ELEMENT_FIELDS =
  `startIndex,endIndex,textRun(${TEXT_RUN_FIELDS}),` +
  "autoText,columnBreak,dateElement,equation,footnoteReference,horizontalRule," +
  "inlineObjectElement,pageBreak,person,richLink";

const PARAGRAPH_FIELDS =
  `elements(${ELEMENT_FIELDS}),` +
  "paragraphStyle(namedStyleType,spaceAbove,spaceBelow)," +
  "suggestedParagraphStyleChanges";

/** Table cell `content` is selected whole: the cell interior is recursive
 * StructuralElement JSON that a flat field mask cannot re-select, tables are
 * rare in debate docs, and their paragraphs are kept whole anyway (never
 * style-targeted), so the over-fetch is bounded and harmless. */
const STRUCTURAL_FIELDS =
  `startIndex,endIndex,paragraph(${PARAGRAPH_FIELDS}),table(tableRows(tableCells(content)))`;

/** Everything the engine reads out of one content segment. Shared verbatim
 * between the legacy top-level fields and tabs[].documentTab so the two read
 * paths can never drift apart. namedRanges is a name-keyed map (arbitrary
 * keys), which a field mask cannot reach into — selected whole. */
const SEGMENT_FIELDS =
  `body(content(${STRUCTURAL_FIELDS})),namedRanges,` +
  "namedStyles(styles(namedStyleType,textStyle(fontSize)))";

/**
 * The fields mask every documents.get carries (plan A13 — ONE masked get per
 * verb). `childTabs` is selected whole because leaf-tab counting (plan A3)
 * needs the recursive tab tree and field masks cannot recurse; single-tab docs
 * have no child tabs so this costs nothing, and multi-tab docs refuse outright
 * right after the count, so the over-fetch never reaches a hot path.
 */
export const DOC_FIELDS_MASK: string =
  `revisionId,${SEGMENT_FIELDS},tabs(tabProperties(tabId),childTabs,documentTab(${SEGMENT_FIELDS}))`;

// ---------------------------------------------------------------------------
// Narrowing primitives — every raw access funnels through these so "malformed
// piece -> safe default" is a property of the file, not of each call site.
// ---------------------------------------------------------------------------

/** Plain-object check; arrays are NOT records (a junk array where an object
 * belongs must fall to the default path, not be walked as keyed fields). */
function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** Junk-tolerant array read: anything that is not an array reads as empty. */
function asArray(v: unknown): unknown[] {
  return Array.isArray(v) ? v : [];
}

function asString(v: unknown, fallback: string): string {
  return typeof v === "string" ? v : fallback;
}

/** Finite-number read. NaN/Infinity/junk fall back — indexes and sizes that
 * are not real numbers must never propagate into range arithmetic. */
function finiteOr(v: unknown, fallback: number): number {
  return typeof v === "number" && Number.isFinite(v) ? v : fallback;
}

/** The closed set of Docs named paragraph styles. An unknown/missing value
 * decodes to NORMAL_TEXT — the conservative reading: body text is the most
 * cautious thing to call a paragraph we cannot classify (headings would be
 * kept whole on style alone, which a junk value has not earned). */
const NAMED_STYLE_TYPES: ReadonlySet<string> = new Set<GNamedStyleType>([
  "NORMAL_TEXT",
  "TITLE",
  "SUBTITLE",
  "HEADING_1",
  "HEADING_2",
  "HEADING_3",
  "HEADING_4",
  "HEADING_5",
  "HEADING_6"
]);

function asNamedStyleType(v: unknown): GNamedStyleType {
  return typeof v === "string" && NAMED_STYLE_TYPES.has(v) ? (v as GNamedStyleType) : "NORMAL_TEXT";
}

// ---------------------------------------------------------------------------
// Wire-value decoders
// ---------------------------------------------------------------------------

/** One rgb channel (0.0-1.0 float) to its 0-255 byte. ABSENT CHANNELS ARE 0.0:
 * real payloads omit zero channels (#ffff00 arrives as {red:1,green:1}), so
 * the missing-key path is the NORMAL path, not an error path. Junk values and
 * out-of-range floats clamp rather than poison the hex. */
function channelByte(v: unknown): number {
  const n = finiteOr(v, 0);
  return Math.round(Math.min(1, Math.max(0, n)) * 255);
}

function byteHex(b: number): string {
  return b.toString(16).padStart(2, "0");
}

/**
 * OptionalColor -> lower-case "#rrggbb" | null. The API's three-layer nesting
 * encodes real states: backgroundColor absent = inherited (null), present
 * with `color` unset = explicitly transparent (also null — no highlight to
 * keep), `color` set = an opaque color whose omitted channels are zeros. A
 * bare `color: {}` therefore decodes to "#000000", not null — it claims
 * opacity, and the conservative reading of an opaque unknown is black (a
 * keepable highlight), never "no highlight".
 */
function decodeBackgroundHex(textStyle: Record<string, unknown>): string | null {
  const optional = textStyle.backgroundColor;
  if (!isRecord(optional)) return null;
  const color = optional.color;
  if (!isRecord(color)) return null;
  const rgb = isRecord(color.rgbColor) ? color.rgbColor : {};
  return `#${byteHex(channelByte(rgb.red))}${byteHex(channelByte(rgb.green))}${byteHex(channelByte(rgb.blue))}`;
}

/** TextStyle.fontSize -> points | null (null = inherits from the named style).
 * A fontSize object without a magnitude stays null: 0pt text cannot exist, so
 * unlike spacing there is no omitted-zero to re-materialize here. */
function fontSizePtOf(textStyle: Record<string, unknown>): number | null {
  const dim = textStyle.fontSize;
  if (!isRecord(dim)) return null;
  const m = dim.magnitude;
  return typeof m === "number" && Number.isFinite(m) ? m : null;
}

/**
 * Direct paragraph spacing -> points | null. PRESENCE of the dimension is the
 * signal: a paragraph with no spaceAbove key inherits (null), while
 * `spaceAbove: {unit:"PT"}` is an EXPLICIT zero whose magnitude the API
 * omitted (the omitted-means-zero rule). Conflating those two would make the
 * styles retro pass (plan A5) skip exactly the direct-zero paragraphs it
 * exists to clear.
 */
function directSpacingPt(v: unknown): number | null {
  if (!isRecord(v)) return null;
  return finiteOr(v.magnitude, 0);
}

// ---------------------------------------------------------------------------
// Element / paragraph decoding
// ---------------------------------------------------------------------------

/**
 * One ParagraphElement -> GElement. THE WHITELIST LIVES HERE (plan A9): only
 * an element carrying a `textRun` object becomes `kind: "text"`; every other
 * shape — chips, breaks, objects, junk, and element types Google has not
 * shipped yet — is `kind: "other"`, which the planner treats as an immovable
 * wall (kept, spans break around it). `fallbackStart` chains the previous
 * element's end so a missing index degrades to "zero-width here" instead of
 * teleporting the element to 0 and corrupting downstream range math.
 */
function parseElement(raw: unknown, fallbackStart: number): GElement {
  const rec = isRecord(raw) ? raw : {};
  const run = isRecord(rec.textRun) ? rec.textRun : null;
  const text = run !== null ? asString(run.content, "") : "";
  const startIndex = finiteOr(rec.startIndex, fallbackStart);
  // A text element's span must cover its content when the wire omits endIndex;
  // an "other" element defaults to zero width (we know nothing about it).
  const endIndex = finiteOr(rec.endIndex, run !== null ? startIndex + text.length : startIndex);
  if (run === null) {
    return { startIndex, endIndex, kind: "other", text: "", fontSizePt: null, bold: false, backgroundHex: null };
  }
  const style = isRecord(run.textStyle) ? run.textStyle : {};
  return {
    startIndex,
    endIndex,
    kind: "text",
    text,
    fontSizePt: fontSizePtOf(style),
    bold: style.bold === true,
    backgroundHex: decodeBackgroundHex(style)
  };
}

/** A structural element + its paragraph payload, queued for ordinal stamping.
 * Collection and stamping are separate passes because `index` (receipt
 * ordinal) and `isLastInSegment` (final-newline clamp) are properties of the
 * FLATTENED sequence, unknowable while still descending into tables. */
interface PendingParagraph {
  se: Record<string, unknown>;
  para: Record<string, unknown>;
  inTable: boolean;
}

/**
 * Flatten one content array (body or table cell) into pending paragraphs, in
 * document order. Table cells recurse with inTable=true — their paragraphs are
 * REPRESENTED (receipts count them; the keeper marks them kept) but never
 * style-targeted (Word decision #16 parity). Anything that is neither a
 * paragraph nor a table (the masked-out sectionBreak stub, tableOfContents,
 * future structural types) contributes nothing: the engine never targets what
 * it never saw, so skipping is the conservative move.
 */
function collectParagraphs(content: unknown[], inTable: boolean, out: PendingParagraph[]): void {
  for (const item of content) {
    if (!isRecord(item)) continue;
    if (isRecord(item.paragraph)) {
      out.push({ se: item, para: item.paragraph, inTable });
      continue;
    }
    if (isRecord(item.table)) {
      for (const row of asArray(item.table.tableRows)) {
        if (!isRecord(row)) continue;
        for (const cell of asArray(row.tableCells)) {
          if (!isRecord(cell)) continue;
          // Nested tables stay inTable — depth never re-arms eligibility.
          collectParagraphs(asArray(cell.content), true, out);
        }
      }
    }
  }
}

/** Decode one pending paragraph at its flattened ordinal. */
function parseParagraph(pending: PendingParagraph, ordinal: number, isLast: boolean): GParagraph {
  const style = isRecord(pending.para.paragraphStyle) ? pending.para.paragraphStyle : {};
  const startIndex = finiteOr(pending.se.startIndex, 0);
  // Elements chain their fallback starts so one missing index stays local.
  let cursor = startIndex;
  const elements: GElement[] = asArray(pending.para.elements).map((e) => {
    const el = parseElement(e, cursor);
    cursor = el.endIndex;
    return el;
  });
  const lastElement = elements.length > 0 ? elements[elements.length - 1] : null;
  return {
    index: ordinal,
    startIndex,
    // A paragraph missing endIndex closes at its last element (or is empty) —
    // never past content it does not have.
    endIndex: finiteOr(pending.se.endIndex, lastElement !== null ? lastElement.endIndex : startIndex),
    namedStyleType: asNamedStyleType(style.namedStyleType),
    inTable: pending.inTable,
    isLastInSegment: isLast,
    spaceAbovePt: directSpacingPt(style.spaceAbove),
    spaceBelowPt: directSpacingPt(style.spaceBelow),
    elements
  };
}

// ---------------------------------------------------------------------------
// Document-level pieces
// ---------------------------------------------------------------------------

/** Total LEAF tabs (plan A3): a tab with children is a container whose leaves
 * are what the user perceives as "tabs"; counting both would double-count. */
function countLeafTabs(tabs: unknown[]): number {
  let count = 0;
  for (const tab of tabs) {
    if (!isRecord(tab)) continue;
    const children = asArray(tab.childTabs);
    count += children.length > 0 ? countLeafTabs(children) : 1;
  }
  return count;
}

/**
 * namedRanges map -> flat GNamedRange[]. The wire shape is doubly nested
 * (name-keyed map -> NamedRanges group -> NamedRange entries -> ranges[]):
 * EVERY entry and EVERY segment is preserved, because split ranges (edit in
 * the middle of a hidden region) and same-name siblings (two regions hidden
 * with identical RLE) are both normal states the restore path must see whole.
 * A segment's omitted startIndex re-materializes as 0 (omitted-means-zero).
 */
function parseNamedRanges(v: unknown): GNamedRange[] {
  if (!isRecord(v)) return [];
  const out: GNamedRange[] = [];
  for (const key of Object.keys(v)) {
    const group = v[key];
    if (!isRecord(group)) continue;
    for (const entry of asArray(group.namedRanges)) {
      if (!isRecord(entry)) continue;
      const segments: GRange[] = [];
      for (const range of asArray(entry.ranges)) {
        if (!isRecord(range)) continue;
        const startIndex = finiteOr(range.startIndex, 0);
        segments.push({ startIndex, endIndex: finiteOr(range.endIndex, startIndex) });
      }
      // The entry's own name wins; the group name (the map's shared name) is
      // the fallback so a masked/partial entry still claims its family.
      out.push({
        id: asString(entry.namedRangeId, ""),
        name: asString(entry.name, asString(group.name, "")),
        segments
      });
    }
  }
  return out;
}

/** namedStyles -> per-style font size map. Only sizes actually STATED land in
 * the map (Partial by design): the cite predicate and receipts resolve null
 * run sizes through it and must know the difference between "11pt" and
 * "the read did not say". */
function parseNamedStyleSizes(v: unknown): GDoc["namedStyleSizesPt"] {
  const out: Partial<Record<GNamedStyleType, number>> = {};
  if (!isRecord(v)) return out;
  for (const entry of asArray(v.styles)) {
    if (!isRecord(entry)) continue;
    const type = entry.namedStyleType;
    if (typeof type !== "string" || !NAMED_STYLE_TYPES.has(type)) continue;
    const size = fontSizePtOf(isRecord(entry.textStyle) ? entry.textStyle : {});
    if (size !== null) out[type as GNamedStyleType] = size;
  }
  return out;
}

/**
 * True when ANY key matching /^suggested/ appears anywhere in the read (plan
 * A16). The walk is GENERIC — not a list of known suggestion fields — so a
 * suggestion surface this build has never heard of still trips the Hide gate.
 * Key PRESENCE is the signal: the API omits empty suggestion arrays entirely,
 * so a key that made it onto the wire means a live suggestion.
 */
function hasSuggestedKey(v: unknown): boolean {
  if (Array.isArray(v)) return v.some(hasSuggestedKey);
  if (!isRecord(v)) return false;
  for (const key of Object.keys(v)) {
    if (key.startsWith("suggested")) return true;
    if (hasSuggestedKey(v[key])) return true;
  }
  return false;
}

// ---------------------------------------------------------------------------
// parseDocument — the public entry point
// ---------------------------------------------------------------------------

/**
 * Decode a raw (fields-masked) documents.get response into the engine's GDoc
 * view. Defensive end to end: ANY input — null, a string, a half-shaped
 * object — yields a structurally valid GDoc; malformed pieces decode to safe
 * defaults and nothing here ever throws. Content resolution order:
 *
 *   * `tabs` present (includeTabsContent:true reality — every doc has >= 1
 *     tab): count leaf tabs, then read the FIRST tab's documentTab segment.
 *   * First tab carries no documentTab (includeTabsContent:false responses
 *     keep content in the legacy top-level fields) or no `tabs` key at all
 *     (older fixtures/payloads): read the legacy top-level segment.
 *
 * Only the single segment is read in v1 — the tab gate (guards.assertSingleTab)
 * refuses multi-tab docs before any index from here is acted on (plan A3).
 */
export function parseDocument(raw: unknown): GDoc {
  const root = isRecord(raw) ? raw : {};

  // --- tab resolution (plan A3) ---
  const tabs = Array.isArray(root.tabs) ? root.tabs : null;
  let tabCount = 1;
  // The segment holding body/namedRanges/namedStyles: first tab's documentTab
  // when present, else the legacy top-level fields (same key names by design).
  let segment: Record<string, unknown> = root;
  if (tabs !== null) {
    // Clamped to >= 1: an empty/garbage tabs array is still ONE readable doc
    // (the legacy fallback below), and a 0 would read as nonsense everywhere
    // tabCount surfaces. Real multi-tab counts pass through untouched.
    tabCount = Math.max(1, countLeafTabs(tabs));
    const first = tabs.find(isRecord);
    if (first !== undefined && isRecord(first.documentTab)) {
      segment = first.documentTab;
    }
  }

  // --- body flattening ---
  const body = isRecord(segment.body) ? segment.body : {};
  const pending: PendingParagraph[] = [];
  collectParagraphs(asArray(body.content), false, pending);
  const paragraphs = pending.map((p, i) => parseParagraph(p, i, i === pending.length - 1));

  return {
    revisionId: asString(root.revisionId, ""),
    tabCount,
    paragraphs,
    namedRanges: parseNamedRanges(segment.namedRanges),
    namedStyleSizesPt: parseNamedStyleSizes(segment.namedStyles),
    // The suggestion walk covers the WHOLE raw read (not just the first tab's
    // segment): a suggestion anywhere makes every index untrustworthy.
    suggestionsPresent: hasSuggestedKey(raw)
  };
}
