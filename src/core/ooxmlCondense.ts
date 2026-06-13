// The OOXML editor for Condense & Shrink — multi-`<w:p>` RANGE FRAGMENT surgery.
//
// WHY THIS IS A SEPARATE MODULE (and not part of ooxml.ts). `ooxml.ts` is "the one and only place"
// that edits a SINGLE body paragraph's run visibility (`<w:vanish/>` for Invisibility Mode), and
// `ooxmlPackage.ts` splices a WHOLE-body package. Condense & Shrink need a THIRD shape: read every
// `<w:p>` in the active-range fragment, merge paragraphs / collapse whitespace / resize runs / insert
// and remove reversible break markers, then re-serialize the whole fragment. That is enough distinct
// OOXML mechanics to warrant its own focused, separately-tested module — exactly the precedent
// `ooxmlPackage.ts` set ("mirrors ooxml.ts's rationale; kept local to stay self-contained"). The pure
// POLICY (the size ladder, mode selection, omission scan) lives in `shrink.ts` / `condense.ts`; this
// file is the mechanical string-in / string-out transform they drive, so the whole feature stays
// unit-tested in Node with no Word host.
//
// Everything here is pure: string in, string out, via @xmldom/xmldom — no Office.js.

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { MARK_SIGNATURE, MARK_SIGNATURES, CITE_STYLE_ID } from "./styles";
import { CondenseOptions, FragmentRunView } from "./types";

// xmldom node types vary across versions; we keep public signatures fully typed (string in/out,
// number/null) and use a localized `any` for node handles — the same pragmatic choice ooxml.ts and
// ooxmlPackage.ts make for a version-robust XML adapter.
/* eslint-disable @typescript-eslint/no-explicit-any */

const ELEMENT_NODE = 1;
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

/**
 * The two visible boundary GLYPHS a merge-mode marker can carry: a plain space (full-merge) or a 6pt
 * `¶` (pilcrow mode). Each is exactly ONE UTF-16 code unit, so the split tokenizer drops a boundary's
 * glyph with a single `.slice(1)` (the SIGNATURE's own length is consumed by the capture-split). They
 * double as a GUARD: a signature is only honored as a boundary when the char that follows it is one of
 * these — so a stray signature in user text (essentially nonexistent for the ZWSP+WJ pair, near-zero
 * for legacy U+2063, but never exactly zero) can never silently eat the following character.
 */
const BOUNDARY_GLYPHS = new Set<string>([" ", "¶"]);

// ---------------------------------------------------------------------------
// Signature-set regexes. WRITE paths emit only MARK_SIGNATURE (ZWSP+WJ); every READ/STRIP path below
// must honor the FULL set (current pair + legacy single U+2063) so docs condensed by the 2026-06-10
// deployed build still uncondense losslessly — and a single run may carry BOTH kinds after
// re-condensing a legacy doc. Signatures are multi-code-unit SEQUENCES, so these are ALTERNATIONS
// (longest first — a shorter alternative must never shadow a longer one), NOT a character class: a
// lone organic U+200B (endemic in web-pasted text) must match NOTHING — single-char containment is
// exactly what let organic ZWSPs fabricate breaks and poison the shrink view (2026-06-10 repro).
// Built once from MARK_SIGNATURES so adding/retiring a signature is a styles.ts-only change.
// ---------------------------------------------------------------------------

/** Alternation body matching ANY known signature, longest first (no member needs regex-escaping). */
const SIGNATURE_ALT = [...MARK_SIGNATURES].sort((a, b) => b.length - a.length).join("|");
/** Test: does a string contain any signature? */
const ANY_SIGNATURE = new RegExp(SIGNATURE_ALT, "u");
/** Strip: remove every signature of every kind (fresh `g` regex — no lastIndex state to share). */
const ALL_SIGNATURES_G = new RegExp(SIGNATURE_ALT, "gu");
/**
 * Split keeping the separator: `text.split(SIGNATURE_SPLIT)` yields `[body0, sig1, part1, sig2, …]`
 * — the captured group preserves WHICH signature preceded each part, so the unmerge tokenizer can
 * re-emit a non-boundary signature VERBATIM as its exact original sequence (lossless even for a
 * legacy stray in a mixed run; never normalized to the current written form).
 */
const SIGNATURE_SPLIT = new RegExp(`(${SIGNATURE_ALT})`, "u");

/**
 * Namespace declarations for the standalone wrapper used to re-parse a stored `<w:pPr>`/`<w:rPr>`
 * payload (see {@link importPPr}/{@link importRPr}). A payload is serialized out of the source document
 * by xmldom, which does NOT re-emit namespace declarations the node merely inherited from the package
 * root — so a mark that carried e.g. a `w14:` property would lose that prefix on a bare `xmlns:w`-only
 * re-parse. Declaring the prefixes Word actually emits keeps the round-trip lossless.
 */
const WRAP_NS =
  `xmlns:w="${W_NS}"` +
  ` xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"` +
  ` xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml"` +
  ` xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"` +
  ` xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"`;

/**
 * One-or-more of: ASCII control whitespace, or any Unicode Space_Separator (`\p{Zs}` — ASCII space,
 * NBSP, figure/thin/narrow spaces, …). Condense normalizes every such run to a single ASCII space
 * (Verbatim's `^t/^l/^n` + NBSP + double-space passes). `<w:tab>`/`<w:br>`/`<w:cr>` are ELEMENTS and
 * are rendered to "\n" by the run walker first, so they fall into this class too.
 */
const WHITESPACE_RUN = /(?:[\t\n\v\f\r]|\p{Zs})+/gu;

/**
 * The CT_RPr children that, per the OOXML schema, come AFTER `<w:sz>`/`<w:szCs>`. `setRunSize` inserts
 * the size elements just before the first of these (else appends), so a resized run stays schema-valid
 * — Word's `insertOoxml` is lenient but a wrong-ordered rPr is a needless corruption risk.
 */
const RPR_AFTER_SZ = new Set<string>([
  "w:highlight",
  "w:u",
  "w:effect",
  "w:bdr",
  "w:shd",
  "w:fitText",
  "w:vertAlign",
  "w:rtl",
  "w:cs",
  "w:em",
  "w:lang",
  "w:eastAsianLayout",
  "w:specVanish",
  "w:oMath"
]);

/** Half-points for a 6pt mark/pilcrow (Verbatim's shrunk-paragraph and pilcrow size). */
const SIX_PT_HALF = 12;

// ---------------------------------------------------------------------------
// Parse / serialize / scope (local, mirroring ooxml.ts + ooxmlPackage.ts)
// ---------------------------------------------------------------------------

function parse(xml: string): any {
  return new DOMParser({
    onError: (level: string, message: string) => {
      if (level === "fatalError") throw new Error(message);
    }
  }).parseFromString(xml, "text/xml");
}

function serialize(node: any): string {
  return new XMLSerializer().serializeToString(node);
}

/**
 * An opaque pre-parsed fragment handle for read-then-write callers. A Shrink press reads the
 * fragment (`readFragmentParagraphs`) and then writes sizes back (`applyFragmentShrink`) over the
 * SAME unchanged string, so parsing it twice is pure waste — and parsing dominates a press because
 * the flat-OPC fragment always bundles the full styles.xml, even for a one-card selection.
 * `parseFragment` lets such callers parse ONCE and thread the tree through both calls — the same
 * read-then-mutate-one-tree fusion `ParsedParagraph` gave classifyParagraph in ooxml.ts. The handle
 * is opaque (callers never touch xmldom internals), and the trailing optional params keep the
 * legacy parse-per-call path callable so tests can assert fused/legacy byte-identity forever.
 */
export interface ParsedFragment {
  /** @internal The xmldom document — only this module may read or mutate it. */
  doc: any;
}

/** Parse a fragment once, for threading through `readFragmentParagraphs` + `applyFragmentShrink`. */
export function parseFragment(fragmentXml: string): ParsedFragment {
  return { doc: parse(fragmentXml) };
}

/** The element to search paragraphs within: the first `<w:body>` if present, else the document root. */
function bodyScope(doc: any): any {
  const bodies = doc.getElementsByTagName("w:body");
  return bodies && bodies.length > 0 ? bodies.item(0) : doc;
}

/** True when a node lives inside a text box (`<w:txbxContent>`) — not part of the body story. */
function isInTextbox(node: any): boolean {
  let n = node ? node.parentNode : null;
  while (n) {
    if (n.nodeName === "w:txbxContent") return true;
    n = n.parentNode;
  }
  return false;
}

/** STORY `<w:p>` (every body paragraph except textbox-nested), document order — what Shrink iterates. */
function storyParagraphs(scope: any): any[] {
  const live = scope.getElementsByTagName("w:p");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) {
    const p = live.item(i);
    if (p && !isInTextbox(p)) out.push(p);
  }
  return out;
}

/** Direct-child `<w:p>` of a scope (top-level paragraphs — what Condense's merge operates on, so it
 *  never folds table-cell or textbox paragraphs into a body paragraph). Document order. */
function directParagraphs(scope: any): any[] {
  const out: any[] = [];
  const kids = scope.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item ? kids.item(i) : kids[i];
    if (k && k.nodeType === ELEMENT_NODE && k.nodeName === "w:p") out.push(k);
  }
  return out;
}

function directChildren(el: any, name: string): any[] {
  const out: any[] = [];
  const kids = el.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE && k.nodeName === name) out.push(k);
  }
  return out;
}

function firstDirectChild(el: any, name: string): any | null {
  const found = directChildren(el, name);
  return found.length > 0 ? found[0] : null;
}

/** The `<w:r>` descendants of a paragraph in document order (matches `readFragmentParagraphs`). */
function paragraphRuns(pEl: any): any[] {
  const live = pEl.getElementsByTagName("w:r");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) out.push(live.item(i));
  return out;
}

// ---------------------------------------------------------------------------
// Run inspection
// ---------------------------------------------------------------------------

function runRPr(runEl: any): any | null {
  return firstDirectChild(runEl, "w:rPr");
}

/** Concatenated `<w:t>` text of a run (no tab/break conversion) — for marker / payload detection. */
function runTextRaw(runEl: any): string {
  let text = "";
  const ts = runEl.getElementsByTagName("w:t");
  for (let i = 0; i < ts.length; i++) {
    const t = ts.item(i);
    text += t.textContent != null ? t.textContent : "";
  }
  return text;
}

/**
 * Run text rendered for a view: `<w:t>` verbatim, `<w:tab>` -> \t, `<w:br>`/`<w:cr>` -> \n. The marker
 * SIGNATURE is filtered OUT of the view: it is an internal boundary signal, never user-visible text, so
 * a view (Shrink's per-run read, the whitespace-collapse rebuild, restored segment text) must never
 * carry it. Detection/splitting use `runTextRaw` instead, which keeps signatures intact. Only FULL
 * signatures are filtered — a lone organic ZWSP in user text is real content and passes through.
 */
function runTextView(runEl: any): string {
  let text = "";
  const walk = (node: any): void => {
    const kids = node.childNodes;
    for (let i = 0; i < kids.length; i++) {
      const k = kids.item(i);
      if (k.nodeType !== ELEMENT_NODE) continue;
      if (k.nodeName === "w:t") text += k.textContent != null ? k.textContent : "";
      else if (k.nodeName === "w:tab") text += "\t";
      else if (k.nodeName === "w:br" || k.nodeName === "w:cr") text += "\n";
      else walk(k);
    }
  };
  walk(runEl);
  return text.replace(ALL_SIGNATURES_G, ""); // filter EVERY signature kind (current pair + legacy)
}

function runHighlight(runEl: any): string | null {
  const rPr = runRPr(runEl);
  if (!rPr) return null;
  const hl = firstDirectChild(rPr, "w:highlight");
  if (!hl) return null;
  const val = (hl.getAttribute("w:val") || "").toLowerCase();
  return val === "" || val === "none" ? null : val;
}

function runRStyle(runEl: any): string | null {
  const rPr = runRPr(runEl);
  if (!rPr) return null;
  const rStyle = firstDirectChild(rPr, "w:rStyle");
  return rStyle ? rStyle.getAttribute("w:val") || null : null;
}

/**
 * Tri-state DIRECT underline on a run: true (`<w:u>` present with a value other than none/0/false),
 * false (explicitly none/0/false), or undefined (no direct `<w:u>` → inherit from the character style).
 * Kept tri-state so a direct `<w:u w:val="none"/>` can OVERRIDE a character style that underlines.
 */
function directUnderlineTri(runEl: any): boolean | undefined {
  const rPr = runRPr(runEl);
  if (!rPr) return undefined;
  const u = firstDirectChild(rPr, "w:u");
  if (!u) return undefined;
  const val = (u.getAttribute("w:val") || "").toLowerCase();
  return !(val === "none" || val === "0" || val === "false");
}

/**
 * Tri-state DIRECT character box (`<w:bdr>`) on a run: true (a real border), false (nil/none), or
 * undefined (no direct `<w:bdr>` → inherit from the character style). Same override semantics as underline.
 */
function directBoxTri(runEl: any): boolean | undefined {
  const rPr = runRPr(runEl);
  if (!rPr) return undefined;
  const b = firstDirectChild(rPr, "w:bdr");
  if (!b) return undefined;
  const val = (b.getAttribute("w:val") || "").toLowerCase();
  return !(val === "nil" || val === "none" || val === "");
}

function vanishOn(rPr: any): boolean {
  if (!rPr) return false;
  const v = firstDirectChild(rPr, "w:vanish");
  if (!v) return false;
  const val = (v.getAttribute("w:val") || "").toLowerCase();
  return !(val === "false" || val === "0" || val === "off");
}

const INELIGIBLE_RUN_TAGS = ["w:fldChar", "w:instrText", "w:footnoteReference", "w:endnoteReference", "w:drawing", "w:object", "w:pict"];

/** The internal-part tags: a `<w:drawing>`/`<w:object>`/`<w:pict>` is an embedded media/OLE part. */
const INTERNAL_PART_TAGS = ["w:drawing", "w:object", "w:pict"];

/** True when a run embeds an internal part (drawing/object/pict) anywhere in its subtree. */
function runHasInternalPart(runEl: any): boolean {
  for (const tag of INTERNAL_PART_TAGS) {
    if (runEl.getElementsByTagName(tag).length > 0) return true;
  }
  return false;
}

/** Allowed element children of a "simple text" run (whitespace-normalizable shape). */
const SIMPLE_TEXT_RUN_TAGS = new Set<string>(["w:rPr", "w:t", "w:tab", "w:br", "w:cr"]);

/** Allowed element children of a "text only" run (`runTextRaw`-rebuildable shape — see below). */
const TEXT_ONLY_RUN_TAGS = new Set<string>(["w:rPr", "w:t"]);

/** True when every ELEMENT child of a run is in `allowed` (the shared walk behind both shape checks). */
function runHasOnlyChildren(runEl: any, allowed: ReadonlySet<string>): boolean {
  const kids = runEl.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE && !allowed.has(k.nodeName)) return false;
  }
  return true;
}

/** True when a run is plain text (only rPr / w:t / w:tab / w:br / w:cr) — safe to normalize whitespace. */
function isSimpleTextRun(runEl: any): boolean {
  return runHasOnlyChildren(runEl, SIMPLE_TEXT_RUN_TAGS);
}

/**
 * True when a run carries NOTHING but its rPr and `<w:t>` text. This is the shape a glyph marker must
 * have to be split safely: the tokenizer rebuilds a marker run from `runTextRaw`, which concatenates
 * the `<w:t>`s and silently DROPS every other child — so honoring a marker inside a run that also
 * carries a `<w:tab/>`/`<w:br/>`/drawing would destroy that content (e.g. `<w:t>AAA{U+200B}</w:t>`
 * `<w:tab/><w:t>{U+2060} BBB</w:t>` would fabricate a break AND lose the tab). Condense only ever writes
 * simple `rPr+t` glyph markers, so requiring this shape costs no genuine detection.
 */
function isTextOnlyRun(runEl: any): boolean {
  return runHasOnlyChildren(runEl, TEXT_ONLY_RUN_TAGS);
}

function runEligible(runEl: any): boolean {
  for (const tag of INELIGIBLE_RUN_TAGS) {
    if (runEl.getElementsByTagName(tag).length > 0) return false;
  }
  let ancestor = runEl.parentNode;
  while (ancestor && ancestor.nodeName !== "w:p") {
    if (ancestor.nodeName === "w:fldSimple") return false;
    ancestor = ancestor.parentNode;
  }
  return true;
}

/** Explicit `<w:sz>` value (half-points) on a run, or null when it inherits from the style. */
function runSizeHalfPts(runEl: any): number | null {
  const rPr = runRPr(runEl);
  if (!rPr) return null;
  const sz = firstDirectChild(rPr, "w:sz");
  if (!sz) return null;
  const n = parseInt(sz.getAttribute("w:val") || "", 10);
  return Number.isFinite(n) ? n : null;
}

/** True when a run carries ANY marker signature (current pair or legacy) in its text — the condense signal. */
function containsSignature(runEl: any): boolean {
  return ANY_SIGNATURE.test(runTextRaw(runEl));
}

/** A condense MARKER run (glyph boundary OR data payload), identified by the intrinsic text sentinel. */
function isMarkerRun(runEl: any): boolean {
  return containsSignature(runEl);
}

/**
 * A marker run's payload text: its raw text with every signature stripped and edges trimmed. A data run
 * stores serialized XML as `SIGNATURE + "<w:pPr…/>"`; the signature must be removed before the XML is
 * classified or re-parsed (a leading signature would make `<w:pPr` fail both `startsWith` and the parser).
 */
function payloadText(runEl: any): string {
  return runTextRaw(runEl).replace(ALL_SIGNATURES_G, "").trim();
}

/** A marker run that stores a divergent following-paragraph `<w:pPr>` (vanished; payload begins `<w:pPr`). */
function isPPrDataRun(runEl: any): boolean {
  return containsSignature(runEl) && vanishOn(runRPr(runEl)) && payloadText(runEl).startsWith("<w:pPr");
}

/**
 * A marker run that stores a retain-mode dropped blank's ORIGINAL mark `<w:rPr>` (vanished; payload
 * begins `<w:rPr`), so Uncondense can restore the user's paragraph mark verbatim — see
 * {@link dropBlankParagraphs}. (`<w:rPr/>` means the blank's mark had no own properties.)
 */
function isMarkRPrDataRun(runEl: any): boolean {
  return containsSignature(runEl) && vanishOn(runRPr(runEl)) && payloadText(runEl).startsWith("<w:rPr");
}

/** A marker run carrying a serialized-XML payload (a stored pPr or mark rPr) rather than a visible glyph. */
function isDataRun(runEl: any): boolean {
  return isPPrDataRun(runEl) || isMarkRPrDataRun(runEl);
}

/**
 * A marker run that is a VISIBLE boundary glyph (a space or `¶`) — the actual paragraph break.
 * REQUIRES the text-only shape: the split rebuilds the run from `runTextRaw`, so a run that also
 * carries non-text children (`<w:tab/>`, `<w:br/>`, a drawing…) must FAIL detection and be left
 * verbatim — losslessness outranks marker detection (fail-safe). Genuine condense-written glyph
 * markers are always `rPr+t` only, so this rejects no real marker.
 */
function isGlyphMarkerRun(runEl: any): boolean {
  return isMarkerRun(runEl) && !isDataRun(runEl) && isTextOnlyRun(runEl);
}

// ---------------------------------------------------------------------------
// rPr mutation helpers
// ---------------------------------------------------------------------------

function ensureRPr(doc: any, runEl: any): any {
  const existing = runRPr(runEl);
  if (existing) return existing;
  const rPr = doc.createElement("w:rPr");
  runEl.insertBefore(rPr, runEl.firstChild);
  return rPr;
}

/** Insert `el` into `rPr` in schema order relative to `<w:sz>` (before the first AFTER_SZ child, else append). */
function insertSizeChild(rPr: any, el: any): void {
  const kids = rPr.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE && RPR_AFTER_SZ.has(k.nodeName)) {
      rPr.insertBefore(el, k);
      return;
    }
  }
  rPr.appendChild(el);
}

/**
 * Set or clear a run's explicit font size on the GIVEN rPr. `halfPts` number → write `<w:sz>` AND
 * `<w:szCs>` (both, so complex-script text shrinks too); null → remove both (revert to inherited
 * Normal). Returns true when the rPr actually changed.
 */
function setSizeOnRPr(doc: any, rPr: any, halfPts: number | null): boolean {
  if (!rPr) return false;
  let changed = false;
  for (const tag of ["w:sz", "w:szCs"]) {
    for (const el of directChildren(rPr, tag)) {
      rPr.removeChild(el);
      changed = true;
    }
  }
  if (halfPts !== null) {
    const sz = doc.createElement("w:sz");
    sz.setAttribute("w:val", String(halfPts));
    insertSizeChild(rPr, sz);
    const szCs = doc.createElement("w:szCs");
    szCs.setAttribute("w:val", String(halfPts));
    insertSizeChild(rPr, szCs);
    changed = true;
  }
  return changed;
}

/** True when an element has at least one ELEMENT child (text/attrs don't count). */
function hasElementChild(el: any): boolean {
  const kids = el.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k && k.nodeType === ELEMENT_NODE) return true;
  }
  return false;
}

/** Remove `el` from its parent when it has become an empty element (keeps round-trips clean). */
function removeIfEmpty(el: any): void {
  if (el && el.parentNode && !hasElementChild(el)) el.parentNode.removeChild(el);
}

/** Set/clear a run's explicit size (creating its rPr only when setting a value). */
function setRunSize(doc: any, runEl: any, halfPts: number | null): boolean {
  const rPr = halfPts === null ? runRPr(runEl) : ensureRPr(doc, runEl);
  const changed = setSizeOnRPr(doc, rPr, halfPts);
  // Don't leave a bare `<w:rPr/>` behind when clearing emptied it (cleanliness + idempotency).
  if (halfPts === null) removeIfEmpty(rPr);
  return changed;
}

/** The paragraph-mark run properties (`<w:pPr><w:rPr>`), creating both if absent (mark rPr is last in pPr). */
function ensureMarkRPr(doc: any, pEl: any): any {
  let pPr = firstDirectChild(pEl, "w:pPr");
  if (!pPr) {
    pPr = doc.createElement("w:pPr");
    pEl.insertBefore(pPr, pEl.firstChild);
  }
  const existing = firstDirectChild(pPr, "w:rPr");
  if (existing) return existing;
  const rPr = doc.createElement("w:rPr");
  pPr.appendChild(rPr);
  return rPr;
}

/** The existing paragraph-mark rPr without creating anything (for clearing), or null. */
function existingMarkRPr(pEl: any): any | null {
  const pPr = firstDirectChild(pEl, "w:pPr");
  return pPr ? firstDirectChild(pPr, "w:rPr") : null;
}

// ---------------------------------------------------------------------------
// Reads
// ---------------------------------------------------------------------------

/** A character style's resolved emphasis: whether it ultimately underlines and/or character-boxes its text. */
export interface StyleEmphasis {
  underline: boolean;
  boxed: boolean;
}

/** Tri-state underline parsed from a style's `<w:rPr>` XML (true / explicit-off / undefined=absent). */
function ownUnderlineTri(rPrXml: string): boolean | undefined {
  const m = /<w:u\b([^>]*?)\/?>/.exec(rPrXml);
  if (!m) return undefined;
  const val = ((/\bw:val="([^"]*)"/.exec(m[1]) || [])[1] || "").toLowerCase();
  return !(val === "none" || val === "0" || val === "false");
}

/** Tri-state character box parsed from a style's `<w:rPr>` XML (true / explicit-off / undefined=absent). */
function ownBoxTri(rPrXml: string): boolean | undefined {
  const m = /<w:bdr\b([^>]*?)\/?>/.exec(rPrXml);
  if (!m) return undefined;
  const val = ((/\bw:val="([^"]*)"/.exec(m[1]) || [])[1] || "").toLowerCase();
  return !(val === "nil" || val === "none" || val === "");
}

/**
 * Map each style id to its RESOLVED emphasis (underline / character-box), following the `basedOn`
 * cascade, parsed from the fragment's bundled styles.xml. The real keep-signal in docs is applied
 * through a character STYLE (StyleUnderline, Emphasis, …), so Shrink must resolve the STRUCTURAL signal
 * through the style — never by hardcoding style NAMES (lesson #3). Regex scan (like
 * `resolveNormalSizeHalfPts` / outline.ts) — styles.xml is flat and well-formed as Word emits it. A bare
 * `<w:p>` fixture with no styles part yields an empty map (a run's direct rPr still resolves).
 */
export function resolveStyleEmphasis(fragmentXml: string): Map<string, StyleEmphasis> {
  // 1) Each style's OWN tri-state underline/box + its basedOn parent.
  interface RawStyle {
    basedOn: string | null;
    ownU: boolean | undefined;
    ownBox: boolean | undefined;
  }
  const raw = new Map<string, RawStyle>();
  const styleRe = /<w:style\b([^>]*)>([\s\S]*?)<\/w:style>/g;
  let m: RegExpExecArray | null;
  while ((m = styleRe.exec(fragmentXml)) !== null) {
    const idMatch = /\bw:styleId="([^"]*)"/.exec(m[1]);
    if (!idMatch) continue;
    const body = m[2];
    const basedMatch = /<w:basedOn\b[^>]*\bw:val="([^"]*)"/.exec(body);
    // Only the style's own <w:rPr> carries character formatting (skip a paragraph style's pPr borders).
    const rPrMatch = /<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(body);
    const rPr = rPrMatch ? rPrMatch[1] : "";
    raw.set(idMatch[1], { basedOn: basedMatch ? basedMatch[1] : null, ownU: ownUnderlineTri(rPr), ownBox: ownBoxTri(rPr) });
  }
  // 2) Resolve each style through its basedOn chain (memoized; cycle-guarded).
  const resolved = new Map<string, StyleEmphasis>();
  const resolveOne = (id: string, seen: Set<string>): StyleEmphasis => {
    const cached = resolved.get(id);
    if (cached) return cached;
    const r = raw.get(id);
    if (!r || seen.has(id)) return { underline: false, boxed: false };
    seen.add(id);
    const base = r.basedOn ? resolveOne(r.basedOn, seen) : { underline: false, boxed: false };
    const out: StyleEmphasis = {
      underline: r.ownU !== undefined ? r.ownU : base.underline,
      boxed: r.ownBox !== undefined ? r.ownBox : base.boxed
    };
    resolved.set(id, out);
    return out;
  };
  for (const id of raw.keys()) resolveOne(id, new Set());
  return resolved;
}

/**
 * One run's read-only view for the pure engines. `underline`/`boxed` are resolved DIRECT-then-STYLE: a
 * direct `<w:u>`/`<w:bdr>` on the run wins (tri-state, so an explicit none overrides), otherwise the
 * run's character style's resolved emphasis (from `styleMap`) decides — because in real docs the cut is
 * underlined/boxed through a character STYLE (StyleUnderline / Emphasis), not a direct rPr.
 */
function readRun(runEl: any, index: number, styleMap: Map<string, StyleEmphasis>): FragmentRunView {
  const rStyle = runRStyle(runEl);
  const styleEmph = rStyle ? styleMap.get(rStyle) : undefined;
  const directU = directUnderlineTri(runEl);
  const directBox = directBoxTri(runEl);
  return {
    index,
    text: runTextView(runEl),
    highlight: runHighlight(runEl),
    citeStyled: rStyle === CITE_STYLE_ID,
    underline: directU !== undefined ? directU : styleEmph ? styleEmph.underline : false,
    boxed: directBox !== undefined ? directBox : styleEmph ? styleEmph.boxed : false,
    hidden: vanishOn(runRPr(runEl)),
    eligible: runEligible(runEl),
    hasInternalPart: runHasInternalPart(runEl),
    sizeHalfPts: runSizeHalfPts(runEl),
    breakMarker: isMarkerRun(runEl)
  };
}

/**
 * Read every STORY paragraph in the fragment as an array of run views (document order). The Shrink
 * engine consumes this to decide per-run sizes; the per-paragraph indexing is what `applyFragmentShrink`
 * writes back against. An empty fragment (no `<w:p>`) reads as `[]`. Pass `parsed` (from
 * `parseFragment`) to reuse an existing tree instead of re-parsing `fragmentXml`; reading never
 * mutates it, so the handle stays valid for a later `applyFragmentShrink` over the same string.
 */
export function readFragmentParagraphs(fragmentXml: string, parsed?: ParsedFragment): FragmentRunView[][] {
  const doc = parsed ? parsed.doc : parse(fragmentXml);
  // The keep-signal (underline / box) is usually applied through a character style, so resolve the
  // fragment's bundled styles.xml ONCE and thread the map down to each run view.
  const styleMap = resolveStyleEmphasis(fragmentXml);
  const paras = storyParagraphs(bodyScope(doc));
  return paras.map((p) => paragraphRuns(p).map((r, i) => readRun(r, i, styleMap)));
}

/**
 * The Normal/default font size (half-points) the fragment's `styles.xml` resolves to: `docDefaults`
 * `<w:rPrDefault>` first, then the `Normal` style's own `<w:sz>`. Returns null when the fragment
 * carries no styles part (bare-`<w:p>` test fixtures) so the caller can fall back to a sensible default.
 * Regex scan (like outline.ts) — `styles.xml` is flat and well-formed as Word emits it.
 */
export function resolveNormalSizeHalfPts(fragmentXml: string): number | null {
  const dd = /<w:docDefaults>[\s\S]*?<w:rPrDefault>[\s\S]*?<w:rPr>([\s\S]*?)<\/w:rPr>/.exec(fragmentXml);
  if (dd) {
    const m = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(dd[1]);
    if (m) return Number(m[1]);
  }
  const normal = /<w:style\b[^>]*\bw:styleId="Normal"[^>]*>([\s\S]*?)<\/w:style>/.exec(fragmentXml);
  if (normal) {
    const m = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(normal[1]);
    if (m) return Number(m[1]);
  }
  return null;
}

// ---------------------------------------------------------------------------
// Shrink apply
// ---------------------------------------------------------------------------

/** The size instructions for ONE paragraph, produced by the pure Shrink engine. */
export interface ParagraphShrinkPlan {
  /**
   * New explicit size (half-points) per run, aligned to `readFragmentParagraphs` order:
   * a number SETS the size, `null` CLEARS it (→ inherited Normal), `undefined` LEAVES it unchanged.
   */
  runSizes: (number | null | undefined)[];
  /**
   * The paragraph-mark size: a number SETS it (Verbatim's "Shrink ¶" → 6pt), `null` CLEARS it (Unshrink
   * reverting the mark to its inherited size), `undefined` LEAVES it unchanged.
   */
  markSizeHalfPts?: number | null;
}

/**
 * Apply per-run (and optional paragraph-mark) sizes to the fragment. `plans[i]` aligns to the i-th story
 * paragraph from `readFragmentParagraphs`. Pure: returns the new XML and whether anything changed
 * (computed by comparing the serialized tree before and after, so xmldom's own normalization never
 * reports a spurious change — keeping Shrink idempotency assertions honest). Pass `parsed` (from
 * `parseFragment`) to reuse an existing tree instead of re-parsing `fragmentXml`; this MUTATES
 * (consumes) the handle — the tree is edited in place, so the caller must not use it afterwards.
 */
export function applyFragmentShrink(
  fragmentXml: string,
  plans: ParagraphShrinkPlan[],
  parsed?: ParsedFragment
): { xml: string; changed: boolean } {
  const doc = parsed ? parsed.doc : parse(fragmentXml);
  const before = serialize(doc);
  const paras = storyParagraphs(bodyScope(doc));
  for (let p = 0; p < paras.length && p < plans.length; p++) {
    const plan = plans[p];
    const runs = paragraphRuns(paras[p]);
    for (let i = 0; i < runs.length && i < plan.runSizes.length; i++) {
      const target = plan.runSizes[i];
      if (target === undefined) continue;
      setRunSize(doc, runs[i], target);
    }
    if (plan.markSizeHalfPts !== undefined) {
      // undefined = leave; number = set; null = clear. Only materialize the mark rPr when setting.
      const markRPr = plan.markSizeHalfPts === null ? existingMarkRPr(paras[p]) : ensureMarkRPr(doc, paras[p]);
      setSizeOnRPr(doc, markRPr, plan.markSizeHalfPts);
      if (plan.markSizeHalfPts === null && markRPr) {
        removeIfEmpty(markRPr); // don't leave a bare mark `<w:rPr/>` after Unshrink clears it
        removeIfEmpty(firstDirectChild(paras[p], "w:pPr"));
      }
    }
  }
  const after = serialize(doc);
  return { xml: after, changed: before !== after };
}

// ---------------------------------------------------------------------------
// Whitespace collapse (shared by both Condense modes)
// ---------------------------------------------------------------------------

/**
 * Collapse whitespace inside ONE paragraph in place (Verbatim's `^t/^l/^n` + double-space passes):
 * every `<w:tab>`/`<w:br>`/`<w:cr>` and Space_Separator char becomes one ASCII space, runs of spaces
 * collapse to one — ACROSS run boundaries — and a simple run left empty is dropped. Structural runs
 * (fields, drawings, anything not plain text) are left untouched and reset the cross-run space state so
 * a space is never deleted across them. Returns true when the paragraph changed.
 */
function collapseParagraphWhitespace(doc: any, pEl: any): boolean {
  const runs = paragraphRuns(pEl);
  let changed = false;
  let prevEndedWithSpace = true; // leading spaces of the FIRST run collapse away (paragraph start)
  for (const runEl of runs) {
    // A break marker (a space/¶ glyph, or a vanished pPr-payload run) is a structural BOUNDARY, not
    // collapsible text — if we let its space collapse away on a re-condense it would silently destroy
    // the paragraph break it stands for (the adversarial "marker-eating" defect). Treat it like a field:
    // skip it and reset the cross-run space state so neither it nor the spaces around it are merged.
    if (isMarkerRun(runEl)) {
      prevEndedWithSpace = false;
      continue;
    }
    if (!isSimpleTextRun(runEl)) {
      prevEndedWithSpace = false; // structural content: never collapse a space across it
      continue;
    }
    const raw = runTextView(runEl); // tab/break rendered as \n; all whitespace handled uniformly below
    let normalized = raw.replace(WHITESPACE_RUN, " ");
    if (prevEndedWithSpace) normalized = normalized.replace(/^ +/, "");
    if (normalized !== raw) changed = true;
    if (normalized === "") {
      // The whole run collapsed away — drop it (keeps merged text clean, no empty <w:t>).
      if (runEl.parentNode) {
        runEl.parentNode.removeChild(runEl);
        changed = true;
      }
      // prevEndedWithSpace unchanged: a removed pure-whitespace run leaves the prior state intact.
      continue;
    }
    rebuildSimpleRunText(doc, runEl, normalized);
    prevEndedWithSpace = normalized.endsWith(" ");
  }
  return changed;
}

/** Replace a simple run's text/tab/break children with a single `<w:t xml:space="preserve">text</w:t>`. */
function rebuildSimpleRunText(doc: any, runEl: any, text: string): void {
  const toRemove: any[] = [];
  const kids = runEl.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (!(k.nodeType === ELEMENT_NODE && k.nodeName === "w:rPr")) toRemove.push(k);
  }
  for (const k of toRemove) runEl.removeChild(k);
  const t = doc.createElement("w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(text));
  runEl.appendChild(t);
}

// ---------------------------------------------------------------------------
// Marker construction
// ---------------------------------------------------------------------------

/**
 * Build a visible boundary glyph run whose text is `SIGNATURE + glyph` (glyph = a space, or a 6pt `¶`).
 * The intrinsic signature — NOT a style reference — is what makes it a marker, because run TEXT survives
 * `insertOoxml` where a net-new custom `<w:rStyle>` does not (see {@link MARK_SIGNATURE}). The pilcrow's
 * 6pt size is direct formatting (`<w:sz>`/`<w:szCs>`), which also survives. A space marker needs no rPr.
 */
function makeGlyphMarker(doc: any, glyph: string, sizeHalfPts: number | null): any {
  const r = doc.createElement("w:r");
  if (sizeHalfPts != null) {
    const rPr = doc.createElement("w:rPr");
    const sz = doc.createElement("w:sz");
    sz.setAttribute("w:val", String(sizeHalfPts));
    rPr.appendChild(sz);
    const szCs = doc.createElement("w:szCs");
    szCs.setAttribute("w:val", String(sizeHalfPts));
    rPr.appendChild(szCs);
    r.appendChild(rPr);
  }
  const t = doc.createElement("w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(MARK_SIGNATURE + glyph));
  r.appendChild(t);
  return r;
}

/**
 * Build a hidden, vanished marker run whose `<w:t>` carries `SIGNATURE + payloadXml` — either a divergent
 * following paragraph's `<w:pPr>` (merge mode) or a dropped blank's original mark `<w:rPr>` (retain mode).
 * Stored as TEXT: the serializer escapes it (`<`->`&lt;`), Uncondense strips the signature, un-escapes and
 * re-parses. Self-describing (the signature + leading tag identify it), so it travels with the paragraph
 * (copy/paste-safe; no global state, no style table). `<w:vanish/>` is direct formatting that both hides
 * the payload AND keeps the run from coalescing into an adjacent VISIBLE glyph/body run on import.
 */
function makeDataRun(doc: any, payloadXml: string): any {
  const r = doc.createElement("w:r");
  const rPr = doc.createElement("w:rPr");
  rPr.appendChild(doc.createElement("w:vanish"));
  r.appendChild(rPr);
  const t = doc.createElement("w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(MARK_SIGNATURE + payloadXml));
  r.appendChild(t);
  return r;
}

// ---------------------------------------------------------------------------
// Condense
// ---------------------------------------------------------------------------

export interface CondenseOoxmlResult {
  xml: string;
  changed: boolean;
  paragraphsScanned: number;
  boundariesMarked: number;
}

/**
 * Condense the fragment's top-level paragraphs per `opts`. Always collapses intra-paragraph whitespace.
 * Then EITHER merges every paragraph into the first (inserting a reversible boundary marker at each
 * former break — a visible space, or a `¶` at 6pt for pilcrow mode, or a plain space for destructive
 * `reversal:"none"`) OR, in retain-paragraphs mode, drops blank/whitespace-only paragraphs (marking
 * each removed one's paragraph mark so Uncondense can restore it, unless destructive).
 *
 * A divergent following-paragraph `<w:pPr>` is preserved in a hidden payload run after its boundary
 * marker, so Uncondense restores paragraph properties losslessly; uniform card bodies (the common case)
 * store no payload and are exact by construction.
 */
export function condenseFragmentOoxml(fragmentXml: string, opts: CondenseOptions): CondenseOoxmlResult {
  const doc = parse(fragmentXml);
  const before = serialize(doc);
  const scope = bodyScope(doc);
  const paras = directParagraphs(scope);

  for (const p of paras) collapseParagraphWhitespace(doc, p);

  let boundariesMarked = 0;
  if (opts.retainParagraphs) {
    boundariesMarked = dropBlankParagraphs(doc, paras, opts.reversal);
  } else {
    boundariesMarked = mergeParagraphs(doc, paras, opts);
  }

  // No styles-part injection: markers self-identify through the intrinsic text {@link MARK_SIGNATURE}
  // + direct run formatting, both of which survive `insertOoxml` into a populated doc — unlike the
  // retired custom `<w:rStyle>`, which Word stripped on import (the 2026-06-09 irreversibility bug).

  const after = serialize(doc);
  return { xml: after, changed: before !== after, paragraphsScanned: paras.length, boundariesMarked };
}

/** The serialized `<w:pPr>` of a paragraph, or "" when it has none (for divergence comparison + storage). */
function paragraphPPrXml(pEl: any): string {
  const pPr = firstDirectChild(pEl, "w:pPr");
  return pPr ? serialize(pPr) : "";
}

/** Every child of `<w:p>` that is NOT its `<w:pPr>` (the run-level content to move on merge). */
function paragraphContentNodes(pEl: any): any[] {
  const out: any[] = [];
  const kids = pEl.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item ? kids.item(i) : kids[i];
    if (k.nodeName !== "w:pPr") out.push(k);
  }
  return out;
}

/**
 * Merge `paras[1..]` into `paras[0]`. Between each former paragraph, insert a boundary marker (and, when
 * the following paragraph's pPr differs from the first's, a hidden payload run carrying it). The merged
 * paragraph keeps the first paragraph's pPr. Returns the number of boundary markers inserted.
 */
function mergeParagraphs(doc: any, paras: any[], opts: CondenseOptions): number {
  if (paras.length < 2) return 0;
  const target = paras[0];
  const basePPr = paragraphPPrXml(target);
  const useMarker = !(opts.reversal === "none" && !opts.usePilcrows); // destructive only when pilcrows off
  const glyph = opts.usePilcrows ? "¶" : " ";
  const glyphSize = opts.usePilcrows ? SIX_PT_HALF : null;
  let boundaries = 0;

  for (let k = 1; k < paras.length; k++) {
    const src = paras[k];
    // 1) Boundary marker (or a plain space for the destructive escape hatch).
    if (useMarker) {
      target.appendChild(makeGlyphMarker(doc, glyph, glyphSize));
      const srcPPr = paragraphPPrXml(src);
      // Store a payload whenever the source paragraph's pPr DIFFERS from the merged (first) paragraph's
      // — including the case where the source had NO pPr but the base does, which must NOT silently
      // inherit the base's pPr on split (the adversarial "wrong-paragraph pPr" defect). "No pPr" is
      // stored as an explicit empty `<w:pPr/>` sentinel so the split restores a default paragraph, not
      // a clone of the base. When src == base (the uniform card-body common case) no payload is stored
      // and the split clones the base pPr — exact by construction.
      if (srcPPr !== basePPr) {
        target.appendChild(makeDataRun(doc, srcPPr || "<w:pPr/>"));
      }
      boundaries++;
    } else {
      const sp = doc.createElement("w:r");
      const t = doc.createElement("w:t");
      t.setAttribute("xml:space", "preserve");
      t.appendChild(doc.createTextNode(" "));
      sp.appendChild(t);
      target.appendChild(sp);
    }
    // 2) Move the source paragraph's run-level content into the target, in order.
    for (const node of paragraphContentNodes(src)) target.appendChild(node);
    // 3) Remove the now-empty source paragraph.
    if (src.parentNode) src.parentNode.removeChild(src);
  }
  return boundaries;
}

/**
 * Retain-paragraphs mode: drop blank / whitespace-only paragraphs. Lossless (`reversal:"marker"`) marks
 * each removed paragraph's mark with the break style + `<w:vanish/>` so Uncondense can bring it back;
 * destructive (`reversal:"none"`) deletes the node. Never removes the only remaining paragraph (a range
 * must end with one). Returns how many paragraphs were dropped/marked.
 */
function dropBlankParagraphs(doc: any, paras: any[], reversal: CondenseOptions["reversal"]): number {
  let count = 0;
  // `paras` is exactly the body's direct paragraphs (see condenseFragmentOoxml), so a destructive
  // removal shrinks that count by one, and marker-mode (vanish-stamp) never removes a node at all.
  // Tracking the count live keeps the same guard semantics while replacing a full-document
  // bodyScope+directParagraphs re-scan per blank — quadratic on selections with many blanks.
  let remaining = directParagraphs(bodyScope(doc)).length;
  for (const p of paras) {
    if (!isBlankParagraph(p)) continue;
    if (remaining <= 1) break; // keep at least one paragraph
    if (reversal === "marker") {
      // Lossless removal: PARK the pristine paragraph-mark rPr in a hidden, self-describing payload run,
      // then add `<w:vanish/>` to the live mark so the blank line collapses now. Uncondense detects the
      // dropped blank by the parked mark-rPr payload and restores the user's mark VERBATIM from it — one
      // uniform path whether the mark had no rPr, a plain rPr, or a foreign character style (real docs
      // underline an empty newline through a char STYLE). The mark carries no break-style reference at
      // all, so nothing depends on a custom style surviving the host round-trip.
      //
      // Capture the original BEFORE mutating: serialize the existing mark rPr (or the empty sentinel
      // `<w:rPr/>` when the mark had none) so the parked copy never includes the `<w:vanish/>` we add next.
      const existing = existingMarkRPr(p);
      p.appendChild(makeDataRun(doc, existing ? serialize(existing) : "<w:rPr/>"));
      const markRPr = ensureMarkRPr(doc, p);
      if (!firstDirectChild(markRPr, "w:vanish")) markRPr.appendChild(doc.createElement("w:vanish"));
      count++;
    } else if (p.parentNode) {
      p.parentNode.removeChild(p);
      remaining--;
      count++;
    }
  }
  return count;
}

/** True when a paragraph has no visible text (only whitespace, or nothing but its mark). */
function isBlankParagraph(pEl: any): boolean {
  for (const r of paragraphRuns(pEl)) {
    if (!isSimpleTextRun(r)) return false; // a field/drawing run is content — don't drop it
    const t = runTextView(r);
    if (t.length > 0 && !/^\s*$/u.test(t)) return false;
  }
  return true;
}

// ---------------------------------------------------------------------------
// Uncondense (the exact inverse of the marker modes)
// ---------------------------------------------------------------------------

export interface UncondenseOoxmlResult {
  xml: string;
  changed: boolean;
  breaksRestored: number;
}

/**
 * Reverse Condense: every boundary marker becomes a paragraph break again (restoring stored pPr), and
 * every retain-mode hidden blank-paragraph mark is un-hidden. Operates on the fragment's top-level
 * paragraphs. Returns the number of breaks restored.
 */
export function uncondenseFragmentOoxml(fragmentXml: string): UncondenseOoxmlResult {
  const doc = parse(fragmentXml);
  const before = serialize(doc);
  const scope = bodyScope(doc);
  let breaks = 0;

  // 1) Retain-mode: restore every blank we dropped, detected by its parked mark-rPr payload run.
  for (const p of directParagraphs(scope)) {
    const dataRun = paragraphRuns(p).find((r) => isMarkRPrDataRun(r));
    if (!dataRun) continue;
    // Restore the user's paragraph mark VERBATIM from the parked payload (signature already stripped by
    // `payloadText`). An empty `<w:rPr/>` payload means the blank's mark had no own properties → drop the
    // live mark rPr entirely (which also removes the `<w:vanish/>` we added to collapse it).
    const original = importRPr(doc, payloadText(dataRun));
    if (!original) continue; // payload unparseable (should never happen) — leave it parked, never lose it
    if (dataRun.parentNode) dataRun.parentNode.removeChild(dataRun);
    const pPr = firstDirectChild(p, "w:pPr");
    const liveMarkRPr = pPr ? firstDirectChild(pPr, "w:rPr") : null;
    if (original && hasElementChild(original)) {
      if (liveMarkRPr) pPr.replaceChild(original, liveMarkRPr);
      else if (pPr) pPr.appendChild(original);
    } else if (liveMarkRPr && liveMarkRPr.parentNode) {
      liveMarkRPr.parentNode.removeChild(liveMarkRPr);
    }
    // Leave no empty `<w:pPr/>` residue behind (keeps the retain round-trip clean).
    removeIfEmpty(pPr);
    breaks++;
  }

  // 2) Merge-mode: split each paragraph that contains boundary glyph markers.
  for (const p of directParagraphs(scope)) {
    breaks += splitParagraphAtMarkers(doc, p);
  }

  const after = serialize(doc);
  return { xml: after, changed: before !== after, breaksRestored: breaks };
}

/**
 * Non-run paragraph children Word may interleave between a boundary glyph run and its pPr payload run
 * on an import/export round-trip (proofing marks, bookmark edges). The payload scan steps over these —
 * requiring strict adjacency would silently drop the stored pPr (a real-Word formatting-loss bug).
 */
const IGNORABLE_BETWEEN_MARKER_AND_PAYLOAD = new Set<string>(["w:proofErr", "w:bookmarkStart", "w:bookmarkEnd"]);

/**
 * Split a merged paragraph at its boundary markers into the original paragraphs, in document order.
 * Segment 0 inherits the merged paragraph's own pPr; each later segment gets the pPr stored in the
 * boundary's payload run, or a clone of the merged pPr (the uniform-card-body common case). The markers
 * and payload runs are dropped. Returns the number of breaks restored (segments created − 1).
 *
 * PRECISE PER SIGNATURE OCCURRENCE, not whole-run-atomic, because the marker signal lives IN run
 * text and Word coalesces adjacent identically-formatted runs on `insertOoxml`. Two consequences this
 * must survive: (a) two consecutive boundaries (an empty former paragraph between cards) merge into ONE
 * run `SIGNATURE glyph SIGNATURE glyph` → still TWO breaks; (b) a glyph marker can merge with adjacent
 * BODY text of identical formatting → the run text becomes `body SIGNATURE glyph body`. Tokenizing on
 * the signature set (text before the first = body for the current segment; each signature = one
 * boundary; the ONE glyph char after it is dropped; the remainder is body for the new segment,
 * re-emitted with the marker run's own rPr — which equals the body's, since identical rPr is exactly
 * what let them coalesce) makes the split lossless under ANY coalescing. A signature NOT followed by a
 * boundary glyph is treated as literal text (kept verbatim as its exact original sequence), so a stray
 * signature can never eat a character — and a lone organic ZWSP never tokenizes at all.
 */
function splitParagraphAtMarkers(doc: any, pEl: any): number {
  const content = paragraphContentNodes(pEl);
  if (!content.some((n) => n.nodeType === ELEMENT_NODE && n.nodeName === "w:r" && isGlyphMarkerRun(n))) {
    return 0;
  }

  const basePPr = firstDirectChild(pEl, "w:pPr");
  const segments: any[][] = [[]];
  const segmentPPr: (string | null)[] = [null]; // serialized pPr per segment (null → clone basePPr)

  /** Append a body run (cloning the marker run's rPr, so coalesced-in body keeps its formatting). */
  const pushBody = (rPrNode: any, text: string): void => {
    if (text === "") return;
    const r = doc.createElement("w:r");
    if (rPrNode) r.appendChild(rPrNode.cloneNode(true));
    const t = doc.createElement("w:t");
    t.setAttribute("xml:space", "preserve");
    t.appendChild(doc.createTextNode(text));
    r.appendChild(t);
    segments[segments.length - 1].push(r);
  };

  for (let i = 0; i < content.length; i++) {
    const node = content[i];
    if (node.nodeType === ELEMENT_NODE && node.nodeName === "w:r" && isGlyphMarkerRun(node)) {
      const rPrNode = runRPr(node);
      // Capture-split on the SIGNATURE SET: tokens = [body0, sig1, part1, sig2, part2, …]. A single
      // run can interleave BOTH signature kinds (re-condensing a legacy-condensed doc writes new-pair
      // boundaries beside surviving U+2063 ones, and Word coalesces them into one run), so the split
      // must honor every kind — and the captured separator lets a NON-boundary signature be re-emitted
      // verbatim as exactly the sequence it was (never normalized to the current written form). The
      // split consumes each signature's FULL length (1 or 2 code units); the glyph after it is always
      // exactly one code unit, dropped via `.slice(1)`.
      const tokens = runTextRaw(node).split(SIGNATURE_SPLIT);
      pushBody(rPrNode, tokens[0]); // text before the first signature stays in the current segment
      let boundariesHere = 0;
      for (let k = 1; k < tokens.length; k += 2) {
        const signature = tokens[k]; // the actual signature sequence found (current pair OR legacy)
        const part = tokens[k + 1] !== undefined ? tokens[k + 1] : "";
        const glyph = part.length > 0 ? part[0] : "";
        if (BOUNDARY_GLYPHS.has(glyph)) {
          // Real boundary: open a new segment, drop the one glyph char, keep any trailing body.
          segments.push([]);
          segmentPPr.push(null);
          boundariesHere++;
          pushBody(rPrNode, part.slice(1));
        } else {
          // Signature not followed by a known glyph (collision / host re-split): keep it verbatim so no
          // character is ever lost; do NOT open a boundary here.
          pushBody(rPrNode, signature + part);
        }
      }
      if (boundariesHere === 0) continue; // nothing split off this run — no payload to consume
      // The pPr payload follows its glyph run, possibly past Word-inserted noise. It belongs to the
      // run's LAST boundary: merged glyphs never had a payload between them — a payload run's distinct
      // (vanished) formatting would have prevented the adjacency that makes Word merge in the first place.
      let j = i + 1;
      const skipped: any[] = [];
      while (
        j < content.length &&
        content[j].nodeType === ELEMENT_NODE &&
        IGNORABLE_BETWEEN_MARKER_AND_PAYLOAD.has(content[j].nodeName)
      ) {
        skipped.push(content[j]);
        j++;
      }
      const next = content[j];
      if (next && next.nodeType === ELEMENT_NODE && next.nodeName === "w:r" && isPPrDataRun(next)) {
        segmentPPr[segmentPPr.length - 1] = payloadText(next); // signature-stripped pPr XML
        // Keep the stepped-over noise (bookmark edges matter); it lands in the boundary's new segment.
        for (const s of skipped) segments[segments.length - 1].push(s);
        i = j; // consume through the payload run
      }
      // No payload found: don't advance i — the scanned nodes re-enter the loop and are kept normally.
    } else if (node.nodeType === ELEMENT_NODE && node.nodeName === "w:r" && isPPrDataRun(node)) {
      continue; // stray payload run with no preceding glyph (shouldn't happen) — drop it
    } else {
      segments[segments.length - 1].push(node);
    }
  }
  if (segments.length < 2) return 0;

  const parent = pEl.parentNode;
  // The existing paragraph becomes segment 0: clear its run-level content, re-append segment 0.
  for (const node of paragraphContentNodes(pEl)) pEl.removeChild(node);
  for (const node of segments[0]) pEl.appendChild(node);

  let prev = pEl;
  for (let s = 1; s < segments.length; s++) {
    const np = doc.createElement("w:p");
    const pPrXml = segmentPPr[s];
    if (pPrXml) {
      const imported = importPPr(doc, pPrXml);
      if (imported) np.appendChild(imported);
    } else if (basePPr) {
      np.appendChild(basePPr.cloneNode(true));
    }
    for (const node of segments[s]) np.appendChild(node);
    parent.insertBefore(np, prev.nextSibling);
    prev = np;
  }
  return segments.length - 1;
}

/** Strip any stray signature (every kind) before re-parsing a payload (callers pass `payloadText`, but be defensive). */
function stripSignatures(xml: string): string {
  return xml.replace(ALL_SIGNATURES_G, "");
}

/** Parse a stored `<w:pPr>` string and import it into `doc` (namespace-wrapped so every prefix resolves). */
function importPPr(doc: any, pPrXml: string): any | null {
  try {
    const wrapped = `<w:p ${WRAP_NS}>${stripSignatures(pPrXml)}</w:p>`;
    const frag = parse(wrapped);
    const pPr = frag.getElementsByTagName("w:pPr").item(0);
    return pPr ? doc.importNode(pPr, true) : null;
  } catch {
    return null;
  }
}

/** Parse a stored mark `<w:rPr>` string and import it into `doc` (namespace-wrapped so every prefix resolves). */
function importRPr(doc: any, rPrXml: string): any | null {
  try {
    const wrapped = `<w:p ${WRAP_NS}><w:pPr>${stripSignatures(rPrXml)}</w:pPr></w:p>`;
    const frag = parse(wrapped);
    const rPr = frag.getElementsByTagName("w:rPr").item(0);
    return rPr ? doc.importNode(rPr, true) : null;
  } catch {
    return null;
  }
}
