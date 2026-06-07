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
import { CONDENSE_MARK_STYLE, CITE_STYLE_ID } from "./styles";
import { CondenseOptions, FragmentRunView } from "./types";

// xmldom node types vary across versions; we keep public signatures fully typed (string in/out,
// number/null) and use a localized `any` for node handles — the same pragmatic choice ooxml.ts and
// ooxmlPackage.ts make for a version-robust XML adapter.
/* eslint-disable @typescript-eslint/no-explicit-any */

const ELEMENT_NODE = 1;
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

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

/** Run text rendered for a view: `<w:t>` verbatim, `<w:tab>` -> \t, `<w:br>`/`<w:cr>` -> \n. */
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
  return text;
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

/** True when a run is plain text (only rPr / w:t / w:tab / w:br / w:cr) — safe to normalize whitespace. */
function isSimpleTextRun(runEl: any): boolean {
  const kids = runEl.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType !== ELEMENT_NODE) continue;
    const n = k.nodeName;
    if (n !== "w:rPr" && n !== "w:t" && n !== "w:tab" && n !== "w:br" && n !== "w:cr") return false;
  }
  return true;
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

/** A run carrying the condense break-marker character style. */
function isMarkerRun(runEl: any): boolean {
  return runRStyle(runEl) === CONDENSE_MARK_STYLE;
}

/** A marker run that stores a divergent following-paragraph `<w:pPr>` (vanished; text begins `<w:pPr`). */
function isPPrDataRun(runEl: any): boolean {
  return isMarkerRun(runEl) && vanishOn(runRPr(runEl)) && runTextRaw(runEl).trimStart().startsWith("<w:pPr");
}

/**
 * A marker run that stores a retain-mode dropped blank's ORIGINAL mark `<w:rPr>` (vanished; text begins
 * `<w:rPr`). Parked when the dropped paragraph's mark carried a foreign character style we had to swap for
 * our break style, so Uncondense can restore the user's mark exactly.
 */
function isMarkRPrDataRun(runEl: any): boolean {
  return isMarkerRun(runEl) && vanishOn(runRPr(runEl)) && runTextRaw(runEl).trimStart().startsWith("<w:rPr");
}

/** A marker run carrying a serialized-XML payload (a stored pPr or mark rPr) rather than a visible glyph. */
function isDataRun(runEl: any): boolean {
  return isPPrDataRun(runEl) || isMarkRPrDataRun(runEl);
}

/** A marker run that is the VISIBLE boundary glyph (a space or `¶`) — the actual paragraph break. */
function isGlyphMarkerRun(runEl: any): boolean {
  return isMarkerRun(runEl) && !isDataRun(runEl);
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
 * cascade, parsed from the fragment's bundled styles.xml. The real keep-signal in briefs is applied
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
 * run's character style's resolved emphasis (from `styleMap`) decides — because in real briefs the cut is
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
    sizeHalfPts: runSizeHalfPts(runEl),
    breakMarker: isMarkerRun(runEl)
  };
}

/**
 * Read every STORY paragraph in the fragment as an array of run views (document order). The Shrink
 * engine consumes this to decide per-run sizes; the per-paragraph indexing is what `applyFragmentShrink`
 * writes back against. An empty fragment (no `<w:p>`) reads as `[]`.
 */
export function readFragmentParagraphs(fragmentXml: string): FragmentRunView[][] {
  const doc = parse(fragmentXml);
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
 * reports a spurious change — keeping Shrink idempotency assertions honest).
 */
export function applyFragmentShrink(
  fragmentXml: string,
  plans: ParagraphShrinkPlan[]
): { xml: string; changed: boolean } {
  const doc = parse(fragmentXml);
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

/** Build a visible boundary glyph run (a space or `¶`), styled `CONDENSE_MARK_STYLE`. */
function makeGlyphMarker(doc: any, glyph: string, sizeHalfPts: number | null): any {
  const r = doc.createElement("w:r");
  const rPr = doc.createElement("w:rPr");
  const rStyle = doc.createElement("w:rStyle");
  rStyle.setAttribute("w:val", CONDENSE_MARK_STYLE);
  rPr.appendChild(rStyle);
  if (sizeHalfPts != null) {
    const sz = doc.createElement("w:sz");
    sz.setAttribute("w:val", String(sizeHalfPts));
    rPr.appendChild(sz);
    const szCs = doc.createElement("w:szCs");
    szCs.setAttribute("w:val", String(sizeHalfPts));
    rPr.appendChild(szCs);
  }
  r.appendChild(rPr);
  const t = doc.createElement("w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(glyph));
  r.appendChild(t);
  return r;
}

/**
 * Build a hidden, vanished, break-styled run whose `<w:t>` carries a serialized-XML payload — either a
 * divergent following paragraph's `<w:pPr>` (merge mode) or a dropped blank's original mark `<w:rPr>`
 * (retain mode). Stored as TEXT: the serializer escapes it (`<`->`&lt;`), Uncondense un-escapes and re-
 * parses. Self-describing (its own tag identifies it), so it travels with the paragraph (copy/paste-safe;
 * no global state).
 */
function makeDataRun(doc: any, payloadXml: string): any {
  const r = doc.createElement("w:r");
  const rPr = doc.createElement("w:rPr");
  const rStyle = doc.createElement("w:rStyle");
  rStyle.setAttribute("w:val", CONDENSE_MARK_STYLE);
  rPr.appendChild(rStyle);
  rPr.appendChild(doc.createElement("w:vanish"));
  r.appendChild(rPr);
  const t = doc.createElement("w:t");
  t.setAttribute("xml:space", "preserve");
  t.appendChild(doc.createTextNode(payloadXml));
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
  for (const p of paras) {
    if (!isBlankParagraph(p)) continue;
    if (directParagraphs(bodyScope(doc)).length <= 1) break; // keep at least one paragraph
    if (reversal === "marker") {
      // Lossless removal stamps the paragraph mark with the break style + `<w:vanish/>`, so it collapses
      // now and Uncondense (which keys on the break style) restores it. rPr allows only ONE `<w:rStyle>`,
      // so when the mark already carries a FOREIGN character style — e.g. an underlined-but-empty newline
      // styled `StyleUnderline` (real briefs underline via a char style) — we can't simply add ours. We
      // park the pristine original mark rPr in a hidden, self-describing payload run, THEN swap the live
      // rStyle to our break style: the blank condenses like any other, and Uncondense restores the user's
      // mark verbatim from the payload — lossless, where we previously had to skip such paragraphs.
      const existing = existingMarkRPr(p);
      const existingRStyle = existing ? firstDirectChild(existing, "w:rStyle") : null;
      const foreignStyle =
        !!existingRStyle && (existingRStyle.getAttribute("w:val") || "") !== CONDENSE_MARK_STYLE;
      if (foreignStyle) {
        p.appendChild(makeDataRun(doc, serialize(existing))); // park the original (captured pre-mutation)
        existingRStyle.setAttribute("w:val", CONDENSE_MARK_STYLE);
        if (!firstDirectChild(existing, "w:vanish")) existing.appendChild(doc.createElement("w:vanish"));
      } else {
        const markRPr = ensureMarkRPr(doc, p);
        if (!firstDirectChild(markRPr, "w:rStyle")) {
          const rStyle = doc.createElement("w:rStyle");
          rStyle.setAttribute("w:val", CONDENSE_MARK_STYLE);
          markRPr.insertBefore(rStyle, markRPr.firstChild);
        }
        if (!firstDirectChild(markRPr, "w:vanish")) markRPr.appendChild(doc.createElement("w:vanish"));
      }
      count++;
    } else if (p.parentNode) {
      p.parentNode.removeChild(p);
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

  // 1) Retain-mode: un-hide every paragraph mark stamped with the break style.
  for (const p of directParagraphs(scope)) {
    const pPr = firstDirectChild(p, "w:pPr");
    const markRPr = pPr ? firstDirectChild(pPr, "w:rPr") : null;
    const rStyle = markRPr ? firstDirectChild(markRPr, "w:rStyle") : null;
    if (rStyle && (rStyle.getAttribute("w:val") || "") === CONDENSE_MARK_STYLE) {
      // If we parked a foreign mark rPr when dropping this blank (see dropBlankParagraphs), restore it
      // verbatim from the payload run; otherwise just strip our break style + vanish back off the mark.
      const dataRun = paragraphRuns(p).find((r) => isMarkRPrDataRun(r));
      const original = dataRun ? importRPr(doc, runTextRaw(dataRun).trim()) : null;
      if (dataRun && dataRun.parentNode) dataRun.parentNode.removeChild(dataRun);
      if (original) {
        pPr.replaceChild(original, markRPr);
      } else {
        for (const v of directChildren(markRPr, "w:vanish")) markRPr.removeChild(v);
        markRPr.removeChild(rStyle);
        removeIfEmpty(markRPr);
      }
      // Leave no empty `<w:rPr/>`/`<w:pPr/>` residue behind (keeps the retain round-trip clean).
      removeIfEmpty(pPr);
      breaks++;
    }
  }

  // 2) Merge-mode: split each paragraph that contains boundary glyph markers.
  for (const p of directParagraphs(scope)) {
    breaks += splitParagraphAtMarkers(doc, p);
  }

  const after = serialize(doc);
  return { xml: after, changed: before !== after, breaksRestored: breaks };
}

/**
 * Split a merged paragraph at its boundary glyph markers into the original paragraphs, in document
 * order. Segment 0 inherits the merged paragraph's own pPr; each later segment gets the pPr stored in
 * the boundary's payload run, or a clone of the merged pPr (the uniform-card-body common case). The
 * markers and payload runs are dropped. Returns the number of breaks restored (segments created − 1).
 */
function splitParagraphAtMarkers(doc: any, pEl: any): number {
  const content = paragraphContentNodes(pEl);
  if (!content.some((n) => n.nodeType === ELEMENT_NODE && n.nodeName === "w:r" && isGlyphMarkerRun(n))) {
    return 0;
  }

  const basePPr = firstDirectChild(pEl, "w:pPr");
  const segments: any[][] = [[]];
  const segmentPPr: (string | null)[] = [null]; // serialized pPr per segment (null → clone basePPr)

  for (let i = 0; i < content.length; i++) {
    const node = content[i];
    if (node.nodeType === ELEMENT_NODE && node.nodeName === "w:r" && isGlyphMarkerRun(node)) {
      let pPr: string | null = null;
      const next = content[i + 1];
      if (next && next.nodeType === ELEMENT_NODE && next.nodeName === "w:r" && isPPrDataRun(next)) {
        pPr = runTextRaw(next).trim();
        i++; // consume the payload run
      }
      segments.push([]);
      segmentPPr.push(pPr);
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

/** Parse a stored `<w:pPr>` string and import it into `doc` (namespace-wrapped so `w:` resolves). */
function importPPr(doc: any, pPrXml: string): any | null {
  try {
    const wrapped = `<w:p xmlns:w="${W_NS}">${pPrXml}</w:p>`;
    const frag = parse(wrapped);
    const pPr = frag.getElementsByTagName("w:pPr").item(0);
    return pPr ? doc.importNode(pPr, true) : null;
  } catch {
    return null;
  }
}

/** Parse a stored mark `<w:rPr>` string and import it into `doc` (namespace-wrapped so `w:` resolves). */
function importRPr(doc: any, rPrXml: string): any | null {
  try {
    const wrapped = `<w:p xmlns:w="${W_NS}"><w:pPr>${rPrXml}</w:pPr></w:p>`;
    const frag = parse(wrapped);
    const rPr = frag.getElementsByTagName("w:rPr").item(0);
    return rPr ? doc.importNode(rPr, true) : null;
  } catch {
    return null;
  }
}
