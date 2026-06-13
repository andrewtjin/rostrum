// The one and only place that parses or edits OOXML.
//
// WHY THIS EXISTS: Office.js has no run object and cannot split or target a range
// on a *formatting* boundary (verified against learn.microsoft.com). So the only
// reliable way to (a) detect a cite character-style inside a paragraph, (b) hide
// only *part* of a paragraph, and (c) hide a paragraph mark for the condensed
// view (decision #5), is to round-trip the paragraph's OOXML: read `<w:r>` runs,
// set/clear `<w:vanish/>`, and write it back. `<w:vanish/>` is exactly the XML
// the "Hidden" font attribute writes, so this is behaviorally identical to
// font.hidden — and natively reversible the same way (decisions #1/#3).
//
// Everything here is a pure string-in / string-out transform; it runs in Node
// (tests) and the task-pane browser (prod) via @xmldom/xmldom, with no Word host.

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { RunView, BridgeSplit } from "./types";
import { CITE_STYLE_ID } from "./styles";

// xmldom nodes are structurally the DOM but their TS types differ across
// versions; we keep public signatures fully typed (string in/out) and use a
// localized `any` for node handles, which is the pragmatic, version-robust choice
// for an XML adapter.
/* eslint-disable @typescript-eslint/no-explicit-any */

const ELEMENT_NODE = 1;

/** Tags whose presence in a run makes it ineligible to hide (decision #16). */
const INELIGIBLE_RUN_TAGS = [
  "w:fldChar",
  "w:instrText",
  "w:footnoteReference",
  "w:endnoteReference",
  "w:drawing",
  "w:object",
  "w:pict"
];

function parse(xml: string): any {
  // Throw on structurally fatal OOXML so the orchestrator can SKIP that paragraph
  // (leaving it unchanged) instead of silently writing back a half-parsed,
  // corrupted tree. Warnings are ignored, and we never console.error (keeps the
  // task-pane console clean on the expected malformed-input path).
  return new DOMParser({
    onError: (level: string, message: string) => {
      if (level === "fatalError") throw new Error(message);
    }
  }).parseFromString(xml, "text/xml");
}

function serialize(doc: any): string {
  return new XMLSerializer().serializeToString(doc);
}

/**
 * The single body paragraph a WordPort hands us. `Paragraph.getOoxml()` returns a
 * flat-OPC *package* that may also contain header/footer/footnote parts whose
 * `<w:p>` precede the document body in document order — so we scope the search to
 * `<w:body>` (the document story) and never target another story (decision #16).
 * Bare `<w:p>` fragments (tests) have no `<w:body>` and fall back to the whole doc.
 */
function firstParagraph(doc: any): any | null {
  const bodies = doc.getElementsByTagName("w:body");
  const scope = bodies && bodies.length > 0 ? bodies.item(0) : doc;
  const paras = scope.getElementsByTagName("w:p");
  return paras && paras.length > 0 ? paras.item(0) : null;
}

/** Direct element children of `el` whose nodeName equals `name`. */
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

/** Concatenated visible text of a run: w:t verbatim, w:tab -> \t, w:br/w:cr -> \n. */
function runText(runEl: any): string {
  let text = "";
  const walk = (node: any): void => {
    const kids = node.childNodes;
    for (let i = 0; i < kids.length; i++) {
      const k = kids.item(i);
      if (k.nodeType !== ELEMENT_NODE) continue;
      if (k.nodeName === "w:t") {
        text += k.textContent != null ? k.textContent : "";
      } else if (k.nodeName === "w:tab") {
        text += "\t";
      } else if (k.nodeName === "w:br" || k.nodeName === "w:cr") {
        text += "\n";
      } else {
        walk(k);
      }
    }
  };
  walk(runEl);
  return text;
}

/** The run's `<w:rPr>` if present (does not create one). */
function runRPr(runEl: any): any | null {
  return firstDirectChild(runEl, "w:rPr");
}

/** Lower-cased `w:highlight` value of a run, or null when absent / "none". */
function runHighlight(runEl: any): string | null {
  const rPr = runRPr(runEl);
  if (!rPr) return null;
  const hl = firstDirectChild(rPr, "w:highlight");
  if (!hl) return null;
  const val = (hl.getAttribute("w:val") || "").toLowerCase();
  if (val === "" || val === "none") return null;
  return val;
}

/** True when the run carries the cite character style. */
function runIsCite(runEl: any): boolean {
  const rPr = runRPr(runEl);
  if (!rPr) return false;
  const rStyle = firstDirectChild(rPr, "w:rStyle");
  if (!rStyle) return false;
  return (rStyle.getAttribute("w:val") || "") === CITE_STYLE_ID;
}

/**
 * True when the run is underlined: a `<w:u>` is present and not explicitly turned off. OOXML
 * `<w:u w:val="…">` carries the underline STYLE ("single", "double", "wave", …); "none"/"0"/"false"
 * mean no underline. We treat any other value (and a bare `<w:u/>`, though Word always writes a val)
 * as underlined. This is Shrink's keep-signal — it must not key on a specific style, just presence.
 */
function runUnderline(runEl: any): boolean {
  const rPr = runRPr(runEl);
  if (!rPr) return false;
  const u = firstDirectChild(rPr, "w:u");
  if (!u) return false;
  const val = (u.getAttribute("w:val") || "").toLowerCase();
  return !(val === "none" || val === "0" || val === "false");
}

/** True when a `<w:vanish>` is present and not explicitly turned off. */
function vanishIsOn(rPr: any): boolean {
  if (!rPr) return false;
  const v = firstDirectChild(rPr, "w:vanish");
  if (!v) return false;
  const val = (v.getAttribute("w:val") || "").toLowerCase();
  return !(val === "false" || val === "0" || val === "off");
}

/** True when a run contains any tag that makes it ineligible to hide. */
function runIsEligible(runEl: any): boolean {
  for (const tag of INELIGIBLE_RUN_TAGS) {
    const hits = runEl.getElementsByTagName(tag);
    if (hits && hits.length > 0) return false;
  }
  // A run that is the *result* of a simple field (PAGE, DATE, REF, …) lives inside
  // <w:fldSimple>; keep it (decision #16). Walk ancestors up to the paragraph.
  let ancestor = runEl.parentNode;
  while (ancestor && ancestor.nodeName !== "w:p") {
    if (ancestor.nodeName === "w:fldSimple") return false;
    ancestor = ancestor.parentNode;
  }
  return true;
}

/** The internal-part tags — a STRICT SUBSET of INELIGIBLE_RUN_TAGS (no field/footnote refs). */
const INTERNAL_PART_TAGS = ["w:drawing", "w:object", "w:pict"];

/**
 * Build the complete `RunView` for one `<w:r>` in ONE fused full-depth traversal (A3), replacing the
 * legacy four separate scans (`runText` walk + `runIsEligible`'s seven `getElementsByTagName` subtree
 * scans + the highlight/cite/underline/vanish rPr reads + a `hasInternalPart` probe). The single DFS
 * over the run subtree simultaneously (a) accumulates the visible text by the EXACT `runText` rules
 * (`<w:t>` verbatim, `<w:tab>`→"\t", `<w:br>`/`<w:cr>`→"\n", recurse through anything else and never
 * into those leaves) and (b) flags, by nodeName, any INELIGIBLE tag and any INTERNAL_PART tag anywhere
 * in the subtree — which is what `getElementsByTagName` saw, so `eligible`/`hasInternalPart` are
 * byte-identical to the legacy reads. The rPr is a direct child, so highlight/cite/underline/vanish are
 * read straight off it (no extra walk). The `<w:fldSimple>` ancestor rule (an ancestor scan, not a
 * subtree scan) stays separate, exactly as `runIsEligible` did it. `index` is supplied by the caller.
 *
 * EXACTNESS NOTE (why this matches the legacy six-field reads precisely): `getElementsByTagName(tag)`
 * inspects the WHOLE subtree including under `<w:rPr>`; our DFS visits every descendant element too, so
 * a (pathological) ineligible tag nested anywhere is caught identically. Text accumulation visits the
 * same nodes `runText` did in the same document order, so the strings are equal character-for-character.
 */
function buildRunView(runEl: any, index: number): RunView {
  let text = "";
  let eligibleByTag = true; // false once any INELIGIBLE tag is seen (mirrors runIsEligible's tag loop)
  let hasInternalPart = false;

  // One DFS over the run subtree. We stop recursing at the text-leaf nodes (w:t/w:tab/w:br/w:cr) exactly
  // as runText does, and we inspect EVERY visited element's name for the ineligible/internal-part sets.
  const walk = (node: any): void => {
    const kids = node.childNodes;
    if (!kids) return;
    for (let i = 0; i < kids.length; i++) {
      const k = kids.item(i);
      if (k.nodeType !== ELEMENT_NODE) continue;
      const name = k.nodeName;
      // Tag classification (whole-subtree, == getElementsByTagName semantics).
      if (eligibleByTag && INELIGIBLE_RUN_TAGS.indexOf(name) !== -1) eligibleByTag = false;
      if (!hasInternalPart && INTERNAL_PART_TAGS.indexOf(name) !== -1) hasInternalPart = true;
      // Text accumulation (the runText rules — leaves terminate the walk on that branch).
      if (name === "w:t") {
        text += k.textContent != null ? k.textContent : "";
      } else if (name === "w:tab") {
        text += "\t";
      } else if (name === "w:br" || name === "w:cr") {
        text += "\n";
      } else {
        walk(k); // rPr, hyperlink wrappers, fldSimple, drawing internals, … — recurse for tag scan + text
      }
    }
  };
  walk(runEl);

  // The fldSimple ancestor rule (ancestor scan up to <w:p>) — only relevant when no ineligible tag was
  // already found, but we run it whenever tag-eligibility survived, matching runIsEligible's order.
  let eligible = eligibleByTag;
  if (eligible) {
    let ancestor = runEl.parentNode;
    while (ancestor && ancestor.nodeName !== "w:p") {
      if (ancestor.nodeName === "w:fldSimple") {
        eligible = false;
        break;
      }
      ancestor = ancestor.parentNode;
    }
  }

  // rPr is a direct child — read the formatting signals straight off it (no second walk).
  const rPr = runRPr(runEl);
  return {
    index,
    text,
    highlight: runHighlight(runEl),
    citeStyled: runIsCite(runEl),
    underline: runUnderline(runEl),
    hidden: vanishIsOn(rPr),
    eligible,
    hasInternalPart
  };
}

/**
 * Get the run's `<w:rPr>`, creating it as the first child if absent. OOXML
 * requires rPr to lead the run's content, so a freshly created one is inserted
 * before everything else.
 */
function ensureRPr(doc: any, runEl: any): any {
  const existing = runRPr(runEl);
  if (existing) return existing;
  const rPr = doc.createElement("w:rPr");
  runEl.insertBefore(rPr, runEl.firstChild);
  return rPr;
}

/**
 * Get the paragraph's `<w:pPr>`, creating it as the first child if absent (OOXML
 * requires pPr to lead the paragraph).
 */
function ensurePPr(doc: any, pEl: any): any {
  const existing = firstDirectChild(pEl, "w:pPr");
  if (existing) return existing;
  const pPr = doc.createElement("w:pPr");
  pEl.insertBefore(pPr, pEl.firstChild);
  return pPr;
}

/**
 * Get the paragraph-mark run properties (`<w:pPr><w:rPr>`), creating the rPr if
 * absent. It is appended to pPr — the mark's rPr is the last element in the pPr
 * content model, so appending keeps the ordering schema-valid.
 */
function ensureMarkRPr(doc: any, pEl: any): any {
  const pPr = ensurePPr(doc, pEl);
  const existing = firstDirectChild(pPr, "w:rPr");
  if (existing) return existing;
  const rPr = doc.createElement("w:rPr");
  pPr.appendChild(rPr);
  return rPr;
}

/** The existing paragraph-mark rPr without creating anything (for clearing). */
function existingMarkRPr(pEl: any): any | null {
  const pPr = firstDirectChild(pEl, "w:pPr");
  return pPr ? firstDirectChild(pPr, "w:rPr") : null;
}

/**
 * Force `<w:vanish>` on or off inside an rPr. Returns true if the DOM actually
 * changed, so callers can avoid spurious writes / re-serialization.
 */
function setVanish(doc: any, rPr: any, hidden: boolean): boolean {
  const vanishEls = directChildren(rPr, "w:vanish");
  const currentlyOn = vanishIsOn(rPr);
  if (hidden) {
    // Already exactly hidden by a single on-vanish: nothing to do (idempotent).
    if (currentlyOn && vanishEls.length === 1) return false;
    // Normalize duplicates / a stray w:val="false" down to one bare <w:vanish/>.
    for (const v of vanishEls) rPr.removeChild(v);
    rPr.appendChild(doc.createElement("w:vanish"));
    return true;
  }
  // Want visible. If nothing is actually hiding the run (no vanish, or only an
  // explicit w:val="false"), leave the XML untouched to avoid a spurious write.
  if (!currentlyOn) return false;
  // Remove ALL vanish elements in one pass — convergent even with duplicates.
  for (const v of vanishEls) rPr.removeChild(v);
  return true;
}

/**
 * Move ONE boundary space character out of `runEl` (which is being hidden) into a
 * new VISIBLE sibling run, so a single space survives the hide as a word separator
 * (wet-test bug 1). Returns true if it split.
 *
 * The space is MOVED, not duplicated: we drop one space from the run's leading or
 * trailing `<w:t>` and re-add it in the new run, so the paragraph's concatenated
 * text is byte-identical — native Font-dialog reversal and Show All stay lossless.
 * The new run clones the hidden run's `<w:rPr>` minus `<w:vanish>` (so the space is
 * actually visible); whitespace carries no perceptible formatting regardless.
 */
function exposeBoundarySpace(doc: any, runEl: any, side: "lead" | "trail"): boolean {
  // Operate on the run's first (lead) or last (trail) <w:t>; tabs/breaks are not
  // exposable spaces, so the planner never targets them.
  const tNodes = directChildren(runEl, "w:t");
  if (tNodes.length === 0) return false;
  const tEl = side === "lead" ? tNodes[0] : tNodes[tNodes.length - 1];
  const text = tEl.textContent != null ? tEl.textContent : "";

  let remaining: string;
  if (side === "lead") {
    if (!text.startsWith(" ")) return false;
    remaining = text.slice(1);
  } else {
    if (!text.endsWith(" ")) return false;
    remaining = text.slice(0, -1);
  }

  // Rewrite the source <w:t> to the remaining text (replace children directly; we
  // don't rely on a textContent setter being present across xmldom versions).
  while (tEl.firstChild) tEl.removeChild(tEl.firstChild);
  tEl.appendChild(doc.createTextNode(remaining));
  // Significant leading/trailing whitespace needs xml:space="preserve" to survive.
  if (remaining !== remaining.trim()) tEl.setAttribute("xml:space", "preserve");

  // Build the visible space run: clone the (now-hidden) rPr, strip vanish, one space.
  const spaceRun = doc.createElement("w:r");
  const srcRPr = runRPr(runEl);
  if (srcRPr) {
    const clone = srcRPr.cloneNode(true);
    for (const v of directChildren(clone, "w:vanish")) clone.removeChild(v);
    spaceRun.appendChild(clone);
  }
  const wt = doc.createElement("w:t");
  wt.setAttribute("xml:space", "preserve");
  wt.appendChild(doc.createTextNode(" "));
  spaceRun.appendChild(wt);

  // Insert before (lead) or after (trail) the hidden run so the space lands between
  // the two visible chunks.
  if (side === "lead") {
    runEl.parentNode.insertBefore(spaceRun, runEl);
  } else {
    runEl.parentNode.insertBefore(spaceRun, runEl.nextSibling);
  }
  return true;
}

/**
 * 3-way split for an INTERIOR bridge space (wet-test bug 2 follow-up). The hidden run
 * is divided at `offset` (a space character) into [before](hidden) [space](visible)
 * [after](hidden), so two kept chunks on either side of a glued clause read with one
 * space instead of fusing — e.g. "...animals" + hidden ", such as scorpions." + "would"
 * reads "...animals would". The space is MOVED (the run's text is partitioned, not
 * duplicated), so the paragraph's concatenated text stays byte-identical and Show All
 * / native reversal restore the exact original. Returns true if it split.
 *
 * Only a structurally simple run (one <w:t>, no tab/break) is split, so `offset` maps
 * directly to the <w:t> string index; anything else is a no-op (words stay fused —
 * faithful). The keeper picks `offset` so BOTH fragments re-derive as hidden on
 * Re-hide (idempotent); see planCrossGapSeparators priority 4.
 */
function exposeInteriorSpace(doc: any, runEl: any, offset: number): boolean {
  const tNodes = directChildren(runEl, "w:t");
  if (tNodes.length !== 1) return false;
  if (
    directChildren(runEl, "w:tab").length > 0 ||
    directChildren(runEl, "w:br").length > 0 ||
    directChildren(runEl, "w:cr").length > 0
  )
    return false;
  const tEl = tNodes[0];
  const text = tEl.textContent != null ? tEl.textContent : "";
  // Must be a true interior space: a " " with content on both sides.
  if (offset <= 0 || offset >= text.length - 1 || text.charAt(offset) !== " ") return false;
  const before = text.slice(0, offset);
  const after = text.slice(offset + 1);

  // Original run keeps `before` (its rPr already carries <w:vanish/> from the caller).
  while (tEl.firstChild) tEl.removeChild(tEl.firstChild);
  tEl.appendChild(doc.createTextNode(before));
  if (before !== before.trim()) tEl.setAttribute("xml:space", "preserve");

  const srcRPr = runRPr(runEl);

  // Visible space run: clone the rPr minus vanish, one space.
  const spaceRun = doc.createElement("w:r");
  if (srcRPr) {
    const clone = srcRPr.cloneNode(true);
    for (const v of directChildren(clone, "w:vanish")) clone.removeChild(v);
    spaceRun.appendChild(clone);
  }
  const sp = doc.createElement("w:t");
  sp.setAttribute("xml:space", "preserve");
  sp.appendChild(doc.createTextNode(" "));
  spaceRun.appendChild(sp);

  // Hidden after-run: clone the original rPr verbatim (keeps <w:vanish/>), text `after`.
  const afterRun = doc.createElement("w:r");
  if (srcRPr) afterRun.appendChild(srcRPr.cloneNode(true));
  const at = doc.createElement("w:t");
  if (after !== after.trim()) at.setAttribute("xml:space", "preserve");
  at.appendChild(doc.createTextNode(after));
  afterRun.appendChild(at);

  // Insert [before = runEl] [spaceRun] [afterRun] in document order.
  const parent = runEl.parentNode;
  parent.insertBefore(spaceRun, runEl.nextSibling);
  parent.insertBefore(afterRun, spaceRun.nextSibling);
  return true;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * The full per-paragraph hide decision the node-direct apply consumes (Loop 002 B1). It is exactly the
 * three arguments `applyVisibility` already takes, packaged as one object so the wiring agent (CORE-3)
 * passes a single plan from the Phase-A classify into the Phase-B in-place apply. `hideFlags` is aligned
 * to `runs` order (index i hides run i); `splits` (bridge separators) is optional and defaults to none.
 */
export interface VisibilityPlan {
  /** Per-run hide flags, aligned to `runs`/`readRuns` document order. */
  hideFlags: boolean[];
  /** Hide the paragraph mark (`<w:pPr><w:rPr><w:vanish/>`) for the condensed view. */
  hideParaMark: boolean;
  /** Bridge-space exposures so kept chunks separated by hidden prose don't fuse (default: none). */
  splits?: readonly BridgeSplit[];
}

/**
 * A paragraph's OOXML parsed ONCE into a reusable tree.
 *
 * WHY THIS EXISTS (performance — the engine's hot path). The hide pass needs BOTH
 * the run VIEWS (`runs`, to decide what to keep) AND the live run NODES (to toggle
 * `<w:vanish/>`). The original API exposed only string→string functions, so
 * `classifyParagraph` parsed every paragraph TWICE — once in `readRuns` and again in
 * `applyRunVisibility`/`makeAllVisible`. xmldom's `DOMParser` dominates the
 * per-paragraph cost (it builds a full DOM tree from the package string), so that
 * second parse roughly DOUBLED the engine's work on every doc — the same CPU the
 * task-pane browser spends live. Parsing once and reading + mutating through one tree
 * halves it, with byte-identical output (only `<w:vanish/>` is ever toggled).
 *
 * The three string→string functions below (`readRuns`, `applyRunVisibility`,
 * `makeAllVisible`) are now thin wrappers over this class so every existing caller and
 * test keeps its exact contract; the orchestrator (`classifyParagraph`) holds ONE
 * instance and calls `runs` then `applyVisibility`/`makeAllVisible` on it. Policy stays
 * out of here entirely — a caller only ever reads the `RunView[]` and hands back flat
 * index-aligned flags, exactly as it did with the standalone functions.
 */
export class ParsedParagraph {
  /** The parsed package tree (mutated in place by the apply methods). */
  private readonly doc: any;
  /** The single body `<w:p>` we operate on, or null for an unparseable/empty fragment. */
  private readonly pEl: any | null;
  /**
   * Snapshot of the paragraph's run elements in document order (the SAME order as
   * `runs`). Captured at construction — before any mutation — so the apply methods can
   * index into it even after bridge splits insert new `<w:r>` siblings (which would
   * otherwise shift a live NodeList mid-pass). This is exactly the snapshot the old
   * standalone `applyRunVisibility` took at call time; nothing mutates between
   * construction and apply, so it is identical.
   */
  private readonly runEls: any[];
  /**
   * The original OOXML, returned unchanged when a STRING-mode mutation is a no-op (no reserialize).
   * NULL in node mode (`fromNode`): a node-backed paragraph never serializes itself — its owner
   * (`ooxmlPackage.ts`) serializes the whole package once — so there is no per-paragraph "original"
   * string to fall back to, and `applyVisibilityInPlace` reports `changed` only (never an `xml`).
   */
  private readonly original: string | null;
  /** Flat run views in document order — the pure read the keeper policy consumes. */
  readonly runs: RunView[];

  /**
   * True when this is a NODE-backed instance (`fromNode`): it never serializes itself, and the
   * orchestrator must apply visibility through the in-place methods (`applyVisibilityInPlace` /
   * `makeAllVisibleInPlace`), NOT the serializing string-mode ones (which throw in node mode —
   * CONTRACT A). The two-phase Hide reads this to pick the apply method per paragraph, keeping the
   * node/string branch decision on the instance itself rather than re-deriving it from `RawParagraph`.
   */
  get isNodeBacked(): boolean {
    return this.original === null;
  }

  /**
   * String mode (the legacy path, byte-for-byte unchanged): parse the package ONCE and read +
   * mutate through that tree. `readRuns`/`applyRunVisibility`/`makeAllVisible` and every existing
   * caller use this; `parseCount.test.ts`'s "exactly one parse" meaning is preserved because this is
   * the ONLY constructor that calls `parse()`. Use `ParsedParagraph.fromNode` for the zero-parse
   * node-direct path (Loop 002 B1).
   */
  constructor(paragraphOoxml: string) {
    this.original = paragraphOoxml;
    this.doc = parse(paragraphOoxml);
    this.pEl = firstParagraph(this.doc);
    this.runEls = [];
    this.runs = [];
    this.extractRuns();
  }

  /**
   * Node mode (Loop 002 B1 — the ANCHOR): construct over an ALREADY-ATTACHED `<w:p>` Element and its
   * live owner document with ZERO string parse. The package was parsed ONCE by `ooxmlPackage.ts`; we
   * reuse that exact live node + document so the hide pass reads and mutates the persistent tree in
   * place, deleting the per-paragraph serialize→parse the string path forces (the win 002-S1 names,
   * falsifiable as "parseCount===0 on the node-direct path"). The run extraction is the SAME fused
   * walk the string ctor uses, so the `RunView[]` is six-field-(plus-hasInternalPart)-identical.
   *
   * Mutations go through `applyVisibilityInPlace` (NOT `applyVisibility`, which serializes): they edit
   * the attached nodes and the owner serializes the whole package later. `pEl` must be a `<w:p>` owned
   * by `doc` (the caller guarantees this); we do not re-scope or re-find it.
   */
  static fromNode(doc: any, pEl: any): ParsedParagraph {
    // Object.create bypasses the string constructor's parse() — the whole point of node mode. We then
    // populate the same private fields the string ctor sets, sharing extractRuns() for the run views.
    const inst: ParsedParagraph = Object.create(ParsedParagraph.prototype);
    // Assign the readonly fields once here (legal at construction-time via the any-cast on a fresh
    // instance); after this the instance is immutable in the same way a constructed one is.
    const w = inst as any;
    w.doc = doc;
    w.pEl = pEl ?? null;
    w.original = null; // node mode never serializes itself — see the field comment.
    w.runEls = [];
    w.runs = [];
    inst.extractRuns();
    return inst;
  }

  /**
   * Populate `runEls` + `runs` from `pEl` via ONE fused full-depth traversal per `<w:r>` (A3). Shared by
   * both modes so the node path's views are identical to the string path's. Snapshots the run elements in
   * document order BEFORE any mutation, so the apply methods index correctly even after bridge splits
   * insert new `<w:r>` siblings (which would otherwise shift a live NodeList mid-pass).
   */
  private extractRuns(): void {
    if (!this.pEl) return;
    const runList = this.pEl.getElementsByTagName("w:r");
    for (let i = 0; i < runList.length; i++) {
      const r = runList.item(i);
      (this.runEls as any[]).push(r);
      (this.runs as RunView[]).push(buildRunView(r, i));
    }
  }

  /**
   * Apply per-run hide flags (aligned to `runs` order) and optionally hide the
   * paragraph mark. `splits` (wet-test bug 1) name hidden runs whose boundary space
   * must be exposed as a visible separator so isolated kept chunks don't fuse — see
   * `BridgeSplit`. Returns the new OOXML plus whether anything actually changed.
   */
  applyVisibility(
    hideFlags: boolean[],
    hideParaMark: boolean,
    splits: readonly BridgeSplit[] = []
  ): { xml: string; changed: boolean } {
    // CONTRACT (A) — node-mode guard. A node-backed instance (`fromNode`, `original===null`) never
    // serializes itself: its owner package serializes the whole body once. `applyVisibility` ends by
    // returning `serialize(doc)` only on change but `this.original` on a no-op — which would emit
    // `null` as the "unchanged" paragraph OOXML and silently corrupt the commit. The node path MUST
    // call `applyVisibilityInPlace` instead, so refuse the string-mode method here rather than emit null.
    if (this.original === null) {
      throw new Error(
        "ParsedParagraph.applyVisibility called on a node-mode instance — use applyVisibilityInPlace " +
          "(a node-backed paragraph never serializes itself; its owner serializes the whole package)."
      );
    }
    // String-mode only (the wrappers call it on a string-constructed instance), so `original` is set.
    if (!this.pEl) return { xml: this.original as string, changed: false };
    const { doc, pEl, runEls } = this;

    let changed = false;
    const n = Math.min(runEls.length, hideFlags.length);
    for (let i = 0; i < n; i++) {
      if (hideFlags[i]) {
        // Hiding: ensure an rPr exists to carry <w:vanish/>.
        if (setVanish(doc, ensureRPr(doc, runEls[i]), true)) changed = true;
      } else {
        // Showing: don't synthesize an rPr just to mark "visible". A freshly-inserted
        // bridge-space run has no rPr; creating an empty one here would be pointless
        // churn on every Re-hide. Only clear an existing vanish (mirrors the
        // paragraph-mark handling below). Keeps Re-hide idempotent by construction.
        const rPr = runRPr(runEls[i]);
        if (rPr && setVanish(doc, rPr, false)) changed = true;
      }
    }

    // Expose bridge spaces AFTER setting vanish, so the moved space inherits the
    // (now-hidden) run's formatting minus vanish and renders visibly. "interior" does a
    // 3-way split at sp.offset; "lead"/"trail" move a boundary space.
    for (const sp of splits) {
      if (sp.index < 0 || sp.index >= runEls.length) continue;
      const did =
        sp.side === "interior"
          ? exposeInteriorSpace(doc, runEls[sp.index], sp.offset == null ? -1 : sp.offset)
          : exposeBoundarySpace(doc, runEls[sp.index], sp.side);
      if (did) changed = true;
    }

    if (hideParaMark) {
      if (setVanish(doc, ensureMarkRPr(doc, pEl), true)) changed = true;
    } else {
      // Don't create nodes just to set "visible"; only clear an existing mark vanish.
      const markRPr = existingMarkRPr(pEl);
      if (markRPr && setVanish(doc, markRPr, false)) changed = true;
    }

    return changed
      ? { xml: serialize(doc), changed: true }
      : { xml: this.original as string, changed: false };
  }

  /**
   * Node-direct visibility apply (Loop 002 B1): the SAME `<w:vanish>` toggles + bridge splits as
   * `applyVisibility`, but performed IN PLACE on the live attached nodes and returning ONLY whether the
   * DOM changed — it NEVER serializes. The owner package (`ooxmlPackage.ts`) serializes the whole body
   * once after the pass, so paying a per-paragraph serialize here would defeat the entire optimization.
   *
   * It reuses the EXACT same primitives the string path uses — `ensureRPr`/`setVanish`/`ensureMarkRPr`/
   * `exposeBoundarySpace`/`exposeInteriorSpace`, all of which build nodes with `doc.createElement("w:…")`
   * (NOT `createElementNS`; loss-3: a namespaced node serializes a redundant `xmlns:w` the semantic oracle
   * flags as an out-of-position delta). The interior-split "after" run clones the source rPr VERBATIM
   * (the oracle checks rPr-minus-vanish equality). Because `setVanish` is convergent and the show branch
   * never synthesizes an rPr, re-applying the same plan is a no-op → `changed:false` (idempotent).
   *
   * On a node-less instance (no `<w:p>`) it is a safe no-op. Any throw from a primitive propagates to the
   * caller UNCAUGHT — Phase B's whole-op abort contract (002-F4) depends on it surfacing, not being
   * swallowed; this method does not try/catch.
   */
  applyVisibilityInPlace(plan: VisibilityPlan): { changed: boolean } {
    if (!this.pEl) return { changed: false };
    const { doc, pEl, runEls } = this;
    const { hideFlags, hideParaMark, splits = [] } = plan;

    let changed = false;
    const n = Math.min(runEls.length, hideFlags.length);
    for (let i = 0; i < n; i++) {
      if (hideFlags[i]) {
        if (setVanish(doc, ensureRPr(doc, runEls[i]), true)) changed = true;
      } else {
        // Show: never synthesize an rPr just to mark visible (keeps Re-hide idempotent); only clear
        // an existing vanish. Identical to applyVisibility's show branch.
        const rPr = runRPr(runEls[i]);
        if (rPr && setVanish(doc, rPr, false)) changed = true;
      }
    }

    // Bridge splits AFTER vanish (moved space inherits the now-hidden run's rPr minus vanish).
    for (const sp of splits) {
      if (sp.index < 0 || sp.index >= runEls.length) continue;
      const did =
        sp.side === "interior"
          ? exposeInteriorSpace(doc, runEls[sp.index], sp.offset == null ? -1 : sp.offset)
          : exposeBoundarySpace(doc, runEls[sp.index], sp.side);
      if (did) changed = true;
    }

    if (hideParaMark) {
      if (setVanish(doc, ensureMarkRPr(doc, pEl), true)) changed = true;
    } else {
      const markRPr = existingMarkRPr(pEl);
      if (markRPr && setVanish(doc, markRPr, false)) changed = true;
    }

    return { changed };
  }

  /**
   * Reveal everything in the paragraph: clear `<w:vanish>` from every run and from
   * the paragraph mark. Convergent/idempotent — Show All's per-paragraph operation
   * (decision #14). May reveal a user's own manually-hidden run; that over-reveal edge
   * is documented and warned once (decision #10).
   */
  makeAllVisible(): { xml: string; changed: boolean } {
    // CONTRACT (A) — node-mode guard (see applyVisibility). A node-backed instance would return
    // `this.original` (null) on a no-op reveal, emitting null OOXML; the node path must use
    // `makeAllVisibleInPlace`. Refuse here rather than serialize null.
    if (this.original === null) {
      throw new Error(
        "ParsedParagraph.makeAllVisible called on a node-mode instance — use makeAllVisibleInPlace " +
          "(a node-backed paragraph never serializes itself; its owner serializes the whole package)."
      );
    }
    // String-mode only (the wrappers call it on a string-constructed instance), so `original` is set.
    if (!this.pEl) return { xml: this.original as string, changed: false };
    const { doc, pEl, runEls } = this;

    let changed = false;
    for (let i = 0; i < runEls.length; i++) {
      const rPr = runRPr(runEls[i]);
      if (rPr && setVanish(doc, rPr, false)) changed = true;
    }
    const markRPr = existingMarkRPr(pEl);
    if (markRPr && setVanish(doc, markRPr, false)) changed = true;

    return changed
      ? { xml: serialize(doc), changed: true }
      : { xml: this.original as string, changed: false };
  }

  /**
   * Node-direct reveal (Loop 002 B1, CONTRACT A): clear `<w:vanish>` from every run AND the paragraph
   * mark IN PLACE on the live attached nodes, returning ONLY whether the DOM changed — it NEVER
   * serializes (the owner package serializes the whole body once after the pass). This is the node-mode
   * twin of `makeAllVisible`, used by the keepWhole / heading / cite / table SELF-HEAL: where the string
   * path forces such a paragraph visible via `makeAllVisible`, the node path forces it visible via this.
   *
   * It reuses the EXACT same primitive (`setVanish(...,false)` via `doc.createElement`-built nodes — no
   * createElementNS, loss-3) the string `makeAllVisible` uses, so the in-place reveal is byte-equivalent
   * to a string-mode reveal then re-parse. Convergent/idempotent: a second reveal clears nothing new →
   * `changed:false`. On a node-less instance (no `<w:p>`) it is a safe no-op. Any primitive throw
   * propagates UNCAUGHT (Phase B's whole-op abort, 002-F4, depends on it surfacing).
   */
  makeAllVisibleInPlace(): { changed: boolean } {
    if (!this.pEl) return { changed: false };
    const { doc, pEl, runEls } = this;

    let changed = false;
    for (let i = 0; i < runEls.length; i++) {
      const rPr = runRPr(runEls[i]);
      if (rPr && setVanish(doc, rPr, false)) changed = true;
    }
    const markRPr = existingMarkRPr(pEl);
    if (markRPr && setVanish(doc, markRPr, false)) changed = true;

    return { changed };
  }
}

/**
 * TEST-ONLY legacy six-field (+hasInternalPart) reference reader (A3 differential oracle, 002-S3).
 *
 * WHY IT EXISTS. The production read path is now the FUSED single traversal `buildRunView`. To prove
 * that fusion is faithful, the A3 test compares it against the ORIGINAL, independent per-signal scans —
 * the separate `runText` walk + `runIsEligible`'s seven `getElementsByTagName` subtree scans + the
 * individual `runHighlight`/`runIsCite`/`runUnderline`/`vanishIsOn` rPr reads + a standalone internal-part
 * probe. This function reconstructs EXACTLY that legacy logic over the same `<w:p>` so the differential
 * is genuine (fused vs unfused), not fused-vs-itself. It is NOT used by any production path; it exists so
 * `runText`/`runIsEligible` stay the live reference the proof gates against. `parseFromString` is never
 * called here — it reads an already-parsed package, so it is safe in a parseCount===0 assertion.
 */
export function legacyRunViewsForTest(paragraphOoxml: string): RunView[] {
  const doc = parse(paragraphOoxml);
  const pEl = firstParagraph(doc);
  const out: RunView[] = [];
  if (!pEl) return out;
  const runList = pEl.getElementsByTagName("w:r");
  for (let i = 0; i < runList.length; i++) {
    const r = runList.item(i);
    // Internal-part probe via the same whole-subtree getElementsByTagName the rider specifies.
    let hasInternalPart = false;
    for (const tag of INTERNAL_PART_TAGS) {
      if (r.getElementsByTagName(tag).length > 0) {
        hasInternalPart = true;
        break;
      }
    }
    out.push({
      index: i,
      text: runText(r), // the standalone walk
      highlight: runHighlight(r),
      citeStyled: runIsCite(r),
      underline: runUnderline(r),
      hidden: vanishIsOn(runRPr(r)),
      eligible: runIsEligible(r), // the standalone seven-scan + fldSimple-ancestor reader
      hasInternalPart
    });
  }
  return out;
}

/** Read a paragraph's runs as flat, document-order views. Pure read. */
export function readRuns(paragraphOoxml: string): RunView[] {
  return new ParsedParagraph(paragraphOoxml).runs;
}

/**
 * Apply per-run hide flags (aligned to `readRuns` order) and optionally hide the
 * paragraph mark. Thin wrapper over `ParsedParagraph.applyVisibility` for external
 * callers/tests; the hide path reuses one `ParsedParagraph` instead (one parse).
 */
export function applyRunVisibility(
  paragraphOoxml: string,
  hideFlags: boolean[],
  hideParaMark: boolean,
  splits: readonly BridgeSplit[] = []
): { xml: string; changed: boolean } {
  return new ParsedParagraph(paragraphOoxml).applyVisibility(hideFlags, hideParaMark, splits);
}

/**
 * Reveal everything in the paragraph. Thin wrapper over
 * `ParsedParagraph.makeAllVisible` for external callers/tests.
 */
export function makeAllVisible(paragraphOoxml: string): { xml: string; changed: boolean } {
  return new ParsedParagraph(paragraphOoxml).makeAllVisible();
}
