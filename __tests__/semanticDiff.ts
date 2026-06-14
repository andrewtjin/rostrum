// THE SEMANTIC DIFF ORACLE — the R1 losslessness gate for Loop 002 (PLAN.md §8
// loss-1 / loss-4; CASES.md 002-S1/002-F1).
//
// WHY THIS EXISTS (and why it is NOT a byte-diff). The hide path commits with
// `insertOoxml(serialize(...), "Replace")`, and @xmldom/xmldom re-serializes the WHOLE
// body on every Hide. So the committed package is NOT byte-identical to Word's input —
// the prior team parked P1 for exactly this (optimization-sweep verdict #0). A naive
// "output bytes == input bytes" gate is therefore structurally unachievable and would
// either always fail or be deleted. The REAL invariant 002-F1 needs is SEMANTIC: the
// engine may add/remove `<w:vanish>` and may MOVE a single bridge space into a new
// visible run — and NOTHING ELSE. This oracle parses both sides with the SAME parser
// the engine uses (@xmldom/xmldom, so it sees exactly what the engine sees) and THROWS,
// with a precise paragraph-index + offending-node message, on any other delta.
//
// THE PERMITTED DELTAS (the whitelist — everything else is a violation):
//   (a) An added or removed `<w:vanish/>` element inside a run's `<w:rPr>` (and the
//       equivalent on the paragraph mark `<w:pPr><w:rPr>`). This is the entire Hide /
//       Show-All mechanism (ooxml.ts `setVanish`).
//   (b) A bridge-split insertion (ooxml.ts `exposeBoundarySpace` / `exposeInteriorSpace`):
//       a NEW `<w:r>` whose visible text is exactly one moved space, carrying
//       `xml:space="preserve"`, paired with the SOURCE run's `<w:t>` having that one
//       space removed. The space is MOVED, never duplicated — so the paragraph's
//       concatenated visible text is unchanged (asserted hard, below). For an INTERIOR
//       split the source run is additionally divided into a [before] (source) and an
//       [after] (a new hidden run cloning the source rPr verbatim) — both permitted.
//
// THE HARD ASSERTIONS (these ARE the 002-F1 losslessness checks):
//   1. For EVERY body paragraph, the CONCATENATED visible run text (w:t verbatim,
//      w:tab -> "\t", w:br/w:cr -> "\n", over ALL runs in document order — the exact
//      `runText` walk the engine uses) is BYTE-IDENTICAL input vs output. A moved bridge
//      space must not change this; a changed/added/dropped character is caught here.
//   2. Run order is preserved MODULO the inserted bridge runs (alignment walk below).
//   3. Every `<w:rPr>`'s children/attributes are identical input vs output EXCEPT the
//      presence/absence of `<w:vanish>`.
//   4. Paragraph COUNT and paragraph STRUCTURE (table nesting, pPr/pStyle, the
//      element shape around runs) are identical.
//
// Parser-faithful by construction: we never compare strings of XML, only parsed nodes,
// so attribute order / self-closing form / whitespace re-emission never produce a false
// positive (those are xmldom serialization artifacts, not semantic deltas). The oracle
// is deliberately STRICT — a permissive oracle that green-passes bugs is worthless, so
// `semanticDiff.selftest.test.ts` proves it throws on every violation class.

import { DOMParser } from "@xmldom/xmldom";

// xmldom node types vary across versions; the engine itself uses a localized `any` for
// node handles for exactly this reason (see ooxml.ts), and we match that choice so the
// oracle inspects nodes the same way the engine does.
/* eslint-disable @typescript-eslint/no-explicit-any */

const ELEMENT_NODE = 1;
const TEXT_NODE = 3;

/**
 * Parse with the SAME fatal-only error policy the engine uses (ooxml.ts / ooxmlPackage.ts),
 * so the oracle accepts exactly the inputs the engine accepts and rejects exactly what it
 * rejects. A non-fatal warning is ignored; a fatal error throws (a malformed package is
 * itself a losslessness failure worth surfacing).
 */
function parse(xml: string): any {
  return new DOMParser({
    onError: (level: string, message: string) => {
      if (level === "fatalError") throw new Error(message);
    }
  }).parseFromString(xml, "text/xml");
}

/**
 * The body-story scope, matching ooxmlPackage.ts `bodyScope` EXACTLY: the first
 * `<w:body>` when present, else the Document node (for a bare `<w:p>` fragment — the
 * Document, not documentElement, so `getElementsByTagName` finds the root `<w:p>` as a
 * descendant). Using the engine's own scoping is what keeps the oracle honest: it diffs
 * precisely the paragraphs the engine reads and writes.
 */
function bodyScope(doc: any): any {
  const bodies = doc.getElementsByTagName("w:body");
  return bodies && bodies.length > 0 ? bodies.item(0) : doc;
}

/** True when a node lives inside a text box — matches ooxmlPackage.ts `isInTextbox`. */
function isInTextbox(node: any): boolean {
  let n = node ? node.parentNode : null;
  while (n) {
    if (n.nodeName === "w:txbxContent") return true;
    n = n.parentNode;
  }
  return false;
}

/**
 * STORY paragraphs only — every `<w:p>` in the scope EXCEPT textbox-nested ones, matching
 * ooxmlPackage.ts `storyParagraphsIn`. These are the paragraphs that align 1:1 with the
 * engine's whole-body splice index, so the oracle compares the same population.
 */
function storyParagraphs(doc: any): any[] {
  const live = bodyScope(doc).getElementsByTagName("w:p");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) {
    const p = live.item(i);
    if (p && !isInTextbox(p)) out.push(p);
  }
  return out;
}

/** Direct element children of `el` whose nodeName equals `name`. */
function directChildren(el: any, name: string): any[] {
  const out: any[] = [];
  const kids = el.childNodes;
  if (!kids) return out;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE && k.nodeName === name) out.push(k);
  }
  return out;
}

/** The first direct child named `name`, or null. */
function firstDirectChild(el: any, name: string): any | null {
  const found = directChildren(el, name);
  return found.length > 0 ? found[0] : null;
}

/**
 * Concatenated visible text of a run — the EXACT `runText` walk from ooxml.ts (w:t
 * verbatim, w:tab -> "\t", w:br/w:cr -> "\n", recursing through anything else). Reusing
 * the engine's own walk is what makes assertion #1 a faithful text-preservation check.
 */
function runText(runEl: any): string {
  let text = "";
  const walk = (node: any): void => {
    const kids = node.childNodes;
    if (!kids) return;
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

/**
 * The concatenated visible text of an ENTIRE paragraph: every `<w:r>` (in document order)
 * run through `runText`, EXCLUDING runs nested in a text box (which belong to a different
 * story). This is the quantity assertion #1 requires to be byte-identical input vs output;
 * because a bridge split MOVES a space (drops it from the source `<w:t>`, re-adds it in a
 * new run), this total is invariant across a legal hide — and changes the instant a real
 * character is added, dropped, or altered.
 */
function paragraphVisibleText(pEl: any): string {
  const runs = pEl.getElementsByTagName("w:r");
  let out = "";
  for (let i = 0; i < runs.length; i++) {
    const r = runs.item(i);
    if (!isInTextbox(r)) out += runText(r);
  }
  return out;
}

/** All `<w:r>` of a paragraph, in document order, EXCLUDING textbox-nested runs. */
function paragraphRuns(pEl: any): any[] {
  const live = pEl.getElementsByTagName("w:r");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) {
    const r = live.item(i);
    if (r && !isInTextbox(r)) out.push(r);
  }
  return out;
}

/** A stable, order-sensitive signature of an element's attributes (name="value", sorted). */
function attrSignature(el: any): string {
  const attrs = el.attributes;
  if (!attrs) return "";
  const pairs: string[] = [];
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item ? attrs.item(i) : attrs[i];
    const name: string = a && (a.name ?? a.nodeName);
    if (!name) continue;
    pairs.push(`${name}=${a.value ?? a.nodeValue ?? ""}`);
  }
  // Sort so attribute ORDER (an xmldom serialization detail) never reads as a delta;
  // attribute SET and VALUES still must match exactly.
  pairs.sort();
  return pairs.join("");
}

/**
 * The attribute signature of an element with one specific attribute filtered out (matched by
 * exact name+value). Used ONLY to compare two `<w:t>` elements modulo an OUTPUT-added
 * `xml:space="preserve"` — see `wtAttrsEqualModuloAddedPreserve`. A pure read; never mutates.
 */
function attrSignatureWithout(el: any, dropName: string, dropValue: string): string {
  const attrs = el.attributes;
  if (!attrs) return "";
  const pairs: string[] = [];
  for (let i = 0; i < attrs.length; i++) {
    const a = attrs.item ? attrs.item(i) : attrs[i];
    const name: string = a && (a.name ?? a.nodeName);
    if (!name) continue;
    const value: string = a.value ?? a.nodeValue ?? "";
    if (name === dropName && value === dropValue) continue; // drop exactly the targeted attr
    pairs.push(`${name}=${value}`);
  }
  pairs.sort();
  return pairs.join("");
}

/**
 * PERMITTED DELTA (b), companion clause — `xml:space="preserve"` newly required by a bridge trim.
 *
 * A bridge split (`exposeBoundarySpace`/`exposeInteriorSpace` in ooxml.ts) TRIMS one space out of the
 * source run's `<w:t>` and MOVES it into a new visible run. When that trim leaves the source `<w:t>` with
 * leading or trailing whitespace (e.g. the interior-split [before] " point" or [after] "to a … they "),
 * the engine MUST add `xml:space="preserve"` — without it Word collapses the now-boundary space, which
 * would itself BREAK losslessness. So this attribute is not a smuggled change: it is the inescapable,
 * reversibility-PRESERVING consequence of the already-permitted text trim.
 *
 * Word omits `xml:space` whenever a run's whitespace is interior (` point to … they ` mid-run needs no
 * preserve), so a real-corpus `<w:t>` frequently has NO `xml:space` on input yet REQUIRES it on the
 * trimmed output — the exact xlarge paragraph[2330]/run[27] case. The string path and node-direct path
 * produce byte-identical output here (proven), and whole-paragraph visible text stays byte-identical
 * (assertion #1), so this is provably lossless and reversible.
 *
 * This clause is DELIBERATELY narrow — it accepts ONLY the OUTPUT gaining `xml:space="preserve"` that the
 * INPUT lacked, and ONLY on a `<w:t>`. It returns true iff the two `<w:t>` attribute sets are equal after
 * removing an `xml:space="preserve"` present on the OUTPUT but absent on the INPUT. It therefore still
 * REJECTS: any non-`xml:space` attribute change, an `xml:space` whose value isn't "preserve", the INPUT
 * carrying an `xml:space` the OUTPUT dropped (a real space-collapse risk), or an `xml:space` already on
 * the INPUT (this clause licenses a pure ADDITION, never a change/removal).
 */
function wtAttrsEqualModuloAddedPreserve(inWt: any, outWt: any): boolean {
  const PRESERVE = "preserve";
  const SPACE = "xml:space";
  const inSpace = inWt.getAttribute ? inWt.getAttribute(SPACE) : null;
  const outSpace = outWt.getAttribute ? outWt.getAttribute(SPACE) : null;
  // The ONLY licensed shape: the INPUT had NO xml:space and the OUTPUT gained xml:space="preserve".
  // If the input already had one, this clause licenses nothing (a change/removal is never permitted);
  // if the output's value isn't exactly "preserve", it isn't the engine's bridge-trim companion.
  if (inSpace != null) return false;
  if (outSpace !== PRESERVE) return false;
  // Every OTHER attribute must still be byte-identical — so an injected/changed attribute (or any
  // attribute beyond the single added xml:space) is still caught.
  return attrSignature(inWt) === attrSignatureWithout(outWt, SPACE, PRESERVE);
}

/**
 * True when `<w:rPr>` (or `<w:pPr>`'s mark rPr) is "vanish-only": after removing every
 * `<w:vanish>` it has NO element children left. The engine's `ensureRPr`+`setVanish` path
 * creates exactly such an rPr when hiding a run that had none — so a vanish-only rPr in the
 * OUTPUT must read as equivalent to NO rPr in the INPUT (and vice-versa for Show All). This
 * is what makes permitted delta (a) work for rPr-less runs (the `rPrLessRuns` fixture).
 */
function isVanishOnlyRPr(el: any): boolean {
  if (el.nodeName !== "w:rPr") return false;
  const kids = el.childNodes;
  if (!kids) return true;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE && k.nodeName !== "w:vanish") return false;
  }
  return true;
}

/**
 * True when a `<w:pPr>` is "mark-vanish-only": its ONLY element child is a vanish-only
 * `<w:rPr>` (the paragraph-mark rPr the engine synthesized purely to hide the mark on a
 * paragraph that had no `<w:pPr>`). Such a pPr is equivalent to an absent pPr, so dropping
 * it keeps "hide the mark on a markless paragraph" a permitted delta (a) rather than a
 * spurious structural change. A pPr carrying ANY other property (pStyle, numPr, outlineLvl,
 * a non-vanish mark rPr) is NOT droppable — those are real structure.
 */
function isMarkVanishOnlyPPr(el: any): boolean {
  if (el.nodeName !== "w:pPr") return false;
  const kids = el.childNodes;
  if (!kids) return true;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType !== ELEMENT_NODE) continue;
    if (k.nodeName === "w:rPr" && isVanishOnlyRPr(k)) continue;
    return false; // any other property element makes the pPr real
  }
  return true;
}

/**
 * The children of an element to compare structurally, applying the permitted-delta (a)
 * normalizations so the engine's vanish mechanics never read as a structural change:
 *   * Inside a `<w:rPr>`, drop every `<w:vanish>` child (the toggle itself).
 *   * Drop a CHILD `<w:rPr>` that is vanish-only after that filtering — i.e. an rPr the
 *     engine synthesized solely to carry `<w:vanish/>` on a previously rPr-less run. With
 *     this, "no rPr" and "rPr containing only vanish" compare equal (both → no rPr child).
 * Whitespace-only text nodes between elements are xmldom artifacts and are skipped so
 * indentation never reads as a delta; significant text (inside `<w:t>`) is always kept.
 */
function elementChildrenForCompare(el: any): any[] {
  const dropVanish = el.nodeName === "w:rPr";
  const out: any[] = [];
  const kids = el.childNodes;
  if (!kids) return out;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE) {
      if (dropVanish && k.nodeName === "w:vanish") continue; // permitted delta (a)
      // A synthesized vanish-only rPr is equivalent to an absent rPr — drop it entirely.
      if (k.nodeName === "w:rPr" && isVanishOnlyRPr(k)) continue;
      // A pPr synthesized solely to hide the paragraph mark is equivalent to an absent pPr.
      if (k.nodeName === "w:pPr" && isMarkVanishOnlyPPr(k)) continue;
      out.push(k);
    } else if (k.nodeType === TEXT_NODE) {
      const t = k.textContent ?? "";
      // Keep ALL text under a <w:t> (significant); elsewhere keep only non-whitespace
      // (structural elements don't carry meaningful mixed text — only formatting markup).
      if (el.nodeName === "w:t" || t.trim().length > 0) out.push(k);
    }
  }
  return out;
}

/** True when a run is exactly a bridge-space run: visible text is one space, `xml:space="preserve"`. */
function isBridgeSpaceRun(runEl: any): boolean {
  if (runEl.nodeName !== "w:r") return false;
  // Its visible text (the runText walk) must be exactly a single ASCII space. A bridge
  // run never carries a tab/break, and it carries no <w:vanish> (it is the VISIBLE space).
  if (runText(runEl) !== " ") return false;
  // Must hold a <w:t xml:space="preserve"> — a moved significant space requires it.
  const t = firstDirectChild(runEl, "w:t");
  if (!t) return false;
  const space = t.getAttribute ? t.getAttribute("xml:space") : null;
  if (space !== "preserve") return false;
  // A bridge run must not itself be hidden (it exists to keep a space VISIBLE).
  const rPr = firstDirectChild(runEl, "w:rPr");
  if (rPr && directChildren(rPr, "w:vanish").length > 0) return false;
  return true;
}

/**
 * Compare the run sequences of one paragraph, allowing the engine's bridge-split insertions
 * in the OUTPUT (delta (b)) while forbidding anything else (assertions #2/#3).
 *
 * THE MODEL. A legal hide transforms the input runs into the output runs by, per run:
 *   * toggling `<w:vanish>` (no run-count change), and/or
 *   * a bridge split, which expands ONE source run into a contiguous group of output runs:
 *       - "lead"/"trail": a new VISIBLE single-space run is inserted adjacent to the source
 *         run, and one space is removed from the source run's `<w:t>` (boundary). The space
 *         run sits just before (lead) or after (trail) the hidden source run.
 *       - "interior": the source run is divided into [before](source, hidden) + [space]
 *         (new, visible) + [after](new, hidden, clones source rPr verbatim). One source run
 *         becomes THREE output runs whose texts concatenate to the source text.
 *   * No run is ever reordered or deleted; the only NEW runs are bridge fragments.
 *
 * THE ALGORITHM (consume each input run exactly once, in order). Walk the output runs with
 * cursor `j`; track the next unconsumed input run with cursor `i` and the rPr (modulo vanish)
 * of the most-recently-consumed source run. For each output run:
 *   (1) If it STRUCTURALLY matches the next input run modulo `<w:vanish>` and modulo `<w:t>`
 *       TEXT (same nodeName/attrs/rPr-minus-vanish/non-text shape), it is that run's PRIMARY
 *       image — consume input run `i` (i++), remember its rPr signature.
 *   (2) Else if it is a permitted INSERTED bridge fragment — a bridge-space run, OR a
 *       text-only "after" run whose rPr-minus-vanish equals the just-consumed source run's —
 *       skip it (a delta-(b) insertion).
 *   (3) Else it is a violation (a smuggled run, a reordered/foreign run, an injected format).
 * At the end EVERY input run must have been consumed (else a run was deleted). Per-run TEXT
 * is intentionally NOT compared here — assertion #1 (whole-paragraph text byte-identical)
 * already guarantees the moved space was moved, not duplicated or dropped; here we only
 * police STRUCTURE and rPr, so a single moved space never reads as a violation while an
 * injected character or format still does (it shows up as a text delta in #1 or a structural
 * delta in (1)/(3)). Returns null on success, else a precise reason.
 */
function diffRunSequence(inRuns: any[], outRuns: any[], paraPath: string): string | null {
  let i = 0; // next unconsumed input run
  let lastSourceRPrSig: string | null = null; // rPr-minus-vanish of the last consumed source run

  for (let j = 0; j < outRuns.length; j++) {
    const outR = outRuns[j];
    const inR = i < inRuns.length ? inRuns[i] : null;

    // (1) Primary image of the next input run: structurally equal modulo vanish AND modulo
    // <w:t> text (the source run of a split has its text trimmed; a plain run keeps it). We
    // compare modulo text so a trimmed source still matches; the trim itself is validated by
    // assertion #1, and any NON-text structural change (injected <w:b/>, changed attr) is
    // caught right here.
    if (inR && structuralDiffIgnoringWtText(inR, outR, `${paraPath}/run[${i}]`) === null) {
      lastSourceRPrSig = rPrSignatureMinusVanish(inR);
      i++;
      continue;
    }

    // (2) Permitted inserted bridge fragment.
    //   (a) a visible single-space run (boundary or interior space carrier), or
    //   (b) an "after" fragment: a text-only run whose rPr (minus vanish) matches the source
    //       run we just consumed — exactly what `exposeInteriorSpace` clones.
    if (isBridgeSpaceRun(outR)) continue;
    if (
      lastSourceRPrSig !== null &&
      isTextOnlyRun(outR) &&
      rPrSignatureMinusVanish(outR) === lastSourceRPrSig
    ) {
      continue;
    }

    // (3) Anything else is a disallowed run.
    if (inR) {
      // We have an unconsumed input run but this output run is neither its image nor a
      // permitted insertion → a structural change on that run (report its diff for clarity).
      const why = structuralDiffIgnoringWtText(inR, outR, `${paraPath}/run[${i}]`);
      return why ?? `${paraPath}/run[${i}]: run changed (${describe(outR)})`;
    }
    return `${paraPath}: output has an extra run that is not a bridge-space insertion (${describe(outR)})`;
  }

  if (i < inRuns.length) {
    return `${paraPath}: output is missing input run #${i} (${describe(inRuns[i])}) — a run was deleted or reordered`;
  }
  return null;
}

/** A short description of a run for error messages (its visible text, truncated). */
function describe(runEl: any): string {
  const t = runText(runEl);
  return `text="${t.length > 24 ? t.slice(0, 24) + "…" : t}"`;
}

/** True when a run's only content (besides an optional rPr) is `<w:t>` text — no field/break/drawing. */
function isTextOnlyRun(runEl: any): boolean {
  if (runEl.nodeName !== "w:r") return false;
  const kids = runEl.childNodes;
  if (!kids) return false;
  let sawT = false;
  for (let k = 0; k < kids.length; k++) {
    const c = kids.item(k);
    if (c.nodeType !== ELEMENT_NODE) continue;
    if (c.nodeName === "w:rPr") continue;
    if (c.nodeName === "w:t") {
      sawT = true;
      continue;
    }
    return false; // a tab/break/drawing/field → not a plain text run
  }
  return sawT;
}

/**
 * A stable signature of a run's `<w:rPr>` with every `<w:vanish>` removed — so two runs with
 * the "same formatting modulo hidden-ness" compare equal. Used to confirm a bridge "after"
 * fragment really did clone its source run's properties (and didn't smuggle a format change).
 * A run with no rPr (or a vanish-only rPr) signs as the empty string, equal to another such.
 */
function rPrSignatureMinusVanish(runEl: any): string {
  const rPr = firstDirectChild(runEl, "w:rPr");
  if (!rPr) return "";
  // Serialize the rPr's non-vanish element children in order, name+attrs deep — enough to
  // distinguish bold/italic/size/highlight/rStyle/etc. without depending on xmldom's
  // serializer (we build the signature from the parsed nodes directly).
  return childSignature(rPr);
}

/** Deep, order-sensitive signature of an element's non-vanish children (names + attrs + nested). */
function childSignature(el: any): string {
  const parts: string[] = [];
  const kids = el.childNodes;
  if (kids) {
    for (let i = 0; i < kids.length; i++) {
      const k = kids.item(i);
      if (k.nodeType !== ELEMENT_NODE) continue;
      if (el.nodeName === "w:rPr" && k.nodeName === "w:vanish") continue; // ignore hidden-ness
      parts.push(`${k.nodeName}[${attrSignature(k)}]{${childSignature(k)}}`);
    }
  }
  return parts.join(",");
}

/**
 * Deep structural equality of two element subtrees, IGNORING two things on either side:
 *   * any `<w:vanish>` direct child of a `<w:rPr>` (permitted delta (a) — done via
 *     `elementChildrenForCompare`, which also collapses a synthesized vanish-only rPr/pPr), and
 *   * the TEXT PAYLOAD of a `<w:t>` (a bridge split trims one space from the source run's
 *     `<w:t>`, and the run-level matcher relies on this to recognize a trimmed source as the
 *     same run; whole-paragraph text is policed separately by assertion #1).
 * This is the workhorse for the run matcher and the per-paragraph structure check: it compares
 * element name, attribute SET (order-insensitive, value-sensitive), and children in order, so
 * an injected `<w:b/>`, a changed `w:highlight`, a foreign run, or any non-text shape change is
 * still reported. Returns null on equality, else a human-readable reason.
 */
function structuralDiffIgnoringWtText(a: any, b: any, path: string): string | null {
  if (a.nodeType === ELEMENT_NODE && b.nodeType === ELEMENT_NODE) {
    if (a.nodeName !== b.nodeName) {
      return `${path}: element name differs (${a.nodeName} vs ${b.nodeName})`;
    }
    if (attrSignature(a) !== attrSignature(b)) {
      // A <w:t> may legitimately GAIN `xml:space="preserve"` when a bridge split trims it to a
      // leading/trailing-whitespace fragment (permitted delta (b)'s companion clause). When the strict
      // signatures differ ONLY by that one added attribute it is lossless, so accept it; otherwise (any
      // other attribute change, or any non-<w:t> element) the differing attributes are a real violation.
      const licensedPreserveAdd = a.nodeName === "w:t" && wtAttrsEqualModuloAddedPreserve(a, b);
      if (!licensedPreserveAdd) {
        return `${path}/${a.nodeName}: attributes differ (in=[${attrSignature(a)}] out=[${attrSignature(b)}])`;
      }
    }
    // Inside a <w:t>, skip the text-payload comparison (the moved space lives here).
    if (a.nodeName === "w:t") return null;
    const aKids = elementChildrenForCompare(a);
    const bKids = elementChildrenForCompare(b);
    if (aKids.length !== bKids.length) {
      return `${path}/${a.nodeName}: child count differs (in=${aKids.length} out=${bKids.length})`;
    }
    for (let i = 0; i < aKids.length; i++) {
      const r = structuralDiffIgnoringWtText(aKids[i], bKids[i], `${path}/${a.nodeName}`);
      if (r) return r;
    }
    return null;
  }
  if (a.nodeType === TEXT_NODE && b.nodeType === TEXT_NODE) {
    const at = a.textContent ?? "";
    const bt = b.textContent ?? "";
    return at === bt ? null : `${path}: text differs ("${at}" vs "${bt}")`;
  }
  if (a.nodeType !== b.nodeType) return `${path}: node kind differs (${a.nodeType} vs ${b.nodeType})`;
  return null;
}

/**
 * The paragraph-mark rPr (`<w:pPr><w:rPr>`) of a paragraph, or null. The mark's vanish is
 * the hidden-paragraph-mark signal (condensed view); it is a permitted delta (a) exactly
 * like a run's vanish, and is covered automatically because the whole-`<w:p>` structural
 * diff below filters `<w:vanish>` out of every `<w:rPr>` — including the mark's.
 */

/**
 * Compare one paragraph pair end to end. Assertion #1 (whole-paragraph text identical) runs
 * FIRST — it is the strongest, simplest losslessness check and catches the largest class of
 * bugs (changed/added/dropped characters) immediately. Then the run sequence is aligned
 * (assertions #2/#3, allowing bridge insertions + moved spaces). Finally the paragraph's
 * NON-run structure (pPr, table cell wrappers, the element shape around runs) is checked by
 * a whole-`<w:p>` structural diff that filters out the runs themselves (already compared)
 * and `<w:vanish>` (permitted) — catching any change to pStyle, outline level, numbering,
 * or table nesting (assertion #4, per-paragraph half).
 */
function diffParagraph(inP: any, outP: any, index: number): void {
  const path = `paragraph[${index}]`;

  // #1 — concatenated visible text byte-identical (the true 002-F1 text-preservation check).
  const inText = paragraphVisibleText(inP);
  const outText = paragraphVisibleText(outP);
  if (inText !== outText) {
    throw new Error(
      `[semanticDiff] ${path}: concatenated visible run text changed — this violates 002-F1 ` +
        `(text must be byte-identical). in=${JSON.stringify(inText)} out=${JSON.stringify(outText)}`
    );
  }

  // #2/#3 — run order + rPr contents, allowing bridge insertions and one moved space.
  const reasonRuns = diffRunSequence(paragraphRuns(inP), paragraphRuns(outP), path);
  if (reasonRuns) {
    throw new Error(`[semanticDiff] ${reasonRuns} — only <w:vanish> toggles and bridge-space insertions are permitted`);
  }

  // #4 (per-paragraph) — the paragraph's NON-run structure (pPr/pStyle/outline/table shape).
  // We compare the whole <w:p> structurally but with every <w:r> collapsed to a placeholder
  // so the run-level moved-space/insertion deltas (already validated above) don't read as a
  // structural change here, while ANY change to the surrounding markup does.
  const reasonStruct = structuralDiffCollapsingRuns(inP, outP, path);
  if (reasonStruct) {
    throw new Error(
      `[semanticDiff] ${reasonStruct} — paragraph structure (pPr/style/table nesting) must be identical`
    );
  }
}

/**
 * Whole-paragraph structural diff that DROPS every `<w:r>` (run internals + order are
 * validated by `diffRunSequence`, and bridge splits legitimately change the run count) and
 * compares EVERYTHING ELSE — pPr, pStyle, outlineLvl, numPr, sectPr, table-cell wrappers,
 * hyperlink wrappers' own attributes — exactly, modulo `<w:vanish>` (and modulo the
 * synthesized vanish-only rPr/pPr the mark-hide path creates). This catches a moved /
 * added / removed pStyle, a changed outline level, a paragraph that gained or lost a
 * table-cell ancestor, etc. Returns null on equality, else a reason.
 */
function structuralDiffCollapsingRuns(a: any, b: any, path: string): string | null {
  if (a.nodeType === ELEMENT_NODE && b.nodeType === ELEMENT_NODE) {
    if (a.nodeName !== b.nodeName) return `${path}: element name differs (${a.nodeName} vs ${b.nodeName})`;
    if (attrSignature(a) !== attrSignature(b)) {
      return `${path}/${a.nodeName}: attributes differ (in=[${attrSignature(a)}] out=[${attrSignature(b)}])`;
    }
    // Drop runs ENTIRELY for this comparison: bridge splits legitimately change the run
    // count (and positions) of a <w:p>, and run internals/order are already validated by
    // `diffRunSequence`. What remains — pPr, pStyle, numPr, outlineLvl, sectPr, hyperlink
    // wrappers (their own attrs), table-cell ancestors — must match positionally and exactly.
    // Comparing only NON-run children means an added/removed bridge run never reads as a
    // structural delta here, while a changed pStyle / lost table nesting still does.
    const aKids = nonRunChildrenForCompare(a);
    const bKids = nonRunChildrenForCompare(b);
    if (aKids.length !== bKids.length) {
      return `${path}/${a.nodeName}: non-run child count differs (in=${aKids.length} out=${bKids.length})`;
    }
    for (let i = 0; i < aKids.length; i++) {
      const r = structuralDiffCollapsingRuns(aKids[i], bKids[i], `${path}/${a.nodeName}`);
      if (r) return r;
    }
    return null;
  }
  if (a.nodeType === TEXT_NODE && b.nodeType === TEXT_NODE) {
    const at = a.textContent ?? "";
    const bt = b.textContent ?? "";
    return at === bt ? null : `${path}: text differs ("${at}" vs "${bt}")`;
  }
  if (a.nodeType !== b.nodeType) return `${path}: node kind differs (${a.nodeType} vs ${b.nodeType})`;
  return null;
}

/**
 * Children for the run-collapsing structural diff: every `<w:r>` is DROPPED (its internals and
 * order are validated by `diffRunSequence`, and bridge splits change its count legitimately),
 * `<w:vanish>` inside an `<w:rPr>` is dropped (permitted), a synthesized vanish-only rPr and a
 * mark-vanish-only pPr are dropped (permitted delta (a) for hiding the mark on a markless
 * paragraph), and whitespace-only inter-element text is skipped (xmldom artifact). Everything
 * else — pPr, sectPr, hyperlink/smartTag wrappers, table-cell ancestors — is kept and compared
 * positionally, so a structural change to the markup AROUND the runs is still caught.
 */
function nonRunChildrenForCompare(el: any): any[] {
  const dropVanish = el.nodeName === "w:rPr";
  const out: any[] = [];
  const kids = el.childNodes;
  if (!kids) return out;
  for (let i = 0; i < kids.length; i++) {
    const k = kids.item(i);
    if (k.nodeType === ELEMENT_NODE) {
      if (dropVanish && k.nodeName === "w:vanish") continue;
      if (k.nodeName === "w:rPr" && isVanishOnlyRPr(k)) continue;
      if (k.nodeName === "w:pPr" && isMarkVanishOnlyPPr(k)) continue;
      if (k.nodeName === "w:r") continue; // runs validated by diffRunSequence
      out.push(k);
    } else if (k.nodeType === TEXT_NODE) {
      const t = k.textContent ?? "";
      if (t.trim().length > 0) out.push(k);
    }
  }
  return out;
}

// ---------------------------------------------------------------------------
// Public API
// ---------------------------------------------------------------------------

/**
 * Assert that `outputPackageXml` differs from `inputPackageXml` ONLY by the permitted
 * vanish + bridge-split deltas, across EVERY body-story paragraph. Throws (with a precise
 * paragraph-index + offending-node message) on any other difference, including a changed
 * paragraph count. This is the whole-body R1 losslessness gate (002-S1 / 002-F1): pass it
 * to a committed package vs the package the engine read.
 */
export function assertVanishBridgeOnlyDelta(inputPackageXml: string, outputPackageXml: string): void {
  const inDoc = parse(inputPackageXml);
  const outDoc = parse(outputPackageXml);
  const inParas = storyParagraphs(inDoc);
  const outParas = storyParagraphs(outDoc);

  // #4 (whole-document half) — paragraph COUNT identical. A hide never adds or removes a
  // paragraph; a count change means the splice corrupted the body (e.g. dropped a <w:p>).
  if (inParas.length !== outParas.length) {
    throw new Error(
      `[semanticDiff] body paragraph count changed: in=${inParas.length} out=${outParas.length} — ` +
        `a hide must never add or remove a paragraph (002-F1)`
    );
  }
  for (let i = 0; i < inParas.length; i++) {
    diffParagraph(inParas[i], outParas[i], i);
  }
}

/**
 * The single-paragraph variant: assert one paragraph's OOXML (a bare `<w:p>` fragment or a
 * one-paragraph package — anything ooxml.ts `firstParagraph`/`bodyScope` accepts) differs
 * only by the permitted deltas. For the per-paragraph commit path (002-S2's compat proof)
 * and for focused tests. Throws on any disallowed difference; throws if either side has no
 * body paragraph (an empty diff would otherwise vacuously "pass").
 */
export function assertVanishBridgeOnlyDeltaPara(inputParaXml: string, outputParaXml: string): void {
  const inP = firstStoryParagraph(parse(inputParaXml));
  const outP = firstStoryParagraph(parse(outputParaXml));
  if (!inP || !outP) {
    throw new Error(
      `[semanticDiff] assertVanishBridgeOnlyDeltaPara: a side had no <w:p> ` +
        `(in=${inP ? "present" : "absent"} out=${outP ? "present" : "absent"}) — nothing to compare`
    );
  }
  diffParagraph(inP, outP, 0);
}

/** The first story paragraph of a fragment/package, or null (mirrors ooxml.ts `firstParagraph`). */
function firstStoryParagraph(doc: any): any | null {
  const paras = storyParagraphs(doc);
  return paras.length > 0 ? paras[0] : null;
}
