// Pure, host-free logic for the WHOLE-BODY OOXML strategy and for outline-level
// normalization. Everything here is string-in / DOM / string-out and runs in Node
// tests and the task-pane browser via @xmldom/xmldom — there is no Office.js in
// this file. The adapter (officeWordPort.ts) is the only thing that turns these
// pure transforms into `Word.run` calls.
//
// Two responsibilities:
//
//   1. OUTLINE-LEVEL NORMALIZATION. The engine's keep rule keys on a canonical
//      0-based outline level (0 = Heading 1 … 8 = Heading 9, null = body). A paragraph
//      exposes its level ONLY as the NUMBER `Paragraph.outlineLevel` (WordApi 1.1):
//      Heading 1 = 1 … Heading 9 = 9, body text = 10. (The string-enum form
//      `ParagraphFormat.outlineLevel` lives on `Word.Style`, NOT on a paragraph, so it
//      cannot be read off a paragraph — see LESSONS #12.) The numeric base is not
//      officially documented; community/VBA parity says 1-based, so we normalize with a
//      configurable base (default `oneBased`) and the adapter logs raw values so a wrong
//      base surfaces in diagnostics rather than silently mis-hiding.
//
//   2. WHOLE-BODY SPLICE. `body.getOoxml()` returns a flat-OPC package wrapping the
//      whole document story; we split it into per-`<w:p>` fragments for the engine
//      and splice the engine's edited fragments back into the SAME parsed DOM, so
//      only `<w:vanish/>` ever changes (byte-preserving except the hide flag).
//      `WholeBodyPackage` caches that parsed DOM between the adapter's read and
//      write so the two halves operate on one tree.

import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { parseStyleDefs, outlineNumberFromProps, type StyleDef } from "./outline";
import { ParsedParagraph } from "./ooxml";

// xmldom's node types vary across versions; we keep public signatures fully typed
// (string in/out, number/null) and use a localized `any` for node handles — the
// same pragmatic choice ooxml.ts makes.
/* eslint-disable @typescript-eslint/no-explicit-any */

// Namespaces / content types used to wrap a single paragraph in a host-valid flat-OPC
// package (see WholeBodyPackage.paragraphXml). `W_NS` is the WordprocessingML main
// namespace (the only prefix the engine inspects); `PKG_NS` is the flat-OPC package
// namespace; `RELS_NS` is the package-relationships namespace; the content types are what
// Word's own getOoxml stamps on the document and relationships parts.
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
const WML_DOCUMENT_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const RELS_CONTENT_TYPE = "application/vnd.openxmlformats-package.relationships+xml";
// The OPC "start part" relationship type: identifies which part is the main document.
const OFFICE_DOCUMENT_REL_TYPE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument";

/**
 * The package-level relationships part (`/_rels/.rels`) — REQUIRED in any flat-OPC package
 * handed to the live host's `insertOoxml`. It declares which part is the main document via the
 * single `officeDocument` relationship to `/word/document.xml`. Omitting it (as the original
 * `paragraphXml` minimal wrapper did) leaves the host unable to identify the main document part,
 * which it rejects with `GeneralException` at `Paragraph.insertOoxml` — the Stage 4.2 commit
 * regression. Word's own `body.getOoxml()` always includes this; the old per-paragraph fallback
 * only worked because it re-inserted Word's COMPLETE package. The Id/Type/Target are fixed OPC
 * constants (and `/_rels/.rels` IDs are scoped to the package root, independent of any
 * `document.xml.rels` rId the paragraph references), so we emit it verbatim. `pkg:padding="512"`
 * mirrors Word's own emission (a serialization hint; harmless). Verified against learn.microsoft.com
 * "Use Office Open XML (OOXML) in Word add-ins" — the minimal formatted-text package.
 */
const PACKAGE_RELS_PART =
  `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS_CONTENT_TYPE}" pkg:padding="512">` +
  `<pkg:xmlData><Relationships xmlns="${RELS_NS}">` +
  `<Relationship Id="rId1" Type="${OFFICE_DOCUMENT_REL_TYPE}" Target="word/document.xml"/>` +
  `</Relationships></pkg:xmlData></pkg:part>`;
const ELEMENT_NODE = 1;

// The base relationship-type namespace shared by the document-part relationships below.
const OFFICE_REL_BASE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

/**
 * The "auxiliary" document parts a paragraph can INHERIT its rendering from, keyed by the
 * `pkg:contentType` Word stamps on each. `commitXml` bundles these (verbatim) into the
 * per-paragraph commit package so that any formatting resolved through a referenced style
 * (`<w:pStyle>`/`<w:rStyle>`/`docDefaults` → underline, character border/box, font size) or
 * through list numbering / theme fonts survives `insertOoxml`. WITHOUT them the host renders
 * every style-derived property as a document DEFAULT (the "underline/box/18pt collapse to plain
 * 11pt, inline highlight survives" wet-test bug) — because the minimal `paragraphXml` package
 * carries no styles part. `settings.xml` is deliberately EXCLUDED: it holds document-level state
 * (track-changes, document protection, rsids) that an inserted fragment must never re-impose.
 */
const AUX_PART_CONTENT_TYPES = new Set<string>([
  "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml",
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml",
  "application/vnd.openxmlformats-officedocument.theme+xml"
]);

/**
 * The relationship TYPES (in `/word/_rels/document.xml.rels`) that link `document.xml` to each
 * aux part above. We copy the matching `<Relationship>` entries verbatim so the host can FIND
 * each part (a body references a style by its styleId string and a list by `numId`, NOT by an
 * `r:id`, so the part is located purely through these typed relationships). Their Ids come from
 * the same source `document.xml.rels` as any hyperlink relationship, so they never collide.
 */
const AUX_REL_TYPES = new Set<string>([
  `${OFFICE_REL_BASE}/styles`,
  `${OFFICE_REL_BASE}/numbering`,
  `${OFFICE_REL_BASE}/theme`
]);

// ---------------------------------------------------------------------------
// Outline-level normalization
// ---------------------------------------------------------------------------

/**
 * How to interpret the numeric `Paragraph.outlineLevel`. `oneBased` is the
 * community/VBA-verified convention (Heading 1 = 1 … 9 = 9, body ≥ 10); `zeroBased`
 * is provided as an escape hatch if a host is ever observed to differ. The adapter
 * defaults to `oneBased` and surfaces the raw values for empirical confirmation.
 */
export type OutlineNumberBase = "oneBased" | "zeroBased";

/**
 * Normalize the NUMBER form (`Paragraph.outlineLevel`) to the canonical 0-based
 * level. With the default `oneBased` convention a value of 1–9 maps to 0–8 and
 * everything else (0, ≥10 body text, NaN, null) is body → null. `zeroBased` maps
 * 0–8 directly. Always total and side-effect-free.
 */
export function normalizeOutlineNumber(
  level: number | null | undefined,
  base: OutlineNumberBase = "oneBased"
): number | null {
  if (level == null || !Number.isFinite(level)) return null;
  const n = Math.trunc(level);
  if (base === "oneBased") {
    return n >= 1 && n <= 9 ? n - 1 : null;
  }
  return n >= 0 && n <= 8 ? n : null;
}

// ---------------------------------------------------------------------------
// Parsing / scoping (mirrors ooxml.ts's rationale; kept local to stay self-
// contained — both are unit-tested, and the divergence risk is a 5-line scope.)
// ---------------------------------------------------------------------------

/** Parse, throwing only on fatal XML so the adapter can surface a clear error. */
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
 * The element to search paragraphs within: the first `<w:body>` (the document
 * story) when present, else the document root (bare `<w:p>` fragments in tests).
 * Scoping to `<w:body>` is the C1 fix — a flat-OPC package can carry header/footer
 * `<w:p>` that must never be touched.
 */
function bodyScope(doc: any): any {
  const bodies = doc.getElementsByTagName("w:body");
  // Fall back to the Document node (NOT documentElement): `getElementsByTagName`
  // only returns descendants, and the Document includes the root `<w:p>` of a bare
  // fragment as a descendant, whereas `documentElement` (the `<w:p>` itself) would
  // report zero. This mirrors ooxml.ts's `firstParagraph` scoping exactly.
  return bodies && bodies.length > 0 ? bodies.item(0) : doc;
}

/** All `<w:p>` within a scope, snapshotted to a static array in document order. */
function paragraphsIn(scope: any): any[] {
  const live = scope.getElementsByTagName("w:p");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) out.push(live.item(i));
  return out;
}

/**
 * True when a node lives inside a text box (`<w:txbxContent>`). Such paragraphs are
 * NOT part of the body STORY: `body.getOoxml()` serializes them as `<w:p>` descendants
 * of `<w:body>`, but Word's `body.paragraphs` collection excludes them — the exact
 * source of the whole-body alignment mismatch (Stage 4.1: a textbox inflates the OOXML
 * `<w:p>` count above the proxy count). We walk ancestors rather than match a fixed
 * depth because a textbox can sit inside a run inside a paragraph at any nesting.
 */
function isInTextbox(node: any): boolean {
  let n = node ? node.parentNode : null;
  while (n) {
    if (n.nodeName === "w:txbxContent") return true;
    n = n.parentNode;
  }
  return false;
}

/**
 * Visible text of a paragraph node in document order, mirroring Word's `Paragraph.text`
 * closely enough for tolerant alignment: `<w:t>` content, `<w:tab>`→"\t", `<w:br>`/
 * `<w:cr>`→"\n", and EXCLUDING text nested in a text box (`<w:txbxContent>`). The caller
 * whitespace-normalizes the result, so rendering tabs/breaks as whitespace (rather than
 * dropping them) is what stops a cite's tab from spuriously failing alignment → a needless
 * whole-document fallback to the slow per-paragraph read.
 */
function collectParagraphText(node: any): string {
  let out = "";
  const walk = (n: any): void => {
    const name = n.nodeName;
    if (name === "w:txbxContent") return; // text box: not part of this paragraph's text
    if (name === "w:t") {
      out += n.textContent ?? "";
      return;
    }
    if (name === "w:tab") {
      out += "\t";
      return;
    }
    if (name === "w:br" || name === "w:cr") {
      out += "\n";
      return;
    }
    const kids = n.childNodes;
    if (kids) for (let k = 0; k < kids.length; k++) walk(kids.item ? kids.item(k) : kids[k]);
  };
  walk(node);
  return out;
}

/**
 * STORY paragraphs only — every `<w:p>` in the scope EXCEPT those nested in a text box
 * (`<w:txbxContent>`). This is what aligns 1:1 with Office.js `body.paragraphs`, so the
 * whole-body splice index maps correctly (① of the Stage 4.1 fix). Document order.
 */
function storyParagraphsIn(scope: any): any[] {
  const live = scope.getElementsByTagName("w:p");
  const out: any[] = [];
  for (let i = 0; i < live.length; i++) {
    const p = live.item(i);
    if (p && !isInTextbox(p)) out.push(p);
  }
  return out;
}

/**
 * Count the `<w:p>` in a fragment's body scope. Used by the adapter's whole-body
 * ALIGNMENT GUARD: if this disagrees with Office.js's `paragraphs.items.length`,
 * the `<w:p>`↔proxy index mapping is unsafe and the adapter falls back.
 */
export function countBodyParagraphs(packageXml: string): number {
  return storyParagraphsIn(bodyScope(parse(packageXml))).length;
}

/**
 * GUARD CORE — return the single body `<w:p>` of an ALREADY-PARSED writeback
 * fragment, throwing the audit-#5 guard error when the count is not exactly one.
 * Node-in (not string-in) so a caller that needs the paragraph node anyway
 * (`WholeBodyPackage.replace`) can guard and select on ONE parsed tree: the
 * string-in `assertSingleParagraph` + a second `parse()` used to cost a full
 * extra DOMParser pass per spliced paragraph on the DEFAULT whole-body hide
 * commit and on cite repair. `__tests__/parseCount.test.ts` locks the fusion.
 */
function singleBodyParagraph(fragDoc: any, context: string): any {
  const paras = paragraphsIn(bodyScope(fragDoc));
  if (paras.length !== 1) {
    throw new Error(
      `Rostrum ${context} guard: expected exactly one <w:p> in the fragment, found ${paras.length}. ` +
        `Refusing to write back potentially corrupted OOXML.`
    );
  }
  return paras[0];
}

/**
 * Assert a writeback fragment contains EXACTLY ONE body `<w:p>` (audit #5). The
 * engine only ever edits a single paragraph's runs; a fragment carrying zero or
 * many `<w:p>` means something upstream corrupted the OOXML, and writing it back
 * could splice a table or empty the paragraph. Throws a clear, contextual error.
 * String-in convenience over `singleBodyParagraph` for callers that never parse
 * the fragment themselves (the per-paragraph commit path) — same error text.
 */
export function assertSingleParagraph(ooxml: string, context = "writeback"): void {
  singleBodyParagraph(parse(ooxml), context);
}

/**
 * Normalize a single paragraph's host OOXML to exactly ONE body paragraph.
 *
 * `Paragraph.getOoxml()` / `Range.getOoxml()` can return a flat-OPC package whose
 * `<w:body>` holds the target paragraph FOLLOWED BY a trailing empty `<w:p>` — Word
 * emits one for the paragraph mark when a range includes it (observed on the live
 * host). That 2-paragraph fragment trips the single-`<w:p>` write guard. This keeps
 * the FIRST direct-child `<w:p>` of the body scope and drops any later direct-child
 * paragraphs, leaving the wrapper, namespaces, and `<w:sectPr>` intact so the result
 * is still valid for `insertOoxml`. A fragment that already holds one paragraph (or a
 * bare `<w:p>`) is returned UNCHANGED (byte-identical) — no reparse/reserialize.
 */
export function keepFirstBodyParagraph(packageXml: string): string {
  const doc = parse(packageXml);
  const scope = bodyScope(doc);
  const directParas: any[] = [];
  const kids = scope.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const n = kids.item ? kids.item(i) : kids[i];
    if (n && n.nodeName === "w:p") directParas.push(n);
  }
  if (directParas.length <= 1) return packageXml; // already single-paragraph; untouched
  for (let i = 1; i < directParas.length; i++) {
    const parent = directParas[i].parentNode;
    if (parent) parent.removeChild(directParas[i]);
  }
  return serialize(doc);
}

/** Escape a value for use inside a double-quoted XML attribute. */
function escapeAttr(v: string): string {
  return v.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

/**
 * The REAL `<w:document>` start-tag attributes, verbatim, for re-wrapping a single paragraph
 * (see `paragraphXml`). We copy the host's own `<w:document>` element attributes — every
 * `xmlns:*` decl AND `mc:Ignorable` (which the MC processor needs for `mc:AlternateContent`
 * emoji to resolve) — so the per-paragraph fragment is exactly the namespace/MC context Word
 * itself emitted. Scoped to the `<w:document>` element ONLY (NOT the whole package) so the
 * package-relationships default `xmlns` and `pkg:` prefix never leak onto `<w:document>`.
 * Falls back to a bare `xmlns:w` when there is no `<w:document>` (bare-`<w:p>` test fixtures).
 */
function documentAttrs(doc: any): string {
  const docs = doc.getElementsByTagName("w:document");
  const src = docs && docs.length > 0 ? docs.item(0) : doc.documentElement;
  const out: string[] = [];
  let hasW = false;
  const attrs = src ? src.attributes : null;
  if (attrs) {
    for (let i = 0; i < attrs.length; i++) {
      const a = attrs.item ? attrs.item(i) : attrs[i];
      const name: string = a && (a.name ?? a.nodeName);
      if (!name) continue;
      if (name === "xmlns:w") hasW = true;
      out.push(`${name}="${escapeAttr(a.value ?? a.nodeValue ?? "")}"`);
    }
  }
  if (!hasW) out.unshift(`xmlns:w="${W_NS}"`); // the one prefix the engine always needs
  return out.join(" ");
}

/**
 * Map of relationship id → serialized `<Relationship/>` from the source package's `.rels`
 * parts. A hyperlink/image run references a relationship by `r:id`/`r:embed`/`r:link`; the
 * relationship itself lives in `/word/_rels/document.xml.rels`, which `body.getOoxml()`
 * includes. `paragraphXml` re-emits ONLY the relationships a given paragraph references, so a
 * fragment with a hyperlink isn't a DANGLING reference (Word's `insertOoxml` rejects a flat-OPC
 * whose document.xml cites an `r:id` with no matching relationship — the live "problem with its
 * contents" failure). Keyed by Id; value is the element's own XML. (Stage 4.2 fix, audit C1.)
 */
function collectRelationships(doc: any): Map<string, string> {
  const map = new Map<string, string>();
  const rels = doc.getElementsByTagName("Relationship");
  for (let i = 0; i < rels.length; i++) {
    const r = rels.item ? rels.item(i) : rels[i];
    const id = r && r.getAttribute ? r.getAttribute("Id") : null;
    if (id) map.set(id, serialize(r));
  }
  return map;
}

/** Relationship ids referenced (via any `r:*` attribute) anywhere inside a paragraph node. */
function referencedRelIds(node: any): Set<string> {
  const ids = new Set<string>();
  const visit = (el: any): void => {
    const attrs = el.attributes;
    if (attrs) {
      for (let i = 0; i < attrs.length; i++) {
        const a = attrs.item ? attrs.item(i) : attrs[i];
        const name: string = a && (a.name ?? a.nodeName);
        if (name && name.startsWith("r:")) {
          const v = a.value ?? a.nodeValue;
          if (v) ids.add(v);
        }
      }
    }
    const kids = el.childNodes;
    if (kids) {
      for (let i = 0; i < kids.length; i++) {
        const k = kids.item ? kids.item(i) : kids[i];
        if (k && k.nodeType === ELEMENT_NODE) visit(k);
      }
    }
  };
  visit(node);
  return ids;
}

/**
 * The concatenated `<pkg:part>` XML (verbatim) for every aux part present in the source package
 * — styles / numbering / theme (see `AUX_PART_CONTENT_TYPES`). Captured ONCE per package because
 * it is identical for every paragraph; `commitXml` prepends it so style/numbering/theme references
 * resolve on the host. Each part keeps its own inner xmlns decls (Word emits self-contained parts);
 * the `pkg:`/`pkg:xmlData` prefixes resolve against the wrapper `commitXml` re-declares.
 */
function collectAuxParts(doc: any): string {
  const parts = doc.getElementsByTagName("pkg:part");
  let out = "";
  for (let i = 0; i < parts.length; i++) {
    const part = parts.item ? parts.item(i) : parts[i];
    const ct = part && part.getAttribute ? part.getAttribute("pkg:contentType") : null;
    if (ct && AUX_PART_CONTENT_TYPES.has(ct)) out += serialize(part);
  }
  return out;
}

/**
 * The concatenated `<Relationship>` XML (verbatim) for the aux parts — the entries in
 * `/word/_rels/document.xml.rels` whose `Type` is a styles/numbering/theme relationship
 * (see `AUX_REL_TYPES`). `commitXml` emits these in the fragment's `document.xml.rels` so the
 * host can locate each aux part. Mirrors how `collectRelationships` serializes a `<Relationship>`.
 */
function collectAuxRels(doc: any): string {
  const rels = doc.getElementsByTagName("Relationship");
  let out = "";
  for (let i = 0; i < rels.length; i++) {
    const r = rels.item ? rels.item(i) : rels[i];
    const type = r && r.getAttribute ? r.getAttribute("Type") : null;
    if (type && AUX_REL_TYPES.has(type)) out += serialize(r);
  }
  return out;
}

// ---------------------------------------------------------------------------
// P4-REVISED — string-level flat-OPC part segmentation (Loop 002 A2).
//
// WHY (perf-1 / PLAN §2 A2). The ctor used to DOM-parse the WHOLE flat-OPC package — the
// document part PLUS the styles part (≈1.27MB on a real debate doc), numbering, theme, rels,
// settings — into one xmldom tree, even though the node-direct Hide only ever MUTATES the
// document part. That whole-package parse is the dominant read-side ctor cost. A2 cuts the
// package at `<pkg:part>` boundaries with a linear string scan, DOM-parses ONLY the document
// part (so its `<w:p>` nodes can still be mutated in place + re-serialized — the CORE-3
// contract), and keeps every other part as a VERBATIM substring. `serialize()` stitches the
// captured substrings back in ORIGINAL ORDER with the document part swapped for the
// re-serialized (mutation-bearing) `<w:document>` subtree.
//
// WHY THIS IS BYTE-IDENTICAL TO TODAY (002-S2/F2). The synthetic fixtures are authored in
// xmldom's canonical round-trip form, so `serialize(parse(fixture)) === fixture` and each
// part's verbatim substring already EQUALS xmldom's standalone serialization of that part.
// Re-serializing only the `<w:document>` subtree reproduces its bytes exactly (xmldom never
// walks above a prefixed start node — the same property `paragraphXml` relies on). So the
// stitched output equals today's whole-DOM `serialize()` byte-for-byte. On REAL docs (no byte
// control) keeping non-document parts verbatim is a FIDELITY IMPROVEMENT over xmldom
// re-serialization, and the document part still DOM-reserializes identically to today.
// ---------------------------------------------------------------------------

/** `<w:t>` content is entity-escaped, so a literal `</pkg:part>` can never appear inside text. */
const PART_OPEN = "<pkg:part";
const PART_CLOSE = "</pkg:part>";

/** One captured `<pkg:part>` of a flat-OPC package — its verbatim bytes plus the pieces A2 needs. */
interface PackagePart {
  /** `pkg:name` attribute value (e.g. `/word/document.xml`), or "" if absent. */
  name: string;
  /** `pkg:contentType` attribute value, or "" if absent. */
  contentType: string;
  /** The part's ENTIRE `<pkg:part …>…</pkg:part>` substring, verbatim from the input. */
  full: string;
  /** The inner content of this part's `<pkg:xmlData>` (the actual part XML), or "" if none. */
  xmlData: string;
}

/**
 * A flat-OPC package cut into its `<pkg:package>` head, ordered parts, and tail — purely by
 * string scanning, no DOM parse. `head` is everything up to (and including) the package
 * element's open tag; `tail` is `</pkg:package>` plus any trailing bytes. The `between` array
 * holds the bytes between consecutive parts (empty for the canonical fixtures, but captured so
 * any whitespace/XML-decl in a real package round-trips verbatim). `null` when the input is NOT
 * a flat-OPC package (a raw `document.xml` or a bare `<w:p>` fragment) — those degrade to the
 * pre-segmentation whole-input DOM parse, unchanged.
 */
interface SegmentedPackage {
  head: string;
  parts: PackagePart[];
  /** `between[k]` = bytes after parts[k-1] and before parts[k]; `between[0]` = bytes after head. */
  between: string[];
  tail: string;
}

/** Read a quoted attribute value from a `<pkg:part …>` open tag substring (returns "" if absent). */
function attrValue(openTag: string, attr: string): string {
  const m = new RegExp(`\\b${attr}="([^"]*)"`).exec(openTag);
  return m ? m[1] : "";
}

/**
 * Cut a flat-OPC package string into `{ head, parts, between, tail }` with a single linear scan.
 * `<pkg:part>` is never nested and `w:t` text is entity-escaped, so matching `<pkg:part`→
 * `</pkg:part>` linearly is unambiguous. Returns null when the string is not a `<pkg:package>`
 * (a raw document.xml / bare fragment), signalling the caller to fall back to a whole-input parse.
 */
function segmentPackage(packageXml: string): SegmentedPackage | null {
  const pkgOpen = packageXml.indexOf("<pkg:package");
  if (pkgOpen < 0) return null; // not a flat-OPC package — caller parses the whole input
  // End of the `<pkg:package …>` open tag. Attribute values are quoted and can't contain a raw
  // `>` (XML rule), so the first `>` after the tag name closes the open tag.
  const headEnd = packageXml.indexOf(">", pkgOpen);
  if (headEnd < 0) return null;
  const head = packageXml.slice(0, headEnd + 1);

  const parts: PackagePart[] = [];
  const between: string[] = [];
  let cursor = headEnd + 1;
  for (;;) {
    const open = packageXml.indexOf(PART_OPEN, cursor);
    if (open < 0) break;
    between.push(packageXml.slice(cursor, open)); // bytes since the previous part (or head)
    const close = packageXml.indexOf(PART_CLOSE, open);
    if (close < 0) return null; // malformed — let the whole-input parse surface the error
    const fullEnd = close + PART_CLOSE.length;
    const full = packageXml.slice(open, fullEnd);
    const openTagEnd = full.indexOf(">");
    const openTag = openTagEnd >= 0 ? full.slice(0, openTagEnd + 1) : full;
    parts.push({
      name: attrValue(openTag, "pkg:name"),
      contentType: attrValue(openTag, "pkg:contentType"),
      full,
      xmlData: extractXmlData(full)
    });
    cursor = fullEnd;
  }
  // Nothing matched `<pkg:part>` despite a `<pkg:package>` — treat as non-segmentable so the
  // caller's whole-input parse keeps today's behavior (and surfaces any real malformation).
  if (parts.length === 0) return null;
  const tail = packageXml.slice(cursor);
  return { head, parts, between, tail };
}

/** The inner content of a part's `<pkg:xmlData>…</pkg:xmlData>` (the part's own XML), or "". */
function extractXmlData(partXml: string): string {
  const open = partXml.indexOf("<pkg:xmlData");
  if (open < 0) return "";
  const openEnd = partXml.indexOf(">", open);
  const close = partXml.lastIndexOf("</pkg:xmlData>");
  if (openEnd < 0 || close < 0 || close < openEnd) return "";
  return partXml.slice(openEnd + 1, close);
}

// ---------------------------------------------------------------------------
// P7 — referenced-closure styles prune (Loop 002 B3, DEFAULT-OFF runtime flag).
//
// WHY (PLAN §2 B3 / perf-1 / CASES 002-S6/F6). Word's `styles.xml` carries the FULL latent/default
// style table (~2.15MB on a real ndca doc) regardless of how few styles the body actually uses. The
// node-direct Hide re-serializes the whole package back to the host, so every byte of that styles
// part is on the WRITE side. When the prune flag is ON, we shrink the styles part to the transitive
// closure of the styles REACHABLE from the retained content — a write-side size win — while keeping
// the document body, `docDefaults`, `latentStyles`, and the `<w:styles>` root verbatim so the
// rendered formatting is unchanged. When the flag is OFF (the shipped default) NONE of this runs and
// the part stays byte-identical to today (002-F6).
//
// CLOSED-WORLD CORRECTNESS (loss-5). The closure is computed by SEEDING from every content part
// (document.xml + numbering backlinks + footnotes/headers/footers, post-splice) plus all `w:default`
// styles, then EXPANDING by a transitive fixpoint over every style cross-reference
// (basedOn / link both directions / next / numStyleLink / styleLink). A retained style therefore
// never dangles: any id it points at is, by construction, also in the closure.
//
// REGEX, not DOM, for the heavy part — consistent with `parseStyleDefs`/`outline.ts` (the styles
// part is flat and well-formed as Word emits it). Seeds scan the (small) content substrings the same
// way. The ONE structural operation — removing whole `<w:style>` elements — is done by re-emitting
// only the retained `<w:style>…</w:style>` spans and the verbatim non-style remainder, so docDefaults
// / latentStyles / root attributes are preserved byte-for-byte.
// ---------------------------------------------------------------------------

/** One style's outgoing cross-references, parsed from its `<w:style>` inner XML (for the closure walk). */
interface StyleRefs {
  /** `<w:basedOn w:val>` — the parent this style inherits from. */
  basedOn: string | null;
  /** `<w:link w:val>` — the paired para/char style. Retained in BOTH directions (a linked pair). */
  link: string | null;
  /** `<w:next w:val>` — the style applied to the NEXT paragraph after one in this style. */
  next: string | null;
  /** `<w:numStyleLink w:val>` — points a numbering style at the style that owns the real numbering. */
  numStyleLink: string | null;
  /** `<w:styleLink w:val>` — the reverse of numStyleLink (the numbering-owner points back). */
  styleLink: string | null;
}

/** Read the FIRST `<w:TAG … w:val="…">` value inside a `<w:style>` inner XML, or null when absent. */
function styleRefVal(inner: string, tag: string): string | null {
  // Attribute-order tolerant (mirrors outline.ts parseStyleDefs C-2): never assume `w:val` is first.
  const m = new RegExp(`<w:${tag}\\b[^>]*\\bw:val="([^"]+)"`).exec(inner);
  return m ? m[1] : null;
}

/**
 * Parse the styles part into `styleId → StyleRefs` plus the set of `w:default="1"` styleIds. Regex over
 * each `<w:style>` element (same matcher shape as `parseStyleDefs`), so it sees exactly the styles the
 * resolver sees. A style with no cross-references still gets an (all-null) entry so the closure walk can
 * terminate cleanly at it.
 */
function parseStyleGraph(stylesXml: string): {
  refs: Map<string, StyleRefs>;
  defaults: Set<string>;
} {
  const refs = new Map<string, StyleRefs>();
  const defaults = new Set<string>();
  // Capture the style's open tag (group 1) separately from its inner XML (group 2): `w:default` lives
  // on the open tag, the cross-references live in the inner XML.
  const styleRe = /<w:style\b([^>]*)\bw:styleId="([^"]+)"([^>]*)>([\s\S]*?)<\/w:style>/g;
  let m: RegExpExecArray | null;
  while ((m = styleRe.exec(stylesXml)) !== null) {
    const openAttrsBefore = m[1];
    const id = m[2];
    const openAttrsAfter = m[3];
    const inner = m[4];
    if (/\bw:default="(?:1|true)"/.test(openAttrsBefore + openAttrsAfter)) defaults.add(id);
    refs.set(id, {
      basedOn: styleRefVal(inner, "basedOn"),
      link: styleRefVal(inner, "link"),
      next: styleRefVal(inner, "next"),
      numStyleLink: styleRefVal(inner, "numStyleLink"),
      styleLink: styleRefVal(inner, "styleLink")
    });
  }
  return { refs, defaults };
}

/** Every `w:val` of the given style-reference tags found anywhere in a content substring. */
function collectStyleValRefs(xml: string, tags: readonly string[]): Set<string> {
  const out = new Set<string>();
  for (const tag of tags) {
    const re = new RegExp(`<w:${tag}\\b[^>]*\\bw:val="([^"]+)"`, "g");
    let m: RegExpExecArray | null;
    while ((m = re.exec(xml)) !== null) out.add(m[1]);
  }
  return out;
}

/**
 * The content-part style references that SEED the closure: paragraph/character/table style ids cited in
 * document.xml and any footnotes/headers/footers, PLUS the numbering backlinks (a `w:lvl`'s `w:pStyle`
 * and its `w:rPr`'s `w:rStyle` — the verified gap where ndca styles are referenced ONLY through
 * numbering). All three tags are scanned in every content part: numbering's own `w:pStyle`/`w:rStyle`
 * are exactly the backlinks we need, and scanning document.xml for them is harmless (it has them too).
 * Computed POST-splice so a cite-repair-injected net-new `w:rStyle` is included (loss-5).
 */
const CONTENT_STYLE_REF_TAGS = ["pStyle", "rStyle", "tblStyle"] as const;

/**
 * True when a non-document, non-styles part can carry style REFERENCES that must seed the closure:
 * `numbering.xml` (the `w:lvl` backlinks — the verified ndca gap), and the auxiliary story parts
 * `footnotes` / `endnotes` / `header*` / `footer*` (their `<w:p>`/`<w:r>` cite styles the same way
 * the body does). Recognized by part NAME (Word's conventional `/word/<name>.xml`) — the durable,
 * cheap signal. `settings`/`theme`/`fontTable`/rels parts carry no `w:pStyle`/`w:rStyle`, so excluding
 * them keeps the seed scan tight (and harmless if a stray ref ever appeared — it would only RETAIN
 * more, never drop a needed style).
 */
function isStyleReferencingPart(part: PackagePart): boolean {
  return /\/word\/(numbering|footnotes|endnotes|header\d*|footer\d*)\.xml$/.test(part.name);
}

/**
 * Expand a seed set to its transitive closure over the style graph (FIXPOINT, not one-hop). From each
 * reached style follow basedOn / link / next / numStyleLink / styleLink AND the REVERSE link edge (a
 * linked para/char pair is retained whichever half is reached). A worklist to fixpoint; cycle-safe
 * because a style is enqueued only the first time it enters the closure.
 */
function expandStyleClosure(seeds: Iterable<string>, refs: Map<string, StyleRefs>): Set<string> {
  // Reverse link adjacency: if A declares `<w:link w:val="B">`, reaching B must also retain A.
  const linkedFrom = new Map<string, string[]>();
  for (const [id, r] of refs) {
    if (r.link) {
      const arr = linkedFrom.get(r.link);
      if (arr) arr.push(id);
      else linkedFrom.set(r.link, [id]);
    }
  }
  const closure = new Set<string>();
  const work: string[] = [];
  const push = (id: string | null | undefined): void => {
    if (id && !closure.has(id)) {
      closure.add(id);
      work.push(id);
    }
  };
  for (const s of seeds) push(s);
  while (work.length) {
    const id = work.pop()!;
    const r = refs.get(id);
    if (r) {
      push(r.basedOn);
      push(r.link);
      push(r.next);
      push(r.numStyleLink);
      push(r.styleLink);
    }
    // Reverse link edge (both directions of a linked pair).
    const back = linkedFrom.get(id);
    if (back) for (const b of back) push(b);
  }
  return closure;
}

/**
 * Re-emit a `<w:styles>` part keeping ONLY the `<w:style>` children whose styleId is in `keep`, with
 * every NON-`<w:style>` byte (the `<w:styles …>` root open tag + its attrs/namespaces, `docDefaults`,
 * `latentStyles`, inter-element whitespace, the `</w:styles>` close) preserved VERBATIM. Works by
 * walking the part as a string: each `<w:style …>…</w:style>` span is dropped unless its id is kept;
 * everything between/around the style spans is copied byte-for-byte. A `<w:style>` with no styleId is
 * always kept (it can't be referenced, so dropping it could change rendering — be conservative).
 */
function pruneStylesXml(stylesXml: string, keep: Set<string>): string {
  const styleRe = /<w:style\b[^>]*?>[\s\S]*?<\/w:style>/g;
  let out = "";
  let last = 0;
  let m: RegExpExecArray | null;
  while ((m = styleRe.exec(stylesXml)) !== null) {
    // Copy the verbatim bytes BEFORE this style element (root open tag, docDefaults, whitespace, …).
    out += stylesXml.slice(last, m.index);
    const span = m[0];
    const idMatch = /\bw:styleId="([^"]+)"/.exec(span);
    const id = idMatch ? idMatch[1] : null;
    if (id === null || keep.has(id)) out += span; // keep retained (or unidentifiable) styles verbatim
    last = m.index + span.length;
  }
  out += stylesXml.slice(last); // the verbatim tail (latentStyles, </w:styles>, …)
  return out;
}

// ---------------------------------------------------------------------------
// Whole-body package: parse once, hand out per-paragraph fragments, splice edits
// back into the same tree, re-serialize.
// ---------------------------------------------------------------------------

/**
 * Construction-time options for `WholeBodyPackage`. The ONLY option today is the P7 (Loop 002 B3)
 * `pruneStyles` runtime flag — a code-level, DEFAULT-OFF opt-in (NOT a manifest/build flag, so the
 * manifest stays byte-identical, scope-4). Kept as an interface (not a positional bool) so future
 * package-level options are additive and self-documenting at call sites.
 */
export interface WholeBodyPackageOptions {
  /**
   * P7 referenced-closure styles prune (DEFAULT FALSE). When false (the shipped default), the styles
   * part is emitted VERBATIM and never DOM-parsed (002-F6 byte-identity). When true, the commit hook
   * may call `pruneStylesToClosure()` to shrink the styles part to the styles actually reachable from
   * the retained content — a WRITE-SIDE size reduction (perf-1). GO is a WET decision (see PLAN §5).
   */
  pruneStyles?: boolean;
}

/**
 * A parsed whole-body flat-OPC package that the adapter caches between its read
 * and write halves. Construction parses the package and snapshots the ordered
 * body `<w:p>` nodes; `paragraphXml(i)` serializes one standalone (with ambient
 * namespaces re-declared so the engine can parse it in isolation); `replace(i,…)`
 * splices an edited single-`<w:p>` fragment back into the SAME node; `serialize()`
 * yields the package to hand to `body.insertOoxml(…, "Replace")`.
 *
 * The index `i` is the position of the paragraph in document order within
 * `<w:body>`. The adapter is responsible for confirming this aligns with Word's
 * `paragraphs.items[i]` (the alignment guard) before trusting per-index metadata.
 */
export class WholeBodyPackage {
  /**
   * The DOM-parsed DOCUMENT part — its root is `<w:document>` when the input is a flat-OPC package
   * (A2 parses ONLY this part), or the whole-input root for a raw document.xml / bare `<w:p>`
   * fragment. Its `<w:p>` nodes (`this.paras`) are mutated in place by the node-direct Hide and
   * re-serialized by `serialize()` — the CORE-3 contract.
   */
  private readonly doc: any;
  private readonly scope: any;
  private readonly paras: any[];
  /** The `<w:document>` element to re-serialize on `serialize()` (the document part's subtree). */
  private readonly docElement: any;
  /** The real `<w:document>` start-tag attrs (xmlns:* + mc:Ignorable), re-applied per fragment. */
  private readonly docAttrs: string;
  /** styleId → outline definition, parsed once from the package's styles.xml (for `headingLevel`). */
  private readonly styleDefs: Map<string, StyleDef>;
  /** Whether the package actually carried a `<w:styles>` part (to distinguish "absent" from "unparseable"). */
  private readonly stylesPresent: boolean;
  /**
   * The flat-OPC segmentation (null for a raw document.xml / bare fragment, which serialize via the
   * whole-input DOM directly). When present, `serialize()` stitches the verbatim part substrings in
   * original order with the document part swapped for the re-serialized `<w:document>` subtree.
   */
  private readonly segments: SegmentedPackage | null;
  /** Index of the document part within `segments.parts`, and the bytes around its `<w:document>`. */
  private readonly docPartIndex: number;
  private readonly docHead: string; // the document part's bytes before `<w:document`
  private readonly docTail: string; // the document part's bytes from `</w:document>` onward
  /**
   * The small auxiliary parts (rels + styles/numbering/theme) the COMMIT/READ fragments need, parsed
   * LAZILY on first access — so the node-direct read ctor never DOM-parses the 1.27MB styles part (or
   * numbering/theme). `paragraphXml`/`commitXml` are off the node-direct hot path, so deferring this
   * is free there and books the read-side ctor-parse reduction (perf-1). Memoized via the fields below.
   */
  private auxParsed: { relsById: Map<string, string>; auxParts: string; auxRels: string } | null = null;
  /** Captured aux/rels part substrings (verbatim) used to build `auxParsed` on demand. */
  private readonly auxPartSubstrings: string[];
  /**
   * P7 (Loop 002 B3) — the DEFAULT-OFF runtime flag enabling `pruneStylesToClosure()`. A code-level
   * option (NOT a manifest/build flag — scope-4 keeps the manifest byte-identical), default false so
   * the SHIPPED path emits the styles part VERBATIM and never DOM-parses it (002-F6). When true, the
   * commit hook may shrink the styles part to its referenced closure between splice and `serialize()`.
   */
  private readonly pruneStyles: boolean;
  /**
   * P7 — the index of the `<w:styles>` part within `segments.parts` (the SAME part `extractStylesSubstring`
   * found, keyed to a `<w:styles>` ROOT, NOT a content-type). `pruneStylesToClosure()` REPLACES that
   * part's verbatim `full` substring with the pruned bytes so `serialize()`'s in-order stitch emits the
   * shrunk part. `-1` when there is no styles part (or the input isn't segmentable) → prune is a no-op.
   */
  private readonly stylesPartIndex: number;
  /**
   * P7 — guards prune IDEMPOTENCE + correct re-seed (002-S6). Set once the styles part has been pruned
   * so a second `pruneStylesToClosure()` (e.g. a reused `lastRead.pkg`) is a no-op rather than re-parsing
   * an already-pruned part. Re-seeding correctness is preserved because the FIRST prune computes the
   * closure over the (post-splice) document part, which the prune never mutates.
   */
  private stylesPruned = false;

  constructor(packageXml: string, options: WholeBodyPackageOptions = {}) {
    this.pruneStyles = options.pruneStyles ?? false;
    // A2: cut the package at `<pkg:part>` boundaries by string scan; DOM-parse ONLY document.xml.
    // A raw document.xml / bare `<w:p>` fragment isn't a `<pkg:package>` → `segments` is null and we
    // parse the whole input exactly as before (byte-identical degrade).
    this.segments = segmentPackage(packageXml);
    const docPartIndex = this.segments
      ? this.segments.parts.findIndex((p) => p.name === "/word/document.xml")
      : -1;
    // A `<pkg:package>` is only safely segmentable when its `/word/document.xml` part exists AND
    // carries non-empty `<pkg:xmlData>` content (the `<w:document>` we must parse + re-serialize).
    // A self-closing/empty `<pkg:xmlData>` would make `indexOf("")` return 0 and corrupt the stitch,
    // so we fall back to the whole-input parse there — byte-identical to today on that odd shape.
    const docXmlData = docPartIndex >= 0 ? this.segments!.parts[docPartIndex].xmlData : "";
    const segmentable = this.segments !== null && docPartIndex >= 0 && docXmlData.length > 0;
    this.docPartIndex = segmentable ? docPartIndex : -1;

    if (segmentable) {
      const docPart = this.segments!.parts[docPartIndex];
      // Parse ONLY the document part's xmlData (its single root `<w:document>` — or a bare root for an
      // unusual fixture). Capture the verbatim bytes around that root within `docPart.full` (the
      // `<pkg:part …><pkg:xmlData>` head and `</pkg:xmlData></pkg:part>` tail) so `serialize()` can
      // splice the re-serialized (mutation-bearing) root back without touching the wrapper. The
      // boundary is the xmlData inner: `docPart.xmlData` is exactly the root element's source (non-empty,
      // guaranteed by `segmentable` above), and `<pkg:xmlData>` wraps exactly one element, so
      // reserializing the parsed root reproduces it.
      const xdStart = docPart.full.indexOf(docXmlData);
      this.docHead = docPart.full.slice(0, xdStart);
      this.docTail = docPart.full.slice(xdStart + docXmlData.length);
      this.doc = parse(docXmlData);
    } else {
      // Not segmentable — today's behavior exactly: DOM-parse the whole input.
      this.docHead = "";
      this.docTail = "";
      this.doc = parse(packageXml);
    }

    this.scope = bodyScope(this.doc);
    // STORY paragraphs only (exclude textbox-nested) so indices align 1:1 with Office.js
    // `body.paragraphs` — the ① half of the Stage 4.1 whole-body alignment fix.
    this.paras = storyParagraphsIn(this.scope);
    // The `<w:document>` element to re-serialize: the parsed document part's root when it IS a
    // `<w:document>`, else the whole-input document element (raw document.xml / synthetic fixtures).
    const docs = this.doc.getElementsByTagName("w:document");
    this.docElement = docs && docs.length > 0 ? docs.item(0) : this.doc.documentElement;
    // Captured ONCE so every fragment we hand out is the host's own namespace/MC context.
    this.docAttrs = documentAttrs(this.doc);

    // STYLES: feed the styles part's xmlData SUBSTRING straight to the regex-only resolver — NEVER
    // DOM-parse it (the read-side win: the styles part is the largest single part). When not
    // segmentable, recover the styles substring from the parsed document tree's sibling parts is
    // impossible (only document.xml was parsed) — but a raw document.xml carries no styles part at
    // all, so `stylesXml` is "" there, exactly as `extractStylesXml` returned before.
    const stylesXml = this.extractStylesSubstring();
    this.stylesPresent = stylesXml.length > 0;
    this.styleDefs = parseStyleDefs(stylesXml);
    // P7 (Loop 002 B3): remember WHICH part is the styles part so `pruneStylesToClosure()` can REPLACE
    // its verbatim `full` substring in place (the `serialize()` stitch then emits the pruned bytes).
    // Keyed to the SAME `<w:styles>` ROOT signal `extractStylesSubstring` uses, so the two never diverge.
    this.stylesPartIndex = this.segments
      ? this.segments.parts.findIndex((p) => /<w:styles[\s>]/.test(p.xmlData))
      : -1;

    // Capture the small aux/rels part substrings (verbatim) for the LAZY `paragraphXml`/`commitXml`
    // path — NOT parsed here, so the node-direct read ctor stays document-only.
    this.auxPartSubstrings = segmentable
      ? this.segments!.parts.filter((_, k) => k !== docPartIndex).map((p) => p.full)
      : [];
  }

  /**
   * The styles part's xmlData SUBSTRING (the `<w:styles>…</w:styles>` XML), or "" when none.
   * Regex-only: A2 never DOM-parses styles.xml (perf-1, the largest part). Keyed to a `<w:styles>`
   * ROOT — exactly the old `extractStylesXml` semantics (`getElementsByTagName("w:styles")`), so a
   * styles part is recognized by its ROOT element, NOT by a particular `pkg:contentType` (a producer
   * may omit/differ on the attribute). For a non-segmentable input (raw document.xml) there is no
   * styles part, so "" — matching the old code which returned "" for a doc with no `<w:styles>`.
   * `styleParseSuspect` then keys to that root being present-but-zero-defs.
   */
  private extractStylesSubstring(): string {
    if (!this.segments) return "";
    // A `<w:styles` open tag in a part's xmlData marks the styles part (the part is named
    // `/word/styles.xml` and/or carries the styles content-type, but the ROOT is the durable signal).
    const part = this.segments.parts.find((p) => /<w:styles[\s>]/.test(p.xmlData));
    return part ? part.xmlData : "";
  }

  /**
   * Build (once, memoized) the aux artifacts the per-paragraph fragments need: the rel-id map, the
   * serialized styles/numbering/theme `<pkg:part>`s, and their `<Relationship>`s. Parses ONLY the
   * small rels + aux part substrings through xmldom and reuses the existing `collect*` helpers, so
   * the output is BYTE-IDENTICAL to the pre-segmentation path (same nodes, same xmldom serialization
   * — including the `xmlns`/`xmlns:pkg` decls xmldom re-emits when serializing a part standalone).
   * Deferred off the node-direct hot path (which never calls paragraphXml/commitXml).
   */
  private aux(): { relsById: Map<string, string>; auxParts: string; auxRels: string } {
    if (this.auxParsed) return this.auxParsed;
    // The DOM the `collect*` helpers scan. On the segmented path it is a TINY synthetic package of
    // just the rels + aux part substrings (so the styles/document parts are never re-parsed here); on
    // the non-segmentable path it is `this.doc` itself — the WHOLE-INPUT parse the old ctor scanned,
    // so a raw document.xml / bare fragment / pkg-without-document-part collects EXACTLY what
    // `collectRelationships(this.doc)`/`collectAuxParts`/`collectAuxRels` returned before (preserving
    // the byte-identical-degrade contract even if such an input carried inline `<Relationship>`s).
    const auxDoc =
      this.docPartIndex < 0
        ? this.doc
        : parse(`<pkg:package xmlns:pkg="${PKG_NS}">${this.auxPartSubstrings.join("")}</pkg:package>`);
    this.auxParsed = {
      relsById: collectRelationships(auxDoc),
      auxParts: collectAuxParts(auxDoc),
      auxRels: collectAuxRels(auxDoc)
    };
    return this.auxParsed;
  }

  /** relationship id → `<Relationship/>` XML (lazy). */
  private get relsById(): Map<string, string> {
    return this.aux().relsById;
  }
  /** Serialized styles/numbering/theme `<pkg:part>`s, bundled into the COMMIT fragment only (lazy). */
  private get auxParts(): string {
    return this.aux().auxParts;
  }
  /** Serialized styles/numbering/theme `<Relationship>`s for the commit fragment's doc rels (lazy). */
  private get auxRels(): string {
    return this.aux().auxRels;
  }

  /**
   * True when the package carries a `<w:styles>` part but parsing it yielded ZERO style definitions
   * — a serialization shape the resolver can't read (e.g. a default-namespace `<styles>` instead of
   * the `w:`-prefixed form `body.getOoxml()` emits). The pure whole-body path treats this as
   * untrustworthy and falls back to the proven proxy read rather than risk mass-hiding every
   * style-derived heading (adversarial-review C-3). A package with NO styles part is NOT suspect:
   * inline `<w:outlineLvl>` and the heading-name fallback still classify it.
   */
  get styleParseSuspect(): boolean {
    return this.stylesPresent && this.styleDefs.size === 0;
  }

  /** Number of STORY body paragraphs (document order, includes table paragraphs, excludes textboxes). */
  get count(): number {
    return this.paras.length;
  }

  /**
   * The concatenated visible text of the i-th story paragraph (its `<w:t>` runs, minus
   * any nested-textbox text), used by the adapter's tolerant alignment (⑧) to match
   * package paragraphs to Word's `body.paragraphs` proxies when the counts differ by
   * Word's serialization quirk. Cheap, read-only; never mutates the tree.
   */
  paragraphText(i: number): string {
    const node = this.paras[i];
    if (!node) throw new RangeError(`paragraph index ${i} out of range (count ${this.count})`);
    return collectParagraphText(node);
  }

  /**
   * The 0-based heading level (0 = Heading 1 … 8 = Heading 9; null = body) of the i-th story
   * paragraph, resolved PURELY from the package: the paragraph's inline `<w:outlineLvl>`, else its
   * `<w:pStyle>` effective level via the `basedOn` cascade in the bundled styles.xml. This is what
   * lets the pure whole-body classify path (avenue ⑦) attach a keep level WITHOUT a Word proxy
   * (the host reports `canGetStyles:false`). `outlineNumberFromProps` returns the 1-based
   * `Paragraph.outlineLevel` equivalent (1–9 = headings, 10 = body), so the 0-based mapping here is
   * exactly `normalizeOutlineNumber(…, "oneBased")` — keeping the pure path and the proxy path on
   * one outline convention.
   *
   * READS THE NODE DIRECTLY — NEVER SERIALIZES. Both outline signals live, per the OOXML schema,
   * only in the paragraph's own direct-child `<w:pPr>`, so a two-level child walk is all the
   * extraction needed. The old form serialized the ENTIRE `<w:p>` subtree (runs, hyperlinks,
   * nested textboxes) just so `outlineNumberOf`'s two regexes could re-find those elements —
   * the dominant per-paragraph cost of the pure whole-body read, paid AGAIN by `repairCites`.
   * The walk is also semantically TIGHTER than the regex, which could falsely match an
   * `<w:outlineLvl>`/`<w:pStyle>` from a nested `w:txbxContent` paragraph or a `<w:pPrChange>`
   * revision when the outer paragraph declares none; direct-child scoping matches the live
   * `Paragraph.outlineLevel` semantics (0 corpus mismatches across all real sample docs).
   * __tests__/ooxmlPackage.test.ts guards the zero-serialize count and both delta directions.
   */
  headingLevel(i: number): number | null {
    const node = this.paras[i];
    if (!node) throw new RangeError(`paragraph index ${i} out of range (count ${this.count})`);
    // The paragraph's OWN properties: the first direct element child named <w:pPr>. The schema
    // mandates pPr-first when present, but scan all children to tolerate lenient producers.
    let pPr: any = null;
    for (let k = node.firstChild; k; k = k.nextSibling) {
      if (k.nodeType === ELEMENT_NODE && k.nodeName === "w:pPr") {
        pPr = k;
        break;
      }
    }
    let inlineLvl: number | null = null;
    let styleId: string | null = null;
    if (pPr) {
      // Direct children only — a <w:pPrChange> revision's nested old pPr must NOT leak through.
      // First occurrence wins (the schema allows at most one of each anyway).
      for (let k = pPr.firstChild; k; k = k.nextSibling) {
        if (k.nodeType !== ELEMENT_NODE) continue;
        if (k.nodeName === "w:outlineLvl" && inlineLvl === null) {
          const v = k.getAttribute("w:val");
          // All-digits only — mirrors the string form's `w:val="(\d+)"`, so a malformed value
          // falls through to the style cascade instead of poisoning the level with NaN.
          if (v && /^\d+$/.test(v)) inlineLvl = Number(v);
        } else if (k.nodeName === "w:pStyle" && styleId === null) {
          const v = k.getAttribute("w:val");
          if (v) styleId = v; // empty/missing val = no style, same as the regex's [^"]+
        }
      }
    }
    const n = outlineNumberFromProps(inlineLvl, styleId, this.styleDefs);
    return n >= 1 && n <= 9 ? n - 1 : null;
  }

  /**
   * True when the i-th story paragraph sits inside a table (a `<w:tc>`/`<w:tbl>` ancestor). Table
   * paragraphs ARE story paragraphs (only textbox-nested ones are excluded by `storyParagraphsIn`),
   * and classification suppresses hiding inside tables, so the pure whole-body path resolves table
   * membership from the package structure itself — no proxy `parentTableOrNullObject` probe.
   */
  inTable(i: number): boolean {
    let n: any = this.paras[i] ? this.paras[i].parentNode : null;
    while (n) {
      if (n.nodeName === "w:tc" || n.nodeName === "w:tbl") return true;
      n = n.parentNode;
    }
    return false;
  }

  /**
   * Serialize the i-th body paragraph as a standalone fragment, wrapped in a MINIMAL flat-OPC
   * package whose `<w:document>` reuses the host's REAL start-tag attributes (`this.docAttrs`:
   * every `xmlns:*` decl + `mc:Ignorable`).
   *
   * WHY THE PACKAGE WRAPPER (Stage 4.2): a bare cloned `<w:p>` handed to the LIVE HOST's
   * `insertOoxml` was rejected ("we found a problem with its contents"). Wrapping in
   * `<w:document {host's own attrs}><w:body>…</w:body>` is the exact namespace/MC envelope Word
   * itself emits, so it's accepted — but kept MINIMAL (no styles/numbering parts) because the
   * host's per-paragraph package bundles every document part, which is what made the old
   * fallback's classify ≈100× slower. Using the host's real attrs (vs a synthesized union)
   * preserves `mc:Ignorable` for the emoji `mc:AlternateContent` and avoids leaking the
   * package-relationships default `xmlns` onto `<w:document>`.
   *
   * THE START-PART RELATIONSHIP IS MANDATORY (Stage 4.2 commit fix). Every flat-OPC package
   * MUST carry `/_rels/.rels` declaring the `officeDocument` relationship to `/word/document.xml`
   * (`PACKAGE_RELS_PART`); without it the host can't identify the main document and `insertOoxml`
   * throws `GeneralException`. The original minimal wrapper omitted it, so the per-paragraph
   * commit failed on EVERY paragraph (plain ones too) — it only ever "worked" when the read had
   * fully fallen back to Word's own complete packages. (See the constant's doc.)
   *
   * The engine still sees exactly one `<w:body>` paragraph, so `readRuns`/`assertSingleParagraph`
   * work identically to the Stage-1 bare-`<w:p>` fixtures.
   *
   * If the paragraph references relationships (a hyperlink `r:id`, etc.), a `/word/_rels/document.xml.rels`
   * part carrying ONLY those relationships is ALSO included so the reference isn't dangling (Word
   * rejects a dangling `r:id` — audit C1). That document-level rels part is scoped to `document.xml`
   * and is independent of the package-level `/_rels/.rels` start part. Paragraphs with no `r:*`
   * reference get the start part + the one document part (still minimal).
   */
  paragraphXml(i: number): string {
    const node = this.paras[i];
    if (!node) throw new RangeError(`paragraph index ${i} out of range (count ${this.count})`);
    // Serialize the ATTACHED node — no deep clone. xmldom's serializer is subtree-self-contained
    // for prefixed nodes: it only consults ancestors (lookupPrefix) when the start node's prefix
    // is null, and every node here is a prefixed `w:p`, so it never walks above the subtree and
    // never mutates it (same reason `headingLevel` serializes attached). The old
    // `cloneNode(true)` was byte-identical defensive copying at ~10× the cost of serialization
    // itself — the single largest per-paragraph cost on the pure whole-body read (avenue ⑦).
    // __tests__/ooxmlPackage.test.ts guards both the zero-clone count and the non-mutation.
    const inner = serialize(node);
    // Carry only the relationships THIS paragraph references (keeps the package minimal — no
    // styles/numbering parts — so classify stays fast, while hyperlinks resolve).
    const usedRels: string[] = [];
    for (const id of referencedRelIds(node)) {
      const rel = this.relsById.get(id);
      if (rel) usedRels.push(rel);
    }
    const docRelsPart = usedRels.length
      ? `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="${RELS_CONTENT_TYPE}">` +
        `<pkg:xmlData><Relationships xmlns="${RELS_NS}">${usedRels.join("")}</Relationships></pkg:xmlData>` +
        `</pkg:part>`
      : "";
    return (
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      PACKAGE_RELS_PART + // REQUIRED: declares /word/document.xml as the main document part
      docRelsPart +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${WML_DOCUMENT_CT}">` +
      `<pkg:xmlData><w:document ${this.docAttrs}><w:body>${inner}</w:body></w:document></pkg:xmlData>` +
      `</pkg:part></pkg:package>`
    );
  }

  /**
   * The i-th story paragraph as a node-backed `ParsedParagraph` (Loop 002 B1 — the ANCHOR of the
   * node-direct hide path). Constructs `ParsedParagraph.fromNode` over the SAME live `<w:p>` node this
   * package already snapshotted at construction (`this.paras[i]`) and the SAME owner document — ZERO
   * string parse and ZERO serialize (contrast `paragraphXml(i)`, which clones+serializes+rewraps the node
   * into a standalone flat-OPC package). The engine reads its `RunView[]` and mutates `<w:vanish>` IN
   * PLACE via `applyVisibilityInPlace`; this package then `serialize()`s the whole body once, with the
   * mutated node already in the tree. Because the node IS the package's node, the pure commit needs NO
   * `replace()` — the mutation is already spliced in by reference. The `cleanAlign` identity mapping the
   * pure read produces guarantees engine index === this index, so the caller indexes directly.
   *
   * `fromNode` requires a `<w:p>` owned by `this.doc`; `this.paras[i]` is exactly that (it was selected
   * from `this.doc` at construction and is mutated in place, never re-imported), so the contract holds.
   */
  parsedParagraph(i: number): ParsedParagraph {
    const node = this.paras[i];
    if (!node) throw new RangeError(`paragraph index ${i} out of range (count ${this.count})`);
    return ParsedParagraph.fromNode(this.doc, node);
  }

  /**
   * Wrap the engine's EDITED single-paragraph fragment for the per-paragraph
   * `Paragraph.insertOoxml(…, "Replace")` commit, INCLUDING the document's style / numbering /
   * theme parts so that any formatting the paragraph INHERITS from a referenced style — underline,
   * character border/box, font size, anything that is NOT an inline run property — survives the
   * round-trip onto the host.
   *
   * WHY THIS IS SEPARATE FROM `paragraphXml`. `paragraphXml` is the READ fragment the engine
   * re-parses for EVERY paragraph during classify; it is deliberately MINIMAL (no styles part)
   * because bundling the full `styles.xml` into all N read fragments is what made the old
   * per-paragraph read ≈100× slower. The COMMIT fragment is the engine's OUTPUT — parsed once
   * here, once by the host, and ONLY for the partial-hide paragraphs that take the OOXML path —
   * so it can afford to carry the style parts that make the host render the paragraph faithfully.
   *
   * THE BUG THIS FIXES. Paragraphs aligned to the whole-body package were read from this minimal
   * (style-less) package, so committing them through it made the host resolve every style-derived
   * property to a document DEFAULT: underlined/boxed/18pt text collapsed to plain 11pt while inline
   * `<w:highlight>` survived. (Targeted-`getOoxml` paragraphs never hit it because Word's own
   * per-paragraph package already bundles styles.) Bundling the captured aux parts here closes that.
   *
   * Hyperlink relationships the paragraph references are carried too (same as `paragraphXml`), so a
   * fragment with both a cite hyperlink AND a style resolves both. All rel Ids come from the one
   * source `document.xml.rels`, so the aux and hyperlink Ids never collide.
   */
  commitXml(paragraphOoxml: string): string {
    const node = paragraphsIn(bodyScope(parse(paragraphOoxml)))[0];
    if (!node) throw new Error("commitXml: fragment held no <w:p> to wrap for commit");
    const inner = serialize(node.cloneNode(true));
    // document.xml.rels = the aux-part rels (styles/numbering/theme) PLUS any relationship THIS
    // paragraph references (a hyperlink r:id). Both come from the source document.xml.rels.
    const usedRels: string[] = [];
    for (const id of referencedRelIds(node)) {
      const rel = this.relsById.get(id);
      if (rel) usedRels.push(rel);
    }
    const relsXml = this.auxRels + usedRels.join("");
    const docRelsPart = relsXml
      ? `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="${RELS_CONTENT_TYPE}">` +
        `<pkg:xmlData><Relationships xmlns="${RELS_NS}">${relsXml}</Relationships></pkg:xmlData>` +
        `</pkg:part>`
      : "";
    return (
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      PACKAGE_RELS_PART + // REQUIRED start part (see paragraphXml)
      docRelsPart +
      this.auxParts + // styles/numbering/theme parts — the fidelity fix
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${WML_DOCUMENT_CT}">` +
      `<pkg:xmlData><w:document ${this.docAttrs}><w:body>${inner}</w:body></w:document></pkg:xmlData>` +
      `</pkg:part></pkg:package>`
    );
  }

  /**
   * Replace the i-th body paragraph with the single `<w:p>` from `paragraphOoxml`
   * (the engine's edited fragment). Guards single-paragraph-ness, imports the new
   * node into this package's document, and swaps it in place — leaving every other
   * node, and the package structure, byte-identical to what Word handed us.
   */
  replace(i: number, paragraphOoxml: string): void {
    const old = this.paras[i];
    if (!old) throw new RangeError(`paragraph index ${i} out of range (count ${this.count})`);
    // ONE parse serves both the single-<w:p> guard and the node selection. The old
    // assertSingleParagraph(string) call followed by parse() re-parsed the IDENTICAL
    // fragment twice per changed paragraph — the dominant avoidable cost of the
    // whole-body commit loop (one replace per changed paragraph on the default hide
    // and on cite repair). parseCount.test.ts pins the single parse; the RangeError
    // above stays FIRST so an out-of-range index never pays a parse at all.
    const fragDoc = parse(paragraphOoxml);
    const newPara = singleBodyParagraph(fragDoc, `whole-body splice @${i}`);
    const imported = this.doc.importNode(newPara, true);
    const parent = old.parentNode;
    if (!parent) throw new Error(`paragraph index ${i} has no parent to replace within`);
    parent.replaceChild(imported, old);
    this.paras[i] = imported;
  }

  /**
   * Serialize the whole package for `body.insertOoxml(pkg, "Replace")`.
   *
   * A2 (segmented input): stitch the captured part substrings in ORIGINAL ORDER, with the document
   * part swapped for `docHead + serialize(<w:document> subtree) + docTail`. The `<w:document>`
   * subtree carries the engine's in-place `<w:vanish>`/bridge mutations to `this.paras[i]` (the
   * CORE-3 contract); every OTHER part is re-emitted VERBATIM from the input. This is byte-identical
   * to the pre-segmentation whole-DOM `serialize()` on the canonical fixtures (each verbatim part
   * substring already equals xmldom's standalone serialization of that part, and the document subtree
   * re-serializes identically), and on real docs it preserves Word's own non-document bytes exactly
   * (a fidelity improvement) while the document part DOM-reserializes as before.
   *
   * Non-segmentable input (raw document.xml / bare fragment): exactly today — serialize the one
   * parsed tree.
   */
  serialize(): string {
    if (!this.segments || this.docPartIndex < 0) return serialize(this.doc);
    const { head, parts, between, tail } = this.segments;
    let out = head;
    for (let k = 0; k < parts.length; k++) {
      out += between[k] ?? ""; // verbatim inter-part bytes (empty for canonical packages)
      if (k === this.docPartIndex) {
        // The document part: verbatim wrapper bytes around the re-serialized (mutation-bearing)
        // `<w:document>` subtree. docHead/docTail are the `<pkg:part …><pkg:xmlData>` and
        // `</pkg:xmlData></pkg:part>` bytes captured at construction; serialize(docElement) carries
        // the engine's in-place vanish/bridge edits.
        out += this.docHead + serialize(this.docElement) + this.docTail;
      } else {
        out += parts[k].full; // every non-document part, byte-for-byte from the input
      }
    }
    return out + tail;
  }

  /**
   * Whether this package will actually prune (the flag is ON, the input is a segmentable flat-OPC
   * package, AND it carries a `<w:styles>` part). Exposed so the commit hook / tests can assert that a
   * flag-OFF or styles-less package is left untouched without reaching into private state.
   */
  get willPruneStyles(): boolean {
    return this.pruneStyles && this.stylesPartIndex >= 0;
  }

  /**
   * The styles part's CURRENT verbatim `<w:styles>…</w:styles>` xmlData substring (post-prune if
   * `pruneStylesToClosure()` ran), or "" when there is no styles part. Read-only view for the
   * closure-correctness + byte-reduction tests; never used on the hot path.
   */
  stylesXmlForTest(): string {
    if (this.stylesPartIndex < 0 || !this.segments) return "";
    return this.segments.parts[this.stylesPartIndex].xmlData;
  }

  /**
   * P7 — shrink the styles part to the transitive closure of the styles REACHABLE from the retained
   * content, in place (DEFAULT-OFF flag; CASES 002-S6). A NO-OP when the flag is off, the input is not
   * a segmentable flat-OPC package, there is no styles part, or a prior call already pruned this package
   * (idempotence — re-seeding off the SAME post-splice document yields the same closure).
   *
   * THE HOOK CONTRACT. Call this AFTER the node-direct splice (the in-place `<w:vanish>`/bridge
   * mutations to `this.paras[i]`) and BEFORE `serialize()`. Seeds are computed POST-splice so a
   * cite-repair-injected net-new `w:rStyle` is captured (loss-5); the prune mutates ONLY the styles
   * part's captured `full`/`xmlData` substring, which `serialize()` then emits in place — the document
   * body, every other part, and the package framing are untouched (002-F1 on the body holds).
   *
   * Returns the byte delta `(before − after)` of the styles part (≥ 0) — write-side evidence (perf-1),
   * 0 when nothing was pruned.
   */
  pruneStylesToClosure(): number {
    // Flag OFF / nothing to prune / already pruned → leave the verbatim substring untouched (002-F6).
    if (!this.pruneStyles || this.stylesPartIndex < 0 || this.stylesPruned || !this.segments) return 0;
    this.stylesPruned = true; // mark first so a re-entrant/repeat call is a guaranteed no-op (idempotence)

    const stylesPart = this.segments.parts[this.stylesPartIndex];
    const stylesXml = stylesPart.xmlData;
    if (!stylesXml) return 0;

    // DOM-free parse of the full style table — same regex the resolver trusts (this is the ONLY place
    // the styles part is read structurally, and ONLY when the flag is on; flag-OFF never gets here).
    const { refs, defaults } = parseStyleGraph(stylesXml);

    // SEEDS: every content part's style references (post-splice) + all `w:default` styles.
    const seeds = new Set<string>(defaults);
    // document.xml — the POST-SPLICE body (serialize the mutated subtree so injected refs count).
    for (const id of collectStyleValRefs(serialize(this.docElement), CONTENT_STYLE_REF_TAGS)) seeds.add(id);
    // numbering / footnotes / endnotes / headers / footers — verbatim part substrings (untouched by the
    // splice). Each is scanned for the same three style-reference tags: numbering's `w:lvl` backlinks
    // (`w:pStyle` + the `w:rStyle` inside its `w:rPr`) are exactly the gap-closing refs (PLAN §2 B3).
    for (const part of this.segments.parts) {
      if (part === stylesPart || part.name === "/word/document.xml") continue;
      if (!isStyleReferencingPart(part)) continue;
      for (const id of collectStyleValRefs(part.xmlData, CONTENT_STYLE_REF_TAGS)) seeds.add(id);
    }

    // EXPAND to the transitive fixpoint, then PRUNE every `<w:style>` not in the closure.
    const closure = expandStyleClosure(seeds, refs);
    const pruned = pruneStylesXml(stylesXml, closure);
    if (pruned === stylesXml) return 0; // nothing removable — keep the byte-identical substring

    // REPLACE the captured substrings so `serialize()`'s in-order stitch emits the shrunk part. We swap
    // the `xmlData` inside the part's `full` wrapper, leaving the `<pkg:part …><pkg:xmlData>` head and
    // `</pkg:xmlData></pkg:part>` tail byte-identical (only the styles XML between them changes).
    const newFull = stylesPart.full.replace(stylesXml, pruned);
    this.segments.parts[this.stylesPartIndex] = { ...stylesPart, full: newFull, xmlData: pruned };
    return stylesXml.length - pruned.length;
  }
}
