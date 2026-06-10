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
import { parseStyleDefs, outlineNumberOf, type StyleDef } from "./outline";

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
 * Assert a writeback fragment contains EXACTLY ONE body `<w:p>` (audit #5). The
 * engine only ever edits a single paragraph's runs; a fragment carrying zero or
 * many `<w:p>` means something upstream corrupted the OOXML, and writing it back
 * could splice a table or empty the paragraph. Throws a clear, contextual error.
 */
export function assertSingleParagraph(ooxml: string, context = "writeback"): void {
  const count = paragraphsIn(bodyScope(parse(ooxml))).length;
  if (count !== 1) {
    throw new Error(
      `Rostrum ${context} guard: expected exactly one <w:p> in the fragment, found ${count}. ` +
        `Refusing to write back potentially corrupted OOXML.`
    );
  }
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

/**
 * The serialized `<w:styles>` part from a whole-body package (`word/styles.xml`), or "" when the
 * package carries none. `body.getOoxml()` always bundles it; the PURE whole-body classify path
 * (avenue ⑦) needs it to resolve outline levels through the `basedOn` cascade WITHOUT a Word proxy
 * (the host reports `canGetStyles:false`). `<w:styles>` is unique in the package, so a single
 * `getElementsByTagName` lookup is unambiguous; `parseStyleDefs` then regex-scans the result.
 */
function extractStylesXml(doc: any): string {
  const styles = doc.getElementsByTagName("w:styles");
  return styles && styles.length > 0 ? serialize(styles.item(0)) : "";
}

// ---------------------------------------------------------------------------
// Whole-body package: parse once, hand out per-paragraph fragments, splice edits
// back into the same tree, re-serialize.
// ---------------------------------------------------------------------------

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
  private readonly doc: any;
  private readonly scope: any;
  private readonly paras: any[];
  /** The real `<w:document>` start-tag attrs (xmlns:* + mc:Ignorable), re-applied per fragment. */
  private readonly docAttrs: string;
  /** relationship id → `<Relationship/>` XML, so a fragment can carry the rels it references. */
  private readonly relsById: Map<string, string>;
  /** Serialized styles/numbering/theme `<pkg:part>`s, bundled into the COMMIT fragment only. */
  private readonly auxParts: string;
  /** Serialized styles/numbering/theme `<Relationship>`s for the commit fragment's doc rels. */
  private readonly auxRels: string;
  /** styleId → outline definition, parsed once from the package's styles.xml (for `headingLevel`). */
  private readonly styleDefs: Map<string, StyleDef>;
  /** Whether the package actually carried a `<w:styles>` part (to distinguish "absent" from "unparseable"). */
  private readonly stylesPresent: boolean;

  constructor(packageXml: string) {
    this.doc = parse(packageXml);
    this.scope = bodyScope(this.doc);
    // STORY paragraphs only (exclude textbox-nested) so indices align 1:1 with Office.js
    // `body.paragraphs` — the ① half of the Stage 4.1 whole-body alignment fix.
    this.paras = storyParagraphsIn(this.scope);
    // Captured ONCE so every fragment we hand out is the host's own namespace/MC context and
    // carries any relationship it references — accepted by the live host's insertOoxml (4.2).
    this.docAttrs = documentAttrs(this.doc);
    this.relsById = collectRelationships(this.doc);
    // The style/numbering/theme parts + their rels, captured ONCE (identical for every paragraph)
    // and bundled into `commitXml` so style-inherited formatting survives the per-paragraph commit.
    this.auxParts = collectAuxParts(this.doc);
    this.auxRels = collectAuxRels(this.doc);
    // Parsed ONCE for the pure whole-body classify path (avenue ⑦): resolve a paragraph's outline
    // level from the package's OWN styles.xml instead of a Word proxy (`canGetStyles:false`).
    const stylesXml = extractStylesXml(this.doc);
    this.stylesPresent = stylesXml.length > 0;
    this.styleDefs = parseStyleDefs(stylesXml);
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
   * (the host reports `canGetStyles:false`). `outlineNumberOf` returns the 1-based
   * `Paragraph.outlineLevel` equivalent (1–9 = headings, 10 = body), so the 0-based mapping here is
   * exactly `normalizeOutlineNumber(…, "oneBased")` — keeping the pure path and the proxy path on
   * one outline convention.
   */
  headingLevel(i: number): number | null {
    const node = this.paras[i];
    if (!node) throw new RangeError(`paragraph index ${i} out of range (count ${this.count})`);
    const n = outlineNumberOf(serialize(node), this.styleDefs);
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
    assertSingleParagraph(paragraphOoxml, `whole-body splice @${i}`);
    const fragDoc = parse(paragraphOoxml);
    const newPara = paragraphsIn(bodyScope(fragDoc))[0];
    const imported = this.doc.importNode(newPara, true);
    const parent = old.parentNode;
    if (!parent) throw new Error(`paragraph index ${i} has no parent to replace within`);
    parent.replaceChild(imported, old);
    this.paras[i] = imported;
  }

  /** Serialize the whole package for `body.insertOoxml(pkg, "Replace")`. */
  serialize(): string {
    return serialize(this.doc);
  }
}
