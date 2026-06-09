// Helpers for the REAL-document test suite (realDocs.test.ts).
//
// These read the actual `.docx` files dropped into `rostrum/samples/`, pull out
// `word/document.xml` (the real OOXML Word produces) plus `word/styles.xml`, and turn
// the body paragraphs into engine-ready `FakePara` fixtures. This lets the FULL
// engine+adapter stack run against genuine debate-doc OOXML — every messy real-world
// run, hyperlink, field, footnote, content control, and numbering — with NO Word host.
// It complements the synthetic fixtures: those prove specific keep rules; these prove
// the engine never corrupts or chokes on reality, and produce real timing at scale.
//
// OUTLINE-LEVEL FIDELITY: a paragraph's outline level is resolved EXACTLY as the live
// `Paragraph.outlineLevel` would report it — a direct `<w:pPr><w:outlineLvl>` wins,
// else the paragraph style's effective level via the `basedOn` cascade in styles.xml
// (so `Analytics`/`Analytic` → `Heading4` → level 3 are kept, matching the host), else
// body. No hardcoded style-name guessing.
//
// SIZE TIERS (from a leading `[tag]` in the filename, else by file size):
//   small  — proof of concept (fast correctness)
//   medium — thoroughness (varied real content)
//   large / xlarge — timing + optimization (deferred unless ROSTRUM_PERF=1, because a
//                    26M-char document.xml is slow + memory-heavy to parse)
// Drop a new `.docx` into samples/ and it is discovered automatically.

import * as fs from "fs";
import * as path from "path";
import JSZip from "jszip";
import { DOMParser, XMLSerializer } from "@xmldom/xmldom";
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { parseStyleDefs, outlineNumberOf, type StyleDef } from "../src/core/outline";
import { FakePara } from "./fakeWord";

// xmldom node handles vary across versions; public signatures stay fully typed (string in/out) and we
// use a localized `any` for the node objects — the same pragmatic choice ooxmlCondense.ts/ooxmlPackage.ts make.
/* eslint-disable @typescript-eslint/no-explicit-any */

/** WordprocessingML main namespace — the only prefix the Shrink/Condense engine inspects. Exported so the
 *  real-doc test suites share ONE source of the literal (no per-file re-declaration to drift). */
export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
/** Flat-OPC package namespace, as Word's `Range.getOoxml()` emits it. */
export const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
/** DOM Node.ELEMENT_NODE — named for the same reason ooxmlPackage.ts/ooxmlCondense.ts name it (no bare `1`). */
const ELEMENT_NODE = 1;

/** rostrum/samples — one level above the add-in package (this file is in __tests__). */
export const SAMPLES_DIR = path.resolve(__dirname, "../../samples");

export type SizeTier = "small" | "medium" | "large" | "xlarge";

export interface SampleRef {
  /** Basename (e.g. `[medium] IT---ExFlex Good---AJin.docx`). */
  file: string;
  fullPath: string;
  tier: SizeTier;
  bytes: number;
}

/** Discover sample `.docx` files (skips Word lock files `~$…`), smallest first. */
export function discoverSamples(): SampleRef[] {
  if (!fs.existsSync(SAMPLES_DIR)) return [];
  return fs
    .readdirSync(SAMPLES_DIR)
    .filter((f) => f.toLowerCase().endsWith(".docx") && !f.startsWith("~$"))
    .map((file) => {
      const fullPath = path.join(SAMPLES_DIR, file);
      const bytes = fs.statSync(fullPath).size;
      return { file, fullPath, tier: tierOf(file, bytes), bytes };
    })
    // A .docx is a zip; a 0-byte file (e.g. an empty doc accidentally dropped into samples/ during a
    // wet-test) can never be a valid package — JSZip throws "Corrupted zip". Skip it with a warning
    // rather than reding the whole real-doc suite on local junk. (CI has no samples/ dir at all.)
    .filter((ref) => {
      if (ref.bytes > 0) return true;
      // eslint-disable-next-line no-console
      console.warn(`[realdoc] skipping empty/corrupt sample (0 bytes): ${ref.file}`);
      return false;
    })
    .sort((a, b) => a.bytes - b.bytes);
}

/** Tier from a leading `[tag]` (e.g. `[medium] …`), falling back to raw file size. */
function tierOf(file: string, bytes: number): SizeTier {
  const tag = (/^\[([^\]]+)\]/.exec(file)?.[1] ?? "").toLowerCase();
  if (/extremely large|x-?large|huge|full/.test(tag)) return "xlarge";
  if (/large/.test(tag)) return "large";
  if (/medium|med/.test(tag)) return "medium";
  if (/small|test|poc/.test(tag)) return "small";
  if (bytes > 3_000_000) return "xlarge";
  if (bytes > 1_000_000) return "large";
  if (bytes > 150_000) return "medium";
  return "small";
}

/** Read `word/document.xml` and `word/styles.xml` out of a `.docx` (a zip container). */
export async function readDocxParts(
  docxPath: string
): Promise<{ documentXml: string; stylesXml: string | null }> {
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const docEntry = zip.file("word/document.xml");
  if (!docEntry) throw new Error(`${path.basename(docxPath)}: missing word/document.xml`);
  const documentXml = await docEntry.async("string");
  const stylesEntry = zip.file("word/styles.xml");
  const stylesXml = stylesEntry ? await stylesEntry.async("string") : null;
  return { documentXml, stylesXml };
}

/**
 * The relationship TYPE suffixes `word/document.xml` declares (from
 * `word/_rels/document.xml.rels`) — e.g. "styles", "numbering", "theme", "fontTable",
 * "hyperlink", "image", "footnotes". Powers the Stage 4.2g coverage guard: a per-paragraph
 * `insertOoxml` commit renders against ONLY the parts `commitXml` bundles, so a body paragraph
 * that inherits formatting from an UNBUNDLED part silently collapses to defaults. Reads ONLY
 * the tiny .rels entry (never the multi-MB document.xml), so it's cheap on every tier.
 */
export async function readDocumentRelTypes(docxPath: string): Promise<string[]> {
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const relsEntry = zip.file("word/_rels/document.xml.rels");
  if (!relsEntry) return [];
  const xml = await relsEntry.async("string");
  const types = [...xml.matchAll(/Type="[^"]*\/relationships\/([^"/]+)"/g)].map((m) => m[1]);
  return [...new Set(types)].sort();
}

// `StyleDef`, `parseStyleDefs`, `styleOutlineLevel`, and `outlineNumberOf` were promoted
// into the shared `src/core/outline.ts` module (so the whole-body classify path and this
// test harness share ONE implementation) and are imported above.

/**
 * Split a real `document.xml` into engine-ready `FakePara[]` (body paragraphs in
 * document order), resolving each paragraph's outline level against `stylesXml`
 * (null = none available → outline from an inline `<w:outlineLvl>` only). Reuses the
 * production `WholeBodyPackage` so the split logic under test is the one shipped.
 * `inTable` is left `false`: a standalone `<w:p>` has lost its table context, and table
 * membership only suppresses hiding — the invariants asserted here don't depend on it.
 */
export function paragraphsFromDocumentXml(documentXml: string, stylesXml: string | null): FakePara[] {
  const defs = stylesXml ? parseStyleDefs(stylesXml) : new Map<string, StyleDef>();
  const pkg = new WholeBodyPackage(documentXml);
  const out: FakePara[] = [];
  for (let i = 0; i < pkg.count; i++) {
    const xml = pkg.paragraphXml(i);
    out.push({ xml, outlineNumber: outlineNumberOf(xml, defs), inTable: false });
  }
  return out;
}

// ---------------------------------------------------------------------------
// Flat-OPC RANGE-package reconstruction (for the Shrink real-doc suite).
//
// The Stage-4.2g per-paragraph package (`WholeBodyPackage.paragraphXml`) is deliberately style-LESS, so
// it CANNOT drive Shrink: `resolveStyleEmphasis`/`resolveNormalSizeHalfPts` regex-scan the fragment for
// `<w:style>`/`<w:docDefaults>`, and real docs apply the cut through a CHARACTER STYLE — with no styles
// part, every run resolves to "no emphasis" and the genuine cut would shrink. The helpers below instead
// rebuild the TWO-part shape Word's live `Range.getOoxml()` returns — a `/word/document.xml` part AND a
// `/word/styles.xml` part — which is exactly what Shrink consumes on the host. (`rangePort.test.ts::pkg()`
// models the same flat-OPC shape but with only the document part.)
// ---------------------------------------------------------------------------

/**
 * Strip a leading UTF-8 BOM and/or XML prolog (`<?xml …?>`) so a whole part can be embedded MID-package
 * inside `<pkg:xmlData>`. JSZip decodes parts as UTF-8 and may retain the BOM, and Word's
 * `document.xml`/`styles.xml` each begin with a prolog; an interior BOM/prolog is not well-formed, so both
 * are removed before wrapping. The explicit BOM strip is LOAD-BEARING, not decorative: the prolog regex's
 * `^\s*` would absorb a BOM only when one PRECEDES `<?xml`, so a part that has a BOM but no prolog needs
 * the dedicated strip. (JS `\s` does include U+FEFF, but only the prolog branch relies on that.)
 */
export function stripProlog(xml: string): string {
  return xml.replace(/^﻿/, "").replace(/^\s*<\?xml[^>]*\?>\s*/, "");
}

/**
 * Reconstruct a faithful flat-OPC RANGE package — a `document.xml` part PLUS a `styles.xml` part — the
 * exact two-part shape Word's live `Range.getOoxml()` returns and the ONLY shape that can drive Shrink
 * against a real doc (the engine needs both the body paragraphs and the style definitions in one
 * string). Each part keeps its own root namespace declarations (Word emits self-contained parts), so the
 * `w:` prefix resolves inside `<pkg:xmlData>`.
 */
export function buildRangePackage(documentXml: string, stylesXml: string): string {
  return (
    `<pkg:package xmlns:pkg="${PKG_NS}">` +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>${stripProlog(documentXml)}</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/styles.xml"><pkg:xmlData>${stripProlog(stylesXml)}</pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}

/**
 * Extract the first body STORY `<w:p>` whose serialized XML satisfies `predicate` (e.g. "contains
 * `rStyle="StyleUnderline"`") and wrap ONLY that paragraph in a minimal `<w:document><w:body>` that
 * REUSES the source document's verbatim `<w:document …>` start tag — so every namespace the paragraph's
 * runs reference resolves exactly as Word declared it (xmldom also re-declares the prefixes the subtree
 * uses onto the serialized `<w:p>`, so the result is robust even for a richer paragraph). Returns null
 * when no paragraph matches.
 *
 * WHY ONE PARAGRAPH: an end-to-end Shrink over a single non-heading card paragraph runs with
 * `outlineLevels: [null]` — no per-paragraph outline alignment, no heading-refusal edge — so the test
 * exercises the size ladder + keep predicate cleanly. Only TOP-LEVEL `<w:p>` children are considered
 * (mirroring the engine's body story: textbox-nested paragraphs are not part of the story).
 */
export function singleParagraphDocumentXml(
  documentXml: string,
  predicate: (paragraphXml: string) => boolean
): string | null {
  const body = stripProlog(documentXml);
  const doc: any = new DOMParser().parseFromString(body, "text/xml");
  const serializer = new XMLSerializer();
  const bodies = doc.getElementsByTagName("w:body");
  const scope: any = bodies && bodies.length > 0 ? bodies.item(0) : null;
  if (!scope) return null;

  let matchXml: string | null = null;
  const kids = scope.childNodes;
  for (let i = 0; i < kids.length; i++) {
    const k: any = kids.item ? kids.item(i) : kids[i];
    // Direct-child `<w:p>` only — top-level body paragraphs, never a textbox/table-nested one.
    if (k && k.nodeType === ELEMENT_NODE && k.nodeName === "w:p") {
      const xml = serializer.serializeToString(k);
      if (predicate(xml)) {
        matchXml = xml;
        break;
      }
    }
  }
  if (matchXml === null) return null;

  // Reuse the source `<w:document>`'s OWN start tag (every `xmlns:*` decl + `mc:Ignorable`) so the lone
  // paragraph stays in the exact namespace/MC context Word emitted — the same fidelity rationale as
  // `WholeBodyPackage.documentAttrs`. Built by SHALLOW-cloning the parsed root and serializing it (xmldom
  // escapes attribute values and re-emits every namespace correctly), then turning the empty self-closing
  // tag into an open tag — no regex over raw XML, so an attribute value can never break it. Falls back to
  // a bare `xmlns:w` only when the source has no `<w:document>` (synthetic fixtures).
  const root: any = doc.documentElement;
  let openTag = `<w:document xmlns:w="${W_NS}">`;
  if (root && root.nodeName === "w:document") {
    const shell = serializer.serializeToString(root.cloneNode(false)); // attributes only, no children
    openTag = shell.endsWith("/>") ? `${shell.slice(0, -2)}>` : shell.replace(/<\/w:document>\s*$/, "");
  }
  return `${openTag}<w:body>${matchXml}</w:body></w:document>`;
}
