// Helpers for the REAL-document test suite (realDocs.test.ts).
//
// These read the actual `.docx` files dropped into `rostrum/samples/`, pull out
// `word/document.xml` (the real OOXML Word produces) plus `word/styles.xml`, and turn
// the body paragraphs into engine-ready `FakePara` fixtures. This lets the FULL
// engine+adapter stack run against genuine debate-brief OOXML — every messy real-world
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
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { parseStyleDefs, outlineNumberOf, type StyleDef } from "../src/core/outline";
import { FakePara } from "./fakeWord";

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
