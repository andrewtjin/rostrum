// gdocs-synth-fixture.mjs — derive REALISTIC large-doc Docs-JSON fixtures from
// real .docx samples (plan A13 "Synthetic Docs-JSON scale fixtures generated
// dev-side from real samples/*.docx statistics" + step S13 + case 001-S6).
//
// WHY THIS EXISTS: the gdocs scale suite (__tests__/gdocsRealDocs.test.ts)
// must prove the planner's request/byte envelope BEFORE the one wet-test
// round, but no real documents.get dump exists yet. The closest available
// reality is the repo's real debate .docx corpus (rostrum/samples — the same
// gitignored corpus the Word realDocs suites consume): this tool reads each
// .docx with jszip (the house pattern from __tests__/realDocs.ts), extracts
// the per-paragraph structure the gdocs engine actually keys on, and emits a
// documents.get-SHAPED JSON in the exact masked shape parse.ts reads (the
// committed __tests__/fixtures/gdocs/*.json conventions: legacy top-level
// body/namedStyles/revisionId, body content starting at index 1 behind a
// masked sectionBreak stub, trailing newlines INSIDE final text runs, zero
// rgb channels OMITTED). Output lands in rostrum-addin/samples/gdocs/ —
// gitignored, skip-when-absent (lesson #44) — so CI never depends on it.
//
// FIDELITY CHOICES (what is modeled, and why):
//   * Run sizes resolve direct <w:sz> FIRST, then the run's character style
//     via the basedOn cascade — debate docs apply cites/underlining through
//     character styles (Style13ptBold = 14pt bold, StyleUnderline = 11pt), and
//     a Docs import MATERIALIZES those as explicit run sizes. Skipping the
//     cascade would erase every cite signature and wildly distort the RLE/
//     region statistics the envelope depends on.
//   * Paragraph styles map structurally: a direct <w:outlineLvl> wins, else
//     the paragraph style's effective level via the basedOn cascade, else the
//     built-in heading-NAME fallback — the exact resolution order of
//     src/core/outline.ts (kept in sync by comment-contract; this is an .mjs
//     tool, so it cannot import the TS module). Levels 1-6 become HEADING_n;
//     Title/Subtitle map to TITLE/SUBTITLE; everything else is NORMAL_TEXT —
//     mirroring how a Docs import assigns named styles.
//   * <w:highlight> names map to the SAME classic hexes the engine's default
//     keep set pins (gdocs/src/core/settings.ts WORD_HIGHLIGHT_HEXES — that
//     table IS the engine's model of what a .docx import materializes, so the
//     fixture must speak the same dialect or keeper statistics lie).
//   * Adjacent runs with identical (size, bold, background) merge into one
//     textRun — the real API coalesces equal-style runs, and NOT merging
//     would overstate element counts (and fixture bytes) several-fold.
//   * Table/hyperlink/ins-wrapped paragraphs and runs are flattened into the
//     body in document order. Tables are rare in this corpus (one <w:tbl> in
//     the xlarge doc) and flattening OVERSTATES hideable text, which is the
//     conservative direction for an upper-bound envelope.
//   * Not modeled (irrelevant to the planner's envelope): underline/italic
//     (no Docs channel the engine reads), images/footnotes (kind "other"
//     breaks regions — rare here), direct paragraph spacing (collapseSpacing
//     defaults OFF), w:shd paragraph shading (engine reads run backgrounds).
//
// USAGE (from rostrum-addin/):
//   node tools/gdocs-synth-fixture.mjs                    # all ../samples/*.docx
//   node tools/gdocs-synth-fixture.mjs "../samples/[small] 2ac---dds2---finals.docx" ...
//
// Output: samples/gdocs/<docx-basename>.json (compact — real wire payloads
// are compact, and the file's byte size doubles as the documents.get payload
// proxy the scale suite asserts on).

import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";
import JSZip from "jszip";

// ---------------------------------------------------------------------------
// Paths
// ---------------------------------------------------------------------------

/** rostrum-addin/ — this file lives in rostrum-addin/tools/. */
const REPO_ROOT = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
/** Where fixtures land (gitignored; created on demand). */
const OUT_DIR = path.join(REPO_ROOT, "samples", "gdocs");
/** The real .docx corpus — the OUTER workspace's samples/, exactly where the
 * Word realDocs suites discover from (__tests__/realDocs.ts SAMPLES_DIR). */
const DOCX_DIR = path.resolve(REPO_ROOT, "..", "samples");

// ---------------------------------------------------------------------------
// OOXML constants
// ---------------------------------------------------------------------------

/**
 * ECMA-376 ST_HighlightColor name -> classic hex. Values are aligned 1:1 with
 * the engine's WORD_HIGHLIGHT_HEXES (gdocs/src/core/settings.ts) — the keep
 * set those hexes seed is the whole point of emitting them: a fixture
 * highlight must be RECOGNIZED as a keeper exactly when the engine's default
 * settings would keep the real import. "white"/"black" stay mapped (the
 * engine deliberately keeps black and drops white from its DEFAULT set —
 * the fixture's job is only to report what the doc carries). "none" clears.
 */
const HIGHLIGHT_HEX = {
  black: "#000000",
  blue: "#0000ff",
  cyan: "#00ffff",
  darkBlue: "#00008b",
  darkCyan: "#008b8b",
  darkGray: "#808080",
  darkGreen: "#006400",
  darkMagenta: "#8b008b",
  darkRed: "#8b0000",
  darkYellow: "#808000",
  green: "#00ff00",
  lightGray: "#c0c0c0",
  magenta: "#ff00ff",
  red: "#ff0000",
  white: "#ffffff",
  yellow: "#ffff00"
};

/**
 * Built-in heading style NAMES -> 0-based outline level, the LAST-RESORT
 * fallback after the structural cascade (mirrors src/core/outline.ts
 * HEADING_NAME_LEVELS, including the navy debate-tag styles): Word does not
 * always re-emit <w:outlineLvl> for built-in headings into styles.xml, so a
 * purely structural read would demote real headings to body text.
 */
const HEADING_NAME_LEVELS = new Map([
  ["heading1", 0],
  ["heading2", 1],
  ["heading3", 2],
  ["heading4", 3],
  ["heading5", 4],
  ["heading6", 5],
  ["heading7", 6],
  ["heading8", 7],
  ["heading9", 8],
  ["analytics", 3],
  ["analytic", 3]
]);

// ---------------------------------------------------------------------------
// styles.xml — style definition table + cascade resolution
// ---------------------------------------------------------------------------

/**
 * Parse styleId -> definition from styles.xml. Regex-based, not DOM: Word's
 * styles.xml is flat and well-formed, and the realDocs/outline.ts precedent
 * already proved a string scan is fast and reliable on this corpus. Each
 * field is TRI-STATE (value | null = "this style does not declare it") so the
 * basedOn cascade can distinguish "declares false/0" from "silent".
 */
function parseStyleDefs(stylesXml) {
  const defs = new Map();
  const styleRe = /<w:style\b[^>]*\bw:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g;
  let m;
  while ((m = styleRe.exec(stylesXml)) !== null) {
    const inner = m[2];
    // Attribute-order tolerant on purpose (outline.ts review C-2): producers
    // reorder attributes on round-trip, so never assume w:val comes first.
    const ol = /<w:outlineLvl\b[^>]*\bw:val="(\d+)"/.exec(inner);
    const based = /<w:basedOn\b[^>]*\bw:val="([^"]+)"/.exec(inner);
    const sz = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(inner);
    const hl = /<w:highlight\b[^>]*\bw:val="([^"]+)"/.exec(inner);
    defs.set(m[1], {
      outlineLvl: ol ? Number(ol[1]) : null,
      basedOn: based ? based[1] : null,
      szHalfPt: sz ? Number(sz[1]) : null,
      bold: boldDeclared(inner),
      highlight: hl ? hl[1] : null
    });
  }
  return defs;
}

/** Tri-state <w:b> read: true (declared on), false (declared off via
 * w:val="0"/"false"), or null (not declared — defer to the cascade). */
function boldDeclared(xml) {
  const b = /<w:b\b([^>]*?)\/?>/.exec(xml);
  if (b === null) return null;
  return !/\bw:val="(?:0|false)"/.test(b[1]);
}

/**
 * Walk the basedOn chain (cycle-safe) returning the first non-null `pick`
 * value — the single cascade primitive shared by size, bold, highlight and
 * outline resolution, exactly how Word resolves style inheritance.
 */
function chainLookup(styleId, defs, pick) {
  const seen = new Set();
  let cur = styleId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const def = defs.get(cur);
    if (!def) return null;
    const v = pick(def);
    if (v !== null) return v;
    cur = def.basedOn;
  }
  return null;
}

/** Heading-name fallback along the basedOn chain (outline.ts headingNameLevel
 * parity: a custom style based on a built-in heading still counts). */
function headingNameLevel(styleId, defs) {
  const seen = new Set();
  let cur = styleId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const lvl = HEADING_NAME_LEVELS.get(cur.toLowerCase().replace(/\s+/g, ""));
    if (lvl !== undefined) return lvl;
    const def = defs.get(cur);
    cur = def ? def.basedOn : null;
  }
  return null;
}

/**
 * The Docs named style a paragraph would import as. Resolution order is the
 * outline.ts cascade: inline <w:outlineLvl> wins, then the style's structural
 * level, then the heading-name fallback; Title/Subtitle map by name (Docs has
 * dedicated named styles for them); 0-based levels 0-5 become HEADING_1..6
 * (Docs has no deeper headings); everything else is NORMAL_TEXT.
 */
function namedStyleTypeOf(inlineLvl, styleId, defs) {
  if (styleId !== null) {
    const norm = styleId.toLowerCase().replace(/\s+/g, "");
    if (norm === "title") return "TITLE";
    if (norm === "subtitle") return "SUBTITLE";
  }
  let lvl = inlineLvl;
  if (lvl === null && styleId !== null) {
    lvl = chainLookup(styleId, defs, (d) => d.outlineLvl);
    if (lvl === null) lvl = headingNameLevel(styleId, defs);
  }
  // OOXML levels are 0-based; 9 (or absent) means body text.
  if (lvl !== null && lvl >= 0 && lvl <= 5) return `HEADING_${lvl + 1}`;
  return "NORMAL_TEXT";
}

// ---------------------------------------------------------------------------
// document.xml — paragraph/run extraction
// ---------------------------------------------------------------------------

/** Decode the XML entities Word emits inside <w:t> (the five named ones plus
 * numeric character references). Indexes downstream are UTF-16 code units, so
 * String.fromCodePoint keeps surrogate-pair widths honest. */
function decodeEntities(s) {
  return s.replace(/&(amp|lt|gt|quot|apos|#x[0-9a-fA-F]+|#\d+);/g, (_, code) => {
    switch (code) {
      case "amp":
        return "&";
      case "lt":
        return "<";
      case "gt":
        return ">";
      case "quot":
        return '"';
      case "apos":
        return "'";
      default:
        return String.fromCodePoint(
          code.startsWith("#x") ? parseInt(code.slice(2), 16) : parseInt(code.slice(1), 10)
        );
    }
  });
}

/** Visible text of one run: <w:t> content (entity-decoded) plus 1-char stand-
 * ins for tab/break — each occupies exactly one index in Docs too ("\u000b"
 * is the literal vertical-tab Docs uses for in-paragraph line breaks). */
function runText(runXml) {
  let out = "";
  const re = /<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>|<w:(tab|br|cr)\b[^>]*\/>/g;
  let m;
  while ((m = re.exec(runXml)) !== null) {
    if (m[1] !== undefined) out += decodeEntities(m[1]);
    else out += m[2] === "tab" ? "\t" : "\u000b";
  }
  return out;
}

/**
 * Extract one paragraph's merged run list: [{ text, sizePt, bold, bgHex }].
 * Direct rPr properties win; silent channels fall to the run's character
 * style via the cascade (see FIDELITY CHOICES); still-silent channels stay
 * null/false = "inherits from the named style", which is exactly what
 * parse.ts decodes an absent textStyle field as.
 */
function extractRuns(paraXml, defs) {
  const runs = [];
  const runRe = /<w:r(?:\s[^>]*)?>([\s\S]*?)<\/w:r>/g;
  let m;
  while ((m = runRe.exec(paraXml)) !== null) {
    const inner = m[1];
    const text = runText(inner);
    if (text === "") continue; // property-only / empty runs occupy no indexes
    // rPr is the run's FIRST child when present; scanning the whole run inner
    // is safe because w:t bodies are entity-escaped (no raw '<').
    const rPr = /<w:rPr>[\s\S]*?<\/w:rPr>/.exec(inner)?.[0] ?? "";
    const rStyle = /<w:rStyle\b[^>]*\bw:val="([^"]+)"/.exec(rPr)?.[1] ?? null;
    const directSz = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(rPr);
    const szHalfPt =
      directSz !== null
        ? Number(directSz[1])
        : rStyle !== null
          ? chainLookup(rStyle, defs, (d) => d.szHalfPt)
          : null;
    const directBold = boldDeclared(rPr);
    const bold =
      directBold !== null
        ? directBold
        : rStyle !== null
          ? (chainLookup(rStyle, defs, (d) => d.bold) ?? false)
          : false;
    const hlName =
      /<w:highlight\b[^>]*\bw:val="([^"]+)"/.exec(rPr)?.[1] ??
      (rStyle !== null ? chainLookup(rStyle, defs, (d) => d.highlight) : null);
    const bgHex = hlName !== null && hlName !== "none" ? (HIGHLIGHT_HEX[hlName] ?? null) : null;

    const sizePt = szHalfPt !== null ? szHalfPt / 2 : null;
    const prev = runs[runs.length - 1];
    // Merge equal-style neighbors — the real API coalesces them, and element
    // count / fixture bytes are measurements the scale suite reports on.
    if (prev !== undefined && prev.sizePt === sizePt && prev.bold === bold && prev.bgHex === bgHex) {
      prev.text += text;
    } else {
      runs.push({ text, sizePt, bold, bgHex });
    }
  }
  return runs;
}

// ---------------------------------------------------------------------------
// Docs-JSON assembly (the masked documents.get shape parse.ts reads)
// ---------------------------------------------------------------------------

/** "#rrggbb" -> proto3-shaped rgbColor with ZERO channels OMITTED — the wire
 * convention the fixture lint pins (fakeDocs.ts rgbColorOf parity). */
function rgbColorOf(hex) {
  const out = {};
  const channel = (offset) => parseInt(hex.slice(offset, offset + 2), 16) / 255;
  const r = channel(1);
  const g = channel(3);
  const b = channel(5);
  if (r > 0) out.red = r;
  if (g > 0) out.green = g;
  if (b > 0) out.blue = b;
  return out;
}

/** One run -> the wire textStyle. Absent keys ARE the inherit/false encoding
 * (real payloads omit false booleans and unset fields — parse.ts decodes by
 * presence), so this never writes `bold: false` or a null size. */
function textStyleOf(run) {
  const style = {};
  if (run.sizePt !== null) style.fontSize = { magnitude: run.sizePt, unit: "PT" };
  if (run.bold) style.bold = true;
  if (run.bgHex !== null) style.backgroundColor = { color: { rgbColor: rgbColorOf(run.bgHex) } };
  return style;
}

/**
 * Build the full documents.get-shaped JSON for one .docx. Index discipline is
 * the committed-fixture convention exactly: body content starts at index 1
 * behind the masked sectionBreak stub, every paragraph's trailing newline
 * lives INSIDE its final text run, and indexes are UTF-16 code units
 * (JS .length semantics — surrogate pairs count 2 for free).
 */
function buildDocsJson(documentXml, stylesXml, name) {
  const defs = parseStyleDefs(stylesXml ?? "");

  const content = [{ endIndex: 1 }]; // the masked sectionBreak stub
  let cursor = 1;
  let paragraphCount = 0;
  let elementCount = 0;

  // Every <w:p> in document order — self-closing (empty) paragraphs included.
  // Top-level body, table-cell and hyperlink-wrapped paragraphs all flatten
  // into one body sequence (see FIDELITY CHOICES for why that is acceptable
  // and conservative here). `</w:p>` cannot false-match `</w:pPr>` (the next
  // char there is "P", not ">"), so the lazy body scan is exact.
  const paraRe = /<w:p\b[^>]*\/>|<w:p\b[^>]*>([\s\S]*?)<\/w:p>/g;
  let m;
  while ((m = paraRe.exec(documentXml)) !== null) {
    const inner = m[1] ?? ""; // self-closing <w:p/> = empty paragraph
    // Paragraph properties: pPr is the first child; pStyle/outlineLvl only
    // ever appear there, so scanning the whole inner is unambiguous (a run
    // can carry neither).
    const pPr = /<w:pPr>[\s\S]*?<\/w:pPr>/.exec(inner)?.[0] ?? "";
    const styleId = /<w:pStyle\b[^>]*\bw:val="([^"]+)"/.exec(pPr)?.[1] ?? null;
    const inlineOl = /<w:outlineLvl\b[^>]*\bw:val="(\d+)"/.exec(pPr);
    const namedStyleType = namedStyleTypeOf(inlineOl ? Number(inlineOl[1]) : null, styleId, defs);

    const runs = extractRuns(inner, defs);
    // The trailing newline rides in the LAST text run (API reality); an empty
    // paragraph is exactly one 1-char inheriting run.
    if (runs.length > 0) runs[runs.length - 1].text += "\n";
    else runs.push({ text: "\n", sizePt: null, bold: false, bgHex: null });

    const startIndex = cursor;
    const elements = runs.map((run) => {
      const el = {
        startIndex: cursor,
        endIndex: cursor + run.text.length,
        textRun: { content: run.text, textStyle: textStyleOf(run) }
      };
      cursor = el.endIndex;
      return el;
    });
    content.push({
      startIndex,
      endIndex: cursor,
      paragraph: { elements, paragraphStyle: { namedStyleType } }
    });
    paragraphCount++;
    elementCount += elements.length;
  }

  // namedStyles: the per-style sizes parse.ts resolves inherited run sizes
  // through. NORMAL_TEXT falls back docDefaults-ward exactly like Word does;
  // heading/Title sizes come from their style chain when stated.
  const docDefaultsXml = /<w:docDefaults>[\s\S]*?<\/w:docDefaults>/.exec(stylesXml ?? "")?.[0] ?? "";
  const docDefaultSz = /<w:sz\b[^>]*\bw:val="(\d+)"/.exec(docDefaultsXml);
  const styles = [];
  const pushStyle = (type, halfPt) => {
    if (halfPt !== null) styles.push({ namedStyleType: type, textStyle: { fontSize: { magnitude: halfPt / 2, unit: "PT" } } });
  };
  pushStyle(
    "NORMAL_TEXT",
    chainLookup("Normal", defs, (d) => d.szHalfPt) ?? (docDefaultSz ? Number(docDefaultSz[1]) : 22)
  );
  for (let level = 1; level <= 6; level++) {
    pushStyle(`HEADING_${level}`, chainLookup(`Heading${level}`, defs, (d) => d.szHalfPt));
  }
  pushStyle("TITLE", chainLookup("Title", defs, (d) => d.szHalfPt));
  pushStyle("SUBTITLE", chainLookup("Subtitle", defs, (d) => d.szHalfPt));

  return {
    json: {
      revisionId: `synth-${name}`,
      body: { content },
      namedStyles: { styles }
    },
    paragraphCount,
    elementCount,
    charCount: cursor - 1
  };
}

// ---------------------------------------------------------------------------
// Driver
// ---------------------------------------------------------------------------

/** Read the two parts the extractor needs out of one .docx (the realDocs.ts
 * jszip pattern; styles.xml is optional there and optional here). */
async function readDocxParts(docxPath) {
  const zip = await JSZip.loadAsync(fs.readFileSync(docxPath));
  const docEntry = zip.file("word/document.xml");
  if (!docEntry) throw new Error(`${path.basename(docxPath)}: missing word/document.xml`);
  const documentXml = await docEntry.async("string");
  const stylesEntry = zip.file("word/styles.xml");
  const stylesXml = stylesEntry ? await stylesEntry.async("string") : null;
  return { documentXml, stylesXml };
}

/** Default input set: every real .docx in the corpus, Word lock files and
 * empty/corrupt zips skipped (the realDocs discoverSamples conventions). */
function discoverDocx() {
  if (!fs.existsSync(DOCX_DIR)) return [];
  return fs
    .readdirSync(DOCX_DIR)
    .filter((f) => f.toLowerCase().endsWith(".docx") && !f.startsWith("~$"))
    .map((f) => path.join(DOCX_DIR, f))
    .filter((p) => fs.statSync(p).size > 0);
}

async function main() {
  const args = process.argv.slice(2);
  const inputs = args.length > 0 ? args : discoverDocx();
  if (inputs.length === 0) {
    console.error(`gdocs-synth-fixture: no .docx inputs (looked in ${DOCX_DIR})`);
    process.exitCode = 1;
    return;
  }
  fs.mkdirSync(OUT_DIR, { recursive: true });

  for (const input of inputs) {
    const base = path.basename(input).replace(/\.docx$/i, "");
    try {
      const { documentXml, stylesXml } = await readDocxParts(input);
      const { json, paragraphCount, elementCount, charCount } = buildDocsJson(documentXml, stylesXml, base);
      const outPath = path.join(OUT_DIR, `${base}.json`);
      // Compact on purpose: the file size is the suite's documents.get payload
      // proxy, and real wire payloads carry no pretty-print whitespace.
      fs.writeFileSync(outPath, JSON.stringify(json));
      const bytes = fs.statSync(outPath).size;
      console.log(
        `[synth] ${base}.json: ${paragraphCount} paragraphs, ${elementCount} elements, ` +
          `${charCount} chars, ${(bytes / 1024).toFixed(0)} KB`
      );
    } catch (e) {
      // One bad input must not abort the rest of the batch — report and mark
      // the process failed so a scripted caller still sees the problem.
      console.error(`[synth] FAILED ${input}: ${e instanceof Error ? e.message : String(e)}`);
      process.exitCode = 1;
    }
  }
}

await main();
