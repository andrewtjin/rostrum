// Stage B — PROACTIVE separator audit. The fusion class (two kept words separated only by
// HIDDEN text whose lone separator char isn't recognized) has been hit twice one-at-a-time
// (NBSP, em-dash). Instead of waiting for the next wet-test, sweep the real corpus headless for
// EVERY character that sits BETWEEN two word characters (`\w X \w`) — exactly the position where,
// if X lands in a hidden gap, the visible words would fuse. Rank by frequency, show contexts, and
// flag which candidates the keepers.ts WHITESPACE (word-separator) predicate already covers.
//
// This does NOT change behavior — it tells us which separators are REAL in debate docs so we can
// add the genuine ones to the predicate (or confirm the predicate is already complete) in one pass.
//
// Run from rostrum-addin/:  node scripts/auditSeparators.mjs
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SAMPLES = path.resolve(__dirname, "../../samples");

// Mirror keepers.ts WHITESPACE (the current word-separator predicate) so the audit can mark which
// candidates are ALREADY handled vs genuinely uncovered. Keep in sync if the predicate changes.
const COVERED = new Set([
  " ", "\t", "\r", "\n", "\f", "\v",
  " ", " ", " ", " ", // NBSP family
  "–", "—" // en / em dash
]);

// Sentence/clause punctuation handled by the HUGS_LEFT/HUGS_RIGHT bridge rules (they ATTACH to a
// neighbour rather than needing exposure as a free-standing separator). Reported separately so the
// "uncovered separators" list isn't drowned by ordinary punctuation that's already accounted for.
const HUGS = new Set([".", ",", ";", ":", "!", "?", ")", "]", "}", "%", "”", "’", "(", "[", "{", "“", "‘", "'", '"']);

const isWordChar = (ch) => /[\p{L}\p{N}]/u.test(ch);

const ENTITIES = { "&amp;": "&", "&lt;": "<", "&gt;": ">", "&quot;": '"', "&apos;": "'" };
const decodeEntities = (s) => s.replace(/&(amp|lt|gt|quot|apos);/g, (m) => ENTITIES[m]);

const files = fs.existsSync(SAMPLES)
  ? fs.readdirSync(SAMPLES).filter((f) => f.toLowerCase().endsWith(".docx") && !f.startsWith("~$")).sort()
  : [];

if (!files.length) {
  console.log(`No .docx samples found in ${SAMPLES}`);
  process.exit(0);
}

// Aggregate across the whole corpus: char → { count, files:Set, samples:[] }.
const interword = new Map();
const bump = (ch, file, context) => {
  let e = interword.get(ch);
  if (!e) {
    e = { count: 0, files: new Set(), samples: [] };
    interword.set(ch, e);
  }
  e.count++;
  e.files.add(file);
  if (e.samples.length < 4 && !e.samples.includes(context)) e.samples.push(context);
};

for (const file of files) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(SAMPLES, file)));
  const docEntry = zip.file("word/document.xml");
  if (!docEntry) continue;
  const xml = await docEntry.async("string");
  // Concatenate visible run text per paragraph (cheap regex extract of <w:t> contents), then scan
  // each paragraph's text for `wordchar X wordchar`. Per-paragraph (not whole-doc) so a paragraph
  // break never manufactures a false adjacency.
  for (const pMatch of xml.matchAll(/<w:p\b[\s\S]*?<\/w:p>/g)) {
    const paraXml = pMatch[0];
    const text = [...paraXml.matchAll(/<w:t\b[^>]*>([\s\S]*?)<\/w:t>/g)]
      .map((m) => decodeEntities(m[1]))
      .join("");
    const chars = [...text];
    for (let i = 1; i < chars.length - 1; i++) {
      const c = chars[i];
      if (isWordChar(c)) continue;
      if (isWordChar(chars[i - 1]) && isWordChar(chars[i + 1])) {
        const context = chars.slice(Math.max(0, i - 12), i + 13).join("");
        bump(c, file, context);
      }
    }
  }
}

const codeOf = (ch) => "U+" + ch.codePointAt(0).toString(16).toUpperCase().padStart(4, "0");
const ranked = [...interword.entries()].sort((a, b) => b[1].count - a[1].count);

console.log(`Separator audit over ${files.length} sample(s): characters appearing BETWEEN two word chars (\\w X \\w).`);
console.log(`COVERED = already in the keepers.ts word-separator predicate; HUGS = handled by bridge hug rules.\n`);

const uncovered = [];
for (const [ch, e] of ranked) {
  const tag = COVERED.has(ch) ? "COVERED " : HUGS.has(ch) ? "HUGS    " : "UNCOVERED";
  if (tag === "UNCOVERED") uncovered.push(ch);
  const display = ch === "\t" ? "\\t" : ch === "\n" ? "\\n" : ch === "\r" ? "\\r" : ch;
  console.log(
    `  [${tag}] ${codeOf(ch)} '${display}'  ×${e.count}  in ${e.files.size}/${files.length} file(s)` +
      `\n      e.g. ${e.samples.map((s) => JSON.stringify(s)).join("  ")}`
  );
}

console.log(`\n=== UNCOVERED inter-word characters (fusion candidates) ===`);
if (!uncovered.length) {
  console.log(`  (none) — every inter-word separator in the corpus is already covered or hug-handled.`);
} else {
  for (const ch of uncovered) {
    console.log(`  ${codeOf(ch)} '${ch}'  ×${interword.get(ch).count}`);
  }
  console.log(`\n  Decide per char: a genuine WORD SEPARATOR (exposing it prevents fusion, e.g. '/') →`);
  console.log(`  add to keepers.ts WHITESPACE; a WORD-INTERNAL char (e.g. '-') → leave EXCLUDED.`);
}
