// One-off audit (Stage 4.2g validation): for every .docx in samples/, list the OOXML
// parts present and the relationship TYPES that word/document.xml declares, then flag any
// rendering-relevant part/relationship that commitXml does NOT bundle. This answers
// "could a different brief inherit formatting from a part we drop?" entirely headless.
//
// Run from rostrum-addin/:  node scripts/auditDocxParts.mjs
import fs from "fs";
import path from "path";
import { fileURLToPath } from "url";
import JSZip from "jszip";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const SAMPLES = path.resolve(__dirname, "../../samples");

// What commitXml currently bundles (must mirror AUX_*_CONTENT_TYPES / AUX_REL_TYPES).
const BUNDLED_REL_SUFFIXES = new Set(["styles", "numbering", "theme"]);
// Relationship types that affect how body content RENDERS (vs. metadata we can ignore).
const RENDER_REL_SUFFIXES = new Set([
  "styles", "numbering", "theme", "fontTable", "settings", "webSettings",
  "glossaryDocument", "customXml"
]);

const files = fs.existsSync(SAMPLES)
  ? fs.readdirSync(SAMPLES).filter((f) => f.toLowerCase().endsWith(".docx") && !f.startsWith("~$"))
  : [];

for (const file of files) {
  const zip = await JSZip.loadAsync(fs.readFileSync(path.join(SAMPLES, file)));
  const parts = Object.keys(zip.files).filter((n) => !zip.files[n].dir);

  // word parts that matter for rendering
  const wordParts = parts.filter((p) => p.startsWith("word/") && p.endsWith(".xml"));

  // document.xml.rels → relationship types referenced by the main document
  const relsEntry = zip.file("word/_rels/document.xml.rels");
  const relsXml = relsEntry ? await relsEntry.async("string") : "";
  const relTypes = [...relsXml.matchAll(/Type="[^"]*\/relationships\/([^"\/]+)"/g)].map((m) => m[1]);
  const relSet = [...new Set(relTypes)].sort();

  // does the body use THEMED fonts/colors (→ needs theme, which we bundle) or anything else?
  const docXml = await zip.file("word/document.xml").async("string");
  const usesThemeFont = /\b(ascii|hAnsi|cs|eastAsia)Theme=/.test(docXml);
  const usesThemeColor = /w:themeColor=/.test(docXml);
  const usesCharBorder = /<w:bdr\b/.test(docXml); // character border ("box")
  const usesParaBorder = /<w:pBdr\b/.test(docXml);

  // the gap: render-relevant rels the document references that commitXml does NOT bundle
  const unbundled = relSet.filter((t) => RENDER_REL_SUFFIXES.has(t) && !BUNDLED_REL_SUFFIXES.has(t));

  console.log(`\n=== ${file} ===`);
  console.log(`  word parts: ${wordParts.map((p) => p.replace("word/", "")).join(", ")}`);
  console.log(`  doc.xml.rels types: ${relSet.join(", ") || "(none)"}`);
  console.log(`  themed fonts: ${usesThemeFont} | themed colors: ${usesThemeColor} | <w:bdr> box: ${usesCharBorder} | <w:pBdr>: ${usesParaBorder}`);
  console.log(`  >>> render rels NOT bundled by commitXml: ${unbundled.length ? unbundled.join(", ") : "(none — fully covered)"}`);
}
