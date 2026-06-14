// SYNTHETIC, license-clean flat-OPC package fixtures for the Loop 002 engine proof
// harness (PLAN.md §8 tests-1; CASES.md 002-S10).
//
// WHY THIS EXISTS. `samples/` is EMPTY and gitignored on this checkout, so every
// "realDocs byte-gate" is a vacuous no-op in CI (tests-1). These fixtures make the
// semantic-diff oracle (§8 loss-1) and the serialize-baseline control (§8 loss-6) run
// in CI ALWAYS, proving the MECHANICS on every named OOXML shape the engine must
// survive. They are NOT a substitute for the user's real ndca/dds2/IPR/xlarge corpus
// (002-S10 binds: real-corpus byte-equality is proven only where discoverSamples()>0
// or in the wet session) — they prove the instruments work, not that the engine is
// fast or that Word's true bytes round-trip.
//
// NO REAL DEBATE CONTENT. Every string of prose here is invented placeholder text
// ("alpha beta gamma", "the quick brown fox", …); nothing is copied from any speech
// doc, card, or third-party source. License-clean by construction.
//
// AUTHORING CONTRACT (load-bearing — read before editing a fixture):
//   * Each fixture is a complete flat-OPC `<pkg:package>` that TODAY's
//     `new WholeBodyPackage(pkg)` parses WITHOUT throwing. `engineFixtures.test`-style
//     consumers and `serializeBaseline.test.ts` both construct every fixture, so a
//     malformed one fails loudly at construction, never silently.
//   * Each fixture is written in @xmldom/xmldom's CANONICAL serialization form so that
//     `serialize(parse(pkg)) === pkg` byte-for-byte (verified: xmldom 0.9.x re-emits
//     prefixed elements, self-closing empty tags `<w:x/>`, and source-order attributes
//     unchanged; it adds no XML declaration). This is what lets `serializeBaseline.test`
//     pin `serialize()` to the fixture string itself — the P4-REV byte reference is
//     "today's serialize() output on this input," NEVER Word's original .docx bytes
//     (loss-6: xmldom never reproduces those).
//     CONSEQUENCE: do NOT pretty-print, reorder attributes, or add an XML declaration
//     to a fixture, or the round-trip identity (and the baseline test) breaks.
//   * Use the SAME parser the engine uses (@xmldom/xmldom) for any verification — the
//     oracle must see what the engine sees.

// The WordprocessingML main namespace — the only prefix the engine inspects.
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
// The flat-OPC package namespace Word's getOoxml() emits.
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
// The officeDocument relationship-type base (hyperlink / styles / numbering types).
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
// The package-relationships namespace (inside a `.rels` part's `<Relationships>`).
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
// Markup-compatibility namespace (mc:AlternateContent / mc:Choice / mc:Fallback).
const MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
// WordprocessingDrawing namespace (the `<wp:inline>` wrapper around a drawing).
const WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
// VML namespace (the legacy `<v:shape>` inside an `mc:Fallback`'s `<w:pict>`). Word
// declares this on `<w:document>`; an undeclared `v:` prefix is a hard parse error in
// @xmldom/xmldom 0.9.x, so the fixture must declare it.
const V = "urn:schemas-microsoft-com:vml";

const RELS_CT = "application/vnd.openxmlformats-package.relationships+xml";
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const DOC_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const SETTINGS_CT =
  "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

// ---------------------------------------------------------------------------
// Small builders. These keep every fixture in the ONE canonical shape that
// round-trips through xmldom, so the authoring contract above holds by
// construction rather than by per-fixture vigilance (DRY).
// ---------------------------------------------------------------------------

/**
 * Wrap document-body inner XML (the `<w:p>` run, optionally a trailing `<w:sectPr/>`)
 * plus any extra non-document `<pkg:part>`s into a flat-OPC package. The document part
 * carries whatever extra xmlns prefixes a fixture needs via `docNs` (so e.g. the
 * hyperlink fixture can declare `xmlns:r`). `extraParts` is concatenated AHEAD of the
 * document part exactly the way Word orders relationship/aux parts before document.xml
 * in some emissions — the engine reads parts by name/type, not position, so order only
 * needs to be STABLE (it is: a literal string).
 */
function pkg(opts: {
  body: string;
  docNs?: string; // extra attrs on <w:document>, e.g. ` xmlns:r="..."` (leading space included)
  extraParts?: string; // serialized <pkg:part> elements to place before /word/document.xml
}): string {
  const docNs = opts.docNs ?? "";
  const extra = opts.extraParts ?? "";
  return (
    `<pkg:package xmlns:pkg="${PKG}">` +
    extra +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CT}"><pkg:xmlData>` +
    `<w:document xmlns:w="${W}"${docNs}><w:body>${opts.body}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}

/** A `.rels` part (package- or document-level) carrying the given `<Relationship>` XML. */
function relsPart(name: string, relationships: string): string {
  return (
    `<pkg:part pkg:name="${name}" pkg:contentType="${RELS_CT}"><pkg:xmlData>` +
    `<Relationships xmlns="${RELS}">${relationships}</Relationships>` +
    `</pkg:xmlData></pkg:part>`
  );
}

/** A trailing section-properties element — every real body ends with one. */
const SECT = `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;

// ===========================================================================
// The fixtures. One named export per shape; each commented with what it
// exercises and which engine path / case it backs.
// ===========================================================================

/**
 * `bridgeGap` — two visible words ("alpha", "omega") separated by a single hidden run
 * whose text is " beta gamma delta " (leading+trailing space, interior words). This is
 * the canonical input for the cross-gap separator logic (keepers.ts
 * `planCrossGapSeparators` + ooxml.ts `exposeBoundarySpace`): hiding the middle run
 * would fuse "alpha"+"omega" into "alphaomega" unless ONE space is MOVED out into a new
 * visible `xml:space="preserve"` run. The oracle's bridge-split whitelist (loss-1 (b))
 * is exercised against exactly this shape. The two anchors are their OWN runs so the
 * planner sees a real gap between word-bearing kept runs.
 */
export const bridgeGap = pkg({
  body:
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">alpha</w:t></w:r>` +
    `<w:r><w:t xml:space="preserve"> beta gamma delta </w:t></w:r>` +
    `<w:r><w:t xml:space="preserve">omega</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `rPrLessRuns` — runs with NO `<w:rPr>` at all. Exercises the engine path that must
 * CREATE an rPr to carry `<w:vanish/>` (ooxml.ts `ensureRPr`) and the oracle's rule
 * that "an rPr added solely to hold a single `<w:vanish>`" is a permitted delta even
 * when the input run had no rPr. A pure body paragraph the classifier would fully hide.
 */
export const rPrLessRuns = pkg({
  body:
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">the quick brown fox</w:t></w:r>` +
    `<w:r><w:t xml:space="preserve"> jumps over</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `mediaDrawing` — a run carrying a `<w:drawing>` (an inline image). Such a run is
 * INELIGIBLE to hide (ooxml.ts INELIGIBLE_RUN_TAGS / `runIsEligible`), so the engine
 * must keep it whole; the oracle must see the drawing subtree UNCHANGED input vs output
 * (no vanish added anywhere inside it). Backs the rider in PLAN §2 B1 (internal-part
 * detection) and 002-F1 (structural content never modified).
 */
export const mediaDrawing = pkg({
  docNs: ` xmlns:wp="${WP}"`,
  body:
    `<w:p>` +
    `<w:r><w:drawing><wp:inline><wp:extent cx="100" cy="100"/></wp:inline></w:drawing></w:r>` +
    `<w:r><w:t xml:space="preserve">figure caption text</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `hyperlinkDense` — several `<w:hyperlink r:id="…">` runs plus the matching
 * `/word/_rels/document.xml.rels` part. Exercises relationship handling: a fragment
 * carrying a hyperlink must keep its `r:id` resolvable (ooxmlPackage.ts
 * `referencedRelIds` / `collectRelationships`), and the oracle must treat the
 * `<w:hyperlink>` wrapper + its `r:id` as structure that does not change. The body text
 * inside each hyperlink is ordinary prose the classifier may hide.
 */
export const hyperlinkDense = pkg({
  docNs: ` xmlns:r="${R}"`,
  extraParts: relsPart(
    "/word/_rels/document.xml.rels",
    `<Relationship Id="rId1" Type="${R}/hyperlink" Target="https://example.org/one" TargetMode="External"/>` +
      `<Relationship Id="rId2" Type="${R}/hyperlink" Target="https://example.org/two" TargetMode="External"/>` +
      `<Relationship Id="rId3" Type="${R}/hyperlink" Target="https://example.org/three" TargetMode="External"/>`
  ),
  body:
    `<w:p>` +
    `<w:hyperlink r:id="rId1"><w:r><w:t xml:space="preserve">first source</w:t></w:r></w:hyperlink>` +
    `<w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
    `<w:hyperlink r:id="rId2"><w:r><w:t xml:space="preserve">second source</w:t></w:r></w:hyperlink>` +
    `<w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
    `<w:hyperlink r:id="rId3"><w:r><w:t xml:space="preserve">third source</w:t></w:r></w:hyperlink>` +
    `</w:p>` +
    SECT
});

/**
 * `fldSimple` — a `<w:fldSimple>` simple field (a PAGE field) wrapping a run. A run that
 * is the RESULT of a simple field is INELIGIBLE to hide (ooxml.ts `runIsEligible` walks
 * ancestors for `<w:fldSimple>`), so the engine keeps it; the oracle must see the field
 * subtree unchanged. Backs 002-S3's fldSimple fixture requirement.
 */
export const fldSimple = pkg({
  body:
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">page </w:t></w:r>` +
    `<w:fldSimple w:instr=" PAGE "><w:r><w:t xml:space="preserve">1</w:t></w:r></w:fldSimple>` +
    `</w:p>` +
    SECT
});

/**
 * `mcAlternateContent` — a `<w:drawing>` inside `mc:AlternateContent`'s `mc:Choice`,
 * with an `mc:Fallback` carrying a `<w:pict>`. The whole-paragraph subtree scan must
 * detect the internal part (drawing/pict) anywhere in the run and keep it ineligible
 * (the "mc:AlternateContent-nested drawing" case named in PLAN §2 A3 / 002-S3). The
 * oracle must see both the Choice and the Fallback unchanged. `mc:Ignorable` is set so
 * the MC context is the one Word emits.
 */
export const mcAlternateContent = pkg({
  docNs: ` xmlns:mc="${MC}" xmlns:wp="${WP}" xmlns:v="${V}" mc:Ignorable="wp"`,
  body:
    `<w:p>` +
    `<w:r><mc:AlternateContent>` +
    `<mc:Choice Requires="wp"><w:drawing><wp:inline><wp:extent cx="50" cy="50"/></wp:inline></w:drawing></mc:Choice>` +
    `<mc:Fallback><w:pict><v:shape id="_x0000_s1026" style="width:50pt;height:50pt"/></w:pict></mc:Fallback>` +
    `</mc:AlternateContent></w:r>` +
    `<w:r><w:t xml:space="preserve">diagram caption</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `styledNumbering` — a `/word/numbering.xml` part plus a paragraph whose ONLY path to a
 * style is a `<w:numPr>`/`<w:pStyle>` backlink (no inline `<w:outlineLvl>`). This is the
 * "styled-numbering" shape: outline level (and, for P7, style closure) must resolve
 * through the numbering/styles backlinks, never through inline props. The styles part
 * defines a `ListParagraph` style the paragraph references via `<w:pStyle>`, and the
 * numbering part defines the list the paragraph references via `<w:numId>`. Construction
 * proves WholeBodyPackage's `parseStyleDefs` + `headingLevel` tolerate this shape.
 */
export const styledNumbering = pkg({
  extraParts:
    `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
    `<w:styles xmlns:w="${W}">` +
    `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>` +
    `</w:styles>` +
    `</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="${NUMBERING_CT}"><pkg:xmlData>` +
    `<w:numbering xmlns:w="${W}">` +
    `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pStyle w:val="ListParagraph"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>` +
    `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
    `</w:numbering>` +
    `</pkg:xmlData></pkg:part>`,
  body:
    `<w:p>` +
    `<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
    `<w:r><w:t xml:space="preserve">a numbered list item body</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `trackChangesOn` — a `/word/settings.xml` part carrying `<w:trackChanges/>`. The Hide
 * gate (guards.ts) refuses to run under Track Changes; this fixture lets the COM harness
 * and any prefetched-TC test see a package whose bundled settings declare TC on (PLAN §2
 * B2 / 002-S4). The body is ordinary prose. The settings part is deliberately NOT in
 * AUX_PART_CONTENT_TYPES, so the engine never re-imposes it on a commit — but it must
 * still parse cleanly here.
 */
export const trackChangesOn = pkg({
  extraParts:
    `<pkg:part pkg:name="/word/settings.xml" pkg:contentType="${SETTINGS_CT}"><pkg:xmlData>` +
    `<w:settings xmlns:w="${W}"><w:trackChanges/></w:settings>` +
    `</pkg:xmlData></pkg:part>`,
  body:
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">body under track changes</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `astralText` — one run containing an astral (supplementary-plane) character, a grinning
 * face U+1F600 represented in the source as the surrogate pair, plus a SECOND run that
 * deliberately carries a LONE surrogate (U+D800 with no trailing low surrogate) as a
 * standalone edge case. This stresses the A3 separator fast-path's surrogate handling
 * (002-F9: charCodeAt is per-UTF-16-unit, `[...text]` iterates code points): the
 * classifier's whitespace/separator decision must agree between the legacy regex and the
 * charCodeAt path on BOTH the astral and the lone-surrogate text. The oracle only checks
 * that the concatenated text is byte-identical, which it is (both survive an xmldom
 * round-trip verbatim — verified).
 */
export const astralText = pkg({
  body:
    `<w:p>` +
    `<w:r><w:t xml:space="preserve">grin \u{1F600} end</w:t></w:r>` +
    `<w:r><w:t xml:space="preserve">lone\uD800surrogate</w:t></w:r>` +
    `</w:p>` +
    SECT
});

/**
 * `multiHeading` — Heading 1–4 paragraphs (kept by the outline rule, levels 0–3), a body
 * paragraph (hidden), and a bold cite paragraph (kept by the cite rule). This is the
 * many-paragraph control the oracle iterates over to prove paragraph COUNT and per-index
 * structure are preserved across a whole-body hide, and that headings/cites are untouched
 * while body runs gain `<w:vanish>`. Heading levels are declared via inline
 * `<w:outlineLvl>` so the fixture is self-contained (no styles part needed to classify).
 * The cite paragraph carries an explicit `<w:rStyle w:val="…">` plus bold; the oracle does
 * not classify, but downstream P1/A3 tests reuse this fixture.
 */
export const multiHeading = pkg({
  body:
    `<w:p><w:pPr><w:outlineLvl w:val="0"/></w:pPr><w:r><w:t xml:space="preserve">Heading One</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t xml:space="preserve">Heading Two</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t xml:space="preserve">Heading Three</w:t></w:r></w:p>` +
    `<w:p><w:pPr><w:outlineLvl w:val="3"/></w:pPr><w:r><w:t xml:space="preserve">Heading Four</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">ordinary body card text that the engine hides</w:t></w:r></w:p>` +
    `<w:p><w:r><w:rPr><w:rStyle w:val="CiteChar"/><w:b/></w:rPr><w:t xml:space="preserve">Author 2020 Journal</w:t></w:r></w:p>` +
    SECT
});

/**
 * All fixtures keyed by name — for table-driven iteration in the baseline + oracle
 * self-tests, and for any downstream P1/A2/A3/A4 test that wants "every shape." Keeping
 * the names here in ONE place means a new fixture is automatically covered by every
 * loop over `ALL_FIXTURES` (DRY; no per-test fixture lists to keep in sync).
 */
export const ALL_FIXTURES: Readonly<Record<string, string>> = {
  bridgeGap,
  rPrLessRuns,
  mediaDrawing,
  hyperlinkDense,
  fldSimple,
  mcAlternateContent,
  styledNumbering,
  trackChangesOn,
  astralText,
  multiHeading
};
