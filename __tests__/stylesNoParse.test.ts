// LOOP 002 A2 — STYLES-NO-DOM-PARSE proof (PLAN §2 A2 / CASES 002-S2).
//
// P4-REVISED's read-side win (perf-1) is that the styles part — the single largest part in a real
// flat-OPC package (≈1.27MB on a debate doc) — is NEVER DOM-parsed: its xmlData SUBSTRING goes
// straight to the regex-only `parseStyleDefs`. The old `extractStylesXml` DOM-parsed the whole
// package (styles included) and then serialized the `<w:styles>` subtree back to a string just to
// hand it to the SAME regex resolver — pure wasted parse+serialize of the biggest part.
//
// This locks that in falsifiably. It spies on the real xmldom `DOMParser.parseFromString` (counting
// every parse while delegating to the genuine impl, like parseCount.test.ts) and asserts that across
// (a) constructing the `WholeBodyPackage` and (b) running `headingLevel` over every paragraph — the
// full read-side classify surface that consumes the styles — NO parsed `source` ever contains the
// styles part's style definitions. A reintroduced styles DOM-parse (e.g. resurrecting
// extractStylesXml) fails HERE. It also asserts the regex resolver still works (the styled paragraph
// resolves its outline level from the unparsed styles substring), so a zero-styles-parse path that
// dropped the win's CORRECTNESS could never pass.

/* eslint-disable @typescript-eslint/no-explicit-any */

// A unique marker string placed INSIDE a <w:style> definition. If the styles part is ever DOM-parsed,
// its source string (carrying this marker) flows through parseFromString and the spy below sees it.
const STYLE_MARKER = "RostrumStylesNoParseMarker";

const parsedSources: string[] = [];
jest.mock("@xmldom/xmldom", () => {
  const actual = jest.requireActual("@xmldom/xmldom");
  class RecordingDOMParser {
    private readonly inner: any;
    constructor(opts?: any) {
      this.inner = new actual.DOMParser(opts);
    }
    parseFromString(source: string, mimeType: string): any {
      parsedSources.push(source);
      return this.inner.parseFromString(source, mimeType);
    }
  }
  return { ...actual, DOMParser: RecordingDOMParser };
});

import { WholeBodyPackage } from "../src/core/ooxmlPackage";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * A flat-OPC package whose styles part defines a real `basedOn` cascade (Analytics → Heading4 →
 * level 3) and carries STYLE_MARKER inside a `<w:style>` so any DOM-parse of the styles part is
 * detectable in the recorded sources. The document references the styled paragraph via `<w:pStyle>`.
 */
function styledPackage(): string {
  return (
    `<pkg:package xmlns:pkg="${PKG}">` +
    `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
    `<w:styles xmlns:w="${W}">` +
    `<w:style w:type="paragraph" w:styleId="Heading4" w:default="0"><w:name w:val="${STYLE_MARKER} Heading 4"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Analytics"><w:name w:val="${STYLE_MARKER} Analytics"/><w:basedOn w:val="Heading4"/></w:style>` +
    `<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>` +
    `</w:styles>` +
    `</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"><pkg:xmlData>` +
    `<w:document xmlns:w="${W}"><w:body>` +
    `<w:p><w:pPr><w:pStyle w:val="Analytics"/></w:pPr><w:r><w:t xml:space="preserve">a navy analytics tag</w:t></w:r></w:p>` +
    `<w:p><w:r><w:t xml:space="preserve">ordinary body card text</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
    `</w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}

describe("A2 styles part is NEVER DOM-parsed (perf-1 read-side win)", () => {
  beforeEach(() => {
    parsedSources.length = 0;
  });

  it("ctor + headingLevel classify resolve the styled level WITHOUT parsing the styles part", () => {
    const wb = new WholeBodyPackage(styledPackage());
    // The read-side classify surface that consumes the styles: resolve every paragraph's level.
    const levels = [wb.headingLevel(0), wb.headingLevel(1)];

    // CORRECTNESS: the regex resolver read the styles substring — Analytics → Heading4 → level 3.
    expect(levels[0]).toBe(3); // navy analytics tag kept (resolved from the UNPARSED styles substring)
    expect(levels[1]).toBeNull(); // plain body

    // THE WIN: no DOM parse ever saw the styles definitions. The styles marker lives ONLY in the
    // styles part, so its absence from every parsed source proves the styles part was regex-scanned,
    // never DOM-parsed (the old extractStylesXml would have parsed the whole package — marker included).
    const sawStyles = parsedSources.filter((s) => s.includes(STYLE_MARKER));
    expect(sawStyles).toEqual([]);
    // And the document part WAS parsed (the one parse A2 keeps) — sanity that the spy is wired and the
    // styles-absence isn't a vacuous "nothing parsed at all".
    expect(parsedSources.some((s) => s.includes("a navy analytics tag"))).toBe(true);
  });

  it("serialize() (the node-direct commit) also never DOM-parses the styles part", () => {
    const wb = new WholeBodyPackage(styledPackage());
    parsedSources.length = 0; // measure ONLY the serialize() call
    const out = wb.serialize();
    // serialize() stitches verbatim substrings + the re-serialized document subtree — it parses
    // NOTHING (no DOMParser at all), and certainly not the styles part.
    expect(parsedSources).toEqual([]);
    // The styles part round-trips verbatim in the stitched output (byte-preserved, not reparsed).
    expect(out).toContain(`${STYLE_MARKER} Analytics`);
  });
});
