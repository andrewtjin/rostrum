import {
  normalizeOutlineNumber,
  countBodyParagraphs,
  assertSingleParagraph,
  keepFirstBodyParagraph,
  WholeBodyPackage
} from "../src/core/ooxmlPackage";
import { readRuns } from "../src/core/ooxml";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

/** A flat-OPC package wrapping a document body built from the given `<w:p>` strings. */
function pkg(bodyParas: string[], opts: { header?: string; sectPr?: boolean } = {}): string {
  const headerPart = opts.header
    ? `<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData>` +
      `<w:hdr xmlns:w="${W_NS}">${opts.header}</w:hdr>` +
      `</pkg:xmlData></pkg:part>`
    : "";
  const sect = opts.sectPr === false ? "" : `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>`;
  return (
    `<pkg:package xmlns:pkg="${PKG_NS}">` +
    headerPart +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${bodyParas.join("")}${sect}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}

const p = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

// ---------------------------------------------------------------------------
describe("story paragraphs exclude textbox content (Stage 4.1 ①)", () => {
  // An outer story paragraph that contains a text box whose own <w:p> must NOT count
  // as a body-story paragraph (Word's body.paragraphs excludes it; getElementsByTagName
  // would otherwise over-count it — the source of the whole-body alignment mismatch).
  const textboxPara =
    `<w:p><w:r><w:t xml:space="preserve">outer two</w:t></w:r>` +
    `<w:r><w:txbxContent><w:p><w:r><w:t xml:space="preserve">INNER BOX</w:t></w:r></w:p></w:txbxContent></w:r></w:p>`;

  it("counts only story <w:p>, never textbox-nested ones", () => {
    const x = pkg([p("one"), textboxPara, p("three")]);
    expect(countBodyParagraphs(x)).toBe(3); // outer paras only; the textbox <w:p> is excluded
    const wbp = new WholeBodyPackage(x);
    expect(wbp.count).toBe(3);
  });

  it("paragraphText excludes textbox text (so alignment matches Word's paragraph.text)", () => {
    const wbp = new WholeBodyPackage(pkg([p("one"), textboxPara, p("three")]));
    expect(wbp.paragraphText(1)).toBe("outer two"); // NOT "outer twoINNER BOX"
    expect(wbp.paragraphText(0)).toBe("one");
  });

  it("paragraphText renders tabs/breaks as whitespace (so a cite tab doesn't break alignment)", () => {
    const tabPara =
      `<w:p><w:r><w:t xml:space="preserve">Smith</w:t></w:r>` +
      `<w:r><w:tab/></w:r><w:r><w:t xml:space="preserve">2020</w:t></w:r></w:p>`;
    const brPara = `<w:p><w:r><w:t xml:space="preserve">a</w:t><w:br/><w:t xml:space="preserve">b</w:t></w:r></w:p>`;
    const wbp = new WholeBodyPackage(pkg([tabPara, brPara]));
    expect(wbp.paragraphText(0)).toBe("Smith\t2020");
    expect(wbp.paragraphText(1)).toBe("a\nb");
  });
});

// ---------------------------------------------------------------------------
describe("keepFirstBodyParagraph (strip the trailing paragraph-mark artifact)", () => {
  it("drops a trailing empty <w:p> that getOoxml emits for the paragraph mark", () => {
    const twoPara = pkg([p("real card body"), "<w:p/>"]);
    expect(countBodyParagraphs(twoPara)).toBe(2);
    const one = keepFirstBodyParagraph(twoPara);
    expect(countBodyParagraphs(one)).toBe(1);
    expect(readRuns(one).map((r) => r.text)).toEqual(["real card body"]);
  });

  it("preserves <w:sectPr> while stripping the trailing paragraph", () => {
    const one = keepFirstBodyParagraph(pkg([p("body"), "<w:p/>"]));
    expect(one).toContain("w:sectPr");
    expect(countBodyParagraphs(one)).toBe(1);
  });

  it("returns an already-single-paragraph package byte-identical (no reserialize)", () => {
    const onePara = pkg([p("only")]);
    expect(keepFirstBodyParagraph(onePara)).toBe(onePara);
  });

  it("returns a bare <w:p> unchanged", () => {
    const bare = `<w:p xmlns:w="${W_NS}"><w:r><w:t>x</w:t></w:r></w:p>`;
    expect(keepFirstBodyParagraph(bare)).toBe(bare);
  });
});

// ---------------------------------------------------------------------------
describe("normalizeOutlineNumber (Paragraph.outlineLevel number)", () => {
  it("oneBased (VBA parity): 1..9 -> 0..8, body (>=10) -> null", () => {
    expect(normalizeOutlineNumber(1)).toBe(0);
    expect(normalizeOutlineNumber(4)).toBe(3); // the navy Analytics style survives Hide
    expect(normalizeOutlineNumber(9)).toBe(8);
    expect(normalizeOutlineNumber(10)).toBeNull(); // wdOutlineLevelBodyText
    expect(normalizeOutlineNumber(0)).toBeNull();
  });

  it("tolerates null / NaN / fractional input", () => {
    expect(normalizeOutlineNumber(null)).toBeNull();
    expect(normalizeOutlineNumber(undefined)).toBeNull();
    expect(normalizeOutlineNumber(NaN)).toBeNull();
    expect(normalizeOutlineNumber(3.0)).toBe(2);
  });

  it("zeroBased escape hatch: 0..8 -> 0..8, 9 -> null", () => {
    expect(normalizeOutlineNumber(0, "zeroBased")).toBe(0);
    expect(normalizeOutlineNumber(8, "zeroBased")).toBe(8);
    expect(normalizeOutlineNumber(9, "zeroBased")).toBeNull();
  });
});

// ---------------------------------------------------------------------------
describe("countBodyParagraphs", () => {
  it("counts only body-story paragraphs, never header/footer", () => {
    const xml = pkg([p("a"), p("b"), p("c")], { header: "<w:p><w:r><w:t>HEADER</w:t></w:r></w:p>" });
    expect(countBodyParagraphs(xml)).toBe(3);
  });

  it("counts a bare single-paragraph fragment as 1", () => {
    expect(countBodyParagraphs(`<w:p xmlns:w="${W_NS}"><w:r><w:t>x</w:t></w:r></w:p>`)).toBe(1);
  });

  it("counts table paragraphs (document order, like Word 1.3+)", () => {
    const tbl =
      `<w:tbl><w:tr><w:tc>${p("cell")}</w:tc></w:tr></w:tbl>`;
    const xml = pkg([p("before"), tbl, p("after")]);
    expect(countBodyParagraphs(xml)).toBe(3); // before + cell + after
  });
});

describe("assertSingleParagraph (multi-<w:p> guard, audit #5)", () => {
  it("passes for exactly one paragraph", () => {
    expect(() => assertSingleParagraph(`<w:p xmlns:w="${W_NS}"><w:r><w:t>x</w:t></w:r></w:p>`)).not.toThrow();
  });

  it("throws for zero paragraphs", () => {
    expect(() => assertSingleParagraph(`<w:sectPr xmlns:w="${W_NS}"/>`)).toThrow(/found 0/);
  });

  it("throws (with context) for more than one paragraph", () => {
    const two = `<w:body xmlns:w="${W_NS}">${p("a")}${p("b")}</w:body>`;
    expect(() => assertSingleParagraph(two, "test-ctx")).toThrow(/test-ctx.*found 2/s);
  });
});

// ---------------------------------------------------------------------------
describe("WholeBodyPackage", () => {
  it("counts body paragraphs and excludes header parts", () => {
    const wb = new WholeBodyPackage(pkg([p("one"), p("two")], { header: "<w:p><w:r><w:t>H</w:t></w:r></w:p>" }));
    expect(wb.count).toBe(2);
  });

  it("hands out a standalone, parseable per-paragraph fragment (ambient xmlns re-declared)", () => {
    const wb = new WholeBodyPackage(pkg([p("alpha"), p("beta")]));
    const frag = wb.paragraphXml(1);
    expect(frag).toContain(`xmlns:w="${W_NS}"`); // standalone-parseable
    // The engine can read it exactly like a Stage-1 bare fragment.
    expect(readRuns(frag).map((r) => r.text)).toEqual(["beta"]);
    expect(countBodyParagraphs(frag)).toBe(1);
  });

  it("re-declares NON-w namespaces on each fragment so the live host's insertOoxml accepts it (Stage 4.2 regression)", () => {
    // The wet-test bug: a fast-path fragment carrying a hyperlink (`r:`) or emoji
    // (`mc:`/`w16se:`) was serialized bare, dropping those ancestor xmlns decls. xmldom (and
    // so the engine) parses that leniently, but the LIVE host's insertOoxml rejects an
    // undeclared prefix ("we found a problem with its contents"). The fragment MUST carry
    // every prefix its content uses.
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const MC_NS = "http://schemas.openxmlformats.org/markup-compatibility/2006";
    const W16SE_NS = "http://schemas.microsoft.com/office/word/2015/wordml/symex";
    const docPkg =
      `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}" xmlns:mc="${MC_NS}" xmlns:w16se="${W16SE_NS}"><w:body>` +
      `<w:p><w:hyperlink r:id="rId7"><w:r><w:t>link</w:t></w:r></w:hyperlink>` +
      `<mc:AlternateContent><mc:Choice Requires="w16se"><w:r><w16se:symEx w16se:font="Segoe UI Emoji" w16se:char="1F600"/></w:r></mc:Choice></mc:AlternateContent>` +
      `</w:p><w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const frag = new WholeBodyPackage(docPkg).paragraphXml(0);
    expect(frag).toContain(`xmlns:r="${R_NS}"`); // hyperlink prefix declared
    expect(frag).toContain(`xmlns:mc="${MC_NS}"`); // markup-compat prefix declared
    expect(frag).toContain(`xmlns:w16se="${W16SE_NS}"`); // emoji prefix declared
    expect(frag).toContain("r:id="); // ...and the content that needs them survives
    expect(frag).toContain("w16se:symEx");
    expect(countBodyParagraphs(frag)).toBe(1); // still exactly one body paragraph for the engine
  });

  it("carries the relationships a paragraph references so r:id isn't dangling (Stage 4.2 audit C1)", () => {
    // The live insertOoxml rejection ("problem with its contents") is most likely a DANGLING
    // relationship: a hyperlink's r:id with no defining .rels part in the fragment. The package
    // from body.getOoxml() DOES include /word/_rels/document.xml.rels; paragraphXml must carry the
    // specific relationship(s) a paragraph references into its fragment.
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    const docPkg =
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>` +
      `<Relationships xmlns="${RELS_NS}">` +
      `<Relationship Id="rId7" Type="${R_NS}/hyperlink" Target="https://example.com" TargetMode="External"/>` +
      `<Relationship Id="rId9" Type="${R_NS}/hyperlink" Target="https://unused.com" TargetMode="External"/>` +
      `</Relationships></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
      `<w:p><w:hyperlink r:id="rId7"><w:r><w:t>cite</w:t></w:r></w:hyperlink></w:p>` +
      `<w:p><w:r><w:t>plain body</w:t></w:r></w:p>` +
      `<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const wb = new WholeBodyPackage(docPkg);

    const hyper = wb.paragraphXml(0);
    expect(hyper).toContain("/word/_rels/document.xml.rels"); // a doc-level rels part is present
    expect(hyper).toContain(`Id="rId7"`); // the referenced relationship is carried
    expect(hyper).not.toContain(`Id="rId9"`); // ...but ONLY the referenced one (stays minimal)
    expect(countBodyParagraphs(hyper)).toBe(1);

    const plain = wb.paragraphXml(1);
    // A plain paragraph needs NO document-level rels part (no r:ref to resolve)...
    expect(plain).not.toContain("document.xml.rels");
    // ...but it STILL carries the mandatory package start part (officeDocument → document.xml).
    // The only relationship present is that start-part rel — never a document.xml.rels one.
    expect(plain).toContain("/_rels/.rels");
    expect(plain).toContain("officeDocument");
    expect(plain).not.toContain(`Id="rId7"`);
  });

  it("ALWAYS carries the /_rels/.rels start part so the host's insertOoxml can find the main document (Stage 4.2 commit fix)", () => {
    // The commit-regression root cause: the minimal per-paragraph package omitted the OPC
    // package-relationships part, so the live host's insertOoxml couldn't identify which part is
    // the main document and threw GeneralException — on EVERY OOXML paragraph, plain ones too
    // (the diagnostic sample was `hasRel:false`). Microsoft's OOXML coercion guidance: the
    // /_rels/.rels officeDocument relationship is "always required". Assert it on BOTH a plain
    // paragraph and one carrying a hyperlink, with the exact Id/Type/Target Word itself emits.
    const R_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    const docPkg =
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>` +
      `<Relationships xmlns="${RELS_NS}">` +
      `<Relationship Id="rId7" Type="${R_NS}/hyperlink" Target="https://example.com" TargetMode="External"/>` +
      `</Relationships></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}" xmlns:r="${R_NS}"><w:body>` +
      `<w:p><w:r><w:t>plain body</w:t></w:r></w:p>` +
      `<w:p><w:hyperlink r:id="rId7"><w:r><w:t>cite</w:t></w:r></w:hyperlink></w:p>` +
      `<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const wb = new WholeBodyPackage(docPkg);

    // The exact start-part relationship Word emits, verbatim — the OPC constant that names
    // /word/document.xml as the main document. Present on every fragment we hand to the host.
    const officeDocRel =
      `<Relationship Id="rId1" ` +
      `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ` +
      `Target="word/document.xml"/>`;
    for (const frag of [wb.paragraphXml(0), wb.paragraphXml(1)]) {
      expect(frag).toContain(`pkg:name="/_rels/.rels"`); // the start-part exists
      expect(frag).toContain(officeDocRel); // ...and names the main document part exactly
      expect(countBodyParagraphs(frag)).toBe(1); // still one body paragraph for the engine
    }

    // The package-level rId1 (officeDocument) and the document-level rId7 (hyperlink) live in
    // SEPARATE .rels parts, so they don't collide: the hyperlink fragment carries both.
    const hyper = wb.paragraphXml(1);
    expect(hyper).toContain(`pkg:name="/word/_rels/document.xml.rels"`);
    expect(hyper).toContain(`Id="rId7"`);
  });

  // -------------------------------------------------------------------------
  // commitXml — the per-paragraph COMMIT fragment bundles style/numbering/theme
  // parts (the read fragment from paragraphXml deliberately does not), so the host
  // renders style-inherited formatting (underline, character box, font size)
  // faithfully instead of collapsing it to document defaults. (Wet-test fix.)
  // -------------------------------------------------------------------------
  const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
  const PKG_RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
  const CT = "application/vnd.openxmlformats-officedocument";
  /** A whole-body package that carries styles + numbering + theme parts and a hyperlink, like Word's. */
  function styledPkg(): string {
    return (
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>` +
      `<Relationships xmlns="${PKG_RELS_NS}">` +
      `<Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/>` +
      `<Relationship Id="rIdN" Type="${REL}/numbering" Target="numbering.xml"/>` +
      `<Relationship Id="rIdT" Type="${REL}/theme" Target="theme/theme1.xml"/>` +
      `<Relationship Id="rId7" Type="${REL}/hyperlink" Target="https://example.com" TargetMode="External"/>` +
      `</Relationships></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${CT}.wordprocessingml.styles+xml"><pkg:xmlData>` +
      `<w:styles xmlns:w="${W_NS}"><w:style w:styleId="Style13ptBold"><w:rPr><w:u w:val="single"/><w:sz w:val="36"/></w:rPr></w:style></w:styles>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="${CT}.wordprocessingml.numbering+xml"><pkg:xmlData>` +
      `<w:numbering xmlns:w="${W_NS}"><w:abstractNum w:abstractNumId="0"/></w:numbering>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/theme/theme1.xml" pkg:contentType="${CT}.theme+xml"><pkg:xmlData>` +
      `<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office"/>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}" xmlns:r="${REL}"><w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="Style13ptBold"/></w:pPr><w:r><w:t>warrant</w:t></w:r></w:p>` +
      `<w:p><w:hyperlink r:id="rId7"><w:r><w:t>cite</w:t></w:r></w:hyperlink></w:p>` +
      `<w:sectPr/></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`
    );
  }

  it("commitXml bundles styles/numbering/theme parts AND their rels so style-inherited formatting survives", () => {
    const wb = new WholeBodyPackage(styledPkg());
    // commitXml takes the ENGINE's edited fragment (here, the read fragment re-fed) and re-wraps it.
    const frag = wb.commitXml(wb.paragraphXml(0));
    // The aux PARTS are present, with their content (so style defs actually reach the host)...
    expect(frag).toContain(`pkg:name="/word/styles.xml"`);
    expect(frag).toContain(`pkg:name="/word/numbering.xml"`);
    expect(frag).toContain(`pkg:name="/word/theme/theme1.xml"`);
    expect(frag).toContain(`w:styleId="Style13ptBold"`); // the underline+18pt style definition
    expect(frag).toContain("<w:u "); // the underline that was being lost
    // ...and the typed RELS that let the host LOCATE each part.
    expect(frag).toContain(`${REL}/styles`);
    expect(frag).toContain(`${REL}/numbering`);
    expect(frag).toContain(`${REL}/theme`);
    // Still a valid single-paragraph package with the mandatory start part.
    expect(frag).toContain(`pkg:name="/_rels/.rels"`);
    expect(countBodyParagraphs(frag)).toBe(1);
  });

  it("commitXml carries a paragraph's hyperlink relationship ALONGSIDE the style parts", () => {
    const wb = new WholeBodyPackage(styledPkg());
    const frag = wb.commitXml(wb.paragraphXml(1)); // paragraph 1 has the hyperlink
    expect(frag).toContain(`Id="rId7"`); // hyperlink rel carried (not dangling)...
    expect(frag).toContain(`pkg:name="/word/styles.xml"`); // ...and styles still bundled
    expect(countBodyParagraphs(frag)).toBe(1);
  });

  it("the READ fragment (paragraphXml) stays style-LESS so classify stays fast — only commitXml bundles styles", () => {
    const wb = new WholeBodyPackage(styledPkg());
    const read = wb.paragraphXml(0);
    expect(read).not.toContain("/word/styles.xml"); // read path is minimal (the ≈100× classify win)
    expect(read).not.toContain("/word/numbering.xml");
    expect(wb.commitXml(read)).toContain("/word/styles.xml"); // commit path bundles them
  });

  it("commitXml preserves the engine's <w:vanish/> edit while bundling the styles (round-trips the hide)", () => {
    const wb = new WholeBodyPackage(styledPkg());
    const edited = `<w:p xmlns:w="${W_NS}"><w:r><w:rPr><w:vanish/></w:rPr><w:t>warrant</w:t></w:r></w:p>`;
    const frag = wb.commitXml(edited);
    expect(readRuns(frag)[0].hidden).toBe(true); // the hide survived the re-wrap
    expect(frag).toContain(`pkg:name="/word/styles.xml"`); // ...with styles bundled
  });

  it("commitXml on a style-less package emits a valid single-paragraph package (start part, no empty rels part)", () => {
    const wb = new WholeBodyPackage(pkg([p("plain")]));
    const frag = wb.commitXml(wb.paragraphXml(0));
    expect(frag).toContain(`pkg:name="/_rels/.rels"`); // start part still present
    expect(frag).not.toContain("/word/styles.xml"); // nothing to bundle
    expect(frag).not.toContain("document.xml.rels"); // no aux rels + no hyperlink → no doc rels part
    expect(countBodyParagraphs(frag)).toBe(1);
  });

  it("splices an edited paragraph back, leaving siblings + structure byte-stable", () => {
    const wb = new WholeBodyPackage(pkg([p("keep me"), p("change me"), p("also keep")]));
    const edited = `<w:p xmlns:w="${W_NS}"><w:r><w:rPr><w:vanish/></w:rPr><w:t>change me</w:t></w:r></w:p>`;
    wb.replace(1, edited);
    const out = wb.serialize();

    expect(out).toContain("<w:vanish/>"); // the edit landed
    expect(out).toContain("keep me");
    expect(out).toContain("also keep");
    expect(out).toContain("<w:sectPr>"); // section properties untouched
    // Re-parsing the whole package still sees exactly 3 body paragraphs in order.
    expect(countBodyParagraphs(out)).toBe(3);
    const wb2 = new WholeBodyPackage(out);
    expect(readRuns(wb2.paragraphXml(1))[0].hidden).toBe(true);
    expect(readRuns(wb2.paragraphXml(0))[0].hidden).toBe(false);
  });

  it("never touches a header paragraph during a body splice", () => {
    const wb = new WholeBodyPackage(
      pkg([p("body card")], { header: "<w:p><w:r><w:t>RUNNING HEADER</w:t></w:r></w:p>" })
    );
    wb.replace(0, `<w:p xmlns:w="${W_NS}"><w:r><w:rPr><w:vanish/></w:rPr><w:t>body card</w:t></w:r></w:p>`);
    const out = wb.serialize();
    expect(out).toContain("RUNNING HEADER"); // header survives verbatim
  });

  it("rejects a multi-paragraph splice fragment", () => {
    const wb = new WholeBodyPackage(pkg([p("x")]));
    const two = `<w:body xmlns:w="${W_NS}">${p("a")}${p("b")}</w:body>`;
    expect(() => wb.replace(0, two)).toThrow(/found 2/);
  });

  it("throws a RangeError for an out-of-range index", () => {
    const wb = new WholeBodyPackage(pkg([p("only")]));
    expect(() => wb.paragraphXml(5)).toThrow(RangeError);
    expect(() => wb.replace(5, p("x"))).toThrow(RangeError);
  });
});
