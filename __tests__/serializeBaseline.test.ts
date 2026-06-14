// THE P4-REV CONTROL REFERENCE — pins today's pre-segmentation serialization output
// (PLAN.md §8 loss-6 / tests-2; CASES.md 002-S2/002-F2 AMENDED).
//
// WHAT THIS IS THE REFERENCE FOR. The future P4-REV string-level part segmentation
// (PLAN §2 A2) must produce `serialize()` / `commitXml()` / `paragraphXml()` output that
// is BYTE-IDENTICAL to TODAY's `WholeBodyPackage`. The reference is the CURRENT
// pre-segmentation output ON THE SAME INPUT — NEVER Word's original .docx bytes, because
// @xmldom/xmldom re-serializes the whole body and never reproduces Word's exact bytes
// (loss-6: that is precisely why the inherited "byte-diff" gate was unachievable and the
// prior team parked P1). So this file captures the control NOW; when P4-REV lands, its
// implementer points its own byte-equality gate at THESE expected values (or re-runs this
// test against the segmented `serialize()` and demands zero diff). 002-F2 binds: ANY byte
// mismatch on ANY realDocs sample kills the path — this synthetic control is the
// always-on CI floor under that gate (002-S10), proven additionally on the user's real
// corpus where `discoverSamples()>0`.
//
// SNAPSHOT CONVENTION. This repo uses NO jest snapshots (`toMatchSnapshot` appears in
// zero test files; there is no `__snapshots__` dir). So this control is pinned with
// EXPLICIT expected strings — deterministic, reviewable in the diff, and immune to a
// stray `--ci`/`-u` snapshot-update masking a regression. `serialize()` is pinned for
// EVERY fixture (it equals the fixture input string, because each fixture is authored in
// xmldom's canonical round-trip form — see fixtures/engine/index.ts); the freshly-wrapped
// `paragraphXml`/`commitXml` fragments are pinned for representative paragraphs spanning
// the interesting shapes (plain body, hyperlink+rels, styled-numbering aux parts).

import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { ALL_FIXTURES, bridgeGap, hyperlinkDense, styledNumbering, multiHeading } from "./fixtures/engine";

// ===========================================================================
// 1. serialize() — the whole-body control. For every fixture, today's
//    serialize() output IS the fixture string (canonical round-trip form). This
//    is the primary P4-REV reference: a segmented serialize() must equal this.
// ===========================================================================

describe("serialize() control reference (P4-REV byte basis, tests-2)", () => {
  for (const [name, pkg] of Object.entries(ALL_FIXTURES)) {
    it(`${name}: serialize() is byte-identical to the canonical input`, () => {
      const wb = new WholeBodyPackage(pkg);
      // The control: today's pre-segmentation serialize(). Pinned to the canonical input
      // string so a future serialize() change (segmentation, a stray namespace hoist) is a
      // hard, byte-level failure here — exactly the 002-F2 "any mismatch kills" gate.
      expect(wb.serialize()).toBe(pkg);
    });
  }

  it("serialize() is idempotent across repeated calls (no hidden mutation)", () => {
    // The pure read must never mutate its own tree (a serialize() that changed on a second
    // call would mean P4-REV's reference is unstable). Pin determinism explicitly.
    const wb = new WholeBodyPackage(multiHeading);
    const first = wb.serialize();
    expect(wb.serialize()).toBe(first);
    expect(first).toBe(multiHeading);
  });
});

// ===========================================================================
// 2. paragraphXml(i) — the READ fragment control (deliberately style-LESS, per
//    ooxmlPackage.ts). Pinned explicitly for representative paragraphs.
// ===========================================================================

// Shared constants matching the host's emission, so the expected strings below read as
// structure rather than opaque blobs (DRY with the fixtures module's own constants).
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const RELS = "http://schemas.openxmlformats.org/package/2006/relationships";
const RELS_CT = "application/vnd.openxmlformats-package.relationships+xml";
const DOC_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const OFFICE_DOC_REL = `${R}/officeDocument`;

/** The mandatory `/_rels/.rels` start part Word stamps on every flat-OPC package (verbatim). */
const START_PART =
  `<pkg:part pkg:name="/_rels/.rels" pkg:contentType="${RELS_CT}" pkg:padding="512"><pkg:xmlData>` +
  `<Relationships xmlns="${RELS}">` +
  `<Relationship Id="rId1" Type="${OFFICE_DOC_REL}" Target="word/document.xml"/>` +
  `</Relationships></pkg:xmlData></pkg:part>`;

describe("paragraphXml(i) control reference (read fragment, tests-2)", () => {
  it("bridgeGap.paragraphXml(0): minimal package, host docAttrs, no styles/rels parts", () => {
    const wb = new WholeBodyPackage(bridgeGap);
    // Pinned byte reference for the read fragment of a plain body paragraph. Note the
    // re-declared `xmlns:w` on the serialized `<w:p>` (xmldom emits it because the attached
    // node's ancestor declared it) — that detail IS part of the control P4-REV must match.
    const expected =
      `<pkg:package xmlns:pkg="${PKG}">` +
      START_PART +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CT}"><pkg:xmlData>` +
      `<w:document xmlns:w="${W}"><w:body>` +
      `<w:p xmlns:w="${W}">` +
      `<w:r><w:t xml:space="preserve">alpha</w:t></w:r>` +
      `<w:r><w:t xml:space="preserve"> beta gamma delta </w:t></w:r>` +
      `<w:r><w:t xml:space="preserve">omega</w:t></w:r>` +
      `</w:p>` +
      `</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    expect(wb.paragraphXml(0)).toBe(expected);
  });

  it("hyperlinkDense.paragraphXml(0): carries ONLY the referenced hyperlink rels, no styles", () => {
    const wb = new WholeBodyPackage(hyperlinkDense);
    // The read fragment of a hyperlink-bearing paragraph must include a document.xml.rels
    // part with the three referenced hyperlink relationships (so r:id isn't dangling) but NO
    // styles/numbering/theme parts (the read path stays minimal/fast). xmldom re-declares the
    // default xmlns on each copied <Relationship> and xmlns:r on each <w:hyperlink> — both are
    // part of the control.
    const rel = (id: string, target: string): string =>
      `<Relationship Id="${id}" Type="${R}/hyperlink" Target="${target}" TargetMode="External" xmlns="${RELS}"/>`;
    const hyperlink = (id: string, text: string): string =>
      `<w:hyperlink xmlns:r="${R}" r:id="${id}"><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:hyperlink>`;
    const expected =
      `<pkg:package xmlns:pkg="${PKG}">` +
      START_PART +
      `<pkg:part pkg:name="/word/_rels/document.xml.rels" pkg:contentType="${RELS_CT}"><pkg:xmlData>` +
      `<Relationships xmlns="${RELS}">` +
      rel("rId1", "https://example.org/one") +
      rel("rId2", "https://example.org/two") +
      rel("rId3", "https://example.org/three") +
      `</Relationships></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CT}"><pkg:xmlData>` +
      `<w:document xmlns:w="${W}" xmlns:r="${R}"><w:body>` +
      `<w:p xmlns:w="${W}">` +
      hyperlink("rId1", "first source") +
      `<w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
      hyperlink("rId2", "second source") +
      `<w:r><w:t xml:space="preserve"> and </w:t></w:r>` +
      hyperlink("rId3", "third source") +
      `</w:p>` +
      `</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    expect(wb.paragraphXml(0)).toBe(expected);
  });

  it("paragraphXml(i) is deterministic and never mutates the package (control stability)", () => {
    // P4-REV's reference is only meaningful if the current output is stable. Pin that calling
    // paragraphXml twice yields identical bytes AND leaves serialize() unchanged.
    const wb = new WholeBodyPackage(multiHeading);
    const whole = wb.serialize();
    const first = [wb.paragraphXml(0), wb.paragraphXml(4)];
    expect([wb.paragraphXml(0), wb.paragraphXml(4)]).toEqual(first);
    expect(wb.serialize()).toBe(whole);
  });
});

// ===========================================================================
// 3. commitXml(...) — the COMMIT fragment control (bundles styles/numbering/
//    theme aux parts; the read fragment does not). Pinned for the styled-
//    numbering paragraph, the shape where the aux bundle actually appears.
// ===========================================================================

describe("commitXml(...) control reference (commit fragment, tests-2)", () => {
  it("styledNumbering.commitXml: bundles the styles + numbering aux parts verbatim", () => {
    const wb = new WholeBodyPackage(styledNumbering);
    const read = wb.paragraphXml(0); // the engine's read fragment, re-fed to commitXml
    // The commit fragment for a styled/numbered paragraph must prepend the styles.xml and
    // numbering.xml aux parts (each re-declaring xmlns:pkg as xmldom serializes the attached
    // <pkg:part> node) BEFORE the document part, so style/numbering-inherited formatting
    // survives insertOoxml. No document.xml.rels part: this paragraph references no r:id and
    // the aux parts are located by type, not relationship id, in this fixture.
    const stylesPart =
      `<pkg:part xmlns:pkg="${PKG}" pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
      `<w:styles xmlns:w="${W}">` +
      `<w:style w:type="paragraph" w:styleId="ListParagraph"><w:name w:val="List Paragraph"/><w:basedOn w:val="Normal"/></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Normal" w:default="1"><w:name w:val="Normal"/></w:style>` +
      `</w:styles></pkg:xmlData></pkg:part>`;
    const numberingPart =
      `<pkg:part xmlns:pkg="${PKG}" pkg:name="/word/numbering.xml" pkg:contentType="${NUMBERING_CT}"><pkg:xmlData>` +
      `<w:numbering xmlns:w="${W}">` +
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pStyle w:val="ListParagraph"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>` +
      `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>` +
      `</w:numbering></pkg:xmlData></pkg:part>`;
    const expected =
      `<pkg:package xmlns:pkg="${PKG}">` +
      START_PART +
      stylesPart +
      numberingPart +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CT}"><pkg:xmlData>` +
      `<w:document xmlns:w="${W}"><w:body>` +
      `<w:p xmlns:w="${W}">` +
      `<w:pPr><w:pStyle w:val="ListParagraph"/><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr>` +
      `<w:r><w:t xml:space="preserve">a numbered list item body</w:t></w:r>` +
      `</w:p>` +
      `</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    expect(wb.commitXml(read)).toBe(expected);
  });

  it("bridgeGap.commitXml: style-less package equals the read fragment (no aux parts to bundle)", () => {
    // When the source package carries NO styles/numbering/theme parts, commitXml emits the
    // SAME minimal package as paragraphXml (only the start part + document part). Pinning this
    // equality documents that the aux bundle is conditional — the control for a style-less doc.
    const wb = new WholeBodyPackage(bridgeGap);
    expect(wb.commitXml(wb.paragraphXml(0))).toBe(wb.paragraphXml(0));
  });
});
