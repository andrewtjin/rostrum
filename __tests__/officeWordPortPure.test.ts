// Adapter tests for the PURE whole-body path (avenue ⑦) — NO Word host (fake RequestContext).
//
// The pure path (`pureWholeBody: true`) reads the whole story in ONE body.getOoxml(), classifies
// EVERY package paragraph directly (outline from the package's own styles.xml, table membership
// from its structure) with NO proxies and NO text alignment, then commits with ONE
// body.insertOoxml("Replace"). These assert: a single whole-body commit (no per-paragraph ops),
// the correct hidden set, outline resolved from the package (inline AND the basedOn cascade) rather
// than a proxy, table suppression from package structure, the ±1 artifact included (never skipped),
// reversibility, and idempotent re-hide with stable paragraph count (no phantom growth, lesson #28).

import { createOfficeWordPort } from "../src/core/officeWordPort";
import { hide } from "../src/core/invisibility";
import { readRuns } from "../src/core/ooxml";
import { parseManifestOrNull } from "../src/core/manifest";
import { W_NS, PKG_NS, para, run, mkDoc, buildPackage, harness, hiddenFlags, settings } from "./fakeWord";

const lvl0 = `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`; // inline Heading 1 outline level

const purePort = (h: ReturnType<typeof harness>) =>
  createOfficeWordPort({ runner: h.runner, pureWholeBody: true, logger: h.tracer.logger("adapter") });

const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
/** A flat-OPC package with a custom styles part + the given body paragraphs (for outline-resolution tests). */
const pkgWith = (stylesInner: string, parasXml: string): string =>
  `<pkg:package xmlns:pkg="${PKG_NS}">` +
  `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>${stylesInner}</pkg:xmlData></pkg:part>` +
  `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>${parasXml}` +
  `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;

describe("pure whole-body read+commit (avenue ⑦)", () => {
  it("commits the whole body in ONE insertOoxml — no per-paragraph ops — with the correct hidden set", async () => {
    // The heading's PROXY outline is body (10); only its inline <w:outlineLvl> marks it a heading.
    // Keeping it proves the pure path resolves outline from the PACKAGE, never the proxy.
    const doc = mkDoc([
      para(run("Heading"), { outlineNumber: 10, pPr: lvl0 }),
      para(run("a long card body")),
      para(run("Author 2019", { cite: true })),
      para(run("intro ") + run("warrant", { highlight: "yellow" }))
    ]);
    const h = harness(doc);
    const res = await hide(purePort(h), settings(["yellow"]));

    // Exactly one whole-body Replace; zero per-paragraph writes/native toggles.
    expect(h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml")).toHaveLength(1);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.insertOoxml")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.font.hidden")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "range.insertOoxml")).toBe(false);

    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // heading kept (package outline, not proxy)
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // body card hidden
    expect(hiddenFlags(doc.paragraphs[2].xml)).toEqual([false]); // cite kept
    expect(hiddenFlags(doc.paragraphs[3].xml)).toEqual([true, false]); // partial: intro hidden, warrant kept
    expect(parseManifestOrNull(doc.manifest!.xml)!.active).toBe(true);
    expect(res.paragraphsChanged).toBe(2);
    // No alignment machinery runs in the pure path.
    expect(h.warnings.some((w) => /fallback|exhausted|targeted/.test(w))).toBe(false);
  });

  it("resolves a heading through the package's styles.xml basedOn cascade (Analytics → Heading4) — no proxy", async () => {
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    // A package whose styles.xml defines Analytics (no own level) basedOn Heading4 (outlineLvl 3).
    const styledPkg =
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>` +
      `<Relationships xmlns="${RELS_NS}"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/></Relationships>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
      `<w:styles xmlns:w="${W_NS}">` +
      `<w:style w:type="paragraph" w:styleId="Heading4"><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>` +
      `<w:style w:type="paragraph" w:styleId="Analytics"><w:basedOn w:val="Heading4"/></w:style>` +
      `</w:styles></pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p><w:pPr><w:pStyle w:val="Analytics"/></w:pPr>${run("Card tag")}</w:p>` +
      `<w:p>${run("plain card body")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
      `</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    // Backing proxies say BOTH are body (10) — ignored by the pure path; count must match for write-back.
    const doc = mkDoc([para(run("Card tag")), para(run("plain card body"))]);
    const h = harness(doc, styledPkg);
    await hide(purePort(h), settings([]));

    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // kept via Analytics → Heading4 → level 3
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // plain body hidden
    expect(h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml")).toHaveLength(1);
  });

  it("suppresses hiding inside a table — table membership read from the package structure, not a proxy", async () => {
    const tablePkg =
      `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}"><w:body>` +
      `<w:p>${run("before table")}</w:p>` +
      `<w:tbl><w:tr><w:tc><w:p>${run("inside cell")}</w:p></w:tc></w:tr></w:tbl>` +
      `<w:p>${run("after table")}</w:p>` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr>` +
      `</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`;
    const doc = mkDoc([para(run("before table")), para(run("inside cell")), para(run("after table"))]);
    const h = harness(doc, tablePkg);
    await hide(purePort(h), settings([]));

    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([true]); // body hidden
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([false]); // in-table paragraph kept (suppressed)
    expect(hiddenFlags(doc.paragraphs[2].xml)).toEqual([true]); // body hidden
  });

  it("includes a trailing ±1 serialization-artifact <w:p> — never skips it, never runs alignment", async () => {
    const doc = mkDoc([para(run("Heading"), { pPr: lvl0 }), para(run("card"))]);
    const withArtifact = buildPackage([...doc.paragraphs, para(run(""))]); // the extra <w:p> (377 vs 376)
    const h = harness(doc, withArtifact);
    const paras = await purePort(h).readParagraphs();

    expect(paras).toHaveLength(3); // the artifact is classified + round-tripped like any paragraph
    expect(h.warnings.some((w) => /fallback|exhausted|targeted/.test(w))).toBe(false); // no alignment at all
  });

  it("is reversible (text preserved) and idempotent on re-hide with a stable paragraph count", async () => {
    const doc = mkDoc([para(run("Heading"), { pPr: lvl0 }), para(run("secret card body"))]);
    // Freeze a clean package so getOoxml returns the same well-formed body on each read.
    const h = harness(doc, buildPackage(doc.paragraphs));
    const port = createOfficeWordPort({ runner: h.runner, pureWholeBody: true, logger: h.tracer.logger("adapter") });

    await hide(port, settings([]));
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // card hidden
    // Lossless: <w:vanish/> toggles visibility, never deletes — the text is still present.
    expect(readRuns(doc.paragraphs[1].xml).map((r) => r.text).join("")).toContain("secret card body");

    h.ctx.commitLog.length = 0;
    await hide(port, settings([])); // re-hide reads the same frozen package
    expect(doc.paragraphs).toHaveLength(2); // no phantom growth (lesson #28)
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]);
    expect(h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml")).toHaveLength(1);
  });

  it("keeps a built-in Heading whose styles.xml omits <w:outlineLvl> — package name fallback (review C-1)", async () => {
    // styles.xml HAS Heading1 but with no explicit level (the common built-in case). The PROXIES say
    // body (10), so if the pure path consulted the proxy the heading would be hidden; keeping it
    // proves the package name fallback fires end-to-end.
    const styles =
      `<w:styles xmlns:w="${W_NS}">` +
      `<w:style w:styleId="Heading1"><w:basedOn w:val="Normal"/></w:style>` +
      `<w:style w:styleId="Normal"></w:style></w:styles>`;
    const parasXml = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr>${run("Tag")}</w:p><w:p>${run("body card")}</w:p>`;
    const doc = mkDoc([para(run("Tag")), para(run("body card"))]);
    const h = harness(doc, pkgWith(styles, parasXml));
    await hide(purePort(h), settings([]));

    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // Heading1 kept via name fallback
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // body hidden
  });

  it("falls back to the proxy read when styles.xml is present but unparseable (review C-3)", async () => {
    // <w:styles> is present (so it isn't 'absent') but its inner styles use no w: prefix → 0 parsed
    // defs. Trusting that would mis-classify every styled heading as body and HIDE it; the pure path
    // must detect the suspect parse and fall back to the proven proxy read instead.
    const broken = `<w:styles xmlns:w="${W_NS}"><style styleId="Heading1"><outlineLvl val="0"/></style></w:styles>`;
    const parasXml = `<w:p>${run("Big Heading")}</w:p><w:p>${run("body card")}</w:p>`;
    const doc = mkDoc([para(run("Big Heading"), { outlineNumber: 1 }), para(run("body card"))]);
    const h = harness(doc, pkgWith(broken, parasXml));
    const paras = await purePort(h).readParagraphs();

    expect(paras[0].headingLevel).toBe(0); // resolved from the PROXY (fallback), not the broken package
    expect(h.warnings.some((w) => /unparseable|fallback/.test(w))).toBe(true);
  });
});
