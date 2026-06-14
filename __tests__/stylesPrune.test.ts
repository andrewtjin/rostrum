// P7 — referenced-closure styles prune (Loop 002 B3; CASES 002-S6 / 002-F6).
//
// PROVES, all in CI on synthetic license-clean fixtures:
//   * 002-F6 (flag OFF): serialize()/commitXml() byte-identical to today; the styles part is never
//     DOM-parsed and never altered (the SHIPPED default).
//   * CLOSED-WORLD transitive-fixpoint closure-correctness (flag ON, loss-5): after prune, every
//     styleId reachable by ANY retention rule from ANY retained content part is PRESENT, AND no
//     retained style's basedOn/link/next/numStyleLink/styleLink points at a REMOVED id.
//   * Numbering-backlink seeding (the verified ndca gap: a style referenced ONLY by numbering's
//     w:lvl is kept).
//   * Prune idempotence + re-seed on a reused package.
//   * The document body + docDefaults/latentStyles/root are untouched by the prune.
//   * Write-side byte reduction (perf-1 evidence) on a styles part padded with unused styles.
//
// All style XML here is invented (no real debate content). The closure rules under test are the
// OOXML style cross-references basedOn / link (both directions) / next / numStyleLink / styleLink.

import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { parseStyleDefs } from "../src/core/outline";
import { createOfficeWordPort } from "../src/core/officeWordPort";
import { hide } from "../src/core/invisibility";
import { mkDoc, para, run, harness, settings } from "./fakeWord";
import { ALL_FIXTURES } from "./fixtures/engine";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const DOC_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/** A single `<w:style>` element with the given id and inner cross-reference children. */
function style(id: string, opts: { type?: string; default?: boolean; inner?: string } = {}): string {
  const type = opts.type ?? "paragraph";
  const def = opts.default ? ` w:default="1"` : "";
  return `<w:style w:type="${type}" w:styleId="${id}"${def}><w:name w:val="${id}"/>${opts.inner ?? ""}</w:style>`;
}

/** A flat-OPC package with a document body + a styles part (+ optional numbering part). */
function pkgWith(opts: {
  body: string;
  styles: string;
  numbering?: string;
  docNs?: string;
}): string {
  const numberingPart = opts.numbering
    ? `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="${NUMBERING_CT}"><pkg:xmlData>` +
      `<w:numbering xmlns:w="${W}">${opts.numbering}</w:numbering>` +
      `</pkg:xmlData></pkg:part>`
    : "";
  return (
    `<pkg:package xmlns:pkg="${PKG}">` +
    `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    `<w:latentStyles w:count="0"/>` +
    opts.styles +
    `</w:styles></pkg:xmlData></pkg:part>` +
    numberingPart +
    `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CT}"><pkg:xmlData>` +
    `<w:document xmlns:w="${W}"${opts.docNs ?? ""}><w:body>${opts.body}</w:body></w:document>` +
    `</pkg:xmlData></pkg:part>` +
    `</pkg:package>`
  );
}

// ---------------------------------------------------------------------------
// A RICH prunable fixture exercising every closure rule:
//   * Normal (w:default) — kept by the default-style seed even if never cited.
//   * Heading1 — cited by document.xml pStyle; basedOn Normal; next BodyText.
//   * BodyText — reached ONLY via Heading1's `next`.
//   * CiteChar — cited by document.xml rStyle; link CitePara (both directions).
//   * CitePara — reached ONLY via CiteChar's reverse `link`.
//   * ListNum / ListNumChar — numStyleLink / styleLink pair, ListNum cited by numbering's w:lvl pStyle.
//   * Orphan1 / Orphan2 — UNREFERENCED, must be PRUNED. Orphan2 basedOn Orphan1 (a removable chain).
//   * GhostBase — basedOn'd ONLY by Orphan1 (must also be pruned — not reachable from content).
// ---------------------------------------------------------------------------
const RICH_STYLES =
  style("Normal", { default: true }) +
  style("Heading1", { inner: `<w:basedOn w:val="Normal"/><w:next w:val="BodyText"/>` }) +
  style("BodyText", { inner: `<w:basedOn w:val="Normal"/>` }) +
  style("CiteChar", { type: "character", inner: `<w:link w:val="CitePara"/>` }) +
  style("CitePara", { inner: `<w:basedOn w:val="Normal"/>` }) +
  style("ListNum", { inner: `<w:numStyleLink w:val="ListNumChar"/>` }) +
  style("ListNumChar", { inner: `<w:styleLink w:val="ListNum"/>` }) +
  style("Orphan2", { inner: `<w:basedOn w:val="Orphan1"/>` }) +
  style("Orphan1", { inner: `<w:basedOn w:val="GhostBase"/>` }) +
  style("GhostBase", {});

/** A body that cites Heading1 (pStyle), CiteChar (rStyle), and a numbered ListNum paragraph. */
const RICH_BODY =
  `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">A heading</w:t></w:r></w:p>` +
  `<w:p><w:r><w:rPr><w:rStyle w:val="CiteChar"/></w:rPr><w:t xml:space="preserve">a cite</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="ListNum"/><w:numPr><w:numId w:val="1"/></w:numPr></w:pPr>` +
  `<w:r><w:t xml:space="preserve">item</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t xml:space="preserve">plain body</w:t></w:r></w:p>`;

/** Numbering whose w:lvl backlinks ListNum (the ndca gap: a style referenced ONLY through numbering). */
const RICH_NUMBERING =
  `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pStyle w:val="ListNum"/><w:numFmt w:val="decimal"/></w:lvl></w:abstractNum>` +
  `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>`;

const richPackage = pkgWith({ body: RICH_BODY, styles: RICH_STYLES, numbering: RICH_NUMBERING });

/** Every styleId present in a styles XML (the retained set after a prune). */
function styleIds(stylesXml: string): Set<string> {
  return new Set(parseStyleDefs(stylesXml).keys());
}

/** The five cross-reference target ids declared by a retained style (for the no-dangling assertion). */
function refTargets(inner: string): string[] {
  const out: string[] = [];
  for (const tag of ["basedOn", "link", "next", "numStyleLink", "styleLink"]) {
    const m = new RegExp(`<w:${tag}\\b[^>]*\\bw:val="([^"]+)"`).exec(inner);
    if (m) out.push(m[1]);
  }
  return out;
}

// ===========================================================================
// 1. FLAG OFF (the shipped default) — 002-F6 byte-identity.
// ===========================================================================
describe("flag OFF: styles part is verbatim (002-F6)", () => {
  it("default ctor: pruneStylesToClosure() is a no-op; serialize() unchanged", () => {
    const wb = new WholeBodyPackage(richPackage); // no options → flag OFF
    const before = wb.serialize();
    expect(wb.willPruneStyles).toBe(false);
    expect(wb.pruneStylesToClosure()).toBe(0); // returns 0 bytes removed
    expect(wb.serialize()).toBe(before);
    expect(wb.serialize()).toBe(richPackage); // == the canonical input
  });

  it("explicit pruneStyles:false behaves identically to the default", () => {
    const wb = new WholeBodyPackage(richPackage, { pruneStyles: false });
    expect(wb.willPruneStyles).toBe(false);
    wb.pruneStylesToClosure();
    expect(wb.serialize()).toBe(richPackage);
  });

  it("every engine fixture: a flag-OFF prune call leaves serialize() byte-identical", () => {
    // The serializeBaseline control already pins serialize()==input; here we additionally prove that
    // CALLING the prune method with the flag OFF never perturbs any fixture (the dormant-path guarantee).
    for (const [name, fixture] of Object.entries(ALL_FIXTURES)) {
      const wb = new WholeBodyPackage(fixture);
      wb.pruneStylesToClosure();
      expect(`${name}: ${wb.serialize()}`).toBe(`${name}: ${fixture}`);
    }
  });
});

// ===========================================================================
// 2. FLAG ON — transitive-fixpoint closure-correctness (loss-5).
// ===========================================================================
describe("flag ON: closure-correctness (loss-5, 002-S6)", () => {
  it("removes unreferenced styles, keeps the full referenced closure", () => {
    const wb = new WholeBodyPackage(richPackage, { pruneStyles: true });
    expect(wb.willPruneStyles).toBe(true);
    const removed = wb.pruneStylesToClosure();
    expect(removed).toBeGreaterThan(0);
    const kept = styleIds(wb.stylesXmlForTest());

    // (a) EVERY style reachable by a retention rule is PRESENT.
    expect(kept).toContain("Normal"); // w:default seed
    expect(kept).toContain("Heading1"); // document.xml pStyle seed
    expect(kept).toContain("BodyText"); // via Heading1 `next`
    expect(kept).toContain("CiteChar"); // document.xml rStyle seed
    expect(kept).toContain("CitePara"); // via CiteChar reverse `link`
    expect(kept).toContain("ListNum"); // numbering w:lvl pStyle backlink seed
    expect(kept).toContain("ListNumChar"); // via ListNum `numStyleLink`

    // The unreferenced chain is GONE.
    expect(kept).not.toContain("Orphan1");
    expect(kept).not.toContain("Orphan2");
    expect(kept).not.toContain("GhostBase");
  });

  it("(b) no retained style's cross-references dangle at a removed id", () => {
    const wb = new WholeBodyPackage(richPackage, { pruneStyles: true });
    wb.pruneStylesToClosure();
    const stylesXml = wb.stylesXmlForTest();
    const kept = styleIds(stylesXml);
    // For each retained <w:style>, every basedOn/link/next/numStyleLink/styleLink target must be retained.
    const styleRe = /<w:style\b[^>]*\bw:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g;
    let m: RegExpExecArray | null;
    let checked = 0;
    while ((m = styleRe.exec(stylesXml)) !== null) {
      for (const target of refTargets(m[2])) {
        checked++;
        expect(kept).toContain(target); // CLOSED-WORLD: no dangling ref
      }
    }
    expect(checked).toBeGreaterThan(0); // the assertion actually ran on real edges
  });

  it("link is retained in BOTH directions (para reached → char kept; char reached → para kept)", () => {
    // CitePara is reached ONLY because CiteChar (cited by content) declares <w:link CitePara>. Prove the
    // reverse edge too: a body citing the PARA style must keep the linked CHAR style.
    const body =
      `<w:p><w:pPr><w:pStyle w:val="CitePara"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`;
    const styles =
      style("Normal", { default: true }) +
      style("CitePara", { inner: `<w:link w:val="CiteChar"/>` }) +
      style("CiteChar", { type: "character" }) +
      style("Unused", {});
    const wb = new WholeBodyPackage(pkgWith({ body, styles }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const kept = styleIds(wb.stylesXmlForTest());
    expect(kept).toEqual(new Set(["Normal", "CitePara", "CiteChar"]));
    expect(kept).not.toContain("Unused");
  });
});

// ===========================================================================
// 3. Numbering-backlink seeding (the verified ndca gap).
// ===========================================================================
describe("flag ON: numbering w:lvl backlinks seed the closure", () => {
  it("a style referenced ONLY by numbering's w:lvl (not by the body) is kept", () => {
    // ListNum is cited NOWHERE in document.xml; its only reference is numbering's w:lvl pStyle.
    const body = `<w:p><w:r><w:t xml:space="preserve">no style refs here</w:t></w:r></w:p>`;
    const styles =
      style("Normal", { default: true }) +
      style("ListNum", { inner: `<w:basedOn w:val="Normal"/>` }) +
      style("Orphan", {});
    const numbering =
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:pStyle w:val="ListNum"/></w:lvl></w:abstractNum>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles, numbering }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const kept = styleIds(wb.stylesXmlForTest());
    expect(kept).toContain("ListNum"); // kept via the numbering backlink — the gap is closed
    expect(kept).not.toContain("Orphan");
  });

  it("a w:rStyle inside a numbering w:lvl's w:rPr also seeds", () => {
    const body = `<w:p><w:r><w:t xml:space="preserve">body</w:t></w:r></w:p>`;
    const styles =
      style("Normal", { default: true }) +
      style("NumRunChar", { type: "character" }) +
      style("Orphan", {});
    const numbering =
      `<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:rPr><w:rStyle w:val="NumRunChar"/></w:rPr></w:lvl></w:abstractNum>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles, numbering }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const kept = styleIds(wb.stylesXmlForTest());
    expect(kept).toContain("NumRunChar");
    expect(kept).not.toContain("Orphan");
  });
});

// ===========================================================================
// 4. Idempotence + re-seed; docDefaults/latentStyles/body preserved.
// ===========================================================================
describe("flag ON: idempotence, re-seed, structural preservation", () => {
  it("pruning twice yields the same bytes (idempotent); the second call removes 0", () => {
    const wb = new WholeBodyPackage(richPackage, { pruneStyles: true });
    const removed1 = wb.pruneStylesToClosure();
    const after1 = wb.serialize();
    const removed2 = wb.pruneStylesToClosure();
    const after2 = wb.serialize();
    expect(removed1).toBeGreaterThan(0);
    expect(removed2).toBe(0); // already pruned → no further removal
    expect(after2).toBe(after1); // stable
  });

  it("docDefaults, latentStyles, and the <w:styles> root survive verbatim", () => {
    const wb = new WholeBodyPackage(richPackage, { pruneStyles: true });
    wb.pruneStylesToClosure();
    const stylesXml = wb.stylesXmlForTest();
    expect(stylesXml).toContain(`<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>`);
    expect(stylesXml).toContain(`<w:latentStyles w:count="0"/>`);
    expect(stylesXml.startsWith(`<w:styles xmlns:w="${W}">`)).toBe(true);
    expect(stylesXml.endsWith(`</w:styles>`)).toBe(true);
  });

  it("the document.xml body is byte-identical after a prune (only styles.xml shrinks)", () => {
    const wb = new WholeBodyPackage(richPackage, { pruneStyles: true });
    const docBefore = /<w:document[\s\S]*?<\/w:document>/.exec(wb.serialize())![0];
    wb.pruneStylesToClosure();
    const docAfter = /<w:document[\s\S]*?<\/w:document>/.exec(wb.serialize())![0];
    expect(docAfter).toBe(docBefore);
  });
});

// ===========================================================================
// 5. Write-side byte reduction (perf-1 evidence) + no-styles-part safety.
// ===========================================================================
describe("flag ON: byte reduction + edge cases", () => {
  it("a styles part padded with many unused styles shrinks materially", () => {
    // Pad the table with 200 unreferenced styles to model Word's bloated latent table; the body cites one.
    let pad = "";
    for (let i = 0; i < 200; i++) pad += style(`Unused${i}`, { inner: `<w:basedOn w:val="Normal"/>` });
    const styles = style("Normal", { default: true }) + style("Used", { inner: `<w:basedOn w:val="Normal"/>` }) + pad;
    const body = `<w:p><w:pPr><w:pStyle w:val="Used"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles }), { pruneStyles: true });
    const before = wb.stylesXmlForTest().length;
    const removed = wb.pruneStylesToClosure();
    const after = wb.stylesXmlForTest().length;
    expect(removed).toBe(before - after);
    expect(after).toBeLessThan(before * 0.2); // dropped the 200 unused → well under a fifth the size
    expect(styleIds(wb.stylesXmlForTest())).toEqual(new Set(["Normal", "Used"]));
    // eslint-disable-next-line no-console
    console.log(`[P7 prune] styles part ${before} -> ${after} chars (removed ${removed})`);
  });

  it("a package with no styles part: prune is a no-op even with the flag ON", () => {
    const wb = new WholeBodyPackage(ALL_FIXTURES.bridgeGap, { pruneStyles: true });
    expect(wb.willPruneStyles).toBe(false); // no <w:styles> part
    expect(wb.pruneStylesToClosure()).toBe(0);
    expect(wb.serialize()).toBe(ALL_FIXTURES.bridgeGap);
  });

  it("a raw document.xml (non-segmentable) with the flag ON is untouched", () => {
    const raw =
      `<w:document xmlns:w="${W}"><w:body><w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p></w:body></w:document>`;
    const wb = new WholeBodyPackage(raw, { pruneStyles: true });
    expect(wb.willPruneStyles).toBe(false);
    expect(wb.pruneStylesToClosure()).toBe(0);
    expect(wb.serialize()).toBe(raw);
  });
});

// ===========================================================================
// 6. END-TO-END through the port: flag-ON node-direct commit shrinks the
//    committed package's styles; flag-OFF commit emits it verbatim.
// ===========================================================================
describe("port node-direct commit honors the prune flag", () => {
  const KEEP = settings(["cyan", "yellow", "green", "lightGray", "magenta", "red"]);

  /** Drive a node-direct Hide, returning the package handed to insertOoxml. */
  async function commitWith(pruneStyles: boolean): Promise<string> {
    const count = new WholeBodyPackage(richPackage).count;
    const doc = mkDoc(Array.from({ length: count }, () => para(run("placeholder body"))));
    const h = harness(doc, richPackage);
    let committed: string | null = null;
    const body: any = (h.ctx as any).document.body;
    const realInsert = body.insertOoxml.bind(body);
    body.insertOoxml = (xml: string, loc?: string): void => {
      committed = xml;
      realInsert(xml, loc);
    };
    const port = createOfficeWordPort({
      runner: h.runner,
      pureWholeBody: true,
      pruneStyles,
      logger: h.tracer.logger("adapter")
    });
    await hide(port, KEEP);
    return committed ?? richPackage;
  }

  it("flag ON: the committed package drops the orphan styles", async () => {
    const committed = await commitWith(true);
    expect(committed).toContain(`w:styleId="Heading1"`);
    expect(committed).not.toContain(`w:styleId="Orphan1"`);
    expect(committed).not.toContain(`w:styleId="Orphan2"`);
    expect(committed).not.toContain(`w:styleId="GhostBase"`);
  });

  it("flag OFF: the committed package keeps the full styles table (002-F6)", async () => {
    const committed = await commitWith(false);
    // The styles part is emitted verbatim — every orphan still present.
    expect(committed).toContain(`w:styleId="Orphan1"`);
    expect(committed).toContain(`w:styleId="Orphan2"`);
    expect(committed).toContain(`w:styleId="GhostBase"`);
  });
});
