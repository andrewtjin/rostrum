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

import { WholeBodyPackage, __styleMatcherAgreementForTest } from "../src/core/ooxmlPackage";
import { parseStyleDefs } from "../src/core/outline";
import { createOfficeWordPort } from "../src/core/officeWordPort";
import { hide } from "../src/core/invisibility";
import { mkDoc, para, run, harness, settings } from "./fakeWord";
import { ALL_FIXTURES } from "./fixtures/engine";
import { discoverSamples, readDocxParts } from "./realDocs";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG = "http://schemas.microsoft.com/office/2006/xmlPackage";
const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const NUMBERING_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const DOC_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

/**
 * A single `<w:style>` element with the given id. PAIRED by default (`<w:style …>…</w:style>`); pass
 * `selfClosing: true` for a childless `<w:style …/>` (the self-closing form the loss-2/reg-2 fixtures
 * exercise). `name` overrides the `<w:name w:val>` text (used by the `$`-bearing-name dollar-safe fixture);
 * a self-closing element has NO children, so `name`/`inner` are ignored there.
 */
function style(
  id: string,
  opts: { type?: string; default?: boolean; inner?: string; name?: string; selfClosing?: boolean } = {}
): string {
  const type = opts.type ?? "paragraph";
  const def = opts.default ? ` w:default="1"` : "";
  if (opts.selfClosing) return `<w:style w:type="${type}" w:styleId="${id}"${def}/>`;
  const nameVal = opts.name ?? id;
  return `<w:style w:type="${type}" w:styleId="${id}"${def}><w:name w:val="${nameVal}"/>${opts.inner ?? ""}</w:style>`;
}

/** A flat-OPC package with a document body + a styles part (+ optional numbering/settings parts). */
function pkgWith(opts: {
  body: string;
  styles: string;
  numbering?: string;
  /** Verbatim settings.xml inner markup (e.g. a `<w:clickAndTypeStyle w:val>` ref) — adds a settings part. */
  settings?: string;
  /** Override the default `<w:latentStyles w:count="0"/>` (e.g. to carry `<w:lsdException>` children). */
  latentStyles?: string;
  docNs?: string;
}): string {
  const numberingPart = opts.numbering
    ? `<pkg:part pkg:name="/word/numbering.xml" pkg:contentType="${NUMBERING_CT}"><pkg:xmlData>` +
      `<w:numbering xmlns:w="${W}">${opts.numbering}</w:numbering>` +
      `</pkg:xmlData></pkg:part>`
    : "";
  const settingsPart = opts.settings
    ? `<pkg:part pkg:name="/word/settings.xml"><pkg:xmlData>` +
      `<w:settings xmlns:w="${W}">${opts.settings}</w:settings>` +
      `</pkg:xmlData></pkg:part>`
    : "";
  const latentStyles = opts.latentStyles ?? `<w:latentStyles w:count="0"/>`;
  return (
    `<pkg:package xmlns:pkg="${PKG}">` +
    `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
    `<w:styles xmlns:w="${W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>` +
    latentStyles +
    opts.styles +
    `</w:styles></pkg:xmlData></pkg:part>` +
    numberingPart +
    settingsPart +
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

/** Every styleId present in a styles XML (the retained set after a prune). PAIRED styles only — the
 *  production `parseStyleDefs` (shared with outline classification) matches only `<w:style>…</w:style>`. */
function styleIds(stylesXml: string): Set<string> {
  return new Set(parseStyleDefs(stylesXml).keys());
}

/** Every styleId present, including SELF-CLOSING `<w:style …/>` — for fixtures that exercise that form
 *  (the prune retains/drops both; parseStyleDefs alone would miss the self-closing ones). */
function styleIdsAll(stylesXml: string): Set<string> {
  const out = new Set<string>();
  const re = /<w:style\b([^>]*?)\/\s*>|<w:style\b([^>]*)>[\s\S]*?<\/w:style>/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(stylesXml)) !== null) {
    const attrs = m[1] !== undefined ? m[1] : m[2];
    const id = /\bw:styleId="([^"]+)"/.exec(attrs);
    if (id) out.add(id[1]);
  }
  return out;
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

// ===========================================================================
// 7. MATCHER AGREEMENT (FIX-2 / 002-F11) — parseStyleGraph's indexed-id set MUST
//    equal pruneStylesXml's spanned-element set, including self-closing styles.
// ===========================================================================
describe("style-matcher agreement (parseStyleGraph ≡ pruneStylesXml id-sets)", () => {
  it("the graph id-set and the span id-set are identical (paired + self-closing)", () => {
    // Mix paired AND self-closing styles. The OLD paired-only regex would skip the self-closing
    // `<w:style/>` and could fuse it into the next paired element — so its id would be MISSING from the
    // span set, diverging from the graph set. With the unified iterator they must match exactly.
    const stylesXml =
      `<w:styles xmlns:w="${W}">` +
      style("Normal", { default: true }) + // paired
      style("SelfA", { selfClosing: true }) + // self-closing — the regression case
      style("Heading1", { inner: `<w:basedOn w:val="Normal"/>` }) + // paired, right after a self-closing
      style("SelfB", { type: "character", selfClosing: true }) +
      `</w:styles>`;
    const { graphIds, spanIds } = __styleMatcherAgreementForTest(stylesXml);
    expect(spanIds).toEqual(graphIds);
    expect(graphIds).toEqual(new Set(["Normal", "SelfA", "Heading1", "SelfB"]));
  });

  it("`w:default` is truthy for 1/true/on (ST_OnOff) and false otherwise (b3-4)", () => {
    // A bare `w:default="on"` must seed the default set; an explicit `="0"`/`="off"`/absent must not. The
    // styleId set is what flows into pkgWith — only the `<w:style>` children, no wrapper.
    const styles =
      `<w:style w:type="paragraph" w:styleId="DefOne" w:default="1"/>` +
      `<w:style w:type="paragraph" w:styleId="DefTrue" w:default="true"/>` +
      `<w:style w:type="paragraph" w:styleId="DefOn" w:default="on"/>` +
      `<w:style w:type="paragraph" w:styleId="NotDefZero" w:default="0"/>` +
      `<w:style w:type="paragraph" w:styleId="NotDefOff" w:default="off"/>` +
      `<w:style w:type="paragraph" w:styleId="NotDefAbsent"/>`;
    // A body citing NO style: only the w:default styles (1/true/on) should survive the prune.
    const body = `<w:p><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const kept = styleIdsAll(wb.stylesXmlForTest()); // all self-closing → use the self-closing-aware id helper
    expect(kept).toEqual(new Set(["DefOne", "DefTrue", "DefOn"]));
  });
});

// ===========================================================================
// 8. SERIALIZE()-BYTE FIXTURES (FIX-5 / tests-1) — assert the COMMITTED bytes, not
//    the intermediate xmlData, so loss-1 (dollar) + loss-2 (self-closing) are proven.
// ===========================================================================
describe("flag ON: serialize()/commitXml() byte-correctness ($-safe, self-closing-safe)", () => {
  it("a retained style whose w:name contains `$` survives serialize() VERBATIM (loss-1)", () => {
    // `Price $1` / `$&` / `$$` are exactly the tokens String.replace would mangle in a replacement STRING.
    // The body cites this style, so it is RETAINED — its bytes must round-trip through the splice unchanged.
    const dollarName = "Price $1 for $$ and $& tokens";
    const styles =
      style("Normal", { default: true }) +
      style("Dollar", { name: dollarName, inner: `<w:basedOn w:val="Normal"/>` }) +
      style("Orphan", {});
    const body = `<w:p><w:pPr><w:pStyle w:val="Dollar"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`;
    const pkg = pkgWith({ body, styles });
    const wb = new WholeBodyPackage(pkg, { pruneStyles: true });
    const removed = wb.pruneStylesToClosure();
    expect(removed).toBeGreaterThan(0); // Orphan was dropped
    const out = wb.serialize();
    // The dollar-bearing name is present BYTE-FOR-BYTE in the committed package (no `$1`/`$&`/`$$` mangling).
    expect(out).toContain(`<w:name w:val="${dollarName}"/>`);
    expect(out).toContain(`w:styleId="Dollar"`);
    expect(out).not.toContain(`w:styleId="Orphan"`);
    // And the styles part is well-formed (re-parse round-trips to the same retained ids).
    expect(styleIds(wb.stylesXmlForTest())).toEqual(new Set(["Normal", "Dollar"]));
  });

  it("a self-closing `<w:style/>` in the closure is KEPT and the next paired style is not fused away (loss-2)", () => {
    // SelfKeep is self-closing AND cited by the body; Heading1 (paired) immediately follows it. The old
    // paired-only regex would skip SelfKeep and start its match at Heading1, fusing the two — dropping the
    // in-closure SelfKeep. Prove both survive and the orphan self-closing is dropped.
    const styles =
      style("Normal", { default: true }) +
      style("SelfKeep", { selfClosing: true }) + // cited by body, self-closing — must be KEPT
      style("Heading1", { inner: `<w:basedOn w:val="Normal"/>` }) + // paired, right after the self-closing
      style("SelfOrphan", { selfClosing: true }); // uncited self-closing — must be DROPPED
    const body =
      `<w:p><w:pPr><w:pStyle w:val="SelfKeep"/></w:pPr><w:r><w:t xml:space="preserve">a</w:t></w:r></w:p>` +
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t xml:space="preserve">b</w:t></w:r></w:p>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const out = wb.serialize();
    const kept = styleIdsAll(wb.stylesXmlForTest()); // self-closing styles present → self-closing-aware ids
    expect(kept).toContain("SelfKeep"); // the self-closing in-closure style is retained
    expect(kept).toContain("Heading1"); // the following paired style is intact (not fused away)
    expect(kept).toContain("Normal");
    expect(kept).not.toContain("SelfOrphan");
    // Byte-level: the self-closing SelfKeep element is present verbatim in the committed package.
    expect(out).toContain(`<w:style w:type="paragraph" w:styleId="SelfKeep"/>`);
    expect(out).not.toContain(`w:styleId="SelfOrphan"`);
  });

  it("commitXml() is dollar-safe AND carries the FULL unpruned styles (loss-5 compat path)", () => {
    // The per-paragraph compat commit deliberately stays UNPRUNED (loss-5). commitXml bundles the FULL
    // styles table (snapshotted at ctor) and round-trips it through xmldom (which XML-escapes `&`, hence no
    // `&` in this fixture — that is correct XML, not a prune concern). The `$` tokens are the dollar-safety
    // proof: they must survive verbatim, and every style (incl. the orphan) must still be present.
    const dollarName = "Cite $0 $$ $1 tokens";
    const styles =
      style("Normal", { default: true }) +
      style("Dollar", { name: dollarName }) +
      style("Orphan", {});
    const body = `<w:p><w:pPr><w:pStyle w:val="Dollar"/></w:pPr><w:r><w:t xml:space="preserve">x</w:t></w:r></w:p>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles }), { pruneStyles: true });
    wb.pruneStylesToClosure(); // prunes the SERIALIZE path; the commitXml aux bundle is independent
    const frag = wb.commitXml(`<w:p xmlns:w="${W}"><w:r><w:t>y</w:t></w:r></w:p>`);
    expect(frag).toContain(`<w:name w:val="${dollarName}"/>`); // dollar tokens verbatim (no $-mangling)
    expect(frag).toContain(`w:styleId="Orphan"`); // FULL unpruned table (loss-5: compat path unpruned)
  });
});

// ===========================================================================
// 9. EXTENDED CLOSURE FIXTURES (FIX-6 / b3-3) — self-closing child, latentStyles
//    with lsdException, settings.xml clickAndTypeStyle ref, tblStyle table style.
// ===========================================================================
describe("flag ON: extended closure fixtures (b3-3)", () => {
  it("a settings.xml clickAndTypeStyle ref seeds the closure (reg-1)", () => {
    // The body cites NOTHING; the only reference to ClickStyle is settings.xml's <w:clickAndTypeStyle>.
    const body = `<w:p><w:r><w:t xml:space="preserve">no body style refs</w:t></w:r></w:p>`;
    const styles =
      style("Normal", { default: true }) +
      style("ClickStyle", { inner: `<w:basedOn w:val="Normal"/>` }) +
      style("Orphan", {});
    const settingsXml = `<w:clickAndTypeStyle w:val="ClickStyle"/><w:defaultTabStop w:val="720"/>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles, settings: settingsXml }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const kept = styleIds(wb.stylesXmlForTest());
    expect(kept).toContain("ClickStyle"); // seeded by settings.xml — the gap is closed
    expect(kept).not.toContain("Orphan");
  });

  it("a tblStyle-cited table style (and its basedOn chain) is kept", () => {
    const body =
      `<w:tbl><w:tblPr><w:tblStyle w:val="MyTable"/></w:tblPr>` +
      `<w:tr><w:tc><w:p><w:r><w:t xml:space="preserve">cell</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`;
    const styles =
      style("Normal", { default: true }) +
      style("TableBase", { type: "table", inner: `<w:basedOn w:val="Normal"/>` }) +
      style("MyTable", { type: "table", inner: `<w:basedOn w:val="TableBase"/>` }) +
      style("Orphan", { type: "table" });
    const wb = new WholeBodyPackage(pkgWith({ body, styles }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const kept = styleIds(wb.stylesXmlForTest());
    expect(kept).toContain("MyTable"); // cited via <w:tblStyle>
    expect(kept).toContain("TableBase"); // via MyTable basedOn
    expect(kept).toContain("Normal");
    expect(kept).not.toContain("Orphan");
  });

  it("a self-closing `<w:style/>` child in the closure + latentStyles with lsdException survive", () => {
    // latentStyles carrying <w:lsdException> is NON-<w:style> markup — it must round-trip VERBATIM (the
    // prune only touches <w:style> spans). The self-closing cited style must be kept.
    const latent =
      `<w:latentStyles w:count="2"><w:lsdException w:name="Normal"/><w:lsdException w:name="heading 1"/></w:latentStyles>`;
    const styles =
      style("Normal", { default: true }) +
      style("SelfCited", { selfClosing: true }) +
      style("Orphan", {});
    const body = `<w:p><w:r><w:rPr><w:rStyle w:val="SelfCited"/></w:rPr><w:t xml:space="preserve">x</w:t></w:r></w:p>`;
    const wb = new WholeBodyPackage(pkgWith({ body, styles, latentStyles: latent }), { pruneStyles: true });
    wb.pruneStylesToClosure();
    const stylesXml = wb.stylesXmlForTest();
    expect(styleIdsAll(stylesXml)).toEqual(new Set(["Normal", "SelfCited"])); // self-closing-aware
    // latentStyles + its lsdException children survive byte-for-byte.
    expect(stylesXml).toContain(latent);
  });
});

// ===========================================================================
// 10. REAL-CORPUS CLOSED-WORLD GATE (FIX-6 / 002-S6 loss-5). Discover a real
//     styles-bearing .docx, prune ON, assert NO retained part references a removed
//     styleId. Presence-gated on discoverSamples()>0 — skip-clean when absent.
// ===========================================================================
describe("flag ON: real-corpus closed-world closure (002-S6, presence-gated)", () => {
  const samples = discoverSamples();
  // Skip-clean when samples/ is absent (CI / a fresh checkout) — a junction supplies them locally.
  const maybe = samples.length > 0 ? it : it.skip;

  maybe("no retained part references a styleId the prune removed (real styles.xml)", async () => {
    // Use the smallest sample that actually carries a styles part (the prune is a no-op without one).
    let chosen: { documentXml: string; stylesXml: string } | null = null;
    for (const s of samples) {
      const { documentXml, stylesXml } = await readDocxParts(s.fullPath);
      if (stylesXml) {
        chosen = { documentXml, stylesXml };
        break;
      }
    }
    expect(chosen).not.toBeNull();
    const { documentXml, stylesXml } = chosen!;

    // Build a flat-OPC package from the REAL document + styles, prune ON, then serialize and re-extract
    // the pruned styles table. Closed-world assert: every styleId still cited by the (verbatim) document
    // part is PRESENT in the pruned styles table, and no retained style's basedOn/link/next/numStyleLink/
    // styleLink dangles at a removed id.
    const pkg =
      `<pkg:package xmlns:pkg="${PKG}">` +
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
      stylesXml.replace(/^﻿/, "").replace(/^\s*<\?xml[^>]*\?>\s*/, "") +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml" pkg:contentType="${DOC_CT}"><pkg:xmlData>` +
      documentXml.replace(/^﻿/, "").replace(/^\s*<\?xml[^>]*\?>\s*/, "") +
      `</pkg:xmlData></pkg:part></pkg:package>`;

    const wb = new WholeBodyPackage(pkg, { pruneStyles: true });
    wb.pruneStylesToClosure();
    const prunedStyles = wb.stylesXmlForTest();
    const kept = styleIds(prunedStyles);
    expect(kept.size).toBeGreaterThan(0);

    // (a) every pStyle/rStyle/tblStyle the document.xml still cites is retained (no needed style dropped).
    const out = wb.serialize();
    const docPart = /<pkg:part pkg:name="\/word\/document\.xml"[\s\S]*?<\/pkg:part>/.exec(out)![0];
    for (const tag of ["pStyle", "rStyle", "tblStyle"]) {
      const re = new RegExp(`<w:${tag}\\b[^>]*\\bw:val="([^"]+)"`, "g");
      let m: RegExpExecArray | null;
      while ((m = re.exec(docPart)) !== null) {
        expect(kept).toContain(m[1]); // a cited style was never pruned away
      }
    }

    // (b) no retained style's cross-references dangle at a removed id (closed-world).
    const styleRe = /<w:style\b[^>]*\bw:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g;
    let sm: RegExpExecArray | null;
    let checked = 0;
    while ((sm = styleRe.exec(prunedStyles)) !== null) {
      for (const target of refTargets(sm[2])) {
        checked++;
        expect(kept).toContain(target);
      }
    }
    expect(checked).toBeGreaterThan(0); // the closed-world assertion ran on real edges
  });
});
