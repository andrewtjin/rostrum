// Adapter tests with NO Word host (fake `RequestContext` in ./fakeWord). These
// assert the adapter's sequencing: single-sync atomic commit (H3), manifest
// set-vs-add, Track-Changes restore errors (TS#2), the multi-<w:p> guard (#5),
// outline normalization, and the whole-body alignment fallback. The real engine
// (hide/showAll) runs on top of the fake unchanged.

import { createOfficeWordPort, CancelledError } from "../src/core/officeWordPort";
import { hide, showAll } from "../src/core/invisibility";
import { TrackChangesActiveError } from "../src/core/guards";
import { readRuns } from "../src/core/ooxml";
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { parseManifestOrNull, MANIFEST_NAMESPACE } from "../src/core/manifest";
import {
  W_NS,
  PKG_NS,
  para,
  run,
  mkDoc,
  buildPackage,
  harness,
  hiddenFlags,
  settings,
  FakeDoc
} from "./fakeWord";

// ===========================================================================
// Read
// ===========================================================================
describe("readParagraphs", () => {
  it("normalizes the numeric outline base and table membership (per-paragraph)", async () => {
    const doc = mkDoc([
      para(run("Heading"), { outlineNumber: 1 }), // 1-based H1 -> canonical 0
      para(run("body")), // 10 -> null
      para(run("cell"), { inTable: true })
    ]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    const paras = await port.readParagraphs();
    expect(paras.map((p) => p.headingLevel)).toEqual([0, null, null]);
    expect(paras.map((p) => p.inTable)).toEqual([false, false, true]);
    expect(readRuns(paras[0].ooxml)[0].text).toBe("Heading");
  });

  it("reads outline level from numeric Paragraph.outlineLevel — never touches `paragraphFormat` (regression: that property doesn't exist on a paragraph)", async () => {
    const doc = mkDoc([
      para(run("H4"), { outlineNumber: 4 }), // 1-based H4 -> canonical 3
      para(run("body")) // 10 -> null
    ]);
    const h = harness(doc);
    // FakeParagraph models the real surface and has NO `paragraphFormat`; if the
    // adapter ever reads `p.paragraphFormat.load(...)` again, this throws the exact
    // live-host error this guards against ("undefined (reading 'load')").
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    const paras = await port.readParagraphs();
    expect(paras.map((p) => p.headingLevel)).toEqual([3, null]);
  });

  it("reports progress across chunks", async () => {
    const doc = mkDoc([para(run("a")), para(run("b")), para(run("c")), para(run("d")), para(run("e"))]);
    const h = harness(doc);
    const progress: number[] = [];
    const port = createOfficeWordPort({
      runner: h.runner,
      chunkSize: 2,
      onProgress: (i) => i.phase === "read" && progress.push(i.done),
      logger: h.tracer.logger("adapter")
    });
    await port.readParagraphs();
    expect(progress).toEqual([2, 4, 5]); // 2-paragraph chunks over 5 paragraphs
  });

  it("uses the whole-body package when its count aligns", async () => {
    const doc = mkDoc([para(run("a"), { outlineNumber: 1 }), para(run("b"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(paras.map((p) => readRuns(p.ooxml)[0].text)).toEqual(["a", "b"]);
    expect(h.warnings).toHaveLength(0); // no alignment fallback
  });

  it("tolerantly skips an EXTRA package paragraph (Word's ±1 getOoxml serialization quirk)", async () => {
    // This is the ndca-semis scenario: body.getOoxml() emits one more <w:p> than
    // body.paragraphs enumerates. Stage 4.1 aligns by text and skips the artifact,
    // so the whole-body fast path engages instead of falling back to the slow read.
    const doc = mkDoc([para(run("alpha")), para(run("bravo"))]);
    const withArtifact = buildPackage([...doc.paragraphs, para(run("phantom"))]);
    const h = harness(doc, withArtifact);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(paras.map((p) => readRuns(p.ooxml)[0].text)).toEqual(["alpha", "bravo"]); // aligned; phantom skipped
    expect(h.warnings.some((w) => /fallback|exhausted/.test(w))).toBe(false); // whole-body used, no fallback
  });

  it("re-reads only unalignable paragraphs (here: all of them) via targeted getOoxml", async () => {
    const doc = mkDoc([para(run("alpha")), para(run("bravo"))]);
    // A package whose paragraph texts don't match the collection at all → none align.
    const unalignable = buildPackage([para(run("zzz")), para(run("yyy"))]);
    const h = harness(doc, unalignable);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(paras).toHaveLength(2); // driven by the collection
    // The targeted re-read reads each paragraph's OWN ooxml (the real backing content).
    expect(paras.map((p) => readRuns(p.ooxml)[0].text)).toEqual(["alpha", "bravo"]);
    expect(h.warnings.some((w) => /fallback|exhausted/.test(w))).toBe(true);
  });

  it("re-reads ONLY the divergent paragraph, keeping the cached package for the rest (Stage 4.2 partial fallback)", async () => {
    // THE regression fix: one paragraph whose proxy `.text` can't be reconciled with the
    // package (the ndca emoji at proxy 259) must cost ONE re-read, not collapse all 376 to
    // the slow path. We distinguish "came from the package" (2 runs) from "re-read from its
    // own backing" (1 run) so the assertion proves only the middle paragraph was re-read.
    const twoRun = (t: string): FakeDoc["paragraphs"][number] => ({
      xml: `<w:p xmlns:w="${W_NS}">${run(t)}${run("")}</w:p>`,
      outlineNumber: 10,
      inTable: false
    });
    const doc = mkDoc([para(run("alpha")), para(run("divergent")), para(run("gamma"))]);
    const pkg = buildPackage([twoRun("alpha"), twoRun("WRONGTEXT"), twoRun("gamma")]); // middle diverges
    const h = harness(doc, pkg);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(paras.map((p) => readRuns(p.ooxml).map((r) => r.text).join(""))).toEqual([
      "alpha",
      "divergent",
      "gamma"
    ]);
    // Outer two kept the cached package fragment (2 runs); the middle was re-read from its own
    // backing (1 run) — exactly one targeted getOoxml, not a whole-document fallback.
    expect(readRuns(paras[0].ooxml).length).toBe(2);
    expect(readRuns(paras[2].ooxml).length).toBe(2);
    expect(readRuns(paras[1].ooxml).length).toBe(1);
    expect(h.warnings.some((w) => /targeted.*fallback/.test(w))).toBe(true);
  });

  it("re-reads EQUAL-COUNT duplicates instead of trusting a count-parity match — text equality ≠ content identity (Stage A REJECTED, lesson #37)", async () => {
    // Adversarial review proved that trusting a duplicate when its proxy/package occurrence COUNTS
    // match (`pkgFreq === proxyFreq`) is UNSOUND: two same-text paragraphs can differ in content,
    // and `body.getOoxml()` may serialize them in a different relative order than `body.paragraphs`
    // enumerates them. Here the proxies are [ "K" highlighted-to-KEEP, "K" plain ] and the package
    // serializes the two "K" paragraphs REVERSED. Count-parity would bind each proxy to the wrong
    // (swapped) package slot — and since the mapping still looks identity-clean, `cleanAlign` would
    // go true and the destructive whole-body Replace would HIDE the highlighted keep (violating the
    // always-show-highlighted ethos). Uniqueness-only re-reads both "K" from their OWN backing, so
    // the highlight is preserved and no swap is possible. THIS TEST FAILS if count-parity returns.
    const twoRun = (t: string): FakeDoc["paragraphs"][number] => ({
      xml: `<w:p xmlns:w="${W_NS}">${run(t)}${run("")}</w:p>`,
      outlineNumber: 10,
      inTable: false
    });
    const doc = mkDoc([
      para(run("head"), { outlineNumber: 1 }),
      para(run("K", { highlight: "yellow" })), // highlighted keep
      para(run("K")) // plain
    ]);
    const pkg = buildPackage([twoRun("head"), twoRun("K"), twoRun("K")]); // 2-run package, same text
    const h = harness(doc, pkg);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    // Both "K" came from their OWN backing (1 run), never a 2-run package slot → no swap.
    expect(readRuns(paras[1].ooxml).length).toBe(1);
    expect(readRuns(paras[2].ooxml).length).toBe(1);
    // The highlighted keep still carries its highlight (it would be lost if swapped/hidden).
    expect(readRuns(paras[1].ooxml)[0].highlight).toBe("yellow");
    expect(readRuns(paras[2].ooxml)[0].highlight).toBeNull();
    expect(h.warnings.some((w) => /targeted.*fallback/.test(w))).toBe(true);
  });

  it("re-reads a duplicate the package holds MORE often than the proxies (real + same-text artifact) — #33 anti-corruption", async () => {
    // The genuine corruption case the uniqueness guard was built for: a real paragraph PLUS a
    // same-text serialization artifact, so the package holds "dup" twice but the proxies hold it
    // once. A non-empty text that is not UNIQUE in the package → unresolved (re-read exactly), so
    // the proxy is never bound to the artifact and no foreign content is committed onto it. (Under
    // the reverted uniqueness-only guard ALL duplicates re-read; this case is the original #33.)
    const twoRun = (t: string): FakeDoc["paragraphs"][number] => ({
      xml: `<w:p xmlns:w="${W_NS}">${run(t)}${run("")}</w:p>`,
      outlineNumber: 10,
      inTable: false
    });
    const doc = mkDoc([para(run("unique head"), { outlineNumber: 1 }), para(run("dup"))]);
    const pkg = buildPackage([twoRun("unique head"), twoRun("dup"), twoRun("dup")]); // extra "dup" artifact
    const h = harness(doc, pkg);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(paras.map((p) => readRuns(p.ooxml).map((r) => r.text).join(""))).toEqual(["unique head", "dup"]);
    expect(readRuns(paras[0].ooxml).length).toBe(2); // unique head: package fragment
    expect(readRuns(paras[1].ooxml).length).toBe(1); // dup: re-read exactly, not bound to the artifact
    expect(h.warnings.some((w) => /targeted.*fallback/.test(w))).toBe(true);
  });

  it("a divergent proxy with NO package slot (pk < N) cannot desync the cursor into a wrong duplicate match (Stage A Finding 2, lesson #37)", async () => {
    // The other half of why count-parity is unsound: the blind `j++` on an unresolved proxy assumes
    // a 1:1 slot. When a proxy diverges and has NO package paragraph (pk < N — a realizable host
    // quirk; lesson #27 never assumes the counts agree), that increment moves the cursor onto a real
    // paragraph, and a count-parity duplicate would then trust the WRONG occurrence. Here proxy 0
    // ("Z") has no package slot; proxies 1 and 2 are duplicate "K"s. Uniqueness-only re-reads all
    // three from backing (each 1 run), so the divergence can't mis-bind a duplicate. THIS TEST
    // FAILS if count-parity returns (proxy 1 would bind to package slot 1, the wrong "K").
    const twoRun = (t: string): FakeDoc["paragraphs"][number] => ({
      xml: `<w:p xmlns:w="${W_NS}">${run(t)}${run("")}</w:p>`,
      outlineNumber: 10,
      inTable: false
    });
    const doc = mkDoc([para(run("Z")), para(run("K")), para(run("K"))]); // 3 proxies
    const pkg = buildPackage([twoRun("K"), twoRun("K")]); // only 2 package paras — Z has no slot
    const h = harness(doc, pkg);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(paras.map((p) => readRuns(p.ooxml).map((r) => r.text).join(""))).toEqual(["Z", "K", "K"]);
    // All re-read from their own backing (1 run each) — no duplicate bound to a desynced slot.
    expect(paras.map((p) => readRuns(p.ooxml).length)).toEqual([1, 1, 1]);
    expect(h.warnings.some((w) => /targeted.*fallback/.test(w))).toBe(true);
  });

  it("aligns a paragraph whose proxy text carries an emoji the package omits — no re-read (Stage 4.2)", async () => {
    // The ndca trigger: `collectParagraphText` can't render the `w16se:symEx` emoji, so the
    // package text lacks it while the proxy `.text` includes it. `normForAlign` strips astral
    // codepoints on BOTH sides, so the paragraph aligns cleanly off the 0.8s package — no
    // targeted re-read. `String.fromCodePoint` keeps the literal emoji out of the source.
    const emoji = String.fromCodePoint(0x1f600);
    const doc = mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run(`card with ${emoji} emoji`))]);
    const pkgNoEmoji = buildPackage([para(run("Heading"), { outlineNumber: 1 }), para(run("card with emoji"))]);
    const h = harness(doc, pkgNoEmoji);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    const paras = await port.readParagraphs();
    expect(h.warnings.some((w) => /fallback|exhausted/.test(w))).toBe(false); // aligned, no re-read
    expect(readRuns(paras[1].ooxml)[0].text).toBe("card with emoji"); // package fragment used
  });

  it("cancels during the read phase before writing anything", async () => {
    const doc = mkDoc([para(run("a")), para(run("b"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({
      runner: h.runner,
      chunkSize: 1,
      cancel: { isCancelled: () => true },
      logger: h.tracer.logger("adapter")
    });
    await expect(port.readParagraphs()).rejects.toBeInstanceOf(CancelledError);
  });
});

// ===========================================================================
// Atomic commit (H3) + manifest set-vs-add, via the real engine
// ===========================================================================
describe("hide() through the adapter — atomic commit", () => {
  function sampleDoc(): FakeDoc {
    return mkDoc([
      para(run("Contention"), { outlineNumber: 1 }),
      para(run("a long card body")),
      para(run("Author 2019", { cite: true })),
      para(run("intro ") + run("warrant", { highlight: "yellow" }))
    ]);
  }

  it("hides the right paragraphs and arms the manifest", async () => {
    const doc = sampleDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    const res = await hide(port, settings(["yellow"]));

    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // heading kept
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // card hidden
    expect(hiddenFlags(doc.paragraphs[2].xml)).toEqual([false]); // cite kept
    expect(hiddenFlags(doc.paragraphs[3].xml)).toEqual([true, false]); // partial
    // The partial paragraph commits via Paragraph.insertOoxml (the others hide natively).
    // That is the package the LIVE host receives, and it REQUIRES the OPC start-part
    // relationship — without it insertOoxml threw GeneralException (the Stage 4.2 commit
    // regression). Assert the committed fragment carries /_rels/.rels → officeDocument, so
    // this path is guarded at the adapter, not only in the paragraphXml unit test.
    expect(doc.paragraphs[3].xml).toContain(`pkg:name="/_rels/.rels"`);
    expect(doc.paragraphs[3].xml).toContain("officeDocument");
    expect(parseManifestOrNull(doc.manifest!.xml)!.active).toBe(true);
    expect(res.paragraphsChanged).toBe(2);
  });

  it("commits paragraph edits AND the manifest in ONE sync (H3 atomicity)", async () => {
    const doc = sampleDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await hide(port, settings(["yellow"]));

    // Stage 4: the whole card body (para 1) hides via a NATIVE font.hidden toggle; only
    // the partial-highlight paragraph (para 3) needs an OOXML replace. Both still land
    // with the manifest in ONE sync — the atomicity invariant is unchanged.
    const native = h.ctx.commitLog.filter((c) => c.op === "paragraph.font.hidden");
    const inserts = h.ctx.commitLog.filter((c) => c.op === "paragraph.insertOoxml");
    const manifest = h.ctx.commitLog.filter((c) => c.op === "xmlParts.add");
    expect(native.length).toBe(1);
    expect(inserts.length).toBe(1);
    expect(manifest.length).toBe(1);
    // The decisive assertion: every paragraph edit (native + OOXML) and the manifest
    // write land in the SAME sync — so a failure can't leave the doc hidden-but-unarmed.
    const syncs = new Set([...native, ...inserts, ...manifest].map((c) => c.sync));
    expect(syncs.size).toBe(1);
  });

  it("uses setXml when a manifest already exists, add when it doesn't", async () => {
    const doc = sampleDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await hide(port, settings(["yellow"])); // first arm -> add
    expect(h.ctx.commitLog.some((c) => c.op === "xmlParts.add")).toBe(true);

    h.ctx.commitLog.length = 0;
    await hide(port, settings(["green"])); // re-arm -> setXml in place
    expect(h.ctx.commitLog.some((c) => c.op === "part.setXml")).toBe(true);
    expect(h.ctx.commitLog.some((c) => c.op === "xmlParts.add")).toBe(false);
    expect(parseManifestOrNull(doc.manifest!.xml)!.keepColors).toEqual(["green"]);
  });

  it("whole-body strategy also commits atomically in one sync", async () => {
    const doc = sampleDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    await hide(port, settings(["yellow"]));
    const body = h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml");
    const manifest = h.ctx.commitLog.filter((c) => c.op === "xmlParts.add");
    expect(body.length).toBe(1);
    expect(new Set([...body, ...manifest].map((c) => c.sync)).size).toBe(1);
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]);
  });

  it("perf mode with a serialization artifact commits PER-PARAGRAPH (no whole-body Replace → no phantom injection)", async () => {
    // Whole-body READ still engages (fast) and aligns tolerantly, but because an artifact
    // <w:p> was skipped the alignment isn't clean, so a full-body Replace would re-inject
    // the artifact. The commit must therefore stay per-paragraph (surgical).
    const doc = mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card body"))]);
    const withArtifact = buildPackage([...doc.paragraphs, para(run("phantom"))]);
    const h = harness(doc, withArtifact);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    await hide(port, settings(["yellow"]));
    expect(h.ctx.commitLog.some((c) => c.op === "body.insertOoxml")).toBe(false); // NOT a full-body Replace
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.font.hidden")).toBe(true); // surgical native hide
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // card hidden
    expect(doc.paragraphs).toHaveLength(2); // no phantom injected into the document
  });

  it("partial-hide commit bundles the document's styles so style-inherited formatting survives", async () => {
    // The whole-body fidelity bug: an aligned paragraph is READ from the minimal (style-LESS)
    // package, so committing it via a bare insertOoxml makes the host resolve style-inherited
    // formatting (underline, character box, font size) to document DEFAULTS — underlined/boxed text
    // collapsed to plain 11pt while inline highlight survived. The per-paragraph commit must
    // re-wrap the fragment through commitXml so the styles part rides along. Model Word's package:
    // the 4 story paragraphs PLUS a styles part and its locating relationship.
    const doc = sampleDoc();
    const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    const RELS_NS = "http://schemas.openxmlformats.org/package/2006/relationships";
    const STYLES_CT = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
    const styledOverride =
      `<pkg:package xmlns:pkg="${PKG_NS}">` +
      `<pkg:part pkg:name="/word/_rels/document.xml.rels"><pkg:xmlData>` +
      `<Relationships xmlns="${RELS_NS}"><Relationship Id="rIdS" Type="${REL}/styles" Target="styles.xml"/></Relationships>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/styles.xml" pkg:contentType="${STYLES_CT}"><pkg:xmlData>` +
      `<w:styles xmlns:w="${W_NS}"><w:style w:styleId="Style13ptBold"><w:rPr><w:u w:val="single"/></w:rPr></w:style></w:styles>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}"><w:body>${doc.paragraphs.map((pp) => pp.xml).join("")}` +
      `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>` +
      `</pkg:xmlData></pkg:part></pkg:package>`;
    const h = harness(doc, styledOverride);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await hide(port, settings(["yellow"]));

    // The partial paragraph (index 3) commits via Paragraph.insertOoxml; its fragment must now
    // carry the styles part AND the relationship that lets the host locate it.
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.insertOoxml")).toBe(true);
    expect(doc.paragraphs[3].xml).toContain(`pkg:name="/word/styles.xml"`);
    expect(doc.paragraphs[3].xml).toContain(`${REL}/styles`);
    expect(doc.paragraphs[3].xml).toContain("officeDocument"); // start part still present
    // The hide still landed correctly: warrant (yellow) kept, intro hidden.
    expect(hiddenFlags(doc.paragraphs[3].xml)).toEqual([true, false]);
  });

});

// ===========================================================================
// showAll
// ===========================================================================
describe("showAll() through the adapter", () => {
  it("safe mode reveals via per-paragraph native font.hidden — NO OOXML rewrite", async () => {
    const doc = mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await hide(port, settings(["yellow"]));
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]);

    h.ctx.commitLog.length = 0;
    await showAll(port);
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([false]); // revealed
    expect(doc.manifest).toBeNull(); // disarmed
    // Stage 4: the reveal is property-only — native font.hidden=false per paragraph,
    // with the slow per-paragraph `insertOoxml` reflow path GONE.
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.font.hidden")).toBe(true);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.insertOoxml")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "range.insertOoxml")).toBe(false);
    expect(h.ctx.commitLog.filter((c) => c.op === "part.delete").length).toBe(1);
  });

  it("perf mode reveals with ONE whole-body font.hidden clear", async () => {
    const doc = mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({
      runner: h.runner,
      commitStrategy: "whole-body",
      logger: h.tracer.logger("adapter")
    });
    await hide(port, settings(["yellow"]));
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]);

    h.ctx.commitLog.length = 0;
    await showAll(port);
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([false]);
    expect(doc.manifest).toBeNull();
    // One blanket clear over the whole story — not a per-paragraph loop (the Verbatim
    // ~instant move, but selective-keep on the Hide side).
    expect(h.ctx.commitLog.filter((c) => c.op === "body.font.hidden").length).toBe(1);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.font.hidden")).toBe(false);
  });
});

// ===========================================================================
// Guards + error surfacing
// ===========================================================================
describe("writeParagraphs guards", () => {
  it("rejects a multi-<w:p> writeback fragment (#5)", async () => {
    const doc = mkDoc([para(run("x"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await port.readParagraphs();
    const twoParas = `<w:body xmlns:w="${W_NS}"><w:p><w:r><w:t>a</w:t></w:r></w:p><w:p><w:r><w:t>b</w:t></w:r></w:p></w:body>`;
    await expect(port.writeParagraphs([{ index: 0, ooxml: twoParas }])).rejects.toThrow(/found 2/);
  });

  it("rejects an out-of-range index", async () => {
    const doc = mkDoc([para(run("x"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await port.readParagraphs();
    await expect(
      port.writeParagraphs([{ index: 9, ooxml: `<w:p xmlns:w="${W_NS}"><w:r><w:t>x</w:t></w:r></w:p>` }])
    ).rejects.toBeInstanceOf(RangeError);
  });
});

describe("Track Changes", () => {
  it("reads the mode", async () => {
    const doc = mkDoc([para(run("x"))], "TrackAll");
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    expect(await port.getChangeTrackingMode()).toBe("TrackAll");
  });

  it("auto-toggles off and restores around a hide (engine + adapter)", async () => {
    const doc = mkDoc([para(run("card"))], "TrackAll");
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    const res = await hide(port, settings(["yellow"]), { autoToggleTrackChanges: true });
    expect(res.trackChangesToggled).toBe(true);
    expect(doc.tcMode).toBe("TrackAll"); // restored
    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([true]);
  });

  it("surfaces a clear, mode-named error when setting TC fails (TS#2)", async () => {
    const doc = mkDoc([para(run("x"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    h.ctx.failNext = { code: "GeneralException", message: "host refused" };
    await expect(port.setChangeTrackingMode("Off")).rejects.toThrow(
      /could not set Track Changes to "Off".*inconsistent/s
    );
  });
});

describe("readManifest", () => {
  it("returns the stored manifest XML, or null when absent", async () => {
    const doc = mkDoc([para(run("x"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    expect(await port.readManifest()).toBeNull();

    doc.manifest = { id: "part-1", xml: `<rostrum xmlns="${MANIFEST_NAMESPACE}"/>` };
    expect(await port.readManifest()).toContain("rostrum");
  });
});

// ===========================================================================
// repairCites — the whole-body cite-repair commit (Apply Styles)
// ===========================================================================
describe("repairCites", () => {
  // A tag paragraph whose XML carries an inline outline level, so the WholeBodyPackage
  // resolves headingLevel 3 (the package read has no proxies). The cite paragraph below it
  // has a bold author+year run with NO rStyle (the mis-styled-cite shape).
  // Schema-valid shape: outlineLvl lives inside <w:pPr> — headingLevel reads ONLY the
  // paragraph's own direct-child pPr, exactly like Word itself (a bare outlineLvl outside
  // pPr is markup Word never emits and never honors).
  const tagPPr = `<w:pPr><w:outlineLvl w:val="3"/></w:pPr>`;
  const boldAuthor = `<w:r><w:rPr><w:b/><w:bCs/></w:rPr><w:t xml:space="preserve">Smith 20</w:t></w:r>`;

  it("detects a mis-styled cite after a tag and commits the rStyle via ONE insertOoxml", async () => {
    const doc = mkDoc([
      para(run("Tag heading"), { pPr: tagPPr }),
      para(`${boldAuthor}<w:r><w:t xml:space="preserve"> [descriptor]</w:t></w:r>`)
    ]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    // Apply-loop reuse guard — the REPAIR-path complement of the lazy-wrap test below: the
    // planner wraps the ONE candidate (the cite paragraph; the tag's headingLevel is read
    // eagerly but its fragment never), and the apply loop must REUSE that cached fragment.
    // The fake host's body.insertOoxml splits the committed body via paragraphXml on its OWN
    // package, so calls are filtered by receiver below: the first call is necessarily the
    // adapter's package (the fake's package only exists once the commit lands).
    const spy = jest.spyOn(WholeBodyPackage.prototype, "paragraphXml");
    const res = await port.repairCites();
    expect(res).toEqual({ paragraphsRepaired: 1, runsRepaired: 1 });
    // The whole body was committed in exactly one insertOoxml (no per-paragraph loop).
    expect(h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml")).toHaveLength(1);
    // The cite paragraph now carries the cite character style; the tag is unchanged.
    expect(doc.paragraphs[1].xml).toContain(`<w:rStyle w:val="Style13ptBold"/>`);
    expect(doc.paragraphs[0].xml).not.toContain(`<w:rStyle w:val="Style13ptBold"/>`);
    // Exactly one wrap on the adapter's package — the planner's probe; the apply loop's read
    // is a cache hit. Re-wrapping per repaired paragraph at apply would make this 2 and fail.
    const adapterPkg = spy.mock.contexts[0];
    expect(spy.mock.contexts.filter((c) => c === adapterPkg)).toHaveLength(1);
    spy.mockRestore();
  });

  it("is a no-op (zeros, no commit) when there is nothing to repair", async () => {
    const doc = mkDoc([para(run("Tag heading"), { pPr: tagPPr }), para(run("ordinary body text"))]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    const res = await port.repairCites();
    expect(res).toEqual({ paragraphsRepaired: 0, runsRepaired: 0 });
    expect(h.ctx.commitLog.some((c) => c.op === "body.insertOoxml")).toBe(false);
  });

  it("wraps fragments lazily — only the tag window is materialized, never all N paragraphs", async () => {
    // 1 tag + 1 plain body paragraph + 20 trailing paragraphs. The planner scans each tag
    // window only to its FIRST real paragraph (here index 1, which fails the cite gates),
    // so exactly ONE fragment is ever wrapped — and the no-repair exit skips the commit,
    // so the spy sees ONLY planner-driven `paragraphXml` calls. The bound is deliberately
    // tight: eager materialization (the old shape) would count 22, and a NON-caching
    // getter would count 3 (the planner's list/empty/evaluate probes each re-wrapping),
    // so either regression fails this guard.
    const doc = mkDoc([
      para(run("Tag heading"), { pPr: tagPPr }),
      para(run("ordinary body text")),
      ...Array.from({ length: 20 }, (_, i) => para(run(`trailing body ${i}`)))
    ]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    const spy = jest.spyOn(WholeBodyPackage.prototype, "paragraphXml");
    const res = await port.repairCites();
    expect(res).toEqual({ paragraphsRepaired: 0, runsRepaired: 0 });
    expect(h.ctx.commitLog.some((c) => c.op === "body.insertOoxml")).toBe(false);
    expect(spy.mock.calls.length).toBeLessThanOrEqual(2); // the one candidate, cached
    expect(spy.mock.calls.length).toBeLessThan(doc.paragraphs.length);
    spy.mockRestore();
  });

  it("does not repair a cite-like paragraph that is NOT after a tag (leak guard)", async () => {
    // No heading anywhere → no window opens, so the bold author+year run is left alone.
    const doc = mkDoc([
      para(run("intro body")),
      para(`${boldAuthor}<w:r><w:t xml:space="preserve"> [descriptor]</w:t></w:r>`)
    ]);
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    const res = await port.repairCites();
    expect(res).toEqual({ paragraphsRepaired: 0, runsRepaired: 0 });
    expect(doc.paragraphs[1].xml).not.toContain(`<w:rStyle w:val="Style13ptBold"/>`);
  });
});

// ===========================================================================
// Loop 002 B2 — single-Word.run read fusion (002-S4)
// ===========================================================================
//
// The TC-mode read is fused into `readParagraphs`' FIRST sync, so the engine's prefetched-mode gate
// consumes the primed value instead of issuing its own TC-read `Word.run`: a clean Hide costs 2 runs
// (read + commit), not 3 (TC read + read + commit). These prove the run-count win on BOTH read paths,
// the abort-before-parse ordering (TC on, no auto-toggle → ZERO classify, ZERO writes), and the
// auto-toggle discard-and-re-read (a TC-on package may carry `w:trackChanges` in bundled settings).
describe("read fusion — clean Hide uses 2 Word.runs, was 3 (B2 / 002-S4)", () => {
  /**
   * Wrap a harness runner so every `Word.run` invocation is counted (one per host round-trip). The
   * count lives on a mutable `{ runs }` holder the caller reads after the op — one entry per `Word.run`.
   */
  function counting(h: ReturnType<typeof harness>): { runner: typeof h.runner; calls: { runs: number } } {
    const calls = { runs: 0 };
    const runner = (<T,>(b: (c: Word.RequestContext) => Promise<T>): Promise<T> => {
      calls.runs++;
      return h.runner(b);
    }) as typeof h.runner;
    return { runner, calls };
  }

  // In the pure node-direct path outline comes from the PACKAGE, not the proxy, so the heading carries
  // an INLINE `<w:outlineLvl>` (the proxy outlineNumber would be ignored).
  const lvl0 = `<w:pPr><w:outlineLvl w:val="0"/></w:pPr>`;

  it("PURE node-direct path: a clean Hide uses exactly 2 Word.runs (read + commit)", async () => {
    const doc = mkDoc([
      para(run("Heading"), { pPr: lvl0 }),
      para(run("a long card body that should vanish")),
      para(run("intro ") + run("warrant", { highlight: "yellow" }))
    ]);
    const h = harness(doc);
    const c = counting(h);
    const port = createOfficeWordPort({
      runner: c.runner,
      pureWholeBody: true,
      logger: h.tracer.logger("adapter")
    });

    await hide(port, settings(["yellow"]));

    // 2 runs: ONE read (TC prefetched in its first sync) + ONE atomic commit. The dedicated TC-read
    // run is gone — that is the 3→2 win. (Pre-B2 this was 3.)
    expect(c.calls.runs).toBe(2);
    // Exactly one whole-body read sync and one commit; no separate TC-read round-trip is observable.
    expect(h.ctx.commitLog.filter((c) => c.op === "body.getOoxml")).toHaveLength(1);
    expect(h.ctx.commitLog.filter((c) => c.op === "body.insertOoxml")).toHaveLength(1);
    // The work still landed correctly (composition change only, not a behavior change).
    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // heading kept
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]); // body hidden
    expect(parseManifestOrNull(doc.manifest!.xml)!.active).toBe(true);
  });

  it("PROXY (per-paragraph) path: a clean Hide uses exactly 2 Word.runs too", async () => {
    // The default (non-pure) read also prefetches TC in its first sync, so even the per-paragraph
    // commit path drops the dedicated TC-read run.
    const doc = mkDoc([para(run("Heading"), { outlineNumber: 1 }), para(run("card body"))]);
    const h = harness(doc);
    const c = counting(h);
    const port = createOfficeWordPort({ runner: c.runner, logger: h.tracer.logger("adapter") });

    await hide(port, settings(["yellow"]));

    expect(c.calls.runs).toBe(2); // read (TC primed) + commit
    expect(h.ctx.commitLog.filter((c) => c.op === "body.getOoxml")).toHaveLength(1);
    expect(hiddenFlags(doc.paragraphs[1].xml)).toEqual([true]);
  });

  it("TC on, NO auto-toggle: aborts BEFORE the parse — ZERO classify, ZERO writes, manifest unarmed", async () => {
    const doc = mkDoc([para(run("card body"))], "TrackAll");
    const h = harness(doc);
    const c = counting(h);
    const port = createOfficeWordPort({
      runner: c.runner,
      pureWholeBody: true,
      logger: h.tracer.logger("adapter")
    });

    await expect(hide(port, settings(["yellow"]))).rejects.toBeInstanceOf(TrackChangesActiveError);

    // The read happened (it primed TC); the gate then threw BEFORE the body ran — so NO commit run.
    expect(c.calls.runs).toBe(1);
    // Abort-before-parse: nothing was classified or written, the manifest is unarmed, the doc is clean.
    expect(h.ctx.commitLog.some((c) => c.op === "body.insertOoxml")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.insertOoxml")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "paragraph.font.hidden")).toBe(false);
    expect(h.ctx.commitLog.some((c) => c.op === "tc.set:Off")).toBe(false); // no toggle either
    expect(doc.manifest).toBeNull();
    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // unchanged
  });

  it("TC on, auto-toggle: DISCARDS the primed read and re-reads after toggling TC off (S-005)", async () => {
    const doc = mkDoc([para(run("card body"))], "TrackAll");
    const h = harness(doc);
    const port = createOfficeWordPort({
      runner: h.runner,
      pureWholeBody: true,
      logger: h.tracer.logger("adapter")
    });

    const res = await hide(port, settings(["yellow"]), { autoToggleTrackChanges: true });

    expect(res.trackChangesToggled).toBe(true);
    // S-005: TC was toggled Off (in the gate) then RESTORED to its prior mode in `finally`.
    expect(doc.tcMode).toBe("TrackAll");
    const tcOps = h.ctx.commitLog.filter((c) => c.op.startsWith("tc.set:")).map((c) => c.op);
    expect(tcOps).toEqual(["tc.set:Off", "tc.set:TrackAll"]);
    // Discard-and-re-read: the body is read TWICE — once to prime TC (under TC-on), then again after
    // toggling TC off (the primed read may carry `w:trackChanges` in bundled settings and isn't trusted).
    expect(h.ctx.commitLog.filter((c) => c.op === "body.getOoxml")).toHaveLength(2);
    // The toggle-off precedes the re-read, which precedes the restore (ordering of the gate body).
    const order = h.ctx.commitLog.map((c) => c.op);
    const firstReRead = order.indexOf("body.getOoxml", order.indexOf("body.getOoxml") + 1);
    expect(order.indexOf("tc.set:Off")).toBeLessThan(firstReRead);
    expect(firstReRead).toBeLessThan(order.lastIndexOf("tc.set:TrackAll"));
    // And the work landed.
    expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([true]);
    expect(parseManifestOrNull(doc.manifest!.xml)!.active).toBe(true);
  });

  it("b2-1: the auto-toggle path's committed package carries settings with NO w:trackChanges", async () => {
    // The discard-and-re-read's PREMISE (S-005): after toggling TC off, the re-read/committed package omits
    // `w:trackChanges` from its verbatim settings part. Previously UNtested in R1 (the fake had no settings
    // part at all). Model one that DROPS `w:trackChanges` once tcMode flips to Off, drive the auto-toggle
    // node-direct Hide, and assert the committed `insertOoxml` package's settings is clean — so a later
    // re-Hide reading that committed doc would NOT see TC active. (R6 002-S4's TC-on gate is the live confirm.)
    const doc = mkDoc([para(run("card body"))], "TrackAll", { modelSettings: true });
    const h = harness(doc);

    // Sanity: the PRIMED read (taken under TC-on) DOES carry `w:trackChanges` — proving the re-read matters.
    expect(h.ctx.bodyOoxml()).toContain("<w:trackChanges/>");

    // Capture the package handed to the node-direct whole-body commit (the serialized re-read pkg).
    let committed: string | null = null;
    const body: any = (h.ctx as any).document.body;
    const realInsert = body.insertOoxml.bind(body);
    body.insertOoxml = (xml: string, loc?: string): void => {
      committed = xml;
      realInsert(xml, loc);
    };

    const port = createOfficeWordPort({
      runner: h.runner,
      pureWholeBody: true, // node-direct: serializes lastRead.pkg (incl. the verbatim settings part)
      logger: h.tracer.logger("adapter")
    });
    const res = await hide(port, settings(["yellow"]), { autoToggleTrackChanges: true });

    expect(res.trackChangesToggled).toBe(true);
    expect(committed).not.toBeNull();
    // The committed package STILL carries a settings part (verbatim carry, A2), but with NO w:trackChanges:
    // it was serialized from the post-toggle re-read taken with TC off.
    expect(committed!).toContain(`pkg:name="/word/settings.xml"`);
    expect(committed!).not.toContain("<w:trackChanges/>");
    expect(committed!).not.toContain("w:trackChanges");
  });
});
