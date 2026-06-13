// LOOP 002 B1 — node-direct Hide over the REAL debate corpus (R4-CORPUS · 002-S1 on real bytes).
//
// The realDocsShard*.test.ts family round-trips every sample through the DEFAULT (per-paragraph
// compat) port. THIS file proves the same losslessness for the NEW default live path — the
// node-direct pure-whole-body pipeline (`pureWholeBody:true`) — on Word's OWN serialization of the
// real ndca/dds2/ExFlex docs, and captures the engine-boundary transformMs for the A/B re-baseline.
//
// For each sample it:
//   1. reads `word/document.xml` from the .docx (Word's real OOXML) and wraps it as a flat-OPC
//      package — the exact shape `body.getOoxml()` returns;
//   2. drives a node-direct Hide, capturing the package handed to the host's `insertOoxml`;
//   3. asserts `assertVanishBridgeOnlyDelta(readPackage, committedPackage)` — the committed body
//      differs from the read body ONLY by `<w:vanish>` toggles + whitelisted bridge-split runs
//      (the whole-body 002-F1 losslessness proof, on real content); and
//   4. round-trips Show All (native reveal in the fake) and asserts nothing stays hidden.
//
// Heavy tiers (xlarge) are deferred unless ROSTRUM_PERF=1 (a 26M-char document.xml is slow to parse),
// exactly like the shard family. Absent samples → a green placeholder (CI has no samples/).

import { createOfficeWordPort } from "../src/core/officeWordPort";
import { hide, showAll } from "../src/core/invisibility";
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { assertVanishBridgeOnlyDelta } from "./semanticDiff";
import { discoverSamples, readDocxParts, SampleRef, SizeTier } from "./realDocs";
import { mkDoc, para, run, harness, settings, hiddenFlags } from "./fakeWord";

const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
/** The default keep set the live add-in ships (mirrors realDocs.ts KEEP_COLORS). */
const KEEP = settings(["cyan", "yellow", "green", "lightGray", "magenta", "red"]);
const HEAVY_TIERS: SizeTier[] = ["large", "xlarge"];
const RUN_HEAVY = process.env.ROSTRUM_PERF === "1";

/** Strip a UTF-8 BOM / XML prolog so a whole part can be embedded inside `<pkg:xmlData>`. */
const stripProlog = (xml: string): string => xml.replace(/^﻿/, "").replace(/^\s*<\?xml[^>]*\?>\s*/, "");
/** Wrap Word's real `document.xml` as the flat-OPC package `body.getOoxml()` returns. */
const wrapPackage = (documentXml: string): string =>
  `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
  `${stripProlog(documentXml)}</pkg:xmlData></pkg:part></pkg:package>`;

describe("node-direct Hide round-trips the REAL corpus losslessly (R4-CORPUS / 002-S1)", () => {
  const all = discoverSamples();
  const active = all.filter((s) => RUN_HEAVY || !HEAVY_TIERS.includes(s.tier));

  if (!active.length) {
    it("no .docx samples — drop files into rostrum/samples to enable the node-direct corpus gate", () => {
      expect(active).toHaveLength(0);
    });
    return;
  }

  it.each(active)(
    "commits a vanish/bridge-only delta and fully reverses — $file [$tier]",
    async (s: SampleRef) => {
      const { documentXml } = await readDocxParts(s.fullPath);
      const readPackage = wrapPackage(documentXml);
      const count = new WholeBodyPackage(readPackage).count;
      expect(count).toBeGreaterThan(0);

      // Back the fake with `count` paragraphs (the pure read classifies the PACKAGE, not the proxies,
      // so only the count must align for the fake's write-back). `bodyOoxmlOverride` returns Word's
      // real package from getOoxml.
      const doc = mkDoc(Array.from({ length: count }, () => para(run("placeholder body"))));
      const h = harness(doc, readPackage);
      let committed: string | null = null;
      const body: any = (h.ctx as any).document.body;
      const realInsert = body.insertOoxml.bind(body);
      body.insertOoxml = (xml: string, loc?: string): void => {
        committed = xml;
        realInsert(xml, loc);
      };
      const port = createOfficeWordPort({ runner: h.runner, pureWholeBody: true, logger: h.tracer.logger("adapter") });

      const t0 = Date.now();
      const res = await hide(port, KEEP);
      const transformMs = Date.now() - t0;

      expect(res.paragraphsScanned).toBe(count);
      expect(res.paragraphsSkipped).toBe(0);
      expect(res.paragraphsChanged).toBeGreaterThan(0);
      // No proxy/alignment fallback ever runs on the pure node-direct path.
      expect(h.warnings.some((w) => /fallback|exhausted|targeted|unparseable/.test(w))).toBe(false);

      // THE LOSSLESSNESS GATE (002-F1) on real content: vanish + bridge only, nothing else.
      const out = committed ?? readPackage;
      assertVanishBridgeOnlyDelta(readPackage, out);

      // Reverse: native Show All clears every vanish; nothing remains hidden.
      const t1 = Date.now();
      await showAll(port);
      const showMs = Date.now() - t1;
      const stillHidden = doc.paragraphs.filter((p) => {
        try {
          return hiddenFlags(p.xml).some(Boolean);
        } catch {
          return false;
        }
      }).length;
      expect(stillHidden).toBe(0);

      // eslint-disable-next-line no-console
      console.log(
        `[realdoc-node ${s.tier}] ${s.file}: ${count} paras | hid ${res.paragraphsChanged} ` +
          `(skipped ${res.paragraphsSkipped}) | node-direct transform ${transformMs}ms | showAll ${showMs}ms`
      );
    },
    300000
  );
});
