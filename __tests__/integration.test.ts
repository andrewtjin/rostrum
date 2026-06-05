// Engine-through-adapter integration on Debate.dotm-derived fixtures, plus a
// 200-page perf pass. These exercise the FULL stack — Office.js adapter (fake host)
// + the unchanged Stage-1 engine — on documents shaped like a real debate brief:
// pocket/hat/block/tag headings (outline levels 0–3), cite paragraphs, body cards,
// highlighted keepers, numbered lists, tables, and empty separators.

import { createOfficeWordPort, CommitStrategy } from "../src/core/officeWordPort";
import { hide, reHide, showAll } from "../src/core/invisibility";
import { isArmed, loadManifest } from "../src/core/manifest";
import { resolveSettings } from "../src/core/settings";
import { para, run, mkDoc, harness, hiddenFlags, settings, FakeDoc, FakePara } from "./fakeWord";

// A numbered-list paragraph's properties: present in the body card's <w:pPr> and
// must survive Hide untouched (the engine only ever adds/removes <w:vanish/>).
const NUM_PPR = `<w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr>`;

/** A realistic single-card brief slice. Indices are referenced by the assertions. */
function briefDoc(): FakeDoc {
  return mkDoc([
    para(run("Pocket: Climate Adv"), { outlineNumber: 1 }), // 0 H1 pocket — kept
    para(run("Hat: Warming bad"), { outlineNumber: 2 }), // 1 H2 hat — kept
    para(run("Block: Sea level"), { outlineNumber: 3 }), // 2 H3 block — kept
    para(run("Tag: Coastal flooding now"), { outlineNumber: 4 }), // 3 H4 tag — kept (level 3)
    para(run("Smith ") + run("’24", { cite: true })), // 4 cite paragraph — kept whole
    para(run("The long warranted card body that should vanish entirely.")), // 5 body — hidden
    para(run("partial ") + run("highlighted keeper", { highlight: "yellow" })), // 6 partial
    para(run("a numbered point that vanishes"), { pPr: NUM_PPR }), // 7 body w/ numbering — hidden
    para(run("cell text"), { inTable: true }), // 8 table — untouched
    para("") // 9 empty separator — untouched
  ]);
}

describe("brief-shaped document (per-paragraph and whole-body)", () => {
  it.each<CommitStrategy>(["per-paragraph", "whole-body"])(
    "hides only non-keeper body text — %s",
    async (commitStrategy) => {
      const doc = briefDoc();
      const h = harness(doc);
      const port = createOfficeWordPort({
        runner: h.runner,
        commitStrategy,
        logger: h.tracer.logger("adapter")
      });
      await hide(port, settings(["yellow"]));

      expect(hiddenFlags(doc.paragraphs[0].xml)).toEqual([false]); // pocket
      expect(hiddenFlags(doc.paragraphs[3].xml)).toEqual([false]); // tag (level 3 kept)
      expect(hiddenFlags(doc.paragraphs[4].xml)).toEqual([false, false]); // cite paragraph
      expect(hiddenFlags(doc.paragraphs[5].xml)).toEqual([true]); // body card
      expect(hiddenFlags(doc.paragraphs[6].xml)).toEqual([true, false]); // partial keep
      expect(hiddenFlags(doc.paragraphs[7].xml)).toEqual([true]); // numbered body card
      expect(hiddenFlags(doc.paragraphs[8].xml)).toEqual([false]); // table untouched
      expect(await isArmed(port)).toBe(true);
    }
  );

  it("preserves numbering markup through hide → showAll (byte-stable except w:vanish)", async () => {
    const doc = briefDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    await hide(port, settings(["yellow"]));
    const hidden = doc.paragraphs[7].xml;
    expect(hidden).toContain("w:numPr"); // numbering survives
    expect(hidden).toContain('w:numId w:val="3"'); // …with its id intact
    expect(hidden).toContain("<w:vanish/>"); // only addition is the hide flag

    await showAll(port);
    const shown = doc.paragraphs[7].xml;
    expect(shown).toContain("w:numPr"); // numbering still intact
    expect(shown).not.toContain("<w:vanish/>"); // fully reversed
  });

  it("Show All reveals everything and disarms; Re-hide catches newly typed cards", async () => {
    const doc = briefDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });

    await hide(port, settings(["yellow"]));
    await showAll(port);
    for (const p of doc.paragraphs) expect(hiddenFlags(p.xml).some((x) => x)).toBe(false);
    expect(await loadManifest(port)).toBeNull();

    // User types a fresh card after revealing, then Re-hides.
    doc.paragraphs.push(para(run("freshly pasted card body")));
    await reHide(port, settings(["yellow"]));
    expect(hiddenFlags(doc.paragraphs[doc.paragraphs.length - 1].xml)).toEqual([true]);
  });

  it("is idempotent — a second hide changes nothing", async () => {
    const doc = briefDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    await hide(port, settings(["yellow"]));
    const second = await hide(port, settings(["yellow"]));
    expect(second.paragraphsChanged).toBe(0);
  });

  it("honors the document manifest's keep-colors on the next session (round-trip)", async () => {
    const doc = briefDoc();
    const h = harness(doc);
    const port = createOfficeWordPort({ runner: h.runner, logger: h.tracer.logger("adapter") });
    // Arm with green-only keepers.
    await hide(port, settings(["green"]));
    // Simulate reopening: resolve settings from the persisted manifest.
    const manifest = await loadManifest(port);
    const resolved = resolveSettings(manifest, null);
    expect([...resolved.keepColors]).toEqual(["green"]);
    // The yellow-highlighted run is NOT a keeper under green-only settings.
    await reHide(port, resolved);
    expect(hiddenFlags(doc.paragraphs[6].xml)).toEqual([true, true]);
  });
});

// ---------------------------------------------------------------------------
// Perf: ~200 pages. A debate brief runs long; the adapter must stay responsive.
// ---------------------------------------------------------------------------
describe("perf — 200-page brief", () => {
  /** ~25 paragraphs/page × 200 pages. Mostly body cards, with periodic headings/cites. */
  function bigDoc(pages = 200, perPage = 25): FakePara[] {
    const out: FakePara[] = [];
    const total = pages * perPage;
    for (let i = 0; i < total; i++) {
      if (i % perPage === 0) out.push(para(run(`Block heading ${i}`), { outlineNumber: 3 }));
      else if (i % 7 === 0) out.push(para(run(`Author ${i}`, { cite: true })));
      else if (i % 5 === 0) out.push(para(run("intro ") + run("keep", { highlight: "yellow" })));
      else out.push(para(run(`Body card paragraph number ${i} with several words of warrant.`)));
    }
    return out;
  }

  it.each<CommitStrategy>(["per-paragraph", "whole-body"])(
    "hides a 5,000-paragraph brief within budget — %s",
    async (commitStrategy) => {
      const doc = mkDoc(bigDoc());
      const h = harness(doc);
      const port = createOfficeWordPort({
        runner: h.runner,
        commitStrategy,
        logger: h.tracer.logger("adapter")
      });

      const t0 = Date.now();
      const res = await hide(port, settings(["yellow"]));
      const ms = Date.now() - t0;
      // eslint-disable-next-line no-console
      console.log(
        `Rostrum perf [${commitStrategy}]: hid ${res.paragraphsChanged}/${res.paragraphsScanned} ` +
          `paragraphs in ${ms}ms`
      );

      expect(res.paragraphsScanned).toBe(5000);
      expect(res.paragraphsSkipped).toBe(0);
      // Generous CI-safe ceiling; the real number is logged above for the spike note.
      expect(ms).toBeLessThan(15000);
    },
    30000
  );
});
