import { classifyParagraph, hide, reHide, showAll } from "../src/core/invisibility";
import { makeAllVisible, readRuns } from "../src/core/ooxml";
import { parseManifestOrNull } from "../src/core/manifest";
import { TrackChangesActiveError } from "../src/core/guards";
import {
  ParagraphUpdate,
  RawParagraph,
  ResolvedSettings,
  TrackChangesMode,
  WordPort
} from "../src/core/types";

// ---------------------------------------------------------------------------
// Fixtures
// ---------------------------------------------------------------------------
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

function run(text: string, opts: { highlight?: string; cite?: boolean; vanish?: boolean } = {}): string {
  const rPr: string[] = [];
  if (opts.cite) rPr.push(`<w:rStyle w:val="Style13ptBold"/>`);
  if (opts.highlight) rPr.push(`<w:highlight w:val="${opts.highlight}"/>`);
  if (opts.vanish) rPr.push(`<w:vanish/>`);
  const rPrXml = rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";
  return `<w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

function para(inner: string): string {
  return `<w:p xmlns:w="${W_NS}">${inner}</w:p>`;
}

function mkPara(
  index: number,
  ooxml: string,
  opts: { headingLevel?: number | null; inTable?: boolean } = {}
): RawParagraph {
  return {
    index,
    headingLevel: opts.headingLevel ?? null,
    inTable: opts.inTable ?? false,
    ooxml
  };
}

function settings(colors: string[]): ResolvedSettings {
  return { keepColors: new Set(colors) };
}

/** In-memory WordPort: the whole engine runs against this with no Word host. */
class FakeWordPort implements WordPort {
  trackChangesMode: TrackChangesMode = "Off";
  manifestXml: string | null = null;
  readonly tcHistory: TrackChangesMode[] = [];

  constructor(public paragraphs: RawParagraph[]) {}

  async getChangeTrackingMode(): Promise<TrackChangesMode> {
    return this.trackChangesMode;
  }
  async setChangeTrackingMode(mode: TrackChangesMode): Promise<void> {
    this.trackChangesMode = mode;
    this.tcHistory.push(mode);
  }
  async readParagraphs(): Promise<RawParagraph[]> {
    return this.paragraphs.map((p) => ({ ...p }));
  }
  async writeParagraphs(updates: ParagraphUpdate[]): Promise<void> {
    for (const u of updates) {
      const p = this.paragraphs.find((x) => x.index === u.index);
      if (p) p.ooxml = u.ooxml;
    }
  }
  async readManifest(): Promise<string | null> {
    return this.manifestXml;
  }
  async writeManifest(xml: string): Promise<void> {
    this.manifestXml = xml;
  }
  async clearManifest(): Promise<void> {
    this.manifestXml = null;
  }
  async clearHidden(): Promise<{ paragraphsScanned: number; paragraphsChanged: number }> {
    // Mirrors the adapter's native clear: reveal everything (clear all <w:vanish/>),
    // counting only paragraphs that actually changed so convergence (a 2nd Show All
    // changes nothing) stays observable at the engine level.
    let changed = 0;
    for (const p of this.paragraphs) {
      const res = makeAllVisible(p.ooxml);
      if (res.changed) {
        p.ooxml = res.xml;
        changed++;
      }
    }
    return { paragraphsScanned: this.paragraphs.length, paragraphsChanged: changed };
  }
}

const hiddenFlags = (ooxml: string): boolean[] => readRuns(ooxml).map((r) => r.hidden);

// ---------------------------------------------------------------------------
// Pure classifier
// ---------------------------------------------------------------------------
describe("classifyParagraph", () => {
  const s = settings(["yellow"]);

  it("keeps a heading whole with no change", () => {
    const plan = classifyParagraph(mkPara(0, para(run("My Heading")), { headingLevel: 0 }), s);
    expect(plan.action).toBe("keepWhole");
    expect(plan.changed).toBe(false);
  });

  it("un-hides a heading that was previously hidden", () => {
    const plan = classifyParagraph(
      mkPara(0, para(run("My Heading", { vanish: true })), { headingLevel: 0 }),
      s
    );
    expect(plan.action).toBe("keepWhole");
    expect(plan.changed).toBe(true);
    expect(hiddenFlags(plan.ooxml)).toEqual([false]);
  });

  it("keeps a paragraph that contains a cite run", () => {
    const plan = classifyParagraph(mkPara(0, para(run("Smith 2020", { cite: true })), {}), s);
    expect(plan.action).toBe("keepWhole");
  });

  it("hides a whole body card and collapses its paragraph mark", () => {
    const plan = classifyParagraph(mkPara(0, para(run("card body text")), {}), s);
    expect(plan.action).toBe("hideWhole");
    expect(hiddenFlags(plan.ooxml)).toEqual([true]);
    expect(plan.ooxml).toMatch(/<w:pPr>[\s\S]*?<w:vanish\/>[\s\S]*?<\/w:pPr>/);
  });

  it("partially hides a body paragraph, keeping the highlighted word", () => {
    const plan = classifyParagraph(mkPara(0, para(run("plain ") + run("keep", { highlight: "yellow" })), {}), s);
    expect(plan.action).toBe("hidePartial");
    expect(hiddenFlags(plan.ooxml)).toEqual([true, false]);
    // partial paragraph keeps its mark visible
    expect(plan.ooxml).not.toMatch(/<w:pPr>[\s\S]*?<w:vanish\/>[\s\S]*?<\/w:pPr>/);
  });

  it("leaves table paragraphs untouched", () => {
    const plan = classifyParagraph(mkPara(0, para(run("in a cell")), { inTable: true }), s);
    expect(plan.action).toBe("keepWhole");
    expect(plan.changed).toBe(false);
  });

  it("leaves an empty separator paragraph alone", () => {
    const plan = classifyParagraph(mkPara(0, para(""), {}), s);
    expect(plan.changed).toBe(false);
  });

  it("keeps a paragraph with an inline image WHOLE (no OOXML write that would dangle the image part — Stage 4.2 P2)", () => {
    // A body paragraph with a `<w:drawing>` (inline image) plus hideable text would otherwise be
    // hidePartial → committed via insertOoxml with the image's internal r:embed dangling → host
    // rejection. It must classify keepWhole (native toggle path) and force the runs visible.
    const drawingRun = `<w:r><w:drawing/></w:r>`; // a `<w:drawing>` (its r:embed lives in a real doc)
    const plan = classifyParagraph(mkPara(0, para(drawingRun + run("hideable body text")), {}), s);
    expect(plan.action).toBe("keepWhole");
    expect(readRuns(plan.ooxml).every((r) => !r.hidden)).toBe(true); // nothing hidden
  });
});

// ---------------------------------------------------------------------------
// Orchestrators
// ---------------------------------------------------------------------------
describe("hide / showAll / reHide", () => {
  const s = settings(["yellow"]);

  function sampleDoc(): FakeWordPort {
    return new FakeWordPort([
      mkPara(0, para(run("Contention 1")), { headingLevel: 0 }),
      mkPara(1, para(run("a long card body that should vanish")), {}),
      mkPara(2, para(run("Author 2019", { cite: true })), {}),
      mkPara(3, para(run("intro ") + run("warrant", { highlight: "yellow" })), {})
    ]);
  }

  it("hides non-keepers, keeps heading/cite/highlight, and arms the manifest", async () => {
    const port = sampleDoc();
    const res = await hide(port, s);

    expect(hiddenFlags(port.paragraphs[0].ooxml)).toEqual([false]); // heading
    expect(hiddenFlags(port.paragraphs[1].ooxml)).toEqual([true]); // body card
    expect(hiddenFlags(port.paragraphs[2].ooxml)).toEqual([false]); // cite
    expect(hiddenFlags(port.paragraphs[3].ooxml)).toEqual([true, false]); // partial

    const manifest = parseManifestOrNull(port.manifestXml);
    expect(manifest).not.toBeNull();
    expect(manifest!.active).toBe(true);
    expect(manifest!.keepColors).toEqual(["yellow"]);
    expect(res.paragraphsChanged).toBe(2); // body card + partial paragraph
    expect(res.trackChangesToggled).toBe(false);
  });

  it("is idempotent — a second hide changes nothing", async () => {
    const port = sampleDoc();
    await hide(port, s);
    const second = await hide(port, s);
    expect(second.paragraphsChanged).toBe(0);
  });

  it("Show All reveals everything and disarms", async () => {
    const port = sampleDoc();
    await hide(port, s);
    await showAll(port);

    for (const p of port.paragraphs) {
      expect(hiddenFlags(p.ooxml).some((h) => h)).toBe(false);
    }
    expect(port.manifestXml).toBeNull();
    // Convergent: a second Show All is a no-op.
    expect((await showAll(port)).paragraphsChanged).toBe(0);
  });

  it("Re-hide catches newly typed text", async () => {
    const port = sampleDoc();
    await hide(port, s);
    // user types a new card after hiding
    port.paragraphs.push(mkPara(4, para(run("freshly typed card")), {}));
    await reHide(port, s);
    expect(hiddenFlags(port.paragraphs[4].ooxml)).toEqual([true]);
  });

  it("Show All also reveals a user's own manually hidden run (documented over-reveal #10)", async () => {
    const port = new FakeWordPort([mkPara(0, para(run("secret", { vanish: true })), { headingLevel: 0 })]);
    await showAll(port);
    expect(hiddenFlags(port.paragraphs[0].ooxml)).toEqual([false]);
  });

  it("skips a malformed paragraph without aborting the pass (M5)", async () => {
    const port = sampleDoc();
    port.paragraphs.push(mkPara(4, `<w:p xmlns:w="${W_NS}"><w:r><w:t>broken`, {})); // truncated OOXML
    const res = await hide(port, s);

    expect(res.paragraphsSkipped).toBe(1);
    expect(parseManifestOrNull(port.manifestXml)!.active).toBe(true); // pass completed + armed
    expect(hiddenFlags(port.paragraphs[1].ooxml)).toEqual([true]); // valid card still hidden
    expect(port.paragraphs[4].ooxml).toContain("broken"); // malformed para left as-is, not corrupted
  });
});

// ---------------------------------------------------------------------------
// Track-Changes gate (decision #14)
// ---------------------------------------------------------------------------
describe("Track Changes gate", () => {
  const s = settings(["yellow"]);
  const doc = (): FakeWordPort => new FakeWordPort([mkPara(0, para(run("body")), {})]);

  it("refuses to hide while Track Changes is on (no auto-toggle)", async () => {
    const port = doc();
    port.trackChangesMode = "TrackAll";
    await expect(hide(port, s)).rejects.toBeInstanceOf(TrackChangesActiveError);
    expect(port.manifestXml).toBeNull(); // nothing written
    expect(hiddenFlags(port.paragraphs[0].ooxml)).toEqual([false]);
  });

  it("auto-toggles Track Changes off and restores it afterward", async () => {
    const port = doc();
    port.trackChangesMode = "TrackAll";
    const res = await hide(port, s, { autoToggleTrackChanges: true });

    expect(res.trackChangesToggled).toBe(true);
    expect(port.trackChangesMode).toBe("TrackAll"); // restored
    expect(port.tcHistory).toEqual(["Off", "TrackAll"]);
    expect(parseManifestOrNull(port.manifestXml)!.active).toBe(true);
    expect(hiddenFlags(port.paragraphs[0].ooxml)).toEqual([true]);
  });
});
