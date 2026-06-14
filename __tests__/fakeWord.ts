// Shared test harness: a fake Office.js `RequestContext` that models the essential
// semantics — reads resolve and mutations apply ONLY on `sync()`, in call order —
// so adapter and integration tests can assert single-sync atomic commits, manifest
// set-vs-add, the alignment fallback, etc., with NO Word host. Not a `*.test.ts`,
// so jest never collects it as a suite; it is plain shared code imported by both
// officeWordPort.test.ts and integration.test.ts (DRY: one fake, not two).

import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { applyRunVisibility, makeAllVisible, readRuns } from "../src/core/ooxml";
import { Tracer } from "../src/core/debug";
import { ResolvedSettings, TrackChangesMode } from "../src/core/types";

export const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
export const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

// ---------------------------------------------------------------------------
// Backing document model + builders
// ---------------------------------------------------------------------------

export interface FakePara {
  /** Standalone, parseable `<w:p xmlns:w=…>…</w:p>`. */
  xml: string;
  /** Raw numeric outline level (as `Paragraph.outlineLevel` would report). */
  outlineNumber: number;
  inTable: boolean;
}

export interface ParaOpts {
  outlineNumber?: number;
  inTable?: boolean;
  /** Extra raw `<w:pPr>` markup (e.g. numbering) to verify byte-preservation. */
  pPr?: string;
}

/** Build a backing paragraph from inner run markup. */
export function para(inner: string, opts: ParaOpts = {}): FakePara {
  const pPr = opts.pPr ?? "";
  const num = opts.outlineNumber ?? 10; // VBA body-text default (≥10 = body text)
  return {
    xml: `<w:p xmlns:w="${W_NS}">${pPr}${inner}</w:p>`,
    outlineNumber: num,
    inTable: opts.inTable ?? false
  };
}

/** Build one `<w:r>` with optional highlight / cite style. */
export function run(text: string, o: { highlight?: string; cite?: boolean } = {}): string {
  const rPr: string[] = [];
  if (o.cite) rPr.push(`<w:rStyle w:val="Style13ptBold"/>`);
  if (o.highlight) rPr.push(`<w:highlight w:val="${o.highlight}"/>`);
  const rPrXml = rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";
  return `<w:r>${rPrXml}<w:t xml:space="preserve">${text}</w:t></w:r>`;
}

export interface FakeDoc {
  paragraphs: FakePara[];
  manifest: { id: string; xml: string } | null;
  tcMode: TrackChangesMode;
  nextId: number;
}

export function mkDoc(paragraphs: FakePara[], tc: TrackChangesMode = "Off"): FakeDoc {
  return { paragraphs, manifest: null, tcMode: tc, nextId: 1 };
}

/** A flat-OPC whole-body package built from the backing paragraphs. */
export function buildPackage(paras: FakePara[]): string {
  return (
    `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
    `<w:document xmlns:w="${W_NS}"><w:body>${paras.map((p) => p.xml).join("")}` +
    `<w:sectPr><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`
  );
}

// ---------------------------------------------------------------------------
// Fake RequestContext
// ---------------------------------------------------------------------------

interface CommitEntry {
  sync: number;
  op: string;
  index?: number;
}

class FakeContext {
  private queue: Array<() => void> = [];
  syncNo = 0;
  readonly commitLog: CommitEntry[] = [];
  /** When set, the NEXT sync throws this and applies nothing (failure injection). */
  failNext: unknown = null;
  readonly document: FakeDocument;
  readonly trackedObjects = { add: (): void => undefined, remove: (): void => undefined };

  constructor(
    public readonly doc: FakeDoc,
    private readonly bodyOoxmlOverride: string | null
  ) {
    this.document = new FakeDocument(this);
  }

  enqueue(fn: () => void): void {
    this.queue.push(fn);
  }

  log(op: string, index?: number): void {
    this.commitLog.push({ sync: this.syncNo, op, index });
  }

  bodyOoxml(): string {
    return this.bodyOoxmlOverride ?? buildPackage(this.doc.paragraphs);
  }

  async sync(): Promise<void> {
    this.syncNo++;
    if (this.failNext) {
      const e = this.failNext;
      this.failNext = null;
      this.queue = [];
      throw e;
    }
    const q = this.queue;
    this.queue = [];
    for (const fn of q) fn();
  }
}

/**
 * A fake `Word.Font` modeling only the `hidden` toggle the adapter sets natively
 * (Stage 4). Applying it produces the SAME effect Word would: `hidden=true` vanishes
 * every run AND the paragraph mark (collapse); `hidden=false` clears all vanish
 * (makeAllVisible). `index === null` targets the whole body (every paragraph); a number
 * targets a single paragraph. Effects apply on `sync()`, like every other fake mutation.
 */
function makeFakeFont(ctx: FakeContext, index: number | null): { hidden: boolean } {
  let state = false;
  return {
    get hidden(): boolean {
      return state;
    },
    set hidden(v: boolean) {
      state = v;
      ctx.enqueue(() => {
        const targets = index === null ? ctx.doc.paragraphs.map((_, i) => i) : [index];
        for (const i of targets) {
          const p = ctx.doc.paragraphs[i];
          p.xml = v
            ? applyRunVisibility(p.xml, readRuns(p.xml).map(() => true), true).xml
            : makeAllVisible(p.xml).xml;
        }
        ctx.log(index === null ? "body.font.hidden" : "paragraph.font.hidden", index ?? undefined);
      });
    }
  };
}

class FakeDocument {
  readonly body: FakeBody;
  readonly customXmlParts: FakeXmlParts;
  constructor(private readonly ctx: FakeContext) {
    this.body = new FakeBody(ctx);
    this.customXmlParts = new FakeXmlParts(ctx);
  }
  load(): void {
    /* values are read live post-sync */
  }
  get changeTrackingMode(): TrackChangesMode {
    return this.ctx.doc.tcMode;
  }
  set changeTrackingMode(mode: TrackChangesMode) {
    this.ctx.enqueue(() => {
      this.ctx.doc.tcMode = mode;
      this.ctx.log(`tc.set:${mode}`);
    });
  }
}

class FakeBody {
  readonly paragraphs: FakeParaCollection;
  /** Native `Word.Body.font` — only `hidden` is modeled (the fast whole-body Show All). */
  readonly font: { hidden: boolean };
  constructor(private readonly ctx: FakeContext) {
    this.paragraphs = new FakeParaCollection(ctx);
    this.font = makeFakeFont(ctx, null);
  }
  getOoxml(): { value: string } {
    const result = { value: "" };
    // Log the whole-body read so tests can count read syncs (e.g. the Loop 002 B2 read-fusion proof:
    // a clean Hide reads ONCE; the auto-toggle path discards the primed read and reads a SECOND time).
    this.ctx.enqueue(() => {
      result.value = this.ctx.bodyOoxml();
      this.ctx.log("body.getOoxml");
    });
    return result;
  }
  insertOoxml(xml: string): void {
    this.ctx.enqueue(() => {
      const pkg = new WholeBodyPackage(xml);
      for (let i = 0; i < pkg.count && i < this.ctx.doc.paragraphs.length; i++) {
        this.ctx.doc.paragraphs[i].xml = pkg.paragraphXml(i);
      }
      this.ctx.log("body.insertOoxml");
    });
  }
}

class FakeParaCollection {
  private cached: FakeParagraph[] | null = null;
  constructor(private readonly ctx: FakeContext) {}
  load(): void {
    /* items read live */
  }
  get items(): FakeParagraph[] {
    // Rebuild when the document's paragraph count changes — real `body.paragraphs`
    // re-enumerates each run, so a paragraph typed between two operations must
    // appear on the next read (the Re-hide-catches-new-text path).
    if (!this.cached || this.cached.length !== this.ctx.doc.paragraphs.length) {
      this.cached = this.ctx.doc.paragraphs.map((_, i) => new FakeParagraph(this.ctx, i));
    }
    return this.cached;
  }
}

// Mirrors the REAL `Word.Paragraph` surface the adapter touches. It deliberately has
// NO `paragraphFormat`: that property exists only on `Word.Style`, never on a
// paragraph, and modeling it here once hid a live-host crash (`undefined (reading
// 'load')`). Outline level comes only from the numeric `outlineLevel` getter. (LESSONS.)
class FakeParagraph {
  readonly parentTableOrNullObject: { load(): void; isNullObject: boolean };
  /** Native `Word.Paragraph.font` — models `hidden` (the safe-mode per-paragraph apply). */
  readonly font: { hidden: boolean };
  constructor(
    private readonly ctx: FakeContext,
    private readonly index: number
  ) {
    const back = (): FakePara => this.ctx.doc.paragraphs[this.index];
    this.parentTableOrNullObject = {
      load: () => undefined,
      get isNullObject() {
        return !back().inTable;
      }
    };
    this.font = makeFakeFont(ctx, index);
  }
  load(): void {
    /* live */
  }
  get outlineLevel(): number {
    return this.ctx.doc.paragraphs[this.index].outlineNumber;
  }
  // `Word.Paragraph.text` — concatenated run text, used by the adapter's whole-body
  // alignment (Stage 4.1) to match package paragraphs to proxies.
  get text(): string {
    return readRuns(this.ctx.doc.paragraphs[this.index].xml)
      .map((r) => r.text)
      .join("");
  }
  // `Word.Paragraph.getOoxml()` (WordApi 1.1) — the OOXML of THIS paragraph (one body
  // `<w:p>`). The adapter reads through this; engine fixtures are bare `<w:p>`, so we
  // hand back the backing xml directly. (The real host wraps it in a flat-OPC package;
  // that package shape is exercised separately by the real-document tests.)
  getOoxml(): { value: string } {
    const result = { value: "" };
    this.ctx.enqueue(() => (result.value = this.ctx.doc.paragraphs[this.index].xml));
    return result;
  }
  // `Word.Paragraph.insertOoxml(ooxml, "Replace")` (WordApi 1.1) — replace this para.
  insertOoxml(xml: string): void {
    this.ctx.enqueue(() => {
      this.ctx.doc.paragraphs[this.index].xml = xml;
      this.ctx.log("paragraph.insertOoxml", this.index);
    });
  }
  getRange(): FakeRange {
    return new FakeRange(this.ctx, this.index);
  }
}

class FakeRange {
  constructor(
    private readonly ctx: FakeContext,
    private readonly index: number
  ) {}
  getOoxml(): { value: string } {
    // REGRESSION MODEL: the real `Paragraph.getRange()` defaults to the "Whole"
    // range, which INCLUDES the paragraph mark, so `Range.getOoxml()` serializes the
    // paragraph PLUS a trailing empty `<w:p>` — TWO body paragraphs in a flat-OPC
    // package. The adapter must read via `Paragraph.getOoxml()` instead; if it ever
    // reverts to `getRange().getOoxml()`, the hide path yields a 2-`<w:p>` fragment
    // and the single-paragraph write guard throws — the exact live-host bug.
    const result = { value: "" };
    this.ctx.enqueue(
      () =>
        (result.value =
          `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml">` +
          `<pkg:xmlData><w:document xmlns:w="${W_NS}"><w:body>` +
          `${this.ctx.doc.paragraphs[this.index].xml}<w:p/>` +
          `</w:body></w:document></pkg:xmlData></pkg:part></pkg:package>`)
    );
    return result;
  }
  insertOoxml(xml: string): void {
    this.ctx.enqueue(() => {
      this.ctx.doc.paragraphs[this.index].xml = xml;
      this.ctx.log("range.insertOoxml", this.index);
    });
  }
}

class FakeXmlParts {
  constructor(private readonly ctx: FakeContext) {}
  getByNamespace(ns: string): FakeScoped {
    return new FakeScoped(this.ctx, ns);
  }
  add(xml: string): FakePart {
    this.ctx.enqueue(() => {
      this.ctx.doc.manifest = { id: `part-${this.ctx.doc.nextId++}`, xml };
      this.ctx.log("xmlParts.add");
    });
    return new FakePart(this.ctx);
  }
}

class FakeScoped {
  constructor(
    private readonly ctx: FakeContext,
    readonly ns: string
  ) {}
  getOnlyItemOrNullObject(): FakePart {
    return new FakePart(this.ctx);
  }
}

class FakePart {
  constructor(private readonly ctx: FakeContext) {}
  load(): void {
    /* live */
  }
  get isNullObject(): boolean {
    return this.ctx.doc.manifest === null;
  }
  get id(): string {
    return this.ctx.doc.manifest?.id ?? "";
  }
  getXml(): { value: string } {
    const result = { value: "" };
    this.ctx.enqueue(() => (result.value = this.ctx.doc.manifest?.xml ?? ""));
    return result;
  }
  setXml(xml: string): void {
    this.ctx.enqueue(() => {
      if (this.ctx.doc.manifest) this.ctx.doc.manifest.xml = xml;
      this.ctx.log("part.setXml");
    });
  }
  delete(): void {
    this.ctx.enqueue(() => {
      this.ctx.doc.manifest = null;
      this.ctx.log("part.delete");
    });
  }
}

// ---------------------------------------------------------------------------
// Harness + small assertion helpers
// ---------------------------------------------------------------------------

export interface Harness {
  ctx: FakeContext;
  doc: FakeDoc;
  runner: <T>(b: (c: Word.RequestContext) => Promise<T>) => Promise<T>;
  tracer: Tracer;
  warnings: string[];
}

/** Wire a fake context, a silent tracer, and a runner the adapter can be given. */
export function harness(doc: FakeDoc, bodyOoxmlOverride: string | null = null): Harness {
  const ctx = new FakeContext(doc, bodyOoxmlOverride);
  const tracer = new Tracer({ console: null });
  const warnings: string[] = [];
  tracer.subscribe((e) => {
    if (e.level === "warn") warnings.push(e.msg);
  });
  const runner = <T,>(b: (c: Word.RequestContext) => Promise<T>): Promise<T> =>
    b(ctx as unknown as Word.RequestContext);
  return { ctx, doc, runner, tracer, warnings };
}

export const settings = (colors: string[]): ResolvedSettings => ({ keepColors: new Set(colors) });
export const hiddenFlags = (xml: string): boolean[] => readRuns(xml).map((r) => r.hidden);
