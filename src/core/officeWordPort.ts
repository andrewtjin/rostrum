// The real WordPort: the ONE place Office.js touches the Rostrum engine.
//
// Stage 1 proved the whole policy against an in-memory fake port. This adapter is
// the production implementation of that same 7-method `WordPort` contract over
// `Word.run`. Everything hard about running inside a live host lives here, and is
// engineered to be (a) correct under the verified Office.js semantics, (b) testable
// with NO Word host (via an injectable runner + the pure helpers in
// ooxmlPackage.ts), and (c) ruthlessly observable (every host round-trip is a timed,
// correlated tracer span; every failure is expanded into its OfficeExtension code +
// debugInfo). See README "Debugging" and core/debug.ts.
//
// KEY DESIGN DECISIONS (and why):
//
//  * STRING-LITERAL API ARGS, NEVER `Word.*` ENUM VALUES. We pass "Replace" /
//    "Off" etc. as string literals (the API accepts them) instead of
//    `Word.InsertLocation.replace`. Referencing a `Word.*` runtime enum would throw
//    in Node, where this module is unit-tested under a fake runner. Only `Word.run`
//    itself is referenced (in the default runner, never called in tests).
//
//  * INJECTABLE RUNNER. `options.runner` defaults to `Word.run`, but tests pass a
//    fake `RequestContext`, so the adapter's sequencing (atomic ordering, manifest
//    set-vs-add, TC restore) is asserted deterministically without office-addin-mock.
//
//  * COMMIT STRATEGY, SPIKE-GATED. The plan's perf path is a whole-document OOXML
//    round-trip (one `insertOoxml`), but its correctness depends on a fidelity spike
//    that needs a live host (xmldom re-serialization acceptable to Word; `<w:p>`↔
//    proxy index alignment). Since that spike can't run host-free, the SAFE
//    `"per-paragraph"` strategy is the DEFAULT (alignment is exact — each paragraph
//    is read and written through its own range), and `"whole-body"` is OPT-IN with a
//    structural ALIGNMENT GUARD that auto-falls-back to per-paragraph on any count
//    mismatch. Flipping the default after a successful spike is a one-line change.
//
//  * SINGLE-SYNC ATOMIC COMMIT (audit H3). The engine calls `writeParagraphs` then a
//    manifest op (`writeManifest`/`clearManifest`) as two WordPort calls. To make the
//    document edit and the manifest land in ONE `context.sync()` — so a failure can
//    never leave the doc "hidden-but-unarmed" — `writeParagraphs` BUFFERS its updates
//    and the following manifest op FLUSHES them together. This buffering is explicit
//    and logged, not hidden.
//
//  * NO PROXIES CROSS RUNS. Rather than tracking Office proxy objects across syncs
//    (the `InvalidObjectPath` trap), each `Word.run` re-fetches what it needs. The
//    whole-body path carries state between read and commit as a PURE JS
//    `WholeBodyPackage`, not a proxy. `context.trackedObjects` is therefore unneeded.

import {
  ParagraphUpdate,
  RangeRead,
  RangeScopedPort,
  RawParagraph,
  TrackChangesMode,
  WordPort
} from "./types";
import {
  WholeBodyPackage,
  assertSingleParagraph,
  countBodyParagraphs,
  keepFirstBodyParagraph,
  normalizeOutlineNumber,
  OutlineNumberBase
} from "./ooxmlPackage";
import {
  applyCiteStyleToParagraphXml,
  CiteRepairParagraph,
  planCiteRepairs
} from "./citeRepair";
import { CancelledError, CancelToken, createPacer, Pacer } from "./cancel";
import { CITE_STYLE_ID } from "./styles";
import { MANIFEST_NAMESPACE } from "./manifest";
import { Logger, describeError, logger as rootLogger } from "./debug";

/* eslint-disable @typescript-eslint/no-explicit-any */
// Office proxy objects are richly typed by @types/office-js, but the per-method
// surface we touch is small; a localized `any` for context handles keeps the fake
// runner (tests) and the real host on one path without fighting generics.

// ---------------------------------------------------------------------------
// Public configuration surface
// ---------------------------------------------------------------------------

/** Which write mechanism the adapter commits with (see file header). */
export type CommitStrategy = "per-paragraph" | "whole-body";

/**
 * The single seam over `Word.run`. Matches its primary overload. Tests inject a
 * fake that runs the batch against an in-memory document and resolves ClientResults
 * on `sync()`.
 */
export type WordRunner = <T>(batch: (context: Word.RequestContext) => Promise<T>) => Promise<T>;

/** How to interpret a paragraph's outline level (see ooxmlPackage.ts). */
export interface OutlineConfig {
  /**
   * Base for the numeric `Word.Paragraph.outlineLevel` (WordApi 1.1) — the ONLY
   * outline API a paragraph exposes. `"oneBased"` = VBA parity (Heading 1 = 1 …
   * body text = 10). There is no paragraph-level *enum* form: `paragraphFormat`
   * exists on `Word.Style`, never on `Word.Paragraph`, so reading the enum off a
   * paragraph is impossible (it threw `undefined (reading 'load')` on every real
   * host) and was removed. See LESSONS.
   */
  numberBase: OutlineNumberBase;
}

/** Read/commit progress, surfaced to the task pane for the progress bar + cancel. */
export interface ProgressInfo {
  phase: "read" | "commit";
  done: number;
  total: number;
}

/** Outcome of a cite-repair pass: how many mis-styled cites were restyled. */
export interface CiteRepairResult {
  /** Number of paragraphs that received the cite style (one per repaired cite). */
  paragraphsRepaired: number;
  /** Total number of runs that received the cite rStyle across all repaired paragraphs. */
  runsRepaired: number;
}

/**
 * A WordPort that can ALSO repair mis-styled cites (re-apply the cite character style to
 * cites that lost it, so the keeper rule keeps them). Kept separate from `WordPort` (the
 * pure-engine contract) because cite-repair is an Apply-Styles concern, not part of the
 * hide/show engine. The factory's return type includes this so the controller can call it.
 */
export interface CiteRepairCapablePort extends WordPort {
  /**
   * Detect and repair mis-styled cites across the body via the whole-body OOXML path
   * (one `getOoxml` → splice → one `insertOoxml("Replace")`). Returns zeros when there is
   * nothing to repair, and is resilient: it never throws for an empty/unrepairable doc.
   */
  repairCites(): Promise<CiteRepairResult>;
}

// `CancelToken`/`CancelledError` moved to core/cancel.ts (so the host-agnostic engine can
// pace/cancel without importing this adapter) and are re-exported here verbatim — every
// existing import site (controller.ts, tests) keeps compiling against the SAME class
// identity, which the controller's `instanceof CancelledError` mapping depends on.
export { CancelledError } from "./cancel";
export type { CancelToken, Pacer } from "./cancel";

export interface OfficeWordPortOptions {
  /** Defaults to the global `Word.run`. Inject a fake in tests. */
  runner?: WordRunner;
  /** Defaults to the SAFE `"per-paragraph"`. See file header for the rationale. */
  commitStrategy?: CommitStrategy;
  /**
   * EXPERIMENTAL (avenue ⑦), default false. The PURE whole-body path: read the whole story in ONE
   * `body.getOoxml()`, classify EACH package paragraph directly (outline from the package's own
   * styles.xml, table membership from its structure) with NO Word proxies and NO text alignment,
   * then commit with ONE `body.insertOoxml(…, "Replace")`. This removes the proxy storm, the
   * targeted re-reads, and the 99× per-paragraph commit, AND gives perfect fidelity by construction
   * (the full original package round-trips, only `<w:vanish/>` toggled). Implies a whole-body commit.
   * It is a DESTRUCTIVE whole-body Replace whose structural fidelity (numbering/sections/fields/
   * bookmarks) is only confirmable on a live host, so it stays opt-in until that wet-test passes.
   */
  pureWholeBody?: boolean;
  /** Defaults to `{ mode: "number", numberBase: "oneBased" }`. */
  outline?: Partial<OutlineConfig>;
  /** Paragraphs loaded per `sync()` during the read phase. Default 200. */
  chunkSize?: number;
  /** Progress callback (read chunks + commit). */
  onProgress?: (info: ProgressInfo) => void;
  /** Cancellation token, polled between read chunks (and per paragraph in the pure read). */
  cancel?: CancelToken;
  /**
   * Paces the PURE whole-body read's package-assembly loop (avenue ⑦): `tick()` is awaited
   * once per paragraph, throwing `CancelledError` on cancellation and yielding a macrotask
   * when its budget elapses — so the task pane can paint progress and a Cancel click can
   * actually land. The loop is otherwise synchronous JS after the one read sync, so a bare
   * `cancel` poll could never observe a click made after that sync resolved. When absent, a
   * never-yielding pacer over `cancel` preserves the previous semantics exactly.
   */
  pacer?: Pacer;
  /** Logger to use; defaults to the shared tracer's "adapter" namespace. */
  logger?: Logger;
}

// ---------------------------------------------------------------------------
// The default runner (only path that names the `Word.run` global)
// ---------------------------------------------------------------------------

/**
 * The production runner: the global `Word.run`. Shared by the adapter, styling, and
 * live mode so the "no host detected" guard lives in exactly one place. `typeof Word`
 * is safe even when office.js isn't loaded (returns "undefined" without a
 * ReferenceError); the `||` short-circuits before touching `Word.run`.
 */
export const defaultWordRunner: WordRunner = (batch) => {
  if (typeof Word === "undefined" || typeof (Word as any).run !== "function") {
    throw new Error(
      "Rostrum: no Word host detected and no `runner` was injected. This must run " +
        "inside Word (office.js loaded) or be given a runner in tests."
    );
  }
  return (Word as any).run(batch);
};

const defaultRunner = defaultWordRunner;

/**
 * The placeholder `ooxml` on a node-direct `RawParagraph` (Loop 002 B1). The engine reads such a
 * paragraph through its `.parsed` node handle, never `.ooxml`, and the node-direct commit serializes
 * the WHOLE cached package (the `<w:p>` was mutated in place) — so a per-paragraph string is both
 * unused and a cost we deliberately drop. A bare valid one-`<w:p>` fragment keeps the type total and
 * any defensive read well-formed.
 */
const NODE_BACKED_OOXML =
  '<w:p xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"/>';

// ---------------------------------------------------------------------------
// Internal read state shared between the read and the (buffered) commit
// ---------------------------------------------------------------------------

interface ReadState {
  /** Which writeback mechanism the last read prepared for. */
  strategy: "range" | "pkg";
  /** Cached whole-body package to splice into (only when strategy === "pkg"). */
  pkg: WholeBodyPackage | null;
  /**
   * True for the LOOP 002 node-direct read (`readBatchPure` handing node-backed paragraphs): the
   * engine mutated this package's `<w:p>` nodes IN PLACE during Phase B, so the commit must NOT
   * re-`replace()` — it serializes the WHOLE package once (the mutations are already spliced by
   * reference) and the per-paragraph `ParagraphUpdate.ooxml` is a placeholder, never consulted. A
   * non-node-direct `pkg` read still re-splices via `replace()` (the proxy/compat path).
   */
  nodeDirect: boolean;
  /** Paragraph count the read observed, used to range-check buffered updates. */
  count: number;
  /**
   * engineIndex (= Office proxy index) → package story-paragraph index, from the
   * tolerant alignment (Stage 4.1 ⑧, non-cascading in 4.2). `-1` marks a proxy the
   * whole-body read couldn't align (re-read via a targeted getOoxml — see readBatch).
   * Consumed by the whole-body commit ONLY when `cleanAlign` is true (identity mapping),
   * so the `-1` sentinels are never dereferenced there.
   */
  pkgIndex: number[] | null;
  /**
   * True only when the package aligned 1:1 with the proxies (count equal AND identity
   * mapping — no artifact skipped). The whole-body COMMIT (`insertOoxml("Replace")`,
   * which rewrites the WHOLE body) is gated on this: if any package `<w:p>` was skipped,
   * re-serializing the whole package would re-inject that artifact, so perf mode instead
   * commits per-paragraph (surgical, can't inject/duplicate). The whole-body READ is
   * unaffected — it benefits even under tolerant (non-clean) alignment.
   */
  cleanAlign: boolean;
}

/** Manifest commit intent passed into the atomic `commit`. */
type ManifestAction = { kind: "set"; xml: string } | { kind: "delete" } | { kind: "none" };

// ---------------------------------------------------------------------------
// The adapter
// ---------------------------------------------------------------------------

class OfficeWordPort implements CiteRepairCapablePort, RangeScopedPort {
  private readonly run: WordRunner;
  private readonly strategy: CommitStrategy;
  /** Avenue ⑦ — read+commit the whole body with no proxies/alignment (the default Hide path). */
  private readonly pureWholeBody: boolean;
  private readonly outline: OutlineConfig;
  private readonly chunkSize: number;
  private readonly onProgress?: (info: ProgressInfo) => void;
  private readonly cancel?: CancelToken;
  /** Ticked per paragraph in the pure read; defaults to a never-yielding wrap of `cancel`. */
  private readonly pacer: Pacer;
  private readonly log: Logger;

  /** Updates buffered by `writeParagraphs`, flushed atomically by the next manifest op. */
  private pending: ParagraphUpdate[] | null = null;
  /** State from the most recent `readParagraphs`, consumed by `commit`. */
  private lastRead: ReadState | null = null;
  /**
   * The Track-Changes mode observed in the FIRST sync of the most recent `readParagraphs` (Loop 002 B2
   * / 002-S4 read fusion). Loaded alongside the body OOXML on that same sync, so a clean Hide spends NO
   * extra Word.run reading TC — the engine's prefetched-mode gate consumes this instead. `null` until a
   * read primes it, and re-nulled at the START of every read so a stale value from a prior op can never
   * leak into a later gate decision.
   */
  private lastReadTrackChanges: TrackChangesMode | null = null;
  /** Last observed manifest part id — advisory/diagnostic only (lookups go by namespace). */
  private manifestPartId: string | null = null;

  constructor(options: OfficeWordPortOptions = {}) {
    this.run = options.runner ?? defaultRunner;
    this.pureWholeBody = options.pureWholeBody ?? false;
    // Avenue ⑦ commits with ONE whole-body insertOoxml, so it implies the whole-body strategy
    // regardless of any commitStrategy passed — keeping read and commit consistent.
    this.strategy = this.pureWholeBody ? "whole-body" : options.commitStrategy ?? "per-paragraph";
    this.outline = {
      numberBase: options.outline?.numberBase ?? "oneBased"
    };
    this.chunkSize = Math.max(1, options.chunkSize ?? 200);
    this.onProgress = options.onProgress;
    this.cancel = options.cancel;
    // The pure-read loop always ticks a pacer. Without an injected one, fall back to a
    // never-yielding (Infinity-budget) pacer over the plain cancel token — byte-for-byte the
    // old bare `isCancelled()` poll, so cancel-only callers keep their exact semantics.
    this.pacer =
      options.pacer ?? createPacer({ cancel: options.cancel, budgetMs: Number.POSITIVE_INFINITY });
    this.log = options.logger ?? rootLogger("adapter");
    this.log.debug("officeWordPort created", {
      strategy: this.strategy,
      pureWholeBody: this.pureWholeBody,
      outline: this.outline,
      chunkSize: this.chunkSize
    });
  }

  // ----- Track Changes ------------------------------------------------------

  async getChangeTrackingMode(): Promise<TrackChangesMode> {
    const op = this.log.child("getTC");
    const span = op.span("getChangeTrackingMode");
    try {
      const mode = await this.run(async (ctx) => {
        const doc: any = ctx.document;
        doc.load("changeTrackingMode");
        await ctx.sync();
        return String(doc.changeTrackingMode) as TrackChangesMode;
      });
      span.end({ mode });
      return normalizeTcMode(mode, op);
    } catch (e) {
      span.fail(e);
      throw clarify(e, "read the Track Changes mode");
    }
  }

  async setChangeTrackingMode(mode: TrackChangesMode): Promise<void> {
    const op = this.log.child("setTC");
    const span = op.span("setChangeTrackingMode", { mode });
    try {
      await this.run(async (ctx) => {
        (ctx.document as any).changeTrackingMode = mode;
        await ctx.sync();
      });
      span.end();
    } catch (e) {
      span.fail(e);
      // TS#2 hardening: name the target mode so a failed RESTORE (the dangerous
      // case — the engine's finally block trying to turn TC back on) produces a
      // clear, actionable error instead of a silent propagation.
      throw new Error(
        `Rostrum could not set Track Changes to "${mode}". The document's Track Changes ` +
          `state may now be inconsistent — check Review ▸ Track Changes. Cause: ` +
          `${summarize(e)}`
      );
    }
  }

  // ----- Read ---------------------------------------------------------------

  /**
   * The Track-Changes mode prefetched by the most recent `readParagraphs` (Loop 002 B2 / 002-S4), or
   * `null` if no read has happened since construction. The Hide engine consumes this through the
   * prefetched-mode TC gate so a clean Hide pays NO separate Word.run for the TC read — the read sync
   * already carried it. A `null` here means "no primed value"; the engine then falls back to the
   * standard gate (its own TC read), so correctness never depends on the prime being present.
   */
  getPrimedTrackChangesMode(): TrackChangesMode | null {
    return this.lastReadTrackChanges;
  }

  async readParagraphs(): Promise<RawParagraph[]> {
    const op = this.log.child("read");
    const span = op.span("readParagraphs", { strategy: this.strategy, pure: this.pureWholeBody });
    // Null any TC mode primed by a prior read BEFORE this read begins: a stale value must never leak
    // into a later gate decision. The read paths below re-prime it inside their first sync.
    this.lastReadTrackChanges = null;
    try {
      const result = await this.run(async (ctx) =>
        this.pureWholeBody ? this.readBatchPure(ctx, op) : this.readBatch(ctx, op)
      );
      this.lastRead = result.state;
      this.lastReadTrackChanges = result.trackChanges;
      span.end({
        count: result.paragraphs.length,
        source: result.state.strategy,
        trackChanges: result.trackChanges,
        outlineRawSample: result.outlineSample
      });
      return result.paragraphs;
    } catch (e) {
      if (e instanceof CancelledError) {
        span.end({ cancelled: true });
        throw e;
      }
      span.fail(e);
      throw clarify(e, "read the document's paragraphs");
    }
  }

  /**
   * The PURE whole-body read (avenue ⑦). ONE `body.getOoxml()`, then classify EVERY package
   * story paragraph DIRECTLY from the package — NO `body.paragraphs` proxy collection, NO outline/
   * text/table proxy loads, NO `alignToProxies`, NO targeted re-reads. Outline level comes from the
   * package's own styles.xml (`pkg.headingLevel`, resolving the `basedOn` cascade — the host's
   * `canGetStyles:false` only blocks the Paragraph styles API, not the bundled styles part); table
   * membership comes from the paragraph's own ancestry (`pkg.inTable`). The package is
   * self-consistent (read N `<w:p>`, the engine edits only run-level `<w:vanish/>` in this same DOM,
   * write N), so the read produces an IDENTITY mapping with `cleanAlign:true` — the commit then takes
   * the single-`insertOoxml` whole-body path with perfect fidelity by construction (the whole
   * original package round-trips). The ±1 serialization artifact, duplicate cards, and the emoji —
   * all the things that made `alignToProxies` fall back — are MOOT here: there is no alignment, each
   * `<w:p>` is just classified and written back as itself.
   *
   * Resilience: if the package fails to parse (a pathological doc), fall back to the robust proxy
   * `readBatch` so pure mode is never WORSE than the default path on an unparseable document.
   */
  private async readBatchPure(
    ctx: any,
    op: Logger
  ): Promise<{ paragraphs: RawParagraph[]; state: ReadState; outlineSample: unknown; trackChanges: TrackChangesMode }> {
    const doc: any = ctx.document;
    const body = doc.body;
    const bodyOoxml = body.getOoxml();
    // Loop 002 B2 (002-S4) — PREFETCH the Track-Changes mode in this same first sync, fusing what was a
    // dedicated TC-read Word.run into the read. The engine's prefetched-mode gate consumes it, so a
    // clean Hide costs 2 runs (read + commit) instead of 3 (TC read + read + commit).
    doc.load("changeTrackingMode");
    await ctx.sync(); // the ONLY read sync: the whole-body OOXML + the TC mode
    const trackChanges = normalizeTcMode(String(doc.changeTrackingMode), op);

    let pkg: WholeBodyPackage;
    try {
      pkg = new WholeBodyPackage(bodyOoxml.value);
    } catch (e) {
      op.warn("pure whole-body parse failed — falling back to the proxy read path", describeError(e));
      // The fallback re-reads the body via the proxy path but has ALREADY primed TC in this sync; pass
      // it through so the fallback's own (second) TC load is unnecessary — preserve the prime.
      return { ...(await this.readBatch(ctx, op)), trackChanges };
    }

    // Review C-3: the package has a styles part but it parsed to ZERO style defs (an unreadable
    // serialization shape). Outline resolution would then mis-classify every style-derived heading
    // as body and the destructive whole-body Replace would HIDE them. Don't trust pure mode here —
    // fall back to the proven proxy read (which reads outline from `Paragraph.outlineLevel`).
    if (pkg.styleParseSuspect) {
      op.warn("pure whole-body: styles.xml present but unparseable — falling back to the proxy read path");
      return { ...(await this.readBatch(ctx, op)), trackChanges };
    }

    const total = pkg.count;
    const out: RawParagraph[] = [];
    let outlineSample: unknown = null;
    for (let i = 0; i < total; i++) {
      // Pre-write, always-safe window: tick the pacer — it observes cancellation AND (when a
      // yielding pacer was injected) slices this pure-JS loop into macrotask-bounded chunks so
      // the pane can paint progress and a Cancel click can land mid-read. Without an injected
      // pacer this is exactly the old bare `isCancelled()` poll (never yields).
      await this.pacer.tick();
      const headingLevel = pkg.headingLevel(i);
      const inTable = pkg.inTable(i);
      // LOOP 002 B1 — node-direct: hand the engine a `ParsedParagraph` over the package's LIVE `<w:p>`
      // node (zero per-paragraph serialize→parse — the win 002-S1 names). The engine reads its run
      // views and mutates `<w:vanish/>` IN PLACE on this same node, then this package serializes once.
      // `ooxml` is a placeholder: with `.parsed` set, the engine never reads it, and the node-direct
      // commit serializes the whole package (never a per-paragraph fragment), so we DROP the
      // `paragraphXml(i)` serialize the proxy path pays here.
      out.push({ index: i, headingLevel, inTable, ooxml: NODE_BACKED_OOXML, parsed: pkg.parsedParagraph(i) });
      if (outlineSample === null && headingLevel !== null) {
        outlineSample = { index: i, headingLevel, inTable, source: "package-styles" };
      }
      if ((i + 1) % this.chunkSize === 0 || i + 1 === total) {
        this.onProgress?.({ phase: "read", done: i + 1, total });
      }
    }

    // Identity mapping + cleanAlign: the package is self-consistent, so engineIndex === packageIndex
    // and the whole-body commit re-serializes exactly what we read (only <w:vanish/> changed).
    const pkgIndex = Array.from({ length: total }, (_, i) => i);
    op.debug("pure whole-body read assembled (node-direct)", { paragraphs: total, source: "package-styles" });
    return {
      paragraphs: out,
      state: { strategy: "pkg", pkg, count: total, pkgIndex, cleanAlign: true, nodeDirect: true },
      outlineSample,
      trackChanges
    };
  }

  /**
   * The read batch (Stage 4.1). ALWAYS reads the whole body story in ONE
   * `body.getOoxml()` — eliminating the per-paragraph `getOoxml()` storm that dominated
   * Hide (≈135s on a 376-paragraph doc). Proxy outline level + table membership + text
   * are loaded as cheap properties (chunked, so progress + cancel still work); decision
   * #7's outline level still comes from the PROXY, not the OOXML. The package is then
   * aligned to the proxies tolerantly (⑧): the body OOXML can hold ±1 `<w:p>` vs
   * `body.paragraphs` (a Word serialization quirk), so we match by text in document order
   * and skip unmatched package artifacts. If alignment can't be established we FALL BACK
   * to the exact (slow) per-paragraph `getOoxml()` path — correctness never depends on the
   * fast path. The COMMIT mechanism (per-paragraph vs whole-body) is chosen independently
   * at commit time from `this.strategy`.
   */
  private async readBatch(
    ctx: any,
    op: Logger
  ): Promise<{ paragraphs: RawParagraph[]; state: ReadState; outlineSample: unknown; trackChanges: TrackChangesMode }> {
    const doc: any = ctx.document;
    const body = doc.body;
    const paras = body.paragraphs;
    paras.load("items");
    const bodyOoxml = body.getOoxml(); // ALWAYS: one whole-body read
    // Loop 002 B2 (002-S4) — PREFETCH the Track-Changes mode in this same first sync (see readBatchPure).
    // Fuses the dedicated TC-read run into the read, so the proxy-path Hide also costs one fewer run.
    doc.load("changeTrackingMode");
    await ctx.sync(); // sync 1: items + whole-body OOXML + the TC mode
    const trackChanges = normalizeTcMode(String(doc.changeTrackingMode), op);

    const total: number = paras.items.length;
    op.debug("paragraph collection materialized", { total });

    // Load cheap proxy properties for every paragraph (outline level, table membership,
    // and text for alignment). NO per-paragraph getOoxml here — that only happens in the
    // fallback below. Chunked so the progress bar + cancellation keep working.
    const outlines: number[] = [];
    const inTables: boolean[] = [];
    const texts: string[] = [];
    for (let start = 0; start < total; start += this.chunkSize) {
      if (this.cancel?.isCancelled()) throw new CancelledError();
      const end = Math.min(start + this.chunkSize, total);
      for (let i = start; i < end; i++) {
        const p = paras.items[i];
        p.load("outlineLevel");
        p.parentTableOrNullObject.load("isNullObject");
        p.load("text");
      }
      await ctx.sync();
      for (let i = start; i < end; i++) {
        const p = paras.items[i];
        outlines[i] = p.outlineLevel;
        inTables[i] = p.parentTableOrNullObject.isNullObject === false;
        texts[i] = typeof p.text === "string" ? p.text : "";
      }
      this.onProgress?.({ phase: "read", done: end, total });
    }

    const headingOf = (i: number): number | null =>
      normalizeOutlineNumber(outlines[i], this.outline.numberBase);
    let outlineSample: unknown = null;
    for (let i = 0; i < total; i++) {
      if (outlines[i] >= 1 && outlines[i] <= 9) {
        outlineSample = { index: i, rawNumber: outlines[i], normalized: headingOf(i), inTable: inTables[i] };
        break;
      }
    }

    // FAST PATH: parse the whole body ONCE, align package <w:p> to the proxies by text.
    // Alignment is NON-CASCADING (Stage 4.2): a paragraph whose proxy `.text` doesn't match
    // our OOXML text rendering — an emoji (`w16se:symEx` inside `mc:AlternateContent`) Word
    // shows in `.text` but `collectParagraphText` omits, a `<w:noBreakHyphen>`, a field/host
    // `.text` quirk — is marked UNRESOLVED for a targeted re-read instead of collapsing the
    // WHOLE doc to the slow path. This is the fix for the ndca-semis regression: ONE divergent
    // paragraph (the emoji at proxy 259) used to send all 376 through per-paragraph getOoxml
    // (≈115s); now it costs ONE extra getOoxml.
    let pkg: WholeBodyPackage | null = null;
    let mapping: number[]; // proxy index → package story index; -1 = unresolved (re-read)
    let unresolved: number[]; // proxy indices needing a targeted per-paragraph re-read
    try {
      pkg = new WholeBodyPackage(bodyOoxml.value);
      const aligned = alignToProxies(texts, pkg, op);
      mapping = aligned.mapping;
      unresolved = aligned.unresolved;
    } catch (e) {
      // Whole-body parse failed outright — every paragraph falls to the targeted re-read
      // (equivalent to the old full per-paragraph path, but down the same one code path).
      op.warn("whole-body parse failed — per-paragraph fallback for all paragraphs", describeError(e));
      pkg = null;
      mapping = new Array(total).fill(-1);
      unresolved = Array.from({ length: total }, (_, i) => i);
    }

    // Targeted per-paragraph getOoxml for ONLY the unresolved proxies (the divergent few —
    // or all of them if the package didn't parse). Outline/table are already loaded above, so
    // this fetches just each unmatched paragraph's own OOXML. The word "fallback" stays in the
    // warning so existing diagnostics/tests still recognize the path.
    const targeted: string[] = [];
    if (unresolved.length) {
      op.warn("whole-body read: targeted per-paragraph getOoxml fallback for unaligned paragraphs", {
        total,
        unaligned: unresolved.length,
        sample: unresolved.slice(0, 8),
        // Stage A wet-test: log each unaligned proxy's `.text` (truncated) so the live run captures
        // the emoji proxy's actual rendering — the one input we can't get headless, needed to decide
        // whether `normForAlign` can be extended to align it or it must stay a targeted re-read.
        sampleTexts: unresolved.slice(0, 8).map((i) => (texts[i] ?? "").slice(0, 80))
      });
      for (let start = 0; start < unresolved.length; start += this.chunkSize) {
        if (this.cancel?.isCancelled()) throw new CancelledError();
        const end = Math.min(start + this.chunkSize, unresolved.length);
        const ooxmlResults: any[] = [];
        for (let s = start; s < end; s++) ooxmlResults[s] = paras.items[unresolved[s]].getOoxml();
        await ctx.sync();
        for (let s = start; s < end; s++) {
          const raw = ooxmlResults[s].value as string;
          if (s === 0) {
            op.debug("targeted getOoxml raw body-<w:p> count (pre-normalize)", {
              proxyIndex: unresolved[s],
              bodyParas: countBodyParagraphs(raw)
            });
          }
          targeted[unresolved[s]] = keepFirstBodyParagraph(raw);
        }
        this.onProgress?.({ phase: "read", done: end, total: unresolved.length });
      }
    }

    // Assemble: confident proxies use the cached whole-body package fragment (cheap, no host
    // round-trip); unresolved ones use their exact targeted OOXML.
    const out: RawParagraph[] = [];
    for (let i = 0; i < total; i++) {
      const ooxml = mapping[i] >= 0 && pkg ? pkg.paragraphXml(mapping[i]) : targeted[i];
      out.push({ index: i, headingLevel: headingOf(i), inTable: inTables[i], ooxml });
    }

    // cleanAlign = the package held EXACTLY the proxy paragraphs in identity order (no artifact
    // skipped, no divergence). UNCHANGED commit gate: only then is a whole-body Replace a
    // faithful round-trip, so the perf-mode whole-body COMMIT stays gated on it (see commit());
    // a ±1 serialization artifact or any unresolved paragraph ⇒ per-paragraph commit.
    const cleanAlign =
      pkg !== null && unresolved.length === 0 && pkg.count === total && mapping.every((v, i) => v === i);
    const strategy: "range" | "pkg" = pkg ? "pkg" : "range";
    op.debug("whole-body read assembled", {
      proxies: total,
      packageParagraphs: pkg ? pkg.count : 0,
      unaligned: unresolved.length,
      cleanAlign,
      strategy
    });
    return {
      paragraphs: out,
      // nodeDirect:false — the proxy read serializes per-paragraph fragments and commits via
      // `replace()` (it aligns the package to LIVE proxies, where in-place node mutation can't be
      // trusted under a ±1 artifact). Only the pure `readBatchPure` path is node-direct.
      state: { strategy, pkg, count: total, pkgIndex: mapping, cleanAlign, nodeDirect: false },
      outlineSample,
      trackChanges
    };
  }

  // ----- Write (buffer) -----------------------------------------------------

  async writeParagraphs(updates: ParagraphUpdate[]): Promise<void> {
    const op = this.log.child("write");
    // On the LOOP 002 node-direct path the update `ooxml` is a placeholder (the `<w:p>` was already
    // mutated IN PLACE inside the cached package and the commit serializes the whole package, never a
    // per-paragraph fragment), so its single-`<w:p>` validity is moot — and `assertSingleParagraph`
    // parses each fragment, which would re-introduce the very per-paragraph parse P1 deletes (it scales
    // with the changed-paragraph count, ~489 on ExFlex). Skip the fragment guard on that path; the
    // index range-check still runs (it never parses).
    const nodeDirect = this.lastRead?.nodeDirect === true;
    // Validate eagerly, BEFORE any host mutation: every fragment must hold exactly
    // one <w:p> (audit #5), and every index must be within the last read.
    for (const u of updates) {
      if (!nodeDirect) assertSingleParagraph(u.ooxml, `writeParagraphs @${u.index}`);
      if (this.lastRead && (u.index < 0 || u.index >= this.lastRead.count)) {
        throw new RangeError(
          `writeParagraphs: update index ${u.index} is out of range (document had ` +
            `${this.lastRead.count} paragraphs at read time).`
        );
      }
    }
    if (this.pending) {
      op.warn("overwriting un-flushed buffered updates", { dropped: this.pending.length });
    }
    this.pending = updates.slice();
    op.debug("buffered paragraph updates (await manifest op to flush atomically)", {
      count: updates.length
    });
  }

  /**
   * Discard any prepared write state (Loop 002 B1 — CONTRACT C / 002-F4). Called by the engine when
   * Phase B of a Hide ABORTS (a throwing `applyVisibilityInPlace`, or a cancel) AFTER it began mutating
   * the cached package's live `<w:p>` nodes but BEFORE any host write. Because the node-direct path
   * mutates `lastRead.pkg` in place, that package is now HALF-MUTATED in memory; nulling `lastRead`
   * (and any `pending` buffer) guarantees no later `commit`/manifest op can ever serialize it, and the
   * NEXT operation re-reads the unchanged on-disk document from scratch. The on-disk doc is untouched
   * (nothing was inserted), so this leaves it byte-identical with the manifest unarmed. Idempotent.
   */
  discardPreparedWrite(): void {
    if (this.lastRead || this.pending) {
      this.log.child("abort").debug("discarding prepared write (Phase-B abort) — re-read on next op", {
        hadRead: !!this.lastRead,
        hadPending: this.pending ? this.pending.length : 0
      });
    }
    this.lastRead = null;
    this.pending = null;
  }

  // ----- Reveal (fast Show All) --------------------------------------------

  /**
   * The fast Show All (Stage 4): clear `font.hidden` NATIVELY across the body story —
   * no OOXML read or write, so this is one/two host round-trips instead of the
   * thousands of `insertOoxml` reflows the old per-paragraph reveal did (~3 min on the
   * extremely-large doc). The strategy decides the shape:
   *   * whole-body (performance mode) → ONE `body.font.hidden = false` over the whole
   *     story — the same single-call move that makes Verbatim's reveal ~instant.
   *   * per-paragraph (safe mode) → `font.hidden = false` on each body paragraph;
   *     per-paragraph scoped (not a whole-doc nuke) but still property-only, no OOXML.
   * Behaviorally identical to the old makeAllVisible pass: it reveals the SAME set,
   * including any run the user hid manually (decision #10). `paragraphsChanged` is
   * best-effort (= scanned) because clearing is unconditional. NOT buffered/atomic with
   * the manifest clear — Show All is convergent, so a partial failure is re-runnable.
   * Intentionally NOT cancellable and emits no progress: this is a property-only pass of
   * one/two syncs with no reflow, so there is nothing meaningful to cancel or report
   * (the old, slow OOXML reveal path was cancellable; there's nothing left to cancel).
   */
  async clearHidden(): Promise<{ paragraphsScanned: number; paragraphsChanged: number }> {
    const op = this.log.child("clearHidden");
    const span = op.span("clearHidden", { strategy: this.strategy });
    try {
      const count = await this.run(async (ctx) => {
        const body = ctx.document.body;
        const paras = body.paragraphs;
        paras.load("items");
        // Performance mode: one blanket clear over the entire story (fastest path).
        if (this.strategy === "whole-body") body.font.hidden = false;
        await ctx.sync(); // sync 1: materialize the collection (+ whole-body clear)
        const total: number = paras.items.length;
        // Safe mode: per-paragraph native clear (still no OOXML).
        if (this.strategy !== "whole-body") {
          for (let i = 0; i < total; i++) paras.items[i].font.hidden = false;
          await ctx.sync(); // sync 2: the per-paragraph clears
        }
        return total;
      });
      span.end({ paragraphs: count });
      return { paragraphsScanned: count, paragraphsChanged: count };
    } catch (e) {
      span.fail(e);
      throw clarify(e, "reveal hidden text");
    }
  }

  // ----- Range-scoped read/replace (Condense & Shrink) ----------------------

  /**
   * Read the active selection — or the CURRENT PARAGRAPH when the selection is collapsed (an insertion
   * point) — as ONE OOXML fragment, plus each paragraph's canonical outline level (for Shrink's heading
   * refusal). This is the second document-access seam, parallel to the whole-body `readParagraphs`: it
   * never asserts a single `<w:p>` because Condense legitimately operates on (and produces) many.
   *
   * Collapsed detection uses `Range.text === ""` (a collapsed range has no text). All proxies live
   * within ONE `Word.run`, so re-fetching the selection in `replaceActiveRangeOoxml` is safe — the
   * controller does read → pure transform → write with no UI yield, so the selection can't shift.
   */
  async readActiveRangeOoxml(): Promise<RangeRead> {
    const op = this.log.child("rangeRead");
    const span = op.span("readActiveRangeOoxml");
    try {
      const result = await this.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        const selParas = sel.paragraphs;
        selParas.load("items");
        await ctx.sync(); // sync 1: selection text + its paragraphs

        const collapsed = (typeof sel.text === "string" ? sel.text : "") === "";
        // Collapsed → operate on the paragraph the cursor sits in; otherwise the selection range itself.
        const range = collapsed && selParas.items.length > 0 ? selParas.items[0].getRange() : sel;
        const paras = range.paragraphs;
        paras.load("items");
        const ooxmlResult = range.getOoxml();
        await ctx.sync(); // sync 2: range paragraphs + the range OOXML

        for (const p of paras.items) p.load("outlineLevel");
        await ctx.sync(); // sync 3: outline levels

        const outlineLevels = paras.items.map((p) =>
          normalizeOutlineNumber(p.outlineLevel, this.outline.numberBase)
        );
        // A collapsed read uses `paragraph.getRange()`, which defaults to the "Whole" range INCLUDING
        // the paragraph mark — so the host serializes the target paragraph PLUS a trailing empty `<w:p>`
        // (two body paragraphs). Strip it to the first body `<w:p>` (exactly as the hide path does via
        // `keepFirstBodyParagraph`) so the fragment holds ONE `<w:p>` aligned with the single outline
        // level — keeping Shrink's single-paragraph heading refusal and Condense's collapsed no-op
        // correct. NOT applied to a real multi-paragraph SELECTION, where it would wrongly drop all but
        // the first selected paragraph.
        const rawOoxml = ooxmlResult.value as string;
        const ooxml = collapsed ? keepFirstBodyParagraph(rawOoxml) : rawOoxml;
        return { ooxml, collapsed, outlineLevels };
      });
      span.end({ collapsed: result.collapsed, paragraphs: result.outlineLevels.length });
      return result;
    } catch (e) {
      span.fail(e);
      throw clarify(e, "read the active selection");
    }
  }

  /**
   * Replace the active range with new OOXML via `insertOoxml(…, "Replace")`. Re-derives the same range
   * `readActiveRangeOoxml` chose (selection, or the current paragraph when collapsed), so Condense's
   * merged single paragraph lands exactly where the cards were, then RE-SELECTS the inserted content so a
   * repeated Shrink/Condense press keeps acting on the whole block (not just the collapsed cursor's
   * paragraph). One `Word.run`, one atomic sync.
   */
  async replaceActiveRangeOoxml(ooxml: string): Promise<void> {
    const op = this.log.child("rangeWrite");
    const span = op.span("replaceActiveRangeOoxml", { length: ooxml.length });
    try {
      await this.run(async (ctx) => {
        const sel = ctx.document.getSelection();
        sel.load("text");
        const selParas = sel.paragraphs;
        selParas.load("items");
        await ctx.sync(); // sync 1: selection text + its paragraphs

        const collapsed = (typeof sel.text === "string" ? sel.text : "") === "";
        const range = collapsed && selParas.items.length > 0 ? selParas.items[0].getRange() : sel;
        // RE-SELECT the inserted content: insertOoxml("Replace") otherwise collapses the selection to a
        // cursor, so a repeated Shrink/Condense press would only act on the current paragraph (the
        // collapsed-cursor fallback) instead of the whole multi-paragraph block the user shrank. Selecting
        // the returned range keeps the original selection live across presses — no keyboard shortcut needed.
        const inserted = range.insertOoxml(ooxml, "Replace");
        inserted.select();
        await ctx.sync(); // sync 2: the atomic replace + reselect
      });
      span.end();
    } catch (e) {
      span.fail(e);
      throw clarify(e, "replace the active selection");
    }
  }

  // ----- Cite repair (Apply Styles) ----------------------------------------

  /**
   * Detect mis-styled cites and re-apply the cite character style, via the WHOLE-BODY
   * OOXML path: ONE `body.getOoxml()` → build a `WholeBodyPackage` → `planCiteRepairs`
   * over every package paragraph (outline from the package's own styles.xml — no proxy
   * needed) → `applyCiteStyleToParagraphXml` on each repaired paragraph → splice each back
   * via `pkg.replace` → ONE `body.insertOoxml(…, "Replace")`.
   *
   * WHY THIS IS SAFE TO WHOLE-BODY-REPLACE (unlike the hide path's gated commit). The hide
   * path gates its whole-body Replace on `cleanAlign` because it aligns the package to LIVE
   * PROXIES and a skipped artifact could be re-injected. Here there are NO proxies: we read
   * N `<w:p>`, mutate only run-level `<w:rStyle>` in THIS SAME parsed DOM, and write the
   * SAME N back — the exact self-consistent identity round-trip the pure ⑦ path relies on.
   * Only the injected rStyle changes; every other byte round-trips.
   *
   * WHY IT NEEDS ONLY getOoxml/insertOoxml (the hard floor). It does NOT touch
   * `getStyles`/`Style.font`, so it runs even when size-application is gated off (a host
   * without WordApi 1.5) — making Apply Styles useful on more hosts.
   *
   * RESILIENCE. A parse failure, an unparseable styles part (`styleParseSuspect` — don't
   * mass-mutate when outline resolution is untrustworthy), or simply no repairs needed all
   * return zeros WITHOUT writing or throwing, so a controller can always call this safely.
   */
  async repairCites(): Promise<CiteRepairResult> {
    const op = this.log.child("citeRepair");
    const span = op.span("repairCites");
    try {
      const result = await this.run(async (ctx) => {
        const body = ctx.document.body;
        const bodyOoxml = body.getOoxml();
        await ctx.sync(); // sync 1: the whole-body OOXML

        // Parse the package; a pathological doc that won't parse is a no-op (not a failure).
        let pkg: WholeBodyPackage;
        try {
          pkg = new WholeBodyPackage(bodyOoxml.value);
        } catch (e) {
          op.warn("cite-repair: whole-body parse failed — skipping (no changes)", describeError(e));
          return { paragraphsRepaired: 0, runsRepaired: 0 };
        }

        // Untrustworthy styles part → outline resolution would mis-classify tags, so the
        // forward scan could pick wrong candidates. Skip rather than risk mass-mutation.
        if (pkg.styleParseSuspect) {
          op.warn("cite-repair: styles.xml present but unparseable — skipping (no changes)");
          return { paragraphsRepaired: 0, runsRepaired: 0 };
        }

        // Build the planner's view from the package: each paragraph's raw XML + its
        // resolved heading level (the tag-boundary signal). `headingLevel` is EAGER —
        // the planner reads it for EVERY paragraph to find tag windows — but the wrapped
        // fragment is a CACHED LAZY getter: `planCiteRepairs` only reads `.xml` for
        // paragraphs inside a tag's forward-scan window (the list/empty probes and the
        // single candidate), ~10% of a real doc, while `paragraphXml` is expensive per
        // call (clone + serialize + rels walk + flat-OPC wrapping). Materializing all N
        // up front cost more than the whole plan (~20s of blocking webview JS on a
        // 13k-paragraph doc, paid even when zero repairs are found). The cache also lets
        // the apply loop below reuse each repaired candidate's already-built fragment.
        const paras: CiteRepairParagraph[] = [];
        for (let i = 0; i < pkg.count; i++) {
          const idx = i; // explicit per-paragraph capture for the getter closure
          let cached: string | null = null;
          paras.push({
            headingLevel: pkg.headingLevel(idx),
            // A getter structurally satisfies the interface's `xml: string`; it defers
            // the wrap until (and unless) this paragraph is actually inspected.
            get xml(): string {
              if (cached === null) cached = pkg.paragraphXml(idx);
              return cached;
            }
          });
        }
        const repairs = planCiteRepairs(paras, CITE_STYLE_ID);
        if (repairs.length === 0) {
          op.debug("cite-repair: no mis-styled cites found", { paragraphs: pkg.count });
          return { paragraphsRepaired: 0, runsRepaired: 0 };
        }

        // Apply each repair to its paragraph fragment and splice it back into the package.
        // `paragraphXml(i)` is a flat-OPC package wrapping one `<w:p>`; `pkg.replace` accepts
        // exactly that shape (it extracts the single body `<w:p>`), so we mutate the fragment
        // and hand it straight back. Every repaired index was a planner candidate, so its
        // lazy fragment is already materialized — this read is a cache hit, not a re-wrap.
        let runsRepaired = 0;
        for (const r of repairs) {
          const fragment = paras[r.paragraphIndex].xml;
          const repaired = applyCiteStyleToParagraphXml(fragment, r.runIndices, CITE_STYLE_ID);
          pkg.replace(r.paragraphIndex, repaired);
          runsRepaired += r.runIndices.length;
        }
        body.insertOoxml(pkg.serialize(), "Replace");
        op.debug("cite-repair: whole-body insertOoxml queued", {
          paragraphsRepaired: repairs.length,
          runsRepaired
        });
        await ctx.sync(); // sync 2: commit the repaired body
        return { paragraphsRepaired: repairs.length, runsRepaired };
      });
      span.end({ paragraphsRepaired: result.paragraphsRepaired, runsRepaired: result.runsRepaired });
      return result;
    } catch (e) {
      span.fail(e);
      throw clarify(e, "repair mis-styled cites");
    }
  }

  // ----- Manifest (flush points) -------------------------------------------

  async readManifest(): Promise<string | null> {
    const op = this.log.child("manifestRead");
    const span = op.span("readManifest");
    try {
      const xml = await this.run(async (ctx) => {
        const scoped = ctx.document.customXmlParts.getByNamespace(MANIFEST_NAMESPACE);
        const part = scoped.getOnlyItemOrNullObject();
        part.load("isNullObject,id");
        await ctx.sync();
        if (part.isNullObject) return null;
        this.manifestPartId = part.id;
        const xmlResult = part.getXml();
        await ctx.sync();
        return xmlResult.value as string;
      });
      span.end({ present: xml !== null, partId: this.manifestPartId });
      return xml;
    } catch (e) {
      span.fail(e);
      throw clarify(e, "read the Rostrum manifest");
    }
  }

  async writeManifest(xml: string): Promise<void> {
    await this.commit({ kind: "set", xml }, "writeManifest");
  }

  async clearManifest(): Promise<void> {
    await this.commit({ kind: "delete" }, "clearManifest");
  }

  /**
   * The OOXML fragment to hand the host for a per-paragraph `insertOoxml` commit (the fallback path
   * when whole-body ⑦ isn't in use). The engine's edited fragment is re-wrapped via `commitXml` so it
   * carries the document's style/numbering/theme parts and STYLE-INHERITED formatting (underline,
   * character box, font size) renders faithfully. With no pkg (a pure range read) the engine fragment
   * is already Word's full styled package, so it's committed unchanged.
   */
  private commitFragment(engineOoxml: string, pkg: WholeBodyPackage | null): string {
    return pkg ? pkg.commitXml(engineOoxml) : engineOoxml;
  }

  /**
   * The atomic commit. Flushes any buffered paragraph updates AND the manifest
   * action in ONE `Word.run` whose SECOND `sync()` is the single transaction
   * (audit H3) — so the document can never end up hidden-but-unarmed. The first
   * sync only reads (manifest existence + paragraph items for per-paragraph mode).
   */
  private async commit(manifest: ManifestAction, label: string): Promise<void> {
    const updates = this.pending ?? [];
    this.pending = null;
    const op = this.log.child("commit");
    // Resolve the effective writeback mechanism (with the per-paragraph fallback if
    // a whole-body read never happened or its package is missing).
    // Whole-body splice (perf mode) ONLY when (a) this.strategy is whole-body, (b) the
    // read produced an aligned package, AND (c) that alignment was CLEAN (identity, no
    // artifact skipped). The READ is always whole-body now, so in safe mode lastRead can
    // be "pkg" too — hence the this.strategy gate. The cleanAlign gate is the fix for the
    // phantom-injection risk: a full-body Replace re-serializes EVERY package paragraph,
    // so re-emitting a skipped artifact would inject it; when alignment isn't clean we
    // commit per-paragraph instead (surgical — can't inject or duplicate).
    const usePkg =
      this.strategy === "whole-body" &&
      !!(this.lastRead?.strategy === "pkg" && this.lastRead.pkg && this.lastRead.cleanAlign);
    // The whole-body read's package (present whenever a whole-body read produced one — even when
    // alignment wasn't clean, so we DON'T use the whole-body commit). The per-paragraph OOXML
    // commit re-wraps each edited fragment through `pkg.commitXml` so style-inherited formatting
    // (underline, character box, font size) survives; without a pkg (pure range read) the engine
    // fragment is already Word's full, styled package and is committed as-is.
    const pkg = this.lastRead?.pkg ?? null;
    const span = op.span(label, { paragraphs: updates.length, manifest: manifest.kind, usePkg });
    try {
      await this.run(async (ctx) => {
        const doc = ctx.document;
        const scoped = doc.customXmlParts.getByNamespace(MANIFEST_NAMESPACE);
        const existing = scoped.getOnlyItemOrNullObject();
        existing.load("isNullObject,id");
        const paras = updates.length && !usePkg ? doc.body.paragraphs : null;
        if (paras) paras.load("items");
        await ctx.sync(); // sync 1: manifest existence (+ items for per-paragraph)

        // --- paragraph edits ---
        if (updates.length) {
          if (usePkg && this.lastRead!.pkg && this.lastRead!.nodeDirect) {
            // LOOP 002 B1 — node-direct commit. The engine already mutated each changed paragraph's
            // `<w:p>` IN PLACE inside THIS cached package (the `ParsedParagraph` handles wrap the
            // package's own nodes), so there is NOTHING to splice: serialize the whole package once
            // and the mutations are already present. We deliberately do NOT call `pkg.replace()` (no
            // re-parse, no re-import of an identical node) — that is the per-paragraph serialize P1
            // deletes — and we ignore each update's placeholder `ooxml`.
            doc.body.insertOoxml(this.lastRead!.pkg.serialize(), "Replace");
            op.debug("node-direct whole-body insertOoxml queued (no per-paragraph splice)", {
              changed: updates.length
            });
          } else if (usePkg && this.lastRead!.pkg) {
            const pkg = this.lastRead!.pkg;
            const map = this.lastRead!.pkgIndex;
            // Map engine index (proxy order) → package story index via the read's
            // tolerant alignment, so a ±1 serialization artifact can't mis-splice.
            for (const u of updates) pkg.replace(map ? map[u.index] : u.index, u.ooxml);
            doc.body.insertOoxml(pkg.serialize(), "Replace");
            op.debug("whole-body insertOoxml queued", { spliced: updates.length });
          } else {
            // Per-paragraph (safe mode). Apply in DESCENDING index order so an earlier
            // paragraph's identity can't be disturbed by a later replacement. For each
            // update pick the CHEAPEST faithful write (Stage 4 perf):
            //   * keepWhole / hideWhole → a NATIVE `font.hidden` toggle on the whole
            //     paragraph (incl. its mark) — no OOXML parse, no per-call reflow. This
            //     is the bulk of a debate doc (whole card bodies) and is what closes
            //     the 3:40-vs-0:51 gap with Verbatim. hideWhole has, by construction, no
            //     kept/structural runs (a body field/footnote-ref ⇒ hidePartial), so a
            //     blunt whole-paragraph hide is exactly the OOXML <w:vanish/> result.
            //   * hidePartial (or no action — pre-Stage-4 callers) → OOXML
            //     `insertOoxml("Replace")`, the only way to hide some runs while keeping
            //     highlighted ones in the same paragraph.
            const ordered = updates.slice().sort((a, b) => b.index - a.index);
            let native = 0;
            let viaOoxml = 0;
            for (const u of ordered) {
              const p = paras!.items[u.index];
              if (u.action === "hideWhole" || u.action === "keepWhole") {
                p.font.hidden = u.action === "hideWhole";
                native++;
              } else {
                // Symmetric with the read: replace the paragraph via
                // `Paragraph.insertOoxml(…, "Replace")` (WordApi 1.1), not the range. Re-wrap the
                // engine's edited fragment so it carries the document's style/numbering/theme parts
                // — an aligned paragraph was READ from a style-LESS minimal package, so committing it
                // bare made the host reset style-inherited formatting (underline, box, font size) to
                // defaults while inline highlight survived. With no pkg (pure range read) u.ooxml is
                // already Word's full styled package, so commit it unchanged.
                p.insertOoxml(this.commitFragment(u.ooxml, pkg), "Replace");
                viaOoxml++;
              }
            }
            op.debug("per-paragraph apply queued", { native, viaOoxml });
          }
        }

        // --- manifest (same batch → atomic with the edits) ---
        if (manifest.kind === "set") {
          if (existing.isNullObject) {
            doc.customXmlParts.add(manifest.xml);
            op.debug("manifest added (no prior part)");
          } else {
            existing.setXml(manifest.xml);
            op.debug("manifest setXml in place", { partId: existing.id });
          }
        } else if (manifest.kind === "delete") {
          if (!existing.isNullObject) {
            existing.delete();
            op.debug("manifest deleted", { partId: existing.id });
          }
        }

        this.onProgress?.({ phase: "commit", done: updates.length, total: updates.length });
        await ctx.sync(); // sync 2: THE atomic commit (edits + manifest together)
      });
      span.end();
    } catch (e) {
      // The buffer was already cleared; surface a clear, contextual failure.
      // DIAGNOSTICS (Stage 4.2): a host `insertOoxml` rejection ("problem with its contents",
      // GeneralException at Paragraph.insertOoxml) is opaque without the OFFENDING OOXML. Log a
      // sample of the OOXML fragments this commit fed to the host — index, byte length, whether
      // they carry a relationship/MC reference, and a head excerpt — so the next wet-test pins
      // the exact rejected XML (Word's "Line N, Column M" indexes into this string) instead of
      // guessing. Only the OOXML-replace updates can be rejected (native font.hidden can't).
      const ooxmlUpdates = updates.filter((u) => u.action !== "hideWhole" && u.action !== "keepWhole");
      op.warn("commit failed — sample of OOXML fragments fed to the host", {
        ooxmlFragments: ooxmlUpdates.length,
        nativeToggles: updates.length - ooxmlUpdates.length,
        sample: ooxmlUpdates.slice(0, 3).map((u) => {
          // Sample the fragment ACTUALLY fed to the host (the re-wrapped, styled package when a
          // pkg was present) so the rejected XML — and now `hasStyles` — matches Word's "Line N
          // Col M". Re-wrapping is pure; guard it so a wrap failure can't mask the original error.
          let fed = u.ooxml;
          try {
            fed = this.commitFragment(u.ooxml, pkg);
          } catch {
            /* fall back to the engine fragment for the sample */
          }
          return {
            index: u.index,
            length: fed.length,
            hasRel: /\sr:[a-zA-Z]+=/.test(fed),
            hasStyles: fed.includes("/word/styles.xml"),
            hasAlternateContent: fed.includes("AlternateContent"),
            head: fed.slice(0, 400)
          };
        })
      });
      span.fail(e);
      throw clarify(e, `commit ${updates.length} paragraph edit(s) + manifest (${manifest.kind})`);
    }
  }

  /** The manifest part id last observed (diagnostics only — lookups go by namespace). */
  get diagnosticManifestPartId(): string | null {
    return this.manifestPartId;
  }
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

/** Validate the host's reported TC mode; default to "Off" on anything unexpected. */
function normalizeTcMode(mode: string, op: Logger): TrackChangesMode {
  if (mode === "Off" || mode === "TrackAll" || mode === "TrackMineOnly") return mode;
  op.warn("unexpected changeTrackingMode from host; treating as Off", { mode });
  return "Off";
}

/** A short human string for an unknown throw — used inside composed error messages. */
function summarize(e: unknown): string {
  const d = describeError(e);
  const code = d.code ? `${d.code}: ` : "";
  return `${code}${d.message ?? d.error ?? "unknown error"}`;
}

/**
 * Wrap a caught Office error in a clear, action-tagged Error while preserving the
 * original (and its Office code/debugInfo) as `cause`, so the task pane can show
 * "Rostrum couldn't <action>" and the diagnostics log still has the raw fields.
 */
function clarify(e: unknown, action: string): Error {
  const wrapped = new Error(`Rostrum couldn't ${action}. ${summarize(e)}`);
  (wrapped as any).cause = e;
  return wrapped;
}

/**
 * Whitespace-normalize for tolerant matching, AND drop characters that Word's
 * `Paragraph.text` includes but our OOXML text walk can't reliably reproduce — astral
 * (surrogate-pair) codepoints and emoji modifiers (variation selectors, ZWJ). The ndca
 * doc carries an emoji (`w16se:symEx` inside `mc:AlternateContent`) that `.text` renders
 * but `collectParagraphText` omits; stripping it on BOTH sides makes that paragraph align
 * cleanly instead of needing a re-read. Anything still divergent (fields, `<w:noBreakHyphen>`)
 * falls to the targeted re-read, so this is an optimization, never a correctness dependency.
 */
function normForAlign(s: string): string {
  return s
    .replace(/[\uD800-\uDFFF︎️‍]/g, "") // astral (surrogate) / variation selectors / ZWJ
    .replace(/\s+/g, " ")
    .trim();
}

/**
 * Tolerant, NON-CASCADING alignment (Stage 4.1 ⑧ → 4.2): map each Office paragraph proxy
 * (index i, text `proxyTexts[i]`) to a package STORY-paragraph index, in document order.
 * Two real-doc discrepancies are handled DIFFERENTLY:
 *
 *   • ARTIFACT — `body.getOoxml()` can serialize ±1 more `<w:p>` than `body.paragraphs`
 *     enumerates (377 vs 376 on the ndca doc, 0 textboxes). These extra package paragraphs
 *     are SKIPPED via a small forward window, capped by a global skip budget (`pk − N`) so we
 *     never skip past where the remaining proxies must live (no cascade into duplicate text).
 *
 *   • DIVERGENCE — a proxy's `.text` ≠ our text rendering of its real package paragraph
 *     (the emoji case above, a field, a host quirk). The OLD code let ONE such paragraph run
 *     the cursor off the end and FALL BACK THE WHOLE DOCUMENT to the ≈115s per-paragraph read.
 *     Now that one proxy is recorded as UNRESOLVED (the adapter re-reads just it via a targeted
 *     getOoxml) and the pointers stay in lockstep, so paragraph 260 onward still align.
 *
 * Returns `mapping` (engineIndex→packageIndex, −1 where unresolved) and `unresolved` (the proxy
 * indices needing a targeted re-read). Empty paragraphs match each other harmlessly (no-ops).
 *
 * UNIQUENESS SAFETY (Stage 4.2, reviewer catch): a text match is only TRUSTED for a non-empty
 * paragraph when that text occurs EXACTLY ONCE in the package. If a real paragraph's text is
 * duplicated by a same-text serialization artifact (or a genuine repeat), positional matching
 * could bind proxy i to the WRONG `<w:p>` and the per-paragraph commit would then write a
 * different paragraph's content onto proxy i (silent corruption). Such ambiguous proxies are
 * marked UNRESOLVED → re-read exactly. A unique text can only match its one true slot, so the
 * fast path is kept for the overwhelming majority of (unique) card bodies. Empty text classifies
 * to a no-op, so an empty↔empty mispairing changes nothing and needs no re-read.
 */
function alignToProxies(
  proxyTexts: string[],
  pkg: WholeBodyPackage,
  op: Logger
): { mapping: number[]; unresolved: number[] } {
  const pk = pkg.count;
  const N = proxyTexts.length;
  const pkgTexts: string[] = [];
  for (let j = 0; j < pk; j++) pkgTexts.push(normForAlign(pkg.paragraphText(j)));
  // Frequency of each normalized package text — a non-empty text seen 2+ times is ambiguous
  // (real paragraph + a same-text artifact, or a genuine duplicate) and must not be trusted.
  const freq = new Map<string, number>();
  for (const t of pkgTexts) freq.set(t, (freq.get(t) ?? 0) + 1);
  // Bound on consecutive artifacts skipped to re-anchor after a local mismatch. Small, so
  // duplicate short lines (empty paras, "AND:", numbering) can't be leap-frogged into a wrong
  // match; the per-step budget below also caps total skips at the artifact count (pk − N).
  const SKIP_WINDOW = 4;
  const mapping: number[] = new Array(N).fill(-1);
  const unresolved: number[] = [];
  let j = 0;
  for (let i = 0; i < N; i++) {
    const want = normForAlign(proxyTexts[i]);
    // Never skip so far that the remaining package paragraphs can't still cover the remaining
    // proxies — keeps artifact-skipping from cascading into duplicate-text mismatches.
    const skipBudget = pk - j - (N - i);
    const window = Math.min(SKIP_WINDOW, Math.max(0, skipBudget));
    let found = -1;
    for (let k = 0; k <= window; k++) {
      if (j + k < pk && pkgTexts[j + k] === want) {
        found = j + k;
        break;
      }
    }
    // Trust the match ONLY when unambiguous: the text is empty (a no-op classify) or UNIQUE in the
    // package. A non-empty, duplicated text → unresolved (re-read), so we never bind a proxy to a
    // wrong same-text paragraph and commit foreign content onto it (lesson #33).
    //
    // REJECTED (Stage A, lesson #37): also trusting a duplicate when its proxy/package occurrence
    // COUNTS match (`pkgFreq === proxyFreq`) is UNSOUND. Adversarial review proved two breaks:
    // (1) text equality ≠ content identity — two same-text paragraphs with DIFFERENT content (one
    // highlighted-to-keep, one plain) serialized by getOoxml in a different relative order than the
    // proxies bind swapped, and because the mapping still looks identity-clean `cleanAlign` goes
    // true and the DESTRUCTIVE whole-body Replace hides the highlighted keep; (2) the blind `j++`
    // below (on an unresolved proxy when pk < N) desyncs the cursor onto a later duplicate it then
    // trusts. We can't verify content identity from a proxy (it exposes only `.text` + outline), so
    // a sound duplicate fast path needs content-aware disambiguation — deferred to the wet-test-
    // gated whole-body-commit unlock. Until then, re-read duplicates exactly (correctness first).
    const trustworthy = found >= 0 && (want === "" || (freq.get(want) ?? 0) === 1);
    if (trustworthy) {
      mapping[i] = found; // skipped j..found-1 are artifacts
      j = found + 1;
    } else {
      // Divergence, package exhausted, OR ambiguous duplicate text: re-read this ONE paragraph.
      unresolved.push(i);
      if (j < pk) j++; // assume 1:1 here so the next proxy resumes at its real slot
    }
  }
  if (unresolved.length) {
    op.debug("whole-body alignment: paragraphs needing a targeted re-read", {
      proxies: N,
      packageParagraphs: pk,
      unresolved: unresolved.length
    });
  }
  return { mapping, unresolved };
}

/**
 * Construct the production WordPort. Pass `runner` (a fake) in tests; in the host,
 * leave it unset to use `Word.run`. The returned object also exposes
 * `diagnosticManifestPartId` for the task pane's diagnostics view.
 */
export function createOfficeWordPort(
  options: OfficeWordPortOptions = {}
): CiteRepairCapablePort & RangeScopedPort & { readonly diagnosticManifestPartId: string | null } {
  return new OfficeWordPort(options);
}
