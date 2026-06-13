// REAL-document scale suite for the gdocs engine (plan S13 + A13 + case
// 001-S6) — the pre-wet evidence that the planner's request/byte envelope
// holds on REALISTIC large documents, not just hand-built fixtures.
//
// Inputs are the synthetic documents.get-shaped JSON fixtures that
// tools/gdocs-synth-fixture.mjs derives from the repo's real debate .docx
// corpus (rostrum/samples). They live in rostrum-addin/samples/gdocs/ —
// GITIGNORED on purpose (personal debate content) — so this suite is
// skip-when-absent (lesson #44: an fs.existsSync gate registering a real
// placeholder test, NEVER an empty it.each, which is a hard Jest error).
// CI has no samples/ and runs only the placeholder.
//
// What each present fixture must prove:
//   1. parseDocument decodes it into a structurally sound view (contiguous
//      UTF-16 indexes from 1, trailing newlines inside final runs).
//   2. planHide invariants hold AT SCALE: only the two fresh-hide request
//      types, anchors decode + RLE-exact + non-overlapping, anchors tile
//      their region's shrink exactly, fontSize-only sentinel writes.
//   3. The SCALE ENVELOPE (plan A13 / 001-S6) holds for the WORST case:
//      < 100k requests, <= 25 chunks, < 45MB payload — for BOTH verbs' plans
//      and for the fixture bytes (the documents.get payload proxy). These
//      bounds are the plan's locked envelope: if one fails, the numbers are
//      reported and the bound is NOT weakened here. (The showAll half of the
//      envelope is asserted wherever planShowAll runs — see HEAVY_TIERS.)
//   4. Restore stays plan-level EXACT, in two layers: (a) on EVERY tier, each
//      anchor's RLE record is replayed directly against the ORIGINAL document
//      (the in-document manifest alone suffices to restore byte-exact); (b)
//      the full planShowAll replay over an armed view — per-write size
//      verification, anchor deletion, sweep adopting exactly the pre-existing
//      tiny passages planHide independently counted — runs on small/medium
//      always and on heavy tiers under ROSTRUM_PERF=1 (HEAVY_TIERS comment
//      has the measured jest-vm pathology that forces the deferral).
//   5. The SMALL fixture additionally runs the LIVE hide -> showAll round
//      trip through FakeDocs + the controller (chunked, revision-chained)
//      and must come back view-equal per character. Bigger tiers stay
//      plan-level only — FakeDocs applies writes by splitting runs across
//      the whole model per request, which is O(requests x runs) and exists
//      to prove semantics, not throughput (the flagship gdocsInvisibility
//      suite owns semantics; here it would just burn minutes).
//
// Measurements are REPORTED (expect-and-log) per fixture and appended to
// samples/gdocs/SCALE.md (gitignored) for the orchestrator's wet-test packet.

import * as fs from "fs";
import * as os from "os";
import * as path from "path";
import { CHUNK_MAX, NAME_MAX, SENTINEL_PT, SENTINELS } from "../google-docs/src/core/constants";
import { hide, showAll } from "../google-docs/src/core/controller";
import { chunkGroups } from "../google-docs/src/core/guards";
import { parseDocument } from "../google-docs/src/core/parse";
import { planHide } from "../google-docs/src/core/planner";
import { decodeRangeName, isRstmName } from "../google-docs/src/core/rangeNames";
import { planShowAll } from "../google-docs/src/core/restore";
import { resolveSettings } from "../google-docs/src/core/settings";
import { DocsRequest, GDoc, GElement, GNamedRange, RequestGroup } from "../google-docs/src/core/types";
import { FakeDocs } from "./fakeDocs";
import { GeSpec, GpSpec } from "./gdocsBuilders";

// ---------------------------------------------------------------------------
// The locked scale envelope (plan A13 + case 001-S6). NOT tuning knobs: a
// failure here is a finding to report, never a number to raise.
// ---------------------------------------------------------------------------

/** Hard ceiling on a single verb's planned request count. */
const MAX_REQUESTS = 100_000;
/** Hard ceiling on a single verb's batch count. */
const MAX_CHUNKS = 25;
/** Hard ceiling on payload bytes — both the summed batchUpdate JSON and the
 * fixture file itself (the documents.get response proxy). */
const MAX_PAYLOAD_BYTES = 45 * 1024 * 1024;

// ---------------------------------------------------------------------------
// Fixture discovery (skip-when-absent, lesson #44)
// ---------------------------------------------------------------------------

/** rostrum-addin/samples/gdocs — the synth tool's gitignored output dir. */
const SYNTH_DIR = path.resolve(__dirname, "../samples/gdocs");

type SizeTier = "small" | "medium" | "large" | "xlarge";

interface FixtureRef {
  /** Basename, e.g. `[small] 2ac---dds2---finals.json`. */
  file: string;
  fullPath: string;
  tier: SizeTier;
  bytes: number;
}

/**
 * Tiers whose planShowAll cross-verification is deferred to ROSTRUM_PERF=1
 * runs — the SAME opt-in convention the Word realDocs family uses for its
 * heavy tiers, forced here by a MEASURED jest pathology, not by the engine:
 * planShowAll over the armed xlarge view takes ~4s in plain node (and that is
 * what an Apps Script runtime resembles) but ~460s inside jest's vm context
 * with the identical transpiled code and identical inputs/outputs. The
 * restore-record exactness still gets proven on EVERY tier by the direct
 * RLE-vs-original check below; what defers is only the planShowAll replay
 * (whose semantics gdocsRestore.test.ts owns at unit scale).
 */
const HEAVY_TIERS: SizeTier[] = ["large", "xlarge"];
/** Opt-in switch for the heavy-tier planShowAll verification. */
const RUN_HEAVY = process.env.ROSTRUM_PERF === "1";

/** Tier from the leading `[tag]` naming convention shared with the Word
 * realDocs family (realDocs.ts tierOf), falling back to raw byte size. The
 * tier picks WHICH fixture earns the live FakeDocs round trip (small only). */
function tierOf(file: string, bytes: number): SizeTier {
  const tag = (/^\[([^\]]+)\]/.exec(file)?.[1] ?? "").toLowerCase();
  if (/extremely large|x-?large|huge|full/.test(tag)) return "xlarge";
  if (/large/.test(tag)) return "large";
  if (/medium|med/.test(tag)) return "medium";
  if (/small|test|poc/.test(tag)) return "small";
  if (bytes > 10_000_000) return "xlarge";
  if (bytes > 3_000_000) return "large";
  if (bytes > 1_200_000) return "medium";
  return "small";
}

/**
 * Discover fixture JSONs in `dir`, smallest first. JUNK-TOLERANT on purpose
 * (the realDocs discoverSamples convention): this is a local scratch dir a
 * human drops files into, so an unparseable .json is warned about and
 * skipped rather than detonating every test in the suite. SCALE.md and any
 * non-.json files are ignored by the extension filter.
 */
function discoverFixtures(dir: string): FixtureRef[] {
  if (!fs.existsSync(dir)) return [];
  return fs
    .readdirSync(dir)
    .filter((f) => f.toLowerCase().endsWith(".json"))
    .flatMap((file): FixtureRef[] => {
      const fullPath = path.join(dir, file);
      try {
        // Parse probe only — the test re-reads; this guards discovery so one
        // truncated download cannot turn into N cryptic per-test failures.
        JSON.parse(fs.readFileSync(fullPath, "utf8"));
      } catch {
        // eslint-disable-next-line no-console
        console.warn(`[gdocs-scale] skipping unparseable fixture: ${file}`);
        return [];
      }
      const bytes = fs.statSync(fullPath).size;
      return [{ file, fullPath, tier: tierOf(file, bytes), bytes }];
    })
    .sort((a, b) => a.bytes - b.bytes);
}

// ---------------------------------------------------------------------------
// Plan dissection helpers
// ---------------------------------------------------------------------------

/** One anchor (createNamedRange) lifted out of a plan, half-open span form. */
interface AnchorInfo {
  name: string;
  start: number;
  end: number;
}

/** Flatten a plan's requests (chunking is asserted separately). */
function requestsOf(groups: RequestGroup[]): DocsRequest[] {
  return groups.flatMap((g) => g.requests);
}

/**
 * Hot-loop assertion: a plain throw, NOT expect(). These suites check
 * hundreds of thousands of per-request/per-element facts on the xlarge
 * fixture, and jest matcher overhead at that volume turned a ~10s test into
 * a 6-minute one (measured). The lazy message keeps the failure as
 * locatable as a matcher diff without paying string-building on every pass.
 */
function check(cond: boolean, msg: () => string): void {
  if (!cond) throw new Error(msg());
}

/** Merge sorted-by-start spans, treating touching as one — the passage unit
 * both planner receipts and these assertions count in. */
function mergedSpans(spans: { start: number; end: number }[]): { start: number; end: number }[] {
  const sorted = [...spans].sort((a, b) => a.start - b.start);
  const out: { start: number; end: number }[] = [];
  for (const s of sorted) {
    const prev = out[out.length - 1];
    if (prev !== undefined && s.start <= prev.end) prev.end = Math.max(prev.end, s.end);
    // Bare {start, end} ON PURPOSE: callers compare merged unions with
    // toEqual, so extra properties (AnchorInfo.name) must never leak through.
    else out.push({ start: s.start, end: s.end });
  }
  return out;
}

/**
 * Per-position style oracle over the ORIGINAL parsed doc: the flat ascending
 * list of every element span with its explicit size (and kind). Restore
 * exactness is judged against THIS — the document as it was before the hide
 * plan — so the verification chain never trusts the planner's own bookkeeping.
 */
interface SpanFact {
  start: number;
  end: number;
  kind: "text" | "other";
  sizePt: number | null;
}

function spanFactsOf(doc: GDoc): SpanFact[] {
  const out: SpanFact[] = [];
  for (const p of doc.paragraphs) {
    for (const el of p.elements) {
      out.push({ start: el.startIndex, end: el.endIndex, kind: el.kind, sizePt: el.fontSizePt });
    }
  }
  return out; // ascending by construction (parse preserves document order)
}

/** Binary search the fact covering `pos` (facts ascending + contiguous). */
function factAt(facts: SpanFact[], pos: number): SpanFact {
  let lo = 0;
  let hi = facts.length - 1;
  let found = -1;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    if (facts[mid].start <= pos) {
      found = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  if (found < 0 || pos >= facts[found].end) throw new Error(`no element fact covers index ${pos}`);
  return facts[found];
}

/** Walk [start, end) handing every covered ORIGINAL element fact to `visit`.
 * Element-granular (style facts are per-element uniform), so the walk is
 * O(elements touched), not O(chars). */
function forEachOriginalFact(
  facts: SpanFact[],
  start: number,
  end: number,
  visit: (fact: SpanFact, at: number) => void
): void {
  let pos = start;
  while (pos < end) {
    const fact = factAt(facts, pos);
    visit(fact, pos);
    pos = fact.end;
  }
}

/**
 * Build the ARMED view the hide plan would leave behind, purely in data
 * space: elements split at anchor boundaries, covered pieces at SENTINEL_PT,
 * anchors materialized as NamedRanges. This is deliberately NOT FakeDocs —
 * a pointer-walk split is O(elements + anchors), which keeps the xlarge
 * fixture's restore verification in seconds (see module header point 5).
 *
 * Two hide-plan invariants are ENFORCED while splitting (cheapest place to
 * see them): an anchor may never cover a non-text element (whitelist, plan
 * A9) and never cover text already AT a sentinel size (pre-existing tiny is
 * counted, not re-anchored — edge rows 8/12).
 */
function armedView(doc: GDoc, anchors: AnchorInfo[]): GDoc {
  const sorted = [...anchors].sort((a, b) => a.start - b.start);
  /** spanAt-style binary search: the anchor covering pos, or null. */
  const anchorAt = (pos: number): AnchorInfo | null => {
    let lo = 0;
    let hi = sorted.length - 1;
    let found = -1;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (sorted[mid].start <= pos) {
        found = mid;
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return found >= 0 && pos < sorted[found].end ? sorted[found] : null;
  };

  // Elements ascend across the whole body, so one cursor into the sorted
  // anchors serves every element (advanced, never rewound).
  let ai = 0;
  const paragraphs = doc.paragraphs.map((p) => ({
    ...p,
    elements: p.elements.flatMap((el): GElement[] => {
      while (ai < sorted.length && sorted[ai].end <= el.startIndex) ai++;
      // Cut points: every anchor edge strictly inside this element.
      const bounds: number[] = [el.startIndex];
      for (let j = ai; j < sorted.length && sorted[j].start < el.endIndex; j++) {
        for (const edge of [sorted[j].start, sorted[j].end]) {
          if (edge > el.startIndex && edge < el.endIndex && edge !== bounds[bounds.length - 1]) {
            bounds.push(edge);
          }
        }
      }
      bounds.push(el.endIndex);

      const pieces: GElement[] = [];
      for (let i = 0; i + 1 < bounds.length; i++) {
        const start = bounds[i];
        const end = bounds[i + 1];
        const covered = anchorAt(start) !== null;
        if (covered) {
          // The two enforced hide-plan invariants (see contract above).
          check(el.kind === "text", () => `anchor covers non-text element at ${start}`);
          check(
            el.fontSizePt === null || !SENTINELS.includes(el.fontSizePt),
            () => `anchor covers pre-existing sentinel text at ${start}`
          );
        }
        pieces.push({
          ...el,
          startIndex: start,
          endIndex: end,
          text: el.kind === "text" ? el.text.slice(start - el.startIndex, end - el.startIndex) : "",
          fontSizePt: covered ? SENTINEL_PT : el.fontSizePt
        });
      }
      return pieces;
    })
  }));

  const namedRanges: GNamedRange[] = anchors.map((a, i) => ({
    id: `armed-${i}`,
    name: a.name,
    segments: [{ startIndex: a.start, endIndex: a.end }]
  }));
  return { ...doc, paragraphs, namedRanges };
}

// ---------------------------------------------------------------------------
// Live round-trip helpers (small tier only)
// ---------------------------------------------------------------------------

/**
 * Parsed GDoc -> the GpSpec vocabulary FakeDocs is built from. The trailing
 * newline is stripped off each paragraph's final text element because
 * FakeDocs (like buildDoc) re-appends it — feeding it twice would shift every
 * index. "other" elements become same-width placeholders (width is all the
 * engine ever reads of them); the synth tool emits none, but the converter
 * must not silently mangle one if a future fixture carries chips.
 */
function toSpecs(doc: GDoc): GpSpec[] {
  return doc.paragraphs.map((p): GpSpec => {
    const elements: GeSpec[] = p.elements.map((el): GeSpec => {
      if (el.kind === "other") {
        return { text: "x".repeat(el.endIndex - el.startIndex), kind: "other" };
      }
      return { text: el.text, size: el.fontSizePt, bold: el.bold, bg: el.backgroundHex };
    });
    const last = elements[elements.length - 1];
    // The synth fixtures' segments-end-with-\n invariant is asserted in the
    // parse test; a violation here would mean discovery handed us junk.
    if (last !== undefined && (last.kind ?? "text") === "text" && last.text.endsWith("\n")) {
      last.text = last.text.slice(0, -1);
    }
    return {
      style: p.namedStyleType,
      spaceAbovePt: p.spaceAbovePt,
      spaceBelowPt: p.spaceBelowPt,
      elements
    };
  });
}

/**
 * Per-character projection of everything Hide may touch and Show All must
 * restore (size/bold/background per char + paragraph style/spacing). A local
 * port of the flagship suite's charView (gdocsInvisibility.test.ts) — kept
 * duplicated because this task may not edit the flagship file to export it,
 * and the projection is the load-bearing equality contract, so inlining it
 * beats importing across test files in a fragile way.
 */
function charProjection(doc: GDoc): string {
  return doc.paragraphs
    .map((p) => {
      const head = `${p.namedStyleType}(${p.spaceAbovePt ?? "i"}/${p.spaceBelowPt ?? "i"}):`;
      const chars = p.elements
        .map((el) => {
          if (el.kind === "other") return `<other:${el.endIndex - el.startIndex}>`;
          let out = "";
          for (let i = 0; i < el.text.length; i++) {
            const ch = el.text[i] === "\n" ? "¶" : el.text[i];
            out += `${ch}@${el.fontSizePt ?? "i"}${el.bold ? "b" : ""}${el.backgroundHex ?? ""};`;
          }
          return out;
        })
        .join("");
      return head + chars;
    })
    .join("\n");
}

// ---------------------------------------------------------------------------
// Discovery failure paths (run ALWAYS — no fixtures needed)
// ---------------------------------------------------------------------------

describe("scale-fixture discovery", () => {
  it("returns empty for a missing directory instead of throwing", () => {
    expect(discoverFixtures(path.join(os.tmpdir(), "gdocs-scale-does-not-exist"))).toEqual([]);
  });

  it("skips unparseable .json and ignores non-.json files with the suite intact", () => {
    // A scratch dir modeling local junk: one truncated fixture, one stray
    // markdown file (SCALE.md lives in the real dir), one valid fixture.
    const dir = fs.mkdtempSync(path.join(os.tmpdir(), "gdocs-scale-junk-"));
    try {
      fs.writeFileSync(path.join(dir, "truncated.json"), '{"revisionId": "x", "body": {');
      fs.writeFileSync(path.join(dir, "SCALE.md"), "| not | a | fixture |");
      fs.writeFileSync(path.join(dir, "[small] ok.json"), '{"revisionId": "ok"}');
      const found = discoverFixtures(dir);
      expect(found.map((f) => f.file)).toEqual(["[small] ok.json"]);
      expect(found[0].tier).toBe("small");
    } finally {
      fs.rmSync(dir, { recursive: true, force: true });
    }
  });
});

// ---------------------------------------------------------------------------
// The scale suite proper (skip-when-absent)
// ---------------------------------------------------------------------------

const fixtures = discoverFixtures(SYNTH_DIR);

describe("gdocs real-doc scale fixtures (samples/gdocs/)", () => {
  // Lesson #44: when the gitignored fixtures are absent (CI, fresh clones),
  // register ONE real placeholder test and return — never an empty it.each.
  if (fixtures.length === 0) {
    it("no synthetic fixtures present — run `node tools/gdocs-synth-fixture.mjs` to enable", () => {
      expect(fixtures).toHaveLength(0);
    });
    return;
  }

  /** Markdown rows collected per fixture, flushed to SCALE.md in afterAll. */
  const scaleRows: string[] = [];

  afterAll(() => {
    if (scaleRows.length === 0) return;
    const header = fs.existsSync(path.join(SYNTH_DIR, "SCALE.md"))
      ? ""
      : "# gdocs scale measurements (generated by gdocsRealDocs.test.ts — gitignored)\n";
    const table =
      `\n## ${new Date().toISOString()} — planHide / planShowAll envelope (${fixtures.length} fixtures)\n\n` +
      "| fixture | tier | fixture MB | paragraphs | elements | hide regions | hide requests | hide chunks | hide batch MB | showAll requests | showAll chunks | pre-existing tiny |\n" +
      "|---|---|---|---|---|---|---|---|---|---|---|---|\n" +
      scaleRows.join("\n") +
      `\n\nEnvelope (locked, plan A13/001-S6): < ${MAX_REQUESTS} requests, <= ${MAX_CHUNKS} chunks, ` +
      `< ${MAX_PAYLOAD_BYTES / (1024 * 1024)}MB payload per verb. ` +
      `"deferred" showAll columns = heavy-tier planShowAll replay needs a ROSTRUM_PERF=1 run (see HEAVY_TIERS).\n`;
    fs.appendFileSync(path.join(SYNTH_DIR, "SCALE.md"), header + table);
  });

  it.each(fixtures)(
    "$file [$tier]: parses, hide/restore plans hold their invariants, envelope holds",
    (ref: FixtureRef) => {
      // Phase timings ride into the log line — scale telemetry for the
      // wet-test packet (plan A13 wants a time-budget chunk model fed by
      // real numbers, not guesses).
      const t0 = Date.now();
      const raw: unknown = JSON.parse(fs.readFileSync(ref.fullPath, "utf8"));
      const doc = parseDocument(raw);
      const tParse = Date.now() - t0;

      // ---- 1. parse soundness -------------------------------------------
      expect(doc.paragraphs.length).toBeGreaterThan(0);
      expect(doc.tabCount).toBe(1);
      expect(doc.suggestionsPresent).toBe(false);
      let cursor = 1; // body content begins at index 1 behind the stub
      let elementCount = 0;
      for (const p of doc.paragraphs) {
        check(p.startIndex === cursor, () => `paragraph ${p.index} starts at ${p.startIndex}, expected ${cursor}`);
        check(p.endIndex > p.startIndex, () => `paragraph ${p.index} is empty-ranged`);
        for (const el of p.elements) {
          check(el.startIndex === cursor, () => `element gap at index ${cursor} (paragraph ${p.index})`);
          check(
            el.kind !== "text" || el.text.length === el.endIndex - el.startIndex,
            () => `text width mismatch at ${el.startIndex} (paragraph ${p.index})`
          );
          cursor = el.endIndex;
          elementCount++;
        }
        check(cursor === p.endIndex, () => `paragraph ${p.index} closes at ${p.endIndex}, walked to ${cursor}`);
        // Wire convention: the trailing newline rides INSIDE the final run.
        const lastEl = p.elements[p.elements.length - 1];
        check(
          lastEl !== undefined && lastEl.kind === "text" && lastEl.text.endsWith("\n"),
          () => `paragraph ${p.index} does not end in a newline-bearing text run`
        );
      }
      const docEnd = cursor;
      expect(docEnd).toBeGreaterThan(1); // the walk above really ran
      expect(doc.paragraphs[doc.paragraphs.length - 1].isLastInSegment).toBe(true);

      // ---- 2. planHide invariants ----------------------------------------
      const t1 = Date.now();
      const settings = resolveSettings(null, null);
      const { groups, result } = planHide(doc, settings);
      const tPlan = Date.now() - t1;
      expect(result.paragraphsScanned).toBe(doc.paragraphs.length);
      // Real debate docs always carry hideable card text; a zero here means
      // the keeper policy collapsed, not that the doc is all-kept.
      expect(result.paragraphsChanged).toBeGreaterThan(0);
      expect(result.regionsHidden).toBeGreaterThan(0);
      // Clean doc: nothing already hidden, nothing to resurface.
      expect(result.regionsAlreadyHidden).toBe(0);
      expect(result.newlyKeptRestored).toBe(0);
      // Fresh hide emits exactly one group per region.
      expect(groups.length).toBe(result.regionsHidden);

      const anchors: AnchorInfo[] = [];
      const shrinks: { start: number; end: number }[] = [];
      for (const group of groups) {
        let sawStyleWrite = false;
        let anchorCursor: number | null = null;
        for (const req of group.requests) {
          if ("createNamedRange" in req) {
            // Group shape: every anchor precedes the region's one style write.
            check(!sawStyleWrite, () => "anchor emitted AFTER its region's style write");
            const { name, range } = req.createNamedRange;
            // Names must decode (round-trip) and the RLE must account for
            // EXACTLY the chars the anchor covers — the restore-exactness
            // precondition (plan A2 / case 001-S4).
            check(name.length <= NAME_MAX, () => `anchor name overflows NAME_MAX: ${name.length} chars`);
            check(isRstmName(name), () => `anchor name not rstm-owned: ${name}`);
            const decoded = decodeRangeName(name);
            check(decoded !== null && decoded.kind === "sizes", () => `anchor name does not decode: ${name}`);
            if (decoded !== null && decoded.kind === "sizes") {
              const rleLen = decoded.entries.reduce((n, e) => n + e.count, 0);
              check(
                rleLen === range.endIndex - range.startIndex,
                () => `RLE length ${rleLen} != anchor width ${range.endIndex - range.startIndex} (${name})`
              );
              // Recorded sizes are ORIGINAL sizes — never the sentinel.
              for (const e of decoded.entries) {
                check(
                  e.sizePt === null || !SENTINELS.includes(e.sizePt),
                  () => `anchor records a sentinel size: ${name}`
                );
              }
            }
            // Split anchors of one region tile it contiguously.
            check(
              anchorCursor === null || range.startIndex === anchorCursor,
              () => `split anchors do not tile: gap before ${range.startIndex}`
            );
            anchorCursor = range.endIndex;
            anchors.push({ name, start: range.startIndex, end: range.endIndex });
          } else if ("updateTextStyle" in req) {
            // Exactly ONE style write per fresh region, fontSize-only, at the
            // sentinel, never touching the unstylable final newline.
            check(!sawStyleWrite, () => "second style write inside one fresh-hide group");
            sawStyleWrite = true;
            const { range, textStyle, fields } = req.updateTextStyle;
            check(fields === "fontSize", () => `hide write fields "${fields}" != "fontSize"`);
            check(
              textStyle.fontSize?.magnitude === SENTINEL_PT,
              () => `hide write magnitude ${textStyle.fontSize?.magnitude} != sentinel`
            );
            check(
              range.startIndex >= 1 && range.endIndex <= docEnd - 1,
              () => `hide write [${range.startIndex}, ${range.endIndex}) out of stylable bounds`
            );
            shrinks.push({ start: range.startIndex, end: range.endIndex });
            // The region's anchors tile exactly its shrink range.
            check(
              anchorCursor === range.endIndex,
              () => `anchors end at ${anchorCursor}, shrink ends at ${range.endIndex}`
            );
          } else {
            // collapseSpacing is OFF by default and the doc is clean, so any
            // other request type is a planner invariant break.
            throw new Error(`unexpected fresh-hide request type: ${Object.keys(req).join(",")}`);
          }
        }
        check(sawStyleWrite, () => "fresh-hide group carries no style write");
      }
      expect(anchors.length).toBeGreaterThanOrEqual(groups.length); // >= 1 anchor per region

      // Anchors never overlap (the enforced rstm invariant), and the anchor
      // union IS the shrink union — nothing shrinks unrecorded, nothing is
      // recorded unshrunk.
      const sortedAnchors = [...anchors].sort((a, b) => a.start - b.start);
      for (let i = 1; i < sortedAnchors.length; i++) {
        check(
          sortedAnchors[i].start >= sortedAnchors[i - 1].end,
          () => `anchors overlap at ${sortedAnchors[i].start}`
        );
      }
      expect(mergedSpans(anchors)).toEqual(mergedSpans(shrinks));

      // ---- 3. chunking + the locked envelope -----------------------------
      const chunks = chunkGroups(groups);
      const hideRequests = chunks.reduce((n, c) => n + c.length, 0);
      expect(hideRequests).toBe(requestsOf(groups).length); // chunker drops nothing
      for (const chunk of chunks) {
        check(chunk.length <= CHUNK_MAX, () => `chunk of ${chunk.length} requests exceeds CHUNK_MAX`);
      }
      const hideBatchBytes = chunks.reduce(
        (n, c) => n + Buffer.byteLength(JSON.stringify({ requests: c })),
        0
      );

      // ---- 4a. restore-RECORD exactness (every tier — fast) ----------------
      // Each anchor's RLE must describe the ORIGINAL document byte-for-byte:
      // walking every entry against the pre-hide element facts proves the
      // in-document manifest alone suffices for an exact restore (case
      // 001-S4), independently of any planner bookkeeping.
      const facts = spanFactsOf(doc);
      let recordedChars = 0;
      for (const a of anchors) {
        const decoded = decodeRangeName(a.name);
        if (decoded === null || decoded.kind !== "sizes") continue; // already failed above
        let entryStart = a.start;
        for (const e of decoded.entries) {
          const entryEnd = entryStart + e.count;
          forEachOriginalFact(facts, entryStart, entryEnd, (fact, at) => {
            // The A9 whitelist, restated as a record property: only textRuns
            // are ever anchored, so a record over a chip is corruption.
            if (fact.kind !== "text") {
              throw new Error(`anchor covers a non-text element at ${at} (${a.name})`);
            }
            if (fact.sizePt !== e.sizePt) {
              throw new Error(`RLE records ${e.sizePt} but original size at ${at} is ${fact.sizePt} (${a.name})`);
            }
          });
          entryStart = entryEnd;
          recordedChars += e.count;
        }
      }
      // Records cover EXACTLY the anchored chars — nothing hidden unrecorded.
      expect(recordedChars).toBe(anchors.reduce((n, a) => n + (a.end - a.start), 0));

      // ---- 4b. planShowAll replay (deferred on heavy tiers — HEAVY_TIERS) --
      const deepRestore = RUN_HEAVY || !HEAVY_TIERS.includes(ref.tier);
      let showRequests: number | null = null;
      let showChunks: number | null = null;
      let restoreMs = "deferred (set ROSTRUM_PERF=1)";
      if (deepRestore) {
        const t2 = Date.now();
        const armed = armedView(doc, anchors);
        const tArm = Date.now() - t2;
        const t3 = Date.now();
        const show = planShowAll(armed, /* sweepUnrecorded */ false);
        const tShow = Date.now() - t3;
        const t4 = Date.now();

        // Every anchor is intact (one segment), so every segment restores
        // EXACTLY; the sweep adopts precisely the pre-existing tiny passages
        // planHide counted through its own, independent geometry.
        expect(show.result.segmentsRestoredExact).toBe(anchors.length);
        expect(show.result.segmentsNormalized).toBe(0);
        expect(show.result.rangesDeleted).toBe(anchors.length);
        expect(show.result.rangesSkippedNewerVersion).toBe(0);
        expect(show.result.sweptOrphans).toBe(result.preexistingTinyCount);

        const anchorUnion = mergedSpans(anchors);
        /** spanAt over the merged union: is pos inside any anchor? */
        const inAnchors = (pos: number): boolean => {
          let lo = 0;
          let hi = anchorUnion.length - 1;
          let found = -1;
          while (lo <= hi) {
            const mid = (lo + hi) >> 1;
            if (anchorUnion[mid].start <= pos) {
              found = mid;
              lo = mid + 1;
            } else {
              hi = mid - 1;
            }
          }
          return found >= 0 && pos < anchorUnion[found].end;
        };

        let restoredChars = 0;
        const deletedIds = new Set<string>();
        showRequests = 0;
        for (const req of requestsOf(show.groups)) {
          showRequests++;
          if ("deleteNamedRange" in req) {
            deletedIds.add(req.deleteNamedRange.namedRangeId);
          } else if ("updateTextStyle" in req) {
            const { range, textStyle, fields } = req.updateTextStyle;
            check(fields === "fontSize", () => `showAll write fields "${fields}" != "fontSize"`);
            const written = textStyle.fontSize?.magnitude ?? null;
            if (inAnchors(range.startIndex)) {
              // RESTORE write: every covered char must come back at its
              // ORIGINAL explicit size (null = cleared back to inherit) —
              // verified against the pre-hide document, not the plan.
              restoredChars += range.endIndex - range.startIndex;
              forEachOriginalFact(facts, range.startIndex, range.endIndex, (fact, at) => {
                if (fact.sizePt !== written) {
                  throw new Error(`restore mismatch at ${at}: original ${fact.sizePt} vs written ${written}`);
                }
              });
            } else {
              // SWEEP write: clear-to-inherit, and only ever over text that
              // was ALREADY at a sentinel size before the hide (user tiny).
              check(written === null, () => `sweep write materializes ${written} instead of clearing`);
              forEachOriginalFact(facts, range.startIndex, range.endIndex, (fact, at) => {
                if (fact.sizePt === null || !SENTINELS.includes(fact.sizePt)) {
                  throw new Error(`sweep over non-tiny text at ${at} (original size ${fact.sizePt})`);
                }
              });
            }
          } else {
            throw new Error(`unexpected showAll request type: ${Object.keys(req).join(",")}`);
          }
        }
        // Every armed anchor dies, and the restore writes cover the ENTIRE
        // hidden area — nothing under an anchor stays sentinel after Show All.
        // (Manual containment, not toEqual on Sets: jest's unordered-collection
        // equality is quadratic, and there are ~27k ids on the xlarge fixture.)
        for (const nr of armed.namedRanges) {
          check(deletedIds.has(nr.id), () => `armed range ${nr.id} (${nr.name}) was never deleted`);
        }
        expect(deletedIds.size).toBe(armed.namedRanges.length);
        expect(restoredChars).toBe(anchorUnion.reduce((n, s) => n + (s.end - s.start), 0));

        showChunks = chunkGroups(show.groups).length;
        restoreMs = `arm ${tArm}, planShow ${tShow}, verify ${Date.now() - t4}`;

        // showAll's half of the locked envelope (only measurable here).
        expect(showRequests).toBeLessThan(MAX_REQUESTS);
        expect(showChunks).toBeLessThanOrEqual(MAX_CHUNKS);

        // ---- armed RE-HIDE: the A1 reconcile through buildAtoms AT SCALE -----
        // planHide above ran on the CLEAN doc; this is the only place the armed
        // reconcile (the interactive mainline — every Hide after the first) is
        // exercised at fixture scale. It also covers buildAtoms' cut-collection
        // on a doc whose coverage is non-empty, guarding against a quadratic
        // regression on long docs. Re-hiding an already-hidden doc is a no-op:
        // it SEES the hidden regions, resurfaces nothing, and emits zero work.
        const reHide = planHide(armed, settings);
        expect(reHide.result.regionsAlreadyHidden).toBeGreaterThan(0);
        expect(reHide.result.regionsHidden).toBe(0);
        expect(reHide.result.newlyKeptRestored).toBe(0);
        expect(reHide.groups.length).toBe(0);
      }

      // ---- the LOCKED envelope (report numbers, never weaken) -------------
      expect(hideRequests).toBeLessThan(MAX_REQUESTS);
      expect(chunks.length).toBeLessThanOrEqual(MAX_CHUNKS);
      expect(hideBatchBytes).toBeLessThan(MAX_PAYLOAD_BYTES);
      expect(ref.bytes).toBeLessThan(MAX_PAYLOAD_BYTES);

      // ---- 5. report (expect-and-log + SCALE.md row) ----------------------
      const mb = (n: number): string => (n / (1024 * 1024)).toFixed(2);
      scaleRows.push(
        `| ${ref.file} | ${ref.tier} | ${mb(ref.bytes)} | ${doc.paragraphs.length} | ${elementCount} | ` +
          `${result.regionsHidden} | ${hideRequests} | ${chunks.length} | ${mb(hideBatchBytes)} | ` +
          `${showRequests ?? "deferred"} | ${showChunks ?? "deferred"} | ${result.preexistingTinyCount} |`
      );
      // eslint-disable-next-line no-console
      console.log(
        `[gdocs-scale ${ref.tier}] ${ref.file}: ${doc.paragraphs.length} paras, ${elementCount} elements, ` +
          `fixture ${mb(ref.bytes)}MB | hide: ${result.regionsHidden} regions -> ${hideRequests} requests, ` +
          `${chunks.length} chunks, ${mb(hideBatchBytes)}MB | showAll: ${showRequests ?? "deferred"} requests, ` +
          `${showChunks ?? "deferred"} chunks | pre-existing tiny: ${result.preexistingTinyCount} | ` +
          `ms: parse ${tParse}, plan ${tPlan}, ${restoreMs}`
      );
    },
    300000
  );

  // ---- live round trip: SMALL tier only (see module header point 5) -------
  const smallFixtures = fixtures.filter((f) => f.tier === "small");
  if (smallFixtures.length === 0) {
    it("no small-tier fixture — live FakeDocs round trip not exercised", () => {
      expect(smallFixtures).toHaveLength(0);
    });
  } else {
    it.each(smallFixtures)(
      "$file: live hide -> showAll round trip through the controller is view-exact",
      async (ref: FixtureRef) => {
        const doc = parseDocument(JSON.parse(fs.readFileSync(ref.fullPath, "utf8")));
        const fake = new FakeDocs(toSpecs(doc));
        const parsedFake = parseDocument(await fake.fetchDocument());
        const before = charProjection(parsedFake);
        // Strict view-equality is only promised for docs with NO pre-existing
        // sentinel-size text: Show All's sweep deliberately normalizes user
        // tiny text on armed docs (edge row 11, documented limitation), which
        // would be a real, intended difference. The current small fixture has
        // none — this guard makes the test fail LOUDLY and explainably if a
        // future fixture changes that, instead of mysteriously diffing.
        const dry = planHide(parsedFake, resolveSettings(null, null));
        expect(dry.result.preexistingTinyCount).toBe(0);

        const hideResult = await hide(fake);
        expect(hideResult.paragraphsChanged).toBeGreaterThan(0);
        // The doc is genuinely armed mid-flight: sentinel text + rstm anchors.
        const mid = parseDocument(await fake.fetchDocument());
        expect(mid.namedRanges.some((nr) => isRstmName(nr.name))).toBe(true);
        expect(charProjection(mid)).not.toBe(before);

        const outcome = await showAll(fake);
        expect(outcome.kind).toBe("done"); // armed doc never needs sweep consent
        if (outcome.kind === "done") {
          // Real-doc records are intact, so every passage restores exactly.
          expect(outcome.result.segmentsNormalized).toBe(0);
        }

        const after = parseDocument(await fake.fetchDocument());
        // The 001-S4 promise at real-doc scale: every character reads back
        // with its original size/bold/background, and no rstm state remains.
        expect(charProjection(after)).toBe(before);
        expect(after.namedRanges.some((nr) => isRstmName(nr.name))).toBe(false);
      },
      300000
    );
  }
});
