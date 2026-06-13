// PERFORMANCE-INVARIANT GUARD — classifyParagraph parses each paragraph exactly ONCE.
//
// The hide pass is parse-bound: xmldom's `DOMParser` (string → full DOM tree) dominates
// the engine's per-paragraph cost on a long doc, and the same CPU is spent live in the
// task-pane browser. The engine USED to parse every paragraph TWICE — once in
// `readRuns` to classify, then again in `applyRunVisibility`/`makeAllVisible` to mutate —
// roughly doubling the work. `ParsedParagraph` collapsed that to one parse (read + mutate
// through one tree, byte-identical output).
//
// This test LOCKS THAT IN. It spies on the real xmldom `DOMParser`, counts
// `parseFromString` calls during a single `classifyParagraph`, and asserts the count is
// exactly 1 for every paragraph shape (heading / body-partial / body-all-hidden /
// bridge-split / table / image). A future change that reintroduces a second parse fails
// HERE — a non-flaky guard (a call COUNT, not a wall-clock threshold). It also asserts the
// classification is still CORRECT on each shape, so it can't pass by parsing once but
// producing the wrong OOXML.
//
// A second block locks the SAME invariant onto Shrink: `shrinkFragment`/`unshrinkFragment`
// used to parse the fragment TWICE (readFragmentParagraphs to read, applyFragmentShrink to
// mutate — over the identical unchanged string). `parseFragment` fused that to one parse,
// threaded through both calls; a byte-identity test pins the fused path to the still-callable
// legacy path so the equivalence stays assertable forever.
//
// A third block locks it onto `WholeBodyPackage.replace` — the write half of the DEFAULT
// pure whole-body hide commit (one replace per changed paragraph) and of cite repair.
// replace() used to call the string-in `assertSingleParagraph` (which parses) and then
// parse the IDENTICAL fragment again to select the `<w:p>`; `singleBodyParagraph` fused
// guard + selection onto one parsed tree, halving the splice loop's DOMParser work.

/* eslint-disable @typescript-eslint/no-explicit-any */

// Count every DOMParser.parseFromString while delegating to the genuine xmldom impl, so
// the engine's real parsing/serialization is exercised (this is a perf guard, not a stub).
let parseCount = 0;
jest.mock("@xmldom/xmldom", () => {
  const actual = jest.requireActual("@xmldom/xmldom");
  class CountingDOMParser {
    private readonly inner: any;
    constructor(opts?: any) {
      this.inner = new actual.DOMParser(opts);
    }
    parseFromString(source: string, mimeType: string): any {
      parseCount++;
      return this.inner.parseFromString(source, mimeType);
    }
  }
  return { ...actual, DOMParser: CountingDOMParser };
});

import { classifyParagraph } from "../src/core/invisibility";
import { DOMParser } from "@xmldom/xmldom";
import { ParsedParagraph, readRuns, legacyRunViewsForTest } from "../src/core/ooxml";
import { computeRunKeepFlags, planCrossGapSeparators } from "../src/core/keepers";
import {
  ParagraphShrinkPlan,
  applyFragmentShrink,
  parseFragment,
  readFragmentParagraphs
} from "../src/core/ooxmlCondense";
import { WholeBodyPackage } from "../src/core/ooxmlPackage";
import { shrinkFragment, unshrinkFragment } from "../src/core/shrink";
import { settings } from "./fakeWord";
import { RawParagraph, ShrinkOptions } from "../src/core/types";

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const KEEP = settings(["yellow", "cyan", "green"]);

/** A bare `<w:p>` fragment (classifyParagraph/readRuns scope to the first paragraph). */
const para = (inner: string): string => `<w:document xmlns:w="${W}"><w:body><w:p>${inner}</w:p></w:body></w:document>`;
const run = (text: string, rPr = ""): string =>
  `<w:r>${rPr ? `<w:rPr>${rPr}</w:rPr>` : ""}<w:t xml:space="preserve">${text}</w:t></w:r>`;
const hl = (color: string): string => `<w:highlight w:val="${color}"/>`;

/** Run `classifyParagraph` once and report how many DOMParser parses it triggered. */
function parsesFor(p: RawParagraph): { parses: number; result: ReturnType<typeof classifyParagraph> } {
  parseCount = 0;
  const result = classifyParagraph(p, KEEP);
  return { parses: parseCount, result };
}

/** Concatenated visible (non-hidden) text of a classified paragraph's OOXML. */
function visibleText(xml: string): string {
  return readRuns(xml)
    .filter((r) => !r.hidden)
    .map((r) => r.text)
    .join("");
}

describe("classifyParagraph parses each paragraph exactly once (perf invariant)", () => {
  it("heading paragraph — one parse, kept whole", () => {
    const p: RawParagraph = { index: 0, headingLevel: 0, inTable: false, ooxml: para(run("Contention One")) };
    const { parses, result } = parsesFor(p);
    expect(parses).toBe(1);
    expect(result.action).toBe("keepWhole");
    expect(visibleText(result.ooxml)).toBe("Contention One");
  });

  it("body paragraph with a highlighted keeper — one parse, only the keeper stays visible", () => {
    const p: RawParagraph = {
      index: 1,
      headingLevel: null,
      inTable: false,
      ooxml: para(run("rising ", hl("yellow")) + run("oil prices fall and ") + run("revenues", hl("yellow")))
    };
    const { parses, result } = parsesFor(p);
    expect(parses).toBe(1);
    expect(result.action).toBe("hidePartial");
    // The hidden middle clause is gone; the two highlighted chunks survive.
    const vis = visibleText(result.ooxml);
    expect(vis).toContain("rising");
    expect(vis).toContain("revenues");
    expect(vis).not.toContain("oil prices fall");
  });

  it("body paragraph fully hidden — one parse, nothing visible", () => {
    const p: RawParagraph = {
      index: 2,
      headingLevel: null,
      inTable: false,
      ooxml: para(run("entirely unhighlighted card body text"))
    };
    const { parses, result } = parsesFor(p);
    expect(parses).toBe(1);
    expect(result.action).toBe("hideWhole");
    expect(visibleText(result.ooxml)).toBe("");
  });

  it("table paragraph — one parse, kept untouched", () => {
    const p: RawParagraph = { index: 3, headingLevel: null, inTable: true, ooxml: para(run("cell body")) };
    const { parses, result } = parsesFor(p);
    expect(parses).toBe(1);
    expect(result.action).toBe("keepWhole");
    expect(visibleText(result.ooxml)).toBe("cell body");
  });

  it("inline-image paragraph — one parse, kept whole (never an OOXML write that dangles the part)", () => {
    const p: RawParagraph = {
      index: 4,
      headingLevel: null,
      inTable: false,
      ooxml: para(`<w:r><w:drawing><wp:inline xmlns:wp="x"/></w:drawing></w:r>` + run("caption"))
    };
    const { parses, result } = parsesFor(p);
    expect(parses).toBe(1);
    expect(result.action).toBe("keepWhole");
  });

  it("never parses more than once even when a bridge split inserts new runs", () => {
    // Two highlighted chunks with a hidden whitespace-bearing gap between them: the planner
    // exposes a separator by splitting a run on the SAME parsed tree — proving the mutate
    // path (including run insertion) adds no extra parse.
    const p: RawParagraph = {
      index: 5,
      headingLevel: null,
      inTable: false,
      ooxml: para(run("gives", hl("cyan")) + run(" Russia leverage over ") + run("Europe", hl("cyan")))
    };
    const { parses, result } = parsesFor(p);
    expect(parses).toBe(1);
    expect(result.action).toBe("hidePartial");
    const vis = visibleText(result.ooxml);
    expect(vis).not.toContain("Russia leverage over");
    // Kept words stay apart (a moved space survives) — never fused into "givesEurope".
    expect(vis).not.toContain("givesEurope");
    expect(vis).toContain("gives");
    expect(vis).toContain("Europe");
  });
});

describe("Shrink parses the fragment exactly once per press (perf invariant)", () => {
  // A multi-paragraph fragment (readFragmentParagraphs/applyFragmentShrink operate on EVERY <w:p>,
  // unlike classifyParagraph's single-paragraph scope).
  const frag = (...paras: string[]): string =>
    `<w:document xmlns:w="${W}"><w:body>${paras.map((p) => `<w:p>${p}</w:p>`).join("")}</w:body></w:document>`;

  // Representative multi-run shape for a press: an underlined keeper (the cut), an explicitly-sized
  // plain body run, a highlighted keeper, and a second body paragraph. The style/Normal resolvers
  // are regex-only, so the DOMParser count below isolates exactly the read+apply parses.
  const fixture = frag(
    run("kept cut ", `<w:u w:val="single"/>`) +
      run("plain body text ", `<w:sz w:val="22"/>`) +
      run("tagline", hl("yellow")),
    run("second paragraph body", `<w:sz w:val="22"/>`)
  );
  const opts: ShrinkOptions = {
    normalHalfPts: 22,
    outlineLevels: [null, null],
    omissionPatterns: [],
    shrinkParagraphMarks: false
  };

  it("shrinkFragment — one parse, and the press still lands (8pt rung, keepers untouched)", () => {
    parseCount = 0;
    const result = shrinkFragment(fixture, opts);
    expect(parseCount).toBe(1); // read + apply share ONE tree (was 2 before the parseFragment fusion)
    expect(result.changed).toBe(true);
    expect(result.appliedSizeHalfPts).toBe(16); // 11pt body → the 8pt rung
    const [p1, p2] = readFragmentParagraphs(result.xml);
    expect(p1[0].sizeHalfPts).toBeNull(); // underlined keeper: no explicit size added
    expect(p1[1].sizeHalfPts).toBe(16); // body run shrunk to the rung
    expect(p1[2].sizeHalfPts).toBeNull(); // highlighted keeper: untouched
    expect(p2[0].sizeHalfPts).toBe(16); // second paragraph's body shrunk too
  });

  it("unshrinkFragment — one parse, explicit sizes cleared back to Normal", () => {
    const pressed = shrinkFragment(fixture, opts).xml;
    parseCount = 0;
    const result = unshrinkFragment(pressed, [null, null]);
    expect(parseCount).toBe(1);
    expect(result.changed).toBe(true);
    // Every run reverts to inherited Normal (no explicit <w:sz> anywhere).
    const runs = readFragmentParagraphs(result.xml).flat();
    expect(runs.length).toBe(4);
    expect(runs.every((r) => r.sizeHalfPts === null)).toBe(true);
  });

  it("the fused (pre-parsed) path is byte-identical to the legacy two-parse path", () => {
    // Read side: identical run views with and without the handle.
    expect(readFragmentParagraphs(fixture, parseFragment(fixture))).toEqual(readFragmentParagraphs(fixture));
    // Apply side: byte-identical XML out (a fresh handle per call — apply consumes its tree).
    const plans: ParagraphShrinkPlan[] = [
      { runSizes: [undefined, 16, undefined] },
      { runSizes: [null], markSizeHalfPts: 12 }
    ];
    const legacy = applyFragmentShrink(fixture, plans);
    const fused = applyFragmentShrink(fixture, plans, parseFragment(fixture));
    expect(fused.xml).toBe(legacy.xml);
    expect(fused.changed).toBe(legacy.changed);
  });
});

describe("WholeBodyPackage.replace parses the fragment exactly once (perf invariant)", () => {
  // The DEFAULT whole-body hide commit calls replace() once per CHANGED paragraph (489 on
  // the ExFlex realDocs sample, 9k+ on the xlarge one), so a stray second parse here scales
  // with document size exactly like the classify/Shrink regressions the blocks above pin.
  const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";
  const bodyP = (text: string): string => `<w:p><w:r><w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;
  /** A minimal flat-OPC package like body.getOoxml() returns (document part + body paras). */
  const bodyPkg = (...paras: string[]): string =>
    `<pkg:package xmlns:pkg="${PKG_NS}">` +
    `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
    `<w:document xmlns:w="${W}"><w:body>${paras.join("")}<w:sectPr/></w:body></w:document>` +
    `</pkg:xmlData></pkg:part></pkg:package>`;

  it("happy-path splice — ONE parse, and the edit still lands", () => {
    const wb = new WholeBodyPackage(bodyPkg(bodyP("keep me"), bodyP("change me")));
    const edited = `<w:p xmlns:w="${W}"><w:r><w:rPr><w:vanish/></w:rPr><w:t>change me</w:t></w:r></w:p>`;
    parseCount = 0;
    wb.replace(1, edited);
    // Guard + node selection share ONE tree (was 2: assertSingleParagraph re-parsed the
    // identical string the next line parsed again).
    expect(parseCount).toBe(1);
    const out = wb.serialize();
    expect(out).toContain("<w:vanish/>"); // the hide landed…
    expect(out).toContain("keep me"); // …and the sibling survived
  });

  it("multi-paragraph fragment — still rejected with the same guard error, after ONE parse", () => {
    const wb = new WholeBodyPackage(bodyPkg(bodyP("x")));
    const two = `<w:body xmlns:w="${W}">${bodyP("a")}${bodyP("b")}</w:body>`;
    parseCount = 0;
    // Context + count survive the fusion verbatim (officeWordPort matches on this shape).
    expect(() => wb.replace(0, two)).toThrow(/whole-body splice @0.*found 2/s);
    expect(parseCount).toBe(1); // the guard no longer pays its own separate parse
  });

  it("out-of-range index — RangeError BEFORE any parse (zero parses)", () => {
    const wb = new WholeBodyPackage(bodyPkg(bodyP("only")));
    parseCount = 0;
    expect(() => wb.replace(5, bodyP("x"))).toThrow(RangeError);
    expect(parseCount).toBe(0); // index check stays first — bad indices never cost a parse
  });
});

// ---------------------------------------------------------------------------
// 002-S1 SUB-GATE (tests-4): the NODE-DIRECT path parses NOTHING.
//
// The compat shim `p.parsed ?? new ParsedParagraph(p.ooxml)` makes every fixture WITHOUT `.parsed`
// take the string branch → still exactly one parse → the blocks above stay green while `fromNode` is
// NEVER hit. So those blocks alone do NOT prove the Loop-002 win. This block exercises the node path
// directly: parse a package ONCE (the cost the node path is meant to PAY ONCE for the whole body, not
// per paragraph), then build `ParsedParagraph.fromNode` over each live <w:p> and run
// `applyVisibilityInPlace` — and assert ZERO further `parseFromString` calls. 002-S1 names exactly this
// ("parseCount===0 on the node-direct path") as the falsifiable proof the win is real. It also asserts
// the node read is six-field-(+hasInternalPart)-identical to the string-path read, so a zero-parse path
// that produced WRONG views could never pass.
// ---------------------------------------------------------------------------
describe("node-direct path (fromNode + applyVisibilityInPlace) parses ZERO times (002-S1 sub-gate)", () => {
  /** Parse a fragment with the same fatal-only policy the engine uses (counted by the spy). */
  const parseFrag = (xml: string): any =>
    new DOMParser({
      onError: (level: string, message: string) => {
        if (level === "fatalError") throw new Error(message);
      }
    }).parseFromString(xml, "text/xml");
  /** The body's first <w:p>, mirroring ooxml.ts firstParagraph scoping. */
  const firstP = (doc: any): any => {
    const bodies = doc.getElementsByTagName("w:body");
    const scope = bodies && bodies.length > 0 ? bodies.item(0) : doc;
    const ps = scope.getElementsByTagName("w:p");
    return ps && ps.length > 0 ? ps.item(0) : null;
  };

  it("fromNode reads + applyVisibilityInPlace mutate with parseCount === 0 after the one setup parse", () => {
    // A multi-run body paragraph: two highlighted keepers around a hidden gap (the bridge-split shape).
    const xml = para(run("gives", hl("cyan")) + run(" Russia leverage over ") + run("Europe", hl("cyan")));
    const doc = parseFrag(xml); // the ONE expected parse — counted, then reset below
    const pEl = firstP(doc);

    parseCount = 0; // from here the node path must add nothing
    const pp = ParsedParagraph.fromNode(doc, pEl);
    expect(parseCount).toBe(0); // fromNode parses nothing

    const keep = computeRunKeepFlags(pp.runs, KEEP.keepColors);
    const { extraKeep, splits } = planCrossGapSeparators(pp.runs, keep);
    for (const i of extraKeep) keep[i] = true;
    const res = pp.applyVisibilityInPlace({ hideFlags: keep.map((k) => !k), hideParaMark: false, splits });

    expect(parseCount).toBe(0); // applyVisibilityInPlace (incl. bridge split) parses nothing
    expect(res.changed).toBe(true); // and the edit still landed
    // Six-field identity vs the string path (proves zero-parse didn't mean wrong/empty views).
    const stringView = legacyRunViewsForTest(xml);
    // Re-read fresh node views from a clean parse to compare pre-mutation reads (the string path reads
    // the unmutated input, so compare against a fresh fromNode over a fresh parse of the same input).
    const docFresh = parseFrag(xml);
    expect(ParsedParagraph.fromNode(docFresh, firstP(docFresh)).runs).toEqual(stringView);
  });

  it("the string ctor still parses exactly once (compat invariant unchanged by the node path)", () => {
    const xml = para(run("entirely unhighlighted card body text"));
    parseCount = 0;
    const pp = new ParsedParagraph(xml); // string mode → one parse
    expect(parseCount).toBe(1);
    expect(pp.runs).toHaveLength(1);
  });
});
