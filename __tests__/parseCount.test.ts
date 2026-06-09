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
import { readRuns } from "../src/core/ooxml";
import { settings } from "./fakeWord";
import { RawParagraph } from "../src/core/types";

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
