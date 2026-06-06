// Unit tests for the Condense engine (src/core/condense.ts): the mode/reversal policy AND the lossless
// guarantee — `uncondense ∘ condense == identity` across the marker + pilcrow modes (the part that
// beats Verbatim), including a divergent-pPr preservation case, plus the destructive negative case.
import { condenseFragment, resolveCondenseOptions, uncondenseFragment } from "../src/core/condense";
import { readFragmentParagraphs } from "../src/core/ooxmlCondense";
import { CondenseOptions } from "../src/core/types";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

const run = (text: string): string => `<w:r><w:t xml:space="preserve">${text}</w:t></w:r>`;
const p = (inner: string, pPr = ""): string => `<w:p>${pPr}${inner}</w:p>`;
const body = (...ps: string[]): string => `<w:body xmlns:w="${W_NS}">${ps.join("")}</w:body>`;
const paraTexts = (xml: string): string[] =>
  readFragmentParagraphs(xml).map((runs) => runs.map((r) => r.text).join(""));
/** Each `<w:p>…</w:p>` block (paragraphs don't nest, so a non-greedy match is exact). */
const paraBlocks = (xml: string): string[] => xml.match(/<w:p>[\s\S]*?<\/w:p>/g) ?? [];

const opt = (over: Partial<CondenseOptions> = {}): CondenseOptions => ({
  usePilcrows: false,
  retainParagraphs: false,
  reversal: "marker",
  ...over
});

describe("resolveCondenseOptions", () => {
  it("coerces destructive (reversal none) to marker when pilcrows are on (a pilcrow IS a marker)", () => {
    expect(resolveCondenseOptions(opt({ usePilcrows: true, reversal: "none" })).reversal).toBe("marker");
  });
  it("honors destructive when pilcrows are off", () => {
    expect(resolveCondenseOptions(opt({ usePilcrows: false, reversal: "none" })).reversal).toBe("none");
  });
});

describe("lossless guarantee: uncondense ∘ condense == identity", () => {
  it("full-merge (hidden-space markers) round-trips the paragraph structure + text", () => {
    const xml = body(p(run("Card one body.")), p(run("Card two body.")), p(run("Card three.")));
    const condensed = condenseFragment(xml, opt());
    expect(paraTexts(condensed.xml)).toHaveLength(1); // merged to one
    const restored = uncondenseFragment(condensed.xml);
    expect(paraTexts(restored.xml)).toEqual(["Card one body.", "Card two body.", "Card three."]);
  });

  it("pilcrow mode round-trips identically", () => {
    const xml = body(p(run("Alpha")), p(run("Beta")));
    const condensed = condenseFragment(xml, opt({ usePilcrows: true }));
    expect(paraTexts(condensed.xml)).toEqual(["Alpha¶Beta"]);
    const restored = uncondenseFragment(condensed.xml);
    expect(paraTexts(restored.xml)).toEqual(["Alpha", "Beta"]);
  });

  it("preserves a divergent paragraph's pPr through the round-trip (sidecar-free, self-describing)", () => {
    // Paragraph 2 carries an indent the others lack — Condense must store it and Uncondense restore it.
    const xml = body(p(run("AAA")), p(run("BBB"), `<w:pPr><w:ind w:left="720"/></w:pPr>`));
    const condensed = condenseFragment(xml, opt());
    const restored = uncondenseFragment(condensed.xml);
    expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB"]);
    const blocks = paraBlocks(restored.xml);
    expect(blocks).toHaveLength(2);
    expect(blocks[0]).toContain("AAA");
    expect(blocks[0]).not.toContain('w:left="720"'); // first paragraph keeps its (no-indent) pPr
    expect(blocks[1]).toContain("BBB");
    expect(blocks[1]).toContain('w:left="720"'); // the indent is restored onto the SECOND paragraph
  });

  it("restores an empty paragraph that sat between two cards", () => {
    const xml = body(p(run("AAA")), p(""), p(run("BBB")));
    const condensed = condenseFragment(xml, opt());
    const restored = uncondenseFragment(condensed.xml);
    expect(paraTexts(restored.xml)).toEqual(["AAA", "", "BBB"]);
  });

  it("survives a RE-condense without eating markers (adversarial marker-eating)", () => {
    // A card body ending in a space, then Condense pressed twice — the space-glyph markers must NOT be
    // collapsed away on the second pass, or the breaks become unrecoverable.
    const xml = body(p(run("Card one body. ")), p(run("Card two.")));
    const first = condenseFragment(xml, opt());
    const second = condenseFragment(first.xml, opt()); // re-condense the already-merged fragment
    const restored = uncondenseFragment(second.xml);
    expect(paraTexts(restored.xml)).toEqual(["Card one body. ", "Card two."]);
  });

  it("does NOT leak the first paragraph's pPr onto a later paragraph that had none (adversarial pPr)", () => {
    // Para 1 is centered; para 2 has no pPr. After the round-trip, para 2 must stay default (NOT centered).
    const xml = body(p(run("AAA"), `<w:pPr><w:jc w:val="center"/></w:pPr>`), p(run("BBB")));
    const condensed = condenseFragment(xml, opt());
    const restored = uncondenseFragment(condensed.xml);
    expect(paraTexts(restored.xml)).toEqual(["AAA", "BBB"]);
    const blocks = paraBlocks(restored.xml);
    expect(blocks[0]).toContain('w:val="center"'); // first paragraph keeps its centering
    expect(blocks[1]).not.toContain('w:val="center"'); // second paragraph does NOT inherit it
  });
});

describe("destructive mode (reversal none) is one-way", () => {
  it("merges with a plain space and is NOT reversible (no markers to restore)", () => {
    const xml = body(p(run("AAA")), p(run("BBB")));
    const condensed = condenseFragment(xml, opt({ reversal: "none" }));
    expect(paraTexts(condensed.xml)).toEqual(["AAA BBB"]);
    const restored = uncondenseFragment(condensed.xml);
    expect(restored.breaksRestored).toBe(0); // nothing to restore
    expect(paraTexts(restored.xml)).toEqual(["AAA BBB"]); // still merged
  });
});

describe("retain-paragraphs mode round-trips blank-line removal", () => {
  it("uncondense restores blank lines dropped by lossless retain mode", () => {
    const xml = body(p(run("AAA")), p("  "), p(run("BBB")));
    const condensed = condenseFragment(xml, opt({ retainParagraphs: true }));
    expect(condensed.boundariesMarked).toBe(1);
    const restored = uncondenseFragment(condensed.xml);
    // structure intact (3 paragraphs throughout); the blank paragraph's hidden mark is restored.
    expect(paraTexts(restored.xml)).toEqual(["AAA", "", "BBB"]);
  });
});
