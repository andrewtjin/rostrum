// gdocs parse suite (plan S3 + A3/A9/A11.ii/A13; edge rows 1, 2, 14, 15, 17).
// Every assert is BEHAVIORAL — fixtures are real-shaped documents.get JSON
// (reality-linted against the discovery schema in gdocsFixtureSchema.test.ts)
// and the expectations here pin what the engine actually depends on: index
// math in UTF-16 code units, the textRun whitelist, omitted-zero decoding,
// tab counting, and the never-throws defensive contract. No snapshots: a
// snapshot would happily bless a wrong decode.

import * as fs from "fs";
import * as path from "path";
import { parseDocument, DOC_FIELDS_MASK } from "../gdocs/src/core/parse";
import { GDoc } from "../gdocs/src/core/types";

const FIXTURE_DIR = path.join(__dirname, "fixtures", "gdocs");

/** Every committed document fixture (the discovery schema is not a doc). The
 * list is read from disk so a new fixture is linted the day it lands. */
const DOC_FIXTURE_NAMES = fs
  .readdirSync(FIXTURE_DIR)
  .filter((f) => f.endsWith(".json") && f !== "discovery-schema.json")
  .sort();

function loadFixture(name: string): unknown {
  return JSON.parse(fs.readFileSync(path.join(FIXTURE_DIR, name), "utf8")) as unknown;
}

function parseFixture(name: string): GDoc {
  return parseDocument(loadFixture(name));
}

describe("fixture roster", () => {
  it("contains exactly the committed scenario docs", () => {
    // A silently deleted fixture must fail loudly, not shrink the suite.
    expect(DOC_FIXTURE_NAMES).toEqual([
      "chips.json",
      "headings.json",
      "namedRanges.json",
      "nearWhite.json",
      "plainCard.json",
      "suggestions.json",
      "surrogate.json",
      "tabbed.json",
      "table.json"
    ]);
  });
});

// ---------------------------------------------------------------------------
// Cross-fixture invariants — the load-bearing index rules hold for EVERY
// fixture, so any future fixture inherits the full battery automatically.
// ---------------------------------------------------------------------------

describe.each(DOC_FIXTURE_NAMES)("invariants: %s", (name) => {
  const doc = parseFixture(name);

  it("keeps every text element's span equal to its UTF-16 text length", () => {
    for (const p of doc.paragraphs) {
      for (const el of p.elements) {
        if (el.kind === "text") {
          expect(el.endIndex - el.startIndex).toBe(el.text.length);
        } else {
          // Whitelist rule: non-text elements carry no visible text.
          expect(el.text).toBe("");
        }
      }
    }
  });

  it("keeps the trailing newline INSIDE each paragraph's final text element", () => {
    for (const p of doc.paragraphs) {
      const last = p.elements[p.elements.length - 1];
      expect(last.kind).toBe("text");
      expect(last.text.endsWith("\n")).toBe(true);
    }
  });

  it("keeps element spans contiguous and flush with the paragraph range", () => {
    for (const p of doc.paragraphs) {
      expect(p.elements.length).toBeGreaterThan(0);
      expect(p.elements[0].startIndex).toBe(p.startIndex);
      expect(p.elements[p.elements.length - 1].endIndex).toBe(p.endIndex);
      for (let i = 1; i < p.elements.length; i++) {
        expect(p.elements[i].startIndex).toBe(p.elements[i - 1].endIndex);
      }
    }
  });

  it("stamps sequential ordinals and marks ONLY the final paragraph last-in-segment", () => {
    doc.paragraphs.forEach((p, i) => {
      expect(p.index).toBe(i);
      expect(p.isLastInSegment).toBe(i === doc.paragraphs.length - 1);
    });
  });

  it("orders paragraphs by strictly increasing start index", () => {
    for (let i = 1; i < doc.paragraphs.length; i++) {
      expect(doc.paragraphs[i].startIndex).toBeGreaterThan(doc.paragraphs[i - 1].startIndex);
    }
  });

  it("starts body content at index 1 and carries a revision id", () => {
    expect(doc.paragraphs[0].startIndex).toBe(1);
    expect(doc.revisionId).not.toBe("");
  });

  it("flags suggestions on the suggestions fixture ONLY", () => {
    expect(doc.suggestionsPresent).toBe(name === "suggestions.json");
  });

  it("counts tabs (multi only on the tabbed fixture)", () => {
    expect(doc.tabCount).toBe(name === "tabbed.json" ? 2 : 1);
  });
});

// ---------------------------------------------------------------------------
// Per-fixture behavior
// ---------------------------------------------------------------------------

describe("plainCard.json (card anatomy + spacing + omitted-zero rgb)", () => {
  const doc = parseFixture("plainCard.json");

  it("reads the tag/cite/body paragraph styles", () => {
    expect(doc.paragraphs.map((p) => p.namedStyleType)).toEqual([
      "HEADING_4",
      "NORMAL_TEXT",
      "NORMAL_TEXT",
      "NORMAL_TEXT"
    ]);
  });

  it("decodes the cite lead as bold 14pt and the trailing run as inherited", () => {
    const [lead, rest] = doc.paragraphs[1].elements;
    expect(lead.bold).toBe(true);
    expect(lead.fontSizePt).toBe(14);
    expect(rest.bold).toBe(false);
    expect(rest.fontSizePt).toBeNull();
  });

  it("decodes a yellow highlight whose ZERO blue channel was omitted on the wire (edge row 15)", () => {
    const highlighted = doc.paragraphs[2].elements[1];
    expect(highlighted.text).toBe("every model predicted");
    expect(highlighted.backgroundHex).toBe("#ffff00");
  });

  it("reads absent background color as null", () => {
    expect(doc.paragraphs[2].elements[0].backgroundHex).toBeNull();
  });

  it("distinguishes inherited spacing (null) from direct values", () => {
    expect(doc.paragraphs[0].spaceAbovePt).toBeNull();
    expect(doc.paragraphs[0].spaceBelowPt).toBeNull();
    expect(doc.paragraphs[1].spaceAbovePt).toBe(12);
    expect(doc.paragraphs[1].spaceBelowPt).toBeNull();
  });

  it("re-materializes an explicit zero spacing whose magnitude the wire omitted", () => {
    // spaceAbove: {unit:"PT"} IS a direct zero — conflating it with inherit
    // would make the styles retro pass skip exactly what it must clear.
    expect(doc.paragraphs[3].spaceAbovePt).toBe(0);
    expect(doc.paragraphs[3].spaceBelowPt).toBe(0);
  });

  it("maps stated named-style sizes and nothing else", () => {
    expect(doc.namedStyleSizesPt).toEqual({ NORMAL_TEXT: 11, HEADING_4: 12 });
  });

  it("carries the revision id through verbatim", () => {
    expect(doc.revisionId).toBe("fixture-rev-plain-card");
  });
});

describe("headings.json (TITLE + H1-H4 + H5, named-style size map)", () => {
  const doc = parseFixture("headings.json");

  it("reads every named style in document order (H5 included — keeper decides its fate, not the parser)", () => {
    expect(doc.paragraphs.map((p) => p.namedStyleType)).toEqual([
      "TITLE",
      "HEADING_1",
      "HEADING_2",
      "HEADING_3",
      "HEADING_4",
      "NORMAL_TEXT",
      "HEADING_5"
    ]);
  });

  it("maps all nine named-style sizes", () => {
    expect(doc.namedStyleSizesPt).toEqual({
      NORMAL_TEXT: 11,
      TITLE: 26,
      SUBTITLE: 15,
      HEADING_1: 20,
      HEADING_2: 16,
      HEADING_3: 14,
      HEADING_4: 12,
      HEADING_5: 11,
      HEADING_6: 11
    });
  });
});

describe("table.json (edge row 1: cell paragraphs kept, flattened in order)", () => {
  const doc = parseFixture("table.json");

  it("flattens cell paragraphs between the surrounding body paragraphs", () => {
    expect(doc.paragraphs.map((p) => p.elements[0].text)).toEqual([
      "Source comparison table:\n",
      "Aff ev quality is higher.\n",
      "Neg ev is older.\n",
      "After the table the body resumes.\n"
    ]);
  });

  it("marks exactly the cell paragraphs inTable", () => {
    expect(doc.paragraphs.map((p) => p.inTable)).toEqual([false, true, true, false]);
  });
});

describe("chips.json (edge row 2 / plan A9: whitelist breaks spans around objects)", () => {
  const doc = parseFixture("chips.json");

  it("maps person/richLink/pageBreak/inlineObject to kind 'other' with 1-unit spans", () => {
    const chipOf = (p: number): GDoc["paragraphs"][number]["elements"][number] => {
      const found = doc.paragraphs[p].elements.find((e) => e.kind === "other");
      if (found === undefined) throw new Error(`no chip in paragraph ${p}`);
      return found;
    };
    for (const p of [0, 1, 2, 3]) {
      const chip = chipOf(p);
      expect(chip.endIndex - chip.startIndex).toBe(1);
      // Style channels on a non-text element are meaningless — pinned inert
      // so a keeper/planner change can never start "reading" a chip.
      expect(chip.fontSizePt).toBeNull();
      expect(chip.bold).toBe(false);
      expect(chip.backgroundHex).toBeNull();
    }
  });

  it("keeps the text runs flush against the chip (break-around, no gaps)", () => {
    const [before, chip, after] = doc.paragraphs[0].elements;
    expect(before.text).toBe("Ask ");
    expect(chip.kind).toBe("other");
    expect(chip.startIndex).toBe(before.endIndex);
    expect(after.startIndex).toBe(chip.endIndex);
    expect(after.text).toBe(" about the aff.\n");
  });

  it("gives a chip-final paragraph its newline as a standalone text run", () => {
    // The API never puts the paragraph newline inside a non-text element.
    const pageBreakPara = doc.paragraphs[2];
    expect(pageBreakPara.elements.map((e) => e.kind)).toEqual(["other", "text"]);
    expect(pageBreakPara.elements[1].text).toBe("\n");
  });
});

describe("surrogate.json (edge row 14: indexes are UTF-16 code units)", () => {
  const doc = parseFixture("surrogate.json");

  it("counts an emoji as two code units inside a run", () => {
    const run = doc.paragraphs[0].elements[0];
    expect(run.text).toBe("Impact \u{1F525} calculus\n");
    expect(run.text.length).toBe(19); // 7 + 2 (surrogate pair) + 9 + 1
    expect(run.endIndex).toBe(run.startIndex + 19);
  });

  it("counts a ZWJ family sequence as eight code units in its own run", () => {
    const family = doc.paragraphs[1].elements[1];
    expect(family.text.length).toBe(8); // 2 + 1 (ZWJ) + 2 + 1 (ZWJ) + 2
    expect(family.endIndex - family.startIndex).toBe(8);
  });

  it("decodes the highlight on the emoji run with its omitted red channel as zero", () => {
    expect(doc.paragraphs[1].elements[1].backgroundHex).toBe("#00ffff");
  });
});

describe("suggestions.json (plan A16: any suggested* key anywhere)", () => {
  const doc = parseFixture("suggestions.json");

  it("raises the suggestions flag", () => {
    expect(doc.suggestionsPresent).toBe(true);
  });

  it("still parses the content normally (the GATE refuses, not the parser)", () => {
    expect(doc.paragraphs).toHaveLength(2);
    expect(doc.paragraphs[0].elements.map((e) => e.text)).toEqual([
      "Baseline text ",
      "suggested addition",
      " continues after it.\n"
    ]);
  });
});

describe("tabbed.json (plan A3: tabs counted, FIRST tab read)", () => {
  const doc = parseFixture("tabbed.json");

  it("counts both leaf tabs", () => {
    expect(doc.tabCount).toBe(2);
  });

  it("reads only the first tab's body segment", () => {
    expect(doc.paragraphs).toHaveLength(1);
    expect(doc.paragraphs[0].elements[0].text).toBe("Tab one card text lives here.\n");
  });

  it("reads named styles from the first tab's segment", () => {
    expect(doc.namedStyleSizesPt).toEqual({ NORMAL_TEXT: 11 });
  });
});

describe("namedRanges.json (split segments + same-name siblings, all preserved)", () => {
  const doc = parseFixture("namedRanges.json");

  it("flattens the name-keyed map to one entry per NamedRange id", () => {
    expect(doc.namedRanges.map((nr) => nr.id)).toEqual(["kix.nr1", "kix.nr2", "kix.nr3"]);
  });

  it("keeps BOTH segments of the split rstm range, in order", () => {
    const split = doc.namedRanges[0];
    expect(split.name).toBe("rstm:v1:7x11,10x14");
    expect(split.segments).toEqual([
      { startIndex: 5, endIndex: 12 },
      { startIndex: 20, endIndex: 30 }
    ]);
  });

  it("keeps foreign ranges too — ownership filtering is rangeNames' job, not the parser's", () => {
    expect(doc.namedRanges[1].name).toBe("docs-anchor");
    expect(doc.namedRanges[2].name).toBe("docs-anchor");
    expect(doc.namedRanges[1].segments).toEqual([{ startIndex: 41, endIndex: 46 }]);
  });
});

describe("nearWhite.json (plan A8: shading decodes EXACTLY; keeper decides later)", () => {
  const doc = parseFixture("nearWhite.json");

  it("decodes near-white web-paste shading to its exact hex", () => {
    expect(doc.paragraphs[0].elements[1].backgroundHex).toBe("#f8f9fa");
  });

  it("decodes the real highlight beside it", () => {
    expect(doc.paragraphs[1].elements[1].backgroundHex).toBe("#ffff00");
  });
});

// ---------------------------------------------------------------------------
// Tab-shape handling beyond the committed fixtures (inline raws)
// ---------------------------------------------------------------------------

/** Minimal valid paragraph SE for inline raw docs. */
function rawPara(startIndex: number, text: string): Record<string, unknown> {
  return {
    startIndex,
    endIndex: startIndex + text.length,
    paragraph: {
      elements: [
        { startIndex, endIndex: startIndex + text.length, textRun: { content: text, textStyle: {} } }
      ],
      paragraphStyle: { namedStyleType: "NORMAL_TEXT" }
    }
  };
}

describe("tab resolution", () => {
  it("treats a single-tab tabs-shaped read as one tab and reads its segment", () => {
    const doc = parseDocument({
      revisionId: "r1",
      tabs: [
        {
          tabProperties: { tabId: "t.0" },
          documentTab: { body: { content: [{ endIndex: 1 }, rawPara(1, "only tab\n")] } }
        }
      ]
    });
    expect(doc.tabCount).toBe(1);
    expect(doc.paragraphs[0].elements[0].text).toBe("only tab\n");
  });

  it("counts LEAF tabs only — a parent with two children is two tabs, content from the first tab", () => {
    const doc = parseDocument({
      tabs: [
        {
          tabProperties: { tabId: "t.parent" },
          documentTab: { body: { content: [{ endIndex: 1 }, rawPara(1, "parent\n")] } },
          childTabs: [
            { tabProperties: { tabId: "t.c1" }, documentTab: { body: { content: [] } } },
            { tabProperties: { tabId: "t.c2" }, documentTab: { body: { content: [] } } }
          ]
        }
      ]
    });
    expect(doc.tabCount).toBe(2);
    expect(doc.paragraphs[0].elements[0].text).toBe("parent\n");
  });

  it("falls back to the legacy body when tabs carry only metadata (includeTabsContent:false shape)", () => {
    const doc = parseDocument({
      tabs: [{ tabProperties: { tabId: "t.0" } }],
      body: { content: [{ endIndex: 1 }, rawPara(1, "legacy body\n")] }
    });
    expect(doc.tabCount).toBe(1);
    expect(doc.paragraphs[0].elements[0].text).toBe("legacy body\n");
  });

  it("clamps a garbage tabs array to one tab and still reads the legacy body", () => {
    const doc = parseDocument({
      tabs: [],
      body: { content: [{ endIndex: 1 }, rawPara(1, "still here\n")] }
    });
    expect(doc.tabCount).toBe(1);
    expect(doc.paragraphs[0].elements[0].text).toBe("still here\n");
  });
});

// ---------------------------------------------------------------------------
// Defensive parsing — malformed pieces decode to safe defaults, never throw.
// ---------------------------------------------------------------------------

describe("defensive parsing (malformed input never throws)", () => {
  const EMPTY: GDoc = {
    revisionId: "",
    tabCount: 1,
    paragraphs: [],
    namedRanges: [],
    namedStyleSizesPt: {},
    suggestionsPresent: false
  };

  it.each([
    ["null", null],
    ["undefined", undefined],
    ["a number", 42],
    ["a string", "junk"],
    ["an array", []],
    ["a boolean", true]
  ])("parses %s to the empty document view", (_label, raw) => {
    expect(parseDocument(raw)).toEqual(EMPTY);
  });

  it("defaults a run with no textStyle to inherit/not-bold/no-background", () => {
    const doc = parseDocument({
      body: {
        content: [
          { startIndex: 1, endIndex: 3, paragraph: { elements: [{ startIndex: 1, endIndex: 3, textRun: { content: "x\n" } }] } }
        ]
      }
    });
    const el = doc.paragraphs[0].elements[0];
    expect(el.kind).toBe("text");
    expect(el.fontSizePt).toBeNull();
    expect(el.bold).toBe(false);
    expect(el.backgroundHex).toBeNull();
  });

  it("treats a junk-typed textRun as kind 'other' (the whitelist holds)", () => {
    const doc = parseDocument({
      body: {
        content: [
          {
            startIndex: 1,
            endIndex: 4,
            paragraph: {
              elements: [
                { startIndex: 1, endIndex: 3, textRun: "zzz" },
                { startIndex: 3, endIndex: 4, textRun: { content: "\n", textStyle: {} } }
              ]
            }
          }
        ]
      }
    });
    expect(doc.paragraphs[0].elements[0].kind).toBe("other");
    expect(doc.paragraphs[0].elements[0].text).toBe("");
  });

  it("derives a missing endIndex from the content length for text elements", () => {
    const doc = parseDocument({
      body: { content: [{ startIndex: 1, paragraph: { elements: [{ startIndex: 1, textRun: { content: "hi\n", textStyle: {} } }] } }] }
    });
    expect(doc.paragraphs[0].elements[0].endIndex).toBe(4);
    expect(doc.paragraphs[0].endIndex).toBe(4); // paragraph closes at its last element
  });

  it("gives a non-text element with no endIndex a zero-width span", () => {
    const doc = parseDocument({
      body: {
        content: [
          {
            startIndex: 1,
            endIndex: 7,
            paragraph: {
              elements: [
                { startIndex: 5, person: {} },
                { startIndex: 5, endIndex: 7, textRun: { content: "a\n", textStyle: {} } }
              ]
            }
          }
        ]
      }
    });
    const chip = doc.paragraphs[0].elements[0];
    expect(chip.startIndex).toBe(5);
    expect(chip.endIndex).toBe(5);
  });

  it("chains missing element start indexes from the previous element's end", () => {
    const doc = parseDocument({
      body: {
        content: [
          {
            startIndex: 1,
            endIndex: 5,
            paragraph: {
              elements: [{ textRun: { content: "ab", textStyle: {} } }, { textRun: { content: "c\n", textStyle: {} } }]
            }
          }
        ]
      }
    });
    expect(doc.paragraphs[0].elements.map((e) => [e.startIndex, e.endIndex])).toEqual([
      [1, 3],
      [3, 5]
    ]);
  });

  it("parses an element-less paragraph as an empty zero-length view", () => {
    const doc = parseDocument({ body: { content: [{ startIndex: 1, paragraph: {} }] } });
    expect(doc.paragraphs[0].elements).toEqual([]);
    expect(doc.paragraphs[0].endIndex).toBe(doc.paragraphs[0].startIndex);
  });

  it("skips junk entries inside the content array", () => {
    const doc = parseDocument({ body: { content: [42, "junk", null, rawPara(1, "real\n")] } });
    expect(doc.paragraphs).toHaveLength(1);
  });

  it("decodes the background color states distinctly: absent, transparent, opaque-empty", () => {
    const mk = (textStyle: unknown): string | null =>
      parseDocument({
        body: {
          content: [
            { startIndex: 1, endIndex: 3, paragraph: { elements: [{ startIndex: 1, endIndex: 3, textRun: { content: "x\n", textStyle } }] } }
          ]
        }
      }).paragraphs[0].elements[0].backgroundHex;
    expect(mk({})).toBeNull(); // no backgroundColor key = inherited
    expect(mk({ backgroundColor: {} })).toBeNull(); // color unset = transparent
    // color set with empty rgb = opaque with all channels omitted (zero) = black.
    expect(mk({ backgroundColor: { color: {} } })).toBe("#000000");
    // junk/out-of-range channels clamp instead of poisoning the hex.
    expect(mk({ backgroundColor: { color: { rgbColor: { red: "x", green: 2 } } } })).toBe("#00ff00");
  });

  it("reads a magnitude-less fontSize as inherited (0pt text cannot exist)", () => {
    const doc = parseDocument({
      body: {
        content: [
          {
            startIndex: 1,
            endIndex: 3,
            paragraph: { elements: [{ startIndex: 1, endIndex: 3, textRun: { content: "x\n", textStyle: { fontSize: { unit: "PT" } } } }] }
          }
        ]
      }
    });
    expect(doc.paragraphs[0].elements[0].fontSizePt).toBeNull();
  });

  it("reads junk spacing as inherited but an empty dimension as explicit zero", () => {
    const mk = (paragraphStyle: unknown): [number | null, number | null] => {
      const p = parseDocument({
        body: {
          content: [
            { startIndex: 1, endIndex: 3, paragraph: { elements: [{ startIndex: 1, endIndex: 3, textRun: { content: "x\n", textStyle: {} } }], paragraphStyle } }
          ]
        }
      }).paragraphs[0];
      return [p.spaceAbovePt, p.spaceBelowPt];
    };
    expect(mk({ spaceAbove: "junk" })).toEqual([null, null]);
    expect(mk({ spaceAbove: {}, spaceBelow: { magnitude: "x" } })).toEqual([0, 0]);
  });

  it("defaults an unknown namedStyleType to NORMAL_TEXT (conservative: body text)", () => {
    const doc = parseDocument({
      body: {
        content: [
          { startIndex: 1, endIndex: 3, paragraph: { elements: [{ startIndex: 1, endIndex: 3, textRun: { content: "x\n", textStyle: {} } }], paragraphStyle: { namedStyleType: "HEADING_9" } } }
        ]
      }
    });
    expect(doc.paragraphs[0].namedStyleType).toBe("NORMAL_TEXT");
  });

  it("skips junk named-range shapes without losing well-formed siblings", () => {
    const doc = parseDocument({
      namedRanges: {
        junk: 42,
        ok: {
          name: "ok",
          namedRanges: [
            "junk-entry",
            { namedRangeId: "kix.ok", name: "ok", ranges: [{ startIndex: 2, endIndex: 5 }, "junk-range"] }
          ]
        }
      }
    });
    expect(doc.namedRanges).toEqual([{ id: "kix.ok", name: "ok", segments: [{ startIndex: 2, endIndex: 5 }] }]);
  });

  it("re-materializes a named-range segment's omitted startIndex as zero", () => {
    const doc = parseDocument({
      namedRanges: { x: { name: "x", namedRanges: [{ namedRangeId: "kix.x", name: "x", ranges: [{ endIndex: 3 }] }] } }
    });
    expect(doc.namedRanges[0].segments).toEqual([{ startIndex: 0, endIndex: 3 }]);
  });

  it("ignores junk named styles and styles without a stated size", () => {
    const doc = parseDocument({
      namedStyles: {
        styles: [
          42,
          { namedStyleType: "HEADING_9", textStyle: { fontSize: { magnitude: 99, unit: "PT" } } },
          { namedStyleType: "HEADING_1", textStyle: {} },
          { namedStyleType: "NORMAL_TEXT", textStyle: { fontSize: { magnitude: 11, unit: "PT" } } }
        ]
      }
    });
    expect(doc.namedStyleSizesPt).toEqual({ NORMAL_TEXT: 11 });
  });

  it("finds a suggested* key buried inside a chip (the walk is generic, not a field list)", () => {
    const doc = parseDocument({
      body: {
        content: [
          {
            startIndex: 1,
            endIndex: 3,
            paragraph: {
              elements: [
                { startIndex: 1, endIndex: 2, person: { suggestedInsertionIds: ["s1"] } },
                { startIndex: 2, endIndex: 3, textRun: { content: "\n", textStyle: {} } }
              ]
            }
          }
        ]
      }
    });
    expect(doc.suggestionsPresent).toBe(true);
  });

  it("does not trip the suggestion flag on suggestion-ADJACENT key names", () => {
    // /^suggested/ must not match e.g. "suggestionsViewMode" — that key is
    // present on every read and says nothing about pending suggestions.
    expect(parseDocument({ suggestionsViewMode: "SUGGESTIONS_INLINE" }).suggestionsPresent).toBe(false);
  });
});

// ---------------------------------------------------------------------------
// DOC_FIELDS_MASK — structural sanity. The schema-reality cross-check (every
// token is a real Document property) lives in gdocsFixtureSchema.test.ts.
// ---------------------------------------------------------------------------

describe("DOC_FIELDS_MASK", () => {
  it("is a single compact mask with balanced parentheses", () => {
    expect(DOC_FIELDS_MASK).not.toMatch(/\s/);
    let depth = 0;
    for (const ch of DOC_FIELDS_MASK) {
      if (ch === "(") depth++;
      if (ch === ")") depth--;
      expect(depth).toBeGreaterThanOrEqual(0);
    }
    expect(depth).toBe(0);
  });

  it.each([
    "revisionId",
    "textStyle(fontSize,bold,backgroundColor)",
    "paragraphStyle(namedStyleType,spaceAbove,spaceBelow)",
    "suggestedInsertionIds",
    "suggestedDeletionIds",
    "suggestedTextStyleChanges",
    "suggestedParagraphStyleChanges",
    "namedRanges",
    "namedStyles(styles(namedStyleType,textStyle(fontSize)))",
    "table(tableRows(tableCells(content)))",
    "documentTab"
  ])("selects %s", (piece) => {
    expect(DOC_FIELDS_MASK).toContain(piece);
  });

  it("selects NO sibling tab-level fields beside documentTab (they tripped the live API reject)", () => {
    // tabProperties/childTabs next to documentTab pushed the mask back into the
    // "tabs and legacy text-level fields in the same request" reject across two
    // wet rounds. The mask must stay at the documented-valid tabs(documentTab(...))
    // shape — the only token inside tabs(...) is documentTab.
    expect(DOC_FIELDS_MASK).not.toContain("tabProperties");
    expect(DOC_FIELDS_MASK).not.toContain("childTabs");
    expect(DOC_FIELDS_MASK).toContain("tabs(documentTab(");
  });

  it("never combines `tabs` with legacy top-level text fields in one mask (the wet-round API reject)", () => {
    // The Docs API refuses a get whose mask asks for BOTH document.tabs and
    // legacy text-level fields: "Field mask cannot retrieve document.tabs and
    // legacy text-level fields from the Document resource in the same request."
    // The verb read pairs this mask with includeTabsContent:true, so the segment
    // selection must live ONLY under tabs[].documentTab — exactly once each, and
    // NEVER at the top level (the portion before `tabs(`).
    const count = (piece: string): number => DOC_FIELDS_MASK.split(piece).length - 1;
    expect(count("body(content(")).toBe(1);
    expect(count(",namedRanges,")).toBe(1);
    expect(count("namedStyles(styles(")).toBe(1);

    const topLevel = DOC_FIELDS_MASK.split("tabs(")[0];
    expect(topLevel).not.toMatch(/body\(|namedRanges|namedStyles/);
    expect(DOC_FIELDS_MASK).toContain("documentTab(body(content(");
  });

  it("whitelists every non-text element type so chip suggestions stay visible to the gate", () => {
    for (const key of [
      "autoText",
      "columnBreak",
      "dateElement",
      "equation",
      "footnoteReference",
      "horizontalRule",
      "inlineObjectElement",
      "pageBreak",
      "person",
      "richLink"
    ]) {
      expect(DOC_FIELDS_MASK).toContain(key);
    }
  });
});
