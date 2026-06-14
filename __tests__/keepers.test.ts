import {
  isHeadingKept,
  paragraphHasCiteRun,
  computeRunKeepFlags,
  planCrossGapSeparators
} from "../src/core/keepers";
import { ptToHalfPoints, buildPocketBorderOoxml, STYLE_MAP, POCKET_BORDER } from "../src/core/styles";
import { RunView } from "../src/core/types";

/** Build a RunView with sensible defaults for the field under test. */
function rv(text: string, over: Partial<RunView> = {}): RunView {
  return {
    index: 0,
    text,
    highlight: null,
    citeStyled: false,
    underline: false,
    hidden: false,
    eligible: true,
    hasInternalPart: false,
    ...over
  };
}

/** Re-index a list of runs so indices are stable/positional. */
function seq(runs: RunView[]): RunView[] {
  return runs.map((r, i) => ({ ...r, index: i }));
}

describe("isHeadingKept (outline level 0–3, decision #7)", () => {
  it.each([
    [null, false],
    [0, true],
    [1, true],
    [3, true],
    [4, false],
    [8, false],
    [-1, false]
  ])("level %s -> %s", (level, expected) => {
    expect(isHeadingKept(level as number | null)).toBe(expected);
  });
});

describe("paragraphHasCiteRun (decision #6b)", () => {
  it("is true when any run carries the cite style", () => {
    expect(paragraphHasCiteRun(seq([rv("a"), rv("b", { citeStyled: true })]))).toBe(true);
  });
  it("is false when no run is cite-styled", () => {
    expect(paragraphHasCiteRun(seq([rv("a"), rv("b")]))).toBe(false);
  });
});

describe("computeRunKeepFlags", () => {
  const all = new Set<string>(["yellow", "green"]);

  it("hides every plain body run when nothing is highlighted", () => {
    const keep = computeRunKeepFlags(seq([rv("card text"), rv("more")]), all);
    expect(keep).toEqual([false, false]);
  });

  it("keeps a run highlighted in a keep-color", () => {
    // Trailing space makes these two distinct words; without it they would be
    // one word and the whole-word rule (#19) would keep both — see next test.
    const keep = computeRunKeepFlags(seq([rv("plain "), rv("hi", { highlight: "yellow" })]), all);
    expect(keep).toEqual([false, true]);
  });

  it("shows ONLY the highlighted run when a highlight splits a word (#19 reversed)", () => {
    // "a"+"b"(yellow) is one word "ab", but only "b" is highlighted -> "a" stays hidden.
    const keep = computeRunKeepFlags(seq([rv("a"), rv("b", { highlight: "yellow" })]), all);
    expect(keep).toEqual([false, true]);
  });

  it("does not keep a highlight outside the keep-set", () => {
    const keep = computeRunKeepFlags(seq([rv("b", { highlight: "red" })]), all);
    expect(keep).toEqual([false]);
  });

  it("keeps ONLY highlighted runs across a split word (#19 reversed)", () => {
    // "inter" + "nation"(yellow) + "al " + "plan": only the highlighted "nation" stays.
    const runs = seq([
      rv("inter"),
      rv("nation", { highlight: "yellow" }),
      rv("al "),
      rv("plan")
    ]);
    expect(computeRunKeepFlags(runs, all)).toEqual([false, true, false, false]);
  });

  it("hides a trailing run glued to a kept word (schulze-makuch '22 ', such as scorpions.')", () => {
    // Only "large desert animals" is highlighted; the unhighlighted ", such as
    // scorpions." run is hidden regardless of the glued comma.
    const runs = seq([
      rv("large desert animals", { highlight: "yellow" }),
      rv(", such as scorpions.")
    ]);
    expect(computeRunKeepFlags(runs, all)).toEqual([true, false]);
  });

  it("hides a lone-punctuation run that is not highlighted (show only highlighted)", () => {
    // "nuclear war"(hl) + ".": the "." is not highlighted, so it is hidden (#19 reversed).
    const runs = seq([rv("nuclear war", { highlight: "yellow" }), rv(".")]);
    expect(computeRunKeepFlags(runs, all)).toEqual([true, false]);
  });

  it("shows only the highlighted fragment of a word (brauner-18 'reduce'/'extinction')", () => {
    // "import"(hl) + "ant cases": only "import" is highlighted -> the view shows "import".
    const importRuns = seq([rv("import", { highlight: "yellow" }), rv("ant cases")]);
    expect(computeRunKeepFlags(importRuns, all)).toEqual([true, false]);
    // "e" + "x"(hl) + "tinction": only the highlighted "x" stays ("extinction" -> "x").
    const extRuns = seq([rv("e"), rv("x", { highlight: "yellow" }), rv("tinction")]);
    expect(computeRunKeepFlags(extRuns, all)).toEqual([false, true, false]);
  });

  it("hides an unhighlighted word glued to a following highlighted run (waite-24 'society')", () => {
    // "myopic"(hl) + " society" + ". People are"(hl): " society" is unhighlighted, so it
    // is hidden even though the next run's "." glues it into the token "society.".
    const runs = seq([
      rv("myopic", { highlight: "yellow" }),
      rv(" society"),
      rv(". People are", { highlight: "yellow" })
    ]);
    expect(computeRunKeepFlags(runs, all)).toEqual([true, false, true]);
  });

  it("never hides structural (ineligible) runs (#16)", () => {
    const keep = computeRunKeepFlags(seq([rv("text"), rv("", { eligible: false })]), all);
    expect(keep).toEqual([false, true]);
  });

  it("keeps nothing by color when the keep-set is empty", () => {
    const keep = computeRunKeepFlags(seq([rv("a", { highlight: "yellow" })]), new Set());
    expect(keep).toEqual([false]);
  });

  it("keeps a whitespace-only run bridging two highlighted words (spacing bug)", () => {
    // "reasons" | " " | "for" — words highlighted, the bare space run between is not.
    // Without the bridge rule the space is hidden -> "reasonsfor".
    const runs = seq([rv("reasons", { highlight: "yellow" }), rv(" "), rv("for", { highlight: "yellow" })]);
    expect(computeRunKeepFlags(runs, all)).toEqual([true, true, true]);
  });

  it("bridges multiple consecutive whitespace runs between kept words", () => {
    const runs = seq([rv("a", { highlight: "yellow" }), rv(" "), rv("\t"), rv("b", { highlight: "yellow" })]);
    expect(computeRunKeepFlags(runs, all)).toEqual([true, true, true, true]);
  });

  it("hides a space run adjacent to hidden body (no leading/trailing space leak)", () => {
    const runs = seq([rv("keep", { highlight: "yellow" }), rv(" "), rv("hiddenbody")]);
    expect(computeRunKeepFlags(runs, all)).toEqual([true, false, false]);
  });

  it("does not promote a space run between two hidden runs", () => {
    expect(computeRunKeepFlags(seq([rv("foo"), rv(" "), rv("bar")]), all)).toEqual([false, false, false]);
  });
});

describe("planCrossGapSeparators (bridge across hidden gaps — wet-test bug 1)", () => {
  const cyan = new Set<string>(["cyan"]);

  // Through the REAL keep policy (computeRunKeepFlags), mirroring classifyParagraph.
  // NOTE: we show only highlighted text, so an unhighlighted run is always hidden — a
  // hidden gap can sit right against a highlighted anchor. The bridge restores ONE space
  // per fused gap (and never before/after hugging punctuation). Fixtures keep the
  // separating space explicit where relevant.
  const plan = (runs: RunView[]) => planCrossGapSeparators(runs, computeRunKeepFlags(runs, cyan));

  it("splits an embedded-space hidden run between two highlighted chunks (the real bug)", () => {
    // Run shapes taken verbatim from samples/[small] 2ac---ndca---semis.docx.
    const runs = seq([
      rv("ultraviolet radiation", { highlight: "cyan" }),
      rv(" are a constant threat, "),
      rv("would", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [{ index: 1, side: "lead" }] });
  });

  it("bridges a hidden gap whose separators are NON-BREAKING spaces (dds2 bunzel-18 'rising oil revenues')", () => {
    // Verbatim shape from samples/[small] 2ac---dds2---finals.docx: the inter-word separators are
    // NBSP (U+00A0), not ASCII spaces — pasted debate text is full of them. Before the fix the
    // separator predicate EXCLUDED NBSP, so the bridge saw no space in the gap and the kept words
    // "rising" + hidden "oil" + "revenues" fused into "risingrevenues". Now the first NBSP-only run
    // in the gap is kept (priority 1), restoring a separator that renders identically and reverses
    // losslessly (we expose the existing NBSP, never rewrite it to an ASCII space).
    const NBSP = "\u00A0";
    const runs = seq([
      rv("rising", { highlight: "cyan" }),
      rv(NBSP),
      rv("oil"),
      rv(NBSP),
      rv("revenues gives", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [1], splits: [] });
  });

  it("rescues a NON-BREAKING space directly between two kept words (dds2 'gives Russia')", () => {
    // Two highlighted words separated by a single hidden NBSP fused into "givesRussia". The
    // whitespace-only rescue in computeRunKeepFlags now treats the NBSP run as a separator and
    // keeps it (both neighbors kept), so the words stay apart and the bridge has nothing to do.
    const runs = seq([
      rv("gives", { highlight: "cyan" }),
      rv("\u00A0"),
      rv("Russia", { highlight: "cyan" })
    ]);
    expect(computeRunKeepFlags(runs, cyan)[1]).toBe(true); // NBSP rescued as a separator
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [] }); // already bridged by the rescue
  });

  it("treats the narrow-NBSP / thin / figure space family as separators too (closes the class)", () => {
    // The fix covers the common typographic spaces pasted text carries, not just U+00A0: narrow
    // NBSP (U+202F), thin space (U+2009), figure space (U+2007) all render as spaces and must
    // bridge identically. Zero-width chars are NOT included (they wouldn't separate visually).
    for (const sp of ["\u202F", "\u2009", "\u2007"]) {
      const runs = seq([
        rv("a", { highlight: "cyan" }),
        rv(sp),
        rv("hidden"),
        rv(sp),
        rv("b", { highlight: "cyan" })
      ]);
      expect(plan(runs)).toEqual({ extraKeep: [1], splits: [] });
    }
  });

  it("bridges a hidden gap whose only separator is an EM DASH (dds2 palmeri 'capital—and')", () => {
    // "c"(kept) + "apital"(hidden) + "—"(hidden, U+2014) + "and"(kept). "capital—and" has NO space,
    // so before the fix the kept words fused into "cand". The em dash IS the separator; exposing the
    // dash-only run gives "c—and" (reversible — the existing char is kept, never a fabricated space).
    const runs = seq([
      rv("c", { highlight: "cyan" }),
      rv("apital"),
      rv("—"),
      rv("and ", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [2], splits: [] });
  });

  it("rescues an EM DASH directly between two kept words (dds2 palmeri 'it—than')", () => {
    // "...to it"(kept) + "—"(hidden, U+2014) + "than..."(kept) fused into "itthan". The dash-only run
    // is now a separator the whitespace-only rescue keeps, so the words stay apart as "it—than".
    const runs = seq([
      rv("to it", { highlight: "cyan" }),
      rv("—"),
      rv("than his", { highlight: "cyan" })
    ]);
    expect(computeRunKeepFlags(runs, cyan)[1]).toBe(true); // em dash rescued as a separator
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [] }); // already bridged by the rescue
  });

  it("does NOT treat a regular hyphen-minus as a separator (word-internal 'term-locked')", () => {
    // A hidden hyphen between two kept fragments stays hidden — a hyphen is word-internal, not a
    // separator. The dash fix is scoped to em/en dashes (U+2014/U+2013), never the hyphen-minus.
    const runs = seq([rv("term", { highlight: "cyan" }), rv("-"), rv("locked", { highlight: "cyan" })]);
    expect(computeRunKeepFlags(runs, cyan)[1]).toBe(false); // hyphen NOT rescued
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [] }); // bridge leaves it fused (faithful)
  });

  // --- Stage B separator audit (`node scripts/auditSeparators.mjs`) ---

  it("treats the whole Unicode Zs space class as separators, incl. four-per-em (Stage B \\p{Zs})", () => {
    // The audit found a four-per-em space (U+2005, "25 kg") the old hand-listed subset missed.
    // Generalizing to \p{Zs} closes the class: every space variant must bridge two kept words.
    for (const sp of [" ", " ", "　", " ", " "]) {
      const runs = seq([rv("a", { highlight: "cyan" }), rv(sp), rv("b", { highlight: "cyan" })]);
      expect(computeRunKeepFlags(runs, cyan)[1]).toBe(true);
    }
  });

  it("rescues a SLASH between two kept words (Stage B 'and/or', 'Good/Bad')", () => {
    // A hidden slash between two highlighted words would fuse them ("andor"); the slash is a prose
    // separator, so the slash-only run is rescued and exposed losslessly as "and/or".
    const runs = seq([rv("and", { highlight: "cyan" }), rv("/"), rv("or", { highlight: "cyan" })]);
    expect(computeRunKeepFlags(runs, cyan)[1]).toBe(true);
  });

  it("does NOT treat underscore as a separator (word-internal identifiers/filenames)", () => {
    // "RAND_RR3139", "Volume-15_Issue-4" — underscore JOINS tokens; keep it excluded like the hyphen.
    const runs = seq([rv("RAND", { highlight: "cyan" }), rv("_"), rv("RR3139", { highlight: "cyan" })]);
    expect(computeRunKeepFlags(runs, cyan)[1]).toBe(false);
  });

  it("prefers keeping an existing whitespace-only run over splitting", () => {
    const runs = seq([
      rv("a", { highlight: "cyan" }),
      rv(" "),
      rv("hidden "),
      rv("b", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [1], splits: [] });
  });

  it("does not bridge when the left chunk already ends with a visible space", () => {
    const runs = seq([
      rv("an ", { highlight: "cyan" }),
      rv("hidden "),
      rv("b", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [] });
  });

  it("emits no separator when only one chunk is highlighted (#19 reversed)", () => {
    // "inter" + "national"(cyan): only "national" is kept now; nothing to bridge.
    const runs = seq([rv("inter"), rv("national", { highlight: "cyan" })]);
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [] });
  });

  it("does not bridge when the next chunk starts with attaching punctuation (waite-24)", () => {
    // "myopic"(hl) + " society"(hidden) + ". People are"(hl): the right chunk starts with
    // "." (hugs left), so NO space is inserted -> reads "myopic. People are".
    const runs = seq([
      rv("myopic", { highlight: "cyan" }),
      rv(" society"),
      rv(". People are", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [] });
  });

  it("bridges every gap across a multi-chunk sentence", () => {
    // ultraviolet radiation | <hidden> | would | <hidden> | give  -> two splits.
    const runs = seq([
      rv("ultraviolet radiation", { highlight: "cyan" }),
      rv(" are a constant threat, "),
      rv("would", { highlight: "cyan" }),
      rv(" seem to "),
      rv("give", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({
      extraKeep: [],
      splits: [
        { index: 1, side: "lead" },
        { index: 3, side: "lead" }
      ]
    });
  });

  it("exposes the first interior space of a glued clause (bug 2 follow-up)", () => {
    // ", such as scorpions." starts "," and ends "." (no boundary space), so priorities
    // 1-3 can't bridge it; priority 4 exposes its FIRST interior space (offset 1). Both
    // fragments ("," / "such as scorpions.") are unhighlighted so they stay hidden, and
    // Re-hide rescues the exposed space — idempotent wherever we split.
    const runs = seq([
      rv("large desert animals", { highlight: "cyan" }),
      rv(", such as scorpions."),
      rv("would", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [{ index: 1, side: "interior", offset: 1 }] });
  });

  it("exposes the interior space of a glued unhighlighted run between two highlights (brauner-18 'reduc x')", () => {
    // "reduc"(hl) + "e e"(not hl: the 'e' of reduce + space + 'e' of extinction) + "x"(hl):
    // the gap "e e" has only an interior space -> priority 4 exposes it so the view reads
    // "reduc x", not "reducx".
    const runs = seq([
      rv("reduc", { highlight: "cyan" }),
      rv("e e"),
      rv("x", { highlight: "cyan" })
    ]);
    expect(plan(runs)).toEqual({ extraKeep: [], splits: [{ index: 1, side: "interior", offset: 1 }] });
  });

  // The trailing-space fallback and the no-space no-op are DEFENSIVE branches that
  // computeRunKeepFlags can't produce (a real hidden gap is space-separated and so
  // hits priority 1 or 2). Exercise them by driving the pure planner with an
  // explicit keep vector.
  it("falls back to a trailing space when the gap has no leading space (direct)", () => {
    const runs = seq([rv("a"), rv("hidden "), rv("b")]);
    expect(planCrossGapSeparators(runs, [true, false, true])).toEqual({
      extraKeep: [],
      splits: [{ index: 1, side: "trail" }]
    });
  });

  it("leaves the words fused when the gap has no exposable space (direct, faithful)", () => {
    const runs = seq([rv("a"), rv("x"), rv("b")]);
    expect(planCrossGapSeparators(runs, [true, false, true])).toEqual({ extraKeep: [], splits: [] });
  });
});

describe("styles constants (decisions #8, #9)", () => {
  it("converts points to half-points", () => {
    expect(ptToHalfPoints(14)).toBe(28);
    expect(ptToHalfPoints(13)).toBe(26);
  });

  it("maps cite to the stable styleId at 14pt", () => {
    expect(STYLE_MAP.cite.styleId).toBe("Style13ptBold");
    expect(STYLE_MAP.cite.sizePt).toBe(14);
    expect(STYLE_MAP.pocket.sizePt).toBe(26);
  });

  it("builds a four-sided pocket border", () => {
    const ooxml = buildPocketBorderOoxml();
    expect(ooxml).toContain("<w:pBdr>");
    expect(ooxml).toContain("</w:pBdr>");
    for (const side of ["top", "left", "bottom", "right"]) {
      expect(ooxml).toContain(`<w:${side} `);
    }
  });

  it("draws the pocket box at 3pt on all four sides (matches Verbatim's w:sz=24)", () => {
    // Contract: the pocket box is 3pt to match Verbatim's pocket (Heading 1) style. OOXML border
    // `w:sz` is in eighths of a point, so 3pt = 24. This guards against the two border
    // representations silently drifting (they previously diverged to 1pt / 0.5pt).
    expect(POCKET_BORDER.sizeEighths).toBe(24);
    expect(POCKET_BORDER.borderWidthToken).toBe("Pt300"); // the live Style.borders enum equivalent
    const ooxml = buildPocketBorderOoxml();
    for (const side of ["top", "left", "bottom", "right"]) {
      expect(ooxml).toContain(`<w:${side} w:val="single" w:sz="24"`);
    }
  });
});
