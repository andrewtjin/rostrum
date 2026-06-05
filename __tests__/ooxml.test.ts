import { readRuns, applyRunVisibility, makeAllVisible } from "../src/core/ooxml";
import { computeRunKeepFlags, planCrossGapSeparators } from "../src/core/keepers";

// ---------------------------------------------------------------------------
// Fixture builders: small, valid `<w:p>` fragments (w: namespace declared so
// xmldom parses cleanly). Each `run()` builds one `<w:r>`.
// ---------------------------------------------------------------------------
const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

interface RunOpts {
  highlight?: string;
  cite?: boolean;
  vanish?: boolean;
  field?: boolean;
}

function run(text: string, opts: RunOpts = {}): string {
  const rPr: string[] = [];
  if (opts.cite) rPr.push(`<w:rStyle w:val="Style13ptBold"/>`);
  if (opts.highlight) rPr.push(`<w:highlight w:val="${opts.highlight}"/>`);
  if (opts.vanish) rPr.push(`<w:vanish/>`);
  const rPrXml = rPr.length ? `<w:rPr>${rPr.join("")}</w:rPr>` : "";
  const content = opts.field
    ? `<w:fldChar w:fldCharType="begin"/>`
    : `<w:t xml:space="preserve">${text}</w:t>`;
  return `<w:r>${rPrXml}${content}</w:r>`;
}

function para(inner: string, pPr = ""): string {
  return `<w:p xmlns:w="${W_NS}">${pPr}${inner}</w:p>`;
}

const vanishCount = (xml: string): number => (xml.match(/<w:vanish\/>/g) || []).length;
const hasHiddenParaMark = (xml: string): boolean =>
  /<w:pPr>[\s\S]*?<w:rPr>[\s\S]*?<w:vanish\/>[\s\S]*?<\/w:rPr>[\s\S]*?<\/w:pPr>/.test(xml);

describe("readRuns", () => {
  it("reads text, highlight (lower-cased), cite, hidden, eligibility", () => {
    const xml = para(
      run("plain") +
        run("hi", { highlight: "Yellow" }) +
        run("cited", { cite: true }) +
        run("gone", { vanish: true }) +
        run("", { field: true })
    );
    const runs = readRuns(xml);
    expect(runs).toHaveLength(5);

    expect(runs[0]).toMatchObject({ text: "plain", highlight: null, citeStyled: false, hidden: false, eligible: true });
    expect(runs[1]).toMatchObject({ text: "hi", highlight: "yellow" });
    expect(runs[2].citeStyled).toBe(true);
    expect(runs[3].hidden).toBe(true);
    expect(runs[4].eligible).toBe(false); // field run is structural
  });

  it("treats highlight 'none' and absent as null", () => {
    const runs = readRuns(para(run("a", { highlight: "none" }) + run("b")));
    expect(runs[0].highlight).toBeNull();
    expect(runs[1].highlight).toBeNull();
  });

  it("converts tabs and breaks to whitespace in run text", () => {
    const xml = para(`<w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:br/><w:t>c</w:t></w:r>`);
    expect(readRuns(xml)[0].text).toBe("a\tb\nc");
  });
});

describe("applyRunVisibility", () => {
  it("hides only the flagged runs and reports a change", () => {
    const xml = para(run("keep") + run("hide"));
    const { xml: out, changed } = applyRunVisibility(xml, [false, true], false);
    expect(changed).toBe(true);
    const runs = readRuns(out);
    expect(runs[0].hidden).toBe(false);
    expect(runs[1].hidden).toBe(true);
    expect(hasHiddenParaMark(out)).toBe(false);
  });

  it("hides the paragraph mark when asked (condensed view)", () => {
    const xml = para(run("a") + run("b"));
    const { xml: out } = applyRunVisibility(xml, [true, true], true);
    expect(readRuns(out).every((r) => r.hidden)).toBe(true);
    expect(hasHiddenParaMark(out)).toBe(true);
  });

  it("is idempotent — re-applying the same flags changes nothing", () => {
    const xml = para(run("a") + run("b"));
    const first = applyRunVisibility(xml, [true, false], false);
    const second = applyRunVisibility(first.xml, [true, false], false);
    expect(second.changed).toBe(false);
  });

  it("clears a previously hidden mark when no longer fully hidden", () => {
    const hidden = applyRunVisibility(para(run("a") + run("b")), [true, true], true).xml;
    const reclassified = applyRunVisibility(hidden, [false, true], false);
    expect(reclassified.changed).toBe(true);
    expect(hasHiddenParaMark(reclassified.xml)).toBe(false);
  });
});

describe("makeAllVisible", () => {
  it("removes every vanish (runs + paragraph mark) and is convergent", () => {
    const hidden = applyRunVisibility(para(run("a") + run("b")), [true, true], true).xml;
    expect(vanishCount(hidden)).toBeGreaterThan(0);

    const { xml: shown, changed } = makeAllVisible(hidden);
    expect(changed).toBe(true);
    expect(vanishCount(shown)).toBe(0);
    expect(hasHiddenParaMark(shown)).toBe(false);

    // Running again is a no-op (convergent).
    expect(makeAllVisible(shown).changed).toBe(false);
  });

  it("round-trips: hide then show returns to an all-visible state", () => {
    const original = para(run("a") + run("b") + run("c"));
    const hidden = applyRunVisibility(original, [true, true, true], true).xml;
    const shown = makeAllVisible(hidden).xml;
    expect(readRuns(shown).some((r) => r.hidden)).toBe(false);
  });
});

describe("robustness (audit fixes)", () => {
  it("targets the document-body paragraph, not a header part (C1)", () => {
    const pkg =
      `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
      `<pkg:part pkg:name="/word/header1.xml"><pkg:xmlData>` +
      `<w:hdr xmlns:w="${W_NS}"><w:p><w:r><w:t>RUNNING HEADER</w:t></w:r></w:p></w:hdr>` +
      `</pkg:xmlData></pkg:part>` +
      `<pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
      `<w:document xmlns:w="${W_NS}"><w:body><w:p><w:r><w:t>CARD BODY</w:t></w:r></w:p></w:body></w:document>` +
      `</pkg:xmlData></pkg:part></pkg:package>`;
    expect(readRuns(pkg).map((r) => r.text)).toEqual(["CARD BODY"]);

    const { xml: out } = applyRunVisibility(pkg, [true], false);
    expect(readRuns(out)[0].hidden).toBe(true); // body run hidden
    expect(out).toContain("RUNNING HEADER"); // header part untouched
  });

  it("never marks a simple-field result run (PAGE) eligible to hide (H2)", () => {
    const xml = para(`<w:fldSimple w:instr="PAGE"><w:r><w:t>3</w:t></w:r></w:fldSimple>` + run("body"));
    const runs = readRuns(xml);
    expect(runs[0].eligible).toBe(false);
    expect(runs[1].eligible).toBe(true);
  });

  it("collapses duplicate <w:vanish> in a single makeAllVisible pass (M4)", () => {
    const xml = para(`<w:r><w:rPr><w:vanish/><w:vanish/></w:rPr><w:t>x</w:t></w:r>`);
    const { xml: out, changed } = makeAllVisible(xml);
    expect(changed).toBe(true);
    expect(vanishCount(out)).toBe(0);
    expect(readRuns(out)[0].hidden).toBe(false);
  });

  it('does not rewrite a run whose only vanish is an explicit w:val="false"', () => {
    const xml = para(`<w:r><w:rPr><w:vanish w:val="false"/></w:rPr><w:t>x</w:t></w:r>`);
    expect(readRuns(xml)[0].hidden).toBe(false);
    expect(makeAllVisible(xml).changed).toBe(false);
  });

  it("throws on fatally malformed OOXML so the orchestrator can skip it (M5)", () => {
    expect(() => readRuns(`<w:p xmlns:w="${W_NS}"><w:r><w:t>oops`)).toThrow();
  });
});

// ---------------------------------------------------------------------------
// Bridge separators across hidden gaps (wet-test bug 1).
// ---------------------------------------------------------------------------
/** Visible (non-hidden) text after a transform — what the reader sees condensed. */
const visibleText = (xml: string): string =>
  readRuns(xml)
    .filter((r) => !r.hidden)
    .map((r) => r.text)
    .join("");
/** Full text of every run — must stay byte-identical for lossless reversibility. */
const allText = (xml: string): string =>
  readRuns(xml)
    .map((r) => r.text)
    .join("");

describe("applyRunVisibility — bridge splits", () => {
  it("moves one leading space out of a hidden run into a visible sibling", () => {
    const xml = para(
      run("radiation", { highlight: "cyan" }) + run(" are a threat, ") + run("would", { highlight: "cyan" })
    );
    const out = applyRunVisibility(xml, [false, true, false], false, [{ index: 1, side: "lead" }]);
    expect(out.changed).toBe(true);
    // Reads with a single bridging space, not fused.
    expect(visibleText(out.xml)).toBe("radiation would");
    // Lossless: concatenation of ALL runs equals the original text exactly.
    expect(allText(out.xml)).toBe("radiation are a threat, would");
  });

  it("moves one trailing space when asked (trail side)", () => {
    const xml = para(run("a", { highlight: "cyan" }) + run(" mid ") + run("b", { highlight: "cyan" }));
    const out = applyRunVisibility(xml, [false, true, false], false, [{ index: 1, side: "trail" }]);
    expect(visibleText(out.xml)).toBe("a b");
    expect(allText(out.xml)).toBe("a mid b");
  });

  it("Show All after a bridge split restores the exact original text (lossless)", () => {
    const original = para(
      run("radiation", { highlight: "cyan" }) + run(" are a threat, ") + run("would", { highlight: "cyan" })
    );
    const hidden = applyRunVisibility(original, [false, true, false], false, [{ index: 1, side: "lead" }]).xml;
    const shown = makeAllVisible(hidden).xml;
    expect(readRuns(shown).some((r) => r.hidden)).toBe(false);
    expect(allText(shown)).toBe("radiation are a threat, would");
  });

  it("does not split when the requested run has no boundary space (no-op, faithful)", () => {
    const xml = para(run("a", { highlight: "cyan" }) + run("/") + run("b", { highlight: "cyan" }));
    const out = applyRunVisibility(xml, [false, true, false], false, [{ index: 1, side: "lead" }]);
    expect(visibleText(out.xml)).toBe("ab"); // genuinely fused; source had no space
  });

  it("splits a hidden run in three to expose an interior space (interior side)", () => {
    const xml = para(run("a", { highlight: "cyan" }) + run(", mid here.") + run("b", { highlight: "cyan" }));
    // Expose the space between "mid" and "here." at offset 5 (", mid| here.").
    const out = applyRunVisibility(xml, [false, true, false], false, [{ index: 1, side: "interior", offset: 5 }]);
    expect(out.changed).toBe(true);
    expect(visibleText(out.xml)).toBe("a b"); // before/after stay hidden, one space visible
    expect(allText(out.xml)).toBe("a, mid here.b"); // lossless: byte-identical concatenation
  });

  it("keeps a surrogate-pair char intact across an interior split (UTF-16 offset is consistent)", () => {
    // The whole interior path indexes UTF-16 units (planner offset, touched map, ooxml
    // slice), and the offset always points at a BMP space, so a 2-unit char before it is
    // never cleaved. "🦂 mid here." — 🦂 = idx 0-1, spaces at idx 2 and 6; split at 6.
    const xml = para(run("a", { highlight: "cyan" }) + run("\u{1F982} mid here.") + run("b", { highlight: "cyan" }));
    const out = applyRunVisibility(xml, [false, true, false], false, [{ index: 1, side: "interior", offset: 6 }]);
    expect(out.changed).toBe(true);
    expect(visibleText(out.xml)).toBe("a b");
    expect(allText(out.xml)).toBe("a\u{1F982} mid here.b"); // emoji preserved, lossless
  });
});

describe("full pipeline — condensed view end-to-end (wet-test bugs 1 & 2)", () => {
  const cyan = new Set<string>(["cyan"]);

  // Mirror classifyParagraph's body path: keep policy -> cross-gap bridge -> apply.
  const condense = (xml: string): { visible: string; xml: string; changed: boolean } => {
    const runs = readRuns(xml);
    const keep = computeRunKeepFlags(runs, cyan);
    const { extraKeep, splits } = planCrossGapSeparators(runs, keep);
    for (const i of extraKeep) keep[i] = true;
    const out = applyRunVisibility(xml, keep.map((k) => !k), false, splits);
    return { visible: visibleText(out.xml), xml: out.xml, changed: out.changed };
  };

  it("renders the real 2ac sentence with restored inter-chunk spaces", () => {
    // Run boundaries verbatim from samples/[small] 2ac---ndca---semis.docx.
    const xml = para(
      run("ultraviolet radiation", { highlight: "cyan" }) +
        run(" are a constant threat, ") +
        run("would", { highlight: "cyan" }) +
        run(" seem to ") +
        run("give", { highlight: "cyan" }) +
        run(" these ") +
        run("microbes an ", { highlight: "cyan" }) +
        run("advantage in surviving", { highlight: "cyan" })
    );
    expect(condense(xml).visible).toBe(
      "ultraviolet radiation would give microbes an advantage in surviving"
    );
  });

  it("is convergent: re-condensing the hidden output yields identical visible text", () => {
    const xml = para(
      run("ultraviolet radiation", { highlight: "cyan" }) +
        run(" are a constant threat, ") +
        run("would", { highlight: "cyan" })
    );
    const first = condense(xml);
    expect(first.visible).toBe("ultraviolet radiation would");
    // Re-hide over the already-split OOXML must not split again or change the view.
    const second = condense(first.xml);
    expect(second.visible).toBe("ultraviolet radiation would");
    // ...and must be a genuine no-op write (idempotent): the bridge-space run is
    // rescued by extraKeep, not re-split, and no rPr churn is emitted.
    expect(second.changed).toBe(false);
    // No run inflation across re-hide: same <w:r> count, exactly one bridge space.
    expect((first.xml.match(/<w:r\b/g) || []).length).toBe((second.xml.match(/<w:r\b/g) || []).length);
  });

  it("hides a non-highlighted clause glued to a kept word (bug 2 end-to-end)", () => {
    // Verbatim run boundaries from samples/[small] 2ac---ndca---semis.docx: only
    // "large desert animals" is highlighted; ", such as scorpions." is one 8pt run.
    // The OLD whole-word rule kept that whole run (the comma completes "animals,"),
    // leaving "such as scorpions" visible. It must now be hidden, while the leading
    // space of the intervening " certain " still bridges "for" -> "large".
    const xml = para(
      run("for", { highlight: "cyan" }) +
        run(" certain ") +
        run("large desert animals", { highlight: "cyan" }) +
        run(", such as scorpions.")
    );
    const out = condense(xml);
    expect(out.visible).toBe("for large desert animals");
    expect(out.visible).not.toContain("scorpions");
    // Lossless reverse: Show All restores the exact original sentence, comma and all.
    expect(visibleText(makeAllVisible(out.xml).xml)).toBe(
      "for certain large desert animals, such as scorpions."
    );
  });

  it("exposes an interior space so chunks around a glued clause don't fuse (bug 2 follow-up)", () => {
    // ", such as scorpions." starts "," and ends "." — no boundary space — so the bridge
    // does a convergence-safe 3-way split at the space between the untouched words "such"
    // and "as": ", such"(hidden) " "(visible) "as scorpions."(hidden). The chunks read
    // "large desert animals would", the clause stays hidden, the concatenation is
    // byte-identical (lossless), and Re-hide is a genuine no-op (idempotent).
    const xml = para(
      run("large desert animals", { highlight: "cyan" }) +
        run(", such as scorpions.") +
        run("would", { highlight: "cyan" })
    );
    const first = condense(xml);
    expect(first.visible).not.toContain("scorpions");
    expect(first.visible).toBe("large desert animals would");
    // Lossless: concatenation of ALL runs is byte-identical to the original.
    expect(allText(first.xml)).toBe("large desert animals, such as scorpions.would");
    // Convergent + idempotent: re-condensing reproduces the view and changes nothing.
    const second = condense(first.xml);
    expect(second.visible).toBe("large desert animals would");
    expect(second.changed).toBe(false);
  });

  it("shows only highlighted text and keeps punctuation flush (waite-24 'society')", () => {
    // "myopic"(hl) + " society"(not hl) + ". People are"(hl): " society" is hidden, and no
    // space is inserted before the period because ". People are" starts with a hugging
    // mark -> reads "myopic. People are" (not "myopic society…" and not "myopic . People").
    const xml = para(
      run("myopic", { highlight: "cyan" }) +
        run(" society") +
        run(". People are", { highlight: "cyan" })
    );
    expect(condense(xml).visible).toBe("myopic. People are");
  });

  it("keeps the space between fused highlighted fragments (brauner-18 'reduc x')", () => {
    // "reduc"(hl) + "e e"(not hl: 'e' of reduce + space + 'e' of extinction) + "x"(hl):
    // only "reduc"/"x" are highlighted; the gap's interior space is exposed so the view
    // reads "reduc x", not "reducx".
    const xml = para(
      run("reduc", { highlight: "cyan" }) + run("e e") + run("x", { highlight: "cyan" })
    );
    expect(condense(xml).visible).toBe("reduc x");
  });
});
