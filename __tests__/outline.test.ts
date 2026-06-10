// Unit tests for the promoted, shared outline resolver (src/core/outline.ts).
//
// These were previously exercised only indirectly through realDocs.test.ts; now that the
// resolver drives the whole-body classify path (avenue ⑦) it gets direct coverage of every
// branch: inline <w:outlineLvl>, the basedOn cascade (the real Analytics→Heading4 case),
// cycle safety, the 0-based→1-based normalization, and the body fallback.

import {
  parseStyleDefs,
  styleOutlineLevel,
  outlineNumberOf,
  outlineNumberFromProps
} from "../src/core/outline";

// A styles.xml mirroring a real debate doc: Heading1..4 carry their own outlineLvl, and
// Analytics (+ its typo Analytic) inherit level 3 from Heading4 via basedOn. Default has none.
const STYLES_XML = `
<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:style w:type="paragraph" w:styleId="Heading1"><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading2"><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading3"><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="2"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Heading4"><w:basedOn w:val="Normal"/><w:pPr><w:outlineLvl w:val="3"/></w:pPr></w:style>
  <w:style w:type="paragraph" w:styleId="Analytics"><w:basedOn w:val="Heading4"/></w:style>
  <w:style w:type="paragraph" w:styleId="Analytic"><w:basedOn w:val="Analytics"/></w:style>
  <w:style w:type="paragraph" w:styleId="Normal"></w:style>
  <w:style w:type="paragraph" w:styleId="BodyText"><w:basedOn w:val="Normal"/></w:style>
</w:styles>`;

describe("parseStyleDefs", () => {
  it("captures each style's own outlineLvl and basedOn", () => {
    const defs = parseStyleDefs(STYLES_XML);
    expect(defs.get("Heading1")).toEqual({ outlineLvl: 0, basedOn: "Normal" });
    expect(defs.get("Heading4")).toEqual({ outlineLvl: 3, basedOn: "Normal" });
    expect(defs.get("Analytics")).toEqual({ outlineLvl: null, basedOn: "Heading4" });
    expect(defs.get("Normal")).toEqual({ outlineLvl: null, basedOn: null });
  });

  it("returns an empty map for empty/garbage input", () => {
    expect(parseStyleDefs("").size).toBe(0);
    expect(parseStyleDefs("<not-styles/>").size).toBe(0);
  });
});

describe("styleOutlineLevel (basedOn cascade)", () => {
  const defs = parseStyleDefs(STYLES_XML);

  it("returns a style's own declared level", () => {
    expect(styleOutlineLevel("Heading1", defs)).toBe(0);
    expect(styleOutlineLevel("Heading4", defs)).toBe(3);
  });

  it("resolves an inheriting style through basedOn (Analytics → Heading4 → 3)", () => {
    expect(styleOutlineLevel("Analytics", defs)).toBe(3);
  });

  it("resolves a two-hop cascade (Analytic → Analytics → Heading4 → 3)", () => {
    expect(styleOutlineLevel("Analytic", defs)).toBe(3);
  });

  it("returns null when no style in the chain declares a level (body)", () => {
    expect(styleOutlineLevel("BodyText", defs)).toBeNull();
    expect(styleOutlineLevel("Normal", defs)).toBeNull();
  });

  it("returns null for an unknown style id", () => {
    expect(styleOutlineLevel("DoesNotExist", defs)).toBeNull();
  });

  it("is cycle-safe (a basedOn loop terminates, returns null)", () => {
    const cyclic = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="A"><w:basedOn w:val="B"/></w:style>
        <w:style w:styleId="B"><w:basedOn w:val="A"/></w:style>
      </w:styles>`);
    expect(styleOutlineLevel("A", cyclic)).toBeNull();
  });
});

describe("outlineNumberOf (1-based Paragraph.outlineLevel equivalent)", () => {
  const defs = parseStyleDefs(STYLES_XML);

  it("prefers a direct inline <w:outlineLvl> over the paragraph style", () => {
    // Inline level 0 (Heading 1) wins even though the pStyle resolves to body.
    const xml = `<w:p><w:pPr><w:pStyle w:val="BodyText"/><w:outlineLvl w:val="0"/></w:pPr></w:p>`;
    expect(outlineNumberOf(xml, defs)).toBe(1);
  });

  it("falls back to the pStyle's cascade level (Analytics → 3 → reported 4)", () => {
    const xml = `<w:p><w:pPr><w:pStyle w:val="Analytics"/></w:pPr><w:r><w:t>card</w:t></w:r></w:p>`;
    expect(outlineNumberOf(xml, defs)).toBe(4);
  });

  it("maps Heading1..4 to 1..4", () => {
    for (const [style, n] of [["Heading1", 1], ["Heading2", 2], ["Heading3", 3], ["Heading4", 4]] as const) {
      const xml = `<w:p><w:pPr><w:pStyle w:val="${style}"/></w:pPr></w:p>`;
      expect(outlineNumberOf(xml, defs)).toBe(n);
    }
  });

  it("reports body (10) for a plain paragraph with no style or level", () => {
    expect(outlineNumberOf(`<w:p><w:r><w:t>plain</w:t></w:r></w:p>`, defs)).toBe(10);
    const bodyStyled = `<w:p><w:pPr><w:pStyle w:val="BodyText"/></w:pPr></w:p>`;
    expect(outlineNumberOf(bodyStyled, defs)).toBe(10);
  });

  it("clamps an out-of-range inline level (>= 9) to body (10)", () => {
    const xml = `<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:p>`;
    expect(outlineNumberOf(xml, defs)).toBe(10);
  });

  // --- Adversarial-review hardening (C-1 name fallback, C-2 attribute order) ---

  it("keeps a built-in Heading whose styles.xml omits an explicit <w:outlineLvl> (C-1 name fallback)", () => {
    // Word often carries a built-in heading's level in its latent style table, not in styles.xml.
    const lossyDefs = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="Heading1"><w:basedOn w:val="Normal"/></w:style>
        <w:style w:styleId="Normal"></w:style>
      </w:styles>`);
    expect(styleOutlineLevel("Heading1", lossyDefs)).toBeNull(); // structural cascade comes up empty…
    const xml = `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Tag</w:t></w:r></w:p>`;
    expect(outlineNumberOf(xml, lossyDefs)).toBe(1); // …but the name fallback still reports Heading 1
  });

  it("resolves a custom style based on a built-in heading with no explicit level (chain name fallback)", () => {
    const d = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="CardTag"><w:basedOn w:val="Heading4"/></w:style>
        <w:style w:styleId="Heading4"><w:basedOn w:val="Normal"/></w:style>
        <w:style w:styleId="Normal"></w:style>
      </w:styles>`);
    const xml = `<w:p><w:pPr><w:pStyle w:val="CardTag"/></w:pPr></w:p>`;
    expect(outlineNumberOf(xml, d)).toBe(4); // CardTag → Heading4 (name) → level 3 → reported 4
  });

  it("keeps the navy Analytics tag via the name fallback when its basedOn linkage is absent", () => {
    const d = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="Analytics"></w:style>
      </w:styles>`);
    const xml = `<w:p><w:pPr><w:pStyle w:val="Analytics"/></w:pPr></w:p>`;
    expect(outlineNumberOf(xml, d)).toBe(4); // Analytics → level 3 → reported 4 (kept)
  });

  it("tolerates attribute order on <w:outlineLvl> and <w:pStyle> (C-2)", () => {
    expect(outlineNumberOf(`<w:p><w:pPr><w:outlineLvl w:dummy="1" w:val="0"/></w:pPr></w:p>`, defs)).toBe(1);
    const reordered = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="Heading2"><w:outlineLvl w:foo="x" w:val="1"/></w:style>
      </w:styles>`);
    expect(outlineNumberOf(`<w:p><w:pPr><w:pStyle w:bar="y" w:val="Heading2"/></w:pPr></w:p>`, reordered)).toBe(2);
  });

  it("a plain custom body style (not heading-named, no level) stays body", () => {
    const d = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="CardBody"><w:basedOn w:val="Normal"/></w:style>
        <w:style w:styleId="Normal"></w:style>
      </w:styles>`);
    expect(outlineNumberOf(`<w:p><w:pPr><w:pStyle w:val="CardBody"/></w:pPr></w:p>`, d)).toBe(10);
  });
});

describe("outlineNumberFromProps (the shared cascade behind both front ends)", () => {
  // The string resolver above and the node-direct WholeBodyPackage.headingLevel both feed
  // their extracted <w:pPr> signals into this ONE cascade — these tests pin its contract so
  // the two front ends can never drift apart on resolution semantics.
  const defs = parseStyleDefs(STYLES_XML);

  it("a non-null inline level wins over any style", () => {
    expect(outlineNumberFromProps(0, "BodyText", defs)).toBe(1); // inline Heading 1 beats body style
    expect(outlineNumberFromProps(2, "Heading1", defs)).toBe(3); // inline 2 beats Heading1's own 0
  });

  it("falls back to the pStyle cascade (Analytics → Heading4 → 3 → reported 4)", () => {
    expect(outlineNumberFromProps(null, "Analytics", defs)).toBe(4);
  });

  it("falls back to the heading-name last resort when the structural cascade is empty (C-1)", () => {
    const lossy = parseStyleDefs(`
      <w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="Heading1"><w:basedOn w:val="Normal"/></w:style>
        <w:style w:styleId="Normal"></w:style>
      </w:styles>`);
    expect(outlineNumberFromProps(null, "Heading1", lossy)).toBe(1);
  });

  it("reports body (10) when both signals are absent or the style resolves nothing", () => {
    expect(outlineNumberFromProps(null, null, defs)).toBe(10);
    expect(outlineNumberFromProps(null, "BodyText", defs)).toBe(10);
    expect(outlineNumberFromProps(null, "DoesNotExist", defs)).toBe(10);
  });

  it("clamps an out-of-range inline level (>= 9) to body (10)", () => {
    expect(outlineNumberFromProps(9, null, defs)).toBe(10);
    expect(outlineNumberFromProps(9, "Heading1", defs)).toBe(10); // clamp is terminal, no style retry
  });
});
