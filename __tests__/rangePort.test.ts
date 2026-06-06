// Adapter-level tests for the RangeScopedPort methods on OfficeWordPort (readActiveRangeOoxml /
// replaceActiveRangeOoxml), driven by a minimal fake Word context. The controller tests use a fake
// RangeScopedPort that bypasses the adapter, so these specifically guard the host-semantics fix: a
// COLLAPSED read uses paragraph.getRange() (the "Whole" range incl. the mark), so the host serializes
// the target paragraph PLUS a trailing empty <w:p>; the adapter must strip it to one paragraph (else
// Shrink's heading refusal and Condense's collapsed no-op break).
import { createOfficeWordPort } from "../src/core/officeWordPort";
import { readFragmentParagraphs } from "../src/core/ooxmlCondense";

const W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const PKG_NS = "http://schemas.microsoft.com/office/2006/xmlPackage";

/** A flat-OPC package wrapping the given body inner XML. */
const pkg = (bodyInner: string): string =>
  `<pkg:package xmlns:pkg="${PKG_NS}"><pkg:part pkg:name="/word/document.xml"><pkg:xmlData>` +
  `<w:document xmlns:w="${W_NS}"><w:body>${bodyInner}</w:body></w:document>` +
  `</pkg:xmlData></pkg:part></pkg:package>`;

/**
 * Build a fake Word.run context. `selText` "" → a collapsed selection. `rangeOoxml` is what BOTH the
 * selection range and the paragraph range report from getOoxml (collapsed paths use the paragraph
 * range). `outlines` are the outline levels of the range's paragraphs (Word's 1-based numbers).
 */
function makeRunner(opts: { selText: string; rangeOoxml: string; outlines: number[] }) {
  const inserts: string[] = [];
  const paragraphRange: any = {
    load() {},
    paragraphs: { load() {}, items: opts.outlines.map((lvl) => ({ load() {}, outlineLevel: lvl })) },
    getOoxml: () => ({ value: opts.rangeOoxml }),
    insertOoxml: (xml: string) => inserts.push(xml)
  };
  const selectionParas = opts.outlines.map((lvl) => ({ load() {}, outlineLevel: lvl, getRange: () => paragraphRange }));
  const selection: any = {
    text: opts.selText,
    load() {},
    paragraphs: { load() {}, items: selectionParas },
    getOoxml: () => ({ value: opts.rangeOoxml }),
    getRange: () => paragraphRange,
    insertOoxml: (xml: string) => inserts.push(xml)
  };
  const ctx: any = { document: { getSelection: () => selection }, sync: async () => undefined };
  const runner = <T,>(batch: (c: any) => Promise<T>): Promise<T> => batch(ctx);
  return { runner, inserts };
}

describe("readActiveRangeOoxml", () => {
  it("collapsed read strips the trailing empty <w:p> to ONE paragraph (heading-refusal fix)", async () => {
    const twoPara = pkg(`<w:p><w:r><w:t>card body</w:t></w:r></w:p><w:p/>`);
    const { runner } = makeRunner({ selText: "", rangeOoxml: twoPara, outlines: [3] });
    const port = createOfficeWordPort({ runner });
    const read = await port.readActiveRangeOoxml();
    expect(read.collapsed).toBe(true);
    expect(readFragmentParagraphs(read.ooxml)).toHaveLength(1); // phantom <w:p/> removed
    expect(read.outlineLevels).toEqual([2]); // 1-based 3 → canonical 0-based 2
  });

  it("non-collapsed selection is NOT stripped (keeps every selected paragraph)", async () => {
    const twoPara = pkg(`<w:p><w:r><w:t>one</w:t></w:r></w:p><w:p><w:r><w:t>two</w:t></w:r></w:p>`);
    const { runner } = makeRunner({ selText: "one two", rangeOoxml: twoPara, outlines: [10, 10] });
    const port = createOfficeWordPort({ runner });
    const read = await port.readActiveRangeOoxml();
    expect(read.collapsed).toBe(false);
    expect(readFragmentParagraphs(read.ooxml)).toHaveLength(2); // both paragraphs preserved
    expect(read.outlineLevels).toEqual([null, null]); // body text
  });
});

describe("replaceActiveRangeOoxml", () => {
  it("writes back via insertOoxml on the active range", async () => {
    const { runner, inserts } = makeRunner({ selText: "x", rangeOoxml: pkg("<w:p/>"), outlines: [10] });
    const port = createOfficeWordPort({ runner });
    await port.replaceActiveRangeOoxml("<NEW/>");
    expect(inserts).toEqual(["<NEW/>"]);
  });
});
