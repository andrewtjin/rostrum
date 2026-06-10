// Pure, host-free OUTLINE-LEVEL RESOLUTION from a document's `styles.xml` plus a
// paragraph's own OOXML — no Office.js, no DOM, just string/regex inspection so it runs
// identically in Node tests, the task-pane browser, and (the reason this module exists)
// the whole-body classify path in officeWordPort.ts.
//
// WHY THIS IS A SHARED `core/` MODULE NOW. The whole-body strategy (avenue ⑦) classifies
// each paragraph DIRECTLY from the `body.getOoxml()` package and therefore must resolve a
// paragraph's outline level WITHOUT a Word proxy — the host reports `canGetStyles:false`,
// so the `Paragraph.outlineLevel` number is the only thing the live API exposes, and the
// pure whole-body path doesn't load proxies at all. The getOoxml package, however, already
// bundles `word/styles.xml`, so the level is fully recoverable from the package itself.
// This resolver was first written and proven inside the real-document test harness
// (`__tests__/realDocs.ts`); promoting it here gives the adapter and the test harness ONE
// implementation instead of two divergent copies.
//
// STRUCTURAL SIGNAL, NOT STYLE NAMES (decision #7 / LESSONS #3). Resolution is purely
// structural: a direct `<w:pPr><w:outlineLvl>` wins, else the paragraph's `<w:pStyle>`
// effective level via the `basedOn` cascade in `styles.xml`. That cascade is exactly how
// Word itself resolves a style's outline level — so `Analytics`/`Analytic` → `Heading4`
// → level 3 are kept with ZERO hardcoded style-name enumeration. There is no friendly-name
// allowlist here; the only knowledge baked in is the OOXML schema (outlineLvl is 0-based,
// `Paragraph.outlineLevel` is 1-based — see `outlineNumberOf`).

/** A single paragraph style's outline contribution, parsed from `styles.xml`. */
export interface StyleDef {
  /** This style's own OOXML outline level (0-based), or null if it declares none. */
  outlineLvl: number | null;
  /** The styleId this one is `basedOn`, or null. */
  basedOn: string | null;
}

/**
 * LAST-RESORT heading-name → 0-based outline level map. Word carries the outline level of its
 * BUILT-IN heading styles in the latent/default style table and does not always re-emit an explicit
 * `<w:outlineLvl>` into a document's `styles.xml` — so the structural `basedOn` cascade can come up
 * empty for a perfectly ordinary `Heading 1` (the live `Paragraph.outlineLevel` would still report
 * it as a heading). Without this fallback the pure whole-body path would HIDE such headings (the
 * adversarial-review C-1 finding). This map is consulted ONLY after the inline level and the
 * structural cascade both fail, so lesson #3 (prefer structural signal over friendly names) still
 * holds — names are the final safety net the original avenue-⑦ plan called for, forced by
 * `canGetStyles:false`. Keys are normalized (lower-cased, spaces stripped). `Analytics`/`Analytic`
 * are the navy debate-tag styles, conventionally a kept heading level (they resolve to Heading4 = 3
 * structurally in the corpus; mapped here for docs whose styles.xml omits that linkage).
 */
const HEADING_NAME_LEVELS = new Map<string, number>([
  ["heading1", 0],
  ["heading2", 1],
  ["heading3", 2],
  ["heading4", 3],
  ["heading5", 4],
  ["heading6", 5],
  ["heading7", 6],
  ["heading8", 7],
  ["heading9", 8],
  ["analytics", 3],
  ["analytic", 3]
]);

/** Normalize a styleId for the heading-name fallback (case- and space-insensitive). */
function normalizeStyleId(id: string): string {
  return id.toLowerCase().replace(/\s+/g, "");
}

/**
 * The heading-name fallback level (0-based) for a style, following the `basedOn` chain so a custom
 * style based on a built-in heading (`CardTag` → `Heading4`) is caught even when neither declares an
 * explicit `<w:outlineLvl>`. Cycle-safe. Returns null when no style in the chain has a known heading
 * name. Used ONLY as the last resort (see HEADING_NAME_LEVELS).
 */
function headingNameLevel(styleId: string, defs: Map<string, StyleDef>): number | null {
  const seen = new Set<string>();
  let cur: string | null = styleId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const lvl = HEADING_NAME_LEVELS.get(normalizeStyleId(cur));
    if (lvl !== undefined) return lvl;
    const d = defs.get(cur);
    cur = d ? d.basedOn : null;
  }
  return null;
}

/**
 * Parse `styleId → { outlineLvl, basedOn }` from a `styles.xml` part. Regex-based (not
 * DOM): `styles.xml` is flat and well-formed as Word emits it, and a string scan keeps
 * this dependency-free and fast on the whole part. A style with no `<w:outlineLvl>` and
 * no `<w:basedOn>` still gets an entry (both null) so the cascade can terminate cleanly.
 */
export function parseStyleDefs(stylesXml: string): Map<string, StyleDef> {
  const defs = new Map<string, StyleDef>();
  const styleRe = /<w:style\b[^>]*\bw:styleId="([^"]+)"[^>]*>([\s\S]*?)<\/w:style>/g;
  let m: RegExpExecArray | null;
  while ((m = styleRe.exec(stylesXml)) !== null) {
    const inner = m[2];
    // Attribute-order tolerant (review C-2): never assume `w:val` is the first attribute, since a
    // producer can emit other attrs first or reorder them on round-trip — a hard `\s+w:val` miss
    // would silently drop the level and HIDE the heading.
    const ol = /<w:outlineLvl\b[^>]*\bw:val="(\d+)"/.exec(inner);
    const based = /<w:basedOn\b[^>]*\bw:val="([^"]+)"/.exec(inner);
    defs.set(m[1], { outlineLvl: ol ? Number(ol[1]) : null, basedOn: based ? based[1] : null });
  }
  return defs;
}

/**
 * A paragraph style's effective OOXML outline level (0-based), following the `basedOn`
 * chain (cycle-safe via a `seen` set). Mirrors how Word resolves a style's outline level
 * through the cascade — e.g. `Analytics` (no own level) → `Heading4` (level 3) → 3.
 * Returns null when no style in the chain declares a level (→ body text).
 */
export function styleOutlineLevel(styleId: string, defs: Map<string, StyleDef>): number | null {
  const seen = new Set<string>();
  let cur: string | null = styleId;
  while (cur && !seen.has(cur)) {
    seen.add(cur);
    const d = defs.get(cur);
    if (!d) return null;
    if (d.outlineLvl !== null) return d.outlineLvl;
    cur = d.basedOn;
  }
  return null;
}

/**
 * The numeric value the LIVE `Paragraph.outlineLevel` would report, resolved from a
 * paragraph's ALREADY-EXTRACTED properties: its inline `<w:outlineLvl>` value (0-based,
 * null when absent) and its `<w:pStyle>` id (null when absent). This is the resolution
 * CASCADE shared by both front ends — the string-regex `outlineNumberOf` below and the
 * node-direct `WholeBodyPackage.headingLevel` (ooxmlPackage.ts), which walks the
 * paragraph's own `<w:pPr>` children instead of serializing the whole subtree just so
 * two regexes can re-find them (the dominant per-paragraph cost on the pure whole-body
 * read). OOXML outline levels are 0-based (0 = Heading 1 … 8 = Heading 9; 9/absent =
 * body), and Word reports them 1-based (Heading 1 = 1 … body = 10) — so the return is
 * normalized to the 1-based convention the adapter already feeds through
 * `normalizeOutlineNumber` (ooxmlPackage.ts), keeping the whole-body path and the proxy
 * path on one downstream normalization.
 */
export function outlineNumberFromProps(
  inlineLvl: number | null,
  styleId: string | null,
  defs: Map<string, StyleDef>
): number {
  // 1) A direct inline <w:outlineLvl> wins.
  if (inlineLvl !== null) {
    return inlineLvl >= 9 ? 10 : inlineLvl + 1;
  }
  // 2) The paragraph style: the STRUCTURAL basedOn cascade first, then the heading-name last resort
  //    (review C-1 — a built-in heading whose styles.xml omits an explicit <w:outlineLvl> would
  //    otherwise resolve to body and be HIDDEN, diverging from the live Paragraph.outlineLevel).
  if (styleId) {
    const structural = styleOutlineLevel(styleId, defs);
    const ol = structural !== null ? structural : headingNameLevel(styleId, defs);
    if (ol !== null) return ol >= 9 ? 10 : ol + 1;
  }
  // 3) Body.
  return 10;
}

/**
 * The numeric value the LIVE `Paragraph.outlineLevel` would report for a paragraph's
 * OOXML string: the two property extractions (attribute-order tolerant — review C-2)
 * feeding the shared cascade above. Kept as the string API for callers that already
 * hold serialized XML (the realDocs test harness); the hot whole-body classify path
 * resolves from the live node instead and never pays this serialization.
 */
export function outlineNumberOf(paraXml: string, defs: Map<string, StyleDef>): number {
  const inline = /<w:outlineLvl\b[^>]*\bw:val="(\d+)"/.exec(paraXml);
  const ps = /<w:pStyle\b[^>]*\bw:val="([^"]+)"/.exec(paraXml);
  return outlineNumberFromProps(inline ? Number(inline[1]) : null, ps ? ps[1] : null, defs);
}
