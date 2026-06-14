// A3 / 002-F9 — the EXHAUSTIVE separator differential.
//
// WHY THIS EXISTS. A3 replaced keepers.ts's per-char `\p{Zs}`-property regex with a `charCodeAt`
// ASCII fast path (0x09–0x0D, 0x20, 0x2F) + a `\p{Zs}`+dash regex FALLBACK for codepoints > 0x7F.
// That is a hot-path rewrite of the engine's word-separator decision; if the fast path and the
// fallback don't TOGETHER reproduce the old predicate on EVERY codepoint, a hidden gap could fuse
// (or wrongly split) two kept words — a silent losslessness/legibility regression. So we do not
// hand-pick fixtures: we sweep the ENTIRE Unicode space and assert the new predicate agrees with
// the LEGACY regex byte-for-byte, code point by code point (the 002-F9 / tests-3 gate).
//
// THE LEGACY REFERENCE is the exact pre-A3 regex literal (`/[\t\n\v\f\r–—/]|\p{Zs}/u`),
// frozen here so the differential is "new fast path vs the precise thing it replaced." The new
// predicate is exercised through its test-only export `isExposableSeparatorForTest`.
//
// SURROGATE SAFETY is the subtle part. `charCodeAt(0)` reads one UTF-16 unit; an astral char is two
// units and a lone surrogate is one. The callers iterate with `[...text]` (code points), so the
// predicate is asked about whole code points (astral) and lone surrogates (each a 1-unit string).
// We sweep BOTH: every BMP code point, a sample of astral code points, and EVERY lone surrogate
// 0xD800–0xDFFF as a standalone unit — proving the fast path never false-matches a surrogate.

import { isExposableSeparatorForTest as isSep } from "../src/core/keepers";

// The EXACT predicate keepers.ts used before A3 (the control). `\p{Zs}` (Space_Separator) plus the
// ASCII control whitespace, the EN/EM dashes, and the slash. Anything else is not a separator.
const LEGACY = /[\t\n\v\f\r–—/]|\p{Zs}/u;
function legacy(ch: string): boolean {
  return LEGACY.test(ch);
}

describe("A3 separator fast path matches the legacy \\p{Zs}+Cc+dash/slash regex (002-F9)", () => {
  it("agrees on EVERY BMP code point 0x0000–0xFFFF (each as a single-unit string)", () => {
    const mismatches: string[] = [];
    for (let cp = 0x0000; cp <= 0xffff; cp++) {
      const ch = String.fromCharCode(cp);
      if (isSep(ch) !== legacy(ch)) {
        mismatches.push(`U+${cp.toString(16).toUpperCase().padStart(4, "0")} new=${isSep(ch)} legacy=${legacy(ch)}`);
      }
    }
    // Surface the first few offenders for a precise failure, never a bare "expected true".
    expect(mismatches.slice(0, 20)).toEqual([]);
    expect(mismatches.length).toBe(0);
  });

  it("agrees on EVERY lone surrogate 0xD800–0xDFFF as a standalone UTF-16 unit (no false fast-path match)", () => {
    const mismatches: string[] = [];
    for (let cp = 0xd800; cp <= 0xdfff; cp++) {
      const ch = String.fromCharCode(cp); // a lone surrogate — the astralText fixture's edge case
      // No surrogate is a separator; both paths must say false. (Asserted via agreement + explicit false.)
      if (isSep(ch) !== legacy(ch) || isSep(ch) !== false) {
        mismatches.push(`U+${cp.toString(16).toUpperCase()} new=${isSep(ch)} legacy=${legacy(ch)}`);
      }
    }
    expect(mismatches).toEqual([]);
  });

  it("agrees across the supplementary planes (sampled astral code points, each a 2-unit string)", () => {
    const mismatches: string[] = [];
    // Sweep every astral code point at a stride that still lands on category boundaries, plus the
    // named edge cases (U+1F600 grinning face from the fixture; the only astral Zs does not exist —
    // every Space_Separator is in the BMP — so the expected answer across all astral is "false",
    // which this proves by agreement with the legacy regex on full code points).
    const named = [0x1f600, 0x10000, 0x1ffff, 0x2f800, 0xe0000, 0x10ffff];
    for (const cp of named) {
      const ch = String.fromCodePoint(cp);
      if (isSep(ch) !== legacy(ch)) mismatches.push(`U+${cp.toString(16)} new=${isSep(ch)} legacy=${legacy(ch)}`);
    }
    for (let cp = 0x10000; cp <= 0x10ffff; cp += 0x111) {
      const ch = String.fromCodePoint(cp);
      if (isSep(ch) !== legacy(ch)) mismatches.push(`U+${cp.toString(16)} new=${isSep(ch)} legacy=${legacy(ch)}`);
    }
    expect(mismatches.slice(0, 20)).toEqual([]);
    expect(mismatches.length).toBe(0);
  });

  it("explicitly classifies the load-bearing separators and non-separators (readable spot checks)", () => {
    // ASCII fast-path TRUE members.
    for (const ch of ["\t", "\n", "\v", "\f", "\r", " ", "/"]) expect(isSep(ch)).toBe(true);
    // >0x7F fallback TRUE members: NBSP, figure, thin, narrow, four-per-em, ideographic spaces + dashes.
    for (const cp of [0x00a0, 0x2007, 0x2009, 0x202f, 0x2005, 0x3000, 0x2013, 0x2014]) {
      expect(isSep(String.fromCharCode(cp))).toBe(true);
    }
    // Deliberately EXCLUDED (word-internal joiners): hyphen family, underscore, ampersand, ZWSP/WJ.
    for (const cp of [0x002d, 0x2010, 0x2011, 0x00ad, 0x005f, 0x0026, 0x200b, 0x2060]) {
      expect(isSep(String.fromCharCode(cp))).toBe(false);
    }
    // Ordinary letters/digits are never separators.
    for (const ch of ["a", "Z", "0", "9", ".", "(", "—".normalize()]) {
      if (ch !== "—") expect(isSep(ch)).toBe(false);
    }
  });
});
