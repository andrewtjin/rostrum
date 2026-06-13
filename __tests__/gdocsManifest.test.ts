// gdocs rangeNames codec suite (plan A2 / S2). The manifest for hidden text
// lives in NamedRange NAMES inside the document, so this codec is the single
// point where a corrupted byte can mis-restore a user's sizes — every test
// here is behavioral (encode/decode round-trips and refusals), no snapshots.

import {
  RSTM_PREFIX,
  encodeSizeEntries,
  encodeSpacingName,
  decodeRangeName,
  isRstmName,
  isKnownRstmName
} from "../google-docs/src/core/rangeNames";
import { NAME_MAX } from "../google-docs/src/core/constants";
import { RleEntry } from "../google-docs/src/core/types";

/** Shorthand: one RLE entry (null size = inherited, encoded "i"). */
function rle(count: number, sizePt: number | null = null): RleEntry {
  return { count, sizePt };
}

/** Sum the chars an entry list covers (negative/zero counts cover nothing). */
function totalChars(entries: RleEntry[]): number {
  return entries.reduce((sum, e) => sum + Math.max(0, e.count), 0);
}

/**
 * Independent oracle for the canonical form (drop non-positive counts, THEN
 * coalesce adjacent same-size entries) — re-derived here ON PURPOSE rather
 * than exported from the codec, so the property test cannot inherit an
 * implementation bug in the codec's own canonicalizer.
 */
function canonicalOracle(entries: RleEntry[]): RleEntry[] {
  const out: RleEntry[] = [];
  for (const e of entries) {
    if (e.count <= 0) continue;
    const prev = out[out.length - 1];
    if (prev !== undefined && prev.sizePt === e.sizePt) prev.count += e.count;
    else out.push({ ...e });
  }
  return out;
}

/**
 * Decode every emitted piece as a sizes record, asserting per piece that the
 * decoded counts sum to its declared charCount, and return the concatenation —
 * the "restore stays exact across splits" contract reduced to data equality.
 */
function decodeAllPieces(pieces: { name: string; charCount: number }[]): RleEntry[] {
  const all: RleEntry[] = [];
  for (const piece of pieces) {
    const decoded = decodeRangeName(piece.name);
    if (decoded === null || decoded.kind !== "sizes") {
      throw new Error(`piece failed to decode as sizes: ${piece.name}`);
    }
    expect(totalChars(decoded.entries)).toBe(piece.charCount);
    all.push(...decoded.entries);
  }
  return all;
}

describe("RSTM_PREFIX (wire contract)", () => {
  it("pins the exact prefix shipped documents depend on", () => {
    // This string is persisted inside real docs; changing it orphans every
    // existing manifest. May only ever gain a sibling version, never move.
    expect(RSTM_PREFIX).toBe("rstm:v1:");
  });
});

describe("encodeSizeEntries — single names", () => {
  it("encodes an integer size", () => {
    expect(encodeSizeEntries([rle(42, 11)])).toEqual([{ name: "rstm:v1:42x11", charCount: 42 }]);
  });

  it("encodes a fractional size without trailing zeros", () => {
    expect(encodeSizeEntries([rle(7, 8.5)])).toEqual([{ name: "rstm:v1:7x8.5", charCount: 7 }]);
  });

  it("trims a whole-number float to its integer form (12.0 -> '12')", () => {
    expect(encodeSizeEntries([rle(3, 12.0)])).toEqual([{ name: "rstm:v1:3x12", charCount: 3 }]);
  });

  it("encodes inherit as 'i'", () => {
    expect(encodeSizeEntries([rle(13, null)])).toEqual([{ name: "rstm:v1:13xi", charCount: 13 }]);
  });

  it("joins multiple entries with commas in order", () => {
    expect(encodeSizeEntries([rle(42, 11), rle(7, 8.5), rle(13, null)])).toEqual([
      { name: "rstm:v1:42x11,7x8.5,13xi", charCount: 62 }
    ]);
  });
});

describe("encodeSizeEntries — canonicalization", () => {
  it("coalesces adjacent same-size entries (incl. inherit/inherit)", () => {
    expect(encodeSizeEntries([rle(10, 11), rle(5, 11), rle(3, null), rle(2, null), rle(4, 11)])).toEqual([
      { name: "rstm:v1:15x11,5xi,4x11", charCount: 24 }
    ]);
  });

  it("does NOT merge equal sizes that are not adjacent (RLE is positional)", () => {
    expect(encodeSizeEntries([rle(2, 11), rle(2, 9), rle(2, 11)])).toEqual([
      { name: "rstm:v1:2x11,2x9,2x11", charCount: 6 }
    ]);
  });

  it("drops zero- and negative-count entries", () => {
    expect(encodeSizeEntries([rle(0, 11), rle(5, 12), rle(-3, null)])).toEqual([
      { name: "rstm:v1:5x12", charCount: 5 }
    ]);
  });

  it("coalesces across a dropped zero-width entry (drop happens first)", () => {
    // A zero-width entry between two 11pt runs means those runs ARE contiguous
    // in the document — the canonical form must merge them.
    expect(encodeSizeEntries([rle(5, 11), rle(0, 12), rle(5, 11)])).toEqual([
      { name: "rstm:v1:10x11", charCount: 10 }
    ]);
  });

  it("excludes dropped counts from charCount (never subtracts)", () => {
    expect(encodeSizeEntries([rle(-5, 11), rle(10, 11)])).toEqual([
      { name: "rstm:v1:10x11", charCount: 10 }
    ]);
  });

  it("returns [] for empty input and for all-dropped input", () => {
    expect(encodeSizeEntries([])).toEqual([]);
    expect(encodeSizeEntries([rle(0, 11), rle(-1, null)])).toEqual([]);
  });

  it("never mutates the caller's entry objects", () => {
    const input = [rle(5, 11), rle(5, 11)];
    encodeSizeEntries(input);
    expect(input).toEqual([rle(5, 11), rle(5, 11)]);
  });
});

describe("encodeSizeEntries — NAME_MAX splitting (plan A2 exactness)", () => {
  // Alternating sizes with distinct counts: already canonical (nothing
  // coalesces), long enough to force several pieces.
  const sizes: (number | null)[] = [11, 9.5, null];
  const entries: RleEntry[] = [];
  for (let i = 0; i < 60; i++) entries.push(rle(100 + i, sizes[i % sizes.length]));

  it("splits at entry boundaries, keeps every name <= NAME_MAX, and restores exactly", () => {
    const pieces = encodeSizeEntries(entries);
    expect(pieces.length).toBeGreaterThan(1);

    for (const piece of pieces) {
      expect(piece.name.length).toBeLessThanOrEqual(NAME_MAX);
      expect(piece.name.startsWith(RSTM_PREFIX)).toBe(true);
    }

    // Sum of declared charCounts == total chars in the region.
    expect(pieces.reduce((sum, p) => sum + p.charCount, 0)).toBe(totalChars(entries));

    // Each piece decodes to EXACTLY its consecutive slice of the input —
    // restore across splits loses nothing.
    let cursor = 0;
    for (const piece of pieces) {
      const decoded = decodeRangeName(piece.name);
      if (decoded === null || decoded.kind !== "sizes") {
        throw new Error(`piece failed to decode: ${piece.name}`);
      }
      expect(decoded.entries).toEqual(entries.slice(cursor, cursor + decoded.entries.length));
      expect(totalChars(decoded.entries)).toBe(piece.charCount);
      cursor += decoded.entries.length;
    }
    expect(cursor).toBe(entries.length);
  });

  it("packs greedily — no piece could have absorbed the next piece's first entry", () => {
    // Guards against a degenerate packer that splits early and floods the doc
    // with ranges; the cap is the ONLY reason to start a new name.
    const pieces = encodeSizeEntries(entries);
    for (let i = 0; i + 1 < pieces.length; i++) {
      const next = decodeRangeName(pieces[i + 1].name);
      if (next === null || next.kind !== "sizes") {
        throw new Error(`piece failed to decode: ${pieces[i + 1].name}`);
      }
      const first = next.entries[0];
      const firstToken = `${first.count}x${first.sizePt === null ? "i" : String(first.sizePt)}`;
      expect(pieces[i].name.length + 1 + firstToken.length).toBeGreaterThan(NAME_MAX);
    }
  });

  it("does not split when the whole RLE fits in one name", () => {
    const pieces = encodeSizeEntries([rle(1000, 11), rle(2000, null), rle(3000, 8.5)]);
    expect(pieces).toEqual([{ name: "rstm:v1:1000x11,2000xi,3000x8.5", charCount: 6000 }]);
  });
});

describe("encodeSpacingName / spacing round-trip (plan A12)", () => {
  it.each<[number | null, number | null, string]>([
    [6, 8, "rstm:v1:p:6x8"],
    [null, 0, "rstm:v1:p:ix0"],
    [0, null, "rstm:v1:p:0xi"],
    [null, null, "rstm:v1:p:ixi"],
    [4.5, 12.0, "rstm:v1:p:4.5x12"], // trailing zero trimmed
    [0.25, 0.5, "rstm:v1:p:0.25x0.5"]
  ])("encodes (%s, %s) as %s and round-trips", (above, below, expected) => {
    const name = encodeSpacingName(above, below);
    expect(name).toBe(expected);
    expect(name.length).toBeLessThanOrEqual(NAME_MAX);
    expect(decodeRangeName(name)).toEqual({
      kind: "spacing",
      spaceAbovePt: above,
      spaceBelowPt: below
    });
    expect(isRstmName(name)).toBe(true);
    expect(isKnownRstmName(name)).toBe(true);
  });
});

describe("decodeRangeName — valid names", () => {
  it("decodes a multi-entry sizes name", () => {
    expect(decodeRangeName("rstm:v1:42x11,7x8.5,13xi")).toEqual({
      kind: "sizes",
      entries: [rle(42, 11), rle(7, 8.5), rle(13, null)]
    });
  });

  it("is lenient about non-canonical adjacency (decode never canonicalizes)", () => {
    // A hand-split or legacy name with adjacent equal sizes still restores the
    // same bytes; only ENCODE owns canonical form (re-encoding coalesces).
    expect(decodeRangeName("rstm:v1:5x11,5x11")).toEqual({
      kind: "sizes",
      entries: [rle(5, 11), rle(5, 11)]
    });
    expect(encodeSizeEntries([rle(5, 11), rle(5, 11)])).toEqual([
      { name: "rstm:v1:10x11", charCount: 10 }
    ]);
  });

  it("decodes zero spacing (0pt spacing is meaningful, unlike 0pt font size)", () => {
    expect(decodeRangeName("rstm:v1:p:0x0")).toEqual({
      kind: "spacing",
      spaceAbovePt: 0,
      spaceBelowPt: 0
    });
  });
});

describe("decodeRangeName — malformed inputs ALL decode to null, never throw", () => {
  it.each<[string]>([
    [""],
    ["rstm:v1:"], // prefix with no payload
    ["rstm:v1:sNaN"], // superseded draft grammar (D2's rstm:v1:s<pt>) — not ours
    ["rstm:v1:s11"], // same: the pre-A2 grammar never shipped and must not decode
    ["rstm:v2:9x9"], // future version: owned (isRstmName) but not decodable
    ["docs-internal-x"], // foreign add-on range
    ["rstm:v1:-3x11"], // negative count
    ["rstm:v1:3x11,-2xi"], // negative count mid-list poisons the whole name
    ["rstm:v1:0x11"], // zero count is non-canonical (covers no chars)
    ["rstm:v1:11x"], // trailing garbage: entry missing its size
    ["rstm:v1:x11"], // entry missing its count
    ["rstm:v1:3x0"], // 0pt cannot be an original font size
    ["rstm:v1:3x11junk"], // trailing garbage after a valid-looking entry
    ["rstm:v1:3X11"], // uppercase X — we only ever emit lowercase
    ["rstm:v1:+3x11"], // signs are never emitted
    ["rstm:v1:3x11,"], // trailing comma -> empty entry
    ["rstm:v1:3x11,,4xi"], // empty entry mid-list
    ["rstm:v1: 3x11"], // stray whitespace
    ["rstm:v1:3x11.5.5"], // double decimal point
    ["rstm:v1:3xNaN"], // non-numeric size
    ["RSTM:V1:3x11"], // prefix is case-sensitive (we always emit lowercase)
    ["rstm:"], // family prefix alone carries nothing
    ["rstm:v1:p:"], // spacing marker with no payload
    ["rstm:v1:p:5"], // spacing missing its second value
    ["rstm:v1:p:5x"], // spacing with empty second value
    ["rstm:v1:p:x5"], // spacing with empty first value
    ["rstm:v1:p:5x6x7"], // spacing with too many values
    ["rstm:v1:p:-1x2"], // negative spacing is never emitted
    ["rstm:v1:p:NaNxi"], // non-numeric spacing
    ["rstm:v1:pp:1x2"] // near-miss marker is a (failing) sizes payload
  ])("decodeRangeName(%j) === null", (name) => {
    expect(decodeRangeName(name)).toBeNull();
    expect(isKnownRstmName(name)).toBe(false);
  });

  it("rejects spacing values that overflow parseFloat to Infinity (codec guard)", () => {
    // A pathologically long digit run still matches SPACING_RE, then parseFloat
    // saturates to Infinity; the never-finite guards in decodeSpacing must
    // reject it as not-ours rather than emit Infinity spacing.
    const huge = "9".repeat(400);
    expect(decodeRangeName(`rstm:v1:p:${huge}x1`)).toBeNull(); // spaceAbove overflows
    expect(decodeRangeName(`rstm:v1:p:1x${huge}`)).toBeNull(); // spaceBelow overflows
    expect(isKnownRstmName(`rstm:v1:p:${huge}x1`)).toBe(false);
  });

  it("flags an unknown future version as ours-but-unknown (edge row 16)", () => {
    // The exact bucket Show All must warn about and leave untouched: owned by
    // the rstm family, but this engine cannot restore from it.
    expect(isRstmName("rstm:v2:9x9")).toBe(true);
    expect(decodeRangeName("rstm:v2:9x9")).toBeNull();
    expect(isKnownRstmName("rstm:v2:9x9")).toBe(false);
  });
});

describe("isRstmName (family ownership, version-agnostic)", () => {
  it.each<[string, boolean]>([
    ["rstm:v1:42x11", true],
    ["rstm:v1:p:6x8", true],
    ["rstm:v2:9x9", true], // future versions are still OURS (warn, don't sweep)
    ["rstm:", true], // degenerate but family-prefixed
    ["rstm:experimental", true],
    ["docs-internal-x", false],
    ["", false],
    [" rstm:v1:3x11", false], // leading space: not a name we ever emitted
    ["Rstm:v1:3x11", false] // case-sensitive on purpose
  ])("isRstmName(%j) === %s", (name, expected) => {
    expect(isRstmName(name)).toBe(expected);
  });
});

describe("fuzz: ~500 random entry lists round-trip exactly (property)", () => {
  /** Tiny deterministic PRNG (mulberry32) so a failing iteration replays
   * exactly on re-run — flaky fuzz is worse than no fuzz. */
  function mulberry32(seed: number): () => number {
    let a = seed >>> 0;
    return () => {
      a = (a + 0x6d2b79f5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const rand = mulberry32(0xc0ffee);
  const randInt = (lo: number, hi: number): number => lo + Math.floor(rand() * (hi - lo + 1));

  /** Realistic size palette: inherited ~25%, else 1–100pt in quarter-point
   * steps (Docs reports fractional point magnitudes; .25 grid covers them). */
  function randSize(): number | null {
    if (rand() < 0.25) return null;
    return randInt(4, 400) / 4;
  }

  it("encode -> decode reproduces the canonical form for every random list", () => {
    for (let i = 0; i < 500; i++) {
      // Up to 40 entries with counts up to 5000 chars: long lists routinely
      // overflow NAME_MAX, so the split path is fuzzed too, not just the happy
      // single-name path. ~15% degenerate counts exercise the drop rule.
      const len = randInt(0, 40);
      const entries: RleEntry[] = [];
      for (let j = 0; j < len; j++) {
        const count = rand() < 0.15 ? randInt(-3, 0) : randInt(1, 5000);
        entries.push(rle(count, randSize()));
      }

      const canonical = canonicalOracle(entries);
      const pieces = encodeSizeEntries(entries);

      if (canonical.length === 0) {
        // Nothing to anchor -> nothing emitted.
        expect({ i, pieces }).toEqual({ i, pieces: [] });
        continue;
      }

      // Every piece respects the cap and is recognized as ours.
      for (const piece of pieces) {
        expect(piece.name.length).toBeLessThanOrEqual(NAME_MAX);
        expect(isRstmName(piece.name)).toBe(true);
        expect(isKnownRstmName(piece.name)).toBe(true);
      }

      // Splits lose nothing: concatenated decode == independent canonical
      // oracle, and the declared char coverage matches the region total.
      // (i is folded into the asserted value so a failure names its iteration.)
      expect({ i, entries: decodeAllPieces(pieces) }).toEqual({ i, entries: canonical });
      expect(pieces.reduce((sum, p) => sum + p.charCount, 0)).toBe(totalChars(canonical));
    }
  });

  it("spacing names round-trip for random above/below values (incl. i and 0)", () => {
    for (let i = 0; i < 100; i++) {
      // null ~30%, else 0–60pt in quarter-point steps (0 included on purpose —
      // zeroed spacing is exactly what the collapse class records around).
      const pick = (): number | null => (rand() < 0.3 ? null : randInt(0, 240) / 4);
      const above = pick();
      const below = pick();
      const name = encodeSpacingName(above, below);
      expect({ i, decoded: decodeRangeName(name) }).toEqual({
        i,
        decoded: { kind: "spacing", spaceAbovePt: above, spaceBelowPt: below }
      });
    }
  });
});
