// Color-codec suite for google-docs/src/core/color.ts (Loop 003 exec-review MAJOR).
//
// WHY THIS SUITE EXISTS: encodeRgbColor is the NEW shared rgb<->hex math that the
// analytics signature, the pocket border, and the keeper's self-recognition all
// depend on, yet the only assertions on it today are tautological (gdocsStyles /
// gdocsAnalytics check it against the SAME constant the code derives from, so a
// wrong codec and a wrong constant would agree and pass). This suite pins the
// codec against HAND-COMPUTED LITERAL objects — values derived by hand from the
// "#rrggbb" math, never read back from the implementation or an engine constant —
// so a regression in the byte<->float conversion, the zero-channel omission, or
// the malformed-hex default fails here loudly.
//
// FALSIFIABILITY: each expected object below is the proto3 wire shape worked out
// independently (red/green/blue = byte/255, zero channels dropped). If the codec
// stopped omitting zeros, started rounding wrong, or threw on junk, these break.
import { decodeOptionalColor, encodeRgbColor } from "../google-docs/src/core/color";

// ---------------------------------------------------------------------------
// 1. encodeRgbColor — literal wire shapes (independent of any engine constant)
// ---------------------------------------------------------------------------

describe("encodeRgbColor — hex to the Docs OptionalColor wire shape", () => {
  it("black is the EMPTY rgbColor (the pocketBorderSide equivalence styles.ts relies on)", () => {
    // All three channels are zero, and proto3 parity drops zero channels, so
    // the entire rgbColor collapses to {}. styles.ts hand-rolled exactly this
    // for the pocket border before centralizing on the codec; if the omission
    // broke, the border color object would carry red:0/green:0/blue:0 and this
    // literal would no longer match.
    expect(encodeRgbColor("#000000")).toEqual({ color: { rgbColor: {} } });
  });

  it("#0b5396 (the analytics navy) carries all three channels as byte/255 floats", () => {
    // 0x0b=11, 0x53=83, 0x96=150 — divided by 255 with NO omission since none is
    // zero. Hand-computed; deliberately NOT sourced from ANALYTICS_FG_HEX so a
    // drift in the constant cannot mask a drift in the codec (the exec-review
    // MAJOR: the existing analytics assertions are circular on that constant).
    expect(encodeRgbColor("#0b5396")).toEqual({
      color: { rgbColor: { red: 11 / 255, green: 83 / 255, blue: 150 / 255 } }
    });
  });

  it("a MIXED-zero color omits ONLY its zero channels (#00ff00 -> green:1 alone)", () => {
    // 0x00 red and 0x00 blue are dropped; 0xff green is 255/255 = exactly 1.
    // Pins the per-channel (not all-or-nothing) omission rule.
    expect(encodeRgbColor("#00ff00")).toEqual({ color: { rgbColor: { green: 1 } } });
  });

  it("malformed hex yields the all-omitted (opaque black) shape, never throws", () => {
    // The regex rejects non-#rrggbb input and the codec returns the empty
    // rgbColor rather than raising — the conservative default for callers that
    // pass known constants. Several junk forms, all collapsing to {}.
    for (const junk of ["not-a-color", "#fff", "#12345", "0b5396", "#0b53961", "", "#gghhii"]) {
      expect(encodeRgbColor(junk)).toEqual({ color: { rgbColor: {} } });
    }
  });

  it("is case-insensitive on the hex digits (uppercase decodes to the same floats)", () => {
    // The /i flag means "#0B5396" must produce the identical payload as the
    // lower-case form; a regression to a case-sensitive match would silently
    // route real uppercase input to the malformed-black default.
    expect(encodeRgbColor("#0B5396")).toEqual(encodeRgbColor("#0b5396"));
  });
});

// ---------------------------------------------------------------------------
// 2. decodeOptionalColor — the raw wire shapes the keeper/parse path sees
// ---------------------------------------------------------------------------

describe("decodeOptionalColor — Docs OptionalColor record to '#rrggbb' | null", () => {
  it("a bare color:{} (opaque, no rgbColor) decodes to black '#000000'", () => {
    // An opaque color with every channel omitted is the proto3 encoding of
    // pure black; the conservative reading of "opaque but unspecified" is black,
    // never null. This is the exact pair to encodeRgbColor('#000000') = {rgbColor:{}}.
    expect(decodeOptionalColor({ color: {} })).toBe("#000000");
  });

  it("an ABSENT color (inherited) and an explicit 'no color' both decode to null", () => {
    // No `color` key = inherited; `color` present but unset would be transparent.
    // Both are "not an opaque color", which the codec reports as null (distinct
    // from the bare-color black above).
    expect(decodeOptionalColor({})).toBeNull(); // inherited
    expect(decodeOptionalColor(undefined)).toBeNull(); // field absent entirely
    expect(decodeOptionalColor({ color: null })).toBeNull(); // explicit none
  });

  it("a populated rgbColor decodes its float channels back to the hex bytes", () => {
    // The decode direction of the navy, asserted on a RAW literal (not the
    // encoder's output) so the two directions are pinned independently.
    expect(
      decodeOptionalColor({ color: { rgbColor: { red: 11 / 255, green: 83 / 255, blue: 150 / 255 } } })
    ).toBe("#0b5396");
  });
});

// ---------------------------------------------------------------------------
// 3. Round-trip: decode(encode(x)) === x — the keeper-recognizes-its-own-write link
// ---------------------------------------------------------------------------

describe("decodeOptionalColor(encodeRgbColor(x)) round-trips exactly", () => {
  // WHY THIS IS THE LOAD-BEARING TEST: analytic-ify WRITES encodeRgbColor(navy)
  // and the keeper later RECOGNIZES an analytics run by decoding the doc back to
  // a hex and comparing strings. If the byte/255 float did not decode (via
  // channelByte's round) to the originating byte, the keeper would fail to see
  // its own write and Delete-analytics would orphan the color. These four hexes
  // exercise the navy itself, a one-byte NEAR-MISS of it, all-zero black, and a
  // mixed-zero color — together covering the omit / round / boundary cases.
  for (const hex of ["#0b5396", "#0b5394", "#000000", "#00ff00"]) {
    it(`${hex} survives encode then decode`, () => {
      expect(decodeOptionalColor(encodeRgbColor(hex))).toBe(hex);
    });
  }

  it("the NEAR-MISS decodes to itself, NOT to the navy (the keeper would reject it)", () => {
    // #0b5394 differs from the analytics navy #0b5396 by one blue byte (0x94 vs
    // 0x96). The round-trip must preserve that distinction or the keeper's exact
    // string-equality detection would false-positive on a neighbouring color.
    expect(decodeOptionalColor(encodeRgbColor("#0b5394"))).toBe("#0b5394");
    expect(decodeOptionalColor(encodeRgbColor("#0b5394"))).not.toBe("#0b5396");
  });
});
