// Shared OptionalColor codec (Loop 003 plan-review DRY fix) — the ONE place the
// Docs API's rgbColor wire shape is decoded to / encoded from a "#rrggbb" hex.
//
// WHY THIS MODULE EXISTS: before this, the rgb<->hex math lived only as a
// DECODE in parse.ts (decodeBackgroundHex + channelByte/byteHex) and as a
// hand-rolled ENCODE inline in styles.ts (the pocket border's black
// `{rgbColor:{}}`). Adding the analytics foreground (decode for the keeper,
// encode for the analytic-ify write) would have made a THIRD copy, so the
// codec is centralized here: parse.ts decodes backgroundColor AND foregroundColor
// through decodeOptionalColor (they differ only by which key they pluck), and
// styles.ts encodes both the pocket border and the analytics color through
// encodeRgbColor. Detection (string ===), keeping, and the apply write are then
// provably consistent — the single source the analytics signature depends on.
//
// PURITY: total, deterministic functions over given values; no Apps Script, no I/O.

/** Plain-object check (arrays are NOT records) — local copy so this module has
 * no import cycle with parse.ts (which imports FROM here). */
function isRecord(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

/** Finite-number read with a fallback (NaN/Infinity/junk fall back). */
function finiteOr(v: unknown, fallback: number): number {
  return typeof v === "number" && Number.isFinite(v) ? v : fallback;
}

/** One rgb channel (0.0-1.0 float) to its 0-255 byte. ABSENT CHANNELS ARE 0.0:
 * real payloads omit zero channels (#ffff00 arrives as {red:1,green:1}), so the
 * missing-key path is the NORMAL path. Junk / out-of-range floats clamp. */
export function channelByte(v: unknown): number {
  const n = finiteOr(v, 0);
  return Math.round(Math.min(1, Math.max(0, n)) * 255);
}

/** A byte (0-255) as a 2-digit lower-case hex pair. */
export function byteHex(b: number): string {
  return b.toString(16).padStart(2, "0");
}

/**
 * A Docs OptionalColor record -> lower-case "#rrggbb" | null. The API's
 * three-layer nesting encodes real states: the field absent = inherited (null),
 * present with `color` unset = explicitly transparent / none (null), `color`
 * set = an opaque color whose omitted channels are zeros. A bare `color: {}`
 * therefore decodes to "#000000" (it claims opacity; the conservative reading
 * of an opaque unknown is black, never "no color"). This is exactly the body
 * the old decodeBackgroundHex carried — now shared so foreground decodes
 * identically.
 */
export function decodeOptionalColor(optional: unknown): string | null {
  if (!isRecord(optional)) return null;
  const color = optional.color;
  if (!isRecord(color)) return null;
  const rgb = isRecord(color.rgbColor) ? color.rgbColor : {};
  return `#${byteHex(channelByte(rgb.red))}${byteHex(channelByte(rgb.green))}${byteHex(channelByte(rgb.blue))}`;
}

/**
 * "#rrggbb" -> the Docs OptionalColor wire shape `{ color: { rgbColor: {…} } }`
 * with ZERO channels OMITTED (proto3 parity — the exact shape the API itself
 * echoes back, and the exact shape styles.ts's pocket border hand-rolls:
 * encodeRgbColor("#000000") === { color: { rgbColor: {} } }). Channels are
 * floats = byte/255, so a value written here reads back (via channelByte) as the
 * same byte — the analytics color round-trips exactly. A malformed hex yields
 * the all-omitted (black) shape rather than throwing (callers pass known
 * constants; the defensive default is the conservative opaque black).
 */
export function encodeRgbColor(hex: string): {
  color: { rgbColor: { red?: number; green?: number; blue?: number } };
} {
  const m = /^#([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(hex);
  const rgb: { red?: number; green?: number; blue?: number } = {};
  if (m !== null) {
    const r = parseInt(m[1], 16) / 255;
    const g = parseInt(m[2], 16) / 255;
    const b = parseInt(m[3], 16) / 255;
    // Omit zero channels (proto3 wire parity); only non-zero channels are sent.
    if (r !== 0) rgb.red = r;
    if (g !== 0) rgb.green = g;
    if (b !== 0) rgb.blue = b;
  }
  return { color: { rgbColor: rgb } };
}
