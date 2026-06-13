// RLE-in-name manifest codec (plan A2). A hidden region's restore manifest IS
// the NAME of the NamedRange that anchors it: "rstm:v1:" followed by
// comma-joined "<charCount>x<sizePt|i>" run-length entries recording the
// region's ORIGINAL font sizes ("i" = the run inherited its size and is
// restored by CLEARING fontSize, never by materializing a number). Putting the
// data in the name — not in document properties or any external store — is the
// whole design: NamedRange indexes auto-track user edits, the manifest travels
// with doc copies, and restore needs nothing but the document itself (repo
// lesson #4, plan D2).
//
// Names must stay <= NAME_MAX, so a region whose RLE overflows is split into
// several names AT ENTRY BOUNDARIES; each emitted piece reports how many
// characters it covers so the planner can pair every piece with its exact
// sub-range and restore stays EXACT across splits (case 001-S4, plan A2).
//
// Decoding NEVER throws: real documents carry ranges from other add-ons
// ("docs-internal-..."), from future rstm versions, and from decayed or
// hand-edited manifests. A name we cannot FULLY parse decodes to null and the
// caller treats the range as foreign (unknown rstm versions are left untouched
// and warned about — edge row 16). Parsing is all-or-nothing on purpose: a
// half-trusted manifest could restore the wrong sizes over the wrong spans,
// which is strictly worse than the amber normalize path.

import { DecodedRangeName, RleEntry } from "./types";
import { NAME_MAX } from "./constants";

/**
 * The exact prefix of every range THIS engine version owns. Persisted into
 * shipped documents — changing it would orphan every existing manifest, so it
 * is pinned by test and may only ever gain a sibling (rstm:v2:...), never move.
 */
export const RSTM_PREFIX = "rstm:v1:";

/** The version-agnostic family prefix. Ownership checks use THIS (not
 * RSTM_PREFIX) so Show All can recognize — and deliberately leave untouched —
 * ranges written by future engine versions (plan edge row 16). */
const RSTM_FAMILY_PREFIX = "rstm:";

/** Marker that distinguishes a paragraph-spacing record (plan A12) from a
 * sizes record. Unambiguous because sizes payloads always begin with a digit. */
const SPACING_MARKER = "p:";

/**
 * One sizes RLE entry: positive integer count, "x", then "i" (inherited) or a
 * positive decimal size. Anchored ^...$ so trailing garbage ("11x", "3x11pt")
 * fails the entry — and with it the whole name (all-or-nothing, see header).
 */
const SIZE_ENTRY_RE = /^(\d+)x(i|\d+(?:\.\d+)?)$/;

/**
 * A spacing payload: "<above>x<below>", each "i" or a NON-negative decimal.
 * Unlike font sizes, 0 is a real, meaningful spacing value (zeroed spacing is
 * exactly what the collapse class writes), so the floor here is >= 0, not > 0.
 */
const SPACING_RE = /^(i|\d+(?:\.\d+)?)x(i|\d+(?:\.\d+)?)$/;

/**
 * Format a point value for a name; null = "i" (inherited). JS's shortest
 * round-trip number formatting already trims trailing zeros (11 -> "11",
 * 8.5 -> "8.5") and parseFloat(String(x)) === x for every finite double, so
 * the codec is lossless for any size the Docs API can report.
 */
function formatPt(pt: number | null): string {
  return pt === null ? "i" : String(pt);
}

/**
 * Canonicalize an RLE list before encoding: drop entries that cover no
 * characters, then coalesce ADJACENT same-size entries. Order matters —
 * dropping happens FIRST so a zero-width entry sitting between two equal-size
 * runs lets its neighbors merge (they really are contiguous in the document).
 * Canonical form keeps encode deterministic and names minimal.
 *
 * Counts are UTF-16 char counts the planner derives from element ranges, so
 * they are positive integers by construction; zero/negative (and, defensively,
 * fractional or non-finite) counts would either cover nothing or encode a name
 * our own decoder rejects — dropping is the safe canonical reading.
 */
function canonicalize(entries: RleEntry[]): RleEntry[] {
  const out: RleEntry[] = [];
  for (const entry of entries) {
    if (!Number.isInteger(entry.count) || entry.count <= 0) continue;
    const prev = out[out.length - 1];
    if (prev !== undefined && prev.sizePt === entry.sizePt) {
      prev.count += entry.count;
    } else {
      // Copy — never mutate the caller's entry objects (out[] is ours to grow).
      out.push({ count: entry.count, sizePt: entry.sizePt });
    }
  }
  return out;
}

/**
 * Encode a region's RLE as one or more range names. Returns [] when nothing
 * survives canonicalization (an empty region anchors nothing — defensive; the
 * planner never asks). Each piece carries the char count it covers so the
 * caller can compute every piece's exact sub-range within the region.
 */
export function encodeSizeEntries(entries: RleEntry[]): { name: string; charCount: number }[] {
  const canonical = canonicalize(entries);
  const pieces: { name: string; charCount: number }[] = [];

  // Greedy packer state for the piece under construction.
  let tokens: string[] = [];
  let nameLen = RSTM_PREFIX.length;
  let charCount = 0;

  /** Close the current piece (no-op while nothing is buffered). */
  const flush = (): void => {
    if (tokens.length === 0) return;
    pieces.push({ name: RSTM_PREFIX + tokens.join(","), charCount });
    tokens = [];
    nameLen = RSTM_PREFIX.length;
    charCount = 0;
  };

  for (const entry of canonical) {
    const token = `${entry.count}x${formatPt(entry.sizePt)}`;
    // Close the piece when this token (plus its joining comma) would push the
    // name past NAME_MAX. A piece always holds >= 1 entry, and a lone token can
    // never overflow by itself: counts are <= Number.MAX_SAFE_INTEGER (16
    // digits) and sizes a handful of chars, so the longest possible
    // single-entry name sits far below NAME_MAX — no intra-entry split needed.
    if (tokens.length > 0 && nameLen + 1 + token.length > NAME_MAX) flush();
    nameLen += tokens.length === 0 ? token.length : 1 + token.length;
    tokens.push(token);
    charCount += entry.count;
  }
  flush();
  return pieces;
}

/**
 * Encode a paragraph-spacing record (plan A12): the DIRECT spaceAbove/Below a
 * fully-hidden paragraph carried before the collapse class zeroed it (null =
 * inherited, restored by clearing). Fixed two-value arity — never splits.
 */
export function encodeSpacingName(spaceAbovePt: number | null, spaceBelowPt: number | null): string {
  return `${RSTM_PREFIX}${SPACING_MARKER}${formatPt(spaceAbovePt)}x${formatPt(spaceBelowPt)}`;
}

/** Decode the payload after "rstm:v1:" as a sizes RLE, or null if malformed. */
function decodeSizes(payload: string): DecodedRangeName | null {
  // "rstm:v1:" with nothing after it carries no manifest — malformed.
  if (payload === "") return null;
  const entries: RleEntry[] = [];
  for (const part of payload.split(",")) {
    const match = SIZE_ENTRY_RE.exec(part);
    if (match === null) return null;
    const count = parseInt(match[1], 10);
    // The regex admits arbitrarily long digit strings; beyond the safe-integer
    // range a count cannot index a real document, so treat it as corruption.
    if (!Number.isSafeInteger(count) || count <= 0) return null;
    const sizePt = match[2] === "i" ? null : parseFloat(match[2]);
    // A 0pt font size cannot be an original size — canonical names never carry
    // one, so its presence means the name is not ours / was mangled.
    if (sizePt !== null && (!Number.isFinite(sizePt) || sizePt <= 0)) return null;
    entries.push({ count, sizePt });
  }
  return { kind: "sizes", entries };
}

/** Decode the payload after "rstm:v1:p:" as a spacing record, or null. */
function decodeSpacing(payload: string): DecodedRangeName | null {
  const match = SPACING_RE.exec(payload);
  if (match === null) return null;
  const spaceAbovePt = match[1] === "i" ? null : parseFloat(match[1]);
  const spaceBelowPt = match[2] === "i" ? null : parseFloat(match[2]);
  // Spacing is >= 0 by the regex (no sign accepted); guard only non-finite
  // parses from pathological digit lengths.
  if (spaceAbovePt !== null && !Number.isFinite(spaceAbovePt)) return null;
  if (spaceBelowPt !== null && !Number.isFinite(spaceBelowPt)) return null;
  return { kind: "spacing", spaceAbovePt, spaceBelowPt };
}

/**
 * Decode a range name into its manifest record. Returns null — NEVER throws —
 * for anything that is not a fully-valid name of OUR version: foreign add-on
 * ranges, future rstm versions, and corrupted names all take the same
 * "not ours to restore" path (see header for why parsing is all-or-nothing).
 */
export function decodeRangeName(name: string): DecodedRangeName | null {
  if (!name.startsWith(RSTM_PREFIX)) return null;
  const payload = name.slice(RSTM_PREFIX.length);
  // Sizes payloads always start with a digit, so the "p:" marker is unambiguous.
  if (payload.startsWith(SPACING_MARKER)) {
    return decodeSpacing(payload.slice(SPACING_MARKER.length));
  }
  return decodeSizes(payload);
}

/**
 * True for ANY rstm-family name, including versions this engine cannot decode.
 * This is the OWNERSHIP test: Show All uses it to warn about (and leave
 * untouched) ranges a future version wrote, instead of sweeping them as
 * foreign tiny text (edge row 16).
 */
export function isRstmName(name: string): boolean {
  return name.startsWith(RSTM_FAMILY_PREFIX);
}

/** True iff the name decodes successfully — i.e. THIS engine can restore from
 * it. isRstmName && !isKnownRstmName is exactly the "ours but newer/corrupt,
 * warn and do not touch" bucket. */
export function isKnownRstmName(name: string): boolean {
  return decodeRangeName(name) !== null;
}
