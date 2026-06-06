// The pure Condense engine — Rostrum's answer to Verbatim's `Condense.bas`.
//
// Verbatim's Condense replaces line breaks / tabs / page breaks with spaces, then collapses paragraph
// marks three ways: full-merge (destructive `^p → space`), pilcrow (reversible `^p → ¶`), and retain-
// paragraphs (drop blank lines). Rostrum unifies all three under ONE reversible break-marker mechanism
// (see ooxmlCondense.ts) so EVERY mode is losslessly reversible by default — the part that beats
// Verbatim — with a destructive escape hatch kept only as an opt-in performance lever.
//
// This module is the thin POLICY layer: it normalizes the mode/reversal combination (a destructive
// merge is only legal with pilcrows off — a visible pilcrow IS a marker) and delegates the mechanical
// OOXML surgery to ooxmlCondense. No Office.js, no DOM here — unit-tested in Node, mirroring how
// keepers.ts sits over ooxml.ts.

import { CondenseOptions, CondenseResult, UncondenseResult } from "./types";
import { condenseFragmentOoxml, uncondenseFragmentOoxml } from "./ooxmlCondense";

/**
 * Normalize a requested option set into the one the engine actually runs with. The single invariant:
 * `reversal:"none"` (destructive) is honored ONLY when pilcrows are off, because a visible pilcrow IS a
 * reversible marker — asking for "no marker" while also asking for a visible pilcrow is contradictory,
 * so we keep the lossless marker. Everything else passes through.
 */
export function resolveCondenseOptions(opts: CondenseOptions): CondenseOptions {
  const reversal = opts.reversal === "none" && opts.usePilcrows ? "marker" : opts.reversal;
  return { usePilcrows: opts.usePilcrows, retainParagraphs: opts.retainParagraphs, reversal };
}

/**
 * Condense the active-range fragment per `opts`. Always collapses intra-paragraph whitespace; then
 * merges all paragraphs into one (with reversible boundary markers) or, in retain-paragraphs mode,
 * drops blank/whitespace-only paragraphs. Lossless by default (`reversal:"marker"`).
 */
export function condenseFragment(
  fragmentXml: string,
  opts: CondenseOptions
): CondenseResult & { xml: string } {
  const resolved = resolveCondenseOptions(opts);
  const out = condenseFragmentOoxml(fragmentXml, resolved);
  return {
    xml: out.xml,
    changed: out.changed,
    paragraphsScanned: out.paragraphsScanned,
    boundariesMarked: out.boundariesMarked
  };
}

/**
 * Uncondense the active-range fragment: every reversible boundary marker becomes a paragraph break
 * again (restoring stored paragraph properties), and every retain-mode hidden blank paragraph is
 * un-hidden. The exact inverse of `condenseFragment` for the marker (lossless) modes.
 */
export function uncondenseFragment(fragmentXml: string): UncondenseResult & { xml: string } {
  const out = uncondenseFragmentOoxml(fragmentXml);
  return { xml: out.xml, changed: out.changed, breaksRestored: out.breaksRestored };
}
