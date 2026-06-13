// Safety gates and batch chunking for the gdocs engine (plan A3, A5, A7,
// A11.viii, A16 — Word reference: src/core/guards.ts).
//
// Two independent pure concerns live here, mirroring the Word module's split:
//   * Verb gates — clean refusals BEFORE any request is emitted. batchUpdate
//     indexes are only trustworthy against a single-tab, suggestion-free read,
//     so the destructive verbs refuse early rather than risk corruption. Each
//     gate throws exactly one named error class so the adapter can map it to
//     exactly one STRINGS entry (failure-copy matrix).
//   * Chunking — packs the planner's atomic RequestGroups into batchUpdate
//     payloads without ever tearing a group apart. A torn createNamedRange /
//     updateTextStyle pair would let an interrupted Hide strand sentinel-size
//     text with no restore anchor, degrading exactly-restorable to
//     amber-normalized (plan A11.viii / case 001-S4).

import { CHUNK_MAX } from "./constants";
import { isRstmName } from "./rangeNames";
import {
  DocsRequest,
  GDoc,
  HiddenStateError,
  MultiTabError,
  RequestGroup,
  SuggestionsActiveError
} from "./types";

/**
 * Tab gate (plan A3): v1 supports exactly one tab, and the parse only carries
 * the single tab's body segment — operating on a multi-tab doc would write
 * against indexes whose meaning the user cannot see. Refuse OUTRIGHT rather
 * than silently acting on tab 1; the error carries the count for the refusal
 * copy ("This doc uses tabs — run Rostrum on a single-tab doc for now").
 */
export function assertSingleTab(doc: GDoc): void {
  if (doc.tabCount > 1) throw new MultiTabError(doc.tabCount);
}

/**
 * Suggestions hard gate — callers are Hide and Apply-styles ONLY (plan D5).
 * Pending suggestions make every index in the read ambiguous (accepting or
 * rejecting one shifts content), so the verbs that CREATE state refuse before
 * emitting anything.
 *
 * SHOW ALL NEVER CALLS THIS. Word parity (plan A7): the reverse path must
 * always be available — a teammate's pending suggestion must never lock a doc
 * in the hidden state. Show All proceeds under revision chaining instead, and
 * its worst case is fail-VISIBLE, which is the design invariant we protect.
 */
export function assertNoSuggestions(doc: GDoc): void {
  if (doc.suggestionsPresent) throw new SuggestionsActiveError();
}

/**
 * True when the doc carries ANY rstm-owned named range — a "sizes" record, a
 * "spacing" record, or a version this build cannot decode (unknown versions
 * COUNT: they are hidden state, just not ours to silently rewrite — edge row
 * 16). Ownership is rangeNames.isRstmName — the version-AGNOSTIC family test,
 * deliberately broader than what decodeRangeName accepts: decode strictness
 * belongs to restore, never to the gate. Drives the sidebar state line and
 * the Apply-styles gate below.
 */
export function hasRstmState(doc: GDoc): boolean {
  return doc.namedRanges.some((nr) => isRstmName(nr.name));
}

/**
 * Apply-styles gate (plan A5): restyling an armed doc would write run-level
 * sizes (the cite-repair pass, the retro heading pass) that collide with the
 * RLE restore records, and reconcile semantics belong to Hide alone. The user
 * path in the refusal copy is "Show All first, then re-apply styles".
 */
export function assertNotHidden(doc: GDoc): void {
  if (hasRstmState(doc)) throw new HiddenStateError();
}

/**
 * Pack WHOLE RequestGroups into batchUpdate chunks of at most `max` requests
 * (default CHUNK_MAX). The load-bearing invariant (plan A11.viii): a chunk
 * boundary may only fall BETWEEN groups, so a region's createNamedRange always
 * lands in the same batch as — and before — its updateTextStyle.
 *
 * A single group larger than `max` becomes its own OVERSIZED chunk on purpose:
 * CHUNK_MAX is our soft batching cap (a payload/latency convention), not a
 * hard API limit — atomicity beats the cap, because tearing the group is the
 * one thing that can turn an interrupted apply from exactly-restorable into
 * amber-normalized. By the same logic a non-positive `max` merely degrades to
 * one chunk per group; it never drops or splits anything.
 *
 * Group order and intra-group request order pass through verbatim — the
 * planner's descending-start-index emission is a convention, not load-bearing
 * (plan A16; the real invariant is that no request type mutates indexes), but
 * the chunker still must not be the place where order silently changes.
 */
export function chunkGroups(groups: RequestGroup[], max: number = CHUNK_MAX): DocsRequest[][] {
  const chunks: DocsRequest[][] = [];
  // Requests accumulated for the chunk currently being filled.
  let current: DocsRequest[] = [];
  for (const group of groups) {
    // Defensive: an empty group carries nothing — skip it rather than let it
    // flush a chunk early or materialize an empty batch (the API rejects an
    // empty request list).
    if (group.requests.length === 0) continue;
    // Flush only when the group will not fit alongside what is already packed.
    // When `current` is empty the group is admitted unconditionally — that IS
    // the oversized-group case described above.
    if (current.length > 0 && current.length + group.requests.length > max) {
      chunks.push(current);
      current = [];
    }
    current.push(...group.requests);
  }
  if (current.length > 0) chunks.push(current);
  return chunks;
}
