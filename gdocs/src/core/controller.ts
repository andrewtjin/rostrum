// The verb orchestrator — the ONE place the pure planners meet the DocsPort
// (plan D5/A5/A7/A13/A14, step S9; Word reference: src/core/invisibility.ts).
//
// Every verb runs the SAME apply protocol (the cross-module pin):
//
//   fetch -> parseDocument -> assertSingleTab -> verb gates -> resolveSettings
//   -> plan -> chunkGroups -> apply chunks chained on the revisionId RETURNED
//   BY EACH applyBatch.
//
// ONE documents.get per verb (plan A13): revision chaining reads writeControl
// from each batchUpdate RESPONSE, never from a fresh get — the only re-fetch
// is the silent re-plan after a first-chunk revision mismatch, where a fresh
// read is the whole point.
//
// Failure discipline (plan D5):
//   * RevisionMismatch BEFORE anything landed (chunk 0 — more precisely,
//     applied === 0, see runVerb) -> silently re-plan from a fresh fetch, up
//     to MAX_REPLAN_ATTEMPTS times, then RevisionConflictError(verb). Retries
//     are IMMEDIATE — no sleeps in core, so the engine stays deterministic
//     and flake-free under test (plan A11.vi).
//   * RevisionMismatch AFTER >= 1 chunk landed -> PartialApplyError(verb,
//     applied, total). Re-planning here would be a lie — the doc already
//     changed under a half-applied verb; the user copy owns the recovery
//     route (Show All unwinds a torn Hide; Show All/styles re-run to finish).
//   * A DocsApiError aborts the verb UNLESS it hit applyStyles' named-style
//     batch — the ONE degradable batch (plan A5/D13): consumer accounts may
//     400 updateNamedStyle, so that batch alone is allowed to fail, flagged
//     namedStylesApplied=false, and the retro pass still runs. The degraded
//     path lives HERE, in core, precisely so it is unit-testable without a
//     host.
//
// Gate placement is deliberate and verb-specific (the A7 asymmetry):
//   * assertSingleTab guards EVERY verb — the parse only carries one tab's
//     body, so no verb may act on a multi-tab doc (plan A3, case 001-F9);
//   * Hide and Apply-styles gate on suggestions (indexes are only
//     trustworthy suggestion-free — they CREATE state);
//   * Apply-styles additionally refuses armed docs (assertNotHidden, plan A5);
//   * Show All gates on NOTHING beyond the tab check: it is the
//     always-available reverse (Word parity), and its one consent boundary —
//     the PURE-sweep path — is a question, not a gate (plan A14).
//
// Everything below is still host-free: the DocsPort seam is the only I/O, so
// the flagship suite drives these verbs against an in-memory fake
// (__tests__/fakeDocs.ts) end to end.

import { MAX_REPLAN_ATTEMPTS } from "./constants";
import { assertNoSuggestions, assertNotHidden, assertSingleTab, chunkGroups } from "./guards";
import { parseDocument } from "./parse";
import { planHide } from "./planner";
import { planShowAll, pureSweepConsentNeeded } from "./restore";
import { resolveSettings } from "./settings";
import { planApplyStyles } from "./styles";
import {
  DocsApiError,
  DocsPort,
  DocsRequest,
  GDoc,
  GdocsSettings,
  GHideResult,
  GShowAllOutcome,
  GStylesResult,
  PartialApplyError,
  RequestGroup,
  RevisionConflictError,
  RevisionMismatchError
} from "./types";

// ---------------------------------------------------------------------------
// Private protocol machinery
// ---------------------------------------------------------------------------

/** The verb names the per-verb error classes carry (types.ts contract). */
type Verb = "hide" | "showAll" | "applyStyles";

/** One batchUpdate payload, already chunked. `degradable` marks applyStyles'
 * named-style batch ONLY: a DocsApiError there flips the degraded flag and
 * the verb continues; on any other batch it propagates (the batch was atomic,
 * so the generic refusal copy stays truthful for it). */
interface PlannedBatch {
  requests: DocsRequest[];
  degradable: boolean;
}

/** What one plan attempt hands the apply loop. `finish` builds the verb's
 * result AFTER the apply outcome is known — applyStyles' namedStylesApplied
 * is an apply fact, unknowable at plan time (styles.ts module header). A
 * zero-batch plan is the legitimate no-op/short-circuit shape: the loop
 * applies nothing and `finish(false)` is returned untouched (Hide's "nothing
 * to hide", Show All's clean doc AND its consent question all ride this). */
interface VerbPlan<R> {
  batches: PlannedBatch[];
  finish: (degraded: boolean) => R;
}

/**
 * The shared protocol runner (see module header). `gate` carries the
 * verb-specific gates beyond the universal tab check; `makePlan` is invoked
 * once per attempt with the FRESH parsed view — a re-plan must never reuse
 * stale indexes, that is its entire reason to exist.
 *
 * Settings are read ONCE, outside the attempt loop: they live in document
 * properties, not document content, so they are not revision-coupled and a
 * re-plan gains nothing from re-reading them.
 *
 * The "first chunk" retry criterion is `applied === 0`, not a literal chunk
 * index: applyStyles' named-style batch may be REJECTED (degraded path)
 * without anything landing, and a mismatch on the next batch is then still a
 * clean nothing-applied state — silently re-planning it is truthful, while
 * keying on index 0 would wrongly report a partial apply.
 */
async function runVerb<R>(
  port: DocsPort,
  verb: Verb,
  gate: (doc: GDoc) => void,
  makePlan: (doc: GDoc, settings: GdocsSettings) => VerbPlan<R>
): Promise<R> {
  // Device defaults are an adapter concern (it embeds UserProperties JSON);
  // core passes null and the doc-properties tier + built-ins resolve (plan S4).
  const settings = resolveSettings(await port.readSettingsJson(), null);

  // Initial attempt + up to MAX_REPLAN_ATTEMPTS silent re-plans (plan D5).
  for (let attempt = 0; ; attempt++) {
    const doc = parseDocument(await port.fetchDocument());
    assertSingleTab(doc);
    gate(doc);
    const plan = makePlan(doc, settings);

    // Revision chain: starts at the fetched revision, then advances to each
    // applyBatch RESPONSE's revision (plan A13 — one get per verb).
    let revisionId = doc.revisionId;
    let degraded = false;
    let applied = 0;
    try {
      for (const batch of plan.batches) {
        try {
          const res = await port.applyBatch(batch.requests, revisionId);
          revisionId = res.revisionId;
          applied++;
        } catch (e) {
          if (batch.degradable && e instanceof DocsApiError) {
            // The named-style 400 (plan A5/D13): the rejected batch was
            // atomic, so the revision did not move — the chain continues
            // from the same id and the retro batches still land.
            degraded = true;
            continue;
          }
          throw e;
        }
      }
    } catch (e) {
      // Only revision mismatches are the controller's to handle; every other
      // failure (DocsApiError on a non-degradable batch, adapter surprises)
      // propagates for strings.errorMessage to map.
      if (!(e instanceof RevisionMismatchError)) throw e;
      if (applied > 0) throw new PartialApplyError(verb, applied, plan.batches.length);
      if (attempt >= MAX_REPLAN_ATTEMPTS) throw new RevisionConflictError(verb);
      continue; // nothing landed — silent, immediate re-plan from a fresh fetch
    }
    return plan.finish(degraded);
  }
}

/** Chunk RequestGroups into non-degradable batches — the shape every verb
 * except applyStyles' named-style batch uses (DRY over the three verbs). */
function regularBatches(groups: RequestGroup[]): PlannedBatch[] {
  return chunkGroups(groups).map((requests) => ({ requests, degradable: false }));
}

// ---------------------------------------------------------------------------
// Public verbs
// ---------------------------------------------------------------------------

/**
 * Hide: the A1 reconcile over the whole doc. Gated on suggestions (plan D5)
 * because Hide writes against absolute indexes; refusal copy owns the escape
 * route. The result is the SUCCESSFUL attempt's planner result — exactly the
 * plan that applied, so the receipt's counts can never describe a stale view.
 */
export async function hide(port: DocsPort): Promise<GHideResult> {
  return runVerb(port, "hide", assertNoSuggestions, (doc, settings) => {
    const { groups, result } = planHide(doc, settings);
    return { batches: regularBatches(groups), finish: () => result };
  });
}

/**
 * Show All: restore + convergence sweep, deliberately ungated beyond the tab
 * check (plan A7 — a teammate's pending suggestion must never lock a doc in
 * the hidden state).
 *
 * The consent handshake (plan A14): when the doc carries ZERO rstm ranges but
 * shows tiny text, the tiny text might be the user's own formatting — so when
 * the caller has not answered (`opts.sweepUnrecorded` undefined) the verb
 * returns `needsConsent` WITHOUT writing anything, and the adapter re-invokes
 * with the answer. An explicit `false` proceeds and leaves the unrecorded
 * text untouched (planShowAll(doc, false) on a pure-sweep doc is a clean
 * zero-group no-op by design); `true` adopts it into the sweep. On an ARMED
 * doc the flag is irrelevant — orphans next to our own state are ours.
 */
export async function showAll(
  port: DocsPort,
  opts?: { sweepUnrecorded?: boolean }
): Promise<GShowAllOutcome> {
  return runVerb<GShowAllOutcome>(
    port,
    "showAll",
    () => undefined, // no verb gates: the always-available reverse (plan A7)
    (doc) => {
      const consent = pureSweepConsentNeeded(doc);
      if (consent !== null && opts?.sweepUnrecorded === undefined) {
        // Ask first — zero batches means zero writes on this path.
        return {
          batches: [],
          finish: (): GShowAllOutcome => ({ kind: "needsConsent", unrecordedTinyCount: consent })
        };
      }
      const plan = planShowAll(doc, opts?.sweepUnrecorded ?? false);
      // Edge row 16 rides INSIDE plan.result (rangesSkippedNewerVersion), so
      // passing the result through is the whole pass-through: the receipt's
      // amber newer-version line renders from the same counted contract as
      // every other Show All outcome.
      return {
        batches: regularBatches(plan.groups),
        finish: (): GShowAllOutcome => ({ kind: "done", result: plan.result })
      };
    }
  );
}

/**
 * Apply debate styles: the two-batch styles pass (plan A5). Gates: suggestions
 * (indexes) AND hidden state (restyling an armed doc would write run-level
 * sizes that collide with the RLE restore records — refusal copy says "Show
 * All first"). The named-style batch goes FIRST and is the one degradable
 * batch; the retro batches then run regardless of its fate, so existing
 * paragraphs are always styled and only future typing depends on the
 * named-style write landing (the documented degraded path, plan D13).
 */
export async function applyStyles(port: DocsPort): Promise<GStylesResult> {
  return runVerb(
    port,
    "applyStyles",
    (doc) => {
      assertNoSuggestions(doc);
      assertNotHidden(doc);
    },
    (doc, settings) => {
      const plan = planApplyStyles(doc, settings);
      const named = chunkGroups(plan.namedStyleGroups).map(
        (requests): PlannedBatch => ({ requests, degradable: true })
      );
      return {
        batches: [...named, ...regularBatches(plan.retroGroups)],
        finish: (degraded): GStylesResult => ({
          namedStylesApplied: !degraded,
          restyled: plan.counts.restyled,
          spacingCleared: plan.counts.spacingCleared,
          citesRepaired: plan.counts.citesRepaired
        })
      };
    }
  );
}
