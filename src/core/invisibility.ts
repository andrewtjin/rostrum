// The orchestrator: hide / re-hide / show-all over a WordPort, plus the pure
// per-paragraph classifier they share. This realizes the keeper algorithm from
// plan §1 — whole-document, idempotent, convergent.
//
// `classifyParagraph` is pure (RawParagraph + settings -> new OOXML) and carries
// all the policy; the async functions are thin: gate Track Changes, map the
// classifier over every paragraph, write the changed ones, update the manifest.

import {
  HideOptions,
  HideResult,
  ParagraphAction,
  ParagraphUpdate,
  RawParagraph,
  ResolvedSettings,
  ShowAllResult,
  WordPort
} from "./types";
import {
  computeRunKeepFlags,
  isHeadingKept,
  paragraphHasCiteRun,
  planCrossGapSeparators
} from "./keepers";
import { applyRunVisibility, makeAllVisible, readRuns } from "./ooxml";
import { assertTrackChangesOff } from "./guards";
import { MANIFEST_SCHEMA_VERSION, clearManifestPart, saveManifest } from "./manifest";

// The canonical `ParagraphAction` now lives in types.ts (so `ParagraphUpdate` can
// reference it for the adapter's native-vs-OOXML dispatch). Re-exported here so the
// existing import path (`from "./invisibility"`) keeps working for callers/tests.
export type { ParagraphAction };

export interface ParagraphPlan {
  index: number;
  action: ParagraphAction;
  /** True only when the OOXML actually changed (avoids spurious writes). */
  changed: boolean;
  /** New OOXML (equals the input when `changed` is false). */
  ooxml: string;
}

/**
 * Inline images / embedded objects (`<w:drawing>`, `<w:object>`, `<w:pict>`) reference an
 * INTERNAL package part (e.g. an `r:embed` to `/word/media/…`). A per-paragraph commit
 * fragment can't carry that binary without re-bloating every fragment (which would erase the
 * whole-body read's speed), so a `hidePartial` OOXML write of such a paragraph would dangle
 * the reference and the live host rejects it ("we found a problem with its contents"). We keep
 * those paragraphs whole via the native `font.hidden` toggle (no OOXML write) — decision #16
 * keeps images untouched regardless. A cheap string probe (no extra parse) keeps classify fast.
 */
const HAS_INTERNAL_PART = /<w:(?:drawing|object|pict)[\s>/]/;

/**
 * PURE. Classify one paragraph and return its (possibly unchanged) OOXML.
 *
 * Order of rules mirrors the plan: table content and headings/cites are kept
 * whole (and forced visible, so re-hide un-hides anything wrongly hidden);
 * otherwise hide every run that isn't a highlighted keeper or structural content,
 * and collapse the paragraph mark when the whole paragraph went hidden.
 */
export function classifyParagraph(
  p: RawParagraph,
  settings: ResolvedSettings
): ParagraphPlan {
  // Tables/images/equations are kept untouched (decision #16). Forcing visible is
  // a no-op on a normal table paragraph and self-heals one wrongly hidden before.
  if (p.inTable) {
    const res = makeAllVisible(p.ooxml);
    return { index: p.index, action: "keepWhole", changed: res.changed, ooxml: res.xml };
  }

  // Inline image / embedded object → keep whole (native toggle, never an OOXML write that would
  // dangle the image's internal package part). See HAS_INTERNAL_PART above.
  if (HAS_INTERNAL_PART.test(p.ooxml)) {
    const res = makeAllVisible(p.ooxml);
    return { index: p.index, action: "keepWhole", changed: res.changed, ooxml: res.xml };
  }

  const runs = readRuns(p.ooxml);

  // Heading rule (#7) or cite rule (#6b): keep the whole paragraph visible.
  if (isHeadingKept(p.headingLevel) || paragraphHasCiteRun(runs)) {
    const res = makeAllVisible(p.ooxml);
    return { index: p.index, action: "keepWhole", changed: res.changed, ooxml: res.xml };
  }

  // Body paragraph: keep highlighted (whole-word) + structural runs; hide the rest.
  const keep = computeRunKeepFlags(runs, settings.keepColors);

  // Bridge separators across hidden gaps so isolated highlighted chunks don't fuse
  // ("radiation" + hidden " are a constant threat, " + "would" -> "radiation would",
  // not "radiationwould" — wet-test bug 1). `extraKeep` promotes an existing
  // whitespace-only run; `splits` move one existing space out of a hidden run into a
  // visible sibling. Neither inserts text, so reversibility stays lossless.
  const { extraKeep, splits } = planCrossGapSeparators(runs, keep);
  for (const i of extraKeep) keep[i] = true;

  const hideFlags = keep.map((k) => !k);
  const anyRuns = runs.length > 0;
  const allHidden = anyRuns && hideFlags.every((h) => h);

  // Collapse the paragraph mark only when an entire *content* paragraph was
  // hidden (decision #5). Empty separator paragraphs (no runs) are left alone.
  const res = applyRunVisibility(p.ooxml, hideFlags, allHidden, splits);
  const action: ParagraphAction = allHidden
    ? "hideWhole"
    : hideFlags.some((h) => h)
      ? "hidePartial"
      : "keepWhole";
  return { index: p.index, action, changed: res.changed, ooxml: res.xml };
}

/**
 * Run a mutation under the Track-Changes gate (decision #14). With TC off, runs
 * directly. With TC on: throw unless the caller opted into auto-toggle, in which
 * case turn TC off, run, and restore the prior mode in `finally`.
 */
async function withTrackChangesGate<T>(
  port: WordPort,
  opts: HideOptions,
  body: () => Promise<T>
): Promise<{ result: T; toggled: boolean }> {
  const mode = await port.getChangeTrackingMode();
  if (mode === "Off") {
    return { result: await body(), toggled: false };
  }
  if (!opts.autoToggleTrackChanges) {
    assertTrackChangesOff(mode); // throws TrackChangesActiveError
  }
  await port.setChangeTrackingMode("Off");
  try {
    return { result: await body(), toggled: true };
  } finally {
    await port.setChangeTrackingMode(mode);
  }
}

/** Hide all non-keeper body text and arm the document (write manifest). */
export async function hide(
  port: WordPort,
  settings: ResolvedSettings,
  opts: HideOptions = {}
): Promise<HideResult> {
  const { result, toggled } = await withTrackChangesGate(port, opts, async () => {
    const paras = await port.readParagraphs();
    const updates: ParagraphUpdate[] = [];
    let skipped = 0;
    for (const p of paras) {
      try {
        const plan = classifyParagraph(p, settings);
        // Carry the action so the adapter can apply a native font.hidden toggle for
        // whole-paragraph cases (keepWhole/hideWhole) and reserve OOXML for hidePartial.
        if (plan.changed) updates.push({ index: plan.index, action: plan.action, ooxml: plan.ooxml });
      } catch {
        // A single malformed paragraph must not abort the whole pass: keep it
        // unchanged (visible) and report it so the UI can warn the user.
        skipped++;
      }
    }
    await port.writeParagraphs(updates);
    await saveManifest(port, {
      active: true,
      keepColors: [...settings.keepColors],
      schemaVersion: MANIFEST_SCHEMA_VERSION
    });
    return { scanned: paras.length, changed: updates.length, skipped };
  });
  return {
    paragraphsScanned: result.scanned,
    paragraphsChanged: result.changed,
    paragraphsSkipped: result.skipped,
    trackChangesToggled: toggled
  };
}

/**
 * Re-derive over the whole document, catching newly typed/pasted text. This is
 * exactly `hide` again (deterministic re-derivation — decision #10), exported
 * under its own name so callers and the ribbon read clearly.
 */
export async function reHide(
  port: WordPort,
  settings: ResolvedSettings,
  opts: HideOptions = {}
): Promise<HideResult> {
  return hide(port, settings, opts);
}

/**
 * Reveal everything Rostrum hid and disarm the document. Convergent/idempotent:
 * safe to run from any partial state, so a stray Undo can't strand the doc
 * (decision #14). Not TC-gated — Show All is the always-available reverse.
 */
export async function showAll(port: WordPort): Promise<ShowAllResult> {
  // Fast path (Stage 4): clear font.hidden NATIVELY over the whole body story in one/two
  // host round-trips, instead of reading + rewriting every paragraph's OOXML (the old
  // path did `insertOoxml` per changed paragraph — thousands of reflows on a long brief,
  // ~3 min on the extremely-large doc). Behaviorally identical to the per-paragraph
  // makeAllVisible pass: it reveals the SAME set, including any run the user hid manually
  // (decision #10), and is convergent — a second Show All clears nothing new. Malformed
  // OOXML is moot here: nothing is parsed, so `paragraphsSkipped` is always 0. Not TC-
  // gated; Show All is the always-available reverse.
  const res = await port.clearHidden();
  await clearManifestPart(port);
  return {
    paragraphsScanned: res.paragraphsScanned,
    paragraphsChanged: res.paragraphsChanged,
    paragraphsSkipped: 0
  };
}
