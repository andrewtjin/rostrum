// Progress phrasing — the ONE place a ProgressInfo tick is turned into human text + a bar
// width. Shared by the in-pane StatusBar (feature panel) and the ribbon pop-out progress
// window, so the two surfaces can never describe the same operation differently (DRY: this
// replaces the formatting that used to be inlined in the invisibility panel).
//
// Pure + host-free (type-only dependency on ProgressInfo), so it unit-tests in Node and adds
// no Office.js / officeWordPort runtime weight to the tiny progress-dialog bundle.
import type { ProgressInfo } from "./officeWordPort";

/**
 * Human label for a progress tick, e.g. "Scanning 612/1041" (read phase) or "Writing 612/1041"
 * (commit phase). With no known total yet (total === 0) it degrades to an indeterminate
 * "Scanning…" / "Writing…", which is also what the pop-out shows when live numbers aren't
 * available (a host without parent→child dialog messaging).
 */
export function formatProgress(p: ProgressInfo): string {
  const verb = p.phase === "read" ? "Scanning" : "Writing";
  return p.total > 0 ? `${verb} ${p.done}/${p.total}` : `${verb}…`;
}

/** Completion percentage (0–100) for a progress-bar width; 0 when the total isn't known yet. */
export function progressPercent(p: ProgressInfo): number {
  if (p.total <= 0) return 0;
  // Clamp to 0–100: a bar width and an aria-valuenow must stay in range even if a caller ever
  // reports done > total (overshoot) — otherwise the bar overflows and the ARIA value is invalid.
  return Math.max(0, Math.min(100, Math.round((p.done / p.total) * 100)));
}
