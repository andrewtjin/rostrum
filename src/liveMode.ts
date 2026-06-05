// Live mode (decision #4): while invisibility is ON, keep the paragraph the user
// is typing in VISIBLE so they can see what they write, then let Re-hide reconcile.
//
// VERIFIED-API REALITY (this corrects the plan's wording):
//   * There is NO Word-specific "selection changed" event. The only selection signal
//     is the COMMON API `Office.context.document.addHandlerAsync(
//     Office.EventType.DocumentSelectionChanged, …)`. We use that.
//   * `context.document.getSelection()` returns a `Word.Range` (WordApi 1.1) — NOT a
//     `Word.Selection`. `Range.font` (WordApi 1.1) exposes `.hidden`, so the
//     `selection.font.hidden = false` reveal below operates on the returned Range.
//   * `Font.hidden` is `WordApiDesktop 1.2`, desktop-only — so live mode is a
//     desktop, best-effort nicety. The HARD guarantee is Re-hide, which deterministically
//     re-derives the whole document; live mode only reduces how often you need it.
//
// Everything host-touching is injected (the selection-subscription + the `Word.run`
// runner), so the re-entrancy guard and reveal logic are unit-tested with no host.

import { WordRunner, defaultWordRunner } from "./core/officeWordPort";
import { Logger, logger as rootLogger } from "./core/debug";

/* eslint-disable @typescript-eslint/no-explicit-any */

/** Subscribe to selection-changed; resolves to an unsubscribe function. */
export type SelectionSubscriber = (handler: () => void) => Promise<() => Promise<void>>;

export interface LiveModeOptions {
  /** How to subscribe to selection changes (defaults to the Office Common API). */
  subscribe?: SelectionSubscriber;
  /** `Word.run` seam (defaults to the host runner). */
  runner?: WordRunner;
  logger?: Logger;
}

/**
 * Best-effort "keep the caret's paragraph visible while typing". Start subscribes to
 * selection changes; each change un-hides the current selection's font. A
 * re-entrancy guard drops overlapping events (typing fires them faster than a
 * `Word.run` round-trips), and every failure is swallowed after logging — live mode
 * must never interrupt the user, because Re-hide is the real guarantee.
 */
export class LiveMode {
  private readonly subscribe: SelectionSubscriber;
  private readonly run: WordRunner;
  private readonly log: Logger;

  private unsubscribe: (() => Promise<void>) | null = null;
  /** True while a reveal `Word.run` is in flight — drops re-entrant selection events. */
  private revealing = false;

  constructor(options: LiveModeOptions = {}) {
    this.subscribe = options.subscribe ?? defaultSelectionSubscriber;
    this.run = options.runner ?? defaultWordRunner;
    this.log = options.logger ?? rootLogger("live");
  }

  get isActive(): boolean {
    return this.unsubscribe !== null;
  }

  /** Begin watching selection changes. Idempotent. */
  async start(): Promise<void> {
    if (this.unsubscribe) return;
    this.log.info("live mode starting");
    this.unsubscribe = await this.subscribe(() => {
      void this.revealSelection();
    });
    this.log.debug("live mode subscribed to selection changes");
  }

  /** Stop watching. Idempotent; safe to call when not started. */
  async stop(): Promise<void> {
    if (!this.unsubscribe) return;
    const off = this.unsubscribe;
    this.unsubscribe = null;
    try {
      await off();
      this.log.info("live mode stopped");
    } catch (e) {
      this.log.caught("error while unsubscribing live mode (ignored)", e);
    }
  }

  /**
   * Un-hide the current selection's font. Guarded against re-entrancy and fully
   * fault-tolerant: a failure here is logged and dropped, never surfaced — the user
   * keeps typing and Re-hide will reconcile the document later.
   */
  private async revealSelection(): Promise<void> {
    if (this.revealing) {
      this.log.debug("reveal skipped (already in flight)");
      return;
    }
    this.revealing = true;
    const span = this.log.span("revealSelection");
    try {
      await this.run(async (ctx) => {
        const selection = (ctx.document as any).getSelection();
        selection.font.hidden = false;
        await ctx.sync();
      });
      span.end();
    } catch (e) {
      // Best-effort: desktop-only API, transient selection states, co-authoring — all
      // tolerable. Re-hide is the guarantee, so we never propagate.
      span.fail(e);
    } finally {
      this.revealing = false;
    }
  }
}

/**
 * The production selection subscriber over the Office Common API. Referenced lazily
 * (inside the returned promise) so importing this module never touches the `Office`
 * global in a non-host environment.
 */
const defaultSelectionSubscriber: SelectionSubscriber = (handler) =>
  new Promise((resolve, reject) => {
    if (typeof Office === "undefined") {
      reject(new Error("Rostrum live mode: no Office host detected."));
      return;
    }
    const O: any = Office;
    O.context.document.addHandlerAsync(
      O.EventType.DocumentSelectionChanged,
      handler,
      (res: any) => {
        if (res.status === O.AsyncResultStatus.Succeeded) {
          resolve(
            () =>
              new Promise<void>((res2, rej2) => {
                O.context.document.removeHandlerAsync(
                  O.EventType.DocumentSelectionChanged,
                  { handler },
                  (r: any) => (r.status === O.AsyncResultStatus.Succeeded ? res2() : rej2(r.error))
                );
              })
          );
        } else {
          reject(res.error);
        }
      }
    );
  });
