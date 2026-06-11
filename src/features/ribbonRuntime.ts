// Shared ribbon-op runtime — the one place a feature's ribbon button turns into a run.
//
// Every ribbon command runs with NO task pane (to keep the document's reading space), so it needs:
//   * a SINGLE in-flight guard ACROSS ALL features — two ribbon clicks (e.g. Hide then Condense) touch
//     the same document, so the second must wait for the first to settle;
//   * a progress pop-out the op streams ticks into, with Cancel wired to the controller;
//   * the OpOutcome → suite-wide CommandResult mapping.
// Invisibility and Condense both used to (or would) hand-roll this; factoring it here removes the
// duplication (plan §5 code-quality) and keeps the cross-feature serialization correct by construction.
//
// React-free (it lives in the ribbon command bundle); `Office`/`window` are touched only inside the
// pop-out driver, so importing this in Node (the manifest generator) stays safe.

import { ProgressInfo } from "../core/officeWordPort";
import { OpOutcome } from "../taskpane/controller";
import { withProgressDialog } from "../progress/host";
import { tracer } from "../core/debug";
import { CommandResult } from "./types";

// A single global progress sink + in-flight flag, SHARED across every feature's runner (they all
// mutate the one open document). `run` points the sink at the active pop-out for the op's duration.
let progressSink: ((p: ProgressInfo) => void) | null = null;
let opInFlight = false;

/** Forward a controller progress tick to the active pop-out (no-op when nothing is running). */
export function emitProgress(p: ProgressInfo): void {
  progressSink?.(p);
}

/** Normalize a controller's OpOutcome into the suite-wide CommandResult. */
function toResult(out: OpOutcome): CommandResult {
  switch (out.status) {
    case "ok":
      return { status: "ok", message: out.message };
    case "trackChanges":
      return {
        status: "blocked",
        message: `Track Changes is "${out.mode}". Open the Rostrum pane and choose "Turn off & continue".`
      };
    case "cancelled":
      return { status: "cancelled" };
    case "error":
      return { status: "error", message: out.message };
  }
}

/** A per-feature ribbon runner: lazily builds + caches its controller and serializes its ops. */
export interface RibbonRunner<C> {
  run(label: string, op: (controller: C) => Promise<OpOutcome>): Promise<CommandResult>;
}

/**
 * Build a ribbon runner for one feature. `build` constructs + initializes the feature's controller
 * (cached after the first successful build; a REJECTED build is never cached, so a transient startup
 * failure retries on the next click). `cancel` (optional) routes the pop-out's Cancel to the controller.
 */
export function createRibbonRunner<C>(opts: {
  feature: string;
  build: () => Promise<C>;
  cancel?: (controller: C) => void;
  /** Bounded ops with nothing to stream (e.g. Condense): suppress the progress pop-out on success;
   *  it still surfaces an error/blocked outcome. See WithProgressOptions.quiet. */
  quiet?: boolean;
}): RibbonRunner<C> {
  let controllerPromise: Promise<C> | null = null;
  const getController = (): Promise<C> => {
    if (!controllerPromise) {
      controllerPromise = opts.build().catch((e) => {
        controllerPromise = null; // never cache a rejected build
        throw e;
      });
    }
    return controllerPromise;
  };

  return {
    async run(label, op) {
      // Refuse overlapping ribbon ops. Set the flag synchronously BEFORE any await so a second
      // near-simultaneous click can't slip past the guard while the first is still awaiting init.
      if (opInFlight) {
        return { status: "blocked", message: "Another Rostrum operation is still running — let it finish first." };
      }
      opInFlight = true;
      try {
        const controller = await getController();
        return await withProgressDialog(label, {
          quiet: opts.quiet,
          onCancel: () => opts.cancel?.(controller),
          // A ribbon op runs in its own runtime, invisible to the pane's diagnostics view — so on a
          // kept-open failure the pop-out is the only place the user can read WHY it failed.
          getDiagnostics: () => tracer.bugReport({ op: label, feature: opts.feature }),
          run: async (forward) => {
            progressSink = forward;
            try {
              return toResult(await op(controller));
            } finally {
              progressSink = null;
            }
          }
        });
      } finally {
        opInFlight = false;
      }
    }
  };
}
