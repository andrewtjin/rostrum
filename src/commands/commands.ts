// Ribbon command handlers — Hide / Re-hide / Show All without opening the pane.
//
// These reuse the SAME tested RostrumController the task pane uses (one source of
// truth for orchestration), so the ribbon and the pane behave identically. The only
// difference is the surface: a function-command can't host the Track-Changes modal,
// so when the gate blocks a Hide, we log a clear instruction (visible in the pane's
// diagnostics) and complete the event — the user finishes the choice in the pane.
// Every path calls `event.completed()` so a button can never hang.

import { assertCanRun, detectFeatureSupport } from "../core/guards";
import { logger } from "../core/debug";
import { OpOutcome, RostrumController } from "../taskpane/controller";

/* eslint-disable @typescript-eslint/no-explicit-any */

const log = logger("ribbon");

// One controller per command runtime, built lazily on first use.
let controllerPromise: Promise<RostrumController> | null = null;

async function getController(): Promise<RostrumController> {
  if (!controllerPromise) {
    controllerPromise = (async () => {
      const features = detectFeatureSupport(Office.context.requirements);
      assertCanRun(features); // throws UnsupportedHostError on web / old perpetual
      const controller = new RostrumController({ features });
      await controller.init();
      return controller;
    })();
  }
  return controllerPromise;
}

/** Shared runner: execute one ribbon action, surface the outcome, always complete. */
async function runRibbon(
  label: string,
  op: (c: RostrumController) => Promise<OpOutcome>,
  event: Office.AddinCommands.Event
): Promise<void> {
  const span = log.span(`ribbon:${label}`);
  try {
    const controller = await getController();
    const out = await op(controller);
    switch (out.status) {
      case "trackChanges":
        log.warn(
          `${label} is blocked because Track Changes is "${out.mode}". Open the Rostrum pane and ` +
            `choose "Turn off & continue", or turn Track Changes off in Review.`,
          { mode: out.mode }
        );
        break;
      case "error":
        log.error(`${label} failed`, { message: out.message });
        break;
      default:
        log.info(`${label} done`, { status: out.status });
    }
    span.end({ status: out.status });
  } catch (e) {
    log.caught(`ribbon ${label} crashed`, e);
    span.fail(e);
  } finally {
    event.completed();
  }
}

function hide(event: Office.AddinCommands.Event): void {
  void runRibbon("hide", (c) => c.hide(), event);
}
function reHide(event: Office.AddinCommands.Event): void {
  void runRibbon("reHide", (c) => c.reHide(), event);
}
function showAll(event: Office.AddinCommands.Event): void {
  void runRibbon("showAll", (c) => c.showAll(), event);
}
function applyStyles(event: Office.AddinCommands.Event): void {
  void runRibbon("applyStyles", (c) => c.applyStyles(), event);
}

// Associate the handlers with the manifest's <FunctionName> ids once Office is ready.
// Guarded so importing this module never touches `Office` off-host.
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => {
    const actions = (Office as any).actions;
    if (actions?.associate) {
      actions.associate("hide", hide as any);
      actions.associate("reHide", reHide as any);
      actions.associate("showAll", showAll as any);
      actions.associate("applyStyles", applyStyles as any);
      log.debug("ribbon handlers associated");
    }
  });
}
