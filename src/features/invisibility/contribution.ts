// Invisibility Mode — the HEADLESS contribution (no React). This is what the ribbon runtime
// associates and what the manifest generator turns into a ribbon group; the React pane lives in
// `panel.tsx` and is attached in `feature.ts`. Keeping this React-free is why the ribbon command
// bundle (commands.js) and the Node manifest generator never pull in the component tree.
import { assertCanRun, detectFeatureSupport } from "../../core/guards";
import { ProgressInfo } from "../../core/officeWordPort";
import { OpOutcome, RostrumController } from "../../taskpane/controller";
import { withProgressDialog } from "../../progress/host";
import { tracer } from "../../core/debug";
import { CommandResult, FeatureContribution, RibbonGroup } from "../types";

// ===========================================================================
// Ribbon-runtime controller (separate JS context from the task pane).
// ===========================================================================

// Mutable progress sink. The ribbon controller's `onProgress` is fixed at construction, so we
// indirect through this module-level hook rather than rebuilding the controller per click:
// `runOp` points it at the active pop-out for the operation's duration, then clears it.
let progressSink: ((p: ProgressInfo) => void) | null = null;
// Serializes ribbon ops. The controller is a shared singleton and `progressSink` is a single
// global, so two overlapping ribbon clicks (e.g. Hide then Apply Styles) would mutate one
// document concurrently and cross-wire progress. We refuse the second op until the first settles.
let opInFlight = false;

let controllerPromise: Promise<RostrumController> | null = null;

async function getRibbonController(): Promise<RostrumController> {
  if (!controllerPromise) {
    controllerPromise = (async () => {
      const features = detectFeatureSupport(Office.context.requirements);
      assertCanRun(features); // throws UnsupportedHostError on web / old perpetual
      const controller = new RostrumController({
        features,
        onProgress: (p) => progressSink?.(p),
      });
      await controller.init();
      return controller;
    })().catch((e) => {
      // Never cache a REJECTED promise: otherwise a transient startup failure would make every
      // later ribbon click fail silently for the session. Reset so the next click retries.
      controllerPromise = null;
      throw e;
    });
  }
  return controllerPromise;
}

/** Normalize the controller's OpOutcome into the suite-wide CommandResult. */
function toResult(out: OpOutcome): CommandResult {
  switch (out.status) {
    case "ok":
      return { status: "ok", message: out.message };
    case "trackChanges":
      return {
        status: "blocked",
        message: `Track Changes is "${out.mode}". Open the Rostrum pane and choose "Turn off & continue".`,
      };
    case "cancelled":
      return { status: "cancelled" };
    case "error":
      return { status: "error", message: out.message };
  }
}

/**
 * Run an invisibility op FROM THE RIBBON: pop a small progress window, forward the controller's
 * progress into it for the op's duration, run the op, and return the normalized result. The
 * dialog auto-closes; its Cancel routes to `controller.cancel()`. Best-effort progress — if the
 * dialog can't open, the op still runs (withProgressDialog handles that).
 */
async function runOp(label: string, op: (c: RostrumController) => Promise<OpOutcome>): Promise<CommandResult> {
  // Refuse overlapping ribbon ops. Set the flag synchronously BEFORE any await so a second
  // near-simultaneous click can't slip past the guard while the first is still awaiting init.
  if (opInFlight) {
    return { status: "blocked", message: "Another Rostrum operation is still running — let it finish first." };
  }
  opInFlight = true;
  try {
    const controller = await getRibbonController();
    return await withProgressDialog(label, {
      onCancel: () => controller.cancel(),
      // A ribbon op runs in its own runtime, so its tracer log is invisible to the task pane's
      // diagnostics view. On a kept-open failure the pop-out surfaces this bug report — the ONLY
      // place the user can read WHY a ribbon Hide/Show/Apply failed (the Office error code +
      // debugInfo.errorLocation/statement that pinpoint an insertOoxml rejection).
      getDiagnostics: () => tracer.bugReport({ op: label, feature: "invisibility" }),
      run: async (forward) => {
        progressSink = forward;
        try {
          return toResult(await op(controller));
        } finally {
          progressSink = null;
        }
      },
    });
  } finally {
    opInFlight = false;
  }
}

// ===========================================================================
// The ribbon group + commands. Command ids == the manifest <FunctionName>s the generator emits.
// ===========================================================================

const ribbon: RibbonGroup = {
  label: "Invisibility",
  controls: [
    // Hide is idempotent + convergent, so it doubles as Re-hide: pressing it again re-derives over
    // the whole document and catches newly typed/pasted text. We surface ONE button (no separate
    // Re-hide control) and say so in the tip — the engine `reHide()` is just `hide()` either way.
    { kind: "action", commandId: "hide", label: "Hide", tip: "Hide all non-keeper body text (cards), keeping headings, cites, analytics, and highlighted runs. Run it again after editing to hide newly typed or pasted text." },
    { kind: "action", commandId: "showAll", label: "Show All", tip: "Reveal everything Rostrum hid. Convergent — safe to run from any state." },
    { kind: "action", commandId: "applyStyles", label: "Apply Styles", tip: "Apply Rostrum heading/cite sizes and the pocket box, and repair mis-styled cites. Reflows the document; needs desktop Word 1.5+." },
    { kind: "pane", label: "Settings", tip: "Open the Rostrum pane for keep-color settings, whole-body mode, and diagnostics." },
  ],
};

export const invisibilityContribution: FeatureContribution = {
  id: "invisibility",
  title: "Invisibility Mode",
  tagline: "Collapse a brief to headings, cites, and highlights — instantly, reversibly.",
  glyph: "🙈",
  status: "stable",
  primarySurface: "pane",
  isAvailable: (f) => f.canHide,
  unavailableReason: "Needs desktop Word (hidden-text formatting is WordApiDesktop 1.2).",
  ribbon,
  commands: [
    {
      id: "hide", // == manifest <FunctionName>
      title: "Hide",
      description: "Hide card bodies; keep headings, cites, analytics, and highlights.",
      isAvailable: (f) => f.canHide,
      run: () => runOp("Hide", (c) => c.hide()),
    },
    {
      id: "showAll",
      title: "Show All",
      description: "Reveal everything Rostrum hid and disarm the document.",
      isAvailable: (f) => f.canHide,
      run: () => runOp("Show All", (c) => c.showAll()),
    },
    {
      id: "applyStyles",
      title: "Apply Styles",
      description: "Apply Rostrum heading/cite sizes + pocket box, and repair mis-styled cites.",
      isAvailable: (f) => f.canGetStyles && f.canStyleFormat,
      run: () => runOp("Apply Styles", (c) => c.applyStyles()),
    },
  ],
};
