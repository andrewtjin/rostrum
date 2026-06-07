// Ribbon command handlers — generalized over the HEADLESS feature contributions. Instead of
// hand-wiring Hide / Re-hide / Show All / Apply-Styles, we associate EVERY contributed command
// by id (which equals its manifest <FunctionName>). Adding a ribbon command for a new feature is
// then just contributing it — this file never changes.
//
// Importing the headless `contributions` (not the React `registry`) keeps the association logic
// itself React-free; it is invoked from the SHARED-RUNTIME startup (src/taskpane/index.tsx), which
// is the manifest's <FunctionFile> + long-lived <Runtime> page now that commands.html is gone.
// (Pre-0.3.0 this module was its own ephemeral commands.js entry that self-ran on Office.onReady;
// under the shared runtime the one runtime page wires the ribbon AND hosts the on-demand pane.)
//
// Each handler runs the command, logs the normalized result (a function-command can't host UI like
// the Track-Changes modal, so a `blocked` op is logged for the user to finish in the pane), and
// ALWAYS calls event.completed() so a ribbon button can never hang.
import { logger } from "../core/debug";
import { contributions } from "../features/contributions";

/* eslint-disable @typescript-eslint/no-explicit-any */

const log = logger("ribbon");

/**
 * Wire every contributed command to its manifest FunctionName via Office.actions.associate. Called
 * once from the shared-runtime startup; guarded so a host without `Office.actions` (e.g. a unit test)
 * is a safe no-op.
 */
export function associateAll(): void {
  const actions = (Office as any).actions;
  if (!actions?.associate) {
    log.warn("Office.actions.associate unavailable — ribbon commands not wired");
    return;
  }
  const commands = contributions.flatMap((feature) => feature.commands);
  for (const command of commands) {
    actions.associate(command.id, (event: Office.AddinCommands.Event) => {
      const span = log.span(`ribbon:${command.id}`);
      command
        .run()
        .then((result) => {
          if (result.status === "error") log.error(`${command.id} failed`, { message: result.message });
          else if (result.status === "blocked") log.warn(`${command.id} blocked`, { message: result.message });
          else log.info(`${command.id} done`, { status: result.status });
          span.end({ status: result.status });
        })
        .catch((e) => {
          log.caught(`ribbon ${command.id} crashed`, e);
          span.fail(e);
        })
        .finally(() => event.completed());
    });
  }
  log.debug("ribbon handlers associated", { count: commands.length });
}
