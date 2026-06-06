// Ribbon command handlers — generalized over the HEADLESS feature contributions. Instead of
// hand-wiring Hide / Re-hide / Show All / Apply-Styles, we associate EVERY contributed command
// by id (which equals its manifest <FunctionName>). Adding a ribbon command for a new feature is
// then just contributing it — this file never changes.
//
// Importing the headless `contributions` (not the React `registry`) keeps this bundle React-free:
// the ribbon function-file has no DOM/React, so it must not drag the component tree in.
//
// Each handler runs the command in the ribbon runtime, logs the normalized result (a
// function-command can't host UI like the Track-Changes modal, so a `blocked` op is logged for
// the user to finish in the pane), and ALWAYS calls event.completed() so a ribbon button can
// never hang.
import { logger } from "../core/debug";
import { contributions } from "../features/contributions";

/* eslint-disable @typescript-eslint/no-explicit-any */

const log = logger("ribbon");

/** Wire every contributed command to its manifest FunctionName via Office.actions.associate. */
function associateAll(): void {
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

// Associate once Office is ready. Guarded so importing this module never touches `Office`
// in a non-host environment (e.g. a unit test).
if (typeof Office !== "undefined" && typeof Office.onReady === "function") {
  Office.onReady(() => associateAll());
}
