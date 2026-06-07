// Headless opener for the full-window workspace dialog. Used by a feature's ribbon command
// (e.g. "Flow ▸ Open") to launch its space-heavy surface directly — no launcher, no
// pane. Lives apart from the React dialog page (dialog/index.tsx) so the ribbon command bundle
// stays React-free. `Office.context.ui` is available in the ribbon function-file runtime too.
//
// A dialog floats over Word and closes fully (it never docks), so it doesn't shrink the
// document the way a task pane does — the right home for browse/import-heavy tools.
import { appPageUrl } from "../core/appUrl";

/** Open `dialog.html#<featureId>` as an Office dialog, routed straight to that one feature. */
export function openWorkspaceDialog(featureId: string): void {
  if (typeof Office === "undefined" || !Office.context?.ui) return;
  // Resolve dialog.html RELATIVE to the current page so it's correct on a project-Pages subpath
  // (location.origin would drop the `/rostrum` segment → 404). See core/appUrl.ts.
  const url = appPageUrl("dialog.html", encodeURIComponent(featureId));
  Office.context.ui.displayDialogAsync(url, { height: 72, width: 64 }, (res) => {
    if (res.status !== Office.AsyncResultStatus.Succeeded) return;
    const dialog = res.value;
    // Let the dialog ask us to close it (its "Done" button posts "close").
    dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
      const message = (arg as { message?: string }).message;
      if (message === "close") dialog.close();
    });
  });
}
