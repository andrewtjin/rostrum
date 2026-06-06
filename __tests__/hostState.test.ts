// Pure tests for host-readiness resolution — the gate the TASK PANE hard-applies and the DIALOG
// soft-applies. The reported dialog bug (a fully capable desktop host shown "lacks WordApiDesktop
// 1.2" inside the document-detached dialog runtime) is guarded by the soft-mode cases below: with
// the SAME under-reporting requirements probe, HARD mode says "unsupported" but SOFT mode comes up
// "ready". Driving the pure resolver directly keeps this host-free (no Office, no React render).
import { resolveHostState } from "../src/taskpane/host";
import { RequirementsLike } from "../src/core/guards";

/** A requirements probe that reports support only for the listed "<Set> <version>" keys. */
function reqWith(supported: Record<string, boolean>): RequirementsLike {
  return {
    isSetSupported: (name: string, version?: string) =>
      supported[`${name} ${version ?? ""}`.trim()] ?? false,
  };
}

// A genuine current desktop Word: every set Rostrum cares about is present.
const fullDesktop = reqWith({
  "WordApiDesktop 1.2": true,
  "WordApiDesktop 1.1": true,
  "WordApi 1.4": true,
  "WordApi 1.5": true,
});

// The dialog runtime UNDER-reports: the desktop sets read false even though the host has them.
// (WordApi 1.4 still reads true — it's the desktop-floor sets that vanish in the dialog window.)
const dialogUnderReport = reqWith({ "WordApi 1.4": true, "WordApi 1.5": true });

describe("resolveHostState", () => {
  it("hard mode: a fully capable desktop host is ready", () => {
    const s = resolveHostState(fullDesktop, { requireSupport: true });
    expect(s.phase).toBe("ready");
    expect(s.features?.canHide).toBe(true);
    expect(s.unsupportedMessage).toBeNull();
  });

  it("hard mode (the default) flags a host without WordApiDesktop 1.2 as unsupported, with the explanatory message", () => {
    const s = resolveHostState(dialogUnderReport);
    expect(s.phase).toBe("unsupported");
    expect(s.unsupportedMessage).toMatch(/WordApiDesktop 1\.2/);
  });

  it("soft mode: the SAME under-reporting probe comes up READY — this is the dialog bug fix", () => {
    const s = resolveHostState(dialogUnderReport, { requireSupport: false });
    expect(s.phase).toBe("ready");
    expect(s.unsupportedMessage).toBeNull();
  });

  it("soft mode: a missing/throwing requirements probe still comes up ready (no false 'failed to start')", () => {
    const s = resolveHostState(undefined, { requireSupport: false });
    expect(s.phase).toBe("ready");
    expect(s.features?.canHide).toBe(false); // nothing detected, but not a blocker in the dialog
  });

  it("hard mode: a missing requirements probe throws — a genuine startup problem useHost's catch surfaces", () => {
    expect(() => resolveHostState(undefined, { requireSupport: true })).toThrow();
  });
});
