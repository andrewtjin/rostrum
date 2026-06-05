// Host-glue tests with no Word host: the PURE style plan, and LiveMode's
// re-entrancy guard + reveal logic driven through injected fakes.

import { planStyleApplications } from "../src/core/officeStyles";
import { LiveMode, SelectionSubscriber } from "../src/liveMode";
import { Tracer } from "../src/core/debug";

const flush = (): Promise<void> => new Promise((r) => setImmediate(r));

// ---------------------------------------------------------------------------
describe("planStyleApplications (pure)", () => {
  it("maps STYLE_MAP to concrete edits and boxes the pocket via Style.borders", () => {
    const plan = planStyleApplications({ canStyleBorders: true });
    const byKey = Object.fromEntries(plan.map((a) => [a.key, a]));

    expect(byKey.pocket).toMatchObject({ styleName: "Heading 1", sizePt: 26, border: "style" });
    expect(byKey.tag).toMatchObject({ styleName: "Heading 4", sizePt: 14, border: "none" });
    expect(byKey.cite).toMatchObject({
      styleName: "Style13ptBold",
      isCharacterStyle: true,
      sizePt: 14,
      border: "none"
    });
  });

  it("chooses the OOXML border fallback when Style.borders is absent", () => {
    const plan = planStyleApplications({ canStyleBorders: false });
    expect(plan.find((a) => a.key === "pocket")?.border).toBe("ooxml");
  });
});

// ---------------------------------------------------------------------------
describe("LiveMode", () => {
  /** A fake selection + context whose sync() is controllable for the re-entrancy test. */
  function fakeRunner() {
    const selection = { font: { hidden: true } };
    let calls = 0;
    let release: (() => void) | null = null;
    const ctx = {
      document: { getSelection: () => selection },
      sync: () =>
        new Promise<void>((resolve) => {
          release = resolve;
        })
    };
    const runner = async <T,>(batch: (c: any) => Promise<T>): Promise<T> => {
      calls++;
      return batch(ctx);
    };
    return { runner, selection, calls: () => calls, releaseSync: () => release?.() };
  }

  it("subscribes on start and reveals the selection on a change", async () => {
    let handler: (() => void) | null = null;
    const subscribe: SelectionSubscriber = (h) => {
      handler = h;
      return Promise.resolve(async () => undefined);
    };
    const f = fakeRunner();
    const live = new LiveMode({
      subscribe,
      runner: f.runner as never,
      logger: new Tracer({ console: null }).logger("live")
    });

    await live.start();
    expect(live.isActive).toBe(true);

    handler!(); // a selection change
    f.releaseSync();
    await flush();
    expect(f.selection.font.hidden).toBe(false); // un-hidden
    expect(f.calls()).toBe(1);
  });

  it("drops re-entrant selection events while a reveal is in flight", async () => {
    let handler: (() => void) | null = null;
    const subscribe: SelectionSubscriber = (h) => {
      handler = h;
      return Promise.resolve(async () => undefined);
    };
    const f = fakeRunner();
    const live = new LiveMode({ subscribe, runner: f.runner as never });
    await live.start();

    handler!(); // first change -> reveal starts, sync pending
    await flush();
    handler!(); // second change while in flight -> must be dropped
    await flush();
    expect(f.calls()).toBe(1);

    f.releaseSync(); // first reveal completes
    await flush();
    handler!(); // now a new change runs again
    f.releaseSync();
    await flush();
    expect(f.calls()).toBe(2);
  });

  it("swallows reveal failures (Re-hide is the guarantee) and stays active", async () => {
    let handler: (() => void) | null = null;
    const subscribe: SelectionSubscriber = (h) => {
      handler = h;
      return Promise.resolve(async () => undefined);
    };
    const runner = async (): Promise<void> => {
      throw { code: "GeneralException", message: "selection gone" };
    };
    const live = new LiveMode({ subscribe, runner: runner as never });
    await live.start();
    handler!();
    await expect(flush()).resolves.toBeUndefined(); // no unhandled rejection
    expect(live.isActive).toBe(true);
  });

  it("unsubscribes on stop", async () => {
    let unsubscribed = false;
    const subscribe: SelectionSubscriber = () =>
      Promise.resolve(async () => {
        unsubscribed = true;
      });
    const live = new LiveMode({ subscribe, runner: (async () => undefined) as never });
    await live.start();
    await live.stop();
    expect(unsubscribed).toBe(true);
    expect(live.isActive).toBe(false);
  });
});
