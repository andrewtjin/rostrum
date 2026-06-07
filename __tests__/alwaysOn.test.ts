// Tests for the Always-On reconciliation policy (src/core/alwaysOn.ts) — the host-free brain that
// drives `Office.addin.setStartupBehavior` so Rostrum auto-loads on every document. Exercised against
// a fake host + fake storage so the "first launch self-registers / opt-out sticks / unsupported host
// no-ops / a throwing host degrades gracefully" decisions are provable without an Office host.
import {
  StartupBehavior,
  StartupBehaviorHost,
  desiredBehavior,
  reconcileStartupBehavior,
  readAlwaysOn,
  setAlwaysOn,
} from "../src/core/alwaysOn";
import { ALWAYS_ON_KEY, StorageLike } from "../src/core/settings";

/** In-memory localStorage stand-in. */
class FakeStorage implements StorageLike {
  private map = new Map<string, string>();
  getItem(key: string): string | null {
    return this.map.has(key) ? this.map.get(key)! : null;
  }
  setItem(key: string, value: string): void {
    this.map.set(key, value);
  }
}

/** A fake shared-runtime host whose state + call counts the tests can inspect. */
class FakeHost implements StartupBehaviorHost {
  behavior: StartupBehavior;
  sets = 0;
  gets = 0;
  constructor(
    private readonly supported: boolean,
    initial: StartupBehavior = "none"
  ) {
    this.behavior = initial;
  }
  isSupported(): boolean {
    return this.supported;
  }
  async getStartupBehavior(): Promise<StartupBehavior> {
    this.gets++;
    return this.behavior;
  }
  async setStartupBehavior(behavior: StartupBehavior): Promise<void> {
    this.sets++;
    this.behavior = behavior;
  }
}

/** A supported host whose probe/set always throw — to prove the startup path never propagates. */
const throwingHost: StartupBehaviorHost = {
  isSupported: () => true,
  getStartupBehavior: () => Promise.reject(new Error("addin api blew up")),
  setStartupBehavior: () => Promise.reject(new Error("addin api blew up")),
};

describe("desiredBehavior (persisted intent → Office token)", () => {
  it("defaults to load (always-on is ON by default)", () => {
    expect(desiredBehavior(new FakeStorage())).toBe("load");
  });

  it("reflects an explicit opt-out", () => {
    const s = new FakeStorage();
    s.setItem(ALWAYS_ON_KEY, "false");
    expect(desiredBehavior(s)).toBe("none");
  });
});

describe("reconcileStartupBehavior (first-launch self-register)", () => {
  it("registers load on first launch when Office is still none", async () => {
    const host = new FakeHost(true, "none");
    const state = await reconcileStartupBehavior(host, new FakeStorage());
    expect(state).toEqual({ supported: true, on: true });
    expect(host.behavior).toBe("load");
    expect(host.sets).toBe(1);
  });

  it("is idempotent — does NOT re-write when Office already matches the intent", async () => {
    const host = new FakeHost(true, "load");
    const state = await reconcileStartupBehavior(host, new FakeStorage());
    expect(state).toEqual({ supported: true, on: true });
    expect(host.sets).toBe(0); // already load → no write
  });

  it("drives Office to none when the user has opted out", async () => {
    const s = new FakeStorage();
    s.setItem(ALWAYS_ON_KEY, "false");
    const host = new FakeHost(true, "load");
    const state = await reconcileStartupBehavior(host, s);
    expect(state).toEqual({ supported: true, on: false });
    expect(host.behavior).toBe("none");
    expect(host.sets).toBe(1);
  });

  it("no-ops on an unsupported host, reporting the persisted intent", async () => {
    const host = new FakeHost(false, "none");
    const state = await reconcileStartupBehavior(host, new FakeStorage());
    expect(state).toEqual({ supported: false, on: true }); // default-on intent
    expect(host.gets).toBe(0);
    expect(host.sets).toBe(0);
  });

  it("degrades gracefully (never throws) when a supported host's api throws", async () => {
    const state = await reconcileStartupBehavior(throwingHost, new FakeStorage());
    expect(state).toEqual({ supported: true, on: true });
  });
});

describe("readAlwaysOn (toggle seed — Office truth when supported)", () => {
  it("reads Office's real behavior on a supported host", async () => {
    const host = new FakeHost(true, "load");
    expect(await readAlwaysOn(host, new FakeStorage())).toEqual({ supported: true, on: true });
    host.behavior = "none";
    expect(await readAlwaysOn(host, new FakeStorage())).toEqual({ supported: true, on: false });
  });

  it("falls back to the persisted intent on an unsupported host", async () => {
    const s = new FakeStorage();
    s.setItem(ALWAYS_ON_KEY, "false");
    expect(await readAlwaysOn(new FakeHost(false), s)).toEqual({ supported: false, on: false });
  });

  it("falls back to default-on intent when nothing is stored on an unsupported host", async () => {
    expect(await readAlwaysOn(new FakeHost(false), new FakeStorage())).toEqual({ supported: false, on: true });
  });

  it("falls back to the persisted intent when a supported host's getStartupBehavior throws", async () => {
    const s = new FakeStorage();
    s.setItem(ALWAYS_ON_KEY, "false");
    // supported, but the probe blows up → report the persisted intent rather than throwing.
    expect(await readAlwaysOn(throwingHost, s)).toEqual({ supported: true, on: false });
  });
});

describe("setAlwaysOn (toggle change — persist + drive Office)", () => {
  it("persists the intent AND drives Office when supported", async () => {
    const s = new FakeStorage();
    const host = new FakeHost(true, "load");
    const state = await setAlwaysOn(host, s, false);
    expect(state).toEqual({ supported: true, on: false });
    expect(s.getItem(ALWAYS_ON_KEY)).toBe("false");
    expect(host.behavior).toBe("none");
  });

  it("persists the intent even on an unsupported host (Office reconciles later)", async () => {
    const s = new FakeStorage();
    const state = await setAlwaysOn(new FakeHost(false), s, false);
    expect(state).toEqual({ supported: false, on: false });
    expect(s.getItem(ALWAYS_ON_KEY)).toBe("false");
  });

  it("keeps the persisted intent even when the Office set throws", async () => {
    const s = new FakeStorage();
    const state = await setAlwaysOn(throwingHost, s, false);
    expect(state).toEqual({ supported: true, on: false });
    expect(s.getItem(ALWAYS_ON_KEY)).toBe("false"); // intent survived the failed Office call
  });
});
