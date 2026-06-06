// Unit tests for the shared progress phrasing (src/core/progress.ts) — the one helper the pane
// StatusBar and the ribbon pop-out both render from, so they can never describe an op differently.
import { formatProgress, progressPercent } from "../src/core/progress";

describe("formatProgress", () => {
  it("labels the read phase as Scanning with done/total", () => {
    expect(formatProgress({ phase: "read", done: 612, total: 1041 })).toBe("Scanning 612/1041");
  });

  it("labels the commit phase as Writing with done/total", () => {
    expect(formatProgress({ phase: "commit", done: 5, total: 9 })).toBe("Writing 5/9");
  });

  it("degrades to an indeterminate label when the total isn't known yet (0)", () => {
    expect(formatProgress({ phase: "read", done: 0, total: 0 })).toBe("Scanning…");
    expect(formatProgress({ phase: "commit", done: 0, total: 0 })).toBe("Writing…");
  });
});

describe("progressPercent", () => {
  it("computes a rounded percentage", () => {
    expect(progressPercent({ phase: "read", done: 1, total: 4 })).toBe(25);
    expect(progressPercent({ phase: "commit", done: 1, total: 3 })).toBe(33);
  });

  it("is 0 when the total is unknown (no divide-by-zero)", () => {
    expect(progressPercent({ phase: "read", done: 0, total: 0 })).toBe(0);
  });

  it("spans 0 → 100 across the operation", () => {
    expect(progressPercent({ phase: "commit", done: 0, total: 10 })).toBe(0);
    expect(progressPercent({ phase: "commit", done: 10, total: 10 })).toBe(100);
  });

  it("clamps overshoot (done > total) to 100 so a bar / aria-valuenow can't go out of range", () => {
    expect(progressPercent({ phase: "commit", done: 11, total: 10 })).toBe(100);
  });
});
