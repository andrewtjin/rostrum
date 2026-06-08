// Tests for the subpath-safe page-URL builder (src/core/appUrl.ts). The whole reason this helper
// exists is that GitHub *project* Pages serve the bundle from a subpath (`…github.io/rostrum/`),
// where `location.origin` drops the `/rostrum` segment and `${origin}/dialog.html` 404s. These
// assertions pin the relative-resolution behavior at BOTH a subpath base and the dev root base,
// so the regression that motivated the helper can never silently come back.
import { appPageUrl } from "../src/core/appUrl";

/** Point `window.location.href` at `href` for one test (node env has no DOM). */
function withLocation(href: string): void {
  (globalThis as Record<string, unknown>).window = { location: { href } };
}

describe("appPageUrl", () => {
  afterEach(() => {
    delete (globalThis as Record<string, unknown>).window;
  });

  it("resolves a sibling page on a project-Pages SUBPATH (the bug this fixes)", () => {
    withLocation("https://andrewtjin.github.io/rostrum/taskpane.html");
    // Must KEEP /rostrum/ — the origin-only form would have produced github.io/dialog.html (404).
    expect(appPageUrl("dialog.html")).toBe("https://andrewtjin.github.io/rostrum/dialog.html");
    expect(appPageUrl("progress.html")).toBe("https://andrewtjin.github.io/rostrum/progress.html");
  });

  it("resolves a sibling page at the dev ROOT base (no subpath) unchanged", () => {
    withLocation("https://localhost:3000/taskpane.html");
    expect(appPageUrl("dialog.html")).toBe("https://localhost:3000/dialog.html");
  });

  it("resolves correctly even when the CALLER is itself the dialog/progress page", () => {
    // taskpane.html is the ribbon function-file page that opens the pop-out — a sibling, not the root.
    withLocation("https://andrewtjin.github.io/rostrum/taskpane.html#somecmd");
    expect(appPageUrl("progress.html")).toBe("https://andrewtjin.github.io/rostrum/progress.html");
  });

  it("appends a fragment after the page (and only with a single #)", () => {
    withLocation("https://localhost:3000/taskpane.html");
    expect(appPageUrl("dialog.html", "invisibility")).toBe(
      "https://localhost:3000/dialog.html#invisibility"
    );
  });

  it("does NOT double-encode an already-encodeURIComponent'd fragment", () => {
    withLocation("https://localhost:3000/taskpane.html");
    // A feature id with a space → caller passes "a%20b"; the URL hash setter must leave %20 as-is
    // (not re-encode the % to %25), so the page reads back the original via decodeURIComponent.
    const encoded = encodeURIComponent("a b"); // "a%20b"
    expect(appPageUrl("dialog.html", encoded)).toBe("https://localhost:3000/dialog.html#a%20b");
  });

  it("emits no trailing '#' when no fragment is given", () => {
    withLocation("https://localhost:3000/taskpane.html");
    expect(appPageUrl("dialog.html")).not.toContain("#");
  });
});
