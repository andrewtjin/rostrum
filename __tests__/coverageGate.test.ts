// Guards the WIRING of the core-engine coverage gate, not the coverage itself.
//
// jest.config.js has always declared coverageThreshold (80/90/90/90 over src/core),
// but Jest only evaluates thresholds when coverage collection is enabled — so until
// the deploy workflow switched from plain `npm test` to `npm run test:coverage`, the
// gate was dead config that never executed anywhere. These tests pin all three links
// of the chain (jest config → npm script → CI step) so no single edit can silently
// disconnect the gate again:
//   • dropping/weakening coverageThreshold fails here even though the workflow stays green;
//   • deleting the test:coverage script (or stripping --coverage) fails here;
//   • reverting deploy-pages.yml to plain `npm test` fails here.
// File-content assertions (vs. actually spawning jest) keep this suite fast; the
// thresholds themselves are exercised for real on every deploy.
import * as fs from "fs";
import * as path from "path";

const repoRoot = path.resolve(__dirname, "..");

// jest.config.js is plain CommonJS, so require() gives the live object Jest itself
// loads — no re-parsing, and a syntax error in the config fails loudly here too.
// eslint-disable-next-line @typescript-eslint/no-var-requires
const jestConfig = require(path.join(repoRoot, "jest.config.js")) as {
  collectCoverageFrom?: string[];
  coverageThreshold?: { global?: Record<string, number> };
};

describe("core-engine coverage gate wiring", () => {
  test("jest.config.js declares the advertised thresholds over src/core", () => {
    // The gate's value is the specific floors the repo advertises; a silent lowering
    // (e.g. branches 80 → 10 to sneak a regression through) must fail the suite.
    const globals = jestConfig.coverageThreshold?.global;
    expect(globals).toBeDefined();
    expect(globals!.branches).toBeGreaterThanOrEqual(80);
    expect(globals!.functions).toBeGreaterThanOrEqual(90);
    expect(globals!.lines).toBeGreaterThanOrEqual(90);
    expect(globals!.statements).toBeGreaterThanOrEqual(90);

    // Scope must stay the pure engine: src/core measured, the Word-host-only
    // adapter excluded (it cannot run headless, so including it would force the
    // thresholds down and dilute the gate).
    expect(jestConfig.collectCoverageFrom).toContain("src/core/**/*.ts");
    expect(jestConfig.collectCoverageFrom).toContain("!src/core/officeWordPort.ts");
  });

  test("package.json exposes test:coverage so the gate is one explicit script", () => {
    // Everyday `npm test` stays coverage-free (fast inner loop); the gate lives in a
    // named script so CI and humans invoke the exact same command.
    const pkg = JSON.parse(
      fs.readFileSync(path.join(repoRoot, "package.json"), "utf8")
    ) as { scripts?: Record<string, string> };
    expect(pkg.scripts?.["test:coverage"]).toBeDefined();
    expect(pkg.scripts!["test:coverage"]).toContain("--coverage");
  });

  test("the deploy workflow runs the coverage gate, not plain `npm test`", () => {
    // The deploy pipeline is the ONLY place the thresholds execute (local npm test
    // skips them by design), so the workflow regressing to `npm test` would kill the
    // gate without any other symptom. Plain-text match is deliberate: the workflow is
    // hand-maintained YAML and a structural parse would over-fit its layout.
    const workflow = fs.readFileSync(
      path.join(repoRoot, ".github", "workflows", "deploy-pages.yml"),
      "utf8"
    );
    expect(workflow).toMatch(/-\s*run:\s*npm run test:coverage/);
    // And the un-gated form must be gone — both lines present would still deploy
    // on the plain run if steps were ever reordered.
    expect(workflow).not.toMatch(/-\s*run:\s*npm test\s*$/m);
  });
});
