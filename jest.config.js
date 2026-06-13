// Jest configuration for the pure core engine.
//
// The core/ modules are deliberately framework-free and Office.js-free, so the
// entire engine is exercised in a plain Node environment with no Word host and
// no DOM mock — OOXML is parsed via @xmldom/xmldom (works identically in Node
// and the task-pane browser). ts-jest transpiles TypeScript on the fly.
//
// tsconfig targets `module: ESNext` for webpack tree-shaking, but Jest runs on
// CommonJS — so we override the emitted module to CommonJS in the transform.
/** @type {import('ts-jest').JestConfigWithTsJest} */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  roots: ["<rootDir>/__tests__", "<rootDir>/src"],
  testMatch: ["**/__tests__/**/*.test.ts"],
  transform: {
    // isolatedModules skips ts-jest's per-file semantic re-check — `tsc --noEmit` is the
    // dedicated type gate (local + CI), so tests only need transpilation. ~15% faster cold
    // runs; warm runs unchanged (transform cache).
    "^.+\\.tsx?$": ["ts-jest", { isolatedModules: true, tsconfig: { module: "CommonJS" } }]
  },
  collectCoverageFrom: [
    "src/core/**/*.ts",
    // The Office.js adapter (officeWordPort.ts) cannot run without a Word host,
    // so it is excluded from unit coverage and validated by Stage 2 integration.
    "!src/core/officeWordPort.ts",
    // The gdocs pure engine is held to the same bar. Its host adapter
    // (google-docs/src/adapter/) is outside this glob for the same reason as
    // officeWordPort.ts — it cannot run without the Apps Script host.
    "google-docs/src/core/**/*.ts"
  ],
  coverageThreshold: {
    // NOTE: a path-scoped entry REMOVES its matched files from the global
    // group (Jest semantics), so adding the gdocs floor leaves the Word
    // engine's global gate numerically identical — the deploy gate is
    // unchanged (case 001-F7, amended).
    global: { branches: 80, functions: 90, lines: 90, statements: 90 },
    "google-docs/src/core/": { branches: 80, functions: 90, lines: 90, statements: 90 }
  },
  clearMocks: true
};
