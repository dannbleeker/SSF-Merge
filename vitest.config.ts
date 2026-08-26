import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    coverage: {
      provider: "v8",
      // The engine is the product, and `src/host` earns the same floor for the
      // part of it that is a pure decision: the probe's verdicts are ordinary
      // functions over numbers and strings. What will NOT be counted here is
      // the Office.js-touching code, because a well-tested engine hiding behind
      // an untestable pane, or the reverse, is exactly what a pooled number
      // lets happen.
      include: ["src/core/**", "src/host/**", "src/pane/**"],
      // The pane's entry point is the one file there that touches Office.js and
      // cannot run in the suite. Everything else in src/pane — the step
      // machine, the copy, the renderer — is checked in jsdom and earns the
      // same floor as the engine. Counting main.ts would drag the number down
      // for a reason nobody can act on, which is how a threshold gets deleted.
      exclude: ["src/pane/main.ts", "src/pane/*.html", "src/pane/*.css"],
      reporter: ["text", "lcov"],
      // Floors sit a little under what the suite achieves today, so an ordinary
      // change does not go red, and they are raised deliberately rather than
      // tracking the current number. A threshold that follows coverage upward
      // on its own only ever ratchets, and the first hard week gets it deleted.
      thresholds: { statements: 90, branches: 65, functions: 92, lines: 95 },
    },
  },
});
