import { defineConfig } from "vitest/config";

export default defineConfig({
  test: {
    include: ["test/**/*.test.ts"],
    environment: "node",
    coverage: {
      provider: "v8",
      // The engine is the product. The host layer and the pane will be covered
      // by different means, and counting them here would let a well-tested
      // engine hide behind an untestable pane, or the reverse.
      include: ["src/core/**"],
      reporter: ["text", "lcov"],
      // Floors sit a little under what the suite achieves today, so an ordinary
      // change does not go red, and they are raised deliberately rather than
      // tracking the current number. A threshold that follows coverage upward
      // on its own only ever ratchets, and the first hard week gets it deleted.
      thresholds: { statements: 90, branches: 65, functions: 92, lines: 95 },
    },
  },
});
