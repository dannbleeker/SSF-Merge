// @ts-check
import js from "@eslint/js";
import globals from "globals";
import tseslint from "typescript-eslint";

/**
 * Type-aware linting, deliberately.
 *
 * The engine is async from top to bottom: every part read, every clone, every
 * write returns a promise. A forgotten `await` there does not throw, it merges
 * the wrong thing quietly, which is the failure mode this project can least
 * afford. `no-floating-promises` and `no-misused-promises` need type
 * information to see it at all, so the cost of the slower lint is the point.
 */
export default tseslint.config(
  // `probe/` is generated for Script Lab, not for this project: it has no
  // imports, targets a runtime with its own globals, and is asserted by
  // test/probe.test.ts rather than by the linter.
  { ignores: ["dist/", "dist-lib/", "coverage/", "public/", "probe/"] },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        // The config files sit outside tsconfig's include, and type-aware
        // linting refuses to parse a file no project owns.
        projectService: { allowDefaultProject: ["*.js", "*.ts", "scripts/*.mjs"] },
        tsconfigRootDir: import.meta.dirname,
      },
    },
    rules: {
      "@typescript-eslint/no-floating-promises": "error",
      "@typescript-eslint/no-misused-promises": "error",
      // A caught error is often re-thrown or reported; the engine does that a
      // lot and does not need a type argument each time.
      "@typescript-eslint/no-unsafe-assignment": "off",
      "@typescript-eslint/no-unsafe-member-access": "off",
      "@typescript-eslint/no-unsafe-argument": "off",
      "@typescript-eslint/no-unsafe-call": "off",
      // An unused argument named with a leading underscore is a documented
      // signature, not an oversight.
      "@typescript-eslint/no-unused-vars": ["error", { argsIgnorePattern: "^_", varsIgnorePattern: "^_" }],
    },
  },
  {
    // The build scripts run under Node, not in a pane.
    files: ["scripts/**", "*.config.{js,ts}"],
    languageOptions: { globals: globals.node },
  },
  {
    // The fixture builder and the tests reach into XML shapes on purpose.
    files: ["test/**"],
    rules: { "@typescript-eslint/no-explicit-any": "off" },
  },
);
