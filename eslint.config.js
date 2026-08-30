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
  //
  // `test-kit/out/` joins them for the same reason: `docs/TEST-KIT.md` tells a
  // tester to write their verification script there, and `.gitignore` already
  // keeps the directory out of the repo. Without this, following the kit's own
  // instructions turns `npm run lint` red — nothing in there is inside
  // tsconfig, and type-aware linting refuses to parse a file no project owns.
  { ignores: ["dist/", "dist-lib/", "coverage/", "public/", "probe/", "test-kit/out/"] },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: {
        // The config files sit outside tsconfig's include, and type-aware
        // linting refuses to parse a file no project owns.
        //
        // The COUNT is raised because the default is eight and `scripts/` went
        // past it the moment the manifest generator arrived — at which point
        // every file over the line fails with a parsing error rather than a
        // finding, so `npm run lint` reports the linter's own limit as though
        // it were a defect in the code. These are small Node scripts; the
        // performance the cap protects is not at stake.
        projectService: {
          allowDefaultProject: ["*.js", "*.ts", "scripts/*.mjs", "test-kit/driver/*.mjs"],
          maximumDefaultProjectFileMatchCount_THIS_WILL_SLOW_DOWN_LINTING: 40,
        },
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
    // The round's browser driver: Node scripts whose `page.evaluate` callbacks
    // are serialised and run inside the browser, so `document` there is real —
    // the same arrangement as scripts/pane-shots.mjs below.
    //
    // Untyped .mjs with no project behind it, so every `JSON.parse` and every
    // CDP result is `any` and a function returning one is an unsafe return by
    // construction. These drive a browser; nothing about their types is load
    // bearing, and what they produce is checked by verify-package.mjs, which is
    // itself checked by mutate.mjs.
    files: ["test-kit/driver/**"],
    languageOptions: { globals: { ...globals.node, ...globals.browser } },
    rules: {
      "@typescript-eslint/no-unsafe-return": "off",
      "@typescript-eslint/no-unsafe-assignment": "off",
    },
  },
  {
    // ...except these two, which are Node scripts whose page.evaluate callbacks
    // are serialised and run inside the browser. `document` there is real.
    files: ["scripts/pane-shots.mjs", "scripts/manual-shots.mjs"],
    languageOptions: { globals: { ...globals.node, ...globals.browser } },
  },
  {
    // These two import `dist-lib/`, which is a BUILD ARTIFACT and is absent on
    // a fresh checkout. Type-aware linting resolves those imports to `error`
    // when it is not there and to real types when it is, so the verdict is a
    // fact about whether somebody has run `npm run build:lib` — green locally,
    // red in CI, and back again on the next build. A rule that answers
    // differently on the same source is not a rule. The scripts are covered by
    // `test/probe.test.ts` and by running them.
    files: ["scripts/build-probe.mjs", "scripts/read-answers.mjs"],
    rules: { "@typescript-eslint/no-unsafe-return": "off" },
  },
  {
    // Untyped `.mjs` with no project behind it: every `RegExp.exec` result and
    // every `JSON.parse` is `any`, so a function returning one is an unsafe
    // return by construction rather than by mistake. The rules these files hold
    // are gated by `test/manifest.test.ts`, which proves each one can still
    // fail — a stronger check than the type of an intermediate.
    files: [
      "scripts/manifest-rules.mjs",
      "scripts/manifest-source.mjs",
      "scripts/release-assets.mjs",
      "scripts/without-prose.mjs",
    ],
    rules: { "@typescript-eslint/no-unsafe-return": "off" },
  },
  {
    // The fixture builder and the tests reach into XML shapes on purpose.
    files: ["test/**"],
    rules: { "@typescript-eslint/no-explicit-any": "off" },
  },
);
