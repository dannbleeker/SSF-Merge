/**
 * Which commit the pane was built from, substituted by Vite at build time.
 *
 * Declared rather than imported because `define` replaces the identifier in the
 * source text: there is no module to import from, and the value does not exist
 * until the bundler puts it there. In the suite and under `tsc` it is undefined,
 * which is why every reader guards.
 */
declare const __BUILD_STAMP__: string | undefined;
