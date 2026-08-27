import { resolve } from "node:path";
import { defineConfig } from "vite";

/**
 * The task pane, built for GitHub Pages.
 *
 * `base: "./"` because the add-in is served from a path that the manifest
 * points at, and absolute asset URLs would break the moment the site moved.
 * `publicDir` carries the CNAME that binds the custom domain and the landing
 * page, so one build produces the whole site.
 */
export default defineConfig({
  // The pane is the root, so it builds to dist/taskpane.html rather than to
  // dist/src/pane/taskpane.html. The manifest points at that URL and a
  // manifest change is the one kind that costs the owner a re-sideload, so the
  // path is worth getting right before there is a manifest to change.
  root: resolve(import.meta.dirname, "src/pane"),
  base: "./",
  publicDir: resolve(import.meta.dirname, "public"),
  // The DEV MANIFEST names https://localhost:3000/taskpane.html, so the dev
  // server has to answer there — a manifest pointing at a port nothing serves
  // is a blank pane with a generic error. `strictPort` because falling back to
  // 3001 would be the same failure with an extra step.
  //
  // Office requires HTTPS for a sideloaded add-in. `npm run dev` serves plain
  // HTTP, which is right for looking at the pane in a browser; sideloading it
  // needs a certificate — `npx office-addin-dev-certs install` and
  // `--https.key`/`--https.cert`. docs/MANUAL.md carries the two commands.
  server: { port: 3000, strictPort: true },
  build: {
    outDir: resolve(import.meta.dirname, "dist"),
    emptyOutDir: true,
    // Named explicitly because the root has no index.html: that name belongs to
    // the landing page in public/, and a pane called index.html would collide
    // with it in dist/.
    rollupOptions: { input: resolve(import.meta.dirname, "src/pane/taskpane.html") },
  },
});
