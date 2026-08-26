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
  build: {
    outDir: resolve(import.meta.dirname, "dist"),
    emptyOutDir: true,
    // Named explicitly because the root has no index.html: that name belongs to
    // the landing page in public/, and a pane called index.html would collide
    // with it in dist/.
    rollupOptions: { input: resolve(import.meta.dirname, "src/pane/taskpane.html") },
  },
});
