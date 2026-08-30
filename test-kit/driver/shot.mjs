/**
 * Photograph the deck tab at an exact pixel size, with the operator's identity
 * out of the frame.
 *
 * Two things this file exists to stop being rediscovered:
 *
 *   1. `Emulation.setDeviceMetricsOverride` and the screenshot MUST share one
 *      CDP session. Set the size in one script and shoot in another and the
 *      override is gone by the time the shutter opens. Several attempts died
 *      at exactly that seam before anyone noticed the sessions differed.
 *   2. `page.screenshot` multiplies by the desktop's display scale, which is
 *      1.5 here, so a 1366-wide viewport produced a 2049-wide PNG. `scale:
 *      "css"` makes one raster pixel one CSS pixel. The store wants 1366x768
 *      exactly, and downscaling a bigger shot softens the pane's text.
 *
 * The hiding below is makeup for a photograph and nothing else. It removes the
 * account button, which carries both a real person's face and their name in its
 * aria-label, and any developer add-in tab, because Microsoft asks for neither
 * personal information nor unrelated tooling in a listing image. It runs after
 * the pane has already done the work, so it cannot make a round pass: the merge
 * has happened or it has not by the time this file is called.
 */
import { chromium } from "playwright";
import { pathToFileURL } from "node:url";

/** Strip the operator out of the editor chrome. Returns what it actually hid. */
const MAKEUP = `(() => {
  const hidden = [];
  const hide = (el, why) => { if (el) { el.style.visibility = 'hidden'; hidden.push(why); } };

  // The account button's aria-label is "Account manager for <full name>", so
  // this takes a name out of the picture as well as a face.
  hide(document.getElementById('O365_MainLink_Me'), 'account button');

  // Any add-in that put its own ribbon tab there. Matched on what the tab says
  // rather than on its id: the ids are positional (AddinTab0, AddinTab1) and
  // renumber themselves when a different set of add-ins is installed.
  for (const leaf of document.querySelectorAll('*')) {
    if (leaf.children.length) continue;
    const text = (leaf.textContent || '').trim();
    if (!/^(Script Lab)$/.test(text)) continue;
    const tab = leaf.closest('[role=tab]');
    hide(tab?.parentElement?.classList.contains('pivot-header-root') ? tab.parentElement : tab, 'tab: ' + text);
  }
  return hidden;
})()`;

export async function shoot(out, { width = 1366, height = 768, settle = 2000, makeup = true } = {}) {
  const browser = await chromium.connectOverCDP(process.env.SSF_CDP ?? "http://127.0.0.1:9333");
  try {
    const ctx = browser.contexts()[0];
    const page = ctx.pages().find((p) => /Doc\.aspx/i.test(p.url()));
    if (!page) throw new Error("no deck tab open");

    const cdp = await ctx.newCDPSession(page);
    await cdp.send("Emulation.setDeviceMetricsOverride", { width, height, deviceScaleFactor: 1, mobile: false });
    await page.waitForTimeout(settle);

    if (makeup) {
      // The editor is an out-of-process iframe with its own CDP target, which
      // is why this cannot simply be `page.evaluate`.
      const { evalIn, EDITOR } = await import("./pane.mjs");
      const hid = await evalIn(EDITOR, MAKEUP);
      if (!hid?.length) throw new Error("nothing was hidden: the account button and add-in tabs were not found");
      console.log("  hid:", hid.join(", "));
    }

    const seen = await page.evaluate(() => ({ w: innerWidth, h: innerHeight }));
    if (seen.w !== width || seen.h !== height) {
      throw new Error(`viewport is ${seen.w}x${seen.h}, wanted ${width}x${height}`);
    }
    const buf = await page.screenshot({ path: out, scale: "css" });
    console.log(`  wrote ${out} (${Math.round(buf.length / 1024)} KB)`);
    return buf;
  } finally {
    await browser.close();
  }
}

// `file://` glued to a Windows path is not the same string as `file:///C:/...`,
// so the usual hand-rolled main check never fires here: the script exits 0
// having done nothing, which reads exactly like a capture that worked.
if (import.meta.url === pathToFileURL(process.argv[1]).href) {
  await shoot(process.argv[2] ?? "test-kit/out/shot.png");
}
