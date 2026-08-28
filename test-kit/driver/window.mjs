#!/usr/bin/env node
/**
 * Resize/maximise the CDP-attached Edge window.
 *
 * The Office Add-ins dialog is wider than the window it opened into, which puts
 * "Upload My Add-in" off-screen. Playwright cannot set a viewport on a browser
 * it did not launch, so go through CDP's Browser domain directly.
 *
 * Usage: node test-kit/driver/window.mjs [maximize|<width> <height>]
 */
import { chromium } from "playwright";

const browser = await chromium.connectOverCDP(process.env.SSF_CDP ?? "http://127.0.0.1:9333");
const ctx = browser.contexts()[0];
const page = ctx.pages().find((p) => /Doc\.aspx|_layouts/i.test(p.url())) ?? ctx.pages()[ctx.pages().length - 1];

const cdp = await ctx.newCDPSession(page);
const { windowId, bounds } = await cdp.send("Browser.getWindowForTarget");
console.log(`current bounds: ${JSON.stringify(bounds)}`);

const mode = process.argv[2] ?? "maximize";
if (mode === "emulate") {
  // The desktop runs at a high display scale, so the CSS viewport is far
  // smaller than the window and wide dialogs get clipped. Force the metrics.
  const width = Number(process.argv[3] ?? 1600);
  const height = Number(process.argv[4] ?? 1000);
  await cdp.send("Emulation.setDeviceMetricsOverride", {
    width,
    height,
    deviceScaleFactor: 1,
    mobile: false,
  });
  console.log(`emulating ${width}x${height} at dsf=1`);
} else if (mode === "clear") {
  await cdp.send("Emulation.clearDeviceMetricsOverride");
  console.log("cleared device metrics override");
} else if (mode === "maximize") {
  // A window already flagged maximised ignores a second maximise, so drop it to
  // normal first and let the state change actually take effect.
  await cdp.send("Browser.setWindowBounds", { windowId, bounds: { windowState: "normal" } });
  await cdp.send("Browser.setWindowBounds", { windowId, bounds: { windowState: "maximized" } });
} else {
  const width = Number(process.argv[2]);
  const height = Number(process.argv[3]);
  await cdp.send("Browser.setWindowBounds", { windowId, bounds: { windowState: "normal" } });
  await cdp.send("Browser.setWindowBounds", { windowId, bounds: { left: 0, top: 0, width, height } });
}

const after = await cdp.send("Browser.getWindowForTarget");
console.log(`new bounds    : ${JSON.stringify(after.bounds)}`);
const size = await page.evaluate(() => ({ w: window.innerWidth, h: window.innerHeight })).catch(() => null);
console.log(`viewport      : ${size ? `${size.w}x${size.h}` : "(could not read - frame busy)"}`);

await browser.close();
