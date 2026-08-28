#!/usr/bin/env node
/**
 * Upload the kit's template into the currently-open OneDrive folder.
 *
 * OneDrive builds its <input type=file> on demand, so there is nothing to call
 * setInputFiles on until the menu item is clicked. Catch the filechooser the
 * click raises instead.
 *
 * Usage: node test-kit/driver/upload-template.mjs <local-file>
 */
import { chromium } from "playwright";

const file = process.argv[2] ?? "test-kit/SSF-Merge-test-template.pptx";
const browser = await chromium.connectOverCDP(process.env.SSF_CDP ?? "http://127.0.0.1:9333");
const ctx = browser.contexts()[0];
const page = ctx.pages().filter((p) => /onedrive|sharepoint/i.test(p.url()))[0] ?? ctx.pages()[0];

console.log(`folder: ${page.url()}`);

/** Click a control by its exact trimmed text, wherever it lives in the DOM. */
async function clickByText(text) {
  const done = await page.evaluate((t) => {
    const all = [...document.querySelectorAll('button, [role="menuitem"], span[role="button"], a')];
    const el = all.find((e) => e.textContent.trim() === t);
    if (!el) return false;
    el.click();
    return true;
  }, text);
  if (!done) throw new Error(`no control with text "${text}"`);
  await page.waitForTimeout(1200);
}

await clickByText("Create or upload");

const [chooser] = await Promise.all([
  page.waitForEvent("filechooser", { timeout: 30_000 }),
  clickByText("Files upload"),
]);
await chooser.setFiles(file);
console.log(`set ${file} on the file chooser`);

// The row appears when the upload commits, not when the bytes finish moving.
const name = file.split(/[\\/]/).pop();
await page
  .waitForFunction((n) => document.body.innerText.includes(n), name, { timeout: 120_000 })
  .then(() => console.log(`UPLOADED: ${name} is listed in the folder`))
  .catch(() => console.log(`the row for ${name} did not appear within 120s`));

await browser.close();
