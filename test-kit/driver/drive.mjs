#!/usr/bin/env node
/**
 * Stepwise driver for the already-running Edge (CDP on 127.0.0.1:9333).
 *
 * The browser stays alive between invocations, so the round can be driven one
 * observation at a time instead of one long blind script — which matters for a
 * UI nobody here has automated before.
 *
 * Usage:
 *   node test-kit/driver/drive.mjs goto <url>
 *   node test-kit/driver/drive.mjs shot <file.png> [--full]
 *   node test-kit/driver/drive.mjs text [--frame <urlpart>] [--max 4000]
 *   node test-kit/driver/drive.mjs ui   [--frame <urlpart>]     interactive elements
 *   node test-kit/driver/drive.mjs frames
 *   node test-kit/driver/drive.mjs click <selector> [--frame <urlpart>]
 *   node test-kit/driver/drive.mjs fill <selector> <value> [--frame <urlpart>]
 *   node test-kit/driver/drive.mjs upload <selector> <file> [--frame <urlpart>]
 *   node test-kit/driver/drive.mjs eval "<js>" [--frame <urlpart>]
 *   node test-kit/driver/drive.mjs url
 */
import { chromium } from "playwright";

const CDP = process.env.SSF_CDP ?? "http://127.0.0.1:9333";
const argv = process.argv.slice(2);
const cmd = argv[0];

function flag(name, fallback) {
  const i = argv.indexOf(`--${name}`);
  return i === -1 ? fallback : argv[i + 1];
}
const positional = argv.filter((a, i) => {
  if (i === 0) return false;
  if (a.startsWith("--")) return false;
  return !(i > 0 && argv[i - 1].startsWith("--") && argv[i - 1] !== "--full");
});

const browser = await chromium.connectOverCDP(CDP);
const ctx = browser.contexts()[0];

/** The Office tab, not an extension background page. */
function pickPage() {
  const pages = ctx.pages().filter((p) => /^https?:/.test(p.url()));
  const want = process.env.SSF_PAGE;
  if (want) {
    const hit = pages.find((p) => p.url().includes(want));
    if (hit) return hit;
  }
  // The editor tab, not the file list it was launched from. Both are on
  // onedrive.live.com, so match the editor's own path first.
  return (
    pages.find((p) => /Doc\.aspx|_layouts|powerpoint|:p:/i.test(p.url())) ??
    pages.find((p) => /onedrive|sharepoint/i.test(p.url())) ??
    pages[pages.length - 1] ??
    ctx.pages()[0]
  );
}
const page = pickPage();

/**
 * Resolve the target: the page, a frame by index (--fi N), or a frame whose URL
 * contains --frame. PowerPoint for the web puts its editor in a child frame
 * whose URL is blank, so index is the only handle that addresses it.
 */
function target() {
  const fi = flag("fi");
  if (fi !== undefined) {
    const f = page.frames()[Number(fi)];
    if (!f) {
      console.error(`no frame at index ${fi}; there are ${page.frames().length}`);
      process.exit(3);
    }
    return f;
  }
  const fp = flag("frame");
  if (!fp) return page;
  const f = page.frames().find((fr) => fr.url().includes(fp));
  if (!f) {
    console.error(`no frame whose url contains "${fp}". Frames:`);
    for (const fr of page.frames()) console.error(`  ${fr.url()}`);
    process.exit(3);
  }
  return f;
}

try {
  switch (cmd) {
    case "url":
      for (const p of ctx.pages()) console.log(`${p === page ? "*" : " "} ${p.url()}`);
      break;

    case "goto":
      await page.goto(positional[0], { waitUntil: "domcontentloaded", timeout: 90_000 });
      await page.waitForTimeout(2500);
      console.log(`now at: ${page.url()}`);
      break;

    case "shot": {
      const full = argv.includes("--full");
      await page.screenshot({ path: positional[0], fullPage: full });
      console.log(`wrote ${positional[0]} (url: ${page.url()})`);
      break;
    }

    case "text": {
      const max = Number(flag("max", 4000));
      const t = await target().evaluate(() => document.body?.innerText ?? "");
      console.log(t.replace(/\n{3,}/g, "\n\n").slice(0, max));
      break;
    }

    case "ui": {
      const items = await target().evaluate(() => {
        const out = [];
        const sel =
          'button, a[href], input, select, textarea, [role="button"], [role="menuitem"], [role="tab"], [role="link"], [role="combobox"], [contenteditable="true"]';
        for (const el of document.querySelectorAll(sel)) {
          const r = el.getBoundingClientRect();
          if (r.width === 0 && r.height === 0) continue;
          const style = getComputedStyle(el);
          if (style.visibility === "hidden" || style.display === "none") continue;
          const label =
            el.getAttribute("aria-label") ||
            (el.innerText || "").trim().slice(0, 60) ||
            el.getAttribute("title") ||
            el.getAttribute("placeholder") ||
            el.getAttribute("value") ||
            "";
          out.push(
            [
              el.tagName.toLowerCase(),
              el.type ? `type=${el.type}` : "",
              el.id ? `#${el.id}` : "",
              el.getAttribute("data-automationid") ? `[dai=${el.getAttribute("data-automationid")}]` : "",
              el.getAttribute("role") ? `role=${el.getAttribute("role")}` : "",
              label ? `"${label}"` : "",
            ]
              .filter(Boolean)
              .join(" "),
          );
        }
        return out;
      });
      console.log(items.join("\n").slice(0, Number(flag("max", 8000))));
      break;
    }

    case "frames": {
      const fs = page.frames();
      for (let i = 0; i < fs.length; i++) {
        let len = -1;
        try {
          len = await fs[i].evaluate(() => (document.body ? document.body.innerText.length : -1));
        } catch {
          /* a frame can be detached or still navigating */
        }
        console.log(`[${i}] textLen=${len} name=${fs[i].name() || "-"} url=${fs[i].url() || "(blank)"}`);
      }
      break;
    }

    case "click": {
      // The Office ribbon fails Playwright's actionability checks a lot — items
      // are overlapped by their own hit-target layers. Fall back to a DOM click,
      // which the app's handlers respond to just as well.
      const sel = positional[0];
      let how = "playwright";
      try {
        await target().locator(sel).first().click({ timeout: 8000 });
      } catch {
        try {
          how = "force";
          await target().locator(sel).first().click({ timeout: 8000, force: true });
        } catch {
          how = "dom";
          const ok = await target().evaluate((s) => {
            const el = document.querySelector(s);
            if (!el) return false;
            el.click();
            return true;
          }, sel);
          if (!ok) {
            console.error(`no element matched ${sel}`);
            process.exit(4);
          }
        }
      }
      await page.waitForTimeout(1500);
      console.log(`clicked ${sel} (via ${how})`);
      break;
    }

    case "press":
      for (const k of positional[0].split("+++")) {
        await page.keyboard.press(k);
        await page.waitForTimeout(600);
      }
      console.log(`pressed ${positional[0]}`);
      break;

    /** A real mouse click at the centre of the element, in page coordinates. */
    case "mouseclick": {
      const box = await target().locator(positional[0]).first().boundingBox();
      if (!box) {
        console.error(`no box for ${positional[0]}`);
        process.exit(4);
      }
      await page.mouse.click(box.x + box.width / 2, box.y + box.height / 2);
      await page.waitForTimeout(1500);
      console.log(`mouse-clicked ${positional[0]} at ${Math.round(box.x)},${Math.round(box.y)}`);
      break;
    }

    case "clicktext": {
      const t = positional[0];
      const ok = await target().evaluate((txt) => {
        const all = [
          ...document.querySelectorAll('button, [role="menuitem"], [role="option"], [role="tab"], a, span, div'),
        ];
        const el = all.find((e) => e.textContent.trim() === txt && e.offsetParent !== null);
        if (!el) return false;
        el.click();
        return true;
      }, t);
      if (!ok) {
        console.error(`no visible control with exact text "${t}"`);
        process.exit(4);
      }
      await page.waitForTimeout(1500);
      console.log(`clicked text "${t}"`);
      break;
    }

    case "fill":
      await target().locator(positional[0]).first().fill(positional[1], { timeout: 30_000 });
      console.log(`filled ${positional[0]}`);
      break;

    case "upload":
      await target().locator(positional[0]).first().setInputFiles(positional[1], { timeout: 30_000 });
      await page.waitForTimeout(2000);
      console.log(`set ${positional[1]} on ${positional[0]}`);
      break;

    case "eval": {
      const r = await target().evaluate((src) => {
        const v = eval(src);
        return typeof v === "object" ? JSON.stringify(v) : String(v);
      }, positional[0]);
      console.log(r);
      break;
    }

    default:
      console.error(`unknown command: ${cmd}`);
      process.exit(2);
  }
} finally {
  // Detach only. Closing would take the browser and the session down with it.
  await browser.close();
}
