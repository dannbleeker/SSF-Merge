#!/usr/bin/env node
/**
 * Pull the open deck's bytes out of the browser, using the page's own session.
 *
 * The round's deliverable is the merged FILE. Clicking Download in the OneDrive
 * UI put nothing anywhere this process could find — the browser's download
 * plumbing is not reachable from here — so fetch from inside the page instead:
 * same cookies, same origin, and the bytes come back over CDP.
 *
 * Everything it needs is read off the open editor tab, so there is nothing to
 * configure and nothing about one machine or one account baked in.
 *
 * IT CAN RACE THE SAVE. PowerPoint for the web autosaves, and the download
 * endpoint serves what OneDrive has COMMITTED — not what is on the screen. On
 * 2026-08-30 a fetch taken straight after a merge returned the pre-merge file:
 * 3 slides, 58 KB, HTTP 200, no error of any kind, and the only thing wrong
 * with it was that it was last minute's deck. A byte count is not a signal
 * here; the slide count is.
 *
 * So the slide count is always printed, and `--expect-slides N` turns it into
 * an assertion that waits: it re-fetches until the deck says N, rather than
 * handing back a stale file that looks perfectly sound.
 *
 * Usage: node test-kit/driver/fetch-deck.mjs <out-file> [--expect-slides N]
 */
import { writeFileSync } from "node:fs";
import JSZip from "jszip";

const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";
const args = process.argv.slice(2).filter((a) => !a.startsWith("--"));
const out = args[0] ?? "test-kit/out/host-deck.pptx";
const expectIdx = process.argv.indexOf("--expect-slides");
const expectSlides = expectIdx === -1 ? null : Number(process.argv[expectIdx + 1]);
if (expectIdx !== -1 && !Number.isInteger(expectSlides)) {
  console.error("--expect-slides needs a whole number");
  process.exit(2);
}

const targets = await (await fetch(`${BASE}/json/list`)).json();

// The EDITOR tab, not the file list it was launched from: only the editor's URL
// carries the document's id.
const target = targets.find((t) => t.type === "page" && /Doc\.aspx|_layouts/i.test(t.url));
if (!target) {
  console.error("no editor tab open (looking for a Doc.aspx / _layouts page)");
  process.exit(3);
}

const drive = /\/(personal\/[0-9a-f]+)\//i.exec(target.url);
const doc = /sourcedoc=%7B([0-9A-Fa-f-]+)%7D/.exec(target.url) ?? /sourcedoc=\{([0-9A-Fa-f-]+)\}/.exec(target.url);
if (!drive || !doc) {
  console.error(`could not read the drive and document id out of:\n  ${target.url}`);
  process.exit(3);
}
const origin = new URL(target.url).origin;
const uniqueId = doc[1].toLowerCase();
console.error(`document: ${uniqueId} on ${drive[1]}`);

const ws = new WebSocket(target.webSocketDebuggerUrl);
let id = 0;
const pending = new Map();
const send = (method, params = {}) => {
  const msgId = ++id;
  ws.send(JSON.stringify({ id: msgId, method, params }));
  return new Promise((resolve, reject) => pending.set(msgId, { resolve, reject }));
};
ws.addEventListener("message", (e) => {
  const m = JSON.parse(e.data);
  if (m.id && pending.has(m.id)) {
    const { resolve, reject } = pending.get(m.id);
    pending.delete(m.id);
    if (m.error) reject(new Error(JSON.stringify(m.error)));
    else resolve(m.result);
  }
});
await new Promise((res, rej) => {
  ws.addEventListener("open", res);
  ws.addEventListener("error", () => rej(new Error("websocket failed to open")));
});

/**
 * Try each download shape and keep the first that answers with a zip.
 *
 * Consumer OneDrive runs on SharePoint now, so `_layouts/15/download.aspx` is
 * the one that answers; the others are kept because the estate moves and a 404
 * costs nothing next to a round lost to a renamed endpoint.
 */
const expression = `
(async () => {
  const id = ${JSON.stringify(uniqueId)};
  const urls = [
    ${JSON.stringify(`${origin}/${drive[1]}/_layouts/15/download.aspx?UniqueId=`)} + id,
    ${JSON.stringify(`${origin}/${drive[1]}/_layouts/15/download.aspx?UniqueId={`)} + id + '}',
    ${JSON.stringify(`${origin}/download?resid=`)} + id
  ];
  for (const url of urls) {
    try {
      const r = await fetch(url, { credentials: 'include' });
      if (!r.ok) continue;
      const buf = new Uint8Array(await r.arrayBuffer());
      // A sign-in page is a 200 too. Only a zip starts "PK".
      if (buf.length < 4 || buf[0] !== 0x50 || buf[1] !== 0x4b) continue;
      let s = '';
      for (let i = 0; i < buf.length; i += 0x8000) s += String.fromCharCode.apply(null, buf.subarray(i, i + 0x8000));
      return JSON.stringify({ url, size: buf.length, b64: btoa(s) });
    } catch { /* try the next shape */ }
  }
  return JSON.stringify({ error: 'no download url returned a zip' });
})()
`;

/** How many slides a fetched package holds, or null if it will not open. */
async function slidesIn(buf) {
  try {
    const zip = await JSZip.loadAsync(buf);
    return Object.keys(zip.files).filter((n) => /^ppt\/slides\/slide\d+\.xml$/.test(n)).length;
  } catch {
    return null;
  }
}

/**
 * Long enough for OneDrive to commit a merge, short enough to fail a round
 * rather than hang it. The observed lag on 2026-08-30 was under thirty seconds.
 */
const TRIES = expectSlides === null ? 1 : 12;

try {
  let payload;
  let slides = null;
  for (let attempt = 1; attempt <= TRIES; attempt++) {
    const res = await send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
    if (res.exceptionDetails) throw new Error(res.exceptionDetails.exception?.description ?? "evaluate threw");
    payload = JSON.parse(res.result.value);
    if (payload.error) {
      console.log(`FAILED: ${payload.error}`);
      process.exit(1);
    }
    slides = await slidesIn(Buffer.from(payload.b64, "base64"));
    if (expectSlides === null || slides === expectSlides) break;
    console.log(`waiting: deck says ${slides} slide(s), expected ${expectSlides} (attempt ${attempt}/${TRIES})`);
    await new Promise((r) => setTimeout(r, 5000));
  }

  writeFileSync(out, Buffer.from(payload.b64, "base64"));
  console.log(`from : ${payload.url}`);
  console.log(`WROTE: ${out} (${payload.size} bytes, ${slides ?? "unreadable"} slide(s))`);

  if (expectSlides !== null && slides !== expectSlides) {
    // Said loudly, and with a non-zero exit. A stale deck is the failure this
    // whole option exists for, and it is indistinguishable from a good one
    // until somebody counts.
    console.log(
      `STALE: expected ${expectSlides} slide(s) and the download still serves ${slides}. ` +
        `OneDrive has not committed the merge — the file written is LAST SAVE, not what is on screen.`,
    );
    process.exit(1);
  }
} finally {
  ws.close();
}
