/**
 * Open a deck in the round's OneDrive folder, by name, without clicking it.
 *
 * Clicking a row in this view is not reliable, and every one of its failures
 * looks exactly like a click that missed:
 *
 *   - a double-click SELECTS the row rather than opening it;
 *   - a row left selected swallows every LATER double-click, so once one attempt
 *     has gone wrong the next twenty go wrong the same way, and reloading the
 *     folder is the only thing that clears it;
 *   - a row below the fold has a rect off the bottom of the window, so the click
 *     is dispatched, accepted, and lands on nothing — the same trap the slide
 *     rail sets in `pane.mjs`;
 *   - Enter on a selected row does not open it either.
 *
 * Between them those cost most of a round twice. Each row does carry the item's
 * GUID, in the id of its Copilot button (`{46807B9C-...}-hero-copilot`), which
 * is enough to build the editor URL and navigate straight there. That has never
 * failed.
 *
 *   import { openDeck } from "./decks.mjs";
 *   await openDeck("round-deck-B2.pptx");
 *
 * The account is read from the open folder tab rather than written down here.
 * A OneDrive editor URL contains a personal account identifier, and this
 * repository is public.
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";

const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

const targets = async () =>
  (await (await fetch(`${BASE}/json/list`)).json()).filter((t) => t.type === "page" || t.type === "iframe");

function socket(target) {
  const ws = new WebSocket(target.webSocketDebuggerUrl);
  let id = 0;
  const pending = new Map();
  ws.addEventListener("message", (e) => {
    const m = JSON.parse(e.data);
    const p = pending.get(m.id);
    if (!p) return;
    pending.delete(m.id);
    if (m.error) p.reject(new Error(m.error.message));
    else p.resolve(m.result);
  });
  return {
    ws,
    ready: new Promise((r) => ws.addEventListener("open", r)),
    send(method, params = {}) {
      const i = ++id;
      ws.send(JSON.stringify({ id: i, method, params }));
      return new Promise((res, rej) => pending.set(i, { resolve: res, reject: rej }));
    },
  };
}

async function folderTab() {
  const tab = (await targets()).find((t) => /onedrive\.live\.com\/my/.test(t.url));
  if (!tab) throw new Error("the OneDrive folder tab is not open; open the round's folder first");
  return tab;
}

/**
 * The `personal/<id>` segment the editor URL needs, taken from the folder tab.
 *
 * The folder's own URL carries it inside the `id` query parameter, encoded:
 * `?id=%2Fpersonal%2F<account>%2FDocuments%2F...`.
 */
function siteFrom(folderUrl) {
  const id = new URL(folderUrl).searchParams.get("id") ?? "";
  const match = /^\/(personal\/[^/]+)\//.exec(decodeURIComponent(id));
  if (!match) throw new Error(`could not read the account from the folder URL: ${folderUrl.slice(0, 80)}`);
  return `https://onedrive.live.com/${match[1]}/_layouts/15/Doc.aspx`;
}

/** Every `.pptx` in the open folder, as `{ name: "{GUID}" }`. */
export async function deckIds() {
  const c = socket(await folderTab());
  await c.ready;
  try {
    const found = await c.send("Runtime.evaluate", {
      expression: `(() => {
        const out = {};
        for (const row of document.querySelectorAll('[role=row]')) {
          const name = (row.innerText || '').split('\\n')[0].trim();
          if (!/\\.pptx$/.test(name)) continue;
          const tagged = [...row.querySelectorAll('[id]')].find((e) => /^\\{[0-9A-F-]+\\}-/i.test(e.id));
          if (tagged) out[name] = tagged.id.slice(0, tagged.id.indexOf('}') + 1);
        }
        return out;
      })()`,
      returnByValue: true,
    });
    return found.result?.value ?? {};
  } finally {
    c.ws.close();
  }
}

/**
 * Navigate to `name` and wait for the editor.
 *
 * Reuses the open deck tab where there is one, so a round does not accumulate
 * a tab per deck — and so `pane.mjs`, which matches targets by URL and takes
 * the first, is never choosing between two editors.
 */
export async function openDeck(name, { timeout = 150_000 } = {}) {
  const ids = await deckIds();
  const guid = ids[name];
  if (!guid) {
    throw new Error(`${name} is not in the open folder; it holds ${Object.keys(ids).join(", ") || "no .pptx files"}`);
  }
  const site = siteFrom((await folderTab()).url);
  const url = `${site}?sourcedoc=${encodeURIComponent(guid)}&file=${encodeURIComponent(name)}&action=edit&mobileredirect=true`;

  const open = (await targets()).find((t) => t.url.includes("Doc.aspx"));
  if (open) {
    const c = socket(open);
    await c.ready;
    await c.send("Page.navigate", { url });
    c.ws.close();
  } else {
    await fetch(`${BASE}/json/new?${encodeURIComponent(url)}`, { method: "PUT" });
  }

  const deadline = Date.now() + timeout;
  while (Date.now() < deadline) {
    await sleep(5000);
    if ((await targets()).some((t) => t.url.includes("ppt.aspx"))) {
      // The editor target exists before the deck is usable; the pane's own
      // waits handle the rest, but a caller reading the rail immediately gets
      // a count of zero without this.
      await sleep(10_000);
      return url;
    }
  }
  throw new Error(`${name} never opened; navigated to ${url.slice(0, 90)}`);
}
