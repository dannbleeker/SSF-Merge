/**
 * Drive the task pane over CDP, by what a control says rather than by selector.
 *
 * Everything in a real-host round funnels through four questions: what does the
 * pane say, which button do I press, what do I type, and what does it look like.
 * Before this file every one of those was an ad-hoc `Runtime.evaluate` string
 * pasted into a one-off script, which is why a round cost dozens of round trips
 * and why each one re-learned the same two traps:
 *
 *   1. PowerPoint for the web nests the editor and the add-in in out-of-process
 *      iframes. Playwright's connectOverCDP does not attach to them here, so
 *      every call has to open its own websocket to that target.
 *   2. The pane is React. Assigning `input.value` updates the DOM and not the
 *      component, so the next render puts the old value back and the step's
 *      button stays disabled. `fill` goes through the native setter.
 *
 * Nothing here knows anything about the merge. `listing-shots.mjs` is the
 * script that knows the steps; this is the vocabulary it is written in.
 */
const BASE = process.env.SSF_CDP ?? "http://127.0.0.1:9333";

/** The pane's own origin. Matching on "taskpane.html" hits Copilot's pane too. */
export const PANE = "struktureretsundfornuft";
/** The PowerPoint editor OOPIF: the ribbon, the slide rail, the canvas. */
export const EDITOR = "ppt.aspx";

export const sleep = (ms) => new Promise((r) => setTimeout(r, ms));

async function targets() {
  const all = await (await fetch(`${BASE}/json/list`)).json();
  return all.filter((t) => t.type === "page" || t.type === "iframe");
}

async function find(match) {
  const hit = (await targets()).find((t) => t.url.includes(match));
  if (!hit) throw new Error(`no target whose url contains "${match}"`);
  return hit;
}

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
      const msgId = ++id;
      ws.send(JSON.stringify({ id: msgId, method, params }));
      return new Promise((res, rej) => pending.set(msgId, { resolve: res, reject: rej }));
    },
  };
}

/**
 * Click at a point inside a target, as a real mouse would.
 *
 * The slide rail does not respond to a dispatched `click` event, and its
 * thumbnails carry no id worth holding on to. Synthesising the press and the
 * release through the browser's own input pipeline does work, including on an
 * out-of-process iframe, which is not obvious: `Page.captureScreenshot` refuses
 * on the same target with "can only be executed on top-level targets".
 *
 * `modifiers` is CDP's bitfield. Shift is 8, and shift-click is how a range of
 * slides gets selected.
 */
export async function clickPoint(match, { x, y }, modifiers = 0) {
  const c = socket(await find(match));
  await c.ready;
  try {
    for (const type of ["mousePressed", "mouseReleased"]) {
      await c.send("Input.dispatchMouseEvent", {
        type,
        x,
        y,
        button: "left",
        clickCount: 1,
        modifiers,
        buttons: type === "mousePressed" ? 1 : 0,
      });
    }
  } finally {
    c.ws.close();
  }
}

/**
 * Send a real key press to a target, the way a keyboard would.
 *
 * `Input.dispatchKeyEvent` and not a synthesised `KeyboardEvent`: the editor
 * ignores dispatched key events for the same reason the rail ignores dispatched
 * clicks. `windowsVirtualKeyCode` is the part that is easy to leave out and
 * makes the press do nothing while still reporting success.
 */
export async function pressKey(match, key, { code = key, vk } = {}) {
  const c = socket(await find(match));
  await c.ready;
  try {
    for (const type of ["rawKeyDown", "keyUp"]) {
      await c.send("Input.dispatchKeyEvent", {
        type,
        key,
        code,
        windowsVirtualKeyCode: vk,
        nativeVirtualKeyCode: vk,
      });
    }
  } finally {
    c.ws.close();
  }
}

/**
 * Every thumbnail in the rail, keyed by the slide number it actually is.
 *
 * Keyed by `aria-posinset` and NOT by position in the returned array. The two
 * agree only while the rail is scrolled to the top, and a merge scrolls it: an
 * array-index version of this selected slide 5 while reporting slide 4, which
 * is the sort of off-by-one that reaches a store listing as a wrong picture.
 */
export async function thumbnails({ timeout = 15_000 } = {}) {
  const read = () =>
    evalIn(
      EDITOR,
      `[...document.querySelectorAll('[role=option][aria-label^=Slide]')].map(e => {
         const r = e.getBoundingClientRect();
         return { n: +e.getAttribute('aria-posinset'), of: +e.getAttribute('aria-setsize'),
                  x: Math.round(r.x + r.width / 2), y: Math.round(r.y + r.height / 2),
                  selected: e.getAttribute('aria-selected') === 'true' };
       })`,
    );

  // The label is matched by prefix, and that is load-bearing. A slide carrying
  // speaker notes is labelled "Slide, Has notes," rather than "Slide", so an
  // exact match cannot see it: the demo deck puts notes on slide 3, and this
  // reported a three-slide deck as holding slides 1 and 2 for as long as the
  // match was exact. It was diagnosed first as a re-render race, which it is
  // not, and the wrong diagnosis survived a fix that appeared to help.
  //
  // Waiting for a complete rail is still right, because the rail does also
  // re-render after a merge. `aria-setsize` is the rail's own count of itself,
  // which makes the completeness check a positive claim rather than a guess
  // about whether a gap means "still rendering" or "cannot see it at all".
  const deadline = Date.now() + timeout;
  let found = [];
  while (Date.now() < deadline) {
    found = await read();
    if (found.length && found.length === found[0].of) break;
    await sleep(500);
  }
  return new Map(found.map((t) => [t.n, t]));
}

/** The slide numbers the rail currently reports as selected. */
async function selection() {
  return [...(await thumbnails()).values()]
    .filter((t) => t.selected)
    .map((t) => t.n)
    .sort((a, b) => a - b);
}

/**
 * Scroll slide `n` into the rail's visible strip and return where to click it.
 *
 * A thumbnail can be in the DOM and still be unclickable: the rail holds all of
 * them, and the ones below the fold report a `y` past the bottom of the
 * viewport, where a synthesised click lands on nothing at all. That failure is
 * silent — the click is dispatched, it is accepted, and the selection simply
 * does not change — so asking for slides 6 and 7 in a nine-slide deck left the
 * selection sitting on 4 and looked like the shift-click had not worked.
 */
async function onScreen(n) {
  await evalIn(
    EDITOR,
    `(() => {
       const el = document.querySelector('[role=option][aria-posinset="${n}"][aria-label^=Slide]');
       el?.scrollIntoView({ block: 'center' });
     })()`,
  );
  await sleep(600);
  const thumb = (await thumbnails()).get(n);
  if (!thumb) throw new Error(`slide ${n} left the rail while scrolling to it`);
  const view = await evalIn(EDITOR, `({ w: innerWidth, h: innerHeight })`);
  if (thumb.y < 0 || thumb.y > view.h || thumb.x < 0 || thumb.x > view.w) {
    throw new Error(`slide ${n} sits at ${thumb.x},${thumb.y}, outside the ${view.w}x${view.h} editor`);
  }
  return thumb;
}

/**
 * Select slides by slide number, and prove it took before returning.
 *
 * Reads the selection back rather than trusting the clicks. A click that lands
 * on the rail's padding instead of a thumbnail changes nothing and says
 * nothing, and the next screenshot is the first place it shows up.
 *
 * The read has to be a poll, not a single look. Straight after a merge the rail
 * re-renders, and during that window every option reports `aria-selected=false`
 * — a single read there says "the rail selected none" about a rail that is
 * plainly showing a selection.
 *
 * Note what this must NOT do: infer selection from a slide's absence from the
 * list. That absence has meant three different things so far, one of which was
 * simply that the selector could not see a slide with speaker notes, and it
 * coincides with the right answer often enough to look like a signal.
 */
export async function selectSlides(numbers, { timeout = 10_000 } = {}) {
  const wanted = [...numbers].sort((a, b) => a - b);
  const thumbs = await thumbnails();
  for (const n of numbers) {
    if (!thumbs.has(n)) {
      throw new Error(`slide ${n} is not in the rail; it holds ${[...thumbs.keys()].join(",") || "nothing"}`);
    }
  }

  const [first, ...rest] = numbers;
  await clickPoint(EDITOR, await onScreen(first));
  await sleep(700);
  for (const n of rest) {
    await clickPoint(EDITOR, await onScreen(n), 8);
    await sleep(500);
  }

  const deadline = Date.now() + timeout;
  let got = [];
  while (Date.now() < deadline) {
    got = await selection();
    if (got.join(",") === wanted.join(",")) return got;
    await sleep(500);
  }
  throw new Error(`asked for slides ${wanted.join(",")}, the rail settled on ${got.join(",") || "none"}`);
}

/**
 * Which slide is showing and how many there are, as `{ slide, of }`.
 *
 * Reconciled against the rail rather than read off the status bar alone. The
 * status bar lags the document, and its two numbers lag TOGETHER, so they stay
 * consistent with each other while both being wrong: after a preview added two
 * slides it read "Slide 2 of 3" — a perfectly sensible pair describing the deck
 * as it was before. An earlier version of this only rejected impossible pairs
 * like "Slide 6 of 3" and let that one straight through, and the caller checked
 * `slide !== 4` against it and stopped a capture run that had not gone wrong.
 *
 * The rail's `aria-setsize` is a genuinely separate source for the total, and
 * the download in `fetch-deck.mjs` is the tiebreak when it matters: it was that
 * download, reporting five slides, that settled which readout was lying.
 */
export async function currentSlide({ timeout = 20_000 } = {}) {
  const deadline = Date.now() + timeout;
  let thumbs = new Map();
  while (Date.now() < deadline) {
    thumbs = await thumbnails();
    const selected = [...thumbs.values()].filter((t) => t.selected).map((t) => t.n);
    if (thumbs.size && selected.length) {
      const answer = { slide: Math.min(...selected), of: thumbs.size };

      // The status bar is kept only as a second opinion, never as the answer.
      // It refreshes when the canvas is next touched rather than when the deck
      // changes, so it can sit for minutes on a stale pair. Disagreement is
      // reported and not fatal, because the download settled which one lies.
      const text = await evalIn(EDITOR, `document.getElementById('GetSlideInformation')?.textContent ?? ''`);
      const m = /Slide (\d+) of (\d+)/.exec(text);
      if (m && (+m[1] !== answer.slide || +m[2] !== answer.of)) {
        console.warn(`  note: status bar says ${JSON.stringify(text)}, the rail says ${answer.slide} of ${answer.of}`);
      }
      return answer;
    }
    await sleep(500);
  }
  throw new Error(
    `the rail never reported a selected slide; it holds ${thumbs.size}. ` +
      `\`node test-kit/driver/fetch-deck.mjs\` downloads the document and counts.`,
  );
}

/** Run an expression in the first target matching `match` and return its value. */
export async function evalIn(match, expression) {
  const c = socket(await find(match));
  await c.ready;
  try {
    const r = await c.send("Runtime.evaluate", { expression, returnByValue: true, awaitPromise: true });
    if (r.exceptionDetails) throw new Error(r.exceptionDetails.exception?.description ?? "threw");
    return r.result?.value;
  } finally {
    c.ws.close();
  }
}

/**
 * What the pane currently says, collapsed to single spaces.
 *
 * Waits for a document that has a body. A frame caught mid-reload has none, and
 * reading `document.body.innerText` there throws "Cannot read properties of
 * null" out of the middle of whatever was driving — which points at this line
 * and not at the reload that caused it. The reload is legitimate and expected:
 * `reset.mjs` performs one deliberately.
 */
export async function says(match = PANE, { timeout = 20_000 } = {}) {
  const deadline = Date.now() + timeout;
  let last = "no document";
  while (Date.now() < deadline) {
    const got = await evalIn(
      match,
      `document.body ? document.body.innerText.replace(/\\s+/g, ' ').trim() : null`,
    ).catch((e) => {
      // The context is torn down and rebuilt during a reload, so "Execution
      // context was destroyed" here is the reload working, not a failure.
      last = e.message.slice(0, 80);
      return null;
    });
    if (got !== null) return got;
    await sleep(500);
  }
  throw new Error(`the pane never produced a document within ${timeout}ms (last: ${last})`);
}

/**
 * Press the control whose accessible name is `label`.
 *
 * Returns "clicked", "disabled" or "not found" rather than throwing, so a
 * caller can decide whether a missing button is a bug or just a step that has
 * not rendered yet. It reports `disabled` instead of clicking into the void:
 * a click on a disabled button succeeds silently and looks identical to a
 * click that worked, which is exactly how a round loses an hour.
 */
export async function click(label, match = PANE) {
  return evalIn(
    match,
    `(() => {
      const wanted = ${JSON.stringify(label)};
      const all = [...document.querySelectorAll('button,a,[role=button],[role=menuitem],[role=tab]')];
      const el = all.find(e => ((e.getAttribute('aria-label') || e.textContent || '').trim()) === wanted);
      if (!el) return 'not found';
      if (el.disabled || el.getAttribute('aria-disabled') === 'true') return 'disabled';
      el.click();
      return 'clicked';
    })()`,
  );
}

/**
 * Type `value` into the field whose label, placeholder or aria-label is `label`.
 *
 * React owns the value, so setting the property directly is a no-op the next
 * time the component renders. This calls the prototype's native setter and
 * then fires `input`, which is what React's synthetic event system listens to.
 */
export async function fill(label, value, match = PANE) {
  return evalIn(
    match,
    `(() => {
      const wanted = ${JSON.stringify(label)};
      const fields = [...document.querySelectorAll('input,textarea')];
      const named = (f) => {
        const byLabel = f.labels?.[0]?.textContent ?? '';
        return [f.getAttribute('aria-label'), f.placeholder, byLabel, f.name, f.id]
          .filter(Boolean).map(s => s.trim());
      };
      const el = fields.find(f => named(f).some(n => n.toLowerCase() === wanted.toLowerCase()));
      if (!el) return 'not found: ' + JSON.stringify(fields.map(named));
      const proto = el instanceof HTMLTextAreaElement ? HTMLTextAreaElement : HTMLInputElement;
      const setter = Object.getOwnPropertyDescriptor(proto.prototype, 'value').set;
      setter.call(el, ${JSON.stringify(value)});
      el.dispatchEvent(new Event('input', { bubbles: true }));
      el.dispatchEvent(new Event('change', { bubbles: true }));
      return 'filled';
    })()`,
  );
}

/** Every control the pane is currently offering, and whether it can be pressed. */
export async function controls(match = PANE) {
  return evalIn(
    match,
    `[...document.querySelectorAll('button,input,textarea,a,[role=button]')].map(e =>
       e.tagName + ' "' + (e.getAttribute('aria-label') || e.textContent || e.placeholder || '').trim().slice(0, 50) + '"'
       + (e.disabled ? ' [disabled]' : ''))`,
  );
}

/**
 * Wait until the pane's text contains `needle`.
 *
 * Every step of the merge is asynchronous inside Office, and a fixed sleep is
 * either a slow round or a flaky one. Throws naming what it waited for and what
 * it saw, because "timed out" alone sends you back to a screenshot to find out.
 */
export async function until(needle, { timeout = 30_000, match = PANE } = {}) {
  const deadline = Date.now() + timeout;
  let seen = "";
  while (Date.now() < deadline) {
    seen = await says(match);
    if (seen.includes(needle)) return seen;
    await sleep(500);
  }
  throw new Error(`waited ${timeout}ms for ${JSON.stringify(needle)}; pane says: ${seen.slice(0, 300)}`);
}
