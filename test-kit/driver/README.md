# Driving the round in PowerPoint for the web

Twelve small scripts that put the add-in through the kit's run in a real host,
and check what came back. Built during the round of 2026-08-28, which found two
defects with them ([#64](https://github.com/dannbleeker/SSF-Merge/pull/64),
[#66](https://github.com/dannbleeker/SSF-Merge/pull/66)), and extended by the
round of 2026-08-30, which found one defect in the product and two in these
tools.

They are here so the next round starts from a working driver. Rebuilding this
was most of the cost of the last one, and none of it was interesting.

Everything they write goes to `test-kit/out/`, which is gitignored: evidence from
one machine on one day does not belong in the repo.

## What each one is for

| Script | What it does |
| --- | --- |
| `web-login.mjs` | Opens a headed browser on a persistent profile and waits while a human signs in. It never types the credentials. |
| `window.mjs` | Resizes or maximises the CDP-attached window, and can force a viewport size. |
| `drive.mjs` | Stepwise driver for the page: `goto`, `shot`, `text`, `ui`, `frames`, `click`, `clicktext`, `mouseclick`, `press`, `fill`, `upload`, `eval`, `url`. |
| `cdp-eval.mjs` | Evaluates JavaScript inside one CDP target, by URL substring. `list` prints the targets. |
| `cdp-type.mjs` | Focuses an element and types into it with real key events; `insert` mode pastes. |
| `cdp-key.mjs` | Sends real ArrowUp/ArrowDown/Enter to an element. What the pane's `<select>` needs — `.value` plus a `change` event does not move React. |
| `cdp-setfile.mjs` | Puts local files on an `<input type=file>`, including one inside a nested frame. |
| `close-popup.mjs` | Closes pages whose URL matches a substring. |
| `upload-template.mjs` | Uploads a local file into the OneDrive folder currently on screen. |
| `fetch-deck.mjs` | Downloads the open deck's bytes through the page's own session, and prints its slide count. `--expect-slides N` waits for OneDrive to commit rather than handing back last save. |
| `verify-package.mjs` | Checks a merged deck, anchored per row to the slide's own title. Knows BOTH kit decks — thirteen checks for the main template, seven for the sunburst — understands a deck that still holds its template, and refuses a deck that is neither. |
| `mutate.mjs` | Breaks a reference deck six ways and asserts each is caught by its own guard — and refuses to count a guard whose check was already red. |

## The run, roughly

```bash
node test-kit/driver/web-login.mjs                 # once; sign in by hand
# then launch Edge yourself with --remote-debugging-port=9333 on the same
# profile, and drive it:
node test-kit/driver/drive.mjs url
node test-kit/driver/upload-template.mjs test-kit/SSF-Merge-test-template.pptx
node test-kit/driver/cdp-eval.mjs list

# The pane is its own CDP target. Match it on the ADD-IN's host: "taskpane.html"
# also matches a Microsoft pane, and answers about the wrong frame convincingly.
PANE=ssf-merge.struktureretsundfornuft.dk
node test-kit/driver/cdp-eval.mjs $PANE "document.body.innerText"

# Text boxes need real typing, and so does the blank-cell <select>.
node test-kit/driver/cdp-type.mjs $PANE "document.querySelectorAll('input[type=text]')[0]" 2
node test-kit/driver/cdp-key.mjs $PANE "[...document.querySelectorAll('select')][0]" ArrowDown 2

# 9 = 3 template slides + 6 merged. Without it you can get last save, silently.
node test-kit/driver/fetch-deck.mjs test-kit/out/round.pptx --expect-slides 9
node test-kit/driver/verify-package.mjs test-kit/out/round.pptx
node test-kit/driver/mutate.mjs test-kit/out/round.pptx   # before believing it
```

`SSF_CDP` overrides the debugging endpoint, `SSF_PROFILE` the browser profile,
`SSF_CHANNEL` the browser channel, and `SSF_PAGE` picks a page by URL substring
where several are open.

## What will bite you

Each of these cost time on 2026-08-28. None is obvious from the outside.

**AppLocker blocks Playwright's bundled Chromium.** It lives under
`%LOCALAPPDATA%\ms-playwright`, and launching it gives `spawn UNKNOWN`; a direct
launch says *"This program is blocked by group policy."* Only Program Files is
allow-listed, so drive a system browser: `channel: "msedge"` or `"chrome"`.

**Playwright cannot reach PowerPoint's frames over CDP.** The editor and the
Office Add-ins dialog are out-of-process iframes, and `frame.evaluate` answers
*"Frame was detached"*. That is what `cdp-eval.mjs` exists for — each OOPIF is
its own CDP target with its own websocket, and talking to it directly works. The
task pane is a separate target and is easy to drive.

**The display scale squeezes the page.** On a high-DPI machine the CSS viewport
came back ~695px wide inside a 1298px window, and wide dialogs were clipped with
no scrollbar to reveal them. Launch with `--force-device-scale-factor=1`. A CDP
`Emulation` override does **not** survive, because it is cleared the moment the
session detaches.

**React ignores the native-setter trick in this pane.** Setting `.value` through
the prototype descriptor left Step 1's button disabled — the value changed and
the state did not. Real key events work, which is what `cdp-type.mjs` sends. Use
its `insert` mode for the pasted table: a typed Tab moves focus and a typed
Enter submits.

**"Upload My Add-in" is invisible to `innerText`.** It is a link in the DOM of
the Office Add-ins dialog. An agent reading the dialog's text concludes the
option is not there and gives up on sideloading. Query for the element.

The path to it, since it is four clicks and none of them obvious: **Home ▸
Add-ins** (a flyout, not a `[role=dialog]`, so a query for dialogs finds
nothing) **▸ More Add-ins** ▸ the **MY ADD-INS** tab ▸ the **Manage My Add-ins**
dropdown at the top right, which is where the link actually lives. The store
listing renders in its own `inclient.store.office.com` OOPIF, and the link is
NOT in it — searching that frame's DOM for "upload" returns nothing, which reads
exactly like the option being absent.

**The Store half of that dialog cannot authenticate on a personal account** —
*"The application is a first party application, the user does not have consent…"*.
Sideloading still works; that error is about the store listing.

**Reloading the task pane's iframe kills it.** It needs Office's handshake, which
a bare `location.reload()` does not redo. Reload the whole editor page and press
the ribbon button again.

**Clicking Download in OneDrive puts the file nowhere you can find it.** Use
`fetch-deck.mjs`, which fetches from inside the page and hands the bytes back
over CDP.

The five below were found on 2026-08-30.

**The `<select>` needs real keys, exactly like the text boxes did.** Setting
`.value` and dispatching `change` on the blank-cell control moved the DOM and
not React: the select read `skip` while the button still said "Add 6 slides".
That looks precisely like the defect the round was sent to look for — a forecast
that does not follow the control — and it is not one. `cdp-key.mjs` sends real
ArrowDown keys, after which the line becomes "2 of 3 rows × 2 slides" and the
button "Add 4 slides". **Never report a control as not responding until it has
been driven with real key events.**

**`fetch-deck.mjs` can race OneDrive's save.** PowerPoint for the web autosaves,
and the download endpoint serves what has been COMMITTED. A fetch straight after
a merge returned the pre-merge deck: 3 slides, 58 KB, HTTP 200, no error. Byte
count will not tell you — 58 KB looks like a plausible deck. Pass
`--expect-slides N` and let it wait.

**Playwright's `page.screenshot` can render the editor at the wrong size.** It
produced 695×705 captures of a page whose own `innerWidth` was 1918, so the
Office Add-ins dialog came out clipped and looked like the known display-scale
trap. It is not that: the window and the page were both fine. Use CDP
`Page.captureScreenshot` on the target for anything you intend to read.

**`web-login.mjs` used to watch for chrome that has moved.** `office.com` now
redirects to `m365.cloud.microsoft/chat`, a Copilot shell carrying none of the
old signed-in markers, so the watcher reported nothing on a session that was
fully signed in. It now also accepts a known signed-in host, and asks every open
tab rather than the first. If it still says nothing, open OneDrive and look
before concluding the login failed.

**Clear the run crumb between experiments.** The pane keeps an interrupted-run
record in `localStorage` under `ssf-merge.run.v1`. It is now keyed to the
document, so it will not follow you to another deck — but re-running the same
deck with a stale crumb still starts you in a recovery state. One line:
`cdp-eval.mjs <pane-host> "localStorage.removeItem('ssf-merge.run.v1')"`.

## The verifier is checked too

`verify-package.mjs` is an **alignment** checker, not a counter: every per-row
assertion is anchored to the slide's own title text, so a correctly-shaped deck
with the rows shuffled fails.

`mutate.mjs` is what makes it trustworthy. It breaks a reference deck six ways —
dangle a relationship, swap a chart's region, revert a workbook to the
placeholder, give two rows the same photo, blank the `{{Nickname}}` placeholder,
put the wrong name in a notes page — and asserts each is caught by its own guard.
Run it before believing a verdict.

Three bugs were found **in the verifier itself** this way, before any could be
reported as a defect in the product:

1. `_rels/.rels` resolves against the package root, not against `_rels/` — the
   first version reported four dangling relationships in a sound deck.
2. `<a:t xml:space="preserve"> 1 Mar</a:t>` carries attributes, and a pattern
   matching only the bare tag dropped exactly those runs — which made a
   correctly-merged date look like a lost one.
3. It assumed every slide maps to a data row, and crashed on a deck that still
   holds its template slides.
4. Having stopped crashing on those decks, it still MARKED THEM DOWN: every
   per-row check treated the template slides as failures, and the part counts
   charged the deck for the template's own chart, workbook and diagram. A round
   done exactly as the manual asks scored 5/13 while being entirely correct, and
   the report had to be re-derived from the XML by hand to get past it. It now
   splits merged from template once and reads 13/13 on that same deck.

A verifier that has only ever seen a good deck is an untested instrument.

**And a mutation harness that cannot fail is the same thing one level up.**
`mutate.mjs` used to ask only whether the expected check was red AFTER the
mutation. If it was red BEFORE, that is true no matter what the mutation did.
On the deck above, four of the six mutants asserted against an already-red
check — and it printed *"Every mutation was caught by its own guard"* and exited
0. It now reports those as `CANNOT PROVE`, counts how many guards were really
exercised, and refuses to certify the run. **Read the count, not the last
line:** `6/6 mutation(s) actually proved their guard` is the sentence that
means something.
