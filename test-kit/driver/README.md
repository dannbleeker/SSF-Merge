# Driving the round in PowerPoint for the web

Eleven small scripts that put the add-in through the kit's run in a real host,
and check what came back. Built during the round of 2026-08-28, which found two
defects with them ([#64](https://github.com/dannbleeker/SSF-Merge/pull/64),
[#66](https://github.com/dannbleeker/SSF-Merge/pull/66)).

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
| `cdp-setfile.mjs` | Puts local files on an `<input type=file>`, including one inside a nested frame. |
| `close-popup.mjs` | Closes pages whose URL matches a substring. |
| `upload-template.mjs` | Uploads a local file into the OneDrive folder currently on screen. |
| `fetch-deck.mjs` | Downloads the open deck's bytes through the page's own session. |
| `verify-package.mjs` | Thirteen checks over a merged deck, anchored per row to the slide's own title. |
| `mutate.mjs` | Breaks a reference deck six ways and asserts each is caught by its own guard. |

## The run, roughly

```bash
node test-kit/driver/web-login.mjs                 # once; sign in by hand
# then launch Edge yourself with --remote-debugging-port=9333 on the same
# profile, and drive it:
node test-kit/driver/drive.mjs url
node test-kit/driver/upload-template.mjs test-kit/SSF-Merge-test-template.pptx
node test-kit/driver/cdp-eval.mjs list
node test-kit/driver/fetch-deck.mjs test-kit/out/round.pptx
node test-kit/driver/verify-package.mjs test-kit/out/round.pptx
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

**The Store half of that dialog cannot authenticate on a personal account** —
*"The application is a first party application, the user does not have consent…"*.
Sideloading still works; that error is about the store listing.

**Reloading the task pane's iframe kills it.** It needs Office's handshake, which
a bare `location.reload()` does not redo. Reload the whole editor page and press
the ribbon button again.

**Clicking Download in OneDrive puts the file nowhere you can find it.** Use
`fetch-deck.mjs`, which fetches from inside the page and hands the bytes back
over CDP.

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

A verifier that has only ever seen a good deck is an untested instrument.
