# Prompt for a local Claude Code session driving the test run

Paste the block below into Claude Code on the machine that has PowerPoint and a
browser. It assumes this repository is checked out there.

**Read the caveats under the prompt before you run it.** Playwright can drive
PowerPoint for the WEB; it cannot drive desktop PowerPoint, and the pane is a
cross-origin iframe. What the browser half is really for is producing the merged
deck; the verdict comes from the file.

**There is already a driver: `test-kit/driver/`.** Eleven scripts that sideload
the add-in, drive the pane, pull the merged deck out of the browser and check the
package — plus a README naming the traps that cost the last round hours, several
of which look like your own mistake when you hit them. Read it before writing
anything. Playwright cannot reach PowerPoint's frames at all, which is the first
thing you will otherwise rediscover.

---

## The prompt

```
You are running the SSF Merge real-host test round on my machine. The kit is in
this repo under test-kit/, and docs/TEST-KIT.md is the checklist a human would
follow. Read both before you start.

Goal: find out whether a deck this add-in merges — one containing a chart and a
SmartArt graphic — opens in real PowerPoint without a repair prompt, and whether
its charts, pictures and diagrams say the right thing.

WHAT IS ALREADY KNOWN, so you do not spend the round re-deriving it:
- The engine merges this exact template correctly when run headlessly. CI proves
  it every run: `npx vitest run test/test-kit.test.ts`. Run it first so you know
  the checkout is sound, then do not treat it as evidence about PowerPoint.
- Everything the suite can check about the merged package is already checked.
  The open question is entirely about the HOST: does PowerPoint accept the file,
  and does it draw what the bytes say.

PLAN, in this order. Stop and tell me if a step is impossible rather than
working around it.

1. `npx vitest run test/test-kit.test.ts`. If it fails, stop — the checkout is
   the problem, not the host.

2. Prepare the template. Open test-kit/SSF-Merge-test-template.pptx and add the
   SmartArt on slide 3 as docs/TEST-KIT.md describes (Insert > SmartArt >
   Process > Basic Process, three boxes reading `{{Name}}`, `{{Region}}` and
   `Renewal {{Renewal|date:d MMM}}`, then delete the grey box). If you cannot
   author SmartArt from a script — you probably cannot — say so and ask me to do
   this one step by hand. It is thirty seconds and it is the whole SmartArt test.

3. Drive the add-in. Two routes; take the first that works and say which.

   a) PowerPoint for the web, with Playwright. Upload the template to OneDrive,
      open it in the editor, sideload the add-in, and drive the pane: block 2 to
      3, paste test-kit/data.txt, choose the three PNGs in the picture picker,
      merge. The pane is a cross-origin iframe — expect to need frameLocator,
      and expect the file input to want setInputFiles on the iframe's own input.
      Save my session state to avoid logging in every run.
   b) Desktop PowerPoint, by hand. If the web route costs more than about thirty
      minutes, stop and hand me a numbered click-list instead. A round done by a
      human today is worth more than an automation that works next week.

4. Capture the evidence. In BOTH routes the deliverable is the merged file, not
   a screenshot:
   - download or save the merged deck to test-kit/out/round-<date>.pptx
   - screenshot each merged slide (in the browser, or File > Export in desktop)
   - copy the pane's own summary line verbatim

5. Verify the merged deck as a package. Do not eyeball the XML.

   `node test-kit/driver/verify-package.mjs <deck.pptx>` already does this and
   prints a table. Run `node test-kit/driver/mutate.mjs` first if you intend to
   trust its verdict — it breaks a reference deck six ways and checks each is
   caught, and it has found three bugs in the verifier itself. Extend it rather
   than starting over. What it asserts:
   - every relationship target resolves to a part that exists
   - three chart parts, three embedded workbooks, three media parts
   - each chart's title and `<c:strCache>` hold a different region, none holds
     `{{Region}}`
   - each workbook's sharedStrings holds that slide's region
   - the SmartArt drawing part (`ppt/diagrams/drawingN.xml`, NOT just dataN)
     holds the row's name — this is the half PowerPoint puts on screen
   - `{{Nickname}}` is still present, because there is no such column
   - the notes pages read "Call Ada before 1 Mar" and so on
   test/test-kit.test.ts does most of this against a headless merge; you are
   doing it against what the HOST produced, which may differ.

6. The three checks a script cannot make. Do these in the app and report what
   you saw:
   - close and reopen the merged deck. Any repair prompt is the headline result
     of the round; quote it exactly.
   - right-click a chart > Edit Data, wait for Excel, close it, and look at the
     chart again. It must still show the merged region. PowerPoint refreshes the
     chart from that workbook when Excel closes, so this is where a half-merge
     shows itself.
   - look at the pictures. The frame is wider than it is tall, so ada.png (wide)
     should keep all four yellow corner dots, grace.png (tall) should lose the
     top and bottom pair. Circles must still be circles — ovals mean it
     stretched, background bands mean it letterboxed.

7. Write test-kit/out/round-<date>.md: what you did, what each check said, the
   pane's summary line, and anything that surprised you. Attach the merged deck
   and the screenshots. Then give me the short version in chat.

RULES:
- Report what happened, including "I could not do step 3a". A round that says
  "the web route defeated me, here is the manual click-list" is a good round.
- Do not change any code in src/ to make the round pass. If you find a bug,
  write it up in the round file and stop.
- Do not commit test-kit/out/ — it is evidence from one machine on one day.
- The deck is throwaway, but the OneDrive account is not: work in a folder you
  are happy to delete, and remember PowerPoint for the web autosaves everything
  as it happens.
```

---

## What to expect, honestly

**Playwright drives the web host, not the desktop one.** If the round you care
about is desktop PowerPoint — and it is the one with the healthier Office.js —
then the browser automation is the wrong tool and the click-list is the right
one. The web host is worth testing too; it is just a different host, and this
project's notes are full of ways the two differ.

**The task pane is a cross-origin iframe.** Playwright can reach into it with
`frameLocator`, and the file picker needs `setInputFiles` against the input
inside that frame. Both work; neither is quick to get right the first time.

**Sideloading is the fiddly part.** The add-in is hosted at
`ssf-merge.struktureretsundfornuft.dk`, and getting the manifest into a
web session by script is more awkward than adding it once by hand and reusing
a saved storage state. Let the agent add it manually on the first run and save
the session.

The route that works on PowerPoint for the web is **Add-ins ▸ My Add-ins ▸
Upload My Add-in**, pointed at `manifest-prod.xml`. Two things about it are
worth knowing before going looking, both found on 2026-08-28:

- On a personal Microsoft account the **Store** half of that dialog cannot
  authenticate. It fails with *"The application is a first party application,
  the user does not have consent, and users are not permitted to consent to
  first party applications."* The upload half still works — the consent error
  is about the store listing, not about sideloading.
- "Upload My Add-in" does not show up in the dialog's `innerText`, so an agent
  reading the text will conclude it is absent. It is a link in the DOM. Look
  for the element, not for the words.

Once it is in, the add-in's button is called **Mail merge** and sits on the
**Home** tab rather than on a tab of its own.

**The merged file is the evidence.** Screenshots show whether something looks
right; the package shows whether it IS right, and the two disagree in exactly
the cases worth finding — a chart that reads correctly until Excel touches it,
a SmartArt whose model merged and whose rendering did not.
