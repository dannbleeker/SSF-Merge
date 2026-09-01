// @vitest-environment jsdom
/**
 * THROWAWAY review harness. Drives the real pane against a simulated deck with
 * the REAL sweepPlan/provenSweep/undoInsert. Delete after the review.
 */
import { beforeEach, describe, expect, it, vi } from "vitest";

const office = vi.hoisted(() => ({
  inspectBlock: vi.fn<(r: { from: number; to: number }) => Promise<unknown>>(),
  runMerge: vi.fn<(r: unknown) => Promise<unknown>>(),
}));

vi.mock("../src/office/merge.js", async (orig) => {
  const real = (await orig()) as Record<string, unknown>;
  return { ...real, inspectBlock: office.inspectBlock, runMerge: office.runMerge };
});

/** The simulated deck: one entry per slide, carrying the run tag it has. */
type Slide = { tag?: string };
let deck: Slide[] = [];
let tagsReadable = true;
let refuseTagRead: (i: number) => boolean = () => false;

function installHost(): void {
  const queue: (() => void)[] = [];
  const ctx = {
    presentation: {
      slides: {
        getCount() {
          const box = { value: -1 };
          queue.push(() => (box.value = deck.length));
          return box;
        },
        getItemAt(i: number) {
          return {
            load() {},
            delete() {
              queue.push(() => deck.splice(i, 1));
            },
            tags: {
              getItemOrNullObject(_key: string) {
                const box: { value?: string; isNullObject: boolean; load: (s: string) => void } = {
                  isNullObject: true,
                  load: () => {},
                };
                queue.push(() => {
                  const slide = deck[i];
                  const hidden = !tagsReadable || refuseTagRead(i);
                  if (!slide || hidden || slide.tag === undefined) {
                    box.isNullObject = true;
                  } else {
                    box.isNullObject = false;
                    box.value = slide.tag;
                  }
                });
                return box;
              },
            },
          };
        },
      },
    },
    async sync() {
      while (queue.length) (queue.shift() as () => void)();
    },
  };
  (globalThis as unknown as { PowerPoint: unknown }).PowerPoint = {
    run: async (cb: (c: typeof ctx) => unknown) => cb(ctx),
  };
  (globalThis as unknown as { Office: unknown }).Office = {
    onReady: (cb: () => void) => {
      onReady = cb;
      return Promise.resolve();
    },
    context: {
      document: { url: "https://example-my.sharepoint.com/x/deck.pptx" },
      requirements: { isSetSupported: (_s: string, v: string) => Number(v) <= 1.3 },
    },
  };
}

let onReady: () => void;

const REPORT = { ok: true, detail: "2 placeholders in slides 4 to 6.", fields: ["First", "Last"] };

async function openPane(): Promise<void> {
  document.body.innerHTML = '<header><b>SSF</b><span>Merge</span></header><div id="pane"></div>';
  vi.resetModules();
  installHost();
  await import("../src/pane/main.js");
  onReady();
  await settle();
}

function pane(): HTMLElement {
  return document.getElementById("pane") as HTMLElement;
}
function primary(): HTMLButtonElement {
  return pane().querySelector("button.primary") as HTMLButtonElement;
}
function byAction(a: string): HTMLButtonElement | null {
  return pane().querySelector(`[data-action="${a}"]`);
}
function field(name: string): HTMLInputElement {
  return pane().querySelector(`[data-field="${name}"]`) as HTMLInputElement;
}
function type(name: string, value: string): void {
  const node = field(name);
  node.value = value;
  node.dispatchEvent(new Event("input", { bubbles: true }));
}
function said(): string[] {
  return Array.from(pane().querySelectorAll(".blocked, .notice, .card")).map((n) => (n.textContent ?? "").trim());
}
async function settle(): Promise<void> {
  for (let i = 0; i < 40; i++) await Promise.resolve();
}

/** Walk to the preview step with a block and data in hand. */
async function reachPreview(): Promise<void> {
  await openPane();
  type("from", "4");
  type("to", "6");
  office.inspectBlock.mockResolvedValueOnce(REPORT);
  primary().click();
  await settle();
  type("paste", "First\tLast\nAda\tLovelace\nGrace\tHopper");
  primary().click();
  await settle();
  office.inspectBlock.mockResolvedValueOnce(REPORT);
  primary().click();
  await settle();
}

beforeEach(() => {
  localStorage.clear();
  deck = Array.from({ length: 12 }, () => ({}));
  tagsReadable = true;
  refuseTagRead = () => false;
  office.inspectBlock.mockReset();
  office.runMerge.mockReset();
});

describe("harness sanity", () => {
  it("reaches the preview step and runs a preview against the simulated deck", async () => {
    await reachPreview();
    // The preview inserts 3 slides tagged r1 onto the deck of 12.
    office.runMerge.mockImplementationOnce(async () => {
      for (let i = 0; i < 3; i++) deck.push({ tag: "r1" });
      return {
        ok: true,
        detail: "3 slides added after slide 12.",
        added: 3,
        accountable: true,
        deckAtStart: 12,
        landedAfter: 12,
        runId: "r1",
        fields: ["First", "Last"],
        imageFields: [],
        slideFields: [[], [], []],
        unknownConditions: [],
      };
    });
    primary().click();
    await settle();
    expect(deck.length).toBe(15);
    console.log("AFTER PREVIEW:", JSON.stringify(said(), null, 1));
    expect(byAction("end-preview")).not.toBeNull();
  });
});
