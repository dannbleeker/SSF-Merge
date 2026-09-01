/**
 * A PowerPoint host, faked far enough to drive `src/office/powerpoint.ts`.
 *
 * That module is the one place in this repo that calls Office.js, and for most
 * of its life nothing in the suite could run a line of it: the tests above it
 * mock `insertDeck` and `undoInsert` wholesale, so what those two functions
 * actually do with a host's answer was checked by nothing. The defects that
 * lived there were real — an uncapped error string that put the merged package
 * on screen, and an undo that reported a 0-based index as a slide number —
 * and both are invisible to a mock of the function that holds them.
 *
 * It is deliberately small. Everything the module touches and nothing else:
 * `Office.context.requirements.isSetSupported`, `PowerPoint.run` with a
 * `presentation.slides` collection that answers `getCount`, `getItemAt`,
 * `exportAsBase64Presentation`, `insertSlidesFromBase64` and
 * `getSelectedSlides`. A fuller fake would be a second implementation of
 * PowerPoint to keep in step, which is how a fake starts lying.
 *
 * The `pending` list is what makes it behave like the real thing in the way
 * that matters: a property loaded in a batch is NOT readable until the batch's
 * `sync()`, so code that reads before syncing gets `undefined` here exactly as
 * it would in the host.
 */
export interface FakeOptions {
  /** The deck's slide ids, in order. Defaults to a three-slide deck. */
  slides?: string[];
  /** What each `context.sync()` costs, for the budget tests. */
  syncMs?: number;
  /** Requirement sets the host claims. Defaults to everything up to 1.10. */
  sets?: string[];
  /** What `getSelectedSlides` answers. */
  selected?: string[];
  /** Called when the deck is inserted, so a test can raise or grow the deck. */
  onInsert?: (base64: string, options: unknown) => void;
  /** What an export answers. */
  exportBytes?: string;
}

export interface FakeHost {
  slides: string[];
  syncs: number;
  inserted: { base64: string; options: unknown }[];
  deleted: number[];
}

export function installFakeHost(options: FakeOptions = {}): FakeHost {
  const state: FakeHost = {
    slides: [...(options.slides ?? ["s1", "s2", "s3"])],
    syncs: 0,
    inserted: [],
    deleted: [],
  };
  const sets = options.sets ?? ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10"];
  const syncMs = options.syncMs ?? 0;

  (globalThis as any).Office = {
    context: {
      requirements: { isSetSupported: (_name: string, version: string) => sets.includes(version) },
      document: { url: "fake.pptx" },
      platform: "OfficeOnline",
      diagnostics: { version: "16.0.fake", host: "PowerPoint" },
    },
    FileType: { Compressed: "compressed" },
    AsyncResultStatus: { Succeeded: "succeeded", Failed: "failed" },
    CoercionType: { Text: "text" },
  };

  (globalThis as any).PowerPoint = {
    run: async (body: (context: any) => Promise<unknown>) => {
      // Filled by each queued call and drained by `sync`, which is what makes a
      // read-before-sync answer `undefined` the way the host does.
      const pending: (() => void)[] = [];
      const removals = new Set<number>();
      const context: any = {
        sync: async () => {
          state.syncs++;
          if (syncMs > 0) await new Promise((resolve) => setTimeout(resolve, syncMs));
          for (const settle of pending.splice(0)) settle();
          // Highest index first, so removing one cannot shift another.
          for (const index of [...removals].sort((a, b) => b - a)) state.slides.splice(index, 1);
          removals.clear();
        },
        presentation: {
          slides: {
            getCount: (): { value?: number } => {
              const box: { value?: number } = {};
              pending.push(() => (box.value = state.slides.length));
              return box;
            },
            getItemAt: (index: number): Record<string, unknown> => {
              const slide: Record<string, unknown> = {
                load: () => undefined,
                // Queued, then applied at `sync`, because that is when a real
                // host performs it — and because the caller deletes several in
                // one batch, so removing one immediately would shift the
                // indexes of the others mid-batch.
                delete: () => {
                  state.deleted.push(index);
                  pending.push(() => removals.add(index));
                },
                tags: {
                  getItemOrNullObject: (): Record<string, unknown> => {
                    const tag: Record<string, unknown> = { load: () => undefined };
                    pending.push(() => (tag.isNullObject = true));
                    return tag;
                  },
                },
              };
              pending.push(() => (slide.id = state.slides[index]));
              return slide;
            },
            exportAsBase64Presentation: (): { value?: string } => {
              const box: { value?: string } = {};
              pending.push(() => (box.value = options.exportBytes ?? "EXPORTED"));
              return box;
            },
          },
          insertSlidesFromBase64: (base64: string, opts: unknown) => {
            state.inserted.push({ base64, options: opts });
            options.onInsert?.(base64, opts);
          },
          getSelectedSlides: (): Record<string, unknown> => {
            const collection: Record<string, unknown> = { load: () => undefined, items: [] };
            pending.push(() => (collection.items = (options.selected ?? []).map((id) => ({ id }))));
            return collection;
          },
        },
      };
      return body(context);
    },
  };
  return state;
}

export function uninstallFakeHost(): void {
  delete (globalThis as any).Office;
  delete (globalThis as any).PowerPoint;
}
