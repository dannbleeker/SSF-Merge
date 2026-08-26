/**
 * Turn a record into the answer a placeholder gets.
 *
 * Separated from the replacement itself so the policy for an empty cell lives
 * in one place and can be tested without any XML.
 */
import { applyFormat } from "../data/format.js";
import type { Resolve } from "./text.js";

export type EmptyPolicy = "blank" | "keep" | "skip";

export interface ResolveOptions {
  /**
   * What an empty cell produces. `"blank"` writes nothing, `"keep"` leaves the
   * placeholder visible. `"skip"` is a planning decision rather than a
   * replacement one: by the time a record reaches here it has been kept, so it
   * behaves as `"blank"`.
   */
  onEmpty?: EmptyPolicy;
}

/**
 * A resolver over one record.
 *
 * A field the record has no column for answers null, which leaves the
 * placeholder on the slide. That is deliberate and it is the same decision the
 * replacement makes: a blank space looks finished, and `{{Territory}}` does
 * not, so the author sees their own mistake rather than shipping 240 slides
 * with a hole in them.
 */
export function makeResolver(row: Record<string, string>, opts: ResolveOptions = {}): Resolve {
  const onEmpty = opts.onEmpty ?? "blank";
  return (name, format) => {
    if (!Object.prototype.hasOwnProperty.call(row, name)) return null;
    const raw = row[name] ?? "";
    if (raw.trim() === "") return onEmpty === "keep" ? null : "";
    return applyFormat(raw, format);
  };
}
