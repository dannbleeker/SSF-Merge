/**
 * Turn a record into the answer a placeholder gets.
 *
 * Separated from the replacement itself so the policy for an empty cell lives
 * in one place and can be tested without any XML.
 */
import { applyFormat } from "../data/format.js";
import { imageMode } from "./images.js";
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
    // An IMAGE field is not text and must never be written as any. The picture
    // pass has already run and either placed it or left the placeholder alone;
    // answering here would print the FILE NAME over the frame that was supposed
    // to hold the photo — which is exactly what happened before this line, and
    // reads as the merge putting data in the wrong place rather than as a
    // missing file.
    //
    // Null is the right answer for BOTH outcomes. Placed, the placeholder text
    // is already gone and there is nothing left to resolve; not placed, the
    // rule is the same one a field with no column follows — stay visible, so
    // the author sees their own gap rather than 240 finished-looking blanks.
    if (imageMode(format)) return null;
    if (!Object.prototype.hasOwnProperty.call(row, name)) return null;
    const raw = row[name] ?? "";
    if (raw.trim() === "") return onEmpty === "keep" ? null : "";
    return foldCellBreaks(applyFormat(raw, format));
  };
}

/**
 * A line break the author typed INSIDE a cell, folded to a single space.
 *
 * `CHAR(11)` is already handled, by `XML_FORBIDDEN` in `text.ts`, and for a
 * different reason: XML cannot carry it at all, so it produced a file
 * PowerPoint refused to open. These are the other spellings of the same
 * keystroke and they are all LEGAL XML, so nothing stopped them — Excel's
 * Alt+Enter puts `CHAR(10)` in the cell and a bare LF on the clipboard, and a
 * `<textarea>` normalises what it is given to CRLF before the pane ever reads
 * it.
 *
 * Legal is not harmless. DrawingML renders a literal newline inside `<a:t>` as
 * a HARD line break, so the merge of 2026-08-31 put `"Ada⏎Lovelace"` on a slide
 * and got a title reading "Ada" on one line and "Lovelace — Nordics" on the
 * next — and the same break in the notes page and in both halves of the
 * SmartArt. The deck opened clean, which is why this survived the sweep that
 * found `CHAR(11)`: there was no repair prompt to notice.
 *
 * A space, for the reason `XML_FORBIDDEN` already gives: dropping it joins two
 * words that were separate. CRLF is ONE break and folds to one space, which is
 * why the pair is matched before the singles.
 *
 * Here rather than in `text.ts` because it is a fact about a CELL, not about a
 * text run: the value reaches a chart's cached labels, an embedded workbook and
 * a SmartArt's model through this same function, and each of those would
 * otherwise need its own copy of the rule.
 */
const CELL_LINE_BREAK = /\r\n|[\r\n]/g;

function foldCellBreaks(text: string): string {
  return text.replace(CELL_LINE_BREAK, " ");
}
