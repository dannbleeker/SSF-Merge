/**
 * A file with its PROSE removed, so a guard reads what the file DOES.
 *
 * Three guards in this repo have gone wrong the same way, each in a different
 * syntax, each found only when it went red on a file that was perfectly correct
 * or green on one that was not:
 *
 * - `test/architecture.test.ts` forbade Office.js in the engine, and matched the
 *   word "Office.js" in the paragraphs explaining WHY the engine avoids it —
 *   four correct files, red.
 * - `scripts/manifest-rules.mjs` forbade a `<Requirements>` block, and matched
 *   the XML comment explaining why the manifest has none — so the generator
 *   refused to write a file that was exactly right.
 * - `test/release.test.ts` checked that the pre-flight runs before the tag is
 *   created, and matched the YAML comment mentioning `gh release create` — so
 *   it compared the comment's position with the check's and reported the order
 *   backwards.
 *
 * A file that explains itself is not a defect; a guard that cannot tell an
 * explanation from an instruction is. Three strippers in three files is three
 * chances to write a fourth, so they are one module — and the next person
 * reaching for one finds it here rather than inventing it.
 *
 * None of these is a PARSER. Each is the smallest thing that makes its own
 * guard honest, and each is checked by `test/without-prose.test.ts` against the
 * case that caught it.
 */

/** `<!-- … -->`. */
export function withoutXmlComments(text) {
  return text.replace(/<!--[\s\S]*?-->/g, "");
}

/** `#` to end of line, for YAML and anything shell-shaped. */
export function withoutHashComments(text) {
  return text
    .split("\n")
    .filter((line) => !/^\s*#/.test(line))
    .join("\n");
}

/**
 * Block comments and comment-only lines, for TypeScript and JavaScript alike.
 *
 * The literals STAY, which is the whole reason this is separate from
 * `withoutTsProse`. A guard over code that builds its output with template
 * literals — `scripts/read-answers.mjs` reads half its sheet inside one — would
 * otherwise have every one of those reads stripped out from under it and report
 * the file as not doing what it plainly does. Reach for this when the thing
 * being checked can legitimately live in a string, and for `withoutTsProse`
 * when a name in a string is prose about a thing rather than a use of it.
 */
export function withoutTsComments(text) {
  return text
    .replace(/\/\*[\s\S]*?\*\//g, "")
    .split("\n")
    .filter((line) => !/^\s*(\/\/|\*)/.test(line))
    .join("\n");
}

/**
 * Block comments, comment-only lines, and STRING LITERALS, for TypeScript.
 *
 * The literals go too, and for the same reason the comments do: a verdict that
 * names `office-js#6105` in its text is a sentence about an issue, not a
 * dependency on one. An import specifier is a string literal as well, so a
 * guard that reads IMPORTS must read the raw source instead — `architecture.ts`
 * does exactly that, deliberately, and says so.
 */
export function withoutTsProse(text) {
  return withoutTsComments(text)
    .replace(/`(?:[^`\\]|\\.)*`/g, '""')
    .replace(/"(?:[^"\\\n]|\\.)*"/g, '""')
    .replace(/'(?:[^'\\\n]|\\.)*'/g, '""');
}
