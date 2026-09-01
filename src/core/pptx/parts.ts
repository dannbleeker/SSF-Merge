/**
 * What a slide's relationships are CALLED, in one place.
 *
 * These strings decide which parts a clone copies, which a clone drops, and
 * which a removal deletes. They were written out in four files: `NOTES_REL_TYPE`
 * in `pkg.ts` and again in `clone.ts`, `COMMENT_REL_TYPES` in both, and
 * `TAGS_REL_TYPE` in `clone.ts` and again in `tags.ts`.
 *
 * The copies agreed, and nothing had gone wrong. What made it worth ending is
 * which decisions they drive: one copy of the comment list says what a CLONE
 * drops, the other says what a REMOVAL deletes. PowerPoint has already added a
 * second spelling of comments once — the modern web one under a Microsoft
 * namespace — and adding a third to one copy and not the other leaves a clone
 * carrying a comment part that the removal will not clean up, or a removal
 * deleting one a surviving slide still points at.
 */
const REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const MS_REL = "http://schemas.microsoft.com/office/2007/relationships";
const MS_REL_2014 = "http://schemas.microsoft.com/office/2014/relationships";

export const REL_TYPE = {
  slide: `${REL}/slide`,
  notesSlide: `${REL}/notesSlide`,
  tags: `${REL}/tags`,
  image: `${REL}/image`,
  chart: `${REL}/chart`,
  /**
   * A callout, arrow or text box drawn ON a chart, in its own drawing part.
   *
   * The chart owns the relationship, not the slide, so a merged copy of the
   * chart pointed at the template's one copy of the drawing — and the text pass
   * never saw it. A `{{Name}}` in a chart callout shipped verbatim on every
   * slide of the merge.
   */
  chartUserShapes: `${REL}/chartUserShapes`,
  /**
   * A MODERN chart — waterfall, funnel, treemap, sunburst, histogram, pareto,
   * box-and-whisker, region map. PowerPoint stores none of those as a
   * `<c:chartSpace>`; they are a separate part under a Microsoft namespace.
   *
   * The RELATIONSHIP is the stable thing to match on. On the slide these charts
   * sit inside `<mc:AlternateContent>`, whose `Requires` token is `cx1`, `cx2`
   * or `cx4` depending on which of three dated namespaces the layout came from
   * — a reader keying on the token misses the other two.
   */
  chartEx: `${MS_REL_2014}/chartEx`,
  diagramData: `${REL}/diagramData`,
  diagramDrawing: `${MS_REL}/diagramDrawing`,
  /** A whole package inside the package: the workbook behind a chart. */
  package: `${REL}/package`,
} as const;

/**
 * A slide's comments, in both spellings.
 *
 * The classic one is `ppt/comments/commentN.xml`; PowerPoint on the web writes
 * MODERN comments, `ppt/comments/modernComment_<id>_<hash>.xml`, under a
 * Microsoft-namespaced relationship. Both hang off the SLIDE.
 */
export const COMMENT_REL_TYPES = [
  `${REL}/comments`,
  "http://schemas.microsoft.com/office/2018/10/relationships/comments",
];

/**
 * The parts a slide OWNS: gone when it goes, if nothing else points at them.
 *
 * Anchored names rather than folders, and that is the security half. A
 * relationship target comes out of the deck and a deck can be sent to somebody:
 * a crafted one naming `/ppt/presentation.xml` put that part into the owned set
 * once, nothing else referred to it, and the sweep deleted it — producing a file
 * PowerPoint cannot open, from a merge that reported success.
 *
 * Tags are here because they were not, and every removed slide left one behind.
 *
 * Media is here for the same reason, and it arrived with modern charts. A
 * modern chart carries a rendered PICTURE for hosts too old to draw it, each
 * merged copy replaces that picture with a notice and stops relating to it, and
 * the template's own copy then goes out with the template slide — leaving the
 * bytes in the package with nothing pointing at them. A rendering of the
 * TEMPLATE's data, shipped in the file that gets sent out.
 *
 * A picture is far likelier than a chart to be SHARED — a logo on the master, a
 * photo used twice — and the referrer scan is what makes that safe: a part any
 * other part still names is left exactly where it is. The scan is why this can
 * be a rule about media in general rather than a special case for a fallback
 * picture, which would need the replacement to remember what it dropped and
 * would miss every other picture a removed slide was the last owner of.
 *
 * One limit, and it is the conservative direction. The referrer scan runs per
 * removal and counts every part that still exists, so a picture two TEMPLATE
 * slides share survives both of them: each removal sees the other slide still
 * pointing at it. That leaves bytes behind rather than taking a part something
 * might still need, and a modern chart's fallback picture is its own — one per
 * chart, per slide — so the case this arrived for is covered.
 *
 * Not anchored to a NUMBER, because media names are not numbered by anything:
 * PowerPoint writes `image7.emf`, Excel writes `Microsoft_Excel_Worksheet.xlsx`,
 * and another tool writes whatever it likes. `[^/]+` keeps it to one segment
 * directly under `ppt/media/`, so a target reaching upwards resolves outside the
 * pattern and is not a candidate.
 */
export const OWNED_BY_SLIDE =
  /^ppt\/(?:(?:charts\/chart|charts\/chartEx|diagrams\/data|tags\/tag)\d+\.xml|media\/[^/]+)$/;

/**
 * What a chart or a SmartArt may drag out of the package with it.
 *
 * Its own styling and colours, the workbook behind it, its diagram parts, its
 * pictures, a theme override. That is the whole list, and it is an allowlist
 * for the reason above.
 */
export const OWNABLE_BY_GRAPHIC = /^ppt\/(charts|diagrams|drawings|embeddings|media|theme)\//;
