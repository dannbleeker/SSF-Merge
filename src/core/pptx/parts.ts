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
 */
export const OWNED_BY_SLIDE = /^ppt\/(charts\/chart|charts\/chartEx|diagrams\/data|tags\/tag)\d+\.xml$/;

/**
 * What a chart or a SmartArt may drag out of the package with it.
 *
 * Its own styling and colours, the workbook behind it, its diagram parts, its
 * pictures, a theme override. That is the whole list, and it is an allowlist
 * for the reason above.
 */
export const OWNABLE_BY_GRAPHIC = /^ppt\/(charts|diagrams|embeddings|media|theme)\//;
