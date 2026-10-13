import type { Inline } from "./docx.js";

/**
 * # What a rebuilt run still knows about itself
 *
 * The page is a rendering of the *resolved* run style, and the editor re-derives the model from it on every
 * keystroke. A view can say how a run looks; it cannot say which of those properties the run itself stated, and it
 * holds no copy of the parts of `w:rPr` the model only carries. Read back on its own, that rebuild leaves the runs
 * of a document claiming their paragraph's bold and the file's floor size, having dropped the `w:kern`, the
 * `w:lang` and the drawing's frame that came in with them.
 *
 * So a rebuilt run is treated as the run it came from, found through the text that survived: `keepRunIdentity` hands
 * back the fields the view cannot express, and the write side's own guard (`asRead`) does the rest — a property the
 * run already showed before the edit is not a property the edit stated.
 */

/** A mark (image, break, tab, anchor, carried structure) is a paragraph's content that is not text: it keeps a
 *  paragraph from being empty, and it stands for one character — which is how a caret offset is counted too.
 *  Document text never holds NUL, so `flatOf` can use one as the placeholder. */
export const isMark = (inline: Inline): boolean =>
  !!inline.image || !!inline.break || !!inline.tab || !!inline.anchor || !!inline.carry;

/** A mark (image, break, tab, anchor, carried structure) stands for one character, which is how a caret offset is
 *  counted too. Document text never holds NUL, so `flatOf` can use one as the placeholder. */
export const inlineLength = (inline: Inline): number => (inline.text !== "" ? inline.text.length : 1);

/** The run list as one comparable string: text as it is, one NUL per mark. */
export const flatOf = (inlines: Inline[]): string =>
  inlines.map((i) => (i.text !== "" ? i.text : "\u0000")).join("");

/**
 * The run fields the page can tell us about. `styleOf` reads the first group off an element's tags and inline CSS,
 * `inlinesOf` reads the rest off the attributes the renderer stamps on the element. Everything outside this list a
 * rebuilt run knows only from the run it came in with: the two faces of `w:rFonts`, the `w:br` and the `w:drawing`
 * as the file wrote them, the `w:rPr` children the model carries as bytes, and the read snapshot that says which of
 * the properties on screen the run stated itself.
 */
export const SEEN_BY_VIEW: readonly (keyof Inline)[] = [
  "text",
  "bold",
  "italic",
  "underline",
  "strike",
  "sizePt",
  "color",
  "highlight",
  "vertAlign",
  "hyperlink",
  "linkAttrs",
  "image",
  "break",
  "tab",
  "anchor",
  "carry",
  "rev",
  "revFormat",
];

/** One element's bytes, not a formatting statement: two runs holding the same drawing would be the same picture
 *  written twice, so only the run that came from it gets it. */
const ONCE_ONLY: readonly (keyof Inline)[] = ["drawing"];

/** The run a character position falls inside. The last run covers the end of the paragraph, where a caret after
 *  the final character sits. */
function runAt(inlines: Inline[], pos: number): Inline | undefined {
  let seen = 0;
  for (const inline of inlines) {
    if (pos < seen + inlineLength(inline)) return inline;
    seen += inlineLength(inline);
  }
  return inlines[inlines.length - 1];
}

/**
 * Gives each rebuilt run the fields the view cannot carry, from the run its text came from.
 *
 * Typing moves text, so the two lists are matched by what did not move: everything up to the first difference
 * lines up one character at a time, everything after the last difference lines up counted back from the end, and
 * the run the difference sits inside is the one the changed text belongs to. A rebuilt run that never had an
 * origin — a paragraph the browser created, with nothing before it — keeps only what the page says.
 *
 * Mutates `next`, which is the list the caller just built from the DOM.
 */
export function keepRunIdentity(next: Inline[], prev: Inline[]): void {
  if (!prev.length || !next.length) return;
  const before = flatOf(prev);
  const after = flatOf(next);
  let head = 0;
  while (head < before.length && head < after.length && before[head] === after[head]) head += 1;
  let tail = 0;
  while (
    tail < before.length - head &&
    tail < after.length - head &&
    before[before.length - 1 - tail] === after[after.length - 1 - tail]
  )
    tail += 1;
  const origin = (pos: number): Inline | undefined => {
    // Inside the changed region, `head` is where the two texts stopped agreeing: the run the change was made in.
    const at = pos < head ? pos : pos >= after.length - tail ? pos + before.length - after.length : Math.min(head, before.length - 1);
    return runAt(prev, at);
  };
  const given = new Set<Inline>();
  let pos = 0;
  for (const run of next) {
    // A carried structure and a comment anchor come back as the objects they were, found by key: they are their
    // own origin, and what they hold is an element's bytes rather than a style the run states.
    const from = run.carry || run.anchor ? undefined : origin(pos);
    if (from && from !== run) {
      for (const key of Object.keys(from) as (keyof Inline)[]) {
        // The page can show a property that is *on*; it has nothing to show for one the file turned off. So a
        // rebuilt run that says nothing is not a contradiction of the origin's `false` — that statement lives
        // only in `w:rPr` (`<w:u w:val="none"/>`, `<w:i w:val="0"/>`, `<w:strike w:val="0"/>`, and for
        // `w:vertAlign` the enum name `baseline`), and letting the silence stand would turn an explicit
        // cancellation back into "the run never mentioned it", which the writer then deletes.
        const cancelled = run[key] === undefined && (from[key] === false || from[key] === "baseline");
        if (!cancelled && (SEEN_BY_VIEW.includes(key) || run[key] !== undefined)) continue;
        if (ONCE_ONLY.includes(key) && given.has(from)) continue;
        (run as unknown as Record<string, unknown>)[key] = from[key];
        if (ONCE_ONLY.includes(key)) given.add(from);
      }
    }
    pos += inlineLength(run);
  }
}
