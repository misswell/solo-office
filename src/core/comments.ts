import type { Inline } from "./docx.js";

/**
 * Comment evidence for this module comes from `.scratch/rea-notes.md` §9: WPS reads
 * `word/comments.xml`, `commentsExtended.xml` and `people.xml` as three separate parts
 * (`WmlCommentsPart::Read` / `WmlCommentsExtendedPart::Read` / `WmlPeoplePart::Read`),
 * and the body carries the anchor as first-class in-stream nodes
 * (`AddCommentRangeStart` / `AddCommentRangeEnd` / `AddCommentEx`).
 *
 * The extension-part names, content types, relationship types and attribute prefixes were
 * read off an independent OOXML implementation installed on this machine (ONLYOFFICE
 * `sdkjs/word/sdk-all.js`, read-only forensics — see `.scratch/doc-comment-evidence-20261011.md`),
 * because WPS keeps OOXML names in a token table and its binaries expose no literals.
 */

/** Where an anchored range begins, ends, or carries Word's reference mark. */
export interface CommentAnchor {
  id: number;
  kind: "start" | "end" | "ref";
}

/**
 * One `w:comment` from `word/comments.xml`.
 *
 * `raw` holds the element exactly as read, so a comment nobody touched is written back
 * byte-for-byte with its rich text, its `w14:paraId` links and any markup this model
 * does not understand. Only a comment whose text or formatting changed — or one created
 * here — is rebuilt from `paragraphs`.
 */
export interface DocComment {
  id: number;
  author: string;
  initials?: string;
  date?: string;
  /** Body runs, one entry per `w:p`. */
  paragraphs: Inline[][];
  /** False when the body holds a table, picture or field: read-only, kept verbatim. */
  editable: boolean;
  raw: string;
  /** The body's run shapes as read, so a save can tell which comment actually changed. */
  baseline: string;
  /**
   * `w14:paraId` of each body paragraph, in order. `commentsExtended.xml` and
   * `commentsIds.xml` key on the first one, so without these a comment cannot carry a
   * resolved flag, a reply link, or the durable id Word holds across sessions.
   */
  paraIds?: string[];
  /** `w15:commentEx/@done`: WPS and Word grey the thread out rather than hiding it. */
  resolved?: boolean;
  /** Set when this comment answers another: `w15:paraIdParent`, mapped back to a comment id. */
  parentId?: number;
  /** `w16cid:commentId/@durableId`. */
  durableId?: string;
}

export const anchorInline = (anchor: CommentAnchor): Inline => ({ text: "", anchor });

export const commentTexts = (c: DocComment): string[] => c.paragraphs.map((runs) => runs.map((r) => r.text).join(""));

/** The whole body as editable text: one line per `w:p`. */
export const commentBody = (c: DocComment): string => commentTexts(c).join("\n");

/**
 * The body reduced to what a save can reproduce: the text of every run together with the style
 * it carries. Comparing this instead of the plain text is what makes a formatting-only edit count
 * as an edit — and comparing it in this canonical form (adjacent runs of one style coalesced) is
 * what stops a comment the reader only looked at from being rebuilt, since the file and the DOM
 * split the same sentence into different runs.
 */
export function bodyShape(paragraphs: Inline[][]): string {
  const styleOf = (r: Inline): string =>
    JSON.stringify([
      r.bold === true, r.italic === true, r.underline === true, r.strike === true,
      r.sizePt ?? 0, r.color ?? "", r.highlight ?? "", r.font ?? "", r.vertAlign ?? "", r.hyperlink ?? "",
    ]);
  const runs = (list: Inline[]): unknown[] => {
    const out: { key: string; text: string }[] = [];
    for (const r of list) {
      // A break or a tab is content of its own: it cannot be folded into a neighbour's text.
      const key = r.break ? "break" : r.tab ? "tab" : r.image ? "image" : styleOf(r);
      const last = out[out.length - 1];
      if (last && last.key === key && key !== "break" && key !== "tab" && key !== "image") last.text += r.text;
      else out.push({ key, text: r.text });
    }
    return out.map((r) => [r.key, r.text]);
  };
  return JSON.stringify(paragraphs.map(runs));
}

export const commentDirty = (c: DocComment): boolean => c.raw === "" || bodyShape(c.paragraphs) !== c.baseline;

/**
 * Everything a save has to reproduce — `raw` never changes, so it is not part of the state.
 * Resolve state, the reply link and the body's run shapes are user-visible, so they belong here;
 * `paraIds` and `durableId` do not, because they only change together with a comment that is
 * already new or edited, and counting them would make an untouched document look modified.
 */
export const commentState = (c: DocComment): unknown[] => [
  c.id,
  c.author,
  c.initials ?? "",
  c.date ?? "",
  bodyShape(c.paragraphs),
  c.resolved === true,
  c.parentId ?? -1,
];

export const isAnchor = (inline: Inline): boolean => inline.anchor !== undefined;

const XML_ESC = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
/** Resolve state and reply links (`w15:commentEx`). */
export const NS_W15 = "http://schemas.microsoft.com/office/word/2012/wordml";
/** Durable paragraph ids (`w16cid:commentId`). */
export const NS_W16CID = "http://schemas.microsoft.com/office/word/2016/wordml/cid";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";

/** Namespace declarations for a package that never had a comments part. */
export const NEW_COMMENTS_ROOT = ` xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" mc:Ignorable="w14"`;

/**
 * Every `w:comment` element in the part text, with its id and its exact source bytes.
 * Comments cannot nest, so a scan from each opening tag to its closing tag is enough.
 */
export function commentSpans(partXml: string): { id: number; xml: string }[] {
  const out: { id: number; xml: string }[] = [];
  // `\s` after the name keeps `w:comments`, `w:commentRangeStart` and `w:commentReference` out.
  const opens = /<([\w.-]+):comment(\s[^>]*)>/g;
  let match: RegExpExecArray | null;
  while ((match = opens.exec(partXml))) {
    const [tag, prefix, attrs] = match;
    const closing = `</${prefix}:comment>`;
    const end = partXml.indexOf(closing, match.index + tag.length);
    if (end < 0) continue;
    const id = Number(/[\w.-]+:id="(-?\d+)"/.exec(attrs)?.[1]);
    if (Number.isFinite(id)) out.push({ id, xml: partXml.slice(match.index, end + closing.length) });
    opens.lastIndex = end + closing.length;
  }
  return out;
}

/** The attribute text of the part's root element, including every namespace declaration. */
export function commentsRootAttrs(partXml: string): string {
  return /<[\w.-]+:comments\b([^>]*)>/.exec(partXml)?.[1] ?? "";
}

/**
 * What a writer needs from a part's root: the prefix bound to the wordprocessingml
 * namespace (created when the part never declared one) and the w14 prefix, whose
 * absence means the document has no paragraph ids to keep.
 *
 * `needsParaIds` covers the case where the model has ids but the part never declared w14 —
 * writing an extension row keyed on an id that does not appear in comments.xml would point
 * Word at a paragraph that is not there, so the namespace is declared here instead.
 */
export function commentsRootInfo(rootAttrs: string, needsParaIds = false): { prefix: string; attrs: string; w14?: string } {
  const found = new RegExp(`xmlns:([\\w.-]+)="\\s*${NS_W}\\s*"`).exec(rootAttrs);
  const prefix = found?.[1] ?? "w";
  let attrs = found ? rootAttrs : `${rootAttrs} xmlns:${prefix}="${NS_W}"`;
  let w14 = new RegExp(`xmlns:([\\w.-]+)="\\s*${NS_W14}\\s*"`).exec(rootAttrs)?.[1];
  if (needsParaIds && w14 === undefined) {
    w14 = "w14";
    attrs = `${attrs} xmlns:${w14}="${NS_W14}"`;
    // Word marks these namespaces ignorable so a reader that predates them can skip the ids.
    if (!/xmlns:mc=/.test(attrs)) attrs = `${attrs} xmlns:mc="${NS_MC}" mc:Ignorable="${w14}"`;
    else if (!/mc:Ignorable=/.test(attrs)) attrs = `${attrs} mc:Ignorable="${w14}"`;
  }
  return { prefix, attrs, w14 };
}

/** `w:paraId` values in document order; `commentsExtended.xml` and `commentsIds.xml` key on them. */
export function commentParaIds(xml: string): string[] {
  return [...xml.matchAll(/[\w.-]+:paraId="([^"]+)"/g)].map((m) => m[1]);
}

/**
 * A `w:comment` rebuilt from the model. `paraIds` are the ones the comment carried when it
 * was read, so `commentsExtended.xml` still resolves the paragraphs of an edited comment.
 */
export function commentElement(comment: DocComment, root: { prefix: string; w14?: string }, write: (runs: Inline[]) => string, paraIds: string[] = []): string {
  const p = `${root.prefix}:`;
  const attrs = [
    `${p}id="${comment.id}"`,
    `${p}author="${XML_ESC(comment.author)}"`,
    comment.initials ? `${p}initials="${XML_ESC(comment.initials)}"` : "",
    comment.date ? `${p}date="${XML_ESC(comment.date)}"` : "",
  ].filter(Boolean);
  const body = comment.paragraphs
    .map((runs, i) => {
      const id = paraIds[i] && root.w14 ? ` ${root.w14}:paraId="${XML_ESC(paraIds[i])}"` : "";
      return `<${p}p${id}>${write(runs)}</${p}p>`;
    })
    .join("");
  return `<${p}comment ${attrs.join(" ")}>${body}</${p}comment>`;
}

// ------------------------------------------------------- extension parts

/** One `w15:commentEx`, kept with its source bytes so an entry we cannot map survives. */
export interface CommentExEntry {
  paraId: string;
  parentParaId?: string;
  resolved: boolean;
  raw: string;
}

/**
 * An attribute value by name, prefixed so `paraId` cannot match inside `paraIdParent`.
 * Reads both the namespaced form Word writes and a bare name.
 */
const attr = (entry: string, name: string): string | undefined =>
  new RegExp(`[\\s:]${name}="([^"]*)"`).exec(entry)?.[1];

const selfClosing = (name: string): RegExp => new RegExp(`<[\\w.-]+:${name}\\b[^>]*/>`, "g");

export function commentExEntries(partXml: string): CommentExEntry[] {
  const out: CommentExEntry[] = [];
  for (const [entry] of partXml.matchAll(selfClosing("commentEx"))) {
    const paraId = attr(entry, "paraId");
    if (!paraId) continue;
    const done = attr(entry, "done");
    out.push({
      paraId,
      parentParaId: attr(entry, "paraIdParent"),
      resolved: done === "1" || done === "true" || done === "on",
      raw: entry,
    });
  }
  return out;
}

/** The `w16cid:commentId` rows: a paragraph id and the durable id Word keeps for it. */
export function commentIdEntries(partXml: string): { paraId: string; durableId?: string; raw: string }[] {
  const out: { paraId: string; durableId?: string; raw: string }[] = [];
  for (const [entry] of partXml.matchAll(selfClosing("commentId"))) {
    const paraId = attr(entry, "paraId");
    if (!paraId) continue;
    out.push({ paraId, durableId: attr(entry, "durableId"), raw: entry });
  }
  return out;
}

export const commentsExRootAttrs = (partXml: string): string =>
  /<[\w.-]+:commentsEx\b([^>]*)>/.exec(partXml)?.[1] ?? "";

export const commentIdsRootAttrs = (partXml: string): string =>
  /<[\w.-]+:commentsIds\b([^>]*)>/.exec(partXml)?.[1] ?? "";

/** The prefix bound to `uri` in a root's attribute text, declaring it under a default when absent. */
function prefixFor(rootAttrs: string, uri: string, fallback: string): { p: string; attrs: string } {
  const found = new RegExp(`xmlns:([\\w.-]+)="\\s*${uri}\\s*"`).exec(rootAttrs);
  return { p: found?.[1] ?? fallback, attrs: found ? rootAttrs : `${rootAttrs} xmlns:${found?.[1] ?? fallback}="${uri}"` };
}

/**
 * `w15:commentsEx` for the comments this model owns, in comment order, followed verbatim by
 * the entries it does not own. A resolve flag or reply link with no entry here is silently
 * dropped by Word, so a document that has comments always gets a row per comment.
 */
export function commentsExXml(
  rootAttrs: string,
  rows: { paraId: string; parentParaId?: string; resolved: boolean }[],
  foreignRaw: string[] = []
): string {
  const { p, attrs } = prefixFor(rootAttrs, NS_W15, "w15");
  const body = rows
    .map(
      (r) =>
        `<${p}:commentEx ${p}:paraId="${XML_ESC(r.paraId)}"` +
        `${r.parentParaId ? ` ${p}:paraIdParent="${XML_ESC(r.parentParaId)}"` : ""}` +
        ` ${p}:done="${r.resolved ? "1" : "0"}"/>`
    )
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n${`<${p}:commentsEx${attrs}>${body}${foreignRaw.join("")}</${p}:commentsEx>`}`;
}

/** `w16cid:commentsIds`, the durable ids keyed on the same paragraph ids. */
export function commentIdsXml(
  rootAttrs: string,
  rows: { paraId: string; durableId: string }[],
  foreignRaw: string[] = []
): string {
  const { p, attrs } = prefixFor(rootAttrs, NS_W16CID, "w16cid");
  const body = rows
    .map((r) => `<${p}:commentId ${p}:paraId="${XML_ESC(r.paraId)}" ${p}:durableId="${XML_ESC(r.durableId)}"/>`)
    .join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<${p}:commentsIds${attrs}>${body}${foreignRaw.join("")}</${p}:commentsIds>`;
}

// ------------------------------------------------------------- id supply

/**
 * `ST_LongHexNumber`: eight upper-case hex digits. Word only ever issues ids whose first
 * digit is 0-7, and reusing one a part already carries would collide the extension rows,
 * so the scan starts above every id in the package and skips the ones taken.
 */
export function allocHexId(taken: Set<string>, start = 0x1000_0001): string {
  for (let value = start; value < 0x8000_0000; value += 7919) {
    const id = value.toString(16).toUpperCase().padStart(8, "0");
    if (!taken.has(id)) {
      taken.add(id);
      return id;
    }
  }
  return "";
}

/** Eight hex digits over everything the package already uses. */
export function allocIds(count: number, taken: Set<string>, start?: number): string[] {
  return Array.from({ length: count }, () => allocHexId(taken, start));
}
