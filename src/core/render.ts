import type { Block, Carried, CarriedInline, HeaderFooter, Inline, NumberLevel, PageBorder, Paragraph, PropSlot, RowInfo, Section, Table, TableBorder, TableCell, WordDocument } from "./docx.js";
import { DEFAULT_SECTION, hasLine, LINE_ONE, MARGIN_SIDES, PAGE_BORDER_SIDES, PBDR_ORDER, headingLevel, lineGridPitchTwips, twipsToPx } from "./docx.js";
import type { CommentAnchor } from "./comments.js";
import type { Presentation, Shape, SlideParagraph } from "./pptx.js";

const esc = (s: string): string =>
  s.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");

/**
 * A format record's `w:*Pr` children as the file wrote them, carried on the page beside the fields.
 *
 * The editor re-derives the model from the DOM on every keystroke, and a record is part of that model: with only
 * the fields on the page, the next save rebuilds the record from them. `prev` can hold the eleven properties the
 * model has a field for and nothing else, so the rebuild takes the record's own `w:kern`, `w:lang` and the run's
 * `w:spacing` down with them, and a reviewer refusing the change afterwards refuses a change no document made.
 * A record signed by this editor holds no bytes, and gets no attribute.
 */
const slotsAttr = (name: string, slots?: PropSlot[]): string =>
  slots?.length ? ` ${name}="${encodeURIComponent(JSON.stringify(slots))}"` : "";

/**
 * §17.3.2.26 splits a run's text by script: `w:ascii` draws the Latin, `w:eastAsia` the Chinese, Japanese and
 * Korean. A plain CSS font stack falls through per glyph by *coverage*, which is a different rule — a Latin face
 * that happens to hold Han glyphs would swallow the Chinese (178 of the 1,783 Chinese runs whose faces differ
 * naming a Chinese family for `w:ascii` too — a pass on the quarantined cohort; `.scratch/run-font-eastasia.ts`
 * has no authorized reading, so that audience is 待重证). So the east-asian face travels
 * as a family scoped to the East Asian ranges by `unicode-range`, and the browser's own matching dispatches by
 * script the way the element says. `src: local()` only resolves against a face the machine has installed; a file
 * naming one it does not falls through to the Latin face, which is exactly what this page drew before the field
 * existed, so an absent font never makes the page worse than it was.
 */
const EAST_RANGES = "U+2E80-9FFF,U+F900-FAFF,U+FE10-FE1F,U+FF00-FFEF,U+20000-2FA1F";
/**
 * A face name as it may appear inside CSS. The name comes from the file, and `<` is legal in an XML attribute
 * value, so `w:eastAsia="evil</style><img>"` is a real thing a document can say: the string delimiters go and
 * `</` is written `<\/`, which CSS reads as `/` and the HTML tokenizer does not read as a closing tag. Both ends
 * of a face use this one spelling — the rule's family and the run's stack — so they always name the same thing.
 */
const cssName = (face: string): string => face.replace(/["\\\n\r]/g, "").replace(/<\//g, "<\\/").trim();
/** The family this page invents for one East Asian face. */
const faceKey = (face: string): string => `solo-ea-${cssName(face)}`;

/** The `@font-face` rules holding these faces to the East Asian ranges. Empty for an empty list. */
export function eastAsiaFaceRules(faces: string[]): string {
  const out: string[] = [];
  const seen = new Set<string>();
  for (const face of faces) {
    const key = faceKey(face);
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(`@font-face{font-family:"${key}";src:local("${cssName(face)}");unicode-range:${EAST_RANGES};}`);
  }
  return out.join("");
}

/** The stack a run draws with: the east-asian family first (it only answers for East Asian codepoints), the
 *  Latin one after it. Quoted, because a face name is a string, not an identifier. A numbered paragraph's mark
 *  builds the same stack from its own resolved faces, so the symbol and the text beside it wear one pair.
 *  Exported because the editor's pending-format placeholder previews a picked face and must not invent a second
 *  spelling of the pair. */
export function fontStack(run: Pick<Inline, "font" | "fontEastAsia">): string {
  const parts: string[] = [];
  if (run.fontEastAsia) parts.push(`"${faceKey(run.fontEastAsia)}"`);
  const latin = run.font?.replace(/[;\\]/g, "").trim();
  if (latin) parts.push(`"${latin}"`);
  return parts.join(",");
}

/** The page marker for one `w:br`. `data-brk` carries the element source verbatim through an edit: without it
 *  a re-typed paragraph comes back as a bare `<br>`, and the writer spells that as `<w:br/>` — losing
 *  `w:type="page"`. `doc-edit` reads the attribute straight back into `Inline.brXml`. */
export const breakHtml = (xml?: string): string => (xml ? `<br data-brk="${esc(xml)}">` : "<br>");

export function renderInline(inline: Inline): string {
  if (inline.image) {
    return `<img src="${inline.image.src}" width="${Math.round(inline.image.width)}" height="${Math.round(inline.image.height)}" alt="${esc(inline.image.name)}">`;
  }
  if (inline.break) return breakHtml(inline.brXml);
  if (inline.tab) return '<span class="tab"></span>';
  let css = "";
  if (inline.bold) css += "font-weight:600;";
  if (inline.italic) css += "font-style:italic;";
  // One declaration for both lines: `text-decoration` is a shorthand, so writing it twice makes the later
  // declaration replace the earlier one — measured in the browser the page runs in, `line-through` then
  // `underline` computes to `underline` alone and a run that is both struck and underlined loses its strike
  // (`.scratch/onoff-dom.ts`). Stated once, both lines draw and `styleOf` reads both back off one value.
  const decorations = [inline.strike ? "line-through" : "", hasLine(inline.underline) ? "underline" : ""].filter((d) => d !== "");
  if (decorations.length) css += `text-decoration:${decorations.join(" ")};`;
  // 线的粗度由名字说：授权 46 份里 `w:u` 说过 1,615 枚，其中 `thick` 20 枚坐在 2 份件里（另有 16 枚在段落标记上，
  // `.scratch/underline-audience.py` 逐枚）。把这 20 枚画成单线，就是把文件说的粗线收成细线。
  // 粗到什么程度是本页的口径（WPS 真机画多粗仍未取证），但写法必须是**相对**长度：`auto`（单线那一档）跟着字号
  // 长粗，`.scratch/underline-ink.mjs` 逐行数截屏的墨量到 11 磅/12 磅/24 磅/36 磅分别是 1/1/3/4 px，而写死的
  // `2px` 在后两档反而比单线还细——那 `thick` 就成了「比正常更细」。`0.1875em` 量出来是 3/3/6/9 px，四档都稳在
  // 单线的 2~3 倍。`text-decoration` 是 shorthand，会把没提到的长项重置，所以这一句跟在它后面。
  if (inline.underline === "thick") css += "text-decoration-thickness:0.1875em;";
  if (inline.sizePt) css += `font-size:${inline.sizePt}pt;`;
  // `w:rPr/w:spacing` speaks twentieths of a point per character, and CSS `letter-spacing` is the same measure in
  // points: one declaration per character, negative allowed, and it is the character advance that changes, so a
  // line breaks where the file's spacing says. Painting the run at its unstated advance is what left WPS's own
  // 加宽 4.1 磅 titles (14 runs saying `w:val="82"` among 授权 46 件) sitting at normal width here.
  if (inline.charSpacingTwips !== undefined) css += `letter-spacing:${inline.charSpacingTwips / 20}pt;`;
  if (inline.color) css += `color:${inline.color};`;
  const family = fontStack(inline);
  if (family) css += `font-family:${esc(family)};`;
  // The name the document chose travels as data, not just as a colour: §17.17.64's members are what the file says,
  // and a colour is what the page makes of one. Reading it back off the CSS is the route that loses a sentence —
  // measured in the browser the page runs in, `background:none` hands its longhand back as `initial`, `blue` comes
  // back as `blue`, and the seven names this renderer spells are the only ones whose `rgb(...)` can be looked up
  // again (`.scratch/hl-pending.ts`). A run that stated `w:highlight w:val="none"` therefore came back from the
  // page saying nothing, and the writer dropped the sentence with it.
  const hl = inline.highlight ? ` data-hl="${esc(inline.highlight)}"` : "";
  if (inline.highlight) css += `background:${highlightColor(inline.highlight)};`;
  // 线名里只有 CSS 带不回来的那几句钉在属性上。`single` 是这页画线时的默认样子，读回来不必再问；`none` 是「说了
  // 不要」，纸上只能读出「没画」，那与「没提过」是同一个样——#85 量到的正是这一眼降级；`thick` 的粗度 CSS 交不回
  // 名字。名字一丢，重写 rPr 时那句 `w:u` 就被当成「这枚 run 从没说过」，文件的话被换成沉默或换成细线。
  const u = inline.underline && inline.underline !== "single" ? ` data-u="${esc(inline.underline)}"` : "";
  if (inline.vertAlign && inline.vertAlign !== "baseline")
    css += `vertical-align:${inline.vertAlign === "superscript" ? "super" : "sub"};font-size:.75em;`;
  let html = inline.text === "" ? "" : `<span${css ? ` style="${css}"` : ""}${hl}${u}>${esc(inline.text)}</span>`;
  if (html === "" && !css) html = esc(inline.text);
  // A link to a place in this document keeps its window to itself: `target="_blank"` on a `#name` href hands
  // the app's own page address to the system browser, where nobody has the bookmark open. The attributes the
  // file put on the link element have no meaning on the page, so they ride along as data for the editor to
  // hand back on the next keystroke — dropped here, they are gone from the saved file.
  if (inline.hyperlink)
    html =
      `<a href="${esc(inline.hyperlink)}"${inline.hyperlink.startsWith("#") ? "" : ' target="_blank" rel="noreferrer"'}` +
      `${inline.linkAttrs ? ` data-link-attrs="${esc(inline.linkAttrs)}"` : ""}>${html || esc(inline.text)}</a>`;
  // A tracked change is marked by class, never by inline style: the editor reads inline styles
  // back into the run, and "underlined because inserted" must not become "underlined because bold".
  const rev = inline.rev;
  if (rev) {
    const what = rev.move === "from" ? "移出" : rev.move === "to" ? "移入" : rev.kind === "del" ? "删除" : "插入";
    const who = `${rev.author}${rev.date ? ` · ${rev.date.slice(0, 10)}` : ""} ${what}`;
    html =
      `<span class="rev rev-${rev.kind}" data-rev="${rev.kind}" data-rev-id="${rev.id}"` +
      ` data-rev-author="${esc(rev.author)}"${rev.date ? ` data-rev-date="${esc(rev.date)}"` : ""}` +
      `${rev.move ? ` data-rev-move="${rev.move}"` : ""}` +
      ` title="${esc(who)}">${html || esc(inline.text)}</span>`;
  }
  // A formatting revision changes how a run looks, not what it says: the run keeps its text and shows
  // its new properties, so the record rides along as data on the run's own span for the editor to read
  // back. It is not a `w:ins`/`w:del` mark and must not be given one class or the edit will read it as
  // an insertion.
  const fmt = inline.revFormat;
  if (fmt) {
    const who = `${fmt.author}${fmt.date ? ` · ${fmt.date.slice(0, 10)}` : ""} 修改格式`;
    html =
      `<span class="rev-fmt" data-rev-format="${fmt.id}" data-rev-format-author="${esc(fmt.author)}"` +
      `${fmt.date ? ` data-rev-format-date="${esc(fmt.date)}"` : ""}` +
      ` data-rev-prev="${encodeURIComponent(JSON.stringify(fmt.prev))}"` +
      slotsAttr("data-rev-prev-slots", fmt.prevSlots) +
      ` title="${esc(who)}">${html || esc(inline.text)}</span>`;
  }
  return html;
}

export function highlightColor(name: string): string {
  const map: Record<string, string> = {
    yellow: "#ffff00",
    green: "#00ff00",
    cyan: "#00ffff",
    magenta: "#ff00ff",
    red: "#ff0000",
    darkYellow: "#808000",
    lightGray: "#d4d4d4",
    // §17.17.64's cancellation draws nothing, and it is still a sentence the document said. Spelled
    // `background:transparent` rather than `background:none` because the shorthand resets more than the colour and
    // hands its longhand back as `initial`, which no reverse lookup can name; `transparent` reads back as itself.
    none: "transparent",
  };
  return map[name] ?? name;
}

// `left`/`start`/`end` are here for the same reason as `center`: a word that is spoken must draw as what it
// says even when something above it (a row) says another. The authorized 46 份 state `left` on 108 paragraphs.
const ALIGN: Record<string, string> = {
  left: "left",
  start: "left",
  center: "center",
  right: "right",
  end: "right",
  both: "justify",
  distribute: "justify",
};

/**
 * Comment rendering state for one pass over the body: which anchored ranges are currently
 * open, and the number each comment shows as its reference mark. Anchors are in-stream, so
 * the set is threaded across paragraphs and table cells exactly as the ranges are.
 */
interface CommentState {
  open: Set<number>;
  numbers: Map<number, number>;
  /** Ids whose thread is resolved, so the mark renders greyed rather than hidden. */
  done: Set<number>;
}

function anchorTag(anchor: CommentAnchor, numbers: Map<number, number>, done: Set<number>): string {
  const label = numbers.get(anchor.id);
  const text = anchor.kind === "ref" && label !== undefined ? String(label) : "";
  // WPS and Word keep a resolved thread in place and grey it out; hiding it would lose the
  // reader's context for the surrounding discussion.
  const resolved = done.has(anchor.id) ? " cmt-done" : "";
  return `<span class="cmt-anchor${anchor.kind === "ref" ? " cmt-ref" : ""}${resolved}" data-cid="${anchor.id}" data-kind="${anchor.kind}" contenteditable="false">${text}</span>`;
}

/** The class an inline carry is drawn with, and the only way the page hands the same one back. */
export const RUN_CARRY_CLASS = "run-carry";

/**
 * A run-level structure the model carries rather than understands. `data-carry` is its identity across a
 * re-render, the same job `data-bid` does for blocks. What shows is what a reader would see inside it — the
 * result of a page field, the date in an inline content control — and nothing at all for a bookmark, which
 * has no visible content to give.
 */
function carryTag(carry: CarriedInline): string {
  return `<span class="${RUN_CARRY_CLASS}" data-carry="${carry.key}" data-tag="${esc(carry.tag)}" contenteditable="false">${esc(carry.show)}</span>`;
}

/**
 * A page field inside a header/footer band, drawn as a labelled box. A page number that reads like the text
 * beside it would be typed over and deleted by accident; `页码` says it is one thing, and the caret cannot
 * enter it. Inline style, because the page's stylesheet is not this module's to change — and the label, not
 * the number the field cached, because this product has no pagination to say which page it would be.
 */
function hfFieldTag(carry: CarriedInline): string {
  return `<span class="${RUN_CARRY_CLASS}" data-carry="${carry.key}" data-tag="${esc(carry.tag)}"${carry.field ? ` data-field="${carry.field}"` : ""} contenteditable="false" style="border:1px dashed #9aa3b0;border-radius:3px;padding:0 5px;margin:0 2px;background:#f2f5fa;color:#3b4453">${esc(carry.show)}</span>`;
}

/** The band's own row of commands, and the button that matches the page's other small buttons.
 *  停着的时候这一排是收起来的：它是编辑器的把手，纸上的带子里没有它。摊在带子的一行里，原件就比纸上的复件
 *  高出一排按钮那么高，而让位只按纸上那一截留（真机对照：同一份 10 行页眉 WPS 让 2005 缇，带把手的那一版让
 *  2670 缇）——多出来那一截不但不留给正文，还直接压在正文第一行上。点开这一枚带子时由 `doc-edit.ts` 亮出来。 */
const HF_BAR_STYLE = "display:none;gap:6px;margin-top:4px";
const HF_BTN_STYLE = "font:inherit;font-size:11px;padding:2px 8px;border:1px solid #c3c9d2;border-radius:5px;background:#f2f4f7;color:#26292e;cursor:pointer";

/** Runs between an open range's markers are wrapped, so the anchored text reads as one block. */
function wordRuns(inlines: Inline[], state: CommentState): string {
  const out: string[] = [];
  let i = 0;
  while (i < inlines.length) {
    const inline = inlines[i];
    if (inline.carry) {
      out.push(carryTag(inline.carry));
      i += 1;
      continue;
    }
    if (inline.anchor) {
      if (inline.anchor.kind === "start") state.open.add(inline.anchor.id);
      else if (inline.anchor.kind === "end") state.open.delete(inline.anchor.id);
      out.push(anchorTag(inline.anchor, state.numbers, state.done));
      i += 1;
      continue;
    }
    const from = i;
    while (i < inlines.length && !inlines[i].anchor && !inlines[i].carry) i += 1;
    const stretch = inlines.slice(from, i).map(renderInline).join("");
    // `data-cids` says which comments hold this stretch anchored: the UI maps a caret in the
    // text back to a comment with it, and one range can sit inside several anchored ranges.
    const ids = [...state.open].join(",");
    out.push(ids === "" ? stretch : `<span class="cmt-in" data-cids="${ids}">${stretch}</span>`);
  }
  return out.join("");
}

/**
 * The paragraph's first-line indent as one CSS value, from whichever of the two spellings the file put its
 * weight behind: the twips value when it states a non-zero one, otherwise the character count (WPS writes 首行
 * 缩进 2 字符 as `w:firstLineChars="200"`), and a `w:hanging` pulls the other way. Both are stated in 1,147 of
 * the quarantined cohort's elements and their ratio is always a font size, so choosing the twips value there
 * draws what WPS drew; the 88 that state a positive `w:firstLine` next to `w:firstLineChars="0"` are the reason
 * the character form does not simply win, as §17.3.1.14 has it. `w:ind` has not been re-counted on the authorized
 * 46 份, so 1,147 / 88 / 236 below are 待重证.
 *
 * A `w:firstLineChars` with no twips counterpart — 236 elements — is drawn in `em`, which makes it follow this
 * page's own font size rather than the document's. 未取证 how wide WPS makes it, since the size WPS measures
 * the character against is the paragraph mark's, and the model does not carry that as a number.
 */
const firstLineIndent = (p: Paragraph): string => {
  const twips = p.indentFirstLine ?? (p.indentHanging === undefined ? undefined : -p.indentHanging);
  if (twips) return `${twipsToPx(twips)}px`;
  const chars = p.indentFirstLineChars ?? 0;
  return chars > 0 ? `${(chars / 100).toFixed(2)}em` : "";
};

/** Three decimals is enough to place a line on a page and keeps the declaration readable in the DOM. */
const roundPx = (px: number): number => Math.round(px * 1000) / 1000;

/** CSS px at 96dpi against the points a `.docx` states sizes in: 1 磅 = 4/3 px. */
const ptToPx = (pt: number): number => roundPx((pt * 96) / 72);

/**
 * The line grid the page measures snapped paragraphs in, in px, or `undefined` when the section asks for none.
 * §17.3.2.9 gives `w:docGrid@w:type` four spellings and the authorized corpus uses three: **58 `w:sectPr` carry 58
 * `w:docGrid`, 49 of them saying `lines`, 6 saying `linesAndChars`, 3 saying nothing**
 * (`.scratch/corpus-recheck2.mjs`, per-section counters). `linesAndChars` is a line grid too — the character grid
 * is what it adds, not what it replaces — so both spellings land in `lineGridPitchTwips`, while `default`,
 * `snapToColumns` (0 instances here) and a whole absent `w:docGrid` all mean "no line grid to draw", which is what
 * every fixture written before this round states. A grid with no `w:linePitch` says where the lines are but not how
 * far apart they lie, and a number this model would have to invent is not a grid: all 58 state the pitch, so the
 * branch is spec-guarded, not corpus-fed.
 * 哪一种 `w:type` 算「有行网格」由 `lineGridPitchTwips` 一处判定，「段落」对话框那句小字问的是同一个答案。
 */
const gridPitchPx = (sec: Section): number | undefined => {
  const pitch = lineGridPitchTwips(sec);
  return pitch === undefined ? undefined : roundPx(twipsToPx(pitch));
};

/**
 * The grid's own height for a line that needs `px`: the next whole number of grid lines, never less than one.
 * This is what 对齐到网格 does to a line taller than the grid — 固定值 18 磅 on a 312 缇 grid draws 41.6px, two
 * lines, which is the well-known "行距莫名其妙变大一倍" whose standard fix is unchecking that box (the authorized
 * corpus writes that box off itself 226 times at a paragraph's own property slot, `.scratch/corpus-recheck2.mjs`).
 * 未取证: how WPS rounds is a font-metrics question no observation here settles; the reading is §17.3.2.9's, and
 * the direction — a snapped line never gets shorter than the height asked for — is the only one that keeps
 * 固定值 a floor rather than a clip.
 */
const gridLinesPx = (px: number, pitchPx: number): number => roundPx(Math.max(1, Math.ceil(px / pitchPx - 1e-9)) * pitchPx);

/**
 * The tallest character height this paragraph asks for, in points — the height a 倍数 or a 最小值 line is really
 * measured against. Two places the file can put it: the run's own resolved size (§17.7.2's ladder, docDefaults
 * included, so a run that states no `w:sz` still has one) and the paragraph mark's (§17.3.1.4). Taking the
 * greater of the two is what keeps a floor from clipping. 待重证: the audience for "text taller than its own
 * paragraph mark" was counted on the quarantined cohort (63 of 201 份 / 1,298 snapped multiplier 段,
 * `.scratch/em-ref-impact.ts`) and has not been re-derived on the authorized 46; the rule itself does not lean on
 * that number, because it follows from §17.3.1.4 (the mark carries a height too) and from a floor being a floor —
 * a mark-only reading would draw those lines under the glyphs. Undefined only when the file says nothing about any
 * of it, which is where `em` and the page's own size are all there is to draw against.
 */
const textHeightPt = (p: Paragraph): number | undefined => {
  const sizes: number[] = [];
  if (p.markSizePt !== undefined && p.markSizePt > 0) sizes.push(p.markSizePt);
  for (const i of p.inlines) if (i.sizePt !== undefined && i.sizePt > 0) sizes.push(i.sizePt);
  return sizes.length ? Math.max(...sizes) : undefined;
};

/**
 * The paragraph's line height as one CSS value, in the unit the file wrote it in. `w:line` counts 240ths of a
 * line under `w:auto` (2,890 of the authorized 46 份's 3,948 stated heights, in the body/band/footnote parts, with
 * 200 more in the style layer) and twips under `w:exact`/`w:atLeast` (1,004 + 54 of them), so drawing every one of
 * them as a multiplier put 固定值 22 磅 at 1.83 倍行距 — a quarter of the heights in the files WPS itself wrote.
 * Every one of those 3,948 signs its own `w:lineRule`: the reading is 2,890 auto / 1,004 exact / 54 atLeast and no
 * height arrived unsigned (`.scratch/authorized-spacing-census.py`; the 13,396 and 6,802 over the old cohort are
 * its arithmetic). A stated
 * `w:line="0"` draws no height at all and no unit turns a zero into a line; the authorized 46 份 state one 0 times,
 * so that guard is for a file the cohort has not shown yet.
 *
 * `exact` is a height, so it is simply that height; glyphs taller than it are cut, which is what 固定值 does.
 * `atLeast` is a floor, and CSS has no line height that says "this much, or whatever the text needs", so it is
 * drawn as the greater of the stated height and the paragraph's own tallest text — the least that cannot clip it.
 * 未取证 how much taller WPS makes such a line when the text needs more than that, since that is a question
 * about font metrics and this page has none.
 *
 * `pitchPx` is the section's line grid, and when it is there and the paragraph has not opted out
 * (`w:snapToGrid w:val="0"`, `undefined` reading as on), the grid is the unit: a paragraph that states no
 * height of its own takes one grid line — 3,507 of the cohort's 7,455 `w:pPr` blocks state none, and they are not
 * the majority there (3,948 do state a height), which is the correction to the old cohort's "24,117 of 27,511 say
 * nothing" — and one that states a height takes that height rounded up to whole grid lines. A
 * `w:auto` multiplier is the one case CSS cannot round: `1.5倍` means one and a half times *the line's own text*,
 * which the browser only knows per line, so on a grid it is drawn as the greater of one grid line and that many
 * times the paragraph's tallest text. That number has to be a length rather than `em`: `em` on a declaration
 * taken by the whole paragraph resolves against the paragraph's own font size, and before this the paragraph
 * carried none, so every one of those floors was measured against the App's 12pt — 735 of the 1,298 such
 * paragraphs in the 201 documents draw a different height once the file's own size answers instead, in 22 of
 * the files (`.scratch/em-ref-impact.ts`; 待重证 on the authorized 46, as the sentence above says). The grid
 * reaching a table cell is the same reading as the same
 * property on any other paragraph (§17.3.1.87 is a paragraph property, and cells hold paragraphs); WPS' own
 * in-cell grid is 未取证.
 */
const lineHeight = (p: Paragraph, pitchPx?: number): string => {
  if (p.lineSpacing === 0) return "";
  const pitch = p.snapToGrid === false ? undefined : pitchPx;
  const text = textHeightPt(p);
  if (!p.lineSpacing) return pitch === undefined ? "" : `${pitch}px`;
  if (p.lineSpacingRule === "exact") {
    const px = twipsToPx(p.lineSpacing);
    return `${pitch === undefined ? px : gridLinesPx(px, pitch)}px`;
  }
  if (p.lineSpacingRule === "atLeast") {
    const floor = pitch === undefined ? twipsToPx(p.lineSpacing) : gridLinesPx(twipsToPx(p.lineSpacing), pitch);
    return `max(${floor}px, ${text === undefined ? "1em" : `${ptToPx(text)}px`})`;
  }
  const multiple = p.lineSpacing / LINE_ONE;
  if (pitch === undefined) return multiple.toFixed(2);
  return `max(${pitch}px, ${text === undefined ? `${multiple}em` : `${ptToPx(text * multiple)}px`})`;
};

/**
 * A ground the page may paint: exactly the shape `hex()` reads a `w:fill` into. Anything else the field holds is
 * the file's own token for no colour (`auto`), and a token goes into a `style` attribute verbatim, so only this
 * form is drawn.
 */
const GROUND = /^#[0-9A-Fa-f]{6}$/;

/**
 * 段落的框与底纹在页上占的那一块：§17.3.2.36 的 `w:pBdr` 与 §17.3.2.33 的 `w:shd`。
 *
 * 每一边各说各的，所以逐边问：`none`/`hidden`/`nil` 三枚字面都是「这一边不画线」（`drawsLine`，与表框、
 * 页框同一个口径），而 #85 那条教训在这里更要紧——样式把四边全说成 `none` 的那 16 枚是**取消**，不是空白，
 * 对着「没写」画就等于替文件编一条它明令不要的线。只有真画线的那一边才把它的 `w:space`（磅）垫成
 * padding：线离字多远是文件自己说的，没线的一边不占这一格。
 *
 * 受众（`.scratch/pbdr-audience.ts`，46 份授权件、7,430 个正文段）：26 枚 `w:pBdr` 里 24 枚在
 * `word/styles.xml`（16 枚四边全 `none`、8 枚一条 `single` 底线），页眉部件的段落 2 枚全取消，正文段落自己
 * 写框的 0 枚。
 * `w:between`（相邻两段中间那条）与 `w:bar`（贯穿左边距的那一条）在这 26 枚里 0 枚，`w:shadow` 也 0 枚。
 * `w:between` 自 #153 起**画**：ribbon 那条「内部横框线」落在段落上说的就是它（见 `betweenOf`），落笔不画就是
 * 一句没有反馈的话——但那一档 `w:space` 与「一条段落的框线该多厚」在语料里都还没有样本，横边按顶/底那一档走。
 * `w:bar` 与 `w:shadow` 不画，也还没有实据可照着画（未取证），字节按 `pPrSlots` 原样带回，读得出（`p.borders`）。
 *
 * 底纹只认 `w:fill`（30 枚段落底纹全是 `w:val="clear"` 加一个十六进制 `w:fill`，`w:color` 30/30 写 `auto`）。
 * 十六进制以外的那一枚 token（`auto`）不画：那是文件说「要页面自己的地」，不是没有说——样式链上那一层填了色，
 * 对着「没说」画就会把样式的地铺进一段明令不要地的段落（`ParaStyle.shading`：段落层 0 例，`w:tcPr` 层 1,600 例，
 * 形状是文件自己的惯用写法）。铺多宽是 未取证 的另一半：这里铺段落自己那一格（缩进之内、含首行缩进让出的那块），
 * WPS 把底纹铺到缩进外还是里没观察过。
 */
function paraBox(p: Paragraph, middle?: TableBorder): string[] {
  const out: string[] = [];
  let padBottom = 0;
  for (const side of PBDR_ORDER) {
    if (side === "between" || side === "bar") continue;
    const edge = p.borders?.[side];
    if (!drawsLine(edge?.val)) continue;
    const css = borderCss(`border-${side}`, edge);
    if (!css) continue;
    out.push(css.slice(0, -1));
    const space = edge?.space;
    if (space !== undefined && space > 0) {
      const px = roundPx(ptToPx(space));
      // 底边那一格与中间那一条要的是同一块地方，取大的那一个，别让后推的那条把前一条挤没。
      if (side === "bottom") padBottom = px;
      else out.push(`padding-${side}:${px}px`);
    }
  }
  if (middle) {
    // 中间那一条画在**这一段自己的下沿**：贴着 padding 盒的底，所以落在底框线更内侧的那一格。用背景而不是再来
    // 一条 border，是因为 `w:bottom` 与 `w:between` 同时说话时那两条线确实各画一条，一条 CSS 边画不出两条。
    // 背景画不出 `double` 那种双线，虚线/点线按 6/12 与 2/4 的重复段近似——这两档的间距是**推断**，语料里
    // `w:between` 一枚样本都没有，落笔的那几条腿只写得 `single`。
    const px = edgeWidthPx(middle);
    const color = edgeColor(middle);
    const stroke = STROKE[middle.val ?? ""] ?? "solid";
    const paint =
      stroke === "dashed"
        ? `repeating-linear-gradient(90deg,${color} 0 6px,transparent 6px 12px)`
        : stroke === "dotted"
          ? `repeating-linear-gradient(90deg,${color} 0 2px,transparent 2px 4px)`
          : `linear-gradient(${color},${color})`;
    out.push(`background-image:${paint}`);
    out.push("background-position:bottom left");
    out.push(`background-size:100% ${px}px`);
    out.push("background-repeat:no-repeat");
    padBottom = Math.max(padBottom, roundPx(ptToPx(middle.space ?? 0)) + px);
  }
  if (padBottom > 0) out.push(`padding-bottom:${padBottom}px`);
  if (p.shading && GROUND.test(p.shading)) out.push(`background-color:${p.shading}`);
  return out;
}

/**
 * `w:between` 该不该在这一段下面画那一条：件里说的是**这一段与下一段之间**（§17.3.2.36，`docx.ts` 的
 * `PBDR_ORDER` 那一句），所以两头都得在——下一块是段落、而且两段是同一个样式（样式不同就没有「之间」那句话）。
 * 一段说完下一段不必再说，所以问的是**这一段**自己。
 */
const betweenOf = (p: Paragraph, next: Block | undefined): TableBorder | undefined => {
  const edge = p.borders?.between;
  if (!edge || !drawsLine(edge.val)) return undefined;
  if (next?.kind !== "paragraph") return undefined;
  return next.styleId === p.styleId ? edge : undefined;
};

function renderParagraph(p: Paragraph, doc: WordDocument, counters: Map<string, number>, state: CommentState, gridPx?: number, next?: Block): string {
  const style: string[] = [];
  if (p.align && ALIGN[p.align]) style.push(`text-align:${ALIGN[p.align]}`);
  if (p.indentLeft) style.push(`margin-left:${twipsToPx(p.indentLeft)}px`);
  if (p.indentRight) style.push(`margin-right:${twipsToPx(p.indentRight)}px`);
  const first = firstLineIndent(p);
  if (first) style.push(`text-indent:${first}`);
  if (p.spacingBefore !== undefined) style.push(`margin-top:${twipsToPx(p.spacingBefore)}px`);
  if (p.spacingAfter !== undefined) style.push(`margin-bottom:${twipsToPx(p.spacingAfter)}px`);
  style.push(...paraBox(p, betweenOf(p, next)));
  // The size the file's ladder resolves for this paragraph's own mark, on the paragraph. Without it the page's
  // 12pt answers for the strut, for a run that states no size, and for the height of an empty paragraph. On the
  // authorized 46 files the ladder speaks for 3,119 of its 3,120 正文段 (one says nothing): 1,422 of them draw a multiple of the grid
  // and 488 of those resolve *above* 12pt (432) or below it (34) — 408 paragraphs whose drawn number really moves
  // once the reference is the file's own size, in 24 of the 46 files — while the plain multiple branch holds 188
  // with 76 off 12pt (`.scratch/em-ref-impact.ts` over `/tmp/wps-authorized.txt`). A ladder that says nothing stays
  // in the stylesheet's number: this draws the file's size where the file states one, never one it does not.
  if (p.markSizePt !== undefined) style.push(`font-size:${p.markSizePt}pt`);
  const lh = lineHeight(p, gridPx);
  if (lh) style.push(`line-height:${lh}`);
  if (p.pageBreakBefore) style.push("page-break-before:always");
  // The heading a paragraph is drawn as. `data-outline` is the outline level the model resolves — its own or
  // the one its style chain states (§17.3.2.26) — and the 大纲 pane counts those marks in the order the page
  // writes them, which is why the attribute carries the resolved number rather than a guessed one. On the authorized
  // 46 files 408 of its 7,430 paragraph blocks resolve a level — 202 stating it themselves, 206 taking it from the
  // chain — spread as 0 × 298, 1 × 60, 9 × 26, 2 × 14, 3 × 8, 4 × 2 (`.scratch/outline-census.ts` over
  // `/tmp/wps-authorized.txt`).
  const outline = headingLevel(p);
  // The tag follows that level, but only down to 4. The page's stylesheet holds those four level-neutral
  // (`font: inherit; line-height: inherit; margin: 0`), so an `h1` the file says nothing about draws exactly the
  // line body text draws (#135); a deeper level stays a `div`, because the browser's own sheet would draw `h5`
  // and below *smaller* than body text — turning a deep level into a heading would put the file's heading under
  // the paragraph it introduces. A `w:pStyle` name still counts on its own, because a Word file can name
  // `Heading2` and state no level — and on the authorized 46 files that branch is the one nobody reaches: of the
  // 172 paragraph styles that state a level, **0 carry a `w:styleId` matching `^Heading([1-9])$`** (110 of them are
  // *named* the English `heading N` but carry some other id), so the tag below is always the resolved level's.
  const named = /^Heading([1-9])$/i.exec(p.styleId ?? "");
  const drawn = outline ?? (named ? Number(named[1]) : undefined);
  const tag = drawn !== undefined && drawn <= 4 ? `h${drawn}` : "div";
  let prefix = "";
  // `w:numId w:val="0"` is the file saying this paragraph is not numbered (§17.3.4.25) while still carrying a
  // `w:numPr`: the authorized 46 files have 298 paragraphs with the element and 52 of them are that cancellation
  // (`.scratch/corpus-recheck2.mjs`; `.scratch/pgmar.mjs` counts 0 that omit `w:numId` altogether). Drawing a
  // bullet for one adds a list member the file refused. `w:ilvl w:val="-1"` — 86 of them on the old, quarantined
  // reading — appears **0** times in the authorized corpus (`.scratch/oneoff-gap.mjs`), so that spelling is
  // 未取证 and nothing here depends on it.
  if (p.numPr && p.numPr.numId !== 0) {
    const { numId, ilvl } = p.numPr;
    const level = doc.levelOf(numId, ilvl);
    // Each numId is its own list, so a bullet list does not consume the numbers of a numbered one.
    const key = `${numId}:${ilvl}`;
    const seen = counters.get(key) ?? (level?.start ?? 1) - 1;
    counters.set(key, seen + 1);
    // The mark wears the font the level names, not the paragraph's: WPS spells a bullet as a private-use
    // codepoint in `w:lvlText`, and that codepoint only has a shape under the font its `w:rPr` points at.
    // Drawn in the paragraph's font it is tofu, and the list loses the one thing that made it that bullet.
    const mark = esc(numberText(doc, counters, numId, ilvl, seen + 1, level));
    // Font, size and colour are all the mark's own character properties (§17.3.1.12: the symbol is a run of its
    // own), and all three travel as a declaration on the mark's span only — never on the paragraph, which the
    // editor reads back through `inlinesOf`; that skips `.bullet`, so none of them can come back as body formatting.
    // Font and size fall back the same way: the level's own statement wins, and where the level is silent it is the
    // paragraph's mark chain (§17.3.1.4) — not the page stylesheet, whose 12pt and Times/Songti pair are the
    // App's own default and no document asked for them.
    // The colour has no level side to beat yet: of the 110 `w:lvl` the authorized 46 份's `numbering.xml` files
    // spell, none states a `w:rPr/w:color`, so nothing here has evidence for what a level's colour should override
    // (`.scratch/authorized-numbering-census.py`; the 617 levels / 8 stating a colour is the quarantined cohort's
    // arithmetic). The mark's own rung answering 82 of the 473 listed paragraphs and 391 stating nothing is
    // 待重证: that pair came off `.scratch/mark-color.ts`, whose file list belongs to the quarantined cohort and no
    // longer exists. What does not move is that no file asked for a colour this renderer would invent, so a mark
    // with no colour of its own stays in the page's ink.
    const markPt = level?.sizePt ?? p.markSizePt;
    const markStyle: string[] = [];
    if (level?.font) markStyle.push(`font-family:${esc(level.font)}`);
    else {
      const markFaces = fontStack({ font: p.markFont, fontEastAsia: p.markFontEastAsia });
      if (markFaces) markStyle.push(`font-family:${esc(markFaces)}`);
    }
    if (markPt) markStyle.push(`font-size:${markPt}pt`);
    if (p.markColor) markStyle.push(`color:${p.markColor}`);
    prefix = `<span class="bullet"${markStyle.length ? ` style="${markStyle.join(";")}"` : ""}>${mark}</span>`;
    style.push(`padding-left:${8 + ilvl * 24}px`);
  }
  const body = wordRuns(p.inlines, state) || "&nbsp;";
  // A paragraph-level revision rides on the paragraph element itself: the break that ends it is not
  // text a run can carry. Class and data attributes only, never inline style, so the editor cannot
  // read a mark back as formatting — and the properties record has to survive the round trip through
  // the page, because the editor re-derives the model from the DOM on every keystroke.
  const mark = p.markRev;
  const fmt = p.revFormat;
  const classes = `para${p.styleName ? " s-" + esc(p.styleName.replace(/\s+/g, "-")) : ""}${mark ? ` rev-mark-${mark.kind}` : ""}${fmt ? " para-rev" : ""}`;
  const carried =
    // 段落自己算好的制表位（缇、档），由 `doc-layout.ts` 的 `applyTabStops` 拿去量每一枚制表符该走多宽。
    // 写在 data 上而不是 style 上：编辑读回的是格式，不该把这些数当成段落自己新说的一句话。
    (p.tabStops?.length
      ? ` data-tab-stops="${p.tabStops.map((s) => `${s.posTwips}:${s.val === "right" ? "r" : s.val === "center" ? "c" : "l"}`).join(",")}"`
      : "") +
    (mark
      ? ` data-rev-mark="${mark.kind}" data-rev-mark-id="${mark.id}" data-rev-mark-author="${esc(mark.author)}"${mark.date ? ` data-rev-mark-date="${esc(mark.date)}"` : ""}${mark.move ? ` data-rev-mark-move="${mark.move}"` : ""}`
      : "") +
    (fmt
      ? ` data-para-rev="${fmt.id}" data-para-rev-author="${esc(fmt.author)}"${fmt.date ? ` data-para-rev-date="${esc(fmt.date)}"` : ""}` +
        ` data-para-prev="${encodeURIComponent(JSON.stringify(fmt.prev))}"` +
        slotsAttr("data-para-prev-slots", fmt.prevSlots)
      : "");
  const shown = mark ? `${mark.author}${mark.date ? ` · ${mark.date.slice(0, 10)}` : ""} ${mark.move ? (mark.move === "from" ? "移出段落标记" : "移入段落标记") : mark.kind === "del" ? "删除段落标记" : "插入段落标记"}` : fmt ? `${fmt.author}${fmt.date ? ` · ${fmt.date.slice(0, 10)}` : ""} 修改段落格式` : "";
  return `<${tag} class="${classes}"${outline === undefined ? "" : ` data-outline="${outline}"`}${p.keepNext ? ` data-keep-next="1"` : ""}${p.keepLines ? ` data-keep-lines="1"` : ""}${p.widowControl === false ? ` data-widow="0"` : ""}${carried}${shown ? ` title="${esc(shown)}"` : ""}${style.length ? ` style="${style.join(";")}"` : ""}>${prefix}${body}</${tag}>`;
}

/**
 * What the symbol in front of a numbered paragraph is made of. The file states it as `w:lvlText`, a template
 * whose `%N` is the number of level N of the same list (§17.3.1.65), so the digits and the punctuation between
 * them both come from the part rather than from this renderer. 授权 46 份里 22 份带 `word/numbering.xml`（去重 13
 * 个模板），**110 枚 `w:lvl` 全写 `w:lvlText`，110 枚全含 `%N`**；最后一个占位符之后跟的东西是（正则切片与
 * ElementTree 两条口径逐枚一致，`.scratch/authorized-families3.mjs`）：
 * `decimal` 88 枚 → `.` 66、`、` 10、`）` 6、`．` 2、`)` 2、什么都不跟 2；`chineseCounting` 14 枚 → `、` 10、`）` 4；
 * `lowerLetter` 4 枚 → `．` 2、`）` 2；`lowerRoman` 2 枚 → 一个空格；`decimalEnclosedCircleChinese` 2 枚 → 空。
 * So the "." this file used to append is right for 66 of the 88 `decimal` levels and wrong for 22 of them, and the
 * split is not even clean per format: the same `chineseCounting` answers with two different marks inside one
 * corpus, which is why the template is copied rather than reconstructed from `w:numFmt`.
 * 待重证: the earlier replay of old-vs-new renderer (501 段 draws a symbol, of which 380 change) counted over the
 * quarantined cohort and has not been re-run; the fixtures that pin 符号按模板画 carry the same facts one paragraph
 * at a time.
 *
 * A `%N` that names another level reads the counter that list is on; one nothing has drawn yet shows that
 * level's `w:start`. A level with no `w:lvlText` at all falls back to the bare number in its own format, and a
 * `w:numId` the part cannot resolve keeps the dot it has always drawn — no paragraph in the corpus is in that
 * shape, so nothing here is measured, only not lost.
 */
function numberText(doc: WordDocument, counters: Map<string, number>, numId: number, ilvl: number, n: number, level: NumberLevel | undefined): string {
  if (!level) return "•";
  const text = level.lvlText;
  if (!text) return ordinal(n, level.numFmt);
  return text.replace(/%(\d)/g, (_all, d: string) => {
    const other = Number(d) - 1;
    if (other === ilvl) return ordinal(n, level.numFmt);
    const below = doc.levelOf(numId, other);
    return ordinal(counters.get(`${numId}:${other}`) ?? below?.start ?? 1, below?.numFmt ?? level.numFmt);
  });
}

/**
 * A number in the shape one `w:numFmt` asks for (§17.18.4), as the number alone: the punctuation around it is
 * the template's, not this function's. Of the values the corpus's numbering definitions use, `decimal`,
 * `lowerLetter`, `upperLetter`, `lowerRoman`, `upperRoman` and `none` were already here; `chineseCounting`
 * (92 levels, drawn by 213 paragraphs) and `decimalEnclosedCircleChinese` (1 level) fell through to the decimal
 * digits, which is neither the number nor the mark the file asked for. The 中文小写 counting — 一…九、十、十一、
 * 二十、二十一…九十九 — and the circled digits are the enumerations themselves; a number past the range they
 * spell, or a value the corpus never writes, keeps the decimal digits.
 *
 * The two letter forms were one code point short of their own first item: `96 + ((n - 1) % 26)` puts U+0060
 * ("`") on n = 1 and 64 ("@") on the first uppercase item, so each of the authorized 46 份's 4 `lowerLetter` levels (and their 2 `lowerRoman` siblings;
 * 110 `w:lvl` in all, of which 88 are `decimal` and 14 `chineseCounting` — `.scratch/authorized-numbering-census.py`,
 * against the quarantined cohort's 75) drew its items one letter early. A letter sequence is also a bijective base-26 count — a, b, …, z, aa — not a
 * modulo, which restarted at a on the 27th item.
 */
function ordinal(n: number, fmt: string): string {
  switch (fmt) {
    case "lowerLetter":
      return letters(n, 97);
    case "upperLetter":
      return letters(n, 65);
    case "lowerRoman":
      return toRoman(n).toLowerCase();
    case "upperRoman":
      return toRoman(n);
    case "chineseCounting":
      return chineseCount(n);
    case "decimalEnclosedCircleChinese":
      return n >= 1 && n <= 20 ? String.fromCodePoint(0x2460 + n - 1) : String(n);
    case "none":
      return "";
    default:
      return String(n);
  }
}

/** A letter sequence: a, b, …, z, aa, ab, with `base` the code point of the first letter. */
function letters(n: number, base: number): string {
  if (!Number.isInteger(n) || n < 1) return String(n);
  let out = "";
  for (let k = n; k > 0; k = Math.floor((k - 1) / 26)) out = String.fromCharCode(base + ((k - 1) % 26)) + out;
  return out;
}

const CN_DIGITS = "零一二三四五六七八九";

/** 中文小写: the counting words for 1..99, and the decimal digits past them. */
function chineseCount(n: number): string {
  if (!Number.isInteger(n) || n < 0 || n > 99) return String(n);
  if (n < 10) return CN_DIGITS[n];
  if (n === 10) return "十";
  const tens = Math.floor(n / 10);
  return `${tens === 1 ? "" : CN_DIGITS[tens]}十${n % 10 === 0 ? "" : CN_DIGITS[n % 10]}`;
}

function toRoman(n: number): string {
  const table: [number, string][] = [[10, "X"], [9, "IX"], [5, "V"], [4, "IV"], [1, "I"]];
  let out = "";
  for (const [v, s] of table) while (n >= v) ((out += s), (n -= v));
  return out;
}

/** The stroke names Word uses for a border, and how CSS draws the same line. */
const STROKE: Record<string, string> = {
  single: "solid",
  thick: "solid",
  double: "double",
  dotted: "dotted",
  dashed: "dashed",
  dotDash: "dashed",
  dash: "dashed",
  dashSmallGap: "dashed",
  dashDotStroked: "dashed",
  dashDot: "dashed",
  dashDotDot: "dashed",
  threeDashSmallGap: "dashed",
  hairline: "solid",
  inset: "inset",
  outset: "outset",
  groove: "groove",
  ridge: "ridge",
};

/** The stroke tokens that name no line, which is what they say. Everything else draws a line. */
const NO_LINE: ReadonlySet<string> = new Set(["none", "hidden", "nil"]);

/**
 * Whether a `w:val` asks for a visible stroke. The file's own answer, not the renderer's fallback: an unknown
 * token still draws, as solid.
 */
export const drawsLine = (val?: string): boolean => val !== undefined && !NO_LINE.has(val);

/** `w:sz`（八分之一磅）落到 CSS 的那一档，最小一条线。 */
const edgeWidthPx = (edge?: TableBorder): number => Math.max(1, Math.round((edge?.size ?? 8) / 6));

/** `w:color` 说的颜色；`auto` 是「跟着字走」，不是没有说。 */
const edgeColor = (edge?: TableBorder): string =>
  !edge?.color || edge.color === "auto" ? "currentColor" : "#" + edge.color.replace(/^#/, "").toUpperCase();

/**
 * One edge as CSS — a table's rule, a cell's, or the page's own border, which is why the edge is taken by
 * structure rather than by name. A token `STROKE` has no entry for draws a solid line of the width asked for,
 * the least-wrong reading of a stroke the file invented. 未取证: how each of these looks in WPS' own page view
 * was never observed.
 */
function borderCss(property: string, edge?: TableBorder): string {
  if (!edge || !drawsLine(edge.val)) return "";
  return `${property}:${edgeWidthPx(edge)}px ${STROKE[edge.val] ?? "solid"} ${edgeColor(edge)};`;
}

const isSet = (edge?: TableBorder): edge is TableBorder => edge !== undefined;

/**
 * 「查看网格线」那一档的屏幕导引线：文件里**没人替这一条边说「画一条线」**（明说 `none` 也算没人说）时，
 * 给这一条边留一条内阴影。为什么用 `box-shadow` 不用 `border`：阴影不进盒模型，开着与关着之间纸上的行高、
 * 表宽一个像素都不动（DOM 闸门逐条量这两档的 `getBoundingClientRect`）。这一档归容器上的 `doc-grid` 决定画不画，
 * 所以关掉之后那些边什么都不画——那才是文件说的样子。
 *
 * 为什么四条文**并成一整条列表**（一枚 `--gl`）而不是四条边各一枚 `--gl-top`：`box-shadow` 的列表里塞一枚
 * `none` 是非法值，整条声明一起作废（`.scratch/bs-td-probe.html` 那趟量到：`inset … , none` → `none`）。
 * 按四条边分槽写，「这一格只缺两条边」的那种最常见的情形就一条也画不出来——正好是这个切片要修的那种东西。
 */
const GRID_EDGE: Record<"top" | "right" | "bottom" | "left", string> = {
  top: "inset 0 1px 0 var(--gl-ink)",
  bottom: "inset 0 -1px 0 var(--gl-ink)",
  left: "inset 1px 0 0 var(--gl-ink)",
  right: "inset -1px 0 0 var(--gl-ink)",
};

const gridGuides = (sides: Record<"top" | "right" | "bottom" | "left", TableBorder | undefined>): string => {
  const guides = (Object.keys(GRID_EDGE) as (keyof typeof GRID_EDGE)[])
    .filter((side) => {
      const edge = sides[side];
      return !edge || !drawsLine(edge.val);
    })
    .map((side) => GRID_EDGE[side]);
  return guides.length ? `--gl:${guides.join(",")};` : "";
};

/** `w:tblW`, counted in its own type: twips, or fiftieths of a percent. */
function tableWidth(t: Table): string {
  if (t.widthTwips === undefined) return "";
  if (t.widthType === "pct") return `width:${t.widthTwips / 50}%;`;
  if (t.widthType === "auto") return "";
  return `width:${twipsToPx(t.widthTwips)}px;`;
}

/** `w:jc` and `w:tblInd`: where the table sits inside the text column, and how far it leans right. */
function tableAlign(t: Table): string {
  const indent = t.indentTwips ? `margin-left:${twipsToPx(t.indentTwips)}px;` : "";
  if (t.align === "center") return `margin-left:auto;margin-right:auto;`;
  if (t.align === "right") return `margin-left:auto;`;
  return indent;
}

/** `w:vAlign` on a cell, in the words CSS uses for the same thing. */
const VALIGN: Record<string, string> = { top: "top", center: "middle", bottom: "bottom" };

/** `w:textDirection`: the vertical spellings are the ones a page can actually draw. */
function writingMode(direction: string): string {
  if (direction.startsWith("tbRl")) return "vertical-rl";
  if (direction.startsWith("tbLr")) return "vertical-lr";
  return "horizontal-tb";
}

/**
 * A row's own height and hiding. `w:trHeight` (§17.3.2.77) states the row's own box, so the declaration has to
 * be one the browser treats as a floor: measured in the engine this page runs in
 * (`.scratch/row-height-css.html`), `min-height` on a `<tr>` — or on its `<td>` — is **ignored** (a row
 * declaring 200px draws 23px), while `height` is the row's height when the content is shorter (200px) and only
 * a floor when it is taller (203px). The files are one-sided about which of the two this is: on the authorized 46 份
 * all 580 `w:trHeight` — in 8 件, i.e. 4 template names each shipped twice (zh_CN and zh_HK) — state
 * `w:hRule="atLeast"` beside their `w:val`, none says `exact`, and none is taller than a page (70..3,244 缇 against
 * A4's 16,838; `.scratch/authorized-row-height.py`). The 2,112 heights over 115 table-bearing files is the
 * quarantined cohort's arithmetic, and counting these by file name would report 4 件, not 8. The page measures its tables off these boxes (`doc-layout.ts` reads one row as one
 * unit), so a floor the browser drops draws — and paginates — a table shorter than the file says.
 *
 * `exact` would additionally clip content that does not fit; a table row here cannot be forced below its
 * content, so the rule this corpus never uses draws as the same floor. The distinction stays in the model and
 * saves back unchanged, and WPS's own clipping on an `exact` row is 未取证.
 */
function rowStyle(row?: RowInfo): string {
  if (!row) return "";
  let css = row.hidden ? "display:none;" : "";
  if (row.heightTwips) css += `height:${twipsToPx(row.heightTwips)}px;`;
  // 行那一句（`w:trPr/w:jc`）落在 `<tr>` 上，格与段都从它继承；段自己说了话的那一处由 `ALIGN[p.align]`
  // 盖回去，正是 §17.3.1.8 的优先级（段 > 格 > 行）。授权 46 份里 4 份模板的 190 行说了这一句（全是
  // `center`），那 190 行里 1,294 格没有第二句可问——`w:tcPr/w:jc` 全语料 0 枚，这些格里的段也没有
  // `w:pPr/w:jc`。在这之前这一句从没读进来，1,294 格都按页面的左对齐画。
  const align = row.jc ? ALIGN[row.jc] : undefined;
  if (align) css += `text-align:${align};`;
  return css;
}

/**
 * A cell's own padding, side by side, from the three places a file may state it: the cell's `w:tcMar`
 * (§17.3.2.69), then the row's `w:tblPrEx` exception band, then the table's `w:tblCellMar` (§17.3.2.82).
 * A side nobody states keeps the page's own spacing, which is what a table the editor built draws —
 * `w:w="0"` is not the same fact as silence, so this asks per side rather than taking the set as a whole.
 *
 * Before this the padding was never read at all: the sheet gave every cell `5px 9px`, while the authorized 46 份's
 * table margin statements are all per side and mostly say `0` above and below against 108 twips at the sides — 10 份
 * hold a table (58 `w:tbl`, 698 `w:tr`, 4,136 `w:tc`), those 10 being 5 template names each shipped twice (zh_CN and
 * zh_HK), and **every one of them** states margins: table-level `w:tblCellMar` on all 58 tables, a row-level
 * `w:tblPrEx/w:tblCellMar` on 693 of the 698 rows, and cell-level `w:tcMar` in 2 份's 888 cells (653 of the
 * table/row sets are top/bottom `0` + left/right `108`, 98 are all `0`; `.scratch/authorized-table-margin.py`).
 * Counting those by file name is what used to report 5 份 and 1 份. Since the page measures each table row off its
 * drawn box to cut the pages (`doc-layout.ts`), the old `5px 9px` was 10px of height per row × 698 rows of paper
 * that is not in the file; the 115 table-bearing files / 204-1,660-1,897 / 2,647 rows figures is the quarantined
 * cohort's arithmetic.
 */
function cellPaddingCss(cell: TableCell, table: Table): string {
  const band = cell.row?.cellMargin;
  return MARGIN_SIDES.map((side) => {
    const twips = [cell.margin?.[side], band?.[side], table.cellMargin?.[side]].find((v) => v !== undefined);
    return twips === undefined ? "" : `padding-${side}:${twipsToPx(twips)}px;`;
  }).join("");
}

/**
 * What each cell of a table draws as: how many rows it covers, and whether it is the part a merge above
 * already drew. `w:vMerge` says only "restart" or "continue", so the height comes from counting the
 * continues that line up under a restart in the same grid column.
 *
 * The covered cell stays in the page, hidden, rather than being left out. The editor re-derives the model
 * from the cells it finds, and a row that has lost one of them would lose its `w:tcPr` with it — a vertical
 * merge that survives viewing but vanishes on save is the failure this is here to avoid. A `continue` with
 * no `restart` above it covers nothing (the original face of a `w:cellMerge` can state exactly that), so it
 * draws as its own one-row cell instead of disappearing with its text.
 */
function mergeOf(t: Table): { rows: number; hidden: boolean }[][] {
  const startOf = t.rows.map((cells) => {
    let column = 0;
    return cells.map((c) => {
      const at = column;
      column += Math.max(1, c.spanCols ?? 1);
      return at;
    });
  });
  const out = t.rows.map((cells) => cells.map(() => ({ rows: 1, hidden: false })));
  t.rows.forEach((cells, r) => {
    cells.forEach((c, index) => {
      if (c.spanRows !== 1) return;
      let spanned = 1;
      const covered: [number, number][] = [];
      for (let below = r + 1; below < t.rows.length; below++) {
        const continued = t.rows[below].findIndex((x, i) => x.spanRows === 0 && startOf[below][i] === startOf[r][index]);
        if (continued < 0) break;
        covered.push([below, continued]);
        spanned += 1;
      }
      out[r][index].rows = spanned;
      for (const [row, at] of covered) out[row][at]!.hidden = true;
    });
  });
  return out;
}

function renderTable(t: Table, doc: WordDocument, state: CommentState, gridPx?: number): string {
  const table = t.borders;
  const merges = mergeOf(t);
  const css = [
    // A table that draws its own lines wants them shared with the cells, which is what collapsing does;
    // one that has none keeps the page's own spacing.
    table ? "border-collapse:collapse;" : "",
    t.layout === "fixed" ? "table-layout:fixed;" : "",
    tableWidth(t),
    tableAlign(t),
    borderCss("border-top", table?.top),
    borderCss("border-left", table?.left ?? table?.start),
    borderCss("border-bottom", table?.bottom),
    borderCss("border-right", table?.right ?? table?.end),
    // 表自己的地（`w:tblPr/w:shd`）铺在格子底下；`auto` 不铺——那是文件说「要页面自己的地」，不是「铺个透明」。
    t.shading && GROUND.test(t.shading) ? `background:${t.shading};` : "",
  ].join("");
  const rows = t.rows
    .map((cells, r) => {
      // A row's own revision rides on the row, as a class plus data attributes the editor reads back —
      // never as inline style, so a mark cannot be mistaken for formatting when the page is re-read.
      const row = cells[0]?.row;
      const rev = row?.rev;
      const carried = rev
        ? ` class="rev-row rev-${rev.kind}" data-row-rev="${rev.kind}" data-row-rev-id="${rev.id}" data-row-rev-author="${esc(rev.author)}"${rev.date ? ` data-row-rev-date="${esc(rev.date)}"` : ""}`
        : "";
      const shown = rev ? `${rev.author}${rev.date ? ` · ${rev.date.slice(0, 10)}` : ""} ${rev.kind === "del" ? "删除行" : "插入行"}` : "";
      const height = rowStyle(row);
      const header = row?.header ? ` ${TBL_HEADER_ATTR}="1"` : "";
      return `<tr${carried}${header}${height ? ` style="${height}"` : ""}${shown ? ` title="${esc(shown)}"` : ""}>${cells
        .map((c, c2) => {
          const merge = merges[r][c2]!;
          // The table's inside lines belong to the seam, so they draw on the cell that faces it, and a
          // border the cell set for itself wins over the one the table set for the grid.
          const seamOfCell = (rr: number, cc: number) => {
            const cellsOf = t.rows[rr] ?? [];
            const that = cellsOf[cc]?.borders;
            return {
              top: [that?.top, rr > 0 ? table?.insideH : undefined].find(isSet),
              bottom: [that?.bottom, rr < t.rows.length - 1 ? table?.insideH : undefined].find(isSet),
              left: [that?.left ?? that?.start, cc > 0 ? table?.insideV : undefined].find(isSet),
              right: [that?.right ?? that?.end, cc < cellsOf.length - 1 ? table?.insideV : undefined].find(isSet),
            };
          };
          const seamOf = seamOfCell(r, c2);
          // 导引线要问的是「这一条**物理线**上有没有人画线」，所以两件事都得问：外圈那四条线是表自己说的
          // （写在上头那枚 `<table>` 上），最外一圈的格没人说话时得替它问一句表；而内部那一条缝由贴着它的
          // **两格任一格**说一句就画一次（collapse 把它们合成同一条线）——只问手上这一格，另一头那格画过线的
          // 缝会被再留一条导引，同一条线上就叠了两条线。
          const edgeOf = {
            top: seamOf.top ?? (r > 0 ? seamOfCell(r - 1, c2).bottom : undefined) ?? (r === 0 ? table?.top : undefined),
            bottom: seamOf.bottom ?? (r < t.rows.length - 1 ? seamOfCell(r + 1, c2).top : undefined) ?? (r === t.rows.length - 1 ? table?.bottom : undefined),
            left: seamOf.left ?? (c2 > 0 ? seamOfCell(r, c2 - 1).right : undefined) ?? (c2 === 0 ? table?.left ?? table?.start : undefined),
            right: seamOf.right ?? (c2 < cells.length - 1 ? seamOfCell(r, c2 + 1).left : undefined) ?? (c2 === cells.length - 1 ? table?.right ?? table?.end : undefined),
          };
          const style = [
            merge.hidden ? "display:none;" : "",
            c.widthTwips ? `width:${c.widthType === "pct" ? `${c.widthTwips / 50}%` : `${twipsToPx(c.widthTwips)}px`};` : "",
            // 认不出是颜色的那一句（`auto`、具名色）不往 style 里放：与段落同一道闸（`GROUND`），
            // 也是「说了不要地」与「没说」的分别——三态读法自 #152 起把 token 原样带进模型。
            c.shading && GROUND.test(c.shading) ? `background:${c.shading};` : "",
            c.valign ? `vertical-align:${VALIGN[c.valign] ?? c.valign};` : "",
            c.noWrap ? "white-space:nowrap;" : "",
            c.textDirection ? `writing-mode:${writingMode(c.textDirection)};` : "",
            cellPaddingCss(c, t),
            borderCss("border-top", seamOf.top),
            borderCss("border-bottom", seamOf.bottom),
            borderCss("border-left", seamOf.left),
            borderCss("border-right", seamOf.right),
            gridGuides(edgeOf),
          ].join("");
          const spans = [c.spanCols && c.spanCols > 1 ? ` colspan="${c.spanCols}"` : "", merge.rows > 1 ? ` rowspan="${merge.rows}"` : ""].join("");
          // 格那一层的结构记录（`w:cellIns` / `w:cellDel`）说的是「这一格在不在」，行那一层的记号管不着它：
          // 一枚记录可以只落在某一行的一格上（OOXML 记「修订的一列」时是给列里每一格盖同一个 `w:id`），
          // 所以这一枚要摆在 `td` 上，视图才有条按接受/拒绝那一句摆的腿——与行同一条规矩，不是猜 WPS 的形状。
          const cellRev = (c.revs ?? []).find((r) => r.tag === "cellIns" || r.tag === "cellDel");
          const cellKind = cellRev ? (cellRev.tag === "cellIns" ? "ins" : "del") : "";
          const cellAttrs = cellRev
            ? ` class="rev-cell rev-${cellKind}" data-cell-rev="${cellKind}" data-cell-rev-id="${cellRev.id}" data-cell-rev-author="${esc(cellRev.author)}"${cellRev.date ? ` data-cell-rev-date="${esc(cellRev.date)}"` : ""}`
            : "";
          const cellShown = cellRev ? `${cellRev.author}${cellRev.date ? ` · ${cellRev.date.slice(0, 10)}` : ""} ${cellKind === "ins" ? "插入单元格" : "删除单元格"}` : "";
          return `<td${style ? ` style="${style}"` : ""}${spans}${cellAttrs}${cellShown ? ` title="${esc(cellShown)}"` : ""}>${renderBlocks(c.paragraphs, doc, new Map(), state, gridPx)}</td>`;
        })
        .join("")}</tr>`;
    })
    .join("");
  return `<table class="doc-table"${css ? ` style="${css}"` : ""}><tbody>${rows}</tbody></table>`;
}

function renderBlocks(blocks: Block[], doc: WordDocument, counters: Map<string, number> = new Map(), state: CommentState = { open: new Set(), numbers: new Map(), done: new Set() }, gridPx?: number): string {
  return blocks
    .map((b, i) =>
      b.kind === "paragraph"
        ? renderParagraph(b, doc, counters, state, gridPx, blocks[i + 1])
        : b.kind === "table"
          ? renderTable(b, doc, state, gridPx)
          : renderCarried(b),
    )
    .join("");
}

/** How a carried container is drawn. The editor re-creates one with this when the page loses it. */
export const CARRY_STYLE = "border:1px dashed #b3b3b3;border-radius:2px;padding:2px 6px;margin:2px 0;color:#555";

/** The class a carried container's element wears: the page's one signal that a block is not editable. */
export const CARRY_CLASS = "doc-carry";

/**
 * 纸上的表头行标记（`w:trPr/w:tblHeader`，§17.4.49）。`data-` 而不用 class：它说的是一句**文件里的话**，
 * 不是页面自己的一层装饰，读回模型时不该被当成格式（与 `render.ts` 给行修订盖 `data-row-rev` 同一条口径）。
 * 铺排（`doc-layout.ts`）与页缝重画按它认「开头连续的那几行」，所以两处读的是同一个字面。
 */
export const TBL_HEADER_ATTR = "data-tbl-header";

/**
 * A block the model keeps as the source text it came in with: what is inside it is shown, and it is
 * read-only — the caret cannot enter, and a save writes the bytes back rather than what the page drew. The
 * box is drawn with inline style because the page's own stylesheet is not this module's to change; `data-tag`
 * names what was kept, and nothing else about the element is read back as formatting.
 */
function renderCarried(b: Carried): string {
  const body = esc(b.text) || "（此处内容暂不可编辑）";
  return `<div class="${CARRY_CLASS}" contenteditable="false" data-tag="${esc(b.tag)}" style="${CARRY_STYLE}">${body}</div>`;
}

/** The class a page border's element wears: drawn inside the paper box, and never a block of text. */
export const PAGE_BORDER_CLASS = "doc-page-border";

/**
 * 后面那些节自己的带子的**母本**（#157）：可点可改的只有纸盒外那两份原件（首张的页眉、末张的页脚），
 * 其余每一节真正显示的那一份画在这里，`display:none`、不可点、不带 `.hf` 这个名——
 * `doc-edit.ts` 的 `paintBandCopies` 拿它做那张纸上的复件。占 `.hf` 的名会被 `wireHF` 认领成可编辑的带子，
 * 一份件里两份能改的带子就都归不了一个主人。
 */
export const BAND_SOURCE_CLASS = "solo-hf-src";
/** 纸盒里那两份能点、能改的带子原件（#157：带子搬进纸内，正文按它占走的高度让位）。数块的时候与条带、
 *  复件同一条守则剔掉：它是纸上的装饰加一个可点的编辑口，不是正文的一块。 */
export const BAND_MASTER_CLASS = "solo-hf-master";

/**
 * The line `w:pgBorders` asks for, drawn as one absolutely positioned element per stated side inside the paper
 * box. Sides are separate because each names its own stroke and its own gap: a document with a thick bottom and
 * no right side has to draw that, and one rectangle with four CSS borders could not — nor could a box whose
 * missing side still holds the gap of its neighbour.
 *
 * The gap is `w:space` in points, measured against what `w:offsetFrom` names: the paper's edge, or the body,
 * whose edge on a side is that side's margin (plus the binding width on the left, as the padding counts it).
 * Each line spans the paper in the other direction, so four sides meet at the corners the way a frame does.
 */
function renderPageBorders(sec: Section): string {
  const borders = sec.pageBorders;
  if (!borders) return "";
  const gutter = sec.margins.gutter ?? 0;
  const inset = (side: (typeof PAGE_BORDER_SIDES)[number]): number => {
    const space = ((borders.sides[side]?.space ?? 0) * 4) / 3;
    if (borders.offsetFrom !== "text") return space;
    const margin =
      side === "top" ? sec.margins.top : side === "bottom" ? sec.margins.bottom : side === "left" ? sec.margins.left + gutter : sec.margins.right;
    return space + twipsToPx(margin);
  };
  const cross = (side: (typeof PAGE_BORDER_SIDES)[number]): string => (side === "top" || side === "bottom" ? "left:0;right:0" : "top:0;bottom:0");
  return PAGE_BORDER_SIDES.map((side) => {
    const css = borderCss(`border-${side}`, borders.sides[side]);
    return css
      ? `<div class="${PAGE_BORDER_CLASS}" data-side="${side}" contenteditable="false" aria-hidden="true" style="position:absolute;pointer-events:none;${cross(side)};${side}:${Math.round(inset(side))}px;${css}"></div>`
      : "";
  }).join("");
}

/**
 * A band's paragraphs as the page shows them. `.hf-p` is `white-space: pre-wrap`, so a tab the part holds stays
 * a tab on screen and comes back as one when the line is read out of the page again. A paragraph that holds a
 * field is drawn piece by piece, because the field is not text and the order the two sit in is the file's, not
 * the reader's.
 */
function bandParagraphsHtml(area: HeaderFooter): string {
  return (
    area.paragraphs
      .map((_, i) => {
        const segs = area.segments?.[i] ?? [{ text: area.paragraphs[i] ?? "" }];
        return `<div class="hf-p">${segs.map((s) => (s.carry ? hfFieldTag(s.carry) : esc(s.text ?? ""))).join("") || "<br>"}</div>`;
      })
      .join("") || '<div class="hf-p"><br></div>'
  );
}

/**
 * 这一页按哪一张脸画：`marks` 是「显示修订」，`final` 是「显示最终状态」，`original` 是「显示原文」。
 * 三档只改**摆出来的**东西，模型与写回一个字都不动。
 */
export type RevisionView = "marks" | "final" | "original";

/**
 * 正文与表/行/格那一叠纸面。`original` 那一档不是「切完 DOM 再补样式」：`doc-edit.ts` 的 `render()` 整块
 * 重铺 `innerHTML`，视图切换期补的任何行内样式下一次按键就没了，而 run 的行内样式又会被 `styleOf` 读回模型，
 * 补上去的改前那套就是把 run 改了。所以视图是**渲染的入参**——原文那一档由 `doc.originalView()` 先按每条
 * 记录的 `prev` 复制一份，再走同一条 style→CSS 映射。
 * 页眉页脚那一叠不在这趟里：带子是按 `segments` 的**文字**画的，属性改动记录从来就没进过带子的纸面，
 * 没有现值可换，也就没有半做。
 */
export function renderWord(doc: WordDocument, view: RevisionView = "marks"): string {
  const shown = view === "original" ? doc.originalView(doc.blocks) : doc.blocks;
  const sec = doc.sections[0] ?? DEFAULT_SECTION;
  const width = twipsToPx(sec.pageWidthTwips);
  // `w:gutter` is binding width: ECMA-376 §17.3.2.13 adds it to the margin on the binding side, and a section
  // that says nothing has none. This model holds no RTL section, so the binding side is the left.
  const gutter = sec.margins.gutter ?? 0;
  // The line grid every snapping paragraph on this page is measured in. It comes from the same section the page
  // box comes from, because one rendered page draws one section's text column — and on this corpus a later
  // section cannot disagree with it anyway: 授权 46 份的 **58 枚 `w:sectPr` 全带 `w:docGrid`，同一件里两节的
  // `(type, linePitch)` 不同的情况 0 件**（opener-walk 切片与 ElementTree 两条口径同为 58/58/0，pitch 值域只有
  // 312/285/360，`.scratch/authorized-families3.mjs`）。取第一节的网格因此不是赌气，是没有可分的第二格。
  const gridPx = gridPitchPx(sec);
  const style = `--page-w:${width}px;--pad-t:${twipsToPx(sec.margins.top)}px;--pad-r:${twipsToPx(sec.margins.right)}px;--pad-b:${twipsToPx(sec.margins.bottom)}px;--pad-l:${twipsToPx(sec.margins.left + gutter)}px`;
  // The distance a band's text is held from the paper's edge, taken from the section's own `w:header`/`w:footer`.
  // Those count from the **page edge** (§17.3.2.9 note on `w:pgMar`), so the band is placed by its own top, not by
  // a sentence of padding that holds its text away from the body — that padding was the substitute drawing while the
  // band hung outside the paper box. The real machine was measured placing it **inside** the margin, with the body
  // yielding below it (`.scratch` 夹具 `/tmp/tall-header.docx`，正文上沿 ≈3445 缇)，所以这里改成摆位、让位由
  // `doc-layout.ts` 的 `applyBandYield` 量出来。A section that states neither keeps the paper edge: inventing
  // Word's docDefaults here would draw a number nobody wrote.
  const bandEdge = (distance: number | undefined): number => (distance === undefined ? 0 : twipsToPx(distance));
  // 这两份原件守着这一叠纸的**首尾**，各自贴在所在那一张的纸沿里：页眉是第一张的，页脚是最后一张的，所以各按
  // `sectionBand` 取那一节真正显示的部件（#157 第二片）。此前两份都取文档级那一对，多节的件就把最后一张的页脚
  // 画成了第一节的字——纸上的复件按节取、能点的那一份却按文档级取，两边问的不是同一份部件，改完写回的也不是纸上那张。
  const lastSection = Math.max(0, doc.sectionRefs.length - 1);
  const visibleBand = (which: "header" | "footer"): HeaderFooter | undefined =>
    which === "header" ? doc.sectionBand("header", 0) : doc.sectionBand("footer", lastSection);
  const hf = (which: "header" | "footer"): string => {
    const area = visibleBand(which);
    if (!area) return "";
    const paras = bandParagraphsHtml(area);
    const name = which === "header" ? "页眉" : "页脚";
    const tip = area.editable ? `点击编辑${name}` : `此${name}含表格或图片等富内容，暂不支持编辑`;
    // The command the band carries, drawn outside the paragraphs so reading them back never picks it up as
    // text. Inline styles again: the page's stylesheet is not this module's to change.
    const bar = area.editable
      ? `<div class="hf-bar" contenteditable="false" style="${HF_BAR_STYLE}"><button class="hf-page" data-hf-page="${which}" style="${HF_BTN_STYLE}" title="在此${name}插入页码域">插入页码</button></div>`
      : "";
    // `hf-bottom` is the rule that puts a footer's padding under its text instead of above it. The band
    // travels with the page's own box variables: it lives inside the paper box now, and the horizontal padding
    // is overridden to the section's margins so the band's text starts at the same column as the body's.
    // Without them the stylesheet answers with its 794px fallback and a landscape section gets a portrait
    // header — same page, two widths.
    // 带子这一侧的纵向数按**它所属那一节**取：页脚原件守的是最后一张，那一节的下边距与 `w:footer` 说的距离
    // 才是它自己的。横向仍是开篇那一节的整套页变量——这一叠纸只有一个纸盒宽度（`doc-layout.ts` 头注「纵向」
    // 那条），带子与正文要落在同一栏里。
    const own = doc.sections[Math.min(which === "header" ? 0 : lastSection, Math.max(0, doc.sections.length - 1))] ?? sec;
    const edge = bandEdge(which === "header" ? own.margins.header : own.margins.footer);
    // 纵向落点先按节说的距离写死在这里，`doc-layout.ts` 的 `applyBandYield` 随后把它挪到「文字贴在距离上」
    // 那一格（样式自己那圈内边距不是距离的一部分，与纸上的复件同一算法），并按带子真正占走的高度让正文让位。
    const place = which === "header" ? `top:${edge}px` : `bottom:${edge}px`;
    // 带子现在住在**可编辑的纸盒里**，这一句 `contenteditable="false"` 是它 resting 状态唯一的来源：此前它挂在
    // 纸盒之外，不可编辑是从宿主继承来的，一句都不用写。缺了这句，`editHF` 开头那句 `isContentEditable` 就
    // 直接 return——带子点不开、打字落进正文、`sync()` 又把它当装饰跳过，两边都丢。
    return `<div class="hf hf-${which}${which === "footer" ? " hf-bottom" : ""} ${BAND_MASTER_CLASS}" data-hf="${which}" contenteditable="false" data-hf-part="${esc(area.part)}" style="${style};position:absolute;left:0;${place};padding-left:var(--pad-l);padding-right:var(--pad-r)" title="${tip}">${paras}${bar}</div>`;
  };
  const borders = renderPageBorders(sec);
  // 后面那些节各自显示的带子的母本（#157）：纸盒外可见的那两份只守着首张与末张，其余每一节真正显示的那一份
  // 画在这里，纸上的复件按节来抄。母本缺了就不画——把上一节的字贴到这一节的纸上，正是这一片要修的。
  // 可见的那两份不进母本（`seenBand` 先记下它们的部件）：一份件里两份能改的带子就都归不了一个主人。
  // 母本用与正文同一套页变量：这一页的纸盒就是按第一节那一套宽度画的（多节各按自己的纸高量，宽仍是这一套，
  // `doc-layout.ts` 头注「纵向」那条），复件落在同一张纸上，换宽就把正文与带子分到两个宽度里。
  const seenBand = new Set([visibleBand("header")?.part, visibleBand("footer")?.part]);
  const bandSources = doc.sectionRefs
    .flatMap((pair, k) =>
      (["header", "footer"] as const).flatMap((which) => {
        const area = which === "header" ? pair.header : pair.footer;
        if (!area || seenBand.has(area.part)) return "";
        seenBand.add(area.part);
        return `<div class="${BAND_SOURCE_CLASS}" data-hf-part="${esc(area.part)}" data-hf-kind="${which}" contenteditable="false" aria-hidden="true" style="${style};display:none;padding-left:var(--pad-l);padding-right:var(--pad-r)">${bandParagraphsHtml(area)}</div>`;
      }),
    )
    .join("");
  // An `@font-face` is a rule, not a declaration, so the east-asian families travel in a stylesheet next to the
  // page rather than in a `style` attribute. It sits beside `.page` for the same reason the borders below do not
  // lead it: the editor counts blocks among `.page`'s children.
  const faceRules = eastAsiaFaceRules(doc.eastAsiaFaces());
  // 装饰都排在**块之后**：编辑器按块在 `.page` 孩子里的下标寻址，任何排在前面的装饰会把每一个下标顶偏。
  // 页面边框与两份带子原件都排在正文之后（#157：带子此前排在纸盒之外，画在纸上看不到，正文也从没为它让位）。
  return `${faceRules ? `<style>${faceRules}</style>` : ""}<div class="page" style="${style};position:relative">${renderBlocks(shown, doc, new Map(), {
    open: new Set(),
    numbers: new Map(doc.comments.map((c, i) => [c.id, i + 1])),
    done: new Set(doc.comments.filter((c) => c.resolved === true).map((c) => c.id)),
  }, gridPx)}${borders}${hf("header")}${hf("footer")}</div>${bandSources}`;
}

function renderSlideParagraph(p: SlideParagraph): string {
  const css: string[] = [];
  if (p.align && ALIGN[p.align]) css.push(`text-align:${ALIGN[p.align]}`);
  const runs = p.runs.map(renderInline).join("");
  return `<div class="sp-p" style="${css.join(";")};padding-left:${p.level * 18}px">${p.bullet ? `<span class="bullet">${esc(p.bullet)}</span>` : ""}${runs || "<br>"}</div>`;
}

function renderShape(s: Shape, pres: Presentation, idx?: number): string {
  // Inline position:absolute keeps exported standalone HTML faithful without a stylesheet.
  const common = `position:absolute;left:${s.x}px;top:${s.y}px;width:${s.width}px;height:${s.height}px`;
  const rot = s.rotation ? `transform:rotate(${s.rotation}deg);` : "";
  // Only top-level shapes carry an identity: the editor selects and drags those.
  const pick = idx === undefined ? "" : ` data-idx="${idx}"`;
  if (s.kind === "picture") {
    return s.src ? `<img class="shape"${pick} style="${common};${rot}object-fit:fill" src="${s.src}" alt="${esc(s.name)}">` : "";
  }
  if (s.kind === "table") {
    const html = (s.rows ?? [])
      .map((r) => `<tr>${r.map((c) => `<td${c.fill ? ` style="background:${c.fill}"` : ""}>${c.runs.map(renderInline).join("") || esc(c.text)}</td>`).join("")}</tr>`)
      .join("");
    return `<div class="shape st-table"${pick} style="${common};${rot}"><table>${html}</table></div>`;
  }
  if (s.kind === "group") {
    return `<div class="shape"${pick} style="${common};${rot}">${(s.children ?? []).map((c) => renderShape(c, pres)).join("")}</div>`;
  }
  const style = [common, rot];
  if (s.geometry === "line") {
    // A line is a zero-height box in the model; its stroke is drawn on that line.
    const width = Math.max(1, s.lineWidth ?? 1);
    const color = s.lineColor ?? "#000000";
    return `<div class="shape"${pick} style="${style.join(";")}"><div style="position:absolute;left:0;right:0;top:0;transform:translateY(-50%);border-top:${width}px solid ${color}"></div></div>`;
  }
  if (s.fill) style.push(`background:${s.fill}`);
  if (s.lineColor) style.push(`border:${Math.max(1, s.lineWidth ?? 1)}px solid ${s.lineColor}`);
  if (s.geometry && s.geometry !== "rect") style.push(`clip-path:${clipFor(s.geometry)}`);
  const text = (s.paragraphs ?? []).map(renderSlideParagraph).join("");
  return `<div class="shape"${pick} style="${style.join(";")}"><div class="shape-text">${text}</div></div>`;
}

function clipFor(geom: string): string {
  switch (geom) {
    case "ellipse":
      return "ellipse(50% 50%)";
    case "roundRect":
      return "inset(0 round 12px)";
    case "triangle":
      return "polygon(50% 0,100% 100%,0 100%)";
    case "parallelogram":
      return "polygon(12% 0,100% 0,88% 100%,0 100%)";
    case "diamond":
      return "polygon(50% 0,100% 50%,50% 100%,0 50%)";
    case "rightArrow":
      return "polygon(0 30%,60% 30%,60% 0,100% 50%,60% 100%,60% 70%,0 70%)";
    case "star5":
      return "polygon(50% 0%,61% 35%,98% 35%,68% 57%,79% 91%,50% 70%,21% 91%,32% 57%,2% 35%,39% 35%)";
    case "chevron":
      return "polygon(25% 0,100% 0,75% 100%,0 100%)";
    default:
      return "";
  }
}

export function renderSlideHtml(pres: Presentation, index: number): string {
  const slide = pres.slides[index];
  if (!slide) return "";
  const bg = slide.background ? `background:${slide.background};` : "";
  return `<div class="slide" style="width:${pres.slideWidthPx}px;height:${pres.slideHeightPx}px;${bg}">${slide.shapes
    .map((s, i) => renderShape(s, pres, i))
    .join("")}</div>`;
}
