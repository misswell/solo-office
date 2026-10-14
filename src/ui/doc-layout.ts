/**
 * 「页数」与「行数」不是模型里的数，是**排版之后**才有的数。包里的 JS API 自己就把它们分在这一类：
 * `office6/addons/kwpsapiex/wps.js` 的 `getDocumentWordCount()` 把 `words`、`characters`、
 * `charactersWithSpaces`、`paragraphs`、`lines`、`pages` 全都交给 `doc.ComputeStatistics(wdStatistic…)`，
 * 而 `lines`/`pages` 那两问要的是字体度量与栏宽决定出来的行盒——第 66 轮做「字数统计」时正因为这个把这两行删了
 * （那时候它们没有受众：`render.ts` 只出一张连续的纸，照画报「页数 1」对一个四十段的件是假读数）。
 * 这一页把那个受众补上：**量的还是这张连续的纸**，按节的 `w:pgSz`/`w:pgMar` 把它切成 N 张，
 * 于是对话框那两行与状态栏那几格读的是同一个量出来的数。纸本身仍画成一张连续的（切成 N 张盒子是下一片）。
 *
 * 量的口径都在这里，别处不另定一份：
 *
 * - 一张纸能放正文的高度 = `w:pgSz@w:h` 换 px − 上边距 − 下边距。两个边距取的是**纸上算出来的 padding**
 *   （`render.ts` 由节写成 `--pad-t`/`--pad-b`），所以装订线、纸张方向、横竖版都在这一处自动跟着走。
 *   纸盒自己的 `min-height:1000px` 不参与：那是画出来好看，不是节说的纸高。
 * - `w:header`/`w:footer` 那两个距离**不进这个式子**：这一页的页眉页脚画在纸盒之外（`render.ts` 的 `.hf`），
 *   没有从正文里占走高度。Word 一族是把带子放在页边距之内的，所以我们量出的页数可能比真机少——**未取证**差多少。
 * - 行盒从**画出来的盒子**里量：`Range.getClientRects()` 给的是每个文字碎片的盒子，同一行上的碎片按 `top`
 *   归堆（容差取本段自己那条 `line-height` 的八成：跨行至少差一整行高，同行上两段字号不同也就差几像素）。
 *   一堆算一行，占的高度取这堆的上下沿，但不小于本段的行高。段与段之间真正画出来的空白（`w:spacing` 的
 *   before/after 在 CSS 里合成后的那一截）并进下一段第一枚单位里，所以流的总高与纸上画的一致。
 * - 段落只在**它自己的行盒之间**断开跨页；表格按**整行**铺（一行算一行，格子里换了几行不改这个数），
 *   只读容器（`.doc-carry`）整块算一枚。`w:cantSplit` 这一版没读，所以「这行不许拆」的文件与没写的同样铺。
 * - 另起一页的来路是**画在纸上的记号**：`br[data-brk]` 里那句源文本的 `w:type` 为 `page`，以及
 *   `render.ts` 由 `w:pageBreakBefore` 写出的 `page-break-before:always`。断点后面没有内容时仍然多算一页
 *   （那一页是空的）——断点说的是「从这里另起一页」，不是「后面有东西才另起」。
 * - **段尾断点多画出来的那几行空白既不计行也不占高**：断点落在段落末尾时浏览器给那一段多画空行，但**最后一条不画**
 *   （`.scratch/br-height-dom.ts` 一枚一枚按出来：1 枚→1.00 行、2 枚→2.00、3 枚→3.00、4 枚→4.00、5 枚→5.00，
 *   也就是 N 枚多画 N−1 条）。这一截空白落在段落的盒子高度里、不落在任何一枚行盒上（`Range.getClientRects()`
 *   给只有 `br` 的那一截零个矩形），所以这里记成一枚 0 高的记号：「翻页，但不占行」。于是行数正好是内容行数，
 *   而页数少算了那 N−1 条空白的高度（默认 A4 一行 25.92px、一页正文 930.53px 放得下 35 行），偏乐观。
 *   WPS 的 `ComputeStatistics` 在这种形状上各报什么**未取证**；把连续纸切成 N 个 `.page` 盒子（下一片）时这一处要连着改。
 *   `.scratch/pagination-dom.ts` 按这条口径钉行数，独立来路扣的是「枚数−1」而不是「枚数」。
 * - **分栏符是一条推导**：这一版模型没有 `w:cols` 字段，画出来就是一栏；一栏的下一栏起点就是下一页同一处，
 *   所以 `w:type="column"` 也按另起一页铺。语料那侧的量是「86 份 WPS 写出的件里 115 枚 `w:cols` 全部
 *   `w:num="1"`」（`.scratch/br-audience.mjs`），单栏是这些件的常态；WPS 真机上分栏符怎么翻页**未取证**。
 * - 一节一节地量**没做**：`render.ts` 只按 `sections[0]` 那张纸画，所以后面几节换了纸大小，也仍是按第一节的
 *   页高铺的。多节的排版是另一片。
 * - **光标落在第几页**（状态栏「第 X 页」，WPS 叫 `KxWpsStatusPageNumberInfoCommand`）走 `pageAtCaret()`：
 *   取光标处（折叠 Range）自己画出来的那枚矩形，用它的**中线**在量好的单位里找落点，再用与 `packPages`
 *   同一趟装箱得出那一枚在第几页。中线而不是上沿：光标矩形的上沿是**字体盒**的上沿，与单位那条**行盒跨度**
 *   差零点几像素且方向不定（实测 1123.47 对 1123.48/1123.5），正好落在相邻两枚相接的缝上——中线在行盒正中，
 *   离任何接缝都有半个行高。矩形落在纸面之上/之下时取它下面那一枚（纸面之下取最后一枚）——那是防呆，
 *   不是「按最近的算」。光标不在纸面上（点了面板、页眉带）这一层由调用方判掉，这里只管坐标。
 * - **反过来问「第 N 页在哪」**走 `pageCaretAnchor()`：「定位」对话框（`doc-goto.ts`）给的是页号，回来的是
 *   文档位置（折叠 Range 的起端），由调用方落光标、滚动。取的是**这一页上第一枚单位**的起点：不经过视口取点
 *   （`caretRangeFromPoint` 只认视口里的点，先滚后量对不上就会落空，落空后光标落去不知道哪里），而是按
 *   「单位→块」的同一趟来路直接指出位置，与滚动、视口都无关。整页只有 0 高记号的空页（断点画在流尾）
 *   没有可落的位置：落在最后一枚有内容的单位的末尾——连续纸模型里最近的就是流尾，WPS 真机落去何处**未取证**。
 *   越界的页号不猜，给 `undefined`。
 *
 * 因此这两行报的是「**本程序按画出来的行盒量出来的**页数/行数」。它与 WPS 自己的排版引擎不会逐页相同
 * （字体度量、行盒取整、页眉占位、孤行控制都不一样，这些也都**未取证**），所以这里不写「与 WPS 一致」。
 * 状态栏的 节／行／列 三格仍未接：行那一格（`KxWpsStatusRowInfoCommand`）要的是「这一页里的第几行」，
 * 用这套单位数得出，但 WPS 把它与「列」（`KxWpsStatusColumnInfoCommand`，字符在一行里的第几个）配对显示，
 * 列要的是字符级映射，是另一片；节则跟着多节排版那片走。
 */
import { twipsToPx, type Section } from "../core/docx.js";
import { CARRY_CLASS, PAGE_BORDER_CLASS } from "../core/render.js";

/**
 * 铺在纸上的一枚单位：一个行盒，或一枚不能拆的东西（表格的一行、一个只读容器）。
 * `height` 为 0 的那一枚是**断点后面没有内容**的记号：不占行，只另起一页。
 */
export type LayoutUnit = {
  height: number;
  forcePage: boolean;
};

/**
 * 量出来的单位还带着它在纸上画出来的**上沿**（视口坐标，与 `getBoundingClientRect()` 同一套）；
 * 光标→页的查找靠它。`height` 里并进来的段间空白落在**声明区间的末尾**（上沿取的是本行第一枚文字碎片的
 * 上沿，并进来的空白是这一枚之后、下一枚之前的），所以相邻两枚的声明区间首尾相接——接缝上像素级的比较
 * 分不出究竟在哪一枚里，查找要用光标矩形的**中线**（中线在行盒正中，离接缝有一整个半行高）。
 */
export type PlacedUnit = LayoutUnit & { top: number };

/** 光标落在第几页；`pages` 是同一趟量出来的总页数。 */
export type PagePlacement = {
  page: number;
  pages: number;
};

/** 「第 N 页在哪里」的答案：落光标的文档位置（折叠 Range 的起端），与它所在的块（滚动到它时用）。 */
export type PageAnchor = {
  node: Node;
  offset: number;
  /** 落点所在的块（页面的直接孩子）。空页的落点在最后一枚有内容的单位所在的块里。 */
  block: HTMLElement;
};

/** 一张连续的纸量出来的结果。 */
export type PageMeasure = {
  /** 一页能放正文的高度，px。 */
  contentHeightPx: number;
  /** 量出来的行盒枚数（表格按行计，0 高的断点记号不算）。 */
  lines: number;
  /** 这些行盒铺出来的页数。空的一列仍是一张纸。 */
  pages: number;
};

/** 行盒高度是浏览器算出来的小数，正好铺满一页时别让它多滚出一页。 */
const FIT_SLACK = 0.01;

/**
 * `w:br` 自己那句 `w:type`。按 local name 取值——前缀是文件的写法，与 `docx.ts` 里 `attrsOf` 同一条理由；
 * 两种引号都认，因为 `data-brk` 存的是元素源文本原样，而单引号也是 XML 允许的写法：认漏一枚就是把一页当成一枚换行。
 */
function breakTypeOf(xml: string): string {
  const gt = xml.indexOf(">");
  const head = gt < 0 ? xml : xml.slice(0, gt);
  const m = /(?:^|\s)(?:[\w.-]+:)?type\s*=\s*(?:"([^"]*)"|'([^']*)')/.exec(head);
  return (m?.[1] ?? m?.[2] ?? "").trim();
}

/** 这一枚 `br` 是不是「另起一页」那一族：`page` 直接是，`column` 在只画得出一个栏的纸上也是（见文件头）。 */
function forcesPage(br: HTMLElement): boolean {
  const type = breakTypeOf(br.getAttribute("data-brk") ?? "");
  return type === "page" || type === "column";
}

/** 这一段自己的行高，px。`line-height` 认不出数时退回这段画出来的高度。 */
function lineHeightOf(el: HTMLElement): number {
  const v = parseFloat(getComputedStyle(el).lineHeight);
  if (Number.isFinite(v) && v > 0) return v;
  return Math.max(1, el.getBoundingClientRect().height);
}

/**
 * 一个范围里量出来的行盒：同一行上的碎片按 `top` 归成一堆。
 * 比的是**这一堆第一条**的 `top`：同行上字号不一的碎片最多差几像素，跨行至少差一整行高。
 */
function lineBoxesOf(range: Range, lh: number): { top: number; height: number }[] {
  const rects = [...range.getClientRects()].filter((r) => r.height > 0);
  if (!rects.length) return [];
  rects.sort((a, b) => a.top - b.top || b.bottom - a.bottom);
  const out: { top: number; height: number }[] = [];
  let top = rects[0]!.top;
  let bottom = rects[0]!.bottom;
  const tol = lh * 0.8;
  const flush = (): void => {
    out.push({ top, height: Math.max(lh, bottom - top) });
  };
  for (const r of rects.slice(1)) {
    if (r.top - top <= tol) {
      bottom = Math.max(bottom, r.bottom);
      continue;
    }
    flush();
    top = r.top;
    bottom = r.bottom;
  }
  flush();
  return out;
}

/** `render.ts` 把 `w:pageBreakBefore` 写成 `page-break-before:always`，这里只认画出来的那一句。 */
function breaksBefore(el: HTMLElement): boolean {
  const cs = getComputedStyle(el);
  const stated = (v: string): boolean => /^(always|page|left|right)$/.test(v);
  return stated(cs.breakBefore) || stated(cs.pageBreakBefore);
}

/**
 * 一块段落（或标题）里的单位：段里若有另起一页的断点，就按断点切成几截分别量，
 * 断点之后那一截的第一行带着 `forcePage`；那一截是空的（断点落在段尾）时记一枚 0 高的记号。
 */
function unitsOfParagraph(el: HTMLElement): PlacedUnit[] {
  const lh = lineHeightOf(el);
  const breaks = [...el.querySelectorAll<HTMLBRElement>("br")].filter(forcesPage);
  const out: PlacedUnit[] = [];
  for (let i = 0; i <= breaks.length; i++) {
    const range = document.createRange();
    if (i === 0) range.setStart(el, 0);
    else range.setStartAfter(breaks[i - 1]!);
    if (i === breaks.length) range.setEnd(el, el.childNodes.length);
    else range.setEndBefore(breaks[i]!);
    const boxes = lineBoxesOf(range, lh);
    if (i === 0) {
      out.push(...boxes.map((b) => ({ height: b.height, forcePage: false, top: b.top })));
      continue;
    }
    // 断点之前那一截不因为后面有个断点而另起一页；断点之后那一截的第一行才是新页的开始。
    if (!boxes.length) {
      const last = out[out.length - 1];
      out.push({ height: 0, forcePage: true, top: last ? last.top + last.height : el.getBoundingClientRect().top });
    } else {
      out.push(
        { height: boxes[0]!.height, forcePage: true, top: boxes[0]!.top },
        ...boxes.slice(1).map((b) => ({ height: b.height, forcePage: false, top: b.top })),
      );
    }
  }
  return out;
}

/** 一块画出来的东西自己的上下沿（不含它与上一块之间那段空白）。 */
function unitsOfBlock(el: HTMLElement): PlacedUnit[] {
  if (el.tagName === "TABLE") {
    // `rows` 只给这张表自己的行：格子里的嵌套表属于那个表对象，不会被这里数两遍。
    return [...(el as HTMLTableElement).rows]
      .map((r) => r.getBoundingClientRect())
      .filter((b) => b.height > 0)
      .map((b) => ({ height: b.height, forcePage: false, top: b.top }));
  }
  if (el.classList.contains(CARRY_CLASS)) {
    const b = el.getBoundingClientRect();
    return [{ height: b.height, forcePage: false, top: b.top }];
  }
  const units = unitsOfParagraph(el);
  if (units.length) return units;
  const b = el.getBoundingClientRect();
  return [{ height: Math.max(1, b.height), forcePage: false, top: b.top }];
}

/**
 * 正文的开始线：纸盒上沿 + 上边框 + 上边距（`--pad-t`）。这一段之前的空白不算进任何一块。
 * 光标不在纸面上时（调用方判掉后）按这条线当作「文首」。
 */
export function contentTopOf(page: HTMLElement): number {
  const box = page.getBoundingClientRect();
  const cs = getComputedStyle(page);
  return box.top + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
}

/**
 * 纸上每一块（页面的直接孩子）量出来的单位，按文档顺序，段间空白并进后一块第一枚。页面边框那一族是画在
 * 纸盒里的装饰，不是正文的一块（`doc-edit.ts` 读回模型时同样跳过它）。`layoutUnits`（枚的序号）与
 * `pageCaretAnchor`（枚回到块里）共用这一趟，两处的铺法与条目一一对应。
 */
function unitBlocksOf(page: HTMLElement): { el: HTMLElement; units: PlacedUnit[] }[] {
  const out: { el: HTMLElement; units: PlacedUnit[] }[] = [];
  let cursor = contentTopOf(page);
  for (const child of [...page.children]) {
    const el = child as HTMLElement;
    if (el.classList.contains(PAGE_BORDER_CLASS)) continue;
    const rect = el.getBoundingClientRect();
    const gap = Math.max(0, rect.top - cursor);
    const units = unitsOfBlock(el);
    if (units.length && gap > 0) units[0] = { ...units[0]!, height: units[0]!.height + gap };
    else if (!units.length && gap > 0) units.push({ height: gap, forcePage: false, top: rect.top });
    if (units.length && breaksBefore(el)) units[0] = { ...units[0]!, forcePage: true };
    out.push({ el, units });
    cursor = Math.max(cursor, rect.bottom);
  }
  return out;
}

/** 把这张连续的纸上画出来的东西读成一列单位（行盒与不可拆的块），按文档顺序。 */
export function layoutUnits(page: HTMLElement): PlacedUnit[] {
  return unitBlocksOf(page).flatMap((b) => b.units);
}

/**
 * 纯的那一半：一列单位在 `contentHeightPx` 高的纸上各落在第几页（与 `units` 一一对应）。
 * 断点必另起一页（哪怕当前这页还空着，那正是连按两次分页符的样子）；放不下的那一行整行挪到下一页，
 * 比一页还高的一枚只占一页——纸上它溢出去，但不会多算出一页来。
 */
export function placementsOf(units: readonly LayoutUnit[], contentHeightPx: number): number[] {
  if (!(contentHeightPx > 0)) return units.map(() => 1);
  const pages: number[] = [];
  let page = 1;
  let used = 0;
  for (const u of units) {
    if (u.forcePage) {
      page++;
      used = 0;
    }
    const h = Math.min(Math.max(u.height, 0), contentHeightPx);
    if (used > 0 && used + h > contentHeightPx + FIT_SLACK) {
      page++;
      used = 0;
    }
    used += h;
    pages.push(page);
  }
  return pages;
}

/** 这一列单位铺出几页。空的一列仍是一张纸。 */
export function packPages(units: readonly LayoutUnit[], contentHeightPx: number): number {
  if (!(contentHeightPx > 0)) return 1;
  const places = placementsOf(units, contentHeightPx);
  return places.length ? Math.max(...places) : 1;
}

/** 一页能放正文的高度 = 节说的页高 − 纸上画出来的上下内边距（口径见文件头）。 */
function contentHeightOf(page: HTMLElement, sec: Section): number | undefined {
  if (!(page.getBoundingClientRect().height > 0)) return undefined;
  const cs = getComputedStyle(page);
  const padTop = parseFloat(cs.paddingTop);
  const padBottom = parseFloat(cs.paddingBottom);
  if (!Number.isFinite(padTop) || !Number.isFinite(padBottom)) return undefined;
  const contentHeightPx = twipsToPx(sec.pageHeightTwips) - padTop - padBottom;
  return contentHeightPx > 0 ? contentHeightPx : undefined;
}

/**
 * 量这一张纸。纸没画上（没布局、高度为 0）或节说的页高连边距都不够留时返回 `undefined`：
 * 读数宁可少那一行，也不报一个量不出来的数。
 */
export function measurePageLayout(page: HTMLElement, sec: Section): PageMeasure | undefined {
  const contentHeightPx = contentHeightOf(page, sec);
  if (contentHeightPx === undefined) return undefined;
  const units = layoutUnits(page);
  return {
    contentHeightPx,
    lines: units.filter((u) => u.height > 0).length,
    pages: packPages(units, contentHeightPx),
  };
}

/**
 * `caretMid`（视口坐标，光标自己那枚矩形的**中线** = 上沿 + 高/2）落在第几页。纸没画上或页高量不出来时
 * `undefined`——与 `measurePageLayout` 同一条守则。中线而不是上沿：相邻两枚单位的声明区间首尾相接
 * （见 `PlacedUnit`），而光标矩形的字体盒上沿与行盒跨度差零点几像素且方向不定，正好落在缝上；中线在行盒
 * 正中，离接缝总有半个行高。中线落在纸面之上/之下时取下（或最后一）枚——防呆，不是「按最近的算」。
 */
export function pageAtCaret(page: HTMLElement, sec: Section, caretMid: number): PagePlacement | undefined {
  const contentHeightPx = contentHeightOf(page, sec);
  if (contentHeightPx === undefined) return undefined;
  const units = layoutUnits(page);
  if (!units.length) return { page: 1, pages: 1 };
  let idx = units.findIndex((u) => u.height > 0 && caretMid >= u.top && caretMid < u.top + u.height);
  if (idx < 0) idx = units.findIndex((u) => u.top > caretMid);
  if (idx < 0) idx = units.length - 1;
  const places = placementsOf(units, contentHeightPx);
  return { page: places[idx]!, pages: places[places.length - 1]! };
}

/** 块里第一枚真文字的起点；一枚文字都没有（空段落）时退回块自己的开头——那里同样落得进光标。 */
function headPositionOf(el: HTMLElement): { node: Node; offset: number } {
  const first = document.createTreeWalker(el, NodeFilter.SHOW_TEXT).nextNode();
  return first ? { node: first, offset: 0 } : { node: el, offset: 0 };
}

/** 块里最后一枚真文字的末尾（与 `headPositionOf` 对称；空页的落点用它）。 */
function tailPositionOf(el: HTMLElement): { node: Node; offset: number } {
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  let last: Node | null = null;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) last = n;
  return last ? { node: last, offset: (last.textContent ?? "").length } : { node: el, offset: el.childNodes.length };
}

/**
 * 「第 `target` 页在哪里」的反向映射（`pageAtCaret` 反着走）：给出**这一页上第一枚单位**的落笔位置，
 * 给「定位」对话框用（`doc-goto.ts` 的 `goToPage` 拿它落光标、滚动）。这一版不经过视口取点
 * （`caretRangeFromPoint` 只认视口里的点，先滚后量对不上就会落空），而是按「枚→块」的同一趟来路直接
 * 指出文档位置，与滚动、视口都无关。枚落在哪一截就落哪一截：表格取那一行，段内另起一页切出来的截
 * 落在断点之后。整页只有 0 高记号的空页（断点画在流尾、后面什么都没有）没有可落的位置：落在最后一枚
 * 有内容的单位的末尾——连续纸模型里最近的就是流尾，WPS 真机落去何处**未取证**。
 * 越界的页号（<1、>pages、不是整数）不猜，返回 `undefined`，由调用方拒绝。
 */
export function pageCaretAnchor(page: HTMLElement, sec: Section, target: number): PageAnchor | undefined {
  const contentHeightPx = contentHeightOf(page, sec);
  if (contentHeightPx === undefined) return undefined;
  const blocks = unitBlocksOf(page);
  const units = blocks.flatMap((b) => b.units);
  const places = placementsOf(units, contentHeightPx);
  const pages = places.length ? places[places.length - 1]! : 1;
  if (!Number.isInteger(target) || target < 1 || target > pages) return undefined;
  let idx = -1;
  for (let i = 0; i < units.length; i++) {
    if (units[i]!.height > 0 && places[i] === target) {
      idx = i;
      break;
    }
  }
  const atFlowEnd = idx < 0;
  if (atFlowEnd) {
    for (let i = units.length - 1; i >= 0; i--) {
      if (units[i]!.height > 0) {
        idx = i;
        break;
      }
    }
  }
  if (idx < 0) return undefined;
  let block: HTMLElement | undefined;
  let segment = idx;
  for (const b of blocks) {
    if (segment < b.units.length) {
      block = b.el;
      break;
    }
    segment -= b.units.length;
  }
  if (!block) return undefined;
  if (atFlowEnd) return { ...tailPositionOf(block), block };
  if (block.tagName === "TABLE") {
    // 表的一枚单位是一行；0 高的行 `unitsOfBlock` 已经撇掉，这里按同一口径取第 segment 行。
    const row = [...(block as HTMLTableElement).rows].filter((r) => r.getBoundingClientRect().height > 0)[segment];
    return row ? { ...headPositionOf(row), block } : undefined;
  }
  if (segment === 0) return { ...headPositionOf(block), block };
  // 段内另起一页切出来的第 segment 截：落点就在第 segment 枚断点之后。
  const br = [...block.querySelectorAll<HTMLBRElement>("br")].filter(forcesPage)[segment - 1];
  if (!br?.parentElement) return undefined;
  return { node: br.parentElement, offset: [...br.parentElement.childNodes].indexOf(br) + 1, block };
}
