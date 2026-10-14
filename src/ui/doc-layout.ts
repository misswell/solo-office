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
 *
 * 因此这两行报的是「**本程序按画出来的行盒量出来的**页数/行数」。它与 WPS 自己的排版引擎不会逐页相同
 * （字体度量、行盒取整、页眉占位、孤行控制都不一样，这些也都**未取证**），所以这里不写「与 WPS 一致」。
 * 状态栏的 页码/页次/行/列 四格仍未接：那几个数要的是**光标所在**那一页那一行，是把这套行盒接着映射到光标位的
 * 下一片（`doc-status-bar.ts` 头注记的是同一件）。
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
 * 一个范围里量出来的行盒高度：同一行上的碎片按 `top` 归成一堆。
 * 比的是**这一堆第一条**的 `top`：同行上字号不一的碎片最多差几像素，跨行至少差一整行高。
 */
function lineBoxesOf(range: Range, lh: number): number[] {
  const rects = [...range.getClientRects()].filter((r) => r.height > 0);
  if (!rects.length) return [];
  rects.sort((a, b) => a.top - b.top || b.bottom - a.bottom);
  const out: number[] = [];
  let top = rects[0]!.top;
  let bottom = rects[0]!.bottom;
  const tol = lh * 0.8;
  const flush = (): void => {
    out.push(Math.max(lh, bottom - top));
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
function unitsOfParagraph(el: HTMLElement): LayoutUnit[] {
  const lh = lineHeightOf(el);
  const breaks = [...el.querySelectorAll<HTMLBRElement>("br")].filter(forcesPage);
  const out: LayoutUnit[] = [];
  for (let i = 0; i <= breaks.length; i++) {
    const range = document.createRange();
    if (i === 0) range.setStart(el, 0);
    else range.setStartAfter(breaks[i - 1]!);
    if (i === breaks.length) range.setEnd(el, el.childNodes.length);
    else range.setEndBefore(breaks[i]!);
    const boxes = lineBoxesOf(range, lh);
    if (i === 0) {
      out.push(...boxes.map((height) => ({ height, forcePage: false })));
      continue;
    }
    // 断点之前那一截不因为后面有个断点而另起一页；断点之后那一截的第一行才是新页的开始。
    if (!boxes.length) out.push({ height: 0, forcePage: true });
    else out.push({ height: boxes[0]!, forcePage: true }, ...boxes.slice(1).map((height) => ({ height, forcePage: false })));
  }
  return out;
}

/** 一块画出来的东西自己的上下沿（不含它与上一块之间那段空白）。 */
function unitsOfBlock(el: HTMLElement): LayoutUnit[] {
  if (el.tagName === "TABLE") {
    // `rows` 只给这张表自己的行：格子里的嵌套表属于那个表对象，不会被这里数两遍。
    return [...(el as HTMLTableElement).rows]
      .map((r) => r.getBoundingClientRect().height)
      .filter((h) => h > 0)
      .map((height) => ({ height, forcePage: false }));
  }
  if (el.classList.contains(CARRY_CLASS)) {
    return [{ height: el.getBoundingClientRect().height, forcePage: false }];
  }
  const units = unitsOfParagraph(el);
  return units.length ? units : [{ height: Math.max(1, el.getBoundingClientRect().height), forcePage: false }];
}

/**
 * 把这张连续的纸上画出来的东西读成一列单位（行盒与不可拆的块），按文档顺序，段间空白并进后一块第一枚。
 * 页面边框那一族是画在纸盒里的装饰，不是正文的一块（`doc-edit.ts` 读回模型时同样跳过它）。
 */
export function layoutUnits(page: HTMLElement): LayoutUnit[] {
  const out: LayoutUnit[] = [];
  const box = page.getBoundingClientRect();
  const cs = getComputedStyle(page);
  // 正文的开始线：纸盒上沿 + 上边框 + 上边距（`--pad-t`）。这一段之前的空白不算进任何一块。
  let cursor = box.top + (parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0);
  for (const child of [...page.children]) {
    const el = child as HTMLElement;
    if (el.classList.contains(PAGE_BORDER_CLASS)) continue;
    const rect = el.getBoundingClientRect();
    const gap = Math.max(0, rect.top - cursor);
    const units = unitsOfBlock(el);
    if (units.length && gap > 0) units[0] = { ...units[0]!, height: units[0]!.height + gap };
    else if (!units.length && gap > 0) units.push({ height: gap, forcePage: false });
    if (units.length && breaksBefore(el)) units[0] = { ...units[0]!, forcePage: true };
    out.push(...units);
    cursor = Math.max(cursor, rect.bottom);
  }
  return out;
}

/**
 * 纯的那一半：一列单位在 `contentHeightPx` 高的纸上铺出几页。
 * 断点必另起一页（哪怕当前这页还空着，那正是连按两次分页符的样子）；放不下的那一行整行挪到下一页，
 * 比一页还高的一枚只占一页——纸上它溢出去，但不会多算出一页来。
 */
export function packPages(units: readonly LayoutUnit[], contentHeightPx: number): number {
  if (!(contentHeightPx > 0)) return 1;
  let pages = 1;
  let used = 0;
  for (const u of units) {
    if (u.forcePage) {
      pages++;
      used = 0;
    }
    const h = Math.min(Math.max(u.height, 0), contentHeightPx);
    if (used > 0 && used + h > contentHeightPx + FIT_SLACK) {
      pages++;
      used = 0;
    }
    used += h;
  }
  return pages;
}

/**
 * 量这一张纸。纸没画上（没布局、高度为 0）或节说的页高连边距都不够留时返回 `undefined`：
 * 读数宁可少那一行，也不报一个量不出来的数。
 */
export function measurePageLayout(page: HTMLElement, sec: Section): PageMeasure | undefined {
  if (!(page.getBoundingClientRect().height > 0)) return undefined;
  const cs = getComputedStyle(page);
  const padTop = parseFloat(cs.paddingTop);
  const padBottom = parseFloat(cs.paddingBottom);
  if (!Number.isFinite(padTop) || !Number.isFinite(padBottom)) return undefined;
  const contentHeightPx = twipsToPx(sec.pageHeightTwips) - padTop - padBottom;
  if (!(contentHeightPx > 0)) return undefined;
  const units = layoutUnits(page);
  return {
    contentHeightPx,
    lines: units.filter((u) => u.height > 0).length,
    pages: packPages(units, contentHeightPx),
  };
}
