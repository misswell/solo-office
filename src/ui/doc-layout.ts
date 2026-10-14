/**
 * 「页数」与「行数」不是模型里的数，是**排版之后**才有的数。包里的 JS API 自己就把它们分在这一类：
 * `office6/addons/kwpsapiex/wps.js` 的 `getDocumentWordCount()` 把 `words`、`characters`、
 * `charactersWithSpaces`、`paragraphs`、`lines`、`pages` 全都交给 `doc.ComputeStatistics(wdStatistic…)`，
 * 而 `lines`/`pages` 那两问要的是字体度量与栏宽决定出来的行盒——第 66 轮做「字数统计」时正因为这个把这两行删了
 * （那时候它们没有受众：`render.ts` 只出一张连续的纸，照画报「页数 1」对一个四十段的件是假读数）。
 * 这一页把那个受众补上：**量的还是这张连续的纸**，按**每一节自己的** `w:pgSz`/`w:pgMar` 一节一张地切，
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
 * - 另起一页的来路有三条，都是**声明**不是几何：`br[data-brk]` 里那句源文本的 `w:type` 为 `page`、
 *   `render.ts` 由 `w:pageBreakBefore` 写出的 `page-break-before:always`，以及**节界**——带
 *   `w:pPr/w:sectPr` 的段落结束它那一节，下一节从新的一页起。`w:type` 这一版没读：语料那侧的 31 枚
 *   段级 `w:sectPr` 全部没写它（`.scratch/sect-audience.mjs`），也就是 §17.6.22 的缺省 `nextPage`；
 *   写了 `continuous` 的件真机上不在此处另起一页，本版会多算一页（未建模、未取证）。断点后面没有内容时
 *   仍然多算一页（那一页是空的）——断点说的是「从这里另起一页」，不是「后面有东西才另起」。
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
 * - **多节按各自节说的纸高量**：块 `b` 属于第 `#{e < b}` 节（`e` 是 `sectionEndsOf` 数出的节尾段落下标，
 *   0 起），每一节用它**自己那一节**的页高 − 上下边距当容量（第一节那张仍走纸上画出来的 padding，因为它
 *   就是纸上那张；其余各节按节自己的声明算，那儿没有画出来的东西可读），节界必另起一页。**纸宽没有重排**：
 *   `render.ts` 只画一张连续的纸，行盒全是按第一节的宽度折的，所以横版那几节的真实折行与真机多半不同
 *   （一个节的宽与高都换的件是语料里的多数：86 份里 18 份多节、14 份各节纸张组合不同，未取证差多少）；
 *   把连续纸切成 N 个 `.page` 盒子（下一片）时这一处连着改。
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
 * - **光标那一行／那一列**（状态栏 `KxWpsStatusRowInfoCommand` 与 `KxWpsStatusColumnInfoCommand`，kuip 里靠
 *   `rightSibling`/`leftSibling` 配成一对）走 `rowColumnAt()`：取点与「第 X 页」同一条中线（先找到枚、
 *   翻出它那一页），行 = **这一页里这一枚排第几**（从这页第一枚有内容的单位数起，0 高的断点记号不数）；
 *   列 = **字符在这一行里的第几个**，把本段文字按字符起点一枚一枚量行盒：同一行上各字符起点的上沿相同、
 *   折行处整行高地下落一次，上沿对字符位置单调不减，二分行首；数的是 Unicode 码元（代理对算两个）。
 *   列只对**段落**成立：表格与只读容器没有「本段文字」这回事，那两处只报行、列由状态栏保留上一次的读数。
 *   行的口径（每页重数而不是全篇连数）与列在行尾空白上的计法对 WPS 真机都**未取证**。
 *
 * - **光标在第几节／第 N 节在哪里**（状态栏 `KxWpsStatusSectionInfoCommand`，detail「@Page Section. Click to
 *   Open "Go To" dialog.」）是这一页里**唯一不量几何**的一对：
 *   节的边界是文件自己声明的（`w:pPr/w:sectPr` 的段落结束它那一节），从模型数（`sectionEndsOf`），落点走
 *   `sectionCaretAnchor`（节首一块的头）。页数那侧已经一节按一张纸地量了（见上一条），但**画出来**的纸
 *   仍只有第一节那张——「第 X 节」与纸盒的宽高可能对不上，这是读数与渲染之间的已知差。该格的中文字面与
 *   真机点开的种子**未取证**。
 *
 * - **光标那行离纸沿多远**（状态栏「纵向位置」，`KxWpsPagePositionInfoCommand`，kuip 452 行）走
 *   `pagePositionAt()`：与「第 X 页」同一条中线先定页，纸沿 = **本页第一枚内容的上沿 − 这一节的上边距**
 *   （第 1 页那张与 `render.ts` 的 `--pad-t` 同源，第 2 页起在流里没有 DOM，但每页内容都从纸沿下
 *   padTop 处起排，所以反推得出），光标行顶走 行/列 那一趟的**字符盒**上沿。真机量到哪条线、什么单位、
 *   几位小数**未取证**。
 *
 * 因此这两行报的是「**本程序按画出来的行盒量出来的**页数/行数」。它与 WPS 自己的排版引擎不会逐页相同
 * （字体度量、行盒取整、页眉占位、孤行控制都不一样，这些也都**未取证**），所以这里不写「与 WPS 一致」。
 */
import { DEFAULT_SECTION, sectionEndsOf, twipsToPx, type Section, type WordDocument } from "../core/docx.js";
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

/** 光标落在这一页的第几行；能数出字符级位置时还有它在这一行的第几个字（都从 1 起）。 */
export type CaretLine = {
  row: number;
  /** 字符在这一行里的序号，1 起。表格/只读容器里没有「本段文字」，不给这个数。 */
  column?: number;
};

/** 「第 N 页／第 N 节在哪里」的答案：落光标的文档位置（折叠 Range 的起端），与它所在的块（滚动到它时用）。 */
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
 * 多节铺排的一「张」纸：某一节自己的正文容量，px。
 */
export type Sheet = { contentHeightPx: number };

/** 一节自己的纸的容量：节说的页高 − 节说的上下边距。量不出正数就不给（与 `contentHeightOf` 同一条守则）。 */
export function sheetOf(section: Section): Sheet | undefined {
  const h = twipsToPx(section.pageHeightTwips - section.margins.top - section.margins.bottom);
  return h > 0 ? { contentHeightPx: h } : undefined;
}

/**
 * 纯的那一半（多节版）：一列单位在 `unitSections[i]` 说的那一节的纸上各落在第几页（与 `units` 一一对应）。
 * 每一节内的铺法与「单节」完全是同一套；**节界必另起一页**（断点是声明，哪怕上一节那一页还空着），
 * 从下一节的第一枚起换用那一节的容量。单节件就是 `sheets` 一张的特例——两处的答案必须逐枚相同。
 */
export function placementsAcrossSections(units: readonly LayoutUnit[], unitSections: readonly number[], sheets: readonly Sheet[]): number[] {
  const pages: number[] = [];
  let page = 1;
  let used = 0;
  let cur = -1;
  for (let i = 0; i < units.length; i++) {
    const s = unitSections[i] ?? 0;
    const contentHeightPx = sheets[s]!.contentHeightPx;
    if (s !== cur) {
      if (cur >= 0) {
        page++;
        used = 0;
      }
      cur = s;
    }
    const u = units[i]!;
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

/**
 * 纯的那一半（单节）：一列单位在 `contentHeightPx` 高的纸上各落在第几页（与 `units` 一一对应）。
 * 断点必另起一页（哪怕当前这页还空着，那正是连按两次分页符的样子）；放不下的那一行整行挪到下一页，
 * 比一页还高的一枚只占一页——纸上它溢出去，但不会多算出一页来。多节件写 `placementsAcrossSections`，
 * 这里是它「只有一张纸」的特例，铺法只有那一份。
 */
export function placementsOf(units: readonly LayoutUnit[], contentHeightPx: number): number[] {
  if (!(contentHeightPx > 0)) return units.map(() => 1);
  return placementsAcrossSections(
    units,
    units.map(() => 0),
    [{ contentHeightPx }],
  );
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

/** 一趟铺排要用的全部中间量：画出来的块与单位、每一枚单位落在第几页、总页数，以及画出来的那张纸的高度。 */
type PagePlan = {
  blocks: { el: HTMLElement; units: PlacedUnit[] }[];
  units: PlacedUnit[];
  /** 与 `units` 一一对应的节号（0 起）——「纵向位置」按它找这一节的上边距。 */
  unitSections: number[];
  places: number[];
  pages: number;
  /** 画出来的那张纸（第一节）的正文高度，px——`PageMeasure.contentHeightPx` 报的就是它。 */
  contentHeightPx: number;
};

/**
 * 一趟量到底：块→单位→节→页，四个读数口（页数/行数、光标第几页、行/列、第 N 页在哪）共用这一份。
 * 每枚单位的节号从模型数（`sectionEndsOf`，块 `b` 属于第 `#{e < b}` 节），每节的纸：第一节按纸上
 * 画出来的 padding（`contentHeightOf`），其余各节按节自己的声明（`sheetOf`）。纸没画上、或哪一节
 * 的容量算不出正数就整趟不给——读数宁可少一行，也不报一个量不出来的数。
 */
function planOf(page: HTMLElement, doc: WordDocument): PagePlan | undefined {
  const sections = doc.sections.length ? doc.sections : [DEFAULT_SECTION];
  const contentHeightPx = contentHeightOf(page, sections[0]!);
  if (contentHeightPx === undefined) return undefined;
  const sheets: Sheet[] = [{ contentHeightPx }];
  for (const s of sections.slice(1)) {
    const sheet = sheetOf(s);
    if (!sheet) return undefined;
    sheets.push(sheet);
  }
  const ends = sectionEndsOf(doc.blocks);
  const blocks = unitBlocksOf(page);
  const units: PlacedUnit[] = [];
  const unitSections: number[] = [];
  let passed = 0;
  blocks.forEach((b, bi) => {
    while (passed < ends.length && ends[passed]! < bi) passed++;
    for (const u of b.units) {
      units.push(u);
      unitSections.push(passed);
    }
  });
  const places = placementsAcrossSections(units, unitSections, sheets);
  return { blocks, units, unitSections, places, pages: places.length ? places[places.length - 1]! : 1, contentHeightPx };
}

/**
 * 量这一张纸。纸没画上（没布局、高度为 0）或哪一节说的页高连边距都不够留时返回 `undefined`：
 * 读数宁可少那一行，也不报一个量不出来的数。`lines` 是全部节加起来的内容行数；`pages` 一节一张地数。
 */
export function measurePageLayout(page: HTMLElement, doc: WordDocument): PageMeasure | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  return {
    contentHeightPx: plan.contentHeightPx,
    lines: plan.units.filter((u) => u.height > 0).length,
    pages: plan.pages,
  };
}

/**
 * `caretMid`（视口坐标，光标自己那枚矩形的**中线** = 上沿 + 高/2）落在第几枚。中线而不是上沿：相邻两枚
 * 单位的声明区间首尾相接（见 `PlacedUnit`），而光标矩形的字体盒上沿与行盒跨度差零点几像素且方向不定，
 * 正好落在缝上；中线在行盒正中，离接缝总有半个行高。中线落在纸面之上/之下时取下（或最后一）枚——
 * 防呆，不是「按最近的算」。`pageAtCaret` 与 `rowColumnAt` 共用这一趟。
 */
function unitIndexAt(units: readonly PlacedUnit[], caretMid: number): number {
  let idx = units.findIndex((u) => u.height > 0 && caretMid >= u.top && caretMid < u.top + u.height);
  if (idx < 0) idx = units.findIndex((u) => u.top > caretMid);
  if (idx < 0) idx = units.length - 1;
  return idx;
}

/**
 * `caretMid`（视口坐标，光标自己那枚矩形的**中线** = 上沿 + 高/2）落在第几页。纸没画上或页高量不出来时
 * `undefined`——与 `measurePageLayout` 同一条守则。取点用中线而不是上沿的理由写在 `unitIndexAt` 上。
 */
export function pageAtCaret(page: HTMLElement, doc: WordDocument, caretMid: number): PagePlacement | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  if (!plan.units.length) return { page: 1, pages: 1 };
  const idx = unitIndexAt(plan.units, caretMid);
  return { page: plan.places[idx]!, pages: plan.pages };
}

/**
 * 纯的那一半：`places`（`placementsOf` 的答案）里第 `idx` 枚在**它那一页里**排第几（1 起）。
 * 行按页重数（不是全篇连数）是本片的口径——WPS 那两格的 detail 是 "@Line Status Bar"／"@Column Status Bar"，
 * 行与列配对显示，页内序数是配对得上的一种读法；真机逐页是否重数**未取证**。0 高的断点记号不占行，不数。
 */
export function rowInPage(places: readonly number[], units: readonly LayoutUnit[], idx: number): number | undefined {
  if (idx < 0 || idx >= units.length) return undefined;
  const page = places[idx]!;
  let row = 0;
  for (let j = 0; j <= idx; j++) {
    if (places[j] === page && units[j]!.height > 0) row++;
  }
  // 取点走到 0 高的记号上时（中线落在所有内容之下的防呆路径）那一页的行算第 1 行，而不是第 0 行。
  return Math.max(1, row);
}

/** 块里拉平的文字与每个文字节点的起点（按文档顺序）。列要在这条串上**按字符**量行盒。 */
type FlatText = { len: number; starts: { node: Text; start: number }[] };

function flattenText(el: HTMLElement): FlatText {
  const walker = document.createTreeWalker(el, NodeFilter.SHOW_TEXT);
  const starts: { node: Text; start: number }[] = [];
  let len = 0;
  for (let n = walker.nextNode(); n; n = walker.nextNode()) {
    const node = n as Text;
    starts.push({ node, start: len });
    len += node.data.length;
  }
  return { len, starts };
}

/** 拉平串里第 `at` 位（0..len）的文档位置；两节点相接处取后一个节点的开头（那是同一处位置）。 */
function positionAt(flat: FlatText, at: number): { node: Node; offset: number } | undefined {
  for (let i = flat.starts.length - 1; i >= 0; i--) {
    const s = flat.starts[i]!;
    if (at >= s.start) return { node: s.node, offset: Math.min(at - s.start, s.node.data.length) };
  }
  return undefined;
}

/** 这一处文档位置画出来的上沿（折叠 Range 自己那枚矩形；与光标取点是同一条路）。 */
function topAtPoint(node: Node, offset: number): number | undefined {
  const range = document.createRange();
  try {
    range.setStart(node, offset);
    range.collapse(true);
  } catch {
    return undefined;
  }
  const own = range.getClientRects()[0];
  if (own) return own.top;
  const box = range.getBoundingClientRect();
  return box.height || box.width || box.top || box.left ? box.top : undefined;
}

/** 光标在拉平串里的位置（0..len）。光标落在元素边界上（空段落、或两枚节点之间）也算得出。 */
function flatOfCaret(block: HTMLElement, node: Node, offset: number): number | undefined {
  if (!block.contains(node)) return undefined;
  const range = document.createRange();
  try {
    range.setStart(block, 0);
    range.setEnd(node, offset);
  } catch {
    return undefined;
  }
  return range.toString().length;
}

/**
 * 光标字符位置 `at` 在这一行里排第几（1 起）。把每个字符起点的上沿排出来：同一行上各处上沿相同，
 * 折行处整行高地下落一次，所以上沿对字符位置单调不减，二分找光标那一行的首字符，列 = 位置 − 首字符 + 1。
 * 容差取半行高：行内因字号不一差几像素的不管，跨行至少差一整行高。量不出上沿时不猜，`undefined`。
 */
function columnOf(flat: FlatText, unit: PlacedUnit, at: number): number | undefined {
  if (!flat.len) return 1;
  let lo = 0;
  let hi = flat.len;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const pos = positionAt(flat, mid);
    const top = pos ? topAtPoint(pos.node, pos.offset) : undefined;
    if (top === undefined) return undefined;
    if (top >= unit.top - unit.height / 2) hi = mid;
    else lo = mid + 1;
  }
  return Math.max(1, at - lo + 1);
}

/**
 * 光标落在这一页的第几行、这一行的第几个字（状态栏 `KxWpsStatusRowInfoCommand`／`KxWpsStatusColumnInfoCommand`）。
 * 行与「第 X 页」出自同一趟取点与装箱；列只认**段落**（表格与只读容器没有「本段文字」，那两处只报行）。
 * 光标与它所在的那一枚对不上时（中线防呆落到别的块上）列也不报数——宁可缺「列」这一格，不报一个算不出的数。
 */
export function rowColumnAt(page: HTMLElement, doc: WordDocument, caret: { mid: number; node: Node; offset: number }): CaretLine | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  const { blocks, units } = plan;
  if (!units.length) return { row: 1 };
  const idx = unitIndexAt(units, caret.mid);
  const row = rowInPage(plan.places, units, idx);
  if (row === undefined) return undefined;
  let block: HTMLElement | undefined;
  let segment = idx;
  for (const b of blocks) {
    if (segment < b.units.length) {
      block = b.el;
      break;
    }
    segment -= b.units.length;
  }
  if (!block || block.tagName === "TABLE" || block.classList.contains(CARRY_CLASS)) return { row };
  const flat = flattenText(block);
  const at = flatOfCaret(block, caret.node, caret.offset);
  if (at === undefined || at > flat.len) return { row };
  const column = columnOf(flat, units[idx]!, at);
  return column === undefined ? { row } : { row, column };
}

/**
 * 光标那行离**纸沿**的纵向距离，px——状态栏「纵向位置」那一格（WPS 的 `KxWpsPagePositionInfoCommand`，
 * kuip 452 行，detail「@Vertical Page Position. Click to Open "Go To" dialog.」）。纸沿是这一页纸的上沿：
 * 第 1 页在纸上画得出来（`contentTopOf`），第 2 页起在流里没有 DOM——但模型里每一页的内容都从纸沿下
 * padTop 处起排（第 1 节与 `render.ts` 的 `--pad-t` 同源，其余各节按节自己的 `w:pgMar@w:top`），
 * 所以反推得出来：纸沿 = **本页第一枚内容的上沿 − 这一节的上边距**。
 * 光标行顶走 行/列 那一趟取点（折叠 Range 自己那枚矩形）；页的归属与「第 X 页」同一条中线（`unitIndexAt`）。
 * 纸上没有内容、或光标行量不出矩形时不猜，`undefined`。取的是**字符盒**的上沿（不是行盒），与 行/列 的
 * 列同一口径；WPS 真机从纸沿量到哪条线、什么单位、几位小数**都未取证**。
 */
export function pagePositionAt(page: HTMLElement, doc: WordDocument, caret: { mid: number; node: Node; offset: number }): number | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  if (!plan.units.length) return undefined;
  const idx = unitIndexAt(plan.units, caret.mid);
  const first = plan.places.indexOf(plan.places[idx]!);
  const top = topAtPoint(caret.node, caret.offset);
  if (top === undefined) return undefined;
  const sections = doc.sections.length ? doc.sections : [DEFAULT_SECTION];
  const sec = sections[plan.unitSections[first]!];
  if (!sec) return undefined;
  return top - plan.units[first]!.top + twipsToPx(sec.margins.top);
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
 * 指出文档位置，与滚动、视口都无关——页从哪儿切来与 `pageAtCaret` 同一份（`planOf`，多节各按自己的
 * 页高切）。枚落在哪一截就落哪一截：表格取那一行，段内另起一页切出来的截落在断点之后。整页只有 0 高
 * 记号的空页（断点画在流尾、后面什么都没有）没有可落的位置：落在最后一枚有内容的单位的末尾——连续纸
 * 模型里最近的就是流尾，WPS 真机落去何处**未取证**。
 * 越界的页号（<1、>pages、不是整数）不猜，返回 `undefined`，由调用方拒绝。
 */
export function pageCaretAnchor(page: HTMLElement, doc: WordDocument, target: number): PageAnchor | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  const { blocks, units, places } = plan;
  const pages = plan.pages;
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

/**
 * 「第 `target` 节在哪里」（`pageCaretAnchor` 的姊妹）：`ends` 是节尾段落（带 `w:pPr/w:sectPr` 的段落）
 * 在块里的下标、`total` 是节数——两者都由模型那侧数出来（`docx.ts` 的 `sectionEndsOf` 与 `ends.length + 1`，
 * 与状态栏「节」格同一份）。返回这一节**第一块**的落笔位置：第一枚真文字（`headPositionOf`），表格取第一行
 * （`pageCaretAnchor` 同一口径）。块的编号就是 `.page` 孩子的编号——`render.ts` 把页面边框排在最后，正是
 * 为了这个编号不挪位。这里**不量**任何几何：节的边界是文件声明的，与铺排没关系。越界的节号不猜，`undefined`。
 */
export function sectionCaretAnchor(page: HTMLElement, ends: readonly number[], total: number, target: number): PageAnchor | undefined {
  if (!Number.isInteger(target) || target < 1 || target > total) return undefined;
  if (target > 1 && ends[target - 2] === undefined) return undefined;
  const block = page.children[target === 1 ? 0 : ends[target - 2]! + 1] as HTMLElement | undefined;
  if (!block || block.classList.contains(PAGE_BORDER_CLASS)) return undefined;
  if (block.tagName === "TABLE") {
    const row = [...(block as HTMLTableElement).rows][0];
    return row ? { ...headPositionOf(row), block } : undefined;
  }
  return { ...headPositionOf(block), block };
}
