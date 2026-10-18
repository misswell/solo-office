/**
 * 「页数」与「行数」不是模型里的数，是**排版之后**才有的数。包里的 JS API 自己就把它们分在这一类：
 * `office6/addons/kwpsapiex/wps.js` 的 `getDocumentWordCount()` 把 `words`、`characters`、
 * `charactersWithSpaces`、`paragraphs`、`lines`、`pages` 全都交给 `doc.ComputeStatistics(wdStatistic…)`，
 * 而 `lines`/`pages` 那两问要的是字体度量与栏宽决定出来的行盒——第 66 轮做「字数统计」时正因为这个把这两行删了
 * （那时候它们没有受众：`render.ts` 只出一张连续的纸，照画报「页数 1」对一个四十段的件是假读数）。
 * 这一页把那个受众补上：**量的还是这张连续的纸**，按**每一节自己的** `w:pgSz`/`w:pgMar` 一节一张地切，
 * 于是对话框那两行与状态栏那几格读的是同一个量出来的数。纸本身仍画成一张连续的白纸：页与页之间由
 * `applySheetGaps` 垫出空档、画出条带（第一片「页缝与纸层」，见该函数的注释）。
 *
 * 量的口径都在这里，别处不另定一份：
 *
 * - 一张纸能放正文的高度 = `w:pgSz@w:h` 换 px − 上边距 − 下边距。两个边距取的是**纸上算出来的 padding**
 *   （`render.ts` 由节写成 `--pad-t`/`--pad-b`，`applyBandYield` 再往上加带子让出的那一截
 *   `--band-yield-t/-b`），所以装订线、纸张方向、横竖版与页眉页脚的高低都在这一处自动跟着走。
 *   纸盒自己的 `min-height:1000px` 不参与：那是画出来好看，不是节说的纸高。
 * - `w:header`/`w:footer` 那两个距离**进这个式子**：带子贴在纸内（`render.ts` 的 `.solo-hf-master`），
 *   越过边距线的那一截从正文里占走——真机量的就是这条（`/tmp/tall-header.docx`：页眉 10 行、四边 1440 缇、
 *   `w:header` 720 缇 → 正文上沿 ≈3445 缇，#159）。带子比整页还高时未取证，本版不裁。
 * - 行盒从**画出来的盒子**里量：`Range.getClientRects()` 给的是每个文字碎片的盒子，同一行上的碎片按 `top`
 *   归堆（容差取本段自己那条 `line-height` 的八成：跨行至少差一整行高，同行上两段字号不同也就差几像素）。
 *   一堆算一行，**这一行占多高按它在纸上推进多少算**（本堆上沿到下一堆上沿，最后一堆到这一截在纸上的下沿），
 *   不小于本段的行高——浏览器的行盒可以比 `line-height` 那句声明高（字体内容盒顶出来的，见 `lineBoxesOf`
 *   的注），只按声明 charging 会一路累出纸的底边距。段与段之间真正画出来的空白（`w:spacing` 的
 *   before/after 在 CSS 里合成后的那一截）并进下一段第一枚单位里，所以流的总高与纸上画的一致。
 * - 段落只在**它自己的行盒之间**断开跨页；表格按**整行**铺（一行算一行，格子里换了几行不改这个数），
 *   只读容器（`.doc-carry`）整块算一枚。`w:cantSplit` 这一版没读，所以「这行不许拆」的文件与没写的同样铺。
 * - 另起一页的来路有三条，都是**声明**不是几何：`br[data-brk]` 里那句源文本的 `w:type` 为 `page`、
 *   `render.ts` 由 `w:pageBreakBefore` 写出的 `page-break-before:always`，以及**节界**——带
 *   `w:pPr/w:sectPr` 的段落结束它那一节，下一节从新的一页起。`w:type` 这一版没读：授权口径 46 份里段级
 *   `w:sectPr` 12 枚（8 件），而整批里**任何** `<w:type>` 元素都是 0 枚（`.scratch/oneoff-gap.mjs`、
 *   `.scratch/corpus-recheck.mjs` #4），也就是 §17.6.22 的缺省 `nextPage`；
 *   写了 `continuous` 的件真机上不在此处另起一页，本版会多算一页（未建模、未取证）。断点后面没有内容时
 *   仍然多算一页（那一页是空的）——断点说的是「从这里另起一页」，不是「后面有东西才另起」。
 * - **段尾断点多画出来的那几行空白既不计行也不占高**：断点落在段落末尾时浏览器给那一段多画空行，但**最后一条不画**
 *   （`.scratch/br-height-dom.ts` 一枚一枚按出来：1 枚→1.00 行、2 枚→2.00、3 枚→3.00、4 枚→4.00、5 枚→5.00，
 *   也就是 N 枚多画 N−1 条）。这一截空白落在段落的盒子高度里、不落在任何一枚行盒上（`Range.getClientRects()`
 *   给只有 `br` 的那一截零个矩形），所以这里记成一枚 0 高的记号：「翻页，但不占行」。于是行数正好是内容行数，
 *   而页数少算了那 N−1 条空白的高度（默认 A4 一行 25.92px、一页正文 930.53px 放得下 35 行），偏乐观。
 *   WPS 的 `ComputeStatistics` 在这种形状上各报什么**未取证**；画页缝（`applySheetGaps`）时这一处连着改：
 *   垫片的高度按真实几何让位（那几行空白落在垫距里），但它们仍不进单位，页数那一侧照旧偏乐观。
 *   `.scratch/pagination-dom.ts` 按这条口径钉行数，独立来路扣的是「枚数−1」而不是「枚数」。
 * - **分栏符是一条推导**：这一版模型没有 `w:cols` 字段，画出来就是一栏；一栏的下一栏起点就是下一页同一处，
 *   所以 `w:type="column"` 也按另起一页铺。授权口径那 46 份的量是「55 枚 `w:cols`，`w:num` 只有 `1` 一档」
 *   （`.scratch/corpus-recheck.mjs` #8；`w:br w:type="column"` 本身 0 枚，`.scratch/oneoff-gap.mjs`），
 *   单栏是这些件的常态；WPS 真机上分栏符怎么翻页**未取证**。
 * - **多节按各自节说的纸高量**：块 `b` 属于第 `#{e < b}` 节（`e` 是 `sectionEndsOf` 数出的节尾段落下标，
 *   0 起），每一节用它**自己那一节**的页高 − 上下边距当容量（第一节那张仍走纸上画出来的 padding，因为它
 *   就是纸上那张；其余各节按节自己的声明算，那儿没有画出来的东西可读），节界必另起一页。**纸宽没有重排**：
 *   `render.ts` 只画一张连续的纸，行盒全是按第一节的宽度折的，所以横版那几节的真实折行与真机多半不同
 *   （授权口径 46 份里 8 件多节，而这 8 件**各节纸张组合全都相同**——`w:pgSz` 的字面一枚都没换过，
 *   `.scratch/corpus-recheck.mjs` #9：宽与高一起换的那一类件在授权语料里没有观众，差多少仍**未取证**）；
 *   画页缝（`applySheetGaps`）时这一处**仍未改**：每张纸按各自节的高让位，宽还是第一节那张的宽。
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
 *   `pagePositionAt()`：与「第 X 页」同一条中线先定页，纸沿 = **本页第一枚内容的块沿 − 这一节的上边距 − 让位**
 *   （块沿与单位自己报的行盒顶之间那一截走 `blockInsetAbove`，与切纸推距同一支；第 1 页那张与 `render.ts` 的
 *   `--pad-t` 同源，第 2 页起在流里没有 DOM，但每页内容都从纸沿下「边距 + 让位」处起排，所以反推得出），
 *   光标行顶走 行/列 那一趟的**字符盒**上沿。真机量到哪条线、什么单位、几位小数**未取证**。
 *
 * 因此这两行报的是「**本程序按画出来的行盒量出来的**页数/行数」。它与 WPS 自己的排版引擎不会逐页相同
 * （字体度量、行盒取整、页眉占位、孤行控制都不一样，这些也都**未取证**），所以这里不写「与 WPS 一致」。
 */
import { DEFAULT_SECTION, sectionEndsOf, twipsToPx, type Section, type WordDocument } from "../core/docx.js";
import { BAND_MASTER_CLASS, BAND_SOURCE_CLASS, CARRY_CLASS, PAGE_BORDER_CLASS } from "../core/render.js";

/**
 * 铺在纸上的一枚单位：一个行盒，或一枚不能拆的东西（表格的一行、一个只读容器）。
 * `height` 为 0 的那一枚是**断点后面没有内容**的记号：不占行，只另起一页。
 */
export type LayoutUnit = {
  height: number;
  forcePage: boolean;
  /** 段尾这一枚与下一块的第一枚**同页**（段落说 `w:keepNext`，量的时候盖在末一枚上）。 */
  keepNext?: boolean;
  /** 段首这一枚身上的**整段高度**（段落说 `w:keepLines`）：这一坨不拆页（比一页还高时不算数）。 */
  keepLines?: number;
  /**
   * 这一枚在**所属段落行盒序列**里的枚序与总枚数（段落没说「关孤行控制」时才有）。装箱那趟用它把落在段中
   * 的页界挪到不落单行的地方：段首行不许独自留在页尾，段尾行不许独自落在页头。
   */
  widowLine?: { index: number; total: number };
};

/**
 * 量出来的单位还带着它在纸上画出来的**上沿**（视口坐标，与 `getBoundingClientRect()` 同一套）；
 * 光标→页的查找靠它。`height` 说的是这一枚在流里**吃掉多高**（`lineBoxesOf` 按画出来的推进量、
 * `unitBlocksOf` 把段间空白并到这一块的第一枚上），对这一块的第一枚来说那截空白画在它的上沿**之上**——
 * 所以 `top + height` 不是这一枚的地盘界线，相邻两枚的「声明区间」会互相压过头。查光标落在哪一枚要按
 * **画出来的上沿**切带（`unitIndexAt`），中线取点则保证离上下两条界都有半个行高。
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

/** 页缝里垫的那段空白（`applySheetGaps` 插进段落里的行内块级 span，空元素、不携带文字）。 */
export const SHEET_GAP_CLASS = "solo-sheet-gap";

/**
 * 页缝落在**两枚表行之间**时垫的那一行（#158）：一行一个空格子，无框、不接鼠标、不可编辑，高就是这一道的推距。
 *
 * 它与段落里那枚垫片同职——只是「缝」落在表行之间时，块级垫片既进不了 `<table>` 的肚子里（HTML 只允许
 * `tr` 做它的孩子），也不该把整张表推开（那会把缝画在表中间）。垫一行是唯一能在两行之间真正让出高度的落笔。
 *
 * 它是**画的装饰，不是表的一枚行**，所以三道口径都不数它：`tableRowsOf`（铺排量的每一枚行）、
 * `doc-edit.ts` 的 `rowsOf`（读回模型）、`main.ts` 冒烟里的行数读数。写进件里的 `<w:tr>` 枚数因此与垫之前
 * 逐枚相同——`sheets-dom` 那一条按重拼后的 document.xml 数 `w:tr` 钉住这一条。
 */
export const SHEET_GAP_ROW_CLASS = "solo-sheet-gap-row";

/** 两张纸之间留出的空档，px。WPS 真机留多宽**未取证**，这一版自定 24。 */
export const SHEET_GAP_PX = 24;

/** 页缝处盖住空档的条带：纸盒里的绝对定位装饰，排在正文块之后（编辑器按位置数块，装饰得垫底）。 */
export const SHEET_STRIP_CLASS = "solo-sheet-strip";

/** 一条页缝：条带（盖住空档的那一段）在**纸盒本地的纸的像素**里的上沿与高度（直接写 CSS 长度，不再除 zoom），以及它收口的那一页与下一页各属于第几节。 */
export type SheetStrip = { top: number; height: number; page: number; section: number; nextSection: number };

/**
 * 页眉页脚带在**第二张纸往后**的复件（`doc-edit.ts` 的 `applySheets` 画）：纸盒里的绝对定位装饰，
 * 与页面边框、条带同一族——编辑器按位置数块，装饰不能占一块。带子自己只在纸盒外画一份（首页那份，
 * 可点、可改），其余每张纸上的是它照着的复件：不可点、不进模型、不进保存。
 */
export const BAND_CLONE_CLASS = "solo-hf-clone";

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

/**
 * 这一张纸被放大到多少（`doc-edit.ts` 的 `setZoom` 写在纸盒的 `style.zoom` 上，缺省 1）。
 *
 * 纸上同时存在两套像素，这一趟量出来的每一个数都属意在其中一套里，比长短之前必须先对齐：
 * - **视口像素**：`getBoundingClientRect()` 与 `Range.getClientRects()` 报的、以及 `caretMid` 那种光标自己
 *   的矩形——`zoom` 过的纸上是真实屏幕距离，放大后比纸的像素大 zoom 倍。
 * - **纸的像素**：`twipsToPx` 从节算出来的页高/边距、`getComputedStyle` 报的 `padding`/`line-height`、
 *   以及最后写进 CSS 的那些长度——它们不随 `zoom` 变，纸放大时跟着纸一起放大。
 *
 * 规矩按「谁开口」定，不按「谁算」定：
 * - 与**量出来的视口读数**比长短的那一节（行高、正文起点、装箱用的纸容量、纵向位置里量出来的那一截），
 *   声明长度先乘 `zoom` 进视口空间。这一片的体重与首尾沿都从画出来的盒子里量，装箱就在视口空间里比。
 * - 交出去当**纸上的一截长度**（`pagePositionAt` 的读数、`PageMeasure.contentHeightPx`）或要写进 CSS 的
 *   （页缝条带、复件摆位、垫片高度 `SheetStrip`），保持纸的像素——`applySheetGaps` 把量回来的视口读数
 *   按纸盒上沿一次除回纸的像素，那条链自己算，不在这规矩外另开一口。
 */
function zoomOf(page: HTMLElement): number {
  return parseFloat(page.style.getPropertyValue("zoom")) || 1;
}

/**
 * 这一段自己的行高，按**视口像素**报（调用方拿它跟画出来的行盒比长短）。
 * `line-height` 那句声明是纸的像素，缩放过的纸要乘 `zoom` 才是画出来的那一格高；认不出数时退回量这段自己的
 * 盒子——那条读数本来就是视口像素，不能再乘一遍。
 */
function lineHeightOf(el: HTMLElement, zoom: number): number {
  const v = parseFloat(getComputedStyle(el).lineHeight);
  if (Number.isFinite(v) && v > 0) return v * zoom;
  return Math.max(1, el.getBoundingClientRect().height);
}

/**
 * 一堆行盒（同一条线上的碎片归成一堆）加上这一截在纸上的流上下沿，算**每一枚占多高**。
 * 归堆比的是**这一堆第一条**的 `top`：同行上字号不一的碎片最多差几像素，跨行至少差一整行高。
 *
 * 一枚行盒占多高，问的是**它在纸上推进了多少**：本堆上沿到下一堆上沿，最后一堆收到「这一截在纸上的总推进
 * 减去前面各枚」。不是问 `line-height` 那句声明，也不是问墨水洒到哪：11 磅的中文行在 `.scratch/lh-drift-dom.ts`
 * 里量到「声明 23.76px、块自己的推进 24.77px」——声明够矮时行盒跟着字体内容高长，所以不能按声明 charging。
 * 反过来，312 缇的网格上一枚量到「声明 20.8px、块推进 20.8px、墨顶比格顶还高 1px」（同一份件在
 * `.scratch/para-gap-dom.ts` 的 `原始账`：块上 206 而墨顶 205）——那 1px 是字自己洒进上一格的墨水，不花流里的
 * 地。按碎片边沿 charging 时每枚量到 21.8px，930.53px 的一页只装 42 行而不是纸上 real 有的 44.7 格，而吸网格
 * 的段落正是授权语料里最大的那一档（正文部件 7,398 枚 `w:p` 里 5,879 枚落在声明了行网格的 39 件下——件级口径，
 * 按一份件的首节判，未按每一节精确）。两件事在这里合成一条：
 * **起点用流上沿（`from`）而不是碎片上沿，末枚收到 `to − from − 前面各枚`；`from`/`to` 不可信（下一个 `br`
 * 零高量不到自己那行）时才退回碎片边沿**。
 * 一枚差 1px 看着不要紧，六十段就累出一张纸——装箱与 `pageAtCaret` 读的是同一列画出来的位置，两处必须
 * 是同一个数。
 */
export function lineCharges(
  runs: readonly { top: number; bottom: number }[],
  lh: number,
  from: number,
  to: number,
): number[] {
  // 第一枚那一格的高从这一截在纸上的**流上沿**（`from`）起算：`from` 比碎片上沿高时，那截空白也是这段占的。
  const starts = runs.map((g, i) => (i === 0 && Number.isFinite(from) ? from : g.top));
  const heights = runs.map((g, i) => {
    const start = starts[i]!;
    const lastBottom = i + 1 < runs.length ? runs[i + 1]!.top : Number.isFinite(to) ? to : g.bottom;
    return Math.max(lh, lastBottom - start);
  });
  // 末枚不量「最后一行的墨水上沿到这一截的下沿」，量「这一截在流里的总推进减去前面各枚」：中文行的墨顶比
  // 它那一格的上沿高 1px（206 的格量到 205 的墨顶，`para-gap-dom` 的 `原始账`），拿墨顶当起点就多出 1px，
  // 而三行的段子里前面每一枚都按墨顶到墨顶量——只有末枚把这一路的偏差都吃进去。收在 `to − from` 上，
  // 一段的枚高之和恒等于它画出来的推进，装箱读的就是纸上走的数。
  if (runs.length > 1 && Number.isFinite(from) && Number.isFinite(to)) {
    const rest = heights.slice(0, -1).reduce((a, b) => a + b, 0);
    heights[heights.length - 1] = Math.max(lh, to - from - rest);
  }
  return heights;
}

/** 一个范围里量出来的行盒：同一行上的碎片按 `top` 归成一堆，每枚占多高见 `lineCharges`。 */
function lineBoxesOf(range: Range, lh: number, from: number, to: number): { top: number; height: number }[] {
  const rects = [...range.getClientRects()].filter((r) => r.height > 0);
  if (!rects.length) return [];
  rects.sort((a, b) => a.top - b.top || b.bottom - a.bottom);
  const runs: { top: number; bottom: number }[] = [];
  const tol = lh * 0.8;
  for (const r of rects) {
    const cur = runs[runs.length - 1];
    if (cur && r.top - cur.top <= tol) {
      cur.bottom = Math.max(cur.bottom, r.bottom);
      continue;
    }
    runs.push({ top: r.top, bottom: r.bottom });
  }
  const heights = lineCharges(runs, lh, from, to);
  // **报出来的上沿仍是碎片的**：切纸的推距（`applySheetGaps`）与「纵向位置」那一格都要把内容对齐到纸沿，
  // 取点换了 1px 就等于把每一张纸的第一行都往上挪 1px。
  return runs.map((g, i) => ({ top: g.top, height: heights[i]! }));
}

/** `render.ts` 把 `w:pageBreakBefore` 写成 `page-break-before:always`，这里只认画出来的那一句。 */
function breaksBefore(el: HTMLElement): boolean {
  const cs = getComputedStyle(el);
  const stated = (v: string): boolean => /^(always|page|left|right)$/.test(v);
  return stated(cs.breakBefore) || stated(cs.pageBreakBefore);
}

/**
 * 一块段落（或标题）里的单位：段里若有另起一页的断点、或有页缝垫片，就按它们切成几截分别量。
 * 断点之后那一截的第一行带着 `forcePage`；那一截是空的（断点落在段尾）时记一枚 0 高的记号。
 * 垫片不是断点也不是内容：它那一截之前的行原样，之后的行从新的一截接着量——垫片自己那一段盒子的
 * `Range.getClientRects()` 会当成一枚矩形报出来（`.scratch/gap-probe.html` 量到 120px 的假行盒），
 * 所以量程必须把它排除在外，靠的就是这一刀。
 */
function unitsOfParagraph(el: HTMLElement, zoom: number): PlacedUnit[] {
  const lh = lineHeightOf(el, zoom);
  const cuts = [...el.querySelectorAll<HTMLElement>(`br, .${SHEET_GAP_CLASS}`)].filter(
    (n) => n.tagName !== "BR" || forcesPage(n as unknown as HTMLBRElement),
  );
  // 每一截在纸上的上下沿：第一截从**块自己的盒子外沿**起，往后各截从前一枚切断元素的下沿起；末尾收到下一枚
  // 切断元素的上沿，最后一截收到块的外沿。`br` 画出来是零高的（量不到自己那一行），不当边界。
  //
  // 外沿而不是内容沿：`w:pBdr` 那条线和 `w:space` 那截间距是这一段在流里真实占掉的地方（`w:space` 量的正是它与
  // 字之间的距离，#144 把它画成这一段的 `padding`），而 `unitBlocksOf` 的游标按 `rect.bottom` 往前推。这里若把
  // 两头的线与间距扣掉，那截地方就两头都不认：`unitBlocksOf` 已经走过去，枚高里却没有，于是页少算、状态栏的
  // 页码/页数、「定位」与「纵向位置」全都跟着错（`.scratch/para-box-dom.ts` 第 20 条量的就是这一条：同 40 段、
  // 只差那一圈，纸数必须变多）。没画框的段落两个沿是同一个数，#132 那套充电量法不动。
  // 已知少一处：段里被页缝垫片切开、且垫片后面没有字时（一段正好跨在页边界上），垫片以下那两截线与间距还没有
  // 主人——它落在最后一截之外，而那一截量不出行盒。授权口径里正文段落自己写 `w:pBdr` 的是 0 枚（26 枚全在
  // `styles.xml` 与页眉部件，`.scratch/corpus-recheck2.mjs`），先记下不动。
  const box = el.getBoundingClientRect();
  const cutBoxes = cuts.map((c) => c.getBoundingClientRect());
  const out: PlacedUnit[] = [];
  for (let i = 0; i <= cuts.length; i++) {
    const range = document.createRange();
    if (i === 0) range.setStart(el, 0);
    else range.setStartAfter(cuts[i - 1]!);
    if (i === cuts.length) range.setEnd(el, el.childNodes.length);
    else range.setEndBefore(cuts[i]!);
    const prev = i > 0 ? cutBoxes[i - 1]! : undefined;
    const next = i < cuts.length ? cutBoxes[i]! : undefined;
    const from = prev ? (prev.height > 0 ? prev.bottom : NaN) : box.top;
    const to = next ? (next.height > 0 ? next.top : NaN) : box.bottom;
    const boxes = lineBoxesOf(range, lh, from, to);
    if (i === 0) {
      out.push(...boxes.map((b) => ({ height: b.height, forcePage: false, top: b.top })));
      continue;
    }
    const cut = cuts[i - 1]!;
    if (cut.tagName !== "BR") {
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

/**
 * 这张表自己的行，按页写出来的顺序（`#158`）：页缝垫下的那一枚 `SHEET_GAP_ROW_CLASS` **不是表的一枚行**，
 * 从这里就撇掉——量、取点、落笔三个口径都必须与「没有缝的时候」逐枚同一份，否则垫一行会把后面的页界整体推歪。
 * 嵌套在格子里的表属于那个表对象，`rows` 本来也不会给过来。
 */
function tableRowsOf(table: HTMLTableElement): HTMLTableRowElement[] {
  return [...table.rows].filter((r) => !r.classList.contains(SHEET_GAP_ROW_CLASS));
}

/** 表里量得出高度的行（0 高的行不产一枚单位，与 `unitsOfBlock` 同一口径），带上各自的矩形。 */
function measuredRows(table: HTMLTableElement): { el: HTMLTableRowElement; rect: DOMRect }[] {
  const out: { el: HTMLTableRowElement; rect: DOMRect }[] = [];
  for (const r of tableRowsOf(table)) {
    const rect = r.getBoundingClientRect();
    if (rect.height > 0) out.push({ el: r, rect });
  }
  return out;
}

/** 一块画出来的东西自己的上下沿（不含它与上一块之间那段空白）。枚的高与沿都按**视口像素**报。 */
function unitsOfBlock(el: HTMLElement, zoom: number): PlacedUnit[] {
  if (el.tagName === "TABLE") {
    // `rows` 只给这张表自己的行：格子里的嵌套表属于那个表对象，不会被这里数两遍。
    return measuredRows(el as HTMLTableElement).map(({ rect }) => ({ height: rect.height, forcePage: false, top: rect.top }));
  }
  if (el.classList.contains(CARRY_CLASS)) {
    const b = el.getBoundingClientRect();
    return [{ height: b.height, forcePage: false, top: b.top }];
  }
  const units = unitsOfParagraph(el, zoom);
  if (units.length) return units;
  const b = el.getBoundingClientRect();
  return [{ height: Math.max(1, b.height), forcePage: false, top: b.top }];
}

/**
 * 正文的开始线：纸盒上沿 + 上边框 + 上边距（`--pad-t`）。这一段之前的空白不算进任何一块。
 * 光标不在纸面上时（调用方判掉后）按这条线当作「文首」。
 * 报的是**视口像素**（与每一枚枚的 `top` 同一个空间，`unitBlocksOf` 拿两者相减得出段间空白）：
 * 上沿是量回来的，边框与内边距是纸的像素，缩放过的纸要乘 `zoom` 才接得上。
 */
export function contentTopOf(page: HTMLElement): number {
  const box = page.getBoundingClientRect();
  const cs = getComputedStyle(page);
  const zoom = zoomOf(page);
  return box.top + ((parseFloat(cs.borderTopWidth) || 0) + (parseFloat(cs.paddingTop) || 0)) * zoom;
}

/**
 * 纸上每一块（页面的直接孩子）量出来的单位，按文档顺序，段间空白并进后一块第一枚。页面边框那一族是画在
 * 纸盒里的装饰，不是正文的一块（`doc-edit.ts` 读回模型时同样跳过它）。`layoutUnits`（枚的序号）与
 * `pageCaretAnchor`（枚回到块里）共用这一趟，两处的铺法与条目一一对应。
 */
function unitBlocksOf(page: HTMLElement): { el: HTMLElement; units: PlacedUnit[] }[] {
  const out: { el: HTMLElement; units: PlacedUnit[] }[] = [];
  const zoom = zoomOf(page);
  let cursor = contentTopOf(page);
  for (const child of [...page.children]) {
    const el = child as HTMLElement;
    if (el.classList.contains(PAGE_BORDER_CLASS) || el.classList.contains(SHEET_STRIP_CLASS) || el.classList.contains(BAND_CLONE_CLASS) || el.classList.contains(BAND_MASTER_CLASS)) continue;
    const rect = el.getBoundingClientRect();
    const gap = Math.max(0, rect.top - cursor);
    const units = unitsOfBlock(el, zoom);
    if (units.length && gap > 0) units[0] = { ...units[0]!, height: units[0]!.height + gap };
    else if (!units.length && gap > 0) units.push({ height: gap, forcePage: false, top: rect.top });
    if (units.length && breaksBefore(el)) units[0] = { ...units[0]!, forcePage: true };
    // 「与下段同页」/「段中不分页」两枚旗子（render.ts 按模型落在 `data-keep-*` 上）盖到units上：
    // keepNext 盖在**末**一枚（无内容的 0 高记号不盖——它自己就是一页的开头），keepLines 盖在**首**一枚
    // （带着整段高度）。段里自己就有声明的断点（`w:br type=page` 那一族切出来的内部截）时两条都不盖：
    // 声明过的断点先说了算，段里都断过一回了，「整段不拆」「和下一段同页」就没了着力点。
    const last = units[units.length - 1];
    const internalBreak = units.some((u, ui) => ui > 0 && u.forcePage);
    if (last && !internalBreak) {
      if (el.dataset.keepNext === "1" && last.height > 0) last.keepNext = true;
      if (el.dataset.keepLines === "1") units[0] = { ...units[0]!, keepLines: units.reduce((n, u) => n + u.height, 0) };
    }
    // 孤行控制（`w:widowControl`）：**缺省即开**，只有 render.ts  stamped 的 `data-widow="0"`（文件显式说关）
    // 才算关。开着且这一段有两枚以上行盒时，给每一枚带上它在段里的枚序与总数，装箱那趟据此避开单行跨页。
    // 表格（一枚是表行不是行盒）与只读容器不在此列；段里自己就有声明断点时也不标——那一截一截是文件说好的
    // 切开处，与 keepNext/keepLines 同一退让。
    if (!internalBreak && units.length > 1 && el.tagName !== "TABLE" && !el.classList.contains(CARRY_CLASS) && el.dataset.widow !== "0") {
      for (let ui = 0; ui < units.length; ui++) units[ui] = { ...units[ui]!, widowLine: { index: ui, total: units.length } };
    }
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
 * 装箱那一趟读的列：每枚的**体重**按 100% 的那一份纸量，再乘回当前档；上沿仍是这一档画出来的那一处。
 *
 * 为什么要另量一次：一枚行盒的体重取的是 `max(line-height, 它在流里的推进)`（`lineCharges`），而推进是
 * 浏览器按**当前档的设备像素**把字体的上下伸缩放量之后报回来的。同一档 1.62 行距、11 磅中文（`.scratch/zoom-line-height-dom.ts`，
 * 80 段夹具逐档量）：声明 `line-height: 23.76px` 每一档都不动，可 100% 上块自己的推进是 **24.766px**（比声明高
 * 1.006，那一截是字洒进格子上沿的墨），150% 上量到 35.625 视口 = **23.75 纸像素**，80% 上 19 = 23.75，200% 上
 * 48.5 = 24.25。除回 zoom 回不到纸的那一份——取整发生在当前档。于是 `max` 在 100% 挑推进（24.766）、在 150% 挑
 * 声明（23.76），一页装的枚数从 37 变 39：**换一下缩放，页界就往前挪两枚**，页数也跟着换（38、39、75、76、77、78
 * 段这六档长度在 100% 与 150% 之间换了页数，同一份探针量的）。
 *
 * 页界是纸的事实：屏幕上放大只是把同一份事实画大，重排一次就等于换了一台打印机。WPS 真机换档怎么表现**未取证**
 * （没有录屏授权），但自家两套读数必须先自洽——状态栏的页数、「第 X 页」与打印出来的那份不能各说一张。
 * 所以这一趟把纸先摆回 100% 量高、再摆回来，量到的高乘 `zoom` 进视口空间，与 `planOf` 里 `contentHeightOf × zoom`
 * 那一格同档（`zoomOf` 写的规矩不动，只是多这一处：与视口读数比长短的体重，得先换算到它自己那一档）。
 *
 * 只在 `zoom ≠ 1` 时多这两趟强制布局。摆回来之后按块逐枚对齐：哪一块在另一档里行数不一样（放大把一行挤下去
 * 或拉上来），那一块整块退回画出来的量——那种块本身就是缩放改变了断行，页界按屏幕那份算，不硬凑。
 */
function unitBlocksForPacking(page: HTMLElement): { el: HTMLElement; units: PlacedUnit[] }[] {
  const zoom = zoomOf(page);
  if (zoom === 1) return unitBlocksOf(page);
  // 摆回 100% 量高在前，量画出来的位置在最后：写 `zoom` 是一次 DOM 改动，浏览器按滚动锚定把这一叠纸的视口
  // 位置挪一挪，而枚的 `top` 是**视口的绝对坐标**——下游 `applySheetGaps` 拿它与它自己现读的纸盒上沿相减
  // （「原点只读这一次」那条守则，见该函数）。先量位置再摆档，两份读数就落在两个滚动空间里：实测 150% 上
  // 推距凭空多垫 192 纸px，整张正文压进下一道缝（`.scratch/zoom-gap-dom.ts` 反向对照）。
  page.style.setProperty("zoom", "1");
  const paper = unitBlocksOf(page);
  page.style.setProperty("zoom", String(zoom));
  const drawn = unitBlocksOf(page);
  return drawn.map((b, i) => {
    const base = paper[i];
    if (!base || base.units.length !== b.units.length) return b;
    return {
      el: b.el,
      units: b.units.map((u, k) => {
        const h = base.units[k]!;
        const scaled: PlacedUnit = { ...u, height: h.height * zoom };
        if (u.keepLines !== undefined && h.keepLines !== undefined) scaled.keepLines = h.keepLines * zoom;
        return scaled;
      }),
    };
  });
}

/**
 * 带子在纸上占走的那一截（纸的像素），一节一个数，挂在纸盒上：铺排、页沿推距、读数口都从这一处取，
 * 别处不另定一份（与文件头「量的口径都在这里」同一条守则）。
 *
 * 真机量到的是「带子贴在页边距之内、正文为它让位」：夹具 `/tmp/tall-header.docx`（页眉 10 行、A4、四边
 * 1440 缇、`w:header` 720 缇）正文上沿 ≈3445 缇。此前带子画在纸盒之外、让位增量恒为 0，同一件差出
 * **2005 缇 = 133.67px**（#157/#159）。距离没说 → 带子贴纸沿；这一节没有这一族带子 → 让 0，
 * 没有东西要从正文里占走。
 */
export const BAND_YIELD_TOP = "data-band-yield-t";
export const BAND_YIELD_BOTTOM = "data-band-yield-b";

/** 纸盒上挂着的那一串让位数；没挂 = 这一趟还没量过 = 一处都不让（与「量不出就不给」同一条守则）。 */
function bandYieldsOf(page: HTMLElement, attr: string): number[] {
  const raw = page.getAttribute(attr);
  if (raw === null) return [];
  return raw.split(",").map((v) => {
    const n = parseFloat(v);
    return Number.isFinite(n) && n > 0 ? n : 0;
  });
}

/** 带子从正文里占走的那一截：只算**越过边距线**的那部分。距离没说 → 带子贴纸沿；带子整个坐在边距线以内 →
 *  让 0——那一截空白本来就是纸的边距，不是从正文里借的。 */
export function bandYieldPx(distanceTwips: number | undefined, heightPx: number, marginTwips: number): number {
  return Math.max(0, (distanceTwips === undefined ? 0 : twipsToPx(distanceTwips)) + heightPx - twipsToPx(marginTwips));
}

/** 量一枚带子在纸上占到多高。母本躲在 `display:none` 里量不到（`offsetHeight` 恒 0），按实测那条路走：
 *  克隆一挂进隐藏副本立刻量得到。克隆要补上带子该有的 `.hf` 字面样式（母本自己不是 `.hf`，字号差一档）。
 *
 *  **扣掉锚边那一圈的内边距**：`bandYieldPx` 的算式假设带子的盒顶就在「纸沿 + 距离」上，而摆位不是这么摆的——
 *  距离说的是**文字**到纸沿，所以 `applyBandYield`（原件）与 `paintBandCopies`（纸上的复件）都把元素顶放在
 *  `距离 − 上内边距`，让第一行正好落在距离上。于是上内边距那一截坐在纸沿与首行之间，**已经算在距离里**；
 *  连着盒高一起交给算式就把同一圈内边距数了两遍。真机对照量到这一处：夹具 `/tmp/tall-header.docx` 页眉 10 行、
 *  四边 1440 缇、`w:header` 720 缇、`.hf` 上内边距 18px，盒高 201px 直接交给算式得到正文上沿 3735 缇，
 *  而 WPS 是 ≈3445 缇——多出来的正好是那 18px = 270 缇。扣掉之后装好的壳量到 231px = **3465 缇**，与真机差
 *  20 缇 = 1.33px，而且正好落在画上那枚带子的下沿（原件 `top:140` + `offsetHeight:201` = 纸沿以下 231px）——
 *  让位与摆位同一口径。剩下的那一档是带子的字面样式（`index.html` 的 `.hf{font-size:9.5pt}` ×
 *  `.hf-p{line-height:1.4}` = 一行 17.73px，真机那一侧反推是 181.67px/10 行 = 18.17px）：页面自己的字号行高，
 *  不是文件说的，**WPS 那一档未取证**。
 *
 *  另一头（页眉的下内边距、页脚的上内边距）留在盒里：那是带子最后一行以下、正文必须绕开的一截，
 *  纸上画着的带子确实占着它。
 *
 *  命令条（`.hf-bar`，「插入页码」那一枚按钮）也摘掉：那是**编辑器的把手**，纸上的带子里没有它
 *  （`paintBandCopies` 画复件时同样摘它）。留着就把这一排把手的高度算进带子的高度、让位跟着多半截——
 *  真机对照量到这一处：同一份 10 行页眉，WPS 让出 2005 缇，带着把手的这一版让出 2670 缇。 */
function bandHeightPx(host: HTMLElement, src: HTMLElement, which: "header" | "footer"): number {
  const probe = src.cloneNode(true) as HTMLElement;
  probe.querySelector(".hf-bar")?.remove();
  probe.classList.remove(BAND_SOURCE_CLASS, BAND_MASTER_CLASS);
  probe.classList.add("hf", `hf-${which}`);
  probe.removeAttribute("title");
  for (const p of ["display", "position", "top", "bottom"]) probe.style.removeProperty(p);
  probe.style.cssText += ";position:absolute;left:-9999px;top:-9999px;visibility:hidden;pointer-events:none";
  host.append(probe);
  // 内边距归 CSS 类管（`.hf{padding:18px 96px 6px}`），摘掉母本身上那份内联声明并不能把它去掉——
  // 要的那一圈得从**量出来的盒**里减，所以读克隆自己的 computed，而不是假设它是 0。
  const anchored = parseFloat(getComputedStyle(probe).getPropertyValue(which === "header" ? "padding-top" : "padding-bottom")) || 0;
  const h = probe.offsetHeight - anchored;
  probe.remove();
  return h;
}

/**
 * 铺排前的那一趟：按每一节**真正显示的那一份**带子量出越过边距线的那一截，写回纸盒。第一节走 CSS 变量
 * （正文顶就是纸盒的上内边距，开始线与容量都跟着这一格走），其余各节挂在属性上由 `planOf`/`applySheetGaps`
 * 取。顺手把能点的那份原件挪到「文字贴在节说的距离上」，与纸上的复件同一算法。
 */
export function applyBandYield(host: HTMLElement, page: HTMLElement, doc: WordDocument): void {
  const sections = doc.sections.length ? doc.sections : [DEFAULT_SECTION];
  const owners = new Map<string, HTMLElement>();
  for (const el of host.querySelectorAll<HTMLElement>(`.${BAND_MASTER_CLASS}, .${BAND_SOURCE_CLASS}`)) {
    if (el.dataset.hfPart) owners.set(el.dataset.hfPart, el);
  }
  const tops: number[] = [];
  const bottoms: number[] = [];
  for (const sec of sections) {
    const over = (which: "header" | "footer", distance: number | undefined, margin: number): number => {
      const band = doc.sectionBand(which, sections.indexOf(sec));
      const src = band ? owners.get(band.part) : undefined;
      if (!band || !src) return 0;
      return bandYieldPx(distance, bandHeightPx(host, src, which), margin);
    };
    tops.push(over("header", sec.margins.header, sec.margins.top));
    bottoms.push(over("footer", sec.margins.footer, sec.margins.bottom));
  }
  page.style.setProperty("--band-yield-t", `${tops[0] ?? 0}px`);
  page.style.setProperty("--band-yield-b", `${bottoms[0] ?? 0}px`);
  page.setAttribute(BAND_YIELD_TOP, tops.join(","));
  page.setAttribute(BAND_YIELD_BOTTOM, bottoms.join(","));
  for (const el of host.querySelectorAll<HTMLElement>(`.${BAND_MASTER_CLASS}`)) {
    // 两份原件各守这一叠纸的一端，锚就在它自己那一头：页眉从纸盒**上沿**往下数（第一张的纸沿就是盒子的上沿），
    // 页脚从纸盒**下沿**往上数（画出来的最后一张的下沿就是盒子的下沿——纸盒的空白永远铺到最后一张的底）。
    // 距离说的是**文字**到纸沿（§17.3.2.7/9），样式自己那圈内边距不算在距离里，所以要扣回来，与纸上的复件同一算法。
    const isHeader = el.dataset.hf !== "footer";
    const distance = (isHeader ? sections[0]?.margins.header : sections[sections.length - 1]?.margins.footer) ?? 0;
    const cs = getComputedStyle(el);
    const pad = parseFloat(cs.getPropertyValue(isHeader ? "padding-top" : "padding-bottom")) || 0;
    if (isHeader) el.style.top = `${twipsToPx(distance) - pad}px`;
    else el.style.bottom = `${twipsToPx(distance) - pad}px`;
  }
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
 * 段中那道页界该落在哪一枚上（`placementsAcrossSections` 的孤行控制修正）：边界本来要落在 `i`，
 * 若 `i` 是段里的第 2 枚（页尾只剩段首一行）或段里只剩最后一枚（页头只落段尾一行），就把边界往前挪一枚
 * ——两条修正都只是把内容**下移**到新的一页，那一页只会更空、不会溢出。往前挪不能越过声明的断点
 * （`units[i-1].forcePage`：那一枚自己就是一页的开头，页界由文件说死，孤行控制没着力点）。
 */
function widowBoundary(units: readonly LayoutUnit[], i: number): number {
  const w = units[i]!.widowLine;
  if (!w || w.total < 2 || w.index === 0) return i;
  if (w.index !== 1 && w.total - w.index !== 1) return i;
  return units[i - 1]!.forcePage ? i : i - 1;
}

/**
 * 纯的那一半（多节版）：一列单位在 `unitSections[i]` 说的那一节的纸上各落在第几页（与 `units` 一一对应）。
 * 每一节内的铺法是「单节」那一套再加三条**文件声明的保持规则**（`w:keepNext` / `w:keepLines` / `w:widowControl`，
 * 段落自己写的与它的样式链说的都算——渲染器盖章时早已按 `docx.ts` 的阶梯把两者合成一个值了）：
 *
 * - `keepNext`（段尾那一枚上的旗子）：它和**下一枚**得同页。装这一枚时把下一枚的位子也带上一起算，
 *   带不下就这一枚先挪到下一页（下一块自己整段不拆时，带的就是它那整坨），到了空页还带不下就不硬凑，
 *   这一枚照放、下一枚自己流（声明互相打架的文件里这是两条规则都退让的唯一出路）。下一枚自己另起一页
 *   （`forcePage`，声明断点）时链不成立，各走各的。
 * - `keepLines`（段首那一枚上带着整段高度）：这一段不拆页——装它的第一枚时若一页装不下整段，整段挪到
 *   下一页起头；比一页还高时这条不算数（那样整段放哪都在拆）。
 * - `widowControl`（**缺省即开**，只有文件显式说关才关）：一段拆在两张纸上时两边都至少留两行——段首行
 *   不许独自留在页尾（widow），段尾行不许独自落在页头（orphan）。落点由 `widowBoundary` 定，两条修正都
 *   只把内容往下一页挪。它与 `keepNext` 撞上时（挪下去的第一枚正是上一块 `keepNext` 指着的那枚）这一片让
 *   孤行控制赢、上一块的链退让——两条都是文件声明的，谁先谁后 WPS 真机**未取证**。
 *
 * **节界必另起一页**（断点是声明，哪怕上一节那一页还空着），从下一节的第一枚起换用那一节的容量；
 * 保持规则不跨节（隔着一道声明的页界，链谈不上）。单节件就是 `sheets` 一张的特例——两处的答案必须逐枚相同。
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
    // 这一枚自己的一坨（keepLines）：整段放不进当前这页就整段挪走（空页上放得下才算数，不然这条作废）。
    const atomic = u.keepLines !== undefined && u.keepLines <= contentHeightPx + FIT_SLACK ? Math.min(u.keepLines, contentHeightPx) : 0;
    // 与下一枚的链（keepNext）：下一枚的代价 = 它自己（若它也是 keepLines 的段首，就是它那一整段）。
    const next = u.keepNext === true && i + 1 < units.length && (unitSections[i + 1] ?? 0) === s && units[i + 1]!.forcePage !== true ? units[i + 1]! : undefined;
    const nextH = next ? (next.keepLines !== undefined && next.keepLines <= contentHeightPx + FIT_SLACK ? Math.min(next.keepLines, contentHeightPx) : Math.min(Math.max(next.height, 0), contentHeightPx)) : 0;
    const need = Math.max(h + nextH, atomic);
    if (used > 0 && used + need > contentHeightPx + FIT_SLACK) {
      // 孤行控制：页界的落点还要再过一道 `widowBoundary`——段首行不许独自留在页尾、段尾行不许独自落在页头。
      // 往前挪出去的那几枚改记到新的一页上，它们的体重一并带下去（那一页因此可能比纸还挤一点：一段的行盒
      // 高过一页时按装箱这趟照常续排，不为它多算一页）。
      const b = widowBoundary(units, i);
      page++;
      used = 0;
      for (let j = b; j < i; j++) {
        pages[j] = page;
        used += Math.min(Math.max(units[j]!.height, 0), contentHeightPx);
      }
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
  // 装箱比的是一枚枚**画出来的**高（视口像素），而纸容量是节说的一截长度（纸的像素）：纸放大 zoom 倍，行盒
  // 跟着放大，纸的容量也得放大同样倍数，两边才是同一个空间。这一记乘在这里——`placementsAcrossSections` 自己
  // 不知道缩放，交出去的 `pages` 只由「量出来的高 ÷ 说出来的容量」定。
  // `PagePlan.contentHeightPx` 交回给外面的仍是纸的像素（它是报给人看的一截物理长度，不是比长短的中间量）。
  const zoom = zoomOf(page);
  // 第一节的两端让位已经写进纸上画出来的内边距（`applyBandYield` 的 `--band-yield-t/-b`），上面那一格
  // `contentHeightOf` 量回来就带着它；其余各节的纸没画在流里，按属性上同一趟写下的数自己扣——两边问的是
  // 同一截高度，谁也不多扣一次。让完放不下一行正文的节不给方案（宁可少一行读数）。
  const yieldT = bandYieldsOf(page, BAND_YIELD_TOP);
  const yieldB = bandYieldsOf(page, BAND_YIELD_BOTTOM);
  const sheets: Sheet[] = [{ contentHeightPx: contentHeightPx * zoom }];
  for (let s = 1; s < sections.length; s++) {
    const sheet = sheetOf(sections[s]!);
    if (!sheet) return undefined;
    const body = sheet.contentHeightPx - (yieldT[s] ?? 0) - (yieldB[s] ?? 0);
    if (!(body > 0)) return undefined;
    sheets.push({ contentHeightPx: body * zoom });
  }
  const ends = sectionEndsOf(doc.blocks);
  // 高不能直接从当前这一档缩放里取：Blink 按设备像素对齐字形上下伸，同一行在 100% 流推 24.766、在 150% 推
  // 23.75（纸的像素），于是 `max(声明, 流推)` 在两档各挑一头，接缝从 37 枚挪到 39 枚，用户滚一下缩放条就
  // 换页。改拿 100% 那一档量高再乘回这一档的缩放（`unitBlocksForPacking`），接缝只由文档内容定。
  const blocks = unitBlocksForPacking(page);
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
 * 第 `idx` 枚单位落在哪一块、哪一块里的第几枚（`rowColumnAt`/`pageCaretAnchor`/`lineCaretAnchor` 与
 * 页缝垫片共用这一趟「枚→块」的来路，铺法的条目一一对应）。
 */
function locateUnit(blocks: readonly { el: HTMLElement; units: PlacedUnit[] }[], idx: number): { el: HTMLElement; seg: number; index: number } | undefined {
  let seg = idx;
  for (let i = 0; i < blocks.length; i++) {
    const b = blocks[i]!;
    if (seg < b.units.length) return { el: b.el, seg, index: i };
    seg -= b.units.length;
  }
  return undefined;
}

/**
 * `caretMid`（视口坐标，光标自己那枚矩形的**中线** = 上沿 + 高/2）落在第几枚。带子按**画出来的上沿**切：
 * 这一枚管「自己的上沿 → 下一枚的上沿」（最后一枚收到 `top + height`）。不能拿 `top + height` 当界——
 * 一块的第一枚把段间空白和块内容上沿以上那截也算进了高（见 `PlacedUnit`），那样它的高会压到第二枚头上：
 * 折行边界那一处的光标（中线在本枚上沿之下半个光标高）会被判回上一行。上一版的说法是「相邻两枚首尾相接，
 * 中线离接缝总有半个行高」——接缝确实在上沿，可第一枚的地盘被高撑到过界，`w:spacing` 的段后距越大越容易踩。
 * 中线而不是上沿：光标矩形的字体盒上沿与碎片上沿差零点几像素且方向不定，正好压在界线上。
 * 中线落在纸面之上/之下时取下（或最后一）枚——防呆，不是「按最近的算」。`pageAtCaret` 与 `rowColumnAt` 共用这一趟。
 */
function unitIndexAt(units: readonly PlacedUnit[], caretMid: number): number {
  let idx = -1;
  for (let i = 0; i < units.length; i++) {
    const u = units[i]!;
    if (u.height <= 0) continue;
    const next = i + 1 < units.length ? units[i + 1]!.top : u.top + u.height;
    if (caretMid >= u.top && caretMid < Math.max(next, u.top)) {
      idx = i;
      break;
    }
  }
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
  return placementInPlan(plan, caretMid);
}

function placementInPlan(plan: PagePlan, caretMid: number): PagePlacement {
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

/**
 * 光标（`node`/`offset`）那一处**字符自己**画出来的上沿：位置后面还有字时量那个字的一字 Range
 * （非折叠），后面没有字了（文本末尾、空段落）才退回折叠光标那条路。折行边界那一处折叠光标的矩形
 * 有亲和性、量到的是**上一行**（`topAtChar` 的注），所以光标那几格的取点（`caretMidOf` 的「行/列」与
 * 「第 X 页」）与行首二分（`topAtChar`）得走同一个地基——这里是它的节点版。
 */
export function caretCharTop(node: Node, offset: number): number | undefined {
  const text = node.nodeType === Node.TEXT_NODE ? (node as Text) : null;
  if (text && offset < text.data.length) {
    const range = document.createRange();
    range.setStart(text, offset);
    range.setEnd(text, offset + 1);
    const own = range.getClientRects()[0];
    if (own) return own.top;
  }
  return topAtPoint(node, offset);
}

/**
 * 拉平串里第 `at` 位**那个字自己**画出来的上沿（非折叠的一字 Range）。行首二分不能量折叠光标：折行边界
 * 那一处折叠光标的矩形有亲和性，量到的是**上一行**（`.scratch/sheets-dom.html` 的长段落上量到整整差一行），
 * 行首那一位会被判成还没到这一行、落下整整一个字。字自己的矩形没有这条暧昧。位置两节点相接处
 * `positionAt` 取后一个节点的开头（`caretCharTop` 就从那一枚字量起）。量不出范围（零宽字符）时退回
 * 折叠光标那条路。
 */
function topAtChar(flat: FlatText, at: number): number | undefined {
  const a = positionAt(flat, at);
  return a ? caretCharTop(a.node, a.offset) : undefined;
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
 * 拉平串里第一条落进 `unit` 这枚行盒的字符位置（0..len；量不出上沿时不猜，`undefined`）：把每个字符的
 * 上沿排出来（那个字自己的矩形，口径写在 `topAtChar`），同一行上各处上沿相同，折行处整行高地下落一次，
 * 所以上沿对位置单调不减，二分找「上沿 ≥ 行盒顶 − 半行高」的第一个位置。容差取半行高：行内因字号不一
 * 差几像素的不管，跨行至少差一整行高。
 * 「列」（`columnOf`）与「第 N 行在哪里」（`lineCaretAnchor`）共用这一趟——行首与列底数必须是同一个位置。
 */
function firstAtUnit(flat: FlatText, unit: PlacedUnit): number | undefined {
  let lo = 0;
  let hi = flat.len;
  while (lo < hi) {
    const mid = (lo + hi) >> 1;
    const top = topAtChar(flat, mid);
    if (top === undefined) return undefined;
    if (top >= unit.top - unit.height / 2) hi = mid;
    else lo = mid + 1;
  }
  return lo;
}

/**
 * 光标字符位置 `at` 在这一行里排第几（1 起）：行首位置由 `firstAtUnit` 找（上沿二分的口径写在那里），
 * 列 = 位置 − 行首 + 1。量不出上沿时不猜，`undefined`。
 */
function columnOf(flat: FlatText, unit: PlacedUnit, at: number): number | undefined {
  if (!flat.len) return 1;
  const lo = firstAtUnit(flat, unit);
  if (lo === undefined) return undefined;
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
  return lineColumnInPlan(plan, caret);
}

function lineColumnInPlan(plan: PagePlan, caret: { mid: number; node: Node; offset: number }): CaretLine | undefined {
  const { blocks, units } = plan;
  if (!units.length) return { row: 1 };
  const idx = unitIndexAt(units, caret.mid);
  const row = rowInPage(plan.places, units, idx);
  if (row === undefined) return undefined;
  const where = locateUnit(blocks, idx);
  if (!where || where.el.tagName === "TABLE" || where.el.classList.contains(CARRY_CLASS)) return { row };
  const flat = flattenText(where.el);
  const at = flatOfCaret(where.el, caret.node, caret.offset);
  if (at === undefined || at > flat.len) return { row };
  const column = columnOf(flat, units[idx]!, at);
  return column === undefined ? { row } : { row, column };
}
/**
 * 光标那行离**纸沿**的纵向距离，px——状态栏「纵向位置」那一格（WPS 的 `KxWpsPagePositionInfoCommand`，
 * kuip 452 行，detail「@Vertical Page Position. Click to Open "Go To" dialog.」）。纸沿是这一页纸的上沿：
 * 第 1 页在纸上画得出来（`contentTopOf`），第 2 页起在流里没有 DOM——但模型里每一页的内容都从纸沿下
 * 「上边距 + 让位」处起排（第 1 节与 `render.ts` 的 `--pad-t` 同源，其余各节按节自己的 `w:pgMar@w:top`，
 * 而 #157 之后开始线还要加上这一节为页眉让出的那一截），所以反推得出来：
 * 纸沿 = **本页第一枚内容的块沿 − 上边距 − 让位**。块沿与单位自己报的行/字符盒顶之间那一截走
 * `blockInsetAbove`——与 `applySheetGaps` 的推距同一支，两条路必须压在同一个口径上（少这一截读数就少 3.5px）。
 * 光标行顶走 行/列 那一趟取点（折叠 Range 自己那枚矩形）；页的归属与「第 X 页」同一条中线（`unitIndexAt`）。
 * 纸上没有内容、或光标行量不出矩形时不猜，`undefined`。取的是**字符盒**的上沿（不是行盒），与 行/列 的
 * 列同一口径；WPS 真机从纸沿量到哪条线、什么单位、几位小数**都未取证**。
 */
export function pagePositionAt(page: HTMLElement, doc: WordDocument, caret: { mid: number; node: Node; offset: number }): number | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  return positionInPlan(plan, page, doc, caret);
}

function positionInPlan(plan: PagePlan, page: HTMLElement, doc: WordDocument, caret: { mid: number; node: Node; offset: number }): number | undefined {
  if (!plan.units.length) return undefined;
  const zoom = zoomOf(page);
  const origin = page.getBoundingClientRect().top;
  const local = (viewportTop: number): number => (viewportTop - origin) / zoom;
  const idx = unitIndexAt(plan.units, caret.mid);
  let first = plan.places.indexOf(plan.places[idx]!);
  // 零高度的单位不算内容（与推距那一趟同一档防呆）：它压在开始线上但不占地方，拿它反推纸沿会少算整截。
  while (first + 1 < plan.units.length && plan.units[first]!.height <= 0) first++;
  const top = topAtPoint(caret.node, caret.offset);
  if (top === undefined) return undefined;
  const sections = doc.sections.length ? doc.sections : [DEFAULT_SECTION];
  const sec = sections[plan.unitSections[first]!];
  if (!sec) return undefined;
  // 交出去的是**纸的像素**里那一截长度（状态栏按厘米显示）：分子是量回来的视口读数相减，缩放过的纸要除回
  // zoom 才与节说的数同在一个空间里相加。相加的那一段有两截：一截是「上边距 + 这一节为页眉让出的那一截」——
  // 正文开始线现在确实比边距低（#157），只加边距会把带子占走的那一截漏掉；另一截是 `blockInsetAbove`，因为
  // 单位自己报的是行/字符盒的顶，而开始线上压着的是**块沿**（#167 的推距按的就是块沿）。少了这一截，纸沿
  // 会被反推得比画出来的低一枚半行距余量，读数跟着少 3.5px（= 52.5 缇）。
  return (
    (top - plan.units[first]!.top) / zoom +
    blockInsetAbove(plan.blocks, plan.units, first, local) +
    twipsToPx(sec.margins.top) +
    (bandYieldsOf(page, BAND_YIELD_TOP)[plan.unitSections[first]!] ?? 0)
  );
}

/** 状态栏一次刷新的全部几何读数（「第 X 页 / 共 Y 页 / 纵向位置 / 行 / 列」）——都出自同一趟铺排。 */
export type CaretStatus = {
  pages: number | undefined;
  placement?: PagePlacement;
  line?: CaretLine;
  positionPx?: number;
};

/**
 * 状态栏那一排格子**一次量完**：`第 X 页`、`共 Y 页`、`纵向位置`、`行`、`列` 五格各自都能单量（`pageAtCaret`
 * ／`measurePageLayout`／`pagePositionAt`／`rowColumnAt`），但逐格量意味着逐格重跑一趟铺排——一次刷新四趟。
 * 铺排自己会读每一枚行盒（#170 之后在 `zoom ≠ 1` 还要多摆一次档量 100% 那一份高），所以逐格量在换档时贵。
 * `.scratch/packcost-dom.ts`（600 段 × A4，同一件同一 DOM 只换档，每档 15 次取均）量到：
 *
 * | 档 | 一趟铺排 | 逐格量四趟 | 合并一趟 |
 * |---|---|---|---|
 * | 100% | 1.88ms | 5.42ms | 1.70ms |
 * | 150% | 8.08ms | **24.32ms** | 8.43ms |
 * | 200% | 8.39ms | 24.01ms | 8.07ms |
 *
 * 打字每敲一下都刷一次，所以这一趟只铺排一次、五格从同一份 `PagePlan` 里取；合并后每一档都落在「一趟铺排」
 * 那一列的量级上。**150% 那一列多出来的不是合并这条路造成的**：把 #170 的第二趟重量关掉（同探针、代码摆回旧算法）
 * 再量一次，150% 档三列各是 1.70 / 4.96 / 1.92ms——多出的六倍全在 #170 那一趟上，合并只是把「四趟各重一次」
 * 收成「重一次」。
 *
 * 取出来的数与逐格量**必须是同一份**（各格共用一支 `unitIndexAt` 与同一个 plan），这也正是这几格的要求：
 * 「第 X 页」与「行 Y」「纵向位置」说的得是同一处光标、同一次铺排。纸没画上或页高量不出来时整排不给数
 * （`undefined`），与 `measurePageLayout` 同一条守则。
 *
 * 这一句由 `.scratch/pos-blockedge-dom.ts` 第 5 节钉住（100%/150%/缩回 100% 三档逐字段比）：它的反向对照是
 * **把这一趟缓存起来**（模块级记忆 `planOf` 的结果），150% 档「行/列」与「纵向位置」各红一条——旧 plan 的 `units`
 * 带着 100% 空间的 `mid`，与现量在 150% 视口里的 `caret.mid` 不在同一个空间，于是选错枚。所以**这一趟不能缓存**：
 * #170 之后铺排的结果是按档位的，纸缝与 `mid` 都只在那一次测量里成框。
 */
export function caretStatusOf(page: HTMLElement, doc: WordDocument, caret: { mid: number; node: Node; offset: number }): CaretStatus | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  return {
    pages: plan.pages,
    placement: placementInPlan(plan, caret.mid),
    line: lineColumnInPlan(plan, caret),
    positionPx: positionInPlan(plan, page, doc, caret),
  };
}

/**
 * 第 `content` 枚单位自己的顶，到它所在那一块**块沿**之间的那一截（同一个原点下的纸的像素）：块沿以下一枚
 * 半行距余量才是第一行的行/字符盒顶（12pt 配 1.62 行距实测 3.5px）。`applySheetGaps` 把块沿按到开始线上，
 * `pagePositionAt` 从开始线反推纸沿，两处问的必须是同一截——分开写就会出现「同一张纸的两条测量路对不上」。
 * 同一块里的续行与首行同一档字面，所以拿块的第一枚量到的就是这一档的余量；表行那一支本来就是行自己的盒子，
 * 量出来 ≈0，不用另开。块找不到、或量出负数（负边距、上标把行盒提到块沿以上）都按 0——与「量不出就不给」
 * 同一条守则。
 */
function blockInsetAbove(
  blocks: readonly { el: HTMLElement; units: PlacedUnit[] }[],
  units: readonly PlacedUnit[],
  content: number,
  local: (viewportTop: number) => number,
): number {
  const where = locateUnit(blocks, content);
  const firstOfBlock = where ? units[content - where.seg] : undefined;
  if (!where || !firstOfBlock) return 0;
  return Math.max(0, local(firstOfBlock.top) - local(where.el.getBoundingClientRect().top));
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
 * 页高切）。枚落在哪一截就落哪一截：表格取那一行；段落取那一枚行盒的行首（`firstAtUnit` 的上沿二分）——
 * 软换行切出来的页与另起一页的断点之后那一截，走的都是同一条取点。整页只有 0 高
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
  const found = locateUnit(blocks, idx);
  if (!found) return undefined;
  const block = found.el;
  const segment = found.seg;
  if (atFlowEnd) return { ...tailPositionOf(block), block };
  if (block.tagName === "TABLE") {
    // 表的一枚单位是一行；0 高的行与页缝垫下的那一枚 `unitsOfBlock` 都撇掉了，这里按同一口径取第 segment 行。
    const row = measuredRows(block as HTMLTableElement)[segment]?.el;
    return row ? { ...headPositionOf(row), block } : undefined;
  }
  // 段落（非表格、非只读容器）：这一页头一枚单位的落笔位置就是那一枚行盒的**行首**——与「行」目标走同一条
  // 取点（`firstAtUnit` 的上沿二分）。软换行切出来的页、另起一页的断点后面那一截的首行，都落在这里。
  // 旧写法把块内**枚序**当断点序（`br[segment−1]`）：段落里第 2 枚往后起页时既找不到断点、又不认识软换行，
  // 一路返回 `undefined`（第 125 轮的 straddle 夹具量到：整段 6 行、页 2 从第 2 行起，跳页其实没落成）。
  // 一枚文字都没有、行首量不出时退回块首（那是这一块的第 1 行，与落不到时的容错同一条）。
  if (block.classList.contains(CARRY_CLASS)) return { ...headPositionOf(block), block };
  const flat = flattenText(block);
  const at = flat.len ? firstAtUnit(flat, units[idx]!) : undefined;
  const pos = at === undefined ? undefined : positionAt(flat, at);
  return pos ? { ...pos, block } : { ...headPositionOf(block), block };
}

/**
 * 光标在**全篇第几行**（绝对行号，从文首连数、换页不重数）——「定位」对话框「行」目标的种子。取点与状态栏
 * 「行」格同一条（`unitIndexAt` 的中线），但那一格数的是**页内**序数（`rowInPage`），这一支数的是全篇序数；
 * 行 = 铺排出来的一枚行盒（表格按行计），0 高的断点记号不占行。WPS 真机「行」按哪种数**未取证**。
 * 光标落到防呆路径上时算第 1 行（`rowInPage` 同一条守则）；纸上什么都没有时也报 1（那是种子不是范围，
 * 范围由 `measurePageLayout().lines` 给）；纸没画上、量不出时不猜，`undefined`。
 */
export function lineNumberAt(page: HTMLElement, doc: WordDocument, caret: { mid: number; node: Node; offset: number }): number | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  const { units } = plan;
  if (!units.length) return 1;
  const idx = unitIndexAt(units, caret.mid);
  let n = 0;
  for (let j = 0; j <= idx; j++) {
    if (units[j]!.height > 0) n++;
  }
  return Math.max(1, n);
}

/**
 * 「第 `target` 行在哪里」（`pageCaretAnchor`／`sectionCaretAnchor` 的姊妹，「行」目标的反向映射）：行号是
 * **绝对**的——从文首连数、换页不重数，与 `lineNumberAt` 同一口径；数的是同一趟铺排出来的正高行盒
 * （0 高的断点记号不占行）。落点落在这一行**头一个字符**上：段落里按上沿二分找行首（`firstAtUnit`，与
 * 「列」的底数同一趟取点），表格取那一行开头，一枚文字都没有、上沿量不出时退回块首（那是这一块的第 1 行，
 * 与落不到时的容错同一条）。越界的行号（<1、>总行数、非整数）不猜，返回 `undefined`，由调用方拒绝。
 * WPS 真机落点落在行首、上下文怎么滚动**未取证**；这一版钉的是「行号 → 那一行」这条回路本身。
 */
export function lineCaretAnchor(page: HTMLElement, doc: WordDocument, target: number): PageAnchor | undefined {
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  const { blocks, units } = plan;
  if (!Number.isInteger(target) || target < 1) return undefined;
  let idx = -1;
  let seen = 0;
  for (let i = 0; i < units.length; i++) {
    if (units[i]!.height > 0 && ++seen === target) {
      idx = i;
      break;
    }
  }
  if (idx < 0) return undefined;
  const found = locateUnit(blocks, idx);
  if (!found) return undefined;
  const block = found.el;
  const segment = found.seg;
  if (block.tagName === "TABLE") {
    // 表的一枚单位是一行；0 高的行与页缝垫下的那一枚 `unitsOfBlock` 已经撇掉，这里按同一口径取第 segment 行（`pageCaretAnchor` 同一条）。
    const row = measuredRows(block as HTMLTableElement)[segment]?.el;
    return row ? { ...headPositionOf(row), block } : undefined;
  }
  if (block.classList.contains(CARRY_CLASS)) return { ...headPositionOf(block), block };
  const flat = flattenText(block);
  const at = flat.len ? firstAtUnit(flat, units[idx]!) : undefined;
  const pos = at === undefined ? undefined : positionAt(flat, at);
  return pos ? { ...pos, block } : { ...headPositionOf(block), block };
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
    const row = tableRowsOf(block as HTMLTableElement)[0];
    return row ? { ...headPositionOf(row), block } : undefined;
  }
  return { ...headPositionOf(block), block };
}

/** 页缝里垫的那枚空白 span（段落之间用；表行之间用 `makeSheetGapRow`）。 */
function makeSheetGap(height: number): HTMLElement {
  const gap = document.createElement("span");
  gap.className = SHEET_GAP_CLASS;
  gap.contentEditable = "false";
  gap.setAttribute("aria-hidden", "true");
  gap.style.cssText = `display:block;height:${height}px;pointer-events:none`;
  return gap;
}

/**
 * 一张表的格子有几列（页缝垫片那一枚 `td` 要横跨的枚数）。
 *
 * 按每一行自己的 `colSpan` 加起来取最大：`render.ts` 把竖合并的续行也写成 `<td>`（`display:none`），
 * 所以横竖合并混在一起时仍是同一个数。这张表里页缝垫下的那一枚不算行——它自己就是按这个数写的，
 * 算进去只会让下一趟比上一趟多一列。数不出列（表是空的）就给 1：一行一格，高照样让出来。
 */
function gridColumnsOf(table: HTMLTableElement): number {
  let n = 1;
  for (const tr of tableRowsOf(table)) {
    let inRow = 0;
    for (const td of tr.children) inRow += (td as HTMLTableCellElement).colSpan || 1;
    if (inRow > n) n = inRow;
  }
  return n;
}

/**
 * 缝落在两枚表行之间时垫下的那一行（#158）：空格子、`border:0`、不接鼠标、不可编辑。
 * `border-style:none` 而不是 `hidden`——合并边框里 `none` 输给邻格真框（上下两行各自的线照画），
 * `hidden` 会把它们一起抹掉。高就是这一道的推距（纸的像素，与段落垫片同一个口径）。
 */
function makeSheetGapRow(height: number, columns: number): HTMLElement {
  const tr = document.createElement("tr");
  tr.className = SHEET_GAP_ROW_CLASS;
  tr.setAttribute("aria-hidden", "true");
  const td = document.createElement("td");
  td.colSpan = Math.max(1, columns);
  td.contentEditable = "false";
  td.style.cssText = `height:${height}px;padding:0;border:0 none;background:transparent;pointer-events:none;user-select:none`;
  tr.append(td);
  return tr;
}

/**
 * 清掉上一趟垫下的页缝，并把被 `splitText` 切开过的文字节点合回去——下一趟还在这一段里找行首。
 * 两种垫片一起清：段落里的空 span，与表行之间的那一行（`#158`）。
 */
function clearSheetGaps(page: HTMLElement): void {
  const parents = new Set<Element>();
  for (const gap of [...page.querySelectorAll<HTMLElement>(`.${SHEET_GAP_CLASS}, .${SHEET_GAP_ROW_CLASS}`)]) {
    if (gap.parentElement) parents.add(gap.parentElement);
    gap.remove();
  }
  for (const parent of parents) parent.normalize();
}

/**
 * 在一枚文字的开头处垫进一段空白；落笔在文字中间就切开它，垫片落在两半之间。
 * 渲染出来的 run 可能包着行内装饰（带样式的 span 一族），垫片要落进**块**这一层：自里向外逐层劈开，
 * 行内祖先照原样复制成后半（文字可以随便切，垫片进不了行内元素的肚子里）。`node` 就是块自己时
 * （行首是编号符号/tab 那种装饰的抬升路径）直接插在块里。
 */
function insertSheetGapAt(node: Node, offset: number, height: number, block: HTMLElement): boolean {
  const gap = makeSheetGap(height);
  if (node === block) {
    block.insertBefore(gap, block.childNodes[offset] ?? null);
    return true;
  }
  if (!block.contains(node)) return false;
  let tail: ChildNode | null = null;
  let cur: Node = node;
  let off = offset;
  for (;;) {
    const parent = cur.parentNode;
    if (!parent) return false;
    if (cur.nodeType === Node.TEXT_NODE) {
      const t = cur as Text;
      tail = off <= 0 ? t : off >= t.data.length ? t.nextSibling : t.splitText(off);
    } else {
      const el = cur as Element;
      const first = el.childNodes[off] ?? null;
      if (first) {
        const copy = el.cloneNode(false) as Element;
        let n: ChildNode | null = first;
        while (n) {
          const nextNode: ChildNode | null = n.nextSibling;
          copy.append(n);
          n = nextNode;
        }
        parent.insertBefore(copy, el.nextSibling);
        tail = copy;
      } else {
        tail = el.nextSibling;
      }
    }
    if (parent === block) break;
    cur = parent;
    off = tail ? [...parent.childNodes].indexOf(tail) : parent.childNodes.length;
  }
  block.insertBefore(gap, tail);
  return true;
}

/** 一枚行盒里第一个字的落笔位置（行首二分与「列」同一趟：`flattenText` + `firstAtUnit` + `positionAt`）。 */
function firstPositionOfUnit(el: HTMLElement, unit: PlacedUnit): { node: Node; offset: number } | undefined {
  const flat = flattenText(el);
  if (!flat.len) return undefined;
  const at = firstAtUnit(flat, unit);
  if (at === undefined) return undefined;
  const pos = positionAt(flat, at);
  if (!pos) return undefined;
  // 行首第一枚可能是画出来的装饰（编号符号、`w:tab` 的定宽 span）。落笔点要抬到装饰**前面**——
  // 垫片塞进 `display:inline-block` 的装饰里面会把那一行拆成两行（装饰留在上一张纸上）。
  // 装饰不是正文的直接孩子（更深）时这一道不垫：宁可这一处不动，也不把段落画坏。
  const deco = pos.node.parentElement?.closest(".bullet, .tab");
  if (deco) {
    const i = [...el.childNodes].indexOf(deco);
    return i >= 0 ? { node: el, offset: i } : undefined;
  }
  return pos;
}

/**
 * 把连续的一张纸画成 N 张（第一片：页缝与纸层）。
 *
 * 纸盒仍是一张连续的白纸，分页是量出来的数（`planOf`）；这里做的是**画**出「一张纸到这儿为止」：
 * 每一道页界，在下页第一枚有内容的单位的正上方垫一段行内块级空白（`SHEET_GAP_CLASS`），把它推到
 * 「上一张纸的下沿 + 页缝 + 下一节的上边距」的位置；返回的空档位置给调用方画条带盖住（`SheetStrip`，
 * 纸盒本地的纸的像素，可直接写进 CSS）。页从哪切与 `pageAtCaret` 同一份 `planOf`。
 *
 * 垫片自己不在量的口径里：`unitsOfParagraph` 按它把量程切段，它那一截不产单位、高度也不进任何一枚
 * 单位，所以垫过之后再量，页界分毫不动。每一趟先 `clearSheetGaps` 再量——幂等，且被切开过的文字
 * 节点合回原文（模型读回 `inlinesOf` 对空元素本来就不产条目，`modelOffset` 也数 0 个字符，垫片对
 * 模型与光标都是透明的）。
 *
 * 垫片插在哪儿，看下页第一枚有内容的单位落在哪：
 * - 块内第 seg>0 枚（一段续到下一页）：把那行头一个字切开，垫片落在两半之间。段里的行盒量程天然
 *   按它分段，行宽不变、折行不变，所以它上下的行盒与没垫时逐枚相同。
 * - 块首、上一块是段落：垫在**上一段末尾**（末尾追加）。垫进下一段开头会丢那一行的 `text-indent`
 *   （匿名块的首行不承接段首缩进，`.scratch/gap-probe.html` 量到 0 对 32px）。
 * - 块首、上一块进不去（表格、只读容器）而本块是段落：垫在**本块开头**——本块首行的缩进会丢，
 *   这一处**接受瑕疵**并记录；WPS 真机怎么画**未取证**。
 * - 落在两枚表行之间（含整张表的第一行）：块级垫片进不了 `<table>` 的肚子，垫的是**一行空格子**
 *   （`SHEET_GAP_ROW_CLASS`，无框、不可编辑）。它不是表的一枚行：量的每一枚（`tableRowsOf`）、读回模型
 *   （`doc-edit.ts` 的 `rowsOf`）、冒烟里的行数读数三处都撇掉它，所以重拼出来的件里 `<w:tr>` 枚数一枚不多。
 * - 本块与上一块都进不去（只读容器紧跟一张表）：垫在那张表的**最后一行之下**，两张纸照样分得开。
 * - 落到只读容器内部、又找不到上述任何一处：这一道界**不动 DOM**，也不画条带（垫片压不进那一块，
 *   画了条带会把字盖住）。
 *
 * 0 高的断点记号不是内容：页界落在记号上时推的是记号之后第一枚 `height>0` 的单位。若那枚已经在更
 * 后面的页上（连着两枚断点＝中间空出一整张纸），这一道不垫——空纸的空白由更下面那一道垫片的推距
 * 一并让出来，条带照画（条带是绝对定位的装饰，与垫片落在哪一段无关）。流尾只剩空纸时没有内容可推，
 * 这一片不画它的条带。
 *
 * 高度差从量出来的几何来：第 p 道缝的推距 H = 目标 − 现位 − **上面各道已让出的推距之和**（上面每垫一段，
 * 下面所有内容都跟着下沉那么多——推距是一组三角形方程，从上往下一趟解完），目标 = 这一页的纸沿 + 这一节的
 * 纸高 + 页缝 + 下一节的上边距。纸沿按链推：第 1 页的上沿就是画出来的纸盒上沿（本地坐标 0，`render.ts` 的
 * `--pad-t` 与节的上边距同源），其后每张 = 上一张纸沿 + 上一节纸高 + 页缝——链不经过单位自己的上沿，整页只有
 * 0 高记号的空纸也算得出。现位是量回来的视口读数换算到同一条链上（见下）。落笔点与推距都按**原 DOM** 解出
 * （这时一个垫片都还没插），插的时候从下往上——同一段文字里先切靠后的位置，靠前的落笔点才不会被切开后的节点
 * 边界带偏。量不出纸（`planOf` 不给）、H 不是正数、行首位置找不着时都不猜：这一道跳过，条带也省掉。页缝多宽、
 * 条带什么颜色，WPS 真机**未取证**（`SHEET_GAP_PX` 是自定的）。
 */
/**
 * 段落元素上带的那一串制表位（`render.ts` 从 `p.tabStops` 写的）：`缇:档` 逗号分隔，档 `l`/`c`/`r`。
 * 数写在 `data-` 上而不是 `style` 上：编辑读回模型时看的是格式，不该把这些位置当成段落新说的一句话。
 */
export const TAB_STOPS_ATTR = "data-tab-stops";

/** `data-tab-stops` 的字面读回来：按缇排好序，认不出的那一节直接丢（半枚画歪的位置不如没有）。 */
export function parseTabStops(raw: string): { posTwips: number; val: "left" | "center" | "right" }[] {
  const out: { posTwips: number; val: "left" | "center" | "right" }[] = [];
  for (const part of raw.split(",")) {
    const at = part.lastIndexOf(":");
    if (at <= 0) continue;
    const pos = Number(part.slice(0, at));
    if (!Number.isFinite(pos)) continue;
    out.push({ posTwips: pos, val: part.slice(at + 1) === "r" ? "right" : part.slice(at + 1) === "c" ? "center" : "left" });
  }
  return out.sort((a, b) => a.posTwips - b.posTwips);
}

/** 制表符至少要走出来的一点宽度：停在原地不算走过去（Word 在同一位置上也照样前进到下一格）。 */
const TAB_MIN_PX = 0.5;

/**
 * 把 `w:pPr/w:tabs` 说的位置画到纸上（#168）。页面本来只有 `.page .tab { width: 32px }` 一枚发明值，
 * 也就是说正文部件写了 318 枚制表位（230 张 `w:pPr/w:tabs` 清单）和 116 枚制表符、三份部件合计 520 枚位，
 * 而纸把每一枚都停在 32px 上——这一段与那一格对齐的说法整个丢了。这一趟按 §17.4.19/§17.3.2.9 的摆位算：
 * 位置从**正文区左沿**（`--pad-l` = 左边距加 `w:gutter`）起算，越过最后一枚明写的位之后按
 * `w:defaultTabStop` 一步步走（授权 46 份：43 份 420、3 份 720，全写了）；
 * `left` 让后头的文字起在位上，`right` 让那一截的**末沿**贴在位上，`center` 把它骑在位上。
 *
 * 318 与 116 这两个数是同一个口径下对出来的账（`.scratch/tabs-roundtrip-46.ts` 与拼法切片）：正文部件里
 * `<w:tab` 起头的标签 434 枚 = 带属性的位定义 318 + 自闭的制表符 116；这 434 枚隔一次真实编辑（43 份可编辑，
 * 3 份 `Normal.dotm` 没有可改的段）进出一致，逐段解析出的位置与类型也一字不差。
 *
 * 116 枚那笔账对得上（`.scratch/tabs-inline-context.py` 按最近的结构祖先分档，六份文件）：正文部件里
 * `w:r/w:tab` 116 枚，其中 **30 枚**住在正文段落（模型 `doc.blocks` 直接那一层，分 14 段），**86 枚**住在
 * `w:sdt` 容器里的段落。后 86 枚画在只带的 carried 盒子（`.doc-carry`，#32 那一档：内容看得见、光标进不去），
 * 不是 `.para`，这一趟不动它们。前 30 枚的分档也量了（`.scratch/tabs-default-audience.ts`）：14 段里 4 段的
 * 样式链上有位，**10 段 / 26 枚**一枚都没写——只按 `data-tab-stops` 选宿主就会把这 26 枚留在发明的 32px 上。
 *
 * 每一枚都先撤掉上一趟的宽度再量，从左往右一枚一枚落笔：读一次 `getBoundingClientRect()` 就同步排一次版，
 * 所以下一枚量到的起点已经带上前一枚垫出来的宽度——这与 Word 的行内推进是同一个方向。那一截的宽按**第一行**
 * 算（跨行的段落里，制表位管的是它这一行的那一截）。都不够用时给一步默认格，而不是硬挤到 0。
 *
 * 两档写在明面上没接：① `w:leader` 只画位置不画点线（94 枚全说 `dot`：正文 92、样式 2）——点线是 WPS 画出来的
 * 东西，手上没有它的实据，按「未取证的不猜」留着；② 格子里那一段（`w:tc` 里的 `w:p`）挡掉不做，它的基准该是
 * 格沿而不是正文区左沿——授权 46 份里这样的 `w:tabs` **0 块**（`.scratch/tabs-inside-table.py`），所以这一挡
 * 眼下没有受众，只是不让将来落进格子里的制表位按纸的边距画。③ 带子里的 12 枚位与 2 枚制表符仍按 `pre-wrap`
 * 的浏览器默认走，不带这一趟的账（带子的字面字号那一档本就未取证）。
 *
 * 落点由 `.scratch/tab-place-dom.ts` 在真 DOM 里量（17 例：三档各贴自己的位、走过第一格找下一格、清单乱序按缇
 * 排、没写位的走 `w:defaultTabStop` 整步、借位与 `clear`、无 settings 时的兜底那一步、zoom 与重跑不改答案）；
 * 反向对照（摘掉下面那一调用）13 条红。
 */
export function applyTabStops(page: HTMLElement, doc: WordDocument): void {
  // 凡是画了制表符的正文段都要走这一趟，**没有** `data-tab-stops` 的那一类也算：授权 46 份里带制表符的正文段
  // 14 段共 30 枚，其中 10 段 / 26 枚的样式链上一枚位都没写（`.scratch/tabs-default-audience.ts`），
  // 它们该按 `w:defaultTabStop` 的步长走，而不是留在 CSS 那枚发明的 32px 上。
  const hosts = [...page.querySelectorAll<HTMLElement>(".para")].filter(
    (el) => !el.closest("table") && el.querySelector("span.tab"),
  );
  if (!hosts.length) return;
  const zoom = zoomOf(page);
  const originLeft = page.getBoundingClientRect().left;
  const cs = getComputedStyle(page);
  const padL = parseFloat(cs.paddingLeft) || 0;
  const padR = parseFloat(cs.paddingRight) || 0;
  const dtsPx = twipsToPx(doc.defaultTabStop() ?? 420);
  // 正文区的右沿：纸宽（视口读数除回纸的像素）减去两头的边距。默认格子只排到这一条线为止。
  const areaRight = page.getBoundingClientRect().width / zoom - padR;
  const localLeft = (viewportLeft: number): number => (viewportLeft - originLeft) / zoom;
  for (const el of hosts) {
    const stops = parseTabStops(el.getAttribute(TAB_STOPS_ATTR) ?? "");
    const tabs = [...el.querySelectorAll<HTMLElement>("span.tab")];
    // 候选位：明写的每一枚，加上从正文区左沿按默认步长排开、且排在最后一枚明写的位之后的那些格子。
    const cands = stops.map((s) => ({ x: padL + twipsToPx(s.posTwips), val: s.val }));
    const lastX = cands.length ? cands.reduce((a, b) => Math.max(a, b.x), 0) : padL;
    for (let x = padL + dtsPx; x <= areaRight; x += dtsPx) if (x > lastX + TAB_MIN_PX) cands.push({ x, val: "left" });
    cands.sort((a, b) => a.x - b.x || 0);
    if (!cands.length) continue;
    for (const t of tabs) t.style.width = "";
    for (let i = 0; i < tabs.length; i++) {
      const t = tabs[i]!;
      const x0 = localLeft(t.getBoundingClientRect().left);
      const segPx = firstLineWidthPx(t, tabs[i + 1] ?? null, el, zoom);
      let width = 0;
      for (const s of cands) {
        if (s.x <= x0 + TAB_MIN_PX) continue;
        const need = s.val === "left" ? s.x - x0 : s.val === "right" ? s.x - segPx - x0 : s.x - segPx / 2 - x0;
        if (need < TAB_MIN_PX) continue;
        width = need;
        break;
      }
      t.style.width = `${(width > 0 ? width : dtsPx).toFixed(2)}px`;
    }
  }
}

/**
 * 这一枚制表符后面那一截文字的宽，只算它**第一行**上的那些矩形（下一枚制表符或这一段末尾为界）。
 * `Range` 的矩形是按碎片给的，行内换行时同一个 Range 会横跨几行，取第一行那一档就是取这一行真正要挤进去的那截。
 * 取的是那一档的**跨度**（最右沿减最左沿），不是各块矩形之和：装机件上实测（`.scratch/tabs-dom-dump.js`，
 * 段落里只有一个 `<span>` 的三字符段）Blink 对同一个 Range 的同一行给了**两枚一模一样的矩形**，按之和量就
 * 把段宽读成两倍，右对齐差一整截、居差半截（2026-10-18 量到 `widthPx` 少 48px）。跨度对重复矩形天然免疫，
 * 对一行里真有几块碎片（换脸、加粗）也正是这一行要挤进去的那一段。
 */
function firstLineWidthPx(from: HTMLElement, to: HTMLElement | null, host: HTMLElement, zoom: number): number {
  const r = document.createRange();
  if (to) {
    r.setStartAfter(from);
    r.setEndBefore(to);
  } else {
    r.selectNodeContents(host);
    r.setStartAfter(from);
  }
  const rects = [...r.getClientRects()].filter((b) => b.width > 0);
  if (!rects.length) return 0;
  const top = rects[0]!.top;
  const line = rects.filter((b) => Math.abs(b.top - top) <= 1);
  return (Math.max(...line.map((b) => b.right)) - Math.min(...line.map((b) => b.left))) / zoom;
}

export type SheetLayout = { strips: SheetStrip[]; pages: number };

export function applySheetGaps(page: HTMLElement, doc: WordDocument): SheetLayout | undefined {
  clearSheetGaps(page);
  const plan = planOf(page, doc);
  if (!plan) return undefined;
  const { blocks, units, unitSections, places, pages } = plan;
  if (pages < 2) return { strips: [], pages };
  const sections = doc.sections.length ? doc.sections : [DEFAULT_SECTION];
  // 下一张纸的正文开始线 = 节说的上边距 + 这一节为页眉让出的那一截。第一张由纸盒画出来的上内边距表达，
  // 往后的每一张由这里的推距表达——同一趟 `applyBandYield` 量出来的同一个数（#157：真机上带子在纸内，
  // 正文一张一张都从它下沿起排，不是只有第一张）。
  const yieldTopOf = bandYieldsOf(page, BAND_YIELD_TOP);
  const padTopOf = (s: number): number => twipsToPx((sections[s] ?? sections[0]!).margins.top) + (yieldTopOf[s] ?? 0);
  const paperHOf = (s: number): number => twipsToPx((sections[s] ?? sections[0]!).pageHeightTwips);
  const firstUnitOfPage = (target: number): number => places.indexOf(target);
  // 两种像素：单位的位置是 `getBoundingClientRect()` 量出来的**视口像素**，纸高/页缝/边距是节自己说的
  // **纸的像素**（`twipsToPx`）。纸盒带 `zoom` 时两者差 zoom 倍，而这一趟算出来的落笔点与条带上沿最后都要
  // 写成 CSS 长度——CSS 长度在缩放过的纸里就是纸的像素。所以纸沿一律从纸盒上沿起算（本地坐标），量回来的
  // 视口读数按同一个原点、同一个 zoom 除回纸的像素：原点只读这一次，读两次会因中间改动 DOM 后的滚动锚定对不上。
  const zoom = parseFloat(page.style.getPropertyValue("zoom")) || 1;
  const origin = page.getBoundingClientRect().top;
  const localTopOf = (viewportTop: number): number => (viewportTop - origin) / zoom;
  // 归属口径（#171 逐枚对拍后才定下来）：某一张纸上有没有正文、那枚算第几张，一律按**行盒的上沿**数——
  // `units[].top` 与这里画的条带是同一条链（`.scratch/seam-straddle-dom.ts`：长段落切 7 张、67 枚行的表，
  // 铺排的 `places` 与按条带反推的张 **0 分岔**，页数 == 条带数 + 1）。**不要按块的 `top` 认领**：一枚跨缝的块
  // 只有一个块顶、留在上一张上，它的续段却压在下一张的开始线之下，按块顶数会把有墨的那一张读成空张
  // （§109 七 那条「第 4 张有条带却没正文块」正是这么来的，不是铺排算错）。
  const sheetTop: number[] = [0, 0]; // 第 1 张纸的上沿就是画出来的纸盒上沿：本地坐标 0
  for (let target = 2; target <= pages; target++) {
    sheetTop.push(sheetTop[target - 1]! + paperHOf(unitSections[firstUnitOfPage(target - 1)] ?? 0) + SHEET_GAP_PX);
  }
  const opaque = (el: HTMLElement | undefined): boolean => !!el && (el.tagName === "TABLE" || el.classList.contains(CARRY_CLASS));
  const strips: SheetStrip[] = [];
  // 推距是三角形方程组：第 p 道缝的垫片会把**下面所有内容**一起推下沉，所以往下解时每一道都要扣掉上面
  // 各道已经让出的距离（`lifted`）。先从上往下一趟把「哪一道垫多少」全解出来、落笔点也全按**原 DOM**
  // 找好（这时一个垫片都还没插），再统一下手插——顺序见下。
  const ops: { h: number; offset: number; node?: Node; prev?: HTMLElement; block?: HTMLElement; row?: HTMLTableRowElement; place?: "before" | "after" }[] = [];
  let lifted = 0;
  for (let p = 1; p <= pages - 1; p++) {
    const idxNext = firstUnitOfPage(p + 1);
    if (idxNext < 0) continue;
    const sheetBottom = sheetTop[p]! + paperHOf(unitSections[firstUnitOfPage(p)] ?? 0);
    const seam = { top: sheetBottom, height: SHEET_GAP_PX, page: p, section: unitSections[firstUnitOfPage(p)] ?? 0, nextSection: unitSections[idxNext] ?? 0 };
    let content = idxNext;
    while (content < units.length && units[content]!.height <= 0) content++;
    if (content >= units.length) continue;
    if (places[content] !== p + 1) {
      strips.push(seam);
      continue;
    }
    const where = locateUnit(blocks, content);
    if (!where) continue;
    const target = sheetBottom + SHEET_GAP_PX + padTopOf(unitSections[content] ?? 0);
    // 目标说的是**块的上沿**该落在哪一条开始线上（第一张纸就是画出来的纸盒上内边距），而 `units[].top` 报的是
    // **行盒**的顶，中间那一截由 `blockInsetAbove` 量（真机实测：第一张的块沿正好压在开始线上差 0，往后的张上
    // 是行盒压在开始线、块沿差 **−3.5px = −52.5 缇**；夹具 `/tmp/tall-header-2sheet.docx`，
    // `.scratch/sheet-2yield-probe.js`）。不补这一截，第二张往后的正文就比第一张高它。
    const above = blockInsetAbove(blocks, units, content, localTopOf);
    const h = target - localTopOf(units[content]!.top) + above - lifted;
    if (!(h > 0)) continue;
    if (where.seg > 0) {
      if (where.el.tagName === "TABLE") {
        // 缝落在第 seg 枚表行之前：块级垫片进不了 `<table>` 的肚子里，垫的那一行落在它前面。
        const row = measuredRows(where.el as HTMLTableElement)[where.seg]?.el;
        if (!row) continue;
        ops.push({ h, row, place: "before", offset: -1 });
      } else if (opaque(where.el)) {
        continue;
      } else {
        const pos = firstPositionOfUnit(where.el, units[content]!);
        if (!pos) continue;
        ops.push({ h, node: pos.node, offset: pos.offset, block: where.el });
      }
    } else {
      const prev = blocks[where.index - 1]?.el;
      // 先后次序按「哪一处推得最小」：上一块进得去就垫在上一块尾（原有的那一条），本块进得去就垫在本块首
      // （段首一枚垫片，不动任何表）。两张表之间的缝、只读带紧跟一张表这两条才有别的可选。
      if (prev && !opaque(prev)) ops.push({ h, prev, offset: -1 });
      else if (!opaque(where.el)) ops.push({ h, node: where.el, offset: 0 });
      else if (where.el.tagName === "TABLE") {
        // 缝正落在这张表的第一行之前，而上一块进不去（另一张表、或只读容器）：垫在这一张的第一行之上。
        const row = measuredRows(where.el as HTMLTableElement)[0]?.el;
        if (!row) continue;
        ops.push({ h, row, place: "before", offset: -1 });
      } else if (prev?.tagName === "TABLE") {
        // 本块进不去（只读容器），上一块是表：垫在那张表的最后一行之下，两道纸才真分得开。
        const rows = measuredRows(prev as HTMLTableElement);
        const row = rows[rows.length - 1]?.el;
        if (!row) continue;
        ops.push({ h, row, place: "after", offset: -1 });
      } else continue;
    }
    lifted += h;
    strips.push(seam);
  }
  // 从下往上插：同一段文字里先切靠后的落笔点，靠前的那个（按原 DOM 解出来的节点与偏移）才不会被切开
  // 的节点边界带偏。跨段落的插入互不影响，顺序只对这一处有约束。
  for (let i = ops.length - 1; i >= 0; i--) {
    const op = ops[i]!;
    if (op.prev) op.prev.append(makeSheetGap(op.h));
    else if (op.row) {
      const table = op.row.closest("table");
      if (!table) continue;
      const gapRow = makeSheetGapRow(op.h, gridColumnsOf(table));
      // DOM 方法不能从宿主上摘下来调（摘了就 `Illegal invocation`，整趟渲染探针全废）：按落笔方向各调一次。
      if (op.place === "after") op.row.after(gapRow);
      else op.row.before(gapRow);
    } else if (op.node && op.block) insertSheetGapAt(op.node, op.offset, op.h, op.block);
  }
  return { strips, pages };
}
