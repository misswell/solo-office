/**
 * 文字页脚的**状态栏**：WPS 的 Writer 在窗口底下常驻一条 `KxWpsStatusCommand id="statusbar"`（`office6/res/wpscommon.kuip`，
 * 442–511 行，底下 63 个直接子项）。其中 `KxWpsStatusWordCountInfoCommand` 带着 `aid="792"`、`id="WordCountInfoCommand"`、
 * `minimumWidth="75"`，detail 说的是「Word Count. Click to open "Word Count" dialog.」——ribbon 那条「字数统计」
 * 就写在**同一个文件的 267 行**、同一个 `aid="792"`，而这一条里其余 62 项一个 `aid` 都没有。所以这一格不是
 * 另一个功能，是同一条命令的另一种长相：报一个数，点一下开同一个对话框。
 *
 * 这一条上挂着**六格**，顺序照 kuip 的子项次序（页码→页数→节→行→列→字数；纵向位置那一格夹在节与行
 * 之间，还没接，见下）：
 *
 * - `KxWpsStatusPageNumberInfoCommand id="PageNumberInfoCommand"`，detail「Page Number. Click to open "Go To" dialog.」
 *   ——**第 X 页**，X 是光标（选区活动端）落在的那一页。量法与「字数统计」的 页数 同一套（`doc-layout.ts`
 *   的 `pageAtCaret`：光标自己那枚矩形落进哪一行盒，就是哪一页）。光标不在纸面上（点了批注面板的输入框、
 *   点了页眉页脚带）时**保留上一次的读数**——WPS 真机上那种时候显示什么**未取证**，这一条是自己定的口径。
 * - `KxWpsStatusPageAreaInfoCommand id="PageAreaInfoCommand"`，detail「Page Area Number. Click to open "Go To" dialog.」
 *   ——**共 N 页**，N 与对话框的 页数 同一份测量；光标跑出纸面也照样量得出来。
 * - `KxWpsStatusSectionInfoCommand`（451 行，detail「@Page Section. Click to open "Go To" dialog.」）——**第 X 节**，
 *   X 是光标落在第几节。这一格**不量几何**：节的边界是文件自己声明的（带 `w:pPr/w:sectPr` 的段落结束它那一节），
 *   从模型数（`sectionEndsOf`），块进第几节、节首一块在哪（`sectionCaretAnchor`）的口径都写在 `doc-layout.ts`
 *   文件头。光标不在纸面上时与页码同一口径保留上一次读数。点一下开「定位」对话框、预选「节」目标；
 *   真机字面与点开时预选什么**未取证**。
 * - `KxWpsStatusRowInfoCommand`（detail「@Line Status Bar」）与 `KxWpsStatusColumnInfoCommand`（detail
 *   「@Column Status Bar」，kuip 里靠它 `rightSibling` 与上一格 `leftSibling` 配成一对）——**行 X／列 X**：
 *   行是光标在**这一页里**排第几（每页重数，0 高的断点记号不数），列是它在这一行的第几个字。
 *   量法在 `doc-layout.ts` 的 `rowColumnAt`（取点与页码同一条中线）。光标不在纸面上时与页码同一口径
 *   保留上一次的读数；表格与只读容器里数不出列，列那一格保留上一次的读数（行照报）。这两格是 **span
 *   不是 button**：kuip 里这两条没有点击动作，没有可点开的东西。真机字面、每页是否重数行、行尾空白
 *   怎么计**都未取证**。
 * - `KxWpsStatusWordCountInfoCommand`（上段那条）——**字数**，点开对话框。
 *
 * 页码/页数/节/字数 四格是**按钮**：kuip 的 detail 说 页码/页数/节 点一下开「定位」对话框（`doc-goto.ts`，
 * 与 ⌘G 同一条命令 `Id="GoTo"`；点节那一格预选「节」目标），字数那一格是我们先接上的同形状。WPS 那几格
 * 点击的细节与中文界面字面都**未取证**（`.qm` 那堵墙与 `doc-word-count.ts` 头注记的是同一件），这里按
 * detail 接上这条腿：种子取光标所在页/节，「定位」把落点交回 `DocEditor.goToPage`／`goToSection`。
 *
 * 还没接的那一格：`KxWpsPagePositionInfoCommand`（纵向位置）等量纵向坐标的那一片。
 *
 * 还有一条**没接**的腿：`res/wpsentries.kuip` 里 `idPath="Status bar Popup Menu.WordCountStatusBar"` 说明 WPS
 * 的状态栏可以右键弹出一串共 15 项的开关，逐格显示或隐藏。本轮不摆那个弹出菜单：逐格开关要先有「哪些格在
 * 显示」这件事，我们六格是写死的。
 *
 * 样式从这个文件注入；这一条挂在 `document.body` 上而不是编辑器的宿主里——`render()` 会把宿主整个换掉。
 * `body` 本身是 `flex-direction: column`，所以一个 `flex: none` 的直接子元素就落在窗口底下，与表格里那排
 * 页签同一种摆法；换开表格或演示时由外壳把它摘掉，不留到别人页脚上。
 */
import type { DocEditor } from "./doc-edit.js";
import { openWordCount } from "./doc-word-count.js";

const STYLE_ID = "solo-doc-status-styles";

const CSS = `
.solo-doc-status {
  flex: none; display: flex; align-items: center; gap: 10px;
  height: 26px; padding: 0 8px;
  background: #e6e9ed; border-top: 1px solid #cfd4da;
  font: 11.5px/1 -apple-system, "Helvetica Neue", "PingFang SC", sans-serif; color: #5b636d;
}
.solo-doc-status button {
  font: inherit; color: inherit; background: none; border: 0; padding: 2px 6px; border-radius: 4px; cursor: pointer;
  /* 数字用等宽数位，打字时那一格不跳宽。 */
  font-variant-numeric: tabular-nums;
}
.solo-doc-status button:hover { background: #d7dce2; color: #1f2328; }
.solo-doc-status .solo-doc-status-loc, .solo-doc-status .solo-doc-status-sec, .solo-doc-status .solo-doc-status-rc {
  padding: 0 2px; color: #6b7480;
  /* 页码/总页/节/行/列 都是数字，同样用等宽数位。 */
  font-variant-numeric: tabular-nums;
}
@media print { .solo-doc-status { display: none; } }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

export class DocStatusBar {
  private readonly root: HTMLDivElement;
  private readonly pageItem: HTMLButtonElement;
  private readonly pagesItem: HTMLButtonElement;
  private readonly sectionItem: HTMLButtonElement;
  private readonly lineItem: HTMLSpanElement;
  private readonly colItem: HTMLSpanElement;
  private readonly wordItem: HTMLButtonElement;
  /** 光标跑出纸面那几格里，页码保留上一次的读数（口径见文件头）；页数不跟光标走，每回现量。 */
  private lastPage = 1;
  /** 节与页码同一口径：光标不在纸面上时保留上一次的读数。 */
  private lastSection = 1;
  /** 行/列 与页码同一口径：光标不在纸面上、或数不出列（表格/只读容器）时保留上一次的读数。 */
  private lastRow: number | undefined;
  private lastCol: number | undefined;

  /**
   * 数是从这一个编辑器**现读**的：这几格报的数与对话框里那几行必须出自同一次计算，所以都走
   * `editor.caretPage()`／`editor.caretLine()`（行/列，与「第 X 页」同一趟取点）与 `editor.counts()`，
   * 不在构造时把数抄下来。
   */
  constructor(private readonly editor: DocEditor) {
    injectStyles();
    this.root = document.createElement("div");
    this.root.className = "solo-doc-status";
    this.pageItem = document.createElement("button");
    this.pageItem.type = "button";
    this.pageItem.className = "solo-doc-status-loc";
    this.pageItem.title = "页码。点击打开「定位」对话框";
    this.pagesItem = document.createElement("button");
    this.pagesItem.type = "button";
    this.pagesItem.className = "solo-doc-status-loc";
    this.pagesItem.title = "页数。点击打开「定位」对话框";
    this.sectionItem = document.createElement("button");
    this.sectionItem.type = "button";
    this.sectionItem.className = "solo-doc-status-sec";
    this.sectionItem.title = "节号：光标在第几节。点击打开「定位」对话框";
    // 页码/页数/节 三格与 ⌘G 开的是同一个对话框，种子与范围都由编辑器现量（口径见 `doc-edit.ts` 的 `showGoTo`）；
    // 点哪一格开哪个目标：前两格「页」，节这一格「节」。
    this.pageItem.addEventListener("click", () => this.editor.showGoTo());
    this.pagesItem.addEventListener("click", () => this.editor.showGoTo());
    this.sectionItem.addEventListener("click", () => this.editor.showGoTo("section"));
    // kuip 里 行/列 是两条没有点击动作的 status 条目（detail 里没有 "Click to open"），画成 span 而不是按钮。
    this.lineItem = document.createElement("span");
    this.lineItem.className = "solo-doc-status-rc";
    this.lineItem.title = "行号：光标在这一页里排第几行";
    this.colItem = document.createElement("span");
    this.colItem.className = "solo-doc-status-rc";
    this.colItem.title = "列号：光标在这一行的第几个字";
    this.wordItem = document.createElement("button");
    this.wordItem.type = "button";
    this.wordItem.className = "solo-doc-status-words";
    this.wordItem.title = "字数。点击打开「字数统计」对话框";
    this.wordItem.addEventListener("click", () => openWordCount(this.editor.counts(), this.editor.selectionCounts(), this.editor.layout()));
    this.root.append(this.pageItem, this.pagesItem, this.sectionItem, this.lineItem, this.colItem, this.wordItem);
    document.body.append(this.root);
    this.refresh();
  }

  refresh(): void {
    const at = this.editor.caretPage();
    if (at) this.lastPage = at.page;
    // 光标不在纸面上时 `caretPage()`／`caretLine()` 不给数，但总页数不跟着光标走：现量一遍，量不出来才留空。
    const total = at?.pages ?? this.editor.layout()?.pages;
    this.pageItem.textContent = total === undefined ? "" : `第 ${this.lastPage} 页`;
    this.pagesItem.textContent = total === undefined ? "" : `共 ${total} 页`;
    const cs = this.editor.caretSection();
    if (cs) this.lastSection = cs.section;
    this.sectionItem.textContent = total === undefined ? "" : `第 ${this.lastSection} 节`;
    const rc = this.editor.caretLine();
    if (rc) {
      this.lastRow = rc.row;
      if (rc.column !== undefined) this.lastCol = rc.column;
    }
    this.lineItem.textContent = total === undefined || this.lastRow === undefined ? "" : `行 ${this.lastRow}`;
    this.colItem.textContent = total === undefined || this.lastCol === undefined ? "" : `列 ${this.lastCol}`;
    const c = this.editor.counts();
    this.wordItem.textContent = `字数：${c.cjk + c.words}`;
  }

  detach(): void {
    this.root.remove();
  }
}
