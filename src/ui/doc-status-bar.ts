/**
 * 文字页脚的**状态栏**：WPS 的 Writer 在窗口底下常驻一条 `KxWpsStatusCommand id="statusbar"`（`office6/res/wpscommon.kuip`，
 * 442–511 行，底下 63 个直接子项）。其中 `KxWpsStatusWordCountInfoCommand` 带着 `aid="792"`、`id="WordCountInfoCommand"`、
 * `minimumWidth="75"`，detail 说的是「Word Count. Click to open "Word Count" dialog.」——ribbon 那条「字数统计」
 * 就写在**同一个文件的 267 行**、同一个 `aid="792"`，而这一条里其余 62 项一个 `aid` 都没有。所以这一格不是
 * 另一个功能，是同一条命令的另一种长相：报一个数，点一下开同一个对话框。
 *
 * 这一条上挂着**三格**，顺序照 kuip 的子项次序（页码/页数 在 字数 之前）：
 *
 * - `KxWpsStatusPageNumberInfoCommand id="PageNumberInfoCommand"`，detail「Page Number. Click to open "Go To" dialog.」
 *   ——**第 X 页**，X 是光标（选区活动端）落在的那一页。量法与「字数统计」的 页数 同一套（`doc-layout.ts`
 *   的 `pageAtCaret`：光标自己那枚矩形落进哪一行盒，就是哪一页）。光标不在纸面上（点了批注面板的输入框、
 *   点了页眉页脚带）时**保留上一次的读数**——WPS 真机上那种时候显示什么**未取证**，这一条是自己定的口径。
 * - `KxWpsStatusPageAreaInfoCommand id="PageAreaInfoCommand"`，detail「Page Area Number. Click to open "Go To" dialog.」
 *   ——**共 N 页**，N 与对话框的 页数 同一份测量；光标跑出纸面也照样量得出来。
 * - `KxWpsStatusWordCountInfoCommand`（上段那条）——**字数**，点开对话框。
 *
 * 这两格**不画成按钮**：kuip 的 detail 说它们点一下开「定位」对话框，而「定位」我们还没有。WPS 那两格的
 * 行为（点击、中文界面字面）都**未取证**（`.qm` 那堵墙与 `doc-word-count.ts` 头注记的是同一件），所以这里
 * 只报数，不挂一个点不开的命令。
 *
 * 还没接的那几格：`KxWpsStatusSectionInfoCommand`（节）跟着多节排版那片走；`KxWpsStatusRowInfoCommand`／
 * `KxWpsStatusColumnInfoCommand`（行/列，kuip 里靠 `rightSibling`/`leftSibling` 配成一对）要的是「这一页里第几行 /
 * 这一行里第几个字」，行能从同一套单位数出来、列要字符级映射，是另一片。`KxWpsPagePositionInfoCommand`
 * （纵向位置）同理。
 *
 * 还有一条**没接**的腿：`res/wpsentries.kuip` 里 `idPath="Status bar Popup Menu.WordCountStatusBar"` 说明 WPS
 * 的状态栏可以右键弹出一串共 15 项的开关，逐格显示或隐藏。本轮不摆那个弹出菜单：逐格开关要先有「哪些格在
 * 显示」这件事，我们三格是写死的。
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
.solo-doc-status .solo-doc-status-loc {
  padding: 0 2px; color: #6b7480;
  /* 页码/总页 也是数字，同样用等宽数位。 */
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
  private readonly pageItem: HTMLSpanElement;
  private readonly pagesItem: HTMLSpanElement;
  private readonly wordItem: HTMLButtonElement;
  /** 光标跑出纸面那几格里，页码保留上一次的读数（口径见文件头）；页数不跟光标走，每回现量。 */
  private lastPage = 1;

  /**
   * 数是从这一个编辑器**现读**的：这几格报的数与对话框里那几行必须出自同一次计算，所以都走
   * `editor.caretPage()`（页码／页数，`layout()` 的另一半）与 `editor.counts()`，不在构造时把数抄下来。
   */
  constructor(private readonly editor: DocEditor) {
    injectStyles();
    this.root = document.createElement("div");
    this.root.className = "solo-doc-status";
    this.pageItem = document.createElement("span");
    this.pageItem.className = "solo-doc-status-loc";
    this.pagesItem = document.createElement("span");
    this.pagesItem.className = "solo-doc-status-loc";
    this.wordItem = document.createElement("button");
    this.wordItem.type = "button";
    this.wordItem.className = "solo-doc-status-words";
    this.wordItem.title = "字数。点击打开「字数统计」对话框";
    this.wordItem.addEventListener("click", () => openWordCount(this.editor.counts(), this.editor.selectionCounts(), this.editor.layout()));
    this.root.append(this.pageItem, this.pagesItem, this.wordItem);
    document.body.append(this.root);
    this.refresh();
  }

  refresh(): void {
    const at = this.editor.caretPage();
    if (at) this.lastPage = at.page;
    // 光标不在纸面上时 `caretPage()` 不给数，但总页数不跟着光标走：现量一遍，量不出来才留空。
    const total = at?.pages ?? this.editor.layout()?.pages;
    this.pageItem.textContent = total === undefined ? "" : `第 ${this.lastPage} 页`;
    this.pagesItem.textContent = total === undefined ? "" : `共 ${total} 页`;
    const c = this.editor.counts();
    this.wordItem.textContent = `字数：${c.cjk + c.words}`;
  }

  detach(): void {
    this.root.remove();
  }
}
