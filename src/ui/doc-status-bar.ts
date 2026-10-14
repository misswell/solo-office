/**
 * 文字页脚的**状态栏**：WPS 的 Writer 在窗口底下常驻一条 `KxWpsStatusCommand id="statusbar"`（`office6/res/wpscommon.kuip`，
 * 442–511 行，底下 63 个直接子项）。其中 `KxWpsStatusWordCountInfoCommand` 带着 `aid="792"`、`id="WordCountInfoCommand"`、
 * `minimumWidth="75"`，detail 说的是「Word Count. Click to open "Word Count" dialog.」——ribbon 那条「字数统计」
 * 就写在**同一个文件的 267 行**、同一个 `aid="792"`，而这一条里其余 62 项一个 `aid` 都没有。所以这一格不是
 * 另一个功能，是同一条命令的另一种长相：报一个数，点一下开同一个对话框。
 *
 * 这一条上只挂了**字数**那一格。状态栏里其余的读数——页码、页次、节、行、列——要的都不是模型里的数：
 * `kwpsapiex/wps.js` 那一份 API 把它们放在 `ComputeStatistics(wdStatisticPages / wdStatisticLines …)` 那一族里，
 * 也就是**排版之后**才知道的结果，而这一页目前把整篇画在一张连续的纸上（`render.ts` 只出一个 `.page`）。
 * 报一个「第 1 页 共 1 页」对一个四十段的件是假读数，所以宁可空着——与对话框里不报「行数」同一条理由。
 *
 * 措辞仍是**未取证**：`.qm` 那堵墙与 `doc-word-count.ts` 头注记的是同一件（`zh_CN` 下按 UTF-8 与 UTF-16 两种
 * 编码都搜不到这些界面串），「字数：」按 Word 一族通用叫法写；`detail` 那句英文提示翻成的中文 tooltip 同理。
 * 行为的实据（同一个 aid、点击开对话框）是有的，字面没有。
 *
 * 还有一条**没接**的腿：`res/wpsentries.kuip` 里 `idPath="Status bar Popup Menu.WordCountStatusBar"` 说明 WPS
 * 的状态栏可以右键弹出一串共 15 项的开关，逐格显示或隐藏。本轮只挂字数这一格，没有可关的东西，所以那个
 * 弹出菜单不做；等状态栏真长出第二格时再说。
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
  private readonly wordItem: HTMLButtonElement;

  /**
   * 数是从这一个编辑器**现读**的：这一格报的数与对话框里那几行必须出自同一次计算，所以两边都走
   * `editor.counts()`，不在构造时把数抄下来。
   */
  constructor(private readonly editor: DocEditor) {
    injectStyles();
    this.root = document.createElement("div");
    this.root.className = "solo-doc-status";
    this.wordItem = document.createElement("button");
    this.wordItem.type = "button";
    this.wordItem.className = "solo-doc-status-words";
    this.wordItem.title = "字数。点击打开「字数统计」对话框";
    this.wordItem.addEventListener("click", () => openWordCount(this.editor.counts()));
    this.root.append(this.wordItem);
    document.body.append(this.root);
    this.refresh();
  }

  refresh(): void {
    const c = this.editor.counts();
    this.wordItem.textContent = `字数：${c.cjk + c.words}`;
  }

  detach(): void {
    this.root.remove();
  }
}
