/**
 * 「字数统计」：WPS 的这条命令是 aid 792、`id="WordCount"`、`idMso="WordCount"`、`dialogHint="1"`，快捷键
 * `Ctrl+Shift+G`（`office6/res/wpscommon.kuip` 第 267 行）。同一个 aid 在包里露出四张脸：ribbon 审阅那一页的
 * `CT_Review.RB_Proofing.WordCountToolBar`、工具菜单的 `ToolsMenu.WordCount`、状态栏上那条
 * `KxWpsStatusWordCountInfoCommand`（`aid="792"`、`minimumWidth="75"`，detail 说的是「点一下开字数统计对话框」），
 * 以及状态栏右键菜单里的 `WordCountStatusBar`（那一串 15 项就是各格的开关）——后三处都在
 * `res/wpscommon.kuip` / `res/wpsentries.kuip`。也就是说**这一页不是新加一个功能，是把同一条命令多接了一条腿**：
 * 状态栏那一格的实现在 `doc-status-bar.ts`，两边读的数都出自这里的 `WordCountReadout`。
 * 快捷键两条按法都接了：包里的 `shortcut cmd="WordCount"` 同时给着 `Ctrl+Shift+G` 与 `Meta+Shift+G`
 * （`res/wpscommon.kuip:1170–1171`，1359–1360 又原样重复一组），实现在 `renderer.ts` 那条窗口级 keydown 分发里，
 * 只有文字页认，且**不带 Shift 的 ⌘G 不算**这条命令。
 *
 * 数不是这里数的，是 `WordDocument.wordCounts()` 数出来的（口径与取舍都写在那一处）；这个文件只把它们排成
 * 一行行的读数。`字数` 那行报的是 `中文字符和朝鲜语单词 + 非中文单词`，与 Word 一族那三行的连法相同。
 *
 * 「行数」与「页数」这两行**不报**：它们要的是排版之后的结果。包里的 JS API 自己就这么分类——
 * `office6/addons/kwpsapiex/wps.js` 的 `getDocumentWordCount()` 把 `words`、`characters`、
 * `charactersWithSpaces`、`paragraphs`、`lines`、`pages` 全都交给 `doc.ComputeStatistics(wdStatistic…)`，
 * 而 `lines`/`pages` 一行在哪里断由字体度量和栏宽决定，模型里没有这个数，编一个出来就是假读数。
 * 这一页目前把整篇画在一张连续的纸上（`render.ts` 只出一个 `.page`），照画报「页数 1」对一个四十段的件
 * 是错的数，所以宁可少一行。「所选文字」那一栏同理：有选区时 WPS 会另报选中部分的字数，本轮只数全文。
 *
 * 对话框标题那句「字数统计」在包里能找到（`office6/addons/kwpsapiex/wps.js`），但**各行标签的中文原话取不到**：
 * `.qm` 是编译过的翻译，本轮把两种编码都扫了一遍——`en_US/wps.qm` 里能按 UTF-16 搜到「Word Count」，
 * `zh_CN/wps.qm` 连这条源文都搜不到，「字数统计／段落数／字符数」在 `zh_CN` 下所有 `.qm` 里 UTF-8 与 UTF-16
 * 全为空（那 74,788 个 CJK 码元说明中文确实在里面，只是存储形式不同：压缩过的 trie）。与「增加缩进量」那次
 * 同一堵墙，所以下面这几行标签是按 Word 一族的通用叫法写的，WPS 真机上逐行的措辞仍是未取证。
 *
 * 样式从这个文件注入：模块自己的样式跟着模块走。外壳那份样式在 `src/ui/index.html`，它顶上那条深色栏
 * 写的是 `body > header`——只写 `header` 会顺着元素名爬到这张卡片头上，把 46px 高、82px 左内边距和深色
 * 渐变画到标题行里（拍干净截图时看到的，同一件事还压着页面设置／编辑图表／筛选那三张卡片）。
 * 对话框挂在 `document.body` 上，不是编辑器的宿主里——`render()` 会把宿主整个换掉。
 */
export type WordCountReadout = {
  paragraphs: number;
  characters: number;
  charactersNoSpace: number;
  cjk: number;
  words: number;
};

const STYLE_ID = "solo-word-count-styles";

const CSS = `
.solo-wc-dialog {
  width: 320px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-wc-dialog form { display: grid; gap: 12px; padding: 14px 16px 16px; }
.solo-wc-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-wc-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-wc-dialog .solo-wc-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-wc-rows { display: grid; gap: 4px; margin: 0; }
.solo-wc-rows div { display: flex; gap: 10px; justify-content: space-between; }
.solo-wc-rows dt, .solo-wc-rows dd { margin: 0; font-size: 12.5px; }
.solo-wc-rows dt { color: #4b5563; }
.solo-wc-rows dd { font-variant-numeric: tabular-nums; }
/* 「字数」是这一堆里用户来看的那个数，所以加粗；余下几行是它的来路。 */
.solo-wc-total dd { font-weight: 600; }
.solo-wc-sel-title { margin: 0; font-size: 12px; font-weight: 600; color: #4b5563; }
/* 两节挨着排，中间那道线把「整篇」与「所选」分开——不写这行，光看数字分不清哪一节是谁。 */
.solo-wc-sel { border-top: 1px solid #e5e9f2; padding-top: 8px; }
.solo-wc-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-wc-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-wc-dialog .solo-wc-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/**
 * 打开读数。这条命令不改件，所以「确定」与「关闭」都只是收摊；一次只挂一个，ribbon 上再点一次换掉旧的。
 *
 * 有选区时多挂一节「所选文字」：WPS 那一份统计接口本来就是**按目标对象**算的
 * （`office6/addons/kwpsapiex/wps.js` 的 `getDocumentWordCount(target)` 里 `target || __target || Application.ActiveDocument`，
 * 再交给 `ComputeStatistics(wdStatistic…)`），把选区当目标传进去就是选中部分的数——所以这一节**不是猜的行为**，
 * 是包里的函数签名允许的。**「所选文字」这四个字与这一节的行序仍是未取证**（同 §48 那堵 `.qm` 墙），
 * 按 Word 一族的通用叫法写。没有选区（折叠光标）时这一节整个不出现，而不是报一排 0。
 */
export function openWordCount(counts: WordCountReadout, selection?: WordCountReadout): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-wc-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-wc-dialog";
  dialog.setAttribute("aria-label", "字数统计");
  const row = (label: string, value: number, cls = ""): string =>
    `<div${cls ? ` class="${cls}"` : ""}><dt>${label}</dt><dd>${value}</dd></div>`;
  // 「中文字符和朝鲜语单词／非中文单词」那两行是「字数」的**来路**，只在整篇那一节排一次：
  // Word 一族的那一节在选区块里就是少这几行。这一条是**推断**（真机的行序未取证），排多排少都不改数。
  const rowsOf = (c: WordCountReadout, full = true): string =>
    `<div class="solo-wc-rows${full ? "" : " solo-wc-sel"}" role="status">
    ${row("段落数", c.paragraphs)}
    ${row("字数", c.cjk + c.words, "solo-wc-total")}
    ${row("字符数（不计空格）", c.charactersNoSpace)}
    ${row("字符数（计空格）", c.characters)}
    ${full ? `${row("中文字符和朝鲜语单词", c.cjk)}${row("非中文单词", c.words)}` : ""}
    </div>`;
  dialog.innerHTML =
    `<form><header><h2>字数统计</h2><button type="button" class="solo-wc-close" aria-label="关闭">×</button></header>
    ${rowsOf(counts)}
    ${selection ? `<h3 class="solo-wc-sel-title">所选文字</h3>${rowsOf(selection, false)}` : ""}
    <footer><button type="button" class="solo-wc-cancel">关闭</button><button type="submit" class="solo-wc-apply">确定</button></footer></form>`;

  const close = (): void => {
    dialog.close();
    dialog.remove();
  };
  dialog.querySelector<HTMLButtonElement>(".solo-wc-close")!.addEventListener("click", close);
  dialog.querySelector<HTMLButtonElement>(".solo-wc-cancel")!.addEventListener("click", close);
  dialog.querySelector("form")!.addEventListener("submit", (e) => {
    e.preventDefault();
    close();
  });
  // The page owns ⌘-style bindings and arrow keys; a dialog that let them fall through would move the caret
  // the reader is not looking at.
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  document.body.append(dialog);
  dialog.showModal();
  dialog.querySelector<HTMLButtonElement>(".solo-wc-apply")!.focus();
}
