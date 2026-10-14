/**
 * 「字数统计」：WPS 的这条命令是 aid 792、`id="WordCount"`、`idMso="WordCount"`、`dialogHint="1"`，快捷键
 * `Ctrl+Shift+G` / `Meta+Shift+G`（`.kuip` 的 `wpsentries.kuip`）。同一份声明里还有
 * `KxWpsStatusWordCountInfoCommand`（`aid="792"`、`minimumWidth="75"`，detail 说的是「点一下开字数统计对话框」）
 * 与 `WordCountStatusBar`，也就是**状态栏那个字数是同一件命令的另一种长相**——本轮只做对话框那一条腿。
 *
 * 数不是这里数的，是 `WordDocument.wordCounts()` 数出来的（口径与取舍都写在那一处）；这个文件只把它们排成
 * 一行行的读数。`字数` 那行报的是 `中文字符和朝鲜语单词 + 非中文单词`，与 Word 一族那三行的连法相同。
 *
 * 「行数」这一行**不报**：它要的是排版之后的结果——一行在哪里断由字体度量和栏宽决定，模型里没有这个数，
 * 编一个出来就是假读数。「所选文字」那一栏同理：有选区时 WPS 会另报选中部分的字数，本轮只数全文。
 *
 * 对话框标题那句「字数统计」在包里能找到（`office6/addons/kwpsapiex/wps.js`），但**各行标签的中文原话
 * 取不到**：`.qm`/`.rcc` 编译过的翻译里 grep 不到这些串（与「增加缩进量」那次同一堵墙），所以下面这几行
 * 标签是按 Word 一族的通用叫法写的，WPS 真机上逐行的措辞仍是未取证。
 *
 * 样式从这个文件注入：`src/ui/index.html` 由构建原样复制，不归我们改。对话框挂在 `document.body` 上，
 * 不是编辑器的宿主里——`render()` 会把宿主整个换掉。
 */
export type WordCountReadout = {
  pages: number;
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
 */
export function openWordCount(counts: WordCountReadout): void {
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
  dialog.innerHTML =
    `<form><header><h2>字数统计</h2><button type="button" class="solo-wc-close" aria-label="关闭">×</button></header>
    <div class="solo-wc-rows" role="status">
    ${row("页数", counts.pages)}
    ${row("段落数", counts.paragraphs)}
    ${row("字数", counts.cjk + counts.words, "solo-wc-total")}
    ${row("字符数（不计空格）", counts.charactersNoSpace)}
    ${row("字符数（计空格）", counts.characters)}
    ${row("中文字符和朝鲜语单词", counts.cjk)}
    ${row("非中文单词", counts.words)}
    </div>
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
