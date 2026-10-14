/**
 * 「定位」（Go To）：kuip 里这条命令是 `aid="757"`、`id="GoTo"`、`idMso="GoTo"`、`dialogHint="1"`、
 * `hotKey="g"`、`rbHotKey="G"`、`shortcutHint="Ctrl+G"`（`office6/res/wpscommon.kuip:225`——与 WordCount 那条
 * `aid="792"` 同样的形状）。快捷键表 1127/1128（另一处 1316/1317 逐字重复一组）给了一对 `Ctrl+G`／`Meta+G`，
 * macOS 上就是 ⌘G，接在 `renderer.ts` 那条窗口级 keydown 分发里（必须排在那条不带 Shift 的 g 上——
 * ⌘⇧G 是字数统计）。状态栏的 `KxWpsStatusPageNumberInfoCommand`／`KxWpsStatusPageAreaInfoCommand`／
 * `KxWpsStatusSectionInfoCommand` 三格（447/449/451 行；451 的 detail 是「@Page Section. Click to open
 * "Go To" dialog.」）的 detail 都写着「Click to open "Go To" dialog.」：本页与 `doc-status-bar.ts` 的三格、
 * 快捷键，是同一条命令的四种长相——与「字数统计」那三处（ribbon／状态栏／快捷键）同一形状。
 *
 * **现在有「页」「节」「行」「书签」「批注」「表格」「图形」「标题」八个目标**：kuip 的浏览器一族（533 行 `KxWpsBrowserCmdTriggerCommand
 * id="BrowseGoTo"`）与 Word 的「定位目标」列表里还有 脚注/公式 等目标；模型里这一刻既落得了点、
 * 又有回路的只有这八个——包里那条 API 就是 `doc.GoTo(wdGoToPage, wdGoToAbsolute, page).Start`
 * （`office6/addons/kwpsapiex/wps.js`，`wdGoToSection` 一族同理）。列表里只摆点得动的目标名：剩下的等
 * 各自的回路接上再进列表。八个目标的种子与合法范围都由打开这一刻的读数给出——页数来自 `caretPage`
 * 的测量，节数来自 `sectionEndsOf` 数的文件声明边界（口径见 `doc-layout.ts` 文件头），行数与光标行号来自
 * 同一趟铺排（`measurePageLayout` 的 `lines` 与 `lineNumberAt`），批注条数来自 `doc.comments` 的条数，
 * 表格张数与图形个数来自同一趟模型走查（`tableCount`／`graphicCount`），标题个数与光标所在标题号来自
 * 导航窗格那趟大纲走查（`outlineEntries` 的长度与 `outlineMarks()` 里那枚标记的下标）。
 *
 * 「书签」这一格的候选是**名字下拉**，名单来自 `DocEditor.bookmarks()`（`w:bookmarkStart` 的 `w:name`，
 * 按文档顺序、同名取第一枚）。含 `_Toc`／`_GoBack` 这类机器书签的名字**不过滤**：那一批 86 份 WPS 写的件
 * 里 86/86 都带 `_` 前缀的名字，而真机的「隐藏书签」开关**未取证**——过滤掉等于把列表做成空的。
 * 没有书签的文档，这一格按时说明并挡住「定位」；名单空与名字不存在是两码事，前者要说得出来。
 *
 * 「批注」这一格的号码就是**纸上批注的编号**：`render.ts` 那枚 `numbers` 映射是 `doc.comments` 的下标 + 1
 * （回复也算一枚，「2 号」可能不显在纸上——回复自己没有标记），这一格数的是同一个数组，「打哪个号落哪条」
 * 与「纸上标着几号」不是两份算法。落点走审阅窗格那条「定位」（`DocEditor.revealComment`）：锚着文字的那条
 * 选中那段文字；**回复跳到父批注那一段**（Word 把回复挂在父批注的段上）；一个标记都找不回来的（锚在页眉带
 * 里、标记被重整吞掉）落不了点。kuip 的旁证是浏览器一族 530 行 `KxWpsBrowserTriggerCommand id="BrowseComment"
 * mode="3" target="3" icon="bubble_kd" text="@Browse by Comment"`（531/533 是并排的 BrowseBookmark／
 * BrowseGoTo）——「按批注浏览」真机上是独立一条命令。真机上这一格按什么顺序数、回复算不算一枚**未取证**；
 * 这里钉的是自己那条回路：号码与纸上编号同源。一条批注都没有的件，这一格按时说明并挡住「定位」。
 *
 * 「行」这一格的行号是**绝对行号**：从文首连数、换页不重数——这与状态栏「行」格（页内序数，`rowInPage`）
 * 不是同一个数，两者出自同一趟铺排（行 = 一枚行盒，表格按行计，0 高的断点记号不占行）；「本文共 N 行」的
 * N 就是 `measurePageLayout().lines`（与「字数统计」那一页同一份测量）。落点落在这一行**头一个字符**上
 * （`DocEditor.goToLine` → `lineCaretAnchor`：段内按上沿二分找行首，与「列」的底数同一趟取点）。WPS 真机
 * 「行」按绝对还是按页重数、落点在行首还是别处**未取证**；这里钉的是自己那条回路：号与行同一趟铺排数的。
 *
 * 「表格」这一格的号按**文档顺序**数，嵌套表算数——父表计入它自己的位置，其格里的表紧随其后（与「书签」
 * 名单的递归同一条：表内按行→格→段）。落点是那张表**第一行第一格里的第一个落得下的位置**（通常就是第一
 * 个段落头一个字符前；格子头一层的块里又有表的，就再往里一层）。WPS 真机数不数嵌套表、落点在格首还是
 * 别处**未取证**；这里钉的是自己那条回路：号与落点出自同一趟模型走查（`DocEditor.goToTable`）。
 *
 * 「图形」这一格数的是**行内图片**（`Inline.image`）：正文段落与表里（含嵌套表）都算，按文档顺序——与
 * 「表格」同一趟走查。落点是那枚图片在纸上的 `<img>` 元素**前面那一处**（图片是行内的一员，光标只能落在
 * 它旁边，进不了它里面）。种子归 1：光标「在图形里」这件事不存在（与书签那格同一条——没有「光标所在的
 * 那一枚」就从头数）。Word 的列表里这一项是「图形」（浮动形状也算图形）；本版只数行内图片，WPS 真机数不数
 * 浮动形状、落点在图形前还是别处**未取证**；这里钉的是自己那条回路：号与落点出自同一趟模型走查
 * （`DocEditor.goToGraphic`）。
 *
 * 「标题」这一格的号按**大纲顺序**数（`DocEditor.outlineEntries`：模型里带大纲级别的段落，头条即 1 号），
 * 表内（含嵌套表）的标题也算——与「表格」「图形」同一趟走查的读者顺序。落点走导航窗格那条「定位」
 * （`DocEditor.revealOutline`）：选中标题文字、光标停在**标题末尾**（窗格那条回路就是末尾——继续往下写
 * 是最常见的意思），纸上那枚标题打上高亮、滚进视野。种子是**光标所在的那枚标题**的号（光标停在标题段落
 * 里就给那一枚，停在正文给 1）——「标题」与「书签/图形」不同，光标真的「在」某一枚里。Word 真机这一格
 * 叫「标题」、`GoTo` 落点在标题开头、数不数表内标题**未取证**；这里钉的是自己那条回路：号与落点出自
 * 同一趟大纲走查，和导航窗格同一份读数。
 *
 * 中文界面里这一家的字面（标题、目标名、标签、按钮）**未取证**（`.qm` 那堵墙，与 `doc-word-count.ts`
 * 头注记的是同一件）；这里按 Word 一族的通用叫法写。对话框的形状（模态与否、「定位」之后关不关）真机上
 * **未取证**：本片与「字数统计」用同一套写法——模态、Esc 关，按「定位」把落点交出去就收摊。这一页不改件，
 * 所以不走 `changed()`；落光标由 `goToPage`/`goToSection`/`goToLine`/`goToBookmark`/`goToComment`/`goToTable`/`goToGraphic`/`goToHeading` 负责，
 * 回来时状态栏自己会重读（走 `selectionchange`）。**取消/关闭之后光标要回到开窗前的位置**：数字框一聚焦
 * 浏览器就把文档选区带进对话框，关窗后它会落去 BODY（隔离探针实测，见 `openGoTo` 里那段注）——开窗前存
 * 一份 Range，收尾时放回去；真机上 WPS 取消「定位」后光标在原处是常理，但我们这一手的实测只在无头
 * Chrome 里做过，Electron 真机**未取证**。
 *
 * 目标的合法范围由对话框自己先挡一道（页/节/行/批注/表格/图形/标题越界、非整数不关窗，就地给一句说明；
 * 书签名单为空、批注一条没有、一行内容都没有、一张表格/一个图形/一枚标题都没有也不关窗），
 * `goToPage`/`goToSection`/`goToLine`/`goToBookmark`/`goToComment`/`goToTable`/`goToGraphic`/`goToHeading`
 * 那层再挡一道——两道都不是猜：页数/节数/行数是打开这一刻量出来的「共 N 页/节/行」，书签名单与批注条数、
 * 表格张数、图形个数、标题个数是同一刻从模型数出来的。
 */
/** 「定位目标」这一格的八个目标：页、节、行、书签、批注、表格、图形、标题。 */
type Target = "page" | "section" | "line" | "bookmark" | "comment" | "table" | "graphic" | "heading";

const STYLE_ID = "solo-goto-styles";

const CSS = `
.solo-goto-dialog {
  width: 300px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-goto-dialog form { display: grid; gap: 12px; padding: 14px 16px 16px; }
.solo-goto-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-goto-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-goto-dialog .solo-goto-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-goto-row { display: flex; align-items: center; gap: 10px; }
.solo-goto-row[hidden] { display: none; }
.solo-goto-row label { white-space: nowrap; color: #4b5563; }
.solo-goto-row select, .solo-goto-row input { flex: 1; min-width: 0; padding: 5px 8px; border: 1px solid #c9d3e6; border-radius: 7px; font: inherit; }
.solo-goto-note { margin: 0; font-size: 12px; color: #6b7280; }
.solo-goto-error { margin: 0; font-size: 12px; color: #b00020; }
.solo-goto-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-goto-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-goto-dialog .solo-goto-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/** 这一格「一枚都没有」时的说法；页/节永远有（一张空纸也是一页），书签空在候选那支另说。 */
function noneText(target: Target): string | undefined {
  if (target === "comment") return "本文没有批注。";
  if (target === "line") return "本文没有内容行。";
  if (target === "table") return "本文没有表格。";
  if (target === "graphic") return "本文没有图形。";
  if (target === "heading") return "本文没有标题。";
  return undefined;
}

/**
 * 打开「定位」。`seed` 是打开这一刻的种子与范围：`page`/`section`/`line`/`comment`/`table`/`heading` 是光标
 * 所在的页号/节号/行号/批注号/表格号/标题号（`graphic` 没有「光标所在的」一枚，永远给 1），`pages`/`sections`/
 * `lines`/`comments`/`tables`/`graphics`/`headings` 是量出来或数出来的共几页/几节/几行/几条批注/几张表格/
 * 几个图形/几枚标题（合法范围的来源），`bookmarks` 是书签名字按文档顺序的名单（「书签」目标的候选），
 * `target` 是预选的目标（点状态栏哪一格就开哪个）。`goTo` 是各目标合法的落点
 * （`DocEditor.goToPage`/`goToSection`/`goToLine`/`goToBookmark`/`goToComment`/`goToTable`/`goToGraphic`/`goToHeading`）。
 * 一次只挂一个，再点一次换掉旧的。
 */
export function openGoTo(
  seed: { page: number; pages: number; section: number; sections: number; line: number; lines: number; bookmarks: readonly string[]; comment: number; comments: number; table: number; tables: number; graphic: number; graphics: number; heading: number; headings: number; target: Target },
  goTo: { page: (n: number) => void; section: (n: number) => void; line: (n: number) => void; bookmark: (name: string) => void; comment: (n: number) => void; table: (n: number) => void; graphic: (n: number) => void; heading: (n: number) => void },
): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-goto-dialog").forEach((old) => {
    if (old.open) old.close();
    old.remove();
  });
  injectStyles();
  const cur = { page: seed.page, section: seed.section, line: seed.line, comment: seed.comment, table: seed.table, graphic: seed.graphic, heading: seed.heading };
  const most = { page: seed.pages, section: seed.sections, line: seed.lines, comment: seed.comments, table: seed.tables, graphic: seed.graphics, heading: seed.headings };
  const unit = { page: "页", section: "节", line: "行", comment: "批注", table: "表格", graphic: "图形", heading: "标题" };
  const dialog = document.createElement("dialog");
  dialog.className = "solo-goto-dialog";
  dialog.setAttribute("aria-label", "定位");
  // `novalidate`：min/max/step 那套原生校验会先把 submit 事件吃掉（越界、非整数时只发 `invalid`、不发
  // `submit`），就地报一句的那条路就永远走不到。范围仍写在 min/max 上，那是给读者的提示，不是闸门；
  // 真闸门是下面 submit 里那行检查。
  dialog.innerHTML =
    `<form novalidate><header><h2>定位</h2><button type="button" class="solo-goto-close" aria-label="关闭">×</button></header>
    <div class="solo-goto-row"><label for="solo-goto-target">定位目标</label><select id="solo-goto-target"><option value="page">页</option><option value="section">节</option><option value="line">行</option><option value="bookmark">书签</option><option value="comment">批注</option><option value="table">表格</option><option value="graphic">图形</option><option value="heading">标题</option></select></div>
    <div class="solo-goto-row solo-goto-number"><label for="solo-goto-value">输入页号</label><input id="solo-goto-value" type="number" min="1" step="1"></div>
    <div class="solo-goto-row solo-goto-name" hidden><label for="solo-goto-bookmark">输入书签名</label><select id="solo-goto-bookmark"></select></div>
    <p class="solo-goto-note"></p>
    <p class="solo-goto-error" hidden></p>
    <footer><button type="button" class="solo-goto-cancel">关闭</button><button type="submit" class="solo-goto-apply">定位</button></footer></form>`;

  const select = dialog.querySelector<HTMLSelectElement>("#solo-goto-target")!;
  const numberRow = dialog.querySelector<HTMLElement>(".solo-goto-number")!;
  const nameRow = dialog.querySelector<HTMLElement>(".solo-goto-name")!;
  const label = dialog.querySelector<HTMLLabelElement>("label[for='solo-goto-value']")!;
  const input = dialog.querySelector<HTMLInputElement>("#solo-goto-value")!;
  const pick = dialog.querySelector<HTMLSelectElement>("#solo-goto-bookmark")!;
  const note = dialog.querySelector<HTMLParagraphElement>(".solo-goto-note")!;
  const error = dialog.querySelector<HTMLParagraphElement>(".solo-goto-error")!;
  for (const name of seed.bookmarks) pick.append(new Option(name, name));
  // 换目标时那一格整体换一套：页/节/行/批注走数字框，书签走名字下拉——下拉里只摆名单里的名字（`bookmarks()`
  // 数的），「名字打错」这条路从根上不存在。数字不跨目标沿用——「5」从页挪到节没有意义，页/节/行/批注各自从
  // 打开这一刻的读数（光标所在页/节/行/批注）重来；书签没有「光标所在的那一枚」，下拉从第一枚起。
  const apply = (target: Target): void => {
    const named = target === "bookmark";
    numberRow.hidden = named;
    nameRow.hidden = !named;
    error.hidden = true;
    if (target === "bookmark") {
      note.textContent = seed.bookmarks.length ? `本文共 ${seed.bookmarks.length} 个书签。` : "本文没有书签。";
      return;
    }
    label.textContent = `输入${unit[target]}号`;
    input.min = "1";
    input.max = String(most[target]);
    input.value = String(cur[target]);
    // 批注的量词是「条」、表格是「张」、图形/标题是「个/枚」不是目标名；一条没有／一行内容都没有／一张表格、
    // 一个图形、一枚标题都没有的件说得出来，与空书签名单同一条路（submit 里也挡）。
    const none = noneText(target);
    note.textContent =
      none && most[target] === 0
        ? none
        : target === "comment"
          ? `本文共 ${most.comment} 条批注。`
          : target === "table"
            ? `本文共 ${most.table} 张表格。`
            : target === "graphic"
              ? `本文共 ${most.graphic} 个图形。`
              : target === "heading"
                ? `本文共 ${most.heading} 个标题。`
                : `本文共 ${most[target]} ${unit[target]}。`;
  };
  // 数字框一聚焦，浏览器就把**文档选区**带进对话框（实测 headless Chrome：开→对话框里，关+摘除→BODY；
  // 焦点落在按钮/下拉上时不动选区——隔离探针 `.scratch/sel-close-probe.html` 的矩阵）。关窗之后光标该回到
  // 开窗前的原处：开窗前存一份 Range，收尾时放回去。四格统一走这一手——书签那格实测不偷，但放回一份没动过
  // 的 Range 是无操作，不值当为它分一条路。
  const sel = document.getSelection();
  const saved = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const restore = (): void => {
    if (!saved || !sel) return;
    sel.removeAllRanges();
    sel.addRange(saved);
  };
  const close = (): void => {
    dialog.close();
    dialog.remove();
    restore();
  };
  // Esc 是浏览器自己关的（那条路不经过 close()）：收尾在 `close` 事件里补一份。close() 那条路连摘除都
  // 做完了（isConnected 是 false），这里就跳过——按「定位」的落点也就不会被这一手覆盖掉。
  dialog.addEventListener("close", () => {
    if (!dialog.isConnected) return;
    dialog.remove();
    restore();
  });
  dialog.querySelector<HTMLButtonElement>(".solo-goto-close")!.addEventListener("click", close);
  dialog.querySelector<HTMLButtonElement>(".solo-goto-cancel")!.addEventListener("click", close);
  select.addEventListener("change", () => apply(select.value as Target));
  input.addEventListener("input", () => {
    error.hidden = true;
  });
  dialog.querySelector("form")!.addEventListener("submit", (e) => {
    e.preventDefault();
    const target = select.value as Target;
    if (target === "bookmark") {
      const name = pick.value;
      // 名单为空时下拉是空的（`pick.value === ""`），名单非空则必有一枚选中——只有前一种情况挡在这里。
      if (name === "") {
        error.textContent = "本文没有书签。";
        error.hidden = false;
        return;
      }
      close();
      goTo.bookmark(name);
      return;
    }
    // 一条批注都没有、一行内容都没有、一张表格/一个图形都没有：任何整数号都落不了，就地按「没有」挡——
    // 落到下面那条通用范围检查上会报一句「请输入 1 到 0 之间的整数……」，那句话对读者不成立。
    const none = noneText(target);
    if (none && most[target] === 0) {
      error.textContent = none;
      error.hidden = false;
      return;
    }
    const n = Number(input.value.trim());
    if (!Number.isInteger(n) || n < 1 || n > most[target]) {
      error.textContent = `请输入 1 到 ${most[target]} 之间的整数${unit[target]}号。`;
      error.hidden = false;
      input.focus();
      input.select();
      return;
    }
    close();
    goTo[target](n);
  });
  // The page owns ⌘-style bindings and arrow keys; a dialog that let them fall through would move the caret
  // the reader is not looking at.
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  document.body.append(dialog);
  select.value = seed.target;
  apply(seed.target);
  dialog.showModal();
  const first = seed.target === "bookmark" ? pick : input;
  first.focus();
  if (first === input) input.select();
}
