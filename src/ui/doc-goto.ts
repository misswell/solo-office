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
 * **现在有「页」「节」「书签」三个目标**：kuip 的浏览器一族（533 行 `KxWpsBrowserCmdTriggerCommand id="BrowseGoTo"`）
 * 与 Word 的「定位目标」列表里还有 行/批注/脚注/表格/图形/公式 等目标；模型里这一刻既落得了点、
 * 又有回路的只有这三个——包里那条 API 就是 `doc.GoTo(wdGoToPage, wdGoToAbsolute, page).Start`
 * （`office6/addons/kwpsapiex/wps.js`，`wdGoToSection` 一族同理）。列表里只摆点得动的目标名：剩下的等
 * 各自的回路接上再进列表。三个目标的种子与合法范围都由打开这一刻的读数给出——页数来自 `caretPage`
 * 的测量，节数来自 `sectionEndsOf` 数的文件声明边界（口径见 `doc-layout.ts` 文件头）。
 *
 * 「书签」这一格的候选是**名字下拉**，名单来自 `DocEditor.bookmarks()`（`w:bookmarkStart` 的 `w:name`，
 * 按文档顺序、同名取第一枚）。含 `_Toc`／`_GoBack` 这类机器书签的名字**不过滤**：那一批 86 份 WPS 写的件
 * 里 86/86 都带 `_` 前缀的名字，而真机的「隐藏书签」开关**未取证**——过滤掉等于把列表做成空的。
 * 没有书签的文档，这一格按时说明并挡住「定位」；名单空与名字不存在是两码事，前者要说得出来。
 *
 * 中文界面里这一家的字面（标题、目标名、标签、按钮）**未取证**（`.qm` 那堵墙，与 `doc-word-count.ts`
 * 头注记的是同一件）；这里按 Word 一族的通用叫法写。对话框的形状（模态与否、「定位」之后关不关）真机上
 * **未取证**：本片与「字数统计」用同一套写法——模态、Esc 关，按「定位」把落点交出去就收摊。这一页不改件，
 * 所以不走 `changed()`；落光标由 `goToPage`/`goToSection`/`goToBookmark` 负责，回来时状态栏自己会重读
 * （走 `selectionchange`）。
 *
 * 目标的合法范围由对话框自己先挡一道（页/节越界、非整数不关窗，就地给一句说明；书签名单为空不关窗），
 * `goToPage`/`goToSection`/`goToBookmark` 那层再挡一道——两道都不是猜：页数/节数是打开这一刻量出来的
 * 「共 N 页/节」，书签名单是同一刻从模型数出来的。
 */
/** 「定位目标」这一格的三个目标：页、节、书签。 */
type Target = "page" | "section" | "bookmark";

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

/**
 * 打开「定位」。`seed` 是打开这一刻的种子与范围：`page`/`section` 是光标所在的页号/节号，`pages`/`sections`
 * 是量出来的共几页/几节（合法范围的来源），`bookmarks` 是书签名字按文档顺序的名单（「书签」目标的候选），
 * `target` 是预选的目标（点状态栏哪一格就开哪个）。`goTo` 是各目标合法的落点
 * （`DocEditor.goToPage`/`goToSection`/`goToBookmark`）。一次只挂一个，再点一次换掉旧的。
 */
export function openGoTo(
  seed: { page: number; pages: number; section: number; sections: number; bookmarks: readonly string[]; target: Target },
  goTo: { page: (n: number) => void; section: (n: number) => void; bookmark: (name: string) => void },
): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-goto-dialog").forEach((old) => {
    if (old.open) old.close();
    old.remove();
  });
  injectStyles();
  const cur = { page: seed.page, section: seed.section };
  const most = { page: seed.pages, section: seed.sections };
  const unit = { page: "页", section: "节" };
  const dialog = document.createElement("dialog");
  dialog.className = "solo-goto-dialog";
  dialog.setAttribute("aria-label", "定位");
  // `novalidate`：min/max/step 那套原生校验会先把 submit 事件吃掉（越界、非整数时只发 `invalid`、不发
  // `submit`），就地报一句的那条路就永远走不到。范围仍写在 min/max 上，那是给读者的提示，不是闸门；
  // 真闸门是下面 submit 里那行检查。
  dialog.innerHTML =
    `<form novalidate><header><h2>定位</h2><button type="button" class="solo-goto-close" aria-label="关闭">×</button></header>
    <div class="solo-goto-row"><label for="solo-goto-target">定位目标</label><select id="solo-goto-target"><option value="page">页</option><option value="section">节</option><option value="bookmark">书签</option></select></div>
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
  // 换目标时那一格整体换一套：页/节走数字框，书签走名字下拉——下拉里只摆名单里的名字（`bookmarks()`
  // 数的），「名字打错」这条路从根上不存在。数字不跨目标沿用——「5」从页挪到节没有意义，页/节各自从
  // 打开这一刻的读数（光标所在页/节）重来；书签没有「光标所在的那一枚」，下拉从第一枚起。
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
    note.textContent = `本文共 ${most[target]} ${unit[target]}。`;
  };
  const close = (): void => {
    dialog.close();
    dialog.remove();
  };
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
