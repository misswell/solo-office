import type { PageSetup, Section } from "../core/docx.js";
import { cmToTwips } from "../core/docx.js";
import { currentUnit, twipText, twipsFromUnit, unitDef } from "./doc-units.js";
import type { Unit } from "./doc-units.js";

/**
 * 「页面设置」: the sheet the page is cut from, which way it lies, the four body margins, the two paper-edge
 * distances, and the binding width. The dialog counts in the status bar's 测量单位 (`doc-units.ts`), and hands
 * the model twips — `w:pgMar` and `w:pgSz` have no other unit (ECMA-376 §17.3.2.13, §17.3.2.12) — converting
 * once here so no number is rounded twice.
 *
 * Three distances may be left **empty**, which changes nothing: they are optional in the part, and a file that
 * says nothing about its binding width stays a file that says nothing. The four body margins and the paper's
 * two edges are always sent, because the page on screen is drawn from them and a save has to write the box it
 * showed.
 *
 * Prefilled to each unit's **twip-exact** decimals — the fewest that bring every whole twip count from 0 to
 * 60000 back as itself (`doc-units.ts` spells the two inequalities; the traversal is `tests/doc-units.test.ts`):
 * 厘米 3, 毫米 2, 英寸 4, 磅 2. 851 twips is 1.501 厘米, and the 1.50 a two-decimal box would show comes back as
 * 850; the same number in inches is 0.591, because 0.59 would come back as 850 too. That is why a dialog nobody
 * retouched cannot move a margin or an edge, and why a sheet that is not on the list — the near-A4 11900×16840
 * four of the corpus sections spell — stays 11900×16840 instead of being pulled to A4. The paper-name row
 * shows the file's own twips to the last digit: A4's 11906 reads 21.001 厘米, not the 21 a name-based table
 * might print, because in inches the same edge has no short spelling at all (8.2681) and one rule has to hold
 * in all four units.
 *
 * Styles are injected from here: `src/ui/index.html` is copied verbatim by the build and is not ours to touch.
 * The dialog is mounted on `document.body` rather than the editor's host, which `render()` replaces whole.
 */
const STYLE_ID = "solo-page-setup-styles";

const CSS = `
.solo-ps-dialog {
  width: 400px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-ps-dialog form { display: grid; gap: 14px; padding: 14px 16px 16px; }
.solo-ps-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-ps-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-ps-dialog .solo-ps-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-ps-fields, .solo-ps-paper { display: grid; gap: 10px 12px; }
/* minmax(0, ...) because a track's floor is otherwise the input's own preferred width: measured at 162px
 * here, two of those left the first column 14px of the dialog's 362 — a list nobody can read. With the floor at
 * zero the tracks split by their factors, and the min-width:0 input inside shrinks to match. */
.solo-ps-fields { grid-template-columns: repeat(2, minmax(0, 1fr)); }
.solo-ps-paper { grid-template-columns: minmax(0, 1.3fr) minmax(0, 1fr) minmax(0, 1fr); }
/* The same floor one level down, or the fix above only moves the overflow: a label's own implicit column is
 * sized by the row beneath it, and that row's floor is the input again. */
.solo-ps-fields label, .solo-ps-paper label { display: grid; grid-template-columns: minmax(0, 1fr); gap: 3px; font-size: 12px; color: #4b5563; }
.solo-ps-num { display: flex; align-items: center; gap: 6px; min-width: 0; }
.solo-ps-fields input, .solo-ps-paper input, .solo-ps-paper select {
  flex: 1; min-width: 0; padding: 4px 6px; font: inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-ps-fields span, .solo-ps-paper span { color: #6b7280; }
.solo-ps-fields small { color: #6b7280; font-size: 11px; }
/* Last of the label rules that could match, so the horizontal spelling of these two wins. */
.solo-ps-orient { display: flex; gap: 14px; align-items: center; font-size: 12px; color: #4b5563; }
.solo-ps-orient label { display: flex; gap: 5px; align-items: center; }
.solo-ps-paper .solo-ps-orient { grid-column: 1 / -1; }
.solo-ps-error { margin: 0; min-height: 18px; color: #b42318; font-size: 12px; }
.solo-ps-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-ps-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-ps-dialog .solo-ps-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
.solo-ps-dialog .solo-ps-apply[disabled] { opacity: .5; cursor: default; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/** 一个缇数按当前测量单位写进格子的字面；`undefined`（文件没说的距离）是空格子。 */
const unitText = (tw: number | undefined): string => (tw === undefined ? "" : twipText(tw));

/**
 * The sheets 「页面设置」 offers by name. `w:pgSz` carries no names — §17.3.2.12 stores two distances and every
 * section spells its own pair (授权语料 46 份 WPS 自带模板的 58 节，`word/document.xml` 里 **58 节全写 `w:pgSz`**，
 * `.scratch/authorized-families2.mjs` 与 ElementTree、xmllint 三口径一致) — so this list is the dialog's
 * convenience, not a vocabulary the part has. Its centimetres are the sizes as their own standards state them
 * (ISO 216 counts in whole tenths of a centimetre; 8.5, 11 and 14 inches are 21.59, 27.94 and 35.56 cm), and
 * each converts to the twip count a producer writes for that sheet: 「A4」 gives 11906×16838, the pair **56 of the
 * 58** spell (the other 2 write 11907×16840 — 同一张纸的另一档取整). A sheet that is not here is 自定义, and its
 * two boxes are what the save owes.
 *
 * 16 开 and 32 开 are deliberately absent: the numbers WPS gives them were never read off a dialog, and a
 * guess under a standard-looking name is how a document ends up on the wrong paper.
 */
const PAPER: readonly { name: string; shortCm: number; longCm: number }[] = [
  { name: "A3", shortCm: 29.7, longCm: 42 },
  { name: "A4", shortCm: 21, longCm: 29.7 },
  { name: "A5", shortCm: 14.8, longCm: 21 },
  { name: "B4", shortCm: 25, longCm: 35.3 },
  { name: "B5", shortCm: 17.6, longCm: 25 },
  { name: "信纸", shortCm: 21.59, longCm: 27.94 },
  { name: "法律专用纸", shortCm: 21.59, longCm: 35.56 },
];
const CUSTOM = "自定义尺寸";

/** What one box holds: nothing typed means "not asked about", and anything unparseable is refused. */
type Reading = { value?: number; bad: boolean };

const read = (input: HTMLInputElement, u: Unit): Reading => {
  const text = input.value.trim();
  if (!text) return { bad: false };
  const v = Number(text);
  // No upper bound is invented: `w:pgMar` takes any twip count, and a producer's own limit is not evidenced
  // in this corpus. What is refused is a value that is not a number at all, or a negative distance.
  if (!Number.isFinite(v) || v < 0) return { bad: true };
  return { value: twipsFromUnit(v, u), bad: false };
};

/** A paper edge: the box has to say a sheet's side, so empty and zero are refused rather than ignored. */
const readEdge = (input: HTMLInputElement, u: Unit): number | undefined => {
  const v = Number(input.value.trim());
  return Number.isFinite(v) && v > 0 ? twipsFromUnit(v, u) : undefined;
};

/**
 * Opens the dialog for the section the page is drawn from. `current` is what the file states — a distance it
 * never stated arrives as no value, and its box is left empty rather than filled with a guess.
 */
export function openPageSetup(current: Section, apply: (set: PageSetup) => void): void {
  // One dialog at a time: reopening from the ribbon replaces what is on screen instead of stacking a second.
  document.querySelectorAll<HTMLDialogElement>(".solo-ps-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-ps-dialog";
  dialog.setAttribute("aria-label", "页面设置");
  // The unit is read once, at open: the dialog is modal, so 「选项」 cannot change it while the boxes are up.
  const u = currentUnit();
  const unitName = unitDef(u).name;
  const box = (name: string, label: string, value: string, hint = ""): string =>
    `<label>${label}<span class="solo-ps-num"><input name="${name}" inputmode="decimal" autocomplete="off" value="${value}"><span>${unitName}</span></span>${hint}</label>`;
  dialog.innerHTML =
    `<form><header><h2>页面设置</h2><button type="button" class="solo-ps-close" aria-label="关闭">×</button></header>
    <div class="solo-ps-paper">
    <label>纸张大小<select name="ps-paper">${[...PAPER.map((p) => p.name), CUSTOM]
      .map((n) => `<option value="${n}">${n}</option>`)
      .join("")}</select></label>
    ${box("ps-width", "宽", unitText(current.pageWidthTwips))}
    ${box("ps-height", "高", unitText(current.pageHeightTwips))}
    <label class="solo-ps-orient">纸张方向
      <span><label><input type="radio" name="ps-orient" value="portrait">纵向</label>
      <label><input type="radio" name="ps-orient" value="landscape">横向</label></span>
    </label>
    </div>
    <div class="solo-ps-fields">
    ${box("ps-top", "上", unitText(current.margins.top))}
    ${box("ps-bottom", "下", unitText(current.margins.bottom))}
    ${box("ps-left", "左", unitText(current.margins.left))}
    ${box("ps-right", "右", unitText(current.margins.right))}
    ${box("ps-header", "页眉距离", unitText(current.margins.header), "<small>从纸张边缘量；留空表示不改</small>")}
    ${box("ps-footer", "页脚距离", unitText(current.margins.footer), "<small>从纸张边缘量；留空表示不改</small>")}
    ${box("ps-gutter", "装订线", unitText(current.margins.gutter), "<small>加在装订侧的页边距上；留空表示不改</small>")}
    </div>
    <p class="solo-ps-error" role="status"></p>
    <footer><button type="button" class="solo-ps-cancel">取消</button><button type="submit" class="solo-ps-apply">确定</button></footer></form>`;

  const field = (name: string): HTMLInputElement => dialog.querySelector(`[name="${name}"]`)!;
  const inputs = {
    width: field("ps-width"),
    height: field("ps-height"),
    top: field("ps-top"),
    bottom: field("ps-bottom"),
    left: field("ps-left"),
    right: field("ps-right"),
    header: field("ps-header"),
    footer: field("ps-footer"),
    gutter: field("ps-gutter"),
  };
  const paper = dialog.querySelector<HTMLSelectElement>('select[name="ps-paper"]')!;
  const error = dialog.querySelector<HTMLElement>(".solo-ps-error")!;
  const applyButton = dialog.querySelector<HTMLButtonElement>(".solo-ps-apply")!;
  const isLandscape = (): boolean => dialog.querySelector<HTMLInputElement>('input[name="ps-orient"]:checked')?.value === "landscape";
  (dialog.querySelector(`[name="ps-orient"][value="${current.pageWidthTwips > current.pageHeightTwips ? "landscape" : "portrait"}"]`) as HTMLInputElement).checked = true;

  /**
   * The sheet the two boxes add up to, ignoring which edge is which — the direction is asked about by its own
   * radios, and the model turns the pair itself. Unreadable boxes are no sheet, which the list shows as
   * 自定义 and `build` refuses.
   */
  const matchedPaper = (): (typeof PAPER)[number] | undefined => {
    const a = readEdge(inputs.width, u);
    const b = readEdge(inputs.height, u);
    if (a === undefined || b === undefined) return undefined;
    const long = Math.max(a, b);
    const short = Math.min(a, b);
    return PAPER.find((p) => cmToTwips(p.longCm) === long && cmToTwips(p.shortCm) === short);
  };

  /**
   * Lays the named sheet out along the direction the page is turned. Called when the user picks from the list
   * and once on open, never while a box is being typed in: the boxes are already re-read on every keystroke's
   * `input`, and retyping over the user's edit would swallow the keystroke that produced the match. The two
   * numbers written are the sheet's own twips in the current unit — A3's 42 厘米 come back exactly, and A4's
   * 11906 is written as 21.001 厘米 — so choosing a name and pressing 确定 writes the bytes that name has
   * always meant.
   */
  const fillPaper = (name: string): void => {
    const sheet = PAPER.find((p) => p.name === name);
    if (!sheet) return;
    const land = isLandscape();
    inputs.width.value = twipText(cmToTwips(land ? sheet.longCm : sheet.shortCm));
    inputs.height.value = twipText(cmToTwips(land ? sheet.shortCm : sheet.longCm));
  };

  /** The setup the boxes say, or nothing at all when a box cannot be read. */
  const build = (): PageSetup | undefined => {
    const width = readEdge(inputs.width, u);
    const height = readEdge(inputs.height, u);
    const taken = [read(inputs.top, u), read(inputs.bottom, u), read(inputs.left, u), read(inputs.right, u)];
    const optional = [read(inputs.header, u), read(inputs.footer, u), read(inputs.gutter, u)];
    const bad = taken.some((r) => r.bad) || optional.some((r) => r.bad);
    const missing = taken.some((r) => r.value === undefined);
    if (width === undefined || height === undefined) {
      error.textContent = `纸张宽与高要填大于 0 ${unitName}的数。`;
      applyButton.disabled = true;
      return undefined;
    }
    if (bad || missing) {
      error.textContent = bad ? "页边距与距离需填 0 或更大的数字。" : "上、下、左、右四个页边距都要填数字。";
      applyButton.disabled = true;
      return undefined;
    }
    error.textContent = "";
    applyButton.disabled = false;
    return {
      paperTwips: { width, height },
      topTwips: taken[0]!.value!,
      bottomTwips: taken[1]!.value!,
      leftTwips: taken[2]!.value!,
      rightTwips: taken[3]!.value!,
      headerTwips: optional[0].value,
      footerTwips: optional[1].value,
      gutterTwips: optional[2].value,
      landscape: isLandscape(),
    };
  };
  const refresh = (): void => {
    const set = build();
    // The list follows the boxes; the boxes never follow the list here, so typing is not interrupted.
    if (set) paper.value = matchedPaper()?.name ?? CUSTOM;
  };
  /**
   * Turn the two boxes with the page. This is for the number the user reads next to 「宽」: the model is handed
   * the pair plus the direction and puts the long edge where the direction says, so swapping cannot change a
   * byte that the same direction and the same sheet would not already have written.
   */
  const swapEdges = (): void => {
    const width = readEdge(inputs.width, u);
    const height = readEdge(inputs.height, u);
    if (width === undefined || height === undefined) return;
    // A sheet the list knows is refilled from the list rather than re-measured: a box the user left as the
    // file's own twips should turn into the sheet's own numbers, not stay whatever was typed there.
    const sheet = matchedPaper();
    if (sheet) fillPaper(sheet.name);
    else {
      inputs.width.value = unitText(height);
      inputs.height.value = unitText(width);
    }
  };
  /**
   * Closes and takes the node out of the document in the same step. The `close` event is a queued task, so
   * relying on it alone leaves a shut dialog lying in the DOM until the next one runs — and the next
   * 「页面设置」 opened before then would find *that* one first, filling boxes on a dialog nobody is looking at.
   *
   * 光标也在这里放回来：`showModal()` 之后 `inputs.top.focus()` 一聚焦输入框，浏览器就把**文档选区**搬进
   * 对话框（headless Chrome 实测：开→对话框里，关+摘除→BODY——与「定位」同一个浏览器行为，矩阵在
   * `.scratch/sel-close-probe.html`）。开窗前存一份 Range，收尾时放回去；按「确定」那条路也走这里，放回
   * 之后 `applyPageSetup` 才读得到开窗前的光标（它随后整张重排、按模型路径再落一次）。
   */
  const sel = document.getSelection();
  const saved = sel && sel.rangeCount ? sel.getRangeAt(0).cloneRange() : null;
  const restore = (): void => {
    if (!saved || !sel) return;
    sel.removeAllRanges();
    sel.addRange(saved);
  };
  const shut = (): void => {
    dialog.close();
    dialog.remove();
    restore();
  };
  dialog.addEventListener("input", refresh);
  dialog.addEventListener("change", (e) => {
    const target = e.target as HTMLElement;
    if (target instanceof HTMLSelectElement && target.name === "ps-paper") fillPaper(paper.value);
    if (target instanceof HTMLInputElement && target.name === "ps-orient") swapEdges();
    refresh();
  });
  // The editor owns the page's shortcuts; a dialog typing over them would move the caret behind it.
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  // Escape closes the dialog without going through any of the buttons, so this is the handler that path has.
  // `shut()` 那条路连摘除都做完了（isConnected 是 false），这里就跳过——按「确定」之后已经被整张重排
  // 换掉的那份押着的 Range 也就不会盖回来（放回它等于把选区指回一堆断开的节点）。
  dialog.addEventListener("close", () => {
    if (!dialog.isConnected) return;
    dialog.remove();
    restore();
  });
  dialog.querySelector<HTMLButtonElement>(".solo-ps-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-ps-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const set = build();
    if (!set) return;
    // Out of the document before the model changes: `apply` re-renders the page, and a stranded modal would
    // have to survive that to be a problem — it should not have to.
    shut();
    apply(set);
  };
  document.body.append(dialog);
  // A file that spells a listed sheet's two edges is shown under the list's own spelling — 21 and 29.7 rather
  // than the 21.001 that 11906 twips is worth — because that is the number a user recognises. Both spellings
  // convert back to 11906, so the page the save owes is the page that was read.
  const opening = matchedPaper();
  if (opening) fillPaper(opening.name);
  refresh();
  dialog.showModal();
  inputs.top.focus();
  inputs.top.select();
}
