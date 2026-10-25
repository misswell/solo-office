import type { MarginTwips, TablePropsPatch } from "../core/docx.js";
import { currentUnit, twipText, twipsFromUnit, unitDef } from "./doc-units.js";

/**
 * 「表格属性」对话框第一片：把表自己 `w:tblPr` 里**有受众、又还没有落笔处**的那几句接上命令面。
 *
 * 这一片摆哪几格是数出来的，不是照另一家软件的对话框猜的（授权 46 份、10 件有表、58 张 `w:tbl`，
 * `.scratch/tblpr-children-audience.py`）：
 * - 对齐 `w:jc`：14 枚 / 4 件，**全是 `center`**；左对齐是「没说」的那一档（0 枚明写 left），右对齐 0 枚。
 * - 左缩进 `w:tblInd`：58 枚 / 10 件，非缺省 28 枚（`108`×10、`288`×6、**负值** `-129`/`-440`/`-72` 各 2 枚），
 *   所以这一格**收负数**——负缩进是件里写过的东西，不是错误输入。
 * - 固定列宽 `w:tblLayout`：58 枚 / 10 件，**58 枚全是 `fixed`**，`autofit` 一枚也没写过。
 *   所以这里是一枚勾：勾上写 `<w:tblLayout w:type="fixed"/>`，不勾把那一句删掉（回到「没说」= 自动列宽），
 *   不代文件写一句它没写过的 `autofit`。
 * - 默认单元格边距 `w:tblCellMar`：58 枚 / 10 件，最常见 `上0 左108 下0 右108`（54 枚），另有 4 枚四条边全 0。
 *
 * **不摆在这一片的三格，各有各的理由**：
 * - 宽度 `w:tblW` 已有一条腿跟着网格走（#190 `setTableWidth`），再摆一格就是两个主人写同一句；
 * - 边框与底纹走同一个「边框与底纹」对话框（#145/#150/#152），那里按选区派发段落/表/格三层；
 * - 表格定位 `w:tblpPr` 只有 2 枚 / 2 件，而且要真画浮动表格才算接上，单独一片。
 *
 * 各格中文措辞（「表格属性」「左缩进」…）是这一族软件的通用说法，本机 `mui/zh_CN/*.qm` 的串池是按哈希排的
 * 平表、分不出哪句属于哪张对话框，**未取证**（见 `.scratch/HANDOFF.md` §143）。
 * 每格留空/停在「不改」= 这一句一个字节也不写。
 */

const STYLE_ID = "solo-table-props-styles";

const CSS = `
.solo-tp-dialog {
  width: 360px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-tp-dialog form { display: grid; gap: 12px; padding: 14px 16px 16px; }
.solo-tp-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-tp-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-tp-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-tp-row { display: grid; grid-template-columns: 5.5em minmax(0, 1fr); align-items: center; gap: 8px; }
.solo-tp-row > span { color: #4b5563; font-size: 12px; }
.solo-tp-margin { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 6px; }
.solo-tp-margin label { display: grid; gap: 2px; font-size: 11px; color: #6b7280; }
.solo-tp-dialog input, .solo-tp-dialog select {
  min-width: 0; padding: 4px 6px; font: 13px/1.4 inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-tp-check { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #4b5563; }
.solo-tp-hint { margin: 0; color: #6b7280; font-size: 11px; min-height: 14px; }
.solo-tp-error { margin: 0; min-height: 18px; color: #b42318; font-size: 12px; }
.solo-tp-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-tp-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-tp-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
.solo-tp-apply[disabled] { opacity: .5; cursor: default; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/** 打开时摆出来的读数：这一张表自己说过的那几句（没有 = 没说，格子停在「不改」）。 */
export type TablePropsState = {
  align?: string;
  indentTwips?: number;
  layout?: string;
  cellMargin?: MarginTwips;
  /** 选中的几张表说法不一致：格子摆第一张的，不一致的那几格摆「不改」。 */
  mixed?: boolean;
  /** 哪一句在所选几张表之间不一致——那一格不预填，免得「原样确定」把别的表改成第一张的说法。 */
  differ?: { align?: boolean; indent?: boolean; layout?: boolean; margin?: Partial<Record<keyof MarginTwips, boolean>> };
};

/** 对齐下拉里「所选几张表不一致」那一档：它与「不改」不同——它**不是**用户选的 left/center/right。 */
const MIXED = "__mixed";

const ALIGNS: [string, string][] = [
  ["", "不改"],
  ["left", "左对齐"],
  ["center", "居中"],
  ["right", "右对齐"],
];

const MARGINS: [keyof MarginTwips, string][] = [
  ["top", "上"],
  ["left", "左"],
  ["bottom", "下"],
  ["right", "右"],
];

/**
 * 一个长度格子：留空 = 不改；`allowNegative` 那一格收负数（`w:tblInd` 在授权件里有 −129/−440/−72 三档）。
 * 返回 `undefined` 表示留空，`null` 表示填坏了。
 */
function readTwips(input: HTMLInputElement, allowNegative: boolean): number | undefined | null {
  const t = input.value.trim();
  if (!t) return undefined;
  const v = Number(t);
  if (!Number.isFinite(v)) return null;
  const tw = twipsFromUnit(v, currentUnit());
  return allowNegative || tw >= 0 ? tw : null;
}

export function showTablePropsDialog(state: TablePropsState, apply: (patch: TablePropsPatch) => void): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-tp-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const unit = unitDef(currentUnit()).name;
  const dialog = document.createElement("dialog");
  dialog.className = "solo-tp-dialog";
  dialog.setAttribute("aria-label", "表格属性");

  const alignOptions = ALIGNS.map(([v, t]) => `<option value="${v}">${t}</option>`).join("");
  const marginCells = MARGINS.map(
    ([side, label]) =>
      `<label>${label}<input name="tp-m-${side}" inputmode="decimal" autocomplete="off" value="" placeholder="不改"></label>`,
  ).join("");
  dialog.innerHTML =
    `<form><header><h2>表格属性</h2><button type="button" class="solo-tp-close" aria-label="关闭">×</button></header>
    <label class="solo-tp-row"><span>对齐</span><select name="tp-align">${alignOptions}</select></label>
    <label class="solo-tp-row"><span>左缩进（${unit}）</span><input name="tp-indent" inputmode="decimal" autocomplete="off" value="" placeholder="不改"></label>
    <label class="solo-tp-check"><input type="checkbox" name="tp-fixed"> 固定列宽（按每列声明的宽度排，不按内容伸缩）</label>
    <div><span class="solo-tp-hint">默认单元格边距（${unit}）</span><div class="solo-tp-margin">${marginCells}</div></div>
    <p class="solo-tp-hint" data-hint="tp-mixed"></p>
    <p class="solo-tp-error" role="status"></p>
    <footer><button type="button" class="solo-tp-cancel">取消</button><button type="submit" class="solo-tp-apply">确定</button></footer></form>`;

  const box = (name: string): HTMLInputElement => dialog.querySelector<HTMLInputElement>(`[name="${name}"]`)!;
  const pick = dialog.querySelector<HTMLSelectElement>('[name="tp-align"]')!;
  const fixed = box("tp-fixed");
  const error = dialog.querySelector<HTMLElement>(".solo-tp-error")!;
  const hint = dialog.querySelector<HTMLElement>('[data-hint="tp-mixed"]')!;

  // 摆出件里那一句：认得出的一档直接选中，认不出来（`start`/`distribute`…）就插成一档额外选项，
  // 而不是让格子停在「不改」——那样读者会以为文件没说过，原样确定也会被当成「点了左对齐」。
  // 所选几张表在这一句上不一致时，摆的是单独那一档 `MIXED`：它按「不改」算（`build` 跳过），
  // 但字面上说得清「这里不是没人说过，是几张表各说各的」。
  if (state.differ?.align) {
    pick.add(new Option("多张表不一致（不改）", MIXED));
    pick.value = MIXED;
  } else if (state.align) {
    if (!ALIGNS.some(([v]) => v === state.align)) pick.add(new Option(`件里：${state.align}`, state.align));
    pick.value = state.align;
  }
  if (state.indentTwips !== undefined && !state.differ?.indent) box("tp-indent").value = twipText(state.indentTwips, currentUnit());
  // 固定列宽那枚勾在几张表之间不一致时摆「混合」态（原生横线），用户一碰它就落到确定的勾/不勾。
  fixed.checked = !state.differ?.layout && state.layout === "fixed";
  fixed.indeterminate = state.differ?.layout === true;
  for (const [side] of MARGINS) {
    const tw = state.differ?.margin?.[side] ? undefined : state.cellMargin?.[side];
    if (tw !== undefined) box(`tp-m-${side}`).value = twipText(tw, currentUnit());
  }
  hint.textContent = state.mixed
    ? "所选的几张表说法不一致：不一致的那几格是「不改」，填了/勾了的格子写给每一张。"
    : "";

  const build = (): TablePropsPatch | undefined => {
    const patch: TablePropsPatch = {};
    if (pick.value !== MIXED && pick.value !== (state.align ?? "")) patch.align = pick.value as TablePropsPatch["align"];
    const indent = readTwips(box("tp-indent"), true);
    if (indent === null) {
      error.textContent = "左缩进要填数字（可以是负数），留空表示不改。";
      return undefined;
    }
    if (indent !== undefined && indent !== state.indentTwips) patch.indentTwips = indent;
    const wasFixed = state.layout === "fixed";
    // 混合态还没被碰过（`indeterminate` 只在用户点一下这枚勾时被浏览器清掉）：这一句不改。
    // 点过之后，几张表混着说时勾与不勾都是用户的意思——得写，哪怕落到的值与第一张相同。
    if (!fixed.indeterminate) {
      const wanted = fixed.checked ? "fixed" : "";
      if (state.differ?.layout || fixed.checked !== wasFixed) patch.layout = wanted as TablePropsPatch["layout"];
    }
    const margins: MarginTwips = {};
    let touched = false;
    for (const [side, label] of MARGINS) {
      const tw = readTwips(box(`tp-m-${side}`), false);
      if (tw === null) {
        error.textContent = `「${label}」边距要填 0 或更大的数字，留空表示不改。`;
        return undefined;
      }
      if (tw === undefined || tw === state.cellMargin?.[side]) continue;
      margins[side] = tw;
      touched = true;
    }
    if (touched) patch.cellMargin = margins;
    error.textContent = "";
    return patch;
  };
  dialog.addEventListener("input", () => {
    if (build()) error.textContent = "";
  });
  dialog.addEventListener("change", () => {
    if (build()) error.textContent = "";
  });
  dialog.addEventListener("keydown", (e) => e.stopPropagation());

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
  dialog.addEventListener("close", () => {
    if (!dialog.isConnected) return;
    dialog.remove();
    restore();
  });
  dialog.querySelector<HTMLButtonElement>(".solo-tp-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-tp-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const patch = build();
    if (!patch) return;
    shut();
    apply(patch);
  };
  document.body.append(dialog);
  dialog.showModal();
  pick.focus();
}
