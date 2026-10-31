import type { TableSortKind, TableSortOrder } from "../core/doc-table-ops.js";

/**
 * 表格「排序」对话框。取据：`CT_TableTools.kuip:90` 的 `<command idQ="Sort" singlerowmergegroup="SingleRow_Data"/>`
 * 把这一条摆在 RB_Data 的「转换成文本」之后，命令本体在 `wpscommon.kuip`：
 * `<KxLegacyTriggerCommand customTip="Sort" hotKey="S" icon="sort_kd" id="Sort" ksoCmd="Sort" rbHotKey="SO"
 * text="@Sort"/>`——它不带 `idMso`、也不带 `txt`，所以 WPS 那张对话框分几栏、每栏什么字面**未取证**
 * （`.qm` 只证词在不在：`mui/zh_CN/wps.qm` 里「排序」有、`kso.qm` 里也有；中文措辞一条也不抄进产品）。
 * 这里摆的是家已有落笔处、纸上真能排出来的四句：排序依据（哪一列）、类型（自动/数字/文字）、顺序（升/降）、
 * 有标题行（第一行不参与排）。WPS 的「笔画／拼音／日期」那几档与「选项」一栏不摆：前者画成什么样没量过，
 * 后者没有对应的落笔处。
 *
 * 对话框只交**用户选了什么**，不碰模型：列号是网格列号（跨列格之后「第几个 `w:tc`」与它是两回事），
 * 钉住的行数由编辑器按件里的 `w:tblHeader` 与这一勾合起来算。
 */
const STYLE_ID = "solo-table-sort-styles";

const CSS = `
.solo-sort-dialog {
  width: 336px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-sort-dialog form { display: grid; gap: 10px; padding: 14px 16px 16px; }
.solo-sort-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-sort-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-sort-dialog .solo-sort-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-sort-dialog label { display: grid; gap: 3px; font-size: 12px; color: #4b5563; min-width: 0; }
.solo-sort-dialog select {
  padding: 4px 6px; font: inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-sort-dialog .solo-sort-head { display: flex; align-items: center; gap: 6px; font-size: 12px; color: #374151; }
.solo-sort-dialog .solo-sort-head input { width: 13px; height: 13px; margin: 0; }
.solo-sort-hint { color: #6b7280; font-size: 11px; min-height: 14px; }
.solo-sort-error { margin: 0; min-height: 18px; color: #b42318; font-size: 12px; }
.solo-sort-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-sort-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-sort-dialog .solo-sort-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
.solo-sort-dialog .solo-sort-apply[disabled] { opacity: .5; cursor: default; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

export type TableSortColumn = { index: number; label: string };

/** 打开对话框时页面上那一张表的形状：列清单、光标那一列、顶部钉住几行。 */
export type TableSortState = {
  columns: TableSortColumn[];
  /** 光标所在格的网格列；不在表里时编辑器不给状态，给 0 是为了让下拉有个落点。 */
  column: number;
  /** 件里写死的标题行（`w:tblHeader` 连续几行），对话框那一勾的默认值。 */
  headerRows: number;
  order: TableSortOrder;
  kind: TableSortKind;
};

/** 对话框交回来的那一句。 */
export type TableSortChoice = {
  column: number;
  order: TableSortOrder;
  kind: TableSortKind;
  hasHeader: boolean;
};

const KIND_LABELS: Record<TableSortKind, string> = { auto: "自动", number: "数字", text: "文字" };
const ORDER_LABELS: Record<TableSortOrder, string> = { asc: "升序", desc: "降序" };

export function showTableSortDialog(state: TableSortState, apply: (choice: TableSortChoice) => void): void {
  document.querySelectorAll<HTMLDialogElement>(".solo-sort-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-sort-dialog";
  dialog.setAttribute("aria-label", "排序");
  const columns = state.columns.length ? state.columns : [{ index: 0, label: "第 1 列" }];
  dialog.innerHTML =
    `<form><header><h2>排序</h2><button type="button" class="solo-sort-close" aria-label="关闭">×</button></header>
    <label>排序依据<select name="sort-column">${columns
      .map((c) => `<option value="${c.index}">${c.label}</option>`)
      .join("")}</select></label>
    <label>类型<select name="sort-kind">${(Object.keys(KIND_LABELS) as TableSortKind[])
      .map((k) => `<option value="${k}">${KIND_LABELS[k]}</option>`)
      .join("")}</select></label>
    <label>顺序<select name="sort-order">${(Object.keys(ORDER_LABELS) as TableSortOrder[])
      .map((o) => `<option value="${o}">${ORDER_LABELS[o]}</option>`)
      .join("")}</select></label>
    <label class="solo-sort-head"><input type="checkbox" name="sort-header"><span>有标题行（第一行不参与排序）</span></label>
    <small class="solo-sort-hint">排序把整行一起挪走：行高、标题行、每格自己的格式都跟着这一行走。</small>
    <p class="solo-sort-error" role="status"></p>
    <footer><button type="button" class="solo-sort-cancel">取消</button><button type="submit" class="solo-sort-apply">确定</button></footer></form>`;

  const column = dialog.querySelector<HTMLSelectElement>('[name="sort-column"]')!;
  const kind = dialog.querySelector<HTMLSelectElement>('[name="sort-kind"]')!;
  const order = dialog.querySelector<HTMLSelectElement>('[name="sort-order"]')!;
  const header = dialog.querySelector<HTMLInputElement>('[name="sort-header"]')!;
  const error = dialog.querySelector<HTMLElement>(".solo-sort-error")!;
  const applyButton = dialog.querySelector<HTMLButtonElement>(".solo-sort-apply")!;

  // 预填：光标那一列、件里已钉住的标题行、升序与自动类型。下拉里认不出光标那一列时落到第一项，
  // 而不是摆一个不存在的值让 `value` 变空串。
  column.value = String(columns.some((c) => c.index === state.column) ? state.column : columns[0]!.index);
  kind.value = state.kind;
  order.value = state.order;
  header.checked = state.headerRows > 0;

  const read = (): TableSortChoice | undefined => {
    const col = Number(column.value);
    if (!Number.isInteger(col) || col < 0) {
      error.textContent = "排序依据要选一列。";
      applyButton.disabled = true;
      return undefined;
    }
    error.textContent = "";
    applyButton.disabled = false;
    return { column: col, order: order.value as TableSortOrder, kind: kind.value as TableSortKind, hasHeader: header.checked };
  };
  dialog.addEventListener("input", read);
  dialog.addEventListener("change", read);
  // The editor owns the page's shortcuts; a dialog typing over them would move the caret behind it.
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
  dialog.querySelector<HTMLButtonElement>(".solo-sort-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-sort-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const choice = read();
    if (!choice) return;
    shut();
    apply(choice);
  };
  document.body.append(dialog);
  dialog.showModal();
  column.focus();
}
