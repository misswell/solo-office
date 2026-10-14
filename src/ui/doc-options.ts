import { UNITS, currentUnit, setUnit, type Unit } from "./doc-units.js";

/**
 * 「选项」对话框——这一版只摆「测量单位」一组。实据是 `wpscommon.kuip` 459 行的
 * `KxWpsStatusUnitCommand`：detail「@Measurement unit. Click to open "Options" dialog.」——状态栏那一格
 * 点开的是选项对话框。真机那张对话框是**多页**的（常规/视图/编辑/打印/…），摆哪些页、页签字面、这一组
 * 排在哪一页**都未取证**；这里只做了「测量单位」一组：四种单位、单选、按「确定」生效、「取消」与 Esc
 * 退出即作废。选择只活在这个窗口里（口径见 `doc-units.ts` 头注）。
 *
 * 开窗与收尾照「页面设置」那一套（`doc-page-setup.ts`）：同一时刻只留这一类里的一张、开窗前存一份选区
 * Range、每条退出路径把光标放回去。开窗会聚焦单选钮——浏览器把**文档选区**搬进对话框（与「定位」「页面
 * 设置」同一个行为），所以取消/Esc 之后状态栏要能照旧读到开窗前的光标：「确定」那条路先收摊（放回选区）
 * 再 `setUnit`，订阅者（状态栏）在通知里现读的才是原来那位光标。
 *
 * 样式从这个文件注入；对话框挂在 `document.body` 上而不是编辑器的宿主里——`render()` 会把宿主整个换掉。
 */
const STYLE_ID = "solo-options-styles";

const CSS = `
.solo-op-dialog {
  width: 320px; border: 1px solid #cdd6ea; border-radius: 12px; padding: 0;
  box-shadow: 0 18px 48px rgba(0, 0, 0, .28);
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif; color: #2b2b2b;
}
.solo-op-dialog form { display: grid; gap: 14px; padding: 14px 16px 16px; }
.solo-op-dialog header { display: flex; align-items: center; justify-content: space-between; }
.solo-op-dialog h2 { margin: 0; font-size: 15px; font-weight: 600; }
.solo-op-dialog .solo-op-close { border: 0; background: none; font-size: 18px; cursor: pointer; color: #6b7280; }
.solo-op-dialog fieldset { border: 1px solid #d8dfeb; border-radius: 8px; padding: 8px 12px 10px; display: grid; gap: 6px; margin: 0; }
.solo-op-dialog legend { padding: 0 4px; font-size: 12px; color: #4b5563; }
.solo-op-dialog fieldset label { display: flex; gap: 6px; align-items: center; font-size: 12.5px; color: #333; }
.solo-op-dialog footer { display: flex; gap: 8px; justify-content: flex-end; }
.solo-op-dialog footer button { padding: 5px 14px; border-radius: 7px; border: 1px solid #c9d3e6; background: #f5f7fb; cursor: pointer; font: inherit; }
.solo-op-dialog .solo-op-apply { background: #2f6fed; border-color: #2f6fed; color: #fff; }
`;

function injectStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

export function openOptions(): void {
  // One dialog at a time: reopening from the status bar replaces what is on screen instead of stacking a second.
  document.querySelectorAll<HTMLDialogElement>(".solo-op-dialog").forEach((old) => {
    old.close();
    old.remove();
  });
  injectStyles();
  const dialog = document.createElement("dialog");
  dialog.className = "solo-op-dialog";
  dialog.setAttribute("aria-label", "选项");
  dialog.innerHTML =
    `<form><header><h2>选项</h2><button type="button" class="solo-op-close" aria-label="关闭">×</button></header>
    <fieldset><legend>测量单位</legend>
    ${UNITS.map((u) => `<label><input type="radio" name="op-unit" value="${u.id}">${u.name}</label>`).join("")}
    </fieldset>
    <footer><button type="button" class="solo-op-cancel">取消</button><button type="submit" class="solo-op-apply">确定</button></footer></form>`;

  const checked = (): HTMLInputElement => dialog.querySelector<HTMLInputElement>('input[name="op-unit"]:checked')!;
  // Prefilled to the unit in force: reopening shows the current choice, not the first entry.
  dialog.querySelector<HTMLInputElement>(`input[name="op-unit"][value="${currentUnit()}"]`)!.checked = true;

  /**
   * Closes and takes the node out of the document in the same step (the `close` event is a queued task, so
   * relying on it alone leaves a shut dialog in the DOM until the next one runs). 光标也在这里放回来：见文件头。
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
  // The editor owns the page's shortcuts; a dialog typing over them would move the caret behind it.
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  // Escape closes the dialog without going through any of the buttons, so this is the handler that path has.
  // `shut()` 那条路连摘除都做完了（isConnected 是 false），这里就跳过——按「确定」之后被摘掉的那份押着的
  // Range 也就不会盖回来（放回它等于把选区指回一堆断开的节点）。
  dialog.addEventListener("close", () => {
    if (!dialog.isConnected) return;
    dialog.remove();
    restore();
  });
  dialog.querySelector<HTMLButtonElement>(".solo-op-close")!.onclick = shut;
  dialog.querySelector<HTMLButtonElement>(".solo-op-cancel")!.onclick = shut;
  dialog.querySelector("form")!.onsubmit = (e) => {
    e.preventDefault();
    const picked = checked().value as Unit;
    shut();
    setUnit(picked);
  };
  document.body.append(dialog);
  dialog.showModal();
  checked().focus();
}
