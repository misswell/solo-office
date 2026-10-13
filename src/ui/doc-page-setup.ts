import type { PageSetup, Section } from "../core/docx.js";
import { cmToTwips, twipsToCm } from "../core/docx.js";

/**
 * 「页面设置」: the four body margins, the two paper-edge distances, the binding width, and which way the paper
 * is turned. The dialog counts in centimetres the way the Word family does, and hands the model twips —
 * `w:pgMar` has no other unit (ECMA-376 §17.3.2.13) — converting once here so no number is rounded twice.
 *
 * Three distances may be left **empty**, which changes nothing: they are optional in the part, and a file that
 * says nothing about its binding width stays a file that says nothing. The four body margins are always sent,
 * because the page on screen is drawn from them and a save has to write the box it showed.
 *
 * Prefilled to three decimals, because that is what brings the twip count back exactly: 851 twips is
 * 1.501 cm, and the 1.50 a two-decimal dialog would show comes back as 850.
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
.solo-ps-fields { display: grid; grid-template-columns: repeat(2, 1fr); gap: 10px 12px; }
.solo-ps-fields label { display: grid; gap: 3px; font-size: 12px; color: #4b5563; }
.solo-ps-fields .solo-ps-num { display: flex; align-items: center; gap: 6px; }
.solo-ps-fields input {
  flex: 1; min-width: 0; padding: 4px 6px; font: inherit; color: inherit;
  border: 1px solid #c9d3e6; border-radius: 6px; background: #fff;
}
.solo-ps-fields span { color: #6b7280; }
.solo-ps-fields small { color: #6b7280; font-size: 11px; }
.solo-ps-orient { display: flex; gap: 14px; align-items: center; font-size: 12px; color: #4b5563; }
.solo-ps-orient label { display: flex; gap: 5px; align-items: center; }
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

/** The centimetres a section's twips are worth, spelled tightly enough to come back as the same integer. */
const cmText = (tw: number | undefined): string => (tw === undefined ? "" : String(Number(twipsToCm(tw).toFixed(3))));

/** What one box holds: nothing typed means "not asked about", and anything unparseable is refused. */
type Reading = { value?: number; bad: boolean };

const read = (input: HTMLInputElement): Reading => {
  const text = input.value.trim();
  if (!text) return { bad: false };
  const cm = Number(text);
  // No upper bound is invented: `w:pgMar` takes any twip count, and a producer's own limit is not evidenced
  // in this corpus. What is refused is a value that is not a number at all, or a negative distance.
  if (!Number.isFinite(cm) || cm < 0) return { bad: true };
  return { value: cmToTwips(cm), bad: false };
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
  const box = (name: string, label: string, value: string, hint = ""): string =>
    `<label>${label}<span class="solo-ps-num"><input name="${name}" inputmode="decimal" autocomplete="off" value="${value}"><span>厘米</span></span>${hint}</label>`;
  dialog.innerHTML =
    `<form><header><h2>页面设置</h2><button type="button" class="solo-ps-close" aria-label="关闭">×</button></header>
    <div class="solo-ps-fields">
    ${box("ps-top", "上", cmText(current.margins.top))}
    ${box("ps-bottom", "下", cmText(current.margins.bottom))}
    ${box("ps-left", "左", cmText(current.margins.left))}
    ${box("ps-right", "右", cmText(current.margins.right))}
    ${box("ps-header", "页眉距离", cmText(current.margins.header), "<small>从纸张边缘量；留空表示不改</small>")}
    ${box("ps-footer", "页脚距离", cmText(current.margins.footer), "<small>从纸张边缘量；留空表示不改</small>")}
    ${box("ps-gutter", "装订线", cmText(current.margins.gutter), "<small>加在装订侧的页边距上；留空表示不改</small>")}
    <label class="solo-ps-orient">纸张方向
      <span><label><input type="radio" name="ps-orient" value="portrait">纵向</label>
      <label><input type="radio" name="ps-orient" value="landscape">横向</label></span>
    </label>
    </div>
    <p class="solo-ps-error" role="status"></p>
    <footer><button type="button" class="solo-ps-cancel">取消</button><button type="submit" class="solo-ps-apply">确定</button></footer></form>`;

  const field = (name: string): HTMLInputElement => dialog.querySelector(`[name="${name}"]`)!;
  const inputs = {
    top: field("ps-top"),
    bottom: field("ps-bottom"),
    left: field("ps-left"),
    right: field("ps-right"),
    header: field("ps-header"),
    footer: field("ps-footer"),
    gutter: field("ps-gutter"),
  };
  const error = dialog.querySelector<HTMLElement>(".solo-ps-error")!;
  const applyButton = dialog.querySelector<HTMLButtonElement>(".solo-ps-apply")!;
  (dialog.querySelector(`[name="ps-orient"][value="${current.pageWidthTwips > current.pageHeightTwips ? "landscape" : "portrait"}"]`) as HTMLInputElement).checked = true;

  /** The setup the boxes say, or nothing at all when a box cannot be read. */
  const build = (): PageSetup | undefined => {
    const taken = [read(inputs.top), read(inputs.bottom), read(inputs.left), read(inputs.right)];
    const optional = [read(inputs.header), read(inputs.footer), read(inputs.gutter)];
    const bad = taken.some((r) => r.bad) || optional.some((r) => r.bad);
    const missing = taken.some((r) => r.value === undefined);
    if (bad || missing) {
      error.textContent = bad ? "页边距与距离需填 0 或更大的数字。" : "上、下、左、右四个页边距都要填数字。";
      applyButton.disabled = true;
      return undefined;
    }
    error.textContent = "";
    applyButton.disabled = false;
    return {
      topTwips: taken[0]!.value!,
      bottomTwips: taken[1]!.value!,
      leftTwips: taken[2]!.value!,
      rightTwips: taken[3]!.value!,
      headerTwips: optional[0].value,
      footerTwips: optional[1].value,
      gutterTwips: optional[2].value,
      landscape: (dialog.querySelector<HTMLInputElement>('input[name="ps-orient"]:checked')?.value ?? "portrait") === "landscape",
    };
  };
  const refresh = (): void => {
    build();
  };
  /**
   * Closes and takes the node out of the document in the same step. The `close` event is a queued task, so
   * relying on it alone leaves a shut dialog lying in the DOM until the next one runs — and the next
   * 「页面设置」 opened before then would find *that* one first, filling boxes on a dialog nobody is looking at.
   */
  const shut = (): void => {
    dialog.close();
    dialog.remove();
  };
  dialog.addEventListener("input", refresh);
  dialog.addEventListener("change", refresh);
  // The editor owns the page's shortcuts; a dialog typing over them would move the caret behind it.
  dialog.addEventListener("keydown", (e) => e.stopPropagation());
  // Escape closes the dialog without going through any of the buttons, so this is the handler that path has.
  dialog.addEventListener("close", () => dialog.remove(), { once: true });
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
  refresh();
  dialog.showModal();
  inputs.top.focus();
  inputs.top.select();
}
