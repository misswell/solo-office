import type { DocEditor } from "./doc-edit.js";

/**
 * The 导航窗格 for a Word document: every heading the document's outline levels name, in reading order, each
 * one indented by its level and each a way to jump to the paragraph it stands for.
 *
 * What the list is made of is the file's own statement, not a guess at its names: `w:outlineLvl`
 * (§17.3.2.26) resolved through the paragraph's style chain. Authorized 语料 46 份 WPS 自带模板的读数（口径分两档，
 * 别混用）：正文层 **7,430 段**（xmllint `//w:p`；配对正则只数到 7,393，少的 37 枚是自闭合的空段）里有
 * **202 段自己写 `w:outlineLvl`**（`.scratch/authorized-families2.mjs`，xmllint 对拍同为 202），走样式链解析之后是
 * **408 段**（`.scratch/authorized-families.mjs` 的解析层口径）。旧句「785 段 / 46,417 段」是已作废的 86 份清单。
 * A paragraph whose text merely *looks* like a heading is not in this list, because nothing in the part says so.
 *
 * Mounted by `DocEditor` for the same reason the review pane is — the editor owns the page and replaces
 * `host.innerHTML` on every render, so a docked pane has to be re-docked there too, and its stylesheet is
 * injected because `index.html` is not ours to touch.
 *
 * Docked on the left, where WPS puts its 导航窗格, and the 审阅窗格 steps aside while this one is open (see the
 * `[data-ol-open]` rule), so the two can be open together without covering each other.
 */

const STYLE_ID = "solo-doc-outline-styles";

const CSS = `
.solo-ol-panel {
  position: fixed; top: 96px; left: 24px; z-index: 30;
  width: 280px; max-height: calc(100vh - 176px);
  display: flex; flex-direction: column;
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif;
  background: #f7f9ff; color: #2b2b2b;
  border: 1px solid #cdd6ea; border-radius: 10px;
  box-shadow: 0 10px 28px rgba(0, 0, 0, .18);
}
.solo-ol-panel[hidden] { display: none; }
.solo-ol-head {
  display: flex; align-items: center; gap: 8px;
  padding: 8px 10px; border-bottom: 1px solid #dde3f0;
  font-weight: 600; color: #33406b;
}
.solo-ol-count { font-weight: 400; color: #7284ac; }
.solo-ol-spacer { flex: 1 1 auto; }
.solo-ol-panel button {
  font: inherit; font-size: 12px; padding: 2px 8px; border-radius: 6px;
  border: 1px solid #cdd6ea; background: #fff; color: #33406b; cursor: pointer;
}
.solo-ol-list { overflow: auto; padding: 6px; }
.solo-ol-item {
  padding: 3px 6px; border-radius: 6px; cursor: pointer;
  white-space: nowrap; overflow: hidden; text-overflow: ellipsis;
}
.solo-ol-item:hover { background: #eaf0ff; }
.solo-ol-item.solo-ol-active { background: #fff2c4; box-shadow: inset 0 0 0 1px #d8b64a; }
.solo-ol-lv { color: #92a0c0; font-size: 11px; margin-right: 4px; }
.solo-ol-empty { padding: 10px; color: #7284ac; text-align: center; }
/* Both panes dock on the left; the outline is the wider habit, so the review pane moves beside it
   rather than under it. The attribute is on the host the panes share, and survives the page re-render. */
[data-ol-open] .solo-rev-panel { left: 320px; }
@media print { .solo-ol-panel { display: none !important; } }
`;

/** The one stylesheet the pane needs, added once per document. */
export function injectOutlineStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

export class OutlinePanel {
  private readonly root: HTMLElement;
  private readonly list: HTMLElement;
  private readonly count: HTMLElement;
  private open = false;
  /**
   * Which row the reader is on, by position in `doc.outline()` rather than by the paragraph's address. The two
   * agree because the page writes one `[data-outline]` mark per entry in that list, in the same order —
   * `renderBlocks` walks the model and `outline()` walks it the same way, table cells included.
   */
  private active = -1;

  constructor(
    private readonly host: HTMLElement,
    private readonly editor: DocEditor,
  ) {
    injectOutlineStyles();
    this.root = document.createElement("aside");
    this.root.className = "solo-ol-panel";
    this.root.hidden = true;

    const head = document.createElement("div");
    head.className = "solo-ol-head";
    const title = document.createElement("span");
    title.textContent = "导航窗格";
    this.count = document.createElement("span");
    this.count.className = "solo-ol-count";
    const close = document.createElement("button");
    close.type = "button";
    close.textContent = "关闭";
    close.addEventListener("click", () => this.setOpen(false));
    const spacer = document.createElement("span");
    spacer.className = "solo-ol-spacer";
    head.append(title, this.count, spacer, close);

    this.list = document.createElement("div");
    this.list.className = "solo-ol-list";
    this.root.append(head, this.list);
    host.append(this.root);

    // The caret leads: the heading the reader is sitting under stays named in the list.
    document.addEventListener("selectionchange", () => this.followCaret());
  }

  /** A re-render replaced the page and with it the docked pane; put ours back. */
  mount(): void {
    if (this.root.parentElement !== this.host) this.host.append(this.root);
    this.refresh();
  }

  setOpen(open: boolean): void {
    this.open = open;
    this.root.hidden = !open;
    if (open) this.host.setAttribute("data-ol-open", "");
    else this.host.removeAttribute("data-ol-open");
    this.refresh();
    this.editor.paneStateChanged();
  }

  isOpen(): boolean {
    return this.open;
  }

  /** Called by the editor when a heading is revealed from the page, so the list names the same one. */
  revealFromPage(index: number): void {
    this.active = index;
    if (!this.open) return;
    this.markActive();
    this.list.querySelector<HTMLElement>(`.solo-ol-item[data-ol-index="${index}"]`)?.scrollIntoView({ block: "nearest" });
  }

  private followCaret(): void {
    if (!this.open) return;
    const marks = this.editor.outlineMarks();
    if (!marks.length) return;
    const sel = document.getSelection();
    const node = sel?.rangeCount ? sel.getRangeAt(0).startContainer : null;
    const el = node instanceof Element ? node : (node?.parentElement ?? null);
    if (!el) return;
    // The last heading that starts at or before the caret — the one the reader is under, even when the caret
    // has moved on into the body paragraphs the heading introduces.
    let at = -1;
    for (let i = 0; i < marks.length; i++) {
      const m = marks[i]!;
      if (m === el || m.contains(el)) {
        at = i;
        break;
      }
      if (el.compareDocumentPosition(m) & Node.DOCUMENT_POSITION_FOLLOWING) break;
      at = i;
    }
    if (at < 0 || at === this.active) return;
    this.active = at;
    this.markActive();
  }

  private markActive(): void {
    for (const row of this.list.querySelectorAll<HTMLElement>(".solo-ol-item")) {
      row.classList.toggle("solo-ol-active", Number(row.dataset.olIndex) === this.active);
    }
  }

  private refresh(): void {
    const entries = this.editor.outlineEntries();
    this.count.textContent = entries.length ? `${entries.length} 个标题` : "无标题";
    if (!entries.length) {
      const none = document.createElement("div");
      none.className = "solo-ol-empty";
      // Word and WPS name a heading by its outline level; a document that states none has nothing to list,
      // and saying why is more use than an empty box.
      none.textContent = "本文没有写出大纲级别（w:outlineLvl）的段落";
      this.list.replaceChildren(none);
      return;
    }
    const rows = entries.map((e, i) => {
      const row = document.createElement("div");
      row.className = "solo-ol-item";
      row.dataset.olIndex = String(i);
      // Indented by level the way the pane in WPS does; the number is kept as the tooltip because a level-4
      // heading two indents deep is hard to tell from a level-3 one by eye.
      row.style.paddingLeft = `${6 + (e.level - 1) * 14}px`;
      row.title = `第 ${e.level} 级标题：${e.text || "（空标题）"}`;
      const lv = document.createElement("span");
      lv.className = "solo-ol-lv";
      lv.textContent = `${e.level}`;
      const text = document.createElement("span");
      text.textContent = e.text || "（空标题）";
      row.append(lv, text);
      row.addEventListener("click", () => this.editor.revealOutline(i));
      return row;
    });
    this.list.replaceChildren(...rows);
    this.markActive();
  }
}
