import { renderInline } from "../core/render.js";
import type { CommentEntry, DocEditor } from "./doc-edit.js";

/**
 * The comment side panel for a Word document.
 *
 * It is mounted by `DocEditor` rather than by the shell, because the editor owns the page
 * it comments on and re-renders that page on every edit: `host.innerHTML` is replaced with
 * each render, so anything docked beside it has to be re-docked there too. Styles come from
 * an injected `<style>` for the same reason — the editor cannot add a stylesheet link.
 *
 * Class names are prefixed `solo-cmt-` throughout: `.comment-panel` and `.cmt-*` already
 * belong to the spreadsheet's cell notes and the inline anchored-range highlight.
 */

const STYLE_ID = "solo-doc-comment-styles";

/**
 * The inline marks are drawn by the renderer (`render.ts`) but styled here, because the only
 * stylesheet the editor may add to lives in index.html and that file is not ours to touch.
 * Print drops the markup entirely: a printed page reads as the document, the way Word's own
 * "Print All Text" default does.
 */
const CSS = `
.page .cmt-in { background: rgba(255, 230, 150, .45); border-radius: 2px; }
.page .cmt-anchor { display: inline; }
.page .cmt-ref {
  display: inline-block; min-width: 13px; margin: 0 2px; padding: 0 4px; border-radius: 7px;
  background: #ffd34d; color: #4a3c00; font-size: 10px; line-height: 14px; text-align: center;
  vertical-align: 2px; user-select: none; cursor: default;
}
.solo-cmt-anchor-hit { background: rgba(255, 214, 102, .55); }
.page .cmt-anchor.cmt-done { opacity: .55; }
.page .cmt-ref.cmt-done { background: #cfcfcf; color: #555; }
.solo-cmt-panel {
  position: fixed; top: 96px; right: 24px; z-index: 30;
  width: 296px; max-height: calc(100vh - 176px);
  display: flex; flex-direction: column;
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif;
  background: #fffdf7; color: #2b2b2b;
  border: 1px solid #e3d9bf; border-radius: 10px;
  box-shadow: 0 10px 28px rgba(0, 0, 0, .18);
}
.solo-cmt-panel[hidden] { display: none; }
.solo-cmt-head {
  display: flex; align-items: center; gap: 8px;
  padding: 8px 10px; border-bottom: 1px solid #ece2c8;
  font-weight: 600; color: #6b5a2a;
}
.solo-cmt-count { font-weight: 400; color: #98875c; }
.solo-cmt-spacer { flex: 1; }
.solo-cmt-head button {
  font: inherit; font-size: 12px; padding: 2px 8px; border-radius: 6px;
  border: 1px solid #ded2b2; background: #fff; color: #5a4d24; cursor: pointer;
}
.solo-cmt-list { overflow: auto; padding: 8px; display: flex; flex-direction: column; gap: 8px; }
.solo-cmt-item { border: 1px solid #ece2c8; border-radius: 8px; padding: 8px; background: #fff; }
.solo-cmt-item.solo-cmt-active { border-color: #d8a800; box-shadow: 0 0 0 2px rgba(216, 168, 0, .18); }
/* WPS and Word keep a resolved thread in the list and grey it out; deleting it would erase
   the discussion the surrounding comments still refer to. */
.solo-cmt-item.solo-cmt-done { background: #f7f6f3; color: #8c8c8c; border-style: dashed; }
.solo-cmt-item.solo-cmt-done .solo-cmt-text, .solo-cmt-item.solo-cmt-done .solo-cmt-edit { text-decoration: line-through; }
.solo-cmt-item.solo-cmt-reply { margin-left: 18px; border-left: 3px solid #efe4c4; }
.solo-cmt-state { color: #8c8c8c; font-size: 12px; }
.solo-cmt-head button[aria-pressed="true"] { background: #6b5a2a; color: #fff; }
.solo-cmt-meta { display: flex; align-items: center; gap: 6px; color: #7a6a3f; font-size: 12px; }
.solo-cmt-avatar {
  width: 18px; height: 18px; border-radius: 50%; background: #efe4c4; color: #6b5a2a;
  display: grid; place-items: center; font-size: 11px;
}
.solo-cmt-author { font-weight: 600; }
.solo-cmt-nr { color: #a29370; }
.solo-cmt-text { margin-top: 6px; white-space: pre-wrap; word-break: break-word; }
.solo-cmt-edit {
  margin-top: 6px; width: 100%; box-sizing: border-box; min-height: 46px; max-height: 190px; overflow: auto;
  font: inherit; padding: 5px 6px; border-radius: 6px; border: 1px solid #ded2b2; background: #fffdf6; color: #2b2b2b;
  white-space: pre-wrap; word-break: break-word; outline: none; cursor: text;
}
.solo-cmt-edit:focus { border-color: #d8a800; box-shadow: 0 0 0 2px rgba(216, 168, 0, .18); }
/* Each paragraph is its own block in the field; an empty one still has to take a line. */
.solo-cmt-text > div, .solo-cmt-edit > div { min-height: 1.4em; }
.solo-cmt-locked { margin-top: 6px; color: #98875c; font-size: 12px; }
.solo-cmt-actions { margin-top: 6px; display: flex; gap: 6px; }
.solo-cmt-actions button {
  font: inherit; font-size: 12px; padding: 2px 8px; border-radius: 6px;
  border: 1px solid #ded2b2; background: #fff; color: #5a4d24; cursor: pointer;
}
.solo-cmt-empty { padding: 10px; color: #98875c; text-align: center; }
@media print {
  .solo-cmt-panel { display: none !important; }
  .page .cmt-in, .solo-cmt-anchor-hit { background: none; }
  .page .cmt-ref { display: none; }
  /* Word's own "打印批注" is off by default and this is the only way to turn it on, so the
     opt-in has to reach the print stylesheet without touching the shell's print path. */
  html.solo-print-comments .page .cmt-ref {
    display: inline-block; background: #ffd34d; color: #4a3c00;
    -webkit-print-color-adjust: exact; print-color-adjust: exact;
  }
  html.solo-print-comments .page .cmt-in {
    background: rgba(255, 230, 150, .45);
    -webkit-print-color-adjust: exact; print-color-adjust: exact;
  }
}
`;

/** The one stylesheet the panel needs, added once per document. */
export function injectCommentStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

export class CommentPanel {
  private readonly root: HTMLElement;
  private readonly list: HTMLElement;
  private readonly count: HTMLElement;
  private dismissed = false;
  private active: number | null = null;

  constructor(
    private readonly host: HTMLElement,
    private readonly editor: DocEditor,
  ) {
    injectCommentStyles();
    this.root = document.createElement("aside");
    this.root.className = "solo-cmt-panel";
    this.root.hidden = true;
    this.count = document.createElement("span");
    this.count.className = "solo-cmt-count";
    this.list = document.createElement("div");
    this.list.className = "solo-cmt-list";

    const head = document.createElement("div");
    head.className = "solo-cmt-head";
    const title = document.createElement("span");
    title.textContent = "批注";
    const newButton = document.createElement("button");
    newButton.type = "button";
    newButton.textContent = "新建";
    newButton.addEventListener("click", () => this.editor.newCommentFromPanel());
    // WPS names these "&Previous Comment" / "&Next Comment" and "DeleteAllComments".
    const prev = document.createElement("button");
    prev.type = "button";
    prev.textContent = "‹";
    prev.title = "上一条批注";
    prev.addEventListener("click", () => this.editor.stepComment(-1));
    const next = document.createElement("button");
    next.type = "button";
    next.textContent = "›";
    next.title = "下一条批注";
    next.addEventListener("click", () => this.editor.stepComment(1));
    const all = document.createElement("button");
    all.type = "button";
    all.textContent = "全部删除";
    all.addEventListener("click", () => {
      this.editor.deleteAllComments();
      this.refresh();
    });
    const print = document.createElement("button");
    print.type = "button";
    print.textContent = "打印含批注";
    print.setAttribute("aria-pressed", "false");
    print.addEventListener("click", () => {
      const on = print.getAttribute("aria-pressed") !== "true";
      print.setAttribute("aria-pressed", on ? "true" : "false");
      this.editor.setPrintComments(on);
    });
    const close = document.createElement("button");
    close.type = "button";
    close.textContent = "关闭";
    close.addEventListener("click", () => this.setOpen(false));
    const spacer = document.createElement("span");
    spacer.className = "solo-cmt-spacer";
    head.append(title, this.count, spacer, prev, next, newButton, all, print, close);
    this.root.append(head, this.list);
    host.append(this.root);

    // The caret leads the selection: a comment whose text the reader is in gets named there.
    document.addEventListener("selectionchange", () => this.followCaret());
  }

  /** A re-render replaced the page and with it the docked panel; put ours back. */
  mount(): void {
    if (this.root.parentElement !== this.host) this.host.append(this.root);
    this.refresh();
  }

  setOpen(open: boolean): void {
    this.dismissed = !open;
    this.refresh();
    this.editor.paneStateChanged();
  }

  isOpen(): boolean {
    return !this.root.hidden;
  }

  /**
   * Rebuilds the list unless the reader is typing in it: a comment box that is being written
   * would lose its caret and half-typed text to any refresh from the model.
   */
  refresh(): void {
    const entries = this.editor.listComments();
    // A document that comes with comments shows them; an empty one waits for the first.
    if (entries.length === 0) this.dismissed = false;
    this.root.hidden = entries.length === 0 || this.dismissed;
    this.count.textContent = entries.length ? `(${entries.length})` : "";
    if (this.list.contains(document.activeElement)) {
      this.markActive();
      return;
    }
    // Threads read as threads: a reply sits under the comment it answers, indented. A reply
    // whose parent was deleted in this session re-attached in the model, so it is a root now.
    const known = new Set(entries.map((e) => e.id));
    const roots = entries.filter((e) => e.parentId === undefined || !known.has(e.parentId));
    const rows = roots.flatMap((root) => [
      this.item(root, false),
      ...entries.filter((e) => e.parentId === root.id).map((child) => this.item(child, true)),
    ]);
    this.list.replaceChildren(...rows);
    if (entries.length === 0) {
      const empty = document.createElement("div");
      empty.className = "solo-cmt-empty";
      empty.textContent = "暂无批注";
      this.list.append(empty);
    }
    this.markActive();
  }

  private item(entry: CommentEntry, reply: boolean): HTMLElement {
    const box = document.createElement("div");
    box.className = `solo-cmt-item${entry.resolved ? " solo-cmt-done" : ""}${reply ? " solo-cmt-reply" : ""}`;
    box.dataset.cid = String(entry.id);

    const meta = document.createElement("div");
    meta.className = "solo-cmt-meta";
    const avatar = document.createElement("span");
    avatar.className = "solo-cmt-avatar";
    avatar.textContent = entry.initials || entry.author.slice(0, 1) || "?";
    const author = document.createElement("span");
    author.className = "solo-cmt-author";
    author.textContent = entry.author;
    const date = document.createElement("span");
    date.textContent = entry.date ? entry.date.slice(0, 10) : "";
    const nr = document.createElement("span");
    nr.className = "solo-cmt-nr";
    nr.textContent = `#${entry.ordinal}`;
    meta.append(avatar, author, date, nr);
    if (entry.resolved) {
      const state = document.createElement("span");
      state.className = "solo-cmt-state";
      state.textContent = "已解决";
      meta.append(state);
    }
    box.append(meta);

    if (!entry.editable) {
      const text = document.createElement("div");
      text.className = "solo-cmt-text";
      text.innerHTML = entry.runs.map((runs) => `<div>${runs.map(renderInline).join("") || "<br>"}</div>`).join("");
      box.append(text);
      const locked = document.createElement("div");
      locked.className = "solo-cmt-locked";
      locked.textContent = "此批注含表格等富内容，暂不支持编辑";
      box.append(locked);
    } else {
      // A contenteditable rather than a textarea: a comment written in bold or in colour keeps
      // that formatting across an edit, which a plain-text field would flatten away. Each
      // paragraph is its own block so the field reads back as the `w:p`s it came from.
      const field = document.createElement("div");
      field.className = "solo-cmt-edit";
      field.contentEditable = "true";
      field.spellcheck = false;
      field.innerHTML = entry.runs.map((runs) => `<div>${runs.map(renderInline).join("") || "<br>"}</div>`).join("");
      const commit = (): void => {
        if (this.editor.setCommentRuns(entry.id, this.editor.commentParagraphsOf(field))) this.refresh();
      };
      field.addEventListener("blur", commit);
      // Pasted as plain text on purpose: a rich paste would carry tables and images this model
      // cannot write back for a comment, and the edit would then silently drop what it showed.
      field.addEventListener("paste", (e) => {
        const clip = (e as ClipboardEvent).clipboardData;
        if (!clip) return;
        e.preventDefault();
        document.execCommand("insertText", false, clip.getData("text/plain"));
      });
      field.addEventListener("keydown", (e) => {
        const key = e as KeyboardEvent;
        // Committing on every modified keystroke would eat ⌘C and ⌘V; Enter is the one meant.
        if (key.key === "Enter" && (key.metaKey || key.ctrlKey)) {
          commit();
          key.preventDefault();
        }
      });
      box.append(field);
    }

    const actions = document.createElement("div");
    actions.className = "solo-cmt-actions";
    const locate = document.createElement("button");
    locate.type = "button";
    locate.textContent = entry.anchored ? "定位" : "正文中已无锚点";
    locate.disabled = !entry.anchored;
    locate.addEventListener("click", () => this.editor.revealComment(entry.id));
    const remove = document.createElement("button");
    remove.type = "button";
    remove.textContent = "删除";
    remove.addEventListener("click", () => {
      this.editor.deleteComment(entry.id);
      this.refresh();
    });
    const resolve = document.createElement("button");
    resolve.type = "button";
    resolve.textContent = entry.resolved ? "重新打开" : "解决";
    resolve.addEventListener("click", () => {
      this.editor.setCommentResolved(entry.id, !entry.resolved);
      this.refresh();
    });
    const answer = document.createElement("button");
    answer.type = "button";
    answer.textContent = "回复";
    answer.addEventListener("click", () => {
      // A reply opens as a thread under its parent, at the same anchor and by the same author.
      const id = this.editor.replyComment(entry.id);
      this.refresh();
      if (id !== undefined) this.openAt(id);
    });
    actions.append(locate, answer, resolve, remove);
    box.append(actions);
    return box;
  }

  /** Which comment the caret sits inside, so the list names it. */
  private followCaret(): void {
    if (this.root.hidden) return;
    if (this.list.contains(document.activeElement)) return;
    const ids = this.editor.commentIdsAtCaret();
    this.active = ids.length ? ids[0] : null;
    this.markActive();
  }

  private markActive(): void {
    for (const el of this.list.querySelectorAll<HTMLElement>(".solo-cmt-item")) {
      el.classList.toggle("solo-cmt-active", el.dataset.cid === String(this.active ?? -1));
    }
  }

  setActive(id: number | null): void {
    this.active = id;
    this.markActive();
  }

  /** Called by the editor when the reader clicks a reference mark. */
  openAt(id: number): void {
    this.dismissed = false;
    this.active = id;
    this.refresh();
  }
}
