import type { RevisionEntry } from "../core/docx.js";
import type { DocEditor } from "./doc-edit.js";

/**
 * The 审阅窗格 for a Word document: every change still waiting, listed in document order, each with its
 * author, date, kind and the characters it covers, and each decidable on its own or together with the
 * rest of what that author signed.
 *
 * Mounted by `DocEditor` rather than by the shell for the same reason the comment panel is — the editor
 * owns the page and replaces `host.innerHTML` on every render, so a docked pane has to be re-docked
 * there too, and its stylesheet has to be injected because `index.html` is not ours to touch.
 *
 * Docked on the left, where Word puts its Reviewing Pane, so it never fights the comment panel for the
 * right-hand column. Class names are `solo-rev-*`; the page's own change marks already use `solo-rev-hit`
 * (styled in `doc-edit.ts`), and no class here collides with it.
 */

const STYLE_ID = "solo-doc-review-styles";

const CSS = `
.solo-rev-panel {
  position: fixed; top: 96px; left: 24px; z-index: 30;
  width: 320px; max-height: calc(100vh - 176px);
  display: flex; flex-direction: column;
  font: 13px/1.5 -apple-system, "Helvetica Neue", sans-serif;
  background: #f7f9ff; color: #2b2b2b;
  border: 1px solid #cdd6ea; border-radius: 10px;
  box-shadow: 0 10px 28px rgba(0, 0, 0, .18);
}
.solo-rev-panel[hidden] { display: none; }
.solo-rev-head {
  display: flex; align-items: center; gap: 8px; flex-wrap: wrap;
  padding: 8px 10px; border-bottom: 1px solid #dde3f0;
  font-weight: 600; color: #33406b;
}
.solo-rev-count { font-weight: 400; color: #7284ac; }
.solo-rev-spacer { flex: 1 1 100%; }
.solo-rev-note { flex: 1 1 100%; font-weight: 400; font-size: 12px; color: #8a5a00; }
.solo-rev-panel button {
  font: inherit; font-size: 12px; padding: 2px 8px; border-radius: 6px;
  border: 1px solid #cdd6ea; background: #fff; color: #33406b; cursor: pointer;
}
.solo-rev-panel button:disabled { color: #a8b2c8; cursor: default; }
.solo-rev-authors { padding: 6px 8px; border-bottom: 1px solid #dde3f0; display: flex; flex-direction: column; gap: 4px; }
.solo-rev-author { display: flex; align-items: center; gap: 6px; }
.solo-rev-author .solo-rev-name { font-weight: 600; color: #33406b; }
.solo-rev-author .solo-rev-on { background: #33406b; color: #fff; }
.solo-rev-list { overflow: auto; padding: 8px; display: flex; flex-direction: column; gap: 8px; }
.solo-rev-item { border: 1px solid #dde3f0; border-radius: 8px; padding: 8px; background: #fff; cursor: pointer; }
.solo-rev-item.solo-rev-active { border-color: #3d6fd6; box-shadow: 0 0 0 2px rgba(61, 111, 214, .18); }
.solo-rev-meta { display: flex; align-items: baseline; gap: 6px; font-size: 12px; color: #5a688f; }
.solo-rev-kind { font-weight: 600; color: #33406b; }
.solo-rev-nr { color: #92a0c0; }
.solo-rev-text {
  margin-top: 6px; max-height: 88px; overflow: auto;
  white-space: pre-wrap; word-break: break-word;
}
.solo-rev-text.solo-rev-removing { text-decoration: line-through; color: #b00020; }
.solo-rev-text.solo-rev-adding { text-decoration: underline; color: #1a6b1f; }
.solo-rev-text.solo-rev-none { color: #8b97b6; }
.solo-rev-actions { margin-top: 6px; display: flex; gap: 6px; }
.solo-rev-empty { padding: 10px; color: #7284ac; text-align: center; }
@media print { .solo-rev-panel { display: none !important; } }
`;

/** The one stylesheet the pane needs, added once per document. */
export function injectReviewStyles(): void {
  if (document.getElementById(STYLE_ID)) return;
  const style = document.createElement("style");
  style.id = STYLE_ID;
  style.textContent = CSS;
  document.head.append(style);
}

/**
 * What the pane calls each kind of pending change. These are our words for the OOXML tag the row came from:
 * WPS' own list is not evidenced in this corpus, where its numeric `WdRevisionType` values are concerned.
 */
const KIND_LABEL: Record<RevisionEntry["type"], string> = {
  insert: "插入文字",
  delete: "删除文字",
  "run-format": "文字格式",
  "paragraph-format": "段落格式",
  "insert-mark": "段落标记：新增",
  "delete-mark": "段落标记：删除",
  "move-from": "移出的文字",
  "move-to": "移入的文字",
  "row-insert": "插入行",
  "row-delete": "删除行",
  "table-format": "表格属性",
  "row-format": "行属性",
  "cell-format": "单元格属性",
  "cell-merge": "合并单元格",
  "cell-insert": "插入单元格",
  "cell-delete": "删除单元格",
};

/** The kinds whose listed text is what goes away if the change stands, so the pane strikes it. */
const REMOVING: ReadonlySet<RevisionEntry["type"]> = new Set<RevisionEntry["type"]>([
  "delete",
  "delete-mark",
  "move-from",
  "row-delete",
  "cell-delete",
]);

/** And the kinds whose text is what the change brings. */
const ADDING: ReadonlySet<RevisionEntry["type"]> = new Set<RevisionEntry["type"]>([
  "insert",
  "insert-mark",
  "move-to",
  "row-insert",
  "cell-insert",
]);

export class ReviewPanel {
  private readonly root: HTMLElement;
  private readonly list: HTMLElement;
  private readonly count: HTMLElement;
  private readonly authors: HTMLElement;
  private readonly note: HTMLElement;
  private open = false;
  /**
   * The change last revealed or stepped to, by its `w:id` rather than by row number. Row numbers are
   * positions in `revisionEntries()`, so they move every time a change is decided: keying the highlight to
   * a number would leave the pane naming some other reader-in-front's change as the one just acted on.
   */
  private activeId?: number;
  /** 只显示某一作者的修订: set when the reader names an author, cleared when that author is decided. */
  private shownAuthor?: string;

  constructor(
    private readonly host: HTMLElement,
    private readonly editor: DocEditor,
  ) {
    injectReviewStyles();
    this.root = document.createElement("aside");
    this.root.className = "solo-rev-panel";
    this.root.hidden = true;

    const head = document.createElement("div");
    head.className = "solo-rev-head";
    const title = document.createElement("span");
    title.textContent = "审阅窗格";
    this.count = document.createElement("span");
    this.count.className = "solo-rev-count";
    const prev = document.createElement("button");
    prev.type = "button";
    prev.textContent = "‹";
    prev.title = "上一条修订";
    prev.addEventListener("click", () => this.step(-1));
    const next = document.createElement("button");
    next.type = "button";
    next.textContent = "›";
    next.title = "下一条修订";
    next.addEventListener("click", () => this.step(1));
    const all = document.createElement("button");
    all.type = "button";
    all.textContent = "全部接受";
    all.addEventListener("click", () => this.decideAll("accept"));
    const none = document.createElement("button");
    none.type = "button";
    none.textContent = "全部拒绝";
    none.addEventListener("click", () => this.decideAll("reject"));
    const close = document.createElement("button");
    close.type = "button";
    close.textContent = "关闭";
    close.addEventListener("click", () => this.setOpen(false));
    this.note = document.createElement("div");
    this.note.className = "solo-rev-note";
    head.append(title, this.count, prev, next, all, none, close, this.note);

    this.authors = document.createElement("div");
    this.authors.className = "solo-rev-authors";
    this.authors.hidden = true;

    this.list = document.createElement("div");
    this.list.className = "solo-rev-list";

    this.root.append(head, this.authors, this.list);
    host.append(this.root);

    // The caret leads: a change the reader is sitting in gets named in the list.
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
    this.refresh();
  }

  isOpen(): boolean {
    return this.open;
  }

  /** Called by the editor when a change in the page is revealed, so the list names the same change. */
  revealFromPage(id: number): void {
    this.activeId = id;
    if (!this.open) return;
    this.markActive();
    // The pane follows the page: a change jumped to from the document is brought back into view.
    this.list.querySelector<HTMLElement>(`.solo-rev-item[data-rev-id="${id}"]`)?.scrollIntoView({ block: "nearest" });
  }

  private step(dir: 1 | -1): void {
    const index = this.editor.stepRevision(dir);
    if (index === undefined) {
      this.say("没有待处理的修订");
      return;
    }
    // `stepRevision` answers the row it landed on; the pane remembers the change itself, which survives
    // the renumbering that any later decision causes.
    this.activeId = this.editor.revisionEntries()[index - 1]?.id;
    this.refresh();
  }

  private decideAll(outcome: "accept" | "reject"): void {
    const waiting = this.editor.revisionEntries().length;
    if (!waiting) {
      this.say("没有待处理的修订");
      return;
    }
    const done = outcome === "accept" ? this.editor.acceptRevisions() : this.editor.rejectRevisions();
    // The count the model answers is decided changes, not listed rows: one `w:id` can cover several.
    this.say(done === 0 ? "这些修订现在判不了，文档未改动" : `${outcome === "accept" ? "已接受" : "已拒绝"} ${done} 处`);
    this.refresh();
  }

  /**
   * Rebuilds the list from the model. Rows are numbered by `revisionEntries()` even while filtered to one
   * author, because that number is what 接受/拒绝/定位 take.
   */
  refresh(): void {
    const all = this.editor.revisionEntries();
    this.count.textContent = all.length ? `(共 ${all.length} 处)` : "";
    if (!this.open) return;
    if (this.shownAuthor !== undefined && !all.some((e) => e.author === this.shownAuthor)) this.shownAuthor = undefined;
    this.buildAuthors(all);
    const shown = this.shownAuthor === undefined ? all : all.filter((e) => e.author === this.shownAuthor);
    const rows = shown.map((entry) => this.item(entry));
    if (!rows.length) {
      const empty = document.createElement("div");
      empty.className = "solo-rev-empty";
      empty.textContent = all.length ? `${this.shownAuthor ?? ""}：已无待处理的修订` : "无待处理的修订";
      rows.push(empty);
    }
    this.list.replaceChildren(...rows);
    this.markActive();
  }

  private buildAuthors(all: RevisionEntry[]): void {
    // One row per author, in the order they first appear. WPS' 审阅窗格 decides by author as well as by row,
    // and 只显示某一作者的修订 is what this bar's 只看 is for.
    const order: string[] = [];
    const counts = new Map<string, number>();
    for (const e of all) {
      if (!counts.has(e.author)) order.push(e.author);
      counts.set(e.author, (counts.get(e.author) ?? 0) + 1);
    }
    if (order.length === 0) {
      this.authors.hidden = true;
      this.authors.replaceChildren();
      return;
    }
    this.authors.hidden = false;
    this.authors.replaceChildren(
      ...order.map((author) => {
        const total = counts.get(author) ?? 0;
        const row = document.createElement("div");
        row.className = "solo-rev-author";
        const name = document.createElement("span");
        name.className = "solo-rev-name";
        name.textContent = author;
        const nr = document.createElement("span");
        nr.className = "solo-rev-nr";
        nr.textContent = `${total} 处`;
        const accept = document.createElement("button");
        accept.type = "button";
        accept.textContent = "接受";
        accept.addEventListener("click", () => this.byAuthor("accept", author, total));
        const reject = document.createElement("button");
        reject.type = "button";
        reject.textContent = "拒绝";
        reject.addEventListener("click", () => this.byAuthor("reject", author, total));
        const only = document.createElement("button");
        only.type = "button";
        only.textContent = "只看";
        only.title = `只显示 ${author} 的修订`;
        only.classList.toggle("solo-rev-on", this.shownAuthor === author);
        only.addEventListener("click", () => {
          this.shownAuthor = this.shownAuthor === author ? undefined : author;
          this.refresh();
        });
        row.append(name, nr, only, accept, reject);
        return row;
      }),
    );
  }

  private byAuthor(outcome: "accept" | "reject", author: string, total: number): void {
    // 按作者判定 is one undo step for the whole author, and answers the rows that actually landed.
    const done =
      outcome === "accept" ? this.editor.acceptRevisionsByAuthor(author) : this.editor.rejectRevisionsByAuthor(author);
    const verb = outcome === "accept" ? "已接受" : "已拒绝";
    if (done === 0) this.say(`${author}：这些修订现在判不了，文档未改动`);
    else if (done < total) this.say(`${author}：${verb} ${done}/${total} 处，其余现在判不了`);
    else this.say(`${author}：${verb}全部 ${total} 处`);
    this.refresh();
  }

  private item(entry: RevisionEntry): HTMLElement {
    const box = document.createElement("div");
    box.className = "solo-rev-item";
    box.dataset.revIndex = String(entry.index);
    box.dataset.revId = String(entry.id);

    const meta = document.createElement("div");
    meta.className = "solo-rev-meta";
    const kind = document.createElement("span");
    kind.className = "solo-rev-kind";
    kind.textContent = KIND_LABEL[entry.type];
    const nr = document.createElement("span");
    nr.className = "solo-rev-nr";
    nr.textContent = `#${entry.index}`;
    const author = document.createElement("span");
    author.textContent = entry.author;
    const date = document.createElement("span");
    date.textContent = entry.date ? entry.date.slice(0, 10) : "";
    meta.append(kind, nr, author, date);
    box.append(meta);

    const text = document.createElement("div");
    text.className = "solo-rev-text";
    if (REMOVING.has(entry.type)) text.classList.add("solo-rev-removing");
    else if (ADDING.has(entry.type)) text.classList.add("solo-rev-adding");
    // A property change lists the part's current words rather than a range of them, and a break has none.
    if (entry.text) text.textContent = entry.text;
    else {
      text.classList.add("solo-rev-none");
      text.textContent = entry.type === "delete-mark" ? "（删掉了段落标记）" : "（无文字内容）";
    }
    box.append(text);

    const actions = document.createElement("div");
    actions.className = "solo-rev-actions";
    const locate = document.createElement("button");
    locate.type = "button";
    locate.textContent = "定位";
    locate.addEventListener("click", () => this.reveal(entry));
    const accept = document.createElement("button");
    accept.type = "button";
    accept.textContent = "接受";
    accept.addEventListener("click", () =>
      this.at("accept", entry, () => this.editor.acceptRevisionAt(entry.index)),
    );
    const reject = document.createElement("button");
    reject.type = "button";
    reject.textContent = "拒绝";
    reject.addEventListener("click", () =>
      this.at("reject", entry, () => this.editor.rejectRevisionAt(entry.index)),
    );
    actions.append(locate, accept, reject);
    box.append(actions);

    // Clicking the row itself jumps, the way a row of WPS' 审阅窗格 does.
    box.addEventListener("click", (e) => {
      if ((e.target as HTMLElement).closest("button")) return;
      this.reveal(entry);
    });
    return box;
  }

  private reveal(entry: RevisionEntry): void {
    // The change may sit where the page draws no mark for it — a header or footer band, a part the
    // renderer does not stamp — in which case the model still lists it but there is nothing to jump to.
    // On success the editor names this change through `revealFromPage`, which marks and scrolls the row.
    if (!this.editor.revealRevision(entry.index)) this.say(`第 ${entry.index} 处在正文里没有可定位的标记`);
  }

  private at(outcome: "accept" | "reject", entry: RevisionEntry, run: () => boolean): void {
    if (!run()) {
      this.say(`第 ${entry.index} 处现在判不了，文档未改动`);
      return;
    }
    this.say(`第 ${entry.index} 处${outcome === "accept" ? "已接受" : "已拒绝"}`);
    // Deciding re-renders the page, which re-mounts and re-lists us.
    this.refresh();
  }

  /** Which change the caret sits in, so the list names it — by `w:id`, since row numbers move when one is decided. */
  private followCaret(): void {
    if (!this.open) return;
    const index = this.editor.revisionIndexAtCaret();
    if (index === undefined) return;
    this.activeId = this.editor.revisionEntries()[index - 1]?.id;
    this.markActive();
  }

  private markActive(): void {
    for (const el of this.list.querySelectorAll<HTMLElement>(".solo-rev-item")) {
      el.classList.toggle("solo-rev-active", el.dataset.revId === String(this.activeId ?? -1));
    }
  }

  private say(message: string): void {
    this.note.textContent = message;
  }
}
