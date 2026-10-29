import type { Inline } from "../core/docx.js";
import { hasLine } from "../core/docx.js";
import type { Presentation, Shape, Slide } from "../core/pptx.js";
import { renderSlideHtml } from "../core/render.js";

/**
 * Editing layer over a presentation: slide rail, stage with selectable shapes,
 * drag/resize, in-place text editing and speaker notes. Shape geometry and text
 * live in the model; a save patches only the slides that were touched.
 */

interface Snap {
  slides: Slide[];
  removed: Set<string>;
  presDirty: boolean;
  addedMedia: [string, string][];
}

const HANDLE_DIRS = ["nw", "n", "ne", "e", "se", "s", "sw", "w"] as const;
type Dir = (typeof HANDLE_DIRS)[number];

/** Insert defaults, close to what a fresh shape looks like in an office suite. */
const SHAPE_FILL = "#dce6f1";
const SHAPE_LINE = "#4472c4";
export type PresetGeometry = "rect" | "roundRect" | "ellipse" | "triangle" | "line" | "diamond" | "parallelogram" | "rightArrow" | "star5" | "chevron";
const SHAPE_NAMES: Record<PresetGeometry, string> = {
  rect: "矩形",
  roundRect: "圆角矩形",
  ellipse: "椭圆",
  triangle: "三角形",
  line: "直线",
  diamond: "菱形",
  parallelogram: "平行四边形",
  rightArrow: "右箭头",
  star5: "五角星",
  chevron: "燕尾形",
};

export class DeckEditor {
  private stack: Snap[] = [];
  private future: Snap[] = [];
  private rail!: HTMLElement;
  private stage!: HTMLElement;
  private notesEl?: HTMLElement;
  private selBox?: HTMLElement;
  private selected: number | null = null;
  private drag: { mode: "move" | "resize"; dir?: Dir; sx: number; sy: number; moved: boolean; box: { x: number; y: number; width: number; height: number }; el: HTMLElement } | null = null;
  private editing: HTMLElement | null = null;
  private burst = 0;
  private fit = 1;

  constructor(private host: HTMLElement, private pres: Presentation, private selectedSlide: number, private hooks: { changed(): void; selected?(): void }) {
    this.render();
  }

  get slide(): Slide | undefined {
    return this.pres.slides[this.selectedSlide];
  }

  /** Zero-based current slide, so a show can start from where the user is editing. */
  get index(): number {
    return this.selectedSlide;
  }

  canUndo(): boolean {
    return this.stack.length > 0;
  }

  canRedo(): boolean {
    return this.future.length > 0;
  }

  // ----------------------------------------------------------------- render

  private render(): void {
    this.host.replaceChildren();
    const deck = document.createElement("div");
    deck.className = "deck";

    const rail = document.createElement("div");
    rail.className = "rail";
    const bar = document.createElement("div");
    bar.className = "rail-bar";
    const btn = (label: string, title: string, fn: () => void): HTMLButtonElement => {
      const b = document.createElement("button");
      b.textContent = label;
      b.title = title;
      b.onclick = fn;
      return b;
    };
    bar.append(
      btn("＋", "在当前页后新建一页", () => this.addSlide()),
      btn("🗑", "删除当前页", () => this.removeSlide()),
      btn("↑", "上移当前页", () => this.moveSlide(-1)),
      btn("↓", "下移当前页", () => this.moveSlide(1)),
    );
    rail.append(bar);
    const scale = 148 / this.pres.slideWidthPx;
    this.pres.slides.forEach((s, i) => {
      const holder = document.createElement("div");
      holder.className = "thumb" + (i === this.selectedSlide ? " on" : "");
      holder.style.width = "148px";
      holder.style.height = `${this.pres.slideHeightPx * scale}px`;
      holder.innerHTML = `<span>${i + 1}</span><div style="transform:scale(${scale});transform-origin:0 0;width:${this.pres.slideWidthPx}px;height:${this.pres.slideHeightPx}px">${renderSlideHtml(this.pres, i)}</div>`;
      holder.onclick = () => {
        this.commitEditing();
        this.selectedSlide = i;
        this.selected = null;
        this.render();
      };
      rail.append(holder);
    });
    this.rail = rail;

    const stage = document.createElement("div");
    stage.className = "stage";
    this.stage = stage;
    const wrap = document.createElement("div");
    wrap.className = "scaler";
    this.fit = this.fitScale();
    wrap.style.transform = `scale(${this.fit})`;
    wrap.style.transformOrigin = "center center";
    wrap.style.width = `${this.pres.slideWidthPx}px`;
    wrap.style.height = `${this.pres.slideHeightPx}px`;
    wrap.innerHTML = renderSlideHtml(this.pres, this.selectedSlide);
    stage.append(wrap);

    this.wireStage(wrap);

    const right = document.createElement("div");
    right.style.cssText = "flex:1;display:flex;flex-direction:column;min-width:0";
    right.append(stage);
    const notes = document.createElement("div");
    notes.className = "notes";
    const label = document.createElement("b");
    label.textContent = "演讲者备注";
    label.style.color = "#2c3138";
    const body = document.createElement("div");
    body.className = "notes-text";
    body.contentEditable = "plaintext-only";
    body.spellcheck = false;
    const slide = this.slide;
    body.textContent = slide?.notes ?? "";
    body.dataset.empty = slide?.notes ? "0" : "1";
    body.oninput = () => (body.dataset.empty = body.textContent ? "0" : "1");
    body.onblur = () => this.commitNotes(body);
    this.notesEl = body;
    notes.append(label, body);
    right.append(notes);

    deck.append(rail, right);
    this.host.append(deck);
    this.placeSelection();
    this.hooks.selected?.();
  }

  private wireStage(wrap: HTMLElement): void {
    const slideEl = wrap.querySelector(".slide");
    if (!slideEl) return;
    slideEl.addEventListener("pointerdown", (e) => this.stageDown(e as PointerEvent, slideEl as HTMLElement));
    slideEl.addEventListener("dblclick", (e) => this.stageDouble(e as MouseEvent, slideEl as HTMLElement));
  }

  private elementFor(index: number): HTMLElement | undefined {
    return this.host.querySelector<HTMLElement>(`.stage .slide [data-idx="${index}"]`) ?? undefined;
  }

  // ------------------------------------------------------------ interaction

  private stageDown(e: PointerEvent, slideEl: HTMLElement): void {
    if (this.editing) return;
    if (e.button !== 0) return;
    const el = (e.target as HTMLElement).closest<HTMLElement>("[data-idx]");
    if (!el || !slideEl.contains(el)) {
      this.select(null);
      return;
    }
    const index = Number(el.dataset.idx);
    const shape = this.slide?.shapes[index];
    if (!shape) return;
    this.select(index);
    this.drag = { mode: "move", sx: e.clientX, sy: e.clientY, moved: false, box: { x: shape.x, y: shape.y, width: shape.width, height: shape.height }, el };
    capture(el, e.pointerId);
    const move = (ev: PointerEvent): void => this.dragMove(ev);
    const up = (ev: PointerEvent): void => {
      release(el, ev.pointerId);
      el.removeEventListener("pointermove", move);
      el.removeEventListener("pointerup", up);
      this.dragEnd();
    };
    el.addEventListener("pointermove", move);
    el.addEventListener("pointerup", up);
    e.preventDefault();
  }

  private dragMove(e: PointerEvent): void {
    const drag = this.drag;
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    if (!drag || !shape) return;
    const dx = (e.clientX - drag.sx) / this.fit;
    const dy = (e.clientY - drag.sy) / this.fit;
    if (!drag.moved) {
      drag.moved = true;
      this.pushUndo();
    }
    if (drag.mode === "move") {
      shape.x = Math.round(drag.box.x + dx);
      shape.y = Math.round(drag.box.y + dy);
    } else {
      const dir = drag.dir!;
      const box = { ...drag.box };
      if (dir.includes("e")) box.width = Math.max(16, drag.box.width + dx);
      if (dir.includes("s")) box.height = Math.max(16, drag.box.height + dy);
      if (dir.includes("w")) {
        box.width = Math.max(16, drag.box.width - dx);
        box.x = drag.box.x + (drag.box.width - box.width);
      }
      if (dir.includes("n")) {
        box.height = Math.max(16, drag.box.height - dy);
        box.y = drag.box.y + (drag.box.height - box.height);
      }
      // A line keeps its height: PowerPoint draws a non-zero ext as a diagonal.
      if (shape.geometry === "line") {
        box.height = drag.box.height;
        box.y = drag.box.y;
      }
      shape.x = Math.round(box.x);
      shape.y = Math.round(box.y);
      shape.width = Math.round(box.width);
      shape.height = Math.round(box.height);
    }
    this.styleShape(drag.el, shape);
    this.placeSelection();
  }

  private dragEnd(): void {
    const drag = this.drag;
    this.drag = null;
    if (!drag || !this.slide) return;
    const shape = this.selected === null ? undefined : this.slide.shapes[this.selected];
    if (!shape) return;
    const changed = shape.x !== drag.box.x || shape.y !== drag.box.y || shape.width !== drag.box.width || shape.height !== drag.box.height;
    if (!changed) return;
    this.pres.markDirty(this.slide);
    this.hooks.changed();
  }

  private styleShape(el: HTMLElement, shape: Shape): void {
    el.style.left = `${shape.x}px`;
    el.style.top = `${shape.y}px`;
    el.style.width = `${shape.width}px`;
    el.style.height = `${shape.height}px`;
  }

  private stageDouble(e: MouseEvent, slideEl: HTMLElement): void {
    const target = e.target as HTMLElement;
    const el = target.closest<HTMLElement>("[data-idx]");
    if (!el || !slideEl.contains(el)) return;
    const index = Number(el.dataset.idx);
    const shape = this.slide?.shapes[index];
    if (!shape) return;
    const td = target.closest("td");
    if (td && shape.kind === "table") {
      this.editCell(td as HTMLElement, shape);
      return;
    }
    const body = el.querySelector<HTMLElement>(".shape-text");
    if (body && (shape.kind === "shape" || shape.kind === "picture")) this.editText(body, shape);
  }

  private editText(body: HTMLElement, shape: Shape): void {
    body.contentEditable = "true";
    body.style.cursor = "text";
    this.editing = body;
    body.focus();
    const range = document.createRange();
    range.selectNodeContents(body);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    body.dataset.before = JSON.stringify(shape.paragraphs ?? []);
    body.onkeydown = (e) => {
      if (e.key === "Escape") {
        e.preventDefault();
        (e.target as HTMLElement).blur();
      }
    };
    plainPaste(body);
    body.onblur = () => this.commitText();
  }

  private editCell(td: HTMLElement, shape: Shape): void {
    td.contentEditable = "true";
    this.editing = td;
    td.focus();
    const range = document.createRange();
    range.selectNodeContents(td);
    const sel = window.getSelection();
    sel?.removeAllRanges();
    sel?.addRange(range);
    td.onkeydown = (e) => {
      if (e.key === "Escape") {
        e.preventDefault();
        td.blur();
      }
    };
    plainPaste(td);
    td.onblur = () => {
      if (this.editing !== td) return;
      this.editing = null;
      td.contentEditable = "false";
      this.commitCell(td, shape);
    };
  }

  private commitCell(td: HTMLElement, shape: Shape): void {
    const table = td.closest("table");
    if (!table || !shape.rows) return;
    const r = [...table.querySelectorAll("tr")].indexOf(td.closest("tr") as HTMLTableRowElement);
    const c = [...(td.closest("tr")?.children ?? [])].indexOf(td);
    const cell = shape.rows[r]?.[c];
    if (!cell) return;
    const text = (td.textContent ?? "").replace(/\u00a0/g, " ");
    if (cell.text === text) return;
    this.pushUndo();
    cell.text = text;
    cell.runs = text ? [{ text }] : [];
    shape.textDirty = true;
    if (this.slide) this.pres.markDirty(this.slide);
    this.hooks.changed();
    this.render();
  }

  /** Flushes whichever edit is open — a text body or a table cell — into the model. */
  private commitEditing(): void {
    const el = this.editing;
    if (!el) return;
    if (el.tagName === "TD") {
      this.editing = null;
      el.contentEditable = "false";
      const holder = el.closest<HTMLElement>("[data-idx]");
      const index = holder ? Number(holder.dataset.idx) : NaN;
      const shape = Number.isFinite(index) ? this.slide?.shapes[index] : undefined;
      if (shape) this.commitCell(el, shape);
      return;
    }
    this.commitText();
  }

  /** Reads an edited text body back into the shape and repaints. */
  commitText(): void {
    const body = this.editing;
    this.editing = null;
    if (!body) return;
    body.contentEditable = "false";
    const el = body.closest<HTMLElement>("[data-idx]");
    const index = el ? Number(el.dataset.idx) : NaN;
    const shape = Number.isFinite(index) ? this.slide?.shapes[index] : undefined;
    if (!shape) return;
    wrapStrays(body);
    const paragraphs = [...body.children]
      .filter((c) => c.tagName === "DIV")
      .map((p) => {
        const pe = p as HTMLElement;
        return {
          level: Math.round(Number(pe.style.paddingLeft.replace("px", "")) / 18) || 0,
          align: pe.style.textAlign === "justify" ? "both" : pe.style.textAlign || undefined,
          runs: this.runsOf(pe),
        };
      });
    const next = paragraphs.length ? paragraphs : [{ level: 0, runs: [] }];
    if (JSON.stringify(next) === JSON.stringify(shape.paragraphs ?? [])) {
      this.render();
      return;
    }
    this.pushUndo();
    shape.paragraphs = next;
    shape.textDirty = true;
    if (this.slide) this.pres.markDirty(this.slide);
    this.hooks.changed();
    this.render();
  }

  private commitNotes(body: HTMLElement): void {
    const slide = this.slide;
    if (!slide) return;
    const text = (body.textContent ?? "").replace(/\u00a0/g, " ");
    if ((slide.notes ?? "") === text) return;
    slide.notes = text || undefined;
    slide.notesDirty = true;
    this.hooks.changed();
  }

  private runsOf(pEl: HTMLElement): Inline[] {
    const out: Inline[] = [];
    const walk = (node: Node, inherited: Partial<Inline>): void => {
      if (node.nodeType === Node.TEXT_NODE) {
        const text = (node.textContent ?? "").replace(/\u00a0/g, " ");
        if (text) out.push({ ...inherited, text } as Inline);
        return;
      }
      if (node.nodeType !== Node.ELEMENT_NODE) return;
      const el = node as HTMLElement;
      if (el.tagName === "BR") {
        out.push({ ...inherited, text: "\n", break: false } as Inline);
        return;
      }
      const style = this.styleOf(el, inherited);
      for (const child of [...el.childNodes]) walk(child, style);
    };
    const base = this.styleOf(pEl, {});
    for (const child of [...pEl.childNodes]) walk(child, base);
    return out;
  }

  private styleOf(el: HTMLElement, inherited: Partial<Inline>): Partial<Inline> {
    const cs = getComputedStyle(el);
    const px = parseFloat(cs.fontSize);
    const sizePt = Number.isFinite(px) ? Math.round((px * 72) / 96) : undefined;
    return {
      bold: Number(cs.fontWeight) >= 600 || inherited.bold,
      italic: cs.fontStyle === "italic" || inherited.italic,
      underline: cs.textDecorationLine.includes("underline") ? "single" : inherited.underline,
      sizePt,
      color: rgbToHex(cs.color),
      font: cs.fontFamily.split(",")[0]?.replace(/["']/g, "") || inherited.font,
    };
  }

  // ---------------------------------------------------------------- selection

  private select(index: number | null): void {
    this.selected = index;
    this.placeSelection();
    this.hooks.selected?.();
  }

  private placeSelection(): void {
    const slideEl = this.host.querySelector<HTMLElement>(".stage .slide");
    this.selBox?.remove();
    this.selBox = undefined;
    if (!slideEl) return;
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    const el = this.selected === null ? undefined : this.elementFor(this.selected);
    if (!shape || !el) return;
    const box = document.createElement("div");
    box.className = "sel-box";
    box.style.left = `${shape.x}px`;
    box.style.top = `${shape.y}px`;
    box.style.width = `${shape.width}px`;
    box.style.height = `${shape.height}px`;
    const handleSize = Math.round(9 / Math.max(this.fit, 0.2));
    for (const dir of HANDLE_DIRS) {
      const handle = document.createElement("i");
      handle.dataset.dir = dir;
      const pos: Record<Dir, [number, number]> = {
        nw: [0, 0],
        n: [50, 0],
        ne: [100, 0],
        e: [100, 50],
        se: [100, 100],
        s: [50, 100],
        sw: [0, 100],
        w: [0, 50],
      };
      const [px, py] = pos[dir];
      handle.style.left = `calc(${px}% - ${(handleSize * px) / 100}px)`;
      handle.style.top = `calc(${py}% - ${(handleSize * py) / 100}px)`;
      handle.style.width = `${handleSize}px`;
      handle.style.height = `${handleSize}px`;
      handle.addEventListener("pointerdown", (e) => {
        e.stopPropagation();
        e.preventDefault();
        const shape2 = this.slide?.shapes[this.selected ?? -1];
        if (!shape2) return;
        this.drag = { mode: "resize", dir, sx: e.clientX, sy: e.clientY, moved: false, box: { x: shape2.x, y: shape2.y, width: shape2.width, height: shape2.height }, el };
        capture(handle, e.pointerId);
        const move = (ev: PointerEvent): void => this.dragMove(ev);
        const up = (ev: PointerEvent): void => {
          release(handle, ev.pointerId);
          handle.removeEventListener("pointermove", move);
          handle.removeEventListener("pointerup", up);
          this.dragEnd();
        };
        handle.addEventListener("pointermove", move);
        handle.addEventListener("pointerup", up);
      });
      box.append(handle);
    }
    // The rotation lever hangs above the middle of the box, always unrotated.
    const rot = document.createElement("i");
    rot.dataset.dir = "rot";
    const gap = Math.round(24 / Math.max(this.fit, 0.2));
    rot.style.left = `calc(50% - ${handleSize / 2}px)`;
    rot.style.top = `${-gap - handleSize}px`;
    rot.style.width = `${handleSize}px`;
    rot.style.height = `${handleSize}px`;
    rot.addEventListener("pointerdown", (e) => {
      e.stopPropagation();
      e.preventDefault();
      const shape = this.slide?.shapes[this.selected ?? -1];
      if (!shape) return;
      const rect = el.getBoundingClientRect();
      const cx = rect.left + rect.width / 2;
      const cy = rect.top + rect.height / 2;
      const at = (ev: PointerEvent): number => (Math.atan2(ev.clientY - cy, ev.clientX - cx) * 180) / Math.PI;
      const a0 = at(e);
      const r0 = shape.rotation ?? 0;
      let pushed = false;
      capture(rot, e.pointerId);
      const move = (ev: PointerEvent): void => {
        if (!pushed) {
          pushed = true;
          this.pushUndo();
        }
        const deg = r0 + at(ev) - a0;
        shape.rotation = ((Math.round(ev.shiftKey ? Math.round(deg / 15) * 15 : deg) % 360) + 360) % 360;
        el.style.transform = `rotate(${shape.rotation}deg)`;
        this.pres.markDirty(this.slide!);
        this.hooks.changed();
      };
      const up = (ev: PointerEvent): void => {
        release(rot, ev.pointerId);
        rot.removeEventListener("pointermove", move);
        rot.removeEventListener("pointerup", up);
        if (pushed) this.render();
      };
      rot.addEventListener("pointermove", move);
      rot.addEventListener("pointerup", up);
    });
    box.append(rot);
    slideEl.append(box);
    this.selBox = box;
  }

  // --------------------------------------------------------------- commands

  addSlide(): void {
    this.commitEditing();
    this.pushUndo();
    this.pres.addSlide();
    const last = this.pres.slides.length - 1;
    this.pres.moveSlide(last, this.selectedSlide + 1 - last);
    this.selectedSlide += 1;
    this.selected = null;
    this.hooks.changed();
    this.render();
  }

  removeSlide(): void {
    if (this.pres.slides.length <= 1) return;
    this.commitEditing();
    this.pushUndo();
    this.pres.removeSlide(this.selectedSlide);
    this.selectedSlide = Math.max(0, Math.min(this.selectedSlide, this.pres.slides.length - 1));
    this.selected = null;
    this.hooks.changed();
    this.render();
  }

  moveSlide(delta: number): void {
    this.commitEditing();
    this.pushUndo();
    this.pres.moveSlide(this.selectedSlide, delta);
    this.selectedSlide = Math.max(0, Math.min(this.selectedSlide + delta, this.pres.slides.length - 1));
    this.hooks.changed();
    this.render();
  }

  /** Inserts an empty text box centred on the slide and starts typing in it. */
  insertTextBox(): void {
    const w = 320;
    const h = 60;
    this.insertShape(
      {
        kind: "shape",
        textBox: true,
        name: "文本框",
        geometry: "rect",
        x: Math.round((this.pres.slideWidthPx - w) / 2),
        y: Math.round((this.pres.slideHeightPx - h) / 2),
        width: w,
        height: h,
        paragraphs: [{ level: 0, runs: [] }],
      },
      true,
    );
  }

  /** Inserts a preset shape centred on the slide; a line is a zero-height box. */
  insertPreset(geometry: PresetGeometry): void {
    const line = geometry === "line";
    const w = 240;
    const h = line ? 0 : 160;
    this.insertShape({
      kind: "shape",
      name: SHAPE_NAMES[geometry],
      geometry,
      x: Math.round((this.pres.slideWidthPx - w) / 2),
      y: Math.round((this.pres.slideHeightPx - h) / 2),
      width: w,
      height: h,
      fill: line ? undefined : SHAPE_FILL,
      lineColor: SHAPE_LINE,
      lineWidth: line ? 2 : 1,
      paragraphs: [{ level: 0, runs: [] }],
    });
  }

  /** Inserts a decoded image, scaled to fit the slide and centred on it. */
  async insertPicture(url: string, name = "图片"): Promise<void> {
    const { width, height } = await imageSize(url);
    const scale = Math.min(1, (this.pres.slideWidthPx * 0.6) / width, (this.pres.slideHeightPx * 0.6) / height);
    const w = Math.max(24, Math.round(width * scale));
    const h = Math.max(24, Math.round(height * scale));
    this.insertShape({
      kind: "picture",
      name,
      src: url,
      mediaPath: this.pres.addPicture(url),
      x: Math.round((this.pres.slideWidthPx - w) / 2),
      y: Math.round((this.pres.slideHeightPx - h) / 2),
      width: w,
      height: h,
    });
  }

  /** Appends a shape to the current slide as one undo step and selects it. */
  private insertShape(shape: Shape, edit = false): void {
    const slide = this.slide;
    if (!slide) return;
    this.commitEditing();
    this.pushUndo();
    slide.shapes.push(shape);
    this.pres.markDirty(slide);
    this.selected = slide.shapes.length - 1;
    this.hooks.changed();
    this.render();
    if (edit) {
      const body = this.elementFor(this.selected)?.querySelector<HTMLElement>(".shape-text");
      if (body) this.editText(body, shape);
    }
  }

  /** Toggles a run style over the text being edited, or over the selected shape's runs. */
  applyStyle(kind: "bold" | "italic" | "underline"): void {
    if (this.editing) {
      document.execCommand(kind, false);
      return;
    }
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    const runs = (shape?.paragraphs ?? []).flatMap((p) => p.runs);
    if (!shape || !runs.length) return;
    const on = runs.every((r) => (kind === "underline" ? hasLine(r.underline) : !!r[kind]));
    this.pushUndo();
    for (const run of runs) {
      // 按钮在两档之间跳，形状名不在这一跳里：关到 `none`，开到 `single`。
      if (kind === "underline") run.underline = on ? "none" : "single";
      else if (kind === "bold") run.bold = !on;
      else run.italic = !on;
    }
    shape.textDirty = true;
    this.pres.markDirty(this.slide!);
    this.hooks.changed();
    this.render();
  }

  /** Ribbon feedback: what the style buttons should light up, when there is a target. */
  styleState(): { bold: boolean; italic: boolean; underline: boolean } | undefined {
    if (this.editing) {
      return { bold: document.queryCommandState("bold"), italic: document.queryCommandState("italic"), underline: document.queryCommandState("underline") };
    }
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    const runs = (shape?.paragraphs ?? []).flatMap((p) => p.runs);
    if (!runs.length) return undefined;
    return { bold: runs.every((r) => r.bold), italic: runs.every((r) => r.italic), underline: runs.every((r) => hasLine(r.underline)) };
  }

  /** Ribbon feedback: what the shape controls should show, when there is a target. */
  shapeStyleState(): { fill?: string; line?: string; sizePt?: number } | undefined {
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    if (!shape) return undefined;
    const runs = (shape.paragraphs ?? []).flatMap((p) => p.runs);
    const sizes = new Set(runs.map((r) => r.sizePt));
    return { fill: shape.fill, line: shape.lineColor, sizePt: runs.length && sizes.size === 1 ? runs[0].sizePt : undefined };
  }

  /** Applies a font size to the text being edited, or to every run of the selected shape. */
  applySize(pt: number): void {
    if (this.editing) {
      const px = Math.round((pt * 96) / 72);
      this.styleEditing((el) => (el.style.fontSize = `${px}px`));
      return;
    }
    this.mutateRuns((run) => (run.sizePt = pt));
  }

  /** Applies a font colour the same way the size is applied. */
  applyRunColor(hex: string): void {
    if (this.editing) {
      const color = hex.toUpperCase();
      this.styleEditing((el) => (el.style.color = color));
      return;
    }
    this.mutateRuns((run) => (run.color = hex.toUpperCase()));
  }

  /**
   * Paragraph alignment: while text is being edited the paragraph divs take
   * `text-align` (commitText reads it back); otherwise the selected shape's
   * paragraphs change in the model, as one undo step.
   */
  applyAlign(kind: "left" | "center" | "right" | "justify"): void {
    if (this.editing) {
      const body = this.editing;
      const sel = window.getSelection();
      const range = sel && sel.rangeCount ? sel.getRangeAt(0) : undefined;
      const paras = [...body.children].filter((c): c is HTMLElement => c.tagName === "DIV");
      const hit = range && !range.collapsed ? paras.filter((p) => range.intersectsNode(p)) : paras;
      for (const p of hit) p.style.textAlign = kind;
      this.hooks.changed();
      return;
    }
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    const runs = (shape?.paragraphs ?? []).flatMap((p) => p.runs);
    if (!shape || !runs.length) return;
    this.pushUndo();
    for (const p of shape.paragraphs ?? []) p.align = kind === "justify" ? "both" : kind;
    shape.textDirty = true;
    this.pres.markDirty(this.slide!);
    this.hooks.changed();
    this.render();
  }

  applyShapeFill(hex: string): void {
    this.mutateShape((shape) => (shape.fill = hex.toUpperCase()));
  }

  applyShapeLine(hex: string): void {
    this.mutateShape((shape) => {
      shape.lineColor = hex.toUpperCase();
      shape.lineWidth = shape.lineWidth ?? 1;
    });
  }

  /** Rewrites every run of the selected shape through `change`, as one undo step. */
  private mutateRuns(change: (run: Inline) => void): void {
    const shape = this.selected === null ? undefined : this.slide?.shapes[this.selected];
    const runs = (shape?.paragraphs ?? []).flatMap((p) => p.runs);
    if (!shape || !runs.length) return;
    this.pushUndo();
    for (const run of runs) change(run);
    shape.textDirty = true;
    this.pres.markDirty(this.slide!);
    this.hooks.changed();
    this.render();
  }

  private mutateShape(change: (shape: Shape) => void): void {
    if (this.editing || this.selected === null || !this.slide) return;
    const shape = this.slide.shapes[this.selected];
    if (!shape) return;
    this.pushUndo();
    change(shape);
    this.pres.markDirty(this.slide);
    this.hooks.changed();
    this.render();
  }

  /**
   * Styling the selection this way (instead of `execCommand`) keeps the computed px value
   * exact, which is what the px→pt read in `commitText` relies on. Each touched piece of
   * text is wrapped on its own: `commitText` reads paragraphs as direct children of the
   * body, so the selection must never be wrapped as one piece around a paragraph.
   */
  private styleEditing(change: (el: HTMLElement) => void): void {
    const body = this.editing;
    if (!body) return;
    const sel = window.getSelection();
    const range = sel?.rangeCount ? sel.getRangeAt(0) : undefined;
    if (!range || range.collapsed || !body.contains(range.commonAncestorContainer)) {
      for (const el of [body, ...body.querySelectorAll<HTMLElement>("*")]) change(el);
      return;
    }
    const texts: Text[] = [];
    const collect = (node: Node): void => {
      if (range.intersectsNode(node)) {
        if (node.nodeType === Node.TEXT_NODE) texts.push(node as Text);
        else for (const child of [...node.childNodes]) collect(child);
      }
    };
    collect(body);
    const spans: HTMLElement[] = [];
    for (const node of texts) {
      const start = node === range.startContainer ? range.startOffset : 0;
      const end = node === range.endContainer ? range.endOffset : node.length;
      if (start >= end) continue;
      const piece = document.createRange();
      piece.setStart(node, start);
      piece.setEnd(node, end);
      const span = document.createElement("span");
      change(span);
      span.append(piece.extractContents());
      piece.insertNode(span);
      spans.push(span);
    }
    if (!spans.length) return;
    sel?.removeAllRanges();
    const after = document.createRange();
    after.setStartBefore(spans[0]);
    after.setEndAfter(spans[spans.length - 1]);
    sel?.addRange(after);
  }

  private fitScale(): number {
    return Math.min((this.host.clientWidth - 220) / this.pres.slideWidthPx, (this.host.clientHeight - 200) / this.pres.slideHeightPx, 1.4);
  }

  /** Ribbon zoom: an extra factor over the fit-to-window scale. */
  setZoom(z: number): void {
    this.zoomExtra = z;
    this.refit();
  }

  private zoomExtra = 1;

  /** Re-fits the stage after the window changed size, without touching history. */
  refit(): void {
    const wrap = this.host.querySelector<HTMLElement>(".stage .scaler");
    if (!wrap) return;
    this.fit = this.fitScale() * this.zoomExtra;
    wrap.style.transform = `scale(${this.fit})`;
    this.placeSelection();
  }

  deleteShape(): void {
    if (this.editing || this.selected === null || !this.slide) return;
    this.pushUndo();
    this.slide.shapes.splice(this.selected, 1);
    this.selected = null;
    this.pres.markDirty(this.slide);
    this.hooks.changed();
    this.render();
  }

  nudge(dx: number, dy: number): void {
    if (this.editing || this.selected === null || !this.slide) return;
    const shape = this.slide.shapes[this.selected];
    if (!shape) return;
    this.pushUndoSoon();
    shape.x += dx;
    shape.y += dy;
    const el = this.elementFor(this.selected);
    if (el) this.styleShape(el, shape);
    this.placeSelection();
    this.pres.markDirty(this.slide);
    this.hooks.changed();
  }

  // ------------------------------------------------------------- undo / redo

  private snap(): Snap {
    const clone = (s: Shape): Shape => ({
      ...s,
      paragraphs: s.paragraphs?.map((p) => ({ ...p, runs: p.runs.map((r) => ({ ...r })) })),
      rows: s.rows?.map((row) => row.map((c) => ({ ...c, runs: c.runs.map((r) => ({ ...r })) }))),
      children: s.children?.map(clone),
    });
    return {
      slides: this.pres.slides.map((s) => ({ ...s, shapes: s.shapes.map(clone) })),
      removed: new Set(this.pres.removedParts),
      presDirty: this.pres.presDirty,
      addedMedia: [...this.pres.addedMedia],
    };
  }

  private pushUndo(): void {
    this.stack.push(this.snap());
    if (this.stack.length > 40) this.stack.shift();
    this.future.length = 0;
    this.burst = Date.now();
  }

  /** Typing and nudging coalesce into one undo step while they keep coming. */
  private pushUndoSoon(): void {
    if (Date.now() - this.burst < 900) return;
    this.pushUndo();
  }

  private restore(snap: Snap): void {
    this.editing = null;
    this.drag = null;
    this.pres.slides.length = 0;
    this.pres.slides.push(...snap.slides);
    this.pres.removedParts = new Set(snap.removed);
    this.pres.presDirty = snap.presDirty;
    this.pres.addedMedia = new Map(snap.addedMedia);
    this.selectedSlide = Math.max(0, Math.min(this.selectedSlide, this.pres.slides.length - 1));
    this.selected = null;
    this.render();
    this.hooks.changed();
  }

  undo(): void {
    this.commitEditing();
    const snap = this.stack.pop();
    if (!snap) return;
    this.future.push(this.snap());
    this.restore(snap);
  }

  redo(): void {
    this.commitEditing();
    const snap = this.future.pop();
    if (!snap) return;
    this.stack.push(this.snap());
    this.restore(snap);
  }

  /** Flushes an in-progress edit before a save reads the model. */
  commit(): void {
    this.commitEditing();
    if (this.notesEl && document.activeElement === this.notesEl) this.commitNotes(this.notesEl);
  }
}

/** Natural pixel size of an image, with a fallback when the browser cannot decode it. */
function imageSize(url: string): Promise<{ width: number; height: number }> {
  return new Promise((resolve) => {
    const img = new Image();
    img.onload = () => resolve({ width: img.naturalWidth || 320, height: img.naturalHeight || 240 });
    img.onerror = () => resolve({ width: 320, height: 240 });
    img.src = url;
  });
}

function rgbToHex(color: string): string | undefined {
  const m = /^rgba?\((\d+),\s*(\d+),\s*(\d+)/.exec(color);
  if (!m) return undefined;
  const hex = (n: string): string => Number(n).toString(16).padStart(2, "0");
  return `#${hex(m[1])}${hex(m[2])}${hex(m[3])}`.toUpperCase();
}

/** Typing can leave bare text next to the block elements; give it a block. */
function wrapStrays(el: Element): void {
  for (const child of [...el.childNodes]) {
    if (child.nodeType === Node.TEXT_NODE && (child.nodeValue ?? "").trim() !== "") {
      const div = document.createElement("div");
      div.className = "sp-p";
      el.insertBefore(div, child);
      div.append(child);
    }
  }
}

/** Pointer capture keeps a drag alive outside the element; not every pointer supports it. */
function capture(el: HTMLElement, pointerId: number): void {
  try {
    el.setPointerCapture(pointerId);
  } catch {
    /* no active pointer to capture */
  }
}

/** Releasing an uncaptured pointer throws the same way, and must not skip the drag teardown. */
function release(el: HTMLElement, pointerId: number): void {
  try {
    el.releasePointerCapture(pointerId);
  } catch {
    /* capture was never taken */
  }
}

/** Keeps a paste from smuggling markup into a slide the writer would not understand. */
function plainPaste(el: HTMLElement): void {
  el.addEventListener("paste", (e) => {
    const text = e.clipboardData?.getData("text/plain") ?? "";
    if (!text) return;
    e.preventDefault();
    document.execCommand("insertText", false, text);
  });
}
