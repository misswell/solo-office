import type { Presentation } from "../core/pptx.js";
import { renderSlideHtml } from "../core/render.js";
import { fitShrunkText } from "./text-fit.js";

/**
 * Full-screen slideshow: a black overlay over the whole window, one slide at a
 * time, scaled to fit. The window itself goes fullscreen through the shell
 * (macOS native fullscreen needs no user gesture that way); if the shell is
 * absent — a plain browser tab — the overlay still covers the viewport.
 */

const IDLE_MS = 2200;

export class DeckShow {
  private el!: HTMLElement;
  private holder!: HTMLElement;
  private counter!: HTMLElement;
  private index: number;
  private blacked = false;
  private timer = 0;
  private closed = false;

  constructor(
    private pres: Presentation,
    start: number,
    private onExit: () => void,
  ) {
    this.index = Math.max(0, Math.min(start, pres.slides.length - 1));
  }

  open(): void {
    if (this.closed || !this.pres.slides.length) return;
    this.el = document.createElement("div");
    this.el.className = "show";
    this.holder = document.createElement("div");
    this.holder.className = "show-slide";
    this.counter = document.createElement("div");
    this.counter.className = "show-count";
    this.el.append(this.holder, this.counter);
    document.body.append(this.el);

    window.soloOffice?.setFullscreen?.(true);
    this.el.addEventListener("mousedown", this.click);
    this.el.addEventListener("mousemove", this.wake);
    this.el.addEventListener("contextmenu", this.menu);
    window.addEventListener("keydown", this.keys, true);
    window.addEventListener("resize", this.fit);
    this.paint();
  }

  close(): void {
    if (this.closed) return;
    this.closed = true;
    clearTimeout(this.timer);
    this.el.remove();
    window.removeEventListener("keydown", this.keys, true);
    window.removeEventListener("resize", this.fit);
    window.soloOffice?.setFullscreen?.(false);
    this.onExit();
  }

  // ----------------------------------------------------------------- display

  private paint(): void {
    const { slideWidthPx: w, slideHeightPx: h } = this.pres;
    const scale = Math.min(window.innerWidth / w, window.innerHeight / h);
    this.holder.style.width = `${w}px`;
    this.holder.style.height = `${h}px`;
    this.holder.style.transform = `scale(${scale})`;
    this.holder.replaceChildren();
    this.holder.innerHTML = renderSlideHtml(this.pres, this.index);
    fitShrunkText(this.holder);
    this.counter.textContent = `${this.index + 1} / ${this.pres.slides.length}`;
    this.wake();
  }

  private fit = (): void => {
    if (this.closed) return;
    const { slideWidthPx: w, slideHeightPx: h } = this.pres;
    this.holder.style.transform = `scale(${Math.min(window.innerWidth / w, window.innerHeight / h)})`;
  };

  // ------------------------------------------------------------- navigation

  private go(delta: number): void {
    const next = Math.max(0, Math.min(this.index + delta, this.pres.slides.length - 1));
    if (next === this.index && !this.blacked) {
      this.wake();
      return;
    }
    this.blacked = false;
    this.el.classList.remove("black");
    this.index = next;
    this.paint();
  }

  private jump(index: number): void {
    this.blacked = false;
    this.el.classList.remove("black");
    this.index = Math.max(0, Math.min(index, this.pres.slides.length - 1));
    this.paint();
  }

  private toggleBlack(): void {
    this.blacked = !this.blacked;
    this.el.classList.toggle("black", this.blacked);
    this.wake();
  }

  // -------------------------------------------------------------- listeners

  private keys = (e: KeyboardEvent): void => {
    if (this.closed) return;
    if (e.metaKey || e.ctrlKey || e.altKey) return;
    const next = ["ArrowRight", "ArrowDown", "PageDown", "Enter", " "];
    const prev = ["ArrowLeft", "ArrowUp", "PageUp", "Backspace"];
    if (e.key === "Escape") {
      e.preventDefault();
      this.close();
      return;
    }
    if (next.includes(e.key)) {
      e.preventDefault();
      this.go(1);
      return;
    }
    if (prev.includes(e.key)) {
      e.preventDefault();
      this.go(-1);
      return;
    }
    if (e.key === "Home") {
      e.preventDefault();
      this.jump(0);
      return;
    }
    if (e.key === "End") {
      e.preventDefault();
      this.jump(this.pres.slides.length - 1);
      return;
    }
    if (e.key === "b" || e.key === "B" || e.key === ".") {
      e.preventDefault();
      this.toggleBlack();
    }
  };

  /** A click turns a black screen back on; otherwise it advances the deck. */
  private click = (e: MouseEvent): void => {
    if (e.button !== 0) return;
    if (this.blacked) {
      this.toggleBlack();
      return;
    }
    this.go(1);
  };

  private menu = (e: Event): void => e.preventDefault();

  /** Screen activity brings the cursor and the counter back, then they fade out. */
  private wake = (): void => {
    if (this.closed) return;
    this.el.classList.remove("idle");
    clearTimeout(this.timer);
    this.timer = window.setTimeout(() => this.el.classList.add("idle"), IDLE_MS);
  };
}
