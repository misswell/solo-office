import { describe, expect, it } from "vitest";
import { paperFit, snapZoom, VIEW_SCROLLBAR_PX, ZOOM_MAX, ZOOM_MIN } from "../src/ui/zoom-fit.js";
import { twipsToPx } from "../src/core/docx.js";
import { PAPER_HEIGHT_TWIPS, PAPER_WIDTH_TWIPS } from "./helpers/make-zoom-fit.js";

/**
 * 视图页 `RB_Zoom` 那两枚按纸定档的腿（WPS `CT_View.kuip` 的 `ZoomPageWidth` / `ZoomOnePage`）的式子。
 * 装机壳开的是 1280×820（`src/main/main.ts` 的窗口），无头探针开的是 756×308——两档的数都从这两只窗口算，
 * 纸则是件里那句 `w:pgSz`（11906×16838 缇，`tests/helpers/make-zoom-fit.ts` 与 DOM／装机两趟共用同一份）。
 */

const W = twipsToPx(PAPER_WIDTH_TWIPS); // 793.7333…px
const H = twipsToPx(PAPER_HEIGHT_TWIPS); // 1122.5333…px

const SHELL = { viewW: 1280, viewH: 820 };
const HEADLESS = { viewW: 756, viewH: 308 };

const fit = (kind: "width" | "page", box = SHELL) => {
  const z = paperFit(kind, box.viewW, box.viewH, W, H);
  expect(z).toBeTypeOf("number");
  return snapZoom(z as number, true);
};

describe("视图页那两枚按纸定档的腿", () => {
  it("纸的像素由件里那句 w:pgSz 来：11906×16838 缇 = 793.7×1122.5px（缇 ÷ 15）", () => {
    expect(W).toBeCloseTo(793.7333333333333, 10);
    expect(H).toBeCloseTo(1122.5333333333334, 10);
  });

  it("页宽档 = 可用宽 ÷ 纸宽：竖滚动条那 15px 从宽里扣，不从高里扣", () => {
    expect(paperFit("width", SHELL.viewW, SHELL.viewH, W, H)).toBeCloseTo((1280 - VIEW_SCROLLBAR_PX) / W, 10);
    // 高那一头不参与页宽档：把可视高换成 1，档不该变。
    expect(paperFit("width", SHELL.viewW, 1, W, H)).toBe(paperFit("width", SHELL.viewW, SHELL.viewH, W, H));
  });

  it("单页档取页宽与「可视高 ÷ 纸高」里小的那一个——装机壳那只窗口里限制的正是高度那一方", () => {
    const byWidth = (SHELL.viewW - VIEW_SCROLLBAR_PX) / W;
    const byHeight = SHELL.viewH / H;
    expect(byHeight).toBeLessThan(byWidth);
    expect(paperFit("page", SHELL.viewW, SHELL.viewH, W, H)).toBeCloseTo(byHeight, 10);
    // 反过来：一列很窄很高的窗口里由宽度那一方说话。
    expect(paperFit("page", 400, 2000, W, H)).toBeCloseTo((400 - VIEW_SCROLLBAR_PX) / W, 10);
  });

  it("按纸定档向下取两位小数：向上取就把纸放得比窗口还大，而那两档说的正是「放得下」", () => {
    const z = fit("page");
    expect(z).toBe(0.73);
    expect(H * z).toBeLessThanOrEqual(SHELL.viewH + 1e-9); // 一整张纸真的放进可视高
    expect(H * (z + 0.01)).toBeGreaterThan(SHELL.viewH); // 再高一档就溢出——所以不能四舍五入上去
    expect(snapZoom(0.7398, true)).toBe(0.73);
    expect(snapZoom(0.7398, false)).toBe(0.74); // ⊖/⊕ 那一族仍四舍五入，同一枚数两种落笔
  });

  it("页宽档落笔后纸宽不溢出可用宽", () => {
    const z = fit("width");
    expect(z).toBe(1.59);
    const usable = SHELL.viewW - VIEW_SCROLLBAR_PX;
    expect(W * z).toBeLessThanOrEqual(usable);
    expect(W * (z + 0.01)).toBeGreaterThan(usable);
  });

  it("单页档不高于页宽档", () => {
    expect(fit("page")).toBeLessThanOrEqual(fit("width"));
    expect(fit("page", HEADLESS)).toBeLessThanOrEqual(fit("width", HEADLESS));
  });

  it("⊕ 的 +0.1 不漂：1.1 加一档是 1.2，不是 1.2000000000000002", () => {
    expect(snapZoom(1.1 + 0.1, false)).toBe(1.2);
    expect(snapZoom(1.1 + 0.1, true)).toBe(1.2);
  });

  it("同一档再点一次不重算出别的数", () => {
    expect(fit("page")).toBe(fit("page"));
  });

  it("无头探针那只 308px 高的窗口里，单页档被自家下界顶住——式子给了 0.27，落笔是 0.4", () => {
    expect(paperFit("page", HEADLESS.viewW, HEADLESS.viewH, W, H)).toBeCloseTo(HEADLESS.viewH / H, 10);
    expect(ZOOM_MIN).toBe(0.4);
    expect(fit("page", HEADLESS)).toBe(ZOOM_MIN);
    expect(snapZoom(9, true)).toBe(ZOOM_MAX); // 上界同一条：250% 以上不再是「看纸」
  });

  it("纸不成数就没有这一档：w:pgSz 没说或说成 0 时交给调用方决定不落笔", () => {
    expect(paperFit("width", SHELL.viewW, SHELL.viewH, 0, H)).toBeUndefined();
    expect(paperFit("page", SHELL.viewW, SHELL.viewH, W, 0)).toBeUndefined();
    expect(paperFit("page", SHELL.viewW, SHELL.viewH, Number.NaN, H)).toBeUndefined();
  });
});
