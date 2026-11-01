import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { basename, join } from "node:path";
import { strFromU8, unzipSync } from "fflate";
import { paperFit, snapZoom, VIEW_SCROLLBAR_PX, ZOOM_MAX, ZOOM_MIN } from "../src/ui/zoom-fit.js";
import { twipsToPx, WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeZoomFitDocx, makeZoomStatedDocx, PAPER_HEIGHT_TWIPS, PAPER_WIDTH_TWIPS } from "./helpers/make-zoom-fit.js";

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

/**
 * 打开时的起手档先问件里那一句：`word/settings.xml` 的 `w:zoom w:percent`。授权口径 docx 家族 28 件里
 * **27 件说了**（100→19 枚、80→6 枚、120→2 枚，8 枚明说的不是 100），槽只有 `w:percent` 一个
 * （`.scratch/zoom-audience.mjs`）。一句都没说的那一件才落到上面那枚「页宽」。
 */
describe("打开时那起手档：件里自己说的那一句 w:zoom", () => {
  const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
  const settingsOf = (out: Uint8Array): string => text(OpcPackage.from(out).bytes("word/settings.xml") ?? new Uint8Array());

  it("件里说了 80/120/100，读出来的就是那一个数", () => {
    for (const p of ["80", "120", "100"]) expect(open(makeZoomStatedDocx(p)).zoomPercent()).toBe(Number(p));
  });

  it("那一个数换算成档后与 ribbon 共用同一道界：80→0.8、120→1.2、3000→2.5、10→0.4", () => {
    const at = (p: string) => {
      const v = open(makeZoomStatedDocx(p)).zoomPercent();
      expect(v).toBeTypeOf("number");
      return snapZoom((v as number) / 100, false);
    };
    expect(at("80")).toBe(0.8);
    expect(at("120")).toBe(1.2);
    expect(at("3000")).toBe(ZOOM_MAX);
    expect(at("10")).toBe(ZOOM_MIN);
    // 回显那一格报的是落笔后的档：件说 80，格上就是 80%，不是页宽档的 159%。
    expect(Math.round(at("80") * 100)).toBe(80);
  });

  it("件里一句没说就是 undefined，交回视图页那枚「页宽」", () => {
    expect(open(makeZoomFitDocx()).zoomPercent()).toBeUndefined();
    expect(open(makeZoomStatedDocx(null)).zoomPercent()).toBeUndefined();
    // 同一张纸、同一只窗口下，没说时才轮到页宽那一个数。
    expect(snapZoom(paperFit("width", SHELL.viewW, SHELL.viewH, W, H) as number, true)).toBe(1.59);
  });

  it("说了槽没给数、以及名字相像的另一枚孩子，都不算「件里说了」", () => {
    expect(open(makeZoomStatedDocx(null, "<w:zoom/>")).zoomPercent()).toBeUndefined();
    expect(open(makeZoomStatedDocx(null, '<w:zoomChange w:percent="55"/>')).zoomPercent()).toBeUndefined();
    expect(open(makeZoomStatedDocx(null, '<w:hyphenationZone w:val="420"/><w:zoom w:percent="120"/>')).zoomPercent()).toBe(120);
  });

  it("保存不无中生有：说了的那一句原样带走，没说的那份连 settings 这一页都不多出来", () => {
    const out = writeDocx(open(makeZoomStatedDocx("80")));
    expect(settingsOf(out)).toContain('<w:zoom w:percent="80"/>');
    expect(settingsOf(out).match(/<[\w-]+:zoom\b/g)?.length).toBe(1);
    expect(open(out).zoomPercent()).toBe(80);
    const silent = writeDocx(open(makeZoomFitDocx()));
    expect(OpcPackage.from(silent).bytes("word/settings.xml")).toBeUndefined();
  });
});

/**
 * 上面那几条问的是我拼的件；这一组把同一枚读腿放到**授权真件**上（口径与 `.scratch/zoom-audience.mjs` 同：
 * `.scratch/wps-corpus` + WPS 包内 `office6` 模板，按 basename 去重的 docx 家族；只读，不碰 ~/Documents 与
 * ~/Desktop；两处都不在就整组跳过，不拿「没有」当「通过」）。字面那一枚独立地从 zip 取，不借产品的读腿。
 *
 * 认的是**精确**那一页 `word/settings.xml`。那 2 枚 `.dotm` 还各带一页 `word/glossary/settings.xml`，
 * 按 `/(^|\/)settings\.xml$/` 抢先撞上的是没有 zoom 的那一页——普查一度因此把 27 件报成 25 件。
 */
const ZOOM_CORPUS = ".scratch/wps-corpus";
const ZOOM_BUNDLE = "/Applications/wpsoffice.app/Contents/Resources/office6";

function authorizedZoomFiles(): string[] {
  const seen = new Map<string, string>();
  const walk = (dir: string): void => {
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      const p = join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else if (/\.(docx|dotx|dotm|docm)$/i.test(e.name) && !seen.has(basename(p))) seen.set(basename(p), p);
    }
  };
  if (existsSync(ZOOM_CORPUS)) walk(ZOOM_CORPUS);
  if (existsSync(ZOOM_BUNDLE)) walk(ZOOM_BUNDLE);
  return [...seen.values()].sort();
}

const ZOOM_FILES = authorizedZoomFiles();

describe.skipIf(ZOOM_FILES.length === 0)("同一枚读腿在授权真件上（字面 vs 产品，逐件对）", () => {
  const literalOf = (path: string): { said: number | undefined; glossarySays: boolean } => {
    const z = unzipSync(new Uint8Array(readFileSync(path)));
    const main = z["word/settings.xml"];
    const glossary = z["word/glossary/settings.xml"];
    const m = main && /<w:zoom\b[^>]*?w:percent="(\d+(?:\.\d+)?)"/.exec(strFromU8(main));
    const g = glossary && /<w:zoom\b[^>]*?w:percent="(\d+(?:\.\d+)?)"/.exec(strFromU8(glossary));
    return { said: m ? Number(m[1]) : undefined, glossarySays: Boolean(g) };
  };

  it(`${ZOOM_FILES.length} 件里说了那一句的件数与分布，与引证那一句对得上`, () => {
    const stated = ZOOM_FILES.map(literalOf).filter((r) => r.said !== undefined);
    expect(ZOOM_FILES.length).toBe(28);
    expect(stated.length).toBe(27);
    const dist = new Map<number, number>();
    for (const r of stated) dist.set(r.said as number, (dist.get(r.said as number) ?? 0) + 1);
    expect([...dist].sort((a, b) => b[1] - a[1])).toEqual([[100, 19], [80, 6], [120, 2]]);
    // 明说的不是 100 的那 8 枚——起手不该都落 100%，「按存的数起手」这一条才有受众。
    expect(stated.filter((r) => r.said !== 100).length).toBe(8);
  });

  it("逐件：产品答的那一个数就是件里写的那一个数；一句没说的答 undefined", () => {
    for (const path of ZOOM_FILES) {
      const { said } = literalOf(path);
      expect(WordDocument.openBytes(new Uint8Array(readFileSync(path))).zoomPercent(), basename(path)).toBe(said);
    }
  });

  it("带两页同名 settings 的那 2 枚：数来自主页，不是 glossary 那一页", () => {
    const twins = ZOOM_FILES.filter((p) => literalOf(p).glossarySays || unzipSync(new Uint8Array(readFileSync(p)))["word/glossary/settings.xml"]);
    expect(twins.length).toBe(2);
    for (const path of twins) {
      const r = literalOf(path);
      expect(r.glossarySays, basename(path)).toBe(false); // glossary 那一页一句 zoom 都没说
      expect(r.said, basename(path)).toBe(100); // 主页说了 100，产品答的也是 100
      expect(WordDocument.openBytes(new Uint8Array(readFileSync(path))).zoomPercent(), basename(path)).toBe(100);
    }
  });
});
