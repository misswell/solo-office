import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeGridlinesDocx, makeRowSignatureDocx, makeSeamInkDocx } from "./helpers/make-tables.js";

/**
 * 「查看网格线」那一条的三条腿（#196）。
 *
 * 要判的事：纸上那一格有没有边，**只问文件**——`w:tcBorders` 与表替这条缝/这一圈说的那句；
 * 而「屏幕上看得见格子的地界」是另一件事，落在导引线上（`--gl` 那一串内阴影），画不画由容器那一枚 `doc-grid` 决定。
 * 原先 `src/ui/index.html` 那条 `.page .doc-table td { border:1px solid #b9bfc7 }` 把两件事并成了一件：
 * 文件明说「这一条边不要线」的地方也照画。授权 46 份里量到 44 条这样的边（去重同名两份模板后 22 条、
 * 落在 2 件模板；`.scratch/gridlines-audience.py`）。
 *
 * 夹具 `makeGridlinesDocx` 把 16 条格边凑齐四种说法：格说的要画（丙底）、格说的不要（甲左）、
 * 表替这一圈说的（top 单线）、表替这条缝说的不要（insideV 单线 none）与两人都没说（insideH 缺席、左右底圈缺席）。
 */
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
/** 页面渲染出的那四枚 `td`，按文档序（甲/乙/丙/丁），取它们各自的 `style`。 */
const cellStyles = (bytes: Uint8Array): string[] =>
  [...renderWord(open(bytes)).matchAll(/<td([^>]*)>/g)].map((m) => /style="([^"]*)"/.exec(m[1] ?? "")?.[1] ?? "");

/**
 * 导引清单里那四条边的写法互不重样（`inset 0 1px 0` 与 `inset 0 -1px 0`、`inset 1px 0 0` 与 `inset -1px 0 0`
 * 两两不构成子串），所以按写法认边，不需要每条边各一枚自定义属性。
 */
const GUIDE: Record<"top" | "right" | "bottom" | "left", string> = {
  top: "inset 0 1px 0",
  bottom: "inset 0 -1px 0",
  left: "inset 1px 0 0",
  right: "inset -1px 0 0",
};
/** 这一格的那串导引（`--gl:` 后面到下一个 `;` 为止），没有导引时是空串。 */
const gl = (style: string): string => /--gl:([^;]*)/.exec(style)?.[1] ?? "";
const guided = (style: string, side: keyof typeof GUIDE): boolean => gl(style).includes(GUIDE[side]);
const guideCount = (style: string): number => (gl(style).match(/inset/g) ?? []).length;

describe("docx 格子边的墨：只按文件说的那句画", () => {
  const styles = cellStyles(makeGridlinesDocx());

  it("导引清单是一条完整的 `box-shadow` 列表：要么整串是内阴影，要么整枚缺席", () => {
    // 分成四槽（`--gl-top` …）再拿 `none` 兜底是错的：`box-shadow` 的列表里塞一枚 `none` 非法，
    // 整条声明作废，「这一格只缺两条边」的那种最常见情形反而一条导引也画不出来。
    for (const style of styles) {
      expect(style).not.toMatch(/--gl-(top|right|bottom|left):/); // 清单里那句 `var(--gl-ink)` 不算槽
      const list = gl(style);
      if (list) expect(list.split(",").every((one) => one.startsWith("inset "))).toBe(true);
    }
  });

  it("格自己说明白不要的那条边，不写行内 border，只留一条导引线", () => {
    const jia = styles[0];
    expect(guided(jia, "left")).toBe(true);
    expect(jia).not.toContain("border-left:");
  });

  it("格自己说要画的那条边，写行内 border 且不留导引线", () => {
    const bing = styles[2];
    expect(bing).toContain("border-bottom:");
    expect(guided(bing, "bottom")).toBe(false);
  });

  it("外圈那条线是表说的：第一行不再替它叠一条导引", () => {
    // `w:tblBorders/w:top` 说了单线，画在 `<table>` 那一圈上；第一行的两格自己没说话，
    // 导引线要认这一句，否则纸上的圈线外侧又压一条。
    const tableTag = /<table class="doc-table"[^>]*>/.exec(renderWord(open(makeGridlinesDocx())))?.[0] ?? "";
    expect(tableTag).toContain("border-collapse:collapse");
    expect(tableTag).toContain("border-top:");
    expect(guided(styles[0], "top")).toBe(false);
    expect(guided(styles[1], "top")).toBe(false);
    // 第二行的顶边归 `w:insideH` 管，而这张表没说过它：那一条才该留导引。
    expect(guided(styles[2], "top")).toBe(true);
  });

  it("表替这条缝说的「不要线」（w:insideV val=none）不再被画上一条", () => {
    // 甲的右边与乙的左边是同一条缝：甲在 c0（右边=缝），乙在 c1（左边=缝），两人都没说过要画。
    expect(guided(styles[0], "right")).toBe(true);
    expect(guided(styles[1], "left")).toBe(true);
    expect(styles[0]).not.toContain("border-right:");
    expect(styles[1]).not.toContain("border-left:");
  });

  it("16 条边各归各位：3 条由文件说要画（写 border），13 条没人说要画（留 --gl）", () => {
    expect(styles.reduce((n, s) => n + guideCount(s), 0)).toBe(13);
    // 甲：顶边是表说的圈线，其余三条没人说。乙：顶边同样是圈线，右边是这一圈的外沿（表没说）。
    expect(guideCount(styles[0])).toBe(3);
    expect(guideCount(styles[1])).toBe(3);
    // 丙：底边自己说要画。丁：底与右都归外沿而表没说，顶与左归两条没线的缝。
    expect(guideCount(styles[2])).toBe(3);
    expect(guideCount(styles[3])).toBe(4);
  });

  it("一张什么都没说过的表：一条边都不写，四格各留四条导引", () => {
    const silent = cellStyles(makeRowSignatureDocx());
    expect(silent.length).toBeGreaterThan(0);
    for (const style of silent) {
      for (const side of ["top", "right", "bottom", "left"] as const) expect(guided(style, side)).toBe(true);
      expect(guideCount(style)).toBe(4);
      expect(style).not.toMatch(/(^|;)border-/);
    }
  });

  it("导引线不落进件里：渲染过一趟之后，写出的字节一句 `--gl` 也没有", () => {    const doc = open(makeGridlinesDocx());
    renderWord(doc);
    const xml = text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
    expect(xml).not.toContain("--gl");
    expect(xml).not.toContain("box-shadow");
    // 文件原本说的那几句边一句不多、一句不少（导引线是屏幕上的东西，不该顺手替格子补 `w:tcBorders`）。
    expect(xml.match(/<w:tcBorders>/g)?.length).toBe(2);
    expect(xml).toContain('<w:insideV w:val="none"');
  });
});

/**
 * 一条缝只有一头说了话（#196 的续）。夹具 `makeSeamInkDocx`：表什么都没说，只有甲自己交代了底边要单线。
 * 甲的底与丙的顶在 `border-collapse:collapse` 下是同一条物理线——丙若还留顶导引，画得出的墨外侧又压一条。
 */
describe("docx 缝上的墨在另一头那格说的：导引要让位", () => {
  const styles = cellStyles(makeSeamInkDocx());

  it("甲自己说要画的那条底边：写 border 且不留导引", () => {
    expect(styles[0]).toContain("border-bottom:");
    expect(guided(styles[0], "bottom")).toBe(false);
  });

  it("丙没说话，但缝的另一头（甲）说要画：丙的顶边不再留导引", () => {
    expect(guided(styles[2], "top")).toBe(false);
    // 其余三条仍没人画线：顶边让位不等于把这一格的其他边一起吞掉。
    expect(guided(styles[2], "bottom")).toBe(true);
    expect(guided(styles[2], "left")).toBe(true);
    expect(guided(styles[2], "right")).toBe(true);
  });

  it("另一头的格子没说话的缝照旧留导引：乙的底、丁的顶", () => {
    expect(guided(styles[1], "bottom")).toBe(true);
    expect(guided(styles[3], "top")).toBe(true);
  });

  it("16 条边凑齐：一条有墨，一条把位置让给同一缝上的墨，其余 14 条留导引", () => {
    expect(styles.reduce((n, s) => n + guideCount(s), 0)).toBe(14);
    // 甲少了底（有墨）、丙少了顶（同一缝）；乙与丁四条边全归导引（表没说圈线也没说缝）。
    expect(guideCount(styles[0])).toBe(3);
    expect(guideCount(styles[1])).toBe(4);
    expect(guideCount(styles[2])).toBe(3);
    expect(guideCount(styles[3])).toBe(4);
  });

  it("导引线不落进件里：缝上另一头那格也不该被补一句 `w:tcBorders`", () => {
    const doc = open(makeSeamInkDocx());
    renderWord(doc);
    const xml = text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
    expect(xml).not.toContain("--gl");
    expect(xml.match(/<w:tcBorders>/g)?.length).toBe(1);
    expect(xml.match(/<w:bottom /g)?.length).toBe(1);
  });
});

describe("页面那套 CSS：发明的格子边切掉，导引线上不了纸", () => {
  const css = readFileSync(new URL("../src/ui/index.html", import.meta.url), "utf8").replace(/\/\*[\s\S]*?\*\//g, "");

  it("`.page .doc-table td` 不再替每一格发明一条边", () => {
    const rule = /\.page \.doc-table td \{([^}]*)\}/.exec(css)?.[1] ?? "";
    expect(rule).not.toContain("border");
    expect(rule).toContain("padding");
  });

  it("导引线由容器那一枚 doc-grid 决定，画的是 box-shadow（不进盒模型）", () => {
    const rule = /#view\.doc-grid \.page \.doc-table td \{([^}]*)\}/.exec(css)?.[1] ?? "";
    expect(rule).toContain("box-shadow");
    expect(rule).toContain("var(--gl,");
    // 四槽各拿 `none` 兜底的那条写法会让整份导引在最常见的格子上作废（列表里塞 `none` 非法）。
    expect(rule).not.toMatch(/--gl-(top|right|bottom|left)/);
    expect(rule).not.toMatch(/(^|;)\s*border:/);
  });

  it("打印那一档把导引线按住：屏幕上的东西不上纸", () => {
    const print = /@media print \{([\s\S]*?)\n      \}/.exec(css)?.[1] ?? "";
    expect(print).toContain("#view .page .doc-table td");
    expect(print).toMatch(/box-shadow:\s*none\s*!important/);
  });
});
