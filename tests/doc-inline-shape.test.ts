import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { makeShapeDocx } from "./helpers/make-shape.js";

/**
 * 行内 `w:drawing` 里的 DrawingML 组（#238）。授权口径：`.scratch/wgp-census.mjs` 在 21 件 WPS 自带模板（40 份文件按
 * basename 去重）+ 语料里那 6 份 Word 件里数到 11 枚 `wps:style`、分布在 4 件——其中 7 枚落在 `business plan.dotm`
 * 那枚 `wp:inline` 的 38 个形状上（7 枚借主题、31 枚自己在 `a:spPr` 里说），另 4 枚落在页脚的 `wp:anchor` 上；114
 * 份语料的 Word 部件里 0 枚形状。夹具把这几种形状各摆一枚，外加一件没有主题色板的副本：借来的那一半到底画不画得
 * 出来，只有把主题抽走才问得清。
 *
 * 三件事各问一处：读（形状、颜色、单位、跳过的那两枚）、画（SVG 的数、覆盖层的像素、文字不再被当流水里的字）、
 * 带回（一个字符没改时，件里那串字节原样落回盘上）。
 */

const doc = WordDocument.openBytes(makeShapeDocx());
const inlines = doc.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines : []));
const shaped = inlines.filter((i) => i.shape);
const group = shaped[0]!.shape!;
const bare = WordDocument.openBytes(makeShapeDocx({ theme: null }));
const bareGroup = bare.blocks
  .flatMap((b) => (b.kind === "paragraph" ? b.inlines : []))
  .find((i) => i.shape)!
  .shape!;
const html = renderWord(doc);
const xml = text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);

describe("读：组的坐标系与四种形状", () => {
  it("纸上的盒子来自 `wp:extent`，刻度来自 `a:chOff`/`a:chExt`——两套数各归各", () => {
    expect({ w: group.width, h: group.height }).toEqual({ w: 500, h: 300 });
    expect(group.viewBox).toEqual({ x: 0, y: 0, w: 500, h: 600 });
  });

  it("向主题借的填充与描边算得出颜色：`phClr` 由引用自己那句 `a:schemeClr` 上色", () => {
    const box = group.glyphs.find((g) => g.kind === "roundRect")!;
    // `a:fillStyleLst[1]` 是无修饰的 solid phClr，`a:lnStyleLst[1]` 是同一句加自己那支笔（9525 缇＝1px）。
    expect(box).toMatchObject({ x: 10, y: 20, w: 100, h: 60, radius: 15, fill: "#4472C4", stroke: "#ED7D31", strokeWidth: 1 });
  });

  it("`a:ext cy=` 后头那个 0 是一句话，不是一枚空形状：横线仍画", () => {
    const line = group.glyphs.find((g) => g.kind === "line")!;
    expect(line).toMatchObject({ x: 200, y: 300, w: 50, h: 0, stroke: "#767171", strokeWidth: 4 });
    // 连接器的 `a:solidFill` 是它被转成形状时该取的颜色，不是内部要填的墨。
    expect(line.fill).toBeUndefined();
  });

  it("`a:custGeom` 的 `a:path@w/@h` 是它自己的尺，按 `a:ext` 摊进去", () => {
    const path = group.glyphs.find((g) => g.kind === "custGeom")!;
    expect(path.d).toBe("M 0 0 L 80 40 L 0 40 Z");
    expect({ rot: path.rotation, flipH: path.flipH, fill: path.fill }).toEqual({ rot: 90, flipH: true, fill: "#000000" });
  });

  it("`a:noFill` 是说了的「没有」，与什么也没说不是一回事", () => {
    const box = group.glyphs.find((g) => g.kind === "rect")!;
    expect(box.noFill).toBe(true);
    expect(box.fill).toBeUndefined();
  });

  it("画不了的形状不静默消失：预设不认与根本没有盒子，各记一句", () => {
    expect(group.skipped).toEqual(["形状 triangle", "形状无位置"]);
    expect(group.glyphs).toHaveLength(4);
  });

  it("形状里的文字属于图形：`carry.show` 不再把那串字打进流水", () => {
    expect(group.texts).toHaveLength(2);
    expect(group.texts.map((t) => t.anchor)).toEqual(["ctr", "ctr"]);
    expect(group.texts[0]!.paragraphs[0]!.inlines.map((r) => r.text).join("")).toBe("公司");
    expect(inlines.filter((i) => i.shape).every((i) => i.carry!.show === "")).toBe(true);
  });

  it("单独一枚 `wps:wsp`（不在组里）以自己的盒子为刻度", () => {
    const lone = shaped[1]!.shape!;
    expect(lone.viewBox).toEqual({ x: 952500, y: 476250, w: 1905000, h: 952500 });
    expect(lone.glyphs[0]!.radius).toBeCloseTo((952500 * 16667) / 100000, 3);
    expect(lone.glyphs[0]!.fill).toBe("#C00000");
  });

  it("`wp:anchor`（浮动的）仍按原件带着，不画：这一片的版面证据只有行内那一半", () => {
    const anchor = inlines.filter((i) => i.carry?.tag === "drawing" && !i.shape);
    expect(anchor).toHaveLength(1);
    expect(anchor[0]!.carry!.xml).toContain("<wp:anchor");
    // 摆不了位置，但字不跟着位置一起丢：那句话说得出，就仍打在那一段的流水里。
    expect(anchor[0]!.carry!.show).toBe("浮动");
    expect((html.match(/<svg viewBox/g) ?? []).length).toBe(shaped.length);
  });

  it("`mc:Choice` 画一次，`mc:Fallback` 那份 VML 不跟着再画一遍", () => {
    const wrapped = shaped[2]!.shape!;
    expect(wrapped.glyphs).toHaveLength(1);
    expect(wrapped.texts[0]!.paragraphs[0]!.inlines.map((r) => r.text).join("")).toBe("重复");
    expect((html.match(/重复/g) ?? []).length).toBe(1);
    expect(xml.match(/<v:shape/g) ?? []).toHaveLength(1);
  });

  it("主题没有色板时，借来的那一半答不出来并记一句，自己说的颜色照旧", () => {
    expect(bareGroup.skipped).toContain("填充");
    expect(bareGroup.skipped).toContain("描边");
    expect(bareGroup.glyphs.find((g) => g.kind === "roundRect")!.fill).toBeUndefined();
    expect(bareGroup.glyphs.find((g) => g.kind === "line")).toMatchObject({ stroke: "#767171", strokeWidth: 4 });
    expect(bareGroup.glyphs.find((g) => g.kind === "custGeom")!.fill).toBeUndefined();
  });
});

describe("画：SVG 与覆盖层落在那套数上", () => {
  it("`viewBox` 是组的儿童空间，且不做等比保持（两轴缩放不同）", () => {
    expect(html).toContain('<svg viewBox="0 0 500 600" preserveAspectRatio="none"');
  });

  it("笔宽按文件说的绝对像素走（`vector-effect`），不被非等比缩放拉粗", () => {
    expect(html).toContain('stroke="#ED7D31" stroke-width="1" vector-effect="non-scaling-stroke"');
    expect(html).toContain('<line x1="200" y1="300" x2="250" y2="300" stroke="#767171" stroke-width="4"');
  });

  it("`a:custGeom` 先落回自己的盒子、再镜像、最后绕中心转", () => {
    expect(html).toContain('transform="rotate(90 340 120) translate(340 120) scale(-1 1) translate(-340 -120) translate(300 100)"');
  });

  it("文件没涂的形状不涂黑：SVG 的默认墨是黑的", () => {
    expect(html).toContain('<rect x="400" y="500" width="80" height="40" fill="none"/>');
  });

  it("文本框覆盖层落在按两轴各自摊开的像素上", () => {
    // sx = 500/500 = 1，sy = 300/600 = 0.5：10×20 的儿童单位是 10px/10px，100×60 是 100px/30px。
    expect(html).toContain('class="shape-words" style="position:absolute;left:10px;top:10px;width:100px;height:30px;overflow:hidden;display:flex;flex-direction:column;justify-content:center"');
    expect(html).toContain('class="shape-words" style="position:absolute;left:400px;top:250px;width:80px;height:20px;overflow:hidden');
  });

  it("形状里那句话仍走段落渲染：居中、加粗、12 磅都在", () => {
    expect(html).toContain('<div class="para" style="text-align:center;font-size:11pt"><span style="font-weight:600;font-size:12pt;">公司</span>');
  });

  it("跳过的形状在盒子的 title 里说明，纸面不说自己画全了", () => {
    expect(html).toContain('title="未画：形状 triangle、形状无位置"');
  });

  it("覆盖层不借演示页那套 `.shape-text` 样式表——它有自己的名字", () => {
    expect(html).not.toContain('class="shape-text"');
  });
});

describe("带回：一个字符没改时字节原样落回", () => {
  it("组、主题引用与两半 AlternateContent 一句不多一句不少", () => {
    for (const needle of ["<wpg:wgp>", '<a:fillRef idx="1">', 'prst="triangle"', "<mc:AlternateContent>", "<v:shape", "<wp:anchor"]) {
      expect(xml).toContain(needle);
    }
    expect(xml.match(/<wpg:wgp>/g) ?? []).toHaveLength(1);
    // 组里六枚 + 单独一枚 + 浮动一枚 + AlternateContent 的 Choice 一枚；VML 那一半不算（它用 `v:shape`）。
    expect(xml.match(/<wps:wsp>/g) ?? []).toHaveLength(9);
    expect(xml.match(/<w:drawing>/g) ?? []).toHaveLength(4);
  });

  it("写出之后再读回，同一组形状还是那一组", () => {
    const again = WordDocument.openBytes(writeDocx(doc));
    const reparsed = again.blocks
      .flatMap((b) => (b.kind === "paragraph" ? b.inlines : []))
      .find((i) => i.shape)!
      .shape!;
    expect(reparsed.glyphs).toEqual(group.glyphs);
    expect(reparsed.skipped).toEqual(group.skipped);
  });
});
