import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import type { Table } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { pack } from "./helpers/make-containers.js";

/**
 * 表的横向摆位（#213）：`w:tblPr/w:jc` 与 `w:tblPr/w:tblInd` 在纸面上各说各的，两句都要上身。
 *
 * 受众（授权 46 件 / 28 份模板名，`.scratch/tbl-place-audience.py` 逐枚问 `w:tblPr`）：正文里 **58 张表**
 * （10 个全路径件 / 5 份模板名）——`w:jc` 说 `center` **14 枚 / 2 名**（`left` / `right` / `start` / `end`
 * **各 0 枚**，谁都没说 44 枚）；`w:tblInd` 说非 0 缇 **28 枚 / 3 名**（−687…+288，负 **4 枚**），说 0 的 30 枚；
 * **两句同时说的 8 枚**。自家此前在 `center` 那一支直接 `return`，把同张表那枚缩进整枚吞掉，
 * 于是 8 枚明写出来的数一页上看不到。
 *
 * 更要紧的是几何：那 14 枚居中的表里 **5 枚比栏宽还宽**（`bidding document.dotm` 纸 11906、左 1588、右 1474
 * → 栏宽 **8844 缇 = 589.6px**，而网格和是 9,231 / 9,500 / 9,637 / 9,803 / 9,806）。CSS 的
 * `margin-left:auto;margin-right:auto` 在盒子比容器宽时**双双归 0**——整枚溢出全落在右边。所以这里按
 * `calc(50% - 宽/2 ± 缩进)` 摆：`calc` 容得下负值，「居中」这句话这才成立。
 *
 * `w:tblInd` 在对齐一说出之后还作不作数——**WPS 侧未取证**（learn.microsoft.com 在本侧取不到正文；
 * WPS 自带的 `WPSQuickLook.appex` 离线渲染挂住不交图；Apple 的 `textutil` 导入器把 `tblW`/`jc`/`tblInd`
 * 三档全丢，当不了裁判）。取「两句都上身」而不是吞数字，理由写进 `tableAlign` 的头注。
 */
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const cssOf = (bytes: Uint8Array): string =>
  / style="([^"]*)"/.exec(renderWord(open(bytes)).match(/<table\b[^>]*>/)?.[0] ?? "")?.[1] ?? "";
const tableOf = (doc: WordDocument): Table => doc.blocks.filter((b): b is Table => b.kind === "table")[0]!;

/** 一栏一张表；`pr` 就是那枚 `w:tblPr` 的内层，逐枚问的就是它。 */
const oneTable = (pr: string, tw = 2500): Uint8Array =>
  pack(
    `<w:tbl><w:tblPr>${pr}</w:tblPr><w:tblGrid><w:gridCol w:w="${tw}"/><w:gridCol w:w="${tw}"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="${tw}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="${tw}" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  );

const DXA_5000 = '<w:tblW w:w="5000" w:type="dxa"/>';
/** 5000 缇 = 333.3333333333333px，一半是 166.66666666666666px。 */
const HALF_5000 = "166.66666666666666";

describe("docx 表的横向摆位：w:jc 与 w:tblInd 各说各的", () => {
  it("件里那两句先读得出来：宽、对齐、缩进各自进模型，不是渲染时拼出来的", () => {
    const t = tableOf(open(oneTable(`${DXA_5000}<w:jc w:val="center"/><w:tblInd w:w="288" w:type="dxa"/>`)));
    expect([t.widthTwips, t.widthType, t.align, t.indentTwips]).toEqual([5000, "dxa", "center", 288]);
  });

  it("居中且说了缇宽：按 calc 摆，并把同张表那枚缩进加在摆好之后", () => {
    const css = cssOf(oneTable(`${DXA_5000}<w:jc w:val="center"/><w:tblInd w:w="288" w:type="dxa"/>`));
    expect(css).toContain(`width:333.3333333333333px;`);
    expect(css).toContain(`margin-left:calc(50% - ${HALF_5000}px + 19.2px);margin-right:auto;`);
    // 不是「居中即两个 auto」：那两个在盒子过宽时双双归 0，缩进更是一个字都不剩。
    expect(css).not.toContain("margin-left:auto;margin-right:auto");
  });

  it("负缩进写成 ` - 45.8px`，不是 ` + -45.8px`：CSS 的加减号两侧都得有空格", () => {
    const css = cssOf(oneTable(`${DXA_5000}<w:jc w:val="center"/><w:tblInd w:w="-687" w:type="dxa"/>`));
    expect(css).toContain(`margin-left:calc(50% - ${HALF_5000}px - 45.8px);margin-right:auto;`);
    expect(css).not.toMatch(/[+] -/);
    expect(css).not.toContain("+ -");
  });

  it("缩进说 0 与压根没说，摆位是同一串：0 不该多写一句偏移", () => {
    const zero = cssOf(oneTable(`${DXA_5000}<w:jc w:val="center"/><w:tblInd w:w="0" w:type="dxa"/>`));
    const none = cssOf(oneTable(`${DXA_5000}<w:jc w:val="center"/>`));
    expect(zero).toBe(none);
    expect(zero).toContain(`margin-left:calc(50% - ${HALF_5000}px);margin-right:auto;`);
  });

  it("比栏宽还宽那一格：calc 允许负值，溢出因此分给两边（auto 会双双归 0）", () => {
    // 件里的形状：9637 缇的表坐在 8844 缇的栏里（bidding document.dotm 实测）。
    const css = cssOf(oneTable('<w:tblW w:w="9637" w:type="dxa"/><w:jc w:val="center"/><w:tblInd w:w="-687" w:type="dxa"/>', 4818));
    expect(css).toContain("margin-left:calc(50% - 321.23333333333335px - 45.8px);");
    // 摆好的左沿是负的（(8844-9637)/2 - 687 = -1090 缇）——这一格只有 calc 摆得出来。
    expect(css).not.toContain("margin-left:auto");
  });

  it("说了 pct 宽、没说缇宽：退回浏览器的 auto 居中，缩进照样上身", () => {
    // 授权口径里 `center` 且不写 dxa 宽的是 **0 枚**；这一支留给 `pct`（自家模板与别的实现会写）。
    const css = cssOf(oneTable('<w:tblW w:w="5000" w:type="pct"/><w:jc w:val="center"/><w:tblInd w:w="108" w:type="dxa"/>'));
    expect(css).toContain("margin-left:auto;margin-right:auto;position:relative;left:7.2px;");
    expect(css).toContain("width:100%;");
  });

  it("没说对齐、只说缩进：还是按缇摆（这一族旧行为不变）", () => {
    const css = cssOf(oneTable(`${DXA_5000}<w:tblInd w:w="108" w:type="dxa"/>`));
    expect(css).toContain("margin-left:7.2px;");
    expect(css).not.toContain("margin-right:auto");
  });

  it("`right` 那一档 0 受众，故不新增说法：件里真出现了也只按 auto 摆", () => {
    const css = cssOf(oneTable(`${DXA_5000}<w:jc w:val="right"/><w:tblInd w:w="108" w:type="dxa"/>`));
    expect(css).toContain("margin-left:auto;");
    expect(css).not.toContain("calc");
  });

  it("画这一族不改件：渲染一趟再存，字节与不渲染时同一枚", () => {
    const bytes = oneTable(`${DXA_5000}<w:jc w:val="center"/><w:tblInd w:w="288" w:type="dxa"/>`);
    const doc = open(bytes);
    renderWord(doc);
    const saved = writeDocx(doc);
    renderWord(open(saved));
    expect(Buffer.from(writeDocx(open(saved))).equals(Buffer.from(saved))).toBe(true);
    expect(Buffer.from(saved).equals(Buffer.from(writeDocx(open(bytes))))).toBe(true);
  });
});
