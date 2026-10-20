import { describe, expect, it } from "vitest";
import type { Table, TableBorder } from "../src/core/docx.js";
import { WordDocument, borderDelta, syncCellBox, syncCellVAlign, syncTableBox } from "../src/core/docx.js";
import { cellOuterSides } from "../src/core/doc-table-ops.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import {
  CELL_BOX_SIDES, PARA_EDGE_SHAPE, TABLE_EDGE_SHAPE, applyBoxPatch, cellLegPatch, cellOuterPatch, tableLegPatch, TABLE_BOX_SIDES,
} from "../src/ui/doc-para-border.js";
import type { BoxPatch, CellBoxSide, EdgePatch, EdgeShape, TableBoxSide } from "../src/ui/doc-para-border.js";
import { pack } from "./helpers/make-containers.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
/** 写出按文档顺序摆表与格，所以「第几张表、第几格」就是寻址——属性集里不含格里的文字，按文字找是找不到的。 */
const setOf = (xml: string, tag: string, nth: number): string =>
  [...xml.matchAll(new RegExp(`<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>|<w:${tag}\\b[^>]*/>`, "g"))][nth]?.[0] ?? "";
/** 一整枚孩子，连它自己的孩子一起——一枚框线集*就是*它的孩子们，只看开标签什么也证明不了。 */
const childOf = (set: string, tag: string): string =>
  new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`).exec(set)?.[0] ?? "";
const sideOf = (set: string, side: string, tag: string): string => childOf(childOf(set, tag), side);
const tableOf = (doc: WordDocument, nth = 0): Table => doc.blocks.filter((b): b is Table => b.kind === "table")[nth];
/** 落一击：与 `setTableBox` 同两步——补丁落模型，只有动过的边写回件里那一枚孩子；`ground` 是同一击填的地。 */
const clickTable = (doc: WordDocument, patch: BoxPatch<TableBoxSide>, nth = 0, ground?: string): void => {
  const t = tableOf(doc, nth);
  const after = applyBoxPatch(t.borders, patch, TABLE_EDGE_SHAPE);
  syncTableBox(t, borderDelta(t.borders, after), ground);
  t.borders = after;
  if (ground !== undefined) t.shading = ground;
};
/** 落一击在某几格上：与 `setCellBox` 同两步，逐格比、逐格写。 */
const clickCells = (doc: WordDocument, patch: BoxPatch<CellBoxSide>, ...cells: [number, number][]): void => {
  const t = tableOf(doc);
  for (const [r, c] of cells) {
    const cell = t.rows[r][c];
    const after = applyBoxPatch(cell.borders, patch, TABLE_EDGE_SHAPE);
    syncCellBox(cell, borderDelta(cell.borders, after));
    cell.borders = after;
  }
};
/** 只填地、一条边也没点：与 `setCellBox` 面对 `{ shading }` 那一支同两步。 */
const groundCells = (doc: WordDocument, ground: string, ...cells: [number, number][]): void => {
  const t = tableOf(doc);
  for (const [r, c] of cells) {
    const cell = t.rows[r][c];
    syncCellBox(cell, {}, ground);
    cell.shading = ground;
  }
};
/**
 * 页面把表那一圈画在 `<table>` 自己的 style 上，把内部横竖画在**朝缝的那一格**的 `<td>` 上（`renderTable`），
 * 所以两处要分开问；第几枚 `<td>`、哪张 `<table>` 与上面那套数法同一顺序。
 */
const cellCss = (doc: WordDocument, nth: number): string =>
  / style="([^"]*)"/.exec([...renderWord(doc).matchAll(/<td\b[^>]*>/g)][nth]?.[0] ?? "")?.[1] ?? "";
const tableCss = (doc: WordDocument, nth: number): string =>
  / style="([^"]*)"/.exec([...renderWord(doc).matchAll(/<table\b[^>]*>/g)][nth]?.[0] ?? "")?.[1] ?? "";

/**
 * 表那一圈（`w:tblBorders`）与格自己那一圈（`w:tcBorders`）能不能改、能不能存。
 *
 * #144/#145 把段落那一圈（`w:pBdr`）做到了「读得出、画得出、改得动、存得回」，表与格这两层却停在半路：
 * 模型有 `borders`、页面照着画（#29），可写出的那一路在件自带 `tblPrXml` / `tcPrXml` 时**整串原样带回**，
 * 于是用户在 ribbon 上点「所有框线」——画对了，保存后没了。这一族把那句话写回件里自己那一枚孩子，
 * 纪律与段落那三条一样：**逐边**、**对着件里的字节比**、**没动的边一个字节不动**。
 *
 * 表与格的框**有**观众，只是旧版那把尺数错了口径（`.scratch/tblbdr-census.ts` 旧版只扫 `wps-corpus/*.docx|dotx`
 * 6 份，把 58 张表全在的 `.dotm` 模板筛掉了，于是报「0 张表」；本轮把它的文件通配符与 `tblPrEx` 归因都修了）。
 * 授权口径实测：46 份件里 10 份有表、共 58 张 / 4,136 格，`w:tblPr/w:tblBorders` 56 枚（另有 650 枚挂在
 * `w:tblPrEx` 那一行）、`w:tcPr/w:tcBorders` 652 枚，模型读回 56/652 与结构口径**逐份对上**。落笔形状照这些件
 * 与 ECMA-376（§17.4.60 `CT_TblPr`、§17.4.76 `CT_TblBorders`、§17.4.73 `CT_TcPr`、§17.4.77 `CT_TcBorders`）与
 * WPS ribbon 那十条腿的顺序（`CT_Home.kuip:272-287`）；**WPS 自己改完表线写出的件仍未取证**，所以既不造
 * `w:tblPrChange`/`w:tcPrChange`，也不替没动过的边补属性。
 *
 * 一条从这些数里读出来的纪律：这两层「新说一条边」的形状与段落那一族**不同**——件里表级 336 枚边孩子全部
 * `w:space="0"`、`w:color="auto"`，`w:sz` 只有 4 与 12 两档（0.5pt / 1.5pt），**没有一枚 6**；格级 1,924 枚同样
 * `space=0`。所以表与格走 `TABLE_EDGE_SHAPE`（`sz=4 space=0`），段落继续走它自己那 8 枚说的 `sz=6 space=1`。
 */
describe("docx 表格与单元格的「边框与底纹」落笔（w:tblBorders / w:tcBorders / w:shd）", () => {
  /**
   * 表把自己的框说得很全：六边各有各的属性写法，顶边还是成对写法（`CT_Border` 两种都合规），左边只写了
   * `w:sz`（没有 `w:val`，模型认不出这一枚）。旁边摆着 `w:shd`、`w:tblLayout`、`w:tblCellMar`、`w:tblLook`
   * 与一枚没裁决的 `w:tblPrChange`，孩子之间留着换行与空格——重建 `w:tblPr` 时那些文字也得在原地。
   */
  const FULL = pack(
    `<w:p><w:r><w:t>表前</w:t></w:r></w:p>
<w:tbl><w:tblPr>
  <w:tblStyle w:val="a7"/>
  <w:tblW w:w="0" w:type="auto"/>
  <w:tblInd w:w="108" w:type="dxa"/>
  <w:tblBorders><w:top w:val="double" w:sz="12" w:space="2" w:color="FF0000"></w:top><w:left w:sz="4"/><w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="auto"/><w:right w:val="single" w:sz="8" w:space="1" w:color="0070C0"/><w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/><w:insideV w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:tblBorders>
  <w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>
  <w:tblLayout w:type="fixed"/>
  <w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>
  <w:tblLook w:val="04A0" w:firstRow="1"/>
  <w:tblPrChange w:id="41" w:author="林晓" w:date="2026-03-01T08:00:00Z"><w:tblW w:w="0" w:type="auto"/></w:tblPrChange>
</w:tblPr><w:tblGrid><w:gridCol w:w="2000"/><w:gridCol w:w="3000"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:tcBorders><w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>甲格</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t>乙格</w:t></w:r></w:p></w:tc></w:tr>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>丙格</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="3000" w:type="dxa"/><w:tcBorders><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/><w:tl2br w:val="dashed" w:sz="4" w:space="0" w:color="auto"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>丁格</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>表后</w:t></w:r></w:p>`,
  );

  /** 一张什么框也没说的表：新写的 `w:tblBorders` 只能按 `CT_TblPr` 的孩子序插进去。 */
  const BARE = pack(
    `<w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="0" w:type="auto"/><w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/><w:tblLook w:val="04A0"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>独格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  );

  /**
   * 三态读的那一枚：表说了 `w:fill="auto"`（「这张表不要地」），甲格同样说 `auto`，乙格压根没写 `w:shd`，
   * 丙格写了一枚颜色。`auto` 与「没说」在 schema 里是两个句子——把前者折成 `undefined`，就是让上一层的地
   * 穿过这一格显示出来，与文件说的那一句相反。语料里那 1,600 枚 `w:tcPr/w:shd w:fill="auto"` 数的正是这一族
   * （`.scratch/shd-audience.ts`）。
   */
  const THREE = pack(
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="900"/><w:gridCol w:w="900"/><w:gridCol w:w="900"/></w:tblGrid><w:tr>` +
      `<w:tc><w:tcPr><w:tcW w:w="900" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="auto"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="900" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="900" w:type="dxa"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>丙</w:t></w:r></w:p></w:tc>` +
      `</w:tr></w:tbl>`,
  );

  it("从没说过框的表：新槽落在 `CT_TblPr` 给它的那一格，左右邻居一个字节不变", () => {
    const doc = open(BARE);
    clickTable(doc, tableLegPatch("all"));
    const pr = setOf(save(doc), "tblPr", 0);
    expect(childOf(pr, "tblBorders")).toBe(
      `<w:tblBorders>${["top", "left", "bottom", "right", "insideH", "insideV"]
        .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
        .join("")}</w:tblBorders>`,
    );
    // 序：`tblW` → `tblBorders` → `shd`，不是拼到末尾。
    expect(pr.indexOf("<w:tblW")).toBeLessThan(pr.indexOf("<w:tblBorders"));
    expect(pr.indexOf("<w:tblBorders")).toBeLessThan(pr.indexOf("<w:shd"));
    // 件里说的底纹与 `tblStyle` 不因重写而让位。
    expect(pr).toContain('<w:shd w:val="clear" w:color="auto" w:fill="EEEEEE"/>');
    expect(pr).toContain('<w:tblStyle w:val="a3"/>');
  });

  it("表下一腿只换那一边：动过的按件的形状重写，没动的连属性写法一起带着走", () => {
    const doc = open(FULL);
    clickTable(doc, tableLegPatch("bottom"));
    const pr = setOf(save(doc), "tblPr", 0);
    // 件里那一枚说的是 `dotted 4 磅 0 距 auto`：这一击只说「要一条线」，所以换的只有线型，
    // 粗细/间距/颜色仍是件自己说的那一句——补丁没点名的一格不替它编。
    expect(sideOf(pr, "bottom", "tblBorders")).toBe('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    // 成对写法的那一枚顶边：模型认得出（`w:val="double"`），但这一击没说它，所以字节照旧。
    expect(sideOf(pr, "top", "tblBorders")).toBe('<w:top w:val="double" w:sz="12" w:space="2" w:color="FF0000"></w:top>');
    // 只写 `w:sz` 的那一枚左边：`bordersOf` 认不出（`CT_Border@w:val` 缺），整槽重拼就会把它抹掉。
    expect(sideOf(pr, "left", "tblBorders")).toBe('<w:left w:sz="4"/>');
    expect(sideOf(pr, "right", "tblBorders")).toBe('<w:right w:val="single" w:sz="8" w:space="1" w:color="0070C0"/>');
    expect(sideOf(pr, "insideH", "tblBorders")).toBe('<w:insideH w:val="none" w:sz="0" w:space="0" w:color="auto"/>');
  });

  it("同一条腿点第二次是空操作：没有一边动过，`w:tblPr` 连孩子之间的换行都还在原地", () => {
    const doc = open(FULL);
    clickTable(doc, tableLegPatch("cross"));
    const once = save(doc);
    clickTable(doc, tableLegPatch("cross"));
    expect(save(doc)).toBe(once);
    const pr = setOf(once, "tblPr", 0);
    // 兄弟孩子们全在：`w:tblCellMar`、没裁决的 `w:tblPrChange`，以及件里的排版文字。
    expect(pr).toContain('<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>');
    expect(pr).toContain('<w:tblPrChange w:id="41" w:author="林晓"');
    expect(pr).toContain("\n  <w:shd");
    expect(sideOf(pr, "insideH", "tblBorders")).toBe('<w:insideH w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    // 「横竖框线」只说内部两条：外圈四条一个字节没动。
    expect(sideOf(pr, "bottom", "tblBorders")).toBe('<w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="auto"/>');
    expect(sideOf(pr, "top", "tblBorders")).toBe('<w:top w:val="double" w:sz="12" w:space="2" w:color="FF0000"></w:top>');
  });

  it("无框线那一腿把六边全说成不要：`w:val=none` 且 `w:sz=0`，页面跟着不画", () => {
    const doc = open(FULL);
    clickTable(doc, tableLegPatch("none"));
    const pr = setOf(save(doc), "tblPr", 0);
    for (const side of ["top", "left", "bottom", "right", "insideH", "insideV"]) {
      expect(sideOf(pr, side, "tblBorders")).toContain('w:val="none"');
      expect(sideOf(pr, side, "tblBorders")).toContain('w:sz="0"');
    }
    expect(tableCss(doc, 0)).not.toContain("border-top");
    expect(tableCss(doc, 0)).toContain("border-collapse:collapse");
  });

  it("画得出来：改完的那一条线在页上就是改完的样子，不是文件原来的", () => {
    // 内部竖线画在朝缝的那一格（第二行第二列，`<td>` 序号 3）：那一格的 style 就是这一击的页面证据。
    const doc = open(FULL);
    clickTable(doc, tableLegPatch("insideV"));
    expect(cellCss(doc, 3)).toContain("border-left:1px solid currentColor");
    // 顶线换了形状就从 2px double #FF0000 变成 1px solid：页面跟着模型走，不跟着文件旧的那一枚走。
    const other = open(FULL);
    expect(tableCss(other, 0)).toContain("border-top:2px double #FF0000");
    clickTable(other, tableLegPatch("top"));
    expect(tableCss(other, 0)).toContain("border-top:1px solid #FF0000");
    expect(tableCss(other, 0)).not.toContain("double");
  });

  it("保存再读开：那一圈线读得回来，第二次保存是同一串字节", () => {
    const doc = open(FULL);
    clickTable(doc, tableLegPatch("all"));
    const once = save(doc);
    expect(save(open(writeDocx(doc)))).toBe(once);
    // 二次读回时六边都在模型里：所以第三次点同一条腿仍是空操作，不会「再落一次」多出字节。
    const reopened = open(writeDocx(doc));
    expect(Object.keys(tableOf(reopened).borders ?? {}).sort()).toEqual([
      "bottom",
      "insideH",
      "insideV",
      "left",
      "right",
      "top",
    ]);
    clickTable(reopened, tableLegPatch("all"));
    expect(save(reopened)).toBe(once);
  });

  it("格那一腿逐格落：本来就有这一边的格换那一边，没被点到的格一枚 `w:tcBorders` 也不许多出来", () => {
    const doc = open(FULL);
    clickCells(doc, cellLegPatch("bottom"), [0, 0], [1, 1]);
    const xml = save(doc);
    // 甲格原本有一条 `w:bottom`（虚线 4 磅 0 距 #0070C0）：同一边不同形状，所以它换。
    expect(sideOf(setOf(xml, "tcPr", 0), "bottom", "tcBorders")).toBe('<w:bottom w:val="single" w:sz="4" w:space="0" w:color="0070C0"/>');
    // 底纹是另一枚孩子，不因为重写框线集而让位或改序。
    expect(setOf(xml, "tcPr", 0)).toContain('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>');
    // 丁格的斜线继续按原件走：这一击只说底线。
    expect(sideOf(setOf(xml, "tcPr", 3), "tl2br", "tcBorders")).toBe('<w:tl2br w:val="dashed" w:sz="4" w:space="0" w:color="auto"/>');
    // 乙、丙两格没被点到。
    expect(childOf(setOf(xml, "tcPr", 1), "tcBorders")).toBe("");
    expect(childOf(setOf(xml, "tcPr", 2), "tcBorders")).toBe("");
  });

  it("格没有 `w:tcBorders` 时新槽按 `CT_TcPr` 的孩子序插，不挤走 `w:vAlign`", () => {
    const doc = open(FULL);
    clickCells(doc, cellLegPatch("all"), [0, 1]);
    const pr = setOf(save(doc), "tcPr", 1);
    expect(pr.indexOf("<w:tcW")).toBeLessThan(pr.indexOf("<w:tcBorders"));
    expect(pr.indexOf("<w:tcBorders")).toBeLessThan(pr.indexOf("<w:vAlign"));
    expect(childOf(pr, "tcBorders")).toBe(
      `<w:tcBorders>${["top", "left", "bottom", "right"]
        .map((side) => `<w:${side} w:val="single" w:sz="4" w:space="0" w:color="auto"/>`)
        .join("")}</w:tcBorders>`,
    );
  });

  it("内部横竖那两句只有表说得出：格的补丁里没有这一族，斜线也不被顺手改写", () => {
    expect(Object.keys(cellLegPatch("all").sides ?? {})).toEqual(["top", "left", "bottom", "right"]);
    const doc = open(FULL);
    const cell = tableOf(doc).rows[1][1];
    const after = applyBoxPatch(cell.borders, cellLegPatch("all"), TABLE_EDGE_SHAPE);
    const changes = borderDelta(cell.borders, after);
    expect(changes).not.toHaveProperty("insideH");
    expect(changes).not.toHaveProperty("tl2br");
    syncCellBox(cell, changes);
    cell.borders = after;
    expect(sideOf(setOf(save(doc), "tcPr", 3), "tl2br", "tcBorders")).toBe('<w:tl2br w:val="dashed" w:sz="4" w:space="0" w:color="auto"/>');
  });

  it("表中之表：外层那一击不改写内层自己的 `w:tblBorders`", () => {
    const nested = pack(
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="4000"/></w:tblGrid><w:tr><w:tc><w:tcPr><w:tcW w:w="4000" w:type="dxa"/></w:tcPr>` +
        `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="dotted" w:sz="6" w:space="1" w:color="7030A0"/></w:tblBorders></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>内层</w:t></w:r></w:p></w:tc></w:tr></w:tbl>` +
        `<w:p><w:r><w:t>外层</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const doc = open(nested);
    clickTable(doc, tableLegPatch("all"));
    const xml = save(doc);
    // 外层：顶边换了，另五边新写；内层：一枚字节也不动。
    expect(sideOf(setOf(xml, "tblPr", 0), "top", "tblBorders")).toBe('<w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/>');
    expect(childOf(setOf(xml, "tblPr", 0), "tblBorders")).toContain("<w:insideV");
    expect(setOf(xml, "tblPr", 1)).toBe(
      '<w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="dotted" w:sz="6" w:space="1" w:color="7030A0"/></w:tblBorders></w:tblPr>',
    );
  });

  it("那几条腿各说什么：外侧框线不碰内部线，横竖框线不碰外圈，所有框线说满六边", () => {
    expect(Object.keys(tableLegPatch("outside").sides ?? {})).toEqual(["top", "left", "bottom", "right"]);
    expect(Object.keys(tableLegPatch("cross").sides ?? {})).toEqual(["insideH", "insideV"]);
    expect(Object.keys(tableLegPatch("all").sides ?? {})).toEqual(["top", "left", "bottom", "right", "insideH", "insideV"]);
    expect(Object.keys(tableLegPatch("none").sides ?? {})).toEqual(["top", "left", "bottom", "right", "insideH", "insideV"]);
    expect(tableLegPatch("insideH").sides).toEqual({ insideH: { val: "single", sizeEighths: 4 } });
    expect(tableLegPatch("none").sides?.top).toEqual({ val: "none" });
    expect(Object.keys(cellLegPatch("none").sides ?? {})).toEqual(["top", "left", "bottom", "right"]);
  });

  it("同一句「要一条线」在两层各说各的磅数：段落 6 磅 1 距、表与格 4 磅 0 距", () => {
    // 段落那一层的 `sz=6 space=1` 是它那 8 枚 `w:pBdr` 边孩子说的；表与格那一层 336 + 1,924 枚边孩子全带
    // `w:space="0"`、`w:sz` 只有 4/12/0 三档。共用一个默认就是替其中一层编数——编出来的线既更厚又带偏移。
    const at = (shape: EdgeShape): string => {
      const edge = applyBoxPatch<"top">({}, { sides: { top: { val: "single" } } }, shape)!.top!;
      return `${edge.size}/${edge.space}/${edge.color}`;
    };
    expect(at(PARA_EDGE_SHAPE)).toBe("6/1/auto");
    expect(at(TABLE_EDGE_SHAPE)).toBe("4/0/auto");
    const off = (shape: EdgeShape): string => {
      const edge = applyBoxPatch<"top">({}, { sides: { top: { val: "none" } } }, shape)!.top!;
      return `${edge.size}/${edge.space}/${edge.color}`;
    };
    // 说了不要线：磅归零，间距按本层的形状（件里 `none` 那 4 枚表级边全是 `sz=0 space=0`）。
    expect(off(TABLE_EDGE_SHAPE)).toBe("0/0/auto");
    expect(off(PARA_EDGE_SHAPE)).toBe("0/1/auto");
  });

  it("对话框那一路「原样确定」：磅填出去再填回来是同一个字节，那一层没摆的斜线也不算改动", () => {
    const doc = open(FULL);
    const table = tableOf(doc)!;
    // 复刻对话框那两只手：格子里写的是 `String(size / 8)`，填回来的是 `Math.round(格子 × 8)`。件里的 4/6/8/12 八分之一磅
    // 过这一对换算必须一个不差，否则「原样确定」会把没动的边重写一遍——那是 #150 把同一枚对话框接到表与格两层之后
    // 才有的路：这两层的 `w:sz` 与段落那一族不同，格子里还多了 `w:space` 与颜色的往返。
    const echo = (edge: TableBorder): EdgePatch => ({
      val: edge.val,
      ...(edge.size === undefined ? {} : { sizeEighths: Math.round(Number(String(edge.size / 8)) * 8) }),
      ...(edge.space === undefined ? {} : { spacePt: Number(String(edge.space)) }),
      ...(edge.color === undefined ? {} : { color: edge.color }),
    });
    const forTable = {} as Partial<Record<TableBoxSide, EdgePatch>>;
    for (const side of TABLE_BOX_SIDES) {
      const edge = table.borders?.[side];
      if (edge) forTable[side] = echo(edge);
    }
    // 件里那句 `<w:left w:sz="4"/>` 没写 `w:val`，模型认不出这一枚（`borderOf`）：所以对话框摆得出的是五边，
    // 那一边在格子里是「不改」，这一击替它编不出一句话。
    expect(Object.keys(forTable)).toEqual(["top", "bottom", "right", "insideH", "insideV"]);
    expect(borderDelta(table.borders, applyBoxPatch(table.borders, { sides: forTable }, TABLE_EDGE_SHAPE))).toEqual({});
    // 丁格：说自己那一圈的顶线，还多带一条 `w:tl2br`——格那一层只摆四条边，没摆出来的那条不该被算成改动。
    const cell = table.rows[1][1]!;
    const forCell = {} as Partial<Record<CellBoxSide, EdgePatch>>;
    for (const side of CELL_BOX_SIDES) {
      const edge = cell.borders?.[side];
      if (edge) forCell[side] = echo(edge);
    }
    expect(Object.keys(forCell)).toEqual(["top"]);
    expect(borderDelta(cell.borders, applyBoxPatch(cell.borders, { sides: forCell }, TABLE_EDGE_SHAPE))).toEqual({});
    // 真动了一格的那一边：这一边按对话框填的那四句换，没摆出来的斜线与它的写法一起留在原地。
    const after = applyBoxPatch(cell.borders, { sides: { top: { val: "thick", sizeEighths: 18, spacePt: 3, color: "C00000" } } }, TABLE_EDGE_SHAPE);
    syncCellBox(cell, borderDelta(cell.borders, after));
    const tc = setOf(save(doc), "tcPr", 3);
    expect(sideOf(tc, "top", "tcBorders")).toBe('<w:top w:val="thick" w:sz="18" w:space="3" w:color="C00000"/>');
    expect(sideOf(tc, "tl2br", "tcBorders")).toBe('<w:tl2br w:val="dashed" w:sz="4" w:space="0" w:color="auto"/>');
  });

  it("三态读：`w:fill=\"auto\"` 是「说了不要地」，没写 `w:shd` 才是「没说」", () => {
    const doc = open(THREE);
    const table = tableOf(doc)!;
    expect(table.shading).toBe("auto");
    expect(table.rows[0]!.map((c) => c.shading)).toEqual(["auto", undefined, "#FFF2CC"]);
    // 画的那一边跟着这一判据：认得出的颜色才上身，`auto` 与「没说」都不铺——但两者回到件里是两种字节。
    expect(tableCss(doc, 0)).not.toContain("background");
    expect(cellCss(doc, 0)).not.toContain("background");
    expect(cellCss(doc, 1)).not.toContain("background");
    expect(cellCss(doc, 2)).toContain("background:#FFF2CC");
    // 往返不能对着「什么都没改的件」问：`writeDocx` 在 `!doc.edited` 时把源件原样交回，那种相同什么也不证明。
    // 所以这一族在**真动了一格**之后问：丙格换成 `auto`，甲、乙两格没被点，各自的字节得留在原地。
    groundCells(doc, "auto", [0, 2]);
    const xml = save(doc);
    expect(childOf(setOf(xml, "tcPr", 2), "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    expect(setOf(xml, "tcPr", 0)).toContain('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    expect(childOf(setOf(xml, "tcPr", 1), "shd")).toBe("");
    const again = open(writeDocx(doc));
    expect(tableOf(again)!.rows[0]!.map((c) => c.shading)).toEqual(["auto", undefined, "auto"]);
    expect(save(again)).toBe(xml);
  });

  it("表那一层填地：`w:shd` 换在原位，`w:tblBorders` 一枚字节不动（连件里那句没 `w:val` 的左边一起）", () => {
    const doc = open(FULL);
    const before = tableOf(doc)!.tblPrXml!;
    clickTable(doc, {} as BoxPatch<TableBoxSide>, 0, "#C00000");
    // 这一击确实要走重生成那条路，不是走「原样带回源件」的快路。
    expect(doc.edited).toBe(true);
    const pr = setOf(save(doc), "tblPr", 0);
    expect(childOf(pr, "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="C00000"/>');
    // 换掉的是同一枚孩子，不是末尾再插一枚。
    expect([...pr.matchAll(/<w:shd\b/g)]).toHaveLength(1);
    // 没点边，所以整槽框线按件的字节走——含 `bordersOf` 认不出的那一句。
    expect(childOf(pr, "tblBorders")).toBe(childOf(before, "tblBorders"));
    expect(sideOf(pr, "left", "tblBorders")).toBe('<w:left w:sz="4"/>');
    // 孩子之间那句排版文字与另外两枚邻居也在原地。
    expect(pr).toContain("\n  <w:shd");
    expect(pr).toContain('<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>');
    expect(pr).toContain('<w:tblPrChange w:id="41" w:author="林晓"');
    expect(pr.indexOf("<w:tblBorders")).toBeLessThan(pr.indexOf("<w:shd"));
    expect(pr.indexOf("<w:shd")).toBeLessThan(pr.indexOf("<w:tblLayout"));
    expect(tableCss(doc, 0)).toContain("background:#C00000");
    const again = open(writeDocx(doc));
    expect(tableOf(again)!.shading).toBe("#C00000");
    expect(save(again)).toBe(save(doc));
  });

  it("「原样确定」那一击：把件里那一枚地填回去，`w:tblPr` 与 `w:tcPr` 都是同一个字节", () => {
    const doc = open(FULL);
    const table = tableOf(doc)!;
    // 对话框那一只手摆出来的是模型给的颜色形状（`#F2F2F2`），填回来还是它自己（`fillValue` 只补 `#` 与大写）。
    // 这一条问的是**写回那一步**的字节，所以对着 `tblPrXml` 本身问：一个字节没动的文档 `edited` 是假，
    // 保存会把源件原样交回，那种「相同」什么也没证明。
    expect(table.shading).toBe("#F2F2F2");
    const pr = table.tblPrXml!;
    syncTableBox(table, {}, table.shading);
    expect(table.tblPrXml).toBe(pr);
    const cell = table.rows[0][0]!;
    const tc = cell.tcPrXml!;
    syncCellBox(cell, {}, cell.shading);
    expect(cell.tcPrXml).toBe(tc);
    // 正控制：同一双手换个数填就得真的动——上面那两条「相同」不是这台机器什么都不做。
    syncTableBox(table, {}, "#C00000");
    expect(table.tblPrXml).not.toBe(pr);
    expect(childOf(table.tblPrXml!, "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="C00000"/>');
    expect(childOf(table.tblPrXml!, "tblBorders")).toBe(childOf(pr, "tblBorders"));
  });

  it("地按件的形状写：`auto` 回到件里还是小写，件没写过 `w:shd` 时按 `CT_TblPr` 的序新插一枚", () => {
    const doc = open(FULL);
    clickTable(doc, {} as BoxPatch<TableBoxSide>, 0, "auto");
    // `shadingFillToken` 只把六位十六进制转大写；`auto` 原样走——转成大写就是替件在颜色的位置上签一个 schema 没有的词。
    expect(childOf(setOf(save(doc), "tblPr", 0), "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    const noShd = pack(
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:tblBorders><w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/></w:tblBorders><w:tblLayout w:type="fixed"/><w:tblLook w:val="04A0"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>独格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const other = open(noShd);
    expect(tableOf(other)!.shading).toBeUndefined();
    clickTable(other, tableLegPatch("top"), 0, "#FFF2CC");
    const pr = setOf(save(other), "tblPr", 0);
    expect(pr.indexOf("<w:tblBorders")).toBeLessThan(pr.indexOf("<w:shd"));
    expect(pr.indexOf("<w:shd")).toBeLessThan(pr.indexOf("<w:tblLayout"));
    // 补头的形状照语料那 1,676 枚唯一的一种组合（`val`+`color`+`fill`，全是 `w:val="clear"`）。
    expect(childOf(pr, "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/>');
    // 这一击点的是「上框线」那一腿，所以顶边的磅数按表这一层的形状换成 4；件里说过的 `w:space="1"` 仍是件的那一句。
    expect(sideOf(pr, "top", "tblBorders")).toBe('<w:top w:val="single" w:sz="4" w:space="1" w:color="auto"/>');
    expect(pr).toContain('<w:tblLook w:val="04A0"/>');
  });

  it("件里没写 `w:tblPr` 的表与编辑器新建的表同一条路：地由字段拼出来，读回来是同一个数", () => {
    const prLess = pack(
      `<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p><w:r><w:t>新表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const doc = open(prLess);
    expect(tableOf(doc)!.tblPrXml).toBeUndefined();
    clickTable(doc, {} as BoxPatch<TableBoxSide>, 0, "#00B050");
    const xml = save(doc);
    expect(setOf(xml, "tblPr", 0)).toBe('<w:tblPr><w:shd w:val="clear" w:color="auto" w:fill="00B050"/></w:tblPr>');
    expect(tableCss(doc, 0)).toContain("background:#00B050");
    const again = open(writeDocx(doc));
    expect(tableOf(again)!.shading).toBe("#00B050");
    expect(save(again)).toBe(xml);
  });

  it("格那一层只填地：`w:tcBorders`（含斜线）一个字节不动，新地按 `CT_TcPr` 的序摆在它后面", () => {
    const doc = open(FULL);
    groundCells(doc, "auto", [0, 0], [1, 1]);
    const xml = save(doc);
    const jia = setOf(xml, "tcPr", 0);
    expect(childOf(jia, "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    expect(sideOf(jia, "bottom", "tcBorders")).toBe('<w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/>');
    const ding = setOf(xml, "tcPr", 3);
    expect(childOf(ding, "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="auto"/>');
    expect(sideOf(ding, "top", "tcBorders")).toBe('<w:top w:val="single" w:sz="6" w:space="1" w:color="auto"/>');
    expect(sideOf(ding, "tl2br", "tcBorders")).toBe('<w:tl2br w:val="dashed" w:sz="4" w:space="0" w:color="auto"/>');
    expect(ding.indexOf("<w:tcBorders")).toBeLessThan(ding.indexOf("<w:shd"));
    // 乙格那一格没有 `w:tcBorders`，只有 `w:vAlign`：新地摆在两者之间，不挤走 vAlign。
    groundCells(doc, "#00B050", [0, 1]);
    const yi = setOf(save(doc), "tcPr", 1);
    expect(yi.indexOf("<w:tcW")).toBeLessThan(yi.indexOf("<w:shd"));
    expect(yi.indexOf("<w:shd")).toBeLessThan(yi.indexOf("<w:vAlign"));
    expect(childOf(yi, "shd")).toBe('<w:shd w:val="clear" w:color="auto" w:fill="00B050"/>');
    // 丙格没被点到，一枚 `w:shd` 也不许多出来。
    expect(childOf(setOf(save(doc), "tcPr", 2), "shd")).toBe("");
    expect(cellCss(doc, 0)).not.toContain("background");
    expect(cellCss(doc, 1)).toContain("background:#00B050");
    // 保存再读开：三格各自的话都读得回来，第二次保存同一串字节。
    const again = open(writeDocx(doc));
    expect(tableOf(again)!.rows.map((r) => r.map((c) => c.shading))).toEqual([
      ["auto", "#00B050"],
      [undefined, "auto"],
    ]);
    expect(save(again)).toBe(save(doc));
  });
});

describe("格竖向对齐 w:tcPr/w:vAlign", () => {
  const cells = (doc: WordDocument) => tableOf(doc)!.rows.flat();
  const FIXTURE = pack(
    `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/><w:tcBorders><w:bottom w:val="dotted" w:sz="4" w:space="0" w:color="0070C0"/></w:tcBorders><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/></w:tcPr><w:p><w:r><w:t>甲格</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t>乙格</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>丙格</w:t></w:r></w:p></w:tc>` +
      `<w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>丁格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  );

  it("改一档是换槽：同一枚孩子只留一处，位置仍按 CT_TcPr 的孩子序", () => {
    const doc = open(FIXTURE);
    const yi = cells(doc)[1]!;
    expect(yi.valign).toBe("bottom");
    syncCellVAlign(yi, "center");
    yi.valign = "center";
    const pr = setOf(save(doc), "tcPr", 1);
    expect([...pr.matchAll(/<w:vAlign\b/g)].length).toBe(1);
    expect(childOf(pr, "vAlign")).toBe('<w:vAlign w:val="center"/>');
    expect(pr.indexOf("<w:tcW")).toBeLessThan(pr.indexOf("<w:vAlign"));
  });

  it("格本来没写这一句时按序插进去，不挤走相邻的底纹", () => {
    const doc = open(FIXTURE);
    const jia = cells(doc)[0]!;
    syncCellVAlign(jia, "top");
    jia.valign = "top";
    const pr = setOf(save(doc), "tcPr", 0);
    expect(childOf(pr, "vAlign")).toBe('<w:vAlign w:val="top"/>');
    expect(pr.indexOf("<w:tcBorders")).toBeLessThan(pr.indexOf("<w:shd"));
    expect(pr.indexOf("<w:shd")).toBeLessThan(pr.indexOf("<w:vAlign"));
  });

  it("保存再读开三档都读得回来，第二次保存字节不变", () => {
    const doc = open(FIXTURE);
    const [jia, yi, bing] = cells(doc);
    syncCellVAlign(jia!, "top");
    jia!.valign = "top";
    syncCellVAlign(yi!, "center");
    yi!.valign = "center";
    syncCellVAlign(bing!, "bottom");
    bing!.valign = "bottom";
    const again = open(writeDocx(doc));
    expect(cells(again).map((c) => c.valign)).toEqual(["top", "center", "bottom", undefined]);
    expect(save(again)).toBe(save(doc));
  });

  it("格子没有原件的 `w:tcPr` 时由写出器按字段拼，槽这一条不动它", () => {
    const bare = pack(
      `<w:tbl><w:tblPr><w:tblW w:w="5000" w:type="pct"/></w:tblPr><w:tblGrid><w:gridCol w:w="5000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:p><w:r><w:t>无属性格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const doc = open(bare);
    const cell = cells(doc)[0]!;
    syncCellVAlign(cell, "center");
    expect(cell.tcPrXml).toBeUndefined();
    cell.valign = "center";
    expect(save(doc)).toContain('<w:vAlign w:val="center"/>');
  });
});

/**
 * #151：「外侧框线」落在格选区时画的是**选区的外圈**，不是每格自己的四条边。WPS 那一枚控件没铺满整张表时
 * 只能逐格说（`CT_TcBorders` 没有「内部」那一枚孩子），所以这里先钉几何：哪一格、哪几条边朝向块外。
 * 选区按整格算——合并格伸出选区的那一段也算在块里；竖合并的一带由起头那一格答。WPS 自己写出的字节未取证。
 */
describe("格选区「外侧框线」的外圈几何（#151）", () => {
  const plated = (rows: string[][]): Uint8Array => {
    const cell = (t: string): string =>
      `<w:tc><w:tcPr><w:tcW w:w="1000" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>${t}</w:t></w:r></w:p></w:tc>`;
    const tr = rows.map((r) => `<w:tr>${r.map(cell).join("")}</w:tr>`).join("");
    return pack(
      `<w:tbl><w:tblPr><w:tblW w:w="${rows[0]!.length * 1000}" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
        `<w:tblGrid>${rows[0]!.map(() => `<w:gridCol w:w="1000"/>`).join("")}</w:tblGrid>${tr}</w:tbl>`,
    );
  };
  const shape = (t: Table, cells: { row: number; col: number }[]): [string, string][] =>
    cellOuterSides(t, cells).map(({ where, sides }) => [`${where.row},${where.col}`, [...sides].sort().join("|")]);

  it("表中间铺 2×2：每格只写朝向块外的那两条边", () => {
    const t = tableOf(open(plated([
      ["甲", "乙", "丙", "丁"],
      ["戊", "己", "庚", "辛"],
      ["壬", "癸", "子", "丑"],
      ["寅", "卯", "辰", "巳"],
    ])));
    expect(shape(t, [{ row: 1, col: 1 }, { row: 1, col: 2 }, { row: 2, col: 1 }, { row: 2, col: 2 }])).toEqual([
      ["1,1", "left|top"],
      ["1,2", "right|top"],
      ["2,1", "bottom|left"],
      ["2,2", "bottom|right"],
    ]);
  });

  it("铺满整张：角上两临边、边上一条、正中一条也不写", () => {
    const t = tableOf(open(plated([["甲", "乙", "丙"], ["丁", "戊", "己"], ["庚", "辛", "壬"]])));
    const all = [0, 1, 2].flatMap((row) => [0, 1, 2].map((col) => ({ row, col })));
    expect(shape(t, all)).toEqual([
      ["0,0", "left|top"], ["0,1", "top"], ["0,2", "right|top"],
      ["1,0", "left"], ["1,1", ""], ["1,2", "right"],
      ["2,0", "bottom|left"], ["2,1", "bottom"], ["2,2", "bottom|right"],
    ]);
  });

  it("横向合并格伸出选区的那一段也按整格算：右边那条边仍算块外", () => {
    const t = tableOf(open(pack(
      `<w:tbl><w:tblPr><w:tblW w:w="2000" w:type="dxa"/><w:tblLayout w:type="fixed"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="1000"/><w:gridCol w:w="1000"/></w:tblGrid>` +
        `<w:tr><w:tc><w:tcPr><w:tcW w:w="2000" w:type="dxa"/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>横跨</w:t></w:r></w:p></w:tc></w:tr>` +
        `<w:tr><w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>丙</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    )));
    expect(shape(t, [{ row: 0, col: 0 }, { row: 1, col: 0 }, { row: 1, col: 1 }])).toEqual([
      ["0,0", "left|right|top"],
      ["1,0", "bottom|left"],
      ["1,1", "bottom|right"],
    ]);
  });

  it("落笔只碰外圈：件里恰 8 条边，块内的两道缝一条也不写", () => {
    const doc = open(plated([
      ["a11", "a12", "a13", "a14"],
      ["a21", "a22", "a23", "a24"],
      ["a31", "a32", "a33", "a34"],
      ["a41", "a42", "a43", "a44"],
    ]));
    const t = tableOf(doc);
    for (const { where, sides } of cellOuterSides(t, [{ row: 1, col: 1 }, { row: 1, col: 2 }, { row: 2, col: 1 }, { row: 2, col: 2 }])) {
      const cell = t.rows[where.row]![where.col]!;
      const after = applyBoxPatch(cell.borders, cellOuterPatch(sides), TABLE_EDGE_SHAPE);
      syncCellBox(cell, borderDelta(cell.borders, after));
      cell.borders = after;
    }
    const xml = save(doc);
    const chunk = (needle: string): string =>
      [...xml.matchAll(/<w:tc\b[\s\S]*?<\/w:tc>/g)].map((m) => m[0]).find((s) => s.includes(needle)) ?? "";
    expect([...xml.matchAll(/<w:top w:val="single"/g)]).toHaveLength(2);
    expect([...xml.matchAll(/<w:bottom w:val="single"/g)]).toHaveLength(2);
    expect([...xml.matchAll(/<w:left w:val="single"/g)]).toHaveLength(2);
    expect([...xml.matchAll(/<w:right w:val="single"/g)]).toHaveLength(2);
    // 块内那两道缝：己格不说右边，庚格不说左边；癸格不说右边，子格不说左边。
    expect(chunk("a22")).toContain('<w:top w:val="single"');
    expect(chunk("a22")).toContain('<w:left w:val="single"');
    expect(chunk("a22")).not.toContain("<w:right");
    expect(chunk("a22")).not.toContain("<w:bottom");
    expect(chunk("a23")).not.toContain("<w:left");
    expect(chunk("a32")).not.toContain("<w:right");
    expect(chunk("a33")).not.toContain("<w:left");
    // 外圈之外那一圈格一枚 `w:tcBorders` 也没多出来。
    expect(chunk("a11")).not.toContain("<w:tcBorders");
    expect(chunk("a44")).not.toContain("<w:tcBorders");
  });
});
