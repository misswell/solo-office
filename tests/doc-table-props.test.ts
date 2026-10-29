import { describe, expect, it } from "vitest";
import { WordDocument, rowShapeOf, signRowRevision, signTableRevision, syncRowAlign, syncTableProps, tableShapeOf } from "../src/core/docx.js";
import type { RowInfo, Table } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { pack } from "./helpers/make-containers.js";
import { EDIT_TABLES, makeTablePlaceDocx } from "./helpers/make-table-place.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
/**
 * 第 n 枚 `w:tblPr` 连同它自己的嵌套孩子——一枚待裁决的 `w:tblPrChange` 里**又有一枚** `w:tblPr`，
 * 所以不能拿「遇到第一个 `</w:tblPr>`」当这枚孩子的结尾（那样记录会被截断，正好把要钉的那一段切掉）。
 * 往里走的那一趟只能在 `depth` 归零的那一刻停：多走一次 `exec` 就把**下一枚** `<w:tblPr>` 吞了，
 * 于是 n>=1 时读到的是隔一张的表（三张表里 n=1 会读到第三张）。
 */
const prOf = (xml: string, nth = 0): string => {
  const re = /<w:tblPr\b[^>]*?(\/?)>|<\/w:tblPr>/g;
  let seen = 0;
  for (let m = re.exec(xml); m; m = re.exec(xml)) {
    if (m[0].startsWith("</")) continue;
    if (m[0].endsWith("/>")) {
      if (seen++ === nth) return m[0];
      continue;
    }
    const start = m.index;
    let depth = 1;
    while (depth > 0) {
      const n = re.exec(xml);
      if (!n) break;
      if (n[0].startsWith("</")) depth -= 1;
      else if (!n[0].endsWith("/>")) depth += 1;
    }
    if (depth !== 0) break;
    if (seen === nth) return xml.slice(start, re.lastIndex);
    seen += 1;
  }
  return "";
};
const childOf = (set: string, tag: string): string =>
  new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`).exec(set)?.[0] ?? "";
const tableOf = (doc: WordDocument): Table => doc.blocks.filter((b): b is Table => b.kind === "table")[0]!;
const tablesAt = (doc: WordDocument): Table[] => doc.blocks.filter((b): b is Table => b.kind === "table");
const tableCss = (doc: WordDocument): string => / style="([^"]*)"/.exec(renderWord(doc).match(/<table\b[^>]*>/)?.[0] ?? "")?.[1] ?? "";
/** 逐张的行内样式，按文档序——装机腿里「两张画到同一处」那一判据的文件层对照。 */
const cssAt = (doc: WordDocument, n: number): string => (renderWord(doc).match(/<table\b[^>]*>/g) ?? []).map((t) => / style="([^"]*)"/.exec(t)?.[1] ?? "")[n] ?? "";
const cellCss = (doc: WordDocument): string => / style="([^"]*)"/.exec(renderWord(doc).match(/<td\b[^>]*>/)?.[0] ?? "")?.[1] ?? "";

/**
 * 「表格属性」第一片的落笔处：`w:jc` / `w:tblInd` / `w:tblLayout` / `w:tblCellMar` 这四句写回表自己的
 * `w:tblPr`。为什么是这四句、为什么别的几句不在这条腿里，见 `src/ui/doc-table-props.ts` 与
 * `syncTableProps` 的头注；受众是授权 46 份的 10 件、58 张表（`.scratch/tblpr-children-audience.py`）。
 *
 * 这一族钉的是与 `syncTableBox` 同一条纪律：**没点名的孩子一个字节不动**，动过的只换它自己那一枚，
 * 并按 `CT_TblPr` 的孩子序留在原位。WPS 那张对话框里摆哪些字段**未取证**，所以这里不拿它的分组当证据，
 * 只钉住「件里说过的那几句能被改、能存、能读回来」。
 */
describe("docx 表格属性：w:tblPr 里那四句的落笔处", () => {
  /** 表把八句都说全（宽高、框、地、布局、边距、look），孩子之间留着换行加两格空格。 */
  const FULL = pack(
    `<w:p><w:r><w:t>表前</w:t></w:r></w:p>
<w:tbl><w:tblPr>
  <w:tblStyle w:val="a7"/>
  <w:tblW w:w="5000" w:type="dxa"/>
  <w:jc w:val="center"/>
  <w:tblInd w:w="288" w:type="dxa"/>
  <w:tblBorders><w:top w:val="single" w:sz="4" w:space="0" w:color="auto"/></w:tblBorders>
  <w:shd w:val="clear" w:color="auto" w:fill="F2F2F2"/>
  <w:tblLayout w:type="fixed"/>
  <w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>
  <w:tblLook w:val="04A0" w:firstRow="1"/>
  <w:tblPrChange w:id="41" w:author="林晓" w:date="2026-03-01T08:00:00Z"><w:jc w:val="left"/></w:tblPrChange>
</w:tblPr><w:tblGrid><w:gridCol w:w="2500"/><w:gridCol w:w="2500"/></w:tblGrid>
<w:tr><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:tcPr><w:tcW w:w="2500" w:type="dxa"/></w:tcPr><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r><w:t>表后</w:t></w:r></w:p>`,
  );

  /** 四句一句都没说的表：新槽只能按 `CT_TblPr` 的孩子序插进去。 */
  const BARE = pack(
    `<w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>独格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
  );

  it("四句各换各的：只动点过名的那一枚，兄弟孩子连孩子之间的换行都留在原地", () => {
    const doc = open(FULL);
    const table = tableOf(doc);
    const before = table.tblPrXml!;
    syncTableProps(table, { align: "right", indentTwips: -129, layout: "", cellMargin: { left: 0 } });
    const pr = prOf(save(doc));

    // 对齐：换在原位，不是末尾再插一枚（表自己的那一层只留一枚；记录里那枚 `w:jc` 是「改之前」的另一套）。
    const own = pr.slice(0, pr.indexOf("<w:tblPrChange"));
    expect([...own.matchAll(/<w:jc\b/g)]).toHaveLength(1);
    expect(childOf(pr, "jc")).toBe('<w:jc w:val="right"/>');
    // 左缩进收负数：件里 −129 / −440 / −72 各 2 枚，负缩进是它写过的东西，不是错误输入。
    expect(childOf(pr, "tblInd")).toBe('<w:tblInd w:w="-129" w:type="dxa"/>');
    // 不勾「固定列宽」= 把那一句删掉：授权 46 份里 `autofit` 一枚也没写过，替文件补一句就是编。
    expect(pr).not.toContain("<w:tblLayout");
    // 默认单元格边距：点名的左边换数，没点名的三边连属性写法一起原样留着。
    expect(childOf(pr, "tblCellMar")).toBe(
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="0" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>',
    );

    // 序：`CT_TblPr` 的孩子序，不是「按这一击的先后」。
    const at = (tag: string): number => pr.indexOf(`<w:${tag}`);
    expect(at("tblW")).toBeLessThan(at("jc"));
    expect(at("jc")).toBeLessThan(at("tblInd"));
    expect(at("tblInd")).toBeLessThan(at("tblBorders"));
    expect(at("tblBorders")).toBeLessThan(at("shd"));
    expect(at("shd")).toBeLessThan(at("tblCellMar"));
    expect(at("tblCellMar")).toBeLessThan(at("tblLook"));
    expect(at("tblLook")).toBeLessThan(at("tblPrChange"));

    // 没点名的孩子逐字节相同（框、地、look、待裁决的记录，以及 `tblStyle`/`tblW`）。
    for (const tag of ["tblStyle", "tblW", "tblBorders", "shd", "tblLook", "tblPrChange"]) {
      expect(childOf(pr, tag)).toBe(childOf(before, tag));
    }
    expect(pr).toContain("\n  <w:shd");
    expect(pr).toContain("\n  <w:tblLook");
  });

  it("原样确定不碰文件：空补丁一个字节不改；同一条补丁点第二次也是空操作", () => {
    const doc = open(FULL);
    const table = tableOf(doc);
    const before = table.tblPrXml!;
    syncTableProps(table, {});
    expect(table.tblPrXml).toBe(before);
    // 正控制：同一双手真点一句就得动——上面那条「相同」不是这台机器什么都不做。
    syncTableProps(table, { align: "left" });
    expect(table.tblPrXml).not.toBe(before);
    // 再点同一句：值已经就是它，所以没有第二轮改写。
    const once = table.tblPrXml;
    syncTableProps(table, { align: "left" });
    expect(table.tblPrXml).toBe(once);
  });

  it("缺的孩子按 `CT_TblPr` 的序插进去：jc → tblInd → tblLayout → tblCellMar，不挤走 tblLook", () => {
    const doc = open(BARE);
    syncTableProps(tableOf(doc), { align: "center", indentTwips: 108, layout: "fixed", cellMargin: { top: 0, left: 108, bottom: 0, right: 108 } });
    const pr = prOf(save(doc));
    expect(childOf(pr, "jc")).toBe('<w:jc w:val="center"/>');
    expect(childOf(pr, "tblInd")).toBe('<w:tblInd w:w="108" w:type="dxa"/>');
    // 勾上「固定列宽」写的正是件里那 58 枚的形状：`w:type="fixed"`，没有别的属性。
    expect(childOf(pr, "tblLayout")).toBe('<w:tblLayout w:type="fixed"/>');
    // 新写的边照件的拼法带 `w:w` 与 `w:type="dxa"`（授权口径 1,502 枚边孩子无一例外）。
    expect(childOf(pr, "tblCellMar")).toBe(
      '<w:tblCellMar><w:top w:w="0" w:type="dxa"/><w:left w:w="108" w:type="dxa"/><w:bottom w:w="0" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar>',
    );
    const at = (tag: string): number => pr.indexOf(`<w:${tag}`);
    expect(at("tblW")).toBeLessThan(at("jc"));
    expect(at("jc")).toBeLessThan(at("tblInd"));
    expect(at("tblInd")).toBeLessThan(at("tblLayout"));
    expect(at("tblLayout")).toBeLessThan(at("tblCellMar"));
    expect(at("tblCellMar")).toBeLessThan(at("tblLook"));
  });

  it("读回来：保存再打开四句都在模型里，第二次保存是同一串字节", () => {
    const doc = open(FULL);
    syncTableProps(tableOf(doc), { align: "right", indentTwips: -129, layout: "", cellMargin: { left: 0, top: 42 } });
    const once = save(doc);
    const again = open(writeDocx(doc));
    const table = tableOf(again);
    expect(table.align).toBe("right");
    expect(table.indentTwips).toBe(-129);
    expect(table.layout).toBeUndefined();
    expect(table.cellMargin).toEqual({ top: 42, left: 0, bottom: 0, right: 108 });
    expect(save(again)).toBe(once);
    // 读回来之后同一条补丁仍是空操作：模型与件的说法一致，不会「再落一次」多出字节。
    syncTableProps(table, { align: "right", indentTwips: -129, layout: "", cellMargin: { left: 0 } });
    expect(save(again)).toBe(once);
  });

  it("画得出来：对齐、缩进、固定列宽与默认边距都按模型走，不是按件里那一句", () => {
    const before = open(FULL);
    // 居中且说了缇宽（5000 缇 = 333.33px）：按 `calc` 摆，同张表那枚 288 缇（19.2px）加在摆好之后。
    // 旧的「两个 auto」在这里会双双归 0，把 288 缇整枚吞掉——见 `tableAlign` 头注与 #213。
    expect(tableCss(before)).toContain("margin-left:calc(50% - 166.66666666666666px + 19.2px);margin-right:auto");
    expect(tableCss(before)).toContain("table-layout:fixed");
    // 108 缇 = 7.2px（`twipsToPx` 的 15 缇/px）。
    expect(cellCss(before)).toContain("padding-left:7.2px");

    const doc = open(FULL);
    syncTableProps(tableOf(doc), { align: "left", indentTwips: 108, layout: "", cellMargin: { left: 54 } });
    const css = tableCss(doc);
    // 左对齐 + 缩进：页面用 `margin-left` 摆，不再居中。
    expect(css).not.toContain("margin-left:auto");
    expect(css).toContain("margin-left:7.2px");
    // 不勾固定列宽：页面回到按内容伸缩。
    expect(css).not.toContain("table-layout:fixed");
    // 左边距 54 缇 = 3.6px。
    expect(cellCss(doc)).toContain("padding-left:3.6px");
  });

  it("认不出的对齐（`start`）不在这一击里被改写：只点缩进时那一枚 `w:jc` 逐字节不动", () => {
    const start = pack(
      `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/><w:jc w:val="start"/><w:tblInd w:w="0" w:type="dxa"/></w:tblPr>` +
        `<w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>格</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const doc = open(start);
    const table = tableOf(doc);
    expect(table.align).toBe("start");
    syncTableProps(table, { indentTwips: 288 });
    const pr = prOf(save(doc));
    expect(childOf(pr, "jc")).toBe('<w:jc w:val="start"/>');
    expect(childOf(pr, "tblInd")).toBe('<w:tblInd w:w="288" w:type="dxa"/>');
  });

  it("开着修订：编辑器造的表留下的记录里带着默认边距（写出器按字段拼 `prev` 那一趟）", () => {
    const bare = pack(
      `<w:tbl><w:tblGrid><w:gridCol w:w="900"/></w:tblGrid><w:tr><w:tc><w:p><w:r><w:t>新表</w:t></w:r></w:p></w:tc></w:tr></w:tbl>`,
    );
    const doc = open(bare);
    const table = tableOf(doc);
    expect(table.tblPrXml).toBeUndefined();
    syncTableProps(table, { align: "center", cellMargin: { left: 108 } });
    const was = { ...tableShapeOf(table), align: undefined, cellMargin: { left: 54 } };
    signTableRevision(table, { id: 7, author: "我", date: "2026-10-25T00:00:00Z" }, was);
    const pr = prOf(save(doc));
    expect(childOf(pr, "jc")).toBe('<w:jc w:val="center"/>');
    expect(childOf(pr, "tblCellMar")).toBe('<w:tblCellMar><w:left w:w="108" w:type="dxa"/></w:tblCellMar>');
    // 记录里那一套是「改之前」：左边距 54、没有说话的对齐（记录不替它补一句）。
    const rec = childOf(pr, "tblPrChange");
    expect(rec).toContain('<w:tblCellMar><w:left w:w="54" w:type="dxa"/></w:tblCellMar>');
    expect(rec).not.toContain("<w:jc");
    expect(pr.indexOf("<w:tblCellMar")).toBeLessThan(pr.indexOf("<w:tblPrChange"));
  });

  /**
   * #214 那三张待改表在**文件层**分得开、在**纸上**分不开的那一本账。
   *
   * 装机腿（`.scratch/table-props-write-install-check.sh`）能证明「点确定之后纸挪了、盘上也落了」，但它量不到
   * 一件东西：「填 0」和「原样不填」画出来是同一处（css 一字不差），只有 `w:tblPr` 里那一枚在不在分得开——
   * 而这一枚是授权口径里真实存在的一句话（46 份里 `w:tblInd` 说 0 的有 30 枚，见 `.scratch/tbl-place-audience.py`）。
   * 所以这一条留在模块测里钉，形状与装机件同取 `EDIT_TABLES`，不另拼一份。
   */
  it("「说了 0」与「没说」画到同一处、盘上是两句话：0 落明写的 `w:w=\"0\"`，没点名的那一枚不许被发明", () => {
    const doc = open(makeTablePlaceDocx(EDIT_TABLES));
    const [s1, s2, s3] = tablesAt(doc);
    expect([s1.align, s1.indentTwips, s2.align, s2.indentTwips, s3.align, s3.indentTwips]).toEqual([
      undefined,
      undefined,
      "center",
      undefined,
      "center",
      288,
    ]);
    const s2Before = s2.tblPrXml!;

    syncTableProps(s1, { align: "center", indentTwips: 283 }); // 「居中」+ 左缩进 0.5 厘米 = 283 缇
    syncTableProps(s2, { align: "center" }); // 原样确定
    syncTableProps(s3, { indentTwips: 0 }); // 左缩进填 0
    const xml = save(doc);

    // 摆1：凭空多出两枚，且各按 `CT_TblPr` 的孩子序坐在 `tblW` 与 `tblLayout` 之间。
    const pr1 = prOf(xml, 0);
    expect(childOf(pr1, "jc")).toBe('<w:jc w:val="center"/>');
    expect(childOf(pr1, "tblInd")).toBe('<w:tblInd w:w="283" w:type="dxa"/>');
    expect(pr1.indexOf("<w:tblW")).toBeLessThan(pr1.indexOf("<w:jc"));
    expect(pr1.indexOf("<w:jc")).toBeLessThan(pr1.indexOf("<w:tblInd"));
    expect(pr1.indexOf("<w:tblInd")).toBeLessThan(pr1.indexOf("<w:tblLayout"));
    // 纸上走带号的那一支：居中那一截 + 那枚缇（283 ÷ 15）。
    expect(cssAt(doc, 0)).toContain("margin-left:calc(50% - 166.66666666666666px + 18.866666666666667px);margin-right:auto;");

    // 摆2：没点名的缩进不许被发明成 0——整枚 `w:tblPr` 逐字节留在原地。
    expect(prOf(xml, 1)).toBe(s2Before);
    expect(childOf(prOf(xml, 1), "tblInd")).toBe("");

    // 摆3：落的是**明写的 0**，不是删掉；删掉就与摆2 在文件里同一长相了。
    expect(childOf(prOf(xml, 2), "tblInd")).toBe('<w:tblInd w:w="0" w:type="dxa"/>');
    expect(prOf(xml, 2)).not.toBe(prOf(xml, 1));
    // 而纸上两张画到同一处：这一本账只有文件层分得开。
    expect(cssAt(doc, 2)).toBe(cssAt(doc, 1));
    expect(cssAt(doc, 1)).toContain("margin-left:calc(50% - 166.66666666666666px);margin-right:auto;");
    expect(cssAt(doc, 1)).not.toContain(" + ");
  });
});

/** 第 n 枚 `w:trPr`（一枚 `w:trPrChange` 里装的是行的孩子、不会再套一枚 `w:trPr`，所以非贪婪就到本枚的尾）。 */
const trPrOf = (xml: string, nth = 0): string => (xml.match(/<w:trPr\b[\s\S]*?<\/w:trPr>/g) ?? [])[nth] ?? "";
/** 这一行自己的那一份属性（`w:trPr` 的每枚孩子在读侧就并到 `RowInfo` 上）。 */
const rowInfoAt = (doc: WordDocument, n: number): RowInfo | undefined => tableOf(doc).rows[n]?.[0]?.row;
const trCssAt = (doc: WordDocument, n: number): string =>
  (renderWord(doc).match(/<tr\b[^>]*>/g) ?? []).map((t) => / style="([^"]*)"/.exec(t)?.[1] ?? "")[n] ?? "";

/**
 * 「表格属性」第二片：行自己那句 `w:trPr/w:jc`（§17.4.75）。读、画、存三腿在 #29 就接上了，这一片补的是落笔处。
 *
 * 摆的是数出来的三种形状（`.scratch/row-jc-shape-audience.py`：授权 46 件里 190 枚说的全是
 * `<w:jc w:val="center"/>`，孩子序只有 `cantSplit+trHeight+jc`(18)、`trHeight+jc`(154)、`jc`(18) 三种，
 * `jc` 永远收尾；`left` 一枚也没明写过）。WPS 那一格摆在哪张对话框**未取证**，所以这里不拿它当分组依据，
 * 只钉「件里说过的那一句改得动、存得下、读得回来，且只换它自己那一枚」。
 */
describe("docx 表格属性：行那一句 w:trPr/w:jc 的落笔处", () => {
  /** 四行按件里的三种形状各摆一行，再加一行整枚 `w:trPr` 都没写的。 */
  const ROWS = pack(
    `<w:tbl><w:tblPr><w:tblStyle w:val="a3"/><w:tblW w:w="0" w:type="auto"/><w:tblLook w:val="04A0"/></w:tblPr>` +
      `<w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid>` +
      `<w:tr><w:trPr><w:cantSplit/><w:trHeight w:val="600"/><w:jc w:val="center"/></w:trPr>` +
      `<w:tc><w:p><w:r><w:t>一</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/><w:jc w:val="center"/></w:trPr>` +
      `<w:tc><w:p><w:r><w:t>二</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>` +
      `<w:tc><w:p><w:r><w:t>三</w:t></w:r></w:p></w:tc></w:tr>` +
      `<w:tr><w:tc><w:p><w:r><w:t>四</w:t></w:r></w:p></w:tc></w:tr>` +
      `</w:tbl>`,
  );

  it("只换点过名的那一枚：兄弟孩子逐字节不动，`CT_TrPr` 的序照旧", () => {
    const doc = open(ROWS);
    expect(rowInfoAt(doc, 0)).toMatchObject({ keepTogether: true, heightTwips: 600, jc: "center" });
    syncRowAlign(rowInfoAt(doc, 0)!, "right");
    expect(trPrOf(save(doc), 0)).toBe('<w:trPr><w:cantSplit/><w:trHeight w:val="600"/><w:jc w:val="right"/></w:trPr>');
    // 正控制：没被点名的三行留在原地（同一趟保存不该顺手改别处的行）。
    expect(trPrOf(save(doc), 1)).toBe('<w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/><w:jc w:val="center"/></w:trPr>');
    expect(trPrOf(save(doc), 2)).toBe('<w:trPr><w:jc w:val="center"/></w:trPr>');
  });

  it("「左对齐」= 摘掉那一句，不是补一句 `val=\"left\"`：这一族说左的方式就是什么都不写", () => {
    const doc = open(ROWS);
    const info = rowInfoAt(doc, 1)!;
    syncRowAlign(info, "");
    expect(info.jc).toBeUndefined();
    // 字段与字节同一步退：`rowShapeOf` 是记录里的「改之前」，留着 `jc` 就是拒绝时把没说过的一句话摆回去。
    expect(rowShapeOf(info)).toEqual({ heightTwips: 600, heightRule: "atLeast" });
    const out = trPrOf(save(doc), 1);
    expect(out).toBe('<w:trPr><w:trHeight w:val="600" w:hRule="atLeast"/></w:trPr>');
    expect(out).not.toContain("left");
  });

  it("行说过别的话但没说这一句时，新的一枚按 `CT_TrPr` 的序坐在 `hidden` 之前", () => {
    const doc = open(ROWS);
    const info: RowInfo = { header: true, hidden: true };
    doc.blocks.filter((b): b is Table => b.kind === "table")[0]!.rows[0]!.forEach((cell) => {
      cell.row = info;
    });
    info.slots = [
      { tag: "tblHeader", xml: '<w:tblHeader w:val="1"/>' },
      { tag: "hidden", xml: "<w:hidden/>" },
    ];
    syncRowAlign(info, "center");
    expect(info.slots?.map((s) => s.tag)).toEqual(["tblHeader", "jc", "hidden"]);
    expect(trPrOf(save(doc), 0)).toBe('<w:trPr><w:tblHeader w:val="1"/><w:jc w:val="center"/><w:hidden/></w:trPr>');
  });

  it("整枚 `w:trPr` 都没有的行（编辑器造的）：字段就够，写出器按模型把这一句拼出来", () => {
    const doc = open(ROWS);
    expect(rowInfoAt(doc, 3)).toBeUndefined(); // 读侧不给没说的行造一份属性
    const table = tableOf(doc);
    const info: RowInfo = {};
    for (const cell of table.rows[3]!) cell.row = info;
    syncRowAlign(info, "center");
    expect(trPrOf(save(doc), 3)).toBe('<w:trPr><w:jc w:val="center"/></w:trPr>');
  });

  it("存了再开还是那一句，第二次存同一串字节：不双写、不漂", () => {
    const doc = open(ROWS);
    syncRowAlign(rowInfoAt(doc, 2)!, "right");
    const once = writeDocx(doc);
    const xml = save(doc);
    expect(xml.match(/<w:trPr><w:jc/g) ?? []).toHaveLength(1);
    const again = open(once);
    expect(rowInfoAt(again, 2)?.jc).toBe("right");
    const twice = text(OpcPackage.from(writeDocx(again)).bytes("word/document.xml")!);
    expect(twice).toBe(xml);
  });

  it("画得出来：行那一句坐在 `<tr>` 上，段自己说了话的那一处按 §17.3.1.8 盖过它", () => {
    const doc = open(ROWS);
    expect(trCssAt(doc, 0)).toContain("text-align:center;");
    // 行高与行对齐落在同一枚 `<tr>` 上：600 缇 ÷ 15 = 40px。
    expect(trCssAt(doc, 1)).toContain("height:40px;text-align:center;");
    expect(trCssAt(doc, 3)).toBe("");
    // 段说了「左」的那一格不再跟着行居中：优先级是段 > 格 > 行，行那句只在段没说时借过去。
    const stated = pack(
      `<w:tbl><w:tblGrid><w:gridCol w:w="1800"/></w:tblGrid>` +
        `<w:tr><w:trPr><w:jc w:val="center"/></w:trPr>` +
        `<w:tc><w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>甲</w:t></w:r></w:p></w:tc></w:tr>` +
        `</w:tbl>`,
    );
    const html = renderWord(open(stated));
    expect(/<tr\b[^>]*>/.exec(html)?.[0]).toContain("text-align:center;");
    // 段自己那一枚 `<div class="para">` 上写的是 `text-align:left`——纸上是两处不同的元素，行那句没有盖住它。
    expect(/<div class="para"[^>]*>/.exec(html)?.[0]).toContain("text-align:left");
  });

  it("开着修订：那一枚 `w:trPrChange` 带的是改之前那一整套，拒绝把件里那句放回原位", () => {
    const doc = open(ROWS);
    const info = rowInfoAt(doc, 0)!;
    const was = rowShapeOf(info);
    const wasSlots = info.slots ?? [];
    syncRowAlign(info, "right");
    signRowRevision(info, { id: 77, author: "顾峰" }, was, wasSlots);
    const bytes = writeDocx(doc);
    const xml = text(OpcPackage.from(bytes).bytes("word/document.xml")!);
    expect(xml).toContain('<w:trPrChange w:id="77" w:author="顾峰">');
    // 记录里是行改之前说过的三句，不是只有这一句——拒绝要摆回的是那一整套。
    expect(/<w:trPrChange[\s\S]*?<\/w:trPrChange>/.exec(xml)?.[0]).toContain('<w:jc w:val="center"/>');

    const reopened = open(bytes);
    expect(reopened.resolveRevisions("reject", 77)).toBe(1);
    expect(rowInfoAt(reopened, 0)).toMatchObject({ keepTogether: true, heightTwips: 600, jc: "center" });
    const back = save(reopened);
    expect(back).not.toContain("<w:trPrChange");
    expect(trPrOf(back, 0)).toBe('<w:trPr><w:cantSplit/><w:trHeight w:val="600"/><w:jc w:val="center"/></w:trPr>');
  });
});
