import { describe, expect, it } from "vitest";
import { WordDocument, signTableRevision, syncTableProps, tableShapeOf } from "../src/core/docx.js";
import type { Table } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderWord } from "../src/core/render.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { pack } from "./helpers/make-containers.js";

const part = (bytes: Uint8Array): string => text(OpcPackage.from(bytes).bytes("word/document.xml")!);
const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const save = (doc: WordDocument): string => part(writeDocx(doc));
/**
 * 第 n 枚 `w:tblPr` 连同它自己的嵌套孩子——一枚待裁决的 `w:tblPrChange` 里**又有一枚** `w:tblPr`，
 * 所以不能拿「遇到第一个 `</w:tblPr>`」当这枚孩子的结尾（那样记录会被截断，正好把要钉的那一段切掉）。
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
    for (let n = re.exec(xml); n && depth > 0; n = re.exec(xml)) {
      if (n[0].startsWith("</")) depth -= 1;
      else if (!n[0].endsWith("/>")) depth += 1;
      if (depth === 0 && seen++ === nth) return xml.slice(start, re.lastIndex);
    }
  }
  return "";
};
const childOf = (set: string, tag: string): string =>
  new RegExp(`<w:${tag}\\b[^>]*/>|<w:${tag}\\b[^>]*>[\\s\\S]*?</w:${tag}>`).exec(set)?.[0] ?? "";
const tableOf = (doc: WordDocument): Table => doc.blocks.filter((b): b is Table => b.kind === "table")[0]!;
const tableCss = (doc: WordDocument): string => / style="([^"]*)"/.exec(renderWord(doc).match(/<table\b[^>]*>/)?.[0] ?? "")?.[1] ?? "";
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
});
