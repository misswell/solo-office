import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { rewriteTocCachedPages, tocPageRefs } from "../src/core/doc-toc.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeTocDocx } from "./helpers/make-carry-toc.js";

/**
 * 「更新目录」把 PAGEREF 的缓存页码按当前分页重算（#248）。
 *
 * 受众与形状照 §176 取证：`event plan.dotm` 的目录条目一段一条，页码是 PAGEREF 复杂域的结果缓存
 * （`PAGEREF _Toc27457 \h` → 结果 `w:t` 里是上次算出的数）。WPS「更新目录」改的正是这截缓存；指令
 * （书签名、`\h` 开关）与三枚 `w:fldChar` 边界一个字不动——更新完仍是一条活的域。
 */

const TOC_XML = (numbers: [string, string][]): string =>
  '<w:sdt><w:sdtContent>' +
  numbers
    .map(
      ([name, n]) =>
        `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:pos="8306"/></w:tabs></w:pPr>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> HYPERLINK \\l ${name} </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t>某一条目</w:t></w:r>` +
        `<w:r><w:tab/></w:r>` +
        `<w:r><w:fldChar w:fldCharType="begin"/></w:r>` +
        `<w:r><w:instrText xml:space="preserve"> PAGEREF ${name} \\h </w:instrText></w:r>` +
        `<w:r><w:fldChar w:fldCharType="separate"/></w:r>` +
        `<w:r><w:t>${n}</w:t></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r>` +
        `<w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>`,
    )
    .join("") +
  '</w:sdtContent></w:sdt>';

/** 目录后面追两处书签：正文里一枚、表格格里一枚（格里的书签归它那一张表那一块）。 */
const BOOKMARKED_BODY =
  `<w:p><w:bookmarkStart w:id="1" w:name="_Toc27457"/><w:r><w:t>一、市场分析</w:t></w:r><w:bookmarkEnd w:id="1"/></w:p>` +
  `<w:tbl><w:tr><w:tc><w:p><w:bookmarkStart w:id="2" w:name="_Toc27458"/><w:r><w:t>格里的标题</w:t></w:r><w:bookmarkEnd w:id="2"/></w:p></w:tc></w:tr></w:tbl>`;

describe("doc-toc：PAGEREF 缓存的读与改", () => {
  it("读得出每条的域名与缓存数字（HYPERLINK 那条不混进来）", () => {
    const refs = tocPageRefs(TOC_XML([["_Toc1", "1"], ["_Toc2", "2"]]));
    expect(refs.map((r) => [r.name, r.numberText])).toEqual([["_Toc1", "1"], ["_Toc2", "2"]]);
  });

  it("按解析器换数：只改结果、指令与边界一字不动", () => {
    const xml = TOC_XML([["_Toc1", "1"], ["_Toc2", "2"]]);
    const out = rewriteTocCachedPages(xml, (name) => (name === "_Toc1" ? 7 : 9));
    expect(out.changed).toBe(2);
    expect(out.xml).toContain("PAGEREF _Toc1 \\h ");
    expect(out.xml).toContain("<w:t>7</w:t>");
    expect(out.xml).toContain("<w:t>9</w:t>");
    // 指令、边界、HYPERLINK 都不动
    expect(out.xml.match(/fldChar w:fldCharType="begin"/g)?.length).toBe(4);
    expect(out.xml.match(/fldChar w:fldCharType="end"/g)?.length).toBe(4);
    expect(out.xml).toContain("HYPERLINK \\l _Toc1");
  });

  it("解析器说不出页码的条目原样留着；已经对的条目不算改", () => {
    const xml = TOC_XML([["_Toc1", "3"], ["_Toc2", "2"]]);
    const out = rewriteTocCachedPages(xml, (name) => (name === "_Toc1" ? 3 : undefined));
    expect(out.changed).toBe(0);
    expect(out.xml).toBe(xml);
  });
});

describe("docx：整块更新目录（写回字节）", () => {
  it("书签落在哪一块报得出来（正文里一枚、表格格里一枚）", () => {
    const doc = WordDocument.openBytes(makeTocDocx(BOOKMARKED_BODY));
    const where = doc.bookmarkBlocks();
    // 块 0 是目录那个 sdt；书签段是块 1；书签所在的表是块 2（格里的书签归它那一张表）。
    expect(where.get("_Toc27457")).toBe(1);
    expect(where.get("_Toc27458")).toBe(2);
    expect(where.size).toBe(2);
  });

  it("解析器说不出位置的块交 undefined：同名只认先出现的那一枚", () => {
    const doc = WordDocument.openBytes(makeTocDocx(BOOKMARKED_BODY));
    expect(doc.updateToc(() => undefined)).toBe(0);
    expect(doc.updateToc((name) => (name === "_Toc27457" ? 568 : undefined))).toBe(1);
  });

  it("更新目录改的是件里那截缓存：写回的 sdt 里数字变了、指令没变", () => {
    const doc = WordDocument.openBytes(makeTocDocx());
    // 夹具两条条目里只有第一条带 PAGEREF 域（第二条是裸数字，照 §176 的「没有域的条目不猜」留着）。
    const changed = doc.updateToc(() => 9);
    expect(changed).toBe(1);
    const saved = text(OpcPackage.from(writeDocx(doc)).bytes("word/document.xml")!);
    expect(saved).toContain("PAGEREF _Toc27457 \\h ");
    expect(saved).toContain("<w:t>9</w:t>");
    expect((saved.match(/<w:t>9<\/w:t>/g) ?? []).length).toBe(1);
    // 条目字还在、制表位还在；第二条那条裸数字原样
    expect(saved).toContain("1、市场分析");
    expect(saved).toContain('<w:tab w:val="right" w:leader="dot" w:pos="8306"/>');
    expect(saved).toContain("<w:t>2</w:t>");
  });

  it("解析器交 undefined 时一个字节都不动", () => {
    const doc = WordDocument.openBytes(makeTocDocx());
    const before = writeDocx(doc);
    expect(doc.updateToc(() => undefined)).toBe(0);
    const after = writeDocx(doc);
    expect(after).toEqual(before);
  });
});
