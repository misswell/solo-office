import { strToU8, zipSync } from "fflate";
import { XML_DECL, tableXml } from "./make-table-place.js";

/**
 * #219 的件：三张各说一种左缩进的表，外加一张一句都没说的对照表。
 *
 * 这一族问的不是「表摆在哪」（#213/#214 管纸面），是「表格属性那一格在**每一种测量单位**下摆什么字面、
 * 填进去的数按哪一种缇数落笔」：三种单位的小数位与每单位缇数在 `doc-units.ts` 里各有各的口径，
 * 厘米那一档 §155 已经走过，这里补毫米 / 英寸 / 磅三档。
 *
 * 三枚起始值刻意挑成一路要变动的量级（288 → 20 毫米、−129 → 0.75 英寸、907 → 36 磅），
 * 落笔之后每张表各自挪几十像素：换错了单位，位移对不上。
 * 对照表（无缩进表）一句 `w:tblInd` 都没写：那一格必须**空着**（「没说」不是 0），按「取消」一个字节都不许动。
 *
 * 形状只有一个主人：装机烟测（`.scratch/make-table-units-smoke.ts`）与它的反读脚本都从这里拿。
 * 表本身的形状借 #213 的 `tableXml`（`w:tblW` 与网格同和、`w:tblLayout=fixed`，与授权 58/58 同形）。
 */
export const UNIT_TABLES = {
  /** 毫米：起始 288 缇，填 20 毫米 → 1134 缇。 */
  mm: { name: "毫米表", startTwips: 288, literal: "20" },
  /** 英寸：起始 −129 缇（授权件里那四枚负缩进之一档），填 0.75 英寸 → 1080 缇。 */
  in: { name: "英寸表", startTwips: -129, literal: "0.75" },
  /** 磅：起始 907 缇，填 36 磅 → 720 缇。 */
  pt: { name: "磅表", startTwips: 907, literal: "36" },
} as const;

/** 一张一句 `w:tblInd` 都没说的对照表。 */
export const PLAIN_NAME = "无缩进表";

const p = (t: string): string => `<w:p><w:r><w:t>${t}</w:t></w:r></w:p>`;
const indent = (twips: number): string => `<w:tblInd w:w="${twips}" w:type="dxa"/>`;

export function makeTableUnitsDocx(): Uint8Array {
  const tables =
    p("单位表之前一段") +
    tableXml(UNIT_TABLES.mm.name, 5000, indent(UNIT_TABLES.mm.startTwips)) +
    p("毫米与英寸之间") +
    tableXml(UNIT_TABLES.in.name, 5000, indent(UNIT_TABLES.in.startTwips)) +
    p("英寸与磅之间") +
    tableXml(UNIT_TABLES.pt.name, 5000, indent(UNIT_TABLES.pt.startTwips)) +
    p("磅与对照之间") +
    tableXml(PLAIN_NAME, 5000, "") +
    p("单位表之后一段");
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
          `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
          `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><w:body>${tables}` +
          `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/>` +
          `<w:pgMar w:top="1440" w:right="1474" w:bottom="1440" w:left="1588" w:header="720" w:footer="720" w:gutter="0"/></w:sectPr></w:body></w:document>`,
      ),
    },
    { level: 6 },
  );
}
