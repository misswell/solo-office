import { strToU8, zipSync } from "fflate";
// The page field comes from the fixture that already spells it the way the files do: an instruction written by
// hand here lost its backslash to the template string, and the band then read as a part no edit may touch.
import { pageFieldRuns } from "./make-header-footer.js";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_FTR = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * A footer that says where each of its paragraphs sits, in the shapes the authorized 46 份 write.
 *
 * Why this exists: 那 46 份的 `word/header*.xml` / `word/footer*.xml` 共 **76 段**，`w:jc` 的三档来源分别是
 * **段自己说 `center` 12 枚**（6 个全路径件 / 3 份模板名）、**样式链借来 52 枚**（`left` 48 / `both` 2 / `center` 2）、
 * **谁都没说 12 段**；另有 **6 段带 `w:pPr/w:framePr`**（4 个全路径件 / 2 份模板名，枚枚都是
 * `w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:y="1"` 加 `w:xAlign` right ×4 / center ×2，坐在带 PAGE 域
 * 的那一段上）（`.scratch/band-para-align-audience.py`）。自家此前把带子里的段落画成清一色左对齐，所以这一族全落在左沿。
 *
 * 每段各占一格，让「谁说的」在断言里分得开：自己说的、链上借的、自己说的盖住链上说的、框摆到右、框摆到中、
 * 框锚在纸上（0 受众，故意不画）、以及什么都没说的对照组。
 *
 * ③ 那一格的 `right` 不是从 46 份的带子里读到的（那里的段自己只说 `center`），它是自家 `STYLED_HEADER` 那一族
 * 早就写着的形状：留这一格是让「自己说的盖住样式说的」这条优先级有断言可站，不代表语料说过 right。
 *
 * 后三格是缩进那一族，各有出处（`.scratch/band-ind-tabs-audience.py`）：段自己说的 `w:ind` **4 枚，枚枚都是
 * `w:right="360"`**（2 个全路径件 / **1 份**模板名 `office rental contract.dotm`）；链上借来的 `w:ind` 落在 **8 段**
 * 上，只说首行（`w:firstLine="422" w:firstLineChars="200"` ×6、`w:firstLine="883" w:firstLineChars="200"` ×2），
 * **枚枚不带 `w:left`/`w:right`**——所以「借来的左右缩进」这一格没有语料，画它等于编它。三格因此只钉住已量到的
 * 两个形状（自己说的 right、链上说的首行）加一条优先级（自己的 `w:firstLineChars="0"` 顶掉链上的首行）。
 * 同一轮重数的 `w:tabs`（自己说 12 段 / 20 枚定义，链上借 58 段各 2 枚，但带子里真正的制表符只有 **2 枚**）不进
 * 这些格：带子的段是 `white-space:pre-wrap` 下的裸 `\t`，给它定位要先改 DOM 结构，动到 `textContent` 读回与可编辑
 * 克隆的光标，是另一笔账。
 */
export const PLACE_FOOTER =
  `${XML_DECL}<w:ftr xmlns:w="${NS_W}" xmlns:r="${NS_REL}">` +
  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:t>自己说居中</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandDeep"/></w:pPr><w:r><w:t>链上借来居中</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandLeft"/><w:jc w:val="right"/></w:pPr><w:r><w:t>自己说的盖住样式说的</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandLeft"/><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="right" w:y="1"/></w:pPr>` +
  `${pageFieldRuns(undefined, "7")}</w:p>` +
  `<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="text" w:hAnchor="margin" w:xAlign="center" w:y="1"/></w:pPr>` +
  `<w:r><w:t>摆到中沿</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:framePr w:wrap="around" w:vAnchor="page" w:hAnchor="page" w:xAlign="right" w:y="1"/></w:pPr>` +
  `<w:r><w:t>锚在纸上</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:framePr w:wrap="around" w:hAnchor="text" w:x="108" w:y="1"/></w:pPr>` +
  `<w:r><w:t>只说偏移不说对齐</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandQuiet"/><w:rPr><w:sz w:val="18"/></w:rPr></w:pPr><w:r><w:t>什么都没说</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:ind w:right="360"/></w:pPr><w:r><w:t>自己说右缩进</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandFirst"/></w:pPr><w:r><w:t>链上借来首行缩进</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandFirst"/><w:ind w:firstLineChars="0"/></w:pPr><w:r><w:t>自己说的不缩首行</w:t></w:r></w:p>` +
  `</w:ftr>`;

/**
 * The two rungs §17.7.2 puts under a band paragraph: the style it names, that style's `w:basedOn`, and the default
 * paragraph style beneath the whole chain. `BandDeep` says nothing of its own and borrows `center` from `BandLent`;
 * `BandLeft` says `left`; `BandQuiet` and the default `Normal` say nothing, so the last paragraph has no answer
 * to draw and the page must invent none.
 */
export const PLACE_STYLES =
  `${XML_DECL}<w:styles xmlns:w="${NS_W}">` +
  `<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BandLent"><w:name w:val="band lend"/><w:pPr><w:jc w:val="center"/></w:pPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BandDeep"><w:name w:val="band deep"/><w:basedOn w:val="BandLent"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BandLeft"><w:name w:val="band left"/><w:pPr><w:jc w:val="left"/></w:pPr></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BandQuiet"><w:name w:val="band quiet"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BandFirst"><w:name w:val="band first"/><w:pPr><w:ind w:firstLine="422" w:firstLineChars="200"/></w:pPr></w:style>` +
  `</w:styles>`;

/** The body's section, pointing at that footer, in the order the authorized templates write the reference. */
const DOCUMENT =
  `${XML_DECL}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_REL}"><w:body>` +
  `<w:p><w:r><w:t>正文</w:t></w:r></w:p>` +
  `<w:sectPr><w:footerReference r:id="rId5" w:type="default"/>` +
  `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>` +
  `</w:sectPr></w:body></w:document>`;

/** A file whose only band is {@link PLACE_FOOTER}, with {@link PLACE_STYLES} behind it. */
export function makeBandPlaceDocx(): Uint8Array {
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
          `<Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="${CT_W}"/>` +
          `<Override PartName="/word/styles.xml" ContentType="${CT_STY}"/>` +
          `<Override PartName="/word/footer1.xml" ContentType="${CT_FTR}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_REL}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(DOCUMENT),
      "word/styles.xml": strToU8(PLACE_STYLES),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_REL}/footer" Target="footer1.xml"/></Relationships>`,
      ),
      "word/footer1.xml": strToU8(PLACE_FOOTER),
    },
    { level: 6 },
  );
}
