import { strToU8, zipSync } from "fflate";

export const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_FTR = "application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * A footer whose tab characters have somewhere to land — the band side of the tab-stops family (#235).
 *
 * 受众先数（`.scratch/band-tab-chars-audience.py`，与 §153 同一只 46 件口径）：带子里真正的 `w:r/w:tab` 只有
 * **2 枚**，去重后是 **1 份件、1 枚**——`event plan.dotm` 页脚段#0：一枚裸制表符、无文字，段自己不写位，
 * 位全从 `footer` 样式（styleId 4）借：`center 4153` + `right 8306`（字面照抄那份件）。**两枚都不带
 * `w:leader`**，所以带子侧的前导符墨在语料里 0 受众；② 格里那枚 `dot` 钉的是机制（token 只跟这一枚真正
 * 停进去的格），不是语料说过的形状——与 `make-band-place.ts` ③ 格「自己说的 right」同一条注法。
 *
 * 每段各占一格：① 链上借位 + 裸制表符（唯一真实受众的形状）；② 段自己说位带前导符、制表符两侧有字
 * （量「跟着这一截文字」那条路）；③ 段自己的位顶掉链上同位置的（§17.7.2 逐位置）；④ 链上 `clear` 花掉一格；
 * ⑤ 无制表符的对照（一个属性都不该多）。
 */
export const TABS_FOOTER =
  `${XML_DECL}<w:ftr xmlns:w="${NS_W}" xmlns:r="${NS_REL}">` +
  `<w:p><w:pPr><w:pStyle w:val="BandFooter"/></w:pPr><w:r><w:tab/></w:r></w:p>` +
  `<w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="2880" w:leader="dot"/></w:tabs></w:pPr>` +
  `<w:r><w:t>前</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>后</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandFooter"/><w:tabs><w:tab w:val="left" w:pos="4153"/></w:tabs></w:pPr>` +
  `<w:r><w:tab/></w:r><w:r><w:t>盖</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="BandFooter"/><w:tabs><w:tab w:val="clear" w:pos="4153"/></w:tabs></w:pPr>` +
  `<w:r><w:tab/></w:r><w:r><w:t>清</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:jc w:val="left"/></w:pPr><w:r><w:t>没有制表符</w:t></w:r></w:p>` +
  `</w:ftr>`;

/**
 * The one rung under the band paragraphs: a `footer` style carrying the very pair `event plan.dotm` writes,
 * so ① and ④ borrow their stops from the chain the way that file's paragraph does.
 */
export const TABS_STYLES =
  `${XML_DECL}<w:styles xmlns:w="${NS_W}">` +
  `<w:docDefaults><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="BandFooter"><w:name w:val="band footer"/><w:pPr>` +
  `<w:tabs><w:tab w:val="center" w:pos="4153"/><w:tab w:val="right" w:pos="8306"/></w:tabs>` +
  `</w:pPr></w:style>` +
  `</w:styles>`;

/** The body's section, pointing at that footer, in the order the authorized templates write the reference.
 * 正文 60 行把纸撑到多页（与 .scratch/band-tabs-dom.ts 同一做法）：页脚原件守最后一张，前面的每一张挂复件，
 * 「复件由补跑那一趟接管」这一句才有断言可站——单页件根本不画复件。 */
const BODY_LINES = Array.from({ length: 60 }, (_, i) => `<w:p><w:r><w:t>分页用第 ${i + 1} 行</w:t></w:r></w:p>`).join("");
const DOCUMENT =
  `${XML_DECL}<w:document xmlns:w="${NS_W}" xmlns:r="${NS_REL}"><w:body>` +
  `${BODY_LINES}` +
  `<w:sectPr><w:footerReference r:id="rId5" w:type="default"/>` +
  `<w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440" w:header="720" w:footer="720"/>` +
  `</w:sectPr></w:body></w:document>`;

/** A file whose only band is {@link TABS_FOOTER}, with {@link TABS_STYLES} behind it. */
export function makeBandTabsDocx(): Uint8Array {
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
      "word/styles.xml": strToU8(TABS_STYLES),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId5" Type="${NS_REL}/footer" Target="footer1.xml"/></Relationships>`,
      ),
      "word/footer1.xml": strToU8(TABS_FOOTER),
    },
    { level: 6 },
  );
}
