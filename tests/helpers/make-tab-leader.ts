import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_SET = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

/** 正文区 8306 缇宽（11906 − 1800 − 1800）；下面每一枚落点都按这一档算：720=48px、1440=96、2880=192、4153≈276.9、8306≈553.7。 */
const ST = (val: string, pos: number, leader?: string) =>
  `<w:tab w:val="${val}" w:pos="${pos}"${leader ? ` w:leader="${leader}"` : ""}/>`;

/**
 * 制表位前导符 `w:leader`（§17.3.2.27 ST_LeaderType）的夹具：五段各摆一处「墨该不该上」分得开的形状。
 * 授权 46 份带 `w:leader` 的 94 枚枚枚说 `dot`（`.scratch/tabs-leader-roundtrip.ts`），其余四档是本机造的
 * 形状，为的是让「token 只跟着这一枚真正停进去的那一格」与「`spaces` 有意留白」两条各有一枚可读的样本。
 *
 * - 甲：唯一一枚位是 `right 8306 dot` ⇒ 这一枚停进去、带点。
 * - 乙：两枚位、两枚制表符。第一枚停 `left 2880 dot`（带点），第二枚停 `right 8306`（那句没说前导符）
 *   ⇒ 同段两枚墨不一样，钉住 token 不是「一段一份」而是「一格一份」。
 * - 丙：段自己什么都不说，`w:tabs` 在样式 `LD` 上（`left 1440 dot`）⇒ 借来的位也带墨，走的是 `mergeTabStops`。
 * - 丁：`left 4153 spaces` ⇒ token 挂着、墨一条不上（`spaces` 是文件说了「留白」）。
 * - 戊：整段没有 `w:tabs`，制表符按 `w:defaultTabStop` 的 720 走 ⇒ 文件本就没说前导符，不许替它铺点。
 * - 尾段：一枚制表符也没有，用来数「不该挂的枚数」。
 */
const BODY =
  `<w:p><w:pPr><w:tabs>${ST("right", 8306, "dot")}</w:tabs></w:pPr>` +
  `<w:r><w:t>甲线 姓名</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>13800</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:tabs>${ST("left", 2880, "dot")}${ST("right", 8306)}</w:tabs></w:pPr>` +
  `<w:r><w:t>乙线</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>第一枚停2880</w:t></w:r>` +
  `<w:r><w:tab/></w:r><w:r><w:t>末</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="LD"/></w:pPr>` +
  `<w:r><w:t>丙线</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>样式借来的位</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:tabs>${ST("left", 4153, "spaces")}</w:tabs></w:pPr>` +
  `<w:r><w:t>丁线</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>留白那一档</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>戊线</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>按默认步长</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>尾段 没有制表符</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/>` +
  `<w:cols w:space="425"/><w:docGrid w:type="lines" w:linePitch="312"/></w:sectPr>`;

const STYLES =
  `<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Calibri" w:hAnsi="Calibri" w:eastAsia="宋体"/><w:sz w:val="24"/><w:szCs w:val="24"/>` +
  `</w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:default="1" w:styleId="1"><w:name w:val="Normal"/><w:qFormat/></w:style>` +
  `<w:style w:type="paragraph" w:styleId="LD"><w:name w:val="带前导符的样式"/><w:qFormat/>` +
  `<w:pPr><w:tabs>${ST("left", 1440, "dot")}</w:tabs></w:pPr></w:style></w:styles>`;

const SETTINGS = `<w:settings xmlns:w="${NS_W}"><w:defaultTabStop w:val="720"/></w:settings>`;

export function makeTabLeaderDocx(): Uint8Array {
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="${CT_STY}"/><Override PartName="/word/settings.xml" ContentType="${CT_SET}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="${NS_R}/settings" Target="word/settings.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${BODY}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
      "word/settings.xml": strToU8(XML_DECL + SETTINGS),
    },
    { level: 6 },
  );
}
