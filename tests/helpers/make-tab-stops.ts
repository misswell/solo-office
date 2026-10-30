import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_SET = "application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml";

const ST = (val: string, pos: number, leader?: string): string =>
  `<w:tab w:val="${val}" w:pos="${pos}"${leader ? ` w:leader="${leader}"` : ""}/>`;
const TABS = (...ss: string[]): string => `<w:tabs>${ss.join("")}</w:tabs>`;

/**
 * #234「制表位」对话框那一整条路共用的夹具：五种「清单该摆什么」分得开的起始态各一段。
 * 单测（`tests/doc-tab-stops.test.ts`）、真浏览器那 35 例（`.scratch/tab-stops-dom.ts`）与装机壳那一腿
 * （`src/main/main.ts` 的 `tabStopsLeg` + `.scratch/tab-stops-install-check.sh`）读的是**同一份字节**，
 * 所以纸上的落点、盘上的字节与对话框里的清单三处对得上。
 *
 * - 甲：段自己写两枚 `left 720` + `right 8306 dot` ⇒ 起手清单两行，逐行读位置、档与前导符。
 * - 乙：段只挂样式 `TB`，两枚从样式链借来（`left 1440` + `right 8306`）⇒ 清单摆的是合并过的那一份，
 *   「原样确定」不许把它抄到段身上（段本来一个字没说）。
 * - 丙：**连 `w:pPr` 都没有** ⇒ 清单空着、只有字段那条路能写；纸上那一枚落在 `w:defaultTabStop` 那一族。
 * - 丁／戊：两份说法不一致的清单 ⇒ 跨段那一问（「多段不一致」那行小字 + 逐位置派发，不整份盖掉）。
 *
 * `w:defaultTabStop` 故意写 **420** 而不是 720：丙段那一枚原本落在 420 的整数格上（两枚汉字已越过第一格），
 * 加上清单里的 720 才走到 48px——两格相同就永远分不开「画的是清单说的」还是「默认格说的」。
 * 它夹在 `w:zoom` 与 `w:hyphenationZone` 中间：改这一枚时那两枚的字节与次序都得留在原位。
 */
const BODY =
  `<w:p><w:pPr>${TABS(ST("left", 720), ST("right", 8306, "dot"))}</w:pPr>` +
  `<w:r><w:t>甲起</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>甲落</w:t></w:r></w:p>` +
  `<w:p><w:pPr><w:pStyle w:val="TB"/></w:pPr>` +
  `<w:r><w:t>乙起</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>乙落</w:t></w:r></w:p>` +
  `<w:p><w:r><w:t>丙起</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>丙落</w:t></w:r></w:p>` +
  `<w:p><w:pPr>${TABS(ST("left", 720))}</w:pPr>` +
  `<w:r><w:t>丁起</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>丁落</w:t></w:r></w:p>` +
  `<w:p><w:pPr>${TABS(ST("left", 1440), ST("right", 4320))}</w:pPr>` +
  `<w:r><w:t>戊起</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>戊落</w:t></w:r></w:p>` +
  `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;

const STYLES =
  `<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr>` +
  `<w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/><w:sz w:val="22"/>` +
  `</w:rPr></w:rPrDefault></w:docDefaults>` +
  `<w:style w:type="paragraph" w:styleId="TB"><w:name w:val="借来的"/><w:pPr>` +
  `${TABS(ST("left", 1440), ST("right", 8306))}</w:pPr></w:style></w:styles>`;

const SETTINGS = `<w:settings xmlns:w="${NS_W}"><w:zoom w:percent="100"/><w:defaultTabStop w:val="420"/><w:hyphenationZone w:val="420"/></w:settings>`;

export function makeTabStopsDocx(): Uint8Array {
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="${CT_STY}"/><Override PartName="/word/settings.xml" ContentType="${CT_SET}"/></Types>`,
      ),
      // 件级的 `.rels` 只挂正文本体；settings 挂在正文那一页的关系表上（授权 46 份份份如此，件级挂它的 0 份：
      // `.scratch/settings-rel-census.py`）。挂错地方，保存默认制表位时就会发出第二张指向同一部件的关系。
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/><Relationship Id="rId2" Type="${NS_R}/settings" Target="settings.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${BODY}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
      "word/settings.xml": strToU8(XML_DECL + SETTINGS),
    },
    { level: 6 },
  );
}
