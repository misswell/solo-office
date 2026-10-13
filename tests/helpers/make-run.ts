import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const NS_V = "urn:schemas-microsoft-com:vml";
const NS_O = "urn:schemas-microsoft-com:office:office";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";

const STYLES_DEFAULT = `<w:styles xmlns:w="${NS_W}"><w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/></w:style></w:styles>`;

const pack = (document: string, styles = STYLES_DEFAULT): Uint8Array =>
  zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/>` +
          `<Relationship Id="rId2" Type="${NS_R}/hyperlink" Target="https://example.com/report" TargetMode="External"/></Relationships>`,
      ),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:mc="${NS_MC}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" ` +
          `xmlns:wps="${NS_WPS}" xmlns:v="${NS_V}" xmlns:o="${NS_O}" xmlns:w14="${NS_W14}"><w:body>${document}</w:body></w:document>`,
      ),
      "word/styles.xml": strToU8(XML_DECL + styles),
    },
    { level: 6 },
  );

const REV = (id: number, author: string): string => `w:id="${id}" w:author="${author}" w:date="2026-03-01T08:00:00Z"`;

/** The children of one `w:rPr`, in the order `CT_RPr` gives them. */
const WHOLE_SET =
  `<w:rStyle w:val="Hyperlink"/><w:rFonts w:hint="eastAsia" w:ascii="Times New Roman" w:eastAsia="宋体" w:cs="Times New Roman"/>` +
  "<w:b/><w:bCs/><w:i w:val=\"0\"/><w:iCs w:val=\"0\"/><w:caps/><w:smallCaps/><w:strike/>" +
  `<w:color w:val="808080" w:themeColor="text1" w:themeShade="BFBFBF"/><w:spacing w:val="20"/><w:w w:val="99"/>` +
  `<w:kern w:val="28"/><w:position w:val="2"/><w:sz w:val="24"/><w:szCs w:val="28"/><w:highlight w:val="yellow"/>` +
  `<w:u w:val="single"/><w:shd w:val="clear" w:color="auto" w:fill="FFF2CC"/><w:vertAlign w:val="baseline"/><w:rtl/>` +
  `<w:lang w:val="en-US" w:eastAsia="zh-CN" w:bidi="ar-SA"/><w14:textFill><w14:solidFill><w14:srgbClr w14:val="C00000"/></w14:solidFill></w14:textFill>`;

/**
 * Runs that say far more about themselves than this model has fields for, in the shapes WPS's own files were
 * measured in (`.scratch/wps-rpr-names.ts` over 86 WPS-written documents, 101825 run property sets):
 * `rFonts` 101175, `szCs` 96165, `sz` 93946, `color` 61554, `lang` 26334, `kern` 18399, `rStyle` 17306,
 * `bCs` 12233, `b` 8889, `spacing` 6666, `vertAlign` 2701, `highlight` 2666, `textFill` 2214, `shd` 1454,
 * `u` 1256, `ligatures` 1066, `i` 971, `iCs` 960, `snapToGrid` 380, `caps` 274, `strike` 38, `dstrike` 2.
 * Nine of those names — bold, italic, underline, strike, size, colour, highlight, font, vertical alignment —
 * are all the model could hold, and rebuilding a run from them deleted the other twenty in every file: 15706
 * `w:rStyle`, 12918 `w:lang`, 10711 `w:kern`, 7657 `w:bCs`, 6658 `w:spacing`, 2110 `w:textFill`, 1188 `w:szCs`.
 *
 * A run can also be built around something the reader has no case for at all — a shape wrapped in
 * `mc:AlternateContent`, an OLE object. Ten such runs were measured across the same 86 files, and a run that
 * reads as nothing is a paragraph a save deletes, so those travel as the bytes they came in with.
 */
export function makeRunPropsDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr>${WHOLE_SET}</w:rPr><w:t>整段自述</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:eastAsia="仿宋" w:cs="仿宋"/><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr><w:t>字号成对</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:sz w:val="24"/></w:rPr><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1"><wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>-24130</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>432435</wp:posOffset></wp:positionV><wp:extent cx="5312410" cy="0"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/><wp:docPr id="1" name="直接连接符 1"/><wp:cNvGraphicFramePr/><a:graphic><a:graphicData uri="${NS_WPS}"><wps:wsp><wps:cNvCnPr/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5312410" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></wps:spPr><wps:bodyPr/></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict><v:line id="_x0000_s1026" o:spid="_x0000_s1026" coordsize="21600,21600"><v:stroke weight="1pt" color="#808080"/><v:imagedata o:title=""/></v:line></w:pict></mc:Fallback></mc:AlternateContent></w:r><w:r><w:t>带形状的段</w:t></w:r></w:p>
<w:p><w:r><w:rPr/><w:br w:type="textWrapping" w:clear="all"/></w:r><w:r><w:t>空的性质</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/><w:kern w:val="28"/><w:sz w:val="28"/><w:szCs w:val="28"/><w:lang w:val="en-US" w:eastAsia="zh-CN"/><w:rPrChange ${REV(11, "顾峰")}><w:rPr><w:rFonts w:ascii="Calibri" w:eastAsia="宋体"/><w:kern w:val="28"/><w:sz w:val="22"/><w:szCs w:val="22"/><w:highlight w:val="green"/></w:rPr></w:rPrChange></w:rPr><w:t>留痕的格式</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/><w:lang w:val="en-US"/></w:rPr><w:t>两截同式</w:t></w:r><w:r><w:rPr><w:b/><w:kern w:val="28"/></w:rPr><w:t>各写各的</w:t></w:r></w:p>
<w:p><w:hyperlink r:id="rId2"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/><w:kern w:val="28"/></w:rPr><w:t>链接里的字</w:t></w:r></w:hyperlink></w:p>`;
  return pack(document);
}

/**
 * The `w:u` shapes, split by who states them.
 *
 * What the files write. Over the 86 WPS documents' body parts, 1586 `w:u` elements sit there: 1256 on a run's own
 * `w:rPr` (670 `none` and 200 `thick` inside table cells, 386 `single` in body paragraphs) and 330 on a paragraph
 * mark (282 `none` in cells, 48 `single` in body). Of the 1256 run-level ones this reader turns 488 into runs with
 * fields — **254 `<w:u w:val="none"/>`**, 194 `single`, 40 `thick`, none leaving `w:val` off — while the other 768
 * are runs it keeps as the bytes they came in with (fields and content controls), and those never pass through a
 * field at all (`.scratch/underline-counted.ts` walks both ways and reconciles them).
 *
 * Who could lend a line: `word/styles.xml` does state `w:u`, 35 times, every one of them inside a `w:style`'s own
 * `w:rPr` — 31 character styles `single`, 4 `none` — while `w:rPrDefault` states it 0 times and the body parts hold
 * 0 `w:hyperlink` elements. Cross the two and the audience of a cancellation is 0 today: all 254 `none` runs sit on
 * a chain whose every rung says nothing about `w:u`, and the 8 runs that do borrow a line say nothing themselves
 * (`.scratch/underline-audience.ts`). So this fixture's `LendU` is a shape **no file in hand has**: §17.7.2's rung
 * that lends a line, built only so the run's sentence above has something it has to beat.
 */
export function makeUnderlineDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr><w:u w:val="none"/></w:rPr><w:t>说明没有</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendU"/><w:u w:val="none"/></w:rPr><w:t>这句话盖掉借来的线</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendU"/></w:rPr><w:t>链上借来的线</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:u w:val="single"/></w:rPr><w:t>自己点了线</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:u w:val="thick"/></w:rPr><w:t>粗线原样回去</w:t></w:r></w:p>`;
  return pack(
    document,
    `<w:styles xmlns:w="${NS_W}"><w:style w:type="character" w:styleId="LendU"><w:name w:val="点了下划线的字符样式"/><w:rPr><w:u w:val="single"/></w:rPr></w:style></w:styles>`,
  );
}

/**
 * The other two members of the on/off family, in the shapes the files write.
 *
 * `w:strike` is §17.3.2.19's `ST_OnOff` like `w:b`, so its cancellation is `0` or `false`. The bodies of the 86 WPS
 * documents state a `w:strike` 56 times: 36 of them on a run's own set with no `w:val` — on, by the element's own
 * default — 18 on a paragraph mark, and **2 state `<w:strike w:val="0"/>`** in 2 files (`.scratch/onoff-where.ts`
 * cuts that by whose property set the tag sits in). Both cancellations are on a run carrying text, and both of them
 * the reader used to resolve to `strike = true`, which is a line drawn through text the file said to leave alone
 * (`.scratch/onoff-audience.ts`). `w:vertAlign` is §17.3.2.24's `ST_VerAlignRun`, whose cancellation is the enum
 * name `baseline`: the bodies state it 3860 times and **every one of them says `baseline`** — 2701 on a run's set,
 * 1159 on a mark — of which this reader models 717, 354 of those naming a character style, and 0 of them sitting on
 * a chain that lends a superscript. So that half is a statement kept rather than pixels rescued. A run stating both
 * a strike and an underline has 0 audience in these files, and is built here only so the drawing of two lines at
 * once has a case.
 */
export function makeOnOffDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr><w:strike w:val="0"/></w:rPr><w:t>说了没划</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendStrike"/><w:strike w:val="false"/></w:rPr><w:t>盖掉借来的删除线</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendStrike"/></w:rPr><w:t>链上借来的删除线</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendSup"/><w:vertAlign w:val="baseline"/></w:rPr><w:t>回到基线</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:strike/><w:u w:val="single"/></w:rPr><w:t>两条线一起</w:t></w:r></w:p>`;
  return pack(
    document,
    `<w:styles xmlns:w="${NS_W}">` +
      `<w:style w:type="character" w:styleId="LendStrike"><w:name w:val="点了删除线的字符样式"/><w:rPr><w:strike/></w:rPr></w:style>` +
      `<w:style w:type="character" w:styleId="LendSup"><w:name w:val="上标的字符样式"/><w:rPr><w:vertAlign w:val="superscript"/></w:rPr></w:style>` +
      `</w:styles>`,
  );
}

/**
 * `w:highlight`'s three states, in the shapes the files write.
 *
 * What the files write. Over the 298 body-bearing parts of the 86 WPS documents (正文、页眉、页脚、脚注、尾注) 4605
 * `w:highlight` elements sit there, and only two enum members appear: **2417 `none`** and 249 `yellow` on a run's own
 * set, 1718 `none` and 221 `yellow` on a paragraph mark (`.scratch/highlight-counted.ts` walks the tag stack and
 * reconciles both ways). Nothing leaves `w:val` off, and no name outside §17.17.64's list appears. Of the 2417
 * run-level cancellations this reader resolves 2369 into runs — 2361 of them carrying text, 1474 inside table cells
 * and 887 in body paragraphs — and the other 48 ride along as the bytes of a run it does not model.
 *
 * Who could lend a colour: **nobody**. `word/styles.xml` states `w:highlight` 0 times, `w:rPrDefault` 0, the
 * numbering part 0, and the header/footer/footnote parts 0 — so every one of those 2417 sentences cancels a rung
 * that says nothing, which makes this a statement kept rather than pixels rescued, exactly as `w:u w:val="none"` and
 * `w:vertAlign w:val="baseline"` were. `LendHl` is therefore a shape **no file in hand has**, built so the run's
 * sentence has something to beat; and `blue`, an §17.17.64 name the page cannot call back off its CSS, has 0
 * audience here too — these files only ever use `yellow`. What WPS itself writes for 无突出显示颜色 is the evidence
 * this fixture is built to keep: `<w:highlight w:val="none"/>`, not silence.
 */
export function makeHighlightDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr><w:highlight w:val="none"/></w:rPr><w:t>说了没高亮</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendHl"/><w:highlight w:val="none"/></w:rPr><w:t>盖掉借来的高亮</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rStyle w:val="LendHl"/></w:rPr><w:t>链上借来的高亮</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>黄底</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:highlight w:val="blue"/></w:rPr><w:t>页面叫不出名字的高亮</w:t></w:r></w:p>
<w:p><w:r><w:t>什么都没点</w:t></w:r></w:p>`;
  return pack(
    document,
    `<w:styles xmlns:w="${NS_W}">` +
      `<w:style w:type="character" w:styleId="LendHl"><w:name w:val="点了高亮的字符样式"/><w:rPr><w:highlight w:val="yellow"/></w:rPr></w:style>` +
      `</w:styles>`,
  );
}

/**
 * The shapes a 悬光标 (collapsed-caret) format click has to land on, taken from what the files actually state
 * (`.scratch/wps-rpr-names.ts`, 86 WPS-written documents, 101825 run property sets): `w:sz` **93946**, `w:szCs`
 * **96165**, `w:color` **61554**, `w:b` **8889**, `w:highlight` 2666. Sizes come in pairs and the two halves are not
 * always equal — 字号成对 is `<w:sz w:val="21"/><w:szCs w:val="24"/>`, the shape the model holds in one field, so a
 * changed size re-spells both (which Word does too when you pick a 字号).
 *
 * The ribbon's 字号 chooser offers 10/10.5/12/14/16/18/24/36 (`DOC_SIZES`), and the gap this fixture exists to keep
 * is the pending-format one: with nothing selected, a click on 字号 used to reach `applyStyle` and return without
 * doing anything, because Blink's `fontSize` command only takes the 1-7 enum and cannot say 22 磅 at all. 高亮 needed
 * the same editor-carried channel (§89), so the two families share it and a click across families stacks.
 */
export function makePendingFormatDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr><w:sz w:val="24"/><w:szCs w:val="24"/></w:rPr><w:t>十二磅</w:t></w:r></w:p>
<w:p><w:r><w:t>没写磅数</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:sz w:val="21"/><w:szCs w:val="24"/></w:rPr><w:t>字号成对</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:color w:val="C00000"/></w:rPr><w:t>红字</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>粗字</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>黄底</w:t></w:r></w:p>
<w:p><w:r><w:t>什么都没点</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * The shapes `w14:textFill` comes in — the 2010 extension's way of naming a run's colour, and the one colour
 * property that has a field of its own (`color`) *and* a carried slot the file wrote. Measured over the 86
 * WPS-written parts (`.scratch/textfill-counted.ts`, `.scratch/textfill-theme.ts`): all 3656 fills in the body
 * and header/footer parts sit under a `w:rPr` that states a `w:color` with a hex value, 3288 of them naming the
 * same theme colour as that `w:color` does, and once the `schemeClr` is resolved through `word/theme/theme1.xml`
 * — with `lumMod`/`lumOff` applied where the file stated them — all 3296 resolvable fills give back exactly the
 * hex `w:color` already gave. So it travels as a mirror of the colour field, and a save that re-spells the colour
 * has to take the mirror with it: 16 of those files hold a run with text whose fill would otherwise go on naming
 * the colour the run no longer has.
 */
export function makeTextFillDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr><w:rFonts w:hint="eastAsia"/><w:kern w:val="28"/><w:color w:val="808080" w:themeColor="text1"/><w:sz w:val="24"/><w:szCs w:val="24"/><w14:textFill><w14:solidFill><w14:schemeClr val="text1"/></w14:solidFill></w14:textFill></w:rPr><w:t>镜像的一对</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:color w:val="C00000"/><w:kern w:val="28"/></w:rPr><w:t>只说了 w:color</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:color w:val="404040" w:themeColor="dark1"/><w:sz w:val="21"/><w14:textFill><w14:solidFill><w14:schemeClr val="dark1"><w14:lumMod val="75000"/><w14:lumOff val="25000"/></w14:schemeClr></w14:solidFill></w14:textFill></w:rPr><w:t>明暗过一道</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w14:textFill><w14:solidFill><w14:srgbClr val="008000"/></w14:solidFill></w14:textFill></w:rPr><w:t>只说了那句扩展色</w:t></w:r></w:p>`;
  return pack(document);
}

/**
 * The shapes a ribbon has to echo. `makeRunPropsDocx` fixes what a *save* owes a run; this fixes what the
 * controls owe the reader sitting on that run, and the states here are the ones the corpus actually holds:
 * over the 86 WPS documents only `yellow` (249 runs with text, all inside table cells) and `none` (2361) ever
 * reach a run the caret can sit on, while 61 of those files hold a run whose size is none of the eight presets
 * the 字号 box lists — 21pt (五号) being the common one, so every stop there used to answer 16
 * (`<w:sz w:val="42"/>`). `darkRed` is the shape no WPS file in hand has: a colour name the palette cannot
 * paint, kept only so the box's third state — 文档说了、控件列不出 — has a case.
 *
 * The table is not decoration: of the corpus's 39,751 caret stops, 27,996 sit in a cell and 21,439 of the
 * paragraphs inside cells state their own `w:jc`, so the box's reading has to survive the table's walls
 * (`.scratch/align-three-state.ts`, `.scratch/cell-jc-audience.ts`).
 */
export function makeRibbonEchoDocx(): Uint8Array {
  const document = `<w:p><w:r><w:rPr><w:highlight w:val="yellow"/></w:rPr><w:t>黄底的一句</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:highlight w:val="none"/></w:rPr><w:t>说了没有的一句</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:highlight w:val="darkRed"/></w:rPr><w:t>控件列不出的一句</w:t></w:r></w:p>
<w:p><w:r><w:t>什么都没点的一句</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:sz w:val="42"/><w:szCs w:val="42"/></w:rPr><w:t>二十一磅的一句</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:b/><w:sz w:val="24"/><w:kern w:val="28"/></w:rPr><w:t>粗而十二磅</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:color w:val="C00000"/><w:u w:val="single"/></w:rPr><w:t>红而带线</w:t></w:r></w:p>
<w:tbl><w:tblGrid><w:gridCol w:w="2400"/></w:tblGrid><w:tr><w:tc>
<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:highlight w:val="yellow"/><w:sz w:val="36"/></w:rPr><w:t>格子里居中而黄底</w:t></w:r></w:p>
<w:p><w:r><w:t>格子里什么都没点</w:t></w:r></w:p>
</w:tc></w:tr></w:tbl>`;
  return pack(document);
}
