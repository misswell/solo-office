import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_WPG = "http://schemas.microsoft.com/office/word/2010/wordprocessingGroup";
const NS_WPS = "http://schemas.microsoft.com/office/word/2010/wordprocessingShape";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_V = "urn:schemas-microsoft-com:vml";
const NS_O = "urn:schemas-microsoft-com:office:office";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";

/**
 * The colour half of a theme part, in the shape `business plan.dotm` writes it: `a:clrScheme` over the slots,
 * then three style lists addressed by position.
 *
 * The lists are deliberately one plain and two this reader refuses: `a:fillStyleLst[1]` is a solid `phClr` with no
 * modifiers (expandable), `[2]` a `gradFill`, and `[3]` a `phClr` carrying `a:tint` — both of which answer
 * undefined, because the arithmetic to expand them is not stated anywhere in the file. `a:lnStyleLst[1]` is a solid
 * `phClr` with the pen the entry itself states (`w="9525"` = 1 px) and `[2]` names its own colour instead of
 * waiting for the shape to supply one.
 */
export function themeColorXml(): string {
  return (
    XML_DECL +
    `<a:theme xmlns:a="${NS_A}" name="Office 主题"><a:themeElements>` +
    `<a:clrScheme name="Office">` +
    `<a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1>` +
    `<a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>` +
    `<a:accent1><a:srgbClr val="4472C4"/></a:accent1>` +
    `<a:accent2><a:srgbClr val="ED7D31"/></a:accent2>` +
    `</a:clrScheme>` +
    `<a:fillStyleLst>` +
    `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>` +
    `<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"/></a:gs></a:gsLst></a:gradFill>` +
    `<a:solidFill><a:schemeClr val="phClr"><a:tint val="60000"/></a:schemeClr></a:solidFill>` +
    `</a:fillStyleLst>` +
    `<a:lnStyleLst>` +
    `<a:ln w="9525" cap="flat"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill></a:ln>` +
    `<a:ln w="19050" cap="flat"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>` +
    `</a:lnStyleLst>` +
    `</a:themeElements></a:theme>`
  );
}

const STYLES = `<w:styles xmlns:w="${NS_W}"><w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults><w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style></w:styles>`;

/** A paragraph inside a shape's text box: the words the file put *in* the figure, not beside it. */
const boxPara = (word: string, sz = "24"): string =>
  `<w:p><w:pPr><w:jc w:val="center"/></w:pPr><w:r><w:rPr><w:b/><w:sz w:val="${sz}"/></w:rPr><w:t>${word}</w:t></w:r></w:p>`;

/**
 * A `wps:wsp`: a box, a geometry, and either the paint it states in `a:spPr` or the theme entry it borrows
 * through `wps:style`. Each of the six below is one of the shapes the authorized diagram is made of.
 */
const wsp = (inner: string): string => `<wps:wsp>${inner}<wps:bodyPr anchor="ctr"/></wps:wsp>`;

/** The six-member group: borrowed paint, a flat-line connector, a custom path, an empty box with words, and two shapes this reader declines. */
const GROUP =
  `<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="4762500" cy="2857500"/>` +
  `<a:chOff x="0" y="0"/><a:chExt cx="500" cy="600"/></a:xfrm></wpg:grpSpPr>` +
  // Borrowed both ways: fill from `a:fillStyleLst[1]` coloured accent1, line from `a:lnStyleLst[1]` with its 1px pen.
  wsp(
    `<wps:cNvPr id="10" name="借来的框"/><wps:spPr><a:xfrm><a:off x="10" y="20"/><a:ext cx="100" cy="60"/></a:xfrm>` +
      `<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 25000"/></a:avLst></a:prstGeom></wps:spPr>` +
      `<wps:style><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:lnRef idx="1"><a:schemeClr val="accent2"/></a:lnRef></wps:style>` +
      `<wps:txbx><w:txbxContent>${boxPara("公司")}</w:txbxContent></wps:txbx>`,
  ) +
  // A horizontal connector: `cy="0"` is a statement, not an absent box. Its pen is 38100 EMU = 4px, and its fill
  // must not paint — a connector's `a:solidFill` is the colour it would take if it became a shape.
  wsp(
    `<wps:cNvPr id="11" name="横线"/><wps:spPr><a:xfrm><a:off x="200" y="300"/><a:ext cx="50" cy="0"/></a:xfrm>` +
      `<a:prstGeom prst="line"><a:avLst/></a:prstGeom>` +
      `<a:ln w="38100"><a:solidFill><a:srgbClr val="767171"/></a:solidFill></a:ln></wps:spPr>`,
  ) +
  // `a:custGeom` on a ruler of its own (`a:path w="100" h="50"`) over a box of 80×40, turned 90° and mirrored first.
  wsp(
    `<wps:cNvPr id="12" name="自定义"/><wps:spPr><a:xfrm rot="5400000" flipH="1">` +
      `<a:off x="300" y="100"/><a:ext cx="80" cy="40"/></a:xfrm>` +
      `<a:custGeom><a:pathLst><a:path w="100" h="50"><a:moveTo><a:pt x="0" y="0"/></a:moveTo>` +
      `<a:lnTo><a:pt x="100" y="50"/></a:lnTo><a:lnTo><a:pt x="0" y="50"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>` +
      `<a:solidFill><a:schemeClr val="dk1"/></a:solidFill></wps:spPr>`,
  ) +
  // A text box: `a:noFill` is an absence the file spoke, and the words belong to the figure, not the flow.
  wsp(
    `<wps:cNvPr id="13" name="文本框"/><wps:spPr><a:xfrm><a:off x="400" y="500"/><a:ext cx="80" cy="40"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
      `<wps:txbx><w:txbxContent>${boxPara("部门")}</w:txbxContent></wps:txbx>`,
  ) +
  // A preset this reader does not draw: keeps its bytes, stays off the page, and says so.
  wsp(
    `<wps:cNvPr id="14" name="三角形"/><wps:spPr><a:xfrm><a:off x="60" y="60"/><a:ext cx="40" cy="40"/></a:xfrm>` +
      `<a:prstGeom prst="triangle"><a:avLst/></a:prstGeom></wps:spPr>`,
  ) +
  // A shape that names no box at all — not zero, nothing.
  wsp(`<wps:cNvPr id="15" name="无位置"/><wps:spPr><a:xfrm><a:off x="1" y="1"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>`) +
  `</wpg:wgp>`;

/** One shape standing alone in its `a:graphicData`: no group, so its own box is the coordinate space. */
const LONE =
  `<wps:wsp><wps:cNvPr id="20" name="单独的形状"/><wps:spPr>` +
  `<a:xfrm><a:off x="952500" y="476250"/><a:ext cx="1905000" cy="952500"/></a:xfrm>` +
  `<a:prstGeom prst="roundRect"><a:avLst/></a:prstGeom>` +
  `<a:solidFill><a:srgbClr val="C00000"/></a:solidFill></wps:spPr><wps:bodyPr/></wps:wsp>`;

const graphic = (inner: string, uri: string): string =>
  `<a:graphic><a:graphicData uri="${uri}">${inner}</a:graphicData></a:graphic>`;

const INLINE = `<wp:inline><wp:extent cx="4762500" cy="2857500"/><wp:effectExtent l="0" t="0" r="0" b="0"/>` + `<wp:docPr id="1" name="图示"/>${graphic(GROUP, NS_WPG)}</wp:inline>`;

/** The lone shape, inline (drawn) and then anchored (not): the two ways a `w:drawing` can be placed. */
const LONE_INLINE = `<wp:inline><wp:extent cx="1905000" cy="952500"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="2" name="单独"/>${graphic(LONE, NS_WPS)}</wp:inline>`;
/**
 * The anchored half spells itself the way the authorized footers do — `wp:anchor` around a text box, not around a
 * bare figure — because the question worth pinning is what the page can still say about a shape it cannot place.
 */
const ANCHOR_BOX =
  `<wps:wsp><wps:cNvPr id="21" name="浮动的框"/><wps:spPr>` +
  `<a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></wps:spPr>` +
  `<wps:txbx><w:txbxContent>${boxPara("浮动")}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`;
const LONE_ANCHOR =
  `<wp:anchor simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
  `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
  `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
  `<wp:extent cx="1905000" cy="952500"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:wrapNone/>` +
  `<wp:docPr id="3" name="浮动"/>${graphic(ANCHOR_BOX, NS_WPS)}</wp:anchor>`;

/** The same figure written twice over, as the corpus writes it: the words live in the Choice and again in the VML copy. */
const ALTERNATE =
  `<mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline>` +
  `<wp:extent cx="1905000" cy="952500"/><wp:effectExtent l="0" t="0" r="0" b="0"/><wp:docPr id="4" name="两写"/>` +
  graphic(
    `<wps:wsp><wps:cNvPr id="40" name="两写的框"/><wps:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>` +
      `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></wps:spPr>` +
      `<wps:txbx><w:txbxContent>${boxPara("重复")}</w:txbxContent></wps:txbx><wps:bodyPr/></wps:wsp>`,
    NS_WPS,
  ) +
  `</wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict>` +
  `<v:shape id="框" o:spid="_x0000_s1026" type="#_x0000_t201"><v:textbox inset="0,0,0,0">` +
  `<w:txbxContent>${boxPara("重复")}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent>`;

/**
 * The fixture `business plan.dotm` reduced to one paragraph each: a `wpg:wgp` group of six shapes drawn from the
 * theme and from itself, a lone shape (whose own box is therefore the coordinate space), an anchored drawing this
 * reader still declines, and the same figure written twice over as a Choice and a VML Fallback.
 *
 * `theme` swaps the colour part for `null` — the shape of a document whose theme states no colours, where every
 * borrowed paint answers unknown and only stated paint draws.
 */
export function makeShapeDocx(opts: { theme?: string | null } = {}): Uint8Array {
  const theme = opts.theme === null ? undefined : (opts.theme ?? themeColorXml());
  const document =
    `<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>${INLINE}</w:drawing></w:r></w:p>` +
    `<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>${LONE_INLINE}</w:drawing></w:r></w:p>` +
    `<w:p><w:r><w:rPr><w:noProof/></w:rPr><w:drawing>${LONE_ANCHOR}</w:drawing></w:r></w:p>` +
    `<w:p><w:r><w:rPr><w:noProof/></w:rPr>${ALTERNATE}</w:r></w:p>` +
    `<w:p><w:r><w:t>正文一句</w:t></w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="${CT_W}"/>` +
          `<Override PartName="/word/styles.xml" ContentType="${CT_STYLES}"/>` +
          (theme ? `<Override PartName="/word/theme/theme1.xml" ContentType="${CT_THEME}"/>` : "") +
          `</Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/>` +
          (theme ? `<Relationship Id="rId2" Type="${NS_R}/theme" Target="theme/theme1.xml"/>` : "") +
          `</Relationships>`,
      ),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:wpg="${NS_WPG}" ` +
          `xmlns:wps="${NS_WPS}" xmlns:mc="${NS_MC}" xmlns:v="${NS_V}" xmlns:o="${NS_O}"><w:body>${document}</w:body></w:document>`,
      ),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
      ...(theme ? { "word/theme/theme1.xml": strToU8(theme) } : {}),
    },
    { level: 6 },
  );
}
