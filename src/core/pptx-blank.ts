/**
 * Canonical parts for a deck that never came from a file. Shapes follow what the
 * WPS blank template writes, which the authorized decks confirm: one slide master,
 * eleven layouts with the standard Chinese names, a theme, presProps/viewProps/
 * tableStyles and document properties. Grid pitch is the WPS-common 72008 EMU.
 */

const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_P = "http://schemas.openxmlformats.org/presentationml/2006/main";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const DECL = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n`;
const P_ROOT = `xmlns:a="${NS_A}" xmlns:r="${NS_R}" xmlns:p="${NS_P}"`;

/** Slide size of the WPS blank template (16:9), the size 30 of 35 authorized decks use. */
export const BLANK_SLIDE_WIDTH_EMU = 12192000;
export const BLANK_SLIDE_HEIGHT_EMU = 6858000;
/** Blank decks start their slide relationship ids after these fixed presentation rels. */
export const BLANK_FIXED_PRES_RELS = 5;

interface Ph {
  type?: string;
  idx?: number;
  sz?: string;
  orient?: "vert";
  custom?: boolean;
  x?: number;
  y?: number;
  w?: number;
  h?: number;
  /** Level-1 size the layout's `a:lstStyle` declares, in points (60/2400 is what the template writes). */
  szPt?: number;
  /** Level-1 alignment the layout declares; empty means the master's own (left) answers. */
  algn?: "l" | "ctr";
  /** `a:bodyPr@anchor`; empty means top. */
  anchor?: "t" | "ctr" | "b";
  /** The layout states `marL="0" indent="0"` for this placeholder, as 节标题/比较/图片与标题 do. */
  zeroIndent?: boolean;
  /** The layout states `<a:buNone/>`: its text carries no bullet even though the master does. */
  noBullet?: boolean;
}

interface LayoutDef {
  name: string;
  type?: string;
  /** WPS marks its three typeless layouts as user-drawn; PowerPoint hides untyped custom layouts without it. */
  userDrawn?: boolean;
  placeholders: Ph[];
}

const title = (x: number, y: number, w: number, h: number, custom = false, extra: Partial<Ph> = {}): Ph => ({ type: "title", custom, x, y, w, h, ...extra });
const body = (idx: number, x: number, y: number, w: number, h: number, extra: Partial<Ph> = {}): Ph => ({ type: "body", idx, x, y, w, h, ...extra });
const content = (idx: number, x: number, y: number, w: number, h: number, extra: Partial<Ph> = {}): Ph => ({ idx, x, y, w, h, ...extra });
const footer = (): Ph[] => [
  { type: "dt", sz: "half", idx: 10 },
  { type: "ftr", sz: "quarter", idx: 11 },
  { type: "sldNum", sz: "quarter", idx: 12 },
];

/** The eleven layouts of the WPS blank template, with the placeholder styles its `a:lstStyle` states. */
export const BLANK_LAYOUTS: LayoutDef[] = [
  {
    name: "标题幻灯片",
    type: "title",
    placeholders: [
      { type: "ctrTitle", x: 1198800, y: 914400, w: 9799200, h: 2570400, szPt: 60, algn: "ctr", anchor: "b" },
      { type: "subTitle", idx: 1, x: 1198800, y: 3560400, w: 9799200, h: 1472400, szPt: 24, algn: "ctr" },
      ...footer(),
    ],
  },
  {
    name: "标题和内容",
    type: "obj",
    placeholders: [title(608400, 608400, 10969200, 705600, false, { anchor: "ctr" }), content(1, 608400, 1490400, 10969200, 4759200), ...footer()],
  },
  {
    name: "节标题",
    type: "secHead",
    placeholders: [
      title(1990800, 3848400, 7768800, 766800, true, { szPt: 44, anchor: "b" }),
      body(1, 1990800, 4615200, 7768800, 867600, { custom: true, szPt: 18, zeroIndent: true, noBullet: true }),
      ...footer(),
    ],
  },
  {
    name: "两栏内容",
    type: "twoObj",
    placeholders: [
      title(608400, 608400, 10969200, 705600, false, { anchor: "ctr" }),
      content(1, 608400, 1501200, 5176800, 4748400, { sz: "half" }),
      content(2, 6411600, 1501200, 5176800, 4748400, { sz: "half" }),
      ...footer(),
    ],
  },
  {
    name: "比较",
    type: "twoTxTwoObj",
    placeholders: [
      title(608400, 608400, 10969200, 705600, false, { anchor: "ctr" }),
      body(1, 608400, 1429200, 5342400, 381600, { custom: true, szPt: 20, anchor: "t", zeroIndent: true, noBullet: true }),
      content(2, 608400, 1854000, 5342400, 4395600, { sz: "half" }),
      body(3, 6235750, 1421729, 5342400, 381600, { sz: "quarter", custom: true, szPt: 20, anchor: "t", zeroIndent: true, noBullet: true }),
      content(4, 6235750, 1854000, 5342400, 4395600, { sz: "quarter" }),
      ...footer(),
    ],
  },
  {
    name: "仅标题",
    type: "titleOnly",
    placeholders: [title(608400, 608400, 10969200, 705600, false, { anchor: "ctr" }), ...footer()],
  },
  {
    name: "空白",
    type: "blank",
    placeholders: footer(),
  },
  {
    name: "图片与标题",
    userDrawn: true,
    placeholders: [
      content(1, 608400, 1555200, 5233077, 4608000, { type: "pic", sz: "half", szPt: 16, zeroIndent: true, noBullet: true }),
      body(2, 6350400, 1555200, 5227200, 4608000, { sz: "half", szPt: 16, zeroIndent: true, noBullet: true }),
      ...footer(),
      { type: "title" },
    ],
  },
  {
    name: "竖排标题与文本",
    type: "vertTitleAndTx",
    placeholders: [
      { ...title(10234800, 914400, 1044000, 5029200, true, { szPt: 28, anchor: "ctr" }), orient: "vert" },
      body(1, 914400, 914400, 9169200, 5029200, { orient: "vert" }),
      ...footer(),
    ],
  },
  {
    name: "内容",
    userDrawn: true,
    placeholders: [...footer(), content(13, 608400, 774000, 10972800, 5482800, { sz: "quarter" })],
  },
  {
    name: "末尾幻灯片",
    userDrawn: true,
    placeholders: [
      ...footer(),
      title(1198800, 2484000, 9799200, 1018800, true, { szPt: 60, algn: "ctr", anchor: "t" }),
      body(13, 1198800, 3560400, 9799200, 471600, { sz: "quarter", szPt: 24, algn: "ctr" }),
    ],
  },
];

function phAttrs(p: Ph): string {
  return (
    (p.type ? ` type="${p.type}"` : "") +
    (p.idx !== undefined ? ` idx="${p.idx}"` : "") +
    (p.sz ? ` sz="${p.sz}"` : "") +
    (p.orient ? ` orient="${p.orient}"` : "") +
    (p.custom ? ` hasCustomPrompt="1"` : "")
  );
}

function spElement(id: number, name: string, p: Ph): string {
  const xfrm =
    p.x !== undefined ? `<a:xfrm><a:off x="${p.x}" y="${p.y}"/><a:ext cx="${p.w}" cy="${p.h}"/></a:xfrm>` : "";
  // The layout states its own level-1 size/alignment; where it stays silent the master's txStyles answer.
  const bodyPr = p.anchor ? `<a:bodyPr anchor="${p.anchor}"/>` : `<a:bodyPr/>`;
  const lvlAttrs =
    (p.algn ? ` algn="${p.algn}"` : "") + (p.zeroIndent ? ' marL="0" indent="0"' : "");
  const lvlKids = (p.noBullet ? "<a:buNone/>" : "") + `<a:defRPr${p.szPt ? ` sz="${p.szPt * 100}"` : ""}/>`;
  const lvl =
    p.algn || p.szPt || p.zeroIndent || p.noBullet
      ? `<a:lstStyle><a:lvl1pPr${lvlAttrs}>${lvlKids}</a:lvl1pPr></a:lstStyle>`
      : `<a:lstStyle/>`;
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph${phAttrs(p)}/></p:nvPr></p:nvSpPr><p:spPr>${xfrm}</p:spPr><p:txBody>${bodyPr}${lvl}<a:p/></p:txBody></p:sp>`;
}

const GROUP_PR = `<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>`;

const PLACEHOLDER_NAMES: Record<string, string> = {
  ctrTitle: "标题",
  title: "标题",
  subTitle: "副标题",
  body: "文本占位符",
  pic: "图片占位符",
  dt: "日期占位符",
  ftr: "页脚占位符",
  sldNum: "幻灯片编号占位符",
};

function layoutXml(layout: LayoutDef): string {
  const shapes = layout.placeholders
    .map((p, i) => spElement(i + 2, `${p.type ? PLACEHOLDER_NAMES[p.type] : "内容占位符"} ${i + 2}`, p))
    .join("");
  const type = layout.type ? ` type="${layout.type}"` : "";
  const userDrawn = layout.userDrawn ? ` userDrawn="1"` : "";
  return `${DECL}<p:sldLayout ${P_ROOT}${type} preserve="1"${userDrawn}><p:cSld name="${layout.name}"><p:spTree>${GROUP_PR}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`;
}

function masterXml(): string {
  const shapes = [
    spElement(2, "标题占位符 1", { type: "title", x: 608400, y: 608400, w: 10969200, h: 705600 }),
    spElement(3, "文本占位符 2", { type: "body", idx: 1, x: 608400, y: 1490400, w: 10969200, h: 4759200 }),
    spElement(4, "日期占位符 3", { type: "dt", sz: "half", idx: 2, x: 612000, y: 6314400, w: 2700000, h: 316800 }),
    spElement(5, "页脚占位符 4", { type: "ftr", sz: "quarter", idx: 3, x: 4116000, y: 6314400, w: 3960000, h: 316800 }),
    spElement(6, "幻灯片编号占位符 5", { type: "sldNum", sz: "quarter", idx: 4, x: 8877600, y: 6314400, w: 2700000, h: 316800 }),
  ].join("");
  const ids = BLANK_LAYOUTS.map((_, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join("");
  const defRPr = (sz: number): string =>
    `<a:defRPr sz="${sz}" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr>`;
  const titleStyle = `<a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">${defRPr(3600)}</a:lvl1pPr>`;
  // WPS's bodyStyle steps in 228600 EMU per level with a hanging -228600 indent; level 4's
  // bullet is Wingdings 0x9F (its filled square), which the XML carries as `&#x9F;`.
  const levelSizes = [1800, 1600, 1600, 1400, 1400, 1800, 1800, 1800, 1800];
  const levelBullets = ["●", "●", "●", "&#x9F;", "•", "•", "•", "•", "•"];
  const bodyStyle = levelSizes
    .map((sz, i) => {
      const face = i === 3 ? "Wingdings" : "Arial";
      return `<a:lvl${i + 1}pPr marL="${228600 + 457200 * i}" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:buFont typeface="${face}"/><a:buChar char="${levelBullets[i]}"/>${defRPr(sz)}</a:lvl${i + 1}pPr>`;
    })
    .join("");
  const txStyles = `<p:txStyles><p:titleStyle>${titleStyle}</p:titleStyle><p:bodyStyle>${bodyStyle}</p:bodyStyle><p:otherStyle><a:defPPr><a:defRPr lang="zh-CN"/></a:defPPr><a:lvl1pPr marL="0" algn="l">${defRPr(1800)}</a:lvl1pPr></p:otherStyle></p:txStyles>`;
  return `${DECL}<p:sldMaster ${P_ROOT}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GROUP_PR}${shapes}</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/><p:sldLayoutIdLst>${ids}</p:sldLayoutIdLst>${txStyles}</p:sldMaster>`;
}

/** The color and font scheme WPS writes in its blank theme, under our own theme name. */
function themeXml(): string {
  const clr = `<a:clrScheme name="Solo Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4874CB"/></a:accent1><a:accent2><a:srgbClr val="EE822F"/></a:accent2><a:accent3><a:srgbClr val="F2BA02"/></a:accent3><a:accent4><a:srgbClr val="75BD42"/></a:accent4><a:accent5><a:srgbClr val="30C0B4"/></a:accent5><a:accent6><a:srgbClr val="E54C5E"/></a:accent6><a:hlink><a:srgbClr val="0026E5"/></a:hlink><a:folHlink><a:srgbClr val="7E1FAD"/></a:folHlink></a:clrScheme>`;
  const font = `<a:fontScheme name="Solo Office"><a:majorFont><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:majorFont><a:minorFont><a:latin typeface="Arial"/><a:ea typeface="微软雅黑"/><a:cs typeface=""/></a:minorFont></a:fontScheme>`;
  const solid = `<a:solidFill><a:schemeClr val="phClr"/></a:solidFill>`;
  const fmt = `<a:fmtScheme name="Solo Office"><a:fillStyleLst>${solid}<a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumOff val="17500"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:hueOff val="-2520000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:gradFill><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:hueOff val="-4200000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"/></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/></a:gradFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst><a:outerShdw blurRad="101600" dist="50800" dir="5400000" algn="ctr" rotWithShape="0"><a:schemeClr val="phClr"><a:alpha val="60000"/></a:schemeClr></a:outerShdw></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:reflection stA="50000" endA="300" endPos="40000" dist="25400" dir="5400000" sy="-100000" algn="bl" rotWithShape="0"/></a:effectLst></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst>${solid}<a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme>`;
  return `${DECL}<a:theme xmlns:a="${NS_A}" name="Solo Office"><a:themeElements>${clr}${font}${fmt}</a:themeElements></a:theme>`;
}

/** Every part a blank deck needs beyond presentation.xml, slides and slide rels. */
export function blankDeckParts(slideCount: number, now = new Date()): Record<string, string> {
  const parts: Record<string, string> = {};
  parts["ppt/theme/theme1.xml"] = themeXml();
  parts["ppt/slideMasters/slideMaster1.xml"] = masterXml();
  parts["ppt/slideMasters/_rels/slideMaster1.xml.rels"] = `${DECL}<Relationships xmlns="${NS_PKG}">${BLANK_LAYOUTS.map(
    (_, i) => `<Relationship Id="rId${i + 1}" Type="${NS_R}/slideLayout" Target="../slideLayouts/slideLayout${i + 1}.xml"/>`,
  ).join("")}<Relationship Id="rId${BLANK_LAYOUTS.length + 1}" Type="${NS_R}/theme" Target="../theme/theme1.xml"/></Relationships>`;
  BLANK_LAYOUTS.forEach((layout, i) => {
    parts[`ppt/slideLayouts/slideLayout${i + 1}.xml`] = layoutXml(layout);
    parts[`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`] =
      `${DECL}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/slideMaster" Target="../slideMasters/slideMaster1.xml"/></Relationships>`;
  });
  parts["ppt/presProps.xml"] = `${DECL}<p:presentationPr ${P_ROOT}/>`;
  parts["ppt/viewProps.xml"] = `${DECL}<p:viewPr ${P_ROOT}><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="13778" autoAdjust="0"/><p:restoredTop sz="94660"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" showGuides="1"><p:cViewPr varScale="1"><p:scale><a:sx n="99" d="100"/><a:sy n="99" d="100"/></p:scale><p:origin x="84" y="582"/></p:cViewPr><p:guideLst><p:guide orient="horz" pos="2160"/><p:guide pos="3840"/></p:guideLst></p:cSldViewPr></p:slideViewPr><p:gridSpacing cx="72008" cy="72008"/></p:viewPr>`;
  parts["ppt/tableStyles.xml"] = `${DECL}<a:tblStyleLst xmlns:a="${NS_A}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`;
  const stamp = now.toISOString().replace(/\.\d+Z$/, "Z");
  parts["docProps/core.xml"] = `${DECL}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title/><dc:creator/><cp:lastModifiedBy/><dcterms:created xsi:type="dcterms:W3CDTF">${stamp}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${stamp}</dcterms:modified></cp:coreProperties>`;
  parts["docProps/app.xml"] = `${DECL}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><Application>Solo Office</Application><PresentationFormat>宽屏</PresentationFormat><Slides>${slideCount}</Slides></Properties>`;
  parts["_rels/.rels"] = `${DECL}<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="ppt/presentation.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="${NS_R}/extended-properties" Target="docProps/app.xml"/></Relationships>`;
  return parts;
}
