import { strToU8, zipSync } from "fflate";

const DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_WP = "http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_PIC = "http://schemas.openxmlformats.org/drawingml/2006/picture";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const NS_MC = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const NS_V = "urn:schemas-microsoft-com:vml";
const NS_W14 = "http://schemas.microsoft.com/office/word/2010/wordml";
const NS_WP14 = "http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing";
const NS_CUST = "urn:fixture:custom-data";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml";
/** A real one-pixel PNG: the media part has to be bytes a reader can resolve, not a placeholder string. */
const PNG_1PX =
  "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

/**
 * A package shaped like the ones the local WPS corpus writes, so a save has something to lose:
 *
 * - `[Content_Types].xml` promises `.rels` and `.xml` by **extension** and carries no `Override` for
 *   `/_rels/.rels` — re-spelling the table per part would invent one for every part in the package.
 * - **no comments part at all**, in any of the three shapes a reader looks for.
 * - a `w:br` with `w:type="page"` and one with `w:type="column" w:clear="none"`: the kinds a field list
 *   that only knows "is this a break?" cannot tell apart.
 * - a `w:bookmarkEnd` sitting **between two rows** of a table, closing a bookmark that started above it —
 *   where OOXML puts it and where WPS puts it.
 * - a **floating** picture (`wp:anchor`) with the frame's own children: `wp:effectExtent`, a crop
 *   (`a:srcRect`), an outline (`a:ln`), and a `a:extLst` hanging off the blip. A drawing re-spelled from
 *   five fields keeps none of them and turns the anchor into an inline one.
 * - an inline picture citing `rId8`, whose id the save must keep.
 * - a root in the order a word processor actually writes it — `mc:` and `wpc:` before `w:` and `r:` — with an
 *   `mc:Ignorable` list, a `xml:lang` this model has no field for, and prefixes no tag here uses. Hoisting
 *   `xmlns:w` to the front or spelling the declaration list from a map is a re-spelling, not the file.
 * - a **VML** picture (`w:pict`/`v:imagedata`) citing `rId9`, whose relationship the file wrote with an
 *   absolute `Target="/word/media/…"`. The shape travels by the byte, so its citation does too, and a save
 *   that re-issued the image relationship under its own spelling left that citation pointing at nothing.
 */
export function makePreservationDocx(): Uint8Array {
  const anchor =
    `<w:drawing><wp:anchor distT="0" distB="0" distL="0" distR="0" simplePos="0" relativeHeight="251659264" behindDoc="0" locked="0" layoutInCell="1" allowOverlap="1">` +
    `<wp:simplePos x="0" y="0"/><wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>` +
    `<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>` +
    `<wp:extent cx="1219200" cy="609600"/><wp:effectExtent l="0" t="0" r="12700" b="12700"/>` +
    `<wp:docPr id="7" name="浮动图"/><wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="7" name="浮动图"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId8"><a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}"><a14:useLocalDpi xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" val="0"/></a:ext></a:extLst></a:blip>` +
    `<a:srcRect l="10000" t="10000" r="10000" b="10000"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1219200" cy="609600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:ln w="12700"><a:noFill/></a:ln></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing>`;
  const inline =
    `<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="1219200" cy="609600"/><wp:docPr id="9" name="行内图"/>` +
    `<a:graphic><a:graphicData uri="${NS_PIC}"><pic:pic>` +
    `<pic:nvPicPr><pic:cNvPr id="9" name="行内图"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rId8"/><a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1219200" cy="609600"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;

  const pict =
    `<w:pict><v:shape id="印章1" type="#_x0000_t75" style="width:96pt;height:48pt">` +
    `<v:fill><v:imagedata r:id="rId9"/></v:fill>` +
    `</v:shape></w:pict>`;

  const document =
    `<w:document xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" xmlns:v="${NS_V}" xmlns:cust="${NS_CUST}" xmlns:wp14="${NS_WP14}" xmlns:wp="${NS_WP}" xmlns:a="${NS_A}" xmlns:pic="${NS_PIC}" xmlns:r="${NS_R}" xmlns:w="${NS_W}" xml:lang="zh-CN" mc:Ignorable="w14 wp14">` +
    `<w:body>` +
    `<w:p><w:r><w:t>正文</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>下一页</w:t></w:r>` +
    `<w:r><w:br w:type="column" w:clear="none"/></w:r><w:r><w:t xml:space="preserve"> 带边空格的 </w:t></w:r></w:p>` +
    `<w:bookmarkStart w:id="3" w:name="_Toc1"/>` +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
    `<w:tr><w:tc><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr>` +
    `<w:bookmarkEnd w:id="3"/>` +
    `<w:tr><w:tblPrEx><w:tblBorders><w:top w:val="single" w:sz="12" w:space="0" w:color="000000"/></w:tblBorders></w:tblPrEx>` +
    `<w:trPr><w:trHeight w:val="600"/></w:trPr>` +
    `<w:tc><w:p><w:r><w:t>丙</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>丁</w:t></w:r></w:p></w:tc></w:tr>` +
    `</w:tbl>` +
    `<w:p><w:r>${anchor}</w:r></w:p>` +
    `<w:p><w:r>${inline}</w:r><w:r><w:t>图后文字</w:t></w:r></w:p>` +
    `<w:p><w:r>${pict}</w:r></w:p>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
    `</w:body></w:document>`;

  const files: Record<string, Uint8Array> = {
    "[Content_Types].xml": strToU8(
      DECL +
        `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Default Extension="png" ContentType="image/png"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/></Types>`,
    ),
    "_rels/.rels": strToU8(
      DECL +
        `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
    ),
    "word/_rels/document.xml.rels": strToU8(
      DECL +
        `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId8" Type="${NS_R}/image" Target="media/trend.png"/><Relationship Id="rId9" Type="${NS_R}/image" Target="/word/media/seal.png"/></Relationships>`,
    ),
    "word/document.xml": strToU8(DECL + document),
    "word/media/trend.png": new Uint8Array(Buffer.from(PNG_1PX, "base64")),
    "word/media/seal.png": new Uint8Array(Buffer.from(PNG_1PX, "base64")),
  };
  return zipSync(files, { level: 6 });
}

/**
 * The `w:br` family on its own, with a table next to it: `w:type` `page` / `column`+`w:clear="none"` /
 * `textWrapping`+`w:clear="all"`, each spelled once, plus one cell whose paragraph holds plain text. A reader
 * that only knows "is this a break?" cannot tell the three apart, and a drawer that draws them all as `<br>`
 * hands the editor a page the model cannot read back — so a keystroke in that paragraph re-spells a page break
 * as a line break.
 *
 * This one carries no media, which is the point of it being separate: the DOM gates import the fixture library
 * and run where there is no `Buffer`, so a package built with base64 media cannot be assembled there.
 */
export function makeBrFamilyDocx(): Uint8Array {
  const document =
    `<w:document xmlns:mc="${NS_MC}" xmlns:w14="${NS_W14}" xmlns:r="${NS_R}" xmlns:w="${NS_W}" xml:lang="zh-CN" mc:Ignorable="w14">` +
    `<w:body>` +
    `<w:p><w:r><w:t>正文</w:t></w:r><w:r><w:br w:type="page"/></w:r><w:r><w:t>下一页</w:t></w:r>` +
    `<w:r><w:br w:type="column" w:clear="none"/></w:r><w:r><w:t>分栏后</w:t></w:r>` +
    `<w:r><w:br w:type="textWrapping" w:clear="all"/></w:r><w:r><w:t>折行后</w:t></w:r></w:p>` +
    `<w:tbl><w:tblPr><w:tblW w:w="0" w:type="auto"/></w:tblPr><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>` +
    `<w:tr><w:tc><w:p><w:r><w:t>甲</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>乙</w:t></w:r></w:p></w:tc></w:tr>` +
    `</w:tbl>` +
    `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1440" w:bottom="1440" w:left="1440"/></w:sectPr>` +
    `</w:body></w:document>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(DECL + document),
    },
    { level: 6 },
  );
}
