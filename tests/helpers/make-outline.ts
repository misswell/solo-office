import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STY = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";

/**
 * The paragraph styles, in the shapes `.scratch/outline-census.ts` found across the 86 WPS-written documents.
 * What matters for the fixtures is who states the level rather than who reads it:
 *
 * - 433 of their styles state a `w:outlineLvl`, and **none** of those styleIds matches `^Heading([1-9])$` — WPS
 *   signs its styles with names like 「标题 1」 over an opaque id, so a renderer that keys the heading off the
 *   id sees no heading in these files at all. `T1`/`T2`/`Deep`/`Body` keep that shape.
 * - 9 of the 433 state `w:val="9"`, which is §17.3.2.26's body text: `Body` naming `T1` as its `w:basedOn` and
 *   then stating 9 is the shape that takes back the level its ancestor lends.
 * - a chain two steps long is in the corpus (`w:basedOn` chains resolve 582 paragraph levels), so `Deep` states
 *   nothing and borrows `T2`'s level through `T2`, which borrowed nothing because it states its own.
 *
 * `Heading2` is the odd one out on purpose: Word spells its built-in heading styles exactly that way, so the
 * id-only branch in the renderer stays tested even though no file in this corpus lands on it.
 *
 * The `w:next` set is the same census applied to §17.7.2.24 (`.scratch/next-resolve.ts`): 668 of their 2,378
 * styles state a next style, every one of those targets is a declared paragraph style and none is the style that
 * states it, and 358 more take a next only from a `w:basedOn` ancestor. The three shapes that are not the headline
 * case are here because they are what the corpus says a *reader* has to get wrong quietly:
 *
 * - `Normal` — the default paragraph style — states a next in 63 of the 86 files, and 16,783 of their 30,322
 *   paragraphs name no style at all. A reader that lends the default style's next to those paragraphs turns body
 *   text into another style on every second Enter (924 of them would land on a heading style).
 * - `Borrow` states no next while its ancestor does: that is the 358.
 * - `Dangling` names a style the part does not declare and `SelfNext` names itself; neither is in the corpus
 *   (both count zero there), and neither is something a split could apply.
 */
const STYLES = `<w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:next w:val="Body1"/></w:style>
<w:style w:type="paragraph" w:styleId="T1"><w:name w:val="标题 1"/><w:basedOn w:val="Normal"/><w:next w:val="Body1"/><w:pPr><w:outlineLvl w:val="0"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Body1"><w:name w:val="正文 1"/><w:basedOn w:val="Normal"/><w:pPr><w:spacing w:before="120" w:after="120" w:line="300" w:lineRule="auto"/><w:jc w:val="both"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Borrow"><w:name w:val="借来的后续"/><w:basedOn w:val="T1"/></w:style>
<w:style w:type="paragraph" w:styleId="Dangling"><w:name w:val="后续样式没声明"/><w:next w:val="NoSuchStyle"/></w:style>
<w:style w:type="paragraph" w:styleId="SelfNext"><w:name w:val="后续是自己"/><w:next w:val="SelfNext"/></w:style>
<w:style w:type="paragraph" w:styleId="T2"><w:name w:val="标题 2"/><w:basedOn w:val="T1"/><w:pPr><w:outlineLvl w:val="1"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Deep"><w:name w:val="借两级"/><w:basedOn w:val="T2"/></w:style>
<w:style w:type="paragraph" w:styleId="Body"><w:name w:val="正文回到9"/><w:basedOn w:val="T1"/><w:pPr><w:outlineLvl w:val="9"/></w:pPr></w:style>
<w:style w:type="paragraph" w:styleId="Heading2"><w:name w:val="heading 2"/></w:style>
</w:styles>`;

/** One paragraph: `pPr` is the caller's own `w:pPr` children, written the way the files write them. */
export interface OutlinePara {
  pPr?: string;
  text: string;
}

/** A body block: paragraphs, or a table whose cells hold paragraphs. */
export type OutlineBlock = OutlinePara | { table: OutlinePara[][] };

/** `w:pStyle` naming a style from `STYLES`. */
export const pStyle = (id: string): string => `<w:pStyle w:val="${id}"/>`;

/** The paragraph's own `w:outlineLvl` (§17.3.2.26): 0…8 for a heading level, 9 for body text. */
export const outlineLvl = (val: number): string => `<w:outlineLvl w:val="${val}"/>`;

/** The shape the files sign: `CT_PPr` puts `w:outlineLvl` after `w:jc`, so a fixture that wants both says so. */
export const jcBefore = (align: string, lvl: number): string => `<w:jc w:val="${align}"/>${outlineLvl(lvl)}`;

function paraXml(e: OutlinePara): string {
  return `<w:p>${e.pPr ? `<w:pPr>${e.pPr}</w:pPr>` : ""}<w:r><w:t>${e.text}</w:t></w:r></w:p>`;
}

function tableXml(rows: OutlinePara[][]): string {
  return `<w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="2400"/><w:gridCol w:w="2400"/></w:tblGrid>${rows
    .map(
      (cells) =>
        `<w:tr>${cells
          .map((c) => `<w:tc><w:tcPr><w:tcW w:w="2400"/></w:tcPr>${paraXml(c)}</w:tc>`)
          .join("")}</w:tr>`,
    )
    .join("")}</w:tbl>`;
}

/**
 * A document of the given blocks over `STYLES`, for the tests about which paragraphs the outline levels name as
 * headings. Cells are the point of the table shape: `doc.outline()` and the page's `[data-outline]` marks both
 * walk the model in the order the blocks are written, and a heading inside a cell has to land between the
 * paragraph before the table and the one after it.
 */
export function makeOutlineDocx(blocks: readonly OutlineBlock[]): Uint8Array {
  const body =
    blocks
      .map((b) => ("table" in b ? tableXml(b.table) : paraXml(b)))
      .join("\n") + `<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="${CT_W}"/><Override PartName="/word/styles.xml" ContentType="${CT_STY}"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${body}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + STYLES),
    },
    { level: 6 },
  );
}
