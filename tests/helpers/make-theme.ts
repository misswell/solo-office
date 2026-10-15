import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";
const CT_W = "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml";
const CT_STYLES = "application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml";
const CT_NUM = "application/vnd.openxmlformats-officedocument.wordprocessingml.numbering+xml";
const CT_THEME = "application/vnd.openxmlformats-officedocument.theme+xml";

/**
 * The package's theme part, in the shape the 201 WPS-written documents write it: `a:theme`/`a:themeElements`/
 * `a:fontScheme` over `a:majorFont` and `a:minorFont`, each of the two naming its Latin face in `a:latin`
 * (201 of 201 do, for both slots), leaving `a:ea` empty (200 of 201) and `a:cs` empty (201 of 201), and keeping
 * the per-writing-system faces as *siblings* of those three rather than children of `a:ea` — 5,947 `a:font`
 * entries per slot, each with a `typeface` (`.scratch/theme-font-slots.ts`, `.scratch/theme-font-shape.ts`,
 * `.scratch/theme-font-nesting.ts`).
 *
 * `eaNamed` switches the one file in 201 that states an East Asian face for a slot: with it the empty `a:ea`
 * becomes a named face, which is the shape that lets `w:eastAsiaTheme` resolve without the run saying a language.
 */
export function themeXml(eaNamed = false): string {
  const scripts = (hans: string): string =>
    `<a:font script="Hans" typeface="${hans}"/><a:font script="Hant" typeface="標楷體"/>` +
    `<a:font script="Jpan" typeface="MS Mincho"/><a:font script="Hang" typeface="Malgun Gothic"/>`;
  const slot = (latin: string, hans: string): string =>
    `<a:latin typeface="${latin}"/><a:ea typeface="${eaNamed ? "等线" : ""}"/><a:cs typeface=""/>${scripts(hans)}`;
  return (
    XML_DECL +
    `<a:theme xmlns:a="${NS_A}" name="Office 主题"><a:themeElements><a:fontScheme name="Office">` +
    `<a:majorFont>${slot("Cambria", "华文中宋")}</a:majorFont>` +
    `<a:minorFont>${slot("Calibri", "仿宋")}</a:minorFont>` +
    `</a:fontScheme></a:themeElements></a:theme>`
  );
}

const DEFAULTS = '<w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:sz w:val="22"/>';

const STYLES = `<w:styles xmlns:w="${NS_W}">
<w:docDefaults><w:rPrDefault><w:rPr>${DEFAULTS}</w:rPr></w:rPrDefault></w:docDefaults>
<w:style w:type="paragraph" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
<w:style w:type="paragraph" w:styleId="Themed"><w:name w:val="主题标题"/><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia"/></w:rPr></w:style>
</w:styles>`;

/** One `w:lvl` whose mark hands its face to the theme, and the number that points at it. */
const NUMBERING = `<w:numbering xmlns:w="${NS_W}">
<w:abstractNum w:abstractNumId="0">
<w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="bullet"/><w:lvlText w:val="${String.fromCharCode(0xf0a7)}"/><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:hint="eastAsia"/><w:lang w:eastAsia="zh-CN"/></w:rPr></w:lvl>
</w:abstractNum>
<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;

/**
 * Every way this model lets a `w:rFonts` name a face without naming one, each in the paragraph shape the files
 * write it in. The audience is measured, not assumed (`.scratch/theme-font-audience.ts` over the 201 documents
 * this machine's WPS wrote): 54,200 `w:rFonts` in the bodies, **5,702 of them carrying a theme attribute**. That is
 * per *element*, and the interesting count is per *slot* (`.scratch/theme-slot-pair.ts`, all 55,585 `w:rFonts` in
 * every part of those files): **no slot ever states a literal face and a theme reference together — 0 elements**,
 * while 5,260 elements do carry both, each one naming one slot literally and pointing another at the theme. So a
 * reference is the slot saying "no face of my own", never an extra opinion next to a face the same slot named.
 * 27,831 state only literal faces, so both shapes live in the same parts. By layer: 2,756 sit in a paragraph's
 * `w:pPr/w:rPr` and 2,946 in a run's own; by part, 264 more in `word/styles.xml` —
 * of which 250 point at a *Latin* slot, and 58 of those are `w:docDefaults`, the whole run floor
 * (`.scratch/theme-font-latin-audience.ts`: in those 58 files, 8,962 body paragraphs state no literal face).
 * The part is reached as the files reach it: a `theme` relationship on `word/document.xml` with a *relative*
 * target, declared by an `Override` of `application/vnd.openxmlformats-officedocument.theme+xml`, in 201 of 201
 * (`.scratch/theme-rel.ts`).
 *
 * What the fixtures therefore hold: the floor read through the theme (`地板给的脸`), a style's `w:rPr`
 * (`样式给的脸`), a literal face and a reference on the *same slot* for the case §17.3.2 leaves unranked
 * (`字面先赢`, stated by 0 of the 55,585 — pinned as a rule, not as a shape), the East Asian slot reached through the
 * run's own `w:lang@w:eastAsia` (`汉字有语言`) and refused without it (`汉字无语言`, `标记只说hint`), the paragraph
 * mark's own pair (`标记的脸`), and all four slots handed to the theme at once (`整行交给主题`) — the shape the
 * write side is measured against, since moving one of its slots must retire that slot's reference and leave the
 * other three as they were. `编号的脸` hands a list mark's PUA glyph to the theme: `word/numbering.xml` holds
 * 421 `w:rFonts` in these files and **none** of them uses a theme attribute, so that one is the rule, not the
 * corpus.
 */
export function makeThemeDocx(opts: { theme?: string | null; defaults?: string } = {}): Uint8Array {
  const document = `<w:p><w:r><w:t>地板给的脸</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Themed"/></w:pPr><w:r><w:t>样式给的脸</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:ascii="Times New Roman" w:asciiTheme="minorHAnsi"/></w:rPr><w:t>字面先赢</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/><w:lang w:eastAsia="zh-CN"/></w:rPr><w:t>汉字有语言</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia"/></w:rPr><w:t>汉字无语言</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:eastAsiaTheme="minorEastAsia" w:hint="eastAsia"/></w:rPr><w:t>标记只说hint</w:t></w:r></w:p>
<w:p><w:r><w:rPr><w:rFonts w:asciiTheme="minorHAnsi" w:hAnsiTheme="minorHAnsi" w:eastAsiaTheme="minorEastAsia" w:cstheme="minorBidi"/><w:lang w:eastAsia="zh-CN"/></w:rPr><w:t>整行交给主题</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hint="eastAsia"/><w:lang w:eastAsia="zh-CN"/></w:rPr></w:pPr><w:r><w:t>标记的脸</w:t></w:r></w:p>
<w:p><w:pPr><w:rPr><w:rFonts w:asciiTheme="majorHAnsi" w:eastAsiaTheme="majorEastAsia" w:hint="eastAsia"/></w:rPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>编号的脸</w:t></w:r></w:p>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr>`;
  const theme = opts.theme === null ? undefined : (opts.theme ?? themeXml());
  const styles = opts.defaults ? STYLES.replace(DEFAULTS, opts.defaults) : STYLES;
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>` +
          `<Override PartName="/word/document.xml" ContentType="${CT_W}"/>` +
          `<Override PartName="/word/styles.xml" ContentType="${CT_STYLES}"/>` +
          `<Override PartName="/word/numbering.xml" ContentType="${CT_NUM}"/>` +
          (theme ? `<Override PartName="/word/theme/theme1.xml" ContentType="${CT_THEME}"/>` : "") +
          `</Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL + `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/_rels/document.xml.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/styles" Target="styles.xml"/>` +
          `<Relationship Id="rId2" Type="${NS_R}/numbering" Target="numbering.xml"/>` +
          (theme ? `<Relationship Id="rId3" Type="${NS_R}/theme" Target="theme/theme1.xml"/>` : "") +
          `</Relationships>`,
      ),
      "word/document.xml": strToU8(XML_DECL + `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>${document}</w:body></w:document>`),
      "word/styles.xml": strToU8(XML_DECL + styles),
      "word/numbering.xml": strToU8(XML_DECL + NUMBERING),
      ...(theme ? { "word/theme/theme1.xml": strToU8(theme) } : {}),
    },
    { level: 6 },
  );
}
