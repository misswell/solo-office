import { strToU8, zipSync } from "fflate";

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';
const NS_W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const NS_CT = "http://schemas.openxmlformats.org/package/2006/content-types";
const NS_PKG = "http://schemas.openxmlformats.org/package/2006/relationships";

/** 单节 A4：`w:pgSz` 说 11906×16838 缇，两枚按纸定档的腿都从这两个数算档。 */
export const PAPER_WIDTH_TWIPS = 11906;
export const PAPER_HEIGHT_TWIPS = 16838;
/** 正文段数：够铺过一张纸，所以「单页」那一档量的是一页，不是全文。 */
export const PARA_COUNT = 40;

/**
 * 视图页 `RB_Zoom` 那两枚按纸定档的腿（WPS `ongmani/CT_View.kuip` 的 `ZoomPageWidth` / `ZoomOnePage`）
 * 用的件：单节 A4、`PARA_COUNT` 段正文。真 Chrome 探针（`.scratch/zoom-fit-dom.ts`）与装机壳
 * （`.scratch/make-zoom-fit-smoke.ts`）这两趟共用同一件，免得两处的纸不一样却报同一组数；
 * `tests/doc-zoom-fit.test.ts` 钉式子时读的也是上面那两个缇数。件名带「视图档」，壳里只有那一腿认领它。
 */
export function makeZoomFitDocx(): Uint8Array {
  return zipSync(
    {
      "[Content_Types].xml": strToU8(
        XML_DECL +
          `<Types xmlns="${NS_CT}"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
      ),
      "_rels/.rels": strToU8(
        XML_DECL +
          `<Relationships xmlns="${NS_PKG}"><Relationship Id="rId1" Type="${NS_R}/officeDocument" Target="word/document.xml"/></Relationships>`,
      ),
      "word/document.xml": strToU8(
        XML_DECL +
          `<w:document xmlns:w="${NS_W}" xmlns:r="${NS_R}"><w:body>` +
          `<w:p><w:r><w:t>甲段</w:t></w:r></w:p>`.repeat(PARA_COUNT) +
          `<w:sectPr><w:pgSz w:w="${PAPER_WIDTH_TWIPS}" w:h="${PAPER_HEIGHT_TWIPS}"/><w:pgMar w:top="1440" w:right="1800" w:bottom="1440" w:left="1800"/></w:sectPr></w:body></w:document>`,
      ),
    },
    { level: 6 },
  );
}
