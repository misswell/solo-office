/**
 * What a save must not do to a file it was not asked to change.
 *
 * Every case here opens a package shaped like the ones the local WPS corpus writes (see
 * `tests/helpers/make-preservation.ts`), makes **one** edit, and then looks at what the save did to the parts
 * the edit never touched. The untouched save is not the measured path: `writeDocx` carries a document nobody
 * edited out byte for byte, so it proves nothing about a writer that has to rebuild a body.
 *
 * The rule these cases share is the one the rest of the writer follows: what the model holds travels from the
 * model, and what it does not travel as the bytes the file gave it. Losing the second half is how a save ends
 * up deleting the 24 page breaks, the 12 bookmark ends between table rows, and the frames around 16 pictures
 * that the census on real files counted.
 */
import { describe, expect, it } from "vitest";
import { OpcPackage } from "../src/core/opc.js";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import type { Inline } from "../src/core/docx.js";
import { makePreservationDocx } from "./helpers/make-preservation.js";

/** One added paragraph: enough to make the document dirty, so the save rebuilds the body from the model. */
function editAndSave(bytes: Uint8Array): Uint8Array {
  const doc = WordDocument.openBytes(bytes);
  doc.blocks.push({ kind: "paragraph", inlines: [{ text: "回归" }] });
  return writeDocx(doc);
}

const bodyOf = (bytes: Uint8Array): string => OpcPackage.from(bytes).text("word/document.xml") ?? "";
const countOf = (hay: string, needle: string): number => hay.split(needle).length - 1;

/** A second real PNG, so a picture the model changed can be told from the one the package came with. */
const PNG_RED = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";

const dataUrl = (base64: string): string => `data:image/png;base64,${base64}`;
const rawOf = (base64: string): Uint8Array => new Uint8Array(Buffer.from(base64, "base64"));

const pictureInlines = (doc: WordDocument): Inline[] =>
  doc.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines.filter((i) => i.image) : []));

describe("docx 编辑保存：不得无中生有", () => {
  it("writes no comments part for a document that never had one", () => {
    const saved = OpcPackage.from(editAndSave(makePreservationDocx()));
    expect([...saved.parts.keys()]).not.toContain("word/comments.xml");
    expect(saved.text("word/_rels/document.xml.rels")).not.toContain("comments");
    // The body is what a reader trusts: no relationship and no override for a list of comments nobody has.
    expect(saved.text("[Content_Types].xml")).not.toContain("comments");
  });

  it("carries the content-type table rather than re-spelling it one part at a time", () => {
    const src = OpcPackage.from(makePreservationDocx());
    const saved = OpcPackage.from(editAndSave(makePreservationDocx()));
    // `.rels` and `.xml` are promised by **extension**, which covers every part of that kind in the package.
    // Rebuilding the table from the part list stamps an Override on each one, and drops a Default the source
    // wrote for a kind this save no longer happens to hold.
    expect(saved.text("[Content_Types].xml")).toBe(src.text("[Content_Types].xml"));
    expect(saved.text("[Content_Types].xml")).not.toContain('PartName="/_rels/.rels"');
  });
});

describe("docx 编辑保存：部件与图片", () => {
  it("keeps the media part a body that still cites it depends on", () => {
    const src = OpcPackage.from(makePreservationDocx());
    const saved = OpcPackage.from(editAndSave(makePreservationDocx()));
    expect([...saved.parts.keys()]).toContain("word/media/trend.png");
    expect(saved.parts.get("word/media/trend.png")).toEqual(src.parts.get("word/media/trend.png"));
    // A part that went missing would leave every `r:embed` in the body pointing at nothing.
    const back = WordDocument.openBytes(editAndSave(makePreservationDocx()));
    for (const inline of pictureInlines(back)) expect(inline.image?.src).toBeTruthy();
  });

  it("writes a picture the model holds from the model's own bytes", () => {
    const doc = WordDocument.openBytes(makePreservationDocx());
    const first = pictureInlines(doc)[0];
    expect(first?.image).toBeDefined();
    first!.image = { ...first!.image!, src: dataUrl(PNG_RED) };
    const saved = OpcPackage.from(writeDocx(doc));
    expect([...saved.parts.get("word/media/trend.png")!]).toEqual([...rawOf(PNG_RED)]);
  });

  it("points a carried drawing at the relationship id this save issued", () => {
    const bytes = editAndSave(makePreservationDocx());
    const saved = OpcPackage.from(bytes);
    const rels = saved.text("word/_rels/document.xml.rels") ?? "";
    const id = /Id="([^"]+)"[^>]*Target="media\/trend\.png"/.exec(rels)?.[1];
    expect(id).toBeTruthy();
    // Whatever the source called it, the drawing now cites the id that resolves in *this* package — and does
    // not keep the old one, which would point at a relationship the save never wrote.
    for (const embed of (bodyOf(bytes).match(/r:embed="([^"]*)"/g) ?? []).map((m) => /="([^"]*)"/.exec(m)![1])) {
      expect(embed).toBe(id);
    }
  });
});

describe("docx 编辑保存：断、书签与图形内层", () => {
  it("keeps a page break and a column break what they were", () => {
    const before = bodyOf(makePreservationDocx());
    const after = bodyOf(editAndSave(makePreservationDocx()));
    expect(countOf(before, "<w:br ")).toBe(2);
    expect(after).toContain('<w:br w:type="page"/>');
    expect(after).toContain('<w:br w:type="column" w:clear="none"/>');
    expect(countOf(after, "<w:br ")).toBe(2);
    // Reading it back has to see two breaks, not one break and a paragraph of nothing.
    const doc = WordDocument.openBytes(editAndSave(makePreservationDocx()));
    const breaks = doc.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines.filter((i) => i.break) : []));
    expect(breaks.map((i) => i.brXml)).toEqual(['<w:br w:type="page"/>', '<w:br w:type="column" w:clear="none"/>']);
  });

  it("keeps a bookmark pair balanced when its end sits between two rows", () => {
    const after = bodyOf(editAndSave(makePreservationDocx()));
    expect(countOf(after, "<w:bookmarkStart ")).toBe(1);
    expect(countOf(after, "<w:bookmarkEnd ")).toBe(1);
    // The end belongs to the table, between the rows — not to the last cell, and not gone.
    const firstRow = after.indexOf("</w:tr>");
    const secondRow = after.indexOf("<w:tr", firstRow);
    const end = after.indexOf("<w:bookmarkEnd ");
    expect(end).toBeGreaterThan(firstRow);
    expect(end).toBeLessThan(secondRow);
  });

  it("keeps a floating picture floating, with the frame the file wrote", () => {
    const src = bodyOf(makePreservationDocx());
    const after = bodyOf(editAndSave(makePreservationDocx()));
    for (const shape of [
      "<wp:anchor",
      "wp:effectExtent",
      "<wp:positionH",
      "<wp:posOffset",
      "<a:srcRect",
      "<a:ln ",
      "<a:noFill/>",
      "<a:extLst>",
      "a14:useLocalDpi",
      "<wp:inline ",
    ]) {
      expect(countOf(after, shape)).toBe(countOf(src, shape));
    }
    // The anchor must not come back as an inline picture: the shape that decided where it sits is gone, and
    // the text no longer flows around it.
    expect(countOf(after, "<wp:anchor")).toBe(1);
  });

  it("re-spells a drawing once the picture's own size has moved", () => {
    const doc = WordDocument.openBytes(makePreservationDocx());
    const first = pictureInlines(doc)[0];
    first!.image = { ...first!.image!, width: 200 };
    const after = bodyOf(writeDocx(doc));
    // 200 px at 96 dpi is 1,905,000 EMU; the carried bytes would still say 1,219,200.
    expect(after).toContain('cx="1905000"');
    expect(after).not.toContain("<a:srcRect");
  });
});

describe("docx 编辑保存：正文的写法", () => {
  it("writes xml:space only where the text would lose an edge", () => {
    const after = bodyOf(editAndSave(makePreservationDocx()));
    expect(after).toContain("<w:t>正文</w:t>");
    expect(after).toContain('<w:t xml:space="preserve"> 带边空格的 </w:t>');
  });
});
