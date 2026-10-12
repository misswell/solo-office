import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makeDocx } from "./helpers/make.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);

describe("docx round trip", () => {
  it("re-reads everything it wrote", () => {
    const doc = open(makeDocx());
    const back = open(writeDocx(doc));
    expect(back.blocks).toEqual(doc.blocks);
    expect(back.sections).toEqual(doc.sections);
    expect(back.title).toBe(doc.title);
  });

  it("keeps the parts the model does not own byte-identical", () => {
    const doc = open(makeDocx());
    const out = writeDocx(doc);
    const pkg = OpcPackage.from(out);
    expect(pkg.kind).toBe("docx");
    for (const part of doc.untouchedParts()) {
      expect(text(pkg.bytes(part.path)!)).toBe(text(part.bytes));
    }
    expect(pkg.bytes("word/media/trend.png")).toEqual(OpcPackage.from(makeDocx()).bytes("word/media/trend.png"));
    expect(pkg.bytes("docProps/core.xml")).toBeDefined();
  });

  it("wires styles, numbering, image and hyperlink relationships", () => {
    const pkg = OpcPackage.from(writeDocx(open(makeDocx())));
    const rels = pkg.rels("word/document.xml");
    const targets = [...rels.values()].map((r) => `${r.type.split("/").pop()}:${r.target}`);
    expect(targets).toContain("styles:styles.xml");
    expect(targets).toContain("numbering:numbering.xml");
    expect(targets).toContain("image:media/trend.png");
    expect(targets).toContain("hyperlink:https://example.com/detail");

    const body = pkg.text("word/document.xml")!;
    const imageRid = /r:embed="(rId\d+)"/.exec(body)?.[1];
    const linkRid = /<w:hyperlink r:id="(rId\d+)"/.exec(body)?.[1];
    expect(rels.get(imageRid!)?.target).toBe("media/trend.png");
    expect(rels.get(linkRid!)?.target).toBe("https://example.com/detail");

    const types = pkg.contentType("word/styles.xml");
    expect(types).toBe("application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml");
    expect(pkg.contentType("word/media/trend.png")).toBe("image/png");
  });

  it("keeps a trailing table trailing", () => {
    const doc = new WordDocument();
    doc.blocks.push({
      kind: "table",
      columns: 1,
      rows: [[{ paragraphs: [{ kind: "paragraph", pageBreakBefore: false, inlines: [{ text: "只有表格" }] }] }]],
    });
    const back = open(writeDocx(doc));
    expect(back.blocks.map((b) => b.kind)).toEqual(["table"]);
  });

  it("carries an explicit un-bold through a style that is bold", () => {
    const doc = open(makeDocx());
    const heading = doc.blocks.find((b) => b.kind === "paragraph" && b.styleId === "Heading1");
    if (heading?.kind !== "paragraph") throw new Error("fixture lost its heading");
    heading.inlines = [{ text: "不再加粗", bold: false }];
    const out = writeDocx(doc);
    expect(OpcPackage.from(out).text("word/document.xml")).toContain('<w:b w:val="0"/>');
    const back = open(out);
    const reread = back.blocks.find((b) => b.kind === "paragraph" && b.styleId === "Heading1");
    expect(reread?.kind === "paragraph" ? reread.inlines[0].bold : undefined).toBe(false);
  });

  it("edits made through the model survive the write", () => {
    const doc = open(makeDocx());
    doc.blocks.splice(1, 0, {
      kind: "paragraph",
      styleId: "Heading1",
      numPr: { numId: 1, ilvl: 0 },
      inlines: [{ text: "新增", bold: true }, { text: "要点" }],
    });
    const back = open(writeDocx(doc));
    const fresh = back.blocks[1];
    if (fresh.kind !== "paragraph") throw new Error("block order changed");
    expect(fresh.inlines.map((i) => i.text)).toEqual(["新增", "要点"]);
    expect(fresh.inlines[0].bold).toBe(true);
    // numId 1 is the fixture's bullet list, and the definition has to come along with it.
    expect(fresh.numPr).toEqual({ numId: 1, ilvl: 0 });
    expect(back.levelOf(1, 0)?.numFmt).toBe("bullet");
    expect(back.blocks.map((b) => b.kind)).toEqual(doc.blocks.map((b) => b.kind));
  });
});

describe("a document created from nothing", () => {
  it("round-trips through its own writer and reader", () => {
    const doc = WordDocument.blank();
    expect(doc.blocks).toHaveLength(1);
    expect(doc.source()).toBeUndefined();
    expect(doc.edited).toBe(false);

    const first = doc.blocks[0];
    if (first.kind !== "paragraph") throw new Error("blank document did not start with a paragraph");
    first.inlines.push({ text: "第一次输入", bold: true });
    const out = writeDocx(doc);
    const parts = [...OpcPackage.from(out).parts.keys()];
    expect(parts).toContain("word/document.xml");
    expect(parts).toContain("_rels/.rels");

    const back = open(out);
    expect(back.blocks).toHaveLength(1);
    const para = back.blocks[0];
    if (para.kind !== "paragraph") throw new Error("block kind changed");
    expect(para.inlines.map((i) => i.text)).toEqual(["第一次输入"]);
  });
});
