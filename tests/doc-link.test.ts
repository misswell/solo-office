import { describe, expect, it } from "vitest";
import { OpcPackage } from "../src/core/opc.js";
import { WordDocument } from "../src/core/docx.js";
import type { Inline } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderInline } from "../src/core/render.js";
import { makeLinkDocx, makeSplitAnchorLinkDocx } from "./helpers/make-link.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const body = (bytes: Uint8Array): string => OpcPackage.from(bytes).text("word/document.xml")!;
const linkRels = (bytes: Uint8Array): string[] => {
  const rels = OpcPackage.from(bytes).rels("word/document.xml");
  return [...rels.values()].filter((r) => r.type.endsWith("/hyperlink")).map((r) => r.target);
};
const links = (doc: WordDocument): string[] =>
  doc.paragraphs().flatMap((p) => p.inlines.filter((x: Inline) => x.hyperlink).map((x: Inline) => `${x.text}=${x.hyperlink}`));

describe("docx links inside the document", () => {
  it("reads an in-document link as a target of its own, and an outside one as its URL", () => {
    expect(links(open(makeLinkDocx()))).toEqual([
      "跳到书签=#_Toc_1",
      "外部链接=https://example.com/detail",
      "修订里的锚点=#_Toc_1",
    ]);
  });

  it("writes an anchor back as `w:anchor` without inventing a relationship for it", () => {
    const out = writeDocx(open(makeLinkDocx()));
    expect(body(out)).toContain('<w:hyperlink w:anchor="_Toc_1">');
    expect(linkRels(out)).toEqual(["https://example.com/detail"]);
    expect(links(open(out))).toEqual(links(open(makeLinkDocx())));
  });

  it("keeps the link when somebody types into the paragraph that holds it", () => {
    // The whole reason the model has to name an anchor: with no field for it the paragraph's bytes only
    // travel intact while nobody edits, and one keystroke turned a table-of-contents entry into plain text.
    const doc = open(makeLinkDocx());
    const para = doc.blocks[0];
    if (para.kind !== "paragraph") throw new Error("block kind changed");
    para.inlines.push({ text: "（打字进来的）" });
    const out = writeDocx(doc);
    expect(body(out)).toContain('<w:hyperlink w:anchor="_Toc_1">');
    expect(links(open(out))).toEqual(links(open(makeLinkDocx())));
    expect(open(out).paragraphs()[0]!.inlines.map((x) => x.text)).toContain("（打字进来的）");
    expect(linkRels(out)).toEqual(["https://example.com/detail"]);
  });

  it("keeps an anchor link inside a revision, where the wrapper sits outside the hyperlink", () => {
    const out = writeDocx(open(makeLinkDocx()));
    expect(body(out)).toContain('<w:ins w:id="7" w:author="甲"><w:hyperlink w:anchor="_Toc_1">');
  });

  it("keeps one link over runs that only differ by format", () => {
    const doc = open(makeSplitAnchorLinkDocx());
    expect(links(doc)).toEqual(["前半=#_Toc_2", "后半=#_Toc_2"]);
    const out = writeDocx(doc);
    expect(body(out).match(/<w:hyperlink w:anchor="_Toc_2">/g)).toHaveLength(1);
    expect(links(open(out))).toEqual(["前半=#_Toc_2", "后半=#_Toc_2"]);
  });

  it("keeps an outside target on the relationship it already had", () => {
    const out = writeDocx(open(makeLinkDocx()));
    const rid = /<w:hyperlink r:id="(rId\d+)"/.exec(body(out))?.[1];
    expect(rid).toBe("rId5");
    expect(OpcPackage.from(out).rels("word/document.xml").get(rid!)?.target).toBe("https://example.com/detail");
    expect(linkRels(out)).toHaveLength(1);
  });

  it("does not send an in-document link to another window", () => {
    expect(renderInline({ text: "跳到书签", hyperlink: "#_Toc_1" })).toBe('<a href="#_Toc_1"><span>跳到书签</span></a>');
    expect(renderInline({ text: "外面", hyperlink: "https://example.com/detail" })).toBe(
      '<a href="https://example.com/detail" target="_blank" rel="noreferrer"><span>外面</span></a>',
    );
  });
});
