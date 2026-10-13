import { describe, expect, it } from "vitest";
import { OpcPackage } from "../src/core/opc.js";
import { WordDocument } from "../src/core/docx.js";
import type { Inline, Paragraph } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { renderInline } from "../src/core/render.js";
import {
  makeHalfAttributedLinkDocx,
  makeLinkDocx,
  makeSingleQuotedLinkDocx,
  makeSplitAnchorLinkDocx,
  makeTargetlessLinkDocx,
} from "./helpers/make-link.js";

const open = (bytes: Uint8Array): WordDocument => WordDocument.openBytes(bytes);
const body = (bytes: Uint8Array): string => OpcPackage.from(bytes).text("word/document.xml")!;
const linkRels = (bytes: Uint8Array): string[] => {
  const rels = OpcPackage.from(bytes).rels("word/document.xml");
  return [...rels.values()].filter((r) => r.type.endsWith("/hyperlink")).map((r) => r.target);
};
const links = (doc: WordDocument): string[] =>
  doc.paragraphs().flatMap((p) => p.inlines.filter((x: Inline) => x.hyperlink).map((x: Inline) => `${x.text}=${x.hyperlink}`));
/** Every `w:hyperlink` opening tag in the written part, in the order they appear. */
const linkTags = (bytes: Uint8Array): string[] => body(bytes).match(/<w:hyperlink[^>]*>/g) ?? [];

/**
 * A save of a document nobody edited returns the source bytes — the right fidelity rule, and the reason a
 * writer test that only opens and saves proves nothing: it reads back the fixture it was handed. Everything
 * asserted about how a link is *re-spelled* therefore has to go through a regeneration first, and one
 * character typed into the last paragraph is how the editor gets there.
 */
function regenerate(doc: WordDocument): WordDocument {
  const last = [...doc.blocks].reverse().find((b) => b.kind === "paragraph") as Paragraph | undefined;
  if (!last) throw new Error("fixture has no paragraph to edit");
  last.inlines.push({ text: "改" });
  return doc;
}
const rewritten = (bytes: Uint8Array): Uint8Array => writeDocx(regenerate(open(bytes)));

describe("docx links inside the document", () => {
  it("reads an in-document link as a target of its own, and an outside one as its URL", () => {
    expect(links(open(makeLinkDocx()))).toEqual([
      "跳到书签=#_Toc_1",
      "外部链接=https://example.com/detail",
      "修订里的锚点=#_Toc_1",
      "按名域取文字=#_Toc_1",
    ]);
  });

  it("writes an anchor back as `w:anchor` without inventing a relationship for it", () => {
    const out = rewritten(makeLinkDocx());
    expect(linkTags(out)).toContain('<w:hyperlink w:anchor="_Toc_1" w:history="1">');
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
    expect(linkTags(out)).toContain('<w:hyperlink w:anchor="_Toc_1" w:history="1">');
    expect(links(open(out))).toEqual(links(open(makeLinkDocx())));
    expect(open(out).paragraphs()[0]!.inlines.map((x) => x.text)).toContain("（打字进来的）");
    expect(linkRels(out)).toEqual(["https://example.com/detail"]);
  });

  it("keeps an anchor link inside a revision, where the wrapper sits outside the hyperlink", () => {
    expect(body(rewritten(makeLinkDocx()))).toContain('<w:ins w:id="7" w:author="甲"><w:hyperlink w:anchor="_Toc_1">');
  });

  it("keeps one link over runs that only differ by format", () => {
    const doc = open(makeSplitAnchorLinkDocx());
    expect(links(doc)).toEqual(["前半=#_Toc_2", "后半=#_Toc_2"]);
    const out = writeDocx(regenerate(doc));
    expect(linkTags(out)).toEqual(['<w:hyperlink w:anchor="_Toc_2">']);
    expect(links(open(out))).toEqual(["前半=#_Toc_2", "后半=#_Toc_2"]);
  });

  it("keeps an outside target on the relationship it already had", () => {
    const out = rewritten(makeLinkDocx());
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

describe("docx attributes a link element carries itself", () => {
  it("reads the ones the model has no field for, off the link and not off its runs", () => {
    const doc = open(makeLinkDocx());
    const at = (n: number): Inline | undefined =>
      doc.blocks[n]?.kind === "paragraph" ? doc.blocks[n]!.inlines.find((x) => x.linkAttrs) : undefined;
    expect(at(0)?.linkAttrs).toBe('w:history="1"');
    expect(at(1)?.linkAttrs).toBe('w:history="1"');
    expect(at(3)?.linkAttrs).toBe('w:docProp="title" w:tm="2024-01-01T00:00:00"');
    // The revision's link in the fixture carries nothing: it has to stay an attribute-free element.
    expect(at(2)?.linkAttrs).toBeUndefined();
  });

  it("writes them back onto the link, in the order the file gave and beside the target the model owns", () => {
    expect(linkTags(rewritten(makeLinkDocx()))).toEqual([
      '<w:hyperlink w:anchor="_Toc_1" w:history="1">',
      '<w:hyperlink r:id="rId5" w:history="1">',
      '<w:hyperlink w:anchor="_Toc_1">',
      '<w:hyperlink w:anchor="_Toc_1" w:docProp="title" w:tm="2024-01-01T00:00:00">',
    ]);
  });

  it("re-reads them from the file it wrote", () => {
    const before = open(makeLinkDocx());
    const out = writeDocx(regenerate(before));
    const attrs = (doc: WordDocument): string[] =>
      doc.blocks.flatMap((b) => (b.kind === "paragraph" ? b.inlines.map((x) => x.linkAttrs ?? "") : []));
    expect(attrs(open(out))).toEqual(attrs(before));
  });

  it("keeps them when somebody types into the paragraph that holds the link", () => {
    // The same trap as the target: the page hands the run back on every keystroke, and a run that comes
    // without these bytes writes a link that lost them.
    const doc = open(makeLinkDocx());
    const para = doc.blocks[0];
    if (para.kind !== "paragraph") throw new Error("block kind changed");
    const linked = para.inlines.find((x) => x.hyperlink);
    if (!linked) throw new Error("link gone from the fixture");
    para.inlines.splice(para.inlines.indexOf(linked) + 1, 0, {
      text: "（打字进来的）",
      hyperlink: linked.hyperlink,
      linkAttrs: linked.linkAttrs,
    });
    const out = writeDocx(doc);
    expect(linkTags(out)).toHaveLength(4); // the typed run joined the link instead of opening a second one
    expect(linkTags(out)[0]).toBe('<w:hyperlink w:anchor="_Toc_1" w:history="1">');
    const inner = /<w:hyperlink w:anchor="_Toc_1" w:history="1">([\s\S]*?)<\/w:hyperlink>/.exec(body(out))?.[1] ?? "";
    expect(inner).toContain("跳到书签");
    expect(inner).toContain("（打字进来的）");
  });

  it("stamps them on the page for the editor to hand back", () => {
    expect(renderInline({ text: "跳到书签", hyperlink: "#_Toc_1", linkAttrs: 'w:history="1"' })).toBe(
      '<a href="#_Toc_1" data-link-attrs="w:history=&quot;1&quot;"><span>跳到书签</span></a>',
    );
    // A link the editor made has no attributes of the file's to carry, and gets none invented.
    expect(renderInline({ text: "新建", hyperlink: "https://example.com/new" })).toBe(
      '<a href="https://example.com/new" target="_blank" rel="noreferrer"><span>新建</span></a>',
    );
  });

  it("does not merge two links that only share their target", () => {
    // Adjacent, same anchor, one with the attributes and one without: joining them by target alone would
    // put `w:history` on text the file never wrapped in it.
    const doc = open(makeHalfAttributedLinkDocx());
    const para = doc.blocks[0];
    if (para.kind !== "paragraph") throw new Error("block kind changed");
    expect(para.inlines.map((x) => `${x.text}|${x.hyperlink ?? ""}|${x.linkAttrs ?? ""}`)).toEqual([
      '带属性|#_Toc_3|w:history="1"',
      "不带属性|#_Toc_3|",
    ]);
    para.inlines.push({ text: "改" });
    expect(linkTags(writeDocx(doc))).toEqual([
      '<w:hyperlink w:anchor="_Toc_3" w:history="1">',
      '<w:hyperlink w:anchor="_Toc_3">',
    ]);
  });

  it("keeps an attribute the file wrote with the other XML quote style", () => {
    expect(linkTags(rewritten(makeSingleQuotedLinkDocx()))).toEqual([`<w:hyperlink w:anchor="_Toc_4" w:history='1'>`]);
  });

  it("leaves no orphan attributes on a link whose target the model cannot name", () => {
    // An empty `w:anchor` is not a place: the element is already plain runs, and bytes that belong to a
    // link would have to be invented onto a run.
    const doc = open(makeTargetlessLinkDocx());
    const para = doc.blocks[0];
    if (para.kind !== "paragraph") throw new Error("block kind changed");
    expect(para.inlines.map((x) => `${x.text}|${x.hyperlink ?? ""}|${x.linkAttrs ?? ""}`)).toEqual(["无目标||"]);
    para.inlines.push({ text: "改" });
    expect(linkTags(writeDocx(doc))).toEqual([]);
  });
});
