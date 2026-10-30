import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { TOC_SDT, makeTocDocx, makeMixedSdtDocx } from "./helpers/make-carry-toc.js";

/**
 * 目录（TOC）的逐行显示（#236）。
 *
 * 受众先数（`.scratch` 语料普查）：授权 46 份里带目录的 **3 份**——business/event/project plan.dotm，
 * `w:docPartGallery` 都说 "Table of Contents"；段上不用 TOC 样式名，位置从 `toc 1/2` 样式链借（toc 2 借
 * `w:ind left=420`），「点线到右沿的页码」是段自己那句 `w:tabs`（右 8306、`w:leader="dot"`），条目字与页码
 * 是 HYPERLINK/PAGEREF 复杂域的**缓存结果**（`w:fldChar` 边界与 `w:instrText` 指令不是文字）。此前的画法
 * 是整块一段流文（`renderCarried` 只画 `text`）——装机截图对比里与 WPS 差得最远的那一格。这一轮按段逐行画：
 * 仍是只带盒子（字节原样回写），可编辑性与回写路径一字不动。
 */

const openDoc = () => WordDocument.openBytes(makeTocDocx());

describe("docx 只带容器：目录逐段的显示账", () => {
  it("sdtContent 枚枚是段 ⇒ paras 逐段给出；每段的字把域缓存收进来、指令排除在外", () => {
    const doc = openDoc();
    const carried = doc.blocks[0];
    expect(carried.kind).toBe("carried");
    if (carried.kind !== "carried") return;
    expect(carried.paras?.length).toBe(3);
    const [title, first, second] = carried.paras!;
    expect(title?.text).toBe("目 录");
    expect(title?.align).toBe("center");
    expect(title?.runs?.[0]).toMatchObject({ text: "目 录", bold: true, szHalf: 36, eastAsia: "宋体" });
    expect(first?.text).toBe("1、市场分析\t1");
    expect(first?.tabStops).toEqual([{ posTwips: 8306, val: "right", leader: "dot" }]);
    expect(second?.text).toBe("2、活动主题\t2");
    // toc 2 从链上借首行缩进（w:ind left=420），条目自己一个字没说缩进。
    expect(second?.ind).toMatchObject({ indentLeft: 420, indentLeftChars: 200 });
    // 三个 run：条目字（加粗）、制表符、页码——HYPERLINK/PAGEREF 的边界与指令一枚都不在。
    expect(first?.runs?.map((r) => r.text)).toEqual(["1、市场分析", "\t", "1"]);
    expect(first?.runs?.every((r) => !/HYPERLINK|PAGEREF/.test(r.text))).toBe(true);
  });

  it("sdtContent 里混着表 ⇒ 不逐行，退回整块一段的旧画法", () => {
    const doc = WordDocument.openBytes(makeMixedSdtDocx());
    const carried = doc.blocks[0];
    if (carried.kind !== "carried") throw new Error("第一块应是 carried");
    expect(carried.paras).toBeUndefined();
    expect(carried.text).toContain("盒子里一段");
  });
});

describe("docx 只带容器：目录逐行的纸面", () => {
  const html = renderWord(openDoc());

  it("一行一条：三枚 .carry-p，条目位上墙、制表符进 span，标题居中", () => {
    const lines = [...html.matchAll(/<div class="carry-p"[^>]*>[\s\S]*?<\/div>/g)].map((m) => m[0]);
    expect(lines.length).toBe(3);
    expect(lines[0]).toContain('style="text-align:center"');
    expect(lines[0]).toContain("目 录");
    expect((html.match(/data-tab-stops="8306:r:dot"/g) ?? []).length).toBe(2);
    expect((html.match(/<span class="tab"><\/span>/g) ?? []).length).toBe(2);
  });

  it("run 的字面上脸：标题宋体加粗 18pt（字族走 eastAsiaFaceRules 的 solo-ea 键），条目加粗", () => {
    expect(html).toContain("font-weight:bold");
    expect(html).toContain("font-size:18pt");
    expect(html).toContain("solo-ea-宋体");
  });

  it("盒子还是只带的：contenteditable=false、旧的一句灰字没有回来", () => {
    expect(html).toMatch(/<div class="doc-carry" contenteditable="false" data-tag="sdt"/);
    expect(html).not.toContain("（此处内容暂不可编辑）");
  });
});

describe("docx 只带容器：字节往返不动", () => {
  it("写回的件里，目录 sdt 逐字节与读进来时同一串", () => {
    const doc = openDoc();
    const out = writeDocx(doc);
    const saved = text(OpcPackage.from(out).bytes("word/document.xml")!);
    expect(saved).toContain(TOC_SDT);
  });
});
