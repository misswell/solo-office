import { describe, expect, it } from "vitest";
import { Presentation } from "../src/core/pptx.js";
import { writePptx } from "../src/core/write-pptx.js";
import { OpcPackage, text } from "../src/core/opc.js";
import { makePptx } from "./helpers/make.js";

const open = (bytes: Uint8Array): Presentation => Presentation.openBytes(bytes);

describe("pptx round trip", () => {
  it("re-reads everything it wrote", () => {
    const pres = open(makePptx());
    const back = open(writePptx(pres));
    expect(back.slides).toHaveLength(1);
    expect(back.slideWidthEmu).toBe(pres.slideWidthEmu);
    const names = (p: Presentation): (string | undefined)[] => p.slides[0].shapes.map((s) => s.name);
    expect(names(back)).toEqual(names(pres));
    const table = back.slides[0].shapes.find((s) => s.kind === "table");
    expect(table?.rows?.flat().map((c) => c.text)).toEqual(["指标", "达成", "收入", "108%"]);
    expect(back.slides[0].notes).toContain("口头补充");
  });

  it("copies every untouched part through byte-identical", () => {
    const pres = open(makePptx());
    const out = writePptx(pres);
    const src = OpcPackage.from(makePptx());
    const pkg = OpcPackage.from(out);
    for (const path of ["ppt/theme/theme1.xml", "ppt/media/trend.png", "ppt/presentation.xml", "ppt/slides/slide1.xml", "ppt/notesSlides/notesSlide1.xml"]) {
      expect(pkg.bytes(path), path).toEqual(src.bytes(path));
    }
    expect(pkg.kind).toBe("pptx");
  });

  it("patches moved geometry into the slide XML", () => {
    const pres = open(makePptx());
    const title = pres.slides[0].shapes.find((s) => s.name === "标题")!;
    title.x = 100;
    title.y = 50;
    title.width = 600;
    title.height = 90;
    pres.markDirty(pres.slides[0]);
    const pkg = OpcPackage.from(writePptx(pres));
    const xml = pkg.text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('x="952500"');
    expect(xml).toContain('y="476250"');
    const back = open(writePptx(pres));
    const moved = back.slides[0].shapes.find((s) => s.name === "标题")!;
    expect(moved.x).toBeCloseTo(100, 3);
    expect(moved.width).toBeCloseTo(600, 3);
  });

  it("rebuilds edited text with its run styles", () => {
    const pres = open(makePptx());
    const ellipse = pres.slides[0].shapes.find((s) => s.name === "椭圆")!;
    ellipse.paragraphs = [{ level: 0, align: "ctr", runs: [{ text: "+25.5%", bold: true, sizePt: 24, color: "#C00000" }] }];
    ellipse.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const back = open(writePptx(pres));
    const run = back.slides[0].shapes.find((s) => s.name === "椭圆")!.paragraphs![0].runs[0];
    expect(run.text).toBe("+25.5%");
    expect(run).toMatchObject({ bold: true, sizePt: 24, color: "#C00000" });
  });

  it("edits table cell text", () => {
    const pres = open(makePptx());
    const table = pres.slides[0].shapes.find((s) => s.kind === "table")!;
    table.rows![1][1] = { text: "125%", runs: [{ text: "125%" }] };
    table.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const back = open(writePptx(pres));
    const cell = back.slides[0].shapes.find((s) => s.kind === "table")!.rows![1][1];
    expect(cell.text).toBe("125%");
  });

  it("rewrites only the cell the user edited", () => {
    const pres = open(makePptx());
    const table = pres.slides[0].shapes.find((s) => s.kind === "table")!;
    table.rows![0][0] = { text: "季度达成", runs: [{ text: "季度达成" }] };
    table.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const xml = OpcPackage.from(writePptx(pres)).text("ppt/slides/slide1.xml")!;
    expect(xml).toContain("<a:p><a:r><a:t>季度达成</a:t></a:r></a:p>");
    expect(xml).toContain("<a:p><a:r><a:t>达成</a:t></a:r></a:p>");
    expect(xml).toContain("<a:p><a:r><a:t>108%</a:t></a:r></a:p>");
    expect(xml).not.toContain("<a:pPr/>");
    expect(xml).not.toContain("<a:rPr/>");
  });

  it("keeps an untouched shape's body exactly as read while another shape moves", () => {
    const pres = open(makePptx());
    const ellipse = pres.slides[0].shapes.find((s) => s.name === "椭圆")!;
    ellipse.x += 10;
    pres.markDirty(pres.slides[0]);
    const pkg = OpcPackage.from(writePptx(pres));
    const xml = pkg.text("ppt/slides/slide1.xml")!;
    expect(xml).toContain("<a:buChar char=");
    const src = text(OpcPackage.from(makePptx()).bytes("ppt/slides/slide1.xml")!);
    const bodyOf = (s: string): string => s.slice(s.indexOf("<p:txBody>"), s.indexOf("</p:txBody>"));
    expect(xml).toContain(bodyOf(src));
  });

  it("adds and deletes slides, updating the presentation list", () => {
    const pres = open(makePptx());
    pres.addSlide();
    pres.addSlide();
    expect(new Set(pres.slides.map((s) => s.partPath)).size).toBe(3);
    expect(new Set(pres.slides.map((s) => s.relId)).size).toBe(3);
    let back = open(writePptx(pres));
    expect(back.slides).toHaveLength(3);
    expect(back.slides.map((s) => s.sldId)).toEqual(["256", "257", "258"]);
    expect(back.slides[2].shapes).toEqual([]);

    back.removeSlide(0);
    back.removeSlide(1);
    const out = writePptx(back);
    const reread = open(out);
    expect(reread.slides).toHaveLength(1);
    const pkg = OpcPackage.from(out);
    expect(pkg.bytes("ppt/slides/slide1.xml")).toBeUndefined();
    expect(pkg.bytes("ppt/notesSlides/notesSlide1.xml")).toBeUndefined();
    expect(pkg.overrideTypes.get("/ppt/slides/slide1.xml")).toBeUndefined();
    expect(pkg.text("[Content_Types].xml")).not.toContain("slide1.xml");
  });

  it("gives a new slide a layout to inherit from", () => {
    const pres = open(makePptx());
    pres.addSlide();
    const pkg = OpcPackage.from(writePptx(pres));
    const rels = pkg.text("ppt/slides/_rels/slide2.xml.rels")!;
    expect(rels).toContain("<Relationships");
    expect(rels).toContain("/slideLayout");
    expect(rels).toContain('Target="../slideLayouts/slideLayout1.xml"');
    expect(pkg.text("[Content_Types].xml")).toContain("/ppt/slideLayouts/slideLayout1.xml");
  });

  it("keeps a slide's numeric id apart from its relationship id", () => {
    const pres = open(makePptx());
    expect(pres.slides[0].sldId).toBe("256");
    expect(pres.slides[0].relId).toBe("rId1");
    pres.markDirty(pres.slides[0]);
    const xml = OpcPackage.from(writePptx(pres)).text("ppt/presentation.xml")!;
    expect(xml).toContain('id="256"');
    expect(xml).toContain('r:id="rId1"');
  });

  it("writes soft line breaks as breaks and reads them back", () => {
    const pres = open(makePptx());
    const ellipse = pres.slides[0].shapes.find((s) => s.name === "椭圆")!;
    ellipse.paragraphs = [{ level: 0, runs: [{ text: "第一行\n第二行" }] }];
    ellipse.textDirty = true;
    pres.markDirty(pres.slides[0]);
    const out = writePptx(pres);
    expect(OpcPackage.from(out).text("ppt/slides/slide1.xml")).toContain("<a:br/>");
    const back = open(out);
    const runs = back.slides[0].shapes.find((s) => s.name === "椭圆")!.paragraphs![0].runs;
    expect(runs.map((r) => r.text).join("")).toBe("第一行\n第二行");
  });

  it("reorders slides and keeps their content", () => {
    const pres = open(makePptx());
    pres.addSlide();
    pres.slides[1].shapes.push({ kind: "shape", name: "新形状", x: 10, y: 20, width: 100, height: 50, paragraphs: [{ level: 0, runs: [{ text: "第二页" }] }] });
    pres.markDirty(pres.slides[1]);
    const back = open(writePptx(pres));
    expect(back.slides[1].shapes.map((s) => s.name)).toContain("新形状");
    back.moveSlide(1, -1);
    const reread = open(writePptx(back));
    expect(reread.slides[0].shapes.map((s) => s.name)).toContain("新形状");
    expect(reread.slides[1].shapes.map((s) => s.name)).toContain("标题");
  });

  it("drops a deleted shape's element from the slide", () => {
    const pres = open(makePptx());
    const slide = pres.slides[0];
    slide.shapes = slide.shapes.filter((s) => s.name !== "椭圆");
    pres.markDirty(slide);
    const out = writePptx(pres);
    const xml = OpcPackage.from(out).text("ppt/slides/slide1.xml")!;
    expect(xml).not.toContain("椭圆");
    expect(xml).toContain("季度经营复盘");
    const back = open(out);
    expect(back.slides[0].shapes.map((s) => s.name)).toEqual(["标题", "正文", "图", "表"]);
  });

  it("stores edited notes", () => {
    const pres = open(makePptx());
    pres.slides[0].notes = "改好的备注\n第二行";
    pres.slides[0].notesDirty = true;
    const back = open(writePptx(pres));
    expect(back.slides[0].notes).toContain("改好的备注");
    expect(back.slides[0].notes).toContain("第二行");
  });
});

describe("a deck created from nothing", () => {
  it("round-trips through its own writer and reader", () => {
    const pres = Presentation.blank();
    expect(pres.slides).toHaveLength(1);
    expect(pres.slides[0].partPath).toBe("ppt/slides/slide1.xml");

    const out = writePptx(pres);
    const pkg = OpcPackage.from(out);
    expect([...pkg.parts.keys()]).toContain("_rels/.rels");
    expect(pkg.text("_rels/.rels")).toContain('Target="ppt/presentation.xml"');

    const back = open(out);
    expect(back.slides).toHaveLength(1);
    expect(back.slides[0].shapes).toEqual([]);
    expect(back.slides[0].name).toBe("第 1 页");

    const added = back.addSlide();
    added.shapes.push({ kind: "shape", name: "标题", x: 10, y: 20, width: 300, height: 60, paragraphs: [{ level: 0, runs: [{ text: "新页标题" }] }] });
    back.markDirty(added);
    const two = open(writePptx(back));
    expect(two.slides.map((s) => s.shapes.map((sh) => sh.name))).toEqual([[], ["标题"]]);
  });
});

describe("inserting shapes and pictures", () => {
  const PNG = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8DwHwAFAAH/q842iQAAAABJRU5ErkJggg==";

  it("writes an inserted text box as a p:sp with txBox", () => {
    const pres = Presentation.blank();
    pres.slides[0].shapes.push({
      kind: "shape",
      textBox: true,
      name: "文本框",
      geometry: "rect",
      x: 100,
      y: 80,
      width: 320,
      height: 60,
      rotation: 30,
      paragraphs: [{ level: 0, runs: [{ text: "新加的文本", sizePt: 20, color: "#C00000" }] }],
    });
    pres.markDirty(pres.slides[0]);

    const out = writePptx(pres);
    const xml = OpcPackage.from(out).text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('txBox="1"');
    expect(xml).toContain('rot="1800000"');
    expect(xml).toContain("新加的文本");

    const back = open(out);
    const box = back.slides[0].shapes[0];
    expect(box.kind).toBe("shape");
    expect(box.textBox).toBe(true);
    expect(box.width).toBeCloseTo(320, 3);
    expect(box.rotation).toBeCloseTo(30, 3);
    expect(box.paragraphs?.[0].runs[0]).toMatchObject({ text: "新加的文本", sizePt: 20, color: "#C00000" });
  });

  it("writes an inserted line as a zero-height shape", () => {
    const pres = Presentation.blank();
    pres.slides[0].shapes.push({ kind: "shape", name: "直线", geometry: "line", x: 40, y: 90, width: 240, height: 0, lineColor: "#4472C4", lineWidth: 2 });
    pres.markDirty(pres.slides[0]);
    const back = open(writePptx(pres));
    const line = back.slides[0].shapes[0];
    expect(line.geometry).toBe("line");
    expect(line.height).toBe(0);
    expect(line.lineWidth).toBeCloseTo(2, 3);
    expect(line.lineColor).toBe("#4472C4");
  });

  it("writes an inserted picture as a part, a relationship and a p:pic", () => {
    const pres = open(makePptx());
    const path = pres.addPicture(PNG);
    expect(path).toBe("ppt/media/image1.png");
    pres.slides[0].shapes.push({ kind: "picture", name: "图片", src: PNG, mediaPath: path, x: 700, y: 300, width: 120, height: 90 });
    pres.markDirty(pres.slides[0]);

    const out = writePptx(pres);
    const pkg = OpcPackage.from(out);
    expect(pkg.parts.has("ppt/media/image1.png")).toBe(true);
    // rId1 layout, rId2 image and rId9 notes belong to the source; the insert takes the lowest free id.
    const rels = pkg.text("ppt/slides/_rels/slide1.xml.rels")!;
    expect(rels).toContain('Id="rId3"');
    expect(rels).toContain('Target="../media/image1.png"');
    expect(rels).toContain('Target="../slideLayouts/slideLayout1.xml"');
    expect(rels).toContain('Target="../notesSlides/notesSlide1.xml"');
    expect(pkg.text("ppt/slides/slide1.xml")).toContain('r:embed="rId3"');
    expect(pkg.text("[Content_Types].xml")).toContain('Extension="png" ContentType="image/png"');

    const back = open(out);
    const pic = back.slides[0].shapes.find((s) => s.name === "图片")!;
    expect(pic.kind).toBe("picture");
    expect(pic.src).toBe(PNG);
    expect(pic.width).toBeCloseTo(120, 3);
  });

  it("generates the rels part a picture needs on a deck built from nothing", () => {
    const pres = Presentation.blank();
    const path = pres.addPicture(PNG);
    pres.slides[0].shapes.push({ kind: "picture", name: "图片", src: PNG, mediaPath: path, x: 0, y: 0, width: 100, height: 80 });
    pres.markDirty(pres.slides[0]);
    const out = writePptx(pres);
    const rels = OpcPackage.from(out).text("ppt/slides/_rels/slide1.xml.rels")!;
    expect(rels).toContain('Id="rId1"');
    expect(rels).toContain('Target="../media/image1.png"');
    expect(open(out).slides[0].shapes[0].src).toBe(PNG);
  });

  it("gives each inserted picture its own media part and relationship", () => {
    const pres = Presentation.blank();
    const first = pres.addPicture(PNG);
    const second = pres.addPicture(PNG);
    expect([first, second]).toEqual(["ppt/media/image1.png", "ppt/media/image2.png"]);
    pres.slides[0].shapes.push(
      { kind: "picture", name: "图一", src: PNG, mediaPath: first, x: 0, y: 0, width: 100, height: 80 },
      { kind: "picture", name: "图二", src: PNG, mediaPath: second, x: 10, y: 10, width: 100, height: 80 },
    );
    pres.markDirty(pres.slides[0]);
    const out = writePptx(pres);
    const rels = OpcPackage.from(out).text("ppt/slides/_rels/slide1.xml.rels")!;
    expect(rels).toContain('Target="../media/image1.png"');
    expect(rels).toContain('Target="../media/image2.png"');
    const back = open(out);
    expect(back.slides[0].shapes.map((s) => s.name)).toEqual(["图一", "图二"]);
    expect(back.slides[0].shapes[1].src).toBe(PNG);
  });
});
