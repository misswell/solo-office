import { describe, expect, it } from "vitest";
import { OpcPackage } from "../src/core/opc.js";
import { Presentation } from "../src/core/pptx.js";
import { renderSlideHtml } from "../src/core/render.js";
import { writePptx } from "../src/core/write-pptx.js";

const open = (bytes: Uint8Array): Presentation => Presentation.openBytes(bytes);
const pkgOf = (pres: Presentation): OpcPackage => OpcPackage.from(writePptx(pres));

describe("blank deck package", () => {
  it("writes the master, eleven layouts, theme and property parts WPS's blank deck has", () => {
    const pkg = pkgOf(Presentation.blank());
    expect(pkg.text("ppt/presentation.xml")).toContain('<p:sldMasterId id="2147483648" r:id="rId1"/>');
    expect(pkg.text("ppt/presentation.xml")).toContain('<p:sldSz cx="12192000" cy="6858000"/>');
    expect(pkg.text("ppt/slideMasters/slideMaster1.xml")).toContain('name="标题占位符 1"');
    expect(pkg.text("ppt/slideMasters/slideMaster1.xml")).toContain("<p:sldLayoutIdLst>");
    expect(pkg.text("ppt/theme/theme1.xml")).toContain('<a:srgbClr val="4874CB"/>');
    expect(pkg.text("ppt/theme/theme1.xml")).toContain('<a:ea typeface="微软雅黑"/>');
    expect(pkg.text("ppt/presProps.xml")).toContain("<p:presentationPr");
    expect(pkg.text("ppt/viewProps.xml")).toContain('<p:gridSpacing cx="72008" cy="72008"/>');
    expect(pkg.text("ppt/tableStyles.xml")).toContain("<a:tblStyleLst");
    expect(pkg.text("docProps/core.xml")).toContain("<cp:coreProperties");
    expect(pkg.text("docProps/app.xml")).toContain("<Slides>1</Slides>");
    const layouts = [...pkg.parts.keys()].filter((p) => /^ppt\/slideLayouts\/slideLayout\d+\.xml$/.test(p));
    expect(layouts).toHaveLength(11);
    expect(pkg.text("ppt/slideLayouts/slideLayout1.xml")).toContain('type="title"');
    expect(pkg.text("ppt/slideLayouts/slideLayout1.xml")).toContain('name="标题幻灯片"');
    expect(pkg.text("ppt/slideLayouts/slideLayout7.xml")).toContain('type="blank"');
    expect(pkg.text("ppt/slideLayouts/slideLayout9.xml")).toContain('orient="vert"');
    // WPS's three typeless layouts are user-drawn; PowerPoint hides untyped custom
    // layouts that do not say so, which would drop 图片与标题/内容/末尾幻灯片 from its picker.
    for (const n of [8, 10, 11]) expect(pkg.text(`ppt/slideLayouts/slideLayout${n}.xml`)).toContain('preserve="1" userDrawn="1"');
    expect(pkg.text("ppt/slideLayouts/slideLayout3.xml")).not.toContain("userDrawn");
  });

  it("wires the relationships the package needs to be a presentation", () => {
    const pkg = pkgOf(Presentation.blank());
    const presRels = pkg.text("ppt/_rels/presentation.xml.rels")!;
    expect(presRels).toContain('Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"');
    expect(presRels).toContain('Target="theme/theme1.xml"');
    expect(presRels).toContain('Target="viewProps.xml"');
    expect(presRels).toContain('Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide1.xml"');
    expect(pkg.text("ppt/slides/_rels/slide1.xml.rels")).toContain('Target="../slideLayouts/slideLayout1.xml"');
    const masterRels = pkg.text("ppt/slideMasters/_rels/slideMaster1.xml.rels")!;
    expect(masterRels).toContain('Id="rId11" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout" Target="../slideLayouts/slideLayout11.xml"');
    expect(masterRels).toContain('Id="rId12" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme1.xml"');
    expect(pkg.text("ppt/slideLayouts/_rels/slideLayout11.xml.rels")).toContain('Target="../slideMasters/slideMaster1.xml"');
    expect(pkg.text("_rels/.rels")).toContain('Target="docProps/core.xml"');
  });

  it("declares every generated part's content type", () => {
    const types = pkgOf(Presentation.blank()).text("[Content_Types].xml")!;
    for (const part of [
      "/ppt/slideMasters/slideMaster1.xml",
      "/ppt/slideLayouts/slideLayout11.xml",
      "/ppt/theme/theme1.xml",
      "/ppt/presProps.xml",
      "/ppt/viewProps.xml",
      "/ppt/tableStyles.xml",
      "/docProps/core.xml",
      "/docProps/app.xml",
    ]) {
      expect(types).toContain(`PartName="${part}"`);
    }
  });

  it("reads its own blank package back with the layouts it wrote", () => {
    const back = open(writePptx(Presentation.blank()));
    expect(back.slides).toHaveLength(1);
    expect(back.slideWidthEmu).toBe(12192000);
    expect(back.gridSpacingEmu).toBe(72008);
    expect(back.newSlideLayoutPath).toBe("ppt/slideLayouts/slideLayout1.xml");
    const added = back.addSlide();
    expect(added.relId).toBe("rId7");
    expect(open(writePptx(back)).slides).toHaveLength(2);
  });

  it("keeps notes parts free of a notes master the blank package never writes", () => {    const pres = Presentation.blank();
    pres.slides[0].notes = "写在空白件上的备注";
    pres.slides[0].notesDirty = true;
    const pkg = pkgOf(pres);
    expect(pkg.text("ppt/notesSlides/notesSlide1.xml")).toContain("写在空白件上的备注");
    expect(pkg.text("ppt/notesSlides/_rels/notesSlide1.xml.rels")).not.toContain("notesMaster");
  });
});

describe("blank deck placeholder prompts", () => {
  it("carries the two title-slide placeholders the layout declares", () => {
    const pres = Presentation.blank();
    const phs = pres.slides[0].layoutPlaceholders ?? [];
    expect(phs.map((p) => p.type)).toEqual(["ctrTitle", "subTitle"]);
    expect(phs[0].x).toBeCloseTo(1198800 / 9525, 3);
    expect(phs[0].height).toBeCloseTo(2570400 / 9525, 3);
  });

  it("prompts where no shape fills the placeholder, and only on editing surfaces", () => {
    const pres = Presentation.blank();
    const editing = renderSlideHtml(pres, 0, { placeholders: true });
    expect(editing).toContain("单击此处添加标题");
    expect(editing).toContain("单击此处添加副标题");
    expect(editing).toContain('data-ph="0"');
    expect(editing).toContain('data-ph="1"');
    expect(renderSlideHtml(pres, 0)).not.toContain("ph-hint");
    expect(renderSlideHtml(pres, 0)).not.toContain("单击此处");
  });

  it("writes a claimed placeholder as p:ph so a save keeps its identity", () => {
    const pres = Presentation.blank();
    pres.slides[0].shapes.push(
      { kind: "shape", name: "标题 1", x: 100, y: 100, width: 400, height: 80, placeholder: "ctrTitle", paragraphs: [] },
      { kind: "shape", name: "副标题 2", x: 100, y: 200, width: 400, height: 80, placeholder: "subTitle", placeholderIdx: 1, paragraphs: [] },
    );
    pres.markDirty(pres.slides[0]);
    const xml = pkgOf(pres).text("ppt/slides/slide1.xml")!;
    expect(xml).toContain('name="标题 1"');
    expect(xml).toContain('type="ctrTitle"');
    expect(xml).toContain('idx="1"');
    const back = open(writePptx(pres));
    const shapes = back.slides[0].shapes;
    expect(shapes.find((s) => s.name === "标题 1")?.placeholder).toBe("ctrTitle");
    expect(shapes.find((s) => s.name === "副标题 2")).toMatchObject({ placeholder: "subTitle", placeholderIdx: 1 });
  });

  it("drops a prompt once a shape claims that placeholder", () => {
    const pres = Presentation.blank();
    pres.slides[0].shapes.push({ kind: "shape", name: "标题 1", x: 100, y: 100, width: 400, height: 80, placeholder: "ctrTitle", paragraphs: [] });
    const html = renderSlideHtml(pres, 0, { placeholders: true });
    expect(html).not.toContain("单击此处添加标题");
    expect(html).toContain("单击此处添加副标题");
  });

  it("reads the placeholders back from the layout part it wrote", () => {
    const back = open(writePptx(Presentation.blank()));
    const phs = back.slides[0].layoutPlaceholders ?? [];
    expect(phs.map((p) => p.type)).toEqual(["ctrTitle", "subTitle"]);
    expect(phs[1].y).toBeCloseTo(3560400 / 9525, 3);
  });
});

describe("placeholder text inherits the layout's declared style", () => {
  it("carries the size, alignment and anchor the layout states", () => {
    const phs = Presentation.blank().slides[0].layoutPlaceholders ?? [];
    expect(phs[0]).toMatchObject({ sizePt: 60, align: "center", anchor: "b" });
    expect(phs[1]).toMatchObject({ sizePt: 24, align: "center" });
    expect(phs[1].anchor).toBeUndefined();
  });

  it("renders a claimed title at the layout's 60pt, centred, bottom-anchored", () => {
    const pres = Presentation.blank();
    pres.slides[0].shapes.push({
      kind: "shape",
      name: "标题 1",
      x: 125.9,
      y: 96,
      width: 1028.8,
      height: 269.9,
      placeholder: "ctrTitle",
      paragraphs: [{ level: 0, runs: [{ text: "季度汇报" }] }],
    });
    const html = renderSlideHtml(pres, 0);
    expect(html).toMatch(/class="shape-text" style="font-size:60pt;justify-content:flex-end"/);
    expect(html).toContain("text-align:center");
    expect(html).toContain("季度汇报");
  });

  it("leaves a silent layout to the master's size and alignment", () => {
    const pres = Presentation.blank();
    pres.slides[0].layoutPlaceholders = [
      { type: "title", x: 64, y: 64, width: 1152, height: 74, anchor: "ctr" },
    ];
    pres.slides[0].shapes.push({
      kind: "shape",
      name: "标题 2",
      x: 64,
      y: 64,
      width: 1152,
      height: 74,
      placeholder: "title",
      paragraphs: [{ level: 0, runs: [{ text: "分析" }] }],
    });
    const html = renderSlideHtml(pres, 0);
    expect(html).toMatch(/class="shape-text" style="font-size:36pt;justify-content:center"/);
    expect(html).not.toContain("text-align:");
  });

  it("draws an empty prompt at the placeholder's own size and alignment", () => {
    const html = renderSlideHtml(Presentation.blank(), 0, { placeholders: true });
    expect(html).toMatch(/data-ph="0"[^>]*font-size:60pt;justify-content:center;align-items:flex-end/);
    expect(html).toMatch(/data-ph="1"[^>]*font-size:24pt;justify-content:center;align-items:flex-start/);
  });

  it("reads the layout lstStyle and master txStyles back from the written package", () => {
    const back = open(writePptx(Presentation.blank()));
    const phs = back.slides[0].layoutPlaceholders ?? [];
    expect(phs[0]).toMatchObject({ sizePt: 60, align: "center", anchor: "b" });
    expect(phs[1]).toMatchObject({ sizePt: 24, align: "center" });
    expect(back.titleStylePt).toBe(36);
    expect(back.bodyStylePt).toBe(18);
    const pkg = pkgOf(Presentation.blank());
    expect(pkg.text("ppt/slideMasters/slideMaster1.xml")).toContain('<a:defRPr sz="3600"');
    expect(pkg.text("ppt/slideLayouts/slideLayout1.xml")).toContain('<a:lvl1pPr algn="ctr"><a:defRPr sz="6000"/>');
  });
});

describe("new slides choose a layout", () => {
  it("offers the eleven blank layouts under the names WPS gives them", () => {
    const choices = Presentation.blank().layoutChoices();
    expect(choices).toHaveLength(11);
    expect(choices[0]).toEqual({ path: "ppt/slideLayouts/slideLayout1.xml", name: "标题幻灯片" });
    expect(choices[2]).toEqual({ path: "ppt/slideLayouts/slideLayout3.xml", name: "节标题" });
    expect(choices[6].name).toBe("空白");
  });

  it("adds a slide on the chosen layout, and a save names that layout in its relationship", () => {
    const pres = Presentation.blank();
    const added = pres.addSlide("ppt/slideLayouts/slideLayout3.xml");
    expect(added.layoutPath).toBe("ppt/slideLayouts/slideLayout3.xml");
    expect((added.layoutPlaceholders ?? []).map((p) => p.type)).toEqual(["title", "body"]);
    expect(added.layoutPlaceholders?.[0]).toMatchObject({ sizePt: 44, anchor: "b" });
    const pkg = pkgOf(pres);
    expect(pkg.text("ppt/slides/_rels/slide2.xml.rels")).toContain('Target="../slideLayouts/slideLayout3.xml"');
    const back = open(writePptx(pres));
    expect(back.slides[1].layoutPath).toBe("ppt/slideLayouts/slideLayout3.xml");
    expect(back.slides[1].layoutPlaceholders?.map((p) => p.type)).toEqual(["title", "body"]);
  });

  it("keeps the default layout when no choice is given, and names layouts of a read deck", () => {
    const pres = Presentation.blank();
    expect(pres.addSlide().layoutPath).toBe("ppt/slideLayouts/slideLayout1.xml");
    const back = open(writePptx(pres));
    expect(back.layoutChoices().map((c) => c.name)).toEqual(pres.layoutChoices().map((c) => c.name));
    expect(back.slides[1].layoutPath).toBe("ppt/slideLayouts/slideLayout1.xml");
  });
});

describe("body text inherits the master's nine levels", () => {
  const claim = (pres: Presentation, layout: string, level: number): string => {
    const slide = pres.addSlide(layout);
    slide.shapes.push({
      kind: "shape",
      name: "文本占位符 2",
      x: 0,
      y: 0,
      width: 400,
      height: 200,
      placeholder: "body",
      placeholderIdx: 1,
      paragraphs: [{ level, runs: [{ text: "甲乙" }] }],
    });
    return renderSlideHtml(pres, pres.slides.indexOf(slide));
  };

  it("reads all nine bodyStyle levels back with sizes, indents and bullets", () => {
    const back = open(writePptx(Presentation.blank()));
    expect(back.bodyLevels.map((l) => l.sizePt)).toEqual([18, 16, 16, 14, 14, 18, 18, 18, 18]);
    expect(back.bodyLevels[0]).toMatchObject({ marLEmu: 228600, indentEmu: -228600, bullet: "●" });
    expect(back.bodyLevels[1].marLEmu).toBe(685800);
    expect(back.bodyLevels[8].marLEmu).toBe(228600 + 457200 * 8);
    expect(back.bodyStylePt).toBe(18);
    expect(pkgOf(Presentation.blank()).text("ppt/slideMasters/slideMaster1.xml")).toContain('<a:lvl9pPr marL="3886200" indent="-228600"');
  });

  it("draws a body placeholder's silent paragraph with the master's bullet and hanging indent", () => {
    const html = claim(Presentation.blank(), "ppt/slideLayouts/slideLayout2.xml", 0);
    expect(html).toContain('<span class="bullet">●</span>');
    expect(html).toContain("padding-left:24px;text-indent:-24px");
  });

  it("honours a layout's buNone: 节标题 body text carries no bullet and starts at the box edge", () => {
    const html = claim(Presentation.blank(), "ppt/slideLayouts/slideLayout3.xml", 0);
    expect(html).not.toContain('class="bullet"');
    expect(html).toContain("padding-left:0px;text-indent:0px");
  });

  it("draws a level-2 paragraph at the master's 16pt, 72px in with the level-2 bullet", () => {
    const html = claim(Presentation.blank(), "ppt/slideLayouts/slideLayout2.xml", 2);
    expect(html).toContain("padding-left:72px;text-indent:-24px");
    expect(html).toContain("font-size:16pt");
    expect(html).toContain('<span class="bullet">●</span>');
  });

  it("writes the layout lstStyle buNone 节标题/比较/图片与标题 declare", () => {
    const pkg = pkgOf(Presentation.blank());
    expect(pkg.text("ppt/slideLayouts/slideLayout3.xml")).toContain('<a:lvl1pPr marL="0" indent="0"><a:buNone/><a:defRPr sz="1800"/>');
    expect(pkg.text("ppt/slideLayouts/slideLayout5.xml")).toContain("<a:buNone/>");
    expect(pkg.text("ppt/slideLayouts/slideLayout8.xml")).toContain("<a:buNone/>");
  });
});
