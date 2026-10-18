import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import type { Inline, Paragraph, WordDocument } from "../src/core/docx.js";
import { WordDocument as Doc } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { text } from "../src/core/opc.js";
import { unzipSync } from "fflate";

/**
 * 授权语料照子（口径见 `AGENTS.md`：`.scratch/wps-corpus` 的 docx 家族 + WPS 包内 `office6` 模板件，只读；
 * **不读 ~/Documents 或 ~/Desktop**）。三条问法都来自 2026-10-16 那一轮普查（`.scratch/unmodeled-census.mjs`
 * 数到正文系部件里 src 没提的只剩 4 个名字）：
 *
 * 1. 件里正文的每一串字，纸上都得有——画不出来的就是「WPS 有、我们没有」的那一块。
 * 2. 页眉页脚的字同理，但**域里那一串除外**：复合域从指令到 `end` 之间是文件算好留在那儿的结果（`PAGE` 那枚
 *    写着 36），我们的铺排自己数页，把结果当字画上会写回一个假页码。域本身另问：纸上要有它那一枚记号。
 * 3. 打开再原样写回：部件一枚不少，`w:sdt` 与它那一格 `w:sdtEndPr`（授权语料 52 枚 / 6 份）一枚不少。
 * 4. 多节件的带子（#163）：**逐节的 `w:headerReference`/`w:footerReference` 表**与**每节纸面显示的那一份**
 *    在写回后逐节一致——既不凭空多 cite，也不把某一节的带子挪到别的节上。授权口径 46 份 58 节里有 8 个多节实例
 *    （4 个模板身份 × 各自的本地化副本），这一条原先只由 `.scratch/band-per-section.ts` 一次性量过。
 * 5. 字符缩放一族（`w:rPr/w:w` 与 `w:rPr/w:fitText`）：真件里那 43 枚在「打开—打字—写回—再开」这一圈里一枚不少。
 *    这一条只问带不带得回，不问画不画——画的那一腿没有 CSS 通道，读数见那条 `it` 的注释。
 *
 * 两处目录不在（换机器、只装了源码）就整组跳过，不拿「没有」当「通过」。
 */
const CORPUS = ".scratch/wps-corpus";
const BUNDLE = "/Applications/wpsoffice.app/Contents/Resources/office6";
const EXT = /\.(docx|dotx|dotm|docm)$/i;

function authorizedDocx(): string[] {
  if (!existsSync(CORPUS)) return [];
  const out = readdirSync(CORPUS).sort().filter((f) => EXT.test(f)).map((f) => join(CORPUS, f));
  const walk = (dir: string): void => {
    for (const e of readdirSync(dir, { withFileTypes: true })) {
      const p = join(dir, e.name);
      if (e.isDirectory()) walk(p);
      else if (EXT.test(e.name)) out.push(p);
    }
  };
  if (existsSync(BUNDLE)) walk(BUNDLE);
  return out;
}

const FILES = authorizedDocx();
const decode = (s: string): string =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCharCode(Number(d)))
    .replace(/&amp;/g, "&");
const norm = (s: string): string => s.replace(/[\s&nbsp;]+/g, " ");
/** 去掉域（指令 + 域结果）与简单域整段，只留文件自己写下的字。 */
const outOfFields = (xml: string): string =>
  xml
    .replace(/<w:instrText[\s\S]*?<w:fldChar w:fldCharType="end"\/?>/g, "")
    .replace(/<w:fldSimple[\s\S]*?<\/w:fldSimple>/g, "");
const stringsOf = (xml: string): string[] =>
  [...outOfFields(xml).matchAll(/<w:t(?:\s[^>]*)?>([^<]*)<\/w:t>/g)]
    .map((m) => norm(decode(m[1] ?? "")).trim())
    .filter((v) => v.length >= 2);
const tally = (parts: Record<string, Uint8Array>, tag: string): number => {
  let n = 0;
  for (const k of Object.keys(parts)) if (/^word\/.*\.xml$/.test(k)) n += (text(parts[k]).match(new RegExp(`<w:${tag}[\\s/>]`, "g")) ?? []).length;
  return n;
};

/**
 * 文件自己写的每一节 `w:sectPr`，按节起始顺序切出来（opener-walk，与 `.scratch/authorized-families3.mjs` 数
 * 58 节同一口径）。带子归谁要问的是字节，不是模型解析后的树——两套表示各数一遍，正是「带子挪到别的节上」这类
 * 问题的来路。
 */
function sectPrsOf(bodyXml: string): string[] {
  const out: string[] = [];
  let i = 0;
  for (;;) {
    const at = bodyXml.indexOf("<w:sectPr", i);
    if (at < 0) break;
    const gt = bodyXml.indexOf(">", at);
    const self = bodyXml[gt - 1] === "/";
    const close = bodyXml.indexOf("</w:sectPr>", gt);
    const end = self ? gt + 1 : close < 0 ? bodyXml.length : close + "</w:sectPr>".length;
    out.push(bodyXml.slice(at, end));
    i = end;
  }
  return out;
}

/**
 * 一节 cite 的带子，按它写出的顺序拼成 `kind:type→id`。属性顺序不参与比较：授权 46 份的 38 枚全是 `r:id` 在前，
 * Word 自己的写法相反，按固定顺序的写法去读其中一种就是一枚数不到的 0。
 */
function citesOf(sectPr: string): string[] {
  return [...sectPr.matchAll(/<w:(header|footer)Reference\b[^>]*\/>/g)].map((m) => {
    const type = /\b\w*:?type="([^"]+)"/.exec(m[0])?.[1] ?? "(无type)";
    const rid = /\br:id="([^"]+)"/.exec(m[0])?.[1] ?? "(无id)";
    return `${m[1]}:${type}→${rid}`;
  });
}

describe.skipIf(FILES.length === 0)("授权语料照子", () => {
  it(`正文与页眉页脚的字都画得出、写回不丢（${FILES.length} 份）`, () => {
    expect(FILES.length).toBeGreaterThan(0);
    const missingBody: string[] = [];
    const missingBand: string[] = [];
    const lostOnSave: string[] = [];
    let fields = 0;
    let fieldMarks = 0;

    for (const p of FILES) {
      const bytes = new Uint8Array(readFileSync(p));
      let before: Record<string, Uint8Array>;
      try {
        before = unzipSync(bytes);
      } catch {
        continue;
      }
      const label = p.split("/").slice(-2).join("/");
      const bodyXml = outOfFields(/<w:body>([\s\S]*)<\/w:body>/.exec(text(before["word/document.xml"] ?? new Uint8Array()))?.[1] ?? "");
      const doc: WordDocument = Doc.openBytes(bytes);
      const html = renderWord(doc);
      const drawn = norm(decode(html.replace(/<[^>]*>/g, " ")));

      for (const v of stringsOf(bodyXml)) if (!drawn.includes(v)) missingBody.push(`${label}: ${v.slice(0, 40)}`);

      for (const k of Object.keys(before).sort()) {
        if (!/^word\/(header|footer)\d*\.xml$/.test(k)) continue;
        for (const v of stringsOf(text(before[k]))) {
          if (!drawn.includes(v)) missingBand.push(`${label} ${k}: ${v.slice(0, 40)}`);
        }
      }
      fields += tally(before, "instrText") + tally(before, "fldSimple");
      fieldMarks += (html.match(/data-tag="(?:fldChar|fldSimple)"/g) ?? []).length;

      const after = unzipSync(writeDocx(doc));
      if (Object.keys(after).length < Object.keys(before).length) lostOnSave.push(`${label}: 部件 ${Object.keys(before).length}→${Object.keys(after).length}`);
      for (const tag of ["sdt", "sdtEndPr", "sdtPr", "p", "tbl"]) {
        const a = tally(before, tag);
        const b = tally(after, tag);
        if (a !== b) lostOnSave.push(`${label}: ${tag} ${a}→${b}`);
      }
    }

    console.log(`份数=${FILES.length} 域=${fields} 枚 / 纸上域记号=${fieldMarks} 枚（记号数不进断言，见文件头第 2 条）`);
    expect(missingBody.slice(0, 5)).toEqual([]);
    expect(missingBand.slice(0, 5)).toEqual([]);
    expect(lostOnSave.slice(0, 5)).toEqual([]);
  });

  /**
   * #163：多节件的带子在写回后逐节对得上。三件事各问一次——节数没变、每一节 cite 的那张表没变（既不凭空多
   * 一枚，也不把某一节的引用挪到别节）、每一节纸面显示的那一份部件没变。前两件问的是文件的字节，第三件问的是
   * 生产代码里的 `sectionBand`。
   */
  it("多节件的逐节 cite 表与纸面显示在写回后一致", () => {
    const drift: string[] = [];
    let instances = 0;
    let sections = 0;
    for (const p of FILES) {
      const bytes = new Uint8Array(readFileSync(p));
      let before: Record<string, Uint8Array>;
      try {
        before = unzipSync(bytes);
      } catch {
        continue;
      }
      const bodyOf = (parts: Record<string, Uint8Array>): string => /<w:body>([\s\S]*)<\/w:body>/.exec(text(parts["word/document.xml"] ?? new Uint8Array()))?.[1] ?? "";
      const sects = sectPrsOf(bodyOf(before));
      if (sects.length < 2) continue;
      instances++;
      sections += sects.length;
      const label = p.split("/").slice(-2).join("/");
      const doc = Doc.openBytes(bytes);
      const cites = sects.map(citesOf);
      const shown = sects.map((_, k) => [doc.sectionBand("header", k)?.part ?? "-", doc.sectionBand("footer", k)?.part ?? "-"]);
      if (doc.sectionRefs.length !== sects.length) drift.push(`${label}: 模型数到 ${doc.sectionRefs.length} 节，文件写了 ${sects.length} 节`);
      const out = writeDocx(doc);
      const after = unzipSync(out);
      const sectsAfter = sectPrsOf(bodyOf(after));
      if (sectsAfter.length !== sects.length) {
        drift.push(`${label}: 节数 ${sects.length}→${sectsAfter.length}`);
        continue;
      }
      sectsAfter.forEach((s, k) => {
        const a = cites[k]!.join(" ");
        const b = citesOf(s).join(" ");
        if (a !== b) drift.push(`${label} 第${k + 1}节 cite 表 [${a}]→[${b}]`);
      });
      const again = Doc.openBytes(out);
      const shownAfter = sectsAfter.map((_, k) => [again.sectionBand("header", k)?.part ?? "-", again.sectionBand("footer", k)?.part ?? "-"]);
      shown.forEach((v, k) => {
        const w = shownAfter[k]!;
        if (v[0] !== w[0] || v[1] !== w[1]) drift.push(`${label} 第${k + 1}节 纸面显示 [${v}]→[${w}]`);
      });
    }
    console.log(`多节件=${instances} 个实例 / 节=${sections} 节（口径见文件头第 4 条）`);
    expect(instances, "授权口径里多节的件至少有量到的那 8 个实例").toBeGreaterThanOrEqual(8);
    expect(drift.slice(0, 8)).toEqual([]);
  });

  /**
   * `w:rPr/w:w`（字符缩放）与 `w:rPr/w:fitText`（缩到指定宽度）。授权 46 件里**各 43 枚 / 2 件**——`business plan.dotm`
   * 的 `zh_CN` 与 `zh_HK` 两份（件按全路径计，basename 去重会把 2 读成 1），**枚枚非缺省**：`w:w` 全 `93`、`w:fitText`
   * 全 `7791`，43 枚**成对住在同一枚 `w:rPr` 里**，全在 `w:p/w:r/w:rPr`（正文那 43 枚中文标签：项目名称 / 电话 / 地址…）；
   * `w:pPr/w:fitText` 0 枚（`.scratch/rpr-spacing-audience.py` 同一条 `xml.etree` 路线，2026-10-18 量）。
   *
   * 这一条只问「带不带得回」，不问「画不画」——画的那一腿今天在自家 Blink 上量完是**没有通道**的
   * （`.scratch/scale-css-probe.html`：16 枚 12pt 中文，同量墨宽与推进量）。`transform:scaleX(.93)` 对
   * `display:inline` 的 span 一动不动（比 1.0000，CSS 本就不让 transform 作用于非替换行内盒）；补 `inline-block`
   * 后墨宽 0.9300 而推进 0.9650（断行位置跟着错）；`font-stretch:93%` 的 computed 留住了 `93%`，但墨宽与推进都是
   * 1.0000（这张中文脸不是可变字体，Blink 不合成窄体），`condensed` / `semi-condensed` 同样不动；
   * `font-variation-settings:'wdth' 93` 也是 1.0000。唯一两样等比的是 `letter-spacing:-0.07em`（0.9301 / 0.9301），
   * 可那条通道 #174 已经交给 `w:spacing`，况且它压的是**字距**不是字宽。WPS 把 93% 画成哪一种：**未取证**（无录屏授权）。
   */
  it("真件里那 43 枚字符缩放与 fitText 打字写回一枚不少", () => {
    const paragraphsOf = (doc: WordDocument): Paragraph[] => doc.blocks.filter((b): b is Paragraph => b.kind === "paragraph");
    const scaledRun = (q: Paragraph): { run: Inline; at: number } | undefined => {
      const at = q.inlines.findIndex((r) => r.rPrSlots?.some((s) => s.tag === "w" || s.tag === "fitText"));
      return at < 0 ? undefined : { run: q.inlines[at]!, at };
    };
    const drift: string[] = [];
    let files = 0;
    let scale = 0;
    let fit = 0;
    let typed = 0;
    for (const p of FILES) {
      const bytes = new Uint8Array(readFileSync(p));
      let before: Record<string, Uint8Array>;
      try {
        before = unzipSync(bytes);
      } catch {
        continue;
      }
      const n = tally(before, "w");
      const m = tally(before, "fitText");
      if (n === 0 && m === 0) continue;
      files++;
      scale += n;
      fit += m;
      const label = p.split("/").slice(-2).join("/");
      const doc = Doc.openBytes(bytes);
      // 打字那一腿问的是 #174 学到的那一课：页面重建的那份会盖住原件，缩放若只活在大字节里、没跟着 run 走，一打字就没。
      let hit: { para: Paragraph; run: Inline; at: number } | undefined;
      for (const q of paragraphsOf(doc)) {
        const found = scaledRun(q);
        if (found) {
          hit = { para: q, ...found };
          break;
        }
      }
      if (hit) {
        typed++;
        hit.para.inlines[hit.at] = { ...hit.run, text: `${hit.run.text}打的字` };
      }
      const out = writeDocx(doc);
      const after = unzipSync(out);
      if (tally(after, "w") !== n) drift.push(`${label}: 打字写回后 w:rPr/w:w ${n}→${tally(after, "w")}`);
      if (tally(after, "fitText") !== m) drift.push(`${label}: 打字写回后 w:rPr/w:fitText ${m}→${tally(after, "fitText")}`);
      // 再开一遍再写：第二代件还得说同一句话，否则「带回」只是第一代件的运气。
      const twice = unzipSync(writeDocx(Doc.openBytes(out)));
      if (tally(twice, "w") !== n) drift.push(`${label}: 再开再写后 w:rPr/w:w ${n}→${tally(twice, "w")}`);
      if (tally(twice, "fitText") !== m) drift.push(`${label}: 再开再写后 w:rPr/w:fitText ${m}→${tally(twice, "fitText")}`);
    }
    console.log(`字符缩放件=${files} / w:w=${scale} 枚 / w:fitText=${fit} 枚 / 打过字的件=${typed}`);
    expect(files, "授权口径里说缩放的件至少有量到的那 2 件").toBeGreaterThanOrEqual(2);
    expect(scale, "`w:rPr/w:w` 至少有量到的那 43 枚").toBeGreaterThanOrEqual(43);
    expect(fit, "`w:rPr/w:fitText` 至少有量到的那 43 枚").toBeGreaterThanOrEqual(43);
    expect(typed, "打字那一腿必须真有 run 可打").toBe(files);
    expect(drift.slice(0, 8)).toEqual([]);
  });
});
