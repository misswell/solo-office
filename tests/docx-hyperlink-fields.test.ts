import { describe, expect, it } from "vitest";
import { existsSync, readdirSync, readFileSync } from "node:fs";
import { join } from "node:path";
import { unzipSync } from "fflate";
import { WordDocument } from "../src/core/docx.js";
import { renderWord } from "../src/core/render.js";
import { writeDocx } from "../src/core/write-docx.js";
import { text } from "../src/core/opc.js";

/**
 * #86 的受众照子。授权 46 份（`.scratch/wps-corpus` + WPS 包内 `office6` 模板，只读；不读 owner 的个人目录）里：
 *
 * 1. `w:hyperlink` 元素 **0 枚**——链接只以 HYPERLINK 域出现（8 份 = 4 个模板身份 × zh_CN/zh_HK 两份本地化）。
 *    WPS 写域时，域结果的 run 只有字体与字号，没有 `w:rStyle`、没有 `w:color`、没有 `w:u`：纸上是纯正文，
 *    不该被读物自己加一条只有链接才有的样子。
 * 2. `styles.xml` 里有内置 Hyperlink / FollowedHyperlink 字符样式（`u=single`、色 0000FF / 800080）的模板
 *    有 7 个身份，但**没有一枚 run 引用它们**（`w:rStyle` 0 处）——「Hyperlink 样式那一路」在授权语料里没有受众，
 *    产品不为它发明行为；将来遇到引用的件，样式阶梯本来就会把色与线借上来。
 * 3. 域结果那串字必须画到纸上（8 份共 110 枚，每份本地化 55 枚），原样写回后 `w:instrText` 与 `w:fldChar`
 *    枚数一枚不少。
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

const decode = (s: string): string =>
  s
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#(\d+);/g, (_, d: string) => String.fromCharCode(Number(d)))
    .replace(/&amp;/g, "&");
const norm = (s: string): string => s.replace(/[\s&nbsp;]+/g, " ").trim();
const htmlText = (html: string): string => norm(decode(html.replace(/<[^>]*>/g, " ")));

type Field = { instr: string; range: [number, number] };

/** 复合域的结果半边：separate 之后、配对的 end 之前；只留 HYPERLINK 指令的域。 */
function hyperlinkFields(xml: string): Field[] {
  const marks = [...xml.matchAll(/<w:fldChar\b[^>]*\bw:fldCharType="(begin|separate|end)"[^>]*\/?>/g)].map((m) => ({
    type: m[1],
    tagStart: m.index ?? 0,
    tagEnd: (m.index ?? 0) + m[0].length,
  }));
  const out: Field[] = [];
  const stack: { instr: string; sep?: number; beginEnd: number }[] = [];
  for (const mk of marks) {
    if (mk.type === "begin") stack.push({ instr: "", sep: undefined, beginEnd: mk.tagEnd });
    else if (mk.type === "separate") {
      const top = stack[stack.length - 1];
      if (!top) continue;
      top.instr = xml.slice(top.beginEnd, mk.tagStart);
      top.sep = mk.tagEnd;
    } else {
      const top = stack.pop();
      if (top?.sep === undefined || !/HYPERLINK/i.test(top.instr)) continue;
      out.push({ instr: top.instr, range: [top.sep, mk.tagStart] });
    }
  }
  return out;
}

const resultTexts = (xml: string): string[] =>
  hyperlinkFields(xml).flatMap((f) =>
    [...xml.slice(f.range[0], f.range[1]).matchAll(/<w:t(?:\s[^>]*)?>([\s\S]*?)<\/w:t>/g)]
      .map((m) => norm(decode(m[1])))
      .filter((s) => s.length >= 2),
  );

const FILES = authorizedDocx();
const loaded = FILES.flatMap((p) => {
  try {
    return [{ path: p, parts: unzipSync(new Uint8Array(readFileSync(p))) }];
  } catch {
    return [];
  }
});
const fieldFiles = loaded.flatMap(({ path, parts }) => {
  const texts: string[] = [];
  let fields = 0;
  for (const [k, v] of Object.entries(parts)) {
    if (!/^word\/.*\.xml$/.test(k)) continue;
    const xml = text(v);
    const found = hyperlinkFields(xml);
    fields += found.length;
    texts.push(...resultTexts(xml));
  }
  return fields ? [{ path, parts, texts, fields }] : [];
});
const count = (pattern: RegExp): number =>
  loaded.reduce((n, { parts }) => {
    let m = 0;
    for (const [k, v] of Object.entries(parts)) if (/^word\/.*\.xml$/.test(k)) m += (text(v).match(pattern) ?? []).length;
    return n + m;
  }, 0);

describe.skipIf(loaded.length === 0)("docx 链接的两种形状（授权语料照子）", () => {
  it("链接只以 HYPERLINK 域出现：w:hyperlink 元素 0 枚", () => {
    expect(loaded.length).toBe(46);
    expect(count(/<w:hyperlink\b/g)).toBe(0);
    expect(fieldFiles.length).toBeGreaterThanOrEqual(8);
  });

  it("Hyperlink 字符样式有定义、没有引用：授权语料里 0 处 w:rStyle 指向它", () => {
    let stylesNamed = 0;
    let refs = 0;
    for (const { parts } of loaded) {
      const styles = parts["word/styles.xml"] ? text(parts["word/styles.xml"]) : "";
      const ids = [...styles.matchAll(/<w:style\b[\s\S]*?<\/w:style>/g)]
        .filter((m) => /<w:name\b[^>]*\bw:val="(?:Followed)?Hyperlink"/.test(m[0]) || /超链接/.test(m[0]))
        .map((m) => /\bw:styleId="([^"]*)"/.exec(m[0])?.[1])
        .filter((v): v is string => !!v);
      stylesNamed += ids.length;
      if (!ids.length) continue;
      for (const [k, v] of Object.entries(parts)) {
        if (!/^word\/.*\.xml$/.test(k) || k === "word/styles.xml") continue;
        const xml = text(v);
        for (const id of ids) refs += (xml.match(new RegExp(`<w:rStyle\\b[^>]*\\bw:val="${id}"`, "g")) ?? []).length;
      }
    }
    expect(stylesNamed).toBeGreaterThanOrEqual(14);
    expect(refs).toBe(0);
  });

  it("HYPERLINK 域的结果文字画得出，写回后域指令与标记一枚不少", () => {
    expect(fieldFiles.length).toBeGreaterThanOrEqual(8);
    let drawn = 0;
    const missing: string[] = [];
    const lost: string[] = [];
    for (const { path, parts, texts } of fieldFiles) {
      const label = path.split("/").slice(-2).join("/");
      const bytes = new Uint8Array(readFileSync(path));
      const page = htmlText(renderWord(WordDocument.openBytes(bytes)));
      for (const t of texts) {
        if (page.includes(t)) drawn++;
        else missing.push(`${label}: ${t.slice(0, 20)}`);
      }
      const after = unzipSync(writeDocx(WordDocument.openBytes(bytes)));
      const tally = (p: Record<string, Uint8Array>, re: RegExp): number => {
        let n = 0;
        for (const [k, v] of Object.entries(p)) if (/^word\/.*\.xml$/.test(k)) n += (text(v).match(re) ?? []).length;
        return n;
      };
      for (const re of [/<w:instrText\b/g, /<w:fldChar\b/g]) {
        const a = tally(parts, re);
        const b = tally(after, re);
        if (a !== b) lost.push(`${label}: ${re.source} ${a}→${b}`);
      }
    }
    expect(missing.slice(0, 5)).toEqual([]);
    expect(lost.slice(0, 5)).toEqual([]);
    // 每份本地化各 55 枚（zh_CN 4 个身份 + zh_HK 4 个身份）。
    expect(drawn).toBe(110);
  });
});
