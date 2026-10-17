import { strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vitest";
import { WordDocument } from "../src/core/docx.js";
import { writeDocx } from "../src/core/write-docx.js";
import { makeCommentedDocx } from "./helpers/make-comments.js";
import { unstamped } from "./helpers/unstamp.js";

/** 两枚戳之外都不该动，所以夹具一律 level 0（正文原样躺在包里，好核对）。 */
const zipOf = (text: string) => zipSync({ "a/x.xml": strToU8(text) }, { level: 0 });

/** 本地头与结束记录那两枚签名的字面，用码位拼出来，源码里不留控制字节。 */
const sig = (c: number, d: number) => String.fromCharCode(0x50, 0x4b, c, d);

describe("zip 的 DOS 时间戳抹平（逐字节断言前的一步）", () => {
  it("同一份内容隔多久拼的都算相等", () => {
    const a = zipOf("<a>话</a>");
    const b = zipOf("<a>话</a>");
    expect(unstamped(a)).toEqual(unstamped(b));
  });

  it("抹的正是那两枚戳：戳换掉之后仍相等，原始字节却不等", () => {
    const a = zipOf("<a>话</a>");
    const poked = Uint8Array.from(a);
    // 第一个本地头从 0 开始：时间占 10..12，日期占 12..14。
    for (let i = 10; i < 14; i++) poked[i] = 0xff;
    expect(Uint8Array.from(a)).not.toEqual(poked);
    expect(unstamped(a)).toEqual(unstamped(poked));
  });

  it("正文里长得像签名的那几个字节不受牵连", () => {
    const text = `<a>${sig(0x03, 0x04)}话${sig(0x05, 0x06)}</a>`;
    const a = zipOf(text);
    const back = unzipSync(unstamped(a))["a/x.xml"];
    expect(new TextDecoder().decode(back)).toBe(text);
  });

  it("不是 zip 就直说，不静默原样返回", () => {
    expect(() => unstamped(strToU8("这不是一个包"))).toThrow(/中央目录结束记录/);
  });

  // 这一条是那条随机红的正面照：夹具与写出器各自盖戳，跨过 2 秒粒度后原始字节确实不等了。
  it("隔着戳边界重拼同一份正文，抹平之后仍然相等", async () => {
    const bytes = makeCommentedDocx();
    const doc = WordDocument.openBytes(bytes);
    await new Promise((r) => setTimeout(r, 2100));
    const out = writeDocx(doc);
    expect(Uint8Array.from(out)).not.toEqual(Uint8Array.from(bytes));
    expect(unstamped(out)).toEqual(unstamped(bytes));
  });
});
