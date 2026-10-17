/**
 * 「逐字节相等」这条断言在 zip 里有一格不该由它管：每一项都盖着一枚 DOS 时间（本地文件头 4 字节 + 中央目录 4 字节）。
 * 写出器不认领那一格——它按保存那一刻重新盖戳，那是容器的作息，不是正文的话。
 * 夹具却是在测试里现拼的，也盖同一枚 `new Date()`：两次调用撞在同一个 2 秒粒度里才相等，
 * 跨过去就随机红一条（995 条里天天有红的份）。所以要比正文，先把那几枚戳抹平。
 *
 * 抹的是按格式读出来的位置，不是全文搜签名字节：正文里也会出现 `PK\x03\x04` 这四个字节，跟着抹会把话抹掉。
 * 走中央目录（条目数与起点都写在 EOCD 里），逐条抹自己那 4 个字节，再按记录里的本地头偏移抹那一份。
 */
export function unstamped(bytes: Uint8Array): Uint8Array {
  const out = Uint8Array.from(bytes);
  const dv = new DataView(out.buffer, out.byteOffset, out.byteLength);
  const eocd = findEocd(dv);
  if (eocd < 0) throw new Error("不是 zip：末尾没有中央目录结束记录");
  const count = dv.getUint16(eocd + 10, true);
  let ptr = dv.getUint32(eocd + 16, true);
  for (let i = 0; i < count; i++) {
    if (dv.getUint32(ptr, true) !== 0x02014b50) throw new Error(`中央目录第 ${i} 条签名不对（${ptr}）`);
    const nameLen = dv.getUint16(ptr + 28, true);
    const extraLen = dv.getUint16(ptr + 30, true);
    const commentLen = dv.getUint16(ptr + 32, true);
    const local = dv.getUint32(ptr + 42, true);
    dv.setUint16(ptr + 12, 0, true); // 中央目录那一份：时间 + 日期
    dv.setUint16(ptr + 14, 0, true);
    if (dv.getUint32(local, true) !== 0x04034b50) throw new Error(`第 ${i} 条指到的本地头签名不对（${local}）`);
    dv.setUint16(local + 10, 0, true); // 本地头那一份
    dv.setUint16(local + 12, 0, true);
    ptr += 46 + nameLen + extraLen + commentLen;
  }
  return out;
}

/** 从尾巴往前找 EOCD 签名（0x06054b50）：注释最长 65535，所以只需在末尾 64KiB+22 里找。 */
function findEocd(dv: DataView): number {
  const floor = Math.max(0, dv.byteLength - 22 - 0xffff);
  for (let i = dv.byteLength - 22; i >= floor; i--) if (dv.getUint32(i, true) === 0x06054b50) return i;
  return -1;
}
