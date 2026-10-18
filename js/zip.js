// @ts-check
/*
 * Word（.docx）・Excel（.xlsx）の入れ物である ZIP を開く
 *   listEntries(バイト列) → 中のファイルの一覧（壊れていれば ZipError）
 *   readEntry(バイト列, 一覧の1件, 展開関数) → 中のファイルのバイト列
 * 展開（deflate）は、ブラウザ標準の DecompressionStream を使う。テストでは同じものを Node で使う。
 */
'use strict';

(function (root) {
  /**
   * @typedef {Object} ZipEntry
   * @property {string} name
   * @property {number} method          0: 圧縮なし、8: deflate
   * @property {boolean} encrypted      ZIP のパスワードがかかっている
   * @property {number} compressedSize
   * @property {number} size
   * @property {number} localOffset
   */

  /** @typedef {(data: Uint8Array) => Promise<Uint8Array>} InflateRaw */

  /** 展開後の大きさの上限（壊れたファイルや、極端に膨らむファイルで画面が固まらないように） */
  const MAX_ENTRY_SIZE = 200 * 1024 * 1024;

  class ZipError extends Error {
    /**
     * @param {'broken'|'encrypted'|'too-large'} code
     * @param {string} message
     */
    constructor(code, message) {
      super(message);
      this.name = 'ZipError';
      this.code = code;
    }
  }

  /**
   * @param {Uint8Array} b
   * @param {number} at
   */
  const u16 = (b, at) => b[at] | (b[at + 1] << 8);
  /**
   * @param {Uint8Array} b
   * @param {number} at
   */
  const u32 = (b, at) => (b[at] | (b[at + 1] << 8) | (b[at + 2] << 16)) + b[at + 3] * 0x1000000;
  /**
   * @param {Uint8Array} b
   * @param {number} at
   */
  const u64 = (b, at) => u32(b, at) + u32(b, at + 4) * 0x100000000;

  const SIG_LOCAL = 0x04034b50;
  const SIG_CENTRAL = 0x02014b50;
  const SIG_END = 0x06054b50;
  const SIG_END64 = 0x06064b50;
  const SIG_END64_LOCATOR = 0x07064b50;

  /** @param {Uint8Array} bytes */
  const isZip = (bytes) =>
    bytes.length >= 4 && u32(bytes, 0) === SIG_LOCAL;

  /**
   * 末尾にある「目次の終わり」の位置
   * @param {Uint8Array} b
   */
  const findEnd = (b) => {
    const min = Math.max(0, b.length - 22 - 0xffff);
    for (let i = b.length - 22; i >= min; i -= 1) {
      if (u32(b, i) === SIG_END) return i;
    }
    return -1;
  };

  const utf8 = new TextDecoder('utf-8');

  /**
   * ZIP の中のファイルの一覧を読む
   * @param {Uint8Array} b
   * @returns {ZipEntry[]}
   */
  const listEntries = (b) => {
    const end = findEnd(b);
    if (end < 0) throw new ZipError('broken', 'ZIP の目次が見つかりません');
    let count = u16(b, end + 10);
    let cdSize = u32(b, end + 12);
    let cdOffset = u32(b, end + 16);
    if (count === 0xffff || cdOffset === 0xffffffff || cdSize === 0xffffffff) {
      const loc = end - 20;
      if (loc < 0 || u32(b, loc) !== SIG_END64_LOCATOR) throw new ZipError('broken', 'ZIP64 の目次が見つかりません');
      const end64 = u64(b, loc + 8);
      if (end64 + 56 > b.length || u32(b, end64) !== SIG_END64) throw new ZipError('broken', 'ZIP64 の目次が壊れています');
      count = u64(b, end64 + 32);
      cdSize = u64(b, end64 + 40);
      cdOffset = u64(b, end64 + 48);
    }
    if (cdOffset + cdSize > b.length) throw new ZipError('broken', 'ZIP の目次が途中で切れています');

    /** @type {ZipEntry[]} */
    const entries = [];
    let p = cdOffset;
    for (let n = 0; n < count; n += 1) {
      if (p + 46 > b.length || u32(b, p) !== SIG_CENTRAL) throw new ZipError('broken', 'ZIP の目次が壊れています');
      const flags = u16(b, p + 8);
      const method = u16(b, p + 10);
      let compressedSize = u32(b, p + 20);
      let size = u32(b, p + 24);
      const nameLen = u16(b, p + 28);
      const extraLen = u16(b, p + 30);
      const commentLen = u16(b, p + 32);
      let localOffset = u32(b, p + 42);
      const name = utf8.decode(b.subarray(p + 46, p + 46 + nameLen));
      // ZIP64 の追加情報（大きさ・位置が 32bit に収まらないとき）
      let e = p + 46 + nameLen;
      const extraEnd = e + extraLen;
      while (e + 4 <= extraEnd) {
        const id = u16(b, e);
        const len = u16(b, e + 2);
        if (id === 0x0001) {
          let q = e + 4;
          if (size === 0xffffffff) { size = u64(b, q); q += 8; }
          if (compressedSize === 0xffffffff) { compressedSize = u64(b, q); q += 8; }
          if (localOffset === 0xffffffff) { localOffset = u64(b, q); }
        }
        e += 4 + len;
      }
      entries.push({
        name: name.replace(/\\/g, '/'),
        method,
        encrypted: (flags & 1) === 1 || method === 99,
        compressedSize,
        size,
        localOffset,
      });
      p = extraEnd + commentLen;
    }
    return entries;
  };

  /**
   * ブラウザ標準の DecompressionStream で deflate を展開する
   * @type {InflateRaw}
   */
  const inflateRawWithStream = async (data) => {
    const ds = new DecompressionStream('deflate-raw');
    const writer = ds.writable.getWriter();
    // 書き込み側の失敗は、読み出し側（reader.read）でも同じ理由で失敗するため、ここではそちらに任せる
    writer.write(data).catch(() => undefined);
    writer.close().catch(() => undefined);
    const reader = ds.readable.getReader();
    /** @type {Uint8Array[]} */
    const chunks = [];
    let total = 0;
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      total += value.length;
      if (total > MAX_ENTRY_SIZE) {
        await reader.cancel();
        throw new ZipError('too-large', '展開すると大きすぎるファイルが含まれています');
      }
      chunks.push(value);
    }
    const out = new Uint8Array(total);
    let at = 0;
    for (const c of chunks) {
      out.set(c, at);
      at += c.length;
    }
    return out;
  };

  /**
   * ZIP の中の1つのファイルを取り出す
   * @param {Uint8Array} b
   * @param {ZipEntry} entry
   * @param {InflateRaw} [inflateRaw]
   * @returns {Promise<Uint8Array>}
   */
  const readEntry = async (b, entry, inflateRaw = inflateRawWithStream) => {
    if (entry.encrypted) throw new ZipError('encrypted', `${entry.name} にパスワードがかかっています`);
    if (entry.size > MAX_ENTRY_SIZE) throw new ZipError('too-large', `${entry.name} が大きすぎます`);
    const p = entry.localOffset;
    if (p + 30 > b.length || u32(b, p) !== SIG_LOCAL) throw new ZipError('broken', `${entry.name} の位置が正しくありません`);
    const start = p + 30 + u16(b, p + 26) + u16(b, p + 28);
    const end = start + entry.compressedSize;
    if (end > b.length) throw new ZipError('broken', `${entry.name} が途中で切れています`);
    const raw = b.subarray(start, end);
    if (entry.method === 0) return raw;
    if (entry.method !== 8) throw new ZipError('broken', `${entry.name} は対応していない圧縮方式（${entry.method}）です`);
    try {
      return await inflateRaw(raw);
    } catch (e) {
      if (e instanceof ZipError) throw e;
      const detail = e instanceof Error ? e.message : String(e);
      throw new ZipError('broken', `${entry.name} を展開できません（${detail}）`);
    }
  };

  /**
   * バイト列を文字にする（UTF-8。先頭に UTF-16 の印があれば UTF-16）
   * @param {Uint8Array} bytes
   */
  const decodeText = (bytes) => {
    if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes.subarray(2));
    if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes.subarray(2));
    return utf8.decode(bytes);
  };

  /**
   * 名前で引ける形にする（大文字・小文字は区別しない。ZIP の中の名前は Office でも揺れることがあるため）
   * @param {ZipEntry[]} entries
   * @returns {ReadonlyMap<string, ZipEntry>}
   */
  const indexEntries = (entries) => new Map(entries.map((e) => [e.name.replace(/^\/+/, '').toLowerCase(), e]));

  root.YegZip = Object.freeze({
    ZipError,
    MAX_ENTRY_SIZE,
    isZip,
    listEntries,
    readEntry,
    decodeText,
    indexEntries,
    inflateRawWithStream,
  });
})(globalThis);
