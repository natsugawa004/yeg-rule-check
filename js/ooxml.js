// @ts-check
/*
 * Word・Excel のファイル（Office Open XML）に共通の読み方
 *   openPackage(バイト列) → 中のファイルを名前で読める入れ物
 *   readRels(入れ物, 部品の名前) → その部品からの参照（rId → 参照先）
 * 名前空間は、通常の形式と「厳密（Strict）」形式の両方を受ける
 */
'use strict';

(function (root) {
  /**
   * xml.js の要素（xml.js の XmlElement と同じ形）
   * @typedef {Object} XmlElement
   * @property {string} name
   * @property {string} local
   * @property {string} ns
   * @property {Readonly<Record<string, string>>} attrs
   * @property {ReadonlyMap<string, string>} scope
   * @property {(XmlElement|string)[]} children
   */
  /**
   * @typedef {Object} ZipEntry
   * @property {string} name
   * @property {boolean} encrypted
   */
  /** @typedef {(data: Uint8Array) => Promise<Uint8Array>} InflateRaw */
  /**
   * @typedef {Object} ZipApi
   * @property {(b: Uint8Array) => ZipEntry[]} listEntries
   * @property {(b: Uint8Array, e: ZipEntry, inflateRaw?: InflateRaw) => Promise<Uint8Array>} readEntry
   * @property {(b: Uint8Array) => string} decodeText
   * @property {(entries: ZipEntry[]) => ReadonlyMap<string, ZipEntry>} indexEntries
   */
  /**
   * @typedef {Object} XmlApi
   * @property {(s: string) => XmlElement} parseXml
   * @property {(el: XmlElement, nsSet: ReadonlySet<string>|null, local: string) => boolean} is
   * @property {(el: XmlElement, nsSet?: ReadonlySet<string>|null, local?: string) => XmlElement[]} kids
   * @property {(el: XmlElement, nsSet: ReadonlySet<string>|null, local: string) => XmlElement|null} kid
   * @property {(el: XmlElement, local: string) => string|null} attr
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (globalThis));
  const zip = /** @type {ZipApi} */ (g.YegZip);
  const xml = /** @type {XmlApi} */ (g.YegXml);

  /**
   * @param {...string} uris
   * @returns {ReadonlySet<string>}
   */
  const nsSet = (...uris) => new Set(uris);

  const NS = Object.freeze({
    pkgRel: nsSet('http://schemas.openxmlformats.org/package/2006/relationships'),
    rel: nsSet(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'http://purl.oclc.org/ooxml/officeDocument/relationships',
    ),
    sml: nsSet('http://schemas.openxmlformats.org/spreadsheetml/2006/main', 'http://purl.oclc.org/ooxml/spreadsheetml/main'),
    wml: nsSet('http://schemas.openxmlformats.org/wordprocessingml/2006/main', 'http://purl.oclc.org/ooxml/wordprocessingml/main'),
    xdr: nsSet(
      'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
      'http://purl.oclc.org/ooxml/drawingml/spreadsheetDrawing',
    ),
    a: nsSet('http://schemas.openxmlformats.org/drawingml/2006/main', 'http://purl.oclc.org/ooxml/drawingml/main'),
    math: nsSet('http://schemas.openxmlformats.org/officeDocument/2006/math', 'http://purl.oclc.org/ooxml/officeDocument/math'),
    mc: nsSet('http://schemas.openxmlformats.org/markup-compatibility/2006'),
  });

  /**
   * @typedef {Object} Package
   * @property {Uint8Array} bytes
   * @property {(name: string) => boolean} has
   * @property {(name: string) => Promise<XmlElement>} readXml
   * @property {string[]} names
   */

  /**
   * @typedef {Object} Relationship
   * @property {string} id
   * @property {string} type
   * @property {string} target  パッケージの中の名前（先頭の / なし）。外部への参照は空文字
   */

  /**
   * ZIP を開き、中のファイルを名前で読めるようにする
   * @param {Uint8Array} bytes
   * @param {InflateRaw} [inflateRaw]
   * @returns {Package}
   */
  const openPackage = (bytes, inflateRaw) => {
    const entries = zip.listEntries(bytes);
    const index = zip.indexEntries(entries);
    /** @param {string} name */
    const find = (name) => index.get(name.replace(/^\/+/, '').toLowerCase()) ?? null;
    return {
      bytes,
      names: entries.map((e) => e.name),
      has: (name) => find(name) !== null,
      readXml: async (name) => {
        const entry = find(name);
        if (!entry) throw new Error(`${name} がありません`);
        const data = await zip.readEntry(bytes, entry, inflateRaw);
        return xml.parseXml(zip.decodeText(data));
      },
    };
  };

  /**
   * 部品の場所から見た参照先を、パッケージの中の名前にする
   * @param {string} base  参照元の部品の名前（例: xl/workbook.xml）
   * @param {string} target 例: worksheets/sheet1.xml、../drawings/drawing1.xml、/xl/styles.xml
   */
  const resolvePath = (base, target) => {
    const clean = target.replace(/\\/g, '/');
    const parts = clean.startsWith('/') ? [] : base.split('/').slice(0, -1);
    for (const seg of clean.replace(/^\/+/, '').split('/')) {
      if (seg === '' || seg === '.') continue;
      if (seg === '..') parts.pop();
      else parts.push(seg);
    }
    return parts.join('/');
  };

  /**
   * 部品の参照の一覧（_rels/xxx.rels）。なければ空の一覧
   * @param {Package} pkg
   * @param {string} partName  例: xl/workbook.xml（パッケージ全体の参照は空文字）
   * @returns {Promise<Relationship[]>}
   */
  const readRels = async (pkg, partName) => {
    const dir = partName.includes('/') ? partName.slice(0, partName.lastIndexOf('/') + 1) : '';
    const file = partName === '' ? '_rels/.rels' : `${dir}_rels/${partName.slice(dir.length)}.rels`;
    if (!pkg.has(file)) return [];
    const relsRoot = await pkg.readXml(file);
    return xml.kids(relsRoot, null, 'Relationship').map((r) => {
      const external = (xml.attr(r, 'TargetMode') ?? '').toLowerCase() === 'external';
      const target = xml.attr(r, 'Target') ?? '';
      return {
        id: xml.attr(r, 'Id') ?? '',
        type: xml.attr(r, 'Type') ?? '',
        target: external ? '' : resolvePath(partName === '' ? '' : partName, target),
      };
    });
  };

  /**
   * 参照の種類（URI の最後の部分）で探す
   * @param {Relationship[]} rels
   * @param {string} kind  例: officeDocument、worksheet、drawing
   */
  const relsOfType = (rels, kind) => rels.filter((r) => r.type.split('/').pop() === kind && r.target !== '');

  /**
   * mc:AlternateContent（新旧2通りの書き方）を、最初の書き方の中身に置き換えた子の一覧
   * @param {XmlElement} el
   * @returns {XmlElement[]}
   */
  const unwrapAlternate = (el) =>
    xml.kids(el).flatMap((c) => {
      if (!xml.is(c, NS.mc, 'AlternateContent')) return [c];
      const choice = xml.kid(c, NS.mc, 'Choice') ?? xml.kid(c, NS.mc, 'Fallback');
      return choice ? unwrapAlternate(choice) : [];
    });

  /**
   * Excel の文字の中の _xHHHH_（制御文字などの書き方）を文字に戻し、改行を \n にそろえる
   * @param {string} s
   */
  const decodeXString = (s) =>
    s.replace(/_x([0-9A-Fa-f]{4})_/g, (_, hex) => String.fromCharCode(parseInt(hex, 16))).replace(/\r\n?/g, '\n');

  root.YegOoxml = Object.freeze({ NS, openPackage, resolvePath, readRels, relsOfType, unwrapAlternate, decodeXString });
})(globalThis);
