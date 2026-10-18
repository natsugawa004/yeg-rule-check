// @ts-check
/*
 * Excel（.xlsx）から、判定に使う文字をシートごとに取り出す（画面に依存しない）
 *   readXlsx(入れ物) → シートごとの「ブロック」（セル1つ・図形1つが1ブロック）と、セル・図形の一覧
 * ・セルは画面に見える形（桁区切り・¥・日付の書式）の文字にする。数式は保存されている計算結果を読む
 * ・テキストボックスなどの図形の中の文字も読む。図形の位置は、左上が置かれたセル（例: B11）で表す
 * ・ふりがな（読み）の文字は読まない
 * ・ブロックの並びは、行ごとに左から右（同じ行のセルはタブでつなぐ）。図形はその行の後に置く
 * S3（空欄の判定）・S4（金額チェック）で使えるよう、セルの番地・値・数式の有無・結合範囲・図形の位置を残す
 */
'use strict';

(function (root) {
  /**
   * @typedef {Object} XmlElement
   * @property {string} name
   * @property {string} local
   * @property {string} ns
   * @property {Readonly<Record<string, string>>} attrs
   * @property {ReadonlyMap<string, string>} scope
   * @property {(XmlElement|string)[]} children
   */
  /**
   * @typedef {Object} XmlApi
   * @property {(el: XmlElement, nsSet: ReadonlySet<string>|null, local: string) => boolean} is
   * @property {(el: XmlElement, nsSet?: ReadonlySet<string>|null, local?: string) => XmlElement[]} kids
   * @property {(el: XmlElement, nsSet: ReadonlySet<string>|null, local: string) => XmlElement|null} kid
   * @property {(el: XmlElement, local: string) => string|null} attr
   * @property {(el: XmlElement, nsSet: ReadonlySet<string>, local: string) => string|null} attrNs
   * @property {(el: XmlElement) => string} textOf
   */
  /**
   * @typedef {Object} Package
   * @property {(name: string) => boolean} has
   * @property {(name: string) => Promise<XmlElement>} readXml
   */
  /** @typedef {{ id: string, type: string, target: string }} Relationship */
  /**
   * @typedef {Object} OoxmlApi
   * @property {Record<string, ReadonlySet<string>>} NS
   * @property {(pkg: Package, part: string) => Promise<Relationship[]>} readRels
   * @property {(rels: Relationship[], kind: string) => Relationship[]} relsOfType
   * @property {(el: XmlElement) => XmlElement[]} unwrapAlternate
   * @property {(s: string) => string} decodeXString
   */
  /**
   * @typedef {Object} NumFmtApi
   * @property {(value: number, code: string, opts?: { date1904?: boolean }) => string} formatNumber
   * @property {(text: string, code: string) => string} formatText
   * @property {(id: number, custom: ReadonlyMap<number, string>) => string} formatCodeOf
   */

  /**
   * @typedef {Object} XlsxCell
   * @property {string} ref     番地（例: B4）
   * @property {number} row     1 から
   * @property {number} col     1 から（A=1）
   * @property {string} text    画面に見える文字
   * @property {'string'|'number'|'boolean'|'error'|'date'|'empty'} kind
   * @property {number|null} value  数値のセルの値（数式なら保存されている計算結果）
   * @property {boolean} formula    数式のセル
   * @property {boolean} cached     数式の計算結果がファイルに保存されている（数式でなければ true）
   * @property {string} numFmt      表示形式
   */
  /**
   * @typedef {Object} XlsxShape
   * @property {string} name     図形の名前（例: TextBox 1）
   * @property {boolean} textBox テキストボックスとして作られた図形
   * @property {string} ref      左上が置かれたセルの番地（位置がセルで決まらない図形は空文字）
   * @property {number} row      左上のセルの行（1 から。位置がセルで決まらない図形は 0）
   * @property {number} col
   * @property {number|null} toRow  右下のセルの行（わかるときだけ）
   * @property {number|null} toCol
   * @property {string} text
   * @property {number} areaRow  図形の上の端が実際にかかっている行（左上のセルの下半分より下から始まる図形は、次の行にする。
   *   行の高さから計算する。S3 で図形がどの欄の記入欄にあるかを決めるのに使う。位置がセルで決まらない図形は 0）
   */
  /** @typedef {{ ref: string, r1: number, c1: number, r2: number, c2: number }} MergeRange */
  /**
   * @typedef {Object} XlsxSheet
   * @property {string} name
   * @property {boolean} hidden
   * @property {XlsxCell[]} cells    文字か数式のあるセルだけ（行・列の順）
   * @property {MergeRange[]} merges
   * @property {XlsxShape[]} shapes
   */
  /**
   * @typedef {{ type: 'cell', sheet: string, ref: string, row: number, col: number }
   *   | { type: 'shape', sheet: string, ref: string, row: number, col: number, name: string }} XlsxLoc
   */
  /**
   * @typedef {Object} XlsxBlock
   * @property {string} text
   * @property {'\n'|'\t'} sep  前のブロックとの区切り
   * @property {XlsxLoc} loc
   */
  /**
   * @typedef {Object} XlsxUnit
   * @property {string} id
   * @property {'sheet'} kind
   * @property {string} label
   * @property {boolean} hidden
   * @property {boolean} hasText
   * @property {XlsxBlock[]} blocks
   * @property {XlsxSheet} sheet
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (globalThis));
  const xml = /** @type {XmlApi} */ (g.YegXml);
  const ooxml = /** @type {OoxmlApi} */ (g.YegOoxml);
  const numfmt = /** @type {NumFmtApi} */ (g.YegNumFmt);
  const { NS } = ooxml;

  // ------------------------------------------------------------------
  // 番地
  // ------------------------------------------------------------------

  /**
   * 列の番号（1 から）を A・B…・AA にする
   * @param {number} col
   */
  const colName = (col) => {
    let n = col;
    let s = '';
    while (n > 0) {
      const r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  };

  /**
   * @param {number} row
   * @param {number} col
   */
  const cellRef = (row, col) => `${colName(col)}${row}`;

  /**
   * 番地（B4、$B$4）を行・列にする。読めなければ null
   * @param {string} ref
   * @returns {{ row: number, col: number }|null}
   */
  const parseRef = (ref) => {
    const m = /^\$?([A-Za-z]{1,3})\$?(\d+)$/.exec(ref.trim());
    if (!m) return null;
    const col = [...m[1].toUpperCase()].reduce((n, ch) => n * 26 + (ch.charCodeAt(0) - 64), 0);
    return { row: Number(m[2]), col };
  };

  /**
   * @param {string} ref  例: B18:C18
   * @returns {MergeRange|null}
   */
  const parseRange = (ref) => {
    const [a, b] = ref.split(':');
    const p = parseRef(a);
    const q = b ? parseRef(b) : p;
    if (!p || !q) return null;
    return {
      ref,
      r1: Math.min(p.row, q.row),
      c1: Math.min(p.col, q.col),
      r2: Math.max(p.row, q.row),
      c2: Math.max(p.col, q.col),
    };
  };

  // ------------------------------------------------------------------
  // 文字（共有文字列・リッチテキスト）
  // ------------------------------------------------------------------

  /**
   * <si> や <is> の文字。ふりがな（rPh）は読まない
   * @param {XmlElement} el
   */
  const richText = (el) =>
    ooxml.decodeXString(
      xml
        .kids(el)
        .map((c) => {
          if (xml.is(c, NS.sml, 't')) return xml.textOf(c);
          if (xml.is(c, NS.sml, 'r')) {
            const t = xml.kid(c, NS.sml, 't');
            return t ? xml.textOf(t) : '';
          }
          return '';
        })
        .join(''),
    );

  /**
   * @param {Package} pkg
   * @param {string|null} path
   * @returns {Promise<string[]>}
   */
  const readSharedStrings = async (pkg, path) => {
    if (!path || !pkg.has(path)) return [];
    const sst = await pkg.readXml(path);
    return xml.kids(sst, NS.sml, 'si').map(richText);
  };

  /**
   * セルの書式番号 → 表示形式
   * @param {Package} pkg
   * @param {string|null} path
   * @returns {Promise<string[]>}
   */
  const readCellFormats = async (pkg, path) => {
    if (!path || !pkg.has(path)) return [];
    const styles = await pkg.readXml(path);
    /** @type {Map<number, string>} */
    const custom = new Map();
    const numFmts = xml.kid(styles, NS.sml, 'numFmts');
    for (const f of numFmts ? xml.kids(numFmts, NS.sml, 'numFmt') : []) {
      const id = Number(xml.attr(f, 'numFmtId'));
      const code = xml.attr(f, 'formatCode');
      if (Number.isInteger(id) && code !== null) custom.set(id, code);
    }
    const cellXfs = xml.kid(styles, NS.sml, 'cellXfs');
    return (cellXfs ? xml.kids(cellXfs, NS.sml, 'xf') : []).map((xf) =>
      numfmt.formatCodeOf(Number(xml.attr(xf, 'numFmtId') ?? '0'), custom),
    );
  };

  // ------------------------------------------------------------------
  // セル
  // ------------------------------------------------------------------

  /**
   * @param {XmlElement} c
   * @param {number} row
   * @param {number} col
   * @param {{ shared: string[], formats: string[], date1904: boolean }} ctx
   * @returns {XlsxCell}
   */
  const readCell = (c, row, col, ctx) => {
    const t = xml.attr(c, 't') ?? 'n';
    const s = Number(xml.attr(c, 's') ?? '0');
    const numFmt = ctx.formats[s] ?? 'General';
    const vEl = xml.kid(c, NS.sml, 'v');
    const v = vEl ? xml.textOf(vEl) : null;
    const formula = xml.kid(c, NS.sml, 'f') !== null;
    const base = { ref: cellRef(row, col), row, col, formula, numFmt };

    if (t === 's') {
      const text = v !== null ? ctx.shared[Number(v)] ?? '' : '';
      return { ...base, text: numfmt.formatText(text, numFmt), kind: 'string', value: null, cached: true };
    }
    if (t === 'inlineStr') {
      const is = xml.kid(c, NS.sml, 'is');
      const text = is ? richText(is) : '';
      return { ...base, text: numfmt.formatText(text, numFmt), kind: 'string', value: null, cached: true };
    }
    if (t === 'str') {
      const text = v !== null ? ooxml.decodeXString(v) : '';
      return { ...base, text: numfmt.formatText(text, numFmt), kind: 'string', value: null, cached: !formula || vEl !== null };
    }
    if (t === 'b') {
      if (v === null || v === '') return { ...base, text: '', kind: 'empty', value: null, cached: !formula };
      return { ...base, text: v === '1' || v.toLowerCase() === 'true' ? 'TRUE' : 'FALSE', kind: 'boolean', value: null, cached: true };
    }
    if (t === 'e') {
      return { ...base, text: v ?? '', kind: 'error', value: null, cached: !formula || v !== null };
    }
    if (t === 'd') {
      return { ...base, text: v ?? '', kind: 'date', value: null, cached: !formula || v !== null };
    }
    if (v === null || v.trim() === '') {
      return { ...base, text: '', kind: 'empty', value: null, cached: !formula };
    }
    const num = Number(v);
    if (!Number.isFinite(num)) return { ...base, text: v, kind: 'string', value: null, cached: true };
    return {
      ...base,
      text: numfmt.formatNumber(num, numFmt, { date1904: ctx.date1904 }),
      kind: 'number',
      value: num,
      cached: true,
    };
  };

  /**
   * @param {XmlElement} ws
   * @param {{ shared: string[], formats: string[], date1904: boolean }} ctx
   * @returns {XlsxCell[]}
   */
  const readCells = (ws, ctx) => {
    const sheetData = xml.kid(ws, NS.sml, 'sheetData');
    if (!sheetData) return [];
    /** @type {XlsxCell[]} */
    const cells = [];
    let rowNo = 0;
    for (const r of xml.kids(sheetData, NS.sml, 'row')) {
      const rAttr = Number(xml.attr(r, 'r'));
      rowNo = Number.isInteger(rAttr) && rAttr > 0 ? rAttr : rowNo + 1;
      let colNo = 0;
      for (const c of xml.kids(r, NS.sml, 'c')) {
        const pos = parseRef(xml.attr(c, 'r') ?? '');
        colNo = pos ? pos.col : colNo + 1;
        const row = pos ? pos.row : rowNo;
        const cell = readCell(c, row, colNo, ctx);
        if (cell.text !== '' || cell.formula) cells.push(cell);
      }
    }
    return cells.sort((a, b) => a.row - b.row || a.col - b.col);
  };

  // ------------------------------------------------------------------
  // 図形（テキストボックスなど）
  // ------------------------------------------------------------------

  /**
   * DrawingML の文字（<a:p> ごとに改行）
   * @param {XmlElement} txBody
   */
  const drawingText = (txBody) =>
    xml
      .kids(txBody, NS.a, 'p')
      .map((p) =>
        xml
          .kids(p)
          .map((c) => {
            if (xml.is(c, NS.a, 'r') || xml.is(c, NS.a, 'fld')) {
              const t = xml.kid(c, NS.a, 't');
              return t ? xml.textOf(t) : '';
            }
            return xml.is(c, NS.a, 'br') ? '\n' : '';
          })
          .join(''),
      )
      .join('\n')
      .replace(/\r\n?/g, '\n')
      .replace(/^\n+|\n+$/g, '');

  /**
   * 図形の位置（xdr:from / xdr:to）。rowOff はセルの上の端からのずれ（EMU）
   * @param {XmlElement|null} el
   * @returns {{ row: number, col: number, rowOff: number }|null}
   */
  const anchorPoint = (el) => {
    if (!el) return null;
    const colEl = xml.kid(el, NS.xdr, 'col');
    const rowEl = xml.kid(el, NS.xdr, 'row');
    if (!colEl || !rowEl) return null;
    const col = Number(xml.textOf(colEl));
    const row = Number(xml.textOf(rowEl));
    const offEl = xml.kid(el, NS.xdr, 'rowOff');
    const off = offEl ? Number(xml.textOf(offEl)) : 0;
    return Number.isInteger(col) && Number.isInteger(row)
      ? { row: row + 1, col: col + 1, rowOff: Number.isFinite(off) && off > 0 ? off : 0 }
      : null;
  };

  /** 1ポイントあたりの EMU（図形の位置の単位） */
  const EMU_PER_POINT = 12700;

  /**
   * 行の高さ（ポイント）を返す関数を作る。非表示の行は 0
   * @param {XmlElement} ws
   * @returns {(row: number) => number}
   */
  const rowHeights = (ws) => {
    const fmt = xml.kid(ws, NS.sml, 'sheetFormatPr');
    const def = Number(fmt ? xml.attr(fmt, 'defaultRowHeight') ?? '' : '');
    const fallback = Number.isFinite(def) && def > 0 ? def : 15;
    /** @type {Map<number, number>} */
    const heights = new Map();
    const sheetData = xml.kid(ws, NS.sml, 'sheetData');
    for (const r of sheetData ? xml.kids(sheetData, NS.sml, 'row') : []) {
      const no = Number(xml.attr(r, 'r'));
      if (!Number.isInteger(no)) continue;
      const hidden = ['1', 'true'].includes((xml.attr(r, 'hidden') ?? '').toLowerCase());
      const htAttr = xml.attr(r, 'ht');
      const ht = htAttr === null ? NaN : Number(htAttr);
      if (hidden) heights.set(no, 0);
      else if (Number.isFinite(ht) && ht >= 0) heights.set(no, ht);
    }
    return (row) => heights.get(row) ?? fallback;
  };

  /**
   * 図形の上の端が実際にかかっている行。左上のセルからのずれ（rowOff）が行の高さの半分以上なら、次の行にする
   * （例: 告知方法の欄の最後の行の下の端ぎりぎりから始まる、実施スケジュールのテキストボックス）
   * @param {number} row
   * @param {number} rowOff  EMU
   * @param {(row: number) => number} heightOf
   */
  const effectiveRow = (row, rowOff, heightOf) => {
    let r = row;
    let off = rowOff / EMU_PER_POINT;
    for (let i = 0; i < 1000; i += 1) {
      const h = heightOf(r);
      if (h > 0 && off < h) break;
      off -= Math.max(h, 0);
      r += 1;
    }
    const h = heightOf(r);
    return h > 0 && off >= h / 2 ? r + 1 : r;
  };

  /**
   * 図形（グループの中も）から、文字を持つものを集める
   * @param {XmlElement} el
   * @returns {{ name: string, textBox: boolean, text: string }[]}
   */
  const collectShapes = (el) =>
    ooxml.unwrapAlternate(el).flatMap((c) => {
      if (xml.is(c, NS.xdr, 'grpSp')) return collectShapes(c);
      if (!xml.is(c, NS.xdr, 'sp')) return [];
      const nv = xml.kid(c, NS.xdr, 'nvSpPr');
      const cNvPr = nv ? xml.kid(nv, NS.xdr, 'cNvPr') : null;
      const cNvSpPr = nv ? xml.kid(nv, NS.xdr, 'cNvSpPr') : null;
      if (cNvPr && ['1', 'true'].includes((xml.attr(cNvPr, 'hidden') ?? '').toLowerCase())) return [];
      const txBody = xml.kid(c, NS.xdr, 'txBody');
      return [
        {
          name: cNvPr ? xml.attr(cNvPr, 'name') ?? '' : '',
          textBox: cNvSpPr ? ['1', 'true'].includes((xml.attr(cNvSpPr, 'txBox') ?? '').toLowerCase()) : false,
          text: txBody ? drawingText(txBody) : '',
        },
      ];
    });

  /**
   * @param {XmlElement} wsDr
   * @param {(row: number) => number} heightOf  行の高さ（ポイント）
   * @returns {XlsxShape[]}
   */
  const readShapes = (wsDr, heightOf) =>
    ooxml.unwrapAlternate(wsDr).flatMap((anchor) => {
      const isTwo = xml.is(anchor, NS.xdr, 'twoCellAnchor');
      const isOne = xml.is(anchor, NS.xdr, 'oneCellAnchor');
      const isAbs = xml.is(anchor, NS.xdr, 'absoluteAnchor');
      if (!isTwo && !isOne && !isAbs) return [];
      const from = isAbs ? null : anchorPoint(xml.kid(anchor, NS.xdr, 'from'));
      const to = isTwo ? anchorPoint(xml.kid(anchor, NS.xdr, 'to')) : null;
      return collectShapes(anchor).map((s) => ({
        ...s,
        ref: from ? cellRef(from.row, from.col) : '',
        row: from ? from.row : 0,
        col: from ? from.col : 0,
        toRow: to ? to.row : null,
        toCol: to ? to.col : null,
        areaRow: from ? effectiveRow(from.row, from.rowOff, heightOf) : 0,
      }));
    });

  /**
   * @param {Package} pkg
   * @param {string} sheetPath
   * @param {(row: number) => number} heightOf  行の高さ（ポイント）
   * @returns {Promise<XlsxShape[]>}
   */
  const readSheetShapes = async (pkg, sheetPath, heightOf) => {
    const rels = await ooxml.readRels(pkg, sheetPath);
    /** @type {XlsxShape[]} */
    const out = [];
    for (const rel of ooxml.relsOfType(rels, 'drawing')) {
      if (!pkg.has(rel.target)) continue;
      out.push(...readShapes(await pkg.readXml(rel.target), heightOf));
    }
    return out;
  };

  // ------------------------------------------------------------------
  // ブロック（判定する文字のまとまり）
  // ------------------------------------------------------------------

  /**
   * セル・図形を読む順に並べ、ブロックにする
   * @param {string} sheetName
   * @param {XlsxCell[]} cells
   * @param {XlsxShape[]} shapes
   * @returns {XlsxBlock[]}
   */
  const buildBlocks = (sheetName, cells, shapes) => {
    const items = [
      ...cells
        .filter((c) => c.text.trim() !== '')
        .map((c) => ({ row: c.row, order: 0, col: c.col, text: c.text, cell: c, shape: null })),
      ...shapes
        .filter((s) => s.text.trim() !== '')
        .map((s) => ({ row: s.row === 0 ? Number.MAX_SAFE_INTEGER : s.row, order: 1, col: s.col, text: s.text, cell: null, shape: s })),
    ].sort((a, b) => a.row - b.row || a.order - b.order || a.col - b.col);

    return items.map((it, i) => {
      const prev = items[i - 1];
      /** @type {'\n'|'\t'} */
      const sep = prev && it.cell && prev.cell && prev.row === it.row ? '\t' : '\n';
      /** @type {XlsxLoc} */
      const loc = it.cell
        ? { type: 'cell', sheet: sheetName, ref: it.cell.ref, row: it.cell.row, col: it.cell.col }
        : {
            type: 'shape',
            sheet: sheetName,
            ref: it.shape ? it.shape.ref : '',
            row: it.shape ? it.shape.row : 0,
            col: it.shape ? it.shape.col : 0,
            name: it.shape ? it.shape.name : '',
          };
      return { text: it.text, sep, loc };
    });
  };

  /**
   * Excel のファイルを読む
   * @param {Package} pkg
   * @returns {Promise<{ units: XlsxUnit[], notes: string[] }>}
   */
  const readXlsx = async (pkg) => {
    const rootRels = await ooxml.readRels(pkg, '');
    const wbPath = ooxml.relsOfType(rootRels, 'officeDocument')[0]?.target ?? 'xl/workbook.xml';
    if (!pkg.has(wbPath)) throw new Error('ブック（workbook.xml）が見つかりません');
    const wb = await pkg.readXml(wbPath);
    const wbRels = await ooxml.readRels(pkg, wbPath);
    const relById = new Map(wbRels.map((r) => [r.id, r]));
    const wbPr = xml.kid(wb, NS.sml, 'workbookPr');
    const date1904 = wbPr !== null && ['1', 'true'].includes((xml.attr(wbPr, 'date1904') ?? '').toLowerCase());

    const shared = await readSharedStrings(pkg, ooxml.relsOfType(wbRels, 'sharedStrings')[0]?.target ?? null);
    const formats = await readCellFormats(pkg, ooxml.relsOfType(wbRels, 'styles')[0]?.target ?? null);
    const ctx = { shared, formats, date1904 };

    const sheetsEl = xml.kid(wb, NS.sml, 'sheets');
    const sheetDefs = sheetsEl ? xml.kids(sheetsEl, NS.sml, 'sheet') : [];
    /** @type {XlsxUnit[]} */
    const units = [];
    /** @type {string[]} */
    const uncached = [];
    for (const def of sheetDefs) {
      const name = xml.attr(def, 'name') ?? `シート${units.length + 1}`;
      const rid = xml.attrNs(def, NS.rel, 'id') ?? '';
      const rel = relById.get(rid);
      // グラフだけのシート・マクロのシートは、セルを持たないので読まない
      if (!rel || rel.type.split('/').pop() !== 'worksheet' || !pkg.has(rel.target)) continue;
      const state = (xml.attr(def, 'state') ?? 'visible').toLowerCase();
      const ws = await pkg.readXml(rel.target);
      const cells = readCells(ws, ctx);
      const mergeEl = xml.kid(ws, NS.sml, 'mergeCells');
      const merges = (mergeEl ? xml.kids(mergeEl, NS.sml, 'mergeCell') : []).flatMap((m) => {
        const r = parseRange(xml.attr(m, 'ref') ?? '');
        return r ? [r] : [];
      });
      const shapes = await readSheetShapes(pkg, rel.target, rowHeights(ws));
      const blocks = buildBlocks(name, cells, shapes);
      uncached.push(...cells.filter((c) => !c.cached).map((c) => `${name}／${c.ref}`));
      units.push({
        id: `sheet-${units.length + 1}`,
        kind: 'sheet',
        label: name,
        hidden: state !== 'visible',
        hasText: blocks.length > 0,
        blocks,
        sheet: { name, hidden: state !== 'visible', cells, merges, shapes },
      });
    }
    if (units.length === 0) throw new Error('ワークシートが見つかりません');
    const notes =
      uncached.length > 0
        ? [
            `計算結果が保存されていないセルがあります（${uncached.slice(0, 3).join('、')}${uncached.length > 3 ? ` など${uncached.length}か所` : ''}）。そのセルは空欄として読みました。Excelで開いて保存し直すと読めます。`,
          ]
        : [];
    return { units, notes };
  };

  root.YegXlsx = Object.freeze({ readXlsx, colName, cellRef, parseRef, parseRange, buildBlocks });
})(globalThis);
