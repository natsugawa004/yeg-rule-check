// @ts-check
/*
 * Word（.docx）から、判定に使う文字を取り出す（画面に依存しない）
 *   readDocx(入れ物) → 「ブロック」の並び（本文の段落のまとまり・表のセル1つ・テキストボックス1つが1ブロック）
 * ・本文の段落は1行ずつ改行でつなぐ（貼り付けたときと同じく、改行をまたいだ語も見つける）
 * ・表は1行ごとに、セルをタブでつなぐ。セルの境目はまたがない
 * ・テキストボックスの文字は、置かれている段落の後に出す（新旧2通りの書き方がある場合は1回だけ読む）
 * ・削除の記録が残っている文字、フィールドの命令文、ふりがなは読まない。ヘッダー・フッターは読まない
 * S3（空欄の判定）で使えるよう、表の行・列・結合の情報を残す
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
   */

  /**
   * @typedef {Object} DocxCell
   * @property {number} row      表の中の行（1 から）
   * @property {number} col      行の中のセルの順番（1 から）
   * @property {number} gridCol  表の列の位置（1 から。横に結合されたセルは左端の位置）
   * @property {number} gridSpan 横に結合している列の数
   * @property {'restart'|'continue'|null} vMerge  縦の結合（continue は上のセルと結合された空のセル）
   * @property {string} text
   */
  /** @typedef {{ index: number, rows: DocxCell[][] }} DocxTable */
  /**
   * @typedef {{ type: 'body' }
   *   | { type: 'tableCell', table: number, row: number, col: number, gridCol: number }
   *   | { type: 'textbox', index: number }} DocxLoc
   */
  /**
   * @typedef {Object} DocxBlock
   * @property {string} text
   * @property {'\n'|'\t'} sep
   * @property {DocxLoc} loc
   */
  /**
   * @typedef {Object} DocxUnit
   * @property {string} id
   * @property {'document'} kind
   * @property {string} label
   * @property {boolean} hidden
   * @property {boolean} hasText
   * @property {DocxBlock[]} blocks
   * @property {{ tables: DocxTable[] }} docx
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (globalThis));
  const xml = /** @type {XmlApi} */ (g.YegXml);
  const ooxml = /** @type {OoxmlApi} */ (g.YegOoxml);
  const { NS } = ooxml;
  const W = NS.wml;

  /** 中の文字を読まない要素（書式・削除の記録・フィールドの命令・脚注の印など） */
  const SKIP = new Set([
    'pPr',
    'rPr',
    'tblPr',
    'tblGrid',
    'trPr',
    'tcPr',
    'sectPr',
    'del',
    'delText',
    'instrText',
    'delInstrText',
    'moveFrom',
    'rt',
    'footnoteReference',
    'endnoteReference',
    'commentReference',
    'annotationRef',
    'fldChar',
    'proofErr',
    'bookmarkStart',
    'bookmarkEnd',
    'lastRenderedPageBreak',
    'permStart',
    'permEnd',
    'sdtPr',
    'sdtEndPr',
  ]);

  /**
   * @typedef {Object} InlineOut
   * @property {string[]} parts  段落の文字
   * @property {string[]} boxes  段落の中にあったテキストボックスの文字
   */

  /**
   * 段落の中を読む
   * @param {XmlElement} el
   * @param {InlineOut} out
   */
  const readInline = (el, out) => {
    for (const c of ooxml.unwrapAlternate(el)) {
      if (W.has(c.ns)) {
        if (SKIP.has(c.local)) continue;
        switch (c.local) {
          case 't':
            out.parts.push(xml.textOf(c));
            break;
          case 'tab':
          case 'ptab':
            out.parts.push('\t');
            break;
          case 'br':
          case 'cr':
            out.parts.push('\n');
            break;
          case 'noBreakHyphen':
            out.parts.push('-');
            break;
          case 'sym': {
            const code = parseInt(xml.attr(c, 'char') ?? '', 16);
            // 記号フォント（Wingdings など）の字は私用領域にあり、文字として意味を持たないので読まない
            if (Number.isFinite(code) && !(code >= 0xe000 && code <= 0xf8ff)) out.parts.push(String.fromCodePoint(code));
            break;
          }
          case 'txbxContent':
            out.boxes.push(readBlockText(c));
            break;
          case 'ruby': {
            const base = xml.kid(c, W, 'rubyBase');
            if (base) readInline(base, out);
            break;
          }
          default:
            readInline(c, out);
        }
      } else if (NS.math.has(c.ns) && c.local === 't') {
        out.parts.push(xml.textOf(c));
      } else {
        // 図形（w:drawing の中の wps:txbx、w:pict の中の v:textbox など）はさらに中を探す
        readInline(c, out);
      }
    }
  };

  /**
   * @param {XmlElement} p
   * @returns {{ text: string, boxes: string[] }}
   */
  const readParagraph = (p) => {
    /** @type {InlineOut} */
    const out = { parts: [], boxes: [] };
    readInline(p, out);
    return { text: out.parts.join(''), boxes: out.boxes };
  };

  /** @param {string} s */
  const trimBlankLines = (s) => s.replace(/^(?:[ \t　]*\n)+/, '').replace(/(?:\n[ \t　]*)+$/, '');

  /**
   * 段落・表の並び（テキストボックスやセルの中身）を、1つの文字にまとめる
   * @param {XmlElement} container
   * @returns {string}
   */
  const readBlockText = (container) => {
    /** @type {string[]} */
    const lines = [];
    /** @param {XmlElement} el */
    const walk = (el) => {
      for (const c of ooxml.unwrapAlternate(el)) {
        if (!W.has(c.ns)) continue;
        if (c.local === 'p') {
          const { text, boxes } = readParagraph(c);
          lines.push(text, ...boxes);
        } else if (c.local === 'tbl') {
          for (const tr of tableRows(c)) lines.push(rowCells(tr).map((tc) => readBlockText(tc).replace(/\n/g, ' ')).join('\t'));
        } else if (c.local === 'sdt') {
          const content = xml.kid(c, W, 'sdtContent');
          if (content) walk(content);
        } else if (c.local === 'customXml' || c.local === 'ins' || c.local === 'moveTo') {
          walk(c);
        }
      }
    };
    walk(container);
    return trimBlankLines(lines.join('\n'));
  };

  /**
   * 表の行（コンテンツコントロールなどで包まれた行も）
   * @param {XmlElement} tbl
   * @returns {XmlElement[]}
   */
  const tableRows = (tbl) =>
    ooxml.unwrapAlternate(tbl).flatMap((c) => {
      if (!W.has(c.ns)) return [];
      if (c.local === 'tr') return [c];
      if (c.local === 'sdt') {
        const content = xml.kid(c, W, 'sdtContent');
        return content ? tableRows(content) : [];
      }
      return c.local === 'customXml' ? tableRows(c) : [];
    });

  /**
   * 行のセル
   * @param {XmlElement} tr
   * @returns {XmlElement[]}
   */
  const rowCells = (tr) =>
    ooxml.unwrapAlternate(tr).flatMap((c) => {
      if (!W.has(c.ns)) return [];
      if (c.local === 'tc') return [c];
      if (c.local === 'sdt') {
        const content = xml.kid(c, W, 'sdtContent');
        return content ? rowCells(content) : [];
      }
      return c.local === 'customXml' ? rowCells(c) : [];
    });

  /**
   * @param {XmlElement|null} parent
   * @param {string} local
   */
  const valOf = (parent, local) => {
    const el = parent ? xml.kid(parent, W, local) : null;
    return el ? xml.attr(el, 'val') : null;
  };

  /**
   * @param {XmlElement} tbl
   * @param {number} index
   * @returns {DocxTable}
   */
  const readTable = (tbl, index) => ({
    index,
    rows: tableRows(tbl).map((tr, ri) => {
      const before = Number(valOf(xml.kid(tr, W, 'trPr'), 'gridBefore') ?? '0');
      let gridCol = Number.isInteger(before) ? before + 1 : 1;
      return rowCells(tr).map((tc, ci) => {
        const tcPr = xml.kid(tc, W, 'tcPr');
        const span = Number(valOf(tcPr, 'gridSpan') ?? '1');
        const vMergeEl = tcPr ? xml.kid(tcPr, W, 'vMerge') : null;
        /** @type {'restart'|'continue'|null} */
        const vMerge = vMergeEl ? (xml.attr(vMergeEl, 'val') === 'restart' ? 'restart' : 'continue') : null;
        const gridSpan = Number.isInteger(span) && span > 0 ? span : 1;
        const cell = { row: ri + 1, col: ci + 1, gridCol, gridSpan, vMerge, text: readBlockText(tc) };
        gridCol += gridSpan;
        return cell;
      });
    }),
  });

  /**
   * Word のファイルを読む
   * @param {Package} pkg
   * @returns {Promise<{ units: DocxUnit[], notes: string[] }>}
   */
  const readDocx = async (pkg) => {
    const rootRels = await ooxml.readRels(pkg, '');
    const docPath = ooxml.relsOfType(rootRels, 'officeDocument')[0]?.target ?? 'word/document.xml';
    if (!pkg.has(docPath)) throw new Error('本文（document.xml）が見つかりません');
    const doc = await pkg.readXml(docPath);
    const body = xml.kid(doc, W, 'body');
    if (!body) throw new Error('本文（body）が見つかりません');

    /** @type {DocxBlock[]} */
    const blocks = [];
    /** @type {DocxTable[]} */
    const tables = [];
    /** @type {string[]} */
    let flow = [];
    let boxCount = 0;

    const flushFlow = () => {
      const text = trimBlankLines(flow.join('\n'));
      if (text.trim() !== '') blocks.push({ text, sep: '\n', loc: { type: 'body' } });
      flow = [];
    };

    /** @param {XmlElement} el */
    const walk = (el) => {
      for (const c of ooxml.unwrapAlternate(el)) {
        if (!W.has(c.ns)) continue;
        if (c.local === 'p') {
          const { text, boxes } = readParagraph(c);
          flow.push(text);
          const filled = boxes.filter((b) => b.trim() !== '');
          if (filled.length > 0) {
            flushFlow();
            for (const b of filled) {
              boxCount += 1;
              blocks.push({ text: b, sep: '\n', loc: { type: 'textbox', index: boxCount } });
            }
          }
        } else if (c.local === 'tbl') {
          flushFlow();
          const table = readTable(c, tables.length + 1);
          tables.push(table);
          for (const row of table.rows) {
            let first = true;
            for (const cell of row) {
              if (cell.text.trim() === '') continue;
              blocks.push({
                text: cell.text,
                sep: first ? '\n' : '\t',
                loc: { type: 'tableCell', table: table.index, row: cell.row, col: cell.col, gridCol: cell.gridCol },
              });
              first = false;
            }
          }
        } else if (c.local === 'sdt') {
          const content = xml.kid(c, W, 'sdtContent');
          if (content) walk(content);
        } else if (c.local === 'customXml' || c.local === 'ins' || c.local === 'moveTo') {
          walk(c);
        }
      }
    };
    walk(body);
    flushFlow();

    return {
      units: [
        {
          id: 'document',
          kind: 'document',
          label: '文書全体',
          hidden: false,
          hasText: blocks.length > 0,
          blocks,
          docx: { tables },
        },
      ],
      notes: [],
    };
  };

  root.YegDocx = Object.freeze({ readDocx });
})(globalThis);
