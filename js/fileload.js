// @ts-check
/*
 * ファイルの読み込みの入口と、読み込んだ結果の扱い（画面に依存しない）
 *   checkSize(大きさ) / loadFile({ name, bytes }, 依存) → 読み込んだ書類（ページ・シートの単位と、文字のブロック）
 *   compose(単位の一覧, 判定に含める単位) → 判定する文字と、各ブロックの位置・またがない位置
 *   locate / locationLabel → 指摘の位置を「01計画／B4」「2ページ」のような場所にする
 * 形式は拡張子ではなく中身で見分ける（拡張子は案内の文言にだけ使う）
 */
'use strict';

(function (root) {
  /** @typedef {'docx'|'xlsx'|'pdf'} Format */
  /**
   * @typedef {'too-large'|'empty-file'|'unsupported'|'password'|'broken'|'no-text'|'pdf-no-text'|'pdf-unavailable'} LoadErrorCode
   */
  /**
   * @typedef {Object} LoadError
   * @property {LoadErrorCode} code
   * @property {string} message  利用者に見せる文
   * @property {string} detail   詳しい理由（画面には補足として出す。空文字のこともある）
   */
  /**
   * 位置の種類
   * @typedef {{ type: 'page', page: number }
   *   | { type: 'cell', sheet: string, ref: string, row: number, col: number }
   *   | { type: 'shape', sheet: string, ref: string, row: number, col: number, name: string }
   *   | { type: 'body' }
   *   | { type: 'tableCell', table: number, row: number, col: number, gridCol: number }
   *   | { type: 'textbox', index: number }} Loc
   */
  /**
   * ひとつながりの文字（この中では改行をまたいだ語も見つける。ブロックの境目はまたがない）
   * @typedef {Object} Block
   * @property {string} text
   * @property {'\n'|'\t'} sep  前のブロックとの区切り
   * @property {Loc} loc
   */
  /**
   * 判定に含める・外すの単位（PDF のページ、Excel のシート。Word は文書全体で1つ）
   * @typedef {Object} Unit
   * @property {string} id
   * @property {'page'|'sheet'|'document'} kind
   * @property {string} label
   * @property {boolean} hidden   非表示のシート
   * @property {boolean} hasText
   * @property {Block[]} blocks
   * @property {unknown} [sheet]  Excel のセル・図形の一覧（xlsx.js の XlsxSheet）
   * @property {unknown} [docx]   Word の表の一覧（docx.js の DocxTable[]）
   */
  /**
   * @typedef {Object} LoadedDoc
   * @property {string} name
   * @property {Format} format
   * @property {Unit[]} units
   * @property {string[]} notes  読み込みで気づいたこと（計算結果が保存されていないセルなど）
   */
  /**
   * @typedef {Object} Span
   * @property {number} start
   * @property {number} end
   * @property {string} unitId
   * @property {Loc} loc
   */
  /**
   * @typedef {Object} Composed
   * @property {string} text
   * @property {number[]} breaks  区切りの文字の位置（またいで語を探さない）
   * @property {Span[]} spans     ブロックの位置（start の順）
   * @property {{ unitId: string, start: number, end: number }[]} unitRanges
   */

  /**
   * @typedef {Object} ZipEntry
   * @property {string} name
   * @property {boolean} encrypted
   */
  /** @typedef {(data: Uint8Array) => Promise<Uint8Array>} InflateRaw */
  /**
   * @typedef {Object} ZipApi
   * @property {(b: Uint8Array) => boolean} isZip
   * @property {(b: Uint8Array) => ZipEntry[]} listEntries
   */
  /** @typedef {{ has: (name: string) => boolean, names: string[] }} Package */
  /**
   * @typedef {Object} OoxmlApi
   * @property {(bytes: Uint8Array, inflateRaw?: InflateRaw) => Package} openPackage
   * @property {(pkg: Package, part: string) => Promise<{ id: string, type: string, target: string }[]>} readRels
   */
  /** @typedef {{ units: Unit[], notes: string[] }} ReadResult */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (globalThis));
  const zip = /** @type {ZipApi} */ (g.YegZip);
  const ooxml = /** @type {OoxmlApi} */ (g.YegOoxml);
  const docx = /** @type {{ readDocx: (pkg: Package) => Promise<ReadResult> }} */ (g.YegDocx);
  const xlsx = /** @type {{ readXlsx: (pkg: Package) => Promise<ReadResult> }} */ (g.YegXlsx);

  const MAX_FILE_BYTES = 20 * 1024 * 1024;
  /** 一度に読み込めるファイルの数（付録F-3 AC-05-06） */
  const MAX_FILES = 10;
  const TOO_MANY_FILES = `一度に読み込めるのは${MAX_FILES}ファイルまでです。ファイルを減らして、もう一度選んでください。`;

  const MESSAGES = Object.freeze({
    'too-large': 'ファイルが大きすぎます（20MBまで）。別紙を分けて保存するか、文章を貼り付けてください。',
    'empty-file': 'ファイルの中身が空です。',
    unsupported: 'この形式は読み込めません。Word（.docx）・Excel（.xlsx）・PDFにして読み込むか、文章を貼り付けてください。',
    password: 'パスワード付きのファイルは読み込めません。パスワードを外して保存し直すか、文章を貼り付けてください。',
    broken:
      'ファイルを読み込めませんでした。ファイルが壊れているか、読み込めない作りのファイルです。元のファイルを開いて保存し直すか、文章を貼り付けてください。',
    'no-text': 'このファイルからは文字を読み取れませんでした。',
    'pdf-no-text':
      'このPDFからは文字を読み取れませんでした。WordやExcelから保存したPDFを使うか、Excelのまま読み込んでください。',
    'pdf-unavailable': 'PDFを読み込む部品を準備できませんでした。ページを再読み込みしてから、もう一度お試しください。',
  });

  const XLS_NOTE = 'Excelで .xlsx 形式に保存し直してください。';
  const DOC_NOTE = 'Wordで .docx 形式に保存し直してください。';

  /** 読み込みの途中で、理由がはっきりしている失敗 */
  class FileLoadError extends Error {
    /**
     * @param {LoadErrorCode} code
     * @param {string} [detail]
     */
    constructor(code, detail = '') {
      super(MESSAGES[code]);
      this.name = 'FileLoadError';
      this.code = code;
      this.detail = detail;
    }
  }

  /**
   * @param {LoadErrorCode} code
   * @param {string} [detail]
   * @param {string} [extra]  文の後ろに添える案内
   * @returns {LoadError}
   */
  const loadError = (code, detail = '', extra = '') => ({
    code,
    message: extra ? `${MESSAGES[code]}${extra}` : MESSAGES[code],
    detail,
  });

  /** @param {string} name */
  const extOf = (name) => {
    const m = /\.([A-Za-z0-9]+)$/.exec(name.trim());
    return m ? m[1].toLowerCase() : '';
  };

  /**
   * 読み込む前に大きさを確かめる
   * @param {number} size
   * @returns {LoadError|null}
   */
  const checkSize = (size) => (size > MAX_FILE_BYTES ? loadError('too-large') : null);

  /**
   * @param {Uint8Array} bytes
   * @param {Uint8Array} pattern
   */
  const includesBytes = (bytes, pattern) => {
    const first = pattern[0];
    for (let i = bytes.indexOf(first); i >= 0 && i <= bytes.length - pattern.length; i = bytes.indexOf(first, i + 1)) {
      let ok = true;
      for (let k = 1; k < pattern.length; k += 1) {
        if (bytes[i + k] !== pattern[k]) {
          ok = false;
          break;
        }
      }
      if (ok) return true;
    }
    return false;
  };

  /** @param {string} s */
  const utf16le = (s) => {
    const out = new Uint8Array(s.length * 2);
    for (let i = 0; i < s.length; i += 1) {
      out[i * 2] = s.charCodeAt(i) & 0xff;
      out[i * 2 + 1] = s.charCodeAt(i) >> 8;
    }
    return out;
  };

  /** @param {Uint8Array} b */
  const isPdf = (b) => {
    const head = new TextDecoder('latin1').decode(b.subarray(0, Math.min(b.length, 1024)));
    return head.includes('%PDF-');
  };

  const CFB_SIGNATURE = [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1];
  /** @param {Uint8Array} b */
  const isCfb = (b) => b.length >= 8 && CFB_SIGNATURE.every((v, i) => b[i] === v);

  /**
   * 拡張子に応じて、形式が読めないときに添える案内
   * @param {string} ext
   */
  const legacyNote = (ext) => (ext === 'xls' ? XLS_NOTE : ext === 'doc' ? DOC_NOTE : '');

  /**
   * 古い Office の入れ物（.doc・.xls、パスワード付きの .docx・.xlsx）を見分ける
   * @param {Uint8Array} b
   * @param {string} ext
   * @returns {LoadError}
   */
  const classifyCfb = (b, ext) => {
    if (includesBytes(b, utf16le('EncryptedPackage'))) return loadError('password');
    if (includesBytes(b, utf16le('WordDocument'))) return loadError('unsupported', '古い形式の Word（.doc）です', DOC_NOTE);
    if (includesBytes(b, utf16le('Workbook')) || includesBytes(b, utf16le('Book'))) {
      return loadError('unsupported', '古い形式の Excel（.xls）です', XLS_NOTE);
    }
    return loadError('unsupported', '古い形式の Office のファイルです', legacyNote(ext));
  };

  /**
   * ZIP の中身から、Word か Excel かを見分ける
   * @param {Package} pkg
   * @returns {Promise<Format|'pptx'|null>}
   */
  const officeKind = async (pkg) => {
    const rels = await ooxml.readRels(pkg, '');
    const main = rels.find((r) => r.type.split('/').pop() === 'officeDocument');
    const target = main ? main.target.toLowerCase() : '';
    if (target.startsWith('word/') || (!target && pkg.has('word/document.xml'))) return 'docx';
    if (target.startsWith('xl/') || (!target && pkg.has('xl/workbook.xml'))) return 'xlsx';
    if (target.startsWith('ppt/') || (!target && pkg.has('ppt/presentation.xml'))) return 'pptx';
    return null;
  };

  /**
   * @param {Unit[]} units
   */
  const hasAnyText = (units) => units.some((u) => u.hasText);

  /**
   * PDF のページを単位にする
   * @param {string[]} pages
   * @returns {Unit[]}
   */
  const pdfUnits = (pages) =>
    pages.map((text, i) => {
      const hasText = text.trim() !== '';
      return {
        id: `page-${i + 1}`,
        kind: 'page',
        label: `${i + 1}ページ`,
        hidden: false,
        hasText,
        blocks: hasText ? [{ text, sep: '\n', loc: { type: 'page', page: i + 1 } }] : [],
      };
    });

  /**
   * 読み込みの途中で起きた失敗を、利用者に見せる形にする
   * @param {unknown} e
   * @returns {LoadError}
   */
  const toLoadError = (e) => {
    if (e instanceof FileLoadError) return loadError(e.code, e.detail);
    const name = typeof e === 'object' && e !== null ? String(/** @type {{ name?: unknown }} */ (e).name ?? '') : '';
    const code = typeof e === 'object' && e !== null ? String(/** @type {{ code?: unknown }} */ (e).code ?? '') : '';
    const detail = e instanceof Error ? e.message : String(e);
    if ((name === 'ZipError' && code === 'encrypted') || (name === 'PdfReadError' && code === 'password')) {
      return loadError('password');
    }
    return loadError('broken', detail);
  };

  /**
   * ファイルを読み込む
   * @param {{ name: string, bytes: Uint8Array }} file
   * @param {{ inflateRaw?: InflateRaw, readPdf: (bytes: Uint8Array) => Promise<string[]> }} deps
   * @returns {Promise<{ ok: true, doc: LoadedDoc } | { ok: false, error: LoadError }>}
   */
  const loadFile = async (file, deps) => {
    const { name, bytes } = file;
    const ext = extOf(name);
    const sizeError = checkSize(bytes.length);
    if (sizeError) return { ok: false, error: sizeError };
    if (bytes.length === 0) return { ok: false, error: loadError('empty-file') };
    try {
      if (isPdf(bytes)) {
        const units = pdfUnits(await deps.readPdf(bytes));
        if (!hasAnyText(units)) return { ok: false, error: loadError('pdf-no-text') };
        return { ok: true, doc: { name, format: 'pdf', units, notes: [] } };
      }
      if (isCfb(bytes)) return { ok: false, error: classifyCfb(bytes, ext) };
      if (zip.isZip(bytes)) {
        const entries = zip.listEntries(bytes);
        if (entries.some((e) => e.encrypted)) return { ok: false, error: loadError('password') };
        const pkg = ooxml.openPackage(bytes, deps.inflateRaw);
        const kind = await officeKind(pkg);
        if (kind === 'docx' || kind === 'xlsx') {
          const read = kind === 'docx' ? await docx.readDocx(pkg) : await xlsx.readXlsx(pkg);
          if (!hasAnyText(read.units)) return { ok: false, error: loadError('no-text') };
          return { ok: true, doc: { name, format: kind, units: read.units, notes: read.notes } };
        }
        return {
          ok: false,
          error: loadError('unsupported', kind === 'pptx' ? 'PowerPoint のファイルです' : 'Word・Excel のファイルではありません'),
        };
      }
      // 拡張子は対応している形式なのに、中身がそうなっていない
      if (ext === 'pdf' || ext === 'docx' || ext === 'xlsx') {
        return { ok: false, error: loadError('broken', `拡張子は .${ext} ですが、中身がその形式になっていません`) };
      }
      return { ok: false, error: loadError('unsupported', '', legacyNote(ext)) };
    } catch (e) {
      return { ok: false, error: toLoadError(e) };
    }
  };

  // ------------------------------------------------------------------
  // 判定する文字を組み立てる
  // ------------------------------------------------------------------

  /**
   * 判定に含める単位のブロックを、区切りの文字でつなぐ
   * @param {readonly Unit[]} units
   * @param {ReadonlySet<string>} included
   * @returns {Composed}
   */
  const compose = (units, included) => {
    /** @type {string[]} */
    const parts = [];
    /** @type {number[]} */
    const breaks = [];
    /** @type {Span[]} */
    const spans = [];
    /** @type {{ unitId: string, start: number, end: number }[]} */
    const unitRanges = [];
    let pos = 0;
    for (const u of units) {
      if (!included.has(u.id) || u.blocks.length === 0) continue;
      let unitStart = pos;
      u.blocks.forEach((b, i) => {
        if (parts.length > 0) {
          parts.push(i === 0 ? '\n' : b.sep);
          breaks.push(pos);
          pos += 1;
        }
        if (i === 0) unitStart = pos;
        parts.push(b.text);
        spans.push({ start: pos, end: pos + b.text.length, unitId: u.id, loc: b.loc });
        pos += b.text.length;
      });
      unitRanges.push({ unitId: u.id, start: unitStart, end: pos });
    }
    return { text: parts.join(''), breaks, spans, unitRanges };
  };

  /**
   * 複数のファイルのときに、単位（ページ・シート）を見分ける名前（1つのファイルのときは単位の id のまま）
   * @param {string} fileKey
   * @param {string} unitId
   * @param {boolean} multi
   */
  const unitKeyOf = (fileKey, unitId, multi) => (multi ? `${fileKey}/${unitId}` : unitId);

  /**
   * 複数のファイルの、判定に含める単位をつなぐ（ファイルの境目も区切りにする。付録F-3）。
   * 1つのファイルのときは compose と同じ。parts はファイルごとの組み立て（位置はそのファイルの中）と、つないだ中での位置
   * @param {readonly { key: string, units: readonly Unit[], included: ReadonlySet<string> }[]} files
   * @returns {{ composed: Composed, parts: { key: string, start: number, end: number, composed: Composed }[] }}
   */
  const composeFiles = (files) => {
    const multi = files.length > 1;
    const locals = files.map((f) => ({ key: f.key, composed: compose(f.units, f.included) }));
    if (!multi) {
      const c = locals.length > 0 ? locals[0].composed : compose([], new Set());
      return { composed: c, parts: locals.map((l) => ({ key: l.key, start: 0, end: c.text.length, composed: c })) };
    }
    /** @type {string[]} */
    const texts = [];
    /** @type {number[]} */
    const breaks = [];
    /** @type {Span[]} */
    const spans = [];
    /** @type {{ unitId: string, start: number, end: number }[]} */
    const unitRanges = [];
    /** @type {{ key: string, start: number, end: number, composed: Composed }[]} */
    const parts = [];
    let pos = 0;
    for (const l of locals) {
      const c = l.composed;
      if (c.text === '') {
        parts.push({ key: l.key, start: pos, end: pos, composed: c });
        continue;
      }
      if (texts.length > 0) {
        texts.push('\n');
        breaks.push(pos);
        pos += 1;
      }
      const off = pos;
      texts.push(c.text);
      breaks.push(...c.breaks.map((b) => b + off));
      spans.push(...c.spans.map((s) => ({ ...s, start: s.start + off, end: s.end + off, unitId: unitKeyOf(l.key, s.unitId, true) })));
      unitRanges.push(...c.unitRanges.map((r) => ({ unitId: unitKeyOf(l.key, r.unitId, true), start: r.start + off, end: r.end + off })));
      parts.push({ key: l.key, start: off, end: off + c.text.length, composed: c });
      pos += c.text.length;
    }
    return { composed: { text: texts.join(''), breaks, spans, unitRanges }, parts };
  };

  /**
   * 文字の位置が、どのブロックにあるか
   * @param {Composed} composed
   * @param {number} offset
   * @returns {Span|null}
   */
  const locate = (composed, offset) => {
    const { spans } = composed;
    let lo = 0;
    let hi = spans.length - 1;
    /** @type {Span|null} */
    let found = null;
    while (lo <= hi) {
      const mid = (lo + hi) >> 1;
      if (spans[mid].start <= offset) {
        found = spans[mid];
        lo = mid + 1;
      } else {
        hi = mid - 1;
      }
    }
    return found;
  };

  /**
   * 場所の表示（例: 「2ページ」「01計画／B4」「01計画／テキストボックス（B11付近）」）
   * @param {Loc} loc
   */
  const locationLabel = (loc) => {
    switch (loc.type) {
      case 'page':
        return `${loc.page}ページ`;
      case 'cell':
        return `${loc.sheet}／${loc.ref}`;
      case 'shape':
        return loc.ref ? `${loc.sheet}／テキストボックス（${loc.ref}付近）` : `${loc.sheet}／テキストボックス`;
      case 'tableCell':
        return `表${loc.table}の${loc.row}行目・${loc.col}列目`;
      case 'textbox':
        return `テキストボックス${loc.index}`;
      default:
        return '';
    }
  };

  /**
   * 単位ごとの指摘の件数（位置を持つ指摘だけ）
   * @param {Composed} composed
   * @param {{ range: { start: number, end: number }|null }[]} findings
   * @returns {Map<string, number>}
   */
  const countByUnit = (composed, findings) => {
    /** @type {Map<string, number>} */
    const counts = new Map();
    for (const f of findings) {
      if (!f.range) continue;
      const span = locate(composed, f.range.start);
      if (span) counts.set(span.unitId, (counts.get(span.unitId) ?? 0) + 1);
    }
    return counts;
  };

  /**
   * 最初に判定に含める単位（非表示のシートは外す）
   * @param {readonly Unit[]} units
   * @returns {Set<string>}
   */
  const defaultIncluded = (units) => new Set(units.filter((u) => !u.hidden).map((u) => u.id));

  /**
   * 一覧に出す単位の名前（非表示のシートは「（非表示）」を付ける）
   * @param {Unit} u
   */
  const unitLabel = (u) => (u.hidden ? `${u.label}（非表示）` : u.label);

  /** @type {Readonly<Record<Format, string>>} */
  const FORMAT_LABELS = Object.freeze({ docx: 'Word', xlsx: 'Excel', pdf: 'PDF' });

  root.YegFile = Object.freeze({
    MAX_FILE_BYTES,
    MAX_FILES,
    TOO_MANY_FILES,
    MESSAGES,
    FORMAT_LABELS,
    FileLoadError,
    checkSize,
    loadFile,
    compose,
    composeFiles,
    unitKeyOf,
    locate,
    locationLabel,
    countByUnit,
    defaultIncluded,
    unitLabel,
    extOf,
  });
})(globalThis);
