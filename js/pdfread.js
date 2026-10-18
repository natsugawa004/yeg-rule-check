// @ts-check
/*
 * PDF から、ページごとの文字を取り出す
 *   readPdf(バイト列, { lib }) → ページごとの文字
 *   pageTextFromItems(pdf.js が返す文字の並び) → 1ページ分の文字（純粋関数）
 * PDF の解析には同梱の pdf.js（vendor/pdfjs/）を使う。Worker は作らず、画面と同じ場所で動かす
 * （index.html を file:// で開いても動かすため。ページを1枚読むごとに画面へ制御を返し、進み具合を出す）。
 * pdf.js には通信をさせない: 文字コードの対応表（CMap）は同梱の JS から渡し、それ以外の追加データは使わない。
 */
'use strict';

(function (root) {
  /**
   * @typedef {Object} PdfTextItem
   * @property {string} str
   * @property {boolean} [hasEOL]
   */
  /**
   * @typedef {Object} PdfPage
   * @property {(opts?: Record<string, unknown>) => Promise<{ items: unknown[] }>} getTextContent
   * @property {() => boolean} cleanup
   */
  /**
   * @typedef {Object} PdfDocument
   * @property {number} numPages
   * @property {(n: number) => Promise<PdfPage>} getPage
   */
  /**
   * @typedef {Object} PdfLoadingTask
   * @property {Promise<PdfDocument>} promise
   * @property {() => Promise<void>} destroy
   */
  /**
   * @typedef {Object} PdfLib
   * @property {(src: Record<string, unknown>) => PdfLoadingTask} getDocument
   */

  class PdfReadError extends Error {
    /**
     * @param {'password'|'broken'} code
     * @param {string} message
     */
    constructor(code, message) {
      super(message);
      this.name = 'PdfReadError';
      this.code = code;
    }
  }

  /**
   * @param {unknown} v
   * @returns {v is PdfTextItem}
   */
  const isTextItem = (v) => typeof v === 'object' && v !== null && typeof (/** @type {{ str?: unknown }} */ (v).str) === 'string';

  /** 日本語の文字（かな・漢字・全角の記号。全角の空白は含めない） */
  const CJK = '[\\u3001-\\u303f\\u3040-\\u30ff\\u3400-\\u4dbf\\u4e00-\\u9fff\\uf900-\\ufaff\\uff01-\\uff60\\uffe0-\\uffe6]';
  /**
   * Word は日本語と英数字の間を少し空けて配置する（「9 月」「HP に」）。pdf.js はその隙間を半角の空白として返すため、
   * 英数字と日本語の間にある半角の空白1つは取り除く（取り除かないと「5000 円」のように、金額の表記の誤りを見逃す）
   */
  const AUTO_SPACE = new RegExp(`(?<=[A-Za-z0-9%])\\u0020(?=${CJK})|(?<=${CJK})\\u0020(?=[A-Za-z0-9])`, 'gu');

  /**
   * pdf.js の文字の並びを、行ごとに改行した文字にする
   * （行末の空白と、前後の空行を除き、3行以上続く空行は1行にする）
   * @param {unknown[]} items
   * @returns {string}
   */
  const pageTextFromItems = (items) => {
    const raw = items
      .filter(isTextItem)
      .map((it) => `${it.str}${it.hasEOL ? '\n' : ''}`)
      .join('')
      .replace(/\r\n?/g, '\n')
      .replace(/\u0000/g, '');
    return raw
      .split('\n')
      .map((line) => line.replace(/[ \t]+$/u, '').replace(AUTO_SPACE, ''))
      .join('\n')
      .replace(/\n{3,}/g, '\n\n')
      .replace(/^\n+|\n+$/g, '');
  };

  /** 画面の描き直し（進み具合の表示）ができるよう、いったん制御を返す */
  const yieldToBrowser = () => new Promise((resolve) => setTimeout(resolve, 0));

  /**
   * base64 の文字をバイト列にする
   * @param {string} b64
   */
  const fromBase64 = (b64) => {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
    return out;
  };

  /**
   * pdf.js が CMap などの追加データを求めたときに渡す仕組み（通信はしない）。
   * pdf.js はこのクラスから作った物の「fetch」という名前の関数を呼ぶので、その名前で、同梱のデータを返す関数を置く
   * @param {() => Readonly<Record<string, string>>|null} getCmaps  CMap の名前 → base64
   */
  const makeBinaryDataFactory = (getCmaps) => {
    /**
     * @param {{ kind: string, filename: string }} req
     * @returns {Promise<Uint8Array>}
     */
    const readBundled = async ({ kind, filename }) => {
      if (kind !== 'cMapUrl') throw new Error(`${kind} のデータは同梱していません`);
      const name = filename.replace(/\.bcmap$/, '');
      const table = getCmaps();
      const b64 = table && Object.prototype.hasOwnProperty.call(table, name) ? table[name] : null;
      if (b64 === null) throw new Error(`文字コードの対応表 ${name} を同梱していません`);
      return fromBase64(b64);
    };
    class BinaryDataFactory {}
    Object.defineProperty(BinaryDataFactory.prototype, 'fetch', { value: readBundled });
    return BinaryDataFactory;
  };

  /**
   * @param {unknown} e
   * @returns {PdfReadError}
   */
  const toReadError = (e) => {
    const name = typeof e === 'object' && e !== null ? String(/** @type {{ name?: unknown }} */ (e).name ?? '') : '';
    const message = e instanceof Error ? e.message : String(e);
    if (name === 'PasswordException') return new PdfReadError('password', 'パスワードが必要な PDF です');
    return new PdfReadError('broken', message || 'PDF を読み取れませんでした');
  };

  /**
   * PDF のページごとの文字を読む
   * @param {Uint8Array} bytes
   * @param {{ lib: PdfLib, getCmaps?: () => Readonly<Record<string, string>>|null, onProgress?: (done: number, total: number) => void, yieldFn?: () => Promise<void> }} deps
   * @returns {Promise<string[]>}
   */
  const readPdf = async (bytes, deps) => {
    const { lib, onProgress, yieldFn = yieldToBrowser } = deps;
    const task = lib.getDocument({
      // pdf.js は渡したバイト列を手放す（使えなくする）ことがあるので、写しを渡す
      data: bytes.slice(),
      isEvalSupported: false,
      disableFontFace: true,
      useSystemFonts: false,
      useWorkerFetch: false,
      useWasm: false,
      enableXfa: false,
      isOffscreenCanvasSupported: false,
      isImageDecoderSupported: false,
      stopAtErrors: false,
      verbosity: 0,
      BinaryDataFactory: makeBinaryDataFactory(deps.getCmaps ?? (() => null)),
    });
    try {
      /** @type {PdfDocument} */
      let doc;
      try {
        doc = await task.promise;
      } catch (e) {
        throw toReadError(e);
      }
      /** @type {string[]} */
      const pages = [];
      for (let n = 1; n <= doc.numPages; n += 1) {
        try {
          const page = await doc.getPage(n);
          const content = await page.getTextContent({ includeMarkedContent: false });
          pages.push(pageTextFromItems(content.items));
          page.cleanup();
        } catch (e) {
          throw new PdfReadError('broken', `${n}ページ目を読み取れませんでした（${e instanceof Error ? e.message : String(e)}）`);
        }
        if (onProgress) onProgress(n, doc.numPages);
        await yieldFn();
      }
      return pages;
    } finally {
      try {
        await task.destroy();
      } catch (e) {
        // 読み取りの結果には影響しない後片付けの失敗なので、記録だけ残す
        console.warn('PDF の読み取りの後片付けに失敗しました', e);
      }
    }
  };

  root.YegPdf = Object.freeze({ PdfReadError, readPdf, pageTextFromItems, makeBinaryDataFactory, fromBase64 });
})(globalThis);
