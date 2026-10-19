// @ts-check
/*
 * 書類の種類に応じたチェック（F-03）。画面に依存しない純粋な関数だけを置く
 *   analyzeDocument(入力, ルール) → 見つかった見出し・欄の範囲・書類の種類・予算か決算か
 *   requiredReport(分析, 選んだ種類, ルール) → 必須項目の表（あり／記入あり／空欄／見つかりません）
 *   checkpointGroups(選んだ種類, ルール) → 確認ポイント
 * 見出しの認め方
 *   ・貼り付け・PDF・Wordの段落: 行の先頭の見出し語（番号・記号の後でもよい）。見出し語だけの行か、後ろに空白をはさんで記入が続く行。
 *     見出しから次の見出しまでをその欄とする。空白・タブで区切った語が並ぶ行（表の見出し行）の語も「見出しあり」とする
 *   ・Excelのセル・Wordの表のセル: 見出し語だけのセル、見出し語＋列名・かっこ書き、見出し語＋「：」や空白に続く記入。
 *     見出しのセルから右（同じ行）と下（同じ列）の、次の見出しの手前までを記入欄とし、空欄かどうかを判定する
 */
'use strict';

(function (root) {
  /**
   * @typedef {Object} DocType
   * @property {string} id
   * @property {string} name
   * @property {string[]} detectWords
   * @property {string[]} required
   * @property {string[]} settlementRequired
   * @property {string[]} blankCheck
   * @property {{ id: string, text: string, checkedBy?: string[] }[]} checkpoints
   * @property {boolean} [useCommonCheckpoints]  共通の確認ポイントも出すか（既定 true）
   */
  /**
   * engine.js の CompiledDocuments
   * @typedef {Object} DocData
   * @property {number} minDetectWords
   * @property {DocType[]} types
   * @property {{ id: string, text: string, checkedBy?: string[] }[]} commonCheckpoints
   * @property {{ key: string, name: string }[]} headingKeys
   * @property {Readonly<Record<string, string[]>>} substitutes
   * @property {{ name: string, regexes: RegExp[] }[]} phrases
   * @property {{ budget: RegExp, settlement: RegExp }} budgetKind
   * @property {string[]} columnSuffixes
   * @property {string[]} otherLabels
   * @property {Readonly<Record<string, string[]>>} fieldLabels
   * @property {RegExp[]} sectionBreaks
   */
  /** @typedef {{ documents: DocData|null }} RulesetLike */
  /**
   * @typedef {Object} EngineApi
   * @property {(s: string) => { text: string, map: number[] }} normalizeLabel
   * @property {(s: string) => string} labelKey
   * @property {(s: string) => string} normalizeNewlines
   */

  /** @typedef {'paste'|'docx'|'xlsx'|'pdf'} Format */
  /** @typedef {{ type: string, [key: string]: unknown }} Loc */
  /** @typedef {{ start: number, end: number, unitId: string, loc: Loc }} Span */
  /** @typedef {{ text: string, breaks: number[], spans: Span[], unitRanges: { unitId: string, start: number, end: number }[] }} Composed */
  /**
   * fileload.js の Unit（使う項目だけ）
   * @typedef {Object} Unit
   * @property {string} id
   * @property {unknown} [sheet]
   * @property {unknown} [docx]
   */
  /** @typedef {{ row: number, col: number, text: string, formula: boolean }} XCell */
  /** @typedef {{ r1: number, c1: number, r2: number, c2: number }} XMerge */
  /** @typedef {{ name: string, ref: string, row: number, col: number, areaRow?: number, text: string }} XShape */
  /** @typedef {{ cells: XCell[], merges: XMerge[], shapes: XShape[] }} XSheet */
  /** @typedef {{ row: number, col: number, gridCol: number, gridSpan: number, vMerge: 'restart'|'continue'|null, text: string }} DCell */
  /** @typedef {{ index: number, rows: DCell[][] }} DTable */

  /**
   * @typedef {Object} DocInput
   * @property {Format} format
   * @property {string} text          判定する文章（ファイルなら compose した文字）
   * @property {Composed|null} composed
   * @property {readonly Unit[]|null} units
   */

  /**
   * 見つかった見出し1つ
   * @typedef {Object} Occurrence
   * @property {string} name
   * @property {'line'|'cell'|'token'} mode   段落の行／表のセル／表の見出し行の語
   * @property {boolean} filled  表のセルのとき、記入欄（または見出しのセルの中）に記入があるか
   */

  /** @typedef {{ start: number, end: number, name: string }} Section */
  /** @typedef {'予算'|'決算'} BudgetKind */
  /** 'table' は、予算・決算の印も文字もなく、Excel の収支表の決算額の列で決めたとき（S4。付録F-2） */
  /** @typedef {'mark'|'text'|'ambiguous'|'default'|'table'} BudgetKindSource */

  /**
   * @typedef {Object} DocAnalysis
   * @property {boolean} available   ルール定義に書類の種類のデータがある
   * @property {Format} format
   * @property {Format[]} [formats]  複数のファイルを合わせたときの、それぞれの形式
   * @property {Occurrence[]} occurrences
   * @property {Section[]} sections
   * @property {string[]} detected    判別できた書類の種類（id。ルール定義の順）
   * @property {Record<string, string[]>} detectWordsFound  種類ごとの、見つかった判別語
   * @property {string[]} phrasesFound  文章のどこかにあった必須項目（予算（案）か決算（案）など）
   * @property {BudgetKind} budgetKind
   * @property {BudgetKindSource} budgetKindSource
   */

  /**
   * 利用者が選んだ（または自動で判別した）書類の種類
   * @typedef {Object} Selection
   * @property {string[]} types
   * @property {BudgetKind} budgetKind
   * @property {boolean} auto
   */

  /** @typedef {'filled'|'blank'|'present'|'missing'} ItemStatus */
  /**
   * @typedef {Object} RequiredRow
   * @property {string} name
   * @property {ItemStatus} status
   * @property {string} label  画面に出す状態（記入あり／空欄／あり／見つかりません）
   * @property {string|null} via  代わりに確認した見出し（例: 対象者 → 参加人数（目標））
   */
  /**
   * @typedef {Object} RequiredTable
   * @property {string} typeId
   * @property {string} title
   * @property {RequiredRow[]} rows
   */
  /**
   * @typedef {Object} RequiredReport
   * @property {RequiredTable[]} tables
   * @property {number} total       必須項目の数（種類をまたいで同じ見出しは1つと数える）
   * @property {number} satisfied   あり・記入ありの数
   * @property {number} shortage    見つかりません・空欄の数
   * @property {string[]} notes     表の下に出す注意
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (root));
  const engine = /** @type {EngineApi} */ (g.YegEngine);

  /** @type {Readonly<Record<ItemStatus, string>>} */
  const STATUS_LABELS = Object.freeze({ filled: '記入あり', blank: '空欄', present: 'あり', missing: '見つかりません' });

  const MESSAGES = Object.freeze({
    undetermined: '書類の種類を判別できませんでした。文章の一部だけを入れた場合は、必須項目のチェックは行いません',
    noneSelected: '書類の種類が選ばれていないため、必須項目のチェックは行いません',
    unavailable: 'ルール定義に書類の種類のデータがないため、書類の種類に応じたチェックは行いません',
    pasteNoBlank: '貼り付け・PDFでは、欄が記入されているかは確認していません（Excelで読み込むと確認できます）',
    docxNoBlank: 'Wordでは、表の中の欄だけ記入されているかを確認しています（表の外の欄は「あり」と表示します）',
    partial: '書類の一部だけを入れた場合は正しく判定できません',
    xlsxBlank: 'Excelの配置によっては、記入があっても空欄と表示される場合があります',
    docxBlank: 'Wordの表の配置によっては、記入があっても空欄と表示される場合があります',
    budgetAmbiguous: '予算か決算かを判別できなかったため、予算として確認しています。違う場合は選び直してください',
    budgetByTable: '予算・決算の印が見つからないため、収支表の決算額の列に金額があるかどうかで判断しました。違う場合は選び直してください',
  });

  /** 見出しの前に付いていてもよい番号・記号 */
  const PREFIX = /^[\s　]*(?:(?:[①-⑳]|[(（][0-9０-９一二三四五六七八九十]{1,2}[)）]|[0-9０-９]{1,2}[.．、)）]|[■□◆◇●・※▼▶★☆])[\s　]*)?/u;
  /** 見出し語のすぐ後ろにあってもよい閉じかっこ（【事業評価】など） */
  const CLOSE = /^[】\]］〕〗]+/u;

  /** @param {string} s */
  const prefixLength = (s) => {
    const m = PREFIX.exec(s);
    return m ? m[0].length : 0;
  };

  /**
   * そろえた文字の n 文字目（n>0）までに当たる、元の文字列での終わりの位置
   * @param {string} original
   * @param {number[]} map
   * @param {number} n
   */
  const endInOriginal = (original, map, n) => {
    const at = map[n - 1];
    const code = original.codePointAt(at) ?? 0;
    return at + (code > 0xffff ? 2 : 1);
  };

  /**
   * 見出し語の後ろが「何もない」「空白・：の後に記入」のどちらか。
   * 段落の行では、数字・英字・「 で始まる記入がすぐ続くものも見出しとみなす
   * （PDF から読んだ文字は「事業名称9月例会」のように、日本語と数字の間の空白がなくなるため）
   * @param {string} rest
   * @param {boolean} [line]
   * @returns {'only'|'content'|null}
   */
  const afterHeading = (rest, line = false) => {
    const r = rest.replace(CLOSE, '');
    if (r.trim() === '') return 'only';
    if (/^[\s　:：]/.test(r)) return 'content';
    return line && /^[0-9A-Za-z０-９Ａ-Ｚａ-ｚ「『]/.test(r) ? 'content' : null;
  };

  // ------------------------------------------------------------------
  // 段落の行の見出し
  // ------------------------------------------------------------------

  /**
   * 行が見出しかどうか。見出し語が次の行にまたがっている（「事業」で改行して「評価」）ときは spansNext
   * @param {string} line
   * @param {string|null} next
   * @param {DocData} docs
   * @returns {{ name: string, spansNext: boolean }|null}
   */
  const parseLineHeading = (line, next, docs) => {
    const start = prefixLength(line);
    const body = line.slice(start);
    const norm = engine.normalizeLabel(body);
    if (norm.text === '') return null;
    /**
     * 見出し語のすぐ後ろに列名（「【支 出】項 目 摘 要」の「項目」）があり、その後ろが空白か行の終わり
     * @param {string} rest
     */
    const columnSuffixFollows = (rest) => {
      const r = rest.replace(CLOSE, '');
      const n = engine.normalizeLabel(r);
      return docs.columnSuffixes.some(
        (sfx) => n.text.startsWith(sfx) && afterHeading(r.slice(endInOriginal(r, n.map, sfx.length))) !== null,
      );
    };
    for (const { key, name } of docs.headingKeys) {
      if (!norm.text.startsWith(key)) continue;
      const rest = body.slice(endInOriginal(body, norm.map, key.length));
      if (afterHeading(rest, true) || columnSuffixFollows(rest)) return { name, spansNext: false };
    }
    if (next === null || norm.text.length < 2) return null;
    // 見出し語の途中で改行している（この行が見出し語の始めの部分だけ）
    const n2 = engine.normalizeLabel(next);
    const combined = norm.text + n2.text;
    for (const { key, name } of docs.headingKeys) {
      if (key.length <= norm.text.length || !key.startsWith(norm.text) || !combined.startsWith(key)) continue;
      const k2 = key.length - norm.text.length;
      if (afterHeading(next.slice(endInOriginal(next, n2.map, k2)), true)) return { name, spansNext: true };
    }
    return null;
  };

  /**
   * 空白・タブで区切った語が並ぶ行（表の見出し行）の中の見出し語
   * @param {string} line
   * @param {ReadonlyMap<string, string>} byKey
   * @returns {string[]}
   */
  const tokenHeadings = (line, byKey) => {
    const tokens = line.split(/[\s　]+/u).filter((t) => t !== '');
    if (tokens.length < 2) return [];
    return tokens.flatMap((t) => {
      const name = byKey.get(engine.labelKey(t.slice(prefixLength(t))));
      return name ? [name] : [];
    });
  };

  /**
   * @typedef {Object} LineScan
   * @property {{ pos: number, name: string|null }[]} events  見出しの行（name）と、欄の区切りの行（name が null）の位置
   * @property {Occurrence[]} occurrences
   */

  /**
   * 段落の文章（範囲ごと）から、行の見出しを探す
   * @param {string} text
   * @param {readonly { start: number, end: number }[]} regions
   * @param {DocData} docs
   * @param {ReadonlyMap<string, string>} byKey
   * @returns {LineScan}
   */
  const scanLines = (text, regions, docs, byKey) => {
    /** @type {LineScan['events']} */
    const events = [];
    /** @type {Occurrence[]} */
    const occurrences = [];
    for (const region of regions) {
      /** @type {{ start: number, text: string }[]} */
      const lines = [];
      let pos = region.start;
      for (const line of text.slice(region.start, region.end).split('\n')) {
        lines.push({ start: pos, text: line });
        pos += line.length + 1;
      }
      for (let i = 0; i < lines.length; i += 1) {
        const line = lines[i];
        for (const name of tokenHeadings(line.text, byKey)) occurrences.push({ name, mode: 'token', filled: false });
        if (docs.sectionBreaks.some((re) => line.text.search(re) >= 0)) {
          events.push({ pos: line.start, name: null });
          continue;
        }
        const h = parseLineHeading(line.text, i + 1 < lines.length ? lines[i + 1].text : null, docs);
        if (!h) continue;
        events.push({ pos: line.start, name: h.name });
        occurrences.push({ name: h.name, mode: 'line', filled: false });
        if (h.spansNext) i += 1;
      }
    }
    return { events, occurrences };
  };

  // ------------------------------------------------------------------
  // 表（Excel のシート・Word の表）の見出しと記入欄
  // ------------------------------------------------------------------

  /**
   * @typedef {Object} GridItem   表のセル（結合されたセルは範囲全体）
   * @property {number} r1
   * @property {number} c1
   * @property {number} r2
   * @property {number} c2
   * @property {string} text
   * @property {boolean} formula
   * @property {string} key  判定する文字のブロックとの対応（ブロックがないセルも持つ）
   */
  /** @typedef {{ row: number, col: number, text: string, key: string }} GridShape */
  /** @typedef {{ kind: 'heading', name: string, inline: boolean } | { kind: 'label' }} CellClass */
  /** @typedef {{ r1: number, c1: number, r2: number, c2: number }} Rect */

  /**
   * セルが見出し（またはほかの欄名）かどうか
   * @param {string} text
   * @param {DocData} docs
   * @returns {CellClass|null}
   */
  const parseCellHeading = (text, docs) => {
    const start = prefixLength(text);
    const body = text.slice(start);
    const norm = engine.normalizeLabel(body);
    if (norm.text === '') return null;
    /** @param {string} key */
    const restOf = (key) => body.slice(endInOriginal(body, norm.map, key.length)).replace(CLOSE, '');
    for (const { key, name } of docs.headingKeys) {
      if (!norm.text.startsWith(key)) continue;
      const suffix = norm.text.slice(key.length);
      if (suffix === '' || docs.columnSuffixes.includes(suffix) || /^\([^()]*\)$/.test(suffix)) {
        return { kind: 'heading', name, inline: false };
      }
      // 「委員会名：みらい創造委員会」「委員会名　みらい創造委員会」（同じ行で、：か空白の後に記入）
      const rest = restOf(key);
      if (/^[ \t　]*[:：]/.test(rest) || /^[ \t　]+\S/.test(rest)) {
        return { kind: 'heading', name, inline: rest.replace(/^[ \t　]*[:：]?/, '').trim() !== '' };
      }
    }
    for (const key of docs.otherLabels) {
      if (!norm.text.startsWith(key)) continue;
      const suffix = norm.text.slice(key.length);
      if (suffix === '' || /^\([^()]*\)$/.test(suffix)) return { kind: 'label' };
      if (/^[ \t　]*(?:\n|[※(（])/.test(restOf(key))) return { kind: 'label' };
    }
    return null;
  };

  /** @param {Rect} r @param {number} row @param {number} col */
  const inRect = (r, row, col) => row >= r.r1 && row <= r.r2 && col >= r.c1 && col <= r.c2;

  /**
   * @typedef {Object} GridResult
   * @property {Occurrence[]} occurrences
   * @property {{ key: string, name: string }[]} owned   欄に属するブロック（見出しのセル自体も含む）
   * @property {string[]} headingKeys  見出しのセルのブロック
   */

  /**
   * 表の見出しのセルを探し、記入欄（右と下）に記入があるかを判定する
   * @param {readonly GridItem[]} items
   * @param {readonly GridShape[]} shapes
   * @param {{ maxRow: number, maxCol: number }} size
   * @param {DocData} docs
   * @returns {GridResult}
   */
  const analyzeGrid = (items, shapes, size, docs) => {
    const classified = items.map((it) => ({ it, cls: it.text.trim() === '' ? null : parseCellHeading(it.text, docs) }));
    const stops = classified.filter((c) => c.cls !== null).map((c) => c.it);
    const plainItems = classified.filter((c) => c.cls === null).map((c) => c.it);
    const headings = classified.flatMap((c) =>
      c.cls && c.cls.kind === 'heading' ? [{ it: c.it, name: c.cls.name, inline: c.cls.inline }] : [],
    );

    const ranged = headings.map((h) => {
      const others = stops.filter((s) => s !== h.it);
      const rightStop = Math.min(
        size.maxCol + 1,
        ...others.filter((s) => s.c1 > h.it.c2 && s.r1 <= h.it.r2 && s.r2 >= h.it.r1).map((s) => s.c1),
      );
      const downStop = Math.min(
        size.maxRow + 1,
        ...others.filter((s) => s.r1 > h.it.r2 && s.c1 <= h.it.c2 && s.c2 >= h.it.c1).map((s) => s.r1),
      );
      /** @type {Rect} */
      const right = { r1: h.it.r1, r2: h.it.r2, c1: h.it.c2 + 1, c2: rightStop - 1 };
      /** @type {Rect} */
      const down = { r1: h.it.r2 + 1, r2: downStop - 1, c1: h.it.c1, c2: h.it.c2 };
      return { ...h, right, down };
    });

    const fieldLabelsOf = (/** @type {string} */ name) => docs.fieldLabels[name] ?? [];
    const occurrences = ranged.map((h) => {
      const labels = fieldLabelsOf(h.name);
      const inRange = (/** @type {number} */ r, /** @type {number} */ c) => inRect(h.right, r, c) || inRect(h.down, r, c);
      const cellFilled = plainItems.some(
        (it) => inRange(it.r1, it.c1) && !it.formula && it.text.trim() !== '' && !labels.includes(engine.labelKey(it.text)),
      );
      const shapeFilled = shapes.some((s) => inRange(s.row, s.col) && s.text.trim() !== '');
      return { name: h.name, mode: /** @type {'cell'} */ ('cell'), filled: h.inline || cellFilled || shapeFilled };
    });

    /**
     * ブロックが属する欄（同じ行の左の見出しを先に、なければ同じ列の上の見出し。どちらも近いものを選ぶ）
     * @param {number} row
     * @param {number} col
     */
    const ownerOf = (row, col) => {
      const byRight = ranged.filter((h) => inRect(h.right, row, col)).sort((a, b) => b.it.c2 - a.it.c2)[0];
      if (byRight) return byRight.name;
      const byDown = ranged.filter((h) => inRect(h.down, row, col)).sort((a, b) => b.it.r2 - a.it.r2)[0];
      return byDown ? byDown.name : null;
    };
    const owned = [
      ...ranged.map((h) => ({ key: h.it.key, name: h.name })),
      ...plainItems.flatMap((it) => {
        const name = it.text.trim() === '' ? null : ownerOf(it.r1, it.c1);
        return name ? [{ key: it.key, name }] : [];
      }),
      ...shapes.flatMap((s) => {
        const name = s.text.trim() === '' ? null : ownerOf(s.row, s.col);
        return name ? [{ key: s.key, name }] : [];
      }),
    ];
    return { occurrences, owned, headingKeys: ranged.map((h) => h.it.key) };
  };

  /**
   * Excel のシートを、表の形にする
   * @param {string} unitId
   * @param {XSheet} sheet
   */
  const sheetGrid = (unitId, sheet) => {
    /** @type {Map<string, XMerge>} */
    const mergeAt = new Map(sheet.merges.map((m) => [`${m.r1}|${m.c1}`, m]));
    /** @type {GridItem[]} */
    const items = sheet.cells.map((c) => {
      const m = mergeAt.get(`${c.row}|${c.col}`);
      return {
        r1: c.row,
        c1: c.col,
        r2: m ? m.r2 : c.row,
        c2: m ? m.c2 : c.col,
        text: c.text,
        formula: c.formula,
        key: `${unitId}|cell|${c.row}|${c.col}`,
      };
    });
    /** @type {GridShape[]} */
    const shapes = sheet.shapes
      .filter((s) => s.row > 0)
      .map((s) => ({ row: s.areaRow && s.areaRow > 0 ? s.areaRow : s.row, col: s.col, text: s.text, key: `${unitId}|shape|${s.name}|${s.ref}` }));
    const withText = items.filter((it) => it.text.trim() !== '');
    const shapesWithText = shapes.filter((s) => s.text.trim() !== '');
    const size = {
      maxRow: Math.max(0, ...withText.map((it) => it.r2), ...shapesWithText.map((s) => s.row)),
      maxCol: Math.max(0, ...withText.map((it) => it.c2), ...shapesWithText.map((s) => s.col)),
    };
    return { items, shapes, size };
  };

  /**
   * Word の表を、表の形にする（縦に結合されたセルは1つの範囲にする）
   * @param {string} unitId
   * @param {DTable} table
   */
  const tableGrid = (unitId, table) => {
    /** @type {GridItem[]} */
    const items = [];
    /** @type {Map<string, GridItem>} */
    const at = new Map();
    for (const row of table.rows) {
      for (const cell of row) {
        const above = cell.vMerge === 'continue' ? at.get(`${cell.row - 1}|${cell.gridCol}`) : undefined;
        if (above) {
          above.r2 = cell.row;
          at.set(`${cell.row}|${cell.gridCol}`, above);
          continue;
        }
        const item = {
          r1: cell.row,
          c1: cell.gridCol,
          r2: cell.row,
          c2: cell.gridCol + cell.gridSpan - 1,
          text: cell.text,
          formula: false,
          key: `${unitId}|t${table.index}|${cell.row}|${cell.col}`,
        };
        items.push(item);
        at.set(`${cell.row}|${cell.gridCol}`, item);
      }
    }
    return { items, size: { maxRow: table.rows.length, maxCol: Math.max(0, ...items.map((it) => it.c2)) } };
  };

  /**
   * 判定する文字のブロックの、表・図形との対応の目印
   * @param {Span} span
   */
  const spanKey = (span) => {
    const loc = span.loc;
    switch (loc.type) {
      case 'cell':
        return `${span.unitId}|cell|${String(loc.row)}|${String(loc.col)}`;
      case 'shape':
        return `${span.unitId}|shape|${String(loc.name)}|${String(loc.ref)}`;
      case 'tableCell':
        return `${span.unitId}|t${String(loc.table)}|${String(loc.row)}|${String(loc.col)}`;
      default:
        return '';
    }
  };

  // ------------------------------------------------------------------
  // 予算か決算か
  // ------------------------------------------------------------------

  /**
   * 正規表現が文字列全体に当てはまるか
   * @param {RegExp} re
   * @param {string} s
   */
  const fullMatch = (re, s) => {
    const m = new RegExp(re.source, re.flags.replace('g', '')).exec(s);
    return m !== null && m.index === 0 && m[0].length === s.length;
  };

  /**
   * Excel のひな形の「予算（案）／予算修正（案）／決算（案）／決算修正（案）」の選択肢のうち、下のセルに印（〇など）があるもの
   * @param {readonly GridItem[]} items
   * @param {DocData} docs
   * @returns {BudgetKind|null}
   */
  const markedBudgetKind = (items, docs) => {
    const options = items.flatMap((it) => {
      const key = engine.labelKey(it.text);
      if (fullMatch(docs.budgetKind.settlement, key)) return [{ it, kind: /** @type {BudgetKind} */ ('決算') }];
      if (fullMatch(docs.budgetKind.budget, key)) return [{ it, kind: /** @type {BudgetKind} */ ('予算') }];
      return [];
    });
    if (!options.some((o) => o.kind === '予算') || !options.some((o) => o.kind === '決算')) return null;
    const marked = new Set(
      options
        .filter((o) =>
          items.some((x) => x.r1 === o.it.r2 + 1 && x.c1 >= o.it.c1 && x.c1 <= o.it.c2 && x.text.trim() !== ''),
        )
        .map((o) => o.kind),
    );
    return marked.size === 1 ? [...marked][0] : null;
  };

  // ------------------------------------------------------------------
  // 分析
  // ------------------------------------------------------------------

  /** @param {unknown} v @returns {v is XSheet} */
  const isSheet = (v) =>
    typeof v === 'object' && v !== null && Array.isArray(/** @type {XSheet} */ (v).cells) && Array.isArray(/** @type {XSheet} */ (v).merges);
  /** @param {unknown} v @returns {v is { tables: DTable[] }} */
  const isDocxInfo = (v) => typeof v === 'object' && v !== null && Array.isArray(/** @type {{ tables: unknown }} */ (v).tables);

  /**
   * @param {Format} format
   * @returns {DocAnalysis}
   */
  const emptyAnalysis = (format) => ({
    available: false,
    format,
    occurrences: [],
    sections: [],
    detected: [],
    detectWordsFound: {},
    phrasesFound: [],
    budgetKind: '予算',
    budgetKindSource: 'default',
  });

  /**
   * 書類を分析する（見出し・欄の範囲・書類の種類・予算か決算か）
   * @param {DocInput} input
   * @param {RulesetLike} ruleset
   * @returns {DocAnalysis}
   */
  const analyzeDocument = (input, ruleset) => {
    const docs = ruleset.documents;
    if (!docs) return emptyAnalysis(input.format);
    const text = engine.normalizeNewlines(input.text);
    /** @type {Map<string, string>} */
    const byKey = new Map(docs.headingKeys.map((h) => [h.key, h.name]));
    const composed = input.composed;
    const spans = composed ? composed.spans : [];
    const included = new Set(composed ? composed.unitRanges.map((r) => r.unitId) : []);
    /** @type {Map<string, Span>} */
    const spanByKey = new Map();
    for (const s of spans) {
      const k = spanKey(s);
      if (k !== '' && !spanByKey.has(k)) spanByKey.set(k, s);
    }

    // 段落の行として見る範囲
    const lineRegions =
      input.format === 'paste' || !composed
        ? [{ start: 0, end: text.length }]
        : spans.filter((s) => ['page', 'body', 'textbox'].includes(s.loc.type)).map((s) => ({ start: s.start, end: s.end }));
    const lines = input.format === 'xlsx' ? { events: [], occurrences: [] } : scanLines(text, lineRegions, docs, byKey);

    // 表（Excel のシート・Word の表）
    /** @type {GridResult[]} */
    const grids = [];
    /** @type {BudgetKind|null} */
    let marked = null;
    for (const unit of input.units ?? []) {
      if (!included.has(unit.id)) continue;
      if (input.format === 'xlsx' && isSheet(unit.sheet)) {
        const grid = sheetGrid(unit.id, unit.sheet);
        grids.push(analyzeGrid(grid.items, grid.shapes, grid.size, docs));
        marked = marked ?? markedBudgetKind(grid.items, docs);
      } else if (input.format === 'docx' && isDocxInfo(unit.docx)) {
        for (const table of unit.docx.tables) {
          const grid = tableGrid(unit.id, table);
          grids.push(analyzeGrid(grid.items, [], grid.size, docs));
        }
      }
    }

    // 欄の範囲: 表は欄に属するブロックそのもの、段落は見出しの行から次の見出し（表の見出しのセルを含む）・区切りの行まで
    /** @type {Section[]} */
    const cellSections = grids.flatMap((gr) =>
      gr.owned.flatMap((o) => {
        const s = spanByKey.get(o.key);
        return s ? [{ start: s.start, end: s.end, name: o.name }] : [];
      }),
    );
    const cellHeadingStarts = grids.flatMap((gr) =>
      gr.headingKeys.flatMap((k) => {
        const s = spanByKey.get(k);
        return s ? [s.start] : [];
      }),
    );
    const events = [
      ...lines.events,
      ...cellHeadingStarts.map((pos) => ({ pos, name: /** @type {string|null} */ (null) })),
    ].sort((a, b) => a.pos - b.pos);
    /** @type {Section[]} */
    const lineSections = events.flatMap((e, i) => {
      if (e.name === null) return [];
      const next = events.slice(i + 1).find((x) => x.pos > e.pos);
      return [{ start: e.pos, end: next ? next.pos : text.length, name: e.name }];
    });

    const occurrences = [...lines.occurrences, ...grids.flatMap((gr) => gr.occurrences)];
    const found = new Set(occurrences.map((o) => o.name));
    const canonical = (/** @type {string} */ w) => byKey.get(engine.labelKey(w)) ?? w;
    /** @type {Record<string, string[]>} */
    const detectWordsFound = {};
    for (const t of docs.types) detectWordsFound[t.id] = [...new Set(t.detectWords.map(canonical))].filter((w) => found.has(w));
    const detected = docs.types.filter((t) => detectWordsFound[t.id].length >= docs.minDetectWords).map((t) => t.id);

    const flat = engine.labelKey(text);
    const phrasesFound = docs.phrases.filter((p) => p.regexes.some((re) => flat.search(re) >= 0)).map((p) => p.name);
    const hasBudget = flat.search(docs.budgetKind.budget) >= 0;
    const hasSettlement = flat.search(docs.budgetKind.settlement) >= 0;
    /** @type {{ kind: BudgetKind, source: BudgetKindSource }} */
    const kind = marked
      ? { kind: marked, source: 'mark' }
      : hasSettlement && !hasBudget
        ? { kind: '決算', source: 'text' }
        : hasBudget && !hasSettlement
          ? { kind: '予算', source: 'text' }
          : { kind: '予算', source: hasBudget && hasSettlement ? 'ambiguous' : 'default' };

    return {
      available: true,
      format: input.format,
      occurrences,
      sections: [...lineSections, ...cellSections],
      detected,
      detectWordsFound,
      phrasesFound,
      budgetKind: kind.kind,
      budgetKindSource: kind.source,
    };
  };

  /**
   * 自動で判別した書類の種類
   * @param {DocAnalysis} analysis
   * @returns {Selection}
   */
  const autoSelection = (analysis) => ({ types: [...analysis.detected], budgetKind: analysis.budgetKind, auto: true });

  /**
   * 選んだ書類の種類の名前（例: 「事業計画書」「収支予算書・決算書（予算）」）
   * @param {Selection} selection
   * @param {RulesetLike} ruleset
   * @returns {string[]}
   */
  const typeNames = (selection, ruleset) => {
    const docs = ruleset.documents;
    if (!docs) return [];
    return docs.types
      .filter((t) => selection.types.includes(t.id))
      .map((t) => (t.settlementRequired.length > 0 ? `${t.name}（${selection.budgetKind}）` : t.name));
  };

  /**
   * 複数のファイルの分析を、1つの書類として合わせる（同じ種類のファイルをまとめて必須項目を判定するのに使う。付録F-3）
   * 欄の範囲（sections）はファイルごとの位置なので合わせない
   * @param {readonly DocAnalysis[]} list
   * @returns {DocAnalysis}
   */
  const mergeAnalyses = (list) => {
    if (list.length === 0) return emptyAnalysis('paste');
    if (list.length === 1) return list[0];
    /** @type {Record<string, string[]>} */
    const detectWordsFound = {};
    for (const a of list) {
      for (const [id, words] of Object.entries(a.detectWordsFound)) {
        detectWordsFound[id] = [...new Set([...(detectWordsFound[id] ?? []), ...words])];
      }
    }
    return {
      available: list.every((a) => a.available),
      format: list[0].format,
      formats: [...new Set(list.flatMap((a) => a.formats ?? [a.format]))],
      occurrences: list.flatMap((a) => a.occurrences),
      sections: [],
      detected: [...new Set(list.flatMap((a) => a.detected))],
      detectWordsFound,
      phrasesFound: [...new Set(list.flatMap((a) => a.phrasesFound))],
      budgetKind: list[0].budgetKind,
      budgetKindSource: list[0].budgetKindSource,
    };
  };

  /** @param {DocAnalysis} a @returns {Format[]} */
  const formatsOf = (a) => a.formats ?? [a.format];

  /**
   * 必須項目の表
   * @param {DocAnalysis} analysis
   * @param {Selection} selection
   * @param {RulesetLike} ruleset
   * @param {Readonly<Record<string, DocAnalysis>>} [byType]  書類の種類ごとに使う分析（複数のファイルのとき。ない種類は analysis）
   * @returns {RequiredReport}
   */
  const requiredReport = (analysis, selection, ruleset, byType = {}) => {
    const docs = ruleset.documents;
    if (!docs || !analysis.available) return { tables: [], total: 0, satisfied: 0, shortage: 0, notes: [] };
    const byKey = new Map(docs.headingKeys.map((h) => [h.key, h.name]));
    const canonical = (/** @type {string} */ w) => byKey.get(engine.labelKey(w)) ?? w;
    const phraseNames = new Set(docs.phrases.map((p) => p.name));
    /** @param {string} typeId */
    const analysisFor = (typeId) => byType[typeId] ?? analysis;

    /**
     * @param {DocAnalysis} a
     * @param {string} name
     * @param {boolean} blankCheck
     * @returns {ItemStatus}
     */
    const statusOf = (a, name, blankCheck) => {
      if (phraseNames.has(name)) return a.phrasesFound.includes(name) ? 'present' : 'missing';
      const occ = a.occurrences.filter((o) => o.name === name);
      if (occ.length === 0) return 'missing';
      const cells = occ.filter((o) => o.mode === 'cell');
      if (!blankCheck || cells.length === 0) return 'present';
      if (cells.some((o) => o.filled)) return 'filled';
      return occ.length > cells.length ? 'present' : 'blank';
    };

    const tables = docs.types
      .filter((t) => selection.types.includes(t.id))
      .map((t) => {
        const a = analysisFor(t.id);
        const names = [
          ...t.required,
          ...(selection.budgetKind === '決算' ? t.settlementRequired : []),
        ].map((n) => (phraseNames.has(n) ? n : canonical(n)));
        const blank = new Set(t.blankCheck.map(canonical));
        /**
         * 見つからないときは、代わりの見出し（documents.substitutes）の欄で確かめる
         * @param {string} name
         * @returns {RequiredRow}
         */
        const rowOf = (name) => {
          const own = statusOf(a, name, blank.has(name));
          const via =
            own === 'missing'
              ? (docs.substitutes[name] ?? []).find((alt) => statusOf(a, alt, blank.has(name)) !== 'missing') ?? null
              : null;
          const status = via ? statusOf(a, via, blank.has(name)) : own;
          return { name, status, label: via ? `${STATUS_LABELS[status]}（${via}の欄）` : STATUS_LABELS[status], via };
        };
        return {
          typeId: t.id,
          title: t.settlementRequired.length > 0 ? `${t.name}（${selection.budgetKind}）` : t.name,
          rows: [...new Set(names)].map(rowOf),
        };
      });

    // 種類をまたいで同じ見出し（委員会名など）は1つと数える
    /** @type {Map<string, ItemStatus>} */
    const unique = new Map();
    for (const row of tables.flatMap((t) => t.rows)) if (!unique.has(row.name)) unique.set(row.name, row.status);
    const statuses = [...unique.values()];
    const shortage = statuses.filter((s) => s === 'missing' || s === 'blank').length;
    const hasMissing = statuses.includes('missing');
    const hasBlank = statuses.includes('blank');

    // 表を作るのに使った分析の形式（複数のファイルのときは、それぞれの形式）
    const formats = new Set(tables.flatMap((t) => formatsOf(analysisFor(t.typeId))));
    const budgetSource = analysisFor('budget').budgetKindSource;
    /** 空欄になった見出しが、どの形式の表から来たか（Word と Excel で注意書きを分ける） */
    const blankFormats = new Set(
      tables.flatMap((t) => (t.rows.some((r) => r.status === 'blank') ? formatsOf(analysisFor(t.typeId)) : [])),
    );
    /** @type {string[]} */
    const notes = tables.length === 0 ? [] : [
      ...(formats.has('paste') || formats.has('pdf') ? [MESSAGES.pasteNoBlank] : []),
      ...(formats.has('docx') ? [MESSAGES.docxNoBlank] : []),
      ...(hasMissing ? [MESSAGES.partial] : []),
      ...(hasBlank && (blankFormats.has('xlsx') || !blankFormats.has('docx')) ? [MESSAGES.xlsxBlank] : []),
      ...(hasBlank && blankFormats.has('docx') ? [MESSAGES.docxBlank] : []),
      ...(selection.types.includes('budget') && selection.auto && budgetSource === 'ambiguous' ? [MESSAGES.budgetAmbiguous] : []),
      ...(selection.types.includes('budget') && selection.auto && budgetSource === 'table' ? [MESSAGES.budgetByTable] : []),
    ];
    return { tables, total: statuses.length, satisfied: statuses.length - shortage, shortage, notes };
  };

  /**
   * 件数のまとめに出す必須項目の文（例: 「必須項目 14／14」「必須項目 13／14／必須項目の不足 1」）
   * @param {RequiredReport} rep
   */
  const requiredSummary = (rep) =>
    rep.tables.length === 0
      ? ''
      : `必須項目 ${rep.satisfied}／${rep.total}${rep.shortage > 0 ? `／必須項目の不足 ${rep.shortage}` : ''}`;

  /**
   * 確認ポイント（共通と、選んだ種類のもの）
   * @param {Selection} selection
   * @param {RulesetLike} ruleset
   * @param {ReadonlySet<string>} [checkedRules]  金額チェックで判定したルール（M-xx）。これで確かめた項目は checkedBy に出す
   * @returns {{ title: string, items: { id: string, text: string, checkedBy: string[] }[] }[]}
   */
  const checkpointGroups = (selection, ruleset, checkedRules = new Set()) => {
    const docs = ruleset.documents;
    if (!docs) return [];
    /** @param {{ id: string, text: string, checkedBy?: string[] }[]} items */
    const withChecked = (items) =>
      items.map((k) => ({ id: k.id, text: k.text, checkedBy: (k.checkedBy ?? []).filter((id) => checkedRules.has(id)) }));
    const selected = docs.types.filter((t) => selection.types.includes(t.id));
    // 共通の確認ポイントは議案向け。委員会事業計画書だけを選んだときは出さない
    const showCommon = selected.length === 0 || selected.some((t) => t.useCommonCheckpoints !== false);
    return [
      ...(showCommon ? [{ title: '共通', items: withChecked(docs.commonCheckpoints) }] : []),
      ...selected.map((t) => ({ title: t.name, items: withChecked(t.checkpoints) })),
    ].filter((grp) => grp.items.length > 0);
  };

  /**
   * 指摘一覧のコピーに入れる、書類の種類と必須項目の行
   * @param {Selection} selection
   * @param {RequiredReport} rep
   * @param {RulesetLike} ruleset
   * @returns {string[]}
   */
  const copyLines = (selection, rep, ruleset) => {
    if (!ruleset.documents) return [];
    const names = typeNames(selection, ruleset);
    const how = selection.auto ? '自動で判別' : '手動で選択';
    if (names.length === 0) {
      return [`書類の種類: ${selection.auto ? '判別できませんでした' : '選ばれていません'}（必須項目のチェックは行っていません）`];
    }
    const lacking = rep.tables
      .flatMap((t) => t.rows)
      .filter((r) => r.status === 'missing' || r.status === 'blank')
      .map((r) => `${r.name}（${r.label}）`);
    return [
      `書類の種類: ${names.join('／')}（${how}）`,
      `${requiredSummary(rep)}${lacking.length > 0 ? `: ${[...new Set(lacking)].join('、')}` : ''}`,
    ];
  };

  /**
   * ルール一覧の画面に出す、書類の種類ごとの必須項目と確認ポイント
   * @param {RulesetLike} ruleset
   */
  const documentsModel = (ruleset) => {
    const docs = ruleset.documents;
    if (!docs) return null;
    return {
      minDetectWords: docs.minDetectWords,
      common: docs.commonCheckpoints,
      types: docs.types.map((t) => ({
        id: t.id,
        name: t.name,
        detectWords: t.detectWords,
        required: t.required,
        settlementRequired: t.settlementRequired,
        blankCheck: t.blankCheck,
        checkpoints: t.checkpoints,
      })),
    };
  };

  root.YegDocCheck = Object.freeze({
    MESSAGES,
    STATUS_LABELS,
    analyzeDocument,
    mergeAnalyses,
    autoSelection,
    typeNames,
    requiredReport,
    requiredSummary,
    checkpointGroups,
    copyLines,
    documentsModel,
    parseLineHeading,
    parseCellHeading,
  });
})(globalThis);
