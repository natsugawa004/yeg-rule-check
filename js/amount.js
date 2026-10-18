// @ts-check
/*
 * 収支予算書・決算書の金額チェック（F-04）。画面に依存しない純粋な関数だけを置く
 *   readSheet(Excel のシート, 金額チェックのデータ) → 収支表の形の読み取り（予算か決算かを決める前）
 *   hasSettlementAmount(読み取り) → 決算額の列に 0 以外の金額があるか（予算・決算の印がないときの判別に使う）
 *   finalize(読み取り, 予算か決算か) → 判定に使える収支表、または読み取れなかった理由
 *   checkTable(収支表, 予算か決算か, ルール) → 金額チェックの指摘のもと（M-01〜M-16）
 *   tableView(収支表, 予算か決算か) → 画面に出す「読み取った収支表」
 * 誤った指摘を出すより判定しないことを優先する（読み取れないときは理由を返し、金額の指摘は出さない）
 */
'use strict';

(function (root) {
  /** @typedef {'予算'|'決算'} BudgetKind */
  /** @typedef {'budget'|'settlement'} Column */
  /** @typedef {'違反'|'要確認'|'参考'} Severity */

  /**
   * xlsx.js の XlsxCell（使う項目だけ）
   * @typedef {{ ref: string, row: number, col: number, text: string, kind: string, value: number|null, formula: boolean, cached: boolean }} XCell
   * @typedef {{ r1: number, c1: number, r2: number, c2: number }} XMerge
   * @typedef {{ name: string, cells: XCell[], merges: XMerge[] }} XSheet
   */

  /**
   * engine.js の CompiledAmounts（語は labelKey でそろえた形）
   * @typedef {Object} AmountData
   * @property {string} documentType
   * @property {string[]} incomeHeadings
   * @property {string[]} expenseHeadings
   * @property {string[]} headingSuffixes
   * @property {string[]} budgetColumns
   * @property {string[]} settlementColumns
   * @property {string[]} noteColumns
   * @property {string[]} budgetNoteWords
   * @property {string[]} settlementNoteWords
   * @property {string[]} incomeTotals
   * @property {string[]} expenseTotals
   * @property {string[]} balanceTotals
   * @property {string[]} summaryBudgetRows
   * @property {string[]} summarySettlementRows
   * @property {string} activityFee
   * @property {string} specialFee
   * @property {string[]} incomeItems
   * @property {string[]} salesWords
   * @property {string[]} purchaseWords
   * @property {string[]} honorariumWords
   * @property {string[]} withholdingWords
   * @property {number} withholdingRatePercent
   * @property {number} withholdingLimit
   * @property {string} lecturerItem
   * @property {string[]} lecturerBreakdownWords
   * @property {string} reserveItem
   * @property {number} reserveMaxPercent
   * @property {string[]} expenseItems
   * @property {string[]} multiplySigns
   * @property {string[]} numberMarks
   * @property {{ documentRules: string[], cellRules: string[] }} replaces
   */

  /**
   * engine.js の CompiledRule（金額チェックで使う項目だけ）
   * @typedef {{ id: string, name: string, category: string, severity: Severity, type: string, check: string,
   *   messages: Readonly<Record<string, string>>, hint: string, description: string, source: string, order: number }} AmountRule
   */

  /** @typedef {{ ref: string, row: number, col: number }} CellPos */

  /**
   * 金額のセル1つ
   * @typedef {Object} Amount
   * @property {number|null} value   数として読めた値（空のセルは 0。読めないときは null）
   * @property {string} ref          番地（セルが空でも、その列の番地）
   * @property {number} row
   * @property {number} col
   * @property {string} shown        画面に見える文字（空のセルは ''）
   * @property {boolean} present     セルに文字か数式がある
   * @property {'text'|'nocache'|null} problem  読めない文字／計算結果が保存されていない数式
   */

  /**
   * 収入・支出の項目の行
   * @typedef {Object} ItemRow
   * @property {string} id
   * @property {'income'|'expense'} section
   * @property {number} row
   * @property {string} label     項目名（先頭の番号を除く）
   * @property {CellPos} labelCell
   * @property {Amount|null} budget
   * @property {Amount|null} settlement
   * @property {string} budgetNote
   * @property {string} settlementNote
   * @property {CellPos|null} budgetNoteCell
   * @property {CellPos|null} settlementNoteCell
   */

  /**
   * 合計の行（収入合計・支出合計・残高合計）
   * @typedef {Object} TotalRow
   * @property {string} id
   * @property {'incomeTotal'|'expenseTotal'|'balance'} kind
   * @property {string} label
   * @property {number} row
   * @property {CellPos} labelCell
   * @property {Amount|null} budget
   * @property {Amount|null} settlement
   */

  /**
   * 収支計算欄の1行（事業収支予算・事業収支決算）
   * @typedef {{ row: number, label: string, income: Amount|null, expense: Amount|null, balance: Amount|null }} SummaryRow
   */

  /**
   * 収支表の読み取り（予算か決算かを決める前）
   * @typedef {Object} RawTable
   * @property {string} sheet
   * @property {ItemRow[]} income
   * @property {ItemRow[]} expense
   * @property {TotalRow} incomeTotal
   * @property {TotalRow} expenseTotal
   * @property {TotalRow|null} balance
   * @property {{ budget: SummaryRow|null, settlement: SummaryRow|null }} summary
   * @property {boolean} hasSettlementColumn  収入・支出の両方に「決算額」の列がある
   * @property {boolean} hasNote              「摘要」の列がある
   */

  /**
   * @typedef {{ status: 'none' } | { status: 'error', reasons: string[] } | { status: 'ok', table: RawTable }} SheetRead
   */

  /**
   * 金額チェックの指摘のもと（ファイル名・文章の中の位置は、呼び出し側で付ける）
   * @typedef {Object} Issue
   * @property {AmountRule} rule
   * @property {string} message
   * @property {CellPos} cell       指摘のセル
   * @property {string} rowLabel    行の項目名（例: 特別会費・支出合計）
   * @property {string} rowId       読み取った収支表の行
   * @property {string} column      '予算額'・'決算額'（決算（案）で列を分けて見せるとき）か ''
   * @property {string} calc        計算に使った数値（例: 「3,000×40＝120,000」）
   */

  /**
   * 画面に出す「読み取った収支表」の1行
   * @typedef {{ id: string, group: string, label: string, budget: string, settlement: string, note: string, total: boolean }} ViewRow
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (root));
  const engine = /** @type {{ labelKey: (s: string) => string }} */ (g.YegEngine);

  const MESSAGES = Object.freeze({
    noIncome: '「収入」の見出しが見つかりません',
    noExpense: '「支出」の見出しが見つかりません',
    noBudgetColumn: '「予算額」の列が見つかりません',
    noSettlementColumn: '「決算額」の列が見つかりません',
    noIncomeTotal: '「収入合計」の行が見つかりません',
    noExpenseTotal: '「支出合計」の行が見つかりません',
    noCache: '計算結果が保存されていないセルがあります（Excelで開いて保存し直してください）',
    /** @param {string} ref */
    notNumber: (ref) => `${ref} の金額を数値として読めません`,
  });

  /** 判定に使う列（予算（案）は予算額、決算（案）は決算額） */
  /** @param {BudgetKind} kind @returns {Column} */
  const mainColumn = (kind) => (kind === '決算' ? 'settlement' : 'budget');

  /** 合計の整合（M-01〜M-03）を見る列（決算（案）は承認済みの予算額も） */
  /** @param {BudgetKind} kind @returns {Column[]} */
  const totalColumns = (kind) => (kind === '決算' ? ['budget', 'settlement'] : ['budget']);

  /** @type {Readonly<Record<Column, string>>} */
  const COLUMN_LABELS = Object.freeze({ budget: '予算額', settlement: '決算額' });

  // ------------------------------------------------------------------
  // 数と文字
  // ------------------------------------------------------------------

  /** @param {number} col */
  const colName = (col) => {
    let s = '';
    for (let c = col; c > 0; c = Math.floor((c - 1) / 26)) s = String.fromCharCode(65 + ((c - 1) % 26)) + s;
    return s;
  };

  /**
   * 金額の表示（3桁区切り。例: 250000 → 「250,000」、-1000 → 「-1,000」）
   * @param {number} n
   */
  const formatYen = (n) => {
    const sign = n < 0 ? '-' : '';
    const [int, frac] = String(Math.abs(n)).split('.');
    return `${sign}${int.replace(/\B(?=(\d{3})+(?!\d))/g, ',')}${frac ? `.${frac}` : ''}`;
  };

  /**
   * 指摘の文の { 名前 } を値に置き換える
   * @param {string} template
   * @param {Readonly<Record<string, string>>} values
   */
  const fillMessage = (template, values) => template.replace(/\{(\w+)\}/g, (m, k) => (k in values ? values[k] : m));

  /**
   * 全角の数字・記号を半角にする（○付き数字などは変えない）。
   * 「\」は円記号にする（日本語の Windows では、円記号を「\」の文字で入力・表示することがあるため）
   */
  const HALF = /** @type {Readonly<Record<string, string>>} */ (
    Object.freeze({ '，': ',', '．': '.', '（': '(', '）': ')', '＝': '=', '＠': '@', '＊': '*', 'Ｘ': 'X', 'ｘ': 'x', '￥': '¥', '＼': '¥', '\\': '¥', '／': '/', '　': ' ', '：': ':' })
  );
  /** @param {string} s */
  const toHalf = (s) =>
    s.replace(/[０-９]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0)).replace(/[，．（）＝＠＊Ｘｘ￥＼\\／　：]/g, (c) => HALF[c] ?? c);

  /**
   * 文字で書かれた金額を数にする（「100,000」「¥100,000」「100,000円」「-1,000」）。読めなければ null
   * @param {string} text
   * @returns {number|null}
   */
  const parseAmountText = (text) => {
    const t = toHalf(text).replace(/\s/g, '');
    if (t === '') return 0;
    if (/^[-−―ー‐]$/.test(t)) return 0;
    const m = /^[¥\\]?([-−△▲])?[¥\\]?((?:\d{1,3}(?:,\d{3})+|\d+)(?:\.\d+)?)円?$/.exec(t);
    if (!m) return null;
    const n = Number(m[2].replace(/,/g, ''));
    return Number.isFinite(n) ? (m[1] ? -n : n) : null;
  };

  /** @param {string} s */
  const escapeRe = (s) => s.replace(/[.*+?^${}()|[\]\\\/-]/g, '\\$&');

  // ------------------------------------------------------------------
  // 摘要の「数×数」（付録E-2）
  // ------------------------------------------------------------------

  /**
   * @typedef {{ kind: 'none' } | { kind: 'many' } | { kind: 'mixed' } | { kind: 'one', expr: string, operands: number[], product: number }} Formula
   */

  const NUM = String.raw`\d[\d,]*(?:\.\d+)?`;

  /** 式のほかに書かれた金額（¥・円の付いた数、桁区切りのある数、3桁以上の数。単位のない金額もあるため、迷う数は金額とみなす） */
  const OTHER_AMOUNT = new RegExp(String.raw`¥\s*${NUM}|${NUM}\s*円|\d{1,3}(?:,\d{3})+|\d{3,}`, 'u');

  /**
   * 摘要の式を読む。括弧の中の数は使わない。式がちょうど1つのときだけ計算する（2つ以上は 'many'）。
   * 式のほかにも金額（¥・円の付いた数、桁区切りのある数。括弧の中も含む）が書かれているときは、
   * 内訳の一部の式とみなして 'mixed'（判定しない）。例: 「照明 4,000円、マイク 1,000円×2本=2,000円」
   * @param {string} note
   * @param {AmountData} data
   * @returns {Formula}
   */
  const parseFormula = (note, data) => {
    let s = toHalf(note);
    /** @type {string[]} 取り除いた括弧の中 */
    const removed = [];
    for (let prev = ''; prev !== s; ) {
      prev = s;
      s = s.replace(/\([^()]*\)/g, (m) => {
        removed.push(m);
        return ' ';
      });
    }
    const signs = data.multiplySigns.map((c) => escapeRe(toHalf(c))).join('');
    const marks = [...new Set(data.numberMarks.map((c) => toHalf(c)))].map(escapeRe).join('');
    // 数の後ろの単位（円・名など）。「2区分」「2部屋」「3セット」のような漢字・カタカナの単位も式の一部として読む
    const operand = String.raw`(?:@\s*)?(${NUM})(?:\s*\/?\s*[${marks}])*[ァ-ヺー一-鿿々]{0,3}`;
    const chain = new RegExp(String.raw`(?<![\d,.])${operand}(?:\s*[${signs}]\s*${operand})+`, 'gu');
    const one = new RegExp(operand, 'u');
    const found = [...s.matchAll(chain)];
    if (found.length === 0) return { kind: 'none' };
    if (found.length > 1) return { kind: 'many' };
    const m = found[0];
    const expr = m[0].trim();
    const operands = expr.split(new RegExp(`\\s*[${signs}]\\s*`, 'u')).map((part) => {
      const om = one.exec(part);
      return om ? Number(om[1].replace(/,/g, '')) : NaN;
    });
    if (operands.some((n) => !Number.isFinite(n))) return { kind: 'none' };
    // 式の後ろの「＝金額」は式の一部として除き、残りに金額があるかを見る
    const at = /** @type {number} */ (m.index);
    const rest = `${s.slice(0, at)} ${s.slice(at + m[0].length).replace(new RegExp(String.raw`^\s*=\s*¥?\s*${NUM}\s*円?`, 'u'), ' ')}`;
    if (OTHER_AMOUNT.test(rest) || removed.some((r) => OTHER_AMOUNT.test(r))) return { kind: 'mixed' };
    return { kind: 'one', expr, operands, product: operands.reduce((a, b) => a * b, 1) };
  };

  /**
   * 語の後ろに続く金額（例: 「源泉3,063円」「謝金25,000円」）
   * @param {string} note
   * @param {readonly string[]} words  labelKey でそろえた語
   * @returns {number|null}
   */
  const amountAfter = (note, words) => {
    if (words.length === 0) return null;
    const s = toHalf(note);
    const re = new RegExp(String.raw`(?:${words.map(escapeRe).join('|')})(?:徴収)?(?:税|分)?額?\s*:?\s*¥?\s*(${NUM})`, 'u');
    const m = re.exec(s);
    if (!m) return null;
    const n = Number(m[1].replace(/,/g, ''));
    return Number.isFinite(n) ? n : null;
  };

  // ------------------------------------------------------------------
  // シートの読み取り（付録E-1）
  // ------------------------------------------------------------------

  /**
   * @typedef {Object} Item   セル（結合されたセルは範囲全体）
   * @property {number} r1
   * @property {number} c1
   * @property {number} r2
   * @property {number} c2
   * @property {string} key   labelKey でそろえた文字
   * @property {XCell} cell
   */

  /**
   * 語が一致するか（「収入合計(自動計算)」のような かっこ書きや、suffixes の列名が後ろに付いていてもよい）
   * @param {string} key
   * @param {readonly string[]} words
   * @param {readonly string[]} [suffixes]
   */
  const matchWord = (key, words, suffixes = []) =>
    words.some((w) => {
      if (key === w) return true;
      if (!key.startsWith(w)) return false;
      const rest = key.slice(w.length);
      return suffixes.includes(rest) || /^\([^()]*\)$/.test(rest);
    });

  /** @param {{ c1: number, c2: number }} range @param {number} col */
  const inCols = (range, col) => col >= range.c1 && col <= range.c2;

  /**
   * @param {XSheet} sheet
   * @returns {Item[]}
   */
  const itemsOf = (sheet) => {
    /** @type {Map<string, XMerge>} */
    const mergeAt = new Map(sheet.merges.map((m) => [`${m.r1}|${m.c1}`, m]));
    return sheet.cells.map((c) => {
      const m = mergeAt.get(`${c.row}|${c.col}`);
      return { r1: c.row, c1: c.col, r2: m ? m.r2 : c.row, c2: m ? m.c2 : c.col, key: engine.labelKey(c.text), cell: c };
    });
  };

  /**
   * セルの金額を読む
   * @param {XCell|undefined} cell
   * @param {number} row
   * @param {number} col  セルがないときに番地に使う列
   * @returns {Amount}
   */
  const readAmount = (cell, row, col) => {
    if (!cell) return { value: 0, ref: `${colName(col)}${row}`, row, col, shown: '', present: false, problem: null };
    const base = { ref: cell.ref, row: cell.row, col: cell.col, shown: cell.text, present: true };
    if (cell.formula && !cell.cached) return { ...base, value: null, problem: 'nocache' };
    if (cell.kind === 'number' && cell.value !== null && Number.isFinite(cell.value)) return { ...base, value: cell.value, problem: null };
    if (cell.kind === 'string' || cell.kind === 'empty') {
      const v = parseAmountText(cell.text);
      return v === null ? { ...base, value: null, problem: 'text' } : { ...base, value: v, problem: null };
    }
    return { ...base, value: null, problem: 'text' };
  };

  /**
   * 見出しの行（または次の行）にある列見出し
   * @param {Item[]} items
   * @param {Item} heading
   * @param {AmountData} data
   */
  const findColumns = (items, heading, data) => {
    const rows = [];
    for (let r = heading.r1; r <= heading.r2 + 1; r += 1) rows.push(r);
    /** @param {number} row */
    const inRow = (row) => items.filter((it) => it.r1 === row && it.c1 > heading.c2).sort((a, b) => a.c1 - b.c1);
    // 「予算額」の見つかる最初の行を列見出しの行とする（なければ見出しと同じ行）
    const headerRow = rows.find((r) => inRow(r).some((it) => matchWord(it.key, data.budgetColumns))) ?? heading.r1;
    const cells = inRow(headerRow);
    const budget = cells.find((it) => matchWord(it.key, data.budgetColumns)) ?? null;
    const settlement = cells.find((it) => matchWord(it.key, data.settlementColumns)) ?? null;
    const notes = cells.filter((it) => matchWord(it.key, data.noteColumns)).slice(0, 2);
    /**
     * 摘要の列が予算・決算のどちらのものか（列見出しか、その上のセルの語で決める）
     * @param {Item} n
     * @returns {'budget'|'settlement'|null}
     */
    const noteRole = (n) => {
      const above = items.filter((it) => it.r2 === n.r1 - 1 && it.c1 <= n.c2 && it.c2 >= n.c1).map((it) => it.key).join('');
      const k = `${n.key}${above}`;
      if (data.settlementNoteWords.some((w) => k.includes(w))) return 'settlement';
      if (data.budgetNoteWords.some((w) => k.includes(w))) return 'budget';
      return null;
    };
    /** @type {{ budget: Item|null, settlement: Item|null }} */
    let noteCols = { budget: null, settlement: null };
    if (notes.length === 1) noteCols = { budget: notes[0], settlement: notes[0] };
    else if (notes.length === 2) {
      const [left, right] = notes;
      const lr = noteRole(left);
      const rr = noteRole(right);
      noteCols = lr === 'settlement' || rr === 'budget' ? { budget: right, settlement: left } : { budget: left, settlement: right };
    }
    return { budget, settlement, notes: noteCols, headerRow };
  };

  /**
   * シートを収支表として読み取る。収入・支出の見出しがなければ 'none'（収支表のシートではない）
   * @param {XSheet} sheet
   * @param {AmountData} data
   * @returns {SheetRead}
   */
  const readSheet = (sheet, data) => {
    const items = itemsOf(sheet).filter((it) => it.key !== '' || it.cell.formula);
    const byPos = (/** @type {Item} */ a, /** @type {Item} */ b) => a.r1 - b.r1 || a.c1 - b.c1;
    const incomeHeads = items.filter((it) => matchWord(it.key, data.incomeHeadings, data.headingSuffixes)).sort(byPos);
    const expenseHeads = items.filter((it) => matchWord(it.key, data.expenseHeadings, data.headingSuffixes)).sort(byPos);
    if (incomeHeads.length === 0 && expenseHeads.length === 0) return { status: 'none' };

    /** 「予算額」の列見出しが並んでいる見出しを優先する */
    const pick = (/** @type {Item[]} */ heads) => heads.find((h) => findColumns(items, h, data).budget !== null) ?? heads[0] ?? null;
    const incomeHead = pick(incomeHeads);
    const expenseHead = pick(expenseHeads);
    if (!incomeHead || !expenseHead) {
      return { status: 'error', reasons: [...(incomeHead ? [] : [MESSAGES.noIncome]), ...(expenseHead ? [] : [MESSAGES.noExpense])] };
    }
    const incomeCols = findColumns(items, incomeHead, data);
    const expenseCols = findColumns(items, expenseHead, data);
    /** @type {string[]} */
    const reasons = [];
    if (!incomeCols.budget || !expenseCols.budget) reasons.push(MESSAGES.noBudgetColumn);

    /**
     * 見出しの下（項目の列）にある合計の行
     * @param {Item} head
     * @param {readonly string[]} words
     * @param {number} limit  この行より下は探さない
     */
    const totalBelow = (head, words, limit) =>
      items
        .filter((it) => it.r1 > head.r2 && it.r1 < limit && inCols(head, it.c1) && matchWord(it.key, words))
        .sort(byPos)[0] ?? null;
    const incomeLimit = expenseHead.r1 > incomeHead.r1 ? expenseHead.r1 : Infinity;
    const expenseLimit = incomeHead.r1 > expenseHead.r1 ? incomeHead.r1 : Infinity;
    const incomeTotalAt = totalBelow(incomeHead, data.incomeTotals, incomeLimit);
    const expenseTotalAt = totalBelow(expenseHead, data.expenseTotals, expenseLimit);
    if (!incomeTotalAt) reasons.push(MESSAGES.noIncomeTotal);
    if (!expenseTotalAt) reasons.push(MESSAGES.noExpenseTotal);
    if (reasons.length > 0 || !incomeTotalAt || !expenseTotalAt || !incomeCols.budget || !expenseCols.budget) {
      return { status: 'error', reasons };
    }

    /** @param {number} row @param {Item|null} col @returns {Amount|null} */
    const amountAt = (row, col) => {
      if (!col) return null;
      const cell = items.filter((it) => it.r1 === row && inCols(col, it.c1)).sort(byPos)[0];
      return readAmount(cell ? cell.cell : undefined, row, col.c1);
    };
    /** @param {number} row @param {Item|null} col */
    const noteAt = (row, col) => {
      if (!col) return { text: '', cell: null };
      const cells = items.filter((it) => it.r1 === row && inCols(col, it.c1) && it.cell.text.trim() !== '').sort(byPos);
      return {
        text: cells.map((it) => it.cell.text.trim()).join('\n'),
        cell: cells.length > 0 ? { ref: cells[0].cell.ref, row, col: cells[0].c1 } : null,
      };
    };

    /**
     * 見出しの次の行から合計の手前までの、項目の行
     * @param {'income'|'expense'} section
     * @param {Item} head
     * @param {Item} total
     * @param {ReturnType<typeof findColumns>} cols
     * @returns {ItemRow[]}
     */
    const rowsBetween = (section, head, total, cols) => {
      const top = Math.max(head.r2, cols.headerRow) + 1;
      const rowNos = [...new Set(items.filter((it) => it.r1 >= top && it.r1 < total.r1).map((it) => it.r1))].sort((a, b) => a - b);
      return rowNos.flatMap((row) => {
        const labelCells = items.filter((it) => it.r1 === row && inCols(head, it.c1)).sort(byPos);
        const label = stripNumber(labelCells.map((it) => it.cell.text.trim()).filter((t) => t !== '').join(' '));
        const budget = amountAt(row, cols.budget);
        const settlement = amountAt(row, cols.settlement);
        const bn = noteAt(row, cols.notes.budget);
        const sn = noteAt(row, cols.notes.settlement);
        const nameCell = labelCells.find((it) => stripNumber(it.cell.text.trim()) !== '') ?? labelCells[0];
        if (label === '' && !(budget && budget.present) && !(settlement && settlement.present) && bn.text === '' && sn.text === '') {
          return [];
        }
        return [
          {
            id: `r${row}`,
            section,
            row,
            label,
            labelCell: nameCell ? { ref: nameCell.cell.ref, row, col: nameCell.c1 } : { ref: `${colName(head.c1)}${row}`, row, col: head.c1 },
            budget,
            settlement,
            budgetNote: bn.text,
            settlementNote: sn.text,
            budgetNoteCell: bn.cell,
            settlementNoteCell: sn.cell,
          },
        ];
      });
    };

    /**
     * @param {'incomeTotal'|'expenseTotal'|'balance'} kind
     * @param {Item} at
     * @param {ReturnType<typeof findColumns>} cols
     * @returns {TotalRow}
     */
    const totalRow = (kind, at, cols) => ({
      id: `r${at.r1}`,
      kind,
      label: at.cell.text.replace(/\s+/g, ''),
      row: at.r1,
      labelCell: { ref: at.cell.ref, row: at.r1, col: at.c1 },
      budget: amountAt(at.r1, cols.budget),
      settlement: amountAt(at.r1, cols.settlement),
    });

    // 残高合計の行（あれば）: 収入合計・支出合計より下の、項目の列
    const lowest = Math.max(incomeTotalAt.r1, expenseTotalAt.r1);
    const balanceAt =
      items
        .filter((it) => it.r1 > lowest && (inCols(incomeHead, it.c1) || inCols(expenseHead, it.c1)) && matchWord(it.key, data.balanceTotals))
        .sort(byPos)[0] ?? null;

    /**
     * 収支計算欄の行（「事業収支予算」などの行に、収入合計・支出合計・残高合計の列が並ぶ形）
     * @param {readonly string[]} words
     * @returns {SummaryRow|null}
     */
    const summaryRow = (words) => {
      const at = items.filter((it) => words.length > 0 && words.includes(it.key)).sort(byPos)[0];
      if (!at) return null;
      const headerOf = (/** @type {number} */ r) =>
        items.filter(
          (it) =>
            it.r1 === r &&
            it.c1 > at.c2 &&
            (matchWord(it.key, data.incomeTotals) || matchWord(it.key, data.expenseTotals) || matchWord(it.key, data.balanceTotals)),
        );
      let header = /** @type {Item[]} */ ([]);
      for (let r = at.r1 - 1; r >= Math.max(1, at.r1 - 3) && header.length === 0; r -= 1) header = headerOf(r);
      if (header.length === 0) return null;
      /** @param {readonly string[]} w */
      const valueOf = (w) => {
        const h = header.find((it) => matchWord(it.key, w));
        return h ? amountAt(at.r1, h) : null;
      };
      return {
        row: at.r1,
        label: at.cell.text.trim(),
        income: valueOf(data.incomeTotals),
        expense: valueOf(data.expenseTotals),
        balance: valueOf(data.balanceTotals),
      };
    };

    return {
      status: 'ok',
      table: {
        sheet: sheet.name,
        income: rowsBetween('income', incomeHead, incomeTotalAt, incomeCols),
        expense: rowsBetween('expense', expenseHead, expenseTotalAt, expenseCols),
        incomeTotal: totalRow('incomeTotal', incomeTotalAt, incomeCols),
        expenseTotal: totalRow('expenseTotal', expenseTotalAt, expenseCols),
        balance: balanceAt ? totalRow('balance', balanceAt, expenseCols) : null,
        summary: { budget: summaryRow(data.summaryBudgetRows), settlement: summaryRow(data.summarySettlementRows) },
        hasSettlementColumn: incomeCols.settlement !== null && expenseCols.settlement !== null,
        hasNote: [incomeCols, expenseCols].some((c) => c.notes.budget !== null || c.notes.settlement !== null),
      },
    };
  };

  /**
   * 項目名の先頭の番号を除く（「1 会場費」「1．会場費」「①会場費」「(1)会場費」）
   * @param {string} s
   */
  const stripNumber = (s) =>
    s
      .replace(/^[\s　]*(?:[(（][0-9０-９]{1,3}[)）]|[①-⑳]|[0-9０-９]{1,3}[.．、)）])[\s　]*/u, '')
      .replace(/^[\s　]*[0-9０-９]{1,3}[\s　]+/u, '')
      .replace(/^[\s　]*[0-9０-９]{1,3}$/u, '')
      .trim();

  // ------------------------------------------------------------------
  // 予算か決算かを決めた後の確かめ
  // ------------------------------------------------------------------

  /**
   * 表の金額のセル（列ごと）
   * @param {RawTable} t
   * @param {Column} col
   * @returns {Amount[]}
   */
  const amountsIn = (t, col) => {
    const sum = col === 'budget' ? t.summary.budget : t.summary.settlement;
    return [
      ...t.income.map((r) => r[col]),
      t.incomeTotal[col],
      ...t.expense.map((r) => r[col]),
      t.expenseTotal[col],
      t.balance ? t.balance[col] : null,
      ...(sum ? [sum.income, sum.expense, sum.balance] : []),
    ].flatMap((a) => (a ? [a] : []));
  };

  /**
   * 決算額の列に 0 以外の金額があるか（予算・決算の印も「予算（案）」「決算（案）」の文字もないときの判別。付録F-2）
   * @param {RawTable} t
   */
  const hasSettlementAmount = (t) => amountsIn(t, 'settlement').some((a) => a.value !== null && a.value !== 0);

  /**
   * 予算か決算かに合わせて、判定に使えるかを確かめる
   * @param {RawTable} t
   * @param {BudgetKind} kind
   * @returns {{ ok: true, table: RawTable } | { ok: false, reasons: string[] }}
   */
  const finalize = (t, kind) => {
    if (kind === '決算' && !t.hasSettlementColumn) return { ok: false, reasons: [MESSAGES.noSettlementColumn] };
    const cells = totalColumns(kind).flatMap((c) => amountsIn(t, c));
    const noCache = cells.some((a) => a.problem === 'nocache');
    const bad = [...new Set(cells.filter((a) => a.problem === 'text').map((a) => a.ref))];
    const reasons = [...(noCache ? [MESSAGES.noCache] : []), ...bad.slice(0, 3).map(MESSAGES.notNumber)];
    return reasons.length > 0 ? { ok: false, reasons } : { ok: true, table: t };
  };

  /**
   * 金額の列のセル（R-15・R-16 の代わりに M-16 で判定するセル）
   * @param {RawTable} t
   * @returns {CellPos[]}
   */
  const amountCells = (t) =>
    [...amountsIn(t, 'budget'), ...amountsIn(t, 'settlement')].filter((a) => a.present).map((a) => ({ ref: a.ref, row: a.row, col: a.col }));

  // ------------------------------------------------------------------
  // 判定（付録E-3）
  // ------------------------------------------------------------------

  /** @param {Amount|null} a */
  const val = (a) => (a && a.value !== null ? a.value : 0);

  /** @param {Amount|null} a @param {CellPos} fallback @returns {CellPos} */
  const posOf = (a, fallback) => (a ? { ref: a.ref, row: a.row, col: a.col } : fallback);

  /**
   * 収支表を判定する
   * @param {RawTable} t  finalize で確かめた収支表
   * @param {BudgetKind} kind
   * @param {readonly AmountRule[]} rules  有効な金額チェックのルール
   * @param {AmountData} data
   * @returns {Issue[]}
   */
  const checkTable = (t, kind, rules, data) => {
    /** @type {Map<string, AmountRule>} */
    const byCheck = new Map(rules.filter((r) => r.type === 'amount').map((r) => [r.check, r]));
    /** @type {Issue[]} */
    const issues = [];
    const main = mainColumn(kind);
    const columnLabel = (/** @type {Column} */ c) => (kind === '決算' ? COLUMN_LABELS[c] : '');
    /**
     * @param {string} check
     * @param {string} variant
     * @param {Record<string, string>} values
     * @param {CellPos} cell
     * @param {{ rowLabel: string, rowId: string, column?: string, calc?: string }} where
     */
    const add = (check, variant, values, cell, where) => {
      const rule = byCheck.get(check);
      if (!rule) return;
      const template = rule.messages[variant] ?? rule.messages.default ?? Object.values(rule.messages)[0] ?? '';
      issues.push({
        rule,
        message: fillMessage(template, values),
        cell,
        rowLabel: where.rowLabel,
        rowId: where.rowId,
        column: where.column ?? '',
        calc: where.calc ?? '',
      });
    };
    /** @param {number[]} nums */
    const sumCalc = (nums) => `${nums.length > 0 ? nums.map(formatYen).join('＋') : '0'}＝${formatYen(nums.reduce((a, b) => a + b, 0))}`;

    // M-01・M-02・M-03: 合計の整合（決算（案）は予算額の列も）
    for (const c of totalColumns(kind)) {
      const inc = t.income.map((r) => val(r[c]));
      const exp = t.expense.map((r) => val(r[c]));
      const incSum = inc.reduce((a, b) => a + b, 0);
      const expSum = exp.reduce((a, b) => a + b, 0);
      const incTotal = val(t.incomeTotal[c]);
      const expTotal = val(t.expenseTotal[c]);
      const where = (/** @type {TotalRow} */ row) => ({ rowLabel: row.label, rowId: row.id, column: columnLabel(c) });
      if (incSum !== incTotal) {
        add('incomeTotal', 'default', { sum: formatYen(incSum), total: formatYen(incTotal) }, posOf(t.incomeTotal[c], t.incomeTotal.labelCell), {
          ...where(t.incomeTotal),
          calc: sumCalc(inc),
        });
      }
      if (expSum !== expTotal) {
        add('expenseTotal', 'default', { sum: formatYen(expSum), total: formatYen(expTotal) }, posOf(t.expenseTotal[c], t.expenseTotal.labelCell), {
          ...where(t.expenseTotal),
          calc: sumCalc(exp),
        });
      }
      const expected = incTotal - expTotal;
      const balCalc = `${formatYen(incTotal)}−${formatYen(expTotal)}＝${formatYen(expected)}`;
      if (t.balance && val(t.balance[c]) !== expected) {
        add('balance', 'balance', { balance: formatYen(val(t.balance[c])), expected: formatYen(expected) }, posOf(t.balance[c], t.balance.labelCell), {
          ...where(t.balance),
          calc: balCalc,
        });
      }
      const sum = c === 'budget' ? t.summary.budget : t.summary.settlement;
      if (sum) {
        /** @type {[Amount|null, string, number][]} */
        const pairs = [
          [sum.income, '収入合計', incTotal],
          [sum.expense, '支出合計', expTotal],
          [sum.balance, '残高合計', expected],
        ];
        for (const [a, label, want] of pairs) {
          if (!a || val(a) === want) continue;
          add('balance', 'summary', { label, value: formatYen(val(a)), expected: formatYen(want) }, posOf(a, t.incomeTotal.labelCell), {
            rowLabel: `${sum.label}・${label}`,
            rowId: '',
            column: '',
            calc: label === '残高合計' ? balCalc : '',
          });
        }
      }
    }

    const incTotal = val(t.incomeTotal[main]);
    const expTotal = val(t.expenseTotal[main]);
    const balance = incTotal - expTotal;
    const mainSummary = main === 'budget' ? t.summary.budget : t.summary.settlement;

    // M-04: 収支の差
    const balanceCell = t.balance
      ? posOf(t.balance[main], t.balance.labelCell)
      : mainSummary && mainSummary.balance
        ? posOf(mainSummary.balance, t.expenseTotal.labelCell)
        : posOf(t.expenseTotal[main], t.expenseTotal.labelCell);
    const balanceWhere = {
      rowLabel: t.balance ? t.balance.label : '残高',
      rowId: t.balance ? t.balance.id : t.expenseTotal.id,
      column: columnLabel(main),
      calc: `${formatYen(incTotal)}−${formatYen(expTotal)}＝${formatYen(balance)}`,
    };
    if (kind === '予算' && balance !== 0) add('surplus', balance > 0 ? 'plus' : 'minus', { balance: formatYen(balance) }, balanceCell, balanceWhere);
    if (kind === '決算' && balance < 0) add('surplus', 'settlementMinus', { balance: formatYen(balance) }, balanceCell, balanceWhere);

    // M-05: 会費などによる収益（決算）
    if (kind === '決算') {
      const fees = t.income.filter((r) => engine.labelKey(r.label) !== data.activityFee);
      const feeSum = fees.reduce((a, r) => a + val(r[main]), 0);
      if (fees.length > 0 && feeSum > expTotal) {
        add('feeSurplus', 'default', { income: formatYen(feeSum), expense: formatYen(expTotal) }, posOf(t.incomeTotal[main], t.incomeTotal.labelCell), {
          rowLabel: t.incomeTotal.label,
          rowId: t.incomeTotal.id,
          column: columnLabel(main),
          calc: `${formatYen(feeSum)}＞${formatYen(expTotal)}`,
        });
      }
    }

    // M-06: 予備費の割合（予算）
    if (kind === '予算') {
      const reserves = t.expense.filter((r) => engine.labelKey(r.label) === data.reserveItem);
      const reserve = reserves.reduce((a, r) => a + val(r.budget), 0);
      if (reserves.length > 0 && expTotal > 0 && reserve * 100 > data.reserveMaxPercent * expTotal) {
        const rate = (Math.round((reserve * 1000) / expTotal) / 10).toFixed(1);
        add('reserveRate', 'default', { rate }, posOf(reserves[0].budget, reserves[0].labelCell), {
          rowLabel: reserves[0].label,
          rowId: reserves[0].id,
          calc: `${formatYen(reserve)}÷${formatYen(expTotal)}＝${rate}%`,
        });
      }
    }

    // 摘要を使うルール（M-07〜M-13）は、摘要の列があるときだけ
    if (t.hasNote) {
      /** @param {ItemRow} r */
      const noteOf = (r) => (main === 'settlement' ? r.settlementNote : r.budgetNote);
      /** @param {ItemRow} r @returns {CellPos} */
      const noteCellOf = (r) => (main === 'settlement' ? r.settlementNoteCell : r.budgetNoteCell) ?? posOf(r[main], r.labelCell);
      /** @param {string} s @param {readonly string[]} words */
      const hasWord = (s, words) => {
        const k = engine.labelKey(s);
        return words.some((w) => k.includes(w));
      };
      /** 金額が0で摘要も空の行（使わなかった項目）は、摘要のルールを判定しない */
      const used = (/** @type {ItemRow} */ r) => val(r[main]) !== 0 || noteOf(r).trim() !== '';
      /** @param {ItemRow} r */
      const at = (r) => ({ rowLabel: r.label, rowId: r.id });

      // M-07: 物販の売上と仕入
      const sales = t.income.filter((r) => hasWord(r.label, data.salesWords) || hasWord(noteOf(r), data.salesWords));
      const purchases = t.expense.filter((r) => hasWord(r.label, data.purchaseWords) || hasWord(noteOf(r), data.purchaseWords));
      if (sales.length > 0 && purchases.length > 0) {
        const s = sales.reduce((a, r) => a + val(r[main]), 0);
        const p = purchases.reduce((a, r) => a + val(r[main]), 0);
        if (s < p) {
          add('sales', 'default', { sales: formatYen(s), purchase: formatYen(p) }, posOf(sales[0][main], sales[0].labelCell), {
            ...at(sales[0]),
            calc: `${formatYen(s)}＜${formatYen(p)}`,
          });
        }
      }

      // M-08・M-09: 特別会費
      for (const r of t.income.filter((x) => engine.labelKey(x.label) === data.specialFee && used(x))) {
        const f = parseFormula(noteOf(r), data);
        if (f.kind === 'none') add('feeFormula', 'default', {}, noteCellOf(r), at(r));
        if (f.kind === 'one' && Math.round(f.product) !== val(r[main])) {
          add(
            'feeCalc',
            'default',
            { expr: f.expr, product: formatYen(Math.round(f.product)), amount: formatYen(val(r[main])) },
            posOf(r[main], r.labelCell),
            { ...at(r), calc: `${f.operands.map(formatYen).join('×')}＝${formatYen(Math.round(f.product))}` },
          );
        }
      }

      for (const r of t.expense.filter(used)) {
        const note = noteOf(r);
        // M-10: 支出の算定根拠
        const f = parseFormula(note, data);
        if (f.kind === 'one' && Math.round(f.product) !== val(r[main])) {
          add(
            'basisCalc',
            'default',
            { expr: f.expr, product: formatYen(Math.round(f.product)), amount: formatYen(val(r[main])) },
            posOf(r[main], r.labelCell),
            { ...at(r), calc: `${f.operands.map(formatYen).join('×')}＝${formatYen(Math.round(f.product))}` },
          );
        }
        // M-11: 謝金の源泉徴収の記載
        if ((hasWord(r.label, data.honorariumWords) || hasWord(note, data.honorariumWords)) && !hasWord(note, data.withholdingWords)) {
          add('withholdingNote', 'default', {}, noteCellOf(r), at(r));
        }
        // M-12: 源泉額の計算
        // 「源泉0円（対象外）」のように0円と書いたときは、源泉徴収しない支払いとみなして判定しない
        const written = amountAfter(note, data.withholdingWords);
        if (written !== null && written > 0) {
          const fee = amountAfter(note, data.honorariumWords) ?? val(r[main]);
          const rate = Math.round(data.withholdingRatePercent * 100);
          const expected = Math.floor((fee * rate) / 10000);
          if (fee > 0 && fee <= data.withholdingLimit && expected !== written) {
            const ratePct = String(data.withholdingRatePercent);
            add(
              'withholdingCalc',
              'default',
              { fee: formatYen(fee), rate: ratePct, expected: formatYen(expected), written: formatYen(written) },
              noteCellOf(r),
              { ...at(r), calc: `${formatYen(fee)}×${ratePct}%＝${formatYen(expected)}` },
            );
          }
        }
        // M-13: 講師関係費の内訳
        if (engine.labelKey(r.label) === data.lecturerItem && !hasWord(note, data.lecturerBreakdownWords)) {
          add('lecturerBreakdown', 'default', {}, noteCellOf(r), at(r));
        }
      }
    }

    // M-14: 支出科目一覧にない項目
    for (const r of t.expense) {
      if (r.label !== '' && !data.expenseItems.includes(engine.labelKey(r.label))) {
        add('expenseItem', 'default', { item: r.label }, r.labelCell, { rowLabel: r.label, rowId: r.id });
      }
    }
    // M-15: 収入の項目名
    for (const r of t.income) {
      if (r.label !== '' && !data.incomeItems.includes(engine.labelKey(r.label))) {
        add('incomeItem', 'default', { item: r.label }, r.labelCell, { rowLabel: r.label, rowId: r.id });
      }
    }

    // M-16: 金額の桁区切りの表示（判定する列と、その列の収支計算欄。どちらも R-15・R-16 の対象から外しているため）
    /** @type {[Amount|null, string, string][]} */
    const shownCells = [
      ...t.income.map((r) => /** @type {[Amount|null, string, string]} */ ([r[main], r.label, r.id])),
      [t.incomeTotal[main], t.incomeTotal.label, t.incomeTotal.id],
      ...t.expense.map((r) => /** @type {[Amount|null, string, string]} */ ([r[main], r.label, r.id])),
      [t.expenseTotal[main], t.expenseTotal.label, t.expenseTotal.id],
      ...(t.balance ? [/** @type {[Amount|null, string, string]} */ ([t.balance[main], t.balance.label, t.balance.id])] : []),
      ...(mainSummary
        ? /** @type {[Amount|null, string, string][]} */ ([
            [mainSummary.income, `${mainSummary.label}・収入合計`, ''],
            [mainSummary.expense, `${mainSummary.label}・支出合計`, ''],
            // 残高合計の行がない表では、読み取った収支表に収支計算欄の残高が `s行番号` の行として出る（tableView）ので、その行を強調する
            [mainSummary.balance, `${mainSummary.label}・残高合計`, t.balance ? '' : `s${mainSummary.row}`],
          ])
        : []),
    ];
    for (const [a, label, id] of shownCells) {
      if (!a || !a.present || a.value === null || Math.abs(a.value) < 1000 || /[,，]/.test(a.shown)) continue;
      add('thousands', 'default', { shown: a.shown }, posOf(a, a), { rowLabel: label, rowId: id, column: columnLabel(main) });
    }

    return issues.sort((a, b) => a.rule.order - b.rule.order || a.cell.row - b.cell.row || a.cell.col - b.cell.col);
  };

  // ------------------------------------------------------------------
  // 画面に出す「読み取った収支表」
  // ------------------------------------------------------------------

  /**
   * @param {RawTable} t
   * @param {BudgetKind} kind
   * @returns {{ showSettlement: boolean, rows: ViewRow[] }}
   */
  const tableView = (t, kind) => {
    const main = mainColumn(kind);
    /** @param {Amount|null} a */
    const shown = (a) => (!a || !a.present ? '' : a.value === null ? a.shown : formatYen(a.value));
    /** @param {ItemRow} r @param {string} group @returns {ViewRow} */
    const item = (r, group) => ({
      id: r.id,
      group,
      label: r.label,
      budget: shown(r.budget),
      settlement: shown(r.settlement),
      note: main === 'settlement' ? r.settlementNote : r.budgetNote,
      total: false,
    });
    /** @param {TotalRow} r @returns {ViewRow} */
    const total = (r) => ({ id: r.id, group: '合計', label: r.label, budget: shown(r.budget), settlement: shown(r.settlement), note: '', total: true });
    const sum = main === 'budget' ? t.summary.budget : t.summary.settlement;
    /** @type {ViewRow[]} */
    const balanceRows = t.balance
      ? [total(t.balance)]
      : sum && sum.balance
        ? [
            {
              id: `s${sum.row}`,
              group: '合計',
              label: `残高合計（${sum.label}）`,
              budget: main === 'budget' ? shown(sum.balance) : '',
              settlement: main === 'settlement' ? shown(sum.balance) : '',
              note: '',
              total: true,
            },
          ]
        : [];
    return {
      showSettlement: kind === '決算',
      rows: [
        ...t.income.map((r) => item(r, '収入')),
        total(t.incomeTotal),
        ...t.expense.map((r) => item(r, '支出')),
        total(t.expenseTotal),
        ...balanceRows,
      ],
    };
  };

  root.YegAmount = Object.freeze({
    MESSAGES,
    COLUMN_LABELS,
    readSheet,
    hasSettlementAmount,
    finalize,
    checkTable,
    tableView,
    amountCells,
    parseFormula,
    parseAmountText,
    formatYen,
    fillMessage,
    stripNumber,
    colName,
  });
})(globalThis);
