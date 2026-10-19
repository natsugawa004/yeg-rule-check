// @ts-check
/*
 * 判定の流れ（画面に依存しない純粋な関数）。画面（app.js）とテストの両方から使う
 *   prepare(入力, ルール) → 文章の判定と、ファイルごとの書類の分析・収支表の読み取り（入力が変わったときだけ行う重い処理）
 *   autoSelection(準備) → 自動で判別した書類の種類と、予算か決算か
 *   evaluate(準備, 選んだ種類, ルール) → 書類全体の指摘・金額チェック・必須項目の表（種類を選び直すたびに行う軽い処理）
 *   locationOf(準備, 指摘) → 指摘の場所（複数のファイルのときは先頭にファイル名）
 * 複数のファイル（付録F-3）: ファイルごとに書類の種類を判別し、同じ種類のファイルは合わせて1つの書類として必須項目を判定する。
 * 金額チェック（付録E）は、収支予算書・決算書と判別したファイルの Excel のシートにだけ行う
 */
'use strict';

(function (root) {
  /** @typedef {'違反'|'要確認'|'参考'} Severity */
  /** @typedef {'paste'|'docx'|'xlsx'|'pdf'} Format */
  /** @typedef {'予算'|'決算'} BudgetKind */
  /** @typedef {{ type: string, [key: string]: unknown }} Loc */
  /** @typedef {{ start: number, end: number, unitId: string, loc: Loc }} Span */
  /** @typedef {{ text: string, breaks: number[], spans: Span[], unitRanges: { unitId: string, start: number, end: number }[] }} Composed */
  /** @typedef {{ id: string, label: string, sheet?: unknown, blocks: unknown[], hidden: boolean }} Unit */
  /** @typedef {{ name: string, format: 'docx'|'xlsx'|'pdf', units: Unit[], notes: string[] }} LoadedDoc */

  /**
   * @typedef {Object} Reason
   * @property {string} ruleId
   * @property {string} ruleName
   * @property {string} category
   * @property {Severity} severity
   * @property {string[]} fixes
   * @property {string} hint
   * @property {string} description
   * @property {string} source
   * @property {string} note
   * @property {number} order
   */
  /**
   * @typedef {Object} AmountMeta  金額チェックの指摘の補足
   * @property {string} location  例: 「04収支／B7（特別会費）」（複数のファイルのときは先頭にファイル名）
   * @property {string} calc      計算に使った数値
   * @property {string} tableId   読み取った収支表
   * @property {string} rowId     読み取った収支表の行
   * @property {string} fileKey
   */
  /**
   * @typedef {Object} Finding
   * @property {string} id
   * @property {'word'|'sentence'|'document'|'amount'} kind
   * @property {{ start: number, end: number }|null} range
   * @property {string} text
   * @property {string} before
   * @property {string} after
   * @property {Reason[]} reasons
   * @property {Severity} severity
   * @property {string} key
   * @property {string|null} section
   * @property {AmountMeta} [amount]
   */
  /**
   * @typedef {{ polite: number, plain: number, politeLabel: string, plainLabel: string }} StyleCounts
   * @typedef {{ ok: true, text: string, charCount: number, findings: Finding[], styleCounts: StyleCounts|null }} CheckOk
   * @typedef {{ ok: false, reason: 'empty'|'too-long', charCount: number }} CheckNg
   */

  /**
   * doccheck.js の型（使う項目だけ）
   * @typedef {{ available: boolean, format: Format, formats?: Format[], occurrences: unknown[], sections: { start: number, end: number, name: string }[],
   *   detected: string[], detectWordsFound: Record<string, string[]>, phrasesFound: string[], budgetKind: BudgetKind,
   *   budgetKindSource: 'mark'|'text'|'ambiguous'|'default'|'table' }} DocAnalysis
   * @typedef {{ types: string[], budgetKind: BudgetKind, auto: boolean }} Selection
   * @typedef {{ tables: { typeId: string, title: string, rows: { name: string, status: string, label: string, via: string|null }[] }[], total: number, satisfied: number, shortage: number, notes: string[] }} RequiredReport
   */

  /**
   * amount.js の型（使う項目だけ）
   * @typedef {{ ref: string, row: number, col: number }} CellPos
   * @typedef {{ status: 'none' } | { status: 'error', reasons: string[] } | { status: 'ok', table: unknown }} SheetRead
   * @typedef {{ rule: { id: string, name: string, category: string, severity: Severity, hint: string, description: string, source: string, order: number },
   *   message: string, cell: CellPos, rowLabel: string, rowId: string, column: string, calc: string }} Issue
   * @typedef {{ id: string, group: string, label: string, budget: string, settlement: string, note: string, total: boolean }} ViewRow
   */

  /**
   * 読み込んだファイル1つ（画面の状態から渡す）
   * @typedef {Object} SourceFile
   * @property {string} key
   * @property {string} name
   * @property {LoadedDoc} doc
   * @property {ReadonlySet<string>} included  判定に含めるページ・シート
   */
  /** @typedef {{ kind: 'paste', text: string } | { kind: 'files', files: readonly SourceFile[] }} PipelineInput */

  /**
   * 判定の単位（貼り付けた文章1つ、またはファイル1つ）
   * @typedef {Object} Part
   * @property {string} key
   * @property {string} name     ファイル名（貼り付けは ''）
   * @property {Format} format
   * @property {number} start    つないだ文章の中での位置
   * @property {number} end
   * @property {Composed|null} composed  このファイルだけの組み立て（位置はこのファイルの中）
   * @property {DocAnalysis} analysis
   * @property {{ unitId: string, sheet: string, read: SheetRead }[]} sheets  Excel のシートの収支表の読み取り
   */
  /**
   * @typedef {Object} Prepared
   * @property {true} ok
   * @property {'paste'|'files'} mode
   * @property {string} text
   * @property {Composed|null} composed  つないだもの（貼り付けは null）
   * @property {CheckOk} result
   * @property {Part[]} parts
   * @property {boolean} multi           2つ以上のファイル
   */

  /**
   * @typedef {Object} AmountSheetView
   * @property {string} id         読み取った収支表の目印
   * @property {string} fileKey
   * @property {string} fileName
   * @property {string} sheet
   * @property {boolean} ok
   * @property {string[]} reasons  読み取れなかった理由
   * @property {{ showSettlement: boolean, rows: ViewRow[] }|null} table
   * @property {number} findingCount
   */
  /**
   * @typedef {Object} AmountView
   * @property {'none'|'unavailable'|'needs-excel'|'checked'} status
   * @property {BudgetKind} kind
   * @property {AmountSheetView[]} sheets
   * @property {string[]} checkedRules  判定したルール（M-xx。収支表を読み取れたときだけ）
   */
  /**
   * @typedef {Object} Evaluation
   * @property {Finding[]} textFindings      文章の指摘（金額の列のセルの R-15・R-16 を除いたもの）
   * @property {Finding[]} documentFindings  書類全体の指摘
   * @property {Finding[]} amountFindings    金額チェックの指摘
   * @property {AmountView} amount
   * @property {RequiredReport} report
   * @property {{ typeId: string, files: string[] }[]} groups  種類ごとに使ったファイル
   * @property {DocAnalysis} analysis        すべてを合わせた分析
   * @property {StructureView|null} structure  委員会事業計画書の委員会名と字数（その種類を選んでいないときは null）
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (root));
  const engine = /** @type {{ MAX_CHARS: number, countChars: (s: string) => number,
   *   checkText: (t: string, rs: unknown, o?: { breaks?: readonly number[], sections?: readonly { start: number, end: number, name: string }[], documentTypes?: readonly string[] }) => CheckOk|CheckNg,
   *   checkDocumentRules: (t: string, rs: unknown, typeIds: readonly string[]) => Finding[],
   *   withDynamicWords: (rs: Ruleset, kind: string, words: readonly string[]) => Ruleset }} */ (g.YegEngine);
  const docCheck = /** @type {{ analyzeDocument: (input: { format: Format, text: string, composed: Composed|null, units: readonly Unit[]|null }, rs: unknown) => DocAnalysis,
   *   mergeAnalyses: (list: readonly DocAnalysis[]) => DocAnalysis,
   *   requiredReport: (a: DocAnalysis, s: Selection, rs: unknown, byType?: Readonly<Record<string, DocAnalysis>>) => RequiredReport }} */ (g.YegDocCheck);
  const fileApi = /** @type {{ composeFiles: (files: readonly { key: string, units: readonly Unit[], included: ReadonlySet<string> }[]) => { composed: Composed, parts: { key: string, start: number, end: number, composed: Composed }[] },
   *   locate: (c: Composed, offset: number) => Span|null, locationLabel: (loc: Loc) => string }} */ (g.YegFile);
  const annualApi = /** @type {{ committeeNames: (text: string, type: unknown) => string[],
   *   adjustFindings: (f: readonly Finding[], regions: readonly { start: number, end: number, typeIds: string[], hasScope: boolean }[], rs: unknown) => Finding[],
   *   bodyParagraphs: (text: string, sections: readonly { start: number, end: number, name: string }[], name: string) => { start: number, end: number, text: string }[],
   *   layout: (paragraphs: readonly { text: string }[], params: Readonly<Record<string, unknown>>) => Layout|null,
   *   structureFindings: (paragraphs: readonly { text: string }[], typeId: string, rs: unknown) => Finding[] }} */ (g.YegAnnual);
  const amountApi = /** @type {{ MESSAGES: Readonly<Record<string, unknown>>, readSheet: (sheet: unknown, data: unknown) => SheetRead,
   *   hasSettlementAmount: (t: unknown) => boolean, finalize: (t: unknown, kind: BudgetKind) => { ok: true, table: unknown } | { ok: false, reasons: string[] },
   *   checkTable: (t: unknown, kind: BudgetKind, rules: readonly unknown[], data: unknown) => Issue[],
   *   tableView: (t: unknown, kind: BudgetKind) => { showSettlement: boolean, rows: ViewRow[] }, amountCells: (t: unknown) => CellPos[] }} */ (g.YegAmount);

  /**
   * ルール定義（engine.js の CompiledRuleset。使う項目だけ）
   * @typedef {{ id: string, name: string, scopeSections: string[], committeeNamePattern: RegExp|null }} DocTypeLike
   * @typedef {{ documents: { types: DocTypeLike[] }|null, amounts: { documentType: string, replaces: { documentRules: string[], cellRules: string[] } }|null,
   *   rules: { id: string, type: string, order: number, check: string, documentTypes: string[], dynamicWords: string, params: Readonly<Record<string, unknown>> }[] }} Ruleset
   */
  /** @typedef {{ label: string, chars: number, target: number, status: 'ok'|'short'|'long' }} LayoutRow */
  /** @typedef {{ split: boolean, rows: LayoutRow[], total: LayoutRow }} Layout */
  /**
   * 委員会事業計画書の構成（付録G）
   * @typedef {{ typeId: string, typeName: string, committeeName: string, layout: Layout|null }} StructureView
   */

  const MESSAGES = Object.freeze({
    needsExcel: '金額チェックはExcel（.xlsx）で読み込むと行えます',
    unavailable: 'ルール定義に金額チェックのデータがないため、金額チェックは行いません',
    noTable: '「収入」「支出」の見出しのある収支表が見つかりません',
    /** @param {readonly string[]} reasons */
    notRead: (reasons) => `収支表の形を読み取れなかったため、金額チェックは行っていません（理由: ${reasons.join('／')}）`,
    noProblem: '金額チェック: 問題は見つかりませんでした',
  });

  /** @param {unknown} v @returns {boolean} */
  const isSheet = (v) =>
    typeof v === 'object' && v !== null && Array.isArray(/** @type {{ cells?: unknown }} */ (v).cells) && Array.isArray(/** @type {{ merges?: unknown }} */ (v).merges);

  /**
   * 文章を判定し、ファイルごとに書類を分析する
   * @param {PipelineInput} input
   * @param {Ruleset} ruleset
   * @returns {Prepared|CheckNg}
   */
  const prepare = (input, ruleset) => {
    if (input.kind === 'paste') {
      const analysis = docCheck.analyzeDocument({ format: 'paste', text: input.text, composed: null, units: null }, ruleset);
      const result = engine.checkText(input.text, ruleset, { sections: analysis.sections });
      if (!result.ok) return result;
      return {
        ok: true,
        mode: 'paste',
        text: result.text,
        composed: null,
        result,
        parts: [{ key: 'paste', name: '', format: 'paste', start: 0, end: result.text.length, composed: null, analysis, sheets: [] }],
        multi: false,
      };
    }
    const files = input.files;
    const joined = fileApi.composeFiles(files.map((f) => ({ key: f.key, units: f.doc.units, included: f.included })));
    const amounts = ruleset.amounts;
    /** @type {Part[]} */
    const parts = files.map((f, i) => {
      const cp = joined.parts[i];
      const analysis = docCheck.analyzeDocument({ format: f.doc.format, text: cp.composed.text, composed: cp.composed, units: f.doc.units }, ruleset);
      const sheets =
        f.doc.format === 'xlsx' && amounts
          ? f.doc.units
              .filter((u) => f.included.has(u.id) && isSheet(u.sheet))
              .map((u) => ({ unitId: u.id, sheet: u.label, read: amountApi.readSheet(u.sheet, amounts) }))
          : [];
      return { key: f.key, name: f.name, format: f.doc.format, start: cp.start, end: cp.end, composed: cp.composed, analysis, sheets };
    });
    if (files.length <= 1) {
      const sections = parts.flatMap((p) => p.analysis.sections);
      const result = engine.checkText(joined.composed.text, ruleset, { breaks: joined.composed.breaks, sections });
      if (!result.ok) return result;
      return { ok: true, mode: 'files', text: result.text, composed: joined.composed, result, parts, multi: false };
    }
    // 複数のファイル: 文章の判定はファイルごとに行う（文体の混在などは、ファイルの中だけで見る）。位置はつないだ文章に直す
    const text = joined.composed.text;
    const charCount = engine.countChars(text);
    if (text.trim() === '') return { ok: false, reason: 'empty', charCount };
    if (charCount > engine.MAX_CHARS) return { ok: false, reason: 'too-long', charCount };
    /** @type {Finding[]} */
    const findings = [];
    /** @type {StyleCounts|null} */
    let styleCounts = null;
    for (const p of parts) {
      if (!p.composed || p.composed.text.trim() === '') continue;
      const r = engine.checkText(p.composed.text, ruleset, { breaks: p.composed.breaks, sections: p.analysis.sections });
      if (!r.ok) continue;
      findings.push(...r.findings.map((f) => ({ ...f, range: f.range ? { start: f.range.start + p.start, end: f.range.end + p.start } : null })));
      const c = r.styleCounts;
      if (c) styleCounts = styleCounts ? { ...styleCounts, polite: styleCounts.polite + c.polite, plain: styleCounts.plain + c.plain } : c;
    }
    const result = { ok: /** @type {true} */ (true), text, charCount, findings: findings.map((f, i) => ({ ...f, id: `f-${i + 1}` })), styleCounts };
    return { ok: true, mode: 'files', text, composed: joined.composed, result, parts, multi: true };
  };

  /**
   * その種類の書類として判定に使うファイル（その種類と判別したファイル。どのファイルも判別していなければ全部）
   * @param {Prepared} prep
   * @param {string} typeId
   */
  const partsFor = (prep, typeId) => {
    const detected = prep.parts.filter((p) => p.analysis.detected.includes(typeId));
    return detected.length > 0 ? detected : prep.parts;
  };

  /** @param {Ruleset} ruleset */
  const budgetTypeOf = (ruleset) => (ruleset.amounts ? ruleset.amounts.documentType : 'budget');

  /**
   * 予算か決算か（印 → 「予算（案）」「決算（案）」の文字 → どちらもなければ収支表の決算額の列。付録B-1・F-2）
   * @param {readonly Part[]} parts
   * @returns {{ kind: BudgetKind, source: DocAnalysis['budgetKindSource'] }}
   */
  const budgetKindOf = (parts) => {
    const marked = parts.find((p) => p.analysis.budgetKindSource === 'mark');
    if (marked) return { kind: marked.analysis.budgetKind, source: 'mark' };
    const texts = parts.filter((p) => p.analysis.budgetKindSource === 'text');
    if (texts.length > 0 && texts.every((p) => p.analysis.budgetKind === texts[0].analysis.budgetKind)) {
      return { kind: texts[0].analysis.budgetKind, source: 'text' };
    }
    const ambiguous = texts.length > 0 || parts.some((p) => p.analysis.budgetKindSource === 'ambiguous');
    const tables = parts.flatMap((p) => p.sheets).flatMap((s) => (s.read.status === 'ok' ? [s.read.table] : []));
    if (tables.length > 0) return { kind: tables.some((t) => amountApi.hasSettlementAmount(t)) ? '決算' : '予算', source: 'table' };
    return { kind: '予算', source: ambiguous ? 'ambiguous' : 'default' };
  };

  /**
   * 自動で判別した書類の種類
   * @param {Prepared} prep
   * @param {Ruleset} ruleset
   * @returns {Selection}
   */
  const autoSelection = (prep, ruleset) => {
    const docs = ruleset.documents;
    const types = docs ? docs.types.map((t) => t.id).filter((id) => prep.parts.some((p) => p.analysis.detected.includes(id))) : [];
    return { types, budgetKind: budgetKindOf(partsFor(prep, budgetTypeOf(ruleset))).kind, auto: true };
  };

  /**
   * ファイルの中の位置（セル）を、つないだ文章の中の範囲にする
   * @param {Part} part
   * @param {string} unitId
   * @param {CellPos} cell
   * @returns {{ start: number, end: number }|null}
   */
  const cellRange = (part, unitId, cell) => {
    if (!part.composed) return null;
    const s = part.composed.spans.find(
      (x) => x.unitId === unitId && x.loc.type === 'cell' && x.loc.row === cell.row && x.loc.col === cell.col,
    );
    return s ? { start: part.start + s.start, end: part.start + s.end } : null;
  };

  /**
   * 金額チェック（収支予算書・決算書と判別したファイルの Excel のシートだけ）
   * @param {Prepared} prep
   * @param {readonly Part[]} parts
   * @param {BudgetKind} kind
   * @param {Ruleset} ruleset
   * @returns {{ view: AmountView, findings: Finding[], cellRanges: { start: number, end: number }[] }}
   */
  const checkAmounts = (prep, parts, kind, ruleset) => {
    const data = ruleset.amounts;
    if (!data) return { view: { status: 'unavailable', kind, sheets: [], checkedRules: [] }, findings: [], cellRanges: [] };
    const xlsx = parts.filter((p) => p.format === 'xlsx');
    if (xlsx.length === 0) return { view: { status: 'needs-excel', kind, sheets: [], checkedRules: [] }, findings: [], cellRanges: [] };
    const rules = ruleset.rules.filter((r) => r.type === 'amount');
    /** @type {AmountSheetView[]} */
    const sheets = [];
    /** @type {Finding[]} */
    const findings = [];
    /** @type {{ start: number, end: number }[]} */
    const cellRanges = [];
    for (const part of xlsx) {
      const candidates = part.sheets.filter((s) => s.read.status !== 'none');
      if (candidates.length === 0) {
        sheets.push({ id: `${part.key}/`, fileKey: part.key, fileName: part.name, sheet: '', ok: false, reasons: [MESSAGES.noTable], table: null, findingCount: 0 });
        continue;
      }
      for (const s of candidates) {
        const id = `${part.key}/${s.unitId}`;
        const base = { id, fileKey: part.key, fileName: part.name, sheet: s.sheet };
        if (s.read.status === 'error') {
          sheets.push({ ...base, ok: false, reasons: s.read.reasons, table: null, findingCount: 0 });
          continue;
        }
        if (s.read.status !== 'ok') continue;
        const fin = amountApi.finalize(s.read.table, kind);
        if (!fin.ok) {
          sheets.push({ ...base, ok: false, reasons: fin.reasons, table: null, findingCount: 0 });
          continue;
        }
        cellRanges.push(...amountApi.amountCells(fin.table).flatMap((c) => {
          const r = cellRange(part, s.unitId, c);
          return r ? [r] : [];
        }));
        const issues = amountApi.checkTable(fin.table, kind, rules, data);
        for (const issue of issues) {
          const range = cellRange(part, s.unitId, issue.cell);
          const where = [issue.rowLabel, issue.column].filter((t) => t !== '').join('・');
          const location = `${prep.multi ? `${part.name}／` : ''}${s.sheet}／${issue.cell.ref}${where ? `（${where}）` : ''}`;
          const r = issue.rule;
          findings.push({
            id: '',
            kind: 'amount',
            range,
            text: range ? prep.text.slice(range.start, range.end) : '',
            before: '',
            after: '',
            reasons: [
              {
                ruleId: r.id,
                ruleName: r.name,
                category: r.category,
                severity: r.severity,
                fixes: [],
                hint: r.hint,
                description: r.description,
                source: r.source,
                note: issue.message,
                order: r.order,
              },
            ],
            severity: r.severity,
            key: [r.id, part.name, s.sheet, issue.cell.ref, issue.message].join('\u0000'),
            section: null,
            amount: { location, calc: issue.calc, tableId: id, rowId: issue.rowId, fileKey: part.key },
          });
        }
        sheets.push({ ...base, ok: true, reasons: [], table: amountApi.tableView(fin.table, kind), findingCount: issues.length });
      }
    }
    const anyOk = sheets.some((s) => s.ok);
    return {
      view: { status: 'checked', kind, sheets, checkedRules: anyOk ? rules.map((r) => r.id) : [] },
      findings: findings.map((f, i) => ({ ...f, id: `m-${i + 1}` })),
      cellRanges,
    };
  };

  /**
   * 金額の列のセルにある指摘から、金額チェックに置き換えるルール（R-15・R-16）を除く
   * @param {readonly Finding[]} findings
   * @param {readonly { start: number, end: number }[]} ranges
   * @param {ReadonlySet<string>} ruleIds
   * @returns {Finding[]}
   */
  const withoutReplaced = (findings, ranges, ruleIds) => {
    if (ranges.length === 0 || ruleIds.size === 0) return [...findings];
    return findings.flatMap((f) => {
      const range = f.range;
      if (!range || !ranges.some((r) => range.start >= r.start && range.start < r.end)) return [f];
      const reasons = f.reasons.filter((r) => !ruleIds.has(r.ruleId));
      if (reasons.length === f.reasons.length) return [f];
      if (reasons.length === 0) return [];
      return [{ ...f, reasons, severity: reasons[0].severity, key: [reasons.map((r) => r.ruleId).join('+'), f.text, f.before, f.after].join('\u0000') }];
    });
  };

  /**
   * 選んだ書類の種類に応じた判定
   * @param {Prepared} prep
   * @param {Selection} selection
   * @param {Ruleset} ruleset
   * @returns {Evaluation}
   */
  const evaluate = (prep, selection, ruleset) => {
    const budgetType = budgetTypeOf(ruleset);
    const groups = selection.types.map((typeId) => ({ typeId, parts: partsFor(prep, typeId) }));

    // 金額チェック
    const budgetGroup = groups.find((x) => x.typeId === budgetType);
    const amount = budgetGroup
      ? checkAmounts(prep, budgetGroup.parts, selection.budgetKind, ruleset)
      : { view: /** @type {AmountView} */ ({ status: 'none', kind: selection.budgetKind, sheets: [], checkedRules: [] }), findings: [], cellRanges: [] };
    const tableRead = amount.view.sheets.some((s) => s.ok);
    const replaces = ruleset.amounts && tableRead ? ruleset.amounts.replaces : { documentRules: [], cellRules: [] };

    // 書類全体の指摘（種類ごとに、その種類のファイルの文章で判定する）
    const replacedDocs = new Set(replaces.documentRules);
    /** @type {Map<string, Finding>} */
    const byRule = new Map();
    for (const grp of groups) {
      const text = grp.parts.map((p) => prep.text.slice(p.start, p.end)).join('\n');
      for (const f of engine.checkDocumentRules(text, ruleset, [grp.typeId])) {
        const id = f.reasons[0].ruleId;
        if (!replacedDocs.has(id) && !byRule.has(id)) byRule.set(id, f);
      }
    }
    // 書類の種類に応じた文章ルールの扱い（付録G。委員会事業計画書は判定する欄を絞り、議案のルールを外す・下げる）
    const typeById = new Map((ruleset.documents ? ruleset.documents.types : []).map((t) => [t.id, t]));
    const regions = prep.parts.map((p) => {
      const typeIds = groups.filter((grp) => grp.parts.includes(p)).map((grp) => grp.typeId);
      const scope = typeIds.flatMap((id) => typeById.get(id)?.scopeSections ?? []);
      return { start: p.start, end: p.end, typeIds, hasScope: scope.length > 0 && p.analysis.sections.some((s) => scope.includes(s.name)) };
    });
    // 書類の種類を決めているルール（委員会事業計画書の用語 Y-01〜Y-05 など）は、その種類の範囲だけで判定する。
    // 見出しから読み取った自分の委員会名（Y-03。入力から語を読み取るルール）もここで判定する
    /** @type {Finding[]} */
    const typedFindings = [];
    /** @type {Map<Part, string[]>} */
    const namesByPart = new Map();
    prep.parts.forEach((p, i) => {
      const typeIds = regions[i].typeIds;
      if (typeIds.length === 0) return;
      const forTypes = (/** @type {{ documentTypes: string[] }} */ r) => r.documentTypes.some((id) => typeIds.includes(id));
      const partText = prep.text.slice(p.start, p.end);
      const t = typeIds.map((id) => typeById.get(id)).find((x) => x && x.committeeNamePattern);
      const names = t ? annualApi.committeeNames(partText, t) : [];
      if (t) namesByPart.set(p, names);
      const rules = [
        ...ruleset.rules.filter((r) => r.type === 'pattern' && r.dynamicWords === '' && forTypes(r)),
        ...engine.withDynamicWords(ruleset, 'committeeName', names).rules.filter(forTypes),
      ];
      if (rules.length === 0) return;
      const r = engine.checkText(
        partText,
        { ...ruleset, rules },
        { breaks: p.composed ? p.composed.breaks : [], sections: p.analysis.sections, documentTypes: typeIds },
      );
      if (!r.ok) return;
      for (const f of r.findings) {
        typedFindings.push({ ...f, id: `y-${typedFindings.length + 1}`, range: f.range ? { start: f.range.start + p.start, end: f.range.end + p.start } : null });
      }
    });

    // 構成のチェック（字数・書き出し・締め）と、字数の表
    /** @type {Finding[]} */
    const structureFindings = [];
    /** @type {StructureView|null} */
    let structure = null;
    for (const grp of groups) {
      const t = typeById.get(grp.typeId);
      const checks = ruleset.rules.filter((r) => r.type === 'document' && r.check !== '' && r.documentTypes.includes(grp.typeId));
      if (!t || checks.length === 0) continue;
      const lengthRule = checks.find((r) => r.check === 'length');
      const sectionName =
        lengthRule && typeof lengthRule.params.section === 'string' ? lengthRule.params.section : (t.scopeSections[0] ?? '');
      const part = grp.parts.find((p) => p.analysis.sections.some((s) => s.name === sectionName)) ?? grp.parts[0];
      const paragraphs = annualApi.bodyParagraphs(prep.text.slice(part.start, part.end), part.analysis.sections, sectionName);
      structureFindings.push(...annualApi.structureFindings(paragraphs, grp.typeId, ruleset));
      structure = structure ?? {
        typeId: grp.typeId,
        typeName: t.name,
        committeeName: (namesByPart.get(part) ?? [])[0] ?? '',
        layout: lengthRule ? annualApi.layout(paragraphs, lengthRule.params) : null,
      };
    }

    const documentFindings = [...[...byRule.values()].sort((a, b) => a.reasons[0].order - b.reasons[0].order), ...structureFindings].map(
      (f, i) => ({ ...f, id: `d-${i + 1}` }),
    );

    // 必須項目（同じ種類のファイルは合わせて1つの書類として判定する）
    /** @type {Record<string, DocAnalysis>} */
    const byType = {};
    for (const grp of groups) {
      const merged = docCheck.mergeAnalyses(grp.parts.map((p) => p.analysis));
      if (grp.typeId === budgetType) {
        const bk = budgetKindOf(grp.parts);
        byType[grp.typeId] = { ...merged, budgetKind: bk.kind, budgetKindSource: bk.source };
      } else {
        byType[grp.typeId] = merged;
      }
    }
    const analysis = docCheck.mergeAnalyses(prep.parts.map((p) => p.analysis));

    const textFindings = annualApi
      .adjustFindings(
        [...withoutReplaced(prep.result.findings, amount.cellRanges, new Set(replaces.cellRules)), ...typedFindings],
        regions,
        ruleset,
      )
      .sort((a, b) => (a.range ? a.range.start : -1) - (b.range ? b.range.start : -1));

    return {
      textFindings,
      documentFindings,
      amountFindings: amount.findings,
      amount: amount.view,
      report: docCheck.requiredReport(analysis, selection, ruleset, byType),
      groups: groups.map((grp) => ({ typeId: grp.typeId, files: grp.parts.map((p) => p.name).filter((n) => n !== '') })),
      analysis,
      structure,
    };
  };

  /**
   * 指摘の場所（例: 「01計画／B4」、複数のファイルのときは「基本.xlsx／01計画／B4」）。貼り付けた文章では空文字
   * @param {Prepared} prep
   * @param {Finding} f
   */
  const locationOf = (prep, f) => {
    if (f.kind === 'amount') return f.amount ? f.amount.location : '';
    if (!prep.composed || !f.range) return '';
    const start = f.range.start;
    const span = fileApi.locate(prep.composed, start);
    const label = span ? fileApi.locationLabel(span.loc) : '';
    if (!prep.multi) return label;
    const part = prep.parts.find((p) => start >= p.start && start < p.end);
    if (!part) return label;
    return label ? `${part.name}／${label}` : part.name;
  };

  /**
   * ファイルごとの指摘の件数（複数のファイルの一覧に出す）
   * @param {Prepared} prep
   * @param {readonly Finding[]} findings
   * @returns {Map<string, number>}
   */
  const countByFile = (prep, findings) => {
    /** @type {Map<string, number>} */
    const counts = new Map();
    for (const f of findings) {
      const range = f.range;
      const key = f.amount ? f.amount.fileKey : range ? prep.parts.find((p) => range.start >= p.start && range.start < p.end)?.key : undefined;
      if (key !== undefined) counts.set(key, (counts.get(key) ?? 0) + 1);
    }
    return counts;
  };

  root.YegPipeline = Object.freeze({
    MESSAGES,
    prepare,
    autoSelection,
    evaluate,
    locationOf,
    countByFile,
    budgetKindOf,
  });
})(globalThis);
