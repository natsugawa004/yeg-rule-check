// @ts-check
/*
 * 画面の操作（DOM）。判定そのものは engine.js、結果の整理は report.js が行う
 * 入力した文章・判定結果・無視した指摘は、この画面のメモリの中にだけ置き、端末に保存しない
 */
'use strict';

(function () {
  /** @typedef {'違反'|'要確認'|'参考'} Severity */
  /** @typedef {Severity|'all'} SeverityFilter */
  /** @typedef {'list'|'body'|'doc'} Tab */

  /**
   * @typedef {Object} Finding
   * @property {string} id
   * @property {'word'|'sentence'|'document'|'amount'} kind
   * @property {{ start: number, end: number }|null} range
   * @property {string} text
   * @property {string} before
   * @property {string} after
   * @property {{ ruleId: string, ruleName: string, severity: Severity, fixes: string[], hint: string, description: string, source: string, note: string, order: number }[]} reasons
   * @property {Severity} severity
   * @property {string} key
   * @property {string|null} [section]  指摘のある欄の名前
   * @property {{ location: string, calc: string, tableId: string, rowId: string, fileKey: string }} [amount]  金額チェックの指摘
   */

  /**
   * @typedef {Object} OkResult
   * @property {true} ok
   * @property {string} text
   * @property {number} charCount
   * @property {Finding[]} findings
   * @property {{ polite: number, plain: number, politeLabel: string, plainLabel: string }|null} styleCounts
   */

  /**
   * 読み込んだファイル（fileload.js の形）
   * @typedef {{ type: string }} Loc
   * @typedef {{ text: string, sep: string, loc: Loc }} Block
   * @typedef {{ id: string, kind: 'page'|'sheet'|'document', label: string, hidden: boolean, hasText: boolean, blocks: Block[] }} Unit
   * @typedef {{ name: string, format: 'docx'|'xlsx'|'pdf', units: Unit[], notes: string[] }} LoadedDoc
   * @typedef {{ start: number, end: number, unitId: string, loc: Loc }} Span
   * @typedef {{ text: string, breaks: number[], spans: Span[], unitRanges: { unitId: string, start: number, end: number }[] }} Composed
   * @typedef {{ code: string, message: string, detail: string }} LoadError
   */

  /**
   * 読み込んだファイル（複数のファイルを同時に読み込める。付録F-3）
   * @typedef {Object} FileEntry
   * @property {string} key
   * @property {LoadedDoc} doc
   * @property {ReadonlySet<string>} included  判定に含めるページ・シート
   */
  /** @typedef {{ key: string, name: string, error: LoadError }} FailedFile  読み込めなかったファイル */

  /**
   * @typedef {Object} State
   * @property {OkResult|null} result
   * @property {string} checkedText
   * @property {ReadonlySet<string>} ignored
   * @property {'rule'|'position'} groupMode
   * @property {SeverityFilter} filter
   * @property {boolean} showIgnored
   * @property {Tab} tab
   * @property {readonly FileEntry[]} files    読み込んだファイル（貼り付けのときは空）
   * @property {readonly FailedFile[]} failed  読み込めなかったファイル（一覧に理由を出す）
   * @property {Prepared|null} prepared        判定の準備（文章の判定・ファイルごとの分析・収支表の読み取り）
   * @property {DocState|null} doc             書類の種類・必須項目・金額チェック（判定した後）
   * @property {ReadonlySet<string>} checked   確認ポイントで付けたチェック（保存しない）
   */

  /**
   * doccheck.js・pipeline.js の結果（使う項目だけ）
   * @typedef {{ types: string[], budgetKind: '予算'|'決算', auto: boolean }} Selection
   * @typedef {{ name: string, status: 'filled'|'blank'|'present'|'missing', label: string, via: string|null }} RequiredRow
   * @typedef {{ typeId: string, title: string, rows: RequiredRow[] }} RequiredTable
   * @typedef {{ tables: RequiredTable[], total: number, satisfied: number, shortage: number, notes: string[] }} RequiredReport
   * @typedef {{ available: boolean, format: 'paste'|'docx'|'xlsx'|'pdf', sections: { start: number, end: number, name: string }[], detected: string[], budgetKind: '予算'|'決算', budgetKindSource: string }} DocAnalysis
   * @typedef {{ minDetectWords: number, common: { id: string, text: string }[], types: { id: string, name: string, detectWords: string[], required: string[], settlementRequired: string[], blankCheck: string[], checkpoints: { id: string, text: string }[] }[] }} DocumentsModel
   * @typedef {{ key: string, name: string, format: 'paste'|'docx'|'xlsx'|'pdf', start: number, end: number, analysis: DocAnalysis }} Part
   * @typedef {{ ok: true, mode: 'paste'|'files', text: string, composed: Composed|null, result: OkResult, parts: Part[], multi: boolean }} Prepared
   * @typedef {{ id: string, group: string, label: string, budget: string, settlement: string, note: string, total: boolean }} AmountRow
   * @typedef {{ id: string, fileKey: string, fileName: string, sheet: string, ok: boolean, reasons: string[], table: { showSettlement: boolean, rows: AmountRow[] }|null, findingCount: number }} AmountSheet
   * @typedef {{ status: 'none'|'unavailable'|'needs-excel'|'checked', kind: '予算'|'決算', sheets: AmountSheet[], checkedRules: string[] }} AmountView
   * @typedef {{ textFindings: Finding[], documentFindings: Finding[], amountFindings: Finding[], amount: AmountView, report: RequiredReport, groups: { typeId: string, files: string[] }[], analysis: DocAnalysis }} Evaluation
   */

  /**
   * @typedef {Object} DocState
   * @property {Selection} selection
   * @property {Evaluation} evaluation  選んだ種類に応じた判定（書類全体の指摘・金額チェック・必須項目）
   */

  /**
   * @typedef {Object} DocCheckApi
   * @property {Readonly<Record<string, string>>} MESSAGES
   * @property {(sel: Selection, ruleset: Ruleset) => string[]} typeNames
   * @property {(rep: RequiredReport) => string} requiredSummary
   * @property {(sel: Selection, ruleset: Ruleset, checkedRules?: ReadonlySet<string>) => { title: string, items: { id: string, text: string, checkedBy: string[] }[] }[]} checkpointGroups
   * @property {(sel: Selection, rep: RequiredReport, ruleset: Ruleset) => string[]} copyLines
   * @property {(ruleset: Ruleset) => DocumentsModel|null} documentsModel
   */

  /**
   * @typedef {Object} PipelineApi
   * @property {{ needsExcel: string, unavailable: string, noProblem: string, notRead: (reasons: readonly string[]) => string }} MESSAGES
   * @property {(input: { kind: 'paste', text: string } | { kind: 'files', files: readonly { key: string, name: string, doc: LoadedDoc, included: ReadonlySet<string> }[] }, ruleset: Ruleset) => Prepared | { ok: false, reason: 'empty'|'too-long', charCount: number }} prepare
   * @property {(prep: Prepared, ruleset: Ruleset) => Selection} autoSelection
   * @property {(prep: Prepared, sel: Selection, ruleset: Ruleset) => Evaluation} evaluate
   * @property {(prep: Prepared, f: Finding) => string} locationOf
   * @property {(prep: Prepared, findings: readonly Finding[]) => Map<string, number>} countByFile
   */

  /**
   * @typedef {Object} Ruleset
   * @property {{ toolName: string, manualName: string, manualVersion: string, sourceLabel: string, updated: string }} meta
   * @property {{ ruleLabel: string, message: string }[]} errors
   * @property {string|null} fatal
   * @property {{ caution: string, request: string, links: { label: string, url: string }[] }} aiGuide
   * @property {{ types: { id: string, name: string }[] }|null} documents
   */

  /**
   * @typedef {Object} EngineApi
   * @property {number} MAX_CHARS
   * @property {(raw: unknown) => Ruleset} compileRuleset
   * @property {(text: string, ruleset: Ruleset, options?: { breaks?: readonly number[], sections?: readonly { start: number, end: number, name: string }[] }) => OkResult | { ok: false, reason: 'empty'|'too-long', charCount: number }} checkText
   * @property {(text: string, ruleset: Ruleset, typeIds: readonly string[]) => Finding[]} checkDocumentRules
   * @property {(s: string) => number} countChars
   */

  /**
   * @typedef {Object} FileApi
   * @property {Readonly<Record<'docx'|'xlsx'|'pdf', string>>} FORMAT_LABELS
   * @property {number} MAX_FILES
   * @property {string} TOO_MANY_FILES
   * @property {new (code: string, detail?: string) => Error} FileLoadError
   * @property {(size: number) => LoadError|null} checkSize
   * @property {(file: { name: string, bytes: Uint8Array }, deps: { readPdf: (bytes: Uint8Array) => Promise<string[]> }) => Promise<{ ok: true, doc: LoadedDoc } | { ok: false, error: LoadError }>} loadFile
   * @property {(units: readonly Unit[], included: ReadonlySet<string>) => Composed} compose
   * @property {(files: readonly { key: string, units: readonly Unit[], included: ReadonlySet<string> }[]) => { composed: Composed, parts: { key: string, start: number, end: number, composed: Composed }[] }} composeFiles
   * @property {(fileKey: string, unitId: string, multi: boolean) => string} unitKeyOf
   * @property {(composed: Composed, offset: number) => Span|null} locate
   * @property {(loc: Loc) => string} locationLabel
   * @property {(composed: Composed, findings: Finding[]) => Map<string, number>} countByUnit
   * @property {(units: readonly Unit[]) => Set<string>} defaultIncluded
   * @property {(u: Unit) => string} unitLabel
   */

  /**
   * @typedef {Object} PdfApi
   * @property {(bytes: Uint8Array, deps: { lib: { getDocument: (src: Record<string, unknown>) => unknown }, getCmaps: () => Readonly<Record<string, string>>|null, onProgress: (done: number, total: number) => void }) => Promise<string[]>} readPdf
   */

  /**
   * @typedef {Object} RuleListModel
   * @property {number} total
   * @property {{ name: string, note: string, rules: { id: string, name: string, severity: Severity, examples: string[], fixes: string[], description: string, source: string, scope: string }[] }[]} categories
   * @property {string[]} notChecked
   */

  /**
   * @typedef {Object} ReportApi
   * @property {readonly Severity[]} SEVERITIES
   * @property {string} WHOLE_DOCUMENT
   * @property {(counts: Record<Severity, number>) => string} formatCounts
   * @property {(s: Severity) => string} severityClass
   * @property {(f: Finding[], ignored: ReadonlySet<string>) => Finding[]} visibleFindings
   * @property {(f: Finding[], ignored: ReadonlySet<string>) => Finding[]} ignoredFindings
   * @property {(f: Finding[]) => Record<Severity, number>} countBySeverity
   * @property {(f: Finding[], filter: SeverityFilter) => Finding[]} filterBySeverity
   * @property {(f: Finding[], mode: 'rule'|'position') => { key: string, title: string|null, severity: Severity|null, findings: Finding[] }[]} groupFindings
   * @property {(fixes: string[]) => string} formatFix
   * @property {(ruleset: Ruleset, source: string) => string} formatSource
   * @property {(c: { polite: number, plain: number, politeLabel: string, plainLabel: string }) => string} formatStyleCounts
   * @property {(f: Finding[]) => boolean} hasStyleFinding
   * @property {(s: string) => string} oneLine
   * @property {(s: string) => string} contextLine
   * @property {(args: { findings: Finding[], ignoredCount: number, charCount: number, styleCounts: OkResult['styleCounts'], ruleset: Ruleset, now: Date, source?: { fileName: string, scope: string }, locate?: (f: Finding) => string, docLines?: string[] }) => string} buildCopyText
   * @property {(ruleset: Ruleset) => string} buildAiPrompt
   * @property {(ruleset: Ruleset) => RuleListModel} buildRuleListModel
   */

  const g = /** @type {Record<string, unknown>} */ (/** @type {unknown} */ (globalThis));
  const engine = /** @type {EngineApi|undefined} */ (g.YegEngine);
  const report = /** @type {ReportApi|undefined} */ (g.YegReport);
  const fileApi = /** @type {FileApi|undefined} */ (g.YegFile);
  const pdfApi = /** @type {PdfApi|undefined} */ (g.YegPdf);
  const docCheck = /** @type {DocCheckApi|undefined} */ (g.YegDocCheck);
  const pipeline = /** @type {PipelineApi|undefined} */ (g.YegPipeline);

  /**
   * @param {string} id
   * @returns {HTMLElement}
   */
  const el = (id) => {
    const node = document.getElementById(id);
    if (!node) throw new Error(`画面の部品 #${id} が見つかりません`);
    return node;
  };

  if (!engine || !report || !fileApi || !pdfApi || !docCheck || !pipeline) {
    const status = el('status');
    status.hidden = false;
    status.className = 'status is-error';
    status.textContent =
      '画面の部品を読み込めませんでした。ページを再読み込みしてください。直らない場合は管理者に連絡してください。';
    return;
  }

  const els = {
    input: /** @type {HTMLTextAreaElement} */ (el('input')),
    charCount: el('char-count'),
    btnCheck: /** @type {HTMLButtonElement} */ (el('btn-check')),
    btnSample: /** @type {HTMLButtonElement} */ (el('btn-sample')),
    btnClear: /** @type {HTMLButtonElement} */ (el('btn-clear')),
    status: el('status'),
    results: el('results'),
    summary: el('summary'),
    amountArea: el('amount-area'),
    tabList: el('tab-list'),
    tabBody: el('tab-body'),
    tabDoc: el('tab-doc'),
    panes: el('panes'),
    paneDoc: el('pane-doc'),
    docView: el('doc-view'),
    bodyView: el('body-view'),
    findingList: el('finding-list'),
    ignoredArea: el('ignored-area'),
    btnToggleIgnored: el('btn-toggle-ignored'),
    ignoredList: el('ignored-list'),
    groupRule: el('group-rule'),
    groupPosition: el('group-position'),
    btnCopyList: el('btn-copy-list'),
    ruleErrors: el('rule-errors'),
    manualName: el('manual-name'),
    manualVersion: el('manual-version'),
    rulesUpdated: el('rules-updated'),
    rulesDialog: /** @type {HTMLDialogElement} */ (el('rules-dialog')),
    rulesBody: el('rules-body'),
    btnRulesTop: el('btn-rules-top'),
    btnRulesBottom: el('btn-rules-bottom'),
    btnRulesClose: el('btn-rules-close'),
    aiCaution: el('ai-caution'),
    aiLinks: el('ai-links'),
    btnCopyPrompt: el('btn-copy-prompt'),
    sheet: el('detail-sheet'),
    sheetBody: el('sheet-body'),
    btnSheetClose: el('btn-sheet-close'),
    toast: el('toast'),
    btnFile: /** @type {HTMLButtonElement} */ (el('btn-file')),
    fileInput: /** @type {HTMLInputElement} */ (el('file-input')),
    filePanel: el('file-panel'),
    fileFormat: el('file-format'),
    fileName: el('file-name'),
    fileUnits: el('file-units'),
    fileNotes: el('file-notes'),
    btnToPaste: /** @type {HTMLButtonElement} */ (el('btn-to-paste')),
    pasteArea: el('paste-area'),
    bodyNote: el('body-note'),
    dropOverlay: el('drop-overlay'),
  };

  const PASTE_BODY_NOTE =
    'チェックした文章です。印を押すと指摘の内容が出ます。直すときは上の入力欄で直して、もう一度「チェックする」を押してください。';
  const FILE_BODY_NOTE =
    'ファイルから読み込んだ文章です（ここでは直せません）。印を押すと指摘の内容が出ます。直す場合は元のファイル（Word・Excel）を修正して、もう一度読み込んでください。';

  const ruleset = engine.compileRuleset(g.YEG_RULESET);
  const sampleText = typeof g.YEG_SAMPLE_TEXT === 'string' ? g.YEG_SAMPLE_TEXT : '';
  const narrowQuery = window.matchMedia('(max-width: 959.98px)');

  /** @type {State} */
  const INITIAL_STATE = Object.freeze({
    result: null,
    checkedText: '',
    ignored: new Set(),
    groupMode: 'rule',
    filter: 'all',
    showIgnored: false,
    tab: 'list',
    files: [],
    failed: [],
    prepared: null,
    doc: null,
    checked: new Set(),
  });

  /** @type {State} */
  let state = INITIAL_STATE;
  let busy = false;
  /** 読み込んだファイルの目印（f1, f2, …） */
  let fileSeq = 0;
  /** 複数のファイルを読み込み中の進み具合（PDF のページの進み具合の前に付ける） */
  let progressPrefix = '';

  /** @param {Partial<State>} patch */
  const setState = (patch) => {
    state = { ...state, ...patch };
    render();
  };

  // ------------------------------------------------------------------
  // 部品づくり
  // ------------------------------------------------------------------

  /**
   * @param {string} tag
   * @param {{ className?: string, text?: string, attrs?: Record<string, string>, onClick?: (e: MouseEvent) => void }} [props]
   * @param {(Node|string|null|false)[]} [children]
   * @returns {HTMLElement}
   */
  const h = (tag, props = {}, children = []) => {
    const node = document.createElement(tag);
    if (props.className) node.className = props.className;
    if (props.text !== undefined) node.textContent = props.text;
    for (const [k, v] of Object.entries(props.attrs ?? {})) node.setAttribute(k, v);
    if (props.onClick) node.addEventListener('click', props.onClick);
    for (const c of children) {
      if (c === null || c === false) continue;
      node.append(typeof c === 'string' ? document.createTextNode(c) : c);
    }
    return node;
  };

  /**
   * 条件によって出さない部品（null・false）を取り除く
   * @param {(Node|null|false)[]} items
   * @returns {Node[]}
   */
  const compact = (items) => items.flatMap((n) => (n === null || n === false ? [] : [n]));

  /** @type {Readonly<Record<Severity, string>>} */
  const ICONS = Object.freeze({
    違反:
      '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" focusable="false"><circle cx="8" cy="8" r="7.5" fill="currentColor"/><path d="M5 5l6 6M11 5l-6 6" stroke="#fff" stroke-width="2" stroke-linecap="round"/></svg>',
    要確認:
      '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" focusable="false"><path d="M8 1L15.5 14.5H.5z" fill="currentColor"/><path d="M8 5.5v4.5" stroke="#fff" stroke-width="2" stroke-linecap="round"/><circle cx="8" cy="12.3" r="1.1" fill="#fff"/></svg>',
    参考:
      '<svg viewBox="0 0 16 16" width="16" height="16" aria-hidden="true" focusable="false"><circle cx="8" cy="8" r="7.5" fill="currentColor"/><path d="M8 7v5" stroke="#fff" stroke-width="2" stroke-linecap="round"/><circle cx="8" cy="4.5" r="1.1" fill="#fff"/></svg>',
  });

  /** @param {Severity} sev */
  const icon = (sev) => {
    const t = document.createElement('template');
    t.innerHTML = ICONS[sev];
    return /** @type {Node} */ (t.content.firstChild);
  };

  /**
   * 重さのラベル（色・アイコン・文字で区別する）
   * @param {Severity} sev
   * @param {string} [label]
   */
  const sevBadge = (sev, label) =>
    h('span', { className: `sev sev-${report.severityClass(sev)}` }, [icon(sev), label ?? sev]);

  // ------------------------------------------------------------------
  // お知らせ・トースト
  // ------------------------------------------------------------------

  /**
   * @param {string} message
   * @param {'info'|'warn'|'error'|'busy'} kind
   */
  const showStatus = (message, kind) => {
    els.status.hidden = false;
    els.status.className = `status is-${kind}`;
    els.status.textContent = message;
  };

  const hideStatus = () => {
    els.status.hidden = true;
    els.status.textContent = '';
  };

  /** @type {number|undefined} */
  let toastTimer;
  /**
   * @param {string} message
   * @param {boolean} [isError]
   */
  const showToast = (message, isError = false) => {
    els.toast.textContent = message;
    els.toast.className = `toast${isError ? ' is-error' : ''}`;
    els.toast.hidden = false;
    window.clearTimeout(toastTimer);
    toastTimer = window.setTimeout(() => {
      els.toast.hidden = true;
    }, isError ? 6000 : 2500);
  };

  /**
   * クリップボードに文字を入れる。使えないときは理由つきのエラーにする
   * @param {string} text
   * @returns {Promise<void>}
   */
  const copyToClipboard = async (text) => {
    /** @type {unknown} */
    let firstError = null;
    if (navigator.clipboard && window.isSecureContext) {
      try {
        await navigator.clipboard.writeText(text);
        return;
      } catch (e) {
        firstError = e;
      }
    }
    const ta = document.createElement('textarea');
    ta.value = text;
    ta.setAttribute('readonly', '');
    ta.style.position = 'fixed';
    ta.style.top = '-1000px';
    ta.style.opacity = '0';
    document.body.append(ta);
    ta.select();
    ta.setSelectionRange(0, text.length);
    const ok = document.execCommand('copy');
    ta.remove();
    if (!ok) {
      const detail = firstError instanceof Error ? `（${firstError.message}）` : '';
      throw new Error(`この端末ではコピー機能が使えませんでした${detail}`);
    }
  };

  /**
   * @param {string} text
   * @param {string} doneMessage
   */
  const copyWithToast = (text, doneMessage) => {
    copyToClipboard(text).then(
      () => showToast(doneMessage),
      (e) => showToast(e instanceof Error ? e.message : 'コピーできませんでした', true),
    );
  };

  // ------------------------------------------------------------------
  // 最初に1回だけ作る部分（版の表示・ルール一覧・AIへの案内）
  // ------------------------------------------------------------------

  const renderLoadInfo = () => {
    els.manualName.textContent = ruleset.meta.manualName;
    els.manualVersion.textContent = ruleset.meta.manualVersion;
    els.rulesUpdated.textContent = ruleset.meta.updated;
    /** @type {string[]} */
    const lines = [
      ...(ruleset.fatal ? [ruleset.fatal] : []),
      ...ruleset.errors.map(
        (/** @type {{ ruleLabel: string, message: string }} */ e) =>
          `ルール ${e.ruleLabel} を読み込めませんでした（${e.message}）。ほかのルールはそのまま使えます。`,
      ),
    ];
    els.ruleErrors.replaceChildren(...lines.map((t) => h('p', { text: t })));
    els.ruleErrors.hidden = lines.length === 0;
    if (ruleset.fatal) {
      els.btnCheck.disabled = true;
      els.btnSample.disabled = true;
    }
  };

  /**
   * ルール一覧の画面に出す、書類の種類ごとの判別語・必須項目・確認ポイント
   * @param {(label: string, value: string) => HTMLElement[]} row
   * @returns {HTMLElement[]}
   */
  const docRuleNodes = (row) => {
    const model = docCheck.documentsModel(ruleset);
    if (!model) return [];
    /** @param {{ id: string, text: string }[]} items */
    const points = (items) => h('ul', { className: 'points-list' }, items.map((k) => h('li', { text: `${k.id} ${k.text}` })));
    return [
      h('h3', { className: 'rule-category', attrs: { id: 'rules-doc-title' }, text: '書類の種類ごとの必須項目と確認ポイント' }),
      h('p', {
        text: `書類の種類は、判別語が見出しとして${model.minDetectWords}つ以上見つかったときに自動で判別します（画面で選び直せます）。「空欄も確認」の見出しは、Excel・Wordの表で読み込んだときだけ空欄かどうかを確認します。確認ポイントは、機械で判定できない項目です。`,
      }),
      h('div', { className: 'rule-item', attrs: { 'data-doc-type': 'common' } }, [
        h('h4', { text: '共通の確認ポイント' }),
        points(model.common),
      ]),
      ...model.types.map((t) =>
        h('div', { className: 'rule-item', attrs: { 'data-doc-type': t.id } }, [
          h('h4', { text: t.name }),
          h('dl', {}, [
            ...row('判別語', t.detectWords.join('／')),
            ...row('必須項目', [...t.required, ...t.settlementRequired.map((n) => `${n}（決算のとき）`)].join('／')),
            ...row('空欄も確認', t.blankCheck.join('／')),
          ]),
          h('p', { className: 'points-title', text: `確認ポイント（${t.name}）` }),
          points(t.checkpoints),
        ]),
      ),
    ];
  };

  const renderRuleDialog = () => {
    const model = report.buildRuleListModel(ruleset);
    /** @param {string} label @param {string} value */
    const row = (label, value) => [h('dt', { text: label }), h('dd', { text: value || '—' })];
    els.rulesBody.replaceChildren(
      h('p', {
        text: `有効なルール: ${model.total}件（基準: ${ruleset.meta.manualName} ${ruleset.meta.manualVersion}／ルール定義の更新日: ${ruleset.meta.updated}）`,
      }),
      h('p', {}, [
        sevBadge('違反'),
        ' マニュアルのルールに明確に反する　',
        sevBadge('要確認'),
        ' 文脈によっては問題ない（間違って指摘している場合もある）　',
        sevBadge('参考'),
        ' 直したほうが読みやすい（マニュアルの「参考」）',
      ]),
      ...model.categories.flatMap(
        (/** @type {RuleListModel['categories'][number]} */ c) => [
          h('h3', { className: 'rule-category', text: `${c.name}（${c.rules.length}件）` }),
          ...(c.note ? [h('p', { className: 'rule-category-note', text: c.note })] : []),
          ...c.rules.map((r) =>
            h('div', { className: 'rule-item', attrs: { 'data-rule-id': r.id } }, [
              h('h4', {}, [sevBadge(r.severity), h('span', { text: r.name }), h('span', { className: 'reason-id', text: r.id })]),
              h('dl', {}, [
                ...row('検出する語の例', r.examples.join('／')),
                ...row('正しい書き方', r.fixes.join('／')),
                ...row('説明', r.description),
                ...row('出典', r.source),
                ...row('適用範囲', r.scope),
              ]),
            ]),
          ),
        ],
      ),
      ...docRuleNodes(row),
      h('h3', { className: 'rule-category', text: 'このツールで判定しないもの' }),
      h(
        'ul',
        { className: 'not-checked' },
        model.notChecked.map((/** @type {string} */ t) => h('li', { text: t })),
      ),
    );
  };

  const renderAiGuide = () => {
    els.aiCaution.textContent = ruleset.aiGuide.caution;
    els.aiLinks.replaceChildren(
      ...ruleset.aiGuide.links.map((/** @type {{ label: string, url: string }} */ l) =>
        h('li', {}, [
          h('a', {
            text: `${l.label} を開く（新しいタブ）`,
            attrs: { href: l.url, target: '_blank', rel: 'noopener noreferrer' },
          }),
        ]),
      ),
    );
  };

  // ------------------------------------------------------------------
  // 判定結果の表示
  // ------------------------------------------------------------------

  /** 書類全体の指摘（書類の種類に応じたもの）・金額チェックの指摘を先に、文章の指摘を後に並べる */
  /** @returns {Finding[]} */
  const allFindings = () => {
    if (!state.result) return [];
    if (!state.doc) return state.result.findings;
    const ev = state.doc.evaluation;
    return [...ev.documentFindings, ...ev.amountFindings, ...ev.textFindings];
  };

  /** @returns {Finding[]} */
  const visible = () => report.visibleFindings(allFindings(), state.ignored);

  const isStale = () => state.result !== null && state.files.length === 0 && els.input.value !== state.checkedText;

  /** 2つ以上のファイルを読み込んでいる（指摘の場所の先頭にファイル名を付ける） */
  const isMulti = () => state.files.length > 1;

  /**
   * 指摘の場所（ファイルのページ・シートとセル番地。複数のファイルのときは先頭にファイル名）。貼り付けた文章では空文字
   * @param {Finding} f
   */
  const locationOf = (f) => (state.prepared ? pipeline.locationOf(state.prepared, f) : '');

  /**
   * 指摘一覧のコピーに入れる場所（例: 「2ページ／欄: 事業目的・趣旨」）
   * @param {Finding} f
   */
  const placeOf = (f) =>
    [locationOf(f), f.section ? `欄: ${f.section}` : ''].filter((t) => t !== '').join('／');

  /** @param {LoadedDoc} doc */
  const unitWord = (doc) => (doc.format === 'pdf' ? 'ページ' : 'シート');

  /**
   * 判定に含めた範囲の説明（例: 「判定に含めたページ 2／3」）。Word は1つの単位なので空文字
   * @param {FileEntry} file
   */
  const includedScopeOf = (file) => {
    if (file.doc.format === 'docx') return '';
    const n = file.doc.units.filter((u) => file.included.has(u.id)).length;
    return `判定に含めた${unitWord(file.doc)} ${n}／${file.doc.units.length}`;
  };

  /** 判定に含めた範囲の説明（複数のファイルのときは、ファイルごとに「名前: 範囲」） */
  const includedScope = () => {
    if (state.files.length === 1) return includedScopeOf(state.files[0]);
    return state.files
      .flatMap((f) => {
        const s = includedScopeOf(f);
        return s ? [`${f.doc.name}: ${s.replace('判定に含めた', '')}`] : [];
      })
      .join('、');
  };

  /** 読み込んだファイルの名前（複数のときは「、」でつなぐ） */
  const fileNames = () => state.files.map((f) => f.doc.name).join('、');

  /**
   * 金額チェックのまとめの1行（例: 「金額チェック: 問題は見つかりませんでした」）と、注意して見てほしいか。
   * 金額チェックの対象でなければ text は空文字
   * @param {Finding[]} shown  無視した指摘を除いた指摘
   * @returns {{ text: string, warn: boolean }}
   */
  const amountSummary = (shown) => {
    const view = state.doc ? state.doc.evaluation.amount : null;
    if (!view || view.status === 'none') return { text: '', warn: false };
    if (view.status === 'needs-excel') return { text: pipeline.MESSAGES.needsExcel, warn: false };
    if (view.status === 'unavailable') return { text: pipeline.MESSAGES.unavailable, warn: true };
    const failed = view.sheets.filter((s) => !s.ok);
    if (failed.length === view.sheets.length) {
      const first = failed[0];
      return {
        text: first ? `金額チェック: ${pipeline.MESSAGES.notRead(first.reasons)}${failed.length > 1 ? `（ほか${failed.length - 1}件）` : ''}` : '',
        warn: true,
      };
    }
    const amounts = shown.filter((f) => f.kind === 'amount');
    const partly = failed.length > 0 ? `（収支表の形を読み取れなかったシートが${failed.length}件あります。下の「金額チェック」を見てください）` : '';
    return {
      text: `${amounts.length === 0 ? pipeline.MESSAGES.noProblem : `金額チェック: ${report.formatCounts(report.countBySeverity(amounts))}`}${partly}`,
      warn: amounts.length > 0 || failed.length > 0,
    };
  };

  const renderSummary = () => {
    const result = state.result;
    if (!result) return;
    const shown = visible();
    const ignoredCount = allFindings().length - shown.length;
    const counts = report.countBySeverity(shown);
    /** @type {Node[]} */
    const countNodes = report.SEVERITIES.filter((/** @type {Severity} */ s) => counts[s] > 0).flatMap(
      (/** @type {Severity} */ s, /** @type {number} */ i) => [
        ...(i > 0 ? [document.createTextNode('／')] : []),
        sevBadge(s, `${s} ${counts[s]}件`),
      ],
    );
    const filterLabel = state.filter === 'all' ? '' : `${state.filter}だけ`;
    const amountLine = amountSummary(shown);
    const shownInList = report.filterBySeverity(shown, state.filter).length;

    els.summary.replaceChildren(
      ...compact([
      ...(shown.length > 0
        ? [h('p', { className: 'summary-counts', attrs: { id: 'summary-counts' } }, countNodes)]
        : [
            h('p', {
              className: 'summary-ok',
              text:
                ignoredCount > 0
                  ? `問題は見つかりませんでした（無視した指摘 ${ignoredCount}件を除く）`
                  : '問題は見つかりませんでした',
            }),
            h('p', {
              className: 'summary-line',
              text: '機械で判定できない点もあります。「確認ポイント」で確かめ、最終的には本人と委員会で確認してください。',
            }),
          ]),
      ...summaryDocNodes(),
      amountLine.text
        ? h('p', {
            className: `summary-line summary-amount${amountLine.warn ? ' is-short' : ''}`,
            attrs: { id: 'summary-amount' },
            text: amountLine.text,
          })
        : null,
      state.files.length === 1
        ? h('p', {
            className: 'summary-line summary-file',
            text: `読み込んだファイル: ${state.files[0].doc.name}${includedScope() ? `（${includedScope()}）` : ''}`,
          })
        : null,
      state.files.length > 1
        ? h('p', {
            className: 'summary-line summary-file',
            text: `読み込んだファイル: ${state.files.length}件（${fileNames()}）`,
          })
        : null,
      state.failed.length > 0
        ? h('p', {
            className: 'summary-line summary-failed',
            text: `読み込めなかったファイル: ${state.failed.map((f) => f.name).join('、')}（表記のチェックにも含めていません）`,
          })
        : null,
      h('p', { className: 'summary-line', text: `判定した文字数: ${result.charCount.toLocaleString('ja-JP')}字` }),
      result.styleCounts && report.hasStyleFinding(result.findings)
        ? h('p', { className: 'summary-line', text: `文体: ${report.formatStyleCounts(result.styleCounts)}` })
        : null,
      ignoredCount > 0 && shown.length > 0
        ? h('p', { className: 'summary-line', text: `（無視した指摘 ${ignoredCount}件は件数に含めていません）` })
        : null,
      filterLabel
        ? h('p', {
            className: 'summary-filter',
            text: `絞り込み中: 指摘一覧には「${filterLabel}」を表示しています（${shownInList}件）。件数は全体の数です。`,
          })
        : null,
      h('p', {
        className: 'summary-stale',
        attrs: { id: 'summary-stale', ...(isStale() ? {} : { hidden: '' }) },
        text: '入力欄の文章が変わりました。もう一度「チェックする」を押すと結果が新しくなります。',
      }),
      ]),
    );
  };

  /**
   * 書類の種類ごとに使ったファイル（複数のファイルのとき）
   * @param {string} typeId
   */
  const filesOfType = (typeId) => {
    const doc = state.doc;
    const grp = doc ? doc.evaluation.groups.find((x) => x.typeId === typeId) : undefined;
    return grp ? grp.files : [];
  };

  /**
   * 複数のファイルのとき、件数のまとめに出す「書類の種類ごとの必須項目」の表（付録F-3 AC-05-03）
   * @param {RequiredReport} rep
   */
  const summaryTypesTable = (rep) =>
    h('div', { className: 'req-table-wrap' }, [
      h('table', { className: 'req-table summary-types', attrs: { id: 'summary-types' } }, [
        h('caption', { text: '書類の種類ごとの必須項目' }),
        h('thead', {}, [
          h('tr', {}, [
            h('th', { text: '書類の種類', attrs: { scope: 'col' } }),
            h('th', { text: 'ファイル', attrs: { scope: 'col' } }),
            h('th', { text: '必須項目', attrs: { scope: 'col' } }),
          ]),
        ]),
        h(
          'tbody',
          {},
          rep.tables.map((t) => {
            const ok = t.rows.filter((r) => r.status === 'filled' || r.status === 'present').length;
            const short = t.rows.length - ok;
            return h('tr', { attrs: { 'data-type': t.typeId } }, [
              h('th', { text: t.title, attrs: { scope: 'row' } }),
              h('td', { text: filesOfType(t.typeId).join('、') }),
              h('td', {}, [
                h('span', {
                  className: `req-status ${short > 0 ? 'st-missing' : 'st-ok'}`,
                  text: `${ok}／${t.rows.length}${short > 0 ? `（不足 ${short}）` : ''}`,
                }),
              ]),
            ]);
          }),
        ),
      ]),
    ]);

  /** 件数のまとめに出す、書類の種類と必須項目の行 */
  const summaryDocNodes = () => {
    const doc = state.doc;
    if (!doc || !doc.evaluation.analysis.available) return [];
    const names = docCheck.typeNames(doc.selection, ruleset);
    const how = doc.selection.auto ? '自動で判別' : '手動で選択';
    const rep = doc.evaluation.report;
    const typeLine =
      names.length > 0
        ? h('p', { className: 'summary-line summary-doc', attrs: { id: 'summary-doctype' } }, [
            '書類の種類: ',
            h('strong', { text: names.join('／') }),
            h('span', { className: 'doc-how', text: `（${how}）` }),
          ])
        : h('p', {
            className: 'summary-line summary-doc is-undetermined',
            attrs: { id: 'summary-doctype' },
            text: doc.selection.auto ? docCheck.MESSAGES.undetermined : `書類の種類: 選ばれていません（${docCheck.MESSAGES.noneSelected}）`,
          });
    return compact([
      typeLine,
      names.length > 0
        ? h('p', {
            className: `summary-line summary-required${rep.shortage > 0 ? ' is-short' : ''}`,
            attrs: { id: 'summary-required' },
            text: docCheck.requiredSummary(rep),
          })
        : null,
      names.length > 0 && isMulti() && rep.tables.length > 0 ? summaryTypesTable(rep) : null,
      h('p', { className: 'summary-line summary-doc-action' }, [
        h('button', {
          className: 'btn btn-small',
          text: names.length > 0 ? '必須項目と確認ポイントを見る' : '書類の種類を選ぶ・確認ポイントを見る',
          attrs: { type: 'button', id: 'btn-show-doc' },
          onClick: showDocPane,
        }),
      ]),
    ]);
  };

  /**
   * @param {Finding} f
   * @param {{ ignored: boolean, inSheet: boolean }} opts
   */
  const renderCard = (f, opts) => {
    const primary = f.reasons[0];
    const location = locationOf(f);
    const locRow = location
      ? h('p', { className: 'card-loc' }, [
          h('span', { className: 'card-row-label', text: '場所' }),
          h('span', { className: 'card-loc-text', text: location }),
        ])
      : null;
    const sectionRow = f.section ? h('p', { className: 'card-section', text: `欄: ${f.section}` }) : null;
    const amount = f.kind === 'amount' ? f.amount ?? null : null;
    const context = amount
      ? h('div', { className: 'card-context card-amount' }, [
          h('p', {}, [h('span', { className: 'card-row-label', text: '内容' }), h('strong', { className: 'card-hit', text: primary.note })]),
          amount.calc
            ? h('p', { className: 'card-calc' }, [h('span', { className: 'card-row-label', text: '計算' }), amount.calc])
            : null,
        ])
      : f.range === null
        ? h('p', { className: 'card-context' }, [
            h('span', { className: 'card-row-label', text: '該当箇所' }),
            h('strong', { className: 'card-hit card-whole', text: report.WHOLE_DOCUMENT }),
          ])
        : h('p', { className: 'card-context' }, [
            h('span', { className: 'card-row-label', text: '該当箇所' }),
            h('br'),
            h('span', { className: 'ctx', text: report.contextLine(f.before) }),
            h('strong', { className: 'card-hit', text: report.oneLine(f.text) }),
            h('span', { className: 'ctx', text: report.contextLine(f.after) }),
          ]);

    const reasons = f.reasons.map((r) =>
      h('div', { className: 'reason' }, [
        h('p', { className: 'reason-head' }, [
          sevBadge(r.severity),
          h('span', { className: 'reason-name', text: r.ruleName }),
          h('span', { className: 'reason-id', text: r.ruleId }),
        ]),
        r.fixes.length > 0
          ? h('p', { className: 'card-row' }, [
              h('span', { className: 'card-row-label', text: '正しい書き方' }),
              h('span', { className: 'card-fix', text: report.formatFix(r.fixes) }),
            ])
          : r.hint
            ? h('p', { className: 'card-row' }, [
                h('span', { className: 'card-row-label', text: '直し方' }),
                h('span', { className: 'card-fix', text: r.hint }),
              ])
            : null,
        r.note && !amount ? h('p', { className: 'card-row', text: r.note }) : null,
        h('p', { className: 'card-row' }, [h('span', { className: 'card-row-label', text: '理由' }), r.description]),
        h('p', { className: 'card-row card-source' }, [
          h('span', { className: 'card-row-label', text: '出典' }),
          report.formatSource(ruleset, r.source),
        ]),
      ]),
    );

    const actions = h('div', { className: 'card-actions' }, [
      amount && !opts.ignored
        ? h('button', {
            className: 'btn btn-small',
            text: '収支表で見る',
            attrs: { type: 'button', 'data-action': 'show-in-table' },
            onClick: () => showInTable(amount.tableId, amount.rowId),
          })
        : null,
      f.range && !opts.ignored && !opts.inSheet
        ? h('button', {
            className: 'btn btn-small',
            text: '本文で見る',
            attrs: { type: 'button', 'data-action': 'show-in-body' },
            onClick: () => showInBody(f.id),
          })
        : null,
      opts.inSheet
        ? h('button', {
            className: 'btn btn-small',
            text: '一覧で見る',
            attrs: { type: 'button', 'data-action': 'show-in-list' },
            onClick: () => showInList(f.id),
          })
        : null,
      opts.ignored
        ? h('button', {
            className: 'btn btn-small',
            text: '無視を取り消す',
            attrs: { type: 'button', 'data-action': 'unignore' },
            onClick: () => unignore(f.key),
          })
        : h('button', {
            className: 'btn btn-small',
            text: 'この指摘を無視',
            attrs: { type: 'button', 'data-action': 'ignore' },
            onClick: () => ignore(f.key),
          }),
    ]);

    const card = h(
      'article',
      {
        className: `card sev-border-${report.severityClass(f.severity)}${opts.ignored ? ' is-ignored' : ''}${amount ? ' card-amount-finding' : ''}`,
        attrs: {
          id: opts.inSheet ? `sheet-card-${f.id}` : opts.ignored ? `ignored-card-${f.id}` : `card-${f.id}`,
          'data-finding-id': f.id,
          'data-rule-ids': f.reasons.map((r) => r.ruleId).join(' '),
          'data-severity': f.severity,
          'aria-label': `${primary.severity} ${primary.ruleName}${location ? `（${location}）` : ''}${f.section ? `（欄: ${f.section}）` : ''}`,
          ...(location ? { 'data-location': location } : {}),
          ...(f.section ? { 'data-section': f.section } : {}),
          ...(amount ? { 'data-table-id': amount.tableId, 'data-row-id': amount.rowId } : {}),
        },
      },
      compact([locRow, sectionRow, context, ...reasons, actions]),
    );
    // 金額チェックのカードは、押すと読み取った収支表の該当行を強調する（ボタンを押したときは、そのボタンの動きだけ）
    if (amount && !opts.ignored) {
      card.addEventListener('click', (e) => {
        if (e.target instanceof Element && e.target.closest('button')) return;
        showInTable(amount.tableId, amount.rowId);
      });
    }
    return card;
  };

  const renderList = () => {
    const shown = report.filterBySeverity(visible(), state.filter);
    const groups = report.groupFindings(shown, state.groupMode);
    /** @type {Node[]} */
    const nodes =
      shown.length === 0
        ? [
            h('p', {
              className: 'empty-note',
              text: visible().length === 0 ? '指摘はありません。' : 'この絞り込みで表示する指摘はありません。',
            }),
          ]
        : groups.flatMap(
            (/** @type {{ key: string, title: string|null, severity: Severity|null, findings: Finding[] }} */ grp) => [
              grp.title
                ? h('h3', { className: 'group-title', attrs: { 'data-group': grp.key } }, [
                    grp.severity
                      ? h('span', { className: `group-icon sev-color-${report.severityClass(grp.severity)}` }, [
                          icon(grp.severity),
                        ])
                      : null,
                    grp.title,
                  ])
                : null,
              ...grp.findings.map((f) => renderCard(f, { ignored: false, inSheet: false })),
            ].filter((n) => n !== null),
          );
    els.findingList.replaceChildren(...nodes);

    for (const btn of els.panes.querySelectorAll('[data-filter]')) {
      btn.setAttribute('aria-pressed', String(btn.getAttribute('data-filter') === state.filter));
    }
    els.groupRule.setAttribute('aria-pressed', String(state.groupMode === 'rule'));
    els.groupPosition.setAttribute('aria-pressed', String(state.groupMode === 'position'));

    const ignoredList = report.ignoredFindings(allFindings(), state.ignored);
    els.ignoredArea.hidden = ignoredList.length === 0;
    els.btnToggleIgnored.textContent = state.showIgnored
      ? `無視した指摘（${ignoredList.length}件）を隠す`
      : `無視した指摘（${ignoredList.length}件）を表示`;
    els.btnToggleIgnored.setAttribute('aria-expanded', String(state.showIgnored));
    els.ignoredList.hidden = !state.showIgnored;
    els.ignoredList.replaceChildren(
      ...(state.showIgnored
        ? ignoredList.map((/** @type {Finding} */ f) => renderCard(f, { ignored: true, inSheet: false }))
        : []),
    );
  };

  /**
   * 本文の表示。指摘箇所を重さごとの見た目で強調する（重なった箇所は1つの印にまとめる）
   */
  const renderBody = () => {
    const result = state.result;
    if (!result) {
      els.bodyView.replaceChildren();
      return;
    }
    const text = result.text;
    const ranged = visible()
      .filter((f) => f.range !== null)
      .sort((a, b) => /** @type {{ start: number }} */ (a.range).start - /** @type {{ start: number }} */ (b.range).start);
    // ファイル（PDF・Excel）では、ページ・シートの始まりに見出しを出す（判定する文字には含めない）。
    // 複数のファイルのときは、先頭にファイル名を付ける（Word はファイル名だけ）
    const prep = state.prepared;
    const composed = prep ? prep.composed : null;
    const multi = isMulti();
    /** @type {Map<number, string>} */
    const heads = new Map(
      prep && composed
        ? composed.unitRanges.flatMap((r) => {
            const part = prep.parts.find((p) => r.start >= p.start && r.start < p.end);
            const entry = part ? state.files.find((f) => f.key === part.key) : undefined;
            if (!part || !entry) return [];
            if (entry.doc.format === 'docx') return multi ? [[r.start, entry.doc.name]] : [];
            const unit = entry.doc.units.find((u) => fileApi.unitKeyOf(entry.key, u.id, multi) === r.unitId);
            if (!unit) return [];
            return [[r.start, multi ? `${entry.doc.name}／${fileApi.unitLabel(unit)}` : fileApi.unitLabel(unit)]];
          })
        : [],
    );
    const points = [
      ...new Set([
        0,
        text.length,
        ...heads.keys(),
        ...ranged.flatMap((f) => (f.range ? [f.range.start, f.range.end] : [])),
      ]),
    ].sort((a, b) => a - b);

    const frag = document.createDocumentFragment();
    /** @type {Finding[]} */
    let active = [];
    let next = 0;
    for (let i = 0; i < points.length - 1; i += 1) {
      const a = points[i];
      const b = points[i + 1];
      const head = heads.get(a);
      if (head !== undefined) frag.append(h('span', { className: 'unit-head', text: head }));
      while (next < ranged.length && /** @type {{ start: number }} */ (ranged[next].range).start <= a) {
        active.push(ranged[next]);
        next += 1;
      }
      active = active.filter((f) => /** @type {{ end: number }} */ (f.range).end > a);
      const segment = text.slice(a, b);
      if (active.length === 0) {
        frag.append(document.createTextNode(segment));
        continue;
      }
      const words = active.filter((f) => f.kind !== 'sentence');
      const sentences = active.filter((f) => f.kind === 'sentence');
      const lead = [...(words.length > 0 ? words : sentences)].sort(
        (x, y) => report.SEVERITIES.indexOf(x.severity) - report.SEVERITIES.indexOf(y.severity),
      )[0];
      const classes = [
        'hl',
        `hl-${report.severityClass(lead.severity)}`,
        ...(sentences.length > 0 ? ['hl-sentence'] : []),
      ];
      const label = active.map((f) => `${f.severity} ${f.reasons[0].ruleName}`).join('、');
      frag.append(
        h('mark', {
          className: classes.join(' '),
          text: segment,
          attrs: {
            'data-ids': active.map((f) => f.id).join(' '),
            tabindex: '0',
            role: 'button',
            'aria-label': `${segment.replace(/\n/g, '')}（${label}）`,
          },
        }),
      );
    }
    els.bodyView.replaceChildren(frag);
  };

  const renderTabs = () => {
    els.panes.setAttribute('data-tab', state.tab);
    els.tabList.setAttribute('aria-selected', String(state.tab === 'list'));
    els.tabBody.setAttribute('aria-selected', String(state.tab === 'body'));
    els.tabDoc.setAttribute('aria-selected', String(state.tab === 'doc'));
  };

  // ------------------------------------------------------------------
  // 書類の種類・必須項目・確認ポイント
  // ------------------------------------------------------------------

  /** 描き直した後にフォーカスを戻す部品の id（書類の種類の切り替え） */
  /** @type {string|null} */
  let focusDocId = null;

  /** @type {Readonly<Record<RequiredRow['status'], string>>} */
  const STATUS_CLASS = Object.freeze({ filled: 'st-ok', present: 'st-ok', blank: 'st-blank', missing: 'st-missing' });

  /**
   * チェック欄つきの行（書類の種類・確認ポイント）
   * @param {{ id: string, checked: boolean, label: (Node|string)[], onChange: (checked: boolean) => void, className?: string, disabled?: boolean }} o
   */
  const checkRow = (o) => {
    const box = document.createElement('input');
    box.type = 'checkbox';
    box.id = o.id;
    box.checked = o.checked;
    box.disabled = Boolean(o.disabled);
    box.addEventListener('change', () => o.onChange(box.checked));
    return h('label', { className: `check-row${o.className ? ` ${o.className}` : ''}`, attrs: { for: o.id } }, [
      box,
      h('span', { className: 'check-text' }, o.label),
    ]);
  };

  /**
   * 予算・決算の切り替え
   * @param {Selection} sel
   */
  const budgetKindControl = (sel) =>
    h(
      'div',
      { className: 'budget-kind', attrs: { role: 'radiogroup', 'aria-label': '収支予算書か決算書か' } },
      /** @type {('予算'|'決算')[]} */ (['予算', '決算']).map((kind) => {
        const radio = document.createElement('input');
        radio.type = 'radio';
        radio.name = 'budget-kind';
        radio.id = `budget-kind-${kind === '予算' ? 'budget' : 'settlement'}`;
        radio.checked = sel.budgetKind === kind;
        radio.disabled = busy;
        radio.addEventListener('change', () => {
          if (radio.checked) changeSelection({ ...sel, budgetKind: kind, auto: false }, radio.id);
        });
        return h('label', { className: 'check-row radio-row', attrs: { for: radio.id } }, [
          radio,
          h('span', { className: 'check-text', text: kind === '予算' ? '予算（案）' : '決算（案）' }),
        ]);
      }),
    );

  /** 書類の種類・必須項目の表・確認ポイントを描く */
  const renderDoc = () => {
    const doc = state.doc;
    const prep = state.prepared;
    if (!doc || !prep || !doc.evaluation.analysis.available) {
      els.docView.replaceChildren(h('p', { className: 'doc-note', text: docCheck.MESSAGES.unavailable }));
      return;
    }
    const sel = doc.selection;
    const ev = doc.evaluation;
    const types = ruleset.documents ? ruleset.documents.types : [];
    const names = docCheck.typeNames(sel, ruleset);

    const typeSection = h('section', { className: 'doc-block doc-types', attrs: { 'aria-labelledby': 'doc-types-title' } }, [
      h('h3', { className: 'doc-title', attrs: { id: 'doc-types-title' } }, [
        '書類の種類',
        h('span', { className: 'doc-how', text: sel.auto ? '（自動で判別）' : '（手動で選択）' }),
      ]),
      names.length === 0
        ? h('p', { className: 'doc-message', text: sel.auto ? docCheck.MESSAGES.undetermined : docCheck.MESSAGES.noneSelected })
        : h('p', { className: 'doc-current', text: `いまの種類: ${names.join('／')}` }),
      h('p', { className: 'doc-note', text: '違う場合は、次のチェックで選び直せます（複数選べます）。' }),
      h(
        'div',
        { className: 'type-choices', attrs: { role: 'group', 'aria-label': '書類の種類を選ぶ' } },
        types.map((t) =>
          checkRow({
            id: `doc-type-${t.id}`,
            checked: sel.types.includes(t.id),
            label: [t.name],
            disabled: busy,
            onChange: (on) =>
              changeSelection(
                {
                  ...sel,
                  types: types.map((x) => x.id).filter((id) => (id === t.id ? on : sel.types.includes(id))),
                  auto: false,
                },
                `doc-type-${t.id}`,
              ),
          }),
        ),
      ),
      sel.types.includes('budget') ? budgetKindControl(sel) : null,
      sel.auto
        ? null
        : h('p', {}, [
            h('button', {
              className: 'btn btn-small',
              text: '自動の判別に戻す',
              attrs: { type: 'button', id: 'btn-doc-auto' },
              onClick: () => changeSelection(pipeline.autoSelection(prep, ruleset), 'doc-types-title'),
            }),
          ]),
    ]);

    // 複数のファイルのとき: ファイルごとに判別した書類の種類（判別できないファイルは表記のルールだけで判定する）
    const typeName = (/** @type {string} */ id) => (types.find((t) => t.id === id) ?? { name: id }).name;
    const filesSection =
      prep.parts.length > 1
        ? h('section', { className: 'doc-block doc-files', attrs: { 'aria-labelledby': 'doc-files-title' } }, [
            h('h3', { className: 'doc-title', attrs: { id: 'doc-files-title' } }, ['ファイルごとの書類の種類']),
            h(
              'ul',
              { className: 'doc-file-types' },
              prep.parts.map((p) =>
                h('li', { attrs: { 'data-file-key': p.key } }, [
                  h('span', { className: 'doc-file-name', text: p.name }),
                  ': ',
                  p.analysis.detected.length > 0
                    ? h('strong', { text: p.analysis.detected.map(typeName).join('／') })
                    : h('span', { className: 'doc-file-none', text: '判別できませんでした（表記のルールだけで判定しています）' }),
                ]),
              ),
            ),
            h('p', { className: 'doc-note', text: '同じ種類のファイルは、合わせて1つの書類として必須項目を確かめています。' }),
          ])
        : null;

    const rep = ev.report;
    const requiredSection =
      rep.tables.length === 0
        ? null
        : h('section', { className: 'doc-block doc-required', attrs: { 'aria-labelledby': 'doc-required-title' } }, [
            h('h3', { className: 'doc-title', attrs: { id: 'doc-required-title' } }, [
              '必須項目',
              h('span', { className: `doc-count${rep.shortage > 0 ? ' is-short' : ''}`, text: `（${docCheck.requiredSummary(rep)}）` }),
            ]),
            ...rep.tables.map((t) =>
              h('div', { className: 'req-table-wrap' }, [
                h('table', { className: 'req-table', attrs: { 'data-type': t.typeId } }, [
                  h('caption', { text: prep.parts.length > 1 ? `${t.title}（${filesOfType(t.typeId).join('、')}）` : t.title }),
                  h('thead', {}, [
                    h('tr', {}, [h('th', { text: '見出し', attrs: { scope: 'col' } }), h('th', { text: '状態', attrs: { scope: 'col' } })]),
                  ]),
                  h(
                    'tbody',
                    {},
                    t.rows.map((r) =>
                      h('tr', { attrs: { 'data-item': r.name, 'data-status': r.status } }, [
                        h('th', { text: r.name, attrs: { scope: 'row' } }),
                        h('td', {}, [h('span', { className: `req-status ${STATUS_CLASS[r.status]}`, text: r.label })]),
                      ]),
                    ),
                  ),
                ]),
              ]),
            ),
            rep.notes.length > 0 ? h('ul', { className: 'doc-notes' }, rep.notes.map((n) => h('li', { text: n }))) : null,
          ]);

    const groups = docCheck.checkpointGroups(sel, ruleset, new Set(ev.amount.checkedRules));
    const pointSection = h('section', { className: 'doc-block doc-points', attrs: { 'aria-labelledby': 'doc-points-title' } }, [
      h('h3', { className: 'doc-title', attrs: { id: 'doc-points-title' } }, [
        '確認ポイント',
        h('span', { className: 'doc-how', text: '（機械で判定できない項目）' }),
      ]),
      h('p', { className: 'doc-note', text: '確かめた項目にチェックを付けられます（チェックは保存されず、ページを閉じると消えます）。' }),
      ...groups.flatMap((grp) => [
        h('h4', { className: 'points-group', text: grp.title }),
        h(
          'ul',
          { className: 'points' },
          grp.items.map((k) =>
            h('li', {}, [
              checkRow({
                id: `point-${k.id}`,
                checked: state.checked.has(k.id),
                label: [
                  h('span', { className: 'point-id', text: k.id }),
                  ` ${k.text}`,
                  ...(k.checkedBy.length > 0
                    ? [' ', h('span', { className: 'point-auto', text: `金額チェックで確認済み（${k.checkedBy.join('・')}）` })]
                    : []),
                ],
                onChange: (on) => {
                  // 描き直さずに覚えておく（チェック欄はすでに切り替わっている）
                  const next = new Set(state.checked);
                  if (on) next.add(k.id);
                  else next.delete(k.id);
                  state = { ...state, checked: next };
                },
              }),
            ]),
          ),
        ),
      ]),
    ]);

    els.docView.replaceChildren(...compact([typeSection, filesSection, requiredSection, pointSection]));
    if (focusDocId) {
      const target = document.getElementById(focusDocId);
      if (target) {
        if (!(target instanceof HTMLInputElement) && !(target instanceof HTMLButtonElement) && !target.hasAttribute('tabindex')) {
          target.setAttribute('tabindex', '-1');
        }
        target.focus({ preventScroll: true });
      }
      focusDocId = null;
    }
  };

  /** 判定に含める・外すを切り替えた後、同じチェック欄に戻すための目印 */
  /** @type {string|null} */
  let focusUnitId = null;

  /** 切り替えたチェック欄にフォーカスを戻す（判定中は押せない状態なので、判定が終わってから） */
  const restoreUnitFocus = () => {
    if (!focusUnitId || busy) return;
    const box = document.getElementById(`unit-check-${focusUnitId}`);
    if (box) box.focus({ preventScroll: true });
    focusUnitId = null;
  };

  /**
   * ページ・シートの一覧の1行
   * @param {Unit} u
   * @param {ReadonlyMap<string, number>} counts
   * @param {FileEntry} file
   * @param {boolean} multi  複数のファイル（チェック欄の目印にファイルの目印を付ける）
   */
  const renderUnitRow = (u, counts, file, multi) => {
    const included = file.included.has(u.id);
    const unitKey = fileApi.unitKeyOf(file.key, u.id, multi);
    const status = !u.hasText
      ? '文字なし'
      : !included
        ? '判定から外しています'
        : state.result
          ? `指摘 ${counts.get(unitKey) ?? 0}件`
          : '';
    /** @type {HTMLInputElement|null} */
    let box = null;
    if (u.hasText) {
      box = document.createElement('input');
      box.type = 'checkbox';
      box.checked = included;
      box.disabled = busy;
      box.id = `unit-check-${unitKey}`;
      box.setAttribute('aria-label', `${multi ? `${file.doc.name}の` : ''}${fileApi.unitLabel(u)}を判定に含める`);
      box.addEventListener('change', () => toggleUnit(file.key, u.id));
    }
    return h(
      'li',
      {
        className: `unit-row${included && u.hasText ? '' : ' is-excluded'}`,
        attrs: { 'data-unit-id': unitKey, 'data-unit-label': u.label },
      },
      [
        h('span', { className: 'unit-name', text: fileApi.unitLabel(u) }),
        h('span', { className: `unit-status${!u.hasText ? ' is-empty' : ''}`, text: status }),
        box ? h('label', { className: 'unit-toggle', attrs: { for: box.id } }, [box, '判定に含める']) : null,
      ],
    );
  };

  /**
   * ページ・シートの一覧（Word は1つの単位なので出さない）
   * @param {FileEntry} file
   * @param {ReadonlyMap<string, number>} counts
   * @param {boolean} multi
   */
  const renderUnitList = (file, counts, multi) => {
    const word = unitWord(file.doc);
    return [
      h('p', {
        className: 'units-title',
        text: `${word}の一覧（${file.doc.units.length}${word}）: 別紙など、直す対象でない${word}は「判定に含める」を外せます`,
      }),
      h(
        'ul',
        { className: 'unit-list', attrs: multi ? { 'data-file-key': file.key } : { id: 'unit-list' } },
        file.doc.units.map((u) => renderUnitRow(u, counts, file, multi)),
      ),
    ];
  };

  /**
   * 複数のファイルの一覧の1つ（ファイル名・判別した書類の種類・指摘の件数・ページやシートの一覧・このファイルを外す）
   * @param {FileEntry} file
   * @param {ReadonlyMap<string, number>} unitCounts
   * @param {ReadonlyMap<string, number>} fileCounts
   */
  const renderFileBlock = (file, unitCounts, fileCounts) => {
    const part = state.prepared ? state.prepared.parts.find((p) => p.key === file.key) : undefined;
    const types = ruleset.documents ? ruleset.documents.types : [];
    const detected = part ? part.analysis.detected.map((id) => (types.find((t) => t.id === id) ?? { name: id }).name) : [];
    return h('li', { className: 'file-block', attrs: { 'data-file-key': file.key, 'data-file-name': file.doc.name } }, [
      h('p', { className: 'file-block-head' }, [
        h('span', { className: 'file-format', text: fileApi.FORMAT_LABELS[file.doc.format] }),
        h('span', { className: 'file-block-name', text: file.doc.name }),
      ]),
      h('p', { className: 'file-block-info' }, [
        part
          ? h('span', {
              className: 'file-block-type',
              text: detected.length > 0 ? `書類の種類: ${detected.join('／')}` : '書類の種類: 判別できませんでした（表記のルールだけで判定）',
            })
          : null,
        state.result ? h('span', { className: 'unit-status', text: `指摘 ${fileCounts.get(file.key) ?? 0}件` }) : null,
        h('button', {
          className: 'btn btn-small',
          text: 'このファイルを外す',
          attrs: { type: 'button', 'data-action': 'remove-file', 'aria-label': `${file.doc.name}を外す`, ...(busy ? { disabled: '' } : {}) },
          onClick: () => removeFile(file.key),
        }),
      ]),
      ...(file.doc.format !== 'docx' ? renderUnitList(file, unitCounts, isMulti()) : []),
      file.doc.notes.length > 0 ? h('ul', { className: 'file-notes' }, file.doc.notes.map((n) => h('li', { text: n }))) : null,
    ]);
  };

  /**
   * 読み込めなかったファイル（理由つき。付録F-3 AC-05-05）
   * @param {FailedFile} f
   */
  const renderFailedBlock = (f) =>
    h('li', { className: 'file-block is-failed', attrs: { 'data-file-key': f.key, 'data-file-name': f.name } }, [
      h('p', { className: 'file-block-head' }, [h('span', { className: 'file-block-name', text: f.name })]),
      h('p', { className: 'file-block-error', text: `読み込めませんでした: ${f.error.message}${f.error.detail ? `（詳しい理由: ${f.error.detail}）` : ''}` }),
      h('p', { className: 'file-block-info' }, [
        h('button', {
          className: 'btn btn-small',
          text: 'このファイルを外す',
          attrs: { type: 'button', 'data-action': 'remove-file', 'aria-label': `${f.name}を一覧から外す`, ...(busy ? { disabled: '' } : {}) },
          onClick: () => removeFile(f.key),
        }),
      ]),
    ]);

  /** ファイルを読み込んだときの表示（ファイル名・ページやシートの一覧・注意） */
  const renderFilePanel = () => {
    const files = state.files;
    const hasFiles = files.length > 0;
    els.filePanel.hidden = !hasFiles;
    els.pasteArea.hidden = hasFiles;
    els.bodyNote.textContent = hasFiles ? FILE_BODY_NOTE : PASTE_BODY_NOTE;
    if (!hasFiles) return;
    const prep = state.prepared;
    const unitCounts = prep && prep.composed && state.result ? fileApi.countByUnit(prep.composed, visible()) : new Map();

    if (files.length === 1 && state.failed.length === 0) {
      // 1つのファイル（S2 からの表示）
      const file = files[0];
      els.fileFormat.textContent = fileApi.FORMAT_LABELS[file.doc.format];
      els.fileName.textContent = file.doc.name;
      const showUnits = file.doc.format !== 'docx';
      els.fileUnits.hidden = !showUnits;
      if (showUnits) {
        els.fileUnits.replaceChildren(...renderUnitList(file, unitCounts, false));
        restoreUnitFocus();
      }
      els.fileNotes.hidden = file.doc.notes.length === 0;
      els.fileNotes.replaceChildren(...file.doc.notes.map((n) => h('li', { text: n })));
      return;
    }

    // 複数のファイル（読み込めなかったファイルも理由つきで並べる）
    const fileCounts = prep && state.result ? pipeline.countByFile(prep, visible()) : new Map();
    els.fileFormat.textContent = `${files.length + state.failed.length}ファイル`;
    els.fileName.textContent = files.length === 1 ? files[0].doc.name : 'まとめて判定しています（同じ種類のファイルは1つの書類として確かめます）';
    els.fileUnits.hidden = false;
    els.fileUnits.replaceChildren(
      h('ul', { className: 'file-list', attrs: { id: 'file-list' } }, [
        ...files.map((f) => renderFileBlock(f, unitCounts, fileCounts)),
        ...state.failed.map(renderFailedBlock),
      ]),
    );
    restoreUnitFocus();
    els.fileNotes.hidden = true;
    els.fileNotes.replaceChildren();
  };

  /** 読み取った収支表と、金額チェックの結果（付録E。収支予算書・決算書のときだけ） */
  const renderAmount = () => {
    const view = state.doc ? state.doc.evaluation.amount : null;
    if (!view || view.status === 'none') {
      els.amountArea.hidden = true;
      els.amountArea.replaceChildren();
      return;
    }
    els.amountArea.hidden = false;
    const title = h('h2', { className: 'pane-title', attrs: { id: 'amount-title' }, text: '金額チェック（収支予算書・決算書）' });
    if (view.status !== 'checked') {
      const text = view.status === 'needs-excel' ? pipeline.MESSAGES.needsExcel : pipeline.MESSAGES.unavailable;
      els.amountArea.replaceChildren(title, h('p', { className: 'amount-message', attrs: { id: 'amount-message' }, text }));
      return;
    }
    const shown = visible();
    const multi = isMulti();
    const blocks = view.sheets.map((s) => {
      const name = multi && s.sheet ? `${s.fileName}／${s.sheet}` : s.sheet || s.fileName;
      if (!s.ok || !s.table) {
        return h('div', { className: 'amount-sheet is-error', attrs: { 'data-table-id': s.id } }, [
          h('p', { className: 'amount-sheet-name', text: name }),
          h('p', { className: 'amount-message is-error', text: pipeline.MESSAGES.notRead(s.reasons) }),
        ]);
      }
      const findings = shown.filter((f) => f.amount && f.amount.tableId === s.id);
      const flagged = new Set(findings.flatMap((f) => (f.amount ? [f.amount.rowId] : [])));
      const showSettlement = s.table.showSettlement;
      return h('div', { className: 'amount-sheet', attrs: { 'data-table-id': s.id } }, [
        h('p', {
          className: `amount-message${findings.length > 0 ? ' is-short' : ' is-ok'}`,
          text:
            findings.length === 0
              ? pipeline.MESSAGES.noProblem
              : `金額チェック: ${report.formatCounts(report.countBySeverity(findings))}（指摘一覧の「金額チェック」にあります）`,
        }),
        h('div', { className: 'amount-table-wrap', attrs: { tabindex: '0', role: 'region', 'aria-label': `読み取った収支表（${name}）` } }, [
          h('table', { className: 'amount-table' }, [
            h('caption', { text: `読み取った収支表（${name}・${view.kind}（案））` }),
            h('thead', {}, [
              h('tr', {}, [
                h('th', { text: '区分', attrs: { scope: 'col' } }),
                h('th', { text: '項目', attrs: { scope: 'col' } }),
                h('th', { text: '予算額', attrs: { scope: 'col' } }),
                showSettlement ? h('th', { text: '決算額', attrs: { scope: 'col' } }) : null,
                h('th', { text: '摘要', attrs: { scope: 'col' } }),
              ]),
            ]),
            h(
              'tbody',
              {},
              s.table.rows.map((r) =>
                h(
                  'tr',
                  {
                    className: `${r.total ? 'is-total' : ''}${flagged.has(r.id) ? ' is-flagged' : ''}`.trim(),
                    attrs: { 'data-row-id': r.id },
                  },
                  [
                    h('td', { className: 'amount-group', text: r.group }),
                    h('th', { text: r.label, attrs: { scope: 'row' } }),
                    h('td', { className: 'amount-num', text: r.budget }),
                    showSettlement ? h('td', { className: 'amount-num', text: r.settlement }) : null,
                    h('td', { className: 'amount-note', text: r.note }),
                  ],
                ),
              ),
            ),
          ]),
        ]),
      ]);
    });
    els.amountArea.replaceChildren(title, ...blocks);
  };

  const render = () => {
    renderFilePanel();
    els.results.hidden = state.result === null;
    if (state.result === null) return;
    renderSummary();
    renderAmount();
    renderList();
    renderBody();
    renderDoc();
    renderTabs();
  };

  // ------------------------------------------------------------------
  // 操作
  // ------------------------------------------------------------------

  /** 読み込んだファイルの、判定に含めるページ・シートをつないだもの */
  const composeAll = () => fileApi.composeFiles(state.files.map((f) => ({ key: f.key, units: f.doc.units, included: f.included }))).composed;

  const updateCharCount = () => {
    if (state.files.length > 0) {
      // ファイルのときは、判定に含めたページ・シートの文字数
      const n = state.result ? state.result.charCount : engine.countChars(composeAll().text);
      els.charCount.textContent = `読み込んだ文字数: ${n.toLocaleString('ja-JP')}字`;
      els.charCount.classList.toggle('is-over', n > engine.MAX_CHARS);
      return;
    }
    const n = engine.countChars(els.input.value);
    const over = n > engine.MAX_CHARS;
    els.charCount.textContent = `入力文字数: ${n.toLocaleString('ja-JP')}字${over ? '（10万字を超えています）' : ''}`;
    els.charCount.classList.toggle('is-over', over);
    const stale = document.getElementById('summary-stale');
    if (stale) stale.hidden = !isStale();
  };

  /** @param {boolean} on */
  const setBusy = (on) => {
    busy = on;
    els.btnCheck.disabled = on || Boolean(ruleset.fatal);
    els.btnSample.disabled = on || Boolean(ruleset.fatal);
    els.btnClear.disabled = on;
    els.btnFile.disabled = on || Boolean(ruleset.fatal);
    els.btnToPaste.disabled = on;
    els.btnCheck.textContent = on ? '判定中…' : 'チェックする';
    for (const box of els.fileUnits.querySelectorAll('input[type="checkbox"], button[data-action="remove-file"]')) {
      /** @type {HTMLInputElement|HTMLButtonElement} */ (box).disabled = on;
    }
    for (const box of els.docView.querySelectorAll('.type-choices input, .budget-kind input')) {
      /** @type {HTMLInputElement} */ (box).disabled = on;
    }
    restoreUnitFocus();
  };

  const closeSheet = () => {
    els.sheet.hidden = true;
    els.sheetBody.replaceChildren();
  };

  /** @param {number} charCount */
  const tooLongMessage = (charCount) =>
    `長すぎます（10万字まで）。いまは${charCount.toLocaleString('ja-JP')}字です。文章を分けてチェックしてください。`;

  /**
   * 書類の種類（自動で判別したもの・選び直したもの）に応じた、書類全体の指摘・金額チェック・必須項目の表
   * @param {Prepared} prepared
   * @param {Selection} selection
   * @returns {DocState}
   */
  const makeDocState = (prepared, selection) => ({ selection, evaluation: pipeline.evaluate(prepared, selection, ruleset) });

  /**
   * 書類の種類を選び直す（文章の判定はやり直さず、書類全体の指摘・金額チェック・必須項目の表だけを出し直す）
   * @param {Selection} selection
   * @param {string} focusId  描き直した後にフォーカスを戻す部品
   */
  const changeSelection = (selection, focusId) => {
    const prepared = state.prepared;
    if (!state.doc || !prepared || busy) return;
    focusDocId = focusId;
    closeSheet();
    setState({ doc: makeDocState(prepared, selection) });
  };

  /** 必須項目と確認ポイントを見せる（スマホではタブを切り替える） */
  const showDocPane = () => {
    closeSheet();
    if (narrowQuery.matches) {
      setTab('doc');
      els.tabDoc.scrollIntoView({ block: 'start' });
    } else {
      els.paneDoc.scrollIntoView({ block: 'start' });
    }
    els.paneDoc.focus({ preventScroll: true });
  };

  /**
   * @typedef {Object} CheckJob
   * @property {{ kind: 'paste', text: string } | { kind: 'files', files: readonly { key: string, name: string, doc: LoadedDoc, included: ReadonlySet<string> }[] }} input
   * @property {boolean} scroll
   * @property {boolean} keepSelection  手動で選んだ書類の種類をそのまま使う（ページ・シートの切り替え・ファイルを外したとき）
   */

  /**
   * 「判定中」を画面に出してから判定し、結果を出す
   * @param {CheckJob} job
   */
  const startCheck = (job) => {
    setBusy(true);
    showStatus('判定中…', 'busy');
    window.requestAnimationFrame(() => {
      window.setTimeout(() => {
        try {
          const prepared = pipeline.prepare(job.input, ruleset);
          if (!prepared.ok) {
            setState({ result: null, checkedText: '', prepared: null, doc: null });
            showStatus(
              prepared.reason === 'empty' ? '文章を貼り付けてください' : tooLongMessage(prepared.charCount),
              prepared.reason === 'empty' ? 'warn' : 'error',
            );
            return;
          }
          const prev = state.doc;
          const selection =
            job.keepSelection && prev && !prev.selection.auto ? prev.selection : pipeline.autoSelection(prepared, ruleset);
          setState({
            result: prepared.result,
            checkedText: job.input.kind === 'paste' ? job.input.text : '',
            prepared,
            doc: makeDocState(prepared, selection),
            tab: 'list',
            showIgnored: false,
          });
          hideStatus();
          if (job.scroll) els.summary.scrollIntoView({ block: 'start' });
        } catch (e) {
          const detail = e instanceof Error ? e.message : String(e);
          showStatus(`判定の途中で問題が起きました（${detail}）。ページを再読み込みして、もう一度お試しください。`, 'error');
          console.error(e);
        } finally {
          setBusy(false);
          updateCharCount();
        }
      }, 0);
    });
  };

  /** 貼り付けた文章を判定する */
  const runPasteCheck = () => {
    const text = els.input.value;
    if (text.trim() === '') {
      setState({ result: null, checkedText: '', prepared: null, doc: null });
      showStatus('文章を貼り付けてください', 'warn');
      return;
    }
    const n = engine.countChars(text);
    if (n > engine.MAX_CHARS) {
      setState({ result: null, checkedText: '', prepared: null, doc: null });
      showStatus(tooLongMessage(n), 'error');
      return;
    }
    startCheck({ input: { kind: 'paste', text }, scroll: true, keepSelection: false });
  };

  /**
   * 読み込んだファイルの、判定に含めるページ・シートを判定する
   * @param {boolean} scroll  判定後に結果のまとめまで動かす
   * @param {boolean} keepSelection  手動で選んだ書類の種類をそのまま使う
   */
  const runFileCheck = (scroll, keepSelection) => {
    const files = state.files;
    if (files.length === 0) return;
    const single = files.length === 1 ? files[0] : null;
    const word = single ? unitWord(single.doc) : 'ページ・シート';
    const composed = composeAll();
    if (composed.text.trim() === '') {
      setState({ result: null, checkedText: '', prepared: null, doc: null });
      showStatus(`判定に含める${word}がありません。一覧の「判定に含める」にチェックを入れてください。`, 'warn');
      updateCharCount();
      return;
    }
    const n = engine.countChars(composed.text);
    if (n > engine.MAX_CHARS) {
      setState({ result: null, checkedText: '', prepared: null, doc: null });
      showStatus(`${tooLongMessage(n)}別紙の${word}${single ? '' : '・ファイル'}を判定から外すと、短くできます。`, 'error');
      updateCharCount();
      return;
    }
    startCheck({
      input: { kind: 'files', files: files.map((f) => ({ key: f.key, name: f.doc.name, doc: f.doc, included: f.included })) },
      scroll,
      keepSelection,
    });
  };

  const runCheck = () => {
    if (busy) return;
    closeSheet();
    if (state.files.length > 0) runFileCheck(true, false);
    else runPasteCheck();
  };

  /**
   * ページ・シートを判定に含める・外す
   * @param {string} fileKey
   * @param {string} unitId
   */
  const toggleUnit = (fileKey, unitId) => {
    const file = state.files.find((f) => f.key === fileKey);
    if (!file || busy) return;
    const included = new Set(file.included);
    if (included.has(unitId)) included.delete(unitId);
    else included.add(unitId);
    focusUnitId = fileApi.unitKeyOf(fileKey, unitId, isMulti());
    closeSheet();
    // 画面の描き直しは判定の後にまとめて行う（先に描き直すと、押したチェック欄からフォーカスが外れるため）
    state = { ...state, files: state.files.map((f) => (f.key === fileKey ? { ...f, included } : f)) };
    runFileCheck(false, true);
  };

  /**
   * 一覧からファイルを1つ外し、残りで判定し直す（付録F-3 AC-05-07）
   * @param {string} key
   */
  const removeFile = (key) => {
    if (busy) return;
    closeSheet();
    const failed = state.failed.filter((f) => f.key !== key);
    if (!state.files.some((f) => f.key === key)) {
      // 読み込めなかったファイルを一覧から外すだけ
      setState({ failed });
      return;
    }
    const files = state.files.filter((f) => f.key !== key);
    if (files.length === 0) {
      clearAll();
      return;
    }
    state = { ...state, files, failed };
    showToast('ファイルを外しました。残りのファイルで判定し直します');
    runFileCheck(false, true);
  };

  /** ファイルから読み込んだ文章を貼り付け欄に移し、貼り付けた文章として判定し直す */
  const switchToPaste = () => {
    if (state.files.length === 0 || busy) return;
    els.input.value = composeAll().text;
    closeSheet();
    hideStatus();
    setState({ files: [], failed: [], prepared: null, result: null, checkedText: '', doc: null });
    updateCharCount();
    runPasteCheck();
    showToast('貼り付け欄に切り替えました（ページ・セルの場所は出なくなります）');
  };

  const clearAll = () => {
    els.input.value = '';
    els.fileInput.value = '';
    closeSheet();
    hideStatus();
    state = INITIAL_STATE;
    render();
    updateCharCount();
    window.scrollTo(0, 0);
    els.input.focus();
  };

  // ------------------------------------------------------------------
  // ファイルの読み込み
  // ------------------------------------------------------------------

  /** ページの読み込みが終わる（同梱の pdf.js の読み込みが済む）まで待つ */
  const pageLoaded = () =>
    document.readyState === 'complete'
      ? Promise.resolve()
      : new Promise((resolve) => window.addEventListener('load', () => resolve(undefined), { once: true }));

  /**
   * @param {unknown} v
   * @returns {v is { getDocument: (src: Record<string, unknown>) => unknown }}
   */
  const isPdfLib = (v) =>
    typeof v === 'object' && v !== null && typeof (/** @type {{ getDocument?: unknown }} */ (v).getDocument) === 'function';

  /**
   * PDF のページごとの文字を読む（進み具合を表示する）
   * @param {Uint8Array} bytes
   * @returns {Promise<string[]>}
   */
  const readPdfBytes = async (bytes) => {
    await pageLoaded();
    const lib = g.pdfjsLib;
    if (!isPdfLib(lib) || typeof g.pdfjsWorker !== 'object' || g.pdfjsWorker === null) {
      throw new fileApi.FileLoadError('pdf-unavailable', 'pdf.js が読み込まれていません');
    }
    return pdfApi.readPdf(bytes, {
      lib,
      getCmaps: () => {
        const t = g.YEG_PDF_CMAPS;
        return typeof t === 'object' && t !== null ? /** @type {Readonly<Record<string, string>>} */ (t) : null;
      },
      onProgress: (done, total) => showStatus(`${progressPrefix}読み込み中 ${done}／${total}ページ`, 'busy'),
    });
  };

  /**
   * @param {string} name
   * @param {LoadError} error
   */
  const showFileError = (name, error) => {
    const detail = error.detail ? `（詳しい理由: ${error.detail}）` : '';
    showStatus(`「${name}」: ${error.message}${detail}`, 'error');
  };

  /** 画面に「読み込み中」を描いてから続ける */
  const nextPaint = () =>
    new Promise((resolve) => window.requestAnimationFrame(() => window.setTimeout(() => resolve(undefined), 0)));

  /**
   * ファイルを1つ読み込む（判定はしない）
   * @param {File} file
   * @returns {Promise<{ ok: true, doc: LoadedDoc } | { ok: false, error: LoadError }>}
   */
  const loadOne = async (file) => {
    const sizeError = fileApi.checkSize(file.size);
    if (sizeError) return { ok: false, error: sizeError };
    try {
      await nextPaint();
      const bytes = new Uint8Array(await file.arrayBuffer());
      return await fileApi.loadFile({ name: file.name, bytes }, { readPdf: readPdfBytes });
    } catch (e) {
      const detail = e instanceof Error ? e.message : String(e);
      console.error(e);
      return { ok: false, error: { code: 'broken', message: 'ファイルを読み込めませんでした。もう一度お試しください。', detail } };
    }
  };

  /**
   * ファイルを読み込み（複数のときは1つずつ順に、進み具合を出しながら）、まとめて判定する。
   * 読めないファイルがあれば、そのファイルだけ理由つきで一覧に出し、残りで判定する（付録F-3）
   * @param {readonly File[]} list
   */
  const loadFiles = async (list) => {
    closeSheet();
    setBusy(true);
    /** @type {FileEntry[]} */
    const loaded = [];
    /** @type {FailedFile[]} */
    const failed = [];
    try {
      for (const [i, file] of list.entries()) {
        progressPrefix = list.length > 1 ? `${i + 1}／${list.length}ファイル目「${file.name}」: ` : '';
        showStatus(list.length > 1 ? `読み込み中 ${i + 1}／${list.length}ファイル（「${file.name}」）…` : `「${file.name}」を読み込み中…`, 'busy');
        fileSeq += 1;
        const key = `f${fileSeq}`;
        const r = await loadOne(file);
        if (r.ok) loaded.push({ key, doc: r.doc, included: fileApi.defaultIncluded(r.doc.units) });
        else failed.push({ key, name: file.name, error: r.error });
      }
    } finally {
      progressPrefix = '';
      setBusy(false);
    }
    if (loaded.length === 0) {
      // どれも読み込めなかったときは、前の状態のままにして理由を出す
      if (failed.length === 1) showFileError(failed[0].name, failed[0].error);
      else {
        const lines = failed.map((f) => `「${f.name}」: ${f.error.message}${f.error.detail ? `（詳しい理由: ${f.error.detail}）` : ''}`);
        showStatus(`ファイルを読み込めませんでした。${lines.join(' ')}`, 'error');
      }
      return;
    }
    // 前のファイル・貼り付けた文章と結果は、新しいファイルで置き換える
    els.input.value = '';
    state = {
      ...state,
      files: loaded,
      failed,
      result: null,
      prepared: null,
      doc: null,
      checked: new Set(),
      checkedText: '',
      filter: 'all',
    };
    render();
    runFileCheck(true, false);
    if (failed.length > 0) showToast(`読み込めなかったファイルが${failed.length}件あります（ファイルの一覧に理由があります）`, true);
  };

  /**
   * ドラッグ＆ドロップ・ファイルの選択で渡されたファイル（10ファイルまで。まとめて1つの議案として判定する）
   * @param {ArrayLike<File>|null|undefined} files
   */
  const readFiles = (files) => {
    if (!files || files.length === 0) return;
    if (busy) {
      showToast('読み込み中です。終わってから、もう一度お試しください', true);
      return;
    }
    const list = Array.from(files);
    if (list.length > fileApi.MAX_FILES) {
      showStatus(fileApi.TOO_MANY_FILES, 'error');
      return;
    }
    loadFiles(list).catch((e) => {
      setBusy(false);
      showStatus(`ファイルを読み込めませんでした（${e instanceof Error ? e.message : String(e)}）。`, 'error');
      console.error(e);
    });
  };

  /** @param {string} key */
  const ignore = (key) => {
    closeSheet();
    setState({ ignored: new Set([...state.ignored, key]) });
    showToast('指摘を無視しました（「無視した指摘を表示」から取り消せます）');
  };

  /** @param {string} key */
  const unignore = (key) => {
    setState({ ignored: new Set([...state.ignored].filter((k) => k !== key)) });
  };

  /** @param {Tab} tab */
  const setTab = (tab) => {
    state = { ...state, tab };
    renderTabs();
  };

  /**
   * @param {Element[]} nodes
   */
  const flash = (nodes) => {
    for (const n of nodes) {
      n.classList.remove('is-flash');
      void (/** @type {HTMLElement} */ (n).offsetWidth);
      n.classList.add('is-flash');
    }
    window.setTimeout(() => nodes.forEach((n) => n.classList.remove('is-flash')), 2000);
  };

  /** @param {string} id */
  const showInBody = (id) => {
    closeSheet();
    if (narrowQuery.matches) setTab('body');
    window.requestAnimationFrame(() => {
      const marks = [...els.bodyView.querySelectorAll(`mark[data-ids~="${id}"]`)];
      if (marks.length === 0) {
        showToast('本文の中に該当箇所が見つかりませんでした', true);
        return;
      }
      const first = /** @type {HTMLElement} */ (marks[0]);
      first.scrollIntoView({ block: 'center' });
      first.focus({ preventScroll: true });
      flash(marks);
    });
  };

  /** @param {string} id */
  const showInList = (id) => {
    closeSheet();
    if (narrowQuery.matches) setTab('list');
    window.requestAnimationFrame(() => {
      const card = document.getElementById(`card-${id}`);
      if (!card) {
        showToast('一覧の絞り込みで隠れています。「すべて」を選ぶと表示されます', true);
        return;
      }
      card.scrollIntoView({ block: 'center' });
      card.classList.add('is-target');
      window.setTimeout(() => card.classList.remove('is-target'), 2000);
    });
  };

  /**
   * 読み込んだ収支表の該当行を見せて強調する（金額チェックのカードを押したとき）
   * @param {string} tableId
   * @param {string} rowId
   */
  const showInTable = (tableId, rowId) => {
    closeSheet();
    const block = [...els.amountArea.querySelectorAll('[data-table-id]')].find((n) => n.getAttribute('data-table-id') === tableId);
    const row = block ? [...block.querySelectorAll('tr[data-row-id]')].find((n) => n.getAttribute('data-row-id') === rowId) : undefined;
    const target = row ?? block;
    if (!target) {
      showToast('読み取った収支表の中に、該当する行が見つかりませんでした', true);
      return;
    }
    window.requestAnimationFrame(() => {
      target.scrollIntoView({ block: 'center' });
      for (const r of els.amountArea.querySelectorAll('.is-target')) r.classList.remove('is-target');
      target.classList.add('is-target');
      flash([target]);
    });
  };

  /** @param {string[]} ids */
  const openSheet = (ids) => {
    /** @type {Finding[]} */
    const list = ids.flatMap((id) => allFindings().filter((f) => f.id === id));
    if (list.length === 0) return;
    els.sheetBody.replaceChildren(...list.map((f) => renderCard(f, { ignored: false, inSheet: true })));
    els.sheet.hidden = false;
    els.sheet.scrollTop = 0;
  };

  /** @param {Event} e */
  const onBodyActivate = (e) => {
    const target = e.target instanceof Element ? e.target.closest('mark[data-ids]') : null;
    if (!target) return;
    const ids = (target.getAttribute('data-ids') ?? '').split(' ').filter((s) => s !== '');
    openSheet(ids);
  };

  const copyList = () => {
    const result = state.result;
    if (!result) {
      showToast('先に「チェックする」を押してください', true);
      return;
    }
    const shown = visible();
    const doc = state.doc;
    const ev = doc ? doc.evaluation : null;
    /** 複数のファイルのとき、種類ごとに使ったファイル */
    const typeFileLines =
      ev && isMulti()
        ? ev.report.tables.map((t) => `${t.title}: ${filesOfType(t.typeId).join('、')}`)
        : [];
    const amountLine = amountSummary(shown).text;
    const text = report.buildCopyText({
      findings: shown,
      ignoredCount: allFindings().length - shown.length,
      charCount: result.charCount,
      styleCounts: result.styleCounts,
      ruleset,
      now: new Date(),
      locate: placeOf,
      docLines: [
        ...(doc && ev && ev.analysis.available ? docCheck.copyLines(doc.selection, ev.report, ruleset) : []),
        ...typeFileLines,
        ...(amountLine ? [amountLine] : []),
        ...(state.failed.length > 0 ? [`読み込めなかったファイル: ${state.failed.map((f) => f.name).join('、')}`] : []),
      ],
      ...(state.files.length > 0 ? { source: { fileName: fileNames(), scope: includedScope() } } : {}),
    });
    copyWithToast(text, 'コピーしました');
  };

  // ------------------------------------------------------------------
  // つなぎこみ
  // ------------------------------------------------------------------

  els.btnCheck.addEventListener('click', runCheck);
  els.btnSample.addEventListener('click', () => {
    if (busy) return;
    // ファイルを読み込んでいるときは、貼り付けに戻して例文を入れる
    if (state.files.length > 0) setState({ files: [], failed: [], prepared: null, result: null, checkedText: '', doc: null });
    els.input.value = sampleText;
    updateCharCount();
    runCheck();
  });
  els.btnClear.addEventListener('click', clearAll);
  els.input.addEventListener('input', updateCharCount);

  els.btnFile.addEventListener('click', () => {
    if (busy) return;
    els.fileInput.click();
  });
  els.fileInput.addEventListener('change', () => {
    const files = els.fileInput.files;
    const picked = files && files.length > 0 ? [...files] : [];
    // 同じファイルをもう一度選んでも読み込めるよう、選んだ状態を空に戻す
    els.fileInput.value = '';
    if (picked.length > 0) readFiles(picked);
  });
  els.btnToPaste.addEventListener('click', switchToPaste);

  // ページのどこにファイルを落としても読み込む（文字のドラッグはそのまま入力欄に任せる）
  let dragDepth = 0;
  /** @param {DragEvent} e */
  const hasFiles = (e) => Boolean(e.dataTransfer) && [...(/** @type {DataTransfer} */ (e.dataTransfer).types)].includes('Files');
  document.addEventListener('dragenter', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth += 1;
    els.dropOverlay.hidden = false;
  });
  document.addEventListener('dragover', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    if (e.dataTransfer) e.dataTransfer.dropEffect = busy ? 'none' : 'copy';
  });
  document.addEventListener('dragleave', (e) => {
    if (!hasFiles(e)) return;
    dragDepth = Math.max(0, dragDepth - 1);
    if (dragDepth === 0) els.dropOverlay.hidden = true;
  });
  document.addEventListener('drop', (e) => {
    if (!hasFiles(e)) return;
    e.preventDefault();
    dragDepth = 0;
    els.dropOverlay.hidden = true;
    readFiles(e.dataTransfer ? e.dataTransfer.files : null);
  });

  els.tabList.addEventListener('click', () => setTab('list'));
  els.tabBody.addEventListener('click', () => setTab('body'));
  els.tabDoc.addEventListener('click', () => setTab('doc'));
  els.groupRule.addEventListener('click', () => setState({ groupMode: 'rule' }));
  els.groupPosition.addEventListener('click', () => setState({ groupMode: 'position' }));
  for (const btn of els.panes.querySelectorAll('[data-filter]')) {
    btn.addEventListener('click', () => {
      const f = btn.getAttribute('data-filter');
      const filter = /** @type {SeverityFilter} */ (
        f === '違反' || f === '要確認' || f === '参考' ? f : 'all'
      );
      setState({ filter });
    });
  }
  els.btnToggleIgnored.addEventListener('click', () => setState({ showIgnored: !state.showIgnored }));
  els.btnCopyList.addEventListener('click', copyList);

  els.bodyView.addEventListener('click', onBodyActivate);
  els.bodyView.addEventListener('keydown', (e) => {
    if (e.key === 'Enter' || e.key === ' ') {
      e.preventDefault();
      onBodyActivate(e);
    }
  });
  els.btnSheetClose.addEventListener('click', closeSheet);
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && !els.sheet.hidden) closeSheet();
  });

  const openRules = () => {
    if (typeof els.rulesDialog.showModal === 'function') els.rulesDialog.showModal();
    else els.rulesDialog.setAttribute('open', '');
  };
  const closeRules = () => {
    if (typeof els.rulesDialog.close === 'function') els.rulesDialog.close();
    else els.rulesDialog.removeAttribute('open');
  };
  els.btnRulesTop.addEventListener('click', openRules);
  els.btnRulesBottom.addEventListener('click', openRules);
  els.btnRulesClose.addEventListener('click', closeRules);
  els.rulesDialog.addEventListener('click', (e) => {
    if (e.target === els.rulesDialog) closeRules();
  });

  els.btnCopyPrompt.addEventListener('click', () =>
    copyWithToast(report.buildAiPrompt(ruleset), 'AIへの依頼文をコピーしました'),
  );

  // ブラウザが前回の入力を復元しないよう、開いたときは必ず空にする
  window.addEventListener('pageshow', (e) => {
    if (e.persisted) clearAll();
  });
  els.input.value = '';

  renderLoadInfo();
  renderRuleDialog();
  renderAiGuide();
  updateCharCount();
  render();
})();
