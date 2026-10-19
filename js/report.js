// @ts-check
/*
 * 判定結果の整理と文章づくり（画面に依存しない純粋な関数だけを置く）
 *   件数のまとめ・並べ替え・絞り込み・指摘一覧のコピー用の文章・AIへの依頼文・ルール一覧
 */
'use strict';

(function (root) {
  /** @typedef {'違反'|'要確認'|'参考'} Severity */
  /** @typedef {Severity|'all'} SeverityFilter */
  /** @typedef {'rule'|'position'} GroupMode */

  /**
   * engine.js の型（実行時は YegEngine が作ったものを受け取る）
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
   *
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
   * @property {string|null} [section]  指摘のある欄の名前（S3。欄が分からないときは null）
   * @property {{ location: string, calc: string }} [amount]  金額チェックの指摘（S4）: 場所（シート・セル・行の項目名）と計算に使った数値
   *
   * @typedef {Object} StyleCounts
   * @property {number} polite
   * @property {number} plain
   * @property {string} politeLabel
   * @property {string} plainLabel
   *
   * @typedef {Object} RuleInfo
   * @property {string} id
   * @property {string} name
   * @property {string} category
   * @property {Severity} severity
   * @property {string[]} examples
   * @property {string[]} fixes
   * @property {string} description
   * @property {string} source
   * @property {string} scope
   * @property {number} order
   *
   * @typedef {Object} RulesetView
   * @property {{ toolName: string, manualName: string, manualVersion: string, sourceLabel: string, updated: string }} meta
   * @property {string[]} categories
   * @property {string[]} notChecked
   * @property {{ caution: string, request: string, links: { label: string, url: string }[] }} aiGuide
   * @property {RuleInfo[]} rules
   * @property {Readonly<Record<string, string>>} [categoryNotes]  分類の見出しの下に出す説明
   */

  /**
   * @typedef {Object} FindingGroup
   * @property {string} key
   * @property {string|null} title  null なら見出しを出さない
   * @property {Severity|null} severity
   * @property {Finding[]} findings
   */

  /** @type {readonly Severity[]} */
  const SEVERITIES = Object.freeze(['違反', '要確認', '参考']);

  /** @type {Readonly<Record<Severity, string>>} */
  const SEVERITY_CLASS = Object.freeze({ 違反: 'violation', 要確認: 'check', 参考: 'info' });

  const WHOLE_DOCUMENT = '書類全体';
  const AMOUNT_GROUP = '金額チェック';

  /** @param {Severity} s */
  const severityRank = (s) => SEVERITIES.indexOf(s);

  /** @param {Severity} s */
  const severityClass = (s) => SEVERITY_CLASS[s];

  /**
   * @param {Finding[]} findings
   * @param {ReadonlySet<string>} ignoredKeys
   */
  const visibleFindings = (findings, ignoredKeys) => findings.filter((f) => !ignoredKeys.has(f.key));

  /**
   * @param {Finding[]} findings
   * @param {ReadonlySet<string>} ignoredKeys
   */
  const ignoredFindings = (findings, ignoredKeys) => findings.filter((f) => ignoredKeys.has(f.key));

  /**
   * @param {Finding[]} findings
   * @returns {Record<Severity, number>}
   */
  const countBySeverity = (findings) => ({
    違反: findings.filter((f) => f.severity === '違反').length,
    要確認: findings.filter((f) => f.severity === '要確認').length,
    参考: findings.filter((f) => f.severity === '参考').length,
  });

  /**
   * 例: 「違反 5件／要確認 2件／参考 1件」（0件の重さは省く）
   * @param {Record<Severity, number>} counts
   */
  const formatCounts = (counts) => {
    const parts = SEVERITIES.filter((s) => counts[s] > 0).map((s) => `${s} ${counts[s]}件`);
    return parts.length > 0 ? parts.join('／') : '指摘 0件';
  };

  /**
   * @param {Finding[]} findings
   * @param {SeverityFilter} filter
   */
  const filterBySeverity = (findings, filter) =>
    filter === 'all' ? findings : findings.filter((f) => f.severity === filter);

  /** @param {Finding} a @param {Finding} b */
  const byPosition = (a, b) => {
    const as = a.range ? a.range.start : -1;
    const bs = b.range ? b.range.start : -1;
    return as - bs || a.reasons[0].order - b.reasons[0].order;
  };

  /**
   * 一覧の並べ方。位置を持たない指摘は先頭に「書類全体」として、金額チェックの指摘はその次に「金額チェック」としてまとめる
   * @param {Finding[]} findings
   * @param {GroupMode} mode
   * @returns {FindingGroup[]}
   */
  const groupFindings = (findings, mode) => {
    const amounts = findings.filter((f) => f.kind === 'amount');
    const rest = findings.filter((f) => f.kind !== 'amount');
    const whole = rest.filter((f) => f.range === null);
    const placed = rest.filter((f) => f.range !== null).sort(byPosition);
    /** @type {FindingGroup[]} */
    const head = [
      ...(whole.length > 0 ? [{ key: 'document', title: WHOLE_DOCUMENT, severity: null, findings: whole }] : []),
      ...(amounts.length > 0 ? [{ key: 'amount', title: `${AMOUNT_GROUP}（${amounts.length}件）`, severity: null, findings: amounts }] : []),
    ];
    if (mode === 'position') {
      return placed.length > 0 ? [...head, { key: 'position', title: null, severity: null, findings: placed }] : head;
    }
    /** @type {Map<string, Finding[]>} */
    const byRule = new Map();
    for (const f of placed) {
      const id = f.reasons[0].ruleId;
      byRule.set(id, [...(byRule.get(id) ?? []), f]);
    }
    const groups = [...byRule.entries()].map(([id, list]) => {
      const r = list[0].reasons[0];
      return {
        key: `rule-${id}`,
        title: `${r.ruleName}（${r.severity}）${list.length}件`,
        severity: r.severity,
        order: r.order,
        findings: list,
      };
    });
    groups.sort(
      (a, b) =>
        severityRank(/** @type {Severity} */ (a.severity)) - severityRank(/** @type {Severity} */ (b.severity)) ||
        a.order - b.order,
    );
    return [...head, ...groups.map(({ key, title, severity, findings: f }) => ({ key, title, severity, findings: f }))];
  };

  /**
   * 例: 「→「および」」「→「A」または「B」」
   * @param {string[]} fixes
   */
  const formatFix = (fixes) =>
    fixes.length === 0 ? '' : `→${fixes.map((f) => `「${f}」`).join('または')}`;

  /**
   * @param {RulesetView} ruleset
   * @param {string} source
   */
  const formatSource = (ruleset, source) =>
    // 「【虎の巻（委員会事業計画版）】…」のように【】で始まる出典は、マニュアル以外の資料なのでそのまま出す
    source.startsWith('【') ? source : `${ruleset.meta.sourceLabel}${source}`;

  /** @param {StyleCounts} c */
  const formatStyleCounts = (c) => `${c.politeLabel} ${c.polite}文／${c.plainLabel} ${c.plain}文`;

  /** @param {Finding[]} findings */
  const hasStyleFinding = (findings) => findings.some((f) => f.kind === 'sentence');

  /**
   * 指摘した語を1行で見せる（改行は取り除き、タブは空白にする）
   * @param {string} s
   */
  const oneLine = (s) => s.replace(/\n/g, '').replace(/\t/g, ' ');

  /**
   * 前後の文脈を1行で見せる（改行・タブは空白にする。タブはファイルの表のセルの区切り）
   * @param {string} s
   */
  const contextLine = (s) => s.replace(/[\n\t]/g, ' ');

  /** @param {number} n */
  const pad2 = (n) => String(n).padStart(2, '0');

  /**
   * @param {Date} date
   */
  const formatDateTime = (date) => {
    const engine = root.YegEngine;
    const w = engine.formatWareki({ y: date.getFullYear(), m: date.getMonth() + 1, d: date.getDate() });
    return `${w ?? `${date.getFullYear()}年${date.getMonth() + 1}月${date.getDate()}日`} ${pad2(date.getHours())}:${pad2(date.getMinutes())}`;
  };

  /**
   * 指摘1件を1行の文章にする
   * 例: 「[違反] 接続詞はひらがな: 「及び」→「および」」、場所があれば「…→「および」（2ページ）」
   * @param {Finding} f
   * @param {string} [location]  ページ・シートとセル番地など（ファイルを読み込んだとき）
   */
  const findingLine = (f, location = '') => {
    // 金額チェックの指摘は、指摘の文と場所（シート・セル・行の項目名）と計算を出す
    if (f.kind === 'amount') {
      const r = f.reasons[0];
      const calc = f.amount && f.amount.calc ? `［計算: ${f.amount.calc}］` : '';
      return `[${f.severity}] ${r.ruleName}: ${r.note}${calc}${location ? `（${location}）` : ''}`;
    }
    // 書類全体の指摘（野外事業の中止基準など）は、語の代わりに指摘の文を出す
    if (f.kind === 'document' || f.range === null) {
      const r = f.reasons[0];
      return `[${f.severity}] ${r.ruleName}: ${r.note || r.description}（${WHOLE_DOCUMENT}）`;
    }
    const text = oneLine(f.text);
    const parts = f.reasons.map((r) => {
      const sev = r.severity === f.severity ? '' : `（${r.severity}）`;
      const fix = r.fixes.length > 0 ? formatFix(r.fixes) : r.hint ? `（${r.hint}）` : '';
      return `${r.ruleName}${sev}: 「${text}」${fix}`;
    });
    return `[${f.severity}] ${parts.join(' ／ ')}${location ? `（${location}）` : ''}`;
  };

  /**
   * 「指摘一覧をコピー」で渡す文章。入力した文章の全文と無視した指摘は含めない
   *   source: 読み込んだファイルの名前と、判定に含めた範囲（ファイルを読み込んだときだけ）
   *   locate: 指摘の場所（ページ・シートとセル番地、欄の名前）を返す関数
   *   docLines: 書類の種類と必須項目の結果（S3）
   * @param {{ findings: Finding[], ignoredCount: number, charCount: number, styleCounts: StyleCounts|null, ruleset: RulesetView, now: Date,
   *   source?: { fileName: string, scope: string }, locate?: (f: Finding) => string, docLines?: string[] }} args
   */
  const buildCopyText = ({ findings, ignoredCount, charCount, styleCounts, ruleset, now, source, locate, docLines }) => {
    const { meta } = ruleset;
    const sorted = groupFindings(findings, 'position').flatMap((g) => g.findings);
    const lines = [
      `${meta.toolName} 判定結果`,
      `判定日時: ${formatDateTime(now)}`,
      `基準: ${meta.manualName}（${meta.manualVersion}）／ルール定義の更新日: ${meta.updated}`,
      ...(source ? [`ファイル: ${source.fileName}${source.scope ? `（${source.scope}）` : ''}`] : []),
      `件数: ${formatCounts(countBySeverity(findings))}（判定した文字数: ${charCount.toLocaleString('ja-JP')}字）`,
      ...(docLines ?? []),
      ...(styleCounts && hasStyleFinding(findings) ? [`文体: ${formatStyleCounts(styleCounts)}`] : []),
      ...(ignoredCount > 0 ? [`（無視した指摘 ${ignoredCount}件は含めていません）`] : []),
      '',
      ...(sorted.length > 0
        ? sorted.map((f) => findingLine(f, locate ? locate(f) : ''))
        : ['問題は見つかりませんでした。']),
    ];
    return lines.join('\n');
  };

  /**
   * 「AIへの依頼文をコピー」で渡す文章。ルール定義から作り、入力した文章は含めない
   * @param {RulesetView} ruleset
   */
  const buildAiPrompt = (ruleset) => {
    const { meta, aiGuide } = ruleset;
    const ruleLines = ruleset.rules.map((r, i) => {
      const ex =
        r.examples.length > 0
          ? ` 例: ${r.examples.join('／')}${r.fixes.length > 0 ? ` → ${r.fixes.join('／')}` : ''}`
          : '';
      return `${i + 1}. ${r.name}（${r.severity}）: ${r.description}${ex}`;
    });
    return [
      aiGuide.request,
      '',
      `【ルール】（${meta.manualName} ${meta.manualVersion}より。重さは「違反」＝ルールに明確に反する、「要確認」＝文脈によっては問題ない、「参考」＝できれば直したほうが読みやすい）`,
      ...ruleLines,
      '',
      '【確認したい文章】',
      '（ここに、個人名・電話番号などを伏せた文章を貼り付けてください）',
    ].join('\n');
  };

  /**
   * ルール一覧の画面に出す内容
   * @param {RulesetView} ruleset
   */
  const buildRuleListModel = (ruleset) => {
    const notes = ruleset.categoryNotes ?? {};
    const categories = ruleset.categories
      .map((name) => ({
        name,
        note: notes[name] ?? '',
        rules: ruleset.rules
          .filter((r) => r.category === name)
          .sort((a, b) => a.order - b.order)
          .map((r) => ({
            id: r.id,
            name: r.name,
            severity: r.severity,
            examples: r.examples,
            fixes: r.fixes,
            description: r.description,
            source: formatSource(ruleset, r.source),
            scope: r.scope,
          })),
      }))
      .filter((c) => c.rules.length > 0);
    return { total: ruleset.rules.length, categories, notChecked: ruleset.notChecked };
  };

  root.YegReport = Object.freeze({
    SEVERITIES,
    WHOLE_DOCUMENT,
    AMOUNT_GROUP,
    severityClass,
    visibleFindings,
    ignoredFindings,
    countBySeverity,
    formatCounts,
    filterBySeverity,
    groupFindings,
    formatFix,
    formatSource,
    formatStyleCounts,
    hasStyleFinding,
    oneLine,
    contextLine,
    formatDateTime,
    findingLine,
    buildCopyText,
    buildAiPrompt,
    buildRuleListModel,
  });
})(globalThis);
