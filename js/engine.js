// @ts-check
/*
 * 判定の仕組み（画面に依存しない純粋な関数だけを置く）
 *   compileRuleset(ルール定義) → 読み込めたルールと、読み込めなかったルールの一覧
 *   checkText(文章, 読み込めたルール) → 指摘の一覧
 */
'use strict';

(function (root) {
  /** @typedef {'違反'|'要確認'|'参考'} Severity */

  /**
   * @typedef {Object} CompiledPattern
   * @property {RegExp} regex
   * @property {string[]} fix
   * @property {string|null} fixFn
   * @property {string|null} test
   * @property {RegExp[]} exclude
   * @property {boolean} crossLines
   */

  /**
   * @typedef {Object} StyleConfig
   * @property {string} politeLabel
   * @property {string} plainLabel
   * @property {string[]} politeEndings
   * @property {string[]} neutralEndings
   * @property {string[]} plainEndings
   * @property {string} plainLastChars
   */

  /** @typedef {Severity|'指摘しない'} SectionSeverity */

  /**
   * 書類全体を見て判定するルール（type: 'document'）の条件
   * @typedef {Object} DocumentCondition
   * @property {RegExp[]} anyOf   どれかが文章にあること
   * @property {RegExp[]} noneOf  1つもないこと
   */

  /**
   * @typedef {Object} CompiledRule
   * @property {string} id
   * @property {string} name
   * @property {string} category
   * @property {Severity} severity
   * @property {'pattern'|'style'|'document'|'amount'} type
   * @property {CompiledPattern[]} patterns
   * @property {StyleConfig|null} style
   * @property {string[]} examples
   * @property {string[]} fixes
   * @property {string} hint
   * @property {string} description
   * @property {string} source
   * @property {string} scope
   * @property {number} order
   * @property {Readonly<Record<string, SectionSeverity>>|null} sectionSeverity  欄に応じた重さ（「ほかの欄」は名前のない欄）
   * @property {string[]} documentTypes  type: 'document' の対象の書類の種類。type: 'pattern' でも書けて、そのときはこの種類の書類だけに当てる（空なら全部）
   * @property {string} dynamicWords  type: 'pattern' で、見つける語を入力から読み取るもの（DYNAMIC_WORDS のどれか。'' なら使わない）
   * @property {DocumentCondition|null} condition
   * @property {string} message  type: 'document' の指摘の文
   * @property {string} check    type: 'amount' の計算の名前（AMOUNT_CHECKS のどれか）。type: 'document' で condition の代わりに使う構成のチェック（DOCUMENT_CHECKS のどれか）
   * @property {Readonly<Record<string, unknown>>} params  type: 'document' の構成のチェックに使う値（字数の目安など）
   * @property {Readonly<Record<string, string>>} messages  type: 'amount' の指摘の文（言い分けがあるときは名前ごと。1つなら default）
   */

  /**
   * 確認ポイント
   * @typedef {Object} Checkpoint
   * @property {string} id
   * @property {string} text
   * @property {string[]} checkedBy  Excel の収支表で、金額チェックのこのルールが自動で確かめている（M-xx）
   */

  /**
   * 書類の種類（documents.types の1件）
   * @typedef {Object} DocType
   * @property {string} id
   * @property {string} name
   * @property {string[]} detectWords
   * @property {string[]} required
   * @property {string[]} settlementRequired  決算のときだけ必須
   * @property {string[]} blankCheck
   * @property {Checkpoint[]} checkpoints
   * @property {boolean} useCommonCheckpoints  共通の確認ポイントも出すか（委員会事業計画書は議案向けの共通の確認ポイントを出さない）
   * @property {string[]} scopeSections  判定する欄（この種類の書類では、ここに書いた欄の中の指摘だけを出す。空なら全部）
   * @property {{ off: string[], lower: boolean }} textRules  文章のルールの扱い（off: 当てないルール、lower: 「違反」を「要確認」に下げる）
   * @property {RegExp|null} committeeNamePattern  見出しから自分の委員会名を読み取る正規表現（1番目のかっこが委員会名）
   */

  /**
   * 金額チェックのデータ（ルール定義の amounts）。語は labelKey でそろえた形で持つ
   * @typedef {Object} CompiledAmounts
   * @property {string} documentType  金額チェックを行う書類の種類（documents.types の id）
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
   * @property {string[]} multiplySigns  そろえない（そのままの1文字）
   * @property {string[]} numberMarks    そろえない（そのままの1文字）
   * @property {{ documentRules: string[], cellRules: string[] }} replaces
   */

  /**
   * 書類の種類に応じたチェックに使うデータ（ルール定義の documents を読み込んだもの）
   * 見出しの語は normalizeLabel でそろえた形（key）で持つ
   * @typedef {Object} CompiledDocuments
   * @property {number} minDetectWords
   * @property {DocType[]} types
   * @property {Checkpoint[]} commonCheckpoints
   * @property {{ key: string, name: string }[]} headingKeys  見出し語（別名を含む）。長い順
   * @property {Readonly<Record<string, string[]>>} substitutes  見つからないとき代わりに確認する見出し
   * @property {{ name: string, regexes: RegExp[] }[]} phrases  文章のどこかにあればよい必須項目
   * @property {{ budget: RegExp, settlement: RegExp }} budgetKind
   * @property {string[]} columnSuffixes
   * @property {string[]} otherLabels
   * @property {Readonly<Record<string, string[]>>} fieldLabels
   * @property {RegExp[]} sectionBreaks
   */

  /**
   * 欄の範囲（判定する文章の中の位置）
   * @typedef {Object} Section
   * @property {number} start
   * @property {number} end
   * @property {string} name  欄の名前（見出し）
   */

  /**
   * @typedef {Object} RuleLoadError
   * @property {string} ruleLabel  画面に出すルールの呼び名（R-05 や「12件目のルール」）
   * @property {string} message    読み込めなかった理由
   */

  /**
   * @typedef {Object} RulesetMeta
   * @property {string} toolName
   * @property {string} manualName
   * @property {string} manualVersion
   * @property {string} sourceLabel
   * @property {string} updated
   */

  /**
   * @typedef {Object} AiLink
   * @property {string} label
   * @property {string} url
   */

  /**
   * @typedef {Object} AiGuide
   * @property {string} caution
   * @property {string} request
   * @property {AiLink[]} links
   */

  /**
   * @typedef {Object} OtherPage
   * @property {string} name    書類の名前（例: 委員会事業計画書）
   * @property {string} href    ページ（例: annual.html）
   * @property {string[]} keys  判別語（labelKey でそろえた形）
   */

  /**
   * @typedef {Object} CompiledRuleset
   * @property {RulesetMeta} meta
   * @property {string[]} categories
   * @property {string[]} excludedPhrases
   * @property {string[]} notChecked
   * @property {AiGuide} aiGuide
   * @property {OtherPage[]} otherPages  別のページで判定する書類（議案のページで案内する）
   * @property {CompiledRule[]} rules
   * @property {CompiledDocuments|null} documents  書類の種類のデータ（読めなかったときは null）
   * @property {CompiledAmounts|null} amounts      金額チェックのデータ（読めなかったときは null）
   * @property {Readonly<Record<string, string>>} categoryNotes  ルール一覧で分類の見出しの下に出す説明
   * @property {RuleLoadError[]} errors
   * @property {string|null} fatal  ルール定義全体が読めなかったときの説明
   */

  /**
   * @typedef {Object} Range
   * @property {number} start
   * @property {number} end
   */

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
   * @typedef {Object} Finding
   * @property {string} id
   * @property {'word'|'sentence'|'document'} kind
   * @property {Range|null} range  位置を持たない（書類全体の）指摘は null
   * @property {string} text
   * @property {string} before
   * @property {string} after
   * @property {Reason[]} reasons  1件目がこの指摘の代表（いちばん重いもの）
   * @property {Severity} severity
   * @property {string} key  無視の判定に使う（ルール・語・前後の文脈）
   * @property {string|null} section  指摘のある欄の名前（欄が分からないときは null）
   */

  /**
   * @typedef {Object} StyleCounts
   * @property {number} polite
   * @property {number} plain
   * @property {string} politeLabel
   * @property {string} plainLabel
   */

  /**
   * @typedef {{ ok: true, text: string, charCount: number, findings: Finding[], styleCounts: StyleCounts|null }
   *   | { ok: false, reason: 'empty'|'too-long', charCount: number }} CheckResult
   */

  /**
   * @typedef {Object} Hit
   * @property {number} start
   * @property {number} end
   * @property {string[]} fixes
   * @property {string} hint  正しい書き方が1つに決まらないときの短い直し方
   * @property {string} note  この指摘だけの補足（文体の内訳など）
   * @property {CompiledRule} rule
   * @property {Severity} [severity]  欄に応じて重さを変えたとき（なければルールの重さ）
   */

  /** @type {readonly Severity[]} */
  const SEVERITIES = Object.freeze(['違反', '要確認', '参考']);
  const MAX_CHARS = 100000;
  const CONTEXT_CHARS = 20;
  const REGEX_FLAGS = 'gmu';

  // ------------------------------------------------------------------
  // 文字の変換（純粋関数）
  // ------------------------------------------------------------------

  /** @param {string} s */
  const toHalfWidth = (s) =>
    s.replace(/[０-９Ａ-Ｚａ-ｚ]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xfee0));

  /** 見出しをそろえるときに取り除く文字（空白・改行・【】などのかっこ） */
  const LABEL_DROP = /[\s【】〔〕[\]〖〗]/u;

  /**
   * 見出しの表記ゆれをそろえる（全角・半角の違い、空白・改行、【】の有無、「・」と「･」の違いを無視する）。
   * map[i] は、そろえた文字の i 文字目が元の文字列の何文字目から来たか
   * @param {string} s
   * @returns {{ text: string, map: number[] }}
   */
  const normalizeLabel = (s) => {
    /** @type {string[]} */
    const out = [];
    /** @type {number[]} */
    const map = [];
    let at = 0;
    for (const ch of s) {
      const code = ch.codePointAt(0) ?? 0;
      // ひらがな・カタカナ・漢字・英数字はそのまま（NFKC で形が変わる文字だけ変換する）
      const plain = code < 0x80 || (code >= 0x3040 && code <= 0x30ff) || (code >= 0x4e00 && code <= 0x9fff);
      for (const c of plain ? ch : ch.normalize('NFKC')) {
        if (LABEL_DROP.test(c)) continue;
        out.push(c === '·' || c === '•' ? '・' : c);
        map.push(at);
      }
      at += ch.length;
    }
    return { text: out.join(''), map };
  };

  /** @param {string} s */
  const labelKey = (s) => normalizeLabel(s).text;

  /**
   * 区切りのない数字を3桁ごとに区切る（長い数字でも遅くならないよう、正規表現を使わず先頭から切る）
   * @param {string} digits
   */
  const groupThousands = (digits) => {
    const head = digits.length % 3 || 3;
    const parts = [digits.slice(0, head)];
    for (let i = head; i < digits.length; i += 3) parts.push(digits.slice(i, i + 3));
    return parts.join(',');
  };

  /**
   * 文字列の中の数字のかたまりを、3桁区切りに直す（全角は半角にする）
   * @param {string} s
   */
  const addThousands = (s) =>
    toHalfWidth(s)
      .replace(/，/g, ',')
      .replace(/\d[\d,]*\d|\d/g, (run) => groupThousands(run.replace(/,/g, '')));

  /**
   * 区切りの位置が正しくないか（1,0000・12,00・1,23,456）。
   * 「1,2,3」のように1桁ずつ並べたものは列挙とみなして対象外にする
   * @param {string} s
   */
  const isBadGrouping = (s) => {
    const groups = toHalfWidth(s).replace(/，/g, ',').split(',');
    const rest = groups.slice(1);
    if (rest.every((g) => g.length === 1)) return false;
    return groups[0].length > 3 || rest.some((g) => g.length !== 3);
  };

  const ERAS = Object.freeze([
    { name: '令和', y: 2019, m: 5, d: 1 },
    { name: '平成', y: 1989, m: 1, d: 8 },
    { name: '昭和', y: 1926, m: 12, d: 25 },
    { name: '大正', y: 1912, m: 7, d: 30 },
  ]);

  /**
   * @typedef {Object} SimpleDate
   * @property {number} y
   * @property {number} m
   * @property {number} d
   */

  /**
   * 西暦の年月日の文字列を読む。日付として正しくなければ null
   * @param {string} s
   * @returns {SimpleDate|null}
   */
  const parseWesternDate = (s) => {
    const m = toHalfWidth(s).match(/(\d{4})\D(\d{1,2})\D(\d{1,2})/);
    if (!m) return null;
    const [y, mo, d] = [Number(m[1]), Number(m[2]), Number(m[3])];
    if (y < 1900 || y > 2099 || mo < 1 || mo > 12 || d < 1) return null;
    const lastDay = new Date(Date.UTC(y, mo, 0)).getUTCDate();
    return d <= lastDay ? { y, m: mo, d } : null;
  };

  /**
   * @param {SimpleDate} date
   * @returns {string|null}
   */
  const formatWareki = ({ y, m, d }) => {
    const key = y * 10000 + m * 100 + d;
    const era = ERAS.find((e) => key >= e.y * 10000 + e.m * 100 + e.d);
    if (!era) return null;
    const eraYear = y - era.y + 1;
    return `${era.name}${eraYear === 1 ? '元' : eraYear}年${m}月${d}日`;
  };

  /**
   * 年だけの4桁（19xx・20xx）で、後ろに単位が付かないもの
   * @param {string} digits
   * @param {string} next 直後の1文字
   */
  const isBareYear = (digits, next) =>
    /^(19|20)\d{2}$/.test(toHalfWidth(digits)) && !/^[\p{sc=Han}\p{sc=Katakana}%％]/u.test(next);

  /** @type {Readonly<Record<string, (match: string) => string[]>>} */
  const FIX_FUNCTIONS = Object.freeze({
    halfWidth: (m) => [toHalfWidth(m)],
    thousands: (m) => [addThousands(m)],
    wareki: (m) => {
      const date = parseWesternDate(m);
      const w = date ? formatWareki(date) : null;
      return w ? [w] : [];
    },
  });

  /** @type {Readonly<Record<string, (ctx: { match: string, next: string }) => boolean>>} */
  const TEST_FUNCTIONS = Object.freeze({
    badGrouping: ({ match }) => isBadGrouping(match),
    notBareYear: ({ match, next }) => !isBareYear(match, next),
    validDate: ({ match }) => parseWesternDate(match) !== null,
  });

  // ------------------------------------------------------------------
  // ルール定義の読み込み
  // ------------------------------------------------------------------

  /** @param {unknown} v @returns {v is Record<string, unknown>} */
  const isObject = (v) => typeof v === 'object' && v !== null && !Array.isArray(v);
  /** @param {unknown} v @returns {v is string} */
  const isNonEmptyString = (v) => typeof v === 'string' && v.trim() !== '';
  /** @param {unknown} v @returns {v is string[]} */
  const isStringArray = (v) => Array.isArray(v) && v.every((x) => typeof x === 'string');
  /** @param {string} s */
  const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\\/]/g, '\\$&');

  /**
   * @template T
   * @typedef {{ ok: true, value: T } | { ok: false, message: string }} Outcome
   */

  /**
   * @param {string} source
   * @param {string} what
   * @returns {Outcome<RegExp>}
   */
  const compileRegex = (source, what) => {
    try {
      return { ok: true, value: new RegExp(source, REGEX_FLAGS) };
    } catch (e) {
      const detail = e instanceof Error ? e.message : String(e);
      return { ok: false, message: `${what}の書き方に誤りがあります（${detail}）` };
    }
  };

  /**
   * @param {unknown} raw
   * @param {number} index
   * @returns {Outcome<CompiledPattern>}
   */
  const compilePattern = (raw, index) => {
    const label = `patterns の${index + 1}行目`;
    if (!isObject(raw)) return { ok: false, message: `${label}が { } の形になっていません` };

    /** @type {string|null} */
    let source = null;
    if (typeof raw.re === 'string' && raw.re !== '') {
      source = raw.re;
    } else if (Array.isArray(raw.words) && raw.words.length > 0 && raw.words.every(isNonEmptyString)) {
      source = raw.words.map((w) => escapeRegExp(String(w))).join('|');
    }
    if (source === null) {
      return { ok: false, message: `${label}に、見つける語（words か re）が書かれていません` };
    }
    const regex = compileRegex(source, label);
    if (!regex.ok) return regex;

    if (raw.fix !== undefined && !isStringArray(raw.fix)) {
      return { ok: false, message: `${label}の fix（正しい書き方）は ['…'] の形で書いてください` };
    }
    if (raw.fixFn !== undefined && !(typeof raw.fixFn === 'string' && raw.fixFn in FIX_FUNCTIONS)) {
      return { ok: false, message: `${label}の fixFn「${String(raw.fixFn)}」は使えません` };
    }
    if (raw.test !== undefined && !(typeof raw.test === 'string' && raw.test in TEST_FUNCTIONS)) {
      return { ok: false, message: `${label}の test「${String(raw.test)}」は使えません` };
    }
    if (raw.crossLines !== undefined && typeof raw.crossLines !== 'boolean') {
      return { ok: false, message: `${label}の crossLines は true か false で書いてください` };
    }
    if (raw.exclude !== undefined && !isStringArray(raw.exclude)) {
      return { ok: false, message: `${label}の exclude は ['…'] の形で書いてください` };
    }
    /** @type {RegExp[]} */
    const exclude = [];
    const excludeList = isStringArray(raw.exclude) ? raw.exclude : [];
    for (const [i, ex] of excludeList.entries()) {
      const r = compileRegex(String(ex), `${label}の exclude の${i + 1}つ目`);
      if (!r.ok) return r;
      exclude.push(r.value);
    }
    return {
      ok: true,
      value: {
        regex: regex.value,
        fix: isStringArray(raw.fix) ? raw.fix : [],
        fixFn: typeof raw.fixFn === 'string' ? raw.fixFn : null,
        test: typeof raw.test === 'string' ? raw.test : null,
        exclude,
        crossLines: raw.crossLines !== false,
      },
    };
  };

  /**
   * @param {unknown} raw
   * @returns {Outcome<StyleConfig>}
   */
  const compileStyle = (raw) => {
    if (!isObject(raw)) return { ok: false, message: 'style（文体の判定に使う語）が書かれていません' };
    const lists = ['politeEndings', 'neutralEndings', 'plainEndings'];
    const bad = lists.find((k) => !isStringArray(raw[k]));
    if (bad) return { ok: false, message: `style の ${bad} は ['…'] の形で書いてください` };
    if (typeof raw.plainLastChars !== 'string') {
      return { ok: false, message: 'style の plainLastChars は文字列で書いてください' };
    }
    return {
      ok: true,
      value: {
        politeLabel: isNonEmptyString(raw.politeLabel) ? raw.politeLabel : 'ですます調',
        plainLabel: isNonEmptyString(raw.plainLabel) ? raw.plainLabel : 'である調',
        politeEndings: /** @type {string[]} */ (raw.politeEndings),
        neutralEndings: /** @type {string[]} */ (raw.neutralEndings),
        plainEndings: /** @type {string[]} */ (raw.plainEndings),
        plainLastChars: raw.plainLastChars,
      },
    };
  };

  const OTHER_SECTIONS = 'ほかの欄';
  const NOT_REPORTED = '指摘しない';
  const DOCUMENTS_LABEL = '書類の種類のデータ（documents）';

  /**
   * 欄に応じた重さ（sectionSeverity）を読む。書いていなければ null
   * @param {unknown} raw
   * @returns {Outcome<Readonly<Record<string, SectionSeverity>>|null>}
   */
  const compileSectionSeverity = (raw) => {
    if (raw === undefined) return { ok: true, value: null };
    if (!isObject(raw)) return { ok: false, message: 'sectionSeverity は { 欄の名前: 重さ } の形で書いてください' };
    /** @type {Record<string, SectionSeverity>} */
    const out = {};
    for (const [name, sev] of Object.entries(raw)) {
      if (!(typeof sev === 'string' && (sev === NOT_REPORTED || SEVERITIES.some((s) => s === sev)))) {
        return { ok: false, message: `sectionSeverity の「${name}」の重さは、違反・要確認・参考・指摘しない のどれかで書いてください` };
      }
      out[labelKey(name) === labelKey(OTHER_SECTIONS) ? OTHER_SECTIONS : name] = /** @type {SectionSeverity} */ (sev);
    }
    return { ok: true, value: Object.freeze(out) };
  };

  /**
   * @param {unknown} list
   * @param {string} what
   * @returns {Outcome<RegExp[]>}
   */
  const compileRegexList = (list, what) => {
    if (!isStringArray(list)) return { ok: false, message: `${what} は ['…'] の形で書いてください` };
    /** @type {RegExp[]} */
    const out = [];
    for (const [i, src] of list.entries()) {
      const r = compileRegex(src, `${what}の${i + 1}つ目`);
      if (!r.ok) return r;
      out.push(r.value);
    }
    return { ok: true, value: out };
  };

  /**
   * 書類全体を見るルールの条件
   * @param {unknown} raw
   * @returns {Outcome<DocumentCondition>}
   */
  const compileCondition = (raw) => {
    if (!isObject(raw)) return { ok: false, message: '条件（condition）が書かれていません' };
    const anyOf = compileRegexList(raw.anyOf, 'condition の anyOf');
    if (!anyOf.ok) return anyOf;
    if (anyOf.value.length === 0) return { ok: false, message: 'condition の anyOf に1つ以上書いてください' };
    const noneOf = raw.noneOf === undefined ? { ok: true, value: [] } : compileRegexList(raw.noneOf, 'condition の noneOf');
    if (!noneOf.ok) return noneOf;
    return { ok: true, value: { anyOf: anyOf.value, noneOf: noneOf.value } };
  };

  /**
   * @param {unknown} raw
   * @param {string} what
   * @returns {Outcome<Checkpoint[]>}
   */
  const compileCheckpoints = (raw, what) => {
    if (raw === undefined) return { ok: true, value: [] };
    if (
      !Array.isArray(raw) ||
      !raw.every(
        (k) => isObject(k) && isNonEmptyString(k.id) && isNonEmptyString(k.text) && (k.checkedBy === undefined || isStringArray(k.checkedBy)),
      )
    ) {
      return { ok: false, message: `${what}は [{ id: 'K-01', text: '…' }] の形で書いてください（checkedBy は ['M-01'] の形）` };
    }
    return {
      ok: true,
      value: raw.map((k) => ({ id: String(k.id), text: String(k.text), checkedBy: isStringArray(k.checkedBy) ? k.checkedBy : [] })),
    };
  };

  const AMOUNTS_LABEL = '金額チェックのデータ（amounts）';

  /** 金額チェックの計算の名前（ルール定義の check に書けるもの） */
  const AMOUNT_CHECKS = Object.freeze([
    'incomeTotal',
    'expenseTotal',
    'balance',
    'surplus',
    'feeSurplus',
    'reserveRate',
    'sales',
    'feeFormula',
    'feeCalc',
    'basisCalc',
    'withholdingNote',
    'withholdingCalc',
    'lecturerBreakdown',
    'expenseItem',
    'incomeItem',
    'thousands',
  ]);

  /** 書類全体の構成のチェックの名前（type: 'document' で condition の代わりに check に書けるもの。判定は annual.js） */
  const DOCUMENT_CHECKS = Object.freeze(['length', 'total', 'opening', 'closing']);

  /** 入力から読み取る語の名前（type: 'pattern' の dynamicWords に書けるもの） */
  const DYNAMIC_WORDS = Object.freeze(['committeeName']);

  /**
   * 構成のチェックの params の書き間違い（なければ null）
   * @param {string} check
   * @param {Record<string, unknown>} params
   * @returns {string|null}
   */
  const checkParamsError = (check, params) => {
    if (params.section !== undefined && !isNonEmptyString(params.section)) return "params.section（本文の欄の名前）は '…' の形で書いてください";
    if (check === 'length') {
      const parts = params.parts;
      const okParts =
        Array.isArray(parts) &&
        parts.length === 3 &&
        parts.every((p) => isObject(p) && isNonEmptyString(p.label) && typeof p.target === 'number' && p.target > 0);
      if (!okParts) return "params.parts は上段・中段・下段の3つを [{ label: '上段', target: 200 }, …] の形で書いてください";
      for (const key of ['total', 'low', 'high']) {
        const v = params[key];
        if (v !== undefined && !(typeof v === 'number' && v > 0)) return `params.${key} は 0 より大きい数で書いてください`;
      }
    }
    if (check === 'total' && !(typeof params.max === 'number' && params.max > 0)) {
      return 'params.max（合計の字数の上限）は 0 より大きい数で書いてください';
    }
    if (check === 'opening' && !isNonEmptyString(params.opening)) return "params.opening（中段の書き出し）は '…' の形で書いてください";
    if (check === 'closing' && !(isStringArray(params.closings) && params.closings.some(isNonEmptyString))) {
      return "params.closings（締めの言葉）は ['…'] の形で書いてください";
    }
    return null;
  };

  /**
   * 指摘の文（文字列1つ、または { 名前: 文 }）
   * @param {unknown} raw
   * @returns {Outcome<Readonly<Record<string, string>>>}
   */
  const compileMessages = (raw) => {
    if (isNonEmptyString(raw)) return { ok: true, value: Object.freeze({ default: raw }) };
    if (isObject(raw) && Object.keys(raw).length > 0 && Object.values(raw).every(isNonEmptyString)) {
      return { ok: true, value: Object.freeze(/** @type {Record<string, string>} */ ({ ...raw })) };
    }
    return { ok: false, message: "指摘の文（message）を '…' か { 名前: '…' } の形で書いてください" };
  };

  /**
   * ルール定義の amounts（金額チェックに使う語・科目・割合など）を読み込む
   * @param {unknown} raw
   * @returns {Outcome<CompiledAmounts>}
   */
  const compileAmounts = (raw) => {
    if (!isObject(raw)) return { ok: false, message: 'amounts が { } の形で書かれていません' };
    /** @param {string} key @param {boolean} required @returns {Outcome<string[]>} */
    const keys = (key, required) => {
      const v = raw[key];
      if (v === undefined && !required) return { ok: true, value: [] };
      if (!isStringArray(v) || (required && v.length === 0)) return { ok: false, message: `amounts.${key} は ['…'] の形で書いてください` };
      return { ok: true, value: [...new Set(v.map(labelKey).filter((k) => k !== ''))] };
    };
    /** @param {string} key @returns {Outcome<string>} */
    const word = (key) =>
      isNonEmptyString(raw[key]) ? { ok: true, value: labelKey(raw[key]) } : { ok: false, message: `amounts.${key} を '…' の形で書いてください` };
    /** @param {string} key @returns {Outcome<number>} */
    const positive = (key) => {
      const v = raw[key];
      return typeof v === 'number' && Number.isFinite(v) && v > 0
        ? { ok: true, value: v }
        : { ok: false, message: `amounts.${key} は 0 より大きい数で書いてください` };
    };
    /** @param {string} key @returns {Outcome<string[]>} */
    const chars = (key) => {
      const v = raw[key];
      return isStringArray(v) && v.length > 0 && v.every((c) => [...c].length === 1)
        ? { ok: true, value: v }
        : { ok: false, message: `amounts.${key} は ['×'] のように1文字ずつ書いてください` };
    };

    /** @type {Record<string, unknown>} */
    const out = {};
    const required = ['incomeHeadings', 'expenseHeadings', 'budgetColumns', 'settlementColumns', 'incomeTotals', 'expenseTotals', 'expenseItems', 'incomeItems'];
    const optional = [
      'headingSuffixes',
      'noteColumns',
      'budgetNoteWords',
      'settlementNoteWords',
      'balanceTotals',
      'summaryBudgetRows',
      'summarySettlementRows',
      'salesWords',
      'purchaseWords',
      'honorariumWords',
      'withholdingWords',
      'lecturerBreakdownWords',
    ];
    for (const key of [...required, ...optional]) {
      const r = keys(key, required.includes(key));
      if (!r.ok) return r;
      out[key] = r.value;
    }
    for (const key of ['activityFee', 'specialFee', 'lecturerItem', 'reserveItem']) {
      const r = word(key);
      if (!r.ok) return r;
      out[key] = r.value;
    }
    for (const key of ['withholdingRatePercent', 'withholdingLimit', 'reserveMaxPercent']) {
      const r = positive(key);
      if (!r.ok) return r;
      out[key] = r.value;
    }
    for (const key of ['multiplySigns', 'numberMarks']) {
      const r = chars(key);
      if (!r.ok) return r;
      out[key] = r.value;
    }
    const rep = raw.replaces === undefined ? {} : raw.replaces;
    if (!isObject(rep) || !['documentRules', 'cellRules'].every((k) => rep[k] === undefined || isStringArray(rep[k]))) {
      return { ok: false, message: "amounts.replaces は { documentRules: ['D-06'], cellRules: ['R-16'] } の形で書いてください" };
    }
    return {
      ok: true,
      value: /** @type {CompiledAmounts} */ ({
        ...out,
        documentType: isNonEmptyString(raw.documentType) ? raw.documentType : 'budget',
        replaces: {
          documentRules: isStringArray(rep.documentRules) ? rep.documentRules : [],
          cellRules: isStringArray(rep.cellRules) ? rep.cellRules : [],
        },
      }),
    };
  };

  /**
   * ルール定義の documents（書類の種類・必須項目・確認ポイントなど）を読み込む
   * @param {unknown} raw
   * @returns {Outcome<CompiledDocuments>}
   */
  const compileDocuments = (raw) => {
    if (!isObject(raw)) return { ok: false, message: 'documents が { } の形で書かれていません' };
    if (!Array.isArray(raw.types) || raw.types.length === 0) {
      return { ok: false, message: 'documents.types（書類の種類）が書かれていません' };
    }
    const phrasesRaw = raw.phrases === undefined ? {} : raw.phrases;
    if (!isObject(phrasesRaw)) return { ok: false, message: 'documents.phrases は { 名前: [\'…\'] } の形で書いてください' };
    /** @type {{ name: string, regexes: RegExp[] }[]} */
    const phrases = [];
    for (const [name, list] of Object.entries(phrasesRaw)) {
      const r = compileRegexList(list, `documents.phrases の「${name}」`);
      if (!r.ok) return r;
      phrases.push({ name, regexes: r.value });
    }
    const phraseNames = new Set(phrases.map((p) => p.name));

    /** @type {DocType[]} */
    const types = [];
    for (const [i, t] of raw.types.entries()) {
      const label = `documents.types の${i + 1}件目`;
      if (!isObject(t) || !isNonEmptyString(t.id) || !isNonEmptyString(t.name)) {
        return { ok: false, message: `${label}に id と name を書いてください` };
      }
      for (const key of ['detectWords', 'required', 'blankCheck']) {
        if (!isStringArray(t[key])) return { ok: false, message: `${label}の ${key} は ['…'] の形で書いてください` };
      }
      if (t.settlementRequired !== undefined && !isStringArray(t.settlementRequired)) {
        return { ok: false, message: `${label}の settlementRequired は ['…'] の形で書いてください` };
      }
      const checkpoints = compileCheckpoints(t.checkpoints, `${label}の checkpoints `);
      if (!checkpoints.ok) return checkpoints;
      if (t.commonCheckpoints !== undefined && typeof t.commonCheckpoints !== 'boolean') {
        return { ok: false, message: `${label}の commonCheckpoints は true か false で書いてください` };
      }
      if (t.scopeSections !== undefined && !isStringArray(t.scopeSections)) {
        return { ok: false, message: `${label}の scopeSections は ['…'] の形で書いてください` };
      }
      const tr = t.textRules === undefined ? {} : t.textRules;
      if (!isObject(tr) || (tr.off !== undefined && !isStringArray(tr.off)) || (tr.lower !== undefined && typeof tr.lower !== 'boolean')) {
        return { ok: false, message: `${label}の textRules は { off: ['R-01'], lower: true } の形で書いてください` };
      }
      /** @type {RegExp|null} */
      let committeeNamePattern = null;
      if (t.committeeNamePattern !== undefined) {
        const re = isNonEmptyString(t.committeeNamePattern)
          ? compileRegex(t.committeeNamePattern, `${label}の committeeNamePattern `)
          : { ok: /** @type {false} */ (false), message: `${label}の committeeNamePattern は '…' の形で書いてください` };
        if (!re.ok) return re;
        // 委員会名は1番目のかっこで受け取るので、かっこがないと読み取れない
        const groups = new RegExp(`${re.value.source}|`, 'u').exec('');
        if (!groups || groups.length < 2) {
          return { ok: false, message: `${label}の committeeNamePattern に、委員会名を受け取るかっこ ( ) を書いてください` };
        }
        committeeNamePattern = re.value;
      }
      types.push({
        id: String(t.id),
        name: String(t.name),
        detectWords: /** @type {string[]} */ (t.detectWords),
        required: /** @type {string[]} */ (t.required),
        settlementRequired: isStringArray(t.settlementRequired) ? t.settlementRequired : [],
        blankCheck: /** @type {string[]} */ (t.blankCheck),
        checkpoints: checkpoints.value,
        useCommonCheckpoints: t.commonCheckpoints !== false,
        scopeSections: isStringArray(t.scopeSections) ? t.scopeSections : [],
        textRules: { off: isStringArray(tr.off) ? tr.off : [], lower: tr.lower === true },
        committeeNamePattern,
      });
    }
    const common = compileCheckpoints(raw.commonCheckpoints, 'documents.commonCheckpoints ');
    if (!common.ok) return common;

    const aliasesRaw = raw.aliases === undefined ? {} : raw.aliases;
    if (!isObject(aliasesRaw) || !Object.values(aliasesRaw).every(isStringArray)) {
      return { ok: false, message: "documents.aliases は { 見出し: ['別名'] } の形で書いてください" };
    }
    const aliases = /** @type {Record<string, string[]>} */ (aliasesRaw);

    // 見出し語（必須の見出し・判別語・別名）。そろえた形が同じものは1つにする
    const names = [
      ...new Set(
        types.flatMap((t) => [...t.required, ...t.settlementRequired, ...t.detectWords]).filter((n) => !phraseNames.has(n)),
      ),
    ];
    /** @type {Map<string, string>} */
    const byKey = new Map();
    for (const name of [...names, ...Object.keys(aliases)]) {
      for (const variant of [name, ...(aliases[name] ?? [])]) {
        const key = labelKey(variant);
        if (key !== '' && !byKey.has(key)) byKey.set(key, name);
      }
    }
    const headingKeys = [...byKey.entries()]
      .map(([key, name]) => ({ key, name }))
      .sort((a, b) => b.key.length - a.key.length || (a.key < b.key ? -1 : 1));

    const kind = isObject(raw.budgetKind) ? raw.budgetKind : {};
    const budgetRe = compileRegex(isNonEmptyString(kind.budget) ? kind.budget : String.raw`予算書?\(案\)`, 'documents.budgetKind の budget');
    if (!budgetRe.ok) return budgetRe;
    const settlementRe = compileRegex(
      isNonEmptyString(kind.settlement) ? kind.settlement : String.raw`決算書?\(案\)`,
      'documents.budgetKind の settlement',
    );
    if (!settlementRe.ok) return settlementRe;

    for (const key of ['columnSuffixes', 'otherLabels']) {
      if (raw[key] !== undefined && !isStringArray(raw[key])) {
        return { ok: false, message: `documents.${key} は ['…'] の形で書いてください` };
      }
    }
    const fieldRaw = raw.fieldLabels === undefined ? {} : raw.fieldLabels;
    if (!isObject(fieldRaw) || !Object.values(fieldRaw).every(isStringArray)) {
      return { ok: false, message: "documents.fieldLabels は { 見出し: ['小見出し'] } の形で書いてください" };
    }
    /** @type {Record<string, string[]>} */
    const fieldLabels = {};
    for (const [name, list] of Object.entries(/** @type {Record<string, string[]>} */ (fieldRaw))) {
      fieldLabels[byKey.get(labelKey(name)) ?? name] = list.map(labelKey).filter((k) => k !== '');
    }
    const breaks = raw.sectionBreaks === undefined ? { ok: true, value: [] } : compileRegexList(raw.sectionBreaks, 'documents.sectionBreaks');
    if (!breaks.ok) return breaks;
    const subRaw = raw.substitutes === undefined ? {} : raw.substitutes;
    if (!isObject(subRaw) || !Object.values(subRaw).every(isStringArray)) {
      return { ok: false, message: "documents.substitutes は { 見出し: ['代わりの見出し'] } の形で書いてください" };
    }
    const canonical = (/** @type {string} */ n) => byKey.get(labelKey(n)) ?? n;
    /** @type {Record<string, string[]>} */
    const substitutes = {};
    for (const [name, list] of Object.entries(/** @type {Record<string, string[]>} */ (subRaw))) {
      substitutes[canonical(name)] = list.map(canonical);
    }
    const min = raw.minDetectWords;

    return {
      ok: true,
      value: {
        minDetectWords: typeof min === 'number' && Number.isInteger(min) && min > 0 ? min : 2,
        types,
        commonCheckpoints: common.value,
        headingKeys,
        substitutes: Object.freeze(substitutes),
        phrases,
        budgetKind: { budget: budgetRe.value, settlement: settlementRe.value },
        columnSuffixes: (isStringArray(raw.columnSuffixes) ? raw.columnSuffixes : []).map(labelKey).filter((k) => k !== ''),
        otherLabels: (isStringArray(raw.otherLabels) ? raw.otherLabels : []).map(labelKey).filter((k) => k !== ''),
        fieldLabels: Object.freeze(fieldLabels),
        sectionBreaks: breaks.value,
      },
    };
  };

  const REQUIRED_TEXT_FIELDS = Object.freeze([
    ['id', '番号（id）'],
    ['name', '名前（name）'],
    ['category', '分類（category）'],
    ['description', '説明（description）'],
    ['source', '出典（source）'],
  ]);

  /**
   * @param {Record<string, unknown>} raw
   * @param {number} order
   * @returns {Outcome<CompiledRule>}
   */
  const compileRule = (raw, order) => {
    const missing = REQUIRED_TEXT_FIELDS.find(([key]) => !isNonEmptyString(raw[key]));
    if (missing) return { ok: false, message: `${missing[1]}が書かれていません` };
    const severity = raw.severity;
    if (!(typeof severity === 'string' && SEVERITIES.some((s) => s === severity))) {
      return { ok: false, message: '重さ（severity）が書かれていないか、違反・要確認・参考のどれでもありません' };
    }
    for (const key of ['examples', 'fixes']) {
      if (raw[key] !== undefined && !isStringArray(raw[key])) {
        return { ok: false, message: `${key} は ['…'] の形で書いてください` };
      }
    }
    const type = raw.type === undefined ? 'pattern' : raw.type;
    if (type !== 'pattern' && type !== 'style' && type !== 'document' && type !== 'amount') {
      return { ok: false, message: `type「${String(raw.type)}」は使えません` };
    }
    const sectionSeverity = compileSectionSeverity(raw.sectionSeverity);
    if (!sectionSeverity.ok) return sectionSeverity;

    /** @type {CompiledPattern[]} */
    const patterns = [];
    /** @type {StyleConfig|null} */
    let style = null;
    /** @type {DocumentCondition|null} */
    let condition = null;
    /** @type {string[]} */
    let documentTypes = [];
    let check = '';
    let dynamicWords = '';
    /** @type {Readonly<Record<string, string>>} */
    let messages = Object.freeze({});
    /** @type {Readonly<Record<string, unknown>>} */
    let params = Object.freeze({});
    if (type === 'style') {
      const s = compileStyle(raw.style);
      if (!s.ok) return s;
      style = s.value;
    } else if (type === 'amount') {
      if (!(typeof raw.check === 'string' && AMOUNT_CHECKS.includes(raw.check))) {
        return { ok: false, message: `計算の名前（check）「${String(raw.check)}」は使えません` };
      }
      const m = compileMessages(raw.message);
      if (!m.ok) return m;
      check = raw.check;
      messages = m.value;
    } else if (type === 'document') {
      if (!isStringArray(raw.documentTypes) || raw.documentTypes.length === 0) {
        return { ok: false, message: "対象の書類の種類（documentTypes）を ['plan'] の形で書いてください" };
      }
      if (raw.check !== undefined) {
        // 構成のチェック（字数・書き出し・締め。委員会事業計画書）。条件の代わりに check と params を書く
        if (!(typeof raw.check === 'string' && DOCUMENT_CHECKS.includes(raw.check))) {
          return { ok: false, message: `構成のチェックの名前（check）「${String(raw.check)}」は使えません` };
        }
        if (!isObject(raw.params)) return { ok: false, message: 'チェックに使う値（params）を { } の形で書いてください' };
        const bad = checkParamsError(raw.check, raw.params);
        if (bad) return { ok: false, message: bad };
        const m = compileMessages(raw.message);
        if (!m.ok) return m;
        check = raw.check;
        messages = m.value;
        params = Object.freeze({ ...raw.params });
      } else {
        if (!isNonEmptyString(raw.message)) return { ok: false, message: '指摘の文（message）が書かれていません' };
        const c = compileCondition(raw.condition);
        if (!c.ok) return c;
        condition = c.value;
      }
      documentTypes = raw.documentTypes;
    } else {
      if (raw.documentTypes !== undefined) {
        if (!isStringArray(raw.documentTypes)) return { ok: false, message: "対象の書類の種類（documentTypes）は ['annual'] の形で書いてください" };
        documentTypes = raw.documentTypes;
      }
      if (raw.dynamicWords !== undefined) {
        if (!(typeof raw.dynamicWords === 'string' && DYNAMIC_WORDS.includes(raw.dynamicWords))) {
          return { ok: false, message: `入力から読み取る語（dynamicWords）「${String(raw.dynamicWords)}」は使えません` };
        }
        dynamicWords = raw.dynamicWords;
      }
      const list = Array.isArray(raw.patterns) ? raw.patterns : [];
      // 入力から語を読み取るルール（自分の委員会名など）は、patterns が空でもよい
      if (list.length === 0 && dynamicWords === '') {
        return { ok: false, message: '見つける語（patterns）が書かれていません' };
      }
      for (const [i, p] of list.entries()) {
        const c = compilePattern(p, i);
        if (!c.ok) return c;
        patterns.push(c.value);
      }
    }

    const wordsOf = () =>
      Array.isArray(raw.patterns)
        ? raw.patterns.flatMap((p) => (isObject(p) && isStringArray(p.words) ? p.words : []))
        : [];
    const fixesOf = () =>
      Array.isArray(raw.patterns)
        ? raw.patterns.flatMap((p) => (isObject(p) && isStringArray(p.fix) ? p.fix : []))
        : [];

    return {
      ok: true,
      value: {
        id: String(raw.id),
        name: String(raw.name),
        category: String(raw.category),
        severity: /** @type {Severity} */ (severity),
        type,
        patterns,
        style,
        examples: isStringArray(raw.examples) ? raw.examples : wordsOf(),
        fixes: isStringArray(raw.fixes) ? raw.fixes : [...new Set(fixesOf())],
        hint: typeof raw.hint === 'string' ? raw.hint : '',
        description: String(raw.description),
        source: String(raw.source),
        scope: isNonEmptyString(raw.scope) ? raw.scope : '書類全体',
        order,
        sectionSeverity: sectionSeverity.value,
        documentTypes,
        dynamicWords,
        condition,
        message: isNonEmptyString(raw.message) ? raw.message : '',
        check,
        messages,
        params,
      },
    };
  };

  /** @param {unknown} v @param {string} fallback */
  const textOr = (v, fallback) => (isNonEmptyString(v) ? v : fallback);

  /**
   * @param {unknown} raw
   * @returns {AiGuide}
   */
  const readAiGuide = (raw) => {
    const g = isObject(raw) ? raw : {};
    const links = Array.isArray(g.links)
      ? g.links.flatMap((l) =>
          isObject(l) && isNonEmptyString(l.label) && typeof l.url === 'string' && /^https:\/\//.test(l.url)
            ? [{ label: l.label, url: l.url }]
            : [],
        )
      : [];
    return {
      caution: textOr(
        g.caution,
        'AIサービスに文章を貼り付ける前に、個人名・電話番号などを伏せてください。',
      ),
      request: textOr(g.request, '次のルールに照らして、文章の表記を確認してください。'),
      links,
    };
  };

  /**
   * 別のページで判定する書類（議案のページで、委員会事業計画書の文章が入ったときに案内する）。
   * 書き方の違う項目は読み飛ばす。リンク先は同じサイトの中のページ（例: annual.html）だけ
   * @param {unknown} raw
   * @returns {OtherPage[]}
   */
  const readOtherPages = (raw) =>
    Array.isArray(raw)
      ? raw.flatMap((p) =>
          isObject(p) &&
          isNonEmptyString(p.name) &&
          typeof p.href === 'string' &&
          /^[\w-]+\.html$/.test(p.href) &&
          isStringArray(p.detectWords) &&
          p.detectWords.some(isNonEmptyString)
            ? [{ name: p.name, href: p.href, keys: p.detectWords.map(labelKey).filter((k) => k !== '') }]
            : [],
        )
      : [];

  /**
   * 文章が、別のページで判定する書類のものらしいとき、そのページ（判別語が1つでもあれば）
   * @param {string} text
   * @param {{ otherPages: readonly OtherPage[] }} ruleset
   * @returns {OtherPage|null}
   */
  const otherPageFor = (text, ruleset) => {
    if (ruleset.otherPages.length === 0) return null;
    const key = labelKey(text);
    return ruleset.otherPages.find((p) => p.keys.some((k) => key.includes(k))) ?? null;
  };

  /**
   * ルール定義を読み込む。1件のルールの書き方が間違っていても、そのルールだけを飛ばす
   * @param {unknown} raw
   * @returns {CompiledRuleset}
   */
  const compileRuleset = (raw) => {
    const meta = isObject(raw) && isObject(raw.meta) ? raw.meta : {};
    /** @type {CompiledRuleset} */
    const base = {
      meta: {
        toolName: textOr(meta.toolName, 'YEG 記載ルールチェッカー'),
        manualName: textOr(meta.manualName, '理事会上程資料の書き方'),
        manualVersion: textOr(meta.manualVersion, '（版の記載なし）'),
        sourceLabel: textOr(meta.sourceLabel, '理事会上程資料の書き方'),
        updated: textOr(meta.updated, '（記載なし）'),
      },
      categories: [],
      excludedPhrases: [],
      notChecked: [],
      aiGuide: readAiGuide(undefined),
      otherPages: [],
      rules: [],
      documents: null,
      amounts: null,
      categoryNotes: Object.freeze({}),
      errors: [],
      fatal: null,
    };
    if (!isObject(raw) || !Array.isArray(raw.rules)) {
      return {
        ...base,
        fatal: 'ルール定義（rules/rules.js・rules/rules_annual.js）を読み込めませんでした。ファイルの書き方を確認してください。',
      };
    }

    /** @type {CompiledRule[]} */
    const rules = [];
    /** @type {RuleLoadError[]} */
    const errors = [];
    const seen = new Set();
    raw.rules.forEach((r, i) => {
      const label = isObject(r) && isNonEmptyString(r.id) ? r.id : `${i + 1}件目のルール`;
      if (!isObject(r)) {
        errors.push({ ruleLabel: label, message: '{ } の形になっていません' });
        return;
      }
      if (r.enabled !== undefined && typeof r.enabled !== 'boolean') {
        errors.push({ ruleLabel: label, message: 'enabled は true か false で書いてください' });
        return;
      }
      if (r.enabled === false) return;
      if (seen.has(label)) {
        errors.push({ ruleLabel: label, message: '同じ番号のルールが2つあります' });
        return;
      }
      const c = compileRule(r, i);
      if (!c.ok) {
        errors.push({ ruleLabel: label, message: c.message });
        return;
      }
      seen.add(label);
      rules.push(c.value);
    });

    const declared = isStringArray(raw.categories) ? raw.categories : [];
    const extra = rules.map((r) => r.category).filter((c) => !declared.includes(c));

    // 書類の種類のデータ。読めなくても表記のチェックは続ける（書類の種類に応じたチェックだけを行わない）
    const docs = raw.documents === undefined ? null : compileDocuments(raw.documents);
    if (docs && !docs.ok) errors.push({ ruleLabel: DOCUMENTS_LABEL, message: docs.message });
    const documents = docs && docs.ok ? docs.value : null;
    if (documents) {
      const typeIds = new Set(documents.types.map((t) => t.id));
      for (const r of rules) {
        const unknown = r.documentTypes.filter((t) => !typeIds.has(t));
        if (unknown.length > 0) {
          errors.push({ ruleLabel: r.id, message: `documentTypes の「${unknown.join('・')}」は documents.types にありません` });
        }
      }
    }

    // 金額チェックのデータ。読めなくてもほかの判定は続ける（金額チェックだけを行わない）
    const amt = raw.amounts === undefined ? null : compileAmounts(raw.amounts);
    if (amt && !amt.ok) errors.push({ ruleLabel: AMOUNTS_LABEL, message: amt.message });

    const notesRaw = isObject(raw.categoryNotes) ? raw.categoryNotes : {};
    /** @type {Record<string, string>} */
    const categoryNotes = {};
    for (const [name, note] of Object.entries(notesRaw)) if (isNonEmptyString(note)) categoryNotes[name] = note;

    return {
      ...base,
      categories: [...new Set([...declared, ...extra])],
      excludedPhrases: isStringArray(raw.excludedPhrases) ? raw.excludedPhrases : [],
      notChecked: isStringArray(raw.notChecked) ? raw.notChecked : [],
      aiGuide: readAiGuide(raw.aiGuide),
      otherPages: readOtherPages(raw.otherPages),
      rules,
      documents,
      amounts: amt && amt.ok ? amt.value : null,
      categoryNotes: Object.freeze(categoryNotes),
      errors,
    };
  };

  // ------------------------------------------------------------------
  // 判定
  // ------------------------------------------------------------------

  /** @param {string} s */
  const normalizeNewlines = (s) => s.replace(/\r\n?/g, '\n');

  /**
   * 文字数（改行は数えない）
   * @param {string} s
   */
  const countChars = (s) => {
    let n = 0;
    for (const ch of s) if (ch !== '\n' && ch !== '\r') n += 1;
    return n;
  };

  /**
   * @param {RegExp} regex
   * @param {string} text
   * @returns {RegExpExecArray[]}
   */
  const execAll = (regex, text) => {
    const re = new RegExp(regex.source, regex.flags);
    /** @type {RegExpExecArray[]} */
    const out = [];
    let m;
    while ((m = re.exec(text)) !== null) {
      if (m[0] === '') {
        re.lastIndex += 1;
        continue;
      }
      out.push(m);
    }
    return out;
  };

  /** @param {Range} a @param {Range} b */
  const overlaps = (a, b) => a.start < b.end && b.start < a.end;

  /**
   * @param {string} template
   * @param {RegExpExecArray} m
   */
  const expandTemplate = (template, m) =>
    template.replace(/\$(\$|&|<([^>]+)>|(\d{1,2}))/g, (whole, token, name, num) => {
      if (token === '$') return '$';
      if (token === '&') return m[0];
      if (name !== undefined) return (m.groups && m.groups[name]) ?? '';
      const v = m[Number(num)];
      return v ?? '';
    });

  /**
   * 改行を取り除いた文章と、元の位置への対応表
   * @param {string} text
   */
  const buildJoinedView = (text) => {
    /** @type {string[]} */
    const chars = [];
    /** @type {number[]} */
    const map = [];
    for (let i = 0; i < text.length; i += 1) {
      const c = text[i];
      if (c === '\n') continue;
      chars.push(c);
      map.push(i);
    }
    return { text: chars.join(''), map };
  };

  /**
   * URL・メールアドレスの範囲
   * @param {string} text
   * @returns {Range[]}
   */
  const findUrlRanges = (text) =>
    execAll(
      // メールアドレスは英数字のかたまりの先頭からだけ探す（長い英数字の並びで遅くならないように）
      /https?:\/\/[^\s　、。「」『』（）()<>"']+|www\.[^\s　、。「」『』（）()<>"']+|(?<![A-Za-z0-9._%+\-])[A-Za-z0-9._%+\-]+@[A-Za-z0-9\-]+(?:\.[A-Za-z0-9\-]+)+/g,
      text,
    ).map((m) => ({ start: m.index, end: m.index + m[0].length }));

  /**
   * 除外する定型文の範囲（改行・空白の違いは無視する）
   * @param {string} text
   * @param {string[]} phrases
   * @returns {Range[]}
   */
  const findPhraseRanges = (text, phrases) => {
    const isSpace = (/** @type {string} */ c) => /[\s　]/.test(c);
    /** @type {string[]} */
    const chars = [];
    /** @type {number[]} */
    const map = [];
    for (let i = 0; i < text.length; i += 1) {
      if (isSpace(text[i])) continue;
      chars.push(text[i]);
      map.push(i);
    }
    const compact = chars.join('');
    return phrases.flatMap((phrase) => {
      const p = [...phrase].filter((c) => !isSpace(c)).join('');
      if (p === '') return [];
      /** @type {Range[]} */
      const found = [];
      let from = 0;
      for (;;) {
        const at = compact.indexOf(p, from);
        if (at < 0) break;
        found.push({ start: map[at], end: map[at + p.length - 1] + 1 });
        from = at + p.length;
      }
      return found;
    });
  };

  /**
   * 範囲が、またいではいけない位置（ファイルのセル・図形・ページの境目の区切り文字）を含むか
   * @param {readonly number[]} breaks 昇順
   * @param {Range} range
   */
  const crossesBreak = (breaks, range) => {
    let lo = 0;
    let hi = breaks.length;
    while (lo < hi) {
      const mid = (lo + hi) >> 1;
      if (breaks[mid] < range.start) lo = mid + 1;
      else hi = mid;
    }
    return lo < breaks.length && breaks[lo] < range.end;
  };

  /**
   * 1つのルールの、文字の並びによる指摘を探す
   * @param {CompiledRule} rule
   * @param {string} text
   * @param {{ text: string, map: number[] }} joined
   * @param {Range[]} excluded
   * @param {readonly number[]} breaks
   * @returns {Hit[]}
   */
  const findPatternHits = (rule, text, joined, excluded, breaks) =>
    rule.patterns.flatMap((p) => {
      /**
       * @param {string} view
       * @param {(s: number, e: number) => Range} toOriginal
       * @param {boolean} mustCrossLine
       * @returns {Hit[]}
       */
      const scan = (view, toOriginal, mustCrossLine) => {
        const exRanges = p.exclude.flatMap((re) =>
          execAll(re, view).map((m) => ({ start: m.index, end: m.index + m[0].length })),
        );
        return execAll(p.regex, view).flatMap((m) => {
          const vs = m.index;
          const ve = vs + m[0].length;
          const range = toOriginal(vs, ve);
          const crosses = text.slice(range.start, range.end).includes('\n');
          if (crosses !== mustCrossLine) return [];
          if (breaks.length > 0 && crossesBreak(breaks, range)) return [];
          if (exRanges.some((r) => overlaps(r, { start: vs, end: ve }))) return [];
          if (excluded.some((r) => overlaps(r, range))) return [];
          if (p.test && !TEST_FUNCTIONS[p.test]({ match: m[0], next: view.charAt(ve) })) return [];
          const fixes = p.fixFn ? FIX_FUNCTIONS[p.fixFn](m[0]) : p.fix.map((t) => expandTemplate(t, m));
          return [{ start: range.start, end: range.end, fixes, hint: rule.hint, note: '', rule }];
        });
      };
      const raw = scan(text, (s, e) => ({ start: s, end: e }), false);
      const cross = p.crossLines
        ? scan(joined.text, (s, e) => ({ start: joined.map[s], end: joined.map[e - 1] + 1 }), true)
        : [];
      return [...raw, ...cross];
    });

  /**
   * 文の終わり方から文体を決める
   * @param {string} sentence 「。」で終わる文
   * @param {StyleConfig} style
   * @returns {'polite'|'plain'|null}
   */
  const classifySentence = (sentence, style) => {
    const body = sentence.replace(/。$/, '').replace(/[」』）)】〕"'\s　]+$/u, '');
    if (body === '') return null;
    const endsWithAny = (/** @type {string[]} */ list) => list.some((e) => e !== '' && body.endsWith(e));
    if (endsWithAny(style.politeEndings)) return 'polite';
    if (endsWithAny(style.neutralEndings)) return null;
    if (endsWithAny(style.plainEndings)) return 'plain';
    const last = [...body].pop() ?? '';
    return style.plainLastChars.includes(last) ? 'plain' : null;
  };

  /**
   * 文を探す範囲（行）。またいではいけない位置があれば、そこでも分ける
   * @param {string} text
   * @param {readonly number[]} breaks
   * @returns {{ start: number, line: string }[]}
   */
  const splitLines = (text, breaks) => {
    /** @type {{ start: number, line: string }[]} */
    const out = [];
    let lineStart = 0;
    let b = 0;
    for (const line of text.split('\n')) {
      const end = lineStart + line.length;
      let from = lineStart;
      while (b < breaks.length && breaks[b] < end) {
        if (breaks[b] >= from) {
          out.push({ start: from, line: text.slice(from, breaks[b]) });
          from = breaks[b] + 1;
        }
        b += 1;
      }
      out.push({ start: from, line: text.slice(from, end) });
      lineStart = end + 1;
    }
    return out;
  };

  /**
   * 文体（ですます調・である調）の数を数え、少ない方の文を指摘する
   * @param {string} text
   * @param {CompiledRule} rule
   * @param {Range[]} excluded
   * @param {readonly number[]} breaks
   */
  const analyzeStyle = (text, rule, excluded, breaks) => {
    const style = /** @type {StyleConfig} */ (rule.style);
    /** @type {{ range: Range, kind: 'polite'|'plain' }[]} */
    const sentences = [];
    for (const { start: lineStart, line } of splitLines(text, breaks)) {
      // 「。」までを1文とする（正規表現で探すと、「。」のない長い行で極端に遅くなるため indexOf で区切る）
      let from = 0;
      for (let end = line.indexOf('。'); end >= 0; end = line.indexOf('。', from)) {
        const raw = line.slice(from, end + 1);
        const lead = raw.length - raw.replace(/^[\s　]+/, '').length;
        const range = { start: lineStart + from + lead, end: lineStart + end + 1 };
        from = end + 1;
        if (range.end - range.start <= 1) continue;
        if (excluded.some((r) => overlaps(r, range))) continue;
        const kind = classifySentence(text.slice(range.start, range.end), style);
        if (kind) sentences.push({ range, kind });
      }
    }
    const polite = sentences.filter((s) => s.kind === 'polite').length;
    const plain = sentences.length - polite;
    /** @type {StyleCounts} */
    const counts = { polite, plain, politeLabel: style.politeLabel, plainLabel: style.plainLabel };
    if (polite === 0 || plain === 0) return { counts, hits: /** @type {Hit[]} */ ([]) };

    const breakdown = `${style.politeLabel} ${polite}文／${style.plainLabel} ${plain}文`;
    const majority = polite > plain ? style.politeLabel : style.plainLabel;
    const note =
      polite === plain
        ? `「${style.politeLabel}」と「${style.plainLabel}」の文が同じ数です（${breakdown}）。どちらかにそろえてください。`
        : `この文章は「${majority}」の文が多いので（${breakdown}）、この文も「${majority}」にそろえてください。`;
    const hint = polite === plain ? 'どちらかの文体にそろえる' : `${majority}にそろえる`;
    const target = polite === plain ? null : polite < plain ? 'polite' : 'plain';
    /** @type {Hit[]} */
    const hits = sentences
      .filter((s) => target === null || s.kind === target)
      .map((s) => ({ start: s.range.start, end: s.range.end, fixes: [], hint, note, rule }));
    return { counts, hits };
  };

  /**
   * 同じルールで重なった指摘は、先に見つかった1件だけにする
   * @param {Hit[]} hits
   */
  const dedupeSameRule = (hits) => {
    const sorted = [...hits].sort(
      (a, b) => a.rule.order - b.rule.order || a.start - b.start || b.end - b.start - (a.end - a.start),
    );
    /** @type {Map<string, Range[]>} */
    const kept = new Map();
    return sorted.filter((h) => {
      const list = kept.get(h.rule.id) ?? [];
      if (list.some((r) => overlaps(r, h))) return false;
      kept.set(h.rule.id, [...list, { start: h.start, end: h.end }]);
      return true;
    });
  };

  /** @param {Severity} s */
  const severityRank = (s) => SEVERITIES.indexOf(s);

  /**
   * @param {Hit} h
   * @returns {Reason}
   */
  const toReason = (h) => ({
    ruleId: h.rule.id,
    ruleName: h.rule.name,
    category: h.rule.category,
    severity: h.severity ?? h.rule.severity,
    fixes: h.fixes,
    hint: h.hint,
    description: h.rule.description,
    source: h.rule.source,
    note: h.note,
    order: h.rule.order,
  });

  /**
   * @param {Reason[]} reasons
   * @returns {Reason[]}
   */
  const sortReasons = (reasons) =>
    [...reasons].sort((a, b) => severityRank(a.severity) - severityRank(b.severity) || a.order - b.order);

  /**
   * 位置から欄の名前を引く関数を作る（欄が重なっているときは、後から始まった内側の欄）
   * @param {readonly Section[]} sections
   * @returns {(pos: number) => string|null}
   */
  const makeSectionLookup = (sections) => {
    const sorted = sections
      .filter((s) => Number.isInteger(s.start) && Number.isInteger(s.end) && s.end > s.start && isNonEmptyString(s.name))
      .sort((a, b) => a.start - b.start || b.end - a.end);
    return (pos) => {
      let lo = 0;
      let hi = sorted.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (sorted[mid].start <= pos) lo = mid + 1;
        else hi = mid;
      }
      for (let i = lo - 1; i >= 0; i -= 1) if (sorted[i].end > pos) return sorted[i].name;
      return null;
    };
  };

  /**
   * 欄に応じて重さを変えるルール（sectionSeverity）の指摘を、欄に合わせて直す（「指摘しない」は取り除く）
   * @param {Hit[]} hits
   * @param {(pos: number) => string|null} sectionAt
   * @returns {Hit[]}
   */
  const applySectionSeverity = (hits, sectionAt) =>
    hits.flatMap((h) => {
      const table = h.rule.sectionSeverity;
      if (!table) return [h];
      const name = sectionAt(h.start);
      if (name === null) return [h];
      const sev = table[name] ?? table[OTHER_SECTIONS] ?? h.rule.severity;
      if (sev === NOT_REPORTED) return [];
      if (sev === h.rule.severity) return [h];
      return [{ ...h, severity: sev, note: `「${name}」の欄にあるため「${sev}」にしています。` }];
    });

  /**
   * 同じ箇所の指摘を1つにまとめ、表示用の指摘を作る
   * @param {string} text
   * @param {Hit[]} wordHits
   * @param {Hit[]} sentenceHits
   * @param {(pos: number) => string|null} sectionAt
   * @returns {Finding[]}
   */
  const buildFindings = (text, wordHits, sentenceHits, sectionAt) => {
    const sorted = [...wordHits].sort((a, b) => a.start - b.start || a.rule.order - b.rule.order);
    /** @type {{ start: number, end: number, hits: Hit[] }[]} */
    const clusters = [];
    for (const h of sorted) {
      const last = clusters[clusters.length - 1];
      if (last && h.start < last.end) {
        last.end = Math.max(last.end, h.end);
        last.hits.push(h);
      } else {
        clusters.push({ start: h.start, end: h.end, hits: [h] });
      }
    }

    /**
     * @param {'word'|'sentence'} kind
     * @param {number} start
     * @param {number} end
     * @param {Hit[]} hits
     */
    const make = (kind, start, end, hits) => {
      /** @type {Map<string, Reason>} */
      const byRule = new Map();
      for (const h of hits) {
        const prev = byRule.get(h.rule.id);
        byRule.set(
          h.rule.id,
          prev ? { ...prev, fixes: [...new Set([...prev.fixes, ...h.fixes])] } : toReason(h),
        );
      }
      const reasons = sortReasons([...byRule.values()]);
      const matched = text.slice(start, end);
      const before = text.slice(Math.max(0, start - CONTEXT_CHARS), start);
      const after = text.slice(end, end + CONTEXT_CHARS);
      return {
        kind,
        range: { start, end },
        text: matched,
        before,
        after,
        reasons,
        severity: reasons[0].severity,
        key: [reasons.map((r) => r.ruleId).join('+'), matched, before, after].join('\u0000'),
        section: sectionAt(start),
      };
    };

    const all = [
      ...clusters.map((c) => make('word', c.start, c.end, c.hits)),
      ...sentenceHits.map((h) => make('sentence', h.start, h.end, [h])),
    ].sort(
      (a, b) =>
        a.range.start - b.range.start ||
        (a.kind === b.kind ? 0 : a.kind === 'sentence' ? -1 : 1) ||
        a.reasons[0].order - b.reasons[0].order,
    );
    return all.map((f, i) => ({ id: `f-${i + 1}`, ...f }));
  };

  /**
   * @typedef {Object} CheckOptions
   * @property {readonly number[]} [breaks]  語や文を探すときにまたがない位置（ファイルのセル・図形・ページの境目に置いた
   *   区切り文字の位置）。改行を \n にそろえた文章での位置で、昇順に渡す。貼り付けた文章では使わない
   * @property {readonly Section[]} [sections]  認識できた欄の範囲（doccheck.js の analyzeDocument で作る）。
   *   欄に応じて重さを変えるルール（R-13 など）と、指摘の欄名の表示に使う。渡さなければ、どの箇所も「欄が分からない」として扱う
   * @property {readonly string[]} [documentTypes]  書類の種類。書類の種類を決めているルール（documentTypes）は、ここに含まれるときだけ当てる
   */

  /**
   * 文章を判定する（同じ入力・同じルールなら必ず同じ結果になる）
   * @param {string} input
   * @param {CompiledRuleset} ruleset
   * @param {CheckOptions} [options]
   * @returns {CheckResult}
   */
  const checkText = (input, ruleset, options = {}) => {
    const text = normalizeNewlines(input);
    const charCount = countChars(text);
    if (text.trim() === '') return { ok: false, reason: 'empty', charCount };
    if (charCount > MAX_CHARS) return { ok: false, reason: 'too-long', charCount };
    const breaks = Array.isArray(options.breaks) ? [...options.breaks].sort((a, b) => a - b) : [];
    const sectionAt = makeSectionLookup(Array.isArray(options.sections) ? options.sections : []);

    const joined = buildJoinedView(text);
    const phraseRanges = findPhraseRanges(text, ruleset.excludedPhrases);
    const excluded = [...findUrlRanges(text), ...phraseRanges];

    // 書類の種類を決めているルール（documentTypes。委員会事業計画書の用語など）は、その種類を指定されたときだけ当てる
    const allowedTypes = Array.isArray(options.documentTypes) ? options.documentTypes : [];
    const wordHits = ruleset.rules
      .filter((r) => r.type === 'pattern' && (r.documentTypes.length === 0 || r.documentTypes.some((t) => allowedTypes.includes(t))))
      .flatMap((r) => findPatternHits(r, text, joined, excluded, breaks));
    const styleRule = ruleset.rules.find((r) => r.type === 'style');
    const style = styleRule ? analyzeStyle(text, styleRule, phraseRanges, breaks) : null;

    return {
      ok: true,
      text,
      charCount,
      findings: buildFindings(
        text,
        applySectionSeverity(dedupeSameRule(wordHits), sectionAt),
        style ? style.hits : [],
        sectionAt,
      ),
      styleCounts: style ? style.counts : null,
    };
  };

  /**
   * 書類全体を見て判定するルール（type: 'document'。野外事業の中止基準など）。
   * 対象の書類の種類（typeIds）に当てはまるルールだけを判定し、位置を持たない指摘（書類全体）を返す
   * @param {string} input
   * @param {CompiledRuleset} ruleset
   * @param {readonly string[]} typeIds
   * @returns {Finding[]}
   */
  const checkDocumentRules = (input, ruleset, typeIds) => {
    const text = normalizeNewlines(input);
    const ids = new Set(typeIds);
    /** @param {RegExp} re */
    const has = (re) => text.search(re) >= 0;
    return ruleset.rules
      .filter((r) => r.type === 'document' && r.condition && r.documentTypes.some((t) => ids.has(t)))
      .filter((r) => {
        const c = /** @type {DocumentCondition} */ (r.condition);
        return c.anyOf.some(has) && !c.noneOf.some(has);
      })
      .map((r, i) => {
        /** @type {Reason} */
        const reason = {
          ruleId: r.id,
          ruleName: r.name,
          category: r.category,
          severity: r.severity,
          fixes: [],
          hint: r.hint,
          description: r.description,
          source: r.source,
          note: r.message,
          order: r.order,
        };
        return {
          id: `d-${i + 1}`,
          kind: /** @type {'document'} */ ('document'),
          range: null,
          text: '',
          before: '',
          after: '',
          reasons: [reason],
          severity: r.severity,
          key: [r.id, '', '', ''].join('\u0000'),
          section: null,
        };
      });
  };

  /**
   * 入力から読み取った語（自分の委員会名など）で探すルールだけを取り出し、その語を見つける語にしたルール定義を作る。
   * 返したルール定義で checkText すると、そのルールの指摘だけが出る（語がなければルールは0件）
   * @param {CompiledRuleset} ruleset
   * @param {string} kind     DYNAMIC_WORDS のどれか
   * @param {readonly string[]} words
   * @returns {CompiledRuleset}
   */
  const withDynamicWords = (ruleset, kind, words) => {
    const list = [...new Set(words.filter(isNonEmptyString))].sort((a, b) => b.length - a.length);
    /** @type {CompiledRule[]} */
    const rules = [];
    if (list.length > 0) {
      for (const r of ruleset.rules) {
        if (r.type !== 'pattern' || r.dynamicWords !== kind) continue;
        const p = compilePattern({ words: list, fix: r.fixes }, 0);
        if (p.ok) rules.push({ ...r, patterns: [...r.patterns, p.value] });
      }
    }
    return { ...ruleset, rules };
  };

  const api = Object.freeze({
    SEVERITIES,
    MAX_CHARS,
    OTHER_SECTIONS,
    AMOUNT_CHECKS,
    DOCUMENT_CHECKS,
    DYNAMIC_WORDS,
    compileRuleset,
    checkText,
    checkDocumentRules,
    withDynamicWords,
    otherPageFor,
    normalizeLabel,
    labelKey,
    countChars,
    normalizeNewlines,
    toHalfWidth,
    addThousands,
    isBadGrouping,
    parseWesternDate,
    formatWareki,
    classifySentence,
  });

  root.YegEngine = api;
})(globalThis);
