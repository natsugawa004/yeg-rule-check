// @ts-check
/*
 * 書類の種類ごとの文章ルールの扱いと、委員会事業計画書（年度。仕様書 付録G）の構成のチェック（画面に依存しない純粋な関数）
 *   committeeNames(文章, 種類) → 見出しの「〇〇委員会　委員長名」から読み取った自分の委員会名（読み仮名のかっこを外した形も）
 *   adjustFindings(指摘, 範囲, ルール) → 種類に応じて、判定する欄に絞り・当てないルールを外し・「違反」を「要確認」に下げる
 *   bodyParagraphs(文章, 欄, 欄の名前) → 本文の段落（見出しが見つからなければ全文）
 *   layout(段落, ルール, 合計の上限) → 上段・中段・下段・合計の字数の表
 *   structureFindings(段落, 種類, ルール) → 構成のチェック（段ごとの字数・合計の字数・書き出し・締め）の指摘（書類全体）
 */
'use strict';

(function (root) {
  /** @typedef {'違反'|'要確認'|'参考'} Severity */
  /** @typedef {{ start: number, end: number, name: string }} Section */
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
   * @property {'word'|'sentence'|'document'|'amount'} kind
   * @property {{ start: number, end: number }|null} range
   * @property {string} text
   * @property {string} before
   * @property {string} after
   * @property {Reason[]} reasons
   * @property {Severity} severity
   * @property {string} key
   * @property {string|null} section
   */
  /**
   * engine.js の CompiledRule・DocType（使う項目だけ）
   * @typedef {{ id: string, name: string, category: string, severity: Severity, type: string, documentTypes: string[], check: string,
   *   messages: Readonly<Record<string, string>>, params: Readonly<Record<string, unknown>>, hint: string, description: string, source: string, order: number }} Rule
   * @typedef {{ id: string, name: string, scopeSections: string[], textRules: { off: string[], lower: boolean }, committeeNamePattern: RegExp|null }} DocType
   * @typedef {{ rules: Rule[], documents: { types: DocType[] }|null }} Ruleset
   */
  /**
   * 指摘を調整する範囲（ファイル1つ、または貼り付けた文章）
   * @typedef {Object} Region
   * @property {number} start
   * @property {number} end
   * @property {string[]} typeIds        この範囲を判定に使う書類の種類
   * @property {boolean} hasScope        判定する欄（scopeSections）の見出しが見つかったか。見つからなければ欄で絞らない
   */
  /** @typedef {{ label: string, chars: number, target: number, status: 'ok'|'short'|'long' }} LayoutRow */
  /** @typedef {{ split: boolean, rows: LayoutRow[], total: LayoutRow }} Layout */

  const SEVERITY_RANK = Object.freeze({ 違反: 0, 要確認: 1, 参考: 2 });
  const LOWERED_NOTE = '委員会事業計画書のため「要確認」にしています（議案では違反）。';

  /** かっこ（読み仮名など）を外す: 「evol（エヴォル）研修委員会」→「evol研修委員会」 */
  const withoutBrackets = (/** @type {string} */ s) => s.replace(/[（(][^（()）]*[）)]/g, '').replace(/[ 　]+/g, ' ').trim();

  /**
   * 見出しの「〇〇委員会　委員長名　氏名」の行から、自分の委員会名を読み取る（最初に見つかった行）
   * @param {string} text
   * @param {DocType} type
   * @returns {string[]}  委員会名と、かっこを外した形（4文字以上で「委員会」を含むものだけ）
   */
  const committeeNames = (text, type) => {
    if (!type.committeeNamePattern) return [];
    const re = new RegExp(type.committeeNamePattern.source, type.committeeNamePattern.flags.replace(/[gy]/g, ''));
    for (const line of text.split('\n')) {
      const m = re.exec(line);
      if (!m || typeof m[1] !== 'string') continue;
      const name = m[1].replace(/[ 　]+/g, ' ').trim();
      return [...new Set([name, withoutBrackets(name)])].filter((n) => n.length >= 4 && n.includes('委員会'));
    }
    return [];
  };

  /**
   * 書類の種類に応じて文章の指摘を調整する
   *  - 種類を決めているルール（documentTypes）は、その種類の範囲でだけ出す
   *  - 判定する欄（scopeSections）がある種類は、その欄の中の指摘だけを出す（見出し・ひな形の部分は判定しない）
   *  - textRules.off のルールは出さず、textRules.lower なら書類の種類を決めていないルール（議案の表記ルール）の「違反」を「要確認」に下げる
   * @param {readonly Finding[]} findings
   * @param {readonly Region[]} regions
   * @param {Ruleset} ruleset
   * @returns {Finding[]}
   */
  const adjustFindings = (findings, regions, ruleset) => {
    const types = new Map((ruleset.documents ? ruleset.documents.types : []).map((t) => [t.id, t]));
    const rules = new Map(ruleset.rules.map((r) => [r.id, r]));
    return findings.flatMap((f) => {
      const range = f.range;
      const region = range ? regions.find((r) => range.start >= r.start && range.start < r.end) : undefined;
      const typeIds = region ? region.typeIds : [];
      const selected = typeIds.flatMap((id) => {
        const t = types.get(id);
        return t ? [t] : [];
      });
      const scope = selected.flatMap((t) => t.scopeSections);
      if (region && region.hasScope && scope.length > 0 && !(f.section !== null && scope.includes(f.section))) return [];
      const off = new Set(selected.flatMap((t) => t.textRules.off));
      const lower = selected.some((t) => t.textRules.lower);
      let changed = false;
      const reasons = f.reasons.flatMap((r) => {
        const rule = rules.get(r.ruleId);
        const only = rule ? rule.documentTypes : [];
        if (rule && rule.type === 'pattern' && only.length > 0 && !only.some((id) => typeIds.includes(id))) {
          changed = true;
          return [];
        }
        if (off.has(r.ruleId)) {
          changed = true;
          return [];
        }
        if (lower && r.severity === '違反' && only.length === 0) {
          changed = true;
          return [{ ...r, severity: /** @type {Severity} */ ('要確認'), note: r.note ? `${r.note} ${LOWERED_NOTE}` : LOWERED_NOTE }];
        }
        return [r];
      });
      if (!changed) return [f];
      if (reasons.length === 0) return [];
      const sorted = [...reasons].sort((a, b) => SEVERITY_RANK[a.severity] - SEVERITY_RANK[b.severity] || a.order - b.order);
      return [
        {
          ...f,
          reasons: sorted,
          severity: sorted[0].severity,
          key: [sorted.map((r) => r.ruleId).join('+'), f.text, f.before, f.after].join('\u0000'),
        },
      ];
    });
  };

  /**
   * 本文の段落（見出しの欄が見つからなければ、文章全体を本文とする）
   * @param {string} text
   * @param {readonly Section[]} sections
   * @param {string} sectionName
   * @returns {{ start: number, end: number, text: string }[]}
   */
  const bodyParagraphs = (text, sections, sectionName) => {
    const own = sections.filter((s) => s.name === sectionName).sort((a, b) => a.start - b.start);
    const ranges = own.length > 0 ? own : [{ start: 0, end: text.length, name: sectionName }];
    /** @type {{ start: number, end: number, text: string }[]} */
    const out = [];
    for (const r of ranges) {
      let pos = r.start;
      for (const line of text.slice(r.start, r.end).split('\n')) {
        const body = line.trim();
        // 見出しの行そのもの（欄の名前だけの行）は段落にしない
        if (body !== '' && body !== sectionName) out.push({ start: pos, end: pos + line.length, text: body });
        pos += line.length + 1;
      }
    }
    return out;
  };

  /** 字数（空白・改行を数えない） */
  const charsOf = (/** @type {string} */ s) => [...s.replace(/\s/g, '')].length;

  /**
   * @param {Readonly<Record<string, unknown>>} params
   * @returns {{ parts: { label: string, target: number }[], total: number, low: number, high: number }}
   */
  const layoutParams = (params) => {
    const parts = Array.isArray(params.parts)
      ? params.parts.flatMap((p) =>
          typeof p === 'object' && p !== null && typeof p.label === 'string' && typeof p.target === 'number' ? [{ label: p.label, target: p.target }] : [],
        )
      : [];
    const num = (/** @type {unknown} */ v, /** @type {number} */ d) => (typeof v === 'number' && v > 0 ? v : d);
    return {
      parts,
      total: num(params.total, parts.reduce((a, p) => a + p.target, 0)),
      low: num(params.low, 0.5),
      high: num(params.high, 1.5),
    };
  };

  /** 合計の字数の上限（check: 'total' のルールの params.max。書いていなければ 0） */
  const totalMaxOf = (/** @type {Readonly<Record<string, unknown>>} */ params) =>
    typeof params.max === 'number' && params.max > 0 ? params.max : 0;

  /** 段落の字数の合計 */
  const totalChars = (/** @type {readonly { text: string }[]} */ paragraphs) => paragraphs.reduce((a, x) => a + charsOf(x.text), 0);

  /**
   * 上段・中段・下段・合計の字数（段落が3つ以上なら、最初＝上段・最後＝下段・間＝中段）
   * @param {readonly { text: string }[]} paragraphs
   * @param {Readonly<Record<string, unknown>>} params  字数の目安（length のルールの params）
   * @param {number} [totalMax]  合計の上限（total のルールの params.max）。0 より大きければ、合計の行はこれを超えたときだけ「多め」にする
   * @returns {Layout|null}
   */
  const layout = (paragraphs, params, totalMax = 0) => {
    if (paragraphs.length === 0) return null;
    const p = layoutParams(params);
    /** @param {string} label @param {number} chars @param {number} target @returns {LayoutRow} */
    const row = (label, chars, target) => ({
      label,
      chars,
      target,
      status: chars < target * p.low ? 'short' : chars > target * p.high ? 'long' : 'ok',
    });
    const all = totalChars(paragraphs);
    const split = paragraphs.length >= 3 && p.parts.length === 3;
    const rows = split
      ? [
          row(p.parts[0].label, charsOf(paragraphs[0].text), p.parts[0].target),
          row(p.parts[1].label, paragraphs.slice(1, -1).reduce((a, x) => a + charsOf(x.text), 0), p.parts[1].target),
          row(p.parts[2].label, charsOf(paragraphs[paragraphs.length - 1].text), p.parts[2].target),
        ]
      : [];
    /** @type {LayoutRow} */
    const total =
      totalMax > 0 ? { label: '合計', chars: all, target: totalMax, status: all > totalMax ? 'long' : 'ok' } : row('合計', all, p.total);
    return { split, rows, total };
  };

  /** 指摘の文の {名前} を埋める */
  const fill = (/** @type {string} */ template, /** @type {Record<string, string|number>} */ values) =>
    template.replace(/\{(\w+)\}/g, (m, k) => (k in values ? String(values[k]) : m));

  /**
   * 構成のチェック（type: 'document' で check を書いたルール）の指摘
   * @param {readonly { text: string }[]} paragraphs
   * @param {string} typeId
   * @param {Ruleset} ruleset
   * @returns {Finding[]}
   */
  const structureFindings = (paragraphs, typeId, ruleset) => {
    if (paragraphs.length === 0) return [];
    /** @type {Finding[]} */
    const out = [];
    /** @param {Rule} r @param {string} note @param {string} detail */
    const add = (r, note, detail) => {
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
        note,
        order: r.order,
      };
      out.push({
        id: '',
        kind: 'document',
        range: null,
        text: '',
        before: '',
        after: '',
        reasons: [reason],
        severity: r.severity,
        key: [r.id, detail, '', ''].join('\u0000'),
        section: null,
      });
    };
    const rules = ruleset.rules.filter((r) => r.type === 'document' && r.check !== '' && r.documentTypes.includes(typeId));
    for (const r of rules) {
      const message = r.messages.default ?? '';
      if (r.check === 'length') {
        // 段ごとの字数（段落が3つ未満で段に分けられないときは見ない。合計は check: 'total' のルールで見る）
        const lay = layout(paragraphs, r.params);
        if (!lay || !lay.split) continue;
        for (const x of lay.rows) {
          if (x.status === 'ok') continue;
          const how = x.status === 'short' ? '少なめ' : '多め';
          add(r, fill(message, { label: x.label, chars: x.chars, target: x.target, how }), x.label);
        }
      } else if (r.check === 'total') {
        const max = totalMaxOf(r.params);
        const chars = totalChars(paragraphs);
        if (max > 0 && chars > max) add(r, fill(message, { chars, max }), 'total');
      } else if (r.check === 'opening') {
        if (paragraphs.length < 3) continue;
        const opening = typeof r.params.opening === 'string' ? r.params.opening : '';
        const middle = paragraphs[1].text;
        if (opening !== '' && !middle.startsWith(opening)) add(r, fill(message, { opening }), 'opening');
      } else if (r.check === 'closing') {
        const closings = Array.isArray(r.params.closings) ? r.params.closings.filter((c) => typeof c === 'string' && c !== '') : [];
        const last = paragraphs[paragraphs.length - 1].text.trim();
        if (closings.length > 0 && !closings.some((c) => last.endsWith(c))) {
          add(r, fill(message, { closings: closings.map((c) => `「${c}」`).join('・') }), 'closing');
        }
      }
    }
    return out;
  };

  root.YegAnnual = Object.freeze({
    committeeNames,
    adjustFindings,
    bodyParagraphs,
    layout,
    totalMaxOf,
    structureFindings,
    charsOf,
  });
})(globalThis);
