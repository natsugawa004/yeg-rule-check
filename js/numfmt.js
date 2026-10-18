// @ts-check
/*
 * Excel の表示形式（桁区切り・¥記号・日付・和暦など）で、数値のセルを画面に見える文字にする
 * （画面に依存しない純粋な関数だけを置く）
 *   formatNumber(数値, 表示形式, { date1904 }) → 見える文字（例: 100000 と "#,##0" → "100,000"）
 *   formatText(文字, 表示形式) → 見える文字（表示形式に文字用の区画があるとき）
 * Excel の表示をすべて再現するものではない。判定に効く「桁区切り・通貨・日付・パーセント」を正しく出すことを優先する。
 */
'use strict';

(function (root) {
  /**
   * 番号だけで指定される組み込みの表示形式（日本語版 Excel の値）
   * @type {Readonly<Record<number, string>>}
   */
  const BUILTIN_FORMATS = Object.freeze({
    0: 'General',
    1: '0',
    2: '0.00',
    3: '#,##0',
    4: '#,##0.00',
    5: '"¥"#,##0;"¥"\\-#,##0',
    6: '"¥"#,##0;[Red]"¥"\\-#,##0',
    7: '"¥"#,##0.00;"¥"\\-#,##0.00',
    8: '"¥"#,##0.00;[Red]"¥"\\-#,##0.00',
    9: '0%',
    10: '0.00%',
    11: '0.00E+00',
    12: '# ?/?',
    13: '# ??/??',
    14: 'yyyy/m/d',
    15: 'd-mmm-yy',
    16: 'd-mmm',
    17: 'mmm-yy',
    18: 'h:mm AM/PM',
    19: 'h:mm:ss AM/PM',
    20: 'h:mm',
    21: 'h:mm:ss',
    22: 'yyyy/m/d h:mm',
    27: '[$-411]ge.m.d',
    28: '[$-411]ggge"年"m"月"d"日"',
    29: '[$-411]ggge"年"m"月"d"日"',
    30: 'm/d/yy',
    31: 'yyyy"年"m"月"d"日"',
    32: 'h"時"mm"分"',
    33: 'h"時"mm"分"ss"秒"',
    34: 'yyyy"年"m"月"',
    35: 'm"月"d"日"',
    36: '[$-411]ge.m.d',
    37: '#,##0 ;(#,##0)',
    38: '#,##0 ;[Red](#,##0)',
    39: '#,##0.00;(#,##0.00)',
    40: '#,##0.00;[Red](#,##0.00)',
    41: '_ * #,##0_ ;_ * \\-#,##0_ ;_ * "-"_ ;_ @_ ',
    42: '_ "¥"* #,##0_ ;_ "¥"* \\-#,##0_ ;_ "¥"* "-"_ ;_ @_ ',
    43: '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ ',
    44: '_ "¥"* #,##0.00_ ;_ "¥"* \\-#,##0.00_ ;_ "¥"* "-"??_ ;_ @_ ',
    45: 'mm:ss',
    46: '[h]:mm:ss',
    47: 'mm:ss.0',
    48: '##0.0E+0',
    49: '@',
    50: '[$-411]ge.m.d',
    51: '[$-411]ggge"年"m"月"d"日"',
    52: 'yyyy"年"m"月"',
    53: 'm"月"d"日"',
    54: '[$-411]ggge"年"m"月"d"日"',
    55: 'yyyy"年"m"月"',
    56: 'm"月"d"日"',
    57: '[$-411]ge.m.d',
    58: '[$-411]ggge"年"m"月"d"日"',
  });

  /**
   * @typedef {{ t: 'lit', v: string }
   *   | { t: 'ph', v: string }
   *   | { t: 'dot' } | { t: 'comma' } | { t: 'pct' } | { t: 'slash' } | { t: 'at' } | { t: 'general' }
   *   | { t: 'exp', v: string }
   *   | { t: 'date', v: string }
   *   | { t: 'elapsed', v: string }
   *   | { t: 'ampm', v: string }} Token
   */

  /**
   * @typedef {Object} Section
   * @property {Token[]} tokens
   * @property {{ op: string, value: number }|null} cond
   * @property {boolean} gannen  令和元年のように「元」と出す
   */

  /**
   * 表示形式を「;」で区画に分ける（引用符・角括弧・\ の後の「;」では分けない）
   * @param {string} code
   * @returns {string[]}
   */
  const splitSections = (code) => {
    /** @type {string[]} */
    const out = [];
    let cur = '';
    let quoted = false;
    let bracket = false;
    for (let i = 0; i < code.length; i += 1) {
      const c = code[i];
      if (!quoted && !bracket && c === '\\') {
        cur += c + (code[i + 1] ?? '');
        i += 1;
        continue;
      }
      if (c === '"' && !bracket) quoted = !quoted;
      else if (c === '[' && !quoted) bracket = true;
      else if (c === ']' && !quoted) bracket = false;
      if (c === ';' && !quoted && !bracket) {
        out.push(cur);
        cur = '';
      } else {
        cur += c;
      }
    }
    out.push(cur);
    return out;
  };

  /**
   * 1つの区画を字句に分ける
   * @param {string} s
   * @returns {Section}
   */
  const tokenize = (s) => {
    /** @type {Token[]} */
    const tokens = [];
    /** @type {{ op: string, value: number }|null} */
    let cond = null;
    let gannen = false;
    /** @param {string} v */
    const lit = (v) => {
      if (v !== '') tokens.push({ t: 'lit', v });
    };
    let i = 0;
    while (i < s.length) {
      const c = s[i];
      const rest = s.slice(i);
      if (c === '"') {
        const j = s.indexOf('"', i + 1);
        lit(s.slice(i + 1, j < 0 ? s.length : j));
        i = j < 0 ? s.length : j + 1;
      } else if (c === '\\') {
        lit(s[i + 1] ?? '');
        i += 2;
      } else if (c === '_' || c === '*') {
        // _x は x の幅の空白、*x は x で埋める指定。判定に使う文字としては何も出さない
        i += 2;
      } else if (c === '[') {
        const j = s.indexOf(']', i);
        const inner = s.slice(i + 1, j < 0 ? s.length : j);
        i = j < 0 ? s.length : j + 1;
        const currency = /^\$([^-]*)(?:-(.*))?$/.exec(inner);
        const condition = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)$/.exec(inner);
        if (currency) {
          lit(currency[1]);
          if (/gannen/i.test(currency[2] ?? '')) gannen = true;
        } else if (condition) {
          cond = { op: condition[1], value: Number(condition[2]) };
        } else if (/^(h+|m+|s+)$/i.test(inner)) {
          tokens.push({ t: 'elapsed', v: inner.toLowerCase()[0] });
        }
        // 色（[Red]）・漢数字（[DBNum1]）などは文字にしない
      } else if (/^general/i.test(rest)) {
        tokens.push({ t: 'general' });
        i += 7;
      } else if (/^am\/pm/i.test(rest)) {
        tokens.push({ t: 'ampm', v: 'AM/PM' });
        i += 5;
      } else if (/^a\/p/i.test(rest)) {
        tokens.push({ t: 'ampm', v: 'A/P' });
        i += 3;
      } else if (c === '0' || c === '#' || c === '?') {
        tokens.push({ t: 'ph', v: c });
        i += 1;
      } else if (c === '.') {
        tokens.push({ t: 'dot' });
        i += 1;
      } else if (c === ',') {
        tokens.push({ t: 'comma' });
        i += 1;
      } else if (c === '%') {
        tokens.push({ t: 'pct' });
        i += 1;
      } else if ((c === 'E' || c === 'e') && (s[i + 1] === '+' || s[i + 1] === '-')) {
        tokens.push({ t: 'exp', v: s[i + 1] });
        i += 2;
      } else if (c === '/') {
        tokens.push({ t: 'slash' });
        i += 1;
      } else if (c === '@') {
        tokens.push({ t: 'at' });
        i += 1;
      } else if (/[ymdhsgeaYMDHSGEA]/.test(c)) {
        let j = i + 1;
        while (j < s.length && s[j].toLowerCase() === c.toLowerCase()) j += 1;
        const run = s.slice(i, j).toLowerCase();
        if (run[0] === 'a' && run.length < 3) lit(s.slice(i, j));
        else tokens.push({ t: 'date', v: run });
        i = j;
      } else {
        lit(c);
        i += 1;
      }
    }
    return { tokens, cond, gannen };
  };

  /** @param {Token[]} tokens */
  const isDateSection = (tokens) => tokens.some((t) => t.t === 'date' || t.t === 'elapsed' || t.t === 'ampm');

  // ------------------------------------------------------------------
  // 数値
  // ------------------------------------------------------------------

  /**
   * 0 以上の数を、指数を使わない10進の文字にする（Excel と同じく有効数字15桁）
   * @param {number} v
   */
  const decimalString = (v) => {
    const s = Number(v.toPrecision(15)).toString();
    const m = /^(\d+)(?:\.(\d+))?e([+-]\d+)$/i.exec(s);
    if (!m) return s;
    const digits = m[1] + (m[2] ?? '');
    const point = m[1].length + Number(m[3]);
    if (point <= 0) return `0.${'0'.repeat(-point)}${digits}`;
    if (point >= digits.length) return digits + '0'.repeat(point - digits.length);
    return `${digits.slice(0, point)}.${digits.slice(point)}`;
  };

  /**
   * 数字の文字列に1を足す（例: '199' → '200'）
   * @param {string} digits
   */
  const incrementDigits = (digits) => {
    const arr = [...digits];
    for (let i = arr.length - 1; i >= 0; i -= 1) {
      if (arr[i] === '9') {
        arr[i] = '0';
      } else {
        arr[i] = String(Number(arr[i]) + 1);
        return arr.join('');
      }
    }
    return `1${arr.join('')}`;
  };

  /**
   * 0 以上の数を小数点以下 n 桁に丸める（四捨五入）。整数部と小数部の文字列を返す
   * @param {number} v
   * @param {number} n
   * @returns {[string, string]}
   */
  const roundDigits = (v, n) => {
    const [ip, fp = ''] = decimalString(v).split('.');
    if (fp.length <= n) return [ip, fp.padEnd(n, '0')];
    let digits = ip + fp.slice(0, n);
    if (fp[n] >= '5') digits = incrementDigits(digits);
    const intPart = digits.slice(0, digits.length - n).replace(/^0+(?=\d)/, '') || '0';
    return [intPart, digits.slice(digits.length - n)];
  };

  /** @param {string} digits */
  const groupThousands = (digits) => digits.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

  /**
   * 「標準」の表示（Excel の「G/標準」）
   * @param {number} v
   */
  const formatGeneral = (v) => {
    if (!Number.isFinite(v)) return String(v);
    const abs = Math.abs(v);
    if (abs === 0) return '0';
    if (abs >= 1e11 || abs < 1e-9) {
      const [m, e] = v.toExponential(5).split('e');
      const mant = m.replace(/\.?0+$/, '');
      const exp = Number(e);
      return `${mant}E${exp < 0 ? '-' : '+'}${String(Math.abs(exp)).padStart(2, '0')}`;
    }
    if (Number.isInteger(v)) return String(v);
    const sign = v < 0 ? '-' : '';
    const intLen = Math.max(1, Math.floor(Math.log10(abs)) + 1);
    const [ip, fp] = roundDigits(abs, Math.max(0, 10 - intLen));
    const frac = fp.replace(/0+$/, '');
    return `${sign}${ip}${frac ? `.${frac}` : ''}`;
  };

  /** @param {Token} t */
  const isPh = (t) => t.t === 'ph';

  /**
   * 数字を置く位置に、整数部の数字を右から入れていく
   * @param {string} digits  整数部（0 のときは空文字）
   * @param {Token[]} tokens 整数部の字句（最初の数字の位置から小数点の手前まで）
   * @param {boolean} grouping
   */
  const fillInteger = (digits, tokens, grouping) => {
    const firstPh = tokens.findIndex(isPh);
    const hasLiteral = tokens.some((t) => t.t === 'lit');
    let rest = digits;
    /** @type {string[]} */
    const out = [];
    for (let k = tokens.length - 1; k >= 0; k -= 1) {
      const t = tokens[k];
      if (t.t === 'ph') {
        if (rest.length > 0) {
          out.unshift(rest.slice(-1));
          rest = rest.slice(0, -1);
        } else {
          out.unshift(t.v === '0' ? '0' : t.v === '?' ? ' ' : '');
        }
        if (k === firstPh && rest.length > 0) {
          out.unshift(rest);
          rest = '';
        }
      } else if (t.t === 'lit') {
        out.unshift(t.v);
      }
    }
    const s = out.join('');
    return grouping && !hasLiteral ? s.replace(/\d+/, groupThousands) : s;
  };

  /**
   * 小数部の数字を入れる（末尾の 0 は # なら消し、? なら空白にする）
   * @param {string} digits
   * @param {Token[]} tokens
   */
  const fillFraction = (digits, tokens) => {
    const phs = tokens.filter(isPh);
    const chars = [...digits];
    for (let i = phs.length - 1; i >= 0 && chars[i] === '0'; i -= 1) {
      if (phs[i].v === '#') chars[i] = '';
      else if (phs[i].v === '?') chars[i] = ' ';
      else break;
    }
    let n = 0;
    return tokens
      .map((t) => {
        if (t.t === 'ph') {
          n += 1;
          return chars[n - 1] ?? '';
        }
        return t.t === 'lit' ? t.v : '';
      })
      .join('');
  };

  /**
   * @param {Token[]} tokens
   * @param {number} value
   */
  const plainText = (tokens, value) =>
    tokens
      .map((t) => {
        if (t.t === 'lit') return t.v;
        if (t.t === 'pct') return '%';
        if (t.t === 'general') return formatGeneral(value);
        return '';
      })
      .join('');

  /**
   * 分数の表示（# ?/? など）。近い分数で表す
   * @param {number} v 0 以上
   * @param {Token[]} tokens
   */
  const formatFraction = (v, tokens) => {
    const slash = tokens.findIndex((t) => t.t === 'slash');
    const denTokens = tokens.slice(slash + 1);
    const fixed = denTokens.filter((t) => t.t === 'lit' && /^\d+$/.test(t.v)).map((t) => (t.t === 'lit' ? t.v : '')).join('');
    const denDigits = Math.max(1, denTokens.filter(isPh).length);
    const hasWhole = tokens.slice(0, slash).some((t) => t.t === 'lit' && t.v.trim() === '') || tokens.slice(0, slash).filter(isPh).length > denDigits;
    const whole = hasWhole ? Math.floor(v) : 0;
    const frac = v - whole;
    let best = { num: 0, den: 1, err: frac };
    if (fixed) {
      const den = Number(fixed);
      const num = Math.round(frac * den);
      best = { num, den, err: 0 };
    } else {
      const maxDen = 10 ** denDigits - 1;
      for (let den = 1; den <= maxDen; den += 1) {
        const num = Math.round(frac * den);
        const err = Math.abs(frac - num / den);
        if (err < best.err - 1e-12) best = { num, den, err };
      }
    }
    const fracText = best.num === 0 ? '' : `${best.num}/${best.den}`;
    const wholeText = hasWhole && (whole !== 0 || fracText === '') ? String(whole) : '';
    return [wholeText, fracText].filter((x) => x !== '').join(' ') || '0';
  };

  /**
   * 数値用の区画で数値を文字にする（value は区画の選び方に応じて絶対値で渡す）
   * @param {number} value
   * @param {Token[]} tokens
   */
  const formatNumberSection = (value, tokens) => {
    let v = value * 100 ** tokens.filter((t) => t.t === 'pct').length;
    const firstPh = tokens.findIndex(isPh);
    if (firstPh < 0) return plainText(tokens, v);
    let lastPh = -1;
    tokens.forEach((t, k) => {
      if (isPh(t)) lastPh = k;
    });
    const prefix = plainText(tokens.slice(0, firstPh), v);
    // 最後の数字の位置のすぐ後ろの「,」は、1,000 で割る指定
    let k = lastPh + 1;
    while (k < tokens.length && tokens[k].t === 'comma') {
      v /= 1000;
      k += 1;
    }
    const suffix = plainText(tokens.slice(k), v);
    const core = tokens.slice(firstPh, lastPh + 1);

    if (core.some((t) => t.t === 'slash')) return `${prefix}${formatFraction(v, core)}${suffix}`;

    const expAt = core.findIndex((t) => t.t === 'exp');
    const mantTokens = expAt >= 0 ? core.slice(0, expAt) : core;
    const dotAt = mantTokens.findIndex((t) => t.t === 'dot');
    let intTokens = dotAt >= 0 ? mantTokens.slice(0, dotAt) : mantTokens;
    const fracTokens = dotAt >= 0 ? mantTokens.slice(dotAt + 1) : [];
    // 整数部の末尾の「,」（例: 0,.00）も 1,000 で割る指定
    while (intTokens.length > 0 && intTokens[intTokens.length - 1].t === 'comma') {
      v /= 1000;
      intTokens = intTokens.slice(0, -1);
    }
    const grouping = intTokens.some((t) => t.t === 'comma');
    const fracCount = fracTokens.filter(isPh).length;

    let expText = '';
    if (expAt >= 0) {
      const expToken = core[expAt];
      const expDigits = core.slice(expAt + 1).filter(isPh).length;
      const intCount = Math.max(1, intTokens.filter(isPh).length);
      let exp = v === 0 ? 0 : Math.floor(Math.log10(v));
      if (intCount > 1) exp = Math.floor(exp / intCount) * intCount;
      let mant = v === 0 ? 0 : v / 10 ** exp;
      const [mi] = roundDigits(mant, fracCount);
      if (intCount === 1 && mi.length > 1) {
        exp += 1;
        mant = v / 10 ** exp;
      }
      v = mant;
      const sign = exp < 0 ? '-' : expToken.t === 'exp' && expToken.v === '+' ? '+' : '';
      expText = `E${sign}${String(Math.abs(exp)).padStart(expDigits, '0')}`;
    }

    const [ip, fp] = roundDigits(v, fracCount);
    const intText = fillInteger(ip === '0' ? '' : ip, intTokens, grouping);
    const fracText = dotAt >= 0 ? `.${fillFraction(fp, fracTokens)}` : '';
    return `${prefix}${intText}${fracText}${expText}${suffix}`;
  };

  // ------------------------------------------------------------------
  // 日付・時刻
  // ------------------------------------------------------------------

  const ERAS = Object.freeze([
    { name: '令和', short: '令', abbr: 'R', y: 2019, start: 20190501 },
    { name: '平成', short: '平', abbr: 'H', y: 1989, start: 19890108 },
    { name: '昭和', short: '昭', abbr: 'S', y: 1926, start: 19261225 },
    { name: '大正', short: '大', abbr: 'T', y: 1912, start: 19120730 },
    { name: '明治', short: '明', abbr: 'M', y: 1868, start: 18680908 },
  ]);
  const WEEKDAYS_JA = Object.freeze(['日', '月', '火', '水', '木', '金', '土']);
  const WEEKDAYS_EN = Object.freeze(['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday']);
  const MONTHS_EN = Object.freeze([
    'January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December',
  ]);

  const DAY_MS = 86400000;

  /**
   * Excel の日付の数値（シリアル値）を年月日・時刻にする
   * @param {number} serial
   * @param {boolean} date1904
   * @param {boolean} keepMs  秒の小数を表示するときは、秒に丸めない
   */
  const serialToParts = (serial, date1904, keepMs) => {
    let days = Math.floor(serial);
    let ms = keepMs ? Math.round((serial - days) * DAY_MS) : Math.round((serial - days) * 86400) * 1000;
    if (ms >= DAY_MS) {
      days += 1;
      ms -= DAY_MS;
    }
    let y = 1900;
    let m = 2;
    let d = 29;
    let wd = 3;
    // 1900 年 2 月 29 日（実在しない日。Excel の互換のための日）
    if (date1904 || days !== 60) {
      const base = date1904 ? Date.UTC(1904, 0, 1) : days < 60 ? Date.UTC(1899, 11, 31) : Date.UTC(1899, 11, 30);
      const dt = new Date(base + days * DAY_MS);
      y = dt.getUTCFullYear();
      m = dt.getUTCMonth() + 1;
      d = dt.getUTCDate();
      wd = dt.getUTCDay();
    }
    return {
      y,
      m,
      d,
      wd,
      H: Math.floor(ms / 3600000),
      M: Math.floor((ms % 3600000) / 60000),
      S: Math.floor((ms % 60000) / 1000),
      ms: ms % 1000,
      totalDays: days + ms / DAY_MS,
    };
  };

  /**
   * @param {Token[]} tokens
   * @param {number} k
   * @param {-1|1} dir
   * @returns {{ t: string, v: string }|null}
   */
  const neighborTimeToken = (tokens, k, dir) => {
    for (let i = k + dir; i >= 0 && i < tokens.length; i += dir) {
      const t = tokens[i];
      if (t.t === 'date' || t.t === 'elapsed') return { t: t.t, v: t.v };
    }
    return null;
  };

  /**
   * 日付・時刻用の区画で数値を文字にする
   * @param {number} value
   * @param {Section} section
   * @param {boolean} date1904
   */
  const formatDateSection = (value, section, date1904) => {
    const { tokens } = section;
    if (value < 0) return formatGeneral(value);
    const subSecondDigits = (() => {
      const sIdx = tokens.findIndex((t) => t.t === 'date' && t.v.startsWith('s'));
      if (sIdx < 0 || tokens[sIdx + 1]?.t !== 'dot') return 0;
      let n = 0;
      for (let i = sIdx + 2; i < tokens.length && tokens[i].t === 'ph' && /** @type {{ v: string }} */ (tokens[i]).v === '0'; i += 1) n += 1;
      return n;
    })();
    const p = serialToParts(value, date1904, subSecondDigits > 0);
    const date = p.y * 10000 + p.m * 100 + p.d;
    const era = ERAS.find((e) => date >= e.start) ?? null;
    const eraYear = era ? p.y - era.y + 1 : p.y;
    const ampm = tokens.some((t) => t.t === 'ampm');
    const hour12 = p.H % 12 === 0 ? 12 : p.H % 12;
    const pad = (/** @type {number} */ n, /** @type {number} */ w) => String(n).padStart(w, '0');

    let afterSeconds = false;
    let subUsed = 0;
    return tokens
      .map((t, k) => {
        if (t.t === 'lit') return t.v;
        if (t.t === 'elapsed') {
          if (t.v === 'h') return String(Math.floor(p.totalDays * 24));
          if (t.v === 'm') return String(Math.floor(p.totalDays * 1440));
          return String(Math.floor(p.totalDays * 86400));
        }
        if (t.t === 'ampm') {
          const pm = p.H >= 12;
          return t.v === 'A/P' ? (pm ? 'P' : 'A') : pm ? 'PM' : 'AM';
        }
        if (t.t === 'dot') {
          afterSeconds = subSecondDigits > 0;
          return '.';
        }
        if (t.t === 'ph') {
          if (afterSeconds && subUsed < subSecondDigits) {
            subUsed += 1;
            return pad(p.ms, 3)[subUsed - 1];
          }
          return t.v === '0' ? '0' : '';
        }
        if (t.t !== 'date') return t.t === 'comma' ? ',' : t.t === 'slash' ? '/' : t.t === 'pct' ? '%' : '';
        const v = t.v;
        switch (v[0]) {
          case 'y':
            return v.length <= 2 ? pad(p.y % 100, 2) : String(p.y);
          case 'e':
            if (section.gannen && eraYear === 1 && era) return '元';
            return v.length >= 2 ? pad(eraYear, 2) : String(eraYear);
          case 'g':
            if (!era) return '';
            return v.length === 1 ? era.abbr : v.length === 2 ? era.short : era.name;
          case 'd':
            if (v.length === 1) return String(p.d);
            if (v.length === 2) return pad(p.d, 2);
            return v.length === 3 ? WEEKDAYS_EN[p.wd].slice(0, 3) : WEEKDAYS_EN[p.wd];
          case 'a':
            return v.length === 3 ? WEEKDAYS_JA[p.wd] : `${WEEKDAYS_JA[p.wd]}曜日`;
          case 'h':
            return v.length >= 2 ? pad(ampm ? hour12 : p.H, 2) : String(ampm ? hour12 : p.H);
          case 's':
            return v.length >= 2 ? pad(p.S, 2) : String(p.S);
          case 'm': {
            const prev = neighborTimeToken(tokens, k, -1);
            const next = neighborTimeToken(tokens, k, 1);
            const isMinute =
              v.length <= 2 &&
              ((prev !== null && (prev.v.startsWith('h') || (prev.t === 'elapsed' && prev.v === 'h'))) ||
                (next !== null && next.v.startsWith('s')));
            if (isMinute) return v.length === 2 ? pad(p.M, 2) : String(p.M);
            if (v.length === 1) return String(p.m);
            if (v.length === 2) return pad(p.m, 2);
            if (v.length === 3) return MONTHS_EN[p.m - 1].slice(0, 3);
            return v.length === 5 ? MONTHS_EN[p.m - 1][0] : MONTHS_EN[p.m - 1];
          }
          default:
            return '';
        }
      })
      .join('');
  };

  // ------------------------------------------------------------------
  // 入口
  // ------------------------------------------------------------------

  /** @type {Map<string, Section[]>} */
  const cache = new Map();

  /** @param {string} code */
  const sectionsOf = (code) => {
    const hit = cache.get(code);
    if (hit) return hit;
    const parsed = splitSections(code).map(tokenize);
    if (cache.size > 500) cache.clear();
    cache.set(code, parsed);
    return parsed;
  };

  /**
   * @param {{ op: string, value: number }} cond
   * @param {number} v
   */
  const testCond = ({ op, value }, v) =>
    op === '<' ? v < value : op === '<=' ? v <= value : op === '>' ? v > value : op === '>=' ? v >= value : op === '=' ? v === value : v !== value;

  /**
   * 数値を表示形式で文字にする
   * @param {number} value
   * @param {string} code  表示形式（例: '#,##0'、'yyyy"年"m"月"d"日"'）
   * @param {{ date1904?: boolean }} [opts]
   * @returns {string}
   */
  const formatNumber = (value, code, opts = {}) => {
    if (!Number.isFinite(value)) return String(value);
    const all = sectionsOf(code === '' ? 'General' : code);
    const numeric = all.filter((s) => !s.tokens.some((t) => t.t === 'at')).slice(0, 3);
    if (numeric.length === 0) return formatGeneral(value);

    /** @type {Section} */
    let section = numeric[0];
    let v = value;
    let sign = '';
    if (numeric.some((s) => s.cond !== null)) {
      const matched = numeric.find((s) => s.cond !== null && testCond(s.cond, value));
      section = matched ?? numeric.find((s) => s.cond === null) ?? numeric[numeric.length - 1];
      // 条件つきの区画では、負の数の符号は区画の側で書く（書いていなければ付ける）
      if (value < 0 && !section.tokens.some((t) => t.t === 'lit' && /[-−△▲(（]/.test(t.v))) sign = '-';
      v = Math.abs(value);
    } else if (value < 0) {
      if (numeric.length >= 2) {
        section = numeric[1];
      } else {
        sign = '-';
      }
      v = -value;
    } else if (value === 0 && numeric.length >= 3) {
      section = numeric[2];
    }
    const body = isDateSection(section.tokens)
      ? formatDateSection(v, section, opts.date1904 === true)
      : formatNumberSection(v, section.tokens);
    // 丸めた結果がすべて 0 なら、符号は付けない（例: -0.001 を "0" で表示）。
    // 前後の空白は、列をそろえるための余白（例: '#,##0 ;(#,##0)' の後ろの空白）なので取り除く
    return (sign && /[1-9]/.test(body) ? `${sign}${body}` : body).trim();
  };

  /**
   * 文字のセルを表示形式で文字にする（表示形式に @ の区画があるときだけ形が変わる）
   * @param {string} text
   * @param {string} code
   */
  const formatText = (text, code) => {
    const all = sectionsOf(code === '' ? 'General' : code);
    const textSection = all.length >= 4 ? all[3] : all.find((s) => s.tokens.some((t) => t.t === 'at'));
    if (!textSection) return text;
    return textSection.tokens.map((t) => (t.t === 'at' ? text : t.t === 'lit' ? t.v : '')).join('');
  };

  /**
   * 日付・時刻の表示形式か（最初の区画で判断する）
   * @param {string} code
   */
  const isDateFormat = (code) => isDateSection(sectionsOf(code === '' ? 'General' : code)[0].tokens);

  /**
   * 表示形式の番号から表示形式を引く（ファイルに書かれたもの → 組み込み → 標準）
   * @param {number} id
   * @param {ReadonlyMap<number, string>} custom
   */
  const formatCodeOf = (id, custom) => custom.get(id) ?? BUILTIN_FORMATS[id] ?? 'General';

  root.YegNumFmt = Object.freeze({
    BUILTIN_FORMATS,
    formatNumber,
    formatText,
    formatGeneral,
    isDateFormat,
    formatCodeOf,
  });
})(globalThis);
