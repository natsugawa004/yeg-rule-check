// @ts-check
/*
 * Word・Excel のファイルの中にある XML を読む（画面に依存しない純粋な関数だけを置く）
 *   parseXml(文字列) → 要素の木。名前空間（w: や a: などの接頭辞が指す URI）を解決して持つ
 * ブラウザの DOMParser を使わないのは、Node のテストでも同じ処理を確かめるため。
 * DOCTYPE（外部の定義を読み込む指定）を含む XML は、Office のファイルには現れないため読み込まない。
 */
'use strict';

(function (root) {
  /**
   * @typedef {Object} XmlElement
   * @property {string} name   書かれたとおりの名前（例: w:p）
   * @property {string} local  接頭辞を除いた名前（例: p）
   * @property {string} ns     名前空間の URI（なければ空文字）
   * @property {Readonly<Record<string, string>>} attrs  属性（書かれたとおりの名前で引く）
   * @property {ReadonlyMap<string, string>} scope  この要素で有効な接頭辞→URI
   * @property {XmlNode[]} children
   */
  /** @typedef {XmlElement|string} XmlNode */

  class XmlError extends Error {
    /** @param {string} message */
    constructor(message) {
      super(message);
      this.name = 'XmlError';
    }
  }

  const XML_NS = 'http://www.w3.org/XML/1998/namespace';

  /** @type {Readonly<Record<string, string>>} */
  const NAMED_ENTITIES = Object.freeze({ lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" });

  /**
   * 文字参照・定義済みの実体参照を文字に戻す
   * @param {string} s
   */
  const decodeEntities = (s) =>
    s.indexOf('&') < 0
      ? s
      : s.replace(/&(#x[0-9A-Fa-f]+|#[0-9]+|[A-Za-z]+);/g, (whole, body) => {
          if (body[0] === '#') {
            const code = body[1] === 'x' || body[1] === 'X' ? parseInt(body.slice(2), 16) : parseInt(body.slice(1), 10);
            return Number.isFinite(code) && code >= 0 && code <= 0x10ffff ? String.fromCodePoint(code) : whole;
          }
          return NAMED_ENTITIES[body] ?? whole;
        });

  /**
   * XML の改行の決まり（CR LF と CR を LF にする）
   * @param {string} s
   */
  const normalizeEol = (s) => (s.indexOf('\r') < 0 ? s : s.replace(/\r\n?/g, '\n'));

  const NAME_END = /[\s/>]/;

  /**
   * 開始タグの中（< と > の間）を読む
   * @param {string} body  例: 'w:p w:rsidR="00A1"' または 'w:br/'
   * @returns {{ name: string, attrs: Record<string, string>, selfClosing: boolean }}
   */
  const parseTag = (body) => {
    const selfClosing = body.endsWith('/');
    const inner = selfClosing ? body.slice(0, -1) : body;
    const m = NAME_END.exec(inner);
    const name = m ? inner.slice(0, m.index) : inner;
    if (name === '') throw new XmlError('要素の名前がありません');
    /** @type {Record<string, string>} */
    const attrs = {};
    const re = /([^\s=]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
    const rest = m ? inner.slice(m.index) : '';
    let a;
    while ((a = re.exec(rest)) !== null) {
      attrs[a[1]] = decodeEntities(a[2] ?? a[3] ?? '');
    }
    return { name, attrs, selfClosing };
  };

  /**
   * @param {ReadonlyMap<string, string>} parent
   * @param {Record<string, string>} attrs
   * @returns {ReadonlyMap<string, string>}
   */
  const extendScope = (parent, attrs) => {
    const decls = Object.keys(attrs).filter((k) => k === 'xmlns' || k.startsWith('xmlns:'));
    if (decls.length === 0) return parent;
    const next = new Map(parent);
    for (const k of decls) next.set(k === 'xmlns' ? '' : k.slice(6), attrs[k]);
    return next;
  };

  /** @param {string} name */
  const splitName = (name) => {
    const i = name.indexOf(':');
    return i < 0 ? { prefix: '', local: name } : { prefix: name.slice(0, i), local: name.slice(i + 1) };
  };

  /** @type {ReadonlyMap<string, string>} */
  const ROOT_SCOPE = new Map([['xml', XML_NS]]);

  /**
   * XML を読み、いちばん外側の要素を返す
   * @param {string} source
   * @returns {XmlElement}
   */
  const parseXml = (source) => {
    const xml = source.charCodeAt(0) === 0xfeff ? source.slice(1) : source;
    /** @type {XmlElement[]} */
    const stack = [];
    /** @type {XmlElement|null} */
    let top = null;
    let i = 0;

    /** @param {string} text */
    const pushText = (text) => {
      if (text === '' || stack.length === 0) return;
      stack[stack.length - 1].children.push(text);
    };

    while (i < xml.length) {
      const lt = xml.indexOf('<', i);
      if (lt < 0) {
        if (xml.slice(i).trim() !== '' && stack.length > 0) throw new XmlError('閉じていない要素があります');
        break;
      }
      if (lt > i) pushText(decodeEntities(normalizeEol(xml.slice(i, lt))));
      if (xml.startsWith('<?', lt)) {
        const end = xml.indexOf('?>', lt + 2);
        if (end < 0) throw new XmlError('処理命令が閉じていません');
        i = end + 2;
      } else if (xml.startsWith('<!--', lt)) {
        const end = xml.indexOf('-->', lt + 4);
        if (end < 0) throw new XmlError('コメントが閉じていません');
        i = end + 3;
      } else if (xml.startsWith('<![CDATA[', lt)) {
        const end = xml.indexOf(']]>', lt + 9);
        if (end < 0) throw new XmlError('CDATA が閉じていません');
        pushText(normalizeEol(xml.slice(lt + 9, end)));
        i = end + 3;
      } else if (xml.startsWith('<!', lt)) {
        throw new XmlError('DOCTYPE などの定義を含む XML は読み込めません');
      } else if (xml[lt + 1] === '/') {
        const end = xml.indexOf('>', lt + 2);
        if (end < 0) throw new XmlError('終了タグが閉じていません');
        const name = xml.slice(lt + 2, end).trim();
        const open = stack.pop();
        if (!open || open.name !== name) throw new XmlError(`終了タグ </${name}> の対応がとれません`);
        if (stack.length === 0) top = open;
        i = end + 1;
      } else {
        // 属性の値の中の > で切らないよう、引用符を数えながら閉じ括弧を探す
        let j = lt + 1;
        /** @type {string|null} */
        let quote = null;
        for (; j < xml.length; j += 1) {
          const c = xml[j];
          if (quote) {
            if (c === quote) quote = null;
          } else if (c === '"' || c === "'") {
            quote = c;
          } else if (c === '>') {
            break;
          }
        }
        if (j >= xml.length) throw new XmlError('開始タグが閉じていません');
        const { name, attrs, selfClosing } = parseTag(xml.slice(lt + 1, j).trim());
        const parentScope = stack.length > 0 ? stack[stack.length - 1].scope : ROOT_SCOPE;
        const scope = extendScope(parentScope, attrs);
        const { prefix, local } = splitName(name);
        /** @type {XmlElement} */
        const el = { name, local, ns: scope.get(prefix) ?? '', attrs, scope, children: [] };
        if (stack.length > 0) stack[stack.length - 1].children.push(el);
        else if (top !== null) throw new XmlError('いちばん外側の要素が2つ以上あります');
        if (selfClosing) {
          if (stack.length === 0) top = el;
        } else {
          stack.push(el);
        }
        i = j + 1;
      }
    }
    if (stack.length > 0) throw new XmlError(`要素 <${stack[stack.length - 1].name}> が閉じていません`);
    if (!top) throw new XmlError('XML の中身がありません');
    return top;
  };

  // ------------------------------------------------------------------
  // 木をたどる補助
  // ------------------------------------------------------------------

  /**
   * @param {XmlNode} n
   * @returns {n is XmlElement}
   */
  const isElement = (n) => typeof n !== 'string';

  /**
   * 名前空間の URI の集まりと名前が合うか（Office の「厳密」形式は URI が違うため、複数の URI を受ける）
   * @param {XmlElement} el
   * @param {ReadonlySet<string>|null} nsSet  null なら名前空間を問わない
   * @param {string} local
   */
  const is = (el, nsSet, local) => el.local === local && (nsSet === null || nsSet.has(el.ns));

  /**
   * 子の要素
   * @param {XmlElement} el
   * @param {ReadonlySet<string>|null} [nsSet]
   * @param {string} [local]
   * @returns {XmlElement[]}
   */
  const kids = (el, nsSet = null, local) =>
    el.children.filter(
      /** @returns {c is XmlElement} */ (c) => isElement(c) && (local === undefined || is(c, nsSet, local)),
    );

  /**
   * 最初に見つかった子の要素
   * @param {XmlElement} el
   * @param {ReadonlySet<string>|null} nsSet
   * @param {string} local
   * @returns {XmlElement|null}
   */
  const kid = (el, nsSet, local) => {
    for (const c of el.children) if (isElement(c) && is(c, nsSet, local)) return c;
    return null;
  };

  /**
   * 子孫の要素を、書かれている順にすべて集める
   * @param {XmlElement} el
   * @param {ReadonlySet<string>|null} nsSet
   * @param {string} local
   * @returns {XmlElement[]}
   */
  const descendants = (el, nsSet, local) => {
    /** @type {XmlElement[]} */
    const out = [];
    /** @param {XmlElement} e */
    const walk = (e) => {
      for (const c of e.children) {
        if (!isElement(c)) continue;
        if (is(c, nsSet, local)) out.push(c);
        walk(c);
      }
    };
    walk(el);
    return out;
  };

  /**
   * 属性の値（接頭辞の付かない属性。なければ同じ名前の接頭辞付き属性）
   * @param {XmlElement} el
   * @param {string} local
   * @returns {string|null}
   */
  const attr = (el, local) => {
    if (Object.prototype.hasOwnProperty.call(el.attrs, local)) return el.attrs[local];
    for (const [k, v] of Object.entries(el.attrs)) {
      const i = k.indexOf(':');
      if (i >= 0 && k.slice(i + 1) === local && !k.startsWith('xmlns')) return v;
    }
    return null;
  };

  /**
   * 名前空間を指定して属性を読む（例: r:id）
   * @param {XmlElement} el
   * @param {ReadonlySet<string>} nsSet
   * @param {string} local
   * @returns {string|null}
   */
  const attrNs = (el, nsSet, local) => {
    for (const [k, v] of Object.entries(el.attrs)) {
      const { prefix, local: l } = splitName(k);
      if (prefix !== '' && l === local && nsSet.has(el.scope.get(prefix) ?? '')) return v;
    }
    return null;
  };

  /**
   * 要素の中の文字（子孫の文字をつなげたもの）
   * @param {XmlElement} el
   * @returns {string}
   */
  const textOf = (el) => el.children.map((c) => (isElement(c) ? textOf(c) : c)).join('');

  root.YegXml = Object.freeze({
    XmlError,
    parseXml,
    decodeEntities,
    isElement,
    is,
    kids,
    kid,
    descendants,
    attr,
    attrNs,
    textOf,
  });
})(globalThis);
