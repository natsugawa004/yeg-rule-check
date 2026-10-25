/*
 * YEG 記載ルールチェッカー 委員会事業計画書のルール定義（annual.html で使います）
 * ------------------------------------------------------------------
 * 委員会ごとの年間の事業計画書（令和N年度事業計画提出書。手帳に載る対外向けの文章）を判定するルールです。
 * 出典は組織図スプレッドシートの「虎の巻（委員会事業計画版）」です。
 *
 * rules.js（議案のルール）を先に読み込み、その文章のルール（R-02〜R-21 など）を「要確認」に下げて一緒に使います。
 * 議案のルールを直すと、このページにも自動で反映されます。
 *
 * ■ 用語を足す
 *   下の annualRules の Y-01〜Y-05 の patterns に1行足します（書き方は rules.js の先頭の説明と同じです）。
 * ■ 議案のルールのうち、このページでは当てないもの
 *   skipRuleIds に番号を書きます（「当青年部」の R-01 と「と思う・と考える」の R-13）。
 * ■ このページだけの項目
 *   documentTypes: ['annual']      … 委員会事業計画書だけに当てます（このページではいつも委員会事業計画書として判定します）
 *   dynamicWords: 'committeeName' … 見出しの「〇〇委員会　委員長名」から読み取った委員会名を、見つける語にします（Y-03）
 *   type: 'document' の check と params
 *     'length'  … 上段・中段・下段の字数の目安（段落が3つ以上のときだけ）。low 倍未満か high 倍を超えたら知らせます
 *     'total'   … 本文の合計の字数。max を超えたら知らせます
 *     'opening' … 中段の書き出し
 *     'closing' … 締めの言葉
 *   必須項目・空欄のチェックはしません（本文だけを貼り付けて確かめることもあるため）。
 * ------------------------------------------------------------------
 */
'use strict';

(function (g) {
  /** 議案のルール（rules.js）。読み込めていなければ、このページのルールも作らない（画面に読み込めなかったことを出す） */
  const base = g.YEG_RULESET;
  if (!base || typeof base !== 'object' || !Array.isArray(base.rules)) {
    g.YEG_RULESET = undefined;
    return;
  }

  /** 議案のルールのうち、このページでは当てないもの */
  const skipRuleIds = ['R-01', 'R-13'];

  /** 議案のルールの「違反」は、このページでは「要確認」にする */
  const lower = (/** @type {unknown} */ s) => (s === '違反' ? '要確認' : s);

  // 議案の文章のルール（書類の種類を決めていないもの）だけを使う。書類全体・金額のルール（D-05〜D-08・M-xx）は使わない
  const gianRules = base.rules
    .filter(
      (r) =>
        r &&
        typeof r === 'object' &&
        (r.type === undefined || r.type === 'pattern' || r.type === 'style') &&
        !(Array.isArray(r.documentTypes) && r.documentTypes.length > 0) &&
        !skipRuleIds.includes(r.id),
    )
    .map((r) => ({
      ...r,
      severity: lower(r.severity),
      ...(r.sectionSeverity && typeof r.sectionSeverity === 'object'
        ? { sectionSeverity: Object.fromEntries(Object.entries(r.sectionSeverity).map(([k, v]) => [k, lower(v)])) }
        : {}),
    }));
  const gianCategories = (Array.isArray(base.categories) ? base.categories : []).filter((c) => gianRules.some((r) => r.category === c));
  const gianNote = '議案（理事会上程資料の書き方）のルールです。委員会事業計画書では「違反」のものも「要確認」にしています。';

  const SECTION = '委員会活動の目的・抱負';
  const SOURCE_WORDS = '【虎の巻（委員会事業計画版）】用語の統一リスト';
  const SOURCE_LAYOUT = '【虎の巻（委員会事業計画版）】4〜6 上段・中段・下段';

  const annualRules = [
    {
      id: 'Y-01',
      name: '組織の呼び方',
      category: '委員会事業計画書の用語',
      severity: '違反',
      documentTypes: ['annual'],
      patterns: [
        { words: ['小田原箱根商工会議所青年部'], fix: ['YEG'] },
        // 「当青年部会員」は Y-02
        { re: String.raw`当青年部(?!会員)`, fix: ['当YEG'] },
        // 他団体の名前の一部（神奈川県商工会議所青年部連合会・日本商工会議所青年部など）と「青年部会員」は指摘しない
        { re: String.raw`(?<!商工会議所|当)青年部(?!会員|連合会)`, fix: ['YEG'] },
        // 「当会」「本会」「我が会」（当会議・本会場・当会員などは指摘しない。「当会員」は Y-05）
        { re: String.raw`(?<!\p{sc=Han})(?:当|本|我が)会(?![議場館計社員長期合話食])`, fix: ['当YEG'] },
      ],
      examples: ['小田原箱根商工会議所青年部', '当青年部', '青年部', '当会'],
      fixes: ['YEG', '当YEG'],
      description:
        '委員会事業計画書では、自分たちの組織は「YEG」（「当YEG」も可）と書きます。「小田原箱根商工会議所青年部」「青年部」「当青年部」「会」とは書きません（「神奈川県商工会議所青年部連合会」のような他団体の名前は除く）。',
      source: SOURCE_WORDS,
      scope: '委員会事業計画書',
      enabled: true,
    },
    {
      id: 'Y-02',
      name: '会員の呼び方',
      category: '委員会事業計画書の用語',
      severity: '違反',
      documentTypes: ['annual'],
      patterns: [{ words: ['当青年部会員', '青年部会員', 'メンバー'], fix: ['会員'] }],
      examples: ['メンバー', '青年部会員', '当青年部会員'],
      fixes: ['会員'],
      description: '自分たちの組織の会員は「会員」と書きます。「メンバー」「青年部会員」「当青年部会員」とは書きません。',
      source: SOURCE_WORDS,
      scope: '委員会事業計画書',
      enabled: true,
    },
    {
      id: 'Y-03',
      name: '自分の委員会の呼び方',
      category: '委員会事業計画書の用語',
      severity: '違反',
      documentTypes: ['annual'],
      // 見出しの「〇〇委員会　委員長名　氏名」から読み取った委員会名（読み仮名のかっこを外した形も）を探す
      dynamicWords: 'committeeName',
      patterns: [],
      examples: ['（見出しに書いた自分の委員会名）'],
      fixes: ['当委員会'],
      description:
        '本文で自分の委員会を指すときは、委員会名ではなく「当委員会」と書きます（委員会名は見出しの「〇〇委員会　委員長名」の行から読み取っています。本文だけを貼り付けたときは確かめません）。',
      source: SOURCE_WORDS,
      scope: '委員会事業計画書（見出しの行があるとき）',
      enabled: true,
    },
    {
      id: 'Y-04',
      name: '地域の呼び方',
      category: '委員会事業計画書の用語',
      severity: '違反',
      documentTypes: ['annual'],
      patterns: [
        { words: ['小田原箱根地域', '小田原・箱根地域', '小田原･箱根地域'], fix: ['当地域'] },
        // 後ろに漢字・カタカナが続く「小田原箱根〇〇」（事業名など）は Y-05 で要確認。「小田原箱根商工会議所」は指摘しない
        { re: String.raw`小田原[・･]?箱根(?![・･]?(?:商工会議所|地域))(?![\p{sc=Han}\p{sc=Katakana}ー])`, fix: ['当地域'] },
      ],
      examples: ['小田原箱根地域', '小田原箱根', '小田原・箱根'],
      fixes: ['当地域'],
      description: '小田原箱根を指すときは「当地域」と書きます。「小田原箱根地域」「小田原箱根」とは書きません。',
      source: SOURCE_WORDS,
      scope: '委員会事業計画書',
      enabled: true,
    },
    {
      id: 'Y-05',
      name: '確認したい呼び方',
      category: '委員会事業計画書の用語',
      severity: '要確認',
      documentTypes: ['annual'],
      patterns: [
        { words: ['当会員'], fix: ['会員'] },
        { words: ['本委員会'], fix: ['当委員会'] },
        { re: String.raw`小田原[・･]?箱根(?![・･]?(?:商工会議所|地域))(?=[\p{sc=Han}\p{sc=Katakana}ー])`, fix: ['当地域'] },
      ],
      examples: ['当会員', '本委員会', '小田原箱根大博覧会'],
      fixes: ['会員', '当委員会', '当地域'],
      description:
        '「当会員」は基本的に使わず「会員」と書きます。「本委員会」は「当委員会」と書きます。「小田原箱根〇〇」は、地域を指すなら「当地域」と書きます（「小田原箱根大博覧会」のような事業名などの固有名詞なら問題ありません）。',
      source: SOURCE_WORDS,
      scope: '委員会事業計画書',
      enabled: true,
    },
    {
      id: 'A-01',
      name: '段ごとの字数の目安',
      category: '委員会事業計画書の構成',
      severity: '参考',
      type: 'document',
      documentTypes: ['annual'],
      check: 'length',
      // 本文の欄と、上段・中段・下段の目安の字数（空白・改行は数えない）。目安の low 倍未満か high 倍を超えたら知らせる
      params: {
        section: SECTION,
        parts: [
          { label: '上段（現状・背景・課題）', target: 200 },
          { label: '中段（課題を解決する手法）', target: 400 },
          { label: '下段（解決後の姿）', target: 150 },
        ],
        low: 0.5,
        high: 1.5,
      },
      message: '{label}は{chars}字で、目安（{target}字）より{how}です',
      examples: ['上段が80字（目安200字）', '中段が700字（目安400字）'],
      fixes: [],
      hint: '目安の字数に近づける',
      description:
        '目安は、上段（現状・背景・課題）200字、中段（課題を解決する手法）400字、下段（解決後の姿）150字です。目安の半分より少ないか、1.5倍より多いときに知らせます（段落が3つ以上のときだけ）。',
      source: SOURCE_LAYOUT,
      scope: '委員会事業計画書の「委員会活動の目的・抱負」',
      enabled: true,
    },
    {
      id: 'A-04',
      name: '合計の字数',
      category: '委員会事業計画書の構成',
      severity: '要確認',
      type: 'document',
      documentTypes: ['annual'],
      check: 'total',
      // 本文の欄の合計の字数（空白・改行は数えない）が max を超えたら知らせる（段落の数にかかわらず）
      params: { section: SECTION, max: 750 },
      message: '「委員会活動の目的・抱負」は合計{chars}字で、目安の合計（{max}字）を超えています',
      examples: ['合計が800字（目安の合計750字）'],
      fixes: [],
      hint: '合計750字以内に収める',
      description: '上段200字・中段400字・下段150字の目安を足した750字を超えたら知らせます（段落の数にかかわらず見ます）。',
      source: SOURCE_LAYOUT,
      scope: '委員会事業計画書の「委員会活動の目的・抱負」',
      enabled: true,
    },
    {
      id: 'A-02',
      name: '中段の書き出し',
      category: '委員会事業計画書の構成',
      severity: '参考',
      type: 'document',
      documentTypes: ['annual'],
      check: 'opening',
      params: { section: SECTION, opening: 'まずは' },
      message: '中段（課題を解決する手法）は「{opening}、〇〇するために、〇〇を行い、〇〇が向上します。」のような書き方が好ましいとされています',
      examples: ['中段が「まずは」で始まっていない'],
      fixes: [],
      hint: '「まずは、〇〇するために…」の形にする',
      description:
        '中段は「まずは、〇〇するために、〇〇を行い、〇〇が向上します。」のように、手法とその効果を書く形が好ましいとされています（書き出しの「まずは」だけを、段落が3つ以上のときに確かめます）。',
      source: '【虎の巻（委員会事業計画版）】5 中段',
      scope: '委員会事業計画書の「委員会活動の目的・抱負」',
      enabled: true,
    },
    {
      id: 'A-03',
      name: '締めの言葉',
      category: '委員会事業計画書の構成',
      severity: '参考',
      type: 'document',
      documentTypes: ['annual'],
      check: 'closing',
      params: { section: SECTION, closings: ['目指します。', '目指してまいります。', '目指します', '目指してまいります'] },
      message: '最後の文は「〜を目指します。」「目指してまいります。」で締めるのが通例です',
      examples: ['最後の文が「〜を目指します。」で終わっていない'],
      fixes: ['〜を目指します。', '〜を目指してまいります。'],
      hint: '「〜を目指します。」で締める',
      description:
        '下段は「当委員会・当YEGは〇〇が向上し、今まで以上に〇〇を目指します。」のように、「目指します」「目指してまいります」で締めるのが通例です。',
      source: '【虎の巻（委員会事業計画版）】6 下段',
      scope: '委員会事業計画書の「委員会活動の目的・抱負」',
      enabled: true,
    },
  ];

  const baseMeta = base.meta && typeof base.meta === 'object' ? base.meta : {};
  const baseDocs = base.documents && typeof base.documents === 'object' ? base.documents : {};
  const baseAi = base.aiGuide && typeof base.aiGuide === 'object' ? base.aiGuide : {};

  g.YEG_RULESET = {
    meta: {
      toolName: 'YEG 記載ルールチェッカー（委員会事業計画書）',
      manualName: '虎の巻（委員会事業計画版）＋理事会上程資料の書き方',
      manualVersion: `書き方は${baseMeta.manualVersion ?? '版の記載なし'}`,
      // 議案のルールの出典の前に付ける名前
      sourceLabel: baseMeta.sourceLabel,
      updated: '2026-10-21',
    },

    // 分類の並び順（虎の巻のルールを先に、議案のルールを後に並べます）
    categories: ['委員会事業計画書の用語', '委員会事業計画書の構成', ...gianCategories],

    categoryNotes: {
      委員会事業計画書の用語: '出典は「虎の巻（委員会事業計画版）」の用語の統一リストです。',
      委員会事業計画書の構成:
        '「委員会活動の目的・抱負」を、段落ごとに上段（現状・背景・課題）・中段（手法）・下段（解決後の姿）とみて確かめます。段ごとの字数と中段の書き出しは、段落が3つ以上のときだけ見ます。合計の字数は、段落の数にかかわらず見ます。',
      ...Object.fromEntries(gianCategories.map((c) => [c, gianNote])),
    },

    // 判定から除外する定型文。議案のものに加えて、事業計画提出書の「事業計画スケジュール」欄の注意書き
    excludedPhrases: [
      ...(Array.isArray(base.excludedPhrases) ? base.excludedPhrases : []),
      'ここは、まだ記載しない。最終的に全委員会の職務分掌が確定したら、記載する。事業計画スケジュールの枠は、ここまで必要。',
    ],

    // ルール一覧の画面に出す「このツールで判定しないもの」
    notChecked: [
      '委員会名が YEG 以外の方にもわかりやすいか、会長所信・運営方針と結び付いているかなど、人の判断が必要なもの（確認ポイントで確かめてください）',
      '中段の「〇〇するために、〇〇を行い、〇〇が向上します」の形（書き出しの「まずは」だけを見ます）',
      '下段の主語（「当委員会・当YEGは」）',
      '欄が記入されているか（必須項目・空欄）。本文だけを貼り付けて確かめることもあるため、見ていません',
      '本文だけを貼り付けたときの、自分の委員会名の書き方（委員会名を見出しの「〇〇委員会　委員長名」の行から読み取るため）',
      '事業計画スケジュールの中身（提出の時点では書かないため）',
      'フォントや文字の大きさ、枠の大きさなどの書式',
    ],

    aiGuide: {
      caution: baseAi.caution,
      request:
        'あなたは文章の校正担当です。このあと貼り付ける文章は、小田原箱根商工会議所青年部（YEG）の委員会ごとの年間の事業計画書で、手帳に載る対外向けの文章です。次のルールに合わない箇所を探し、「該当箇所」「ルール」「直し方」の3列の表にしてください。ルールにないことは指摘しないでください。個人名などは伏せ字にしてあります。',
      links: baseAi.links,
    },

    documents: {
      // 欄の区切りに使う語（見出しの別名・表の列名など）は議案と同じ
      ...baseDocs,
      types: [
        {
          // 委員会ごとの年間の事業計画書（令和N年度事業計画提出書）。このページでは書類の種類を判別せず、いつもこの種類として判定する
          id: 'annual',
          name: '委員会事業計画書',
          // 見出しの語（欄の区切りに使う）
          detectWords: ['事業計画提出書', '委員会活動の目的・抱負', '事業計画スケジュール'],
          // 必須項目・空欄は確かめない（本文だけを貼り付けて確かめることもあるため）
          required: [],
          blankCheck: [],
          // 判定する欄。見出しの欄が見つかったときは、見出し（「小田原箱根商工会議所青年部」「令和N年度事業計画提出書」など）を判定しない
          scopeSections: ['委員会活動の目的・抱負', '事業計画スケジュール'],
          // 見出しの「〇〇委員会　委員長名　氏名」の行から委員会名を読み取る（1番目のかっこが委員会名）
          committeeNamePattern: String.raw`^[\s　]*(\S.*?委員会.*?)[\s　]*委員長名`,
          // 共通の確認ポイント（別紙番号・動員目標の書き方など）は議案向けなので出さない
          commonCheckpoints: false,
          checkpoints: [
            { id: 'K-61', text: '会長所信・運営方針を紐解き、当委員会がなぜ設置されたか・何をするべきかと結び付いているか' },
            { id: 'K-62', text: '委員会名は YEG 以外の方にもわかりやすいか' },
            { id: 'K-63', text: '上段の現状・背景・課題は、現状の調査・分析に基づいているか' },
            { id: 'K-64', text: '手帳に掲載される対外向けの文章として読めるか（内輪の言葉・略語がないか）' },
          ],
        },
      ],
      commonCheckpoints: [],
    },

    rules: [...annualRules, ...gianRules],
  };
})(globalThis);
