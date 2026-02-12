/**
 * NameMatcher.gs - 名寄せエンジン
 * シフト表の名前表記と従業員マスタを照合
 *
 * マッチング優先順位:
 * 1. 完全一致: シフト名 == マスタ氏名
 * 2. スペース正規化一致: normalize(シフト名) == normalize(マスタ氏名)
 * 3. 別名一致: シフト名 ∈ マスタ.別名リスト
 * 4. 姓一致(一意): シフト名 == マスタ.姓 かつ 同姓1名のみ
 * 5. 姓+名先頭一致: "佐藤佳" → "佐藤 佳子"
 * 6. 旧字体/異体字変換一致
 */

/**
 * 名寄せエンジンクラス
 */
class NameMatcherEngine {
  /**
   * @param {Array<Object>} employees - 従業員マスタ配列
   */
  constructor(employees) {
    this.employees = employees;
    this.buildIndices();
  }

  /**
   * 検索用インデックスを構築
   */
  buildIndices() {
    // 正規化名→従業員マッピング
    this.normalizedNameMap = new Map();
    // 姓→従業員リスト
    this.surnameMap = new Map();
    // 別名→従業員
    this.aliasMap = new Map();
    // 正規化別名→従業員
    this.normalizedAliasMap = new Map();

    for (const emp of this.employees) {
      if (emp.status === '退職') continue;

      // 正規化名でインデックス
      const normalizedName = this.normalizeName(emp.name);
      this.normalizedNameMap.set(normalizedName, emp);

      // 元の名前もインデックス
      this.normalizedNameMap.set(emp.name, emp);

      // 姓でインデックス
      const surname = this.extractSurname(emp.name);
      if (surname) {
        if (!this.surnameMap.has(surname)) {
          this.surnameMap.set(surname, []);
        }
        this.surnameMap.get(surname).push(emp);
      }

      // 別名でインデックス
      for (const alias of emp.aliases) {
        this.aliasMap.set(alias, emp);
        this.normalizedAliasMap.set(this.normalizeName(alias), emp);
      }
    }
  }

  /**
   * 名前をマッチング
   * @param {string} shiftName - シフト表の名前
   * @param {string} facility - 施設名（絞り込み用）
   * @return {Object|null} マッチ結果 {employeeNo, formalName, matchType}
   */
  match(shiftName, facility) {
    if (!shiftName) return null;

    const trimmedName = shiftName.trim();
    if (!trimmedName) return null;

    // 1. 完全一致
    const exactMatch = this.findExactMatch(trimmedName);
    if (exactMatch) return this.toResult(exactMatch, '完全一致');

    // 2. スペース正規化一致
    const normalizedMatch = this.findNormalizedMatch(trimmedName);
    if (normalizedMatch) return this.toResult(normalizedMatch, 'スペース正規化一致');

    // 3. 別名一致
    const aliasMatch = this.findAliasMatch(trimmedName);
    if (aliasMatch) return this.toResult(aliasMatch, '別名一致');

    // 4. 異体字変換一致
    const variantMatch = this.findVariantMatch(trimmedName);
    if (variantMatch) return this.toResult(variantMatch, '異体字変換一致');

    // 5. 姓一致（一意）
    const surnameMatch = this.findUniqueSurnameMatch(trimmedName, facility);
    if (surnameMatch) return this.toResult(surnameMatch, '姓一致');

    // 6. 姓+名先頭一致 (例: "佐藤佳" → "佐藤 佳子")
    const partialMatch = this.findPartialNameMatch(trimmedName);
    if (partialMatch) return this.toResult(partialMatch, '部分一致');

    return null;
  }

  /**
   * マッチ候補を検索（未マッチ時のレポート用）
   * @param {string} shiftName - シフト表の名前
   * @return {Array<string>} 候補リスト
   */
  findCandidates(shiftName) {
    const candidates = [];
    const normalized = this.normalizeName(shiftName);
    const surname = this.extractSurname(shiftName) || shiftName;

    // 姓が一致する従業員を候補に
    const surnameCandidates = this.surnameMap.get(surname);
    if (surnameCandidates) {
      for (const emp of surnameCandidates) {
        candidates.push(emp.employeeNo + ':' + emp.name);
      }
    }

    // 部分一致する従業員を候補に
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      if (emp.name.includes(shiftName) || shiftName.includes(this.extractSurname(emp.name))) {
        const entry = emp.employeeNo + ':' + emp.name;
        if (!candidates.includes(entry)) {
          candidates.push(entry);
        }
      }
    }

    return candidates.slice(0, 5); // 最大5候補
  }

  /**
   * 完全一致検索
   */
  findExactMatch(name) {
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      if (emp.name === name) return emp;
    }
    return null;
  }

  /**
   * スペース正規化一致検索
   * "高橋百合" → "高橋 百合", "石井祐一" → "石井 祐一"
   */
  findNormalizedMatch(name) {
    const normalized = this.normalizeName(name);
    const match = this.normalizedNameMap.get(normalized);
    if (match && match.status !== '退職') return match;

    // スペースなし版でも検索
    const noSpace = name.replace(/[\s　]+/g, '');
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      if (this.normalizeName(emp.name).replace(/[\s　]+/g, '') === noSpace) {
        return emp;
      }
    }
    return null;
  }

  /**
   * 別名一致検索
   */
  findAliasMatch(name) {
    // 完全一致
    const match = this.aliasMap.get(name);
    if (match && match.status !== '退職') return match;

    // 正規化一致
    const normalized = this.normalizeName(name);
    const normalizedMatch = this.normalizedAliasMap.get(normalized);
    if (normalizedMatch && normalizedMatch.status !== '退職') return normalizedMatch;

    return null;
  }

  /**
   * 異体字変換一致検索
   * 﨑↔崎, 髙↔高, ﾜﾌﾞﾘﾆｰｸ→ワブリニーク 等
   */
  findVariantMatch(name) {
    const variants = this.generateVariants(name);
    for (const variant of variants) {
      const match = this.findExactMatch(variant) || this.findNormalizedMatch(variant);
      if (match) return match;
    }
    return null;
  }

  /**
   * 姓一致検索（同姓が1名のみの場合）
   */
  findUniqueSurnameMatch(name, facility) {
    const candidates = this.surnameMap.get(name);
    if (!candidates) return null;

    // 在職者のみフィルタ
    const active = candidates.filter(e => e.status !== '退職');
    if (active.length === 1) return active[0];

    // 複数いる場合は施設で絞り込み
    if (facility && active.length > 1) {
      const facilityMatch = active.filter(e => e.facility === facility);
      if (facilityMatch.length === 1) return facilityMatch[0];
    }

    return null;
  }

  /**
   * 姓+名先頭一致検索
   * "佐藤佳" → "佐藤 佳子", "高橋百合" → "高橋 百合"
   */
  findPartialNameMatch(name) {
    const noSpace = name.replace(/[\s　]+/g, '');
    if (noSpace.length < 2) return null;

    const matches = [];
    for (const emp of this.employees) {
      if (emp.status === '退職') continue;
      const empNoSpace = emp.name.replace(/[\s　]+/g, '');
      // シフト名がマスタ名の先頭部分と一致
      if (empNoSpace.startsWith(noSpace) && empNoSpace.length > noSpace.length) {
        matches.push(emp);
      }
    }

    return matches.length === 1 ? matches[0] : null;
  }

  /**
   * 名前を正規化
   * - 全角スペース→半角スペース
   * - 連続スペース→単一スペース
   * - 前後のスペース除去
   * - 異体字→標準字
   */
  normalizeName(name) {
    if (!name) return '';
    let normalized = String(name);

    // 全角スペース→半角スペース
    normalized = normalized.replace(/　/g, ' ');
    // 連続スペース→単一
    normalized = normalized.replace(/\s+/g, ' ');
    // 前後トリム
    normalized = normalized.trim();
    // 異体字変換
    normalized = this.convertVariantChars(normalized);

    return normalized;
  }

  /**
   * 姓を抽出
   * @param {string} name - フルネーム
   * @return {string} 姓
   */
  extractSurname(name) {
    if (!name) return '';
    const parts = name.trim().split(/[\s　]+/);
    return parts[0] || '';
  }

  /**
   * 名を抽出
   * @param {string} name - フルネーム
   * @return {string} 名
   */
  extractFirstName(name) {
    if (!name) return '';
    const parts = name.trim().split(/[\s　]+/);
    return parts.length > 1 ? parts.slice(1).join(' ') : '';
  }

  /**
   * 異体字変換テーブル
   */
  static get VARIANT_MAP() {
    return {
      '﨑': '崎', '崎': '﨑',
      '髙': '高', '高': '髙',
      '塚': '塚', '塚': '塚',
      '邊': '辺', '邉': '辺', '辺': '邊',
      '齋': '斎', '齊': '斎', '斎': '齋',
      '澤': '沢', '沢': '澤',
      '櫻': '桜', '桜': '櫻',
      '國': '国', '国': '國',
      '龍': '竜', '竜': '龍',
      '藏': '蔵', '蔵': '藏',
      '壽': '寿', '寿': '壽',
      '廣': '広', '広': '廣',
      '惠': '恵', '恵': '惠',
      '學': '学', '學': '学',
      '實': '実', '実': '實',
      '寶': '宝', '宝': '寶'
    };
  }

  /**
   * 異体字を変換
   * @param {string} text - 入力文字列
   * @return {string} 変換後文字列
   */
  convertVariantChars(text) {
    let result = text;
    // 異体字→標準字（崎系統に統一）
    const standardize = {
      '﨑': '崎', '髙': '高', '邊': '辺', '邉': '辺',
      '齋': '斎', '齊': '斎', '澤': '沢', '櫻': '桜',
      '國': '国', '龍': '竜', '藏': '蔵', '壽': '寿',
      '廣': '広', '惠': '恵'
    };
    for (const [from, to] of Object.entries(standardize)) {
      result = result.replace(new RegExp(from, 'g'), to);
    }
    return result;
  }

  /**
   * 異体字バリエーションを生成
   * @param {string} name - 名前
   * @return {Array<string>} バリエーション配列
   */
  generateVariants(name) {
    const variants = new Set();
    const variantMap = NameMatcherEngine.VARIANT_MAP;

    // 各文字について異体字があれば置換版を生成
    for (let i = 0; i < name.length; i++) {
      const char = name[i];
      if (variantMap[char]) {
        const variant = name.substring(0, i) + variantMap[char] + name.substring(i + 1);
        variants.add(variant);
      }
    }

    // 半角カタカナ→全角カタカナ変換
    const fullWidth = this.halfToFullKatakana(name);
    if (fullWidth !== name) {
      variants.add(fullWidth);
    }

    // 全角カタカナ→半角カタカナ変換
    const halfWidth = this.fullToHalfKatakana(name);
    if (halfWidth !== name) {
      variants.add(halfWidth);
    }

    return Array.from(variants);
  }

  /**
   * 半角カタカナ→全角カタカナ変換
   */
  halfToFullKatakana(str) {
    const halfToFull = {
      'ｱ':'ア','ｲ':'イ','ｳ':'ウ','ｴ':'エ','ｵ':'オ',
      'ｶ':'カ','ｷ':'キ','ｸ':'ク','ｹ':'ケ','ｺ':'コ',
      'ｻ':'サ','ｼ':'シ','ｽ':'ス','ｾ':'セ','ｿ':'ソ',
      'ﾀ':'タ','ﾁ':'チ','ﾂ':'ツ','ﾃ':'テ','ﾄ':'ト',
      'ﾅ':'ナ','ﾆ':'ニ','ﾇ':'ヌ','ﾈ':'ネ','ﾉ':'ノ',
      'ﾊ':'ハ','ﾋ':'ヒ','ﾌ':'フ','ﾍ':'ヘ','ﾎ':'ホ',
      'ﾏ':'マ','ﾐ':'ミ','ﾑ':'ム','ﾒ':'メ','ﾓ':'モ',
      'ﾔ':'ヤ','ﾕ':'ユ','ﾖ':'ヨ',
      'ﾗ':'ラ','ﾘ':'リ','ﾙ':'ル','ﾚ':'レ','ﾛ':'ロ',
      'ﾜ':'ワ','ｦ':'ヲ','ﾝ':'ン',
      'ｧ':'ァ','ｨ':'ィ','ｩ':'ゥ','ｪ':'ェ','ｫ':'ォ',
      'ｯ':'ッ','ｬ':'ャ','ｭ':'ュ','ｮ':'ョ',
      'ﾞ':'゛','ﾟ':'゜','ｰ':'ー'
    };

    let result = '';
    for (let i = 0; i < str.length; i++) {
      const char = str[i];
      const nextChar = str[i + 1];
      if (halfToFull[char]) {
        let fullChar = halfToFull[char];
        // 濁点・半濁点の結合
        if (nextChar === 'ﾞ' && 'カキクケコサシスセソタチツテトハヒフヘホウ'.includes(fullChar)) {
          const dakuten = {'カ':'ガ','キ':'ギ','ク':'グ','ケ':'ゲ','コ':'ゴ',
                          'サ':'ザ','シ':'ジ','ス':'ズ','セ':'ゼ','ソ':'ゾ',
                          'タ':'ダ','チ':'ヂ','ツ':'ヅ','テ':'デ','ト':'ド',
                          'ハ':'バ','ヒ':'ビ','フ':'ブ','ヘ':'ベ','ホ':'ボ',
                          'ウ':'ヴ'};
          if (dakuten[fullChar]) {
            fullChar = dakuten[fullChar];
            i++; // 濁点をスキップ
          }
        } else if (nextChar === 'ﾟ' && 'ハヒフヘホ'.includes(fullChar)) {
          const handakuten = {'ハ':'パ','ヒ':'ピ','フ':'プ','ヘ':'ペ','ホ':'ポ'};
          if (handakuten[fullChar]) {
            fullChar = handakuten[fullChar];
            i++;
          }
        }
        result += fullChar;
      } else {
        result += char;
      }
    }
    return result;
  }

  /**
   * 全角カタカナ→半角カタカナ変換
   */
  fullToHalfKatakana(str) {
    const fullToHalf = {
      'ア':'ｱ','イ':'ｲ','ウ':'ｳ','エ':'ｴ','オ':'ｵ',
      'カ':'ｶ','キ':'ｷ','ク':'ｸ','ケ':'ｹ','コ':'ｺ',
      'サ':'ｻ','シ':'ｼ','ス':'ｽ','セ':'ｾ','ソ':'ｿ',
      'タ':'ﾀ','チ':'ﾁ','ツ':'ﾂ','テ':'ﾃ','ト':'ﾄ',
      'ナ':'ﾅ','ニ':'ﾆ','ヌ':'ﾇ','ネ':'ﾈ','ノ':'ﾉ',
      'ハ':'ﾊ','ヒ':'ﾋ','フ':'ﾌ','ヘ':'ﾍ','ホ':'ﾎ',
      'マ':'ﾏ','ミ':'ﾐ','ム':'ﾑ','メ':'ﾒ','モ':'ﾓ',
      'ヤ':'ﾔ','ユ':'ﾕ','ヨ':'ﾖ',
      'ラ':'ﾗ','リ':'ﾘ','ル':'ﾙ','レ':'ﾚ','ロ':'ﾛ',
      'ワ':'ﾜ','ヲ':'ｦ','ン':'ﾝ',
      'ァ':'ｧ','ィ':'ｨ','ゥ':'ｩ','ェ':'ｪ','ォ':'ｫ',
      'ッ':'ｯ','ャ':'ｬ','ュ':'ｭ','ョ':'ｮ',
      'ー':'ｰ',
      'ガ':'ｶﾞ','ギ':'ｷﾞ','グ':'ｸﾞ','ゲ':'ｹﾞ','ゴ':'ｺﾞ',
      'ザ':'ｻﾞ','ジ':'ｼﾞ','ズ':'ｽﾞ','ゼ':'ｾﾞ','ゾ':'ｿﾞ',
      'ダ':'ﾀﾞ','ヂ':'ﾁﾞ','ヅ':'ﾂﾞ','デ':'ﾃﾞ','ド':'ﾄﾞ',
      'バ':'ﾊﾞ','ビ':'ﾋﾞ','ブ':'ﾌﾞ','ベ':'ﾍﾞ','ボ':'ﾎﾞ',
      'パ':'ﾊﾟ','ピ':'ﾋﾟ','プ':'ﾌﾟ','ペ':'ﾍﾟ','ポ':'ﾎﾟ',
      'ヴ':'ｳﾞ'
    };

    let result = '';
    for (const char of str) {
      result += fullToHalf[char] || char;
    }
    return result;
  }

  /**
   * マッチ結果オブジェクトを生成
   */
  toResult(employee, matchType) {
    return {
      employeeNo: employee.employeeNo,
      formalName: employee.name,
      matchType: matchType
    };
  }
}
