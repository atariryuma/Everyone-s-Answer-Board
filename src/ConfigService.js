/**
 * @fileoverview ConfigService - configJSON の CRUD / validation / sanitize /
 *   動的 URL 生成 / profiles・profileHistory のサニタイズ。
 */

/* global getCurrentEmail, findUserById, updateUser, SYSTEM_LIMITS, validateConfig, validateSpreadsheetId, openSpreadsheet, getSheetInfo, DEFAULT_DISPLAY_SETTINGS, getCachedProperty, logError_ */

/**
 * デフォルト設定取得
 * @param {string} userId - ユーザーID
 * @returns {Object} デフォルト設定
 */
function getDefaultConfig(userId) {
  return {
    userId,
    setupStatus: 'pending',
    isPublished: false,
    hasSeenWelcome: false,
    displaySettings: { ...DEFAULT_DISPLAY_SETTINGS },
    completionScore: 0
  };
}

/**
 * 設定JSONパース・修復
 * @param {string} configJson - 設定JSON文字列
 * @param {string} userId - ユーザーID
 * @returns {Object} パース済み設定
 */
function parseAndRepairConfig(configJson, userId) {
  try {
    const config = JSON.parse(configJson || '{}');

    const repairedConfig = repairNestedConfig(config, userId);

    return ensureRequiredFields(repairedConfig, userId);
  } catch (parseError) {
    // Why: configJson の JSON.parse が失敗するのは「設定破損」「部分書込み」のいずれか。
    //   silently default に戻すと、formUrl/spreadsheetId などの公開情報が消えるサイレントデータ
    //   ロスにつながる。WARN ではなく ERROR レベルで明示し、parsed=null マーカーを付けて
    //   呼び出し元が "復旧不能" と気付けるようにする。デフォルトは fallback として返すが、
    //   __parseFailed フラグで「これは健全な config ではない」を伝播する。
    console.error('parseAndRepairConfig: JSON解析失敗 - 設定破損の可能性', {
      operation: 'parseAndRepairConfig',
      userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
      configLength: configJson?.length || 0,
      error: parseError.message,
      stack: parseError.stack
    });
    const fallback = getDefaultConfig(userId);
    fallback.__parseFailed = true;
    fallback.__parseError = parseError.message;
    return fallback;
  }
}

/**
 * ネストされた設定構造の修復
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 修復済み設定
 */
function repairNestedConfig(config, userId) {
  const repaired = { ...config };

  if (!repaired.displaySettings || typeof repaired.displaySettings !== 'object') {
    repaired.displaySettings = { ...DEFAULT_DISPLAY_SETTINGS };
  }

  if (!repaired.columnMapping || typeof repaired.columnMapping !== 'object') {
    repaired.columnMapping = {};
  }

  return repaired;
}

/**
 * 必須フィールドの確保（残りのフィールドは保持する）
 *
 * Why: 以前は allowlist ベースで必須フィールドだけ残していたので、DB に保存されていた
 *      formTitle / publishedAt / hasSeenWelcome / lastAccessedAt / etag / showDetails が
 *      毎回のリード→ライト往復で silently 消えていた（critical data loss）。
 *      スプレッド演算子で入力を保持し、必須フィールドだけ normalize する。
 */
function ensureRequiredFields(config, userId) {
  return {
    ...config,
    userId,
    setupStatus: config.setupStatus || 'pending',
    isPublished: Boolean(config.isPublished),
    spreadsheetId: config.spreadsheetId || '',
    sheetName: config.sheetName || '',
    formUrl: config.formUrl || '',
    displaySettings: config.displaySettings || { ...DEFAULT_DISPLAY_SETTINGS },
    columnMapping: config.columnMapping,
    completionScore: calculateCompletionScore(config)
  };
}

/**
 * 動的URL付きの設定拡張
 * @param {Object} baseConfig - 基本設定
 * @param {string} userId - ユーザーID
 * @returns {Object} URL拡張済み設定
 */
function enhanceConfigWithDynamicUrls(baseConfig, userId) {
  const enhanced = { ...baseConfig };

  try {
    const webAppUrl = getWebAppUrl();

    enhanced.dynamicUrls = {
      webAppUrl: webAppUrl || '',
      adminPanelUrl: webAppUrl ? `${webAppUrl}?mode=admin&userId=${userId}` : '',
      viewBoardUrl: webAppUrl ? `${webAppUrl}?mode=view&userId=${userId}` : '',
      setupUrl: webAppUrl ? `${webAppUrl}?mode=setup&userId=${userId}` : '',
      manualUrl: webAppUrl ? `${webAppUrl}?mode=manual` : ''
    };

  } catch (error) {
    console.warn('enhanceConfigWithDynamicUrls: URL生成エラー', error.message);
    enhanced.dynamicUrls = {};
  }

  return enhanced;
}

// 公開時用の設定検証（必須フィールドを厳格にチェック）。
function validatePublishConfig(config, userId) {
  try {
    const baseValidation = validateAndSanitizeConfig(config, userId);
    if (!baseValidation.success) {
      return baseValidation;
    }

    const errors = [];

    if (!config.spreadsheetId || typeof config.spreadsheetId !== 'string' || !config.spreadsheetId.trim()) {
      errors.push('公開にはスプレッドシートIDが必要です');
    } else if (typeof validateSpreadsheetId === 'function') {
      // Why: 旧実装は trim() しかチェックしておらず "x" でも通る。validators.js の正規パターン
      //   (^[a-zA-Z0-9-_]{40,50}$) で publish 時の厳格検証を行い、不正 ID の混入を防ぐ。
      const v = validateSpreadsheetId(config.spreadsheetId);
      if (v && !v.isValid) {
        errors.push('スプレッドシートID形式が不正です: ' + ((v.errors && v.errors[0]) || 'unknown'));
      }
    }

    if (!config.sheetName || typeof config.sheetName !== 'string' || !config.sheetName.trim()) {
      errors.push('公開にはシート名が必要です');
    }

    if (!config.columnMapping || typeof config.columnMapping !== 'object') {
      errors.push('公開には列マッピングが必要です');
    } else if (Object.keys(config.columnMapping).length === 0) {
      errors.push('公開には列マッピングが必要です（空のマッピングは無効）');
    } else {
      const answerColumn = config.columnMapping.answer;
      if (answerColumn === undefined || answerColumn === null || (typeof answerColumn === 'number' && answerColumn < 0)) {
        errors.push('公開には回答列（answer）の設定が必要です');
      }
    }

    if (errors.length > 0) {
      return {
        success: false,
        message: '公開に必要な設定が不足しています',
        errors,
        data: baseValidation.data
      };
    }

    return baseValidation;

  } catch (error) {
    logError_('validatePublishConfig', error);
    return {
      success: false,
      message: '公開設定検証中にエラーが発生しました',
      errors: [error.message]
    };
  }
}

/**
 * 設定検証・サニタイズ（統合版）
 * @param {Object} config - 設定オブジェクト
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validateAndSanitizeConfig(config, userId) {
  try {
    const validationResult = validateConfig(config);
    if (!validationResult.isValid) {
      return {
        success: false,
        message: '設定検証エラーがあります',
        errors: validationResult.errors,
        data: validationResult.sanitized || config
      };
    }

    const errors = [];
    if (!validateConfigUserId(userId)) {
      errors.push('無効なユーザーID形式');
    }

    // 既存設定フィールドを保持しつつ、検証済みフィールドで上書きする。
    // Why spread: allowlist 化すると過去に formTitle/publishedAt/etag 等が read→write 往復で
    //   silently 消える data loss を起こしたため、未知の正規フィールドは保持する方針。
    const sanitized = { ...config, ...validationResult.sanitized };

    // ただし「永続化してはいけないキー」だけは明示的に剥がす:
    //   - prototype 汚染ベクタ (__proto__ / constructor / prototype)
    //   - parse 失敗マーカー (__parseFailed/__parseError: 実行時のみ意味を持つ)
    //   - 計算で都度生成される派生データ (dynamicUrls: 保存すると stale URL が残る)
    // これにより spread の data-loss 安全性を保ちつつ未知キー素通しの面を絞る。
    const TRANSIENT_OR_DANGEROUS_KEYS = ['__proto__', 'constructor', 'prototype', '__parseFailed', '__parseError', 'dynamicUrls'];
    TRANSIENT_OR_DANGEROUS_KEYS.forEach((k) => { delete sanitized[k]; });

    sanitized.userId = userId;

    if (sanitized.displaySettings) {
      sanitized.displaySettings = sanitizeDisplaySettings(sanitized.displaySettings);
    }
    if (sanitized.columnMapping) {
      sanitized.columnMapping = sanitizeMapping(sanitized.columnMapping);
    }

    // 可視化モード補助メタデータ。null を返す sanitize 関数で空オブジェクトを排除。
    if ('xAxisLabels' in sanitized) {
      const v = sanitizeAxisLabels(sanitized.xAxisLabels);
      if (v) sanitized.xAxisLabels = v; else delete sanitized.xAxisLabels;
    }
    if ('yAxisLabels' in sanitized) {
      const v = sanitizeAxisLabels(sanitized.yAxisLabels);
      if (v) sanitized.yAxisLabels = v; else delete sanitized.yAxisLabels;
    }
    if ('matrixQuadrantLabels' in sanitized) {
      const v = sanitizeQuadrantLabels(sanitized.matrixQuadrantLabels);
      if (v) sanitized.matrixQuadrantLabels = v; else delete sanitized.matrixQuadrantLabels;
    }
    if ('allowResubmit' in sanitized) {
      sanitized.allowResubmit = Boolean(sanitized.allowResubmit);
    }

    // multi-board: profiles 配列とアクティブプロファイル名
    // Why: 1 ユーザーが複数 Forms を切替えて使えるよう、設定スナップショットを保持。
    //      profiles が空配列のときは config から完全に削除して JSON サイズを節約。
    if ('profiles' in sanitized) {
      const profs = sanitizeProfiles(sanitized.profiles);
      if (profs.length > 0) sanitized.profiles = profs;
      else delete sanitized.profiles;
    }
    if ('activeProfile' in sanitized) {
      const name = typeof sanitized.activeProfile === 'string' ? sanitized.activeProfile.trim() : '';
      if (name) sanitized.activeProfile = name.substring(0, 50);
      else delete sanitized.activeProfile;
    }

    // profileHistory: navigable な過去 phase の時系列ログ。
    // Why cross-ref (v2772 で再定義): profiles[] に存在しない orphan を許容すると、
    //   削除済 profile が strikethrough で client に残り続け UX 混乱を起こしていた。
    //   sanitizeProfileHistory に sanitized.profiles を渡し cross-ref filter させることで、
    //   delete / rename / 直編集 すべての mutation で自動的に同期する (SSoT)。
    //   実行順序が大事: 上の sanitizeProfiles 完了後の sanitized.profiles を参照する。
    if ('profileHistory' in sanitized) {
      const hist = sanitizeProfileHistory(sanitized.profileHistory, sanitized.profiles);
      if (hist.length > 0) sanitized.profileHistory = hist;
      else delete sanitized.profileHistory;
    }

    try {
      const configSize = JSON.stringify(sanitized).length;
      if (configSize > SYSTEM_LIMITS.CONFIG_JSON_MAX_CHARS) {
        errors.push('設定データが大きすぎます（32KB制限）');
      }
    } catch (sizeCheckError) {
      errors.push(`設定サイズ検証に失敗しました: ${sizeCheckError.message}`);
    }

    if (errors.length > 0) {
      return {
        success: false,
        message: '設定検証エラーがあります',
        errors,
        data: sanitized
      };
    }

    return {
      success: true,
      message: '設定検証完了',
      data: sanitized,
      errors: []
    };

  } catch (error) {
    logError_('validateAndSanitizeConfig', error);
    return {
      success: false,
      message: '設定検証中にエラーが発生しました',
      errors: [error.message]
    };
  }
}

/**
 * 表示設定サニタイズ
 * @param {Object} displaySettings - 表示設定
 * @returns {Object} サニタイズ済み表示設定
 */
// Why: boardMode の正規 enum 定義は validators.js の VALIDATOR_BOARD_MODES (single source of truth)。
//   ここでは GAS の単一グローバルスコープでミラー参照する。値が divergent にならないよう
//   validators.js が定義済の場合はそちらを尊重し、未定義のときだけ fallback で定義。
const VALID_BOARD_MODES = (typeof VALIDATOR_BOARD_MODES !== 'undefined' && Array.isArray(VALIDATOR_BOARD_MODES))
  ? VALIDATOR_BOARD_MODES
  : ['auto', 'board', 'numberline', 'matrix', 'wordcloud', 'pie'];

// Why __strictBool: Boolean("false") は truthy 評価で true を返すため、文字列で
//   永続化された display setting (旧クライアント or migration 経由) が読込時に privacy
//   regression を引き起こす ("匿名" だったはずのボードが "実名表示" に化ける) 危険があった。
//   許容: true literal / "true" / "1" / 1 を真とし、それ以外は false に倒す。
function __strictBool(v) {
  if (v === true || v === 1) return true;
  if (typeof v === 'string') {
    const lower = v.trim().toLowerCase();
    return lower === 'true' || lower === '1';
  }
  return false;
}

function sanitizeDisplaySettings(displaySettings) {
  const sanitized = {
    showNames: __strictBool(displaySettings.showNames),
    showReactions: __strictBool(displaySettings.showReactions),
    theme: String(displaySettings.theme || 'default').substring(0, SYSTEM_LIMITS.PREVIEW_LENGTH),
    pageSize: Math.min(Math.max(Number(displaySettings.pageSize) || SYSTEM_LIMITS.DEFAULT_PAGE_SIZE, 1), SYSTEM_LIMITS.MAX_PAGE_SIZE)
  };

  const mode = displaySettings.boardMode;
  if (typeof mode === 'string' && VALID_BOARD_MODES.includes(mode)) {
    sanitized.boardMode = mode;
  }

  return sanitized;
}

/**
 * 列マッピングサニタイズ
 * @param {Object} columnMapping - 列マッピング
 * @returns {Object} サニタイズ済み列マッピング
 */
function sanitizeMapping(columnMapping) {
  const sanitized = {};
  // Why: numericX/numericY は可視化モード用の数値列指標。
  //      Forms「線形尺度」列を指す。answer/reason と独立して任意設定可能。
  const validFields = ['answer', 'reason', 'class', 'name', 'timestamp', 'email', 'numericX', 'numericY'];

  validFields.forEach(field => {
    const index = columnMapping[field];
    if (typeof index === 'number' && index >= 0 && index < 100) {
      sanitized[field] = index;
    }
  });

  return sanitized;
}

/**
 * 軸ラベル等の可視化メタデータをサニタイズ。
 * Why: M1/M2 で「線形尺度」の両端ラベルを管理画面から手入力で持たせるため。
 *      Forms から自動取得は API 非対応なので、保存可能な構造体として扱う。
 */
function sanitizeAxisLabels(input) {
  if (!input || typeof input !== 'object') return null;
  const cap = (v) => String(v == null ? '' : v).substring(0, 40);
  const out = { min: cap(input.min), max: cap(input.max) };
  if (!out.min && !out.max) return null;
  return out;
}

function sanitizeQuadrantLabels(input) {
  if (!input || typeof input !== 'object') return null;
  const cap = (v) => String(v == null ? '' : v).substring(0, 40);
  const out = {
    hh: cap(input.hh),
    hl: cap(input.hl),
    lh: cap(input.lh),
    ll: cap(input.ll)
  };
  // 全て空なら null を返す（DBに不要なノイズを残さない）
  if (!out.hh && !out.hl && !out.lh && !out.ll) return null;
  return out;
}

/**
 * profileHistory のサニタイズ。
 *
 * Why: 生徒が過去フェーズを閲覧する Option B の根拠データ。entry は
 *   `{ name: string, activatedAt: ISO8601 string }` の形を期待。最新が末尾。
 *   無効 entry は捨て、最新 50 件のみ保持（古い順）。
 *   activatedAt は client 側で時系列表示に使うため文字列のまま保存。
 *
 * @param {Array} input - profileHistory 候補
 * @returns {Array<{name:string, activatedAt:string}>}
 */
function sanitizeProfileHistory(input, sanitizedProfiles) {
  if (!Array.isArray(input)) return [];
  const MAX_HISTORY = 50;

  // Why cross-ref: profileHistory は「現在 navigable な過去 phase の時系列ログ」と再定義。
  //   profiles[] に存在しない (= 削除済) 名前は orphan として drop する。
  //   これで delete / rename / 直編集 すべてで自動的に history が同期する
  //   (各 mutation 経路で個別に history を更新する必要がなくなる、SSoT)。
  //   sanitizedProfiles が undefined のときは cross-ref を skip (後方互換 / test 用)。
  const validNames = (Array.isArray(sanitizedProfiles) && sanitizedProfiles.length > 0)
    ? new Set(sanitizedProfiles.filter(p => p && p.name).map(p => String(p.name).trim()))
    : null;

  const out = [];
  for (const e of input) {
    if (!e || typeof e !== 'object') continue;
    const name = String(e.name == null ? '' : e.name).trim().substring(0, 50);
    if (!name) continue;
    if (validNames && !validNames.has(name)) continue;  // orphan: drop
    let activatedAt = '';
    if (typeof e.activatedAt === 'string') {
      activatedAt = e.activatedAt.substring(0, 40);
    } else if (e.activatedAt instanceof Date) {
      activatedAt = e.activatedAt.toISOString();
    }
    out.push({ name, activatedAt });
  }
  // 末尾を新しい側として MAX_HISTORY 件まで残す。
  if (out.length > MAX_HISTORY) return out.slice(out.length - MAX_HISTORY);
  return out;
}

/**
 * profiles 配列のサニタイズ。
 *
 * Why: multi-board のために 1 user の config 内に複数の「表示設定スナップショット」を
 *      保持する。各 profile は active config と同じ形（formUrl/spreadsheetId/sheetName/
 *      columnMapping/displaySettings + 軸ラベル等）の subset。
 *
 *   - 最大 10 profile（config サイズ膨張防止）
 *   - 各 name は最大 50 char
 *   - 既存の sanitize 関数を再利用してフィールド単位で正規化
 *   - サポート外フィールドは silently drop
 *
 * @param {Array} input - profiles 候補配列
 * @returns {Array} sanitized profiles（空配列なら caller で削除する）
 */
function sanitizeProfiles(input) {
  if (!Array.isArray(input)) return [];
  const MAX_PROFILES = 10;
  const out = [];
  for (const p of input) {
    if (!p || typeof p !== 'object') continue;
    const name = String(p.name == null ? '' : p.name).trim().substring(0, 50);
    if (!name) continue; // name 必須

    const cleaned = { name };
    // データソース系（文字列）
    if (typeof p.formUrl === 'string') cleaned.formUrl = p.formUrl.substring(0, 500);
    if (typeof p.formTitle === 'string') cleaned.formTitle = p.formTitle.substring(0, 200);
    if (typeof p.spreadsheetId === 'string') cleaned.spreadsheetId = p.spreadsheetId.substring(0, 100);
    if (typeof p.sheetName === 'string') cleaned.sheetName = p.sheetName.substring(0, 100);
    // 列マッピング・表示設定（既存 sanitize を再利用）
    if (p.columnMapping && typeof p.columnMapping === 'object') {
      cleaned.columnMapping = sanitizeMapping(p.columnMapping);
    }
    if (p.displaySettings && typeof p.displaySettings === 'object') {
      cleaned.displaySettings = sanitizeDisplaySettings(p.displaySettings);
    }
    // 可視化メタデータ
    const x = sanitizeAxisLabels(p.xAxisLabels);
    if (x) cleaned.xAxisLabels = x;
    const y = sanitizeAxisLabels(p.yAxisLabels);
    if (y) cleaned.yAxisLabels = y;
    const q = sanitizeQuadrantLabels(p.matrixQuadrantLabels);
    if (q) cleaned.matrixQuadrantLabels = q;
    if (typeof p.allowResubmit !== 'undefined') cleaned.allowResubmit = Boolean(p.allowResubmit);

    out.push(cleaned);
    if (out.length >= MAX_PROFILES) break;
  }
  return out;
}

/**
 * ユーザーID検証
 * @param {string} userId - ユーザーID
 * @returns {boolean} 有効性
 */
function validateConfigUserId(userId) {
  return typeof userId === 'string' && userId.length > 0 && userId.length <= 100;
}

/**
 * 設定完了スコア計算
 * @param {Object} config - 設定オブジェクト
 * @returns {number} 完了スコア (0-100)
 */
function calculateCompletionScore(config) {
  if (!config) return 0;

  let score = 0;

  if (config.spreadsheetId) score += 25;
  if (config.sheetName) score += 25;
  if (config.formUrl) score += 30;
  if (config.displaySettings) score += 10;
  if (config.columnMapping && Object.keys(config.columnMapping).length > 0) score += 10;

  return score;
}

function getUserCacheKeys_(userId) {
  return [
    `config_${userId}`,
    `user_${userId}`,
    `board_data_${userId}`,
    `admin_panel_${userId}`,
    `question_text_${userId}`
  ];
}

/**
 * 設定キャッシュクリア
 * @param {string} userId - ユーザーID
 */
function clearConfigCache(userId) {
  try {
    CacheService.getScriptCache().removeAll(getUserCacheKeys_(userId));
  } catch (error) {
    console.warn('clearConfigCache: キャッシュクリアエラー', error.message);
  }
}

/**
 * コアシステムプロパティ確認 - 3つの必須項目をすべてチェック
 *
 * Why cache: doGet のどのモードでも少なくとも 1 回呼ばれる hot path。
 *   生の PropertiesService 読み取りを 3 回 + JSON.parse していたのを、
 *   getCachedProperty (30s メモリキャッシュ) に乗せ替える。
 *   setCachedProperty 経由の書き込みで個別キーは即 invalidation されるので、
 *   セットアップ直後の false → true 反映も問題ない。
 *
 * @returns {boolean} 3つすべて存在すれば true
 */
function hasCoreSystemProps() {
  try {
    const adminEmail = getCachedProperty('ADMIN_EMAIL');
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    const creds = getCachedProperty('SERVICE_ACCOUNT_CREDS');

    if (!adminEmail || !dbId || !creds) {
      console.warn('hasCoreSystemProps: 必須項目不足', {
        hasAdmin: !!adminEmail,
        hasDb: !!dbId,
        hasCreds: !!creds
      });
      return false;
    }

    try {
      const parsed = JSON.parse(creds);
      if (!parsed || typeof parsed !== 'object' || !parsed.client_email) {
        console.warn('hasCoreSystemProps: SERVICE_ACCOUNT_CREDS JSON不正');
        return false;
      }
    } catch (jsonError) {
      console.warn('hasCoreSystemProps: SERVICE_ACCOUNT_CREDS JSON解析失敗', jsonError.message);
      return false;
    }

    return true;
  } catch (error) {
    logError_('hasCoreSystemProps', error);
    return false;
  }
}

/**
 * 動的questionText取得（configJson最適化対応 + パフォーマンス最適化済み）
 * headers配列とcolumnMappingから実際の問題文を動的取得
 * preloadedHeadersが提供された場合は重複スプレッドシートアクセスを回避
 * @param {Object} config - ユーザー設定オブジェクト
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @param {Array} preloadedHeaders - 事前取得されたヘッダー配列（パフォーマンス最適化用）
 * @returns {string} 問題文テキスト
 */
function getQuestionText(config, context = {}, preloadedHeaders = null) {
  try {

    const answerIndex = config?.columnMapping?.answer;

    if (typeof answerIndex === 'number' && config?.headers?.[answerIndex]) {
      const questionText = config.headers[answerIndex];
      if (questionText && typeof questionText === 'string' && questionText.trim()) {
        return questionText.trim();
      }
    }

    if (typeof answerIndex === 'number' && preloadedHeaders && preloadedHeaders[answerIndex]) {
      const questionText = preloadedHeaders[answerIndex];
      if (questionText && typeof questionText === 'string' && questionText.trim()) {
        return questionText.trim();
      }
    }

    if (typeof answerIndex === 'number' && config?.spreadsheetId && config?.sheetName) {
      try {
        // accessMode 自動判定: owner = openById、viewer/admin = SA pool 経由。
        try {
          const dataAccess = openSpreadsheet(config.spreadsheetId, { context: 'configService.questionText' });
          const { spreadsheet } = dataAccess;
          const sheet = spreadsheet.getSheetByName(config.sheetName);
          if (sheet) {
            const { lastCol, headers } = getSheetInfo(sheet);
            if (lastCol > 0 && headers && headers[answerIndex]) {
              const questionText = headers[answerIndex];
              if (questionText && typeof questionText === 'string' && questionText.trim()) {
                return questionText.trim();
              }
            }
          }
        } catch (dataAccessError) {
          console.warn('⚠️ getQuestionText: openSpreadsheet access failed:', dataAccessError.message);
        }
      } catch (dynamicError) {
        console.warn('⚠️ getQuestionText: Dynamic headers fetch failed:', dynamicError.message);
      }
    }

    if (config?.formTitle && typeof config.formTitle === 'string' && config.formTitle.trim()) {
      return config.formTitle.trim();
    }

    return 'Everyone\'s Answer Board';
  } catch (error) {
    logError_('getQuestionText', error);
    return 'Everyone\'s Answer Board';
  }
}

/**
 * 統一設定読み込みAPI - V8最適化、変数チェック強化
 * main.gs内の直接JSON.parse()操作を置換する統一関数
 * @param {string} userId - ユーザーID
 * @param {Object} preloadedUser - 事前取得済みユーザーオブジェクト（パフォーマンス最適化用）
 * @returns {Object} {success: boolean, config: Object, message?: string, userId?: string}
 */
function getUserConfig(userId, preloadedUser = null) {
  if (!userId || typeof userId !== 'string' || !userId.trim()) {
    return {
      success: false,
      config: getDefaultConfig(null),
      message: 'Invalid userId provided'
    };
  }

  try {
    const user = preloadedUser || findUserById(userId, {
      requestingUser: getCurrentEmail()
    });
    if (!user) {
      return {
        success: false,
        config: getDefaultConfig(userId),
        message: 'User not found',
        userId
      };
    }

    if (!user.configJson || typeof user.configJson !== 'string') {
      return {
        success: true,
        config: getDefaultConfig(userId),
        message: 'No config found, using defaults',
        userId
      };
    }

    const config = parseAndRepairConfig(user.configJson, userId);

    // __parseFailed = JSON 破損で default に fallback した状態。 read は success:true で
    //   default を返して graceful degradation するが、 write 経路がこの default を保存すると
    //   破損だが復旧可能だった原本を上書きしてしまう。 envelope に corrupted を surface し、
    //   applyConfigPatch_ / __applyPublishStateChange が上書きを拒否できるようにする。
    return {
      success: true,
      config,
      corrupted: Boolean(config && config.__parseFailed),
      userId,
      message: config && config.__parseFailed
        ? 'Config corrupted (parse failed); returning defaults'
        : 'Config loaded successfully'
    };

  } catch (error) {
    logError_('getUserConfig', error, {
      userId: userId ? `${userId.substring(0, 8)}***` : 'N/A'
    });
    return {
      success: false,
      config: getDefaultConfig(userId),
      message: error && error.message ? `Config load error: ${error.message}` : 'Config load error: 詳細不明',
      userId
    };
  }
}

/**
 * 統一設定保存API - 検証+サニタイズ必須
 * main.gs内の直接JSON.stringify()操作を置換する統一関数
 * @param {string} userId - ユーザーID
 * @param {Object} config - 保存する設定オブジェクト
 * @param {Object} options - 保存オプション
 * @returns {Object} {success: boolean, message: string, data?: Object}
 */
function saveUserConfig(userId, config, options = {}) {
  if (!userId || typeof userId !== 'string' || !userId.trim()) {
    return {
      success: false,
      message: 'Invalid userId provided'
    };
  }

  if (!config || typeof config !== 'object') {
    return {
      success: false,
      message: 'Invalid config object provided'
    };
  }

  try {
    const user = findUserById(userId, { requestingUser: getCurrentEmail() });
    if (!user) {
      return {
        success: false,
        message: 'User not found'
      };
    }

    // user.configJson を 1 度だけ parse し、 etag 検証 + 後段の publish-state guard で共有する
    //   (旧実装は同じ文字列を最大 2 回 parse していた)。 parse 失敗時は null = 検証スキップ。
    let persistedConfig = null;
    if (user.configJson) {
      try {
        persistedConfig = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('saveUserConfig: persisted config parse failed:', parseError.message);
      }
    }

    if (config.etag) {
      if (persistedConfig) {
        const currentETag = persistedConfig.etag || user.lastModified;

        if (currentETag && config.etag !== currentETag) {
          console.warn('saveUserConfig: ETag mismatch detected', {
            userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
            requestETag: config.etag,
            currentETag
          });

          return {
            success: false,
            error: 'etag_mismatch',
            message: 'Configuration has been modified by another user',
            currentConfig: persistedConfig
          };
        }
      }
    } else if (persistedConfig && persistedConfig.etag) {
      // etag を渡さない save は楽観ロックをスキップする (last-write-wins)。 既存 config が
      // etag を持つのに etag 無しで上書きする = 別タブ/別 request の編集を黙って飲み込む
      // 可能性がある。 silent lost-update を観測可能にするため WARN を残す
      // (hard reject はしない: profile/lesson 等 etag を渡さない正当な内部 caller が多数あるため)。
      console.warn('saveUserConfig: etag-less overwrite of an etag-bearing config (optimistic lock skipped)', {
        userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
        currentETag: persistedConfig.etag
      });
    }

    const validation = options.isPublish
      ? validatePublishConfig(config, userId)
      : validateAndSanitizeConfig(config, userId);
    if (!validation.success) {
      return {
        success: false,
        message: validation.message,
        errors: validation.errors
      };
    }

    const cleanedConfig = cleanConfigFields(validation.data, options);

    // Why (publish lifecycle single-source-of-truth, v2865): config.isPublished / publishedAt は
    //   __applyPublishStateChange と publishApp の 2 経路だけが書き換えてよい (CLAUDE.md の不変条件)。
    //   それ以外の save (frontend autosave / applyConfigPatch_ / profile 切替 / markWelcomeSeen 等) が
    //   stale な isPublished を運んできても、ここで現在永続化されている値で強制上書きし「公開状態は
    //   変えない」を保証する。strip ではなく現値で上書きするのは、 field を消すと
    //   isUserBoardPublished が false 化して意図せぬ非公開バグになるため (data-loss 回避)。
    //   公開状態を変える正規経路は必ず options.__allowPublishStateWrite を付けて呼ぶ。
    if (!options.__allowPublishStateWrite) {
      // 冒頭で parse 済の persistedConfig を再利用 (parse 失敗 = null なら未公開既定)。
      const persistedPublished = persistedConfig ? persistedConfig.isPublished : undefined;
      const persistedPublishedAt = persistedConfig ? persistedConfig.publishedAt : undefined;
      cleanedConfig.isPublished = persistedPublished === true;
      cleanedConfig.publishedAt = (persistedPublished === true) ? (persistedPublishedAt || null) : null;
    }

    if (!cleanedConfig.lastAccessedAt) {
      cleanedConfig.lastAccessedAt = new Date().toISOString();
    }

    // Why (orphan spreadsheetId 検出): top-level spreadsheetId が profiles のどれにも
    //   一致しない場合、profile 削除や直接編集の残骸である可能性が高い。polling /
    //   getPublishedSheetData が古い ID を見にいって 403 を出し続ける根本原因なので
    //   WARN を出して気付けるようにする (削除はしない — 移行期間の互換性のため)。
    //   activeProfile が空 (= profile 機能未使用) の場合はスキップ。
    if (cleanedConfig.activeProfile && cleanedConfig.spreadsheetId && Array.isArray(cleanedConfig.profiles)) {
      const inProfiles = cleanedConfig.profiles.some(p => p && p.spreadsheetId === cleanedConfig.spreadsheetId);
      if (!inProfiles) {
        console.warn('saveUserConfig: orphan spreadsheetId (top-level does not match any profile)', {
          userId: userId && typeof userId === 'string' ? `${userId.substring(0, 8)}***` : 'N/A',
          topLevelSpreadsheetId: `${cleanedConfig.spreadsheetId.substring(0, 20)}...`,
          activeProfile: cleanedConfig.activeProfile,
          profileCount: cleanedConfig.profiles.length
        });
      }
    }

    const currentTime = new Date().toISOString();
    // Why: Math.random() は予測可能で 2 つの同時 save が同 etag になる衝突リスクがある。
    //   Utilities.getUuid() は GAS が提供する RFC 4122 v4 UUID で衝突確率 ~0。
    //   フォーマット: `<ISO ts>_<uuid 12 文字>` (検索性 + 一意性の両立)。
    const etagUuid = (typeof Utilities !== 'undefined' && Utilities.getUuid)
      ? Utilities.getUuid().replace(/-/g, '').substring(0, 12)
      : Math.random().toString(36).substring(2, 14);
    cleanedConfig.etag = `${currentTime}_${etagUuid}`;

    const updateResult = updateUser(userId, {
      configJson: JSON.stringify(cleanedConfig)
    });

    if (!updateResult || !updateResult.success) {
      return {
        success: false,
        message: updateResult?.message || 'Database update failed'
      };
    }

    clearConfigCache(userId);

    const ssId = user.spreadsheetId || cleanedConfig.spreadsheetId;
    if (ssId) {
      try {
        CacheService.getScriptCache().remove(`sa_validation_${ssId}`);
      } catch (cacheRemoveError) {
        console.warn('saveUserConfig: Failed to clear SA validation cache:', cacheRemoveError.message);
      }
    }

    return {
      success: true,
      message: 'Config saved successfully',
      data: cleanedConfig,
      config: cleanedConfig, // フロントエンド互換性のため
      etag: cleanedConfig.etag, // 楽観的ロック用
      userId
    };

  } catch (error) {
    logError_('saveUserConfig', error, {
      userId: userId ? `${userId.substring(0, 8)}***` : 'N/A'
    });
    return {
      success: false,
      message: error && error.message ? `Config save error: ${error.message}` : 'Config save error: 詳細不明'
    };
  }
}

/**
 * 共通フィールドクリーンアップ - main.gs内の個別delete操作を統一
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - クリーンアップオプション
 * @returns {Object} クリーンアップ済み設定
 */
function cleanConfigFields(config, options = {}) {
  if (!config || typeof config !== 'object') {
    return config;
  }

  const cleanedConfig = { ...config };

  const fieldsToRemove = [
    'setupComplete',  // レガシーフィールド
    'isDraft',        // レガシーフィールド
    'questionText'    // 動的生成されるフィールド
  ];

  fieldsToRemove.forEach(field => {
    if (field in cleanedConfig) {
      delete cleanedConfig[field];
    }
  });

  if (options.updateAccessTime !== false) {
    cleanedConfig.lastAccessedAt = new Date().toISOString();
  }

  return cleanedConfig;
}
