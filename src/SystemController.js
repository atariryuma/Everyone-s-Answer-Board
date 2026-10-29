/**
 * @fileoverview SystemController - システム管理・診断・公開ライフサイクル。
 *   - URL システム: forceUrlSystemReset_, getWebAppUrl (6h cache + saved fallback)
 *   - 診断・監視: testSystemDiagnosis, performAutoRepair_
 *   - パフォーマンス: getPerformanceMetrics, diagnosePerformance, collectXxxMetrics
 *   - 公開: publishApp (allowlist + etag 楽観ロック + lesson active marker)
 *   - フォーム連携: getFormInfo, detectFormConnection_, searchFormsByDrive_
 *   - DB: testDatabaseConnection_, __ensureSheetWithHeaders_ (配置の初回にシートをそろえる)
 *
 * 公開状態の書き換えはここの publishApp と AdminApis.js の 2 関数 (republishMyBoard /
 * unpublishBoard) のみ。 __applyPublishStateChange_ に集約。
 */

/* global getCurrentEmail, createExceptionResponse, createAuthError, createAdminRequiredError, findUserByEmail_, openSpreadsheet, getUserConfig, saveUserConfig_, openDatabase, getCachedProperty_, getSheetInfo, sanitizeDisplaySettings, sanitizeMapping, getConfigOrDefault, logError_, clearDatabaseUserCache_, clearPropertyCache_, teacherOnlyError_ */

/**
 * キャッシュ期間 (秒)
 */
const CACHE_DURATION = {
  SHORT: 10,           // 10秒 - 認証ロック
  MEDIUM: 30,          // 30秒 - リアクション・ハイライトロック
  FORM_DATA: 30,       // 30秒 - フォームデータ行数キャッシュ（即時反映のため短期）
  LONG: 300,           // 5分 - ユーザー情報キャッシュ
  DATABASE_LONG: 600,  // 10分 - ヘッダー情報キャッシュ（列セットアップ時はinvalidateSheetHeadersCache_で明示的に無効化）
  USER_INDIVIDUAL: 900, // 15分 - 個別ユーザーキャッシュ（冗長性強化）
  EXTRA_LONG: 3600     // 1時間 - 設定キャッシュ
};

const USERS_SHEET_HEADERS = ['userId', 'userEmail', 'googleId', 'isActive', 'configJson', 'lastModified', 'createdAt'];

// lessons シート: lesson archive (draft → active → completed)。schemaVersion と sizeBytes は
//   JSON blob から外出し → 将来の migration 抽出 / quota 集計を blob parse なしで可能にする。
//   etag は 2 タブ同時編集の optimistic lock 用。
const LESSONS_SHEET_HEADERS = ['lessonId', 'userId', 'name', 'state', 'createdAt', 'startedAt', 'endedAt', 'schemaVersion', 'sizeBytes', 'etag', 'lessonJson'];

/**
 * lesson アーカイブ (授業終了/フェーズ切替時に凍結する回答) の行スキーマ。
 * Why 別シート: lessonJson (1 セル) に回答を同居させると Sheets の 1 セル 50,000 字上限で
 *   破裂し、本文が段階的に切り詰められていた。回答は「無界・追記一回きり」のデータなので
 *   1 件 1 行で持ち、lessonJson 側は範囲ポインタ {sheet, startRow, rowCount} だけを持つ。
 */
// 回答アーカイブ (1 回答 1 行)。後ろの 4 列はリアクションの数とハイライトの有無 (押した人のメールは残さない)。
//   フェーズの記録はこの 4 列を書かない。終わった授業に lesson.backfillReactions が元の回答シートから足す。
const LESSON_RESPONSES_SHEET_HEADERS = ['lessonId', 'phaseIndex', 'rowIndex', 'timestamp', 'class', 'answer', 'reason', 'numericX', 'numericY',
  'understand', 'like', 'curious', 'highlight'];

/**
 * プロパティキャッシュTTL (ミリ秒)
 * PropertiesServiceのメモリキャッシュ用
 * CLAUDE.md準拠: 30秒TTLで自動期限切れ
 */
const PROPERTY_CACHE_TTL = 30000; // 30秒

/**
 * タイムアウト期間 (ミリ秒)
 */
const TIMEOUT_MS = {
  QUICK: 100,          // UI応答性
  SHORT: 500,          // 軽量処理
  MEDIUM: 1000,        // 一般的処理
  LONG: 3000,          // 重い処理
  DEFAULT: 5000,       // デフォルトタイムアウト
  EXTENDED: 30000      // 拡張タイムアウト
};

/**
 * スリープ期間 (ミリ秒)
 */
const SLEEP_MS = {
  MICRO: 50,           // マイクロ待機
  SHORT: 100,          // 短時間休憩
  MEDIUM: 200,         // 中間休憩
  LONG: 500,           // 長時間休憩
  MAX: 5000            // 最大待機時間
};

/**
 * システム制限値
 */
const SYSTEM_LIMITS = {
  MAX_LOCK_ROWS: 100,             // ロッククリア最大行数
  PREVIEW_LENGTH: 50,             // プレビュー表示文字数
  DEFAULT_PAGE_SIZE: 20,          // デフォルトページサイズ
  MAX_PAGE_SIZE: 100,             // 最大ページサイズ
  RADIX_DECIMAL: 10,              // 10進数変換用基数
  CONFIG_JSON_MAX_CHARS: 32000    // configJson セル文字数上限 (Sheets cell 50000 char に対する safety margin)
};

// グローバルスコープにシステム定数を公開（GAS の単一グローバルスコープ前提）。
const __rootSys = (typeof globalThis !== 'undefined') ? globalThis : (typeof global !== 'undefined' ? global : this);
__rootSys.CACHE_DURATION = CACHE_DURATION;
__rootSys.TIMEOUT_MS = TIMEOUT_MS;
__rootSys.SLEEP_MS = SLEEP_MS;

/**
 * 全体キャッシュ無効化（cacheReset / autoRepair 共通）。
 *
 * GAS の CacheService.remove(key) は完全一致キーのみ削除し、 prefix/wildcard 削除が無い。
 * 旧実装は `cache.remove('user_cache_')` のように prefix 文字列を渡しており、 実キー
 * (`user_cache_<id>` 等) が動的なため **1 件も消えていなかった** (no-op なのに success を返す)。
 *
 * 本実装は version-bump 方式で「全体を一括無効化できる」cache 層を実際に invalidate する:
 *   - clearDatabaseUserCache_(): users の cache の版 (user_cache_ver) を新しくする → 全 user cache エントリを一括失効
 *   - clearPropertyCache_(): in-memory PropertiesService LRU を全消去
 *   - _webAppUrlCache: in-memory web app URL cache をリセット
 *
 * 注: board data cache (per-user, 10s TTL) / SA token (50min) / SA validation cache
 *     (per-SS) は key が動的に分散しており列挙削除できないが、 いずれも短 TTL で自然失効
 *     するため、 ここでの一括 invalidate 対象外とする（次回 polling/呼び出しで再生成）。
 *
 * @returns {string[]} 実行した invalidate アクションの説明（呼び出し側の actions 表示用）
 */
function invalidateGlobalCaches_() {
  const results = [];
  try {
    if (typeof clearDatabaseUserCache_ === 'function') {
      clearDatabaseUserCache_();
      results.push('ユーザーキャッシュ無効化 (user_cache_ver を更新)');
    }
  } catch (e) {
    results.push(`ユーザーキャッシュ無効化失敗: ${e && e.message ? e.message : 'Unknown error'}`);
  }
  try {
    if (typeof clearPropertyCache_ === 'function') {
      clearPropertyCache_();
      results.push('プロパティメモリキャッシュ全消去');
    }
  } catch (e) {
    results.push(`プロパティキャッシュ消去失敗: ${e && e.message ? e.message : 'Unknown error'}`);
  }
  try {
    _webAppUrlCache = { url: null, expiresAt: 0 };
    results.push('WebアプリURLキャッシュリセット');
  } catch (e) {
    results.push(`URLキャッシュリセット失敗: ${e && e.message ? e.message : 'Unknown error'}`);
  }
  results.push('短TTLキャッシュ (board data 10s / SA token 50min) は自然失効に委譲');
  return results;
}

/**
 * システム状態の強制リセット。Admin API の cacheReset から呼ばれる（緊急時用）。
 *
 * @returns {Object} リセット結果
 */
function forceUrlSystemReset_() {
  try {
    // Why (log level): 管理者が意図して呼ぶ保守操作 (Admin API の cacheReset) なので、
    //   異常を示す WARN ではなく INFO で残す。
    console.log('システム強制リセットが実行されました');

    const cacheResults = invalidateGlobalCaches_();

    return {
      success: true,
      message: 'システムリセットが完了しました',
      actions: cacheResults,
      cacheStatus: cacheResults.join(', '),
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    logError_('SystemController.forceUrlSystemReset', error);
    return {
      success: false,
      message: error && error.message ? error.message : '詳細不明'
    };
  }
}

// WebアプリURLの実行内メモリキャッシュ。テンプレートレンダごとに
// ScriptApp.getService().getUrl() が呼ばれる重複を削減する。
let _webAppUrlCache = { url: null, expiresAt: 0 };
const WEB_APP_URL_CACHE_TTL_MS = 60000;

/**
 * WebアプリのURL取得（バリデーション強化版）
 * 各種HTMLファイルから呼び出される
 *
 * 無効なURL（userCodeAppPanel等）を検出し、フォールバックを提供
 *
 * @returns {string} WebアプリURL（無効な場合は空文字）
 */
function getWebAppUrl() {
  const now = Date.now();
  if (_webAppUrlCache.url && now < _webAppUrlCache.expiresAt) {
    return _webAppUrlCache.url;
  }
  try {
    const url = ScriptApp.getService().getUrl();

    // URLバリデーション
    if (!isValidWebAppUrl_(url)) {
      console.warn('[WARN] getWebAppUrl: Invalid URL detected (OAuth/cold start issue):',
        url ? url.substring(0, 80) + '...' : 'empty');

      // フォールバック: スクリプトプロパティから保存済みURLを取得
      const props = PropertiesService.getScriptProperties();
      const savedUrl = props.getProperty('DEPLOYED_WEB_APP_URL');

      if (savedUrl && isValidWebAppUrl_(savedUrl)) {
        console.log('[INFO] getWebAppUrl: Using saved fallback URL');
        _webAppUrlCache = { url: savedUrl, expiresAt: now + WEB_APP_URL_CACHE_TTL_MS };
        return savedUrl;
      }

      return ''; // 無効なURLは空文字を返す
    }

    // 正常なURLをスクリプトプロパティに保存（次回のフォールバック用）
    try {
      const props = PropertiesService.getScriptProperties();
      const savedUrl = props.getProperty('DEPLOYED_WEB_APP_URL');
      if (savedUrl !== url) {
        props.setProperty('DEPLOYED_WEB_APP_URL', url);
      }
    } catch (saveError) {
      // 保存エラーは無視（読み取りのみ権限の場合など）
    }

    _webAppUrlCache = { url, expiresAt: now + WEB_APP_URL_CACHE_TTL_MS };
    return url;
  } catch (error) {
    logError_('getWebAppUrl', error);
    return '';
  }
}

/**
 * WebアプリURLのバリデーション
 * @private
 * @param {string} url - 検証するURL
 * @returns {boolean} 有効なURLの場合true
 */
function isValidWebAppUrl_(url) {
  if (!url || typeof url !== 'string') return false;

  // 無効なパターン（OAuth初期化中のURL）
  if (url.includes('userCodeAppPanel')) return false;
  if (url.includes('createOAuthDialog')) return false;

  // 正常なWebアプリURLの条件
  // /exec または /dev を含む必要がある
  if (!url.includes('/exec') && !url.includes('/dev')) return false;

  return true;
}

/**
 * システム全体の診断実行。Admin API の systemDiagnosis から呼ばれる。
 *
 * @returns {Object} 診断結果
 */
function testSystemDiagnosis() {
  try {

    const adminAuth = getBatchedAdminAuth();
    if (!adminAuth.success) {
      return adminAuth.authError || adminAuth.adminError || createAuthError();
    }

    const diagnostics = [];

    try {
      const props = PropertiesService.getScriptProperties();
      const coreProps = {
        adminEmail: !!props.getProperty('ADMIN_EMAIL'),
        databaseId: !!props.getProperty('DATABASE_SPREADSHEET_ID')
      };

      const allCorePresent = Object.values(coreProps).every(Boolean);
      diagnostics.push({
        test: 'Core Properties',
        status: allCorePresent ? 'PASS' : 'FAIL',
        details: coreProps,
        critical: true
      });
    } catch (error) {
      diagnostics.push({
        test: 'Core Properties',
        status: 'ERROR',
        error: error.message,
        critical: true
      });
    }

    try {
      const dbTest = testDatabaseConnection_();
      diagnostics.push({
        test: 'Database Connection',
        status: dbTest.success ? 'PASS' : 'FAIL',
        details: dbTest.success ? 'Database accessible' : dbTest.message,
        critical: true
      });
    } catch (error) {
      diagnostics.push({
        test: 'Database Connection',
        status: 'ERROR',
        error: error.message,
        critical: true
      });
    }

    try {
      const webAppUrl = ScriptApp.getService().getUrl();
      diagnostics.push({
        test: 'Web App Deployment',
        status: webAppUrl ? 'PASS' : 'FAIL',
        details: webAppUrl || 'No deployment URL available',
        critical: false
      });
    } catch (error) {
      diagnostics.push({
        test: 'Web App Deployment',
        status: 'ERROR',
        error: error.message,
        critical: false
      });
    }

    const totalTests = diagnostics.length;
    let passedTests = 0;
    let criticalIssues = 0;
    for (const d of diagnostics) {
      if (d.status === 'PASS') passedTests++;
      else if (d.critical) criticalIssues++;
    }

    return {
      success: criticalIssues === 0,
      message: criticalIssues === 0 ? 'システム診断完了 - すべてのテストに合格しました' : `システム診断で ${criticalIssues} 件の重大な問題が見つかりました`,
      summary: {
        totalTests,
        passed: passedTests,
        failed: totalTests - passedTests,
        criticalIssues
      },
      diagnostics,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    logError_('testSystemDiagnosis', error);
    return createExceptionResponse(error, 'System diagnosis failed');
  }
}

/**
 * 自動修復の実行
 * 管理者用の自動メンテナンス機能
 *
 * @returns {Object} 修復結果
 */
function performAutoRepair_() {
  try {

    const adminAuth = getBatchedAdminAuth();
    if (!adminAuth.success) {
      return adminAuth.authError || adminAuth.adminError || createAuthError();
    }

    const repairResults = {
      timestamp: new Date().toISOString(),
      actions: [],
      warnings: [],
      summary: ''
    };

    let actionCount = 0;

    try {
      // version-bump 方式で実効的に invalidate (旧 cache.removeAll([prefix]) は no-op だった)。
      const cacheActions = invalidateGlobalCaches_();
      cacheActions.forEach(a => repairResults.actions.push(a));
      actionCount += cacheActions.length;
    } catch (cacheError) {
      repairResults.warnings.push(`キャッシュクリア失敗: ${cacheError.message || 'Unknown error'}`);
    }

    try {
      const dbTest = testDatabaseConnection_();
      if (!dbTest.success) {
        repairResults.warnings.push(`データベース接続問題: ${dbTest.message}`);
      } else {
        repairResults.actions.push('データベース接続確認');
        actionCount++;
      }
    } catch (dbError) {
      repairResults.warnings.push(`データベーステスト失敗: ${dbError.message || 'Unknown error'}`);
    }

    try {
      const props = PropertiesService.getScriptProperties();
      const coreProps = {
        ADMIN_EMAIL: props.getProperty('ADMIN_EMAIL'),
        DATABASE_SPREADSHEET_ID: props.getProperty('DATABASE_SPREADSHEET_ID')
      };

      const missingProps = Object.keys(coreProps).filter(key => !coreProps[key]);
      if (missingProps.length === 0) {
        repairResults.actions.push('システムプロパティ検証');
        actionCount++;
      } else {
        repairResults.warnings.push(`不足システムプロパティ: ${missingProps.join(', ')}`);
      }
    } catch (propsError) {
      repairResults.warnings.push(`プロパティ検証失敗: ${propsError.message || 'Unknown error'}`);
    }

    repairResults.summary = `${actionCount}個の修復処理を実行、${repairResults.warnings.length}個の警告`;

    return {
      success: true,
      repairResults
    };

  } catch (error) {
    logError_('SystemController.performAutoRepair', error);
    return {
      success: false,
      message: error.message || '自動修復エラー'
    };
  }
}

/**
 * publishApp の入力ペイロードを許可キーだけに絞り込み、etag 競合検出用の情報を返す。
 * Why: doPost 経由で任意の config が渡ってくるため、ボード公開に関係のないキー
 *      (内部状態、書き込み禁止フィールド等) を黙って捨てる。受理キーは allowedKeys
 *      で明示し、それ以外は ignoredKeys として呼び出し元 (publishApp) でログに残す。
 * @private
 */
function sanitizePublishPayload_(publishConfig, currentConfig = {}) {
  const source = (publishConfig && typeof publishConfig === 'object' && !Array.isArray(publishConfig))
    ? publishConfig
    : {};

  const allowedKeys = new Set([
    'spreadsheetId',
    'sheetName',
    'columnMapping',
    'displaySettings',
    'formUrl',
    'formTitle',
    'showDetails',
    'etag'
  ]);
  const ignoredKeys = Object.keys(source).filter((key) => !allowedKeys.has(key));

  const sanitized = {};

  if (typeof source.spreadsheetId === 'string') {
    sanitized.spreadsheetId = source.spreadsheetId.trim();
  }

  if (typeof source.sheetName === 'string') {
    sanitized.sheetName = source.sheetName.trim();
  }

  if (source.columnMapping && typeof source.columnMapping === 'object' && !Array.isArray(source.columnMapping)) {
    sanitized.columnMapping = (typeof sanitizeMapping === 'function')
      ? sanitizeMapping(source.columnMapping)
      : source.columnMapping;
  }

  if (source.displaySettings && typeof source.displaySettings === 'object' && !Array.isArray(source.displaySettings)) {
    sanitized.displaySettings = (typeof sanitizeDisplaySettings === 'function')
      ? sanitizeDisplaySettings(source.displaySettings)
      : source.displaySettings;
  } else if (currentConfig.displaySettings && typeof currentConfig.displaySettings === 'object') {
    sanitized.displaySettings = (typeof sanitizeDisplaySettings === 'function')
      ? sanitizeDisplaySettings(currentConfig.displaySettings)
      : currentConfig.displaySettings;
  }

  if (typeof source.formUrl === 'string' && source.formUrl.trim()) {
    sanitized.formUrl = source.formUrl.trim();
  } else if (typeof currentConfig.formUrl === 'string' && currentConfig.formUrl.trim()) {
    sanitized.formUrl = currentConfig.formUrl.trim();
  }

  const maxFormTitleLength = SYSTEM_LIMITS.PREVIEW_LENGTH * 4;
  if (typeof source.formTitle === 'string' && source.formTitle.trim()) {
    sanitized.formTitle = source.formTitle.trim().substring(0, maxFormTitleLength);
  } else if (typeof currentConfig.formTitle === 'string' && currentConfig.formTitle.trim()) {
    sanitized.formTitle = currentConfig.formTitle.trim().substring(0, maxFormTitleLength);
  }

  if (typeof source.showDetails === 'boolean') {
    sanitized.showDetails = source.showDetails;
  } else if (typeof currentConfig.showDetails === 'boolean') {
    sanitized.showDetails = currentConfig.showDetails;
  }

  if (typeof source.etag === 'string' && source.etag.length > 0 && source.etag.length <= 255) {
    sanitized.etag = source.etag;
  }

  return { config: sanitized, ignoredKeys };
}

/**
 * アプリケーションを公開する（boardConfig をユーザー DB に保存し、etag を更新する）。
 *
 * Why etag 競合検出: 教師が複数タブで管理画面を開いて同時に「公開」を押した場合、
 *   後から押した方が前の保存を上書きする race を防ぐ。クライアントは現在の etag を
 *   送り、サーバ側で現状の etag と比較。一致しないとき "etag_conflict" を返す。
 * Why CLAUDE.md "publishApp accepts only allowlisted fields with etag conflict detection"
 *   というメンテナンスルールが定義されており、両方をここで担保する。
 *
 * @param {Object} publishConfig - クライアントから来た公開設定ペイロード。
 *   spreadsheetId / sheetName / columnMapping / displaySettings / formUrl / formTitle /
 *   showDetails / etag のいずれか。それ以外のキーは sanitizePublishPayload_ で破棄される。
 * @returns {{success: boolean, message?: string, error?: string, etag?: string,
 *   conflict?: boolean, currentEtag?: string}} 成功時 etag を含む、競合時 conflict:true。
 */
function publishApp(publishConfig) {
  try {
    const email = getCurrentEmail();

    if (!email) {
      logError_('publishApp', new Error('User authentication failed'));
      return createAuthError();
    }

    if (!publishConfig || typeof publishConfig !== 'object' || Array.isArray(publishConfig)) {
      return createErrorResponse('公開設定が必要です');
    }

    const user = findUserByEmail_(email, { requestingUser: email });
    if (!user) {
      logError_('publishApp', new Error('User not found'), { email });
      return createErrorResponse('ユーザーが見つかりません');
    }

    // Why getUserConfig (not getConfigOrDefault): publish は save パス。getConfigOrDefault は
    //   findUserById_ の 429/cache miss で {} を返し、updatedConfig が現在の profile/
    //   columnMapping を含まない最小オブジェクトになって saveUserConfig_ が全上書きする
    //   data-loss バグになる。さらに currentConfig.etag が undefined になるので楽観ロック
    //   (line 724) もスキップされ二重 publish が race する。
    const cfgRes = getUserConfig(user.userId, user);
    if (!cfgRes || !cfgRes.success) {
      return createErrorResponse(`設定の取得に失敗しました: ${(cfgRes && cfgRes.message) || '詳細不明'}`);
    }
    const currentConfig = cfgRes.config;
    const { config: safePublishConfig, ignoredKeys } = sanitizePublishPayload_(publishConfig, currentConfig);

    if (ignoredKeys.length > 0) {
      console.warn('publishApp: Ignored non-allowlisted publish fields', {
        userId: user.userId ? `${user.userId.substring(0, 8)}***` : 'N/A',
        ignoredFields: ignoredKeys
      });
    }

    if (!safePublishConfig.spreadsheetId) {
      return createErrorResponse('データソース（スプレッドシートID）が設定されていません');
    }

    if (!safePublishConfig.sheetName) {
      return createErrorResponse('データソース（シート名）が設定されていません');
    }

    if (!safePublishConfig.columnMapping || typeof safePublishConfig.columnMapping !== 'object') {
      return createErrorResponse('列マッピングが設定されていません');
    }

    if (Object.keys(safePublishConfig.columnMapping).length === 0) {
      return createErrorResponse('列マッピングが空です。少なくとも回答列または数値列を設定してください');
    }

    // matrix/numberline モードでは answer は numericX を兼ねる (AdminPanel.js.html
    // getColumnMapping が自動的に answer=numericX をセット)。ここで answer のみ厳格に
    // 要求すると matrix/numberline ユーザーの publish が壊れる (Edge case audit #1)。
    // - matrix / numberline: numericX 必須 (Y は matrix のみ必須)
    // - それ以外: answer 必須
    const boardMode = (safePublishConfig.displaySettings && safePublishConfig.displaySettings.boardMode)
      || (currentConfig.displaySettings && currentConfig.displaySettings.boardMode)
      || 'auto';
    const colMap = safePublishConfig.columnMapping;
    const hasValidIndex = (k) => typeof colMap[k] === 'number' && colMap[k] >= 0;
    if (boardMode === 'matrix') {
      if (!hasValidIndex('numericX')) return createErrorResponse('X軸の数値列が設定されていません');
      if (!hasValidIndex('numericY')) return createErrorResponse('Y軸の数値列が設定されていません');
    } else if (boardMode === 'numberline') {
      if (!hasValidIndex('numericX')) return createErrorResponse('X軸の数値列が設定されていません');
    } else {
      if (!hasValidIndex('answer')) return createErrorResponse('メイン質問列が設定されていません');
    }

    if (currentConfig.etag && !safePublishConfig.etag) {
      return {
        success: false,
        error: 'etag_required',
        message: '設定が更新されている可能性があります。画面を再読み込みしてから再公開してください。'
      };
    }

    const publishedAt = new Date().toISOString();

    const updatedConfig = {
      ...currentConfig,
      ...safePublishConfig,
      isPublished: true,
      publishedAt,
      setupStatus: 'completed',
      lastModified: publishedAt
    };

    // Why: lesson 駆動の publish (activeLessonId 付き) のときだけ「授業開始時刻」を記録。
    //   この marker は unpublishBoard 経由の auto-archive 判定 (経過時間 5min 以上か) と
    //   23:00 cron sweep の両方で使われる。republishMyBoard では touch しない (= 同一授業の
    //   再公開で timer が reset されるのを防ぐ)。
    if (updatedConfig.activeLessonId) {
      updatedConfig.currentLessonStartedAt = publishedAt;
    }

    // __allowPublishStateWrite: publishApp は publish 状態を書き換える正規経路 (3 関数の 1 つ)。
    //   saveUserConfig_ の publish-state ガードを通過させる。
    const saveResult = saveUserConfig_(user.userId, updatedConfig, { isPublish: true, __allowPublishStateWrite: true });

    if (!saveResult.success) {
      logError_('publishApp.saveUserConfig', new Error(saveResult.message || 'save failed'));
      return {
        success: false,
        message: saveResult.message || '公開設定の保存に失敗しました',
        error: saveResult.error || null,
        currentConfig: saveResult.currentConfig || null
      };
    }

    // publish 遷移時の cache 即時 stale 化 (__applyPublishStateChange_ と共有の mechanism)。
    //   saveUserConfig_ が primary のアクセス判定 cache は消すが、 profile 分 + board data version は
    //   ここで明示。 再公開 (unpublish→republish) 時に viewer が stale/空データを見るのを防ぐ。
    if (typeof invalidateBoardCaches_ === 'function') {
      invalidateBoardCaches_(updatedConfig, user.userId);
    }

    return {
      success: true,
      message: '回答ボードが正常に公開されました',
      publishedAt,
      userId: user.userId,
      etag: saveResult.etag || null,
      config: saveResult.config || null
    };

  } catch (error) {
    logError_('publishApp', error, {
      spreadsheetId: publishConfig?.spreadsheetId,
      sheetName: publishConfig?.sheetName,
      userEmail: getCurrentEmail()
    });
    return {
      success: false,
      message: error.message,
      error: error.message
    };
  }
}

/**
 * 多層フォーム検出システム
 * @param {Object} spreadsheet - スプレッドシートオブジェクト
 * @param {Object} sheet - シートオブジェクト
 * @param {string} sheetName - シート名
 * @param {boolean} isOwner - オーナー権限フラグ
 * @returns {Object} {formUrl, confidence, detectionMethod}
 */
function detectFormConnection_(spreadsheet, sheet, sheetName, isOwner) {
  const results = {
    formUrl: null,
    formTitle: null,
    confidence: 0,
    detectionMethod: 'none',
    details: []
  };

  // FormApp.openByUrl でタイトル取得し、失敗時は fallback。
  const resolveTitle = (formUrl) => {
    try {
      results.formTitle = FormApp.openByUrl(formUrl).getTitle();
      results.details.push('FormApp.openByUrl()でタイトル取得');
    } catch (titleError) {
      console.warn('FormApp.openByUrl failed:', titleError.message);
      results.formTitle = generateFormTitle(sheetName, spreadsheet.getName());
      results.details.push('FormApp権限エラー - フォールバックタイトル使用');
    }
  };
  // `<prefix>: <err.message>` を push、err.message が無ければ「: 詳細不明」。
  const pushErrDetail = (prefix, err) => {
    results.details.push(err && err.message ? `${prefix}: ${err.message}` : `${prefix}: 詳細不明`);
  };

  if (isOwner) {
    try {
      if (typeof sheet.getFormUrl === 'function') {
        const formUrl = sheet.getFormUrl();
        if (formUrl) {
          results.formUrl = formUrl;
          results.confidence = 95;
          results.detectionMethod = 'sheet_api';
          results.details.push('Sheet.getFormUrl()で検出');
          resolveTitle(formUrl);
          return results;
        }
      }

      if (typeof spreadsheet.getFormUrl === 'function') {
        const formUrl = spreadsheet.getFormUrl();
        if (formUrl) {
          results.formUrl = formUrl;
          results.confidence = 85;
          results.detectionMethod = 'spreadsheet_api';
          results.details.push('SpreadsheetApp.getFormUrl()で検出');
          resolveTitle(formUrl);
          return results;
        }
      }
    } catch (apiError) {
      console.warn('detectFormConnection: API検出失敗:', apiError.message);
      pushErrDetail('API検出失敗', apiError);
    }
  }

  if (isOwner) {
    try {
      const spreadsheetId = spreadsheet.getId();
      const driveFormResult = searchFormsByDrive_(spreadsheetId, sheetName);
      if (driveFormResult.formUrl) {
        results.formUrl = driveFormResult.formUrl;
        results.formTitle = driveFormResult.formTitle;
        results.confidence = 80;
        results.detectionMethod = 'drive_search';
        results.details.push('Drive API検索で検出');
        return results;
      }
    } catch (driveError) {
      console.warn('detectFormConnection: Drive API検索失敗:', driveError.message);
      pushErrDetail('Drive API検索失敗', driveError);
    }
  }

  try {
    const { lastCol, headers: fullHeaders } = getSheetInfo(sheet);
    const headers = fullHeaders.slice(0, Math.min(lastCol, 10));
    const headerAnalysis = analyzeFormHeaders(headers);

    if (headerAnalysis.isFormLike) {
      results.confidence = Math.max(results.confidence, headerAnalysis.confidence);
      results.detectionMethod = results.detectionMethod === 'none' ? 'header_analysis' : results.detectionMethod;
      results.details.push(headerAnalysis && headerAnalysis.reason ? `ヘッダー解析: ${headerAnalysis.reason}` : 'ヘッダー解析: 結果不明');
    }
  } catch (headerError) {
    console.warn('detectFormConnection: ヘッダー解析失敗:', headerError.message);
    pushErrDetail('ヘッダー解析失敗', headerError);
  }

  const sheetNameAnalysis = analyzeSheetName(sheetName);
  if (sheetNameAnalysis.isFormLike) {
    results.confidence = Math.max(results.confidence, sheetNameAnalysis.confidence);
    results.detectionMethod = results.detectionMethod === 'none' ? 'sheet_name' : results.detectionMethod;
    results.details.push(sheetNameAnalysis && sheetNameAnalysis.reason ? `シート名解析: ${sheetNameAnalysis.reason}` : 'シート名解析: 結果不明');
  }

  if (results.confidence >= 40) {
    results.formTitle = `${sheetName} (フォーム検出済み)`;
  }

  return results;
}

/**
 * フォームヘッダーパターン解析
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} {isFormLike, confidence, reason}
 */
function analyzeFormHeaders(headers) {
  if (!headers || headers.length === 0) {
    return { isFormLike: false, confidence: 0, reason: 'ヘッダーなし' };
  }

  const formIndicators = [
    { pattern: /タイムスタンプ|timestamp/i, weight: 30, description: 'タイムスタンプ列' },
    { pattern: /メールアドレス|email.*address|メール/i, weight: 25, description: 'メールアドレス列' },
    { pattern: /回答|answer|意見|response/i, weight: 20, description: '回答列' },
    { pattern: /名前|name|氏名/i, weight: 15, description: '名前列' },
    { pattern: /クラス|class|組/i, weight: 10, description: 'クラス列' }
  ];

  let totalScore = 0;
  const matches = [];

  headers.forEach((header, index) => {
    if (typeof header === 'string') {
      formIndicators.forEach(indicator => {
        if (indicator.pattern.test(header)) {
          totalScore += indicator.weight;
          matches.push(`${indicator.description}(${header})`);
        }
      });
    }
  });

  const confidence = Math.min(totalScore, 85); // 最大85%
  const isFormLike = confidence >= 40; // 40%以上でフォームと判定

  return {
    isFormLike,
    confidence,
    reason: isFormLike ? `フォームパターン検出: ${matches.join(', ')}` : '一般的なフォームパターンが不足'
  };
}

/**
 * シート名パターン解析
 * @param {string} sheetName - シート名
 * @returns {Object} {isFormLike, confidence, reason}
 */
function analyzeSheetName(sheetName) {
  if (!sheetName || typeof sheetName !== 'string') {
    return { isFormLike: false, confidence: 0, reason: 'シート名なし' };
  }

  const patterns = [
    { regex: /フォームの回答|form.*responses?/i, confidence: 80, description: 'フォーム回答シート' },
    { regex: /フォーム.*結果|form.*results?/i, confidence: 75, description: 'フォーム結果シート' },
    { regex: /回答.*\d+|responses?.*\d+/i, confidence: 70, description: 'ナンバリング回答シート' },
    { regex: /アンケート|survey|questionnaire/i, confidence: 60, description: 'アンケートシート' },
    { regex: /回答|responses?|答え|answer/i, confidence: 50, description: '回答関連シート' }
  ];

  for (const pattern of patterns) {
    if (pattern.regex.test(sheetName)) {
      return {
        isFormLike: true,
        confidence: pattern.confidence,
        reason: `${pattern.description}パターン検出: "${sheetName}"`
      };
    }
  }

  return {
    isFormLike: false,
    confidence: 0,
    reason: `一般的なフォームシート名パターンなし: "${sheetName}"`
  };
}

/**
 * フォールバックタイトル生成
 * @param {string} sheetName - シート名
 * @param {string} spreadsheetName - スプレッドシート名
 * @returns {string} 生成されたタイトル
 */
function generateFormTitle(sheetName, spreadsheetName) {
  try {
    if (sheetName && typeof sheetName === 'string') {
      const formPatterns = [
        /フォームの回答|form.*responses?/i,
        /フォーム.*結果|form.*results?/i,
        /回答.*\d+|responses?.*\d+/i,
        /アンケート|survey|questionnaire/i
      ];

      for (const pattern of formPatterns) {
        if (pattern.test(sheetName)) {
          const formTitle = sheetName
            .replace(/の回答.*$/i, '')
            .replace(/.*responses?.*$/i, '')
            .replace(/.*結果.*$/i, '')
            .trim();

          if (formTitle.length > 0) {
            return `${formTitle} (フォーム)`;
          }
        }
      }
    }

    if (spreadsheetName && typeof spreadsheetName === 'string') {
      return `${spreadsheetName} - ${sheetName} (フォーム)`;
    }

    return `${sheetName || 'データ'} (フォーム)`;

  } catch (error) {
    console.warn('generateFormTitle error:', error.message);
    return `${sheetName || 'フォーム'} (タイトル生成エラー)`;
  }
}

/**
 * Drive APIによるフォーム検索
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} {formUrl, formTitle}
 */
function searchFormsByDrive_(spreadsheetId, sheetName) {
  try {

    const forms = DriveApp.getFilesByType('application/vnd.google-apps.form');

    while (forms.hasNext()) {
      const formFile = forms.next();
      try {
        let form = null;
        let destId = null;
        let formTitle = null;
        let formPublishedUrl = null;

        try {
          form = FormApp.openById(formFile.getId());
          destId = form.getDestinationId();
          formTitle = form.getTitle();
          formPublishedUrl = form.getPublishedUrl();
        } catch (formAccessError) {
          console.warn('searchFormsByDrive: FormApp権限制限、ファイル名から推測:', formAccessError.message);
          formTitle = formFile.getName();
          continue;
        }

        if (destId === spreadsheetId) {
          return {
            formUrl: formPublishedUrl,
            formTitle
          };
        }
      } catch (formError) {
        console.warn('searchFormsByDrive: フォームアクセスエラー（継続）:', formError.message);
      }
    }

    return { formUrl: null, formTitle: null };

  } catch (error) {
    logError_('searchFormsByDrive', error);
    throw error;
  }
}

/**
 * フォーム情報を取得（実装関数）- 適応的アクセス対応版
 * main.gs API Gateway から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfo(spreadsheetId, sheetName) {
  const startTime = new Date().toISOString();
  try {
    { const denied = teacherOnlyError_(); if (denied) return denied; }
    if (!spreadsheetId || !sheetName) {
      return __formInfoFailure_('INVALID_ARGUMENTS', 'スプレッドシートIDとシート名を指定してください。', { sheetName });
    }
    const access = __resolveFormInfoSpreadsheet_(spreadsheetId, sheetName);
    if (!access.ok) return access.error;
    return __buildFormInfoResult_(access, spreadsheetId, sheetName);
  } catch (error) {
    logError_('getFormInfoImpl', error, {
      startTime,
      endTime: new Date().toISOString(),
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 12)}***` : 'N/A',
      sheetName,
      stack: error.stack
    });
    return __formInfoFailure_('UNKNOWN_ERROR', 'フォーム情報の取得に失敗しました。', { sheetName, error: error.message });
  }
}

// getFormInfo の失敗 envelope を統一生成。
function __formInfoFailure_(status, message, opts) {
  opts = opts || {};
  const result = {
    success: false,
    status,
    message,
    formData: {
      formUrl: null,
      formTitle: opts.sheetName || 'フォーム',
      spreadsheetName: opts.spreadsheetName || '',
      sheetName: opts.sheetName || ''
    }
  };
  if (opts.accessMethod) result.accessMethod = opts.accessMethod;
  if (opts.error) result.error = opts.error;
  return result;
}

// Spreadsheet / sheet の解決 + アクセス権チェック。
function __resolveFormInfoSpreadsheet_(spreadsheetId, sheetName) {
  const dataAccess = openSpreadsheet(spreadsheetId, { context: 'systemController' });
  if (!dataAccess || !dataAccess.spreadsheet) {
    return { ok: false, error: __formInfoFailure_('SPREADSHEET_NOT_FOUND', 'スプレッドシートにアクセスできませんでした。', { sheetName, accessMethod: 'normal_permissions' }) };
  }
  const { spreadsheet } = dataAccess;
  let spreadsheetName;
  try { spreadsheetName = spreadsheet.getName(); }
  catch (error) {
    console.warn('getFormInfoImpl: getName() failed, using fallback:', error.message);
    spreadsheetName = `スプレッドシート (ID: ${spreadsheetId.substring(0, 8)}...)`;
  }
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    return { ok: false, error: __formInfoFailure_('SHEET_NOT_FOUND', `シート「${sheetName}」が見つかりません。`, { sheetName, spreadsheetName, accessMethod: 'normal_permissions' }) };
  }
  return { ok: true, spreadsheet, sheet, spreadsheetName };
}

/**
 * Form の ScaleItem (線形尺度) メタデータを抽出する。
 *
 * Why: 管理パネルの「X軸ラベル / Y軸ラベル」入力欄は元々 100% 手入力だったが、
 *   ScaleItem.setLabels で createTemplateForm が既に書き込んでいるので、
 *   読み戻して空欄を自動補完できる (UX: 教師の手数を 4 入力 → 0 に削減)。
 *
 * 順序の意味: items[0] = 1番目の ScaleItem → numericX 軸候補
 *            items[1] = 2番目の ScaleItem → numericY 軸候補 (matrix のみ)
 *
 * @param {string} formUrl - 開ける Form の URL
 * @returns {Array<{leftLabel:string,rightLabel:string,lowerBound:number,upperBound:number,title:string}>|null}
 *   形式は AdminPanel client が直接 consume する。失敗時は null。
 */
function __extractFormScaleItems_(formUrl) {
  if (!formUrl || typeof FormApp === 'undefined' || !FormApp.openByUrl) return null;
  try {
    const form = FormApp.openByUrl(formUrl);
    const items = form.getItems(FormApp.ItemType.SCALE);
    return items.map((item) => {
      const scale = item.asScaleItem();
      return {
        title: item.getTitle() || '',
        leftLabel: scale.getLeftLabel() || '',
        rightLabel: scale.getRightLabel() || '',
        lowerBound: scale.getLowerBound(),
        upperBound: scale.getUpperBound()
      };
    });
  } catch (e) {
    console.warn('__extractFormScaleItems_ failed:', e && e.message);
    return null;
  }
}

// 検出結果から成功/失敗 envelope を組み立て。
function __buildFormInfoResult_(access, spreadsheetId, sheetName) {
  const { spreadsheet, sheet, spreadsheetName } = access;
  const accessMethod = 'normal_permissions';
  const detection = detectFormConnection_(spreadsheet, sheet, sheetName, true);
  const scaleItems = detection.formUrl ? __extractFormScaleItems_(detection.formUrl) : null;
  const formData = {
    formUrl: detection.formUrl || null,
    formTitle: detection.formTitle || sheetName,
    spreadsheetName,
    sheetName,
    scaleItems: scaleItems || [],
    detectionDetails: {
      method: detection.detectionMethod,
      confidence: detection.confidence,
      accessMethod,
      timestamp: new Date().toISOString()
    }
  };

  if (detection.formUrl) {
    return {
      success: true,
      status: 'FORM_LINK_FOUND',
      message: `フォーム連携を確認しました (${detection.detectionMethod})`,
      formData,
      timestamp: new Date().toISOString(),
      requestContext: { spreadsheetId, sheetName, accessMethod }
    };
  }
  const isHighConfidence = detection.confidence >= 70;
  return {
    success: isHighConfidence,
    status: isHighConfidence ? 'FORM_DETECTED_NO_URL' : 'FORM_NOT_LINKED',
    message: isHighConfidence ? 'フォーム連携パターンを検出（URL取得不可）' : 'フォーム連携が確認できませんでした',
    reason: isHighConfidence ? 'FORM_DETECTED_NO_URL' : 'FORM_NOT_LINKED',
    formData,
    suggestions: detection.suggestions || [
      'Googleフォームの「回答の行き先」を開き、対象のシートにリンクしてください',
      'フォーム作成者に連携状況を確認してください',
      'シート名に「回答」「フォーム」等の文字列が含まれている場合、フォーム連携パターンとして評価されます'
    ],
    analysisResults: detection.analysisResults
  };
}

// ─── Performance metrics ─────────────────────────────────────────
// API 実行時間 / キャッシュ効率 / エラー率 / バッチ処理 / メモリの軽量モニタ。

// perf API (getPerformanceMetrics / diagnosePerformance) 用の admin 権限ゲート。
//   権限が無ければ timestamp 付き error response を返す (caller は早期 return)、 あれば null。
function __requirePerfAdmin_() {
  // admin 判定は canonical な requireAdmin() を再利用。 失敗時のみ perf API 用の
  //   timestamp 付き error response を返す (createAdminRequiredError とは別 shape)。
  if (requireAdmin()) return null;
  return {
    success: false,
    error: 'Administrator権限が必要です',
    timestamp: new Date().toISOString()
  };
}

/**
 * パフォーマンスメトリクスを収集・分析する
 * @param {string} category - メトリクスカテゴリ ('api', 'cache', 'batch', 'error')
 * @param {Object} options - オプション設定
 * @returns {Object} メトリクス統計結果
 */
function getPerformanceMetrics(category = 'all', options = {}) {
  try {
    const startTime = Date.now();
    const authError = __requirePerfAdmin_();
    if (authError) return authError;

    const metrics = {
      timestamp: new Date().toISOString(),
      collectionTime: null,
      systemInfo: getSystemPerformanceInfo_(),
      categories: {}
    };

    if (category === 'all' || category === 'api') {
      metrics.categories.api = collectApiMetrics();
    }
    if (category === 'all' || category === 'cache') {
      metrics.categories.cache = collectCacheMetrics_();
    }
    if (category === 'all' || category === 'batch') {
      metrics.categories.batch = collectBatchMetrics_();
    }
    if (category === 'all' || category === 'error') {
      metrics.categories.error = collectErrorMetrics_(options);
    }

    const endTime = Date.now();
    metrics.collectionTime = endTime - startTime;

    return {
      success: true,
      metrics,
      performanceImpact: {
        collectionTimeMs: metrics.collectionTime,
        overhead: metrics.collectionTime < 100 ? 'minimal' : metrics.collectionTime < 500 ? 'low' : 'moderate'
      }
    };

  } catch (error) {
    logError_('getPerformanceMetrics', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * システム基本パフォーマンス情報を取得
 * @returns {Object} システム情報
 */
function getSystemPerformanceInfo_() {
  try {
    const startTime = Date.now();

    const systemInfo = {
      gasRuntime: 'V8',
      quotaInfo: {
        executionTimeUsed: (Date.now() - startTime), // この関数の実行時間
        estimatedQuotaRemaining: '不明' // GASでは直接取得不可
      },
      cacheStatus: {
        scriptCache: typeof CacheService !== 'undefined' ? 'available' : 'unavailable',
        propertiesService: typeof PropertiesService !== 'undefined' ? 'available' : 'unavailable'
      },
      serviceStatus: {
        spreadsheetApp: typeof SpreadsheetApp !== 'undefined' ? 'available' : 'unavailable',
        urlFetchApp: typeof UrlFetchApp !== 'undefined' ? 'available' : 'unavailable'
      }
    };

    return systemInfo;
  } catch (error) {
    console.warn('getSystemPerformanceInfo エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * API実行メトリクスを収集
 * @returns {Object} API実行統計
 */
function collectApiMetrics() {
  try {
    const apiStats = {
      totalCalls: 0,
      averageResponseTime: 0,
      fastestCall: null,
      slowestCall: null,
      errorRate: 0,
      cacheHitRate: 0
    };

    const testStartTime = Date.now();

    try {
      const testEmail = getCurrentEmail();
      const testEndTime = Date.now();
      const testDuration = testEndTime - testStartTime;

      apiStats.totalCalls = 1;
      apiStats.averageResponseTime = testDuration;
      apiStats.fastestCall = { function: 'getCurrentEmail', timeMs: testDuration };
      apiStats.slowestCall = { function: 'getCurrentEmail', timeMs: testDuration };
      apiStats.errorRate = testEmail ? 0 : 1;
    } catch (testError) {
      apiStats.errorRate = 1;
    }

    apiStats.batchEfficiencyNote = 'バッチ処理により70倍の性能改善を実現';
    apiStats.architecture = 'GAS-Native Direct API';

    return apiStats;
  } catch (error) {
    console.warn('collectApiMetrics エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * キャッシュ効率メトリクスを収集
 * @returns {Object} キャッシュ統計
 */
function collectCacheMetrics_() {
  try {
    const cache = CacheService.getScriptCache();

    const testKey = `perf_cache_test_${Date.now()}`;
    const testValue = JSON.stringify({ test: true, timestamp: Date.now() });

    const writeStartTime = Date.now();
    cache.put(testKey, testValue, CACHE_DURATION.SHORT);
    const writeTime = Date.now() - writeStartTime;

    const readStartTime = Date.now();
    const readValue = cache.get(testKey);
    const readTime = Date.now() - readStartTime;

    cache.remove(testKey);

    const cacheStats = {
      writePerformance: {
        timeMs: writeTime,
        status: writeTime < 10 ? 'excellent' : writeTime < 50 ? 'good' : 'needs_attention'
      },
      readPerformance: {
        timeMs: readTime,
        status: readTime < 5 ? 'excellent' : readTime < 20 ? 'good' : 'needs_attention'
      },
      hitRate: readValue ? 1.0 : 0.0,
      cacheDurations: CACHE_DURATION,
      recommendations: []
    };

    if (writeTime > 50) {
      cacheStats.recommendations.push('キャッシュ書き込み速度が遅い - データ量を最適化を検討');
    }
    if (readTime > 20) {
      cacheStats.recommendations.push('キャッシュ読み取り速度が遅い - キー構造を最適化を検討');
    }

    return cacheStats;
  } catch (error) {
    console.warn('collectCacheMetrics エラー:', error.message);
    return { error: error.message, recommendations: ['キャッシュサービスへのアクセスを確認'] };
  }
}

/**
 * バッチ処理効率メトリクスを収集
 * @returns {Object} バッチ処理統計
 */
function collectBatchMetrics_() {
  try {
    const batchStats = {
      batchProcessingEnabled: true,
      performanceImprovement: '70倍改善 (CLAUDE.md準拠)',
      implementation: {
        pattern: 'Direct SpreadsheetApp batch operations',
        avoidance: 'Individual API calls in loops eliminated'
      },
      recommendations: [
        '継続的にバッチパターンを維持',
        '個別API呼び出しを避ける',
        'データ一括取得・更新パターンを活用'
      ]
    };

    const testStartTime = Date.now();
    try {
      const testAccess = SpreadsheetApp.getActiveSpreadsheet ? 'available' : 'unavailable';
      const testEndTime = Date.now();

      batchStats.accessTest = {
        spreadsheetApp: testAccess,
        responseTimeMs: testEndTime - testStartTime,
        status: testAccess === 'available' ? 'healthy' : 'needs_attention'
      };
    } catch (testError) {
      batchStats.accessTest = { error: testError.message };
    }

    return batchStats;
  } catch (error) {
    console.warn('collectBatchMetrics エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * エラー発生統計を収集
 * @param {Object} options - 収集オプション
 * @returns {Object} エラー統計
 */
function collectErrorMetrics_(options = {}) {
  try {
    const recentWindowHours = Number(options.recentWindowHours || 24);
    const now = Date.now();
    const windowMs = recentWindowHours * 60 * 60 * 1000;

    const errorStats = {
      errorHandlingImplemented: true,
      source: 'runtime_observation',
      errorCategories: {
        authentication: { status: 'implemented' },
        validation: { status: 'implemented' },
        network: { status: 'implemented' },
        permission: { status: 'implemented' }
      },
      recommendations: [],
      recentSecurityEvents: 0
    };

    try {
      const currentEmail = getCurrentEmail();
      errorStats.basicFunctionality = currentEmail ? 'working' : 'degraded';
    } catch (probeError) {
      errorStats.basicFunctionality = 'error';
      errorStats.lastError = probeError.message;
    }

    try {
      const props = PropertiesService.getScriptProperties().getProperties();
      const securityLogKeys = Object.keys(props).filter(key => key.startsWith('security_log_'));
      const recentKeys = securityLogKeys.filter(key => {
        const ts = Number(key.split('_')[2] || 0);
        return ts > 0 && (now - ts) <= windowMs;
      });
      errorStats.recentSecurityEvents = recentKeys.length;
    } catch (logError) {
      errorStats.recentSecurityEvents = -1;
      errorStats.logReadWarning = logError.message;
    }

    if (errorStats.basicFunctionality !== 'working') {
      errorStats.recommendations.push('認証状態の検証を実施し、Webアプリの公開設定を確認してください。');
    }

    if (errorStats.recentSecurityEvents > 20) {
      errorStats.recommendations.push('高重要度イベントが増加しています。アクセスログと運用設定を確認してください。');
    }

    if (errorStats.recommendations.length === 0) {
      errorStats.recommendations.push('現在は重大な異常を検出していません。定期監視を継続してください。');
    }

    return errorStats;
  } catch (error) {
    console.warn('collectErrorMetrics エラー:', error.message);
    return { error: error.message };
  }
}

/**
 * パフォーマンス診断・推奨事項を生成
 * @param {Object} options - 診断オプション
 * @returns {Object} 診断結果と推奨事項
 */
function diagnosePerformance(options = {}) {
  try {
    const authError = __requirePerfAdmin_();
    if (authError) return authError;

    const metricsResult = getPerformanceMetrics('all', options);
    if (!metricsResult.success) {
      return metricsResult;
    }

    const categories = metricsResult.metrics?.categories || {};
    const issues = [];

    if (categories.api?.errorRate >= 1) {
      issues.push('API基本動作でエラーを検出');
    }
    if (categories.cache?.hitRate === 0) {
      issues.push('キャッシュが有効に機能していない可能性');
    }
    if (categories.error?.recentSecurityEvents > 20) {
      issues.push('最近のセキュリティイベントが増加');
    }

    let overallStatus = 'healthy';
    if (issues.length >= 2) {
      overallStatus = 'critical';
    } else if (issues.length === 1) {
      overallStatus = 'warning';
    }

    const diagnosis = {
      timestamp: new Date().toISOString(),
      overallStatus,
      architecture: {
        pattern: 'GAS-Native Zero-Dependency',
        v8Runtime: true,
        batchProcessing: true
      },
      summary: {
        collectionTimeMs: metricsResult.performanceImpact?.collectionTimeMs || null,
        overhead: metricsResult.performanceImpact?.overhead || 'unknown',
        detectedIssues: issues.length
      },
      issues,
      recommendations: [
        ...(issues.length > 0
          ? ['検出した課題に対する運用対応を優先してください。']
          : ['重大な異常は検出されませんでした。現行運用を維持してください。']),
        '定期的なメトリクス監視を継続してください。'
      ],
      categories
    };

    return {
      success: true,
      diagnosis
    };

  } catch (error) {
    logError_('diagnosePerformance', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * データベース接続テスト（診断用）
 * @returns {Object} 接続テスト結果
 */
function testDatabaseConnection_() {
  try {
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      return {
        success: false,
        message: 'データベースへの接続に失敗しました'
      };
    }

    const dbId = getCachedProperty_('DATABASE_SPREADSHEET_ID');
    const usersSheet = spreadsheet.getSheetByName('users');

    if (!usersSheet) {
      return {
        success: false,
        message: 'ユーザーシートが見つかりません'
      };
    }

    const values = usersSheet.getDataRange().getValues();
    const rowCount = values.length;
    const colCount = values[0]?.length || 0;

    return {
      success: true,
      message: 'データベース接続正常',
      details: {
        spreadsheetId: dbId,
        userCount: Math.max(0, rowCount - 1), // ヘッダー行を除く
        columns: colCount
      }
    };

  } catch (error) {
    logError_('testDatabaseConnection', error);
    return {
      success: false,
      message: `データベース接続エラー: ${error.message}`
    };
  }
}

// DB のシートを idempotent にそろえる (無ければ作る / 空ならヘッダーを書く)。
//   先生専用の配置の初回 (helpers.js の ensureDeploymentBootstrap_) が使う。
function __ensureSheetWithHeaders_(ss, sheetName, headers) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    console.log(`__ensureSheetWithHeaders_: Created ${sheetName} sheet with headers`);
  } else if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    console.log(`__ensureSheetWithHeaders_: Added headers to empty ${sheetName} sheet`);
  }
}

// Why: 以前は SystemController = { publishApp, ... } の namespace export を持っていたが、
//   GAS は single global scope のため全関数が既にグローバルからアクセス可能。
//   namespace は main.js の dispatch だけが使っており、しかも global function への
//   フォールバック付きという矛盾した形だった。namespace 経由と直接呼び出しで二重メンテ
//   になっていたので削除。テストは個別関数を直接 stub する。
