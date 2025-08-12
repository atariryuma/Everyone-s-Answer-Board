/**
 * 統一ユーティリティモジュール
 * StudyQuestの共通機能を統合し、重複コードを削減
 */

/**
 * 統一ユーザー情報管理クラス
 * 複数の重複関数を統合
 */
class UnifiedUserManager {

  /**
   * ユーザー情報を統一的に取得
   * getUserInfoCached, getOrFetchUserInfo, findUserById を統合
   * @param {string} identifier - ユーザーIDまたはメールアドレス
   * @param {object} options - オプション設定
   * @returns {object|null} ユーザー情報
   */
  static getUser(identifier, options = {}) {
    const {
      forceRefresh = false,
      useCache = true,
      authCheck = true
    } = options;

    try {
      // 認証チェック（必要時）
      if (authCheck) {
        verifyUserAccess(identifier);
      }

      // キャッシュ確認（forceRefresh=false の場合）
      if (useCache && !forceRefresh) {
        const cached = this._getCachedUser(identifier);
        if (cached) {
          debugLog(`✅ 統一ユーザー管理: キャッシュヒット (${identifier})`);
          return cached;
        }
      }

      // データベースから取得
      debugLog(`🔍 統一ユーザー管理: DB検索 (${identifier})`);
      const userInfo = this._fetchUserFromDatabase(identifier);

      // キャッシュ保存
      if (userInfo && useCache) {
        this._cacheUser(identifier, userInfo);
      }

      return userInfo;

    } catch (error) {
      errorLog('統一ユーザー管理エラー:', error.message);
      return null;
    }
  }

  /**
   * キャッシュからユーザー情報を取得
   * @private
   */
  static _getCachedUser(identifier) {
    // 統一キャッシュシステムを利用
    const unifiedCache = getUnifiedExecutionCache();
    return unifiedCache.getUserInfo(identifier);
  }

  /**
   * ユーザー情報をキャッシュに保存
   * @private
   */
  static _cacheUser(identifier, userInfo) {
    const unifiedCache = getUnifiedExecutionCache();
    unifiedCache.setUserInfo(identifier, userInfo);
  }

  /**
   * データベースからユーザー情報を取得
   * @private
   */
  static _fetchUserFromDatabase(identifier) {
    // 既存のfindUserById関数を利用（後方互換性）
    if (typeof findUserById === 'function') {
      return findUserById(identifier);
    }

    // フォールバック実装
    return this._fallbackUserFetch(identifier);
  }

  /**
   * フォールバック用ユーザー取得 - 強化版
   * @private
   */
  static _fallbackUserFetch(identifier) {
    warnLog('統一ユーザー管理: フォールバック実装を使用（強化版）');
    
    try {
      // 既存の関数を活用してフォールバック実装を強化
      if (typeof findUserById === 'function') {
        debugLog('フォールバック: findUserByIdを使用');
        return findUserById(identifier);
      }
      
      if (typeof findUserByEmail === 'function') {
        debugLog('フォールバック: findUserByEmailを使用');
        return findUserByEmail(identifier);
      }
      
      // 最後の手段として実行レベルキャッシュから取得を試行
      if (typeof getUnifiedExecutionCache === 'function') {
        const execCache = getUnifiedExecutionCache();
        const cachedUser = execCache.getUserInfo(identifier);
        if (cachedUser) {
          debugLog('フォールバック: UnifiedExecutionCacheからユーザー情報を復旧');
          return cachedUser;
        }
      }
      
      warnLog('⚠️ フォールバック: 全ての取得方法が失敗');
      return null;
      
    } catch (fallbackError) {
      errorLog('❌ フォールバック処理でエラー:', fallbackError.message);
      return null;
    }
  }

  /**
   * ユーザーキャッシュをクリア
   * 複数のclearExecutionUserInfoCache関数を統合
   */
  static clearUserCache(identifier = null) {
    try {
      const unifiedCache = getUnifiedExecutionCache();

      if (identifier) {
        // 特定ユーザーのキャッシュクリア
        unifiedCache.clearUserInfo();
        debugLog(`🗑️ 統一ユーザー管理: 特定ユーザーキャッシュクリア (${identifier})`);
      } else {
        // 全ユーザーキャッシュクリア
        unifiedCache.clearAll();
        debugLog('🗑️ 統一ユーザー管理: 全ユーザーキャッシュクリア');
      }

      // 統一キャッシュマネージャーとの同期
      unifiedCache.syncWithUnifiedCache('userDataChange');

    } catch (error) {
      errorLog('ユーザーキャッシュクリアエラー:', error.message);
    }
  }
}

/**
 * 統一URL管理クラス
 * URL関連の重複機能を統合
 */
class UnifiedURLManager {

  /**
   * WebアプリURLを統一的に取得
   * @returns {string}
   */
  static getWebAppURL() {
    return getWebAppUrl();
  }

  /**
   * ユーザー用URLを生成
   * @param {string} userId - ユーザーID
   * @param {object} options - URL生成オプション
   * @returns {object} 各種URL
   */
  static generateUserURLs(userId, options = {}) {
    return generateUserUrlsWithCacheBusting(userId, options);
  }

  /**
   * URL検証・サニタイズ
   * @param {string} url - 検証対象URL
   * @returns {string} サニタイズされたURL
   */
  static sanitizeURL(url) {
    if (!url || typeof url !== 'string') {
      return '';
    }

    try {
      let sanitized = String(url).trim();
      if ((sanitized.startsWith('"') && sanitized.endsWith('"')) ||
          (sanitized.startsWith("'") && sanitized.endsWith("'"))) {
        sanitized = sanitized.slice(1, -1);
      }
      sanitized = sanitized.replace(/\\"/g, '"').replace(/\\'/g, "'");

      if (sanitized.includes('javascript:') || sanitized.includes('data:')) {
        warnLog('統一URL管理: 危険なURLスキームを検出');
        return '';
      }

      return sanitized;

    } catch (error) {
      errorLog('URL検証エラー:', error.message);
      return '';
    }
  }
}

/**
 * 統一API呼び出しクラス
 * UrlFetchAppの重複パターンを統合
 */
class UnifiedAPIClient {

  /**
   * 統一API呼び出し
   * @param {string} url - API URL
   * @param {object} options - リクエストオプション
   * @returns {object} レスポンス
   */
  static async request(url, options = {}) {
    const {
      method = 'GET',
      headers = {},
      payload = null,
      timeout = 30000,
      retries = 2,
      authToken = null
    } = options;

    // 認証ヘッダー追加
    if (authToken) {
      headers['Authorization'] = `Bearer ${authToken}`;
    }

    const requestConfig = {
      method: method,
      headers: headers,
      muteHttpExceptions: true
    };

    if (payload) {
      requestConfig.payload = payload;
    }

    let lastError = null;

    // リトライロジック
    for (let attempt = 0; attempt <= retries; attempt++) {
      try {
        debugLog(`🌐 統一API: ${method} ${url} (試行 ${attempt + 1}/${retries + 1})`);

        const response = resilientUrlFetch(url, requestConfig);
        
        // レスポンスオブジェクトの検証
        if (!response || typeof response.getResponseCode !== 'function') {
          throw new Error('無効なレスポンスオブジェクトが返されました');
        }
        
        const statusCode = response.getResponseCode();

        // 成功時
        if (statusCode >= 200 && statusCode < 300) {
          infoLog(`✅ 統一API成功: ${statusCode} ${method} ${url}`);
          return {
            success: true,
            status: statusCode,
            data: response.getContentText(),
            response: response
          };
        }

        // 4xxエラーはリトライしない
        if (statusCode >= 400 && statusCode < 500) {
          warnLog(`❌ 統一APIクライアントエラー: ${statusCode} ${method} ${url}`);
          return {
            success: false,
            status: statusCode,
            error: response.getContentText(),
            response: response
          };
        }

        // 5xxエラーはリトライ
        lastError = new Error(`サーバーエラー: ${statusCode}`);

      } catch (error) {
        lastError = error;
        warnLog(`⚠️ 統一API例外 (試行 ${attempt + 1}): ${error.message}`);

        // 最後の試行でない場合は待機
        if (attempt < retries) {
          Utilities.sleep(1000 * (attempt + 1)); // 指数バックオフ
        }
      }
    }

    // 全試行失敗
    errorLog('❌ 統一API全試行失敗:', lastError.message);
    return {
      success: false,
      status: 0,
      error: lastError.message
    };
  }

  /**
   * Google Sheets API専用呼び出し
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} operation - 操作種別
   * @param {object} data - リクエストデータ
   */
  static async sheetsAPI(spreadsheetId, operation, data = {}) {
    const token = getServiceAccountTokenCached();
    const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${spreadsheetId}`;

    const endpoints = {
      'get': `${baseUrl}`,
      'batchUpdate': `${baseUrl}:batchUpdate`,
      'values': `${baseUrl}/values/${data.range || 'A1'}`
    };

    const url = endpoints[operation];
    if (!url) {
      throw new Error(`未対応のSheets API操作: ${operation}`);
    }

    const options = {
      method: operation === 'get' ? 'GET' : 'POST',
      authToken: token,
      headers: {
        'Content-Type': 'application/json'
      }
    };

    if (data && Object.keys(data).length > 0) {
      options.payload = JSON.stringify(data);
    }

    return this.request(url, options);
  }
}

/**
 * 統一バリデーションエンジン
 * 各種検証関数を統合
 */
class UnifiedValidation {

  /**
   * 統一バリデーション関数
   * @param {any} data - 検証対象データ
   * @param {object} rules - バリデーションルール
   * @returns {object} 検証結果
   */
  static validate(data, rules) {
    const errors = [];
    const warnings = [];

    try {
      for (const [field, rule] of Object.entries(rules)) {
        const value = data[field];
        const result = this._validateField(field, value, rule);

        if (result.errors.length > 0) {
          errors.push(...result.errors);
        }
        if (result.warnings.length > 0) {
          warnings.push(...result.warnings);
        }
      }

      return {
        isValid: errors.length === 0,
        errors: errors,
        warnings: warnings
      };

    } catch (error) {
      errorLog('統一バリデーションエラー:', error.message);
      return {
        isValid: false,
        errors: [`バリデーション処理エラー: ${error.message}`],
        warnings: []
      };
    }
  }

  /**
   * フィールド単体のバリデーション
   * @private
   */
  static _validateField(field, value, rule) {
    const errors = [];
    const warnings = [];

    // 必須チェック
    if (rule.required && (value === undefined || value === null || value === '')) {
      errors.push(`${field} は必須項目です`);
      return { errors, warnings };
    }

    // 型チェック
    if (value !== undefined && rule.type && typeof value !== rule.type) {
      errors.push(`${field} の型が正しくありません。期待値: ${rule.type}, 実際: ${typeof value}`);
    }

    // 長さチェック
    if (rule.minLength && value.length < rule.minLength) {
      errors.push(`${field} は ${rule.minLength} 文字以上である必要があります`);
    }

    if (rule.maxLength && value.length > rule.maxLength) {
      errors.push(`${field} は ${rule.maxLength} 文字以下である必要があります`);
    }

    // パターンチェック
    if (rule.pattern && !rule.pattern.test(value)) {
      errors.push(`${field} の形式が正しくありません`);
    }

    return { errors, warnings };
  }

  /**
   * ユーザー情報専用バリデーション
   * @param {object} userInfo - ユーザー情報
   * @returns {object} 検証結果
   */
  static validateUser(userInfo) {
    const rules = {
      userId: { required: true, type: 'string', minLength: 1 },
      email: {
        required: true,
        type: 'string',
        pattern: /^[^\s@]+@[^\s@]+\.[^\s@]+$/
      },
      spreadsheetId: { type: 'string' },
      configJson: { type: 'string' }
    };

    return this.validate(userInfo, rules);
  }

  /**
   * 設定情報専用バリデーション
   * @param {object} config - 設定情報
   * @returns {object} 検証結果
   */
  static validateConfig(config) {
    const rules = {
      setupStatus: {
        required: true,
        type: 'string'
      },
      formCreated: { required: true, type: 'boolean' },
      appPublished: { required: true, type: 'boolean' },
      publishedSheetName: { type: 'string' },
      publishedSpreadsheetId: { type: 'string' }
    };

    return this.validate(config, rules);
  }
}

// 後方互換性のためのラッパー関数群

/**
 * 既存のgetUserInfoCached関数のラッパー
 * @deprecated UnifiedUserManager.getUser を使用してください
 */
function getUserInfoCachedUnified(requestUserId) {
  return UnifiedUserManager.getUser(requestUserId, { useCache: true });
}

/**
 * 既存のgetWebAppUrl関数のラッパー
 * @deprecated UnifiedURLManager.getWebAppURL を使用してください
 */
function getWebAppUrlUnified() {
  return UnifiedURLManager.getWebAppURL({ useCache: true });
}

/**
 * 既存のclearExecutionUserInfoCache関数のラッパー
 * @deprecated UnifiedUserManager.clearUserCache を使用してください
 */
function clearExecutionUserInfoCacheUnified() {
  return UnifiedUserManager.clearUserCache();
}