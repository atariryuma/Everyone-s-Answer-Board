/**
 * @fileoverview 統合ユーザーマネージャー - 15個のuser関数を最適化統合
 * Phase 3: ユーザー情報取得最適化 - 重複機能の排除とパフォーマンス向上
 */

/**
 * 統合ユーザーキャッシュマネージャー
 * 全てのユーザー関連キャッシュを一元管理
 */
const unifiedUserCache = {
  // キャッシュレイヤー設定
  layers: {
    fast: { ttl: 60, prefix: 'user_fast_' }, // 1分 - 高頻度アクセス
    standard: { ttl: 180, prefix: 'user_std_' }, // 3分 - 通常使用
    extended: { ttl: 300, prefix: 'user_ext_' }, // 5分 - 閲覧者向け
    secure: { ttl: 120, prefix: 'user_sec_' }, // 2分 - セキュリティ検証付き
  },

  /**
   * レイヤー別キャッシュ取得
   * @param {string} layer - キャッシュレイヤー
   * @param {string} key - キー
   * @returns {object|null} キャッシュされたデータ
   */
  get(layer, key) {
    try {
      const config = this.layers[layer];
      if (!config) return null;

      const cacheKey = config.prefix + key;
      const cached = CacheService.getScriptCache().get(cacheKey);
      return cached ? JSON.parse(cached) : null;
    } catch (error) {
      // キャッシュエラーは無視
      return null;
    }
  },

  /**
   * レイヤー別キャッシュ設定
   * @param {string} layer - キャッシュレイヤー
   * @param {string} key - キー
   * @param {object} data - データ
   */
  set(layer, key, data) {
    try {
      const config = this.layers[layer];
      if (!config) return;

      const cacheKey = config.prefix + key;
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(data), config.ttl);
    } catch (error) {
      // キャッシュ保存エラーは無視
    }
  },

  /**
   * キャッシュ無効化
   * @param {string} userId - ユーザーID
   * @param {string} email - メールアドレス
   */
  invalidate(userId, email) {
    try {
      const keysToRemove = [];

      // 全レイヤーから該当キーを削除
      Object.values(this.layers).forEach((config) => {
        if (userId) {
          keysToRemove.push(config.prefix + userId);
        }
        if (email) {
          keysToRemove.push(config.prefix + 'email_' + email);
        }
      });

      keysToRemove.forEach((key) => {
        try {
          CacheService.getScriptCache().remove(key);
        } catch (e) {
          // 個別キー削除エラーは無視
        }
      });
    } catch (error) {
      // 無効化エラーは無視してログ出力
      debugLog('Cache invalidation error:', error.message);
    }
  },
};

/**
 * 【コア関数1】データベースからユーザー情報を取得する統一関数
 * 15個の関数の最終的な実装基盤
 * @param {string} field - 検索フィールド ('userId' | 'adminEmail')
 * @param {string} value - 検索値
 * @param {object} options - オプション
 * @returns {object|null} ユーザー情報
 */
function coreGetUserFromDatabase(field, value, options = {}) {
  if (!field || !value) {
    return null;
  }

  const opts = {
    forceFresh: options.forceFresh === true,
    cacheLayer: options.cacheLayer || 'standard',
    securityCheck: options.securityCheck === true,
    ...options,
  };

  try {
    // セキュリティ検証（必要時）
    if (opts.securityCheck) {
      const currentUserId = Session.getActiveUser().getEmail();
      if (!multiTenantSecurity.validateTenantBoundary(currentUserId, value, 'user_info_access')) {
        throw new Error('SECURITY_ERROR: ユーザー情報アクセス拒否 - テナント境界違反');
      }
    }

    // キャッシュ確認（forceFresh時はスキップ）
    if (!opts.forceFresh) {
      const cacheKey = field === 'adminEmail' ? `email_${value}` : value;
      const cached = unifiedUserCache.get(opts.cacheLayer, cacheKey);
      if (cached) {
        debugLog(`✅ Cache hit: ${field}=${value} (layer: ${opts.cacheLayer})`);
        return cached;
      }
    }

    // データベース設定取得
    const dbId = getSecureDatabaseId();
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    const service = getSheetsServiceCached(opts.forceFresh);
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    // 値の正規化（メールアドレス検索時）
    let normalizedValue = value;
    if (field.toLowerCase().includes('email')) {
      normalizedValue = String(value).toLowerCase().trim().replace(/\s+/g, '');
    }

    // 強制フレッシュ時のキャッシュ無効化
    if (opts.forceFresh) {
      const userId = field === 'userId' ? value : null;
      const email = field === 'adminEmail' ? value : null;
      unifiedUserCache.invalidate(userId, email);
    }

    // データベースからデータ取得
    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`]);
    const values = data.valueRanges[0].values || [];

    if (values.length <= 1) {
      debugLog(`❌ No data found: ${field}=${value}`);
      return null;
    }

    const headers = values[0];
    const fieldIndex = headers.findIndex(
      (header) => header && header.toString().toLowerCase() === field.toLowerCase()
    );

    if (fieldIndex === -1) {
      throw new Error(`データベースに ${field} フィールドが見つかりません`);
    }

    // ユーザー検索
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row[fieldIndex]) continue;

      let cellValue = row[fieldIndex].toString().trim();
      if (field.toLowerCase().includes('email')) {
        cellValue = cellValue.toLowerCase().replace(/\s+/g, '');
      }

      if (cellValue === normalizedValue) {
        // ユーザーオブジェクト構築
        const userObj = {};
        headers.forEach((header, index) => {
          if (header && row[index] !== undefined) {
            userObj[header] = row[index];
          }
        });

        // キャッシュに保存
        const cacheKey = field === 'adminEmail' ? `email_${value}` : userObj.userId || value;
        unifiedUserCache.set(opts.cacheLayer, cacheKey, userObj);

        debugLog(`✅ User found: ${field}=${value} (cached in ${opts.cacheLayer})`);
        return userObj;
      }
    }

    debugLog(`❌ User not found: ${field}=${value}`);
    return null;
  } catch (error) {
    console.error(`[ERROR] coreGetUserFromDatabase error (${field}=${value}):`, error.message);
    return null;
  }
}

/**
 * 【コア関数2】現在のユーザーのメールアドレス取得
 * @returns {string} ユーザーメールアドレス
 */
function coreGetCurrentUserEmail() {
  try {
    const email = Session.getActiveUser().getEmail();
    debugLog('coreGetCurrentUserEmail: Retrieved email:', email);
    return email || '';
  } catch (error) {
    console.error('[ERROR] coreGetCurrentUserEmail:', error.message);
    return '';
  }
}

/**
 * 【コア関数3】ユーザーID生成・管理
 * @param {string} requestUserId - リクエストユーザーID
 * @returns {string} ユーザーID
 */
function coreGetUserId(requestUserId) {
  if (requestUserId) return requestUserId;

  const email = coreGetCurrentUserEmail();
  if (!email) {
    throw new Error('ユーザーIDを取得できませんでした。');
  }

  // メールアドレスベースのユニークキー生成
  const userKey =
    'CURRENT_USER_ID_' +
    Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, email, Utilities.Charset.UTF_8)
      .map(function (byte) {
        return (byte + 256).toString(16).slice(-2);
      })
      .join('');

  try {
    const props = getResilientUserProperties();
    let userId = props.getProperty(userKey);

    if (!userId) {
      userId = 'USER_' + Date.now() + '_' + Math.random().toString(36).substring(2, 8);
      props.setProperty(userKey, userId);
      debugLog('New userId generated:', userId, 'for email:', email);
    }

    return userId;
  } catch (error) {
    console.error('[ERROR] coreGetUserId:', error.message);
    throw new Error('ユーザーID生成に失敗しました: ' + error.message);
  }
}

/**
 * 【コア関数4】全ユーザー一覧取得
 * @returns {Array} ユーザー情報配列
 */
function coreGetAllUsers() {
  try {
    const dbId = getSecureDatabaseId();
    if (!dbId) {
      throw new Error('データベースIDが設定されていません');
    }

    const service = getSheetsServiceCached();
    const sheetName = DB_SHEET_CONFIG.SHEET_NAME;

    const data = batchGetSheetsData(service, dbId, [`'${sheetName}'!A:H`]);
    const values = data.valueRanges[0].values || [];

    if (values.length <= 1) {
      return []; // ヘッダーのみの場合は空配列
    }

    const headers = values[0];
    const users = [];

    // データをオブジェクト形式に変換
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const user = {};

      headers.forEach((header, index) => {
        if (header && row[index] !== undefined) {
          user[header] = row[index];
        }
      });

      if (user.userId) {
        // userIdが存在するもののみ追加
        users.push(user);
      }
    }

    debugLog(`✅ Retrieved ${users.length} users from database`);
    return users;
  } catch (error) {
    console.error('[ERROR] coreGetAllUsers:', error.message);
    throw new Error('ユーザー一覧取得に失敗しました: ' + error.message);
  }
}

/**
 * ===== 以下、既存15個の関数を互換性保持ラッパーとして実装 =====
 */

/**
 * 1. fetchUserFromDatabase - 既存コア関数との互換性
 * @param {string} field - 検索フィールド
 * @param {string} value - 検索値
 * @param {object} options - オプション
 * @returns {object|null} ユーザー情報
 */
function fetchUserFromDatabase(field, value, options = {}) {
  return coreGetUserFromDatabase(field, value, {
    cacheLayer: 'standard',
    ...options,
  });
}

/**
 * 2. findUserById - 標準キャッシュでID検索
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function findUserById(userId) {
  return coreGetUserFromDatabase('userId', userId, {
    cacheLayer: 'standard',
  });
}

/**
 * 3. findUserByIdFresh - キャッシュバイパスでID検索
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function findUserByIdFresh(userId) {
  const result = coreGetUserFromDatabase('userId', userId, {
    forceFresh: true,
    cacheLayer: 'standard',
  });

  if (result) {
    infoLog('✅ Fresh user data retrieved for:', userId);
  }

  return result;
}

/**
 * 4. findUserByEmail - 標準キャッシュでメール検索
 * @param {string} email - メールアドレス
 * @returns {object|null} ユーザー情報
 */
function findUserByEmail(email) {
  return coreGetUserFromDatabase('adminEmail', email, {
    cacheLayer: 'standard',
  });
}

/**
 * 5. findUserByIdForViewer - 閲覧者向け長期キャッシュ
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function findUserByIdForViewer(userId) {
  return coreGetUserFromDatabase('userId', userId, {
    cacheLayer: 'extended',
  });
}

/**
 * 6. getSecureUserInfo - セキュリティ検証付きユーザー情報取得
 * @param {string} userId - ユーザーID
 * @param {object} options - オプション
 * @returns {object|null} ユーザー情報
 */
function getSecureUserInfo(userId, options = {}) {
  return coreGetUserFromDatabase('userId', userId, {
    cacheLayer: 'secure',
    securityCheck: true,
    ...options,
  });
}

/**
 * 7. getSecureUserConfig - セキュリティ検証付き設定情報取得
 * @param {string} userId - ユーザーID
 * @param {string} configKey - 設定キー
 * @param {object} options - オプション
 * @returns {any} 設定値
 */
function getSecureUserConfig(userId, configKey, options = {}) {
  const userInfo = coreGetUserFromDatabase('userId', userId, {
    cacheLayer: 'secure',
    securityCheck: true,
    ...options,
  });

  if (!userInfo || !userInfo.configJson) {
    return null;
  }

  try {
    const config = JSON.parse(userInfo.configJson);
    return config[configKey] || null;
  } catch (error) {
    console.error('[ERROR] getSecureUserConfig JSON parse error:', error.message);
    return null;
  }
}

/**
 * 8. getUserIdFromEmail - メールからユーザーID取得
 * @param {string} email - メールアドレス
 * @returns {string|null} ユーザーID
 */
function getUserIdFromEmail(email) {
  const userInfo = coreGetUserFromDatabase('adminEmail', email, {
    cacheLayer: 'standard',
  });
  return userInfo ? userInfo.userId : null;
}

/**
 * 9. getEmailFromUserId - ユーザーIDからメール取得
 * @param {string} userId - ユーザーID
 * @returns {string|null} メールアドレス
 */
function getEmailFromUserId(userId) {
  const userInfo = coreGetUserFromDatabase('userId', userId, {
    cacheLayer: 'standard',
  });
  return userInfo ? userInfo.adminEmail : null;
}

/**
 * 10. getOrFetchUserInfo - 汎用ユーザー情報取得
 * @param {string} identifier - ID またはメール
 * @param {string} type - タイプ指定
 * @param {object} options - オプション
 * @returns {object|null} ユーザー情報
 */
function getOrFetchUserInfo(identifier, type = null, options = {}) {
  try {
    // タイプに基づいて適切なフィールドを選択
    if (type === 'email' || (!type && typeof identifier === 'string' && identifier.includes('@'))) {
      return coreGetUserFromDatabase('adminEmail', identifier, {
        cacheLayer: 'standard',
        ...options,
      });
    } else {
      return coreGetUserFromDatabase('userId', identifier, {
        cacheLayer: 'standard',
        ...options,
      });
    }
  } catch (error) {
    console.error('[ERROR] getOrFetchUserInfo error:', error.message);
    return null;
  }
}

/**
 * 11. getUserWithFallback - フォールバック付きユーザー取得
 * @param {string} userId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function getUserWithFallback(userId) {
  // 入力検証
  if (!userId || typeof userId !== 'string') {
    warnLog('getUserWithFallback: Invalid userId:', userId);
    return null;
  }

  const user = coreGetUserFromDatabase('userId', userId, {
    cacheLayer: 'standard',
  });

  if (!user) {
    handleMissingUser(userId);
  }

  return user;
}

/**
 * 12. getConfigUserInfo - 設定用ユーザー情報取得
 * @param {string} requestUserId - ユーザーID
 * @returns {object|null} ユーザー情報
 */
function getConfigUserInfo(requestUserId) {
  return coreGetUserFromDatabase('userId', requestUserId, {
    cacheLayer: 'standard',
  });
}

/**
 * 13. getCurrentUserEmail - 現在のユーザーメール取得
 * @returns {string} ユーザーメール
 */
function getCurrentUserEmail() {
  return coreGetCurrentUserEmail();
}

/**
 * 14. getUserId - ユーザーID取得・生成
 * @param {string} requestUserId - リクエストユーザーID
 * @returns {string} ユーザーID
 */
function getUserId(requestUserId) {
  return coreGetUserId(requestUserId);
}

/**
 * 15. getAllUsers - 全ユーザー一覧取得
 * @returns {Array} ユーザー情報配列
 */
function getAllUsers() {
  return coreGetAllUsers();
}

/**
 * 管理者用全ユーザー取得（権限チェック付き）
 * @returns {Array} ユーザー情報配列
 */
function getAllUsersForAdmin() {
  try {
    // 管理者権限チェック
    if (!isDeployUser()) {
      throw new Error('この機能にアクセスする権限がありません。');
    }

    return coreGetAllUsers();
  } catch (error) {
    console.error('[ERROR] getAllUsersForAdmin:', error.message);
    throw new Error('全ユーザー取得に失敗しました: ' + error.message);
  }
}

/**
 * ユーザーキャッシュ無効化ヘルパー
 * @param {string} userId - ユーザーID
 * @param {string} email - メールアドレス
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {boolean} broadInvalidation - 広範囲無効化フラグ
 */
function invalidateUserCache(userId, email, spreadsheetId, broadInvalidation = false) {
  // 統一APIを使用（後方互換性維持）
  return unifiedCacheAPI.invalidateUserCache(userId, email, spreadsheetId, broadInvalidation);
}

/**
 * 統合ユーザーマネージャーの健全性チェック
 * @returns {object} 健全性チェック結果
 */
function checkUnifiedUserManagerHealth() {
  try {
    const stats = {
      timestamp: new Date().toISOString(),
      cacheLayersHealthy: true,
      databaseAccessible: false,
      totalFunctions: 15,
      coreComponents: 4,
    };

    // データベースアクセステスト
    try {
      const dbId = getSecureDatabaseId();
      stats.databaseAccessible = !!dbId;
    } catch (e) {
      stats.databaseAccessible = false;
    }

    // キャッシュレイヤーテスト
    try {
      unifiedUserCache.set('fast', 'health_check', { test: true });
      const retrieved = unifiedUserCache.get('fast', 'health_check');
      stats.cacheLayersHealthy = !!retrieved;
    } catch (e) {
      stats.cacheLayersHealthy = false;
    }

    return stats;
  } catch (error) {
    console.error('[ERROR] checkUnifiedUserManagerHealth:', error.message);
    return {
      timestamp: new Date().toISOString(),
      healthy: false,
      error: error.message,
    };
  }
}

/**
 * 統合後のパフォーマンス統計
 * @returns {object} パフォーマンス統計
 */
function getUnifiedUserManagerStats() {
  return {
    timestamp: new Date().toISOString(),
    optimization: {
      originalFunctions: 15,
      coreComponents: 4,
      reductionRatio: '73%',
      cacheLayerCount: 4,
      duplicateElimination: 'Complete',
    },
    cacheConfig: unifiedUserCache.layers,
    healthStatus: checkUnifiedUserManagerHealth(),
  };
}
