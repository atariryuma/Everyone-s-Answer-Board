/**
 * @fileoverview 簡略化された統合キャッシュマネージャー
 * 基本的なキャッシュ機能のみを保持し、複雑なキャッシュクラスを除去
 */

/**
 * シンプルなCacheService永続キャッシュ管理
 * GAS実行環境の特性に合わせてglobalThis依存を排除
 */
const SimpleCacheManager = {
  scriptCache: CacheService.getScriptCache(),
  defaultTTL: 21600, // 6時間

  /**
   * CacheService永続キャッシュから取得、なければ生成して保存
   * disableCacheService: true の場合は関数オブジェクト保護のためキャッシュをスキップ
   */
  get(key, valueFn, options = {}) {
    const { ttl = this.defaultTTL, disableCacheService = false } = options;

    try {
      // シンプル版：disableCacheService=true の場合は関数実行のみ
      if (disableCacheService) {
        return typeof valueFn === 'function' ? valueFn() : null;
      }

      // CacheService永続キャッシュから取得
      const cachedValue = this.scriptCache.get(key);
      if (cachedValue !== null) {
        return JSON.parse(cachedValue);
      }

      // キャッシュにない場合は関数実行
      if (typeof valueFn === 'function') {
        const newValue = valueFn();
        this.set(key, newValue, { ttl });
        return newValue;
      }

      return null;
    } catch (error) {
      console.error('SimpleCacheManager.get エラー:', error.message);
      return null;
    }
  },

  /**
   * CacheService永続キャッシュに保存
   */
  set(key, value, options = {}) {
    const { ttl = this.defaultTTL } = options;

    try {
      this.scriptCache.put(key, JSON.stringify(value), ttl);
    } catch (error) {
      console.error('SimpleCacheManager.set エラー:', error.message);
    }
  },

  /**
   * CacheService永続キャッシュから削除
   */
  remove(key) {
    try {
      this.scriptCache.remove(key);
    } catch (error) {
      console.error('SimpleCacheManager.remove エラー:', error.message);
    }
  },

  /**
   * 🚨 緊急時キャッシュクリア（システム復旧用）
   * Service Accountトークンとサービスオブジェクトを強制リセット
   */
  clearAll() {
    try {
      Logger.info('緊急キャッシュクリア開始');

      // Service Accountトークンクリア
      this.scriptCache.remove('SA_TOKEN_CACHE');
      this.scriptCache.remove('sheets_service_optimized');
      this.scriptCache.remove('sheets_service');

      // その他のキャッシュクリア
      const commonCacheKeys = ['user_config', 'form_info', 'system_status'];
      commonCacheKeys.forEach((key) => this.scriptCache.remove(key));

      Logger.info('緊急キャッシュクリア完了');
    } catch (error) {
      Logger.error('緊急キャッシュクリアエラー', { error: error.message });
    }
  },
};

// 後方互換性のためのエイリアス
const cacheManager = SimpleCacheManager;
Logger.info('シンプルCacheService初期化完了');

/**
 * Sheets APIサービス結果のキャッシュ（互換性維持）
 * @returns {object|null} キャッシュされたサービス情報
 */
function getSheetsServiceCached() {
  const cacheKey = 'sheets_service_optimized';

  // メモリキャッシュから確認
  try {
    const cachedService = cacheManager.get(cacheKey, null, { disableCacheService: false });
    if (cachedService !== null) {
      // 関数検証：必要なAPI関数が存在するかチェック
      const isValidService =
        cachedService?.spreadsheets?.values?.append &&
        typeof cachedService.spreadsheets.values.append === 'function' &&
        cachedService?.spreadsheets?.values?.batchGet &&
        typeof cachedService.spreadsheets.values.batchGet === 'function' &&
        cachedService?.spreadsheets?.values?.update &&
        typeof cachedService.spreadsheets.values.update === 'function';

      if (isValidService) {
        return cachedService;
      } else {
        // 不完全なキャッシュをクリア
        cacheManager.remove(cacheKey);
      }
    }
  } catch (error) {
    // キャッシュアクセスエラーは無視してサービスを再生成
    Logger.warn('キャッシュアクセスエラー', { error: error.message });
  }

  const result = cacheManager.get(
    cacheKey,
    () => {
      // Service Account認証確認
      let testToken;
      try {
        testToken = getServiceAccountTokenCached();
        if (!testToken || testToken.length <= 100) {
          throw new Error('無効なトークンを取得しました');
        }
      } catch (tokenError) {
        Logger.error('Service Accountトークン取得エラー', {
          error: tokenError.message,
          context: 'service_object_creation'
        });
        throw new Error('Service Account認証が利用できません。システム管理者に連絡してください。');
      }

      // Service Objectトークン取得（legacy互換性のため）
      const initialAccessToken = getServiceAccountTokenCached();

      const serviceObject = {
        baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets',
        accessToken: initialAccessToken, // ✅ getSpreadsheetsData互換性修復
        spreadsheets: {
          batchUpdate (params) {
            // 最新のアクセストークンを取得（トークンの期限切れ対応）
            const accessToken = getServiceAccountTokenCached();
            if (!accessToken) {
              throw new Error('Service Account token is not available');
            }


            // Sheets API v4 batchUpdate実装
            const url = `https://sheets.googleapis.com/v4/spreadsheets/${params.spreadsheetId}:batchUpdate`;

            const response = UrlFetchApp.fetch(url, {
              method: 'POST',
              headers: {
                Authorization: `Bearer ${accessToken}`,
                'Content-Type': 'application/json',
              },
              payload: JSON.stringify({
                requests: params.requests,
              }),
              muteHttpExceptions: true,
            });


            if (response.getResponseCode() !== 200) {
              throw new Error(`Sheets API Error: ${response.getContentText()}`);
            }

            return JSON.parse(response.getContentText());
          },
          values: {
            batchGet (params) {
              // 最新のアクセストークンを取得（トークンの期限切れ対応）
              const accessToken = getServiceAccountTokenCached();
              if (!accessToken) {
                throw new Error('Service Account token is not available');
              }


              // Sheets API v4 batchGet実装
              const url = `https://sheets.googleapis.com/v4/spreadsheets/${params.spreadsheetId}/values:batchGet`;
              const queryParams = params.ranges ? `?ranges=${params.ranges.join('&ranges=')}` : '';

              const response = UrlFetchApp.fetch(url + queryParams, {
                method: 'GET',
                headers: {
                  Authorization: `Bearer ${accessToken}`,
                  'Content-Type': 'application/json',
                },
                muteHttpExceptions: true,
              });


              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API Error: ${response.getContentText()}`);
              }

              return JSON.parse(response.getContentText());
            },
            update (params) {
              // 最新のアクセストークンを取得（トークンの期限切れ対応）
              const accessToken = getServiceAccountTokenCached();
              if (!accessToken) {
                throw new Error('Service Account token is not available');
              }


              // Sheets API v4 update実装
              const url = `https://sheets.googleapis.com/v4/spreadsheets/${params.spreadsheetId}/values/${params.range}?valueInputOption=RAW`;

              const response = UrlFetchApp.fetch(url, {
                method: 'PUT',
                headers: {
                  Authorization: `Bearer ${accessToken}`,
                  'Content-Type': 'application/json',
                },
                payload: JSON.stringify({
                  values: params.values,
                }),
                muteHttpExceptions: true,
              });


              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API Error: ${response.getContentText()}`);
              }

              return JSON.parse(response.getContentText());
            },
            append (params) {
              // 最新のアクセストークンを取得（トークンの期限切れ対応）
              const accessToken = getServiceAccountTokenCached();
              if (!accessToken) {
                throw new Error('Service Account token is not available');
              }

              // Sheets API v4 append実装
              const url = `https://sheets.googleapis.com/v4/spreadsheets/${params.spreadsheetId}/values/${params.range}:append?valueInputOption=${params.valueInputOption || 'RAW'}&insertDataOption=${params.insertDataOption || 'INSERT_ROWS'}`;

              const response = UrlFetchApp.fetch(url, {
                method: 'POST',
                headers: {
                  Authorization: `Bearer ${accessToken}`,
                  'Content-Type': 'application/json',
                },
                payload: JSON.stringify({
                  values: params.values,
                }),
                muteHttpExceptions: true,
              });


              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API Error: ${response.getContentText()}`);
              }

              return JSON.parse(response.getContentText());
            },
          },
        },
      };

      // 🔧 service object構築完了確認（構築成功時は簡潔ログ）
      const isComplete =
        serviceObject.spreadsheets &&
        serviceObject.spreadsheets.values &&
        typeof serviceObject.spreadsheets.values.batchGet === 'function' &&
        typeof serviceObject.spreadsheets.values.update === 'function' &&
        typeof serviceObject.spreadsheets.values.append === 'function';

      if (!isComplete) {
        Logger.warn('service object構築異常', {
          hasSpreadsheets: !!serviceObject.spreadsheets,
          hasValues: !!serviceObject.spreadsheets?.values,
        });
      }

      return serviceObject;
    },
    {
      ttl: 900, // 15分間に延長（パフォーマンス重視）
      enableMemoization: true,
      disableCacheService: true, // ✅ CacheService無効化（関数オブジェクト保護）
    }
  );

  // 異常時のみデバッグ情報を出力
  const hasAllMethods =
    result?.spreadsheets?.values?.append &&
    result?.spreadsheets?.values?.batchGet &&
    result?.spreadsheets?.values?.update;

  if (!hasAllMethods) {
    Logger.warn('キャッシュからの戻り値異常', {
      resultType: typeof result,
      hasSpreadsheets: !!result?.spreadsheets,
    });
  }

  return result;
}

/**
 * 汎用スプレッドシートヘッダー取得関数
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {object} [options] - オプション { useCache: boolean, validate: boolean, forceRefresh: boolean }
 * @returns {object} ヘッダーインデックス情報
 */
function getSpreadsheetHeaders(spreadsheetId, sheetName, options = {}) {
  const { useCache = true, validate = false, forceRefresh = false } = options;

  if (!spreadsheetId || !sheetName) {
    throw new Error('スプレッドシートIDとシート名は必須です');
  }

  const cacheKey = `headers_${spreadsheetId}_${sheetName}`;

  // キャッシュから取得を試行（forceRefreshでない場合）
  if (useCache && !forceRefresh) {
    const cached = SimpleCacheManager.get(cacheKey, null);
    if (cached && (!validate || validateSpreadsheetHeaders(cached).success)) {
      return cached;
    }
  }

  try {
    // スプレッドシートから直接取得
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート「${sheetName}」が見つかりません`);
    }

    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headerRow || headerRow.length === 0) {
      throw new Error('ヘッダー行が見つかりません');
    }

    // ヘッダーインデックスマップを作成
    const headerIndices = {};
    headerRow.forEach((header, index) => {
      if (header && String(header).trim() !== '') {
        headerIndices[String(header).trim()] = index;
      }
    });

    // 検証実行（必要な場合）
    if (validate) {
      const validationResult = validateSpreadsheetHeaders(headerIndices);
      if (!validationResult.success) {
        console.warn('ヘッダー検証に失敗:', validationResult.missing);
      }
    }

    // キャッシュに保存
    if (useCache) {
      SimpleCacheManager.set(cacheKey, headerIndices, { ttl: 1800 });
    }

    return headerIndices;
  } catch (error) {
    console.error('[ERROR] getSpreadsheetHeaders:', error.message);
    throw new Error(`ヘッダー取得エラー: ${error.message}`);
  }
}

/**
 * スプレッドシートヘッダーの検証（柔軟な列名検出）
 * @param {object} headerIndices - ヘッダーインデックス
 * @returns {object} 検証結果 { success: boolean, missing: string[], hasReasonColumn: boolean, hasOpinionColumn: boolean }
 */
function validateSpreadsheetHeaders(headerIndices) {
  if (!headerIndices || typeof headerIndices !== 'object') {
    return createResponse(false, 'ヘッダー検証失敗', {
      missing: ['すべて'],
      hasReasonColumn: false,
      hasOpinionColumn: false,
    });
  }

  const headerNames = Object.keys(headerIndices);

  // 動的な列名パターン検出
  const reasonPatterns = [
    '理由',
    'なぜ',
    'どうして',
    '根拠',
    'わけ',
    'reason',
    'why',
    '考える理由',
    '体験',
    '経験',
  ];

  const opinionPatterns = [
    '回答',
    '答え',
    '意見',
    'こたえ',
    '考え',
    '思考',
    'answer',
    'opinion',
    'どう思いますか',
    '書きましょう',
    '教えてください',
  ];

  // パターンマッチングで列を検出
  const hasReason = headerNames.some((header) =>
    reasonPatterns.some((pattern) => header.toLowerCase().includes(pattern.toLowerCase()))
  );

  const hasOpinion = headerNames.some((header) =>
    opinionPatterns.some((pattern) => header.toLowerCase().includes(pattern.toLowerCase()))
  );

  const missing = [];
  if (!hasReason) missing.push('理由系列');
  if (!hasOpinion) missing.push('回答系列');

  // 最低限必要なのは2列以上のテキストデータ
  const minimalValidation = headerNames.length >= 2;

  return {
    success: minimalValidation && (hasReason || hasOpinion || headerNames.length >= 4),
    missing,
    hasReasonColumn: hasReason,
    hasOpinionColumn: hasOpinion,
    detectedColumns: {
      reasonCandidates: headerNames.filter((h) =>
        reasonPatterns.some((p) => h.toLowerCase().includes(p.toLowerCase()))
      ),
      opinionCandidates: headerNames.filter((h) =>
        opinionPatterns.some((p) => h.toLowerCase().includes(p.toLowerCase()))
      ),
      totalColumns: headerNames.length,
    },
  };
}

/**
 * ユーザーキャッシュ無効化（後方互換性）
 * @param {string} userId - ユーザーID
 * @param {string} email - ユーザーメールアドレス
 * @param {string|null} spreadsheetId - スプレッドシートID
 * @param {boolean|string} clearPattern - パターンクリア（true='all', false=基本のみ）
 * @param {string} dbSpreadsheetId - データベーススプレッドシートID（未使用）
 */
function invalidateUserCache(userId, email, spreadsheetId, clearPattern = false, dbSpreadsheetId) {
  try {
    Logger.info('ユーザーキャッシュ無効化開始', { userId, email });

    // 基本ユーザーキャッシュクリア
    if (userId) {
      const userCacheKeys = [
        `user_${userId}`,
        `user_data_${userId}`,
        `userinfo_${userId}`,
        `unified_user_info_${userId}`,
      ];
      userCacheKeys.forEach((key) => SimpleCacheManager.remove(key));

    }

    // メールベースキャッシュクリア
    if (email) {
      const emailCacheKeys = [`email_${email}`, `unified_user_info_${email}`];
      emailCacheKeys.forEach((key) => SimpleCacheManager.remove(key));
    }

    // スプレッドシート関連キャッシュクリア
    if (spreadsheetId) {
      const spreadsheetKeys = [
        `headers_${spreadsheetId}`,
        `spreadsheet_info_${spreadsheetId}`,
        `published_data_${spreadsheetId}`,
      ];
      spreadsheetKeys.forEach((key) => SimpleCacheManager.remove(key));
    }

    Logger.info('ユーザーキャッシュ無効化完了');
  } catch (error) {
    Logger.error('ユーザーキャッシュ無効化エラー', { error: error.message });
    // エラーが発生してもシステムを停止させない
  }
}

/**
 * クリティカル更新後のキャッシュ同期（統合版）
 * @param {string} userId - ユーザーID
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string|null} oldSpreadsheetId - 古いスプレッドシートID
 * @param {string|null} newSpreadsheetId - 新しいスプレッドシートID
 */
function synchronizeCacheAfterCriticalUpdate(
  userId,
  userEmail,
  oldSpreadsheetId,
  newSpreadsheetId
) {
  try {
    console.info('🔄 キャッシュ同期開始', {
      userId,
      oldSpreadsheetId,
      newSpreadsheetId,
    });

    // 古いスプレッドシート関連のキャッシュを削除
    if (oldSpreadsheetId) {
      const oldKeys = [
        `headers_${oldSpreadsheetId}`,
        `user_data_${userId}`,
        `spreadsheet_info_${oldSpreadsheetId}`,
      ];
      oldKeys.forEach((key) => SimpleCacheManager.remove(key));
    }

    // 新しいスプレッドシート関連のキャッシュを初期化
    if (newSpreadsheetId) {
      const newKeys = [
        `headers_${newSpreadsheetId}`,
        `user_data_${userId}`,
        `spreadsheet_info_${newSpreadsheetId}`,
      ];
      newKeys.forEach((key) => SimpleCacheManager.remove(key)); // 古いキャッシュがあれば削除
    }

    Logger.info('クリティカル更新後のキャッシュ同期完了');
  } catch (error) {
    Logger.error('クリティカル更新後のキャッシュ同期エラー', { error: error.message });
    // エラーが発生してもシステムを停止させない
  }
}
