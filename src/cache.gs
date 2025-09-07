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
      // 🔧 関数オブジェクト保護：disableCacheService=true の場合はキャッシュをスキップ
      if (disableCacheService) {
        console.log('🔧 SimpleCacheManager.get: CacheService無効化モード（関数オブジェクト保護）', { key });
        if (typeof valueFn === 'function') {
          return valueFn();
        }
        return null;
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
      console.log('🔥 緊急キャッシュクリア: システム復旧開始');
      
      // Service Accountトークンクリア
      this.scriptCache.remove('SA_TOKEN_CACHE');
      this.scriptCache.remove('sheets_service_optimized');
      this.scriptCache.remove('sheets_service');
      
      // その他のキャッシュクリア
      const commonCacheKeys = ['user_config', 'form_info', 'system_status'];
      commonCacheKeys.forEach(key => this.scriptCache.remove(key));
      
      console.log('✅ 緊急キャッシュクリア: 完了');
    } catch (error) {
      console.error('❌ 緊急キャッシュクリア: エラー', error.message);
    }
  }
};

// 後方互換性のためのエイリアス
const cacheManager = SimpleCacheManager;
console.log('🗄️ シンプルCacheService永続キャッシュが初期化されました（globalThis依存排除）');

/**
 * Sheets APIサービス結果のキャッシュ（互換性維持）
 * @returns {object|null} キャッシュされたサービス情報
 */
function getSheetsServiceCached() {
  console.log('🔧 getSheetsServiceCached: 安定化版キャッシュ確認開始');
  
  // ✅ 修正: CacheServiceは関数オブジェクトを正しく保存できないため、メモリキャッシュのみ使用
  // ✅ 最適化：先にキャッシュ存在確認とヒット率向上
  const cacheKey = 'sheets_service_optimized';
  console.log('🔧 getSheetsServiceCached: キャッシュ確認', { key: cacheKey });

  // 🔧 メモリキャッシュから直接確認（関数オブジェクト保護考慮）
  try {
    // disableCacheService=true の場合はキャッシュチェックをスキップ
    const cachedService = cacheManager.get(cacheKey, null, { disableCacheService: false });
    if (cachedService !== null) {
      // 🔍 厳密な関数検証：実際の関数型と構造をチェック
      const isValidService = cachedService?.spreadsheets?.values?.append && 
                            typeof cachedService.spreadsheets.values.append === 'function' &&
                            cachedService?.spreadsheets?.values?.batchGet &&
                            typeof cachedService.spreadsheets.values.batchGet === 'function' &&
                            cachedService?.spreadsheets?.values?.update &&
                            typeof cachedService.spreadsheets.values.update === 'function';
      
      if (isValidService) {
        console.log('✅ getSheetsServiceCached: メモリキャッシュヒット（完全検証済み）');
        return cachedService;
      } else {
        console.warn('⚠️ getSheetsServiceCached: キャッシュサービス不完全（関数欠損）- 再生成実行', {
          hasAppend: typeof cachedService?.spreadsheets?.values?.append,
          hasBatchGet: typeof cachedService?.spreadsheets?.values?.batchGet,
          hasUpdate: typeof cachedService?.spreadsheets?.values?.update
        });
        // 不完全なキャッシュをクリア
        cacheManager.remove(cacheKey);
      }
    }
  } catch (error) {
    // キャッシュアクセスエラーは無視してサービスを再生成
    console.log('⚠️ キャッシュアクセスエラー（サービス再生成）:', error.message);
  }

  const result = cacheManager.get(
    cacheKey,
    () => {
      console.log('🔧 getSheetsServiceCached: 新しいサービスオブジェクト作成（キャッシュミス）');
      
      // Service Account認証確認
      let testToken;
      try {
        console.log('🔧 getSheetsServiceCached: Service Accountトークン取得開始');
        testToken = getServiceAccountTokenCached();
        console.log('🔧 getSheetsServiceCached: Service Accountトークン確認', { 
          hasToken: !!testToken,
          tokenLength: testToken ? testToken.length : 0,
          tokenPrefix: testToken ? testToken.substring(0, 20) + '...' : 'null'
        });
      } catch (tokenError) {
        console.error('🔧 getSheetsServiceCached: Service Accountトークン取得エラー詳細', {
          error: tokenError.message,
          stack: tokenError.stack,
          context: 'service_object_creation',
          errorName: tokenError.name,
          errorConstructor: tokenError.constructor.name
        });
        
        // 🚨 重要：トークン取得失敗時は不完全なサービスオブジェクトを返さない
        console.error('🚨 Service Accountトークン取得失敗により、service object構築を中止します');
        throw new Error('Service Account認証が利用できません。システム管理者に連絡してください。');
      }

      // Google Sheets APIサービスオブジェクトを返す
      console.log('🔧 getSheetsServiceCached: service object構築開始');
      
      // 🚨 実行コンテキスト情報を記録 - getUser成功/createUser失敗の原因調査
      const executionContext = {
        timestamp: new Date().toISOString(),
        stackTrace: new Error().stack.split('\n').slice(1, 4).join(' -> '),
        memoryUsage: typeof Utilities !== 'undefined' ? 'available' : 'unavailable'
      };
      console.log('🔧 getSheetsServiceCached: 実行コンテキスト', executionContext);
      
      const serviceObject = {
        baseUrl: 'https://sheets.googleapis.com/v4/spreadsheets',
        spreadsheets: {
          batchUpdate: function (params) {
            // 最新のアクセストークンを取得（トークンの期限切れ対応）
            const accessToken = getServiceAccountTokenCached();
            if (!accessToken) {
              throw new Error('Service Account token is not available');
            }

            console.log('getSheetsServiceCached.batchUpdate: API呼び出し開始', {
              spreadsheetId: params.spreadsheetId,
              requestsCount: params.requests ? params.requests.length : 0,
              hasToken: !!accessToken,
            });

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

            console.log('getSheetsServiceCached.batchUpdate: API応答', {
              responseCode: response.getResponseCode(),
              contentLength: response.getContentText().length,
            });

            if (response.getResponseCode() !== 200) {
              throw new Error(`Sheets API Error: ${response.getContentText()}`);
            }

            return JSON.parse(response.getContentText());
          },
          values: {
            batchGet: function (params) {
              // 最新のアクセストークンを取得（トークンの期限切れ対応）
              const accessToken = getServiceAccountTokenCached();
              if (!accessToken) {
                throw new Error('Service Account token is not available');
              }

              console.log('getSheetsServiceCached.batchGet: API呼び出し開始', {
                spreadsheetId: params.spreadsheetId,
                rangesCount: params.ranges ? params.ranges.length : 0,
                hasToken: !!accessToken,
              });

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

              console.log('getSheetsServiceCached.batchGet: API応答', {
                responseCode: response.getResponseCode(),
                contentLength: response.getContentText().length,
              });

              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API Error: ${response.getContentText()}`);
              }

              return JSON.parse(response.getContentText());
            },
            update: function (params) {
              // 最新のアクセストークンを取得（トークンの期限切れ対応）
              const accessToken = getServiceAccountTokenCached();
              if (!accessToken) {
                throw new Error('Service Account token is not available');
              }

              console.log('getSheetsServiceCached.update: API呼び出し開始', {
                spreadsheetId: params.spreadsheetId,
                range: params.range,
                hasToken: !!accessToken,
              });

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

              console.log('getSheetsServiceCached.update: API応答', {
                responseCode: response.getResponseCode(),
                contentLength: response.getContentText().length,
              });

              if (response.getResponseCode() !== 200) {
                throw new Error(`Sheets API Error: ${response.getContentText()}`);
              }

              return JSON.parse(response.getContentText());
            },
            append: function (params) {
              console.log('🔧 cache.gs append function called', { 
                hasParams: !!params,
                spreadsheetId: params?.spreadsheetId,
                range: params?.range
              });
              
              // 最新のアクセストークンを取得（トークンの期限切れ対応）
              const accessToken = getServiceAccountTokenCached();
              if (!accessToken) {
                console.error('🔧 cache.gs append: Service Account token is not available');
                throw new Error('Service Account token is not available');
              }

              console.log('getSheetsServiceCached.append: API呼び出し開始', {
                spreadsheetId: params.spreadsheetId,
                range: params.range,
                valuesCount: params.values ? params.values.length : 0,
                hasToken: !!accessToken,
              });

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

              console.log('getSheetsServiceCached.append: API応答', {
                responseCode: response.getResponseCode(),
                contentLength: response.getContentText().length,
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
      const isComplete = serviceObject.spreadsheets && serviceObject.spreadsheets.values &&
                        typeof serviceObject.spreadsheets.values.batchGet === 'function' &&
                        typeof serviceObject.spreadsheets.values.update === 'function' &&
                        typeof serviceObject.spreadsheets.values.append === 'function';
      
      if (isComplete) {
        console.log('✅ getSheetsServiceCached: service object構築完了（全メソッド確認済み）');
      } else {
        console.warn('🔧 getSheetsServiceCached: service object構築異常', {
          hasSpreadsheets: !!serviceObject.spreadsheets,
          hasValues: !!serviceObject.spreadsheets?.values,
          hasBatchGet: typeof serviceObject.spreadsheets?.values?.batchGet === 'function',
          hasUpdate: typeof serviceObject.spreadsheets?.values?.update === 'function', 
          hasAppend: typeof serviceObject.spreadsheets?.values?.append === 'function',
          valuesKeys: serviceObject.spreadsheets?.values ? Object.keys(serviceObject.spreadsheets.values) : []
        });
      }
      
      return serviceObject;
    },
    { 
      ttl: 900, // 15分間に延長（パフォーマンス重視）
      enableMemoization: true,
      disableCacheService: true // ✅ CacheService無効化（関数オブジェクト保護）
    }
  );
  
  // 🔧 デバッグ：result の詳細情報を出力（異常時のみ）
  const hasAllMethods = result?.spreadsheets?.values?.append && 
                       result?.spreadsheets?.values?.batchGet && 
                       result?.spreadsheets?.values?.update;
  
  if (!hasAllMethods) {
    console.log('🔧 getSheetsServiceCached: キャッシュからの戻り値詳細（異常検出）', {
      resultType: typeof result,
      isNull: result === null,
      isUndefined: result === undefined,
      hasSpreadsheets: !!result?.spreadsheets,
      spreadsheetsType: typeof result?.spreadsheets,
      hasValues: !!result?.spreadsheets?.values,
      valuesType: typeof result?.spreadsheets?.values,
      valuesKeys: result?.spreadsheets?.values ? Object.keys(result.spreadsheets.values) : [],
      resultKeys: result ? Object.keys(result) : []
    });
  }

  // ✅ 安定化：実際の関数動作確認まで行う詳細検証
  const validation = {
    hasResult: !!result,
    hasSpreadsheets: !!result?.spreadsheets,
    hasValues: !!result?.spreadsheets?.values,
    hasAppend: !!result?.spreadsheets?.values?.append,
    appendIsFunction: typeof result?.spreadsheets?.values?.append === 'function',
    hasBatchGet: !!result?.spreadsheets?.values?.batchGet,
    hasUpdate: !!result?.spreadsheets?.values?.update
  };
  
  // 🔍 完全性チェック：全必要メソッドの存在確認
  const isComplete = validation.hasResult && 
                    validation.hasSpreadsheets && 
                    validation.hasValues && 
                    validation.appendIsFunction && 
                    validation.hasBatchGet && 
                    validation.hasUpdate;
  
  // パフォーマンス向上：正常時はログ出力を削減
  if (!isComplete) {
    const missingMethods = [
      !validation.hasAppend && 'append',
      !validation.hasBatchGet && 'batchGet', 
      !validation.hasUpdate && 'update'
    ].filter(Boolean);
    
    console.log('🔧 getSheetsServiceCached: サービス検証失敗', {
      missingMethods: missingMethods,
      methodCount: missingMethods.length
    });
  }
  
  // 🚨 破損したservice objectの自動修復
  // ✅ 安定化：必要な全メソッドの存在確認
  // 🔄 無限ループ防止：修復試行回数を制限
  const retryKey = `${cacheKey}_retry_count`;
  const currentRetryCount = parseInt(cacheManager.get(retryKey) || '0');
  const maxRetryAttempts = 3;
  if (!isComplete) {
    console.error('🚨 Service object破損検出：必要メソッド欠損', {
      hasAppend: validation.appendIsFunction,
      hasBatchGet: validation.hasBatchGet,
      hasUpdate: validation.hasUpdate,
      currentRetryCount: currentRetryCount,
      maxRetryAttempts: maxRetryAttempts
    });
    
    // 🔄 無限ループ防止：リトライ回数が上限を超えた場合はエラー
    if (currentRetryCount >= maxRetryAttempts) {
      console.error('🚨 Service object修復失敗：最大リトライ回数を超過', {
        retryCount: currentRetryCount,
        maxRetryAttempts: maxRetryAttempts
      });
      // リトライカウンターをリセットして次回は再試行できるようにする
      cacheManager.remove(retryKey);
      throw new Error('Service Account認証が利用できません。システム管理者に連絡してください。');
    }
    
    // リトライ回数を増やす
    cacheManager.set(retryKey, (currentRetryCount + 1).toString(), { ttl: 300 }); // 5分でリセット
    
    // ✅ メモリキャッシュのみクリア（CacheService無効のため）
    cacheManager.remove('sheets_service');
    cacheManager.remove('sheets_service_optimized');
    console.log('🔧 破損メモリキャッシュクリア完了', {
      nextRetryCount: currentRetryCount + 1
    });
    
    // ✅ 次回呼び出しで正常なオブジェクトが生成される
    throw new Error('Service object corruption detected - please retry operation');
  }
  
  // 🎉 成功時はリトライカウンターをクリア
  if (currentRetryCount > 0) {
    cacheManager.remove(retryKey);
    console.log('✅ Service object正常生成：リトライカウンターをクリア');
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
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
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

    console.log(`📊 スプレッドシートヘッダーを取得しました: ${spreadsheetId}/${sheetName}`);
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
    console.log('🗑️ ユーザーキャッシュ無効化開始:', {
      userId,
      email,
      spreadsheetId,
      clearPattern,
    });

    // 基本ユーザーキャッシュクリア
    if (userId) {
      const userCacheKeys = [
        `user_${userId}`,
        `user_data_${userId}`,
        `userinfo_${userId}`,
        `unified_user_info_${userId}`,
      ];
      userCacheKeys.forEach((key) => SimpleCacheManager.remove(key));

      // パターンクリア（CacheServiceでは個別削除のみ対応）
      if (clearPattern) {
        console.log('パターンクリア: CacheServiceでは個別キー削除のみ対応');
      }
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

    console.log('✅ ユーザーキャッシュ無効化完了');
  } catch (error) {
    console.error('[ERROR] invalidateUserCache:', error.message);
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
    console.log('🔄 クリティカル更新後のキャッシュ同期開始:', {
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

    console.log('✅ クリティカル更新後のキャッシュ同期完了');
  } catch (error) {
    console.error('[ERROR] synchronizeCacheAfterCriticalUpdate:', error.message);
    // エラーが発生してもシステムを停止させない
  }
}

