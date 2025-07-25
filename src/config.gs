/**
 * config.gs - 軽量化版
 * 新サービスアカウントアーキテクチャで必要最小限の関数のみ
 */

const CONFIG_SHEET_NAME = 'Config';

var runtimeUserInfo = null;

// メモリ管理用の実行レベル変数
var executionUserInfoCache = null;
var lastCacheUserIdKey = null;
var executionStartTime = Date.now();
const EXECUTION_MAX_LIFETIME = 300000; // 5分間の最大実行時間

/**
 * 実行中に一度だけユーザー情報を取得して再利用する。
 * メモリ管理強化版：実行時間制限とキャッシュ自動クリア
 * @param {string} [requestUserId] - リクエスト元のユーザーID (オプション)
 * @returns {object|null} ユーザー情報
 */
function getUserInfoCached(requestUserId) {
  // 実行時間制限チェック
  if (Date.now() - executionStartTime > EXECUTION_MAX_LIFETIME) {
    console.warn('⚠️ 実行時間制限到達、キャッシュを自動クリア');
    clearExecutionUserInfoCache();
    executionStartTime = Date.now();
  }
  
  const userIdKey = requestUserId || getUserId();
  
  // キャッシュヒット条件：同じユーザーIDかつキャッシュが有効
  if (executionUserInfoCache && lastCacheUserIdKey === userIdKey) {
    return executionUserInfoCache;
  }
  
  // キャッシュミス：新規取得
  try {
    const userInfo = findUserById(userIdKey);
    if (userInfo) {
      executionUserInfoCache = userInfo;
      lastCacheUserIdKey = userIdKey;
    }
    return userInfo;
  } catch (error) {
    console.error('getUserInfoCached error:', error.message);
    // エラー時はキャッシュをクリア
    clearExecutionUserInfoCache();
    return null;
  }
}

/**
 * 実行レベルのユーザー情報キャッシュをクリア
 */
function clearExecutionUserInfoCache() {
  executionUserInfoCache = null;
  lastCacheUserIdKey = null;
  debugLog('[Memory] 実行レベルユーザー情報キャッシュをクリアしました');
}

/**
 * リソース管理を自動化するContextクラス
 * メモリ管理改善版：リソース自動解放とライフサイクル管理
 */
class ManagedExecutionContext {
  constructor(requestUserId, userInfo) {
    this.requestUserId = requestUserId;
    this.userInfo = userInfo;
    this.resources = new Map();
    this.startTime = Date.now();
    this.maxLifetime = 300000; // 5分間の最大ライフタイム
    this.isDestroyed = false;
    this.memoryUsage = { peak: 0, current: 0 };
    
    // シーツサービスを安全に作成
    try {
      this.sheetsService = createSheetsService();
      this.resources.set('sheetsService', this.sheetsService);
      this._trackMemoryUsage('sheetsService');
    } catch (error) {
      console.error('❌ サービス作成エラー:', error.message);
      this.sheetsService = getSheetsServiceCached();
      this.resources.set('sheetsService', this.sheetsService);
      this._trackMemoryUsage('sheetsService');
    }
    
    this.pendingUpdates = {};
    this.cachesToInvalidate = [];
    
    // 自動クリーンアップのタイマー設定
    this._scheduleAutoCleanup();
    
    // メモリ使用量監視の開始
    this._startMemoryMonitoring();
  }
  
  /**
   * メモリ使用量を追跡
   */
  _trackMemoryUsage(resourceKey) {
    try {
      // 大まかなメモリ使用量を計算
      const resourcesSize = this.resources.size * 100; // 概算値
      const updatesSize = Object.keys(this.pendingUpdates).length * 50;
      this.memoryUsage.current = resourcesSize + updatesSize;
      
      if (this.memoryUsage.current > this.memoryUsage.peak) {
        this.memoryUsage.peak = this.memoryUsage.current;
      }
      
      // メモリ使用量が過大になった場合の警告
      if (this.memoryUsage.current > 5000) {
        console.warn('⚠️ ExecutionContext高メモリ使用量検出:', this.memoryUsage.current);
      }
    } catch (error) {
      console.warn('メモリ使用量追跡エラー:', error.message);
    }
  }
  
  /**
   * メモリ使用量監視を開始
   */
  _startMemoryMonitoring() {
    this.memoryMonitorTimer = setTimeout(() => {
      if (!this.isDestroyed) {
        this._trackMemoryUsage('monitor');
        this._startMemoryMonitoring(); // 再帰的監視
      }
    }, 30000); // 30秒間隔
  }
  
  /**
   * リソースの自動クリーンアップを予約
   */
  _scheduleAutoCleanup() {
    this.cleanupTimer = setTimeout(() => {
      if (!this.isDestroyed) {
        console.warn('⚠️ ExecutionContext自動クリーンアップ実行:', this.requestUserId);
        this.destroy();
      }
    }, this.maxLifetime);
  }
  
  /**
   * リソースを手動で追加
   */
  addResource(key, resource) {
    if (!this.isDestroyed) {
      this.resources.set(key, resource);
      this._trackMemoryUsage(key);
    }
  }
  
  /**
   * 特定のリソースを安全に削除
   */
  removeResource(key) {
    if (this.resources.has(key)) {
      try {
        const resource = this.resources.get(key);
        if (resource && typeof resource.cleanup === 'function') {
          resource.cleanup();
        }
        this.resources.delete(key);
        debugLog(`🗑️ リソース削除: ${key}`);
      } catch (error) {
        console.warn('リソース削除エラー:', key, error.message);
      }
    }
  }
  
  /**
   * コンテキストを破棄してリソースを解放
   */
  destroy() {
    if (this.isDestroyed) return;
    
    try {
      // タイマーをクリア
      if (this.cleanupTimer) {
        clearTimeout(this.cleanupTimer);
        this.cleanupTimer = null;
      }
      
      if (this.memoryMonitorTimer) {
        clearTimeout(this.memoryMonitorTimer);
        this.memoryMonitorTimer = null;
      }
      
      // リソースを個別にクリーンアップ
      for (const [key, resource] of this.resources) {
        try {
          if (resource && typeof resource.cleanup === 'function') {
            resource.cleanup();
          }
        } catch (cleanupError) {
          console.warn(`リソースクリーンアップエラー (${key}):`, cleanupError.message);
        }
      }
      
      // リソースをクリア
      this.resources.clear();
      this.sheetsService = null;
      this.pendingUpdates = null;
      this.cachesToInvalidate = null;
      this.userInfo = null;
      
      this.isDestroyed = true;
      
      const lifetime = Date.now() - this.startTime;
      debugLog(`✅ ExecutionContextクリーンアップ完了: ${this.requestUserId} (lifetime: ${lifetime}ms, peak memory: ${this.memoryUsage.peak})`);
      
    } catch (error) {
      console.error('❌ ExecutionContextクリーンアップエラー:', error.message);
    }
  }
  
  /**
   * ヘルスチェック
   */
  isHealthy() {
    const age = Date.now() - this.startTime;
    return !this.isDestroyed && 
           this.sheetsService && 
           age < this.maxLifetime &&
           this.resources.size > 0 &&
           this.memoryUsage.current < 10000; // メモリ制限チェック
  }
  
  /**
   * メモリ統計情報を取得
   */
  getMemoryStats() {
    return {
      resourceCount: this.resources.size,
      currentMemory: this.memoryUsage.current,
      peakMemory: this.memoryUsage.peak,
      lifetime: Date.now() - this.startTime,
      isHealthy: this.isHealthy()
    };
  }
}

// グローバルなコンテキスト管理
const globalContextManager = {
  activeContexts: new Map(),
  maxConcurrentContexts: 10,
  
  register(context) {
    if (this.activeContexts.size >= this.maxConcurrentContexts) {
      console.warn('⚠️ 最大コンテキスト数到達、古いコンテキストを強制クリーンアップ');
      this._cleanupOldest();
    }
    
    this.activeContexts.set(context.requestUserId, context);
    debugLog(`📝 コンテキスト登録: ${context.requestUserId} (total: ${this.activeContexts.size})`);
  },
  
  unregister(userId) {
    if (this.activeContexts.has(userId)) {
      const context = this.activeContexts.get(userId);
      context.destroy();
      this.activeContexts.delete(userId);
      debugLog(`🗑️ コンテキスト削除: ${userId} (remaining: ${this.activeContexts.size})`);
    }
  },
  
  _cleanupOldest() {
    let oldestContext = null;
    let oldestTime = Date.now();
    
    for (const [userId, context] of this.activeContexts) {
      if (context.startTime < oldestTime) {
        oldestTime = context.startTime;
        oldestContext = userId;
      }
    }
    
    if (oldestContext) {
      this.unregister(oldestContext);
    }
  },
  
  cleanup() {
    for (const userId of this.activeContexts.keys()) {
      this.unregister(userId);
    }
  }
};

/**
 * グローバルメモリ管理機能
 */
function performGlobalMemoryCleanup() {
  try {
    console.log('🧹 グローバルメモリクリーンアップ開始');
    
    // 1. ExecutionContextの清理
    globalContextManager.cleanup();
    
    // 2. 実行レベルキャッシュのクリア
    clearExecutionUserInfoCache();
    
    // 3. CacheManagerの期限切れキャッシュクリア
    if (typeof cacheManager !== 'undefined' && cacheManager.clearExpired) {
      cacheManager.clearExpired();
    }
    
    // 4. 旧形式のランタイム変数クリア
    runtimeUserInfo = null;
    
    console.log('✅ グローバルメモリクリーンアップ完了');
    
  } catch (error) {
    console.error('❌ グローバルメモリクリーンアップエラー:', error.message);
  }
}

/**
 * 現在のユーザーのスプレッドシートを取得 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getCurrentSpreadsheet(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var userInfo = findUserById(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    // スプレッドシートIDの形式チェック
    const spreadsheetIdPattern = /^[a-zA-Z0-9-_]{44}$/;
    if (!spreadsheetIdPattern.test(userInfo.spreadsheetId)) {
      console.error('❌ 無効なスプレッドシートID形式:', userInfo.spreadsheetId);
      throw new Error(`無効なスプレッドシートID形式です: ${userInfo.spreadsheetId}`);
    }
    
    // SpreadsheetApp.openById()で権限エラーが発生する可能性があるため、事前チェック
    console.log('🔧 getCurrentSpreadsheet: スプレッドシートアクセス試行中:', userInfo.spreadsheetId);
    
    try {
      return SpreadsheetApp.openById(userInfo.spreadsheetId);
    } catch (openError) {
      console.error('❌ SpreadsheetApp.openById 権限エラー:', openError.message);
      
      // 詳細なエラー情報とデバッグ情報を提供
      console.log('🔍 権限エラー診断情報:', {
        spreadsheetId: userInfo.spreadsheetId,
        userId: requestUserId,
        userEmail: userInfo.adminEmail,
        executionContext: 'WebApp',
        errorType: openError.name,
        errorMessage: openError.message
      });
      
      if (openError.message.includes('openById') || openError.message.includes('permission') || openError.message.includes('Unexpected error')) {
        throw new Error(`スプレッドシートへのアクセス権限がありません。
原因: Webアプリの実行コンテキストでスプレッドシートにアクセスできません。
対処法: 
1. スプレッドシートの共有設定でサービスアカウントに編集者権限を付与
2. Webアプリの実行権限設定を確認
3. スプレッドシートID: ${userInfo.spreadsheetId}
詳細エラー: ${openError.message}`);
      } else {
        throw new Error(`スプレッドシートの取得に失敗しました: ${openError.message}`);
      }
    }
  } catch (e) {
    console.error('getCurrentSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートの取得に失敗しました: ' + e.message);
  }
}

/**
 * アクティブなスプレッドシートのURLを取得 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {string} スプレッドシートURL
 */
function openActiveSpreadsheet(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var ss = getCurrentSpreadsheet(requestUserId);
    return ss.getUrl();
  } catch (e) {
    console.error('openActiveSpreadsheet エラー: ' + e.message);
    throw new Error('スプレッドシートのURL取得に失敗しました: ' + e.message);
  }
}

/**
 * 現在のスクリプト実行ユーザーのユニークIDを取得する。
 * セッション分離を強化し、アカウント間の混在を防ぐ。
 * @param {string} [requestUserId] - リクエスト元のユーザーID (オプション)
 * @returns {string} 現在のユーザーのユニークID
 */
function getUserId(requestUserId) {
  if (requestUserId) return requestUserId;

  const email = Session.getActiveUser().getEmail();
  if (!email) {
    throw new Error('ユーザーIDを取得できませんでした。');
  }

  // メールアドレスベースのユニークキーを生成
  const userKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, email, Utilities.Charset.UTF_8)
    .map(function(byte) { return (byte + 256).toString(16).slice(-2); })
    .join('');

  const props = PropertiesService.getUserProperties();
  let userId = props.getProperty(userKey);

  if (!userId) {
    // 既存ユーザーを検索してセッション同期
    const userInfo = findUserByEmail(email);
    if (userInfo) {
      userId = userInfo.userId;
      console.log('既存ユーザーのセッションを復元: ' + email);
    } else {
      // 新規ユーザーの場合はメールアドレスをIDとして使用
      userId = email;
      console.log('新規ユーザーIDを生成: ' + email);
    }

    // アカウント固有のキーで保存
    props.setProperty(userKey, userId);

    // 古いキャッシュをクリア
    clearOldUserCache(email);
  }

  return userId;
}

/**
 * 古いユーザーキャッシュをクリアしてセッション混在を防ぐ
 * @param {string} currentEmail - 現在のユーザーメール
 */
function clearOldUserCache(currentEmail) {
  try {
    const props = PropertiesService.getUserProperties();

    // 古い形式のキャッシュを削除
    props.deleteProperty('CURRENT_USER_ID');

    // 現在のユーザー以外のキャッシュをクリア
    const userCache = CacheService.getUserCache();
    if (userCache) {
      userCache.removeAll(['config_v3_', 'user_', 'email_']);
    }

    console.log('古いユーザーキャッシュをクリアしました: ' + currentEmail);
  } catch (e) {
    console.warn('キャッシュクリア中にエラーが発生しましたが、処理を続行します: ' + e.message);
  }
}

// 他の関数も同様に、存在することを確認
function getConfigUserInfo(requestUserId) {
  return getUserInfoCached(requestUserId);
}

// レガシー互換性のため保持
function getUserInfo(requestUserId) {
  return getConfigUserInfo(requestUserId);
}

/**
 * 現在有効なスプレッドシートIDを取得します。(マルチテナント対応版)
 * 設定に公開中のIDがあればそれを優先します。
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {string} スプレッドシートID
 */
function getEffectiveSpreadsheetId(requestUserId) {
  verifyUserAccess(requestUserId);
  const userInfo = getUserInfo(requestUserId);
  const configJson = userInfo && userInfo.configJson
    ? JSON.parse(userInfo.configJson)
    : {};
  return configJson.publishedSpreadsheetId || userInfo.spreadsheetId;
}

/**
 * シートヘッダーを取得 (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 */
function getSheetHeaders(requestUserId, spreadsheetId, sheetName) {
  verifyUserAccess(requestUserId);
  const ss = SpreadsheetApp.openById(spreadsheetId);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastColumn = sheet.getLastColumn();
  if (lastColumn < 1) {
    console.warn(`シート '${sheetName}' に列が存在しません`);
    return [];
  }

  return sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
}

/**
 * 指定されたシートの設定情報を取得する、新しい推奨関数。(マルチテナント対応版)
 * 保存された設定があればそれを最優先し、なければ自動マッピングを試みる。
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - 設定を取得するシート名
 * @param {boolean} forceRefresh - キャッシュを無視して強制的に再取得するかどうか
 * @returns {object} 統一された設定オブジェクト
 */
function getConfig(requestUserId, sheetName, forceRefresh = false) {
  verifyUserAccess(requestUserId);
  const userId = requestUserId; // requestUserId を使用
  const userCache = CacheService.getUserCache();
  const cacheKey = 'config_v3_' + userId + '_' + sheetName;

  if (!forceRefresh) {
    const cached = userCache.get(cacheKey);
    if (cached) {
      console.log('設定キャッシュヒット: %s', cacheKey);
      return JSON.parse(cached);
    }
  }

  try {
    console.log('設定を取得中: sheetName=%s, userId=%s', sheetName, userId);
    const userInfo = getUserInfo(userId); // 依存関係を明確化
    const headers = getSheetHeaders(userId, userInfo.spreadsheetId, sheetName);

    // 1. 返却する設定オブジェクトの器を準備
    let finalConfig = {
      sheetName: sheetName,
      opinionHeader: '',
      reasonHeader: '',
      nameHeader: '',
      classHeader: '',
      showNames: false,
      showCounts: false,
      availableHeaders: headers || [],
      hasExistingConfig: false
    };

    // 2. 保存済みの設定があるか確認
    const configJson = userInfo.configJson ? JSON.parse(userInfo.configJson) : {};
    const sheetConfigKey = 'sheet_' + sheetName;
    const savedSheetConfig = configJson[sheetConfigKey];

    if (savedSheetConfig && Object.keys(savedSheetConfig).length > 0) {
      // 3.【最優先】保存済みの設定を適用する
      console.log('保存済み設定を適用します:', JSON.stringify(savedSheetConfig));
      finalConfig.hasExistingConfig = true;
      finalConfig.opinionHeader = savedSheetConfig.opinionHeader || '';
      finalConfig.reasonHeader = savedSheetConfig.reasonHeader || '';
      finalConfig.nameHeader = savedSheetConfig.nameHeader || '';
      finalConfig.classHeader = savedSheetConfig.classHeader || '';
      // showNames, showCounts, displayMode はグローバル設定に統一するため、シート固有設定からは削除

    } else if (headers && headers.length > 0) {
      // 4.【保存設定がない場合のみ】新しい自動マッピングを実行
      console.log('保存済み設定がないため、自動マッピングを実行します。');
      const guessedConfig = autoMapHeaders(headers); 
      finalConfig.opinionHeader = guessedConfig.opinionHeader || '';
      finalConfig.reasonHeader = guessedConfig.reasonHeader || '';
      finalConfig.nameHeader = guessedConfig.nameHeader || '';
      finalConfig.classHeader = guessedConfig.classHeader || '';
    }

    // 5. 最終的な設定をキャッシュに保存
    userCache.put(cacheKey, JSON.stringify(finalConfig), 3600); // 1時間キャッシュ

    console.log('getConfig: 最終設定を返します: %s', JSON.stringify(finalConfig));
    return finalConfig;

  } catch (error) {
    console.error('getConfigでエラー:', error.message, error.stack);
    throw new Error('シート設定の取得中にエラーが発生しました: ' + error.message);
  }
}

/**
 * 列ヘッダーのリストから、各項目に最もふさわしい列名を推測する（大幅強化版）
 * @param {Array<string>} headers - スプレッドシートのヘッダーリスト
 * @param {string} sheetName - シート名（データ分析用）
 * @returns {object} 推測されたマッピング結果
 */
function autoMapHeaders(headers, sheetName = null) {
  // 高精度マッピングルール（優先度順）
  const mappingRules = {
    opinionHeader: {
      exact: ['今日の学習について、あなたの考えや感想を聞かせてください', 'あなたの回答・意見', '回答・意見'],
      high: ['今日のテーマ', 'あなたの考え', '意見', '回答', '答え', '質問への回答'],
      medium: ['answer', 'response', 'opinion', 'comment', '投稿', 'コメント', '内容'],
      low: ['テキスト', 'text', '記述', '入力', '自由記述']
    },
    reasonHeader: {
      exact: ['そう考える理由や体験があれば教えてください', '理由・根拠', '根拠・理由'],
      high: ['そう考える理由', '理由', '根拠', '説明'],
      medium: ['詳細', 'reason', '理由説明', '補足'],
      low: ['なぜ', 'why', '背景', '経験']
    },
    nameHeader: {
      exact: ['名前', '氏名', 'name'],
      high: ['ニックネーム', '呼び名', '表示名'],
      medium: ['nickname', 'display_name', '投稿者名'],
      low: ['ユーザー', 'user', '投稿者', '回答者']
    },
    classHeader: {
      exact: ['あなたのクラス', 'クラス名', 'クラス'],
      high: ['組', '学年', 'class'],
      medium: ['グループ', 'チーム', 'group', 'team'],
      low: ['所属', '部門', 'section', '学校']
    }
  };

  let result = {};
  const usedHeaders = new Set();
  const headerScores = {}; // ヘッダー毎のスコア記録

  // 1. ヘッダーの前処理と分析
  const processedHeaders = headers.map((header, index) => {
    const cleaned = String(header || '').trim();
    const lower = cleaned.toLowerCase();
    
    return {
      original: header,
      cleaned: cleaned,
      lower: lower,
      index: index,
      length: cleaned.length,
      wordCount: cleaned.split(/\s+/).length,
      hasJapanese: /[\u3040-\u309F\u30A0-\u30FF\u4E00-\u9FAF]/.test(cleaned),
      hasEnglish: /[A-Za-z]/.test(cleaned),
      hasNumbers: /\d/.test(cleaned),
      isMetadata: isMetadataColumn(cleaned)
    };
  });

  // 2. 各項目に対して最適なヘッダーをスコア計算で決定
  Object.keys(mappingRules).forEach(mappingKey => {
    let bestHeader = null;
    let bestScore = 0;

    processedHeaders.forEach(headerInfo => {
      if (usedHeaders.has(headerInfo.index) || headerInfo.isMetadata) {
        return;
      }

      const score = calculateHeaderScore(headerInfo, mappingRules[mappingKey], mappingKey);
      
      if (score > bestScore) {
        bestScore = score;
        bestHeader = headerInfo;
      }
    });

    if (bestHeader && bestScore > 0.3) { // 閾値以上のスコアが必要
      result[mappingKey] = bestHeader.original;
      usedHeaders.add(bestHeader.index);
      headerScores[mappingKey] = bestScore;
    } else {
      result[mappingKey] = '';
    }
  });

  // 3. 主要な回答列が見つからない場合のフォールバック処理
  if (!result.opinionHeader) {
    const candidateHeaders = processedHeaders.filter(h => 
      !usedHeaders.has(h.index) && 
      !h.isMetadata &&
      h.cleaned.length > 0
    );

    if (candidateHeaders.length > 0) {
      // 最も長い説明的なヘッダーを選択
      const fallbackHeader = candidateHeaders.reduce((best, current) => {
        const currentPriority = calculateFallbackPriority(current);
        const bestPriority = calculateFallbackPriority(best);
        return currentPriority > bestPriority ? current : best;
      });
      
      result.opinionHeader = fallbackHeader.original;
      usedHeaders.add(fallbackHeader.index);
    }
  }

  // 4. データ内容を分析して判定精度を向上（シート名が提供された場合）
  if (sheetName) {
    try {
      const contentAnalysis = analyzeColumnContent(sheetName, processedHeaders);
      result = refineResultsWithContentAnalysis(result, contentAnalysis, processedHeaders);
    } catch (e) {
      console.warn('データ内容分析でエラーが発生しましたが、処理を続行します:', e.message);
    }
  }

  console.log('高精度自動判定結果:', JSON.stringify({
    result: result,
    scores: headerScores,
    totalHeaders: headers.length,
    usedHeaders: usedHeaders.size
  }));

  return result;
}

/**
 * ヘッダーのスコア計算（キーワードマッチング + コンテキスト分析）
 */
function calculateHeaderScore(headerInfo, rules, mappingType) {
  let score = 0;
  const header = headerInfo.lower;

  // 完全一致（最高スコア）
  if (rules.exact) {
    for (const pattern of rules.exact) {
      if (header === pattern.toLowerCase()) {
        return 1.0;
      }
      if (header.includes(pattern.toLowerCase())) {
        score = Math.max(score, 0.9);
      }
    }
  }

  // 高優先度マッチング
  if (rules.high) {
    for (const pattern of rules.high) {
      const patternLower = pattern.toLowerCase();
      if (header.includes(patternLower)) {
        score = Math.max(score, 0.8 * (patternLower.length / header.length));
      }
    }
  }

  // 中優先度マッチング
  if (rules.medium) {
    for (const pattern of rules.medium) {
      const patternLower = pattern.toLowerCase();
      if (header.includes(patternLower)) {
        score = Math.max(score, 0.6 * (patternLower.length / header.length));
      }
    }
  }

  // 低優先度マッチング
  if (rules.low) {
    for (const pattern of rules.low) {
      const patternLower = pattern.toLowerCase();
      if (header.includes(patternLower)) {
        score = Math.max(score, 0.4 * (patternLower.length / header.length));
      }
    }
  }

  // コンテキスト分析によるスコア調整
  score = adjustScoreByContext(score, headerInfo, mappingType);

  return score;
}

/**
 * コンテキストに基づくスコア調整
 */
function adjustScoreByContext(score, headerInfo, mappingType) {
  // 長い説明的なヘッダーは質問項目である可能性が高い
  if (mappingType === 'opinionHeader' && headerInfo.length > 20) {
    score *= 1.2;
  }

  // 短いヘッダーは名前やクラス項目の可能性が高い
  if ((mappingType === 'nameHeader' || mappingType === 'classHeader') && headerInfo.length < 10) {
    score *= 1.1;
  }

  // 日本語が含まれる場合の調整
  if (headerInfo.hasJapanese) {
    score *= 1.1; // 日本語項目を優先
  }

  // 数字が含まれるヘッダーの調整
  if (headerInfo.hasNumbers) {
    if (mappingType === 'classHeader') {
      score *= 1.1; // クラス番号の可能性
    } else {
      score *= 0.9; // その他では若干減点
    }
  }

  return Math.min(score, 1.0); // 最大1.0に制限
}

/**
 * メタデータ列かどうかの判定
 */
function isMetadataColumn(header) {
  const metadataPatterns = [
    'タイムスタンプ', 'timestamp', 'メールアドレス', 'email', 
    'id', 'uuid', '更新日時', '作成日時', 'created_at', 'updated_at'
  ];
  
  const headerLower = header.toLowerCase();
  return metadataPatterns.some(pattern => 
    headerLower.includes(pattern.toLowerCase())
  );
}

/**
 * フォールバック用の優先度計算
 */
function calculateFallbackPriority(headerInfo) {
  let priority = 0;
  
  // 長い説明的なヘッダーを優先
  priority += headerInfo.length * 0.1;
  
  // 日本語が含まれる場合は優先
  if (headerInfo.hasJapanese) {
    priority += 10;
  }
  
  // 質問らしいパターン
  if (headerInfo.lower.includes('ください') || headerInfo.lower.includes('何') || 
      headerInfo.lower.includes('どう') || headerInfo.lower.includes('?')) {
    priority += 15;
  }
  
  return priority;
}

/**
 * 列の実際のデータ内容を分析
 */
function analyzeColumnContent(sheetName, processedHeaders) {
  try {
    const spreadsheet = getCurrentSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) return {};

    const analysis = {};
    const lastRow = Math.min(sheet.getLastRow(), 21); // 最大20行まで分析
    
    processedHeaders.forEach((headerInfo, index) => {
      if (lastRow <= 1) return; // データがない場合

      try {
        const columnData = sheet.getRange(2, index + 1, lastRow - 1, 1).getValues()
          .map(row => String(row[0] || '').trim())
          .filter(value => value.length > 0);

        if (columnData.length === 0) return;

        analysis[index] = {
          avgLength: columnData.reduce((sum, val) => sum + val.length, 0) / columnData.length,
          maxLength: Math.max(...columnData.map(val => val.length)),
          hasLongText: columnData.some(val => val.length > 50),
          hasShortText: columnData.every(val => val.length < 20),
          containsReasonWords: columnData.some(val => 
            val.includes('なぜなら') || val.includes('理由') || val.includes('から') || val.includes('ので')
          ),
          containsClassPattern: columnData.some(val => 
            /^\d+[年組]/.test(val) || /^[A-Z]\d*$/.test(val)
          ),
          dataCount: columnData.length
        };
      } catch (e) {
        console.warn(`列${index + 1}の分析でエラー:`, e.message);
      }
    });

    return analysis;
  } catch (e) {
    console.warn('データ内容分析でエラー:', e.message);
    return {};
  }
}

/**
 * データ内容分析結果を使って判定結果を精緻化
 */
function refineResultsWithContentAnalysis(result, contentAnalysis, processedHeaders) {
  const refinedResult = { ...result };

  // 回答列の精緻化
  if (refinedResult.opinionHeader) {
    const opinionIndex = processedHeaders.findIndex(h => h.original === refinedResult.opinionHeader);
    const analysis = contentAnalysis[opinionIndex];
    
    if (analysis && analysis.avgLength < 10 && !analysis.hasLongText) {
      // 短いテキストばかりの場合、本当に回答列か確認
      const betterCandidate = findBetterOpinionColumn(contentAnalysis, processedHeaders, refinedResult);
      if (betterCandidate) {
        refinedResult.opinionHeader = betterCandidate;
      }
    }
  }

  // 理由列の精緻化
  if (refinedResult.reasonHeader) {
    const reasonIndex = processedHeaders.findIndex(h => h.original === refinedResult.reasonHeader);
    const analysis = contentAnalysis[reasonIndex];
    
    if (analysis && analysis.containsReasonWords) {
      // 理由を示すキーワードが含まれている場合、信頼度を上げる
      console.log('理由列の信頼度が高いことを確認しました');
    }
  }

  // クラス列の精緻化
  if (refinedResult.classHeader) {
    const classIndex = processedHeaders.findIndex(h => h.original === refinedResult.classHeader);
    const analysis = contentAnalysis[classIndex];
    
    if (analysis && analysis.containsClassPattern) {
      console.log('クラス列でクラス番号パターンを確認しました');
    }
  }

  return refinedResult;
}

/**
 * より適切な回答列を探す
 */
function findBetterOpinionColumn(contentAnalysis, processedHeaders, currentResult) {
  const usedHeaders = Object.values(currentResult).filter(v => v);
  
  for (let i = 0; i < processedHeaders.length; i++) {
    const header = processedHeaders[i];
    const analysis = contentAnalysis[i];
    
    if (usedHeaders.includes(header.original) || header.isMetadata) {
      continue;
    }
    
    if (analysis && (analysis.hasLongText || analysis.avgLength > 30)) {
      return header.original;
    }
  }
  
  return null;
}

/**
 * シート設定の保存、アクティブ化、最新ステータスの取得を一つのトランザクションで実行する統合関数 (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} spreadsheetId - 対象のスプレッドシートID
 * @param {string} sheetName - 対象のシート名
 * @param {object} config - 保存する設定オブジェクト
 * @returns {object} { status: 'success'|'error', message: string, [data]: object } 形式のオブジェクト
 */
function saveAndActivateSheet(requestUserId, spreadsheetId, sheetName, config) {
  verifyUserAccess(requestUserId);
  // ★★★ここから追加：引数の検証処理★★★
  if (typeof spreadsheetId !== 'string' || !spreadsheetId) {
    throw new Error('無効なspreadsheetIdです。スプレッドシートIDは必須です。');
  }
  if (typeof sheetName !== 'string' || !sheetName) {
    throw new Error('無効なsheetNameです。シート名は必須です。');
  }
  if (typeof config !== 'object' || config === null) {
    throw new Error('無効なconfigオブジェクトです。設定オブジェクトは必須です。');
  }
  // ★★★ここまで追加★★★

  try {
    console.log('saveAndActivateSheet開始: sheetName=%s', sheetName);

    // 1. 最小限のキャッシュクリア（設定保存時のみ）
    const currentUserId = requestUserId; // requestUserId を使用

    // 2-4. バッチ処理で効率化（設定保存、シート切り替え、表示オプション設定を統合）
    const displayOptions = {
      showNames: !!config.showNames,
      showCounts: config.showCounts !== undefined ? !!config.showCounts : false
    };

    // 統合処理（バッチモード使用）
    saveSheetConfig(requestUserId, spreadsheetId, sheetName, config, { 
      batchMode: true, 
      displayOptions: displayOptions 
    });
    console.log('saveAndActivateSheet: バッチ処理完了');

    // 5. 最新のステータスをキャッシュを無視して取得
    const finalStatus = getAppConfig(requestUserId);
    console.log('saveAndActivateSheet: 統合処理完了');

    // 6. 新しいスプレッドシートまたは設定変更時に自動で公開準備まで進める
    const isNewOrUpdatedForm = checkIfNewOrUpdatedForm(requestUserId, spreadsheetId, sheetName);
    if (finalStatus && isNewOrUpdatedForm) {
      console.log('saveAndActivateSheet: 新規/更新フォーム作成後の自動化開始');
      finalStatus.readyToPublish = true;
      finalStatus.autoShowPublishDialog = true;
      finalStatus.isNewForm = true;
    }

    return finalStatus;

  } catch (error) {
    console.error('saveAndActivateSheetで致命的なエラー:', error.message, error.stack);
    // クライアントには分かりやすいエラーメッセージを返す
    throw new Error('設定の保存・適用中にサーバーエラーが発生しました: ' + error.message);
  }
}



/**
 * 既存の設定で回答ボードを再公開 (Unpublished.htmlから呼び出される)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} 公開結果
 */
function republishBoard(requestUserId) {
  if (!requestUserId) {
    requestUserId = getUserId();
  }
  
  verifyUserAccess(requestUserId);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    console.log('republishBoard開始: userId=%s', requestUserId);

    const userInfo = getConfigUserInfo(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません。');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    
    // 既存の公開設定をチェック
    if (!configJson.publishedSpreadsheetId || !configJson.publishedSheetName) {
      throw new Error('公開するシートが設定されていません。管理パネルから設定を行ってください。');
    }

    // アプリを公開状態に設定
    configJson.appPublished = true;
    configJson.lastModified = new Date().toISOString();

    // 設定を保存
    updateUser(requestUserId, {
      configJson: JSON.stringify(configJson),
      lastUpdated: new Date().toISOString()
    });

    console.log('republishBoard完了: シート=%s', configJson.publishedSheetName);
    
    return {
      success: true,
      message: '回答ボードが再公開されました',
      publishedSheetName: configJson.publishedSheetName,
      publishedSpreadsheetId: configJson.publishedSpreadsheetId
    };

  } catch (error) {
    console.error('republishBoardでエラー:', error.message, error.stack);
    throw new Error('再公開処理中にエラーが発生しました: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * 新しいフォームまたは更新されたフォームかどうかを判定 (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {boolean} 新規または更新フォームの場合true
 */
function checkIfNewOrUpdatedForm(requestUserId, spreadsheetId, sheetName) {
  verifyUserAccess(requestUserId);
  try {
    const userInfo = getUserInfo(requestUserId);
    const configJson = JSON.parse(userInfo.configJson || '{}');

    // 現在のスプレッドシートIDが以前と異なる場合（新しいフォーム）
    const currentSpreadsheetId = configJson.publishedSpreadsheetId;
    if (currentSpreadsheetId !== spreadsheetId) {
      console.log('checkIfNewOrUpdatedForm: 新しいスプレッドシート検出');
      return true;
    }

    // 現在のシート名が以前と異なる場合（新しいシート）
    const currentSheetName = configJson.publishedSheetName;
    if (currentSheetName !== sheetName) {
      console.log('checkIfNewOrUpdatedForm: 新しいシート検出');
      return true;
    }

    // 最近作成されたスプレッドシートの場合（作成から30分以内）
    const createdAt = new Date(userInfo.createdAt || 0);
    const now = new Date();
    const timeDiff = now.getTime() - createdAt.getTime();
    const thirtyMinutes = 30 * 60 * 1000; // 30分をミリ秒で

    if (timeDiff < thirtyMinutes) {
      console.log('checkIfNewOrUpdatedForm: 最近作成されたフォーム検出');
      return true;
    }

    console.log('checkIfNewOrUpdatedForm: 既存フォームの設定更新');
    return false;

  } catch (error) {
    console.error('checkIfNewOrUpdatedForm error:', error.message);
    // エラーの場合は安全側に倒して新規扱い
    return true;
  }
}


/**
 *【新しい推奨関数】設定を下書きとして保存する（公開はしない） (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 * @param {object} config - 設定オブジェクト
 * @returns {object} 保存完了メッセージ
 */
function saveDraftConfig(requestUserId, sheetName, config) {
  verifyUserAccess(requestUserId);
  try {
    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error('無効なsheetNameです。シート名は必須です。');
    }
    if (typeof config !== 'object' || config === null) {
      throw new Error('無効なconfigオブジェクトです。設定オブジェクトは必須です。');
    }

    console.log('saveDraftConfig開始: sheetName=%s', sheetName);

    const userInfo = getUserInfo(requestUserId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません。');
    }

    // 設定を保存
    saveSheetConfig(requestUserId, userInfo.spreadsheetId, sheetName, config);

    // 関連キャッシュをクリア
    invalidateUserCache(userInfo.userId, userInfo.adminEmail, userInfo.spreadsheetId, false);

    console.log('saveDraftConfig: 設定保存完了');

    return {
      success: true,
      message: '設定が下書きとして保存されました。'
    };

  } catch (error) {
    console.error('saveDraftConfigで致命的なエラー:', error.message, error.stack);
    throw new Error('設定の保存中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * シートを公開状態にする（設定は既に保存済みであることを前提） (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} 最新のステータス
 */
function activateSheet(requestUserId, spreadsheetId, sheetName) {
  verifyUserAccess(requestUserId);
  const startTime = Date.now();
  try {
    if (typeof spreadsheetId !== 'string' || !spreadsheetId) {
      throw new Error('無効なspreadsheetIdです。スプレッドシートIDは必須です。');
    }
    if (typeof sheetName !== 'string' || !sheetName) {
      throw new Error('無効なsheetNameです。シート名は必須です。');
    }

    console.log('activateSheet開始: sheetName=%s', sheetName);

    // 最適化: 一度だけユーザー情報を取得してキャッシュクリア
    const currentUserId = requestUserId; // requestUserId を使用
    if (currentUserId) {
      // ユーザー情報を一度だけ取得してキャッシュ
      const userInfo = findUserById(currentUserId);
      if (userInfo) {
        // 最小限のキャッシュクリア（スプレッドシート関連のみ）
        const keysToRemove = [
          'config_v3_' + currentUserId + '_' + sheetName,
          'sheets_' + spreadsheetId,
          'data_' + spreadsheetId
        ];
        keysToRemove.forEach(key => cacheManager.remove(key));
        console.log('activateSheet: 最小限キャッシュクリア完了');
      }
    }

    // シートをアクティブ化（効率化）
    const switchResult = switchToSheet(requestUserId, spreadsheetId, sheetName);
    console.log('activateSheet: シート切り替え完了');

    // 最新のステータスを取得（キャッシュ活用）
    const finalStatus = getAppConfig(requestUserId);
    console.log('activateSheet: 公開処理完了');

    const executionTime = Date.now() - startTime;
    console.log('activateSheet完了: 実行時間 %dms', executionTime);

    return finalStatus;

  } catch (error) {
    const executionTime = Date.now() - startTime;
    console.error('activateSheetで致命的なエラー (実行時間: %dms):', executionTime, error.message, error.stack);
    throw new Error('シートの公開中にサーバーエラーが発生しました: ' + error.message);
  }
}

/**
 * スプレッドシートの列名から自動的にconfig設定を推測する (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 * @param {object} overrides - 上書き設定
 */
function autoMapSheetHeaders(requestUserId, sheetName, overrides) {
  verifyUserAccess(requestUserId);
  try {
    var headers = getSheetHeaders(requestUserId, getCurrentSpreadsheet(requestUserId).getId(), sheetName);
    if (!headers || headers.length === 0) {
      return null;
    }
    
    // 新しい高精度自動判定機能を使用
    const mappingResult = autoMapHeaders(headers, sheetName);

    // モーダル入力値による上書き
    if (overrides) {
      if (overrides.mainQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.mainQuestion.trim());
        if (match) mappingResult.opinionHeader = match;
      }
      if (overrides.reasonQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.reasonQuestion.trim());
        if (match) mappingResult.reasonHeader = match;
      }
      if (overrides.nameQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.nameQuestion.trim());
        if (match) mappingResult.nameHeader = match;
      }
      if (overrides.classQuestion) {
        const match = headers.find(h => String(h).trim() === overrides.classQuestion.trim());
        if (match) mappingResult.classHeader = match;
      }
    }

    // 従来のフォーマットに変換（後方互換性のため）
    var mapping = {
      mainHeader: mappingResult.opinionHeader || '',
      rHeader: mappingResult.reasonHeader || '',
      nameHeader: mappingResult.nameHeader || '',
      classHeader: mappingResult.classHeader || ''
    };
    
    debugLog('高精度自動マッピング結果 for ' + sheetName + ': ' + JSON.stringify(mapping));
    return mapping;
    
  } catch (e) {
    console.error('autoMapSheetHeaders エラー: ' + e.message);
    return null;
  }
}

/**
 * スプレッドシートURLを追加してシート検出を実行 (マルチテナント対応版)
 * Unpublished.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} url - スプレッドシートURL
 */
function addSpreadsheetUrl(requestUserId, url) {
  verifyUserAccess(requestUserId);
  try {
    var spreadsheetId = url.match(/\/d\/([a-zA-Z0-9_-]+)/)[1];
    if (!spreadsheetId) {
      throw new Error('無効なスプレッドシートURLです。');
    }

    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }

    // サービスアカウントをスプレッドシートに追加
    addServiceAccountToSpreadsheet(spreadsheetId);

    // 公開設定をリセットしつつユーザー情報を更新
    var configJson = userInfo.configJson ? JSON.parse(userInfo.configJson) : {};
    configJson.publishedSpreadsheetId = spreadsheetId;
    configJson.publishedSheetName = '';
    configJson.appPublished = false;

    // スプレッドシートからフォームURLを自動検出（可能な場合）
    var formUrl = null;
    try {
      formUrl = detectFormUrlFromSpreadsheet(spreadsheetId);
      console.log('スプレッドシートからフォーム検出:', formUrl ? 'あり' : 'なし');
      
      // 検出されたフォームURLをconfigJsonにも保存
      if (formUrl) {
        configJson.formUrl = formUrl;
      }
    } catch (formDetectionError) {
      console.warn('フォーム自動検出でエラー:', formDetectionError.message);
    }

    var updateData = {
      spreadsheetId: spreadsheetId,
      spreadsheetUrl: url,
      configJson: JSON.stringify(configJson)
    };

    // フォームURLが検出できた場合は一緒に更新
    if (formUrl) {
      updateData.formUrl = formUrl;
    }

    updateUser(currentUserId, updateData);

    // シートリストを即座に取得
    var sheets = [];
    try {
      sheets = getSheetsList(currentUserId);
      console.log('シート検出完了:', {
        spreadsheetId: spreadsheetId,
        sheetCount: sheets.length,
        sheets: sheets.map(s => s.name)
      });
    } catch (sheetError) {
      console.warn('シートリスト取得でエラー:', sheetError.message);
      // シートリスト取得に失敗してもスプレッドシート追加は成功とする
    }

    // 必要最小限のキャッシュ無効化
    invalidateUserCache(currentUserId, userInfo.adminEmail, spreadsheetId, true);

    var message = 'スプレッドシートが正常に追加されました。';
    if (formUrl) {
      message += ' 関連するフォームも自動で連携されました。';
    }
    
    return { 
      status: 'success', 
      message: message, 
      sheets: sheets,
      spreadsheetId: spreadsheetId,
      autoSelectFirst: sheets.length > 0 ? sheets[0].name : null,
      needsRefresh: true, // UI側でのリフレッシュが必要
      formDetected: !!formUrl
    };
  } catch (e) {
    console.error('addSpreadsheetUrl エラー: ' + e.message);
    throw new Error('スプレッドシートの追加に失敗しました: ' + e.message);
  }
}

/**
 * フォームURLを追加 (外部リソース連携)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} url - GoogleフォームのURL
 */
function addFormUrl(requestUserId, url) {
  verifyUserAccess(requestUserId);
  try {
    // GoogleフォームのIDを抽出
    var formIdMatch = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (!formIdMatch) {
      throw new Error('無効なGoogleフォームURLです。');
    }
    var formId = formIdMatch[1];

    var currentUserId = requestUserId;

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }

    // フォームURLをユーザー情報に追加
    updateUser(currentUserId, {
      formUrl: url
    });

    // キャッシュを無効化
    invalidateUserCache(currentUserId, userInfo.adminEmail, userInfo.spreadsheetId, true);

    console.log('✅ フォームURL追加完了: ' + currentUserId);
    return {
      status: 'success',
      message: 'フォームが追加されました。回答ボードでフォームを開くことができるようになりました。'
    };

  } catch (e) {
    console.error('addFormUrl エラー: ' + e.message);
    throw new Error('フォームの追加に失敗しました: ' + e.message);
  }
}

/**
 * スプレッドシートからフォームURLを自動検出
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string|null} フォームURL（見つからない場合はnull）
 */
function detectFormUrlFromSpreadsheet(spreadsheetId) {
  try {
    // スプレッドシートを開く
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheets = spreadsheet.getSheets();
    
    // 各シートでフォームへのリンクを探す
    for (var i = 0; i < sheets.length; i++) {
      var sheet = sheets[i];
      
      try {
        // シートに関連付けられたフォームがあるかチェック
        var form = sheet.getFormUrl();
        if (form) {
          console.log('✅ フォーム自動検出成功 (getFormUrl):', form);
          return form;
        }
      } catch (formError) {
        // getFormUrl()がエラーになる場合があるのでcatchする
      }
      
      // 手動でセルの内容からフォームURLを探す（A1:Z10の範囲）
      try {
        var range = sheet.getRange('A1:Z10');
        var values = range.getValues();
        
        for (var row = 0; row < values.length; row++) {
          for (var col = 0; col < values[row].length; col++) {
            var cellValue = values[row][col];
            if (typeof cellValue === 'string' && cellValue.includes('docs.google.com/forms/')) {
              // GoogleフォームのURLらしき文字列を見つけた
              var urlMatch = cellValue.match(/https:\/\/docs\.google\.com\/forms\/d\/[a-zA-Z0-9_-]+/);
              if (urlMatch) {
                console.log('✅ フォーム自動検出成功 (セル内容):', urlMatch[0]);
                return urlMatch[0];
              }
            }
          }
        }
      } catch (cellError) {
        console.warn('セル内容の検索でエラー:', cellError.message);
      }
    }
    
    console.log('ℹ️ フォームURL検出: 見つかりませんでした');
    return null;
    
  } catch (error) {
    console.error('detectFormUrlFromSpreadsheet エラー:', error.message);
    return null;
  }
}

/**
 * アクティブシートをクリア（公開停止） (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function unpublishBoard(requestUserId) {
  console.log('🔧 unpublishBoard: 公開停止処理開始', { requestUserId });
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId; // requestUserId を使用

    var userInfo = findUserById(currentUserId);
    if (!userInfo) {
      console.error('❌ ユーザー情報が見つかりません:', currentUserId);
      throw new Error('ユーザー情報が見つかりません。');
    }
    
    console.log('✅ ユーザー情報取得成功:', {
      userId: currentUserId,
      hasConfigJson: !!userInfo.configJson,
      adminEmail: userInfo.adminEmail
    });

    var configJson = JSON.parse(userInfo.configJson || '{}');
    console.log('🔍 公開停止前の設定:', {
      publishedSheet: configJson.publishedSheet,
      publishedSheetName: configJson.publishedSheetName,
      publishedSpreadsheetId: configJson.publishedSpreadsheetId,
      appPublished: configJson.appPublished
    });

    // 公開状態のクリア（データソースとシート選択は保持）
    configJson.publishedSheet = ''; // 後方互換性のため残す
    configJson.publishedSheetName = ''; // 正しいプロパティ名
    configJson.publishedSpreadsheetId = ''; // スプレッドシートIDもクリア
    configJson.appPublished = false; // 公開停止
    
    // 回答ボード連携の解除（フォーム情報は保持）
    configJson.setupStatus = 'reconfiguring'; // ステップ1-3を全オープンにするための状態
    configJson.setupStep = 1; // ステップ1にリセット（公開終了後はセットアップから再開）
    
    // 列設定のみクリア（データソース・シート選択は保持）
    configJson.opinionHeader = '';
    configJson.nameHeader = '';
    configJson.reasonHeader = '';
    configJson.classHeader = '';
    configJson.timestampHeader = '';
    
    // 表示設定のクリア
    configJson.showNames = false;
    configJson.showCounts = false;
    configJson.highlightMode = false;
    
    // カラムマッピングのクリア
    if (configJson.columnMappings) {
      configJson.columnMappings = {};
    }
    
    // 注意: spreadsheetId, activeSheetName, formUrl等のデータソース情報は保持
    // 注意: formCreated状態も保持（ユーザーが作成したフォームを維持）

    console.log('🧹 公開停止: 設定を完全クリア完了');

    // データベースを更新
    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    // キャッシュを無効化して即座にUIに反映
    try {
      if (typeof invalidateUserCache === 'function') {
        invalidateUserCache(currentUserId, null, null, false);
      }
    } catch (cacheError) {
      console.warn('キャッシュ無効化でエラーが発生しましたが、処理を続行します:', cacheError.message);
    }

    console.log('回答ボードの公開を正常に停止しました - ユーザーID:', currentUserId);

    // 最新のステータスを取得して返す（UI更新のため）
    const updatedStatus = getAppConfig(requestUserId);
    return {
      success: true,
      message: '✅ 回答ボードの公開を停止しました。',
      timestamp: new Date().toISOString(),
      ...updatedStatus
    };
  } catch (e) {
    console.error('unpublishBoard エラー: ' + e.message);
    throw new Error('回答ボードの公開停止に失敗しました: ' + e.message);
  }
}

/**
 * 表示オプション設定（統合版：Optimized機能統合）
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} displayOptions - 表示オプション
 * @param {object} options - オプション設定
 * @param {object} options.sheetsService - 共有SheetsService（最適化用）
 * @param {object} options.userInfo - 事前取得済みユーザー情報（最適化用）
 */
function setDisplayOptions(requestUserId, displayOptions, options = {}) {
  verifyUserAccess(requestUserId);
  try {
    var currentUserId = requestUserId;

    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    var userInfo = options.userInfo || findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません。');
    }

    var configJson = JSON.parse(userInfo.configJson || '{}');
    
    // 各オプションを個別にチェックして設定（undefinedの場合は既存値を保持）
    if (displayOptions.showNames !== undefined) {
      configJson.showNames = displayOptions.showNames;
    }
    if (displayOptions.showCounts !== undefined) {
      configJson.showCounts = displayOptions.showCounts;
    }
    if (displayOptions.displayMode !== undefined) {
      configJson.displayMode = displayOptions.displayMode;
    } else if (displayOptions.showNames !== undefined) {
      // showNamesからdisplayModeを設定（後方互換性）
      configJson.displayMode = displayOptions.showNames ? 'named' : 'anonymous';
    }
    
    configJson.lastModified = new Date().toISOString();

    console.log('setDisplayOptions: 設定更新', displayOptions);

    updateUser(currentUserId, {
      configJson: JSON.stringify(configJson)
    });

    return '表示オプションを保存しました。';
  } catch (e) {
    console.error('setDisplayOptions エラー: ' + e.message);
    throw new Error('表示オプションの保存に失敗しました: ' + e.message);
  }
}

/**
 * 管理者からボードを作成 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function createBoardFromAdmin(requestUserId) {
  verifyUserAccess(requestUserId);
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userId = requestUserId; // requestUserId を使用

    // フォームとスプレッドシートを作成
    var formAndSsInfo = createStudyQuestForm(activeUserEmail, userId);

    // 中央データベースにユーザー情報を登録
    var initialConfig = {
      formUrl: formAndSsInfo.viewFormUrl || formAndSsInfo.formUrl,
      editFormUrl: formAndSsInfo.editFormUrl,
      createdAt: new Date().toISOString(),
      publishedSheet: formAndSsInfo.sheetName, // 作成時に公開シートを設定
      appPublished: true, // 作成時に公開状態にする
      showNames: false, // グローバル設定の初期値
      showCounts: false, // グローバル設定の初期値
      displayMode: 'anonymous' // グローバル設定の初期値
    };

    var userData = {
      userId: userId,
      adminEmail: activeUserEmail,
      spreadsheetId: formAndSsInfo.spreadsheetId,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      createdAt: new Date().toISOString(),
      configJson: JSON.stringify(initialConfig),
      lastAccessedAt: new Date().toISOString(),
      isActive: 'true'
    };

    createUser(userData);

    // 成功レスポンスを返す
    var appUrls = generateUserUrls(userId);
    return {
      status: 'success',
      message: '新しいボードが作成され、公開されました！',
      formUrl: formAndSsInfo.formUrl,
      spreadsheetUrl: formAndSsInfo.spreadsheetUrl,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      formTitle: formAndSsInfo.formTitle // フォームタイトルも返す
    };
  } catch (e) {
    console.error('createBoardFromAdmin エラー: ' + e.message);
    throw new Error('ボードの作成に失敗しました: ' + e.message);
  }
}

/**
 * 既存ボード情報を取得 (マルチテナント対応版)
 * Login.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getExistingBoard(requestUserId) {
  // 新規ユーザー（requestUserIdがundefinedまたはnull）の場合はverifyUserAccessをスキップ
  if (requestUserId) {
    verifyUserAccess(requestUserId);
  }
  try {
    var activeUserEmail = Session.getActiveUser().getEmail();
    var userInfo = findUserByEmail(activeUserEmail);

    if (userInfo && isTrue(userInfo.isActive)) {
      var appUrls = generateUserUrls(userInfo.userId);
      return {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl
      };
    } else if (userInfo && String(userInfo.isActive).toLowerCase() === 'false') {
      return {
        status: 'setup_required',
        userId: userInfo.userId
      };
    } else {
      return {
        status: 'new_user'
      };
    }
  } catch (e) {
    console.error('getExistingBoard エラー: ' + e.message);
    return { status: 'error', message: '既存ボード情報の取得に失敗しました: ' + e.message };
  }
}

/**
 * ユーザー認証を検証 (マルチテナント対応版)
 * Login.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function verifyUserAuthentication(requestUserId) {
  // 新規ユーザー（requestUserIdがundefinedまたはnull）の場合はverifyUserAccessをスキップ
  if (requestUserId) {
    verifyUserAccess(requestUserId);
  }
  try {
    var email = Session.getActiveUser().getEmail();
    if (email) {
      // ドメイン制限チェック
      var domainInfo = getDeployUserDomainInfo();
      if (domainInfo.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
        console.warn('Domain access denied:', domainInfo.currentDomain, 'vs', domainInfo.deployDomain);
        return { 
          authenticated: false, 
          email: null, 
          error: `ドメインアクセスが制限されています。許可されたドメイン: ${domainInfo.deployDomain}, 現在のドメイン: ${domainInfo.currentDomain}` 
        };
      }
      
      return { authenticated: true, email: email };
    } else {
      return { authenticated: false, email: null };
    }
  } catch (e) {
    console.error('verifyUserAuthentication エラー: ' + e.message);
    return { authenticated: false, email: null, error: e.message };
  }
}

/**
 * セッションをリセットして新しいアカウント選択を促す (マルチテナント対応版)
 * SharedUtilities のアカウント切り替え機能から呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {{status:'success'|'error',message:string}}
 */
function resetUserAuthentication(requestUserId) {
  // 新規ユーザー（requestUserIdがundefinedまたはnull）の場合はverifyUserAccessをスキップ
  if (requestUserId) {
    verifyUserAccess(requestUserId);
  }
  try {
    var email = Session.getActiveUser().getEmail();
    if (typeof cleanupSessionOnAccountSwitch === 'function' && email) {
      cleanupSessionOnAccountSwitch(email);
    }
    return { status: 'success', message: 'ユーザー認証がリセットされました' };
  } catch (e) {
    console.error('resetUserAuthentication エラー: ' + e.message);
    return { status: 'error', message: 'ユーザー認証のリセットに失敗しました: ' + e.message };
  }
}

/**
 * スプレッドシートの列名から自動的にconfig設定を推測する (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 * @param {object} overrides - 上書き設定
 */
/**
 * 指定されたシートのヘッダー情報と既存設定をまとめて取得します。(マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} { allHeaders: Array<string>, guessedConfig: object, existingConfig: object }
 */
function getSheetDetails(requestUserId, spreadsheetId, sheetName) {
  // パラメータ型検証（従来版）
  console.log('🔍 getSheetDetails (userId version) parameter validation:', {
    requestUserIdType: typeof requestUserId,
    spreadsheetIdType: typeof spreadsheetId,
    sheetNameType: typeof sheetName,
    userIdIsString: typeof requestUserId === 'string'
  });
  
  if (typeof requestUserId === 'object') {
    console.error('❌ CRITICAL: getSheetDetails received object as requestUserId parameter');
    console.error('❌ This suggests you meant to call getSheetDetailsFromContext instead');
    throw new Error('Parameter mismatch: Use getSheetDetailsFromContext for context-based calls');
  }
  
  // 強化されたパラメータ検証
  if (typeof requestUserId !== 'string' || requestUserId.trim() === '') {
    console.error('❌ CRITICAL: Invalid requestUserId parameter:', { requestUserId, type: typeof requestUserId });
    throw new Error('requestUserId must be a non-empty string');
  }
  
  // spreadsheetIdが提供されている場合、その形式を検証
  if (spreadsheetId !== null && spreadsheetId !== undefined) {
    if (typeof spreadsheetId !== 'string') {
      console.error('❌ CRITICAL: Invalid spreadsheetId parameter type:', { spreadsheetId, type: typeof spreadsheetId });
      throw new Error('spreadsheetId must be a string or null/undefined');
    }
    // Google Sheets IDの基本的な形式チェック（44文字の英数字とハイフン、アンダースコア）
    if (spreadsheetId.length > 0 && !/^[a-zA-Z0-9_-]{10,100}$/.test(spreadsheetId)) {
      console.error('❌ CRITICAL: Invalid spreadsheetId format:', { 
        spreadsheetId, 
        length: spreadsheetId.length,
        pattern: 'Expected: 10-100 chars of [a-zA-Z0-9_-]'
      });
      throw new Error('spreadsheetId has invalid format for Google Sheets ID');
    }
  }
  
  verifyUserAccess(requestUserId);
  try {
    // sheetNameのバリデーションとフォールバック処理
    if (!sheetName || (typeof sheetName === 'string' && sheetName.trim() === '')) {
      console.warn('⚠️ getSheetDetails: sheetNameが空です。フォールバック処理を実行中...');
      
      // ユーザー情報から publishedSheetName を取得
      try {
        const userInfo = getCachedUserInfo(requestUserId);
        if (userInfo && userInfo.configJson) {
          const configJson = JSON.parse(userInfo.configJson);
          const fallbackSheetName = configJson.publishedSheetName;
          if (fallbackSheetName && typeof fallbackSheetName === 'string' && fallbackSheetName.trim() !== '') {
            console.log('✅ configJsonからフォールバックシート名を使用:', fallbackSheetName);
            sheetName = fallbackSheetName;
          } else {
            throw new Error('getSheetDetails: sheetNameが必須ですが、configJsonからも取得できませんでした');
          }
        } else {
          throw new Error('getSheetDetails: sheetNameが必須ですが、ユーザー情報からも取得できませんでした');
        }
      } catch (fallbackError) {
        console.error('❌ フォールバック処理に失敗:', fallbackError.message);
        throw new Error('sheetNameは必須です');
      }
    }
    var targetId = spreadsheetId || getEffectiveSpreadsheetId(requestUserId);
    if (!targetId) {
      throw new Error('spreadsheetIdが取得できません');
    }
    
    // targetIdがspreadsheetIdではなくsheetNameになっていないかチェック
    if (typeof targetId === 'string' && !/^[a-zA-Z0-9_-]{10,100}$/.test(targetId)) {
      console.error('❌ CRITICAL: targetId appears to be a sheet name instead of spreadsheet ID:', { 
        targetId, 
        sheetName,
        length: targetId.length,
        providedSpreadsheetId: spreadsheetId
      });
      throw new Error(`無効なspreadsheetID形式です: "${targetId}" はシート名のようです。パラメータの順序を確認してください。`);
    }
    // SpreadsheetApp.openById()の代わりにSheets APIを使用（権限問題回避）
    console.log('🔧 Sheets APIを使用してヘッダーを取得中:', { targetId, sheetName });
    
    let headers = [];
    try {
      // Sheets APIサービスを取得
      const sheetsService = getSheetsServiceCached();
      
      // ヘッダー行を取得（1行目）
      const range = `'${sheetName}'!1:1`;
      const batch = batchGetSheetsData(sheetsService, targetId, [range]);
      
      if (batch && batch.valueRanges && batch.valueRanges[0] && batch.valueRanges[0].values) {
        headers = batch.valueRanges[0].values[0] || [];
        console.log('✅ Sheets APIでヘッダー取得成功:', headers.length, '列');
      } else {
        console.warn('⚠️ ヘッダー行が空または取得できませんでした');
        headers = [];
      }
    } catch (apiError) {
      console.error('❌ Sheets API取得エラー:', apiError.message);
      
      // フォールバック: SpreadsheetApp.openById()を試行（権限がある場合のみ）
      console.warn('⚠️ フォールバック: SpreadsheetApp.openById()を試行');
      try {
        const ss = SpreadsheetApp.openById(targetId);
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          throw new Error('シートが見つかりません: ' + sheetName);
        }
        
        const lastColumn = sheet.getLastColumn();
        if (lastColumn < 1) {
          throw new Error(`シート '${sheetName}' に列が存在しません`);
        }
        headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0] || [];
        console.log('✅ SpreadsheetApp フォールバック成功');
      } catch (spreadsheetError) {
        console.error('❌ SpreadsheetApp フォールバックも失敗:', spreadsheetError.message);
        throw new Error(`ヘッダー取得に失敗しました。Sheets API: ${apiError.message}, SpreadsheetApp: ${spreadsheetError.message}`);
      }
    }
    const guessed = autoMapHeaders(headers);

    let existing = {};
    try {
      existing = getConfig(requestUserId, sheetName, true) || {};
    } catch (e) {
      console.warn('getConfig failed in getSheetDetails:', e.message);
    }

    return {
      allHeaders: headers,
      guessedConfig: guessed,
      existingConfig: existing
    };

  } catch (error) {
    console.error('getSheetDetails error:', error.message);
    throw new Error('シート情報の取得に失敗しました: ' + error.message);
  }
}

// =================================================================
// 関数統合完了: Optimized/Batch機能は基底関数に統合済み
// =================================================================

// =================================================================
// PHASE3 OPTIMIZATION: トランザクション型実行アーキテクチャ
// =================================================================

/**
 * 実行コンテキストを作成（リソース一括作成・管理）- 最適化対応版
 * @param {string} requestUserId - ユーザーID
 * @param {object} options - オプション設定（後方互換性のため）
 * @param {object} options.reuseService - 既存のSheetsServiceを再利用
 * @param {object} options.reuseUserInfo - 既存のユーザー情報を再利用
 * @returns {object} 実行コンテキスト
 */
function createExecutionContext(requestUserId, options = {}) {
  const startTime = new Date().getTime();
  console.log('🚀 ExecutionContext作成開始: userId=%s', requestUserId);
  
  try {
    // 1. 共有リソースの取得（最適化：既存リソース再利用対応）
    const originalSheetsService = options.reuseService || getSheetsServiceCached();
    let userInfo = options.reuseUserInfo || getCachedUserInfo(requestUserId);

    if (!userInfo) {
      console.warn('getCachedUserInfo miss: userId=%s', requestUserId);
      userInfo = findUserByIdFresh(requestUserId);

      if (userInfo) {
        _executionUserInfoCache = { userId: requestUserId, userInfo };
        console.log('✅ findUserByIdFresh success: cached for execution');
      } else {
        console.error('❌ findUserByIdFresh failed for userId=%s', requestUserId);
        throw new Error('ユーザー情報が取得できません。ユーザー登録が完了しているか、データベース設定を確認してください');
      }
    }
    
    // SheetsServiceオブジェクトの安全なコピー - 関数のクロージャスコープを保持
    if (!originalSheetsService || !originalSheetsService.baseUrl || !originalSheetsService.accessToken) {
      throw new Error('SheetsServiceオブジェクトが無効です。baseUrlまたはaccessTokenが見つかりません');
    }
    
    // 重要: 新しいサービスインスタンスを作成してクロージャスコープの問題を回避
    let sheetsService = createSheetsService(originalSheetsService.accessToken);
    
    // フォールバック検証: 作成されたサービスが有効であることを確認
    if (!sheetsService || !sheetsService.baseUrl || !sheetsService.accessToken) {
      console.error('❌ SheetsService作成失敗、オリジナルサービスを使用');
      // 最後の手段：オリジナルサービスを直接使用（参照の問題は発生するが機能する）
      sheetsService = originalSheetsService;
    }
    
    console.log('DEBUG: SheetsService安全コピー完了 - baseUrl存在:', !!sheetsService.baseUrl);
    
    // 2. 実行コンテキスト構築（既存プロパティ名を完全保持）
    const context = {
      // 基本情報
      requestUserId: requestUserId,
      startTime: startTime,
      
      // 共有リソース
      sheetsService: sheetsService, // 関数参照を保持した安全なコピー
      userInfo: JSON.parse(JSON.stringify(userInfo)), // Deep copy（userInfoは単純オブジェクトのため有効）
      
      // 変更トラッキング
      pendingUpdates: {},
      configChanges: {},
      hasChanges: false,
      
      // パフォーマンス情報
      stats: {
        sheetsServiceCreations: options.reuseService ? 0 : 1, // 再利用時は0
        dbQueries: options.reuseUserInfo ? 0 : 1, // 再利用時は0
        cacheHits: 0,
        operationsCount: 0
      }
    };
    // デバッグ: SheetsService設定完了（セキュリティのため詳細は非表示）
    console.log('DEBUG: context.sheetsService initialized successfully');
    
    const endTime = new Date().getTime();
    console.log('✅ ExecutionContext作成完了: %dms (リソース再利用: Service=%s, UserInfo=%s)', 
      endTime - startTime, !!options.reuseService, !!options.reuseUserInfo);
    
    return context;
    
  } catch (error) {
    console.error('❌ ExecutionContext作成エラー:', error.message);
    throw new Error('実行コンテキストの初期化に失敗しました: ' + error.message);
  }
}

/**
 * 最適化版updateUser（インメモリ更新）
 * @param {object} context - 実行コンテキスト
 * @param {object} updateData - 更新データ
 */
function updateUserOptimized(context, updateData) {
  console.log('💾 updateUserOptimized: 変更をコンテキストに蓄積');
  
  // 変更をpendingUpdatesに蓄積（DB書き込みはしない）
  context.pendingUpdates = {...context.pendingUpdates, ...updateData};
  
  // メモリ内のuserInfoも即座に更新（後続処理で使用可能）
  context.userInfo = {...context.userInfo, ...updateData};
  context.hasChanges = true;
  context.stats.operationsCount++;
  
  console.log('📊 蓄積された変更数: %d', Object.keys(context.pendingUpdates).length);
}

/**
 * 蓄積された全変更を一括でDBにコミット
 * @param {object} context - 実行コンテキスト
 */
function commitAllChanges(context) {
  const startTime = new Date().getTime();
  console.log('💽 commitAllChanges: 一括DB書き込み開始');
  
  if (!context.hasChanges || Object.keys(context.pendingUpdates).length === 0) {
    console.log('📝 変更なし: DB書き込みをスキップ');
    return;
  }
  
  try {
    // Phase 1: DB更新実行
    updateUserDirect(context.sheetsService, context.requestUserId, context.pendingUpdates);
    
    // Phase 2: アトミックなキャッシュ無効化（DB更新成功後に即座に実行）
    try {
      console.log('🗑️ トランザクション後のキャッシュ無効化開始');
      
      // 1. ユーザー関連キャッシュの無効化
      invalidateUserCacheTransaction(context.requestUserId, context.userInfo.adminEmail, 
                                    context.userInfo.spreadsheetId);
      
      // 2. 実行レベルキャッシュのクリア
      clearExecutionUserInfoCache();
      
      // 3. SheetsServiceキャッシュの部分無効化（必要に応じて）
      if (context.pendingUpdates.spreadsheetId) {
        // スプレッドシートIDが変更された場合は関連キャッシュをクリア
        cacheManager.clearByPattern(`batchGet_${context.pendingUpdates.spreadsheetId}_*`);
      }
      
      console.log('✅ キャッシュ無効化完了');
      
    } catch (cacheError) {
      // キャッシュエラーは致命的ではないが、ログに記録
      console.warn('⚠️ キャッシュ無効化で軽微なエラー:', cacheError.message);
    }
    
    const endTime = new Date().getTime();
    console.log('✅ 一括DB書き込み+キャッシュ無効化完了: %dms, 変更項目数: %d', 
      endTime - startTime, Object.keys(context.pendingUpdates).length);
    
    // 統計更新
    context.stats.dbQueries++; // コミット時の1回をカウント
    
    // コンテキストの変更状態をリセット
    context.hasChanges = false;
    context.pendingUpdates = {};
    
  } catch (error) {
    console.error('❌ 一括DB書き込みエラー:', error.message);
    throw new Error('設定の保存に失敗しました: ' + error.message);
  }
}

/**
 * トランザクション用の安全なキャッシュ無効化
 * @param {string} userId - ユーザーID
 * @param {string} userEmail - ユーザーのメールアドレス
 * @param {string} spreadsheetId - スプレッドシートID
 */
function invalidateUserCacheTransaction(userId, userEmail, spreadsheetId) {
  try {
    // 1. ユーザー情報キャッシュの無効化
    if (typeof invalidateUserCache === 'function') {
      invalidateUserCache(userId, userEmail, spreadsheetId, false);
    }
    
    // 2. CacheManagerを使用した関連キャッシュの無効化
    if (typeof cacheManager !== 'undefined' && cacheManager.remove) {
      // ユーザー関連のキーを無効化
      const userCacheKeys = [
        `user_${userId}`,
        `userInfo_${userId}`,
        `config_${userId}`,
        `appUrls_${userId}`
      ];
      
      userCacheKeys.forEach(key => {
        try {
          cacheManager.remove(key);
        } catch (keyError) {
          console.warn('キャッシュキー削除エラー:', key, keyError.message);
        }
      });
      
      // スプレッドシート関連のキーを無効化
      if (spreadsheetId) {
        try {
          cacheManager.clearByPattern(`sheets_${spreadsheetId}_*`);
          cacheManager.clearByPattern(`batchGet_${spreadsheetId}_*`);
        } catch (patternError) {
          console.warn('パターンベースキャッシュ削除エラー:', patternError.message);
        }
      }
    }
    
    console.log('✅ トランザクション用キャッシュ無効化完了');
    
  } catch (error) {
    console.warn('⚠️ トランザクション用キャッシュ無効化エラー:', error.message);
    // エラーは再スローしない（トランザクションを破綻させない）
  }
}

/**
 * 既存のupdateUser内部実装を直接実行（SheetsService再利用版）
 * @param {object} sheetsService - 再利用するSheetsService
 * @param {string} userId - ユーザーID
 * @param {object} updateData - 更新データ
 */
function updateUserDirect(sheetsService, userId, updateData) {
  var props = PropertiesService.getScriptProperties();
  var dbId = props.getProperty('DATABASE_SPREADSHEET_ID');
  var sheetName = 'Users';
  
  // 現在のデータを取得（提供されたSheetsServiceを使用）
  var data = batchGetSheetsData(sheetsService, dbId, ["'" + sheetName + "'!A:H"]);
  var values = data.valueRanges[0].values || [];
  
  if (values.length === 0) {
    throw new Error('データベースが空です');
  }
  
  var headers = values[0];
  var userIdIndex = headers.indexOf('userId');
  var rowIndex = -1;
  
  // ユーザーの行を特定
  for (var i = 1; i < values.length; i++) {
    if (values[i][userIdIndex] === userId) {
      rowIndex = i + 1; // 1-based index
      break;
    }
  }
  
  if (rowIndex === -1) {
    throw new Error('更新対象のユーザーが見つかりません');
  }
  
  // 更新データを適用
  var updateRequests = [];
  Object.keys(updateData).forEach(function(field) {
    var columnIndex = headers.indexOf(field);
    if (columnIndex !== -1) {
      updateRequests.push({
        range: sheetName + '!' + getColumnLetter(columnIndex + 1) + rowIndex,
        values: [[updateData[field]]]
      });
    }
  });
  
  if (updateRequests.length > 0) {
    batchUpdateSheetsData(sheetsService, dbId, updateRequests);
  }
}

/**
 * 列番号をアルファベットに変換
 * @param {number} num - 列番号（1-based）
 * @returns {string} 列のアルファベット表記
 */
function getColumnLetter(num) {
  var letter = '';
  while (num > 0) {
    num--;
    letter = String.fromCharCode(65 + (num % 26)) + letter;
    num = Math.floor(num / 26);
  }
  return letter;
}

/**
 * コンテキストから統合レスポンスを構築（DB検索なし）
 * @param {object} context - 実行コンテキスト
 * @returns {object} getInitialData形式の統合レスポンス
 */
function buildResponseFromContext(context) {
  const startTime = new Date().getTime();
  console.log('🏗️ buildResponseFromContext: DB検索なしでレスポンス構築');
  
  // 🚨 緊急診断: buildResponseFromContext受信時のcontext確認
  console.log('🚨 CRITICAL: Context at buildResponseFromContext entry:', {
    contextType: typeof context,
    isString: typeof context === 'string',
    isObject: typeof context === 'object' && context !== null,
    contextConstructor: context && context.constructor && context.constructor.name,
    hasUserInfo: context && typeof context.userInfo !== 'undefined',
    hasSheetsService: context && typeof context.sheetsService !== 'undefined'
  });
  
  // contextが文字列の場合の復旧処理
  if (typeof context === 'string') {
    console.error('❌ CRITICAL: buildResponseFromContext received string context, attempting recovery');
    try {
      context = JSON.parse(context);
      console.log('✅ Successfully recovered context from JSON string');
      
      // sheetsServiceが失われている場合は再作成
      if (!context.sheetsService) {
        console.warn('⚠️ SheetsService lost during serialization, recreating...');
        const accessToken = getServiceAccountTokenCached();
        if (accessToken) {
          context.sheetsService = createSheetsService(accessToken);
          console.log('✅ SheetsService recreated after deserialization');
        }
      }
    } catch (parseError) {
      console.error('❌ Failed to recover context from JSON:', parseError.message);
      throw new Error('Context parameter corrupted and cannot be recovered');
    }
  }
  
  try {
    // 最新のuserInfoから必要な情報を取得
    const userInfo = context.userInfo;
    const configJson = JSON.parse(userInfo.configJson || '{}');
    const spreadsheetId = userInfo.spreadsheetId;
    
    // publishedSheetNameの型安全性確保（'true'問題の修正）
    let publishedSheetName = '';
    if (configJson.publishedSheetName) {
      if (typeof configJson.publishedSheetName === 'string') {
        publishedSheetName = configJson.publishedSheetName;
      } else {
        console.error('❌ publishedSheetNameが不正な型です:', typeof configJson.publishedSheetName, configJson.publishedSheetName);
        console.log('🔧 publishedSheetNameを空文字にリセットしました');
        publishedSheetName = '';
        // 不正な値をデータベースから修正する必要があることをログ出力
        console.warn('⚠️ データベースの publishedSheetName を修正してください。現在の値:', configJson.publishedSheetName);
      }
    }

    // 公開シートに紐づく設定を取得
    const sheetConfigKey = publishedSheetName ? 'sheet_' + publishedSheetName : '';
    const activeSheetConfig = sheetConfigKey && configJson[sheetConfigKey]
      ? configJson[sheetConfigKey]
      : {};

    const opinionHeader = activeSheetConfig.opinionHeader || '';
    const nameHeader = activeSheetConfig.nameHeader || '';
    const classHeader = activeSheetConfig.classHeader || '';
    
    // 基本的なレスポンス構造を構築
    const response = {
      userInfo: userInfo,
      isPublished: configJson.appPublished || false,
      setupStep: 3, // saveAndPublish完了時は常にStep 3
      
      // URL情報（キャッシュされた値を使用）
      appUrls: {
        webAppUrl: ScriptApp.getService().getUrl(),
        viewUrl: userInfo.viewUrl || (ScriptApp.getService().getUrl() + '?userId=' + encodeURIComponent(context.requestUserId) + '&mode=view'),
        setupUrl: ScriptApp.getService().getUrl() + '?setup=true',
        adminUrl: ScriptApp.getService().getUrl() + '?mode=admin&userId=' + context.requestUserId,
        status: 'success'
      },
      
      // スプレッドシート情報（既存データから構築）
      activeSheetName: publishedSheetName,

      // 表示・列設定
      config: {
        publishedSheetName: publishedSheetName,
        opinionHeader: opinionHeader,
        nameHeader: nameHeader,
        classHeader: classHeader,
        showNames: configJson.showNames || false,
        showCounts: configJson.showCounts !== undefined ? configJson.showCounts : false,
        displayMode: configJson.displayMode || 'anonymous',
        setupStatus: configJson.setupStatus || 'initial',
        isPublished: configJson.appPublished || false,
      },
      
      // パフォーマンス統計
      _meta: {
        executionTime: new Date().getTime() - context.startTime,
        includedApis: ['buildResponseFromContext'],
        apiVersion: 'optimized_v1',
        stats: context.stats
      }
    };
    
    // スプレッドシートの詳細情報が必要な場合は追加取得
    if (spreadsheetId && publishedSheetName) {
      try {
        // 最終的な型安全性チェック: 'true'/'false'文字列の検出と防止
        if (publishedSheetName === 'true' || publishedSheetName === 'false') {
          console.error('❌ buildResponseFromContext: 無効なシート名を検出:', publishedSheetName);
          console.warn('⚠️ シート詳細取得をスキップします');
        } else {
          console.log('DEBUG: Calling getSheetDetailsFromContext with context service');
          // シート情報を取得（最低限の情報のみ、既存SheetsServiceを使用）
          const sheetDetails = getSheetDetailsFromContext(context, spreadsheetId, publishedSheetName);
          response.sheetDetails = sheetDetails;
          response.allSheets = sheetDetails.allSheets || [];
          response.sheetNames = sheetDetails.sheetNames || [];

          // ヘッダー情報と自動マッピング結果を追加
          response.headers = sheetDetails.allHeaders || [];
          response.autoMappedHeaders = sheetDetails.guessedConfig || {
            opinionHeader: '',
            reasonHeader: '',
            nameHeader: '',
            classHeader: '',
          };
        }
        
      } catch (e) {
        console.warn('buildResponseFromContext: シート詳細取得エラー（基本情報のみで継続）:', e.message);
        response.allSheets = [];
        response.sheetNames = [];
        response.headers = [];
        response.autoMappedHeaders = {
          opinionHeader: '',
          reasonHeader: '',
          nameHeader: '',
          classHeader: '',
        };
      }
    }
    
    const endTime = new Date().getTime();
    console.log('✅ レスポンス構築完了: %dms', endTime - startTime);
    
    return response;
    
  } catch (error) {
    console.error('❌ buildResponseFromContext エラー:', error.message);
    throw new Error('レスポンス構築に失敗しました: ' + error.message);
  }
}

/**
 * 最適化版シート詳細取得（コンテキスト内SheetsService使用）
 * @param {object} context - 実行コンテキスト
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {object} シート詳細情報
 */
function getSheetDetailsFromContext(context, spreadsheetId, sheetName) {
  console.log('DEBUG: getSheetDetailsFromContext called with context object');
  
  // 🚨 パラメータ型検証: 正しい型で呼ばれているか確認
  console.log('🔍 Parameter validation:', {
    contextType: typeof context,
    spreadsheetIdType: typeof spreadsheetId,
    sheetNameType: typeof sheetName,
    contextIsObject: typeof context === 'object' && context !== null,
    spreadsheetIdIsString: typeof spreadsheetId === 'string',
    sheetNameIsString: typeof sheetName === 'string'
  });
  
  // contextが文字列の場合はエラー（パラメータ順序間違いの可能性）
  if (typeof context === 'string') {
    console.error('❌ CRITICAL: getSheetDetailsFromContext received string as context parameter');
    console.error('❌ This suggests parameter order mismatch. Expected: (context:object, spreadsheetId:string, sheetName:string)');
    console.error('❌ Received parameters:', {
      param1: typeof context === 'string' ? context.substring(0, 50) : context,
      param2: spreadsheetId,
      param3: sheetName
    });
    throw new Error('Parameter order mismatch: context must be an object, not a string. Check function call parameters.');
  }
  
  // 基本的なパラメータ検証
  if (!context || typeof context !== 'object') {
    throw new Error('getSheetDetailsFromContext: context parameter must be a valid object');
  }
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    throw new Error('getSheetDetailsFromContext: spreadsheetId parameter must be a valid string');
  }
  if (!sheetName || typeof sheetName !== 'string' || sheetName.trim() === '') {
    console.error('❌ CRITICAL: Invalid sheetName parameter:', {
      sheetName: sheetName,
      sheetNameType: typeof sheetName,
      sheetNameTrimmed: typeof sheetName === 'string' ? sheetName.trim() : 'not string',
      contextUserInfo: context && context.userInfo ? {
        hasConfigJson: !!(context.userInfo.configJson),
        publishedSheetName: JSON.parse(context.userInfo.configJson || '{}').publishedSheetName
      } : 'no userInfo'
    });
    
    // フォールバック: configJsonからpublishedSheetNameを取得
    if (context && context.userInfo && context.userInfo.configJson) {
      try {
        const configJson = JSON.parse(context.userInfo.configJson);
        const fallbackSheetName = configJson.publishedSheetName;
        if (fallbackSheetName && typeof fallbackSheetName === 'string' && fallbackSheetName.trim() !== '') {
          console.warn('⚠️ Using fallback publishedSheetName from config:', fallbackSheetName);
          sheetName = fallbackSheetName;
        } else {
          throw new Error('getSheetDetailsFromContext: sheetName parameter must be a valid non-empty string. Received: ' + sheetName);
        }
      } catch (parseError) {
        throw new Error('getSheetDetailsFromContext: sheetName parameter must be a valid non-empty string and configJson fallback failed');
      }
    } else {
      throw new Error('getSheetDetailsFromContext: sheetName parameter must be a valid non-empty string and no fallback available');
    }
  }
  
  try {
    // 入力パラメータの検証強化
    if (!spreadsheetId) {
      throw new Error('spreadsheetIdは必須です');
    }
    if (!sheetName) {
      throw new Error('sheetNameは必須です');
    }
    
    // サービスオブジェクトの健全性を確認し、必要に応じて復旧
    if (!context?.sheetsService?.baseUrl || !context?.sheetsService?.accessToken) {
      console.warn('⚠️ ExecutionContextのSheetsServiceが無効、復旧中...');
      try {
        // Step 1: 新しいサービスを取得（代入前の検証）
        console.log('🔍 Step 1: getSheetsServiceCached呼び出し前');
        const newService = getSheetsServiceCached();
        console.log('🔍 Step 2: getSheetsServiceCached返り値の検証:', {
          hasService: !!newService,
          hasBaseUrl: !!(newService && newService.baseUrl),
          hasAccessToken: !!(newService && newService.accessToken),
          baseUrl: newService && newService.baseUrl,
          serviceType: typeof newService,
          serviceKeys: newService ? Object.keys(newService) : null
        });
        
        // Step 2: 代入前のコンテキスト状態確認
        console.log('🔍 Step 3: 代入前のコンテキスト状態:', {
          hasContext: !!context,
          contextType: typeof context,
          contextKeys: context ? Object.keys(context) : null,
          currentService: context ? typeof context.sheetsService : 'no context'
        });
        
        // Step 3: 実際の代入
        console.log('🔍 Step 4: サービス代入実行');
        context.sheetsService = newService;
        console.log('🔍 Step 5: 代入直後の確認');
        
        // Step 4: 代入後の検証
        console.log('🔍 復旧後のサービスオブジェクト検証:', {
          hasService: !!context.sheetsService,
          hasBaseUrl: !!(context.sheetsService && context.sheetsService.baseUrl),
          hasAccessToken: !!(context.sheetsService && context.sheetsService.accessToken),
          baseUrl: context.sheetsService && context.sheetsService.baseUrl,
          serviceType: typeof context.sheetsService,
          serviceKeys: context.sheetsService ? Object.keys(context.sheetsService) : null
        });
        
        // 代入が失敗した場合のフォールバック: 直接サービスを作成
        if (!context.sheetsService || !context.sheetsService.baseUrl || !context.sheetsService.accessToken) {
          console.warn('⚠️ 代入が失敗、直接サービス作成でリトライ');
          
          // トークンを取得して直接サービスを作成
          const accessToken = getServiceAccountTokenCached();
          if (!accessToken) {
            throw new Error('SheetsService復旧に失敗: アクセストークンが取得できません');
          }
          
          // 直接createSheetsServiceを呼び出してコンテキストに設定
          context.sheetsService = createSheetsService(accessToken);
          
          console.log('🔍 直接作成後の検証:', {
            hasService: !!context.sheetsService,
            hasBaseUrl: !!(context.sheetsService && context.sheetsService.baseUrl),
            hasAccessToken: !!(context.sheetsService && context.sheetsService.accessToken),
            baseUrl: context.sheetsService && context.sheetsService.baseUrl,
            serviceType: typeof context.sheetsService
          });
          
          // 最終検証
          if (!context.sheetsService || !context.sheetsService.baseUrl || !context.sheetsService.accessToken) {
            throw new Error('SheetsService復旧に失敗: 直接作成でも有効なサービスオブジェクトを作成できません');
          }
          
          console.log('✅ SheetsService直接作成による復旧完了');
        } else {
          console.log('✅ SheetsService復旧完了');
        }
      } catch (serviceError) {
        console.error('❌ SheetsService復旧エラー:', serviceError.message);
        console.error('❌ Error stack:', serviceError.stack);
        
        // 最後の手段: 強制リフレッシュでリトライ
        try {
          console.warn('🔄 最後の手段: 強制リフレッシュでサービス復旧をリトライ');
          
          // キャッシュを強制クリアして新しいトークンで再試行
          cacheManager.remove('service_account_token');
          const freshToken = generateNewServiceAccountToken();
          
          if (!freshToken) {
            throw new Error('強制リフレッシュでもトークン取得に失敗');
          }
          
          context.sheetsService = createSheetsService(freshToken);
          
          console.log('🔍 強制リフレッシュ後の検証:', {
            hasService: !!context.sheetsService,
            hasBaseUrl: !!(context.sheetsService && context.sheetsService.baseUrl),
            hasAccessToken: !!(context.sheetsService && context.sheetsService.accessToken),
            serviceType: typeof context.sheetsService
          });
          
          if (!context.sheetsService || !context.sheetsService.baseUrl || !context.sheetsService.accessToken) {
            throw new Error('強制リフレッシュでも復旧失敗');
          }
          
          console.log('✅ 強制リフレッシュによるSheetsService復旧成功');
          
        } catch (retryError) {
          console.error('❌ 強制リフレッシュリトライも失敗:', retryError.message);
          throw new Error(`Sheets APIサービスの復旧に完全に失敗しました。初期エラー: ${serviceError.message}, リトライエラー: ${retryError.message}`);
        }
      }
    }
    
    // コンテキスト内のSheetsServiceを使用してシート情報を取得
    console.log('DEBUG: Calling getSpreadsheetsData with context service');
    
    // 最終的なサービス状態の検証（API呼び出し前）
    console.log('🔍 API呼び出し前の最終サービス検証:', {
      hasService: !!context.sheetsService,
      hasBaseUrl: !!(context.sheetsService && context.sheetsService.baseUrl),
      hasAccessToken: !!(context.sheetsService && context.sheetsService.accessToken),
      hasSpreadsheets: !!(context.sheetsService && context.sheetsService.spreadsheets),
      hasGet: !!(context.sheetsService && context.sheetsService.spreadsheets && typeof context.sheetsService.spreadsheets.get === 'function'),
      serviceType: typeof context.sheetsService,
      baseUrl: context.sheetsService && context.sheetsService.baseUrl
    });
    let data;
    try {
      data = getSpreadsheetsData(context.sheetsService, spreadsheetId);
      if (!data) {
        throw new Error('getSpreadsheetsData returned null');
      }
      console.log('DEBUG: getSpreadsheetsData success, sheets count:', data.sheets?.length || 0);
    } catch (apiError) {
      console.error('❌ getSpreadsheetsData failed:', apiError.message);
      throw new Error('スプレッドシート情報の取得に失敗: ' + apiError.message);
    }

    if (!data.sheets || data.sheets.length === 0) {
      throw new Error('スプレッドシートにシートが存在しません');
    }
    
    // 指定されたシートが存在するか確認
    const targetSheet = data.sheets.find(sheet => sheet.properties.title === sheetName);
    if (!targetSheet) {
      throw new Error(`指定されたシート '${sheetName}' が見つかりません。利用可能なシート: ${data.sheets.map(s => s.properties.title).join(', ')}`);
    }

    // ヘッダー行をAPIで取得
    const range = `'${sheetName}'!1:1`;
    console.log('DEBUG: Fetching headers with range:', range);
    let batch;
    try {
      batch = batchGetSheetsData(context.sheetsService, spreadsheetId, [range]);
      if (!batch || !batch.valueRanges) {
        throw new Error('batchGetSheetsData returned invalid response');
      }
      console.log('DEBUG: batchGetSheetsData success, valueRanges count:', batch.valueRanges.length);
    } catch (batchError) {
      console.error('❌ batchGetSheetsData failed:', batchError.message);
      throw new Error('ヘッダー行の取得に失敗: ' + batchError.message);
    }
    
    const headers = (batch.valueRanges && batch.valueRanges[0] && batch.valueRanges[0].values)
      ? batch.valueRanges[0].values[0] || []
      : [];
    
    console.log('DEBUG: Retrieved headers:', headers.length, 'columns');
    
    if (headers.length === 0) {
      throw new Error(`シート '${sheetName}' の1行目にヘッダーが見つかりません`);
    }

    const guessed = autoMapHeaders(headers);
    const existing = getConfigFromContext(context, sheetName);

    const result = {
      allHeaders: headers,
      guessedConfig: guessed,
      existingConfig: existing,
      allSheets: data.sheets,
      sheetNames: data.sheets.map(sheet => ({
        name: sheet.properties.title,
        id: sheet.properties.sheetId
      }))
    };
    
    console.log('✅ getSheetDetails success:', {
      headersCount: headers.length,
      sheetsCount: data.sheets.length,
      hasExistingConfig: Object.keys(existing).length > 0
    });
    
    return result;

  } catch (error) {
    console.error('❌ getSheetDetails エラー:', error.message);
    console.error('❌ Error stack:', error.stack);
    // エラーの詳細を含めてスローし直す（空のオブジェクトを返さない）
    throw new Error('シート詳細情報の取得に失敗しました: ' + error.message);
  }
}

/**
 * コンテキストから設定を取得（DB検索なし）
 * @param {object} context - 実行コンテキスト  
 * @param {string} sheetName - シート名
 * @returns {object} 設定オブジェクト
 */
function getConfigFromContext(context, sheetName) {
  try {
    const configJson = JSON.parse(context.userInfo.configJson || '{}');
    const sheetKey = 'sheet_' + sheetName;
    const sheetConfig = configJson[sheetKey] || {};

    // グローバル設定と重複するキーを削除して返す
    delete sheetConfig.showNames;
    delete sheetConfig.showCounts;
    delete sheetConfig.displayMode;

    return sheetConfig;
  } catch (e) {
    console.warn('getConfigFromContext エラー:', e.message);
    return {};
  }
}

/**
 * コンテキスト版: シート設定保存（インメモリ更新のみ）
 * @param {object} context - 実行コンテキスト
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {object} config - 設定オブジェクト
 */
function saveSheetConfigInContext(context, spreadsheetId, sheetName, config) {
  console.log('💾 saveSheetConfigInContext: インメモリ更新');
  
  try {
    const configJson = JSON.parse(context.userInfo.configJson || '{}');
    const sheetKey = 'sheet_' + sheetName;
    
    // シート固有の設定を準備し、グローバル設定と重複するキーを削除
    const sheetConfig = { ...config };
    delete sheetConfig.showNames;
    delete sheetConfig.showCounts;
    delete sheetConfig.displayMode;

    // シート設定を更新
    configJson[sheetKey] = {
      ...sheetConfig,
      lastModified: new Date().toISOString()
    };
    
    // updateUserOptimizedを使用してコンテキストに変更を蓄積
    updateUserOptimized(context, { 
      configJson: JSON.stringify(configJson) 
    });
    
    console.log('✅ シート設定をコンテキストに保存: %s', sheetKey);
    
  } catch (error) {
    console.error('❌ saveSheetConfigInContext エラー:', error.message);
    throw new Error('シート設定の保存に失敗しました: ' + error.message);
  }
}

/**
 * コンテキスト版: シート切り替え（インメモリ更新のみ）
 * @param {object} context - 実行コンテキスト
 * @param {string} spreadsheetId - スプレッドシートID  
 * @param {string} sheetName - シート名
 */
function switchToSheetInContext(context, spreadsheetId, sheetName) {
  console.log('🔄 switchToSheetInContext: インメモリ更新');
  
  try {
    const configJson = JSON.parse(context.userInfo.configJson || '{}');
    
    // アクティブシート情報を更新
    configJson.publishedSpreadsheetId = spreadsheetId;
    configJson.publishedSheetName = sheetName;
    configJson.appPublished = true;
    configJson.lastModified = new Date().toISOString();
    
    // 6時間自動停止機能の設定
    configJson.publishedAt = new Date().toISOString(); // 公開開始時間
    configJson.autoStopEnabled = true; // 6時間自動停止フラグ
    configJson.autoStopMinutes = 360; // 6時間 = 360分
    
    // updateUserOptimizedを使用してコンテキストに変更を蓄積
    updateUserOptimized(context, { 
      configJson: JSON.stringify(configJson) 
    });
    
    console.log('✅ シート切り替えをコンテキストに保存: %s', sheetName);
    
  } catch (error) {
    console.error('❌ switchToSheetInContext エラー:', error.message);
    throw new Error('シート切り替えに失敗しました: ' + error.message);
  }
}

/**
 * コンテキスト版: 表示オプション設定（インメモリ更新のみ）
 * @param {object} context - 実行コンテキスト
 * @param {object} displayOptions - 表示オプション
 */
function setDisplayOptionsInContext(context, displayOptions) {
  console.log('🎛️ setDisplayOptionsInContext: インメモリ更新');
  
  try {
    const configJson = JSON.parse(context.userInfo.configJson || '{}');
    
    // 表示オプションを更新
    if (displayOptions.showNames !== undefined) {
      configJson.showNames = displayOptions.showNames;
    }
    if (displayOptions.showCounts !== undefined) {
      configJson.showCounts = displayOptions.showCounts;
    }
    if (displayOptions.displayMode !== undefined) {
      configJson.displayMode = displayOptions.displayMode;
    } else if (displayOptions.showNames !== undefined) {
      // 後方互換性
      configJson.displayMode = displayOptions.showNames ? 'named' : 'anonymous';
    }
    
    configJson.lastModified = new Date().toISOString();
    
    // updateUserOptimizedを使用してコンテキストに変更を蓄積
    updateUserOptimized(context, { 
      configJson: JSON.stringify(configJson) 
    });
    
    console.log('✅ 表示オプションをコンテキストに保存:', displayOptions);
    
  } catch (error) {
    console.error('❌ setDisplayOptionsInContext エラー:', error.message);
    throw new Error('表示オプションの設定に失敗しました: ' + error.message);
  }
}

/**
 * 【最適化版】設定を保存し、ボードを公開する統合関数（トランザクション型）
 * @param {string} requestUserId - ユーザーID
 * @param {string} sheetName - シート名
 * @param {object} config - 設定オブジェクト
 * @returns {object} 最新のステータスオブジェクト
 */
function saveAndPublish(requestUserId, sheetName, config) {
  verifyUserAccess(requestUserId);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    console.log('🚀 saveAndPublishOptimized開始: sheetName=%s', sheetName);
    const overallStartTime = new Date().getTime();

    // Phase 1: 実行コンテキスト作成（リソース一括作成）
    const context = createExecutionContext(requestUserId);
    const spreadsheetId = context.userInfo.spreadsheetId;

    if (!spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません。');
    }

    // Phase 2: インメモリ更新（DB書き込みなし）
    console.log('💾 Phase 2: インメモリ更新開始');
    
    // 2-1. シート設定保存
    saveSheetConfigInContext(context, spreadsheetId, sheetName, config);
    
    // 2-2. シート切り替え
    switchToSheetInContext(context, spreadsheetId, sheetName);
    
    // 2-3. 表示オプション設定
    const displayOptions = {
      showNames: !!config.showNames,
      showCounts: config.showCounts !== undefined ? !!config.showCounts : false
    };
    setDisplayOptionsInContext(context, displayOptions);
    
    console.log('✅ Phase 2完了: 全設定をコンテキストに蓄積');

    // Phase 2.5: キャッシュクリア（DB書き込み前に実行）
    console.log('🗑️ DB書き込み前にキャッシュをクリア（古いデータ配信防止）');
    try {
      // フロントエンドが古いキャッシュを参照しないよう事前にクリア
      synchronizeCacheAfterCriticalUpdate(
        context.requestUserId, 
        context.userInfo.adminEmail,
        context.userInfo.spreadsheetId, 
        context.userInfo.spreadsheetId
      );
      console.log('✅ 事前キャッシュクリア完了');
    } catch (cacheError) {
      console.warn('⚠️ 事前キャッシュクリアでエラー（処理は続行）:', cacheError.message);
    }
    
    // Phase 3: 一括DB書き込み（1回のみ）
    console.log('💽 Phase 3: 一括DB書き込み開始');
    commitAllChanges(context);
    console.log('✅ Phase 3完了: DB書き込み完了');

    // Phase 3.5: DB書き込み後の追加キャッシュウォーミング
    console.log('🔥 新しいデータでキャッシュウォーミング中...');
    try {
      // 新しいデータを事前にキャッシュに読み込み
      const freshUserInfo = findUserByIdFresh(context.requestUserId);
      if (freshUserInfo) {
        console.log('✅ 新しいデータでキャッシュをウォーミング完了');
      }
    } catch (warmingError) {
      console.warn('⚠️ キャッシュウォーミングでエラー（処理は続行）:', warmingError.message);
    }
    
    // 最新ユーザー情報を再取得してコンテキストを更新
    const updatedUserInfo = findUserByIdFresh(context.requestUserId);
    if (updatedUserInfo) {
      context.userInfo = updatedUserInfo;
      console.log('✅ コンテキストをDBの最新情報で更新しました。');
    } else {
      console.warn('DBからのユーザー再取得に失敗しましたが、処理を続行します。');
    }

    // Phase 4: 統合レスポンス生成（DB検索なし）
    console.log('🏗️ Phase 4: レスポンス構築開始');
    
    // 🚨 緊急診断: buildResponseFromContext呼び出し前のcontext検証
    console.log('🚨 CRITICAL: Context before buildResponseFromContext:', {
      contextType: typeof context,
      isString: typeof context === 'string',
      isObject: typeof context === 'object' && context !== null,
      hasUserInfo: context && typeof context.userInfo !== 'undefined',
      hasSheetsService: context && typeof context.sheetsService !== 'undefined',
      contextKeys: context && typeof context === 'object' ? Object.keys(context) : 'not object'
    });
    
    const finalResponse = buildResponseFromContext(context);
    console.log('✅ Phase 4完了: レスポンス構築完了');

    // パフォーマンス統計
    const totalTime = new Date().getTime() - overallStartTime;
    console.log('📊 saveAndPublishOptimized パフォーマンス統計:');
    console.log('  ⏱️ 総実行時間: %dms', totalTime);
    console.log('  🔧 SheetsService作成: %d回', context.stats.sheetsServiceCreations);
    console.log('  🗄️ DB検索: %d回', context.stats.dbQueries);
    console.log('  ⚡ 操作回数: %d回', context.stats.operationsCount);

    // レスポンスに統計情報を追加
    finalResponse._meta.totalExecutionTime = totalTime;
    finalResponse._meta.optimizationStats = {
      sheetsServiceCreations: context.stats.sheetsServiceCreations,
      dbQueries: context.stats.dbQueries,
      operationsCount: context.stats.operationsCount,
      improvementMessage: 'トランザクション型最適化により高速化'
    };

    return finalResponse;

  } catch (error) {
    console.error('❌ saveAndPublishOptimized致命的エラー:', error.message, error.stack);
    throw new Error('設定の保存と公開中にサーバーエラーが発生しました: ' + error.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * configJsonを初期値にリセット
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} リセット結果
 */
function resetConfigJson(requestUserId) {
  verifyUserAccess(requestUserId);
  
  console.log('🔄 ConfigJsonリセット開始 for user:', requestUserId);
  
  try {
    // 初期configJsonを定義
    const initialConfigJson = {
      // フォーム設定
      title: '',
      questionText: '',
      responseType: 'text',
      options: [],
      
      // 表示設定  
      displayMode: 'anonymous',
      showCounts: false,
      sortOrder: 'newest',
      
      // システム設定
      formCreated: false,
      formUrl: '',
      editFormUrl: '',
      spreadsheetId: '',
      targetSheetName: '',
      
      // メタデータ
      version: '1.0.0',
      lastModified: new Date().toISOString(),
      resetAt: new Date().toISOString()
    };
    
    // データベースでconfigJsonを更新
    const updateResult = updateUser(requestUserId, {
      configJson: JSON.stringify(initialConfigJson)
    });
    
    if (updateResult.status === 'success') {
      console.log('✅ ConfigJsonリセット完了');
      return {
        success: true,
        message: '設定を初期値にリセットしました',
        resetAt: initialConfigJson.resetAt,
        configJson: initialConfigJson
      };
    } else {
      throw new Error('データベース更新に失敗しました: ' + updateResult.message);
    }
    
  } catch (error) {
    console.error('❌ ConfigJsonリセットエラー:', error.message);
    throw new Error('設定リセットに失敗しました: ' + error.message);
  }
}

// =================================================================
// アプリケーション有効/無効機能（isActive概念）
// =================================================================

/**
 * アプリケーション全体の有効/無効状態を取得
 * @returns {boolean} アプリケーションが有効な場合はtrue
 */
function getApplicationEnabled() {
  try {
    const props = PropertiesService.getScriptProperties();
    const value = props.getProperty('APPLICATION_ENABLED');
    
    // デフォルト値: true（有効）
    if (value === null || value === undefined) {
      return true;
    }
    
    // boolean型に変換
    return value === 'true' || value === true;
  } catch (error) {
    console.error('getApplicationEnabled エラー:', error.message);
    // エラー時はデフォルトで有効とする
    return true;
  }
}

/**
 * アプリケーション全体の有効/無効状態を設定
 * @param {boolean} enabled - 有効にする場合はtrue、無効にする場合はfalse
 * @returns {object} 設定結果
 */
function setApplicationEnabled(enabled) {
  try {
    // システム管理者権限チェック
    if (!isSystemAdmin()) {
      throw new Error('この機能にアクセスする権限がありません。システム管理者のみが利用できます。');
    }
    
    const props = PropertiesService.getScriptProperties();
    const enabledValue = enabled ? 'true' : 'false';
    
    props.setProperty('APPLICATION_ENABLED', enabledValue);
    
    const message = enabled 
      ? 'アプリケーションが有効化されました。すべてのユーザーがアクセス可能です。'
      : 'アプリケーションが無効化されました。システム管理者以外のアクセスが制限されます。';
    
    console.log('setApplicationEnabled:', message);
    
    return {
      success: true,
      enabled: enabled,
      message: message,
      timestamp: new Date().toISOString(),
      adminEmail: Session.getActiveUser().getEmail()
    };
    
  } catch (error) {
    console.error('setApplicationEnabled エラー:', error.message);
    throw new Error('アプリケーション状態の設定に失敗しました: ' + error.message);
  }
}

/**
 * アプリケーションの有効/無効状態を確認し、アクセス可否を判定
 * @returns {object} アクセス可否の詳細情報
 */
function checkApplicationAccess() {
  try {
    const isEnabled = getApplicationEnabled();
    const currentUserEmail = Session.getActiveUser().getEmail();
    const isAdmin = isSystemAdmin();
    
    // アプリケーションが有効、またはシステム管理者の場合はアクセス許可
    const hasAccess = isEnabled || isAdmin;
    
    return {
      hasAccess: hasAccess,
      isApplicationEnabled: isEnabled,
      isSystemAdmin: isAdmin,
      userEmail: currentUserEmail,
      accessReason: hasAccess 
        ? (isEnabled ? 'アプリケーションが有効です' : 'システム管理者権限でアクセス')
        : 'アプリケーションが無効化されています'
    };
    
  } catch (error) {
    console.error('checkApplicationAccess エラー:', error.message);
    return {
      hasAccess: false,
      isApplicationEnabled: false,
      isSystemAdmin: false,
      userEmail: '',
      accessReason: 'アクセス確認中にエラーが発生しました',
      error: error.message
    };
  }
}






