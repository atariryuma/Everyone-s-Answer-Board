/**
 * @fileoverview 超最適化コアシステム - 2024年最新技術の結集
 * V8ランタイム、最新パフォーマンス技術、安定性強化を統合
 */

// V8ランタイムのES6+機能を活用した定数定義
const ULTRA_CONFIG = {
  EXECUTION_LIMITS: {
    MAX_TIME: 330000, // 5.5分（安全マージン）
    BATCH_SIZE: 100,
    API_RATE_LIMIT: 90 // 100秒間隔での制限
  },
  
  CACHE_STRATEGY: {
    L1_TTL: 300,     // Level 1: 5分
    L2_TTL: 3600,    // Level 2: 1時間  
    L3_TTL: 21600    // Level 3: 6時間（最大）
  }
};

/**
 * 超高速データベース操作クラス
 */
class UltraOptimizedDatabase {
  
  constructor() {
    this.circuitBreaker = StabilityEnhancer.createCircuitBreaker(
      this._makeApiCall.bind(this),
      { failureThreshold: 3, resetTimeoutMs: 30000 }
    );
    
    this.healthMonitor = StabilityEnhancer.createHealthMonitor();
    this.resourceMonitor = StabilityEnhancer.createResourceMonitor();
  }
  
  /**
   * 超高速ユーザー検索（多層キャッシュ + メモ化）
   */
  async findUserOptimized(identifier, field = 'userId') {
    const profiler = globalProfiler;
    profiler.start('findUser');
    
    try {
      // Level 1: メモ化キャッシュ
      const memoKey = `${field}_${identifier}`;
      const memoized = AdvancedCacheManager.smartGet(
        memoKey,
        null, // データ取得関数は後で指定
        { 
          enableMemoization: true,
          ttl: ULTRA_CONFIG.CACHE_STRATEGY.L1_TTL,
          scope: 'script'
        }
      );
      
      if (memoized) {
        profiler.end('findUser');
        this.healthMonitor.recordSuccess(profiler.end('findUser'));
        return memoized;
      }
      
      // Level 2: 高速データベース検索
      const user = await this._fetchUserWithResilience(field, identifier);
      
      if (user) {
        // 複数キーでキャッシュ（前方・後方参照）
        const cachePromises = [
          AdvancedCacheManager.smartGet(`userId_${user.userId}`, () => user, {
            ttl: ULTRA_CONFIG.CACHE_STRATEGY.L2_TTL
          }),
          AdvancedCacheManager.smartGet(`email_${user.adminEmail}`, () => user, {
            ttl: ULTRA_CONFIG.CACHE_STRATEGY.L2_TTL
          })
        ];
        
        // 並列キャッシュ更新（ノンブロッキング）
        Promise.all(cachePromises).catch(err => 
          console.warn('キャッシュ更新警告:', err.message)
        );
      }
      
      const duration = profiler.end('findUser');
      this.healthMonitor.recordSuccess(duration);
      
      return user;
      
    } catch (error) {
      profiler.end('findUser');
      this.healthMonitor.recordFailure(error);
      
      // フェイルセーフ - キャッシュから古いデータを取得
      return this._getStaleDataFallback(identifier, field);
    }
  }
  
  /**
   * 回復力のあるデータ取得
   */
  async _fetchUserWithResilience(field, value) {
    return await StabilityEnhancer.resilientExecute(
      async () => {
        // リソース制限チェック
        if (!this.resourceMonitor.shouldContinue()) {
          throw new Error('Resource limits exceeded');
        }
        
        // サーキットブレーカー経由でAPI呼び出し
        return await this.circuitBreaker.execute(
          this._performDatabaseQuery.bind(this),
          field,
          value
        );
      },
      {
        maxRetries: 3,
        baseDelayMs: 1000,
        retryCondition: (error) => !error.message.includes('Resource limits')
      }
    );
  }
  
  /**
   * 最適化されたデータベースクエリ
   */
  async _performDatabaseQuery(field, value) {
    const props = PropertiesService.getScriptProperties();
    const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    
    if (!dbId) {
      throw new Error('Database not configured');
    }
    
    // 最適化されたAPI経由でデータ取得
    const api = new OptimizedSheetsAPI(getAuthTokenCached());
    const response = await api.batchGetOptimized(dbId, [`${DB_SHEET_CONFIG.SHEET_NAME}!A:H`]);
    
    const values = response.valueRanges[0]?.values || [];
    if (values.length === 0) return null;
    
    const headers = values[0];
    const fieldIndex = headers.indexOf(field);
    
    if (fieldIndex === -1) return null;
    
    // 効率的な検索（V8最適化を活用）
    const dataRow = values.slice(1).find(row => row[fieldIndex] === value);
    
    if (!dataRow) return null;
    
    // オブジェクト構築（最適化）
    return headers.reduce((user, header, index) => {
      user[header] = dataRow[index] || '';
      return user;
    }, {});
  }
  
  /**
   * フェイルセーフ古いデータ取得
   */
  _getStaleDataFallback(identifier, field) {
    try {
      // PropertiesServiceから古いキャッシュを検索
      const props = PropertiesService.getScriptProperties();
      const staleKey = `STALE_${field}_${identifier}`;
      const staleData = props.getProperty(staleKey);
      
      if (staleData) {
        console.warn(`Using stale data for ${field}:${identifier}`);
        return JSON.parse(staleData);
      }
    } catch (error) {
      console.error('Stale data fallback failed:', error.message);
    }
    
    return null;
  }
  
  /**
   * 超高速バッチユーザー更新
   */
  async batchUpdateUsers(updates) {
    const profiler = globalProfiler;
    profiler.start('batchUpdate');
    
    try {
      // バッチサイズでチャンク分割
      const chunks = this._chunkArray(updates, ULTRA_CONFIG.EXECUTION_LIMITS.BATCH_SIZE);
      const results = [];
      
      for (const chunk of chunks) {
        // リソース制限チェック
        if (!this.resourceMonitor.shouldContinue()) {
          console.warn('Stopping batch update due to resource limits');
          break;
        }
        
        const chunkResult = await this._processBatchChunk(chunk);
        results.push(...chunkResult);
        
        // API制限対策の小休止
        if (chunks.length > 1) {
          await this._smartDelay();
        }
      }
      
      profiler.end('batchUpdate');
      return results;
      
    } catch (error) {
      profiler.end('batchUpdate');
      this.healthMonitor.recordFailure(error);
      throw error;
    }
  }
  
  /**
   * インテリジェント遅延（適応的）
   */
  async _smartDelay() {
    const health = this.healthMonitor.getHealth();
    
    // 成功率に基づいて遅延調整
    let delay = 50; // ベース50ms
    
    if (health.successRate < 0.9) {
      delay = 200; // 成功率低下時は200ms
    } else if (health.successRate < 0.95) {
      delay = 100; // 軽微な低下時は100ms
    }
    
    // ジッター追加（雷鳴効果回避）
    delay += Math.random() * 50;
    
    return new Promise(resolve => setTimeout(resolve, delay));
  }
  
  /**
   * 配列チャンク分割ユーティリティ
   */
  _chunkArray(array, chunkSize) {
    const chunks = [];
    for (let i = 0; i < array.length; i += chunkSize) {
      chunks.push(array.slice(i, i + chunkSize));
    }
    return chunks;
  }
  
  /**
   * バッチチャンク処理
   */
  async _processBatchChunk(chunk) {
    // 時分割並列処理でチャンクを高速処理
    return ParallelProcessor.timeSlicedParallel(
      chunk.map(update => () => this._processSingleUpdate(update)),
      { maxConcurrent: 3, sliceTimeMs: 100 }
    );
  }
  
  /**
   * 単一更新処理
   */
  async _processSingleUpdate(update) {
    // 個別更新処理の実装
    return await this._updateUserRecord(update.userId, update.data);
  }
  
  /**
   * API呼び出し（サーキットブレーカー用）
   */
  async _makeApiCall(...args) {
    // 実際のAPI呼び出し処理
    return await this._performDatabaseQuery(...args);
  }
}

/**
 * 超高速リアクションシステム
 */
class UltraOptimizedReactions {
  
  /**
   * 並列リアクション処理
   */
  static async processReactionBatch(reactions) {
    const profiler = globalProfiler;
    profiler.start('reactionBatch');
    
    try {
      // リアクションをスプレッドシート別にグループ化
      const groupedReactions = this._groupBySpreadsheet(reactions);
      
      // 並列処理でスプレッドシートごとに処理
      const promises = Object.keys(groupedReactions).map(spreadsheetId =>
        this._processSpreadsheetReactions(spreadsheetId, groupedReactions[spreadsheetId])
      );
      
      const results = await Promise.all(promises);
      
      profiler.end('reactionBatch');
      return results.flat();
      
    } catch (error) {
      profiler.end('reactionBatch');
      throw error;
    }
  }
  
  /**
   * スプレッドシート別リアクション処理
   */
  static async _processSpreadsheetReactions(spreadsheetId, reactions) {
    // バッチ処理でヘッダーと該当行を一括取得
    const ranges = [
      ...reactions.map(r => `${r.sheetName}!${r.rowIndex}:${r.rowIndex}`),
      `${reactions[0].sheetName}!1:1` // ヘッダー
    ];
    
    const api = new OptimizedSheetsAPI(getAuthTokenCached());
    const response = await api.batchGetOptimized(spreadsheetId, ranges);
    
    // 更新リクエストを構築
    const updateRequests = this._buildReactionUpdates(reactions, response);
    
    // バッチ更新実行
    return await api.batchUpdateOptimized(spreadsheetId, updateRequests);
  }
  
  /**
   * リアクション更新リクエスト構築
   */
  static _buildReactionUpdates(reactions, responseData) {
    const headers = responseData.valueRanges[responseData.valueRanges.length - 1].values[0];
    const updates = [];
    
    reactions.forEach((reaction, index) => {
      const rowData = responseData.valueRanges[index].values[0] || [];
      const columnIndex = headers.indexOf(COLUMN_HEADERS[reaction.reactionKey]);
      
      if (columnIndex !== -1) {
        const currentReactions = this._parseReactionString(rowData[columnIndex] || '');
        const userIndex = currentReactions.indexOf(reaction.userEmail);
        
        if (userIndex >= 0) {
          currentReactions.splice(userIndex, 1);
        } else {
          currentReactions.push(reaction.userEmail);
        }
        
        updates.push({
          range: `${reaction.sheetName}!${String.fromCharCode(65 + columnIndex)}${reaction.rowIndex}`,
          values: [[currentReactions.join(', ')]]
        });
      }
    });
    
    return updates;
  }
  
  /**
   * スプレッドシート別グループ化
   */
  static _groupBySpreadsheet(reactions) {
    return reactions.reduce((groups, reaction) => {
      const key = reaction.spreadsheetId;
      if (!groups[key]) {
        groups[key] = [];
      }
      groups[key].push(reaction);
      return groups;
    }, {});
  }
  
  /**
   * リアクション文字列パース
   */
  static _parseReactionString(val) {
    if (!val) return [];
    return val.toString().split(',').map(s => s.trim()).filter(Boolean);
  }
}

/**
 * 超最適化メイン制御システム
 */
class UltraOptimizedController {
  
  constructor() {
    this.database = new UltraOptimizedDatabase();
    this.isInitialized = false;
  }
  
  /**
   * 初期化（遅延ロード）
   */
  async initialize() {
    if (this.isInitialized) return;
    
    try {
      // システム健康状態チェック
      const recoveryCheck = AutoRecoveryService.performRecoveryCheck();
      
      if (recoveryCheck.overallStatus === 'EMERGENCY') {
        throw new Error('System in emergency state');
      }
      
      // 自動修復実行（必要時）
      if (recoveryCheck.overallStatus !== 'HEALTHY') {
        AutoRecoveryService.performAutoRepair();
      }
      
      this.isInitialized = true;
      console.log('Ultra-optimized system initialized successfully');
      
    } catch (error) {
      console.error('Initialization failed:', error.message);
      throw error;
    }
  }
  
  /**
   * 超高速doGet処理
   */
  async doGetOptimized(e) {
    const profiler = globalProfiler;
    profiler.start('doGet');
    
    try {
      await this.initialize();
      
      const { userId, mode, setup } = e.parameter;
      
      // セットアップページ（キャッシュ対象外）
      if (setup === 'true') {
        return HtmlService.createTemplateFromFile('SetupPage')
          .evaluate()
          .setTitle('StudyQuest - Setup');
      }
      
      // 新規登録ページ
      if (!userId) {
        return HtmlService.createTemplateFromFile('Registration')
          .evaluate()
          .setTitle('新規登録');
      }
      
      // ユーザー情報取得（超高速キャッシュ利用）
      const userInfo = await this.database.findUserOptimized(userId);
      
      if (!userInfo) {
        return HtmlService.createHtmlOutput('無効なユーザーIDです。');
      }
      
      // 非同期で最終アクセス日時更新（レスポンス遅延なし）
      this._updateLastAccessAsync(userId).catch(err => 
        console.warn('Last access update failed:', err.message)
      );
      
      // ユーザーコンテキスト設定
      PropertiesService.getUserProperties().setProperty('CURRENT_USER_ID', userId);
      
      // テンプレート生成とレスポンス
      const template = mode === 'admin' ? 
        this._createAdminTemplate(userInfo, userId) :
        this._createUserTemplate(userInfo, userId);
      
      profiler.end('doGet');
      return template;
      
    } catch (error) {
      profiler.end('doGet');
      
      // 緊急シャットダウン条件チェック
      if (error.message.includes('Resource limits') || 
          error.message.includes('emergency')) {
        EmergencyShutdown.initiate(error.message);
      }
      
      throw error;
    }
  }
  
  /**
   * 非同期最終アクセス更新
   */
  async _updateLastAccessAsync(userId) {
    // 低優先度での更新（エラーは無視）
    try {
      await this.database.batchUpdateUsers([{
        userId: userId,
        data: { lastAccessedAt: new Date().toISOString() }
      }]);
    } catch (error) {
      // 無視（ログ出力のみ）
      console.warn('Last access update ignored:', error.message);
    }
  }
  
  /**
   * 管理テンプレート作成
   */
  _createAdminTemplate(userInfo, userId) {
    const template = HtmlService.createTemplateFromFile('AdminPanel');
    template.userInfo = userInfo;
    template.userId = userId;
    return template.evaluate().setTitle('管理パネル - みんなの回答ボード');
  }
  
  /**
   * ユーザーテンプレート作成
   */
  _createUserTemplate(userInfo, userId) {
    const template = HtmlService.createTemplateFromFile('Page');
    template.userInfo = userInfo;
    template.userId = userId;
    return template.evaluate().setTitle('みんなの回答ボード');
  }
}

// グローバルインスタンス
const ultraController = new UltraOptimizedController();

/**
 * 公開エントリーポイント（超最適化版）
 */
function doGet(e) {
  return ultraController.doGetOptimized(e);
}

/**
 * パフォーマンス監視エンドポイント
 */
function getPerformanceMetrics() {
  return {
    profiler: globalProfiler.getReport(),
    health: ultraController.database?.healthMonitor?.getHealth() || {},
    cache: AdvancedCacheManager.getHealth(),
    system: AutoRecoveryService.performRecoveryCheck()
  };
}