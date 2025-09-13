/**
 * @fileoverview DatabaseCore - データベースコア機能
 *
 * 🎯 責任範囲:
 * - データベース接続・認証
 * - 基本CRUD操作
 * - サービスアカウント管理
 */

/* global PROPS_KEYS, CONSTANTS, SecurityService, AppCacheService, ConfigService, UserService, UnifiedLogger */

/**
 * DatabaseCore - データベースコア機能
 * 基本的なデータベース操作とサービス管理
 */
const DatabaseCore = Object.freeze({

  // ==========================================
  // 🔐 データベース接続・認証
  // ==========================================

  /**
   * セキュアなデータベースID取得
   * @returns {string} データベースID
   */
  getSecureDatabaseId() {
    try {
      return PropertiesService.getScriptProperties().getProperty(PROPS_KEYS.DATABASE_SPREADSHEET_ID);
    } catch (error) {
      UnifiedLogger.error('DatabaseCore', {
        operation: 'getSecureDatabaseId',
        error: error.message
      });
      throw new Error('データベース設定の取得に失敗しました');
    }
  },

  /**
   * バッチデータ取得
   * @param {Object} service - Sheetsサービス
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {Array} ranges - 取得範囲配列
   * @returns {Object} バッチ取得結果
   */
  batchGetSheetsData(service, spreadsheetId, ranges) {
    try {
      const timer = UnifiedLogger.startTimer('DatabaseCore.batchGetSheetsData');

      if (!ranges || ranges.length === 0) {
        return { valueRanges: [] };
      }

      const result = service.spreadsheets.values.batchGet({
        spreadsheetId,
        ranges
      });

      timer.end();
      return result;
    } catch (error) {
      UnifiedLogger.error('DatabaseCore', {
        operation: 'batchGetSheetsData',
        spreadsheetId,
        rangesCount: ranges?.length,
        error: error.message
      });
      throw error;
    }
  },

  /**
   * Sheetsサービス取得（キャッシュ付き）
   * @returns {Object} Sheetsサービス
   */
  getSheetsServiceCached() {
    const cacheKey = 'sheets_service';

    try {
      const cachedService = AppCacheService.get(cacheKey, null);
      if (cachedService) {
        const isValidService =
          cachedService.spreadsheets &&
          cachedService.spreadsheets.values &&
          typeof cachedService.spreadsheets.values.get === 'function';

        if (isValidService) {
          return cachedService;
        }
      }

      // 新しいサービス作成
      const service = this.createSheetsService();
      AppCacheService.set(cacheKey, service, AppCacheService.TTL.MEDIUM);

      return service;
    } catch (error) {
      UnifiedLogger.error('DatabaseCore', {
        operation: 'getSheetsServiceCached',
        error: error.message
      });
      throw error;
    }
  },

  /**
   * Sheetsサービス作成
   * @returns {Object} 新しいSheetsサービス
   */
  createSheetsService() {
    try {
      const serviceAccountKey = PropertiesService.getScriptProperties()
        .getProperty(PROPS_KEYS.SERVICE_ACCOUNT_KEY);

      if (!serviceAccountKey) {
        throw new Error('サービスアカウントキーが設定されていません');
      }

      const parsedKey = JSON.parse(serviceAccountKey);
      const service = Sheets.newService();
      service.useServiceAccountCreds(parsedKey);

      UnifiedLogger.success('DatabaseCore', {
        operation: 'createSheetsService',
        serviceType: parsedKey.type || 'unknown'
      });

      return service;
    } catch (error) {
      UnifiedLogger.error('DatabaseCore', {
        operation: 'createSheetsService',
        error: error.message
      });
      throw error;
    }
  },

  /**
   * リトライ付きSheetsサービス取得
   * @param {number} maxRetries - 最大リトライ回数
   * @returns {Object} Sheetsサービス
   */
  getSheetsServiceWithRetry(maxRetries = 2) {
    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        return this.getSheetsServiceCached();
      } catch (error) {
        UnifiedLogger.warn('DatabaseCore', {
          operation: 'getSheetsServiceWithRetry',
          attempt,
          maxRetries,
          error: error.message
        });

        if (attempt === maxRetries) {
          throw error;
        }

        Utilities.sleep(1000 * attempt); // 指数バックオフ
      }
    }
  },

  // ==========================================
  // 🔧 診断・ユーティリティ
  // ==========================================

  /**
   * データベース接続診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const results = {
      service: 'DatabaseCore',
      timestamp: new Date().toISOString(),
      checks: []
    };

    try {
      // データベースID確認
      const databaseId = this.getSecureDatabaseId();
      results.checks.push({
        name: 'Database ID',
        status: databaseId ? '✅' : '❌',
        details: databaseId ? 'Database ID configured' : 'Database ID missing'
      });

      // サービスアカウント確認
      try {
        const service = this.createSheetsService();
        results.checks.push({
          name: 'Service Account',
          status: service ? '✅' : '❌',
          details: 'Service account authentication working'
        });
      } catch (serviceError) {
        results.checks.push({
          name: 'Service Account',
          status: '❌',
          details: serviceError.message
        });
      }

      // キャッシュサービス確認
      try {
        AppCacheService.get('test_key', null);
        results.checks.push({
          name: 'Cache Service',
          status: '✅',
          details: 'Cache service accessible'
        });
      } catch (cacheError) {
        results.checks.push({
          name: 'Cache Service',
          status: '⚠️',
          details: cacheError.message
        });
      }

      results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    } catch (error) {
      results.checks.push({
        name: 'Core Diagnosis',
        status: '❌',
        details: error.message
      });
      results.overall = '❌';
    }

    return results;
  }

});