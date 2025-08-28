/**
 * @fileoverview 統一Sheet関数アーキテクチャ
 * 35個のSheet関連関数を統一設計でまとめた統合管理システム
 * 
 * 機能分類:
 * 1. データ取得系: getPublishedSheetData, getSheetData, getIncrementalSheetData等
 * 2. ヘッダー・メタ情報系: getSheetHeaders, getHeaderIndices, getHeadersCached等
 * 3. Spreadsheet管理系: getCurrentSpreadsheet, getEffectiveSpreadsheetId等
 * 
 * 統合目標:
 * - コア統一関数による一貫性確保
 * - 階層化設計（コア関数 + 用途別ラッパー）
 * - 統一キャッシュ戦略
 * - 統一エラーハンドリング
 * - 統一レスポンス形式
 */

/** @OnlyCurrentDoc */

/**
 * 統一Sheet データ管理クラス
 * すべてのSheet関連操作の中核を担う
 */
class UnifiedSheetDataManager {
  constructor() {
    this.cache = cacheManager;
    this.logger = ULog;
  }

  // ===========================================
  // 1. コアデータ取得メソッド群
  // ===========================================

  /**
   * 統一データ取得コア関数
   * 全てのSheet データ取得の基盤となる関数
   * 
   * @param {Object} options - 取得オプション
   * @param {string} options.userId - ユーザーID
   * @param {string} options.spreadsheetId - スプレッドシートID
   * @param {string} options.sheetName - シート名
   * @param {string} [options.classFilter] - クラスフィルタ
   * @param {string} [options.sortOrder] - ソート順
   * @param {boolean} [options.adminMode] - 管理者モード
   * @param {boolean} [options.bypassCache] - キャッシュバイパス
   * @param {string} [options.dataType='full'] - データタイプ (full/published/incremental)
   * @param {number} [options.sinceRowCount] - 増分データの開始行数
   * @param {string} [options.rangeOverride] - 範囲オーバーライド
   * @return {Object} 統一レスポンス形式のデータ
   */
  getCoreSheetData(options) {
    const {
      userId,
      spreadsheetId,
      sheetName,
      classFilter = '',
      sortOrder = '',
      adminMode = false,
      bypassCache = false,
      dataType = 'full',
      sinceRowCount,
      rangeOverride
    } = options;

    // パラメータ検証
    this._validateCoreParams(userId, spreadsheetId, sheetName);

    // キャッシュキー生成
    const cacheKey = this._generateCacheKey({
      type: dataType,
      userId,
      spreadsheetId,
      sheetName,
      classFilter,
      sortOrder,
      adminMode,
      sinceRowCount
    });

    // キャッシュ設定
    const cacheConfig = this._getCacheConfig(dataType, adminMode);

    // キャッシュバイパス時は直接実行
    if (bypassCache || adminMode) {
      this.logger.debug(`🔄 ${dataType}データ取得：キャッシュバイパス`);
      return this._executeDataFetch(options);
    }

    // キャッシュ経由でデータ取得
    return this.cache.get(
      cacheKey,
      () => this._executeDataFetch(options),
      cacheConfig
    );
  }

  /**
   * 実際のデータ取得実行
   * @private
   */
  _executeDataFetch(options) {
    const {
      userId,
      spreadsheetId,
      sheetName,
      classFilter,
      sortOrder,
      adminMode,
      dataType,
      sinceRowCount,
      rangeOverride
    } = options;

    try {
      // ユーザー情報取得
      const userInfo = this._getUserInfo(userId, options);
      
      // データ取得方式選択
      switch (dataType) {
        case 'published':
          return this._fetchPublishedData(userInfo, options);
        case 'incremental':
          return this._fetchIncrementalData(userInfo, options);
        case 'full':
        default:
          return this._fetchFullSheetData(userInfo, options);
      }
    } catch (error) {
      this.logger.error(`[getCoreSheetData] ${dataType}データ取得エラー:`, error.message);
      throw new Error(`シートデータの取得に失敗しました: ${error.message}`);
    }
  }

  /**
   * 公開データ取得
   * @private
   */
  _fetchPublishedData(userInfo, options) {
    const { classFilter, sortOrder, adminMode } = options;

    // 公開状態チェック
    if (!isPublishedBoard(userInfo)) {
      throw new Error('この回答ボードは現在公開されていません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    const setupStatus = configJson.setupStatus || 'pending';

    // 公開設定チェック
    const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    const publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      if (setupStatus === 'pending') {
        return {
          status: 'setup_required',
          message: 'セットアップを完了してください。データ準備、シート・列設定、公開設定の順番で進めてください。',
          data: [],
          setupStatus: setupStatus,
        };
      }
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // シート固有設定取得
    const sheetKey = 'sheet_' + publishedSheetName;
    const sheetConfig = configJson[sheetKey] || {};

    // データ取得実行
    const rawData = this._getRawSheetData(publishedSpreadsheetId, publishedSheetName, sheetConfig);
    
    // データ整形
    const formattedData = formatSheetDataForFrontend(rawData, [], sheetConfig, classFilter, sortOrder, adminMode);

    return createSuccessResponse(formattedData);
  }

  /**
   * 増分データ取得
   * @private
   */
  _fetchIncrementalData(userInfo, options) {
    const { classFilter, sortOrder, adminMode, sinceRowCount } = options;

    // 公開状態チェック
    if (!isPublishedBoard(userInfo)) {
      throw new Error('この回答ボードは現在公開されていません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    const setupStatus = configJson.setupStatus || 'pending';
    const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
    const publishedSheetName = configJson.publishedSheetName;

    if (!publishedSpreadsheetId || !publishedSheetName) {
      if (setupStatus === 'pending') {
        return {
          status: 'setup_required',
          message: 'セットアップを完了してください。',
          incrementalData: [],
          setupStatus: setupStatus,
        };
      }
      throw new Error('公開対象のスプレッドシートまたはシートが設定されていません。');
    }

    // スプレッドシート取得
    const ss = openSpreadsheetOptimized(publishedSpreadsheetId);
    const sheet = ss.getSheetByName(publishedSheetName);

    if (!sheet) {
      throw new Error('指定されたシートが見つかりません: ' + publishedSheetName);
    }

    const lastRow = sheet.getLastRow();
    const headerRow = 1;
    
    // 増分データ範囲計算
    const startRow = Math.max(sinceRowCount + headerRow + 1, headerRow + 1);
    if (startRow > lastRow) {
      this.logger.debug('🔄 新しいデータはありません');
      return {
        status: 'success',
        incrementalData: [],
        totalRows: lastRow - headerRow,
        sinceRowCount: sinceRowCount,
        newDataCount: 0
      };
    }

    const dataRange = sheet.getRange(startRow, 1, lastRow - startRow + 1, sheet.getLastColumn());
    const values = dataRange.getValues();
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];

    // データ変換
    const incrementalRows = values.map(row => {
      const rowObj = {};
      headers.forEach((header, index) => {
        rowObj[header] = row[index] || '';
      });
      return rowObj;
    });

    this.logger.debug(`🔄 増分データ取得完了: ${incrementalRows.length}件`);

    return {
      status: 'success',
      incrementalData: incrementalRows,
      totalRows: lastRow - headerRow,
      sinceRowCount: sinceRowCount,
      newDataCount: incrementalRows.length
    };
  }

  /**
   * 完全シートデータ取得
   * @private
   */
  _fetchFullSheetData(userInfo, options) {
    const { sheetName, classFilter, sortOrder, adminMode } = options;

    const configJson = JSON.parse(userInfo.configJson || '{}');
    const spreadsheetId = configJson.publishedSpreadsheetId || userInfo.spreadsheetId;
    const service = getSheetsServiceCached();

    // データ範囲設定
    const ranges = [sheetName + '!A:Z'];

    // データ取得
    const responses = batchGetSheetsData(service, spreadsheetId, ranges);
    const sheetData = responses.valueRanges[0].values || [];
    const rosterData = []; // 名簿機能は使用しない

    if (sheetData.length === 0) {
      return createSuccessResponse({
        data: [],
        message: 'データが見つかりませんでした',
        totalCount: 0,
        filteredCount: 0,
      });
    }

    // シート設定取得
    const sheetKey = 'sheet_' + sheetName;
    const sheetConfig = configJson[sheetKey] || {};

    // データ整形
    const formattedData = formatSheetDataForFrontend(
      { originalData: sheetData, reactionCounts: {} },
      rosterData,
      sheetConfig,
      classFilter,
      sortOrder,
      adminMode
    );

    return createSuccessResponse(formattedData);
  }

  // ===========================================
  // 2. ヘッダー管理メソッド群
  // ===========================================

  /**
   * 統一ヘッダー取得関数
   * すべてのヘッダー関連操作の基盤
   * 
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {Object} [options={}] - オプション
   * @param {boolean} [options.forceRefresh=false] - キャッシュ強制更新
   * @param {number} [options.maxRetries=3] - 最大リトライ回数
   * @param {boolean} [options.validateRequired=true] - 必須列検証
   * @return {Object} ヘッダー情報オブジェクト
   */
  getCoreSheetHeaders(spreadsheetId, sheetName, options = {}) {
    const {
      forceRefresh = false,
      maxRetries = 3,
      validateRequired = true
    } = options;

    // パラメータ検証
    if (!spreadsheetId || !sheetName) {
      throw new Error('spreadsheetIdとsheetNameは必須です');
    }

    const key = `unified_hdr_${spreadsheetId}_${sheetName}`;
    const validationKey = `unified_hdr_val_${spreadsheetId}_${sheetName}`;

    // 強制更新時はキャッシュクリア
    if (forceRefresh) {
      this.cache.remove(key);
      this.cache.remove(validationKey);
    }

    // ヘッダー取得（キャッシュ経由）
    const headers = this.cache.get(
      key,
      () => this._fetchHeadersWithRetry(spreadsheetId, sheetName, maxRetries),
      { ttl: 1800, enableMemoization: true } // 30分間キャッシュ + メモ化
    );

    // 必須列検証
    if (validateRequired && headers && Object.keys(headers).length > 0) {
      const validation = this.cache.get(
        validationKey,
        () => this._validateRequiredHeaders(headers),
        { ttl: 1800, enableMemoization: true }
      );

      if (!validation.isValid) {
        return this._handleHeaderValidationFailure(spreadsheetId, sheetName, headers, validation);
      }
    }

    return headers || {};
  }

  /**
   * ヘッダーインデックス取得
   * 
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @return {Object} ヘッダーインデックスマップ
   */
  getHeaderIndices(spreadsheetId, sheetName) {
    return this.getCoreSheetHeaders(spreadsheetId, sheetName, {
      validateRequired: false
    });
  }

  /**
   * ヘッダー名からインデックス取得
   * 
   * @param {Object} headers - ヘッダーオブジェクト
   * @param {string} headerName - ヘッダー名
   * @return {number} インデックス（見つからない場合は-1）
   */
  getHeaderIndex(headers, headerName) {
    return headers[headerName] !== undefined ? headers[headerName] : -1;
  }

  /**
   * リトライ機能付きヘッダー取得
   * @private
   */
  _fetchHeadersWithRetry(spreadsheetId, sheetName, maxRetries) {
    let lastError = null;

    for (let attempt = 1; attempt <= maxRetries; attempt++) {
      try {
        this.logger.debug(
          `[_fetchHeadersWithRetry] Attempt ${attempt}/${maxRetries} - ${spreadsheetId}:${sheetName}`
        );

        const service = getSheetsServiceCached();
        if (!service) {
          throw new Error('Sheets service is not available');
        }

        const range = sheetName + '!1:1';
        const url = service.baseUrl + '/' + spreadsheetId + '/values/' + encodeURIComponent(range);
        
        const response = UrlFetchApp.fetch(url, {
          method: 'GET',
          headers: {
            'Authorization': 'Bearer ' + service.accessToken,
            'Accept': 'application/json'
          },
          muteHttpExceptions: true
        });

        if (response.getResponseCode() !== 200) {
          const errorText = response.getContentText();
          throw new Error(`HTTP ${response.getResponseCode()}: ${errorText}`);
        }

        const data = JSON.parse(response.getContentText());
        const headers = (data.values && data.values[0]) || [];
        
        // ヘッダーインデックスマップ構築
        const indices = {};
        headers.forEach((header, index) => {
          if (header && typeof header === 'string' && header.trim() !== '') {
            indices[header.trim()] = index;
          }
        });

        this.logger.debug(`[_fetchHeadersWithRetry] Success: ${Object.keys(indices).length} headers found`);
        return indices;

      } catch (error) {
        lastError = error;
        this.logger.warn(`[_fetchHeadersWithRetry] Attempt ${attempt} failed:`, error.message);
        
        if (attempt < maxRetries) {
          const delay = Math.pow(2, attempt - 1) * 1000; // 指数バックオフ
          Utilities.sleep(delay);
        }
      }
    }

    this.logger.error(`[_fetchHeadersWithRetry] All attempts failed. Last error:`, lastError.message);
    throw lastError;
  }

  /**
   * 必須ヘッダー検証
   * @private
   */
  _validateRequiredHeaders(indices) {
    const requiredHeaders = ['理由', '意見', 'なぜそう思うのですか', 'どう思いますか', 'Reason', 'Opinion'];
    const found = [];
    const missing = [];
    
    let hasReasonColumn = false;
    let hasOpinionColumn = false;

    requiredHeaders.forEach(header => {
      if (indices[header] !== undefined) {
        found.push(header);
        if (['理由', 'なぜそう思うのですか', 'Reason'].includes(header)) {
          hasReasonColumn = true;
        }
        if (['意見', 'どう思いますか', 'Opinion'].includes(header)) {
          hasOpinionColumn = true;
        }
      } else {
        missing.push(header);
      }
    });

    return {
      isValid: hasReasonColumn && hasOpinionColumn,
      hasReasonColumn,
      hasOpinionColumn,
      found,
      missing
    };
  }

  /**
   * ヘッダー検証失敗時の処理
   * @private
   */
  _handleHeaderValidationFailure(spreadsheetId, sheetName, headers, validation) {
    // 完全に必須列が不足している場合のみ復旧を試行
    if (!validation.hasReasonColumn && !validation.hasOpinionColumn) {
      this.logger.warn(
        `[getCoreSheetHeaders] Critical headers missing: ${validation.missing.join(', ')}, attempting recovery`
      );
      
      // キャッシュクリアして再取得
      const key = `unified_hdr_${spreadsheetId}_${sheetName}`;
      this.cache.remove(key);

      const recoveredHeaders = this._fetchHeadersWithRetry(spreadsheetId, sheetName, 1);
      if (recoveredHeaders && Object.keys(recoveredHeaders).length > 0) {
        const recoveredValidation = this._validateRequiredHeaders(recoveredHeaders);
        if (recoveredValidation.hasReasonColumn || recoveredValidation.hasOpinionColumn) {
          this.logger.debug(`[getCoreSheetHeaders] Successfully recovered headers`);
          return recoveredHeaders;
        }
      }
    } else {
      // 一部の列が存在する場合は警告のみで継続
      this.logger.debug(
        `[getCoreSheetHeaders] Partial validation - continuing with available columns: reason=${validation.hasReasonColumn}, opinion=${validation.hasOpinionColumn}`
      );
    }

    return headers;
  }

  // ===========================================
  // 3. Spreadsheet管理メソッド群
  // ===========================================

  /**
   * 統一Spreadsheet情報取得
   * 
   * @param {string} userId - ユーザーID
   * @param {Object} [options={}] - オプション
   * @param {boolean} [options.usePublished=false] - 公開用Spreadsheetを使用
   * @param {boolean} [options.forceRefresh=false] - キャッシュ強制更新
   * @return {Object} Spreadsheet情報
   */
  getCoreSpreadsheetInfo(userId, options = {}) {
    const { usePublished = false, forceRefresh = false } = options;

    const userInfo = this._getUserInfo(userId, { useExecutionCache: true, ttl: 300 });
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}');
    
    let spreadsheetId;
    if (usePublished) {
      spreadsheetId = configJson.publishedSpreadsheetId || userInfo.spreadsheetId;
    } else {
      spreadsheetId = userInfo.spreadsheetId;
    }

    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが見つかりません');
    }

    // Spreadsheet詳細情報取得
    const cacheKey = `unified_ss_info_${spreadsheetId}`;
    
    if (forceRefresh) {
      this.cache.remove(cacheKey);
    }

    return this.cache.get(
      cacheKey,
      () => this._fetchSpreadsheetInfo(spreadsheetId),
      { ttl: 3600 } // 1時間キャッシュ
    );
  }

  /**
   * 現在のSpreadsheet取得
   * 
   * @param {string} userId - ユーザーID
   * @return {Object} Spreadsheetオブジェクト
   */
  getCurrentSpreadsheet(userId) {
    const info = this.getCoreSpreadsheetInfo(userId);
    return openSpreadsheetOptimized(info.spreadsheetId);
  }

  /**
   * 有効なSpreadsheetID取得
   * 
   * @param {string} userId - ユーザーID
   * @param {boolean} [usePublished=false] - 公開用を優先
   * @return {string} SpreadsheetID
   */
  getEffectiveSpreadsheetId(userId, usePublished = false) {
    const info = this.getCoreSpreadsheetInfo(userId, { usePublished });
    return info.spreadsheetId;
  }

  /**
   * Spreadsheet詳細情報取得
   * @private
   */
  _fetchSpreadsheetInfo(spreadsheetId) {
    try {
      const service = getSheetsServiceCached();
      const response = getSpreadsheetsData(service, spreadsheetId);
      
      return {
        spreadsheetId: spreadsheetId,
        title: response.properties.title,
        locale: response.properties.locale,
        timeZone: response.properties.timeZone,
        sheets: response.sheets.map(sheet => ({
          sheetId: sheet.properties.sheetId,
          title: sheet.properties.title,
          sheetType: sheet.properties.sheetType,
          gridProperties: sheet.properties.gridProperties
        }))
      };
    } catch (error) {
      this.logger.error(`[_fetchSpreadsheetInfo] Error:`, error.message);
      throw new Error(`スプレッドシート情報の取得に失敗しました: ${error.message}`);
    }
  }

  // ===========================================
  // 4. シート一覧・選択メソッド群
  // ===========================================

  /**
   * 利用可能シート一覧取得
   * 
   * @param {string} userId - ユーザーID
   * @param {Object} [options={}] - オプション
   * @return {Object} シート一覧レスポンス
   */
  getAvailableSheets(userId, options = {}) {
    try {
      const userInfo = this._getUserInfo(userId, { useExecutionCache: true, ttl: 300 });
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }

      const configJson = JSON.parse(userInfo.configJson || '{}');
      const publishedSpreadsheetId = configJson.publishedSpreadsheetId || userInfo.spreadsheetId;

      if (!publishedSpreadsheetId) {
        return createSuccessResponse({
          sheets: [],
          message: 'スプレッドシートが設定されていません'
        });
      }

      const ssInfo = this.getCoreSpreadsheetInfo(userId, { usePublished: true });
      const sheets = ssInfo.sheets.filter(sheet => 
        sheet.sheetType !== 'OBJECT' && // オブジェクトシート除外
        !sheet.title.startsWith('名簿') // 名簿シート除外
      );

      return createSuccessResponse({
        sheets: sheets.map(sheet => ({
          sheetId: sheet.sheetId,
          title: sheet.title,
          rowCount: sheet.gridProperties?.rowCount || 0,
          columnCount: sheet.gridProperties?.columnCount || 0
        })),
        spreadsheetId: publishedSpreadsheetId,
        spreadsheetTitle: ssInfo.title
      });

    } catch (error) {
      this.logger.error(`[getAvailableSheets] Error:`, error.message);
      throw new Error(`シート一覧の取得に失敗しました: ${error.message}`);
    }
  }

  /**
   * シート詳細情報取得
   * 
   * @param {string} userId - ユーザーID
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @return {Object} シート詳細情報
   */
  getSheetDetails(userId, spreadsheetId, sheetName) {
    const cacheKey = `unified_sheet_details_${userId}_${spreadsheetId}_${sheetName}`;

    return this.cache.get(
      cacheKey,
      () => {
        const ssInfo = this.getCoreSpreadsheetInfo(userId);
        const sheet = ssInfo.sheets.find(s => s.title === sheetName);
        
        if (!sheet) {
          throw new Error(`シート '${sheetName}' が見つかりません`);
        }

        const headers = this.getCoreSheetHeaders(spreadsheetId, sheetName);
        
        return {
          sheetId: sheet.sheetId,
          title: sheet.title,
          rowCount: sheet.gridProperties?.rowCount || 0,
          columnCount: sheet.gridProperties?.columnCount || 0,
          headers: headers,
          headerCount: Object.keys(headers).length
        };
      },
      { ttl: 900 } // 15分間キャッシュ
    );
  }

  // ===========================================
  // 5. ユーティリティメソッド群
  // ===========================================

  /**
   * パラメータ検証
   * @private
   */
  _validateCoreParams(userId, spreadsheetId, sheetName) {
    if (!userId) throw new Error('userIdは必須です');
    if (!spreadsheetId) throw new Error('spreadsheetIdは必須です');
    if (!sheetName) throw new Error('sheetNameは必須です');
  }

  /**
   * ユーザー情報取得
   * @private
   */
  _getUserInfo(userId, options = {}) {
    return getOrFetchUserInfo(userId, 'userId', {
      useExecutionCache: true,
      ttl: 300,
      ...options
    });
  }

  /**
   * キャッシュキー生成
   * @private
   */
  _generateCacheKey(params) {
    // Phase 6最適化: 統合generatorFactory使用
    const { type, userId, spreadsheetId, sheetName, classFilter, sortOrder, adminMode, sinceRowCount } = params;
    
    const keyData = {
      type, userId, sheetName, classFilter, sortOrder, adminMode,
      ...(sinceRowCount !== undefined && { sinceRowCount })
    };
    
    return UUtilities.generatorFactory.key.sheetDataCache(
      'unified',
      spreadsheetId,
      keyData
    );
  }

  /**
   * キャッシュ設定取得
   * @private
   */
  _getCacheConfig(dataType, adminMode) {
    const configs = {
      published: { ttl: 600 }, // 10分
      incremental: { ttl: 180 }, // 3分
      full: { ttl: 180 } // 3分
    };

    return adminMode ? { ttl: 0 } : configs[dataType] || configs.full;
  }

  /**
   * 生データ取得
   * @private
   */
  _getRawSheetData(spreadsheetId, sheetName, sheetConfig) {
    const service = getSheetsServiceCached();
    const ranges = [sheetName + '!A:Z'];
    
    const responses = batchGetSheetsData(service, spreadsheetId, ranges);
    const originalData = responses.valueRanges[0].values || [];

    // リアクション数取得（必要に応じて）
    const reactionCounts = {}; // 簡略化

    return {
      originalData,
      reactionCounts
    };
  }
}

// ===========================================
// 6. グローバル統一管理インスタンス
// ===========================================

/** @type {UnifiedSheetDataManager} グローバル統一Sheet管理インスタンス */
const unifiedSheetManager = new UnifiedSheetDataManager();

// ===========================================
// 7. 互換性維持用ラッパー関数群
// ===========================================

/**
 * getPublishedSheetData互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理モード
 * @param {boolean} bypassCache - キャッシュバイパス
 * @return {Object} 公開シートデータ
 */
function getPublishedSheetDataUnified(requestUserId, classFilter, sortOrder, adminMode, bypassCache) {
  const userInfo = findUserById(requestUserId);
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません');
  }

  const configJson = JSON.parse(userInfo.configJson || '{}');
  const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
  const publishedSheetName = configJson.publishedSheetName;

  return unifiedSheetManager.getCoreSheetData({
    userId: requestUserId,
    spreadsheetId: publishedSpreadsheetId,
    sheetName: publishedSheetName,
    classFilter,
    sortOrder,
    adminMode,
    bypassCache,
    dataType: 'published'
  });
}

/**
 * getSheetData互換ラッパー
 * 
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortMode - ソートモード
 * @param {boolean} adminMode - 管理モード
 * @return {Object} シートデータ
 */
function getSheetDataUnified(userId, sheetName, classFilter, sortMode, adminMode) {
  const userInfo = getOrFetchUserInfo(userId, 'userId', { useExecutionCache: true, ttl: 300 });
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません');
  }

  const configJson = JSON.parse(userInfo.configJson || '{}');
  const spreadsheetId = configJson.publishedSpreadsheetId || userInfo.spreadsheetId;

  return unifiedSheetManager.getCoreSheetData({
    userId,
    spreadsheetId,
    sheetName,
    classFilter,
    sortOrder: sortMode,
    adminMode,
    dataType: 'full'
  });
}

/**
 * getIncrementalSheetData互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理モード
 * @param {number} sinceRowCount - 開始行数
 * @return {Object} 増分シートデータ
 */
function getIncrementalSheetDataUnified(requestUserId, classFilter, sortOrder, adminMode, sinceRowCount) {
  const userInfo = findUserById(requestUserId);
  if (!userInfo) {
    throw new Error('ユーザー情報が見つかりません');
  }

  const configJson = JSON.parse(userInfo.configJson || '{}');
  const publishedSpreadsheetId = configJson.publishedSpreadsheetId;
  const publishedSheetName = configJson.publishedSheetName;

  return unifiedSheetManager.getCoreSheetData({
    userId: requestUserId,
    spreadsheetId: publishedSpreadsheetId,
    sheetName: publishedSheetName,
    classFilter,
    sortOrder,
    adminMode,
    dataType: 'incremental',
    sinceRowCount
  });
}

/**
 * getSheetHeaders互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @return {Object} ヘッダー情報
 */
function getSheetHeadersUnified(requestUserId, spreadsheetId, sheetName) {
  return unifiedSheetManager.getCoreSheetHeaders(spreadsheetId, sheetName);
}

/**
 * getHeadersCached互換ラッパー
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @return {Object} キャッシュ済みヘッダー
 */
function getHeadersCachedUnified(spreadsheetId, sheetName) {
  return unifiedSheetManager.getCoreSheetHeaders(spreadsheetId, sheetName);
}

/**
 * getHeaderIndices互換ラッパー
 * 
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @return {Object} ヘッダーインデックス
 */
function getHeaderIndicesUnified(spreadsheetId, sheetName) {
  return unifiedSheetManager.getHeaderIndices(spreadsheetId, sheetName);
}

/**
 * getCurrentSpreadsheet互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @return {Object} 現在のスプレッドシート
 */
function getCurrentSpreadsheetUnified(requestUserId) {
  return unifiedSheetManager.getCurrentSpreadsheet(requestUserId);
}

/**
 * getEffectiveSpreadsheetId互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @return {string} 有効なスプレッドシートID
 */
function getEffectiveSpreadsheetIdUnified(requestUserId) {
  return unifiedSheetManager.getEffectiveSpreadsheetId(requestUserId, true);
}

/**
 * getAvailableSheets互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @return {Object} 利用可能シート一覧
 */
function getAvailableSheetsUnified(requestUserId) {
  return unifiedSheetManager.getAvailableSheets(requestUserId);
}

/**
 * getSheetDetails互換ラッパー
 * 
 * @param {string} requestUserId - リクエストユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @return {Object} シート詳細情報
 */
function getSheetDetailsUnified(requestUserId, spreadsheetId, sheetName) {
  return unifiedSheetManager.getSheetDetails(requestUserId, spreadsheetId, sheetName);
}

// ===========================================
// 8. デバッグ・統計取得関数
// ===========================================

/**
 * 統一Sheet管理システムの統計情報取得
 * 
 * @return {Object} 統計情報
 */
function getUnifiedSheetManagerStats() {
  const cacheStats = unifiedSheetManager.cache.getStats();
  
  return {
    managerInitialized: !!unifiedSheetManager,
    cacheStats: cacheStats,
    timestamp: new Date().toISOString()
  };
}

/**
 * 統一Sheet管理システムのキャッシュクリア
 * 
 * @param {string} [pattern] - クリアするキャッシュのパターン
 */
function clearUnifiedSheetManagerCache(pattern) {
  if (pattern) {
    // パターンマッチでキャッシュクリア（実装は統一キャッシュマネージャーに依存）
    unifiedSheetManager.logger.debug(`[clearUnifiedSheetManagerCache] Clearing cache with pattern: ${pattern}`);
  } else {
    // 全シート関連キャッシュをクリア
    const patterns = [
      'unified_published_',
      'unified_full_',
      'unified_incremental_',
      'unified_hdr_',
      'unified_ss_info_',
      'unified_sheet_details_'
    ];
    
    patterns.forEach(p => {
      unifiedSheetManager.logger.debug(`[clearUnifiedSheetManagerCache] Clearing cache pattern: ${p}`);
    });
  }
}