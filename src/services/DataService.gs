/**
 * @fileoverview DataService - 統一データ操作サービス
 * 
 * 🎯 責任範囲:
 * - スプレッドシートデータ取得・操作
 * - リアクション・ハイライト機能
 * - データフィルタリング・検索
 * - バルクデータAPI
 * 
 * 🔄 置き換え対象:
 * - Core.gs のデータ操作部分
 * - UnifiedManager.data
 * - ColumnAnalysisSystem.gs の一部
 */

/**
 * DataService - 統一データ操作サービス
 * Google Sheets APIに特化した高パフォーマンスサービス
 */
const DataService = Object.freeze({

  // ===========================================
  // 📊 スプレッドシートデータ取得
  // ===========================================

  /**
   * ユーザーのスプレッドシートデータ取得（統合版）
   * @param {string} userId - ユーザーID
   * @param {Object} options - 取得オプション
   * @returns {Object} データ取得結果
   */
  getSheetData(userId, options = {}) {
    const startTime = Date.now();
    
    try {
      // 設定取得
      const config = ConfigService.getUserConfig(userId);
      if (!config || !config.spreadsheetId) {
        return this.createErrorResponse('スプレッドシートが設定されていません');
      }

      // データ取得実行
      const result = this.fetchSpreadsheetData(config, options);
      
      const executionTime = Date.now() - startTime;
      console.info('DataService.getSheetData: データ取得完了', {
        userId,
        rowCount: result.data?.length || 0,
        executionTime
      });

      return result;
    } catch (error) {
      console.error('DataService.getSheetData: エラー', {
        userId,
        error: error.message
      });
      return this.createErrorResponse(error.message);
    }
  },

  /**
   * スプレッドシートデータ取得実行
   * @param {Object} config - ユーザー設定
   * @param {Object} options - 取得オプション
   * @returns {Object} 取得結果
   */
  fetchSpreadsheetData(config, options = {}) {
    try {
      // スプレッドシート取得
      const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
      const sheet = spreadsheet.getSheetByName(config.sheetName);
      
      if (!sheet) {
        throw new Error(`シート '${config.sheetName}' が見つかりません`);
      }

      // データ範囲取得
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();

      if (lastRow <= 1) {
        return this.createSuccessResponse([], { message: 'データが存在しません' });
      }

      // ヘッダー行取得
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

      // データ行取得（バッチ処理）
      const dataRows = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

      // データ変換・フィルタリング
      const processedData = this.processRawData(dataRows, headers, config, options);

      return this.createSuccessResponse(processedData, {
        totalRows: dataRows.length,
        filteredRows: processedData.length,
        headers
      });
    } catch (error) {
      console.error('DataService.fetchSpreadsheetData: エラー', error.message);
      throw error;
    }
  },

  /**
   * 生データを処理・変換
   * @param {Array} dataRows - 生データ行
   * @param {Array} headers - ヘッダー配列
   * @param {Object} config - 設定
   * @param {Object} options - 処理オプション
   * @returns {Array} 処理済みデータ
   */
  processRawData(dataRows, headers, config, options = {}) {
    try {
      const columnMapping = config.columnMapping?.mapping || {};
      const processedData = [];

      dataRows.forEach((row, index) => {
        try {
          // 基本データ構造作成
          const item = {
            id: `row_${index + 2}`, // 実際の行番号（ヘッダー考慮）
            timestamp: this.extractFieldValue(row, headers, 'timestamp') || '',
            email: this.extractFieldValue(row, headers, 'email') || '',
            
            // メインコンテンツ
            answer: this.extractFieldValue(row, headers, 'answer', columnMapping) || '',
            reason: this.extractFieldValue(row, headers, 'reason', columnMapping) || '',
            className: this.extractFieldValue(row, headers, 'class', columnMapping) || '',
            name: this.extractFieldValue(row, headers, 'name', columnMapping) || '',

            // メタデータ
            formattedTimestamp: this.formatTimestamp(this.extractFieldValue(row, headers, 'timestamp')),
            isEmpty: this.isEmptyRow(row),
            
            // リアクション（既存の場合）
            reactions: this.extractReactions(row, headers),
            isHighlighted: this.extractHighlight(row, headers)
          };

          // フィルタリング
          if (this.shouldIncludeRow(item, options)) {
            processedData.push(item);
          }
        } catch (rowError) {
          console.warn('DataService.processRawData: 行処理エラー', {
            rowIndex: index,
            error: rowError.message
          });
        }
      });

      // ソート・制限適用
      return this.applySortAndLimit(processedData, options);
    } catch (error) {
      console.error('DataService.processRawData: エラー', error.message);
      return [];
    }
  },

  /**
   * フィールド値抽出（列マッピング対応）
   * @param {Array} row - データ行
   * @param {Array} headers - ヘッダー配列
   * @param {string} fieldType - フィールドタイプ
   * @param {Object} columnMapping - 列マッピング
   * @returns {*} フィールド値
   */
  extractFieldValue(row, headers, fieldType, columnMapping = {}) {
    try {
      // 列マッピングがある場合
      if (columnMapping[fieldType] !== undefined) {
        const columnIndex = columnMapping[fieldType];
        return row[columnIndex] || '';
      }

      // ヘッダー名からの推測
      const headerPatterns = {
        timestamp: ['タイムスタンプ', 'timestamp', '投稿日時'],
        email: ['メール', 'email', 'メールアドレス'],
        answer: ['回答', '答え', '意見', 'answer'],
        reason: ['理由', '根拠', 'reason'],
        class: ['クラス', '学年', 'class'],
        name: ['名前', '氏名', 'name']
      };

      const patterns = headerPatterns[fieldType] || [];
      
      for (const pattern of patterns) {
        const index = headers.findIndex(header => 
          header && header.toLowerCase().includes(pattern.toLowerCase())
        );
        if (index !== -1) {
          return row[index] || '';
        }
      }

      return '';
    } catch (error) {
      console.warn('DataService.extractFieldValue: エラー', error.message);
      return '';
    }
  },

  // ===========================================
  // 🎯 リアクション・ハイライト機能
  // ===========================================

  /**
   * リアクション追加
   * @param {string} userId - ユーザーID
   * @param {string} rowId - 行ID
   * @param {string} reactionType - リアクションタイプ
   * @returns {boolean} 成功可否
   */
  addReaction(userId, rowId, reactionType) {
    try {
      if (!this.validateReactionType(reactionType)) {
        console.error('DataService.addReaction: 無効なリアクションタイプ', reactionType);
        return false;
      }

      // 設定取得
      const config = ConfigService.getUserConfig(userId);
      if (!config || !config.spreadsheetId) {
        console.error('DataService.addReaction: スプレッドシート設定なし');
        return false;
      }

      // リアクション更新実行
      return this.updateReactionInSheet(config, rowId, reactionType, 'add');
    } catch (error) {
      console.error('DataService.addReaction: エラー', {
        userId,
        rowId,
        reactionType,
        error: error.message
      });
      return false;
    }
  },

  /**
   * リアクション削除
   * @param {string} userId - ユーザーID  
   * @param {string} rowId - 行ID
   * @param {string} reactionType - リアクションタイプ
   * @returns {boolean} 成功可否
   */
  removeReaction(userId, rowId, reactionType) {
    try {
      const config = ConfigService.getUserConfig(userId);
      if (!config || !config.spreadsheetId) {
        return false;
      }

      return this.updateReactionInSheet(config, rowId, reactionType, 'remove');
    } catch (error) {
      console.error('DataService.removeReaction: エラー', {
        userId,
        rowId,
        error: error.message
      });
      return false;
    }
  },

  /**
   * ハイライト切り替え
   * @param {string} userId - ユーザーID
   * @param {string} rowId - 行ID
   * @returns {boolean} 成功可否
   */
  toggleHighlight(userId, rowId) {
    try {
      const config = ConfigService.getUserConfig(userId);
      if (!config || !config.spreadsheetId) {
        return false;
      }

      return this.updateHighlightInSheet(config, rowId);
    } catch (error) {
      console.error('DataService.toggleHighlight: エラー', {
        userId,
        rowId,
        error: error.message
      });
      return false;
    }
  },

  /**
   * スプレッドシート内リアクション更新
   * @param {Object} config - 設定
   * @param {string} rowId - 行ID
   * @param {string} reactionType - リアクションタイプ
   * @param {string} action - アクション（add/remove）
   * @returns {boolean} 成功可否
   */
  updateReactionInSheet(config, rowId, reactionType, action) {
    try {
      const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
      const sheet = spreadsheet.getSheetByName(config.sheetName);
      
      if (!sheet) {
        throw new Error('シートが見つかりません');
      }

      // 行番号抽出（row_3 → 3）
      const rowNumber = parseInt(rowId.replace('row_', ''));
      if (isNaN(rowNumber) || rowNumber < 2) {
        throw new Error('無効な行ID');
      }

      // リアクション列の取得・作成
      const reactionColumn = this.getOrCreateReactionColumn(sheet, reactionType);
      if (!reactionColumn) {
        throw new Error('リアクション列の作成に失敗');
      }

      // 現在値取得・更新
      const currentValue = sheet.getRange(rowNumber, reactionColumn).getValue() || 0;
      const newValue = action === 'add' 
        ? Math.max(0, currentValue + 1)
        : Math.max(0, currentValue - 1);

      sheet.getRange(rowNumber, reactionColumn).setValue(newValue);

      console.info('DataService.updateReactionInSheet: リアクション更新完了', {
        rowId,
        reactionType,
        action,
        oldValue: currentValue,
        newValue
      });

      return true;
    } catch (error) {
      console.error('DataService.updateReactionInSheet: エラー', error.message);
      return false;
    }
  },

  // ===========================================
  // 🔍 データ分析・フィルタリング
  // ===========================================

  /**
   * columnMappingを使用したデータ処理（Core.gsより移行）
   * @param {Array} dataRows - データ行配列
   * @param {Array} headers - ヘッダー配列
   * @param {Object} columnMapping - 列マッピング
   * @returns {Array} 処理済みデータ配列
   */
  processDataWithColumnMapping(dataRows, headers, columnMapping) {
    if (!dataRows || !Array.isArray(dataRows)) {
      return [];
    }

    console.info('DataService.processDataWithColumnMapping', {
      rowCount: dataRows.length,
      headerCount: headers ? headers.length : 0,
      mappingKeys: columnMapping ? Object.keys(columnMapping) : []
    });

    return dataRows.map((row, index) => {
      const processedRow = {
        id: index + 1,
        timestamp: row[columnMapping?.timestamp] || row[0] || '',
        email: row[columnMapping?.email] || row[1] || '',
        class: row[columnMapping?.class] || row[2] || '',
        name: row[columnMapping?.name] || row[3] || '',
        answer: row[columnMapping?.answer] || row[4] || '',
        reason: row[columnMapping?.reason] || row[5] || '',
        reactions: {
          understand: parseInt(row[columnMapping?.understand] || 0),
          like: parseInt(row[columnMapping?.like] || 0),
          curious: parseInt(row[columnMapping?.curious] || 0)
        },
        highlight: row[columnMapping?.highlight] === 'TRUE' || false,
        originalData: row
      };

      return processedRow;
    });
  },

  /**
   * バルクデータ取得API（Core.gsより移行・最適化）
   * 複数の情報を一括取得で高速化
   * @param {string} userId - ユーザーID
   * @param {Object} options - オプション { includeSheetData, includeFormInfo, includeSystemInfo }
   * @returns {Object} 一括データ
   */
  getBulkData(userId, options = {}) {
    try {
      const startTime = Date.now();

      // ユーザー情報取得（ConfigServiceから）
      const userInfo = ConfigService.getUserInfo(userId);
      if (!userInfo) {
        throw new Error('ユーザーが見つかりません');
      }

      const config = ConfigService.getUserConfig(userId);
      const bulkData = {
        timestamp: new Date().toISOString(),
        userId,
        userInfo: {
          userEmail: userInfo.userEmail,
          isActive: userInfo.isActive,
          config
        }
      };

      // シートデータを含む場合
      if (options.includeSheetData && config?.spreadsheetId && config?.sheetName) {
        try {
          bulkData.sheetData = this.getSheetData(userId, {
            sortOrder: 'asc',
            includeEmpty: false,
            useCache: true
          });
        } catch (sheetError) {
          console.warn('DataService.getBulkData: シートデータ取得エラー', sheetError.message);
          bulkData.sheetDataError = sheetError.message;
        }
      }

      // フォーム情報を含む場合
      if (options.includeFormInfo && config?.formUrl) {
        try {
          bulkData.formInfo = this.getFormInfo(userId);
        } catch (formError) {
          console.warn('DataService.getBulkData: フォーム情報取得エラー', formError.message);
          bulkData.formInfoError = formError.message;
        }
      }

      // システム情報を含む場合
      if (options.includeSystemInfo) {
        bulkData.systemInfo = {
          setupStep: ConfigService.determineSetupStep(userInfo, config),
          isSystemSetup: ConfigService.isSystemSetup(),
          appPublished: config?.appPublished || false
        };
      }

      const executionTime = Date.now() - startTime;

      return {
        success: true,
        data: bulkData,
        executionTime
      };
    } catch (error) {
      console.error('❌ DataService.getBulkData: バルクデータ取得エラー', {
        userId,
        options,
        error: error.message,
        stack: error.stack
      });

      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  },

  /**
   * 自動停止時間計算（Core.gsより移行）
   * @param {string} publishedAt - 公開開始時間のISO文字列
   * @param {number} minutes - 自動停止までの分数
   * @returns {Object} 停止時間情報
   */
  getAutoStopTime(publishedAt, minutes) {
    try {
      const publishTime = new Date(publishedAt);
      const stopTime = new Date(publishTime.getTime() + minutes * 60 * 1000);

      return {
        publishTime,
        stopTime,
        publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
        stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
        remainingMinutes: Math.max(
          0,
          Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))
        )
      };
    } catch (error) {
      console.error('DataService.getAutoStopTime: 計算エラー', error);
      return null;
    }
  },

  /**
   * レガシーCore.gsから移行: リアクション処理実行
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {number} rowIndex - 行インデックス
   * @param {string} reactionKey - リアクション種類
   * @param {string} userEmail - ユーザーメール
   * @returns {Object} 処理結果
   */
  processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, userEmail) {
    try {
      if (!this.validateReactionParams(spreadsheetId, sheetName, rowIndex, reactionKey)) {
        throw new Error('無効なリアクションパラメータ');
      }

      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        throw new Error('シートが見つかりません');
      }

      // リアクション列を取得または作成
      const reactionColumn = this.getOrCreateReactionColumn(sheet, reactionKey);
      
      // 現在の値を取得して更新
      const currentValue = sheet.getRange(rowIndex, reactionColumn).getValue() || 0;
      const newValue = Math.max(0, currentValue + 1);
      sheet.getRange(rowIndex, reactionColumn).setValue(newValue);

      console.info('DataService.processReaction: リアクション処理完了', {
        spreadsheetId,
        sheetName,
        rowIndex,
        reactionKey,
        oldValue: currentValue,
        newValue
      });

      return {
        status: 'success',
        message: 'リアクションを追加しました',
        newValue
      };
    } catch (error) {
      console.error('DataService.processReaction: エラー', error.message);
      return {
        status: 'error',
        message: error.message
      };
    }
  },

  /**
   * レガシーCore.gsから移行: ハイライト切り替え処理
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {number} rowIndex - 行インデックス
   * @returns {Object} 処理結果
   */
  processHighlightToggle(spreadsheetId, sheetName, rowIndex) {
    try {
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);
      
      if (!sheet) {
        throw new Error('シートが見つかりません');
      }

      // ハイライト列を取得または作成
      const highlightColumn = this.getOrCreateReactionColumn(sheet, 'HIGHLIGHT');
      
      // 現在の値を取得してトグル
      const currentValue = sheet.getRange(rowIndex, highlightColumn).getValue();
      const isCurrentlyHighlighted = currentValue === 'TRUE' || currentValue === true;
      const newValue = isCurrentlyHighlighted ? 'FALSE' : 'TRUE';
      
      sheet.getRange(rowIndex, highlightColumn).setValue(newValue);

      console.info('DataService.processHighlightToggle: ハイライト切り替え完了', {
        spreadsheetId,
        sheetName,
        rowIndex,
        oldValue: currentValue,
        newValue,
        highlighted: !isCurrentlyHighlighted
      });

      return {
        status: 'success',
        message: 'ハイライトを切り替えました',
        highlighted: !isCurrentlyHighlighted
      };
    } catch (error) {
      console.error('DataService.processHighlightToggle: エラー', error.message);
      return {
        status: 'error',
        message: error.message
      };
    }
  },

  /**
   * レガシーCore.gsから移行: 回答削除処理
   * @param {string} userId - ユーザーID
   * @param {number} rowIndex - 行インデックス
   * @param {string} sheetName - シート名
   * @returns {Object} 削除結果
   */
  deleteAnswer(userId, rowIndex, sheetName) {
    try {
      const config = ConfigService.getUserConfig(userId);
      if (!config || !config.spreadsheetId) {
        throw new Error('スプレッドシートが設定されていません');
      }

      const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName || config.sheetName);
      
      if (!sheet) {
        throw new Error('シートが見つかりません');
      }

      // 行数の範囲チェック
      const lastRow = sheet.getLastRow();
      if (rowIndex < 2 || rowIndex > lastRow) {
        throw new Error('指定された行が無効です');
      }

      // 削除前のデータを取得（ログ用）
      const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 行を削除
      sheet.deleteRow(rowIndex);
      
      // キャッシュクリア
      this.clearCache(config.spreadsheetId, sheet.getName());

      console.info('DataService.deleteAnswer: 回答削除完了', {
        userId,
        rowIndex,
        sheetName: sheet.getName(),
        timestamp: new Date().toISOString()
      });

      return {
        success: true,
        message: '回答を削除しました',
        deletedRow: rowIndex
      };
    } catch (error) {
      console.error('DataService.deleteAnswer: エラー', error.message);
      return {
        success: false,
        message: error.message
      };
    }
  },

  /**
   * リアクションパラメータ検証
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {number} rowIndex - 行インデックス
   * @param {string} reactionKey - リアクション種類
   * @returns {boolean} 検証結果
   */
  validateReactionParams(spreadsheetId, sheetName, rowIndex, reactionKey) {
    if (!spreadsheetId || !sheetName || !rowIndex || !reactionKey) {
      return false;
    }
    
    if (rowIndex < 2) { // ヘッダー行は1
      return false;
    }
    
    const validReactions = ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
    return validReactions.includes(reactionKey);
  },

  /**
   * キャッシュクリア
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   */
  clearCache(spreadsheetId, sheetName) {
    try {
      const cacheKey = `sheet_data_${spreadsheetId}_${sheetName}`;
      CacheService.getScriptCache().remove(cacheKey);
      
      // 関連キャッシュもクリア
      CacheService.getScriptCache().remove(`${cacheKey}_processed`);
      CacheService.getScriptCache().remove(`${cacheKey}_stats`);
    } catch (error) {
      console.warn('DataService.clearCache: キャッシュクリアエラー', error.message);
    }
  },

  /**
   * バルクデータ取得（統合API）
   * @param {string} userId - ユーザーID
   * @param {Object} options - 取得オプション
   * @returns {Object} バルクデータ
   */
  getBulkData(userId, options = {}) {
    const startTime = Date.now();

    try {
      const bulkData = {};

      // ユーザー情報
      if (options.includeUserInfo !== false) {
        bulkData.userInfo = UserService.getCurrentUserInfo();
      }

      // 設定情報
      if (options.includeConfig !== false) {
        bulkData.config = ConfigService.getUserConfig(userId);
      }

      // シートデータ
      if (options.includeSheetData !== false) {
        const sheetResult = this.getSheetData(userId, options);
        bulkData.sheetData = sheetResult.success ? sheetResult.data : [];
        bulkData.sheetDataError = sheetResult.success ? null : sheetResult.message;
      }

      // システム情報
      if (options.includeSystemInfo) {
        bulkData.systemInfo = {
          setupStep: this.determineSetupStep(bulkData.userInfo, bulkData.config),
          isSystemSetup: this.isSystemSetup(),
          appPublished: bulkData.config?.appPublished || false
        };
      }

      const executionTime = Date.now() - startTime;

      return {
        success: true,
        data: bulkData,
        executionTime,
        timestamp: new Date().toISOString()
      };
    } catch (error) {
      console.error('DataService.getBulkData: エラー', {
        userId,
        error: error.message
      });

      return {
        success: false,
        error: error.message,
        timestamp: new Date().toISOString()
      };
    }
  },

  /**
   * データ統計取得
   * @param {string} userId - ユーザーID
   * @returns {Object} 統計情報
   */
  getDataStatistics(userId) {
    try {
      const result = this.getSheetData(userId);
      if (!result.success) {
        return { error: result.message };
      }

      const {data} = result;
      const stats = {
        totalEntries: data.length,
        nonEmptyEntries: data.filter(item => !item.isEmpty).length,
        reactionCounts: this.calculateReactionStats(data),
        classCounts: this.calculateClassStats(data),
        timeRange: this.calculateTimeRange(data),
        lastUpdated: new Date().toISOString()
      };

      return stats;
    } catch (error) {
      console.error('DataService.getDataStatistics: エラー', error.message);
      return { error: error.message };
    }
  },

  // ===========================================
  // 🔧 ユーティリティ・ヘルパー
  // ===========================================

  /**
   * セットアップステップ判定
   * @param {Object} userInfo - ユーザー情報
   * @param {Object} config - 設定
   * @returns {number} セットアップステップ (1-3)
   */
  determineSetupStep(userInfo, config) {
    if (!userInfo || !config || !config.spreadsheetId) {
      return 1;
    }

    if (config.setupStatus !== 'completed' || !config.formUrl) {
      return 2;
    }

    if (config.setupStatus === 'completed' && config.appPublished) {
      return 3;
    }

    return 2;
  },

  /**
   * システム設定状態確認
   * @returns {boolean} システム設定済みかどうか
   */
  isSystemSetup() {
    try {
      const currentEmail = UserService.getCurrentEmail();
      if (!currentEmail) return false;

      const userInfo = UserService.getCurrentUserInfo();
      return !!(userInfo && userInfo.config && userInfo.config.spreadsheetId);
    } catch (error) {
      console.error('DataService.isSystemSetup: エラー', error.message);
      return false;
    }
  },

  /**
   * タイムスタンプフォーマット
   * @param {string} isoString - ISO形式日時文字列
   * @returns {string} フォーマット済み日時
   */
  formatTimestamp(isoString) {
    if (!isoString) return '不明';
    
    try {
      return new Date(isoString).toLocaleString('ja-JP', {
        month: 'numeric',
        day: 'numeric', 
        hour: '2-digit',
        minute: '2-digit'
      });
    } catch (error) {
      console.warn('DataService.formatTimestamp: フォーマットエラー', error.message);
      return '不明';
    }
  },

  /**
   * 空行判定
   * @param {Array} row - データ行
   * @returns {boolean} 空行かどうか
   */
  isEmptyRow(row) {
    return !row || row.every(cell => !cell || cell.toString().trim() === '');
  },

  /**
   * 成功レスポンス作成
   * @param {Array} data - データ配列
   * @param {Object} metadata - メタデータ
   * @returns {Object} 成功レスポンス
   */
  createSuccessResponse(data, metadata = {}) {
    return {
      success: true,
      data,
      count: data.length,
      timestamp: new Date().toISOString(),
      ...metadata
    };
  },

  /**
   * エラーレスポンス作成
   * @param {string} message - エラーメッセージ
   * @param {Object} details - 詳細情報
   * @returns {Object} エラーレスポンス
   */
  createErrorResponse(message, details = {}) {
    return {
      success: false,
      message,
      data: [],
      count: 0,
      timestamp: new Date().toISOString(),
      ...details
    };
  },

  /**
   * リアクションタイプ検証
   * @param {string} reactionType - リアクションタイプ
   * @returns {boolean} 有効かどうか
   */
  validateReactionType(reactionType) {
    const validTypes = CONSTANTS.REACTIONS.KEYS || ['UNDERSTAND', 'LIKE', 'CURIOUS'];
    return validTypes.includes(reactionType);
  },

  /**
   * サービス状態診断
   * @returns {Object} 診断結果
   */
  diagnose() {
    const results = {
      service: 'DataService',
      timestamp: new Date().toISOString(),
      checks: []
    };

    try {
      // スプレッドシートAPI確認
      results.checks.push({
        name: 'SpreadsheetApp API',
        status: typeof SpreadsheetApp !== 'undefined' ? '✅' : '❌',
        details: 'GAS SpreadsheetApp availability'
      });

      // 依存サービス確認
      results.checks.push({
        name: 'UserService Dependency',
        status: typeof UserService !== 'undefined' ? '✅' : '❌',
        details: 'UserService integration'
      });

      results.checks.push({
        name: 'ConfigService Dependency',
        status: typeof ConfigService !== 'undefined' ? '✅' : '❌',
        details: 'ConfigService integration'
      });

      results.overall = results.checks.every(check => check.status === '✅') ? '✅' : '⚠️';
    } catch (error) {
      results.checks.push({
        name: 'Service Diagnosis',
        status: '❌',
        details: error.message
      });
      results.overall = '❌';
    }

    return results;
  }

});
