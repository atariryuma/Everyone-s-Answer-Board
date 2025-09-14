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

/* global DB, DataFormatter, CONSTANTS, ResponseFormatter, PropertiesService, Session, PROPS_KEYS, SpreadsheetApp, DriveApp, Utilities, CacheService */

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
      // ✅ GAS Best Practice: 直接DB呼び出し（ConfigService依存除去）
      const user = DB.findUserById(userId);
      if (!user || !user.configJson) {
        console.error('DataService.getSheetData: ユーザー設定が見つかりません', { userId });
        return this.createErrorResponse('ユーザー設定を取得できませんでした');
      }

      const config = JSON.parse(user.configJson);
      if (!config.spreadsheetId) {
        console.warn('DataService.getSheetData: スプレッドシートIDが設定されていません', { userId });
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
      return this.createErrorResponse(error.message || 'データ取得エラー');
    }
  },

  /**
   * 公開されたスプレッドシートデータを取得（API Gateway互換）
   * @param {string} userId - ユーザーID
   * @param {Object} options - 取得オプション
   * @returns {Object} データ取得結果
   */
  getPublishedSheetData(userId, options = {}) {
    try {
      // ✅ GAS Best Practice: 内部関数直接呼び出し
      return this.getSheetData(userId, options);
    } catch (error) {
      console.error('DataService.getPublishedSheetData error:', error);
      return this.createErrorResponse(error.message || '公開データ取得エラー');
    }
  },

  /**
   * スプレッドシートデータ取得実行
   * ✅ バッチ処理対応 - GAS制限対応（実行時間・メモリ制限）
   * @param {Object} config - ユーザー設定
   * @param {Object} options - 取得オプション
   * @returns {Object} 取得結果
   */
  fetchSpreadsheetData(config, options = {}) {
    const startTime = Date.now();
    const MAX_EXECUTION_TIME = 180000; // 3分制限（安全マージン拡大）
    const MAX_BATCH_SIZE = 200; // バッチサイズ削減（メモリ制限対応）
    
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
        return ResponseFormatter.createSuccessResponse([], { message: 'データが存在しません' });
      }

      // ヘッダー行取得
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

      // ✅ 大量データ対応: バッチ処理で安全に取得
      const totalDataRows = lastRow - 1;
      let processedData = [];
      let processedCount = 0;
      
      // バッチごとに処理（メモリ・実行時間制限対応）
      for (let startRow = 2; startRow <= lastRow; startRow += MAX_BATCH_SIZE) {
        // 実行時間チェック
        if (Date.now() - startTime > MAX_EXECUTION_TIME) {
          console.warn('DataService.fetchSpreadsheetData: 実行時間制限のため処理を中断', {
            processedRows: processedCount,
            totalRows: totalDataRows
          });
          break;
        }
        
        const endRow = Math.min(startRow + MAX_BATCH_SIZE - 1, lastRow);
        const batchSize = endRow - startRow + 1;
        
        try {
          // バッチデータ取得
          const batchRows = sheet.getRange(startRow, 1, batchSize, lastCol).getValues();
          
          // バッチ処理実行
          const batchProcessed = this.processRawDataBatch(batchRows, headers, config, options, startRow - 2);
          
          processedData = processedData.concat(batchProcessed);
          processedCount += batchSize;
          
          console.log(`DataService.fetchSpreadsheetData: バッチ処理完了 ${processedCount}/${totalDataRows}`);
          
          // API制限対策: 100行毎に短い休憩
          if (processedCount % 1000 === 0) {
            Utilities.sleep(100); // 0.1秒休憩
          }
          
        } catch (batchError) {
          console.error('DataService.fetchSpreadsheetData: バッチ処理エラー', {
            startRow,
            endRow,
            error: batchError.message
          });
          // バッチエラーは継続（他のバッチは処理）
        }
      }

      const executionTime = Date.now() - startTime;
      console.info('DataService.fetchSpreadsheetData: バッチ処理完了', {
        totalRows: totalDataRows,
        processedRows: processedCount,
        filteredRows: processedData.length,
        executionTime,
        batchCount: Math.ceil(totalDataRows / MAX_BATCH_SIZE)
      });

      return ResponseFormatter.createSuccessResponse(processedData, {
        totalRows: totalDataRows,
        processedRows: processedCount,
        filteredRows: processedData.length,
        headers,
        executionTime,
        wasTruncated: processedCount < totalDataRows
      });
    } catch (error) {
      console.error('DataService.fetchSpreadsheetData: エラー', error.message);
      throw error;
    }
  },

  /**
   * バッチ処理用データ変換（メモリ効率重視）
   * @param {Array} batchRows - バッチデータ行
   * @param {Array} headers - ヘッダー配列
   * @param {Object} config - 設定
   * @param {Object} options - 処理オプション
   * @param {number} startOffset - 開始オフセット（行番号計算用）
   * @returns {Array} 処理済みバッチデータ
   */
  processRawDataBatch(batchRows, headers, config, options = {}, startOffset = 0) {
    try {
      const columnMapping = config.columnMapping?.mapping || {};
      const processedBatch = [];

      batchRows.forEach((row, batchIndex) => {
        try {
          // グローバル行インデックス計算
          const globalIndex = startOffset + batchIndex;
          
          // 基本データ構造作成
          const item = {
            id: `row_${globalIndex + 2}`, // 実際の行番号（ヘッダー考慮）
            timestamp: this.extractFieldValue(row, headers, 'timestamp') || '',
            email: this.extractFieldValue(row, headers, 'email') || '',
            
            // メインコンテンツ
            answer: this.extractFieldValue(row, headers, 'answer', columnMapping) || '',
            reason: this.extractFieldValue(row, headers, 'reason', columnMapping) || '',
            className: this.extractFieldValue(row, headers, 'class', columnMapping) || '',
            name: this.extractFieldValue(row, headers, 'name', columnMapping) || '',

            // メタデータ
            formattedTimestamp: ResponseFormatter.formatTimestamp(this.extractFieldValue(row, headers, 'timestamp')),
            isEmpty: this.isEmptyRow(row),
            
            // リアクション（既存の場合）
            reactions: this.extractReactions(row, headers),
            isHighlighted: this.extractHighlight(row, headers)
          };

          // フィルタリング
          if (this.shouldIncludeRow(item, options)) {
            processedBatch.push(item);
          }
        } catch (rowError) {
          console.warn('DataService.processRawDataBatch: 行処理エラー', {
            batchIndex,
            globalIndex: startOffset + batchIndex,
            error: rowError.message
          });
        }
      });

      return processedBatch;
    } catch (error) {
      console.error('DataService.processRawDataBatch: エラー', error.message);
      return [];
    }
  },

  /**
   * 生データを処理・変換（レガシー互換）
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
            formattedTimestamp: ResponseFormatter.formatTimestamp(this.extractFieldValue(row, headers, 'timestamp')),
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

      // ✅ GAS Best Practice: 直接DB呼び出し（ConfigService依存除去）
      const user = DB.findUserById(userId);
      if (!user || !user.configJson) {
        console.error('DataService.addReaction: ユーザー設定なし');
        return false;
      }

      const config = JSON.parse(user.configJson);
      if (!config.spreadsheetId) {
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
      // ✅ GAS Best Practice: 直接DB呼び出し（ConfigService依存除去）
      const user = DB.findUserById(userId);
      if (!user || !user.configJson) {
        return false;
      }

      const config = JSON.parse(user.configJson);
      if (!config.spreadsheetId) {
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
      // ✅ GAS Best Practice: 直接DB呼び出し（ConfigService依存除去）
      const user = DB.findUserById(userId);
      if (!user || !user.configJson) {
        return false;
      }

      const config = JSON.parse(user.configJson);
      if (!config.spreadsheetId) {
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
      // ✅ GAS Best Practice: 直接DB呼び出し（ConfigService依存除去）
      const user = DB.findUserById(userId);
      if (!user || !user.configJson) {
        throw new Error('ユーザー設定が見つかりません');
      }

      const config = JSON.parse(user.configJson);
      if (!config.spreadsheetId) {
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
      // const rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

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
   * タイムスタンプフォーマット（formatters.gsに統一）
   * @param {string} isoString - ISO形式日時文字列
   * @returns {string} フォーマット済み日時
   */
  // formatTimestamp - formatters.jsに統一 (重複削除完了)

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
  // createSuccessResponse - ResponseFormatterに統一 (重複削除完了)

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
   * リアクション情報抽出
   * @param {Array} row - データ行
   * @param {Array} headers - ヘッダー行
   * @returns {Object} リアクション情報
   */
  extractReactions(row, headers) {
    try {
      const reactions = {
        UNDERSTAND: 0,
        LIKE: 0,
        CURIOUS: 0
      };

      // リアクション列を探して値を抽出
      headers.forEach((header, index) => {
        const headerStr = String(header).toLowerCase();
        if (headerStr.includes('understand') || headerStr.includes('理解')) {
          reactions.UNDERSTAND = parseInt(row[index]) || 0;
        } else if (headerStr.includes('like') || headerStr.includes('いいね')) {
          reactions.LIKE = parseInt(row[index]) || 0;
        } else if (headerStr.includes('curious') || headerStr.includes('気になる')) {
          reactions.CURIOUS = parseInt(row[index]) || 0;
        }
      });

      return reactions;
    } catch (error) {
      console.warn('DataService.extractReactions: エラー', error.message);
      return { UNDERSTAND: 0, LIKE: 0, CURIOUS: 0 };
    }
  },

  /**
   * ハイライト情報抽出
   * @param {Array} row - データ行
   * @param {Array} headers - ヘッダー行
   * @returns {boolean} ハイライト状態
   */
  extractHighlight(row, headers) {
    try {
      // ハイライト列を探して値を抽出
      for (let i = 0; i < headers.length; i++) {
        const header = String(headers[i]).toLowerCase();
        if (header.includes('highlight') || header.includes('ハイライト')) {
          const value = String(row[i]).toUpperCase();
          return value === 'TRUE' || value === '1' || value === 'YES';
        }
      }
      return false;
    } catch (error) {
      console.warn('DataService.extractHighlight: エラー', error.message);
      return false;
    }
  },

  /**
   * 行フィルタリング判定
   * @param {Object} item - データアイテム
   * @param {Object} options - フィルタリングオプション
   * @returns {boolean} 含めるかどうか
   */
  shouldIncludeRow(item, options = {}) {
    try {
      // 空行のフィルタリング
      if (options.excludeEmpty !== false && item.isEmpty) {
        return false;
      }

      // メイン回答が空の行をフィルタリング
      if (options.requireAnswer !== false && (!item.answer || item.answer.trim() === '')) {
        return false;
      }

      // 日付範囲フィルタリング
      if (options.dateFrom && item.timestamp) {
        const itemDate = new Date(item.timestamp);
        const fromDate = new Date(options.dateFrom);
        if (itemDate < fromDate) {
          return false;
        }
      }

      if (options.dateTo && item.timestamp) {
        const itemDate = new Date(item.timestamp);
        const toDate = new Date(options.dateTo);
        if (itemDate > toDate) {
          return false;
        }
      }

      // クラスフィルタリング
      if (options.classFilter && options.classFilter.length > 0) {
        if (!options.classFilter.includes(item.className)) {
          return false;
        }
      }

      // キーワード検索
      if (options.searchKeyword && options.searchKeyword.trim() !== '') {
        const keyword = options.searchKeyword.toLowerCase();
        const searchText = `${item.answer || ''} ${item.reason || ''} ${item.name || ''}`.toLowerCase();
        if (!searchText.includes(keyword)) {
          return false;
        }
      }

      return true;
    } catch (error) {
      console.warn('DataService.shouldIncludeRow: エラー', error.message);
      return true; // エラー時は含める
    }
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
  },

  // ===========================================
  // 📊 管理パネル用API関数（main.gsから移動）
  // ===========================================

  /**
   * スプレッドシート一覧を取得
   * AdminPanel.js.html, AppSetupPage.html から呼び出される
   *
   * @returns {Object} スプレッドシート一覧
   */
  getSpreadsheetList() {
    const started = Date.now();
    try {
      console.log('DataService.getSpreadsheetList: 開始 - GAS独立化完了');

      // ✅ GAS Best Practice: 直接API呼び出し（依存除去）
      const currentUser = Session.getActiveUser().getEmail();
      console.log('DataService.getSpreadsheetList: ユーザー情報', { currentUser });

      // DriveApp直接使用（効率重視）
      const files = DriveApp.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');
      console.log('DataService.getSpreadsheetList: DriveApp検索完了', {
        hasFiles: typeof files !== 'undefined',
        hasNext: files.hasNext()
      });

      // 権限テスト（必要最小限）
      let driveAccessOk = true;
      try {
        const testFiles = DriveApp.getFiles();
        driveAccessOk = testFiles.hasNext();
        console.log('DataService.getSpreadsheetList: Drive権限OK');
      } catch (driveError) {
        console.error('DataService.getSpreadsheetList: Drive権限エラー', driveError.message);
        driveAccessOk = false;
      }

      if (!driveAccessOk) {
        return {
          success: false,
          message: 'Driveアクセス権限がありません',
          spreadsheets: [],
          executionTime: `${Date.now() - started}ms`
        };
      }

      // スプレッドシート取得（高速処理）
      const spreadsheets = [];
      let count = 0;
      const maxCount = 25; // GAS制限対応

      console.log('DataService.getSpreadsheetList: スプレッドシート列挙開始');

      while (files.hasNext() && count < maxCount) {
        try {
          const file = files.next();
          spreadsheets.push({
            id: file.getId(),
            name: file.getName(),
            url: file.getUrl(),
            lastUpdated: file.getLastUpdated()
          });
          count++;
        } catch (fileError) {
          console.warn('DataService.getSpreadsheetList: ファイル処理スキップ', fileError.message);
          continue; // エラー時はスキップして継続
        }
      }

      console.log('DataService.getSpreadsheetList: 列挙完了', {
        totalFound: count,
        maxReached: count >= maxCount
      });

      const response = {
        success: true,
        cached: false,
        executionTime: `${Date.now() - started}ms`,
        spreadsheets,
        totalCount: count,
        maxReached: count >= maxCount
      };

      // レスポンスサイズ監視（GAS制限対応）
      const responseSize = JSON.stringify(response).length;
      const responseSizeKB = Math.round(responseSize / 1024 * 100) / 100;

      if (responseSizeKB > 50) { // 50KB警告
        console.warn('DataService.getSpreadsheetList: 大きなレスポンス', {
          responseSize: `${responseSizeKB  }KB`
        });
      }

      console.log('DataService.getSpreadsheetList: 成功', {
        spreadsheetsCount: response.spreadsheets.length,
        executionTime: response.executionTime
      });

      return response;
    } catch (error) {
      console.error('DataService.getSpreadsheetList: エラー', {
        error: error.message,
        executionTime: `${Date.now() - started}ms`
      });

      return {
        success: false,
        cached: false,
        executionTime: `${Date.now() - started}ms`,
        message: error.message || 'スプレッドシート一覧取得エラー',
        spreadsheets: []
      };
    }
  },

  /**
   * シート一覧を取得
   * AdminPanel.js.html, AppSetupPage.html から呼び出される
   *
   * @param {string} spreadsheetId - スプレッドシートID
   * @returns {Object} シート一覧
   */
  getSheetList(spreadsheetId) {
    try {
      if (!spreadsheetId) {
        return {
          success: false,
          message: 'スプレッドシートIDが指定されていません'
        };
      }

      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheets = spreadsheet.getSheets();
      const sheetList = sheets.map(sheet => ({
        name: sheet.getName(),
        index: sheet.getIndex(),
        rowCount: sheet.getMaxRows(),
        columnCount: sheet.getMaxColumns()
      }));

      return {
        success: true,
        sheets: sheetList
      };
    } catch (error) {
      console.error('DataService.getSheetList エラー:', error.message);
      return {
        success: false,
        message: error.message,
        sheets: []
      };
    }
  },

  /**
   * 列を分析
   * AdminPanel.js.html, AppSetupPage.html から呼び出される
   *
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {Object} 列分析結果
   */
  analyzeColumns(spreadsheetId, sheetName) {
    const started = Date.now();
    try {
      console.log('DataService.analyzeColumns: 開始', {
        spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
        sheetName: sheetName || 'null'
      });

      if (!spreadsheetId || !sheetName) {
        const errorResponse = {
          success: false,
          message: 'スプレッドシートIDとシート名が必要です',
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        };
        console.error('DataService.analyzeColumns: 必須パラメータ不足', {
          spreadsheetId: !!spreadsheetId,
          sheetName: !!sheetName
        });
        return errorResponse;
      }

      // スプレッドシート接続テスト
      let spreadsheet;
      try {
        console.log('DataService.analyzeColumns: スプレッドシート接続開始');
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
        console.log('DataService.analyzeColumns: スプレッドシート接続成功');
      } catch (openError) {
        const errorResponse = {
          success: false,
          message: `スプレッドシートへのアクセスに失敗しました: ${openError.message}`,
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        };
        console.error('DataService.analyzeColumns: スプレッドシート接続失敗', {
          error: openError.message,
          spreadsheetId: `${spreadsheetId.substring(0, 10)}...`
        });
        return errorResponse;
      }

      // シート取得テスト
      let sheet;
      try {
        sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet) {
          const errorResponse = {
            success: false,
            message: `シート "${sheetName}" が見つかりません`,
            headers: [],
            columns: [],
            columnMapping: { mapping: {}, confidence: {} }
          };
          console.error('DataService.analyzeColumns: シートが見つかりません', {
            sheetName
          });
          return errorResponse;
        }
        console.log('DataService.analyzeColumns: シート取得成功');
      } catch (sheetError) {
        const errorResponse = {
          success: false,
          message: `シートアクセスエラー: ${sheetError.message}`,
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        };
        console.error('DataService.analyzeColumns: シートアクセスエラー', {
          error: sheetError.message,
          sheetName
        });
        return errorResponse;
      }

      // ヘッダー行を取得（1行目）- 堅牢化
      let headers = [];
      let sampleData = [];
      let lastColumn = 1;
      let lastRow = 1;

      try {
        console.log('DataService.analyzeColumns: シートサイズ取得開始');
        lastColumn = sheet.getLastColumn();
        lastRow = sheet.getLastRow();
        console.log('DataService.analyzeColumns: シートサイズ', {
          lastColumn,
          lastRow
        });

        if (lastColumn === 0 || lastRow === 0) {
          const errorResponse = {
            success: false,
            message: 'スプレッドシートが空です',
            headers: [],
            columns: [],
            columnMapping: { mapping: {}, confidence: {} }
          };
          console.error('DataService.analyzeColumns: 空のスプレッドシート', {
            lastColumn,
            lastRow
          });
          return errorResponse;
        }

        // ヘッダー行取得
        console.log('DataService.analyzeColumns: ヘッダー取得開始');
        headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
        console.log('DataService.analyzeColumns: ヘッダー取得成功', {
          headersCount: headers.length
        });

        // サンプルデータを取得（最大5行）
        const sampleRowCount = Math.min(5, lastRow - 1);
        if (sampleRowCount > 0) {
          console.log('DataService.analyzeColumns: サンプルデータ取得開始');
          sampleData = sheet.getRange(2, 1, sampleRowCount, lastColumn).getValues();
          console.log('DataService.analyzeColumns: サンプルデータ取得成功', {
            sampleRowCount: sampleData.length
          });
        }
      } catch (rangeError) {
        const errorResponse = {
          success: false,
          message: `データ取得エラー: ${rangeError.message}`,
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        };
        console.error('DataService.analyzeColumns: データ取得エラー', {
          error: rangeError.message,
          lastColumn,
          lastRow
        });
        return errorResponse;
      }

      // 列情報を分析
      const columns = headers.map((header, index) => {
        const samples = sampleData.map(row => row[index]).filter(v => v);

        // 列タイプを推測
        let type = 'text';
        if (header.toLowerCase().includes('timestamp') || header.toLowerCase().includes('日時')) {
          type = 'datetime';
        } else if (header.toLowerCase().includes('email') || header.toLowerCase().includes('メール')) {
          type = 'email';
        } else if (header.toLowerCase().includes('class') || header.toLowerCase().includes('クラス')) {
          type = 'class';
        } else if (header.toLowerCase().includes('name') || header.toLowerCase().includes('名前')) {
          type = 'name';
        } else if (samples.length > 0 && samples.every(s => !isNaN(s))) {
          type = 'number';
        }

        return {
          index,
          header,
          type,
          samples: samples.slice(0, 3) // 最大3つのサンプル
        };
      });

      // AI検出シミュレーション（シンプルなパターンマッチング）
      const mapping = { mapping: {}, confidence: {} };

      headers.forEach((header, index) => {
        const headerLower = header.toLowerCase();

        // 回答列の検出
        if (headerLower.includes('回答') || headerLower.includes('answer') || headerLower.includes('意見')) {
          mapping.mapping.answer = index;
          mapping.confidence.answer = 85;
        }

        // 理由列の検出
        if (headerLower.includes('理由') || headerLower.includes('根拠') || headerLower.includes('reason')) {
          mapping.mapping.reason = index;
          mapping.confidence.reason = 80;
        }

        // クラス列の検出
        if (headerLower.includes('クラス') || headerLower.includes('class') || headerLower.includes('組')) {
          mapping.mapping.class = index;
          mapping.confidence.class = 90;
        }

        // 名前列の検出
        if (headerLower.includes('名前') || headerLower.includes('name') || headerLower.includes('氏名')) {
          mapping.mapping.name = index;
          mapping.confidence.name = 85;
        }
      });

      const result = {
        success: true,
        headers,
        columns,
        columnMapping: mapping,
        sampleData: sampleData.slice(0, 3), // 最大3行のサンプル
        executionTime: `${Date.now() - started}ms`
      };

      console.log('DataService.analyzeColumns: 成功', {
        headersCount: headers.length,
        columnsCount: columns.length,
        mappingDetected: Object.keys(mapping.mapping).length,
        executionTime: result.executionTime
      });

      return result;

    } catch (error) {
      const errorResponse = {
        success: false,
        message: `予期しないエラーが発生しました: ${error.message}`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} },
        executionTime: `${Date.now() - started}ms`
      };

      console.error('DataService.analyzeColumns: 予期しないエラー', {
        error: error.message,
        stack: error.stack,
        executionTime: errorResponse.executionTime
      });

      console.error('DataService.analyzeColumns 予期しないエラー:', {
        error: error.message,
        stack: error.stack,
        spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)  }...` : 'null',
        sheetName
      });

      // 絶対にnullを返さない
      return errorResponse;
    }
  },

  /**
   * 軽量ヘッダー取得 - 列分析に失敗してもヘッダー名だけは取得
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @returns {Object} ヘッダー取得結果
   */
  getLightweightHeaders(spreadsheetId, sheetName) {
    const started = Date.now();
    try {
      console.log('DataService.getLightweightHeaders: 開始', {
        spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
        sheetName: sheetName || 'null'
      });

      if (!spreadsheetId || !sheetName) {
        return {
          success: false,
          message: 'スプレッドシートIDとシート名が必要です',
          headers: []
        };
      }

      // 最小限のアクセスでヘッダーのみ取得
      const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      const sheet = spreadsheet.getSheetByName(sheetName);

      if (!sheet) {
        return {
          success: false,
          message: `シート "${sheetName}" が見つかりません`,
          headers: []
        };
      }

      const lastColumn = sheet.getLastColumn();
      if (lastColumn === 0) {
        return {
          success: false,
          message: 'スプレッドシートが空です',
          headers: []
        };
      }

      // ヘッダー行のみ取得（データは取得しない）
      const headers = sheet.getRange(1, 1, 1, Math.min(lastColumn, 20)).getValues()[0];

      const result = {
        success: true,
        headers: headers.map(h => h || ''),
        executionTime: `${Date.now() - started}ms`
      };

      console.log('DataService.getLightweightHeaders: 成功', {
        headersCount: result.headers.length,
        executionTime: result.executionTime
      });

      return result;

    } catch (error) {
      const errorResponse = {
        success: false,
        message: `ヘッダー取得エラー: ${error.message}`,
        headers: [],
        executionTime: `${Date.now() - started}ms`
      };

      console.error('DataService.getLightweightHeaders: エラー', {
        error: error.message,
        executionTime: errorResponse.executionTime
      });

      return errorResponse;
    }
  }

});
