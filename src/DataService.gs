/**
 * @fileoverview DataService - 統一データ操作サービス (遅延初期化対応)
 *
 * 🎯 責任範囲:
 * - スプレッドシートデータ取得・操作
 * - リアクション・ハイライト機能
 * - データフィルタリング・検索
 * - バルクデータAPI
 *
 * 🔄 GAS Best Practices準拠:
 * - 遅延初期化パターン (各公開関数先頭でinit)
 * - ファイル読み込み順序非依存設計
 * - グローバル副作用排除
 */

/* global ServiceFactory, formatTimestamp, DatabaseOperations, createErrorResponse, createExceptionResponse, getSheetData, columnAnalysis */

// ===========================================
// 🔧 Zero-Dependency DataService (ServiceFactory版)
// ===========================================

/**
 * DataService - ゼロ依存アーキテクチャ
 * ServiceFactoryパターンによる依存関係除去
 * DB, CONSTANTS依存を完全排除
 */

/**
 * ServiceFactory統合初期化（DataService版）
 * 依存関係チェックなしの即座初期化
 * @returns {boolean} 初期化成功可否
 */
function initDataServiceZero() {
  return ServiceFactory.getUtils().initService('DataService');
}

/**
 * ユーザーのスプレッドシートデータ取得（統合版）
 * GAS公式ベストプラクティス：シンプルな関数形式
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} GAS公式推奨レスポンス形式
 */
function getUserSheetData(userId, options = {}) {
  const startTime = Date.now();

  try {
    // 🚀 Zero-dependency initialization
    if (!initDataServiceZero()) {
      console.error('getSheetData: ServiceFactory not available');
      return createErrorResponse('サービス初期化エラー', { data: [], headers: [], sheetName: '' });
    }

    // 🔧 ServiceFactory経由でデータベース取得
    const db = ServiceFactory.getDB();
    if (!db) {
      console.error('DataService.getUserSheetData: Database not available');
      return { success: false, message: 'データベース接続エラー', data: [], headers: [], sheetName: '' };
    }

    const user = db.findUserById(userId);
    if (!user || !user.configJson) {
      console.error('DataService.getUserSheetData: ユーザー設定が見つかりません', { userId });
      // Direct return format like admin panel getSheetList
      return { success: false, message: 'ユーザー設定を取得できませんでした', data: [], headers: [], sheetName: '' };
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId) {
      console.warn('DataService.getUserSheetData: スプレッドシートIDが設定されていません', { userId });
      // Direct return format like admin panel getSheetList
      return { success: false, message: 'スプレッドシートが設定されていません', data: [], headers: [], sheetName: '' };
    }

    // データ取得実行
    const result = fetchSpreadsheetData(config, options);

    const executionTime = Date.now() - startTime;
    console.info('getSheetData: データ取得完了', {
      userId,
      rowCount: result.data?.length || 0,
      executionTime
    });

    // Standardized response format
    if (result.success) {
      return {
        ...result,
        header: config.header || config.title || result.sheetName || '回答一覧',
        showDetails: config.showDetails !== false // デフォルトはtrue
      };
    }

    return result;
  } catch (error) {
    console.error('DataService.getUserSheetData: エラー', {
      userId,
      error: error.message
    });
    // Direct return format like admin panel getSheetList
    return { success: false, message: error.message || 'データ取得エラー', data: [], headers: [], sheetName: '' };
  }
}


/**
 * スプレッドシートデータ取得実行
 * ✅ バッチ処理対応 - GAS制限対応（実行時間・メモリ制限）
 * GAS公式ベストプラクティス：シンプルな関数形式
 * @param {Object} config - ユーザー設定
 * @param {Object} options - 取得オプション
 * @returns {Object} GAS公式推奨レスポンス形式
 */
function fetchSpreadsheetData(config, options = {}) {
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
      // ✅ google.script.run 互換: シンプル形式
      return { success: true, data: [], headers: [], sheetName: config.sheetName || '不明' };
    }

    // ヘッダー行取得
    const [headers] = sheet.getRange(1, 1, 1, lastCol).getValues();

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
        const batchProcessed = processRawDataBatch(batchRows, headers, config, options, startRow - 2);

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
    // クラスフィルタリングとソートを適用
    if (options.classFilter) {
      processedData = processedData.filter(item => item.class === options.classFilter);
    }

    // ソート処理
    if (options.sortBy) {
      processedData = applySortAndLimit(processedData, {
        sortBy: options.sortBy,
        limit: options.limit
      });
    }

    console.info('DataService.fetchSpreadsheetData: バッチ処理完了', {
      totalRows: totalDataRows,
      processedRows: processedCount,
      filteredRows: processedData.length,
      executionTime,
      batchCount: Math.ceil(totalDataRows / MAX_BATCH_SIZE)
    });

    // ✅ google.script.run 互換: フロントエンド期待形式で直接返却
    return {
      success: true,
      data: processedData,
      headers,
      sheetName: config.sheetName || '不明',
      // デバッグ情報（オプショナル）
      totalRows: totalDataRows,
      processedRows: processedCount,
      filteredRows: processedData.length,
      executionTime
    };
  } catch (error) {
    console.error('DataService.fetchSpreadsheetData: エラー', error.message);
    throw error;
  }
}

/**
 * バッチ処理用データ変換（メモリ効率重視）
 * @param {Array} batchRows - バッチデータ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} config - 設定
 * @param {Object} options - 処理オプション
 * @param {number} startOffset - 開始オフセット（行番号計算用）
 * @returns {Array} 処理済みバッチデータ
 */
function processRawDataBatch(batchRows, headers, config, options = {}, startOffset = 0) {
  try {
    const columnMapping = config.columnMapping?.mapping || {};
    const processedBatch = [];

    batchRows.forEach((row, batchIndex) => {
      try {
        // グローバル行インデックス計算
        const globalIndex = startOffset + batchIndex;

        // 基本データ構造作成
        const item = {
          id: `row_${globalIndex + 2}`,
          rowIndex: globalIndex + 2, // 1-based row number including header
          timestamp: extractFieldValue(row, headers, 'timestamp') || '',
          email: extractFieldValue(row, headers, 'email') || '',

          // メインコンテンツ（フロントエンドと統一）
          answer: extractFieldValue(row, headers, 'answer', columnMapping) || '',
          opinion: extractFieldValue(row, headers, 'answer', columnMapping) || '', // Alias for answer field
          reason: extractFieldValue(row, headers, 'reason', columnMapping) || '',
          class: extractFieldValue(row, headers, 'class', columnMapping) || '',
          name: extractFieldValue(row, headers, 'name', columnMapping) || '',

          // メタデータ
          formattedTimestamp: formatTimestamp(extractFieldValue(row, headers, 'timestamp')),
          isEmpty: isEmptyRow(row),

          // リアクション（{count, reacted} 形式に統一）
          reactions: extractReactions(row, headers),
          highlight: extractHighlight(row, headers)
        };

        // フィルタリング
        if (shouldIncludeRow(item, options)) {
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
}

/**
 * 生データを処理・変換（レガシー互換）
 * @param {Array} dataRows - 生データ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} config - 設定
 * @param {Object} options - 処理オプション
 * @returns {Array} 処理済みデータ
 */
function processRawData(dataRows, headers, config, options = {}) {
  try {
    const columnMapping = config.columnMapping?.mapping || {};
    const processedData = [];

    dataRows.forEach((row, index) => {
      try {
        // 基本データ構造作成
        const item = {
          id: `row_${index + 2}`,
          rowIndex: index + 2,
          timestamp: extractFieldValue(row, headers, 'timestamp') || '',
          email: extractFieldValue(row, headers, 'email') || '',

          // メインコンテンツ
          answer: extractFieldValue(row, headers, 'answer', columnMapping) || '',
          opinion: extractFieldValue(row, headers, 'answer', columnMapping) || '', // Alias for answer field
          reason: extractFieldValue(row, headers, 'reason', columnMapping) || '',
          class: extractFieldValue(row, headers, 'class', columnMapping) || '',
          name: extractFieldValue(row, headers, 'name', columnMapping) || '',

          // メタデータ
          formattedTimestamp: formatTimestamp(extractFieldValue(row, headers, 'timestamp')),
          isEmpty: isEmptyRow(row),

          // リアクション（{count, reacted} 形式）
          reactions: extractReactions(row, headers),
          highlight: extractHighlight(row, headers)
        };

        // フィルタリング
        if (shouldIncludeRow(item, options)) {
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
    return applySortAndLimit(processedData, options);
  } catch (error) {
    console.error('DataService.processRawData: エラー', error.message);
    return [];
  }
}

/**
 * フィールド値抽出（列マッピング対応）
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @returns {*} フィールド値
 */
function extractFieldValue(row, headers, fieldType, columnMapping = {}) {
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
}

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



/**
 * スプレッドシート内リアクション更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @param {string} reactionType - リアクションタイプ
 * @param {string} action - アクション（add/remove）
 * @returns {boolean} 成功可否
 */
function updateReactionInSheet(config, rowId, reactionType, action) {
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
    const reactionColumn = getOrCreateReactionColumn(sheet, reactionType);
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
}

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
function processDataWithColumnMapping(dataRows, headers, columnMapping) {
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
}

/**
 * 自動停止時間計算（Core.gsより移行）
 * @param {string} publishedAt - 公開開始時間のISO文字列
 * @param {number} minutes - 自動停止までの分数
 * @returns {Object} 停止時間情報
 */
function getAutoStopTime(publishedAt, minutes) {
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
}

/**
 * レガシーCore.gsから移行: リアクション処理実行
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類
 * @param {string} userEmail - ユーザーメール
 * @returns {Object} 処理結果
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, _userEmail) {
  // 🚀 Zero-dependency: ServiceFactory経由で初期化
  try {
    if (!validateReaction(spreadsheetId, sheetName, rowIndex, reactionKey)) {
      throw new Error('無効なリアクションパラメータ');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // リアクション列を取得または作成
    const reactionColumn = getOrCreateReactionColumn(sheet, reactionKey);

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
      success: true,
      status: 'success',
      message: 'リアクションを追加しました',
      newValue
    };
  } catch (error) {
    console.error('DataService.processReaction: エラー', error.message);
    return {
      success: false,
      status: 'error',
      message: error.message
    };
  }
}

// ===========================================
// 🔧 ユーティリティ・ヘルパー
// ===========================================

/**
 * 空行判定
 * @param {Array} row - データ行
 * @returns {boolean} 空行かどうか
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
}

/**
 * リアクションタイプ検証
 * @param {string} reactionType - リアクションタイプ
 * @returns {boolean} 有効かどうか
 */
function validateReactionType(reactionType) {
  // 🔧 CONSTANTS依存除去: 直接定義
  const validTypes = ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
  return validTypes.includes(reactionType);
}

/**
 * リアクション情報抽出
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行
 * @returns {Object} リアクション情報
 */
function extractReactions(row, headers) {
  try {
    const counts = {
      UNDERSTAND: 0,
      LIKE: 0,
      CURIOUS: 0
    };

    // リアクション列を探して値を抽出
    headers.forEach((header, index) => {
      const headerStr = String(header).toLowerCase();
      if (headerStr.includes('understand') || headerStr.includes('理解')) {
        counts.UNDERSTAND = parseInt(row[index]) || 0;
      } else if (headerStr.includes('like') || headerStr.includes('いいね')) {
        counts.LIKE = parseInt(row[index]) || 0;
      } else if (headerStr.includes('curious') || headerStr.includes('気になる')) {
        counts.CURIOUS = parseInt(row[index]) || 0;
      }
    });

    return {
      UNDERSTAND: { count: counts.UNDERSTAND, reacted: false },
      LIKE: { count: counts.LIKE, reacted: false },
      CURIOUS: { count: counts.CURIOUS, reacted: false }
    };
  } catch (error) {
    console.warn('DataService.extractReactions: エラー', error.message);
    return {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };
  }
}

/**
 * ハイライト情報抽出
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー行
 * @returns {boolean} ハイライト状態
 */
function extractHighlight(row, headers) {
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
}

/**
 * 行フィルタリング判定
 * @param {Object} item - データアイテム
 * @param {Object} options - フィルタリングオプション
 * @returns {boolean} 含めるかどうか
 */
function shouldIncludeRow(item, options = {}) {
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
      if (!options.classFilter.includes(item.class)) {
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
}

/**
 * ソート・制限適用
 * @param {Array} data - データ配列
 * @param {Object} options - オプション
 * @returns {Array} ソート済みデータ
 */
function applySortAndLimit(data, options = {}) {
  try {
    let sortedData = [...data];

    // ソート適用
    if (options.sortBy) {
      switch (options.sortBy) {
        case 'newest':
          sortedData.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
          break;
        case 'oldest':
          sortedData.sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));
          break;
        case 'reactions':
          sortedData.sort((a, b) => {
            const aTotal = (a.reactions?.UNDERSTAND || 0) + (a.reactions?.LIKE || 0) + (a.reactions?.CURIOUS || 0);
            const bTotal = (b.reactions?.UNDERSTAND || 0) + (b.reactions?.LIKE || 0) + (b.reactions?.CURIOUS || 0);
            return bTotal - aTotal;
          });
          break;
        case 'random':
          // Fisher-Yates シャッフル
          for (let i = sortedData.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [sortedData[i], sortedData[j]] = [sortedData[j], sortedData[i]];
          }
          break;
      }
    }

    // 制限適用
    if (options.limit && options.limit > 0) {
      sortedData = sortedData.slice(0, options.limit);
    }

    return sortedData;
  } catch (error) {
    console.error('DataService.applySortAndLimit: エラー', error.message);
    return data;
  }
}

// ===========================================
// 📊 管理パネル用API関数（main.gsから移動）
// ===========================================

/**
 * スプレッドシート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @returns {Object} スプレッドシート一覧
 */
function getSpreadsheetList() {
  // 🚀 Zero-dependency: ServiceFactory経由で初期化
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

    // ✅ google.script.run互換 - シンプル形式に最適化
    const response = {
      success: true,
      spreadsheets,
      executionTime: `${Date.now() - started}ms`
    };

    // レスポンスサイズ監視（GAS制限対応）
    const responseSize = JSON.stringify(response).length;
    const responseSizeKB = Math.round(responseSize / 1024 * 100) / 100;

    console.log('DataService.getSpreadsheetList: 成功 - google.script.run最適化', {
      spreadsheetsCount: spreadsheets.length,
      executionTime: response.executionTime,
      responseSizeKB,
      responseValid: response !== null && typeof response === 'object'
    });

    // ✅ google.script.run互換性チェック
    if (!response || typeof response !== 'object' || !Array.isArray(response.spreadsheets)) {
      console.error('DataService.getSpreadsheetList: 無効なレスポンス形式', response);
      return {
        success: false,
        spreadsheets: [],
        message: 'レスポンス形式エラー'
      };
    }

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
}

/**
 * シート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */

/**
 * 列を分析
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */

// ===========================================
// 📊 Column Analysis - Refactored Functions
// ===========================================

/**
 * 列分析のメイン関数（リファクタリング版）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */

// ===========================================
// 📊 Column Analysis - Refactored Functions
// ===========================================

/**
 * 列分析のメイン関数（コンテキスト対応版）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Object} options - 分析オプション
 * @param {boolean} options.basicOnly - 基本ヘッダー情報のみ取得
 * @param {boolean} options.useConfigJson - configJsonからマッピング復元
 * @param {string} options.userId - ユーザーID（設定復元用）
 * @param {boolean} options.forceFullAnalysis - フル分析を強制実行（設定復元・基本ヘッダーをスキップ）
 * @returns {Object} 列分析結果
 */
function columnAnalysisImpl(spreadsheetId, sheetName, options = {}) {
  const started = Date.now();
  try {
    console.log('DataService.columnAnalysis: 開始', {
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
      sheetName: sheetName || 'null',
      options
    });

    // 🎯 GAS Best Practice: パラメータ検証を別関数に分離
    const paramValidation = validateSheetParams(spreadsheetId, sheetName);
    if (!paramValidation.isValid) {
      console.error('DataService.columnAnalysis: パラメータ検証失敗');
      return paramValidation.errorResponse;
    }

    // 🎯 フル分析強制実行の場合は設定復元・基本ヘッダーをスキップ
    if (!options.forceFullAnalysis) {
      // 🎯 configJsonからの設定復元（優先実行）
      if (options.useConfigJson && options.userId) {
        const configResult = restoreColumnConfig(options.userId, spreadsheetId, sheetName);
        if (configResult.success) {
          console.log('DataService.columnAnalysis: configJson復元成功');
          return configResult;
        }
      }

      // 🎯 基本ヘッダー情報のみ取得
      if (options.basicOnly) {
        return getSheetHeaders(spreadsheetId, sheetName, started);
      }
    } else {
      console.log('DataService.columnAnalysis: フル分析を強制実行');
    }

    // 🎯 GAS Best Practice: スプレッドシート接続を別関数に分離
    console.log('DataService.columnAnalysis: スプレッドシート接続開始');
    const connectionResult = connectToSheetInternal(spreadsheetId, sheetName);
    console.log('DataService.columnAnalysis: 接続結果', { success: connectionResult?.success });
    if (!connectionResult.success) {
      console.error('DataService.columnAnalysis: スプレッドシート接続失敗');
      return connectionResult.errorResponse;
    }

    // 🎯 GAS Best Practice: データ取得を別関数に分離
    console.log('DataService.columnAnalysis: データ取得開始');
    const dataResult = extractSheetHeaders(connectionResult.sheet);
    console.log('DataService.columnAnalysis: データ取得結果', { success: dataResult?.success, headerCount: dataResult?.headers?.length });
    if (!dataResult.success) {
      console.error('DataService.columnAnalysis: データ取得失敗');
      return dataResult.errorResponse;
    }

    // 🎯 GAS Best Practice: 列分析を別関数に分離
    console.log('DataService.columnAnalysis: 列分析開始');
    const analysisResult = detectColumnTypes(dataResult.headers, dataResult.sampleData);
    console.log('DataService.columnAnalysis: 列分析結果', {
      mappingKeys: Object.keys(analysisResult?.mapping?.mapping || {}),
      confidenceKeys: Object.keys(analysisResult?.mapping?.confidence || {})
    });

    console.log('DataService.columnAnalysis: 最終結果構築開始');

    // google.script.run互換性: フロントエンド期待フォーマットに変換
    const frontendMapping = {
      ...analysisResult.mapping.mapping,        // answer: 4, reason: 5, etc
      confidence: analysisResult.mapping.confidence  // confidence: { answer: 90, ... }
    };

    const finalResult = {
      success: true,
      headers: dataResult.headers,
      columnMapping: frontendMapping,
      executionTime: `${Date.now() - started}ms`
    };

    console.log('DataService.columnAnalysis: 正常終了', {
      headersCount: dataResult.headers.length,
      mappingKeys: Object.keys(analysisResult.mapping?.mapping || {}),
      success: true,
      finalResultType: typeof finalResult,
      finalResultKeys: Object.keys(finalResult)
    });

    console.log('DataService.columnAnalysis: 戻り値準備完了', finalResult);
    return finalResult;

  } catch (error) {
    console.error('DataService.columnAnalysis: 予期しないエラー', {
      error: error.message,
      stack: error.stack,
      executionTime: `${Date.now() - started}ms`
    });

    return {
      success: false,
      message: `予期しないエラーが発生しました: ${error.message}`,
      headers: [],
      columns: [],
      columnMapping: { mapping: {}, confidence: {} },
      executionTime: `${Date.now() - started}ms`
    };
  }
}

/**
 * パラメータ検証（GAS Best Practice: 単一責任）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 検証結果
 */
function validateSheetParams(spreadsheetId, sheetName) {
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
    return { isValid: false, errorResponse };
  }

  return { isValid: true };
}

/**
 * スプレッドシート接続（GAS Best Practice: エラーハンドリング分離）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 接続結果
 */
function connectToSheetInternal(spreadsheetId, sheetName) {
  try {
    console.log('DataService.connectToSheetInternal: スプレッドシート接続開始');
    console.log('DataService.connectToSheetInternal: パラメータ確認', {
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
      sheetName: sheetName || 'null'
    });

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    console.log('DataService.connectToSheetInternal: スプレッドシート接続成功');

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: `シート "${sheetName}" が見つかりません`,
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        }
      };
    }

    console.log('DataService.connectToSheetInternal: シート取得成功');
    return { success: true, sheet };

  } catch (error) {
    console.error('DataService.connectToSheetInternal: 接続エラー', {
      error: error.message,
      spreadsheetId: `${spreadsheetId.substring(0, 10)}...`
    });

    return {
      success: false,
      errorResponse: {
        success: false,
        message: `スプレッドシートアクセスエラー: ${error.message}`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      }
    };
  }
}

/**
 * シートデータ抽出（GAS Best Practice: データアクセス最適化）
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} データ抽出結果
 */
function extractSheetHeaders(sheet) {
  try {
    console.log('DataService.extractSheetHeaders: シートサイズ取得開始');
    const lastColumn = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    console.log('DataService.extractSheetHeaders: シートサイズ', { lastColumn, lastRow });

    if (lastColumn === 0 || lastRow === 0) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: 'スプレッドシートが空です',
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        }
      };
    }

    // 🎯 GAS Best Practice: バッチでデータ取得
    console.log('DataService.analyzeColumns: ヘッダー取得開始');
    const [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues();
    console.log('DataService.analyzeColumns: ヘッダー取得成功', { headersCount: headers.length });

    // サンプルデータを取得（最大5行）
    let sampleData = [];
    const sampleRowCount = Math.min(5, lastRow - 1);
    if (sampleRowCount > 0) {
      console.log('DataService.extractSheetHeaders: サンプルデータ取得開始');
      sampleData = sheet.getRange(2, 1, sampleRowCount, lastColumn).getValues();
      console.log('DataService.extractSheetHeaders: サンプルデータ取得成功', {
        sampleRowCount: sampleData.length
      });
    }

    return { success: true, headers, sampleData };

  } catch (error) {
    console.error('DataService.extractSheetHeaders: データ取得エラー', {
      error: error.message
    });

    return {
      success: false,
      errorResponse: {
        success: false,
        message: `データ取得エラー: ${error.message}`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      }
    };
  }
}

/**
 * configJsonベースの列マッピング取得
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 設定ベースの結果
 */
function restoreColumnConfig(userId, spreadsheetId, sheetName) {
  try {
    const user = DatabaseOperations.findUserByEmail(Session.getActiveUser().getEmail());
    if (!user || !user.configJson) {
      return { success: false, message: 'User config not found' };
    }

    const config = JSON.parse(user.configJson);
    if (config.spreadsheetId !== spreadsheetId || config.sheetName !== sheetName) {
      return { success: false, message: 'Config mismatch' };
    }

    // 基本ヘッダー情報を取得
    const basicHeaders = getSheetHeaders(spreadsheetId, sheetName, Date.now());
    if (!basicHeaders.success) {
      return basicHeaders;
    }

    return {
      success: true,
      headers: basicHeaders.headers,
      columnMapping: {
        mapping: config.columnMapping || {},
        confidence: config.confidence || {}
      },
      source: 'configJson',
      executionTime: basicHeaders.executionTime
    };
  } catch (error) {
    console.error('restoreColumnConfig error:', error.message);
    return { success: false, message: error.message };
  }
}

/**
 * 基本ヘッダー情報のみ取得（軽量版の代替）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} started - 開始時刻
 * @returns {Object} ヘッダー情報
 */
function getSheetHeaders(spreadsheetId, sheetName, started) {
  try {
    console.log('getSheetHeaders: ServiceFactory.getSpreadsheet()呼び出し開始');
    const spreadsheetService = ServiceFactory.getSpreadsheet();
    console.log('getSheetHeaders: spreadsheetService取得結果', { isNull: spreadsheetService === null });

    if (!spreadsheetService) {
      console.error('getSheetHeaders: ServiceFactory.getSpreadsheet()がnullを返しました');
      return {
        success: false,
        message: 'ServiceFactory.getSpreadsheet()がnullを返しました',
        headers: []
      };
    }

    const spreadsheet = spreadsheetService.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return {
        success: false,
        message: `シート "${sheetName}" が見つかりません`,
        headers: []
      };
    }

    const lastColumn = sheet.getLastColumn();
    const headers = lastColumn > 0 ?
      sheet.getRange(1, 1, 1, lastColumn).getValues()[0] : [];

    return {
      success: true,
      headers: headers.map(h => String(h || '')),
      sheetName,
      columnCount: lastColumn,
      source: 'basic',
      executionTime: `${Date.now() - started}ms`
    };
  } catch (error) {
    console.error('getSheetHeaders error:', error.message);
    return {
      success: false,
      message: error.message || 'ヘッダー取得エラー',
      headers: []
    };
  }
}

/**
 * 列分析実行（GAS Best Practice: ビジネスロジック分離）
 * @param {Array} headers - ヘッダー配列
 * @param {Array} sampleData - サンプルデータ配列
 * @returns {Object} 分析結果
 */
function detectColumnTypes(headers, sampleData) {
  try {
    console.log('DataService.detectColumnTypes: 高精度AI分析開始', {
      headersCount: headers.length,
      sampleDataCount: sampleData.length
    });

    // 防御的プログラミング: 入力値検証
    if (!Array.isArray(headers) || headers.length === 0) {
      console.warn('DataService.detectColumnTypes: 無効なheaders', headers);
      return { columns: [], mapping: { mapping: {}, confidence: {} } };
    }

    if (!Array.isArray(sampleData)) {
      console.warn('DataService.detectColumnTypes: 無効なsampleData', sampleData);
      sampleData = [];
    }

    // 列情報を分析
    const columns = headers.map((header, index) => {
      const samples = sampleData.map(row => row && row[index]).filter(v => v != null && v !== '');

      // 列タイプを推測
      let type = 'text';
      const headerLower = String(header || '').toLowerCase();

      if (headerLower.includes('timestamp') || headerLower.includes('日時') || headerLower.includes('タイムスタンプ')) {
        type = 'datetime';
      } else if (headerLower.includes('email') || headerLower.includes('メール')) {
        type = 'email';
      } else if (headerLower.includes('class') || headerLower.includes('クラス')) {
        type = 'class';
      } else if (headerLower.includes('name') || headerLower.includes('名前')) {
        type = 'name';
      } else if (samples.length > 0 && samples.every(s => !isNaN(s))) {
        type = 'number';
      }

      return {
        index,
        header: String(header || ''),
        type,
        samples: samples.slice(0, 3) // 最大3つのサンプル
      };
    });

    // 🎯 高精度AI検出システム（5次元統計分析）
    const mapping = { mapping: {}, confidence: {} };
    const analysisResults = performHighPrecisionAnalysis(headers, sampleData);

    // 結果をマッピングに反映 - 設定可能な信頼度閾値
    const confidenceThreshold = 50; // Zero-dependency: 日本語ヘッダー対応で50%に調整
    console.log(`DataService.detectColumnTypes: 信頼度閾値=${confidenceThreshold}% で検出開始`);

    Object.entries(analysisResults).forEach(([columnType, result]) => {
      console.log(`DataService.detectColumnTypes: ${columnType}列分析結果 - index=${result.index}, confidence=${result.confidence}%, header="${headers[result.index] || 'undefined'}"`);

      if (result.confidence >= confidenceThreshold) {
        mapping.mapping[columnType] = result.index;
        mapping.confidence[columnType] = Math.round(result.confidence);
        console.log(`✅ DataService.detectColumnTypes: 検出採用 - ${columnType}列 "${headers[result.index]}" at index ${result.index} (confidence: ${result.confidence}%)`);
      } else {
        // 閾値未満でも信頼度情報は記録（フロントエンド用）
        mapping.confidence[columnType] = Math.round(result.confidence);
        console.log(`⚠️ DataService.detectColumnTypes: 信頼度不足 - ${columnType}列 "${headers[result.index] || 'undefined'}" (${result.confidence}% < ${confidenceThreshold}%)`);
      }
    });

    const result = { columns, mapping };

    console.log('DataService.detectColumnTypes: 高精度分析完了', {
      headersCount: headers.length,
      columnsCount: columns.length,
      mappingDetected: Object.keys(mapping.mapping).length,
      mappingDetails: mapping.mapping,
      confidenceDetails: mapping.confidence,
      analysisMethod: '5次元統計分析'
    });

    return result;

  } catch (error) {
    console.error('DataService.detectColumnTypes: エラー', {
      error: error.message,
      stack: error.stack
    });

    // エラー時のフォールバック
    return {
      columns: [],
      mapping: { mapping: {}, confidence: {} }
    };
  }
}

/**
 * 🎯 高精度AI検出システム - 5次元統計分析
 * @param {Array} headers - 列ヘッダー
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} 分析結果
 */
function performHighPrecisionAnalysis(headers, sampleData) {
  console.log('🔍 performHighPrecisionAnalysis: 開始', {
    headersCount: headers.length,
    sampleDataCount: sampleData.length,
    headers
  });

  const results = {
    answer: { index: -1, confidence: 0, factors: {} },
    reason: { index: -1, confidence: 0, factors: {} },
    class: { index: -1, confidence: 0, factors: {} },
    name: { index: -1, confidence: 0, factors: {} }
  };

  headers.forEach((header, index) => {
    if (!header) return;

    console.log(`🔍 分析中: ヘッダー[${index}] = "${header}"`);

    const samples = sampleData.map(row => row && row[index]).filter(v => v != null && v !== '');

    // 各列タイプに対する分析を実行
    Object.keys(results).forEach(columnType => {
      const analysis = analyzeColumnForType(header, samples, index, headers, columnType);
      console.log(`📊 ${columnType}分析結果[${index}]: 信頼度=${analysis.confidence}%, factors=`, analysis.factors);

      if (analysis.confidence > results[columnType].confidence) {
        results[columnType] = analysis;
        console.log(`✅ ${columnType}の最優秀候補更新: ヘッダー[${index}]="${header}" (信頼度: ${analysis.confidence}%)`);
      }
    });
  });

  console.log('🎯 performHighPrecisionAnalysis: 最終結果', {
    answer: `index=${results.answer.index}, confidence=${results.answer.confidence}%`,
    reason: `index=${results.reason.index}, confidence=${results.reason.confidence}%`,
    class: `index=${results.class.index}, confidence=${results.class.confidence}%`,
    name: `index=${results.name.index}, confidence=${results.name.confidence}%`
  });

  return results;
}

/**
 * 特定の列タイプに対する多次元分析
 * @param {string} header - 列ヘッダー
 * @param {Array} samples - サンプルデータ
 * @param {number} index - 列インデックス
 * @param {Array} allHeaders - 全ヘッダー
 * @param {string} targetType - 対象列タイプ
 * @returns {Object} 分析結果
 */
function analyzeColumnForType(header, samples, index, allHeaders, targetType) {
  const headerLower = String(header).toLowerCase();
  let totalConfidence = 0;
  const factors = {};

  // 1️⃣ ヘッダーパターン分析
  const headerScore = analyzeHeaderPattern(headerLower, targetType);
  factors.headerPattern = headerScore;

  // 🎯 日本語完全一致対応: 高精度パターンマッチの重み配分を動的調整
  let headerWeight, contentWeight, linguisticWeight, contextWeight, semanticWeight;

  if (headerScore >= 90) {
    // 日本語完全一致 (95%等) - ヘッダーパターンを最重要視
    headerWeight = 0.6;      // 30% → 60% (倍増)
    contentWeight = 0.15;    // 25% → 15%
    linguisticWeight = 0.1;  // 20% → 10%
    contextWeight = 0.1;     // 15% → 10%
    semanticWeight = 0.05;   // 10% → 5%
  } else if (headerScore >= 70) {
    // 強パターンマッチ - ヘッダー重視
    headerWeight = 0.45;
    contentWeight = 0.2;
    linguisticWeight = 0.15;
    contextWeight = 0.125;
    semanticWeight = 0.075;
  } else {
    // 標準分析 - バランス型
    headerWeight = 0.3;
    contentWeight = 0.25;
    linguisticWeight = 0.2;
    contextWeight = 0.15;
    semanticWeight = 0.1;
  }

  totalConfidence += headerScore * headerWeight;

  // 2️⃣ コンテンツ統計分析
  const contentScore = analyzeContentStatistics(samples, targetType);
  factors.contentStatistics = contentScore;
  totalConfidence += contentScore * contentWeight;

  // 3️⃣ 言語パターン分析
  const linguisticScore = analyzeLinguisticPatterns(samples, targetType);
  factors.linguisticPatterns = linguisticScore;
  totalConfidence += linguisticScore * linguisticWeight;

  // 4️⃣ コンテキスト推論
  const contextScore = analyzeContextualClues(header, index, allHeaders, targetType);
  factors.contextualClues = contextScore;
  totalConfidence += contextScore * contextWeight;

  // 5️⃣ セマンティック分析
  const semanticScore = analyzeSemanticCharacteristics(samples, targetType);
  factors.semanticCharacteristics = semanticScore;
  totalConfidence += semanticScore * semanticWeight;

  const finalConfidence = Math.min(Math.max(totalConfidence, 0), 100);

  // 詳細デバッグログ（動的重み対応版）
  if (headerScore > 0 || finalConfidence > 10) {
    const weightProfile = headerScore >= 90 ? 'Japanese-Optimized' :
                         headerScore >= 70 ? 'Pattern-Focused' : 'Balanced';

    console.log(`🔬 詳細分析[${index}] "${header}" → ${targetType} [${weightProfile}]:`, {
      headerLower,
      headerScore: `${headerScore} (${(headerScore * headerWeight).toFixed(1)})`,
      contentScore: `${contentScore} (${(contentScore * contentWeight).toFixed(1)})`,
      linguisticScore: `${linguisticScore} (${(linguisticScore * linguisticWeight).toFixed(1)})`,
      contextScore: `${contextScore} (${(contextScore * contextWeight).toFixed(1)})`,
      semanticScore: `${semanticScore} (${(semanticScore * semanticWeight).toFixed(1)})`,
      weights: `H:${(headerWeight*100).toFixed(0)}% C:${(contentWeight*100).toFixed(0)}% L:${(linguisticWeight*100).toFixed(0)}%`,
      totalConfidence: totalConfidence.toFixed(1),
      finalConfidence: finalConfidence.toFixed(1),
      samplesCount: samples.length
    });
  }

  return {
    index,
    confidence: finalConfidence,
    factors
  };
}

/**
 * 1️⃣ ヘッダーパターン分析 - 高度な正規表現と重み付きキーワード
 */
function analyzeHeaderPattern(headerLower, targetType) {
  const patterns = {
    answer: {
      primary: [/^回答$/, /^答え$/, /^answer$/, /^response$/],
      strong: [/回答/, /答え/, /answer/, /意見/, /予想/, /考え/, /思う/, /選択/, /choice/],
      medium: [/結果/, /result/, /値/, /value/, /内容/, /content/],
      weak: [/データ/, /data/, /情報/, /info/]
    },
    reason: {
      primary: [/^理由$/, /^根拠$/, /^reason$/, /^説明$/],
      strong: [/理由/, /根拠/, /reason/, /なぜ/, /why/, /わけ/, /説明/, /explanation/],
      medium: [/詳細/, /detail/, /背景/, /background/, /コメント/, /comment/],
      weak: [/その他/, /other/, /備考/, /note/]
    },
    class: {
      primary: [/^クラス$/, /^class$/, /^組$/, /^年組$/],
      strong: [/クラス/, /class/, /組/, /年組/, /学級/, /grade/],
      medium: [/学年/, /year/, /グループ/, /group/],
      weak: [/チーム/, /team/]
    },
    name: {
      primary: [/^名前$/, /^氏名$/, /^name$/],
      strong: [/名前/, /氏名/, /name/, /お名前/, /ネーム/, /ニックネーム/],
      medium: [/ユーザー/, /user/, /学生/, /student/],
      weak: [/id/, /アカウント/, /account/]
    }
  };

  const typePatterns = patterns[targetType] || {};

  // デバッグ用の一時変数
  let matchedPattern = null;
  let matchedLevel = null;
  let score = 0;

  // 段階的マッチング
  for (const pattern of typePatterns.primary || []) {
    if (pattern.test(headerLower)) {
      matchedPattern = pattern.toString();
      matchedLevel = 'primary';
      score = 95;
      break;
    }
  }

  if (score === 0) {
    for (const pattern of typePatterns.strong || []) {
      if (pattern.test(headerLower)) {
        matchedPattern = pattern.toString();
        matchedLevel = 'strong';
        score = 85;
        break;
      }
    }
  }

  if (score === 0) {
    for (const pattern of typePatterns.medium || []) {
      if (pattern.test(headerLower)) {
        matchedPattern = pattern.toString();
        matchedLevel = 'medium';
        score = 60;
        break;
      }
    }
  }

  if (score === 0) {
    for (const pattern of typePatterns.weak || []) {
      if (pattern.test(headerLower)) {
        matchedPattern = pattern.toString();
        matchedLevel = 'weak';
        score = 35;
        break;
      }
    }
  }

  // マッチした場合のデバッグログ
  if (score > 0) {
    console.log(`🎯 ヘッダーパターンマッチ: "${headerLower}" → ${targetType} (${matchedLevel}: ${matchedPattern}) = ${score}%`);
  }

  return score;
}

/**
 * 2️⃣ コンテンツ統計分析 - データの特性を統計的に分析
 */
function analyzeContentStatistics(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  // 文字数統計
  const lengths = textSamples.map(s => s.trim().length);
  const avgLength = lengths.reduce((a, b) => a + b, 0) / lengths.length;
  const lengthVariance = lengths.reduce((sum, len) => sum + Math.pow(len - avgLength, 2), 0) / lengths.length;

  // 統計的特徴に基づく判定
  switch (targetType) {
    case 'answer':
      // 回答は一般的に短く、バリエーションが少ない
      if (avgLength <= 20 && lengthVariance <= 100) return 75;
      if (avgLength <= 50 && lengthVariance <= 500) return 60;
      if (avgLength <= 100) return 40;
      return 20;

    case 'reason':
      // 理由は一般的に長く、バリエーションが多い
      if (avgLength >= 50 && lengthVariance >= 200) return 80;
      if (avgLength >= 30 && lengthVariance >= 100) return 65;
      if (avgLength >= 20) return 45;
      return 25;

    case 'class':
      // クラスは短く、パターンが限定的
      if (avgLength <= 10 && lengthVariance <= 20) return 85;
      if (avgLength <= 20 && lengthVariance <= 50) return 65;
      return 30;

    case 'name':
      // 名前は中程度の長さで、適度なバリエーション
      if (avgLength >= 5 && avgLength <= 30 && lengthVariance <= 100) return 75;
      if (avgLength >= 3 && avgLength <= 50) return 55;
      return 35;
  }

  return 0;
}

/**
 * 3️⃣ 言語パターン分析 - 言語的特徴を分析
 */
function analyzeLinguisticPatterns(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  let score = 0;
  const sampleText = textSamples.join(' ').toLowerCase();

  switch (targetType) {
    case 'answer':
      // 回答によく現れるパターン
      if (/[あいうえお]だと思[う|います]/.test(sampleText)) score += 30;
      if (/[はい|いいえ|yes|no]/.test(sampleText)) score += 25;
      if (/\d+[番号]/.test(sampleText)) score += 20;
      if (/[選択肢]/.test(sampleText)) score += 15;
      break;

    case 'reason':
      // 理由によく現れるパターン
      if (/[だから|なぜなら|because]/.test(sampleText)) score += 35;
      if (/[と思う|と考える|だと思います]/.test(sampleText)) score += 25;
      if (/[ため|理由|根拠]/.test(sampleText)) score += 20;
      if (/[経験|体験|感じ]/.test(sampleText)) score += 15;
      break;

    case 'class':
      // クラス情報によく現れるパターン
      if (/\d+[年組班]/.test(sampleText)) score += 40;
      if (/[a-z]+[class|group]/.test(sampleText)) score += 30;
      if (/[グループ|チーム]\d+/.test(sampleText)) score += 20;
      break;

    case 'name':
      // 名前によく現れるパターン
      if (/^[ぁ-んァ-ン一-龯]+$/.test(sampleText)) score += 30; // 日本語名
      if (/^[a-zA-Z\s]+$/.test(sampleText)) score += 25; // 英語名
      if (/[さん|くん|ちゃん]$/.test(sampleText)) score += 20; // 敬称
      break;
  }

  return Math.min(score, 100);
}

/**
 * 4️⃣ コンテキスト推論 - 列位置と関係性を分析
 */
function analyzeContextualClues(header, index, allHeaders, targetType) {
  let score = 0;
  const headerLower = header.toLowerCase();

  // 列位置による推論
  const totalColumns = allHeaders.length;
  const position = index / (totalColumns - 1); // 0-1の相対位置

  switch (targetType) {
    case 'answer':
      // 回答列は通常中央付近に位置
      if (position >= 0.3 && position <= 0.7) score += 20;
      // タイムスタンプの後に来ることが多い
      if (index > 0 && allHeaders[index - 1] &&
          allHeaders[index - 1].toLowerCase().includes('timestamp')) score += 15;
      break;

    case 'reason':
      // 理由列は回答の後に来ることが多い
      if (index > 0) {
        const prevHeader = allHeaders[index - 1].toLowerCase();
        if (prevHeader.includes('回答') || prevHeader.includes('answer')) score += 25;
      }
      // 通常後半に位置
      if (position >= 0.5) score += 15;
      break;

    case 'class':
      // クラス情報は通常最初の方に位置
      if (position <= 0.3) score += 25;
      if (index <= 2) score += 20;
      break;

    case 'name':
      // 名前は通常最初の方に位置
      if (position <= 0.2) score += 30;
      if (index <= 1) score += 25;
      break;
  }

  // 隣接列との関係性分析
  const adjacentHeaders = [
    index > 0 ? allHeaders[index - 1] : null,
    index < allHeaders.length - 1 ? allHeaders[index + 1] : null
  ].filter(h => h).map(h => h.toLowerCase());

  for (const adjacent of adjacentHeaders) {
    if (targetType === 'answer' && adjacent.includes('reason')) score += 10;
    if (targetType === 'reason' && adjacent.includes('answer')) score += 10;
    if (targetType === 'name' && adjacent.includes('class')) score += 10;
  }

  return Math.min(score, 100);
}

/**
 * 5️⃣ セマンティック分析 - 意味的特徴を分析
 */
function analyzeSemanticCharacteristics(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  let score = 0;
  const uniqueValues = [...new Set(textSamples)];
  const uniquenessRatio = uniqueValues.length / textSamples.length;

  switch (targetType) {
    case 'answer':
      // 回答は選択肢的で重複が多い
      if (uniquenessRatio <= 0.3) score += 30;
      if (uniquenessRatio <= 0.5) score += 20;
      // 数値や選択肢パターン
      if (textSamples.some(s => /^[1-9]$/.test(s))) score += 25;
      if (textSamples.some(s => /^[A-D]$/.test(s))) score += 25;
      break;

    case 'reason':
      // 理由は個別性が高く、重複が少ない
      if (uniquenessRatio >= 0.8) score += 35;
      if (uniquenessRatio >= 0.6) score += 25;
      // 説明的な言葉
      if (textSamples.some(s => s.includes('ため'))) score += 15;
      break;

    case 'class':
      // クラス情報は限定的なパターン
      if (uniquenessRatio <= 0.2) score += 40;
      if (textSamples.some(s => /\d/.test(s))) score += 20;
      break;

    case 'name':
      // 名前は個別性が高い
      if (uniquenessRatio >= 0.7) score += 35;
      if (uniquenessRatio >= 0.5) score += 20;
      break;
  }

  return Math.min(score, 100);
}

// ===========================================
// 🔧 ヘルパー関数（必要な実装）
// ===========================================

/**
 * リアクション列の取得または作成
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} reactionType - リアクションタイプ
 * @returns {number} 列番号
 */
function getOrCreateReactionColumn(sheet, reactionType) {
  try {
    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    const reactionHeader = reactionType.toUpperCase();

    // 既存の列を探す
    const existingIndex = headers.findIndex(header =>
      String(header).toUpperCase().includes(reactionHeader)
    );

    if (existingIndex !== -1) {
      return existingIndex + 1; // 1-based index
    }

    // 新しい列を作成
    const newColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, newColumn).setValue(reactionHeader);
    return newColumn;
  } catch (error) {
    console.error('getOrCreateReactionColumn: エラー', error.message);
    return null;
  }
}

/**
 * ハイライト列の更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @returns {boolean} 成功可否
 */
function updateHighlightInSheet(config, rowId) {
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

    // ハイライト列の取得・作成
    const highlightColumn = getOrCreateReactionColumn(sheet, 'HIGHLIGHT');
    if (!highlightColumn) {
      throw new Error('ハイライト列の作成に失敗');
    }

    // 現在値取得・トグル
    const currentValue = sheet.getRange(rowNumber, highlightColumn).getValue();
    const isCurrentlyHighlighted = currentValue === 'TRUE' || currentValue === true;
    const newValue = isCurrentlyHighlighted ? 'FALSE' : 'TRUE';

    sheet.getRange(rowNumber, highlightColumn).setValue(newValue);

    const highlighted = newValue === 'TRUE';

    console.info('DataService.updateHighlightInSheet: ハイライト切り替え完了', {
      rowId,
      oldValue: currentValue,
      newValue,
      highlighted
    });

    return {
      success: true,
      highlighted
    };
  } catch (error) {
    console.error('DataService.updateHighlightInSheet: エラー', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * リアクションパラメータ検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類
 * @returns {boolean} 検証結果
 */
function validateReaction(spreadsheetId, sheetName, rowIndex, reactionKey) {
  if (!spreadsheetId || !sheetName || !rowIndex || !reactionKey) {
    return false;
  }

  if (rowIndex < 2) { // ヘッダー行は1
    return false;
  }

  const validReactions = ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
  return validReactions.includes(reactionKey);
}

// ===========================================
// 🌍 Public DataService Namespace
// ===========================================

/**
 * addReaction (user context)
 * @param {string} userId
 * @param {number|string} rowId - number or 'row_#'
 * @param {string} reactionType
 * @returns {Object}
 */
function dsAddReaction(userId, rowId, reactionType) {
  try {
    const db = ServiceFactory.getDB();
    const user = db && db.findUserById ? db.findUserById(userId) : null;
    if (!user || !user.configJson) {
      return createErrorResponse('User configuration not found');
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    const rowIndex = typeof rowId === 'string' ? parseInt(String(rowId).replace('row_', ''), 10) : parseInt(rowId, 10);
    if (!rowIndex || rowIndex < 2) {
      return createErrorResponse('Invalid row ID');
    }

    const res = processReaction(config.spreadsheetId, config.sheetName, rowIndex, reactionType, null);
    if (res && (res.success || res.status === 'success')) {
      return {
        success: true,
        message: res.message || 'Reaction added',
        newValue: res.newValue
      };
    }

    return {
      success: false,
      message: res?.message || 'Failed to add reaction'
    };
  } catch (error) {
    console.error('DataService.dsAddReaction: エラー', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * toggleHighlight (user context)
 * @param {string} userId
 * @param {number|string} rowId - number or 'row_#'
 * @returns {Object}
 */
function dsToggleHighlight(userId, rowId) {
  try {
    const db = ServiceFactory.getDB();
    const user = db && db.findUserById ? db.findUserById(userId) : null;
    if (!user || !user.configJson) {
      return createErrorResponse('User configuration not found');
    }

    const config = JSON.parse(user.configJson);
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    // updateHighlightInSheet expects 'row_#'
    const rowNumber = typeof rowId === 'string' && rowId.startsWith('row_')
      ? rowId
      : `row_${parseInt(rowId, 10)}`;

    const result = updateHighlightInSheet(config, rowNumber);
    if (result?.success) {
      return {
        success: true,
        message: 'Highlight toggled successfully',
        highlighted: Boolean(result.highlighted)
      };
    }

    return {
      success: false,
      message: result?.error || 'Failed to toggle highlight'
    };
  } catch (error) {
    console.error('DataService.dsToggleHighlight: エラー', error.message);
    return createExceptionResponse(error);
  }
}

// Expose a stable namespace for non-global access patterns
if (typeof global !== 'undefined') {
  global.DataService = {
    getUserSheetData,
    processReaction,
    addReaction: dsAddReaction,
    toggleHighlight: dsToggleHighlight
  };
} else {
  this.DataService = {
    getUserSheetData,
    processReaction,
    addReaction: dsAddReaction,
    toggleHighlight: dsToggleHighlight
  };
}
