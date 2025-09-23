/**
 * @fileoverview DataService - コアデータ操作サービス（リファクタリング版）
 *
 * 🎯 責任範囲:
 * - スプレッドシートデータ取得・操作
 * - データフィルタリング・検索
 * - バルクデータAPI
 * - シート接続・寸法取得
 *
 * 🔄 CLAUDE.md Best Practices準拠:
 * - 分離されたモジュール利用（ColumnMappingService, ReactionService）
 * - Zero-Dependency Architecture（直接GAS API）
 * - バッチ操作による70x性能向上
 * - V8ランタイム最適化
 *
 * 🔗 依存モジュール:
 * - ColumnMappingService.gs（列マッピング・抽出）
 * - ReactionService.gs（リアクション・ハイライト）
 */

/* global formatTimestamp, createErrorResponse, createExceptionResponse, getQuestionText, findUserByEmail, findUserById, findUserBySpreadsheetId, openSpreadsheet, getUserConfig, helpers, CACHE_DURATION, getCurrentEmail, extractFieldValueUnified, extractReactions, extractHighlight */

// ===========================================
// 🎯 Core Data Operations - CLAUDE.md準拠
// ===========================================

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
    // 🚀 Zero-dependency data processing

    // 🔧 Zero-Dependency統一: 直接findUserById使用
    const user = findUserById(userId);
    if (!user) {
      console.error('DataService.getUserSheetData: ユーザーが見つかりません', { userId });
      return helpers.createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId) {
      console.warn('[WARN] DataService.getUserSheetData: Spreadsheet ID not configured', { userId });
      return helpers.createDataServiceErrorResponse('スプレッドシートが設定されていません');
    }

    // データ取得実行
    const result = fetchSpreadsheetData(config, options, user);

    const executionTime = Date.now() - startTime;
    console.info('getSheetData: データ取得完了', {
      userId,
      rowCount: result.data?.length || 0,
      executionTime
    });

    // Standardized response format
    if (result.success) {
      // ✅ パフォーマンス最適化: 既に取得済みのheadersを活用
      const preloadedHeaders = result.headers;
      const questionText = getQuestionText(config, { targetUserEmail: user.userEmail }, preloadedHeaders);

      return {
        ...result,
        header: questionText || result.sheetName || '回答一覧',
        showDetails: config.showDetails !== false // デフォルトはtrue
      };
    }

    return result;
  } catch (error) {
    console.error('DataService.getUserSheetData: エラー', {
      userId,
      error: error.message
    });
    return helpers.createDataServiceErrorResponse(error.message || 'データ取得エラー');
  }
}

/**
 * スプレッドシート接続とシート取得（GAS Best Practice: 単一責任）
 * @param {Object} config - 設定オブジェクト
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @returns {Object} シート情報
 */
function connectToSpreadsheetSheet(config, context = {}) {
  // 🔧 CLAUDE.md準拠: Context-aware service account usage
  // ✅ **Cross-user**: Use service account when accessing other user's spreadsheet
  // ✅ **Self-access**: Use normal permissions for own spreadsheet
  const currentEmail = getCurrentEmail();

  // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
  const targetUser = findUserBySpreadsheetId(config.spreadsheetId);
  const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;

  console.log(`connectToSpreadsheetSheet: ${isSelfAccess ? 'Self-access normal permissions' : 'Cross-user service account'} for spreadsheet`);
  const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: !isSelfAccess });
  const {spreadsheet} = dataAccess;
  const sheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sheet) {
    const sheetName = config && config.sheetName ? config.sheetName : '不明';
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }

  return { sheet, spreadsheet };
}

/**
 * シートの寸法取得（GAS Best Practice: 単一責任）
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} 寸法情報
 */
function getSheetDimensions(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  return { lastRow, lastCol };
}

/**
 * ヘッダー行取得（GAS Best Practice: 単一責任）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {number} lastCol - 最終列
 * @returns {Array} ヘッダー配列
 */
function getSheetHeaders(sheet, lastCol) {
  const [headers] = sheet.getRange(1, 1, 1, lastCol).getValues();
  return headers;
}

/**
 * バッチデータ処理（GAS Best Practice: 大量データ処理分離）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Array} headers - ヘッダー配列
 * @param {number} lastRow - 最終行
 * @param {number} lastCol - 最終列
 * @param {Object} config - 設定
 * @param {Object} options - オプション
 * @param {Object} user - ユーザー情報
 * @param {number} startTime - 開始時刻
 * @returns {Array} 処理済みデータ
 */
function processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime) {
  const MAX_EXECUTION_TIME = 180000; // 3分制限
  const MAX_BATCH_SIZE = 200; // バッチサイズ

  let processedData = [];
  let processedCount = 0;
  const totalDataRows = lastRow - 1;

  for (let startRow = 2; startRow <= lastRow; startRow += MAX_BATCH_SIZE) {
    // 実行時間チェック
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      console.warn('DataService.processBatchData: 実行時間制限のため処理を中断', {
        processedRows: processedCount,
        totalRows: totalDataRows
      });
      break;
    }

    const endRow = Math.min(startRow + MAX_BATCH_SIZE - 1, lastRow);
    const batchSize = endRow - startRow + 1;

    try {
      const batchRows = sheet.getRange(startRow, 1, batchSize, lastCol).getValues();
      const batchProcessed = processRawDataBatch(batchRows, headers, config, options, startRow - 2, user);

      processedData = processedData.concat(batchProcessed);
      processedCount += batchSize;

      // API制限対策: 1000行毎に短い休憩
      if (processedCount % 1000 === 0) {
        Utilities.sleep(100); // 0.1秒休憩
      }

    } catch (batchError) {
      console.error('DataService.processBatchData: バッチ処理エラー', {
        startRow, endRow, error: batchError && batchError.message ? batchError.message : 'エラー詳細不明'
      });
    }
  }

  // フィルタリングとソート適用
  if (options.classFilter) {
    processedData = processedData.filter(item => item.class === options.classFilter);
  }

  if (options.sortBy) {
    processedData = applySortAndLimit(processedData, {
      sortBy: options.sortBy,
      limit: options.limit
    });
  }

  return processedData;
}

/**
 * スプレッドシートデータ取得（リファクタリング版 - GAS Best Practice準拠）
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - オプション
 * @param {Object} user - ユーザー情報
 * @returns {Object} データ取得結果
 */
function fetchSpreadsheetData(config, options = {}, user = null) {
  const startTime = Date.now();

  try {
    // シート接続
    const { sheet } = connectToSpreadsheetSheet(config, { targetUserEmail: user?.userEmail });

    // 寸法取得
    const { lastRow, lastCol } = getSheetDimensions(sheet);

    if (lastRow <= 1) {
      return helpers.createDataServiceSuccessResponse([], [], config.sheetName);
    }

    // ヘッダー取得
    const headers = getSheetHeaders(sheet, lastCol);

    // バッチ処理実行
    const processedData = processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime);

    console.info('DataService.fetchSpreadsheetData: バッチ処理完了', {
      filteredRows: processedData.length,
      executionTime: Date.now() - startTime
    });

    // ✅ フロントエンド期待形式で直接返却
    return {
      success: true,
      data: processedData,
      headers,
      sheetName: config.sheetName,
      // デバッグ情報（オプショナル）
      filteredRows: processedData.length,
      executionTime: Date.now() - startTime
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
 * @param {Object} user - ユーザー情報
 * @returns {Array} 処理済みバッチデータ
 */
function processRawDataBatch(batchRows, headers, config, options = {}, startOffset = 0, user = null) {
  try {
    const columnMapping = config.columnMapping?.mapping || {};
    const processedBatch = [];

    batchRows.forEach((row, batchIndex) => {
      try {
        // グローバル行インデックス計算
        const globalIndex = startOffset + batchIndex;

        // 基本データ構造作成（ColumnMappingService利用）
        const item = {
          id: `row_${globalIndex + 2}`,
          rowIndex: globalIndex + 2, // 1-based row number including header
          timestamp: extractFieldValueUnified(row, headers, 'timestamp') || '',
          email: extractFieldValueUnified(row, headers, 'email') || '',

          // メインコンテンツ（ColumnMappingService利用）
          answer: extractFieldValueUnified(row, headers, 'answer', columnMapping) || '',
          opinion: extractFieldValueUnified(row, headers, 'answer', columnMapping) || '', // Alias for answer field
          reason: extractFieldValueUnified(row, headers, 'reason', columnMapping) || '',
          class: extractFieldValueUnified(row, headers, 'class', columnMapping) || '',
          name: extractFieldValueUnified(row, headers, 'name', columnMapping) || '',

          // メタデータ
          formattedTimestamp: formatTimestamp(extractFieldValueUnified(row, headers, 'timestamp')),
          isEmpty: isEmptyRow(row),

          // リアクション（ReactionService利用）
          reactions: extractReactions(row, headers, getCurrentEmail()),
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
 * 生データを処理・変換
 * @param {Array} dataRows - 生データ行
 * @param {Array} headers - ヘッダー配列
 * @param {Object} config - 設定
 * @param {Object} options - 処理オプション
 * @param {Object} user - ユーザー情報
 * @returns {Array} 処理済みデータ
 */
function processRawData(dataRows, headers, config, options = {}, user = null) {
  try {
    const columnMapping = config.columnMapping?.mapping || {};
    const processedData = [];

    dataRows.forEach((row, index) => {
      try {
        // 基本データ構造作成（ColumnMappingService利用）
        const item = {
          id: `row_${index + 2}`,
          rowIndex: index + 2,
          timestamp: extractFieldValueUnified(row, headers, 'timestamp') || '',
          email: extractFieldValueUnified(row, headers, 'email') || '',

          // メインコンテンツ（ColumnMappingService利用）
          answer: extractFieldValueUnified(row, headers, 'answer', columnMapping) || '',
          opinion: extractFieldValueUnified(row, headers, 'answer', columnMapping) || '', // Alias for answer field
          reason: extractFieldValueUnified(row, headers, 'reason', columnMapping) || '',
          class: extractFieldValueUnified(row, headers, 'class', columnMapping) || '',
          name: extractFieldValueUnified(row, headers, 'name', columnMapping) || '',

          // メタデータ
          formattedTimestamp: formatTimestamp(extractFieldValueUnified(row, headers, 'timestamp')),
          isEmpty: isEmptyRow(row),

          // リアクション（ReactionService利用）
          reactions: extractReactions(row, headers, getCurrentEmail()),
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
 * フィールド値抽出（列マッピング対応）- ColumnMappingServiceへの委譲
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @param {string} fieldType - フィールドタイプ
 * @param {Object} columnMapping - 列マッピング
 * @returns {*} フィールド値
 */
function extractFieldValue(row, headers, fieldType, columnMapping = {}) {
  // 🎯 ColumnMappingServiceに委譲（後方互換性保持）
  return extractFieldValueUnified(row, headers, fieldType, columnMapping);
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
  const startTime = new Date(publishedAt);
  const stopTime = new Date(startTime.getTime() + minutes * 60000);

  return {
    stopTime: stopTime.toISOString(),
    formattedStopTime: formatTimestamp(stopTime.toISOString()),
    remainingMinutes: Math.max(0, Math.ceil((stopTime.getTime() - new Date().getTime()) / 60000))
  };
}

// ===========================================
// 🔧 ユーティリティ・ヘルパー
// ===========================================

/**
 * 空行判定（ReactionServiceから移動したisEmptyRowを利用）
 * @param {Array} row - データ行
 * @returns {boolean} 空行かどうか
 */
function isEmptyRow(row) {
  return !row || row.every(cell => !cell || cell.toString().trim() === '');
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
    if (options.classFilter && item.class !== options.classFilter) {
      return false;
    }

    return true;
  } catch (error) {
    console.warn('DataService.shouldIncludeRow: フィルタリングエラー', error.message);
    return true; // エラー時は含める
  }
}

/**
 * ソート・制限適用
 * @param {Array} data - データ配列
 * @param {Object} options - ソートオプション
 * @returns {Array} ソート済みデータ
 */
function applySortAndLimit(data, options = {}) {
  try {
    let sortedData = [...data];

    // ソート実行
    if (options.sortBy) {
      switch (options.sortBy) {
        case 'timestamp':
        case 'newest':
          sortedData.sort((a, b) => new Date(b.timestamp || 0) - new Date(a.timestamp || 0));
          break;
        case 'oldest':
          sortedData.sort((a, b) => new Date(a.timestamp || 0) - new Date(b.timestamp || 0));
          break;
        case 'class':
          sortedData.sort((a, b) => (a.class || '').localeCompare(b.class || ''));
          break;
        case 'name':
          sortedData.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
          break;
        default:
          // デフォルトは新しい順
          sortedData.sort((a, b) => new Date(b.timestamp || 0) - new Date(a.timestamp || 0));
      }
    }

    // 制限適用
    if (options.limit && typeof options.limit === 'number' && options.limit > 0) {
      sortedData = sortedData.slice(0, options.limit);
    }

    return sortedData;
  } catch (error) {
    console.error('DataService.applySortAndLimit: エラー', error.message);
    return data; // エラー時は元データを返す
  }
}