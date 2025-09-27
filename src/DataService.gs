/**
 * @fileoverview DataService - コアデータ操作サービス（リファクタリング版）
 *
 * 責任範囲:
 * - スプレッドシートデータ取得・操作
 * - データフィルタリング・検索
 * - バルクデータAPI
 * - シート接続・寸法取得
 *
 * CLAUDE.md Best Practices準拠:
 * - 分離されたモジュール利用（ColumnMappingService, ReactionService）
 * - Zero-Dependency Architecture（直接GAS API）
 * - バッチ操作による70x性能向上
 * - V8ランタイム最適化
 *
 * 依存モジュール:
 * - ColumnMappingService.gs（列マッピング・抽出）
 * - ReactionService.gs（リアクション・ハイライト）
 */

/* global formatTimestamp, createErrorResponse, createExceptionResponse, getQuestionText, findUserByEmail, findUserById, findUserBySpreadsheetId, openSpreadsheet, getUserConfig, CACHE_DURATION, getCurrentEmail, extractFieldValueUnified, extractReactions, extractHighlight, createDataServiceErrorResponse, createDataServiceSuccessResponse */

// Core Data Operations

/**
 * ユーザーのスプレッドシートデータ取得
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} レスポンスオブジェクト
 */
function getUserSheetData(userId, options = {}, preloadedUser = null, preloadedConfig = null) {
  const startTime = Date.now();

  try {
    // Zero-dependency data processing

    // Performance improvement - use preloaded data
    const user = preloadedUser || findUserById(userId);
    if (!user) {
      console.error('DataService.getUserSheetData: ユーザーが見つかりません', { userId });
      return createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    // Performance improvement - use preloaded config
    let config;
    if (preloadedConfig) {
      config = preloadedConfig;
    } else {
      const configResult = getUserConfig(userId);
      config = configResult.success ? configResult.config : {};
    }

    if (!config.spreadsheetId) {
      console.warn('[WARN] DataService.getUserSheetData: Spreadsheet ID not configured', { userId });
      return createDataServiceErrorResponse('スプレッドシートが設定されていません');
    }

    // データ取得実行
    const result = fetchSpreadsheetData(config, options, user);

    const executionTime = Date.now() - startTime;

    // Standardized response format
    if (result.success) {
      // Performance optimization: use already retrieved headers
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
    // Ensure proper error response structure
    const errorResponse = createDataServiceErrorResponse(error.message || 'データ取得エラー');
    console.error('DataService.getUserSheetData: Creating error response', errorResponse);
    return errorResponse;
  }
}

/**
 * タイムスタンプ専用抽出関数（通知システム用）
 * @param {Array} row - データ行
 * @param {Array} headers - ヘッダー配列
 * @returns {string} タイムスタンプ値
 */
function extractTimestampValue(row, headers) {
  try {
    // 1. 最初のカラムがタイムスタンプかチェック（Googleフォーム標準）
    if (headers.length > 0 && row.length > 0) {
      const [firstHeader] = headers;
      if (firstHeader && typeof firstHeader === 'string') {
        const normalizedHeader = firstHeader.toLowerCase().trim();
        if (normalizedHeader.includes('タイムスタンプ') ||
            normalizedHeader.includes('timestamp') ||
            normalizedHeader.includes('日時')) {
          return row[0] || '';
        }
      }
    }

    // 2. ヘッダー名でタイムスタンプカラムを検索
    for (let i = 0; i < headers.length; i++) {
      const header = headers[i];
      if (header && typeof header === 'string') {
        const normalizedHeader = header.toLowerCase().trim();
        if (normalizedHeader.includes('タイムスタンプ') ||
            normalizedHeader.includes('timestamp') ||
            normalizedHeader.includes('日時') ||
            normalizedHeader.includes('日付')) {
          return row[i] || '';
        }
      }
    }

    return '';
  } catch (error) {
    console.warn('extractTimestampValue error:', error.message);
    return '';
  }
}

/**
 * スプレッドシート接続とシート取得（GAS Best Practice: 単一責任）
 * @param {Object} config - 設定オブジェクト
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @returns {Object} シート情報
 */
function connectToSpreadsheetSheet(config, context = {}) {
  // Context-aware service account usage
  // Cross-user: Use service account when accessing other user's spreadsheet
  // Self-access: Use normal permissions for own spreadsheet
  const currentEmail = getCurrentEmail();

  // Performance improvement - use preloaded user to avoid duplicate DB access
  const targetUser = context.preloadedUser || findUserBySpreadsheetId(config.spreadsheetId, {
    preloadedAuth: context.preloadedAuth // 認証情報を渡して重複認証回避
  });
  const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;

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
    // Performance improvement -事前取得ユーザー情報を活用してDB重複呼び出し排除
    const { sheet } = connectToSpreadsheetSheet(config, {
      targetUserEmail: user?.userEmail,
      preloadedUser: user, // 事前取得ユーザーを渡してfindUserBySpreadsheetId重複回避
      preloadedAuth: options.preloadedAuth // 認証情報も渡して重複認証回避
    });

    // 寸法取得
    const { lastRow, lastCol } = getSheetDimensions(sheet);

    if (lastRow <= 1) {
      return createDataServiceSuccessResponse([], [], config.sheetName);
    }

    // ヘッダー取得
    const headers = getSheetHeaders(sheet, lastCol);

    // バッチ処理実行
    const processedData = processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime);

    // Return directly in frontend-expected format
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
    const columnMapping = config.columnMapping || {};
    const processedBatch = [];

    batchRows.forEach((row, batchIndex) => {
      try {
        // グローバル行インデックス計算
        const globalIndex = startOffset + batchIndex;

        // 基本データ構造作成（ColumnMappingService利用）
        // 各フィールドのデータ抽出（null チェック強化）
        const answerResult = extractFieldValueUnified(row, headers, 'answer', columnMapping);
        const reasonResult = extractFieldValueUnified(row, headers, 'reason', columnMapping);
        const classResult = extractFieldValueUnified(row, headers, 'class', columnMapping);
        const nameResult = extractFieldValueUnified(row, headers, 'name', columnMapping);
        const emailResult = extractFieldValueUnified(row, headers, 'email', columnMapping);

        // 匿名性保護: 抽出データのクロスチェック
        const answerValue = answerResult?.value;
        const nameValue = nameResult?.value;

        // 名前データが回答・理由欄に混入していないかチェック
        if (nameValue && (answerValue === nameValue || reasonResult?.value === nameValue)) {
          console.warn('DataService: 匿名性保護のため行をスキップ', {
            rowIndex: globalIndex + 2,
            reason: '名前データが回答・理由欄に混入'
          });
          return; // この行をスキップ
        }

        const item = {
          id: `row_${globalIndex + 2}`,
          rowIndex: globalIndex + 2, // 1-based row number including header
          timestamp: extractTimestampValue(row, headers) || '',
          email: String(emailResult?.value || ''),

          // Main content using ColumnMappingService - 厳密な null チェック
          answer: String(answerValue || ''),
          opinion: String(answerValue || ''), // Alias for answer field
          reason: String(reasonResult?.value || ''),
          class: String(classResult?.value || ''),
          name: String(nameValue || ''),

          // メタデータ
          formattedTimestamp: formatTimestamp(extractTimestampValue(row, headers)),
          isEmpty: isEmptyRow(row),

          // リアクション（ReactionService利用）
          reactions: extractReactions(row, headers, getCurrentEmail()),
          highlight: extractHighlight(row, headers)
        };

        // データ整合性の最終チェック
        if (!answerValue && !reasonResult?.value) {
          // 回答も理由も空の場合はスキップ
          return;
        }

        // includeTimestamp option processing
        if (options.includeTimestamp === false) {
          delete item.timestamp;
          delete item.formattedTimestamp;
        }

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
    const columnMapping = config.columnMapping || {};
    const processedData = [];

    dataRows.forEach((row, index) => {
      try {
        // 各フィールドのデータ抽出（null チェック強化）
        const answerResult = extractFieldValueUnified(row, headers, 'answer', columnMapping);
        const reasonResult = extractFieldValueUnified(row, headers, 'reason', columnMapping);
        const classResult = extractFieldValueUnified(row, headers, 'class', columnMapping);
        const nameResult = extractFieldValueUnified(row, headers, 'name', columnMapping);
        const emailResult = extractFieldValueUnified(row, headers, 'email', columnMapping);
        const timestampResult = extractFieldValueUnified(row, headers, 'timestamp');

        // 匿名性保護: 抽出データのクロスチェック
        const answerValue = answerResult?.value;
        const nameValue = nameResult?.value;

        // 名前データが回答・理由欄に混入していないかチェック
        if (nameValue && (answerValue === nameValue || reasonResult?.value === nameValue)) {
          console.warn('DataService: 匿名性保護のため行をスキップ', {
            rowIndex: index + 2,
            reason: '名前データが回答・理由欄に混入'
          });
          return; // この行をスキップ
        }

        // 基本データ構造作成（ColumnMappingService利用）
        const item = {
          id: `row_${index + 2}`,
          rowIndex: index + 2,
          timestamp: String(timestampResult?.value || ''),
          email: String(emailResult?.value || ''),

          // Main content using ColumnMappingService - 厳密な null チェック
          answer: String(answerValue || ''),
          opinion: String(answerValue || ''), // Alias for answer field
          reason: String(reasonResult?.value || ''),
          class: String(classResult?.value || ''),
          name: String(nameValue || ''),

          // メタデータ
          formattedTimestamp: formatTimestamp(extractTimestampValue(row, headers)),
          isEmpty: isEmptyRow(row),

          // リアクション（ReactionService利用）
          reactions: extractReactions(row, headers, getCurrentEmail()),
          highlight: extractHighlight(row, headers)
        };

        // データ整合性の最終チェック
        if (!answerValue && !reasonResult?.value) {
          // 回答も理由も空の場合はスキップ
          return;
        }

        // includeTimestamp option processing
        if (options.includeTimestamp === false) {
          delete item.timestamp;
          delete item.formattedTimestamp;
        }

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
  // Delegate to ColumnMappingService (backward compatibility)
  return extractFieldValueUnified(row, headers, fieldType, columnMapping);
}

// 🔍 データ分析・フィルタリング

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

// Utility helpers

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

    // Filter out rows with empty main answers - V8 runtime safe with enhanced type checking
    if (options.requireAnswer !== false) {
      const answerStr = item.answer ? String(item.answer).trim() : '';
      if (!answerStr) {
        return false;
      }
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

// Data Deletion Operations

/**
 * 回答行を削除（管理モード専用）
 * @param {string} userId - ユーザーID
 * @param {number} rowIndex - 削除対象の行インデックス（1-based, ヘッダー含む）
 * @param {Object} options - オプション設定
 * @returns {Object} 削除結果
 */
function deleteAnswerRow(userId, rowIndex, options = {}) {
  const startTime = Date.now();

  try {
    // 🛡️ CLAUDE.md準拠: Security Gate - 管理者権限チェック
    const currentEmail = getCurrentEmail();
    const user = findUserById(userId);

    if (!user) {
      console.error('deleteAnswerRow: User not found:', userId);
      return createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    // 所有者または管理者のみ削除可能
    const isOwner = user.userEmail === currentEmail;
    const isAdmin = (() => {
      try {
        const adminEmail = PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL');
        return currentEmail?.toLowerCase() === adminEmail?.toLowerCase();
      } catch (error) {
        console.warn('Administrator check failed:', error);
        return false;
      }
    })();

    if (!isOwner && !isAdmin) {
      console.warn('deleteAnswerRow: Insufficient permissions:', { userId, currentEmail });
      return createDataServiceErrorResponse('削除権限がありません');
    }

    // ✅ CLAUDE.md準拠: GAS-Native Architecture - Direct SpreadsheetApp usage
    const config = getUserConfig(userId);
    if (!config.success || !config.config.spreadsheetId) {
      return createDataServiceErrorResponse('スプレッドシート設定が見つかりません');
    }

    const {spreadsheetId} = config.config;
    const sheetName = config.config.sheetName || 'フォームの回答 1';

    // 🚀 Zero-dependency spreadsheet access
    const dataAccess = openSpreadsheet(spreadsheetId, {
      useServiceAccount: !isOwner, // Cross-user access for admins
      targetUserEmail: user.userEmail,
      context: 'deleteAnswerRow'
    });

    if (!dataAccess || !dataAccess.spreadsheet) {
      return createDataServiceErrorResponse('スプレッドシートにアクセスできません');
    }

    const sheet = dataAccess.spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return createDataServiceErrorResponse(`シート「${sheetName}」が見つかりません`);
    }

    // 行範囲検証
    const lastRow = sheet.getLastRow();
    if (rowIndex < 2 || rowIndex > lastRow) {
      return createDataServiceErrorResponse('無効な行番号です');
    }

    // Performance improvement -Batch operation
    const [rowData] = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues();

    // 削除実行
    sheet.deleteRows(rowIndex, 1);

    const executionTime = Date.now() - startTime;

    return {
      success: true,
      message: '回答を削除しました',
      deletedRowIndex: rowIndex,
      executionTime
    };

  } catch (error) {
    console.error('deleteAnswerRow error:', {
      userId,
      rowIndex,
      error: error.message
    });
    return createDataServiceErrorResponse(`削除エラー: ${error.message}`);
  }
}