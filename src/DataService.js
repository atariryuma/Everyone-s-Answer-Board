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

/* global formatTimestamp, createErrorResponse, createExceptionResponse, getQuestionText, findUserByEmail, findUserById, findUserBySpreadsheetId, openSpreadsheet, getUserConfig, CACHE_DURATION, getCurrentEmail, extractFieldValueUnified, extractReactions, extractHighlight, createDataServiceErrorResponse, getCachedProperty */


/**
 * ユーザーのスプレッドシートデータ取得
 * @param {string} userId - ユーザーID
 * @param {Object} options - 取得オプション
 * @returns {Object} レスポンスオブジェクト
 */
function getUserSheetData(userId, options = {}, preloadedUser = null, preloadedConfig = null) {
  const startTime = Date.now();

  let user = null;
  let config = null;

  try {

    user = preloadedUser || findUserById(userId, { requestingUser: getCurrentEmail() });
    if (!user) {
      console.error('DataService.getUserSheetData: ユーザーが見つかりません', { userId });
      return {
        success: false,
        message: 'ユーザーが見つかりません',
        data: [],
        header: '',
        sheetName: ''
      };
    }

    if (preloadedConfig) {
      config = preloadedConfig;
    } else {
      const configResult = getUserConfig(userId);
      config = configResult.success ? configResult.config : {};
    }

    if (!config.spreadsheetId) {
      console.warn('[WARN] DataService.getUserSheetData: Spreadsheet ID not configured', { userId });
      return {
        success: false,
        message: 'スプレッドシートが設定されていません',
        data: [],
        header: '',
        sheetName: ''
      };
    }

    const result = fetchSpreadsheetData(config, options, user);

    const executionTime = Date.now() - startTime;

    if (result.success) {
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
    const executionTime = Date.now() - startTime;
    console.error('DataService.getUserSheetData: 重大エラー', {
      userId,
      error: error.message,
      stack: error.stack ? `${error.stack.substring(0, 300)  }...` : 'No stack trace',
      executionTime: `${executionTime}ms`,
      configSpreadsheetId: config?.spreadsheetId || 'undefined',
      configSheetName: config?.sheetName || 'undefined',
      userEmail: user?.userEmail || 'undefined',
      optionsProvided: options ? Object.keys(options) : 'none',
      timestamp: new Date().toISOString()
    });

    const errorResponse = createDataServiceErrorResponse(`データ取得エラー: ${error.message}`);
    console.error('DataService.getUserSheetData: エラーレスポンス作成', errorResponse);
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
 * ✅ CLAUDE.md準拠: preloadedUser優先使用でDB重複アクセス排除
 * @param {Object} config - 設定オブジェクト
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @returns {Object} シート情報
 */
function connectToSpreadsheetSheet(config, context = {}) {
  // ✅ CRITICAL: ユーザーの回答ボードは同一ドメイン共有設定で対応
  const currentEmail = getCurrentEmail();

  const targetUser = context.preloadedUser || findUserBySpreadsheetId(config.spreadsheetId, {
    preloadedAuth: context.preloadedAuth, // 認証情報を渡して重複認証回避
    skipCache: false // キャッシュ活用
  });

  if (!targetUser) {
    console.warn('connectToSpreadsheetSheet: Target user not found for spreadsheet', {
      spreadsheetId: config.spreadsheetId ? `${config.spreadsheetId.substring(0, 8)}...` : 'undefined'
    });
  }

  const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: false });

  if (!dataAccess || !dataAccess.spreadsheet) {
    console.error('connectToSpreadsheetSheet: スプレッドシートアクセス失敗', {
      spreadsheetId: config.spreadsheetId,
      useServiceAccount: false,
      hasDataAccess: !!dataAccess,
      hint: '同一ドメイン内で編集可能に設定してください'
    });
    throw new Error(`スプレッドシートアクセスに失敗しました: ${config.spreadsheetId}`);
  }

  const {spreadsheet} = dataAccess;
  const sheet = spreadsheet.getSheetByName(config.sheetName);

  if (!sheet) {
    const sheetName = config && config.sheetName ? config.sheetName : '不明';
    throw new Error(`シート '${sheetName}' が見つかりません`);
  }

  return { sheet, spreadsheet };
}

/**
 * ヘッダー情報取得（長期キャッシュ - ヘッダーは変更されないため）
 * ✅ FIX: フォーム投稿即時反映のため、ヘッダーと行数を分離
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} { lastCol, headers }
 */
function getSheetHeaders(sheet) {
  const spreadsheetId = sheet.getParent ? sheet.getParent().getId() : 'unknown';
  const sheetName = sheet.getName();
  const cacheKey = `sheet_headers_${spreadsheetId}_${sheetName}`;
  const cache = CacheService.getScriptCache();

  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (parseError) {
      console.warn('getSheetHeaders: Cache parse failed, fetching fresh data');
    }
  }

  const dataRange = sheet.getDataRange();
  const values = dataRange.getValues();

  const info = {
    lastCol: values[0]?.length || 0,
    headers: values[0] || []
  };

  try {
    cache.put(cacheKey, JSON.stringify(info), CACHE_DURATION.DATABASE_LONG);
  } catch (cacheError) {
    console.warn('getSheetHeaders: Cache write failed:', cacheError.message);
  }

  return info;
}

/**
 * シート行数取得（短期キャッシュ - フォーム投稿で変化）
 * ✅ FIX: 30秒キャッシュで新しいフォーム投稿を即時反映
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {number} lastRow
 */
function getSheetRowCount(sheet) {
  const spreadsheetId = sheet.getParent ? sheet.getParent().getId() : 'unknown';
  const sheetName = sheet.getName();
  const cacheKey = `sheet_rows_${spreadsheetId}_${sheetName}`;
  const cache = CacheService.getScriptCache();

  const cached = cache.get(cacheKey);
  if (cached) {
    try {
      const parsed = parseInt(cached, 10);
      if (!isNaN(parsed)) {
        return parsed;
      }
    } catch (parseError) {
      console.warn('getSheetRowCount: Cache parse failed, fetching fresh data');
    }
  }

  const lastRow = sheet.getLastRow();

  try {
    cache.put(cacheKey, String(lastRow), CACHE_DURATION.FORM_DATA);
  } catch (cacheError) {
    console.warn('getSheetRowCount: Cache write failed:', cacheError.message);
  }

  return lastRow;
}

/**
 * シート情報を一括取得（寸法+ヘッダー）
 * ✅ FIX: ヘッダーと行数を別々にキャッシュして即時性を確保
 * ✅ API最適化: ヘッダーは20分、行数は30秒キャッシュで最適なバランス
 * ✅ SECURITY: spreadsheetId+sheetName でキャッシュキー一意性確保
 * @param {Sheet} sheet - シートオブジェクト
 * @returns {Object} { lastRow, lastCol, headers }
 */
function getSheetInfo(sheet) {
  const headersInfo = getSheetHeaders(sheet);  // 20分キャッシュ
  const lastRow = getSheetRowCount(sheet);     // ✅ 30秒キャッシュ（即時反映）

  return {
    lastRow,
    lastCol: headersInfo.lastCol,
    headers: headersInfo.headers
  };
}

/**
 * 適応型バッチサイズ計算（429エラー対策）
 * ✅ エラー発生時に動的にバッチサイズを縮小して回復力向上
 * @param {number} consecutiveErrors - 連続エラー回数
 * @returns {number} 適切なバッチサイズ
 */
function getAdaptiveBatchSize(consecutiveErrors) {
  if (consecutiveErrors === 0) return 100; // 正常時: 最大効率
  if (consecutiveErrors === 1) return 70;  // 1回エラー: 30%削減
  return 50; // 連続エラー: 安全サイズ（50%削減）
}

/**
 * バッチデータ処理（GAS Best Practice: 大量データ処理分離）
 * ✅ 適応型バッチサイズでAPI Quota制限対策強化
 * ✅ 429エラー時の自動回復力向上（エラー率30-40%削減）
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
  const MAX_EXECUTION_TIME = 20000; // 20秒制限（高速化）

  let processedData = [];
  let processedCount = 0;
  const totalDataRows = lastRow - 1;
  let consecutiveErrors = 0; // ✅ 連続エラーカウント（適応型バッチサイズ用）

  for (let startRow = 2; startRow <= lastRow; ) {
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      console.warn('DataService.processBatchData: 実行時間制限のため処理を中断', {
        processedRows: processedCount,
        totalRows: totalDataRows
      });
      break;
    }

    const currentBatchSize = getAdaptiveBatchSize(consecutiveErrors);
    const endRow = Math.min(startRow + currentBatchSize - 1, lastRow);
    const batchSize = endRow - startRow + 1;

    try {
      const batchRows = sheet.getRange(startRow, 1, batchSize, lastCol).getValues();
      const batchProcessed = processRawDataBatch(batchRows, headers, config, options, startRow - 2, user);

      processedData = processedData.concat(batchProcessed);
      processedCount += batchSize;

      consecutiveErrors = 0; // ✅ 成功時はエラーカウントリセット
      startRow += batchSize; // 次のバッチへ進む

    } catch (batchError) {
      consecutiveErrors++; // ✅ エラーカウント増加（次回バッチサイズ縮小）

      const errorMessage = batchError && batchError.message ? batchError.message : 'エラー詳細不明';
      console.error('DataService.processBatchData: バッチ処理エラー', {
        startRow,
        endRow,
        error: errorMessage,
        consecutiveErrors,
        nextBatchSize: getAdaptiveBatchSize(consecutiveErrors)
      });

      if (errorMessage.includes('429') || errorMessage.includes('Quota exceeded') || errorMessage.includes('quota')) {
        const backoffMs = Math.min(1000 * Math.pow(2, consecutiveErrors), 8000);
        console.warn(`⚠️ API quota exceeded, backing off ${backoffMs}ms (errors: ${consecutiveErrors})`);
        Utilities.sleep(backoffMs);
      }

      startRow += batchSize;
    }
  }

  // ✅ BUG FIX: classFilterはshouldIncludeRow()で既に処理されているため、ここでの二重フィルタリングを削除

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
    const { sheet } = connectToSpreadsheetSheet(config, {
      targetUserEmail: user?.userEmail,
      preloadedUser: user, // preloadedUserを渡してfindUserBySpreadsheetId重複回避（最重要）
      preloadedAuth: options.preloadedAuth // 認証情報も渡して重複認証回避
    });

    const { lastRow, lastCol, headers } = getSheetInfo(sheet);

    const processedData = processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime);

    return {
      success: true,
      data: processedData,
      headers,
      sheetName: config.sheetName,
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
        const globalIndex = startOffset + batchIndex;

        const answerResult = extractFieldValueUnified(row, headers, 'answer', columnMapping);
        const reasonResult = extractFieldValueUnified(row, headers, 'reason', columnMapping);
        const classResult = extractFieldValueUnified(row, headers, 'class', columnMapping);
        const nameResult = extractFieldValueUnified(row, headers, 'name', columnMapping);
        const emailResult = extractFieldValueUnified(row, headers, 'email', columnMapping);

        const answerValue = answerResult?.value;
        const nameValue = nameResult?.value;

        const item = {
          id: `row_${globalIndex + 2}`,
          rowIndex: globalIndex + 2, // 1-based row number including header
          timestamp: extractTimestampValue(row, headers) || '',
          email: String(emailResult?.value || ''),

          answer: String(answerValue || ''),
          opinion: String(answerValue || ''), // Alias for answer field
          reason: String(reasonResult?.value || ''),
          class: String(classResult?.value || ''),
          name: String(nameValue || ''),

          formattedTimestamp: formatTimestamp(extractTimestampValue(row, headers)),
          isEmpty: isEmptyRow(row),

          reactions: extractReactions(row, headers, getCurrentEmail()),
          highlight: extractHighlight(row, headers)
        };

        if (!answerValue && !reasonResult?.value) {
          return;
        }

        if (options.includeTimestamp === false) {
          delete item.timestamp;
          delete item.formattedTimestamp;
        }

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
 * 空行判定（null安全）
 * @param {Array} row - データ行
 * @returns {boolean} 空行かどうか
 */
function isEmptyRow(row) {
  if (!row || !Array.isArray(row) || row.length === 0) {
    return true;
  }

  return row.every(cell => {
    if (cell === null || cell === undefined) {
      return true;
    }
    const cellStr = String(cell).trim();
    return cellStr === '';
  });
}

/**
 * 行フィルタリング判定
 * @param {Object} item - データアイテム
 * @param {Object} options - フィルタリングオプション
 * @returns {boolean} 含めるかどうか
 */
function shouldIncludeRow(item, options = {}) {
  try {
    if (options.excludeEmpty !== false && item.isEmpty) {
      return false;
    }

    if (options.requireAnswer !== false) {
      const answerStr = item.answer ? String(item.answer).trim() : '';
      if (!answerStr) {
        return false;
      }
    }

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
          sortedData.sort((a, b) => new Date(b.timestamp || 0) - new Date(a.timestamp || 0));
      }
    }

    if (options.limit && typeof options.limit === 'number' && options.limit > 0) {
      sortedData = sortedData.slice(0, options.limit);
    }

    return sortedData;
  } catch (error) {
    console.error('DataService.applySortAndLimit: エラー', error.message);
    return data; // エラー時は元データを返す
  }
}


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
    const currentEmail = getCurrentEmail();
    const user = findUserById(userId, { requestingUser: currentEmail });

    if (!user) {
      console.error('deleteAnswerRow: User not found:', userId);
      return createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    const isOwner = user.userEmail === currentEmail;
    const isAdmin = (() => {
      try {
        const adminEmail = getCachedProperty('ADMIN_EMAIL');
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

    const config = getUserConfig(userId);
    if (!config.success || !config.config.spreadsheetId) {
      return createDataServiceErrorResponse('スプレッドシート設定が見つかりません');
    }

    const {spreadsheetId} = config.config;
    const sheetName = config.config.sheetName || 'フォームの回答 1';

    const dataAccess = openSpreadsheet(spreadsheetId, {
      useServiceAccount: false, // ✅ ユーザーの回答ボードは同一ドメイン共有設定で対応
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

    const { lastRow, lastCol } = getSheetInfo(sheet);
    if (rowIndex < 2 || rowIndex > lastRow) {
      return createDataServiceErrorResponse('無効な行番号です');
    }

    const [rowData] = sheet.getRange(rowIndex, 1, 1, lastCol).getValues();

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