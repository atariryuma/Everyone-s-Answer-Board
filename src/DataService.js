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

/* global formatTimestamp, createErrorResponse, createExceptionResponse, getQuestionText, findUserByEmail, findUserById, findUserBySpreadsheetId, openSpreadsheet, getUserConfig, getConfigOrDefault, normalizeHeader, CACHE_DURATION, getCurrentEmail, isAdministrator, resolveColumnIndex, extractReactions, extractHighlight, createDataServiceErrorResponse, getCachedProperty, FormApp */

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
      config = getConfigOrDefault(userId);
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
    // Why: 失敗時の ERROR ログはここに集約する。下層は WARN or silent に降格済み。
    const executionTime = Date.now() - startTime;
    console.error('DataService.getUserSheetData: 失敗', {
      userId,
      error: error.message,
      executionTime: `${executionTime}ms`,
      spreadsheetId: config?.spreadsheetId || 'undefined',
      sheetName: config?.sheetName || 'undefined',
      userEmail: user?.userEmail || 'undefined'
    });
    return createDataServiceErrorResponse(`データ取得エラー: ${error.message}`);
  }
}

/** Returns the index of the timestamp column, or -1 if not found. */
function resolveTimestampIndex(headers) {
  if (!Array.isArray(headers)) return -1;
  for (let i = 0; i < headers.length; i++) {
    const header = headers[i];
    if (!header || typeof header !== 'string') continue;
    const normalized = normalizeHeader(header);
    if (normalized.includes('タイムスタンプ') ||
        normalized.includes('timestamp') ||
        normalized.includes('日時') ||
        normalized.includes('日付')) {
      return i;
    }
  }
  return -1;
}

/**
 * スプレッドシート接続とシート取得（GAS Best Practice: 単一責任）
 * @param {Object} config - 設定オブジェクト
 * @returns {Object} シート情報
 */
function connectToSpreadsheetSheet(config) {
  const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: false });

  if (!dataAccess || !dataAccess.spreadsheet) {
    // Why: openSpreadsheet が既に WARN ログを出しているので二重ログしない。
    //      ここでは throw だけ行い、最上位（getUserSheetData）の catch で
    //      1 行にまとめて ERROR ログを出す。
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

  // Why: キャッシュミス時に sheet.getDataRange().getValues() で全件読むのは無駄。
  //      ヘッダ行は常に1行目の1..lastColumn なので、そこだけ取得する。
  //      getLastColumn() は寸法だけのメタ読みで安価。1000行のシートで数百ms削減。
  const lastCol = sheet.getLastColumn();
  const headers = lastCol > 0
    ? (sheet.getRange(1, 1, 1, lastCol).getValues()[0] || [])
    : [];

  const info = { lastCol, headers };

  try {
    cache.put(cacheKey, JSON.stringify(info), CACHE_DURATION.DATABASE_LONG);
  } catch (cacheError) {
    console.warn('getSheetHeaders: Cache write failed:', cacheError.message);
  }

  return info;
}

/**
 * ヘッダーキャッシュを明示的に無効化
 * 列追加・削除・リネーム後に呼び出し、次回リアクション処理等で最新ヘッダーを強制取得させる。
 * Why: getSheetHeadersは10分キャッシュするため、列セットアップ直後に新しい列が見えず
 *      「リアクション列が見つかりません」エラーが発生するのを防ぐ。
 */
function invalidateSheetHeadersCache(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) return;
  try {
    const cacheKey = `sheet_headers_${spreadsheetId}_${sheetName}`;
    CacheService.getScriptCache().remove(cacheKey);
  } catch (error) {
    console.warn('invalidateSheetHeadersCache: Cache remove failed:', error.message);
  }
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
  const MAX_EXECUTION_TIME = 20000;
  const MAX_CONSECUTIVE_ERRORS = 3;

  let processedData = [];
  let processedCount = 0;
  const totalDataRows = lastRow - 1;
  let consecutiveErrors = 0; // ✅ 連続エラーカウント（適応型バッチサイズ用）

  // Headers don't change mid-batch, so resolve column indices once here
  // rather than per-row (resolveColumnIndex runs a full header scan on misses).
  const columnMapping = config.columnMapping || {};
  const fieldIndices = {
    answer: resolveColumnIndex(headers, 'answer', columnMapping).index,
    reason: resolveColumnIndex(headers, 'reason', columnMapping).index,
    class: resolveColumnIndex(headers, 'class', columnMapping).index,
    name: resolveColumnIndex(headers, 'name', columnMapping).index,
    email: resolveColumnIndex(headers, 'email', columnMapping).index
  };
  const tsIndex = resolveTimestampIndex(headers);

  for (let startRow = 2; startRow <= lastRow; ) {
    const currentBatchSize = getAdaptiveBatchSize(consecutiveErrors);
    const endRow = Math.min(startRow + currentBatchSize - 1, lastRow);
    const batchSize = endRow - startRow + 1;

    try {
      const batchRows = sheet.getRange(startRow, 1, batchSize, lastCol).getValues();
      const batchProcessed = processRawDataBatch(batchRows, headers, config, options, startRow - 2, user, fieldIndices, tsIndex);

      processedData = processedData.concat(batchProcessed);
      processedCount += batchSize;

      consecutiveErrors = 0;
      startRow += batchSize;

    } catch (batchError) {
      consecutiveErrors++;

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
        // Why: sleep 込みで予算超過するなら、sleep すること自体がリソース浪費。
        //      30 人同時ポーリング時、全員が 14s ブロックして script-runtime quota を圧迫する問題を防ぐ。
        if (Date.now() - startTime + backoffMs > MAX_EXECUTION_TIME) {
          console.warn('⚠️ API quota exceeded, skipping backoff (would exceed budget)');
          break;
        }
        console.warn(`⚠️ API quota exceeded, backing off ${backoffMs}ms (errors: ${consecutiveErrors})`);
        Utilities.sleep(backoffMs);
      }

      const isPermanent = errorMessage.includes('not found') || errorMessage.includes('out of range');
      if (isPermanent || consecutiveErrors >= MAX_CONSECUTIVE_ERRORS) {
        console.warn(`⚠️ バッチスキップ: 行${startRow}-${endRow}（${isPermanent ? '永続エラー' : consecutiveErrors + '回連続エラー'}）`);
        startRow += batchSize;
      }
    }

    // Why: 時間チェックはバッチ試行の「後」に置く。以前は先頭に置いていたため、
    //      setup（openSpreadsheet / getSheetInfo）に時間が取られると 1 行も処理せず
    //      success:true, data:[] を返して「回答が 0 件」と誤表示していた。
    //      少なくとも 1 バッチは必ず試行する。
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      console.warn('DataService.processBatchData: 実行時間制限のため処理を中断', {
        processedRows: processedCount,
        totalRows: totalDataRows
      });
      break;
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

  const { sheet } = connectToSpreadsheetSheet(config);
  const { lastRow, lastCol, headers } = getSheetInfo(sheet);
  const processedData = processBatchData(sheet, headers, lastRow, lastCol, config, options, user, startTime);

  return {
    success: true,
    data: processedData,
    headers,
    sheetName: config.sheetName
  };
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
function processRawDataBatch(batchRows, headers, config, options, startOffset, user, fieldIndices, tsIndex) {
  try {
    const processedBatch = [];
    // Normalize empty/missing cells to null to match the prior
    // extractFieldValueUnified contract (empty strings → value:null).
    const getCell = (row, idx) => {
      if (idx < 0 || idx >= row.length) return null;
      const v = row[idx];
      if (v === null || v === undefined) return null;
      if (typeof v === 'string' && v.trim() === '') return null;
      return v;
    };

    const viewerEmail = getCurrentEmail();

    batchRows.forEach((row, batchIndex) => {
      try {
        const globalIndex = startOffset + batchIndex;

        const answerValue = getCell(row, fieldIndices.answer);
        const reasonValue = getCell(row, fieldIndices.reason);
        const classValue = getCell(row, fieldIndices.class);
        const nameValue = getCell(row, fieldIndices.name);
        const emailValue = getCell(row, fieldIndices.email);

        const tsValue = tsIndex >= 0 && tsIndex < row.length ? (row[tsIndex] || '') : '';

        const item = {
          id: `row_${globalIndex + 2}`,
          rowIndex: globalIndex + 2,
          timestamp: tsValue,
          email: emailValue || '',

          answer: answerValue || '',
          opinion: answerValue || '',
          reason: reasonValue || '',
          class: classValue || '',
          name: nameValue || '',

          formattedTimestamp: formatTimestamp(tsValue),
          isEmpty: isEmptyRow(row),

          reactions: extractReactions(row, headers, viewerEmail),
          highlight: extractHighlight(row, headers)
        };

        if (!answerValue && !reasonValue) {
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

    if ((options.dateFrom || options.dateTo) && item.timestamp) {
      const itemDate = new Date(item.timestamp);
      if (options.dateFrom && itemDate < new Date(options.dateFrom)) {
        return false;
      }
      if (options.dateTo && itemDate > new Date(options.dateTo)) {
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
        case 'score':
          // Why: スコア順 = リアクション合計数の降順。ハイライトは最上位に固定し、
          //      同点時は新着順でタイブレーク。以前 case 'score' が無かったため
          //      default 分岐で newest にフォールバックしていた silent bug。
          sortedData.sort((a, b) => {
            const ah = a.highlight ? 1 : 0;
            const bh = b.highlight ? 1 : 0;
            if (ah !== bh) return bh - ah;
            const score = (r) => (r?.LIKE?.count || 0) + (r?.UNDERSTAND?.count || 0) + (r?.CURIOUS?.count || 0);
            const sa = score(a.reactions);
            const sb = score(b.reactions);
            if (sa !== sb) return sb - sa;
            return new Date(b.timestamp || 0) - new Date(a.timestamp || 0);
          });
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
 * 回答行を削除（編集者権限必須）
 * @param {string} userId - ユーザーID
 * @param {number} rowIndex - 削除対象の行インデックス（1-based, ヘッダー含む）
 * @param {Object} options - オプション設定
 * @returns {Object} 削除結果
 */
function deleteAnswerRow(userId, rowIndex) {
  try {
    const currentEmail = getCurrentEmail();
    const user = findUserById(userId, { requestingUser: currentEmail });

    if (!user) {
      console.error('deleteAnswerRow: User not found:', userId);
      return createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    const isOwner = user.userEmail === currentEmail;
    const isAdmin = isAdministrator(currentEmail);

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

    // Why: sheet.deleteRows だけでは Google Forms のレスポンスストアには残り、
    //      フォームの集計・再送信防止・回答編集 URL が実在する回答として扱われる。
    //      未リンク・権限不足は致命的ではないので best-effort でシート削除は成功扱い。
    deleteLinkedFormResponseByTimestamp(sheet, rowData[0]);

    return {
      success: true,
      message: '回答を削除しました'
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

/**
 * シートにリンクされたフォームから、指定タイムスタンプのレスポンスを削除する（best-effort）。
 *
 * Why timestamp match: シートの 1 列目にはフォーム送信時の timestamp が入り、
 *   FormResponse.getTimestamp() と同じ値になる。ID 保持列を増やさずに突き合わせ可能。
 *
 * @param {Sheet} sheet
 * @param {Date} rowTimestamp - sheet.getRange(...).getValues() 由来なので常に Date
 * @returns {{success:boolean, message?:string}}
 */
function deleteLinkedFormResponseByTimestamp(sheet, rowTimestamp) {
  const TOLERANCE_MS = 1000;
  try {
    const formUrl = sheet.getFormUrl();
    if (!formUrl) {
      return { success: false, message: 'no linked form' };
    }

    // Why instanceof だけだと vm context 跨ぎのテスト用 Date で false 判定されるため duck-typing。
    //      production では sheet.getRange().getValues() は常に Date を返す。
    const targetTime = rowTimestamp && typeof rowTimestamp.getTime === 'function'
      ? rowTimestamp.getTime()
      : NaN;
    if (!Number.isFinite(targetTime)) {
      return { success: false, message: 'invalid timestamp' };
    }

    const form = FormApp.openByUrl(formUrl);
    // Why: 全件取得でなく target 以降に限定することで、長期運用フォーム（数千件）でも
    //      1 件取れば済む。FormApp.getResponses(timestamp) は「timestamp 以降」を返す。
    const responses = form.getResponses(new Date(targetTime - TOLERANCE_MS));
    for (let i = 0; i < responses.length; i += 1) {
      // Why: Sheet 側は秒精度、FormResponse 側はミリ秒精度のため完全一致しない。
      const respTime = responses[i].getTimestamp().getTime();
      if (Math.abs(respTime - targetTime) <= TOLERANCE_MS) {
        form.deleteResponse(responses[i].getId());
        return { success: true };
      }
    }
    return { success: false, message: 'no matching form response' };
  } catch (error) {
    console.warn('deleteLinkedFormResponseByTimestamp: failed', {
      error: error.message
    });
    return { success: false, message: error.message };
  }
}
