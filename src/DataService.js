/**
 * @fileoverview DataService - スプレッドシートデータの取得・フィルタ・整形、
 *   シート寸法/ヘッダー取得（キャッシュ付き）、適応型バッチ読込。
 */

/* global formatTimestamp, getQuestionText, findUserById, openSpreadsheet, getUserConfig, getConfigOrDefault, normalizeHeader, CACHE_DURATION, getCurrentEmail, isAdministrator, resolveColumnIndex, extractReactions, extractHighlight, createDataServiceErrorResponse, logError_, sameEmail_ */

/**
 * ユーザーのスプレッドシートデータ取得
 * @param {string} userId - ユーザーID
 * @param {Object} [options] - 取得オプション
 * @param {Object} [preloadedUser] - 既に取得済みの user (batch 呼び出しで重複 lookup を防ぐ)
 * @param {Object} [preloadedConfig] - 既に取得済みの config
 * @returns {Object} レスポンスオブジェクト
 */
function getUserSheetData(userId, options = {}, preloadedUser = null, preloadedConfig = null) {
  const startTime = Date.now();

  let user = null;
  let config = null;

  try {

    user = preloadedUser || findUserById(userId, { requestingUser: getCurrentEmail() });
    if (!user) {
      logError_('DataService.getUserSheetData', new Error('ユーザーが見つかりません'), { userId });
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

    const result = fetchSpreadsheetData(config, options);

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
    logError_('DataService.getUserSheetData', error, {
      userId,
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
 * config からスプレッドシートを開いて該当シートを返す。
 * @param {Object} config
 * @returns {{sheet:Sheet, spreadsheet:Spreadsheet}}
 */
function connectToSpreadsheetSheet(config) {
  // accessMode は openSpreadsheet が自動判定 (owner なら openById、 viewer なら SA pool)。
  const dataAccess = openSpreadsheet(config.spreadsheetId, { context: 'connectToSpreadsheetSheet' });

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
 * ヘッダー情報取得（20分キャッシュ。新規投稿の即時反映は行数側 (`getSheetRowCount`) が担う）。
 * @param {Sheet} sheet
 * @returns {{lastCol:number, headers:Array}}
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
 * シート行数取得（30秒キャッシュ — 新規フォーム投稿を即時反映するため短期）。
 * @param {Sheet} sheet
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
 * シート情報を一括取得（寸法 + ヘッダー）。ヘッダーは長期、行数は短期キャッシュ。
 * @param {Sheet} sheet
 * @returns {{lastRow:number, lastCol:number, headers:Array}}
 */
function getSheetInfo(sheet) {
  const headersInfo = getSheetHeaders(sheet);
  const lastRow = getSheetRowCount(sheet);
  return { lastRow, lastCol: headersInfo.lastCol, headers: headersInfo.headers };
}

// 適応型バッチサイズ (429 対策)。連続エラー時に段階的に縮小。
function getAdaptiveBatchSize(consecutiveErrors) {
  if (consecutiveErrors === 0) return 100; // 正常時: 最大効率
  if (consecutiveErrors === 1) return 70;  // 1回エラー: 30%削減
  return 50; // 連続エラー: 安全サイズ（50%削減）
}

/**
 * バッチでシート行を読み出し、フィルタ・整形して返す。
 * 適応型バッチサイズで API quota 制限に追従。
 * @param {Sheet} sheet
 * @param {Array} headers
 * @param {number} lastRow
 * @param {number} lastCol
 * @param {Object} config
 * @param {Object} options
 * @param {Object} user
 * @param {number} startTime
 */
/**
 * @param {Object} ctx - {sheet, headers, lastRow, lastCol, config, options, startTime}
 */
function processBatchData(ctx) {
  const { sheet, headers, lastRow, lastCol, config, options, startTime } = ctx;
  const MAX_EXECUTION_TIME = 20000;
  const MAX_CONSECUTIVE_ERRORS = 3;

  let processedData = [];
  let processedCount = 0;
  const totalDataRows = lastRow - 1;
  let consecutiveErrors = 0;

  // Headers don't change mid-batch, so resolve column indices once here
  // rather than per-row (resolveColumnIndex runs a full header scan on misses).
  const columnMapping = config.columnMapping || {};
  // Why: numericX/Y はパターン検出がない（教師が手動 or detectNumericScaleColumns で選ぶ）。
  //      columnMapping に明示インデックスがあるときだけ取り込む。-1 のとき item は null を入れる。
  const fieldIndices = {
    answer: resolveColumnIndex(headers, 'answer', columnMapping).index,
    reason: resolveColumnIndex(headers, 'reason', columnMapping).index,
    class: resolveColumnIndex(headers, 'class', columnMapping).index,
    name: resolveColumnIndex(headers, 'name', columnMapping).index,
    email: resolveColumnIndex(headers, 'email', columnMapping).index,
    numericX: (typeof columnMapping.numericX === 'number' && columnMapping.numericX >= 0 && columnMapping.numericX < headers.length) ? columnMapping.numericX : -1,
    numericY: (typeof columnMapping.numericY === 'number' && columnMapping.numericY >= 0 && columnMapping.numericY < headers.length) ? columnMapping.numericY : -1
  };
  const tsIndex = resolveTimestampIndex(headers);

  for (let startRow = 2; startRow <= lastRow; ) {
    const currentBatchSize = getAdaptiveBatchSize(consecutiveErrors);
    const endRow = Math.min(startRow + currentBatchSize - 1, lastRow);
    const batchSize = endRow - startRow + 1;

    try {
      const batchRows = sheet.getRange(startRow, 1, batchSize, lastCol).getValues();
      const batchProcessed = processRawDataBatch({
        batchRows, headers, options,
        startOffset: startRow - 2,
        fieldIndices, tsIndex
      });

      // push(...batchProcessed) は in-place で O(n)。concat は新 array を毎回作るので O(n²)。
      for (let i = 0; i < batchProcessed.length; i++) processedData.push(batchProcessed[i]);
      processedCount += batchSize;

      consecutiveErrors = 0;
      startRow += batchSize;

    } catch (batchError) {
      consecutiveErrors++;
      // batchError から message 文字列を取り出す。これを忘れると catch 自身が
      // ReferenceError を投げ、429 backoff / 永続エラースキップが一切動かなくなる。
      const errorMessage = batchError && batchError.message ? batchError.message : String(batchError);
      logError_('DataService.processBatchData', batchError, {
        startRow,
        endRow,
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
    //
    // Why (WARN 抑制): 全行処理済みで時間超過の場合は「中断」ではなく「完走後の時間超過」。
    //      processedCount >= totalDataRows なら break するが WARN は出さない。
    //      旧版は 42/42 完了時にも WARN を吐いて false-positive アラート扱いされていた。
    if (Date.now() - startTime > MAX_EXECUTION_TIME) {
      if (processedCount < totalDataRows) {
        console.warn('DataService.processBatchData: 実行時間制限のため処理を中断', {
          processedRows: processedCount,
          totalRows: totalDataRows
        });
      }
      break;
    }
  }

  // Why: applySortAndLimit は sortBy と limit を独立に適用する。以前はここで
  //      `if (options.sortBy)` でガードしていたため、limit 単体を渡す caller が
  //      silently 無視されていた。limit も sortBy も未指定なら関数内で no-op。
  if (options.sortBy || options.limit) {
    processedData = applySortAndLimit(processedData, {
      sortBy: options.sortBy,
      limit: options.limit
    });
  }

  return processedData;
}

/**
 * config に従ってシートを開き、フィルタ・整形済みのデータ配列を返す。
 * @param {Object} config
 * @param {Object} options
 */
function fetchSpreadsheetData(config, options = {}) {
  const startTime = Date.now();

  const { sheet } = connectToSpreadsheetSheet(config);
  const { lastRow, lastCol, headers } = getSheetInfo(sheet);
  const processedData = processBatchData({ sheet, headers, lastRow, lastCol, config, options, startTime });

  return {
    success: true,
    data: processedData,
    headers,
    sheetName: config.sheetName
  };
}

/**
 * バッチ処理用データ変換（メモリ効率重視）。
 * @param {Object} ctx - {batchRows, headers, options, startOffset, fieldIndices, tsIndex}
 * @returns {Array} 処理済みバッチデータ
 */
function processRawDataBatch(ctx) {
  const { batchRows, headers, options, startOffset, fieldIndices, tsIndex } = ctx;
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

        // Why: numericX/Y は Forms「線形尺度」を想定。整数として正規化、
        //      非数値や空欄は null（描画側で除外判定に使う）。
        const toScaleNumber = (raw) => {
          if (raw === null || raw === undefined || raw === '') return null;
          const n = typeof raw === 'number' ? raw : Number(String(raw).trim());
          return Number.isFinite(n) ? n : null;
        };
        const numericXValue = fieldIndices.numericX >= 0 ? toScaleNumber(row[fieldIndices.numericX]) : null;
        const numericYValue = fieldIndices.numericY >= 0 ? toScaleNumber(row[fieldIndices.numericY]) : null;

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
          numericX: numericXValue,
          numericY: numericYValue,

          formattedTimestamp: formatTimestamp(tsValue),
          isEmpty: isEmptyRow(row),

          reactions: extractReactions(row, headers, viewerEmail),
          highlight: extractHighlight(row, headers)
        };

        // Why: answer/reason は board モードの必須コンテンツ。両方空ならスキップ。
        //      ただし numericX/Y が入っているならビジュアル化対象として残す
        //      （Forms 線形尺度のみで自由記述なし、もありうる）。
        const hasNumeric = numericXValue !== null || numericYValue !== null;
        if (!answerValue && !reasonValue && !hasNumeric) {
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
    logError_('DataService.processRawDataBatch', error);
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

// 行フィルタリング判定
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
        case 'score': {
          // Why: ハイライト＞リアクション合計数＞新着 の優先順位。スコアを事前計算
          //      しないと比較のたびに再計算され O(n log n) 回走ってしまう。
          const scored = sortedData.map((item) => {
            const r = item.reactions;
            return {
              item,
              highlight: item.highlight ? 1 : 0,
              score: (r?.LIKE?.count || 0) + (r?.UNDERSTAND?.count || 0) + (r?.CURIOUS?.count || 0),
              ts: new Date(item.timestamp || 0).getTime(),
            };
          });
          scored.sort((a, b) => (b.highlight - a.highlight) || (b.score - a.score) || (b.ts - a.ts));
          sortedData = scored.map((x) => x.item);
          break;
        }
        default:
          sortedData.sort((a, b) => new Date(b.timestamp || 0) - new Date(a.timestamp || 0));
      }
    }

    if (options.limit && typeof options.limit === 'number' && options.limit > 0) {
      sortedData = sortedData.slice(0, options.limit);
    }

    return sortedData;
  } catch (error) {
    logError_('DataService.applySortAndLimit', error);
    return data; // エラー時は元データを返す
  }
}

// timestamp 値を比較可能な正規形に揃える。 sheet cell は Date object、 client から戻る
//   値は google.script.run の JSON 化で ISO 文字列になるため、 両者を epoch ms 文字列に
//   正規化して同一瞬間かを判定する。 日付として解釈できない値はそのまま trim 比較。
function toComparableTimestamp_(v) {
  if (v === null || v === undefined || v === '') return '';
  if (v instanceof Date) return String(v.getTime());
  const d = new Date(v);
  if (!isNaN(d.getTime())) return String(d.getTime());
  return String(v).trim();
}

/**
 * 回答行を削除（編集者権限必須）
 * @param {string} userId - ユーザーID
 * @param {number} rowIndex - 削除対象の行インデックス（1-based, ヘッダー含む）
 * @param {string} [expectedTimestamp] - 削除対象として client が想定する行の timestamp。
 *   指定時、 削除直前に実シートの当該行 timestamp (列A) と照合し、 不一致なら abort する。
 *   deleteRows は後続行を 1 行ずつ繰り上げるため、 古い snapshot に基づく rowIndex で
 *   別の生徒の回答を誤削除する race を防ぐ (identity verification)。
 * @returns {Object} 削除結果
 */
function deleteAnswerRow(userId, rowIndex, expectedTimestamp) {
  try {
    const currentEmail = getCurrentEmail();
    const user = findUserById(userId, { requestingUser: currentEmail });

    if (!user) {
      logError_('deleteAnswerRow', new Error('User not found'), { userId });
      return createDataServiceErrorResponse('ユーザーが見つかりません');
    }

    const isOwner = sameEmail_(user.userEmail, currentEmail);
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

    // accessMode は openSpreadsheet が自動判定 (owner = openById, viewer/admin = SA pool)。
    const dataAccess = openSpreadsheet(spreadsheetId, {
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

    // bounds は **fresh** な getLastRow() で取る。 getSheetInfo の lastRow は
    //   getSheetRowCount の 30s cache 由来で、 新規投稿直後は stale になりうるため。
    const { lastCol } = getSheetInfo(sheet);
    const lastRow = sheet.getLastRow();
    if (rowIndex < 2 || rowIndex > lastRow) {
      return createDataServiceErrorResponse('無効な行番号です');
    }

    const [rowData] = sheet.getRange(rowIndex, 1, 1, lastCol).getValues();

    // identity verification: client が想定した行と実シートの行が一致するか timestamp (列A) で確認。
    //   不一致 = snapshot 以降に行が削除/挿入され index がずれた = 別の回答を指している。
    //   誤削除を防ぐため abort し、 画面更新を促す。
    if (expectedTimestamp !== undefined && expectedTimestamp !== null && expectedTimestamp !== '') {
      const actual = toComparableTimestamp_(rowData[0]);
      const expected = toComparableTimestamp_(expectedTimestamp);
      if (actual && expected && actual !== expected) {
        console.warn('deleteAnswerRow: timestamp mismatch (row shifted?)', { userId, rowIndex });
        return createDataServiceErrorResponse('対象の回答が変更されています。画面を更新してからもう一度お試しください。');
      }
    }

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
    logError_('deleteAnswerRow', error, { userId, rowIndex });
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
