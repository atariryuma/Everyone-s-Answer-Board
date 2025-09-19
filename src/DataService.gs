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

/* global ServiceFactory, formatTimestamp, createErrorResponse, createExceptionResponse, getSheetData, columnAnalysis, getQuestionText, Data, Config, getConfigSafe */

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
    if (!user) {
      console.error('DataService.getUserSheetData: ユーザーが見つかりません', { userId });
      return { success: false, message: 'ユーザーが見つかりません', data: [], headers: [], sheetName: '' };
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId) {
      console.warn('DataService.getUserSheetData: スプレッドシートIDが設定されていません', { userId });
      // Direct return format like admin panel getSheetList
      return { success: false, message: 'スプレッドシートが設定されていません', data: [], headers: [], sheetName: '' };
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
      return {
        ...result,
        header: getQuestionText(config) || result.sheetName || '回答一覧',
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
function fetchSpreadsheetData(config, options = {}, user = null) {
  const startTime = Date.now();
  const MAX_EXECUTION_TIME = 180000; // 3分制限（安全マージン拡大）
  const MAX_BATCH_SIZE = 200; // バッチサイズ削減（メモリ制限対応）

  try {
    // スプレッドシート取得
    const dataAccess = Data.open(config.spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      throw new Error(`シート '${config.sheetName}' が見つかりません`);
    }

    // データ範囲取得
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1) {
      // ✅ シンプル形式で返却
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
        const batchProcessed = processRawDataBatch(batchRows, headers, config, options, startRow - 2, user);

        processedData = processedData.concat(batchProcessed);
        processedCount += batchSize;


        // API制限対策: 100行毎に短い休憩
        if (processedCount % 1000 === 0) {
          Utilities.sleep(100); // 0.1秒休憩
        }

      } catch (batchError) {
        console.error('DataService.fetchSpreadsheetData: バッチ処理エラー', {
          operation: 'fetchSpreadsheetData',
          batchIndex: Math.floor(startRow / options.batchSize),
          startRow,
          endRow,
          totalRows: totalDataRows,
          sheetName: sheet?.getName() || 'unknown',
          error: batchError.message,
          stack: batchError.stack
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

    // ✅ フロントエンド期待形式で直接返却
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
function processRawDataBatch(batchRows, headers, config, options = {}, startOffset = 0, user = null) {
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
          reactions: extractReactions(row, headers, user.userEmail),
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
 * @returns {Array} 処理済みデータ
 */
function processRawData(dataRows, headers, config, options = {}, user = null) {
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
          reactions: extractReactions(row, headers, user.userEmail),
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
    const dataAccess = Data.open(config.spreadsheetId);
    const {spreadsheet} = dataAccess;
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
 * 🎯 排他的リアクション処理システム - CLAUDE.md準拠Zero-Dependency実装
 *
 * 仕様:
 * - 排他的リアクション: ユーザーは1つのリアクションのみ選択可能
 * - 同じリアクションをクリック: トグル（削除）
 * - 異なるリアクションをクリック: 古いリアクションを削除し、新しいリアクションを追加
 * - ユーザーベース管理: メールアドレスで重複防止
 * - カウントベース表示: フロントエンド向けに適切に変換
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類 (LIKE, UNDERSTAND, CURIOUS)
 * @param {string} userEmail - ユーザーメール
 * @returns {Object} 処理結果 {success, status, message, action, reactions, userReaction, newValue}
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, userEmail) {
  // 🚀 Zero-dependency: ServiceFactory経由で初期化
  try {
    if (!validateReaction(spreadsheetId, sheetName, rowIndex, reactionKey)) {
      throw new Error('無効なリアクションパラメータ');
    }

    if (!userEmail) {
      throw new Error('ユーザー情報が必要です');
    }

    const dataAccess = Data.open(spreadsheetId);
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // Get all reaction columns for this row to implement exclusive reactions
    const reactionColumns = {
      'LIKE': getOrCreateReactionColumn(sheet, 'LIKE'),
      'UNDERSTAND': getOrCreateReactionColumn(sheet, 'UNDERSTAND'),
      'CURIOUS': getOrCreateReactionColumn(sheet, 'CURIOUS')
    };

    // Get current reaction states for all reaction types
    const currentReactions = {};
    const allReactionsData = {};
    let userCurrentReaction = null;

    Object.keys(reactionColumns).forEach(key => {
      const col = reactionColumns[key];
      const cellValue = sheet.getRange(rowIndex, col).getValue() || '';
      const reactionUsers = parseReactionUsers(cellValue);
      currentReactions[key] = reactionUsers;
      allReactionsData[key] = {
        count: reactionUsers.length,
        reacted: reactionUsers.includes(userEmail)
      };

      if (reactionUsers.includes(userEmail)) {
        userCurrentReaction = key;
      }
    });

    // Apply reaction rules with simplified logic
    let action = 'added';
    let newUserReaction = null;

    if (userCurrentReaction === reactionKey) {
      // User clicking same reaction -> remove (toggle)
      currentReactions[reactionKey] = currentReactions[reactionKey].filter(u => u !== userEmail);
      action = 'removed';
      newUserReaction = null;
    } else {
      // User clicking different reaction -> remove old (if any), add new
      if (userCurrentReaction) {
        currentReactions[userCurrentReaction] = currentReactions[userCurrentReaction].filter(u => u !== userEmail);
        action = 'changed';
      }
      currentReactions[reactionKey].push(userEmail);
      newUserReaction = reactionKey;
    }

    // Update all reaction columns
    Object.keys(reactionColumns).forEach(key => {
      const col = reactionColumns[key];
      const users = currentReactions[key];
      const cellValue = serializeReactionUsers(users);
      sheet.getRange(rowIndex, col).setValue(cellValue);

      // Update response data
      allReactionsData[key] = {
        count: users.length,
        reacted: users.includes(userEmail)
      };
    });

    console.info('🎯 排他的リアクション処理完了 - CLAUDE.md準拠', {
      spreadsheetId: `${spreadsheetId.substring(0, 10)}***`,
      sheetName,
      rowIndex,
      reactionKey,
      userEmail: `${userEmail.substring(0, 5)  }***`,
      action,
      exclusive: true,  // 排他的リアクションであることを明示
      previousReaction: userCurrentReaction,
      newReaction: newUserReaction,
      reactionCounts: Object.keys(allReactionsData).map(key => `${key}:${allReactionsData[key].count}`).join(', ')
    });

    return {
      success: true,
      status: 'success',
      message: `リアクションを${action === 'removed' ? '削除' : action === 'changed' ? '変更' : '追加'}しました`,
      action,
      reactions: allReactionsData,
      userReaction: newUserReaction,
      newValue: allReactionsData[reactionKey]?.count || 0  // For backwards compatibility
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

/**
 * リアクションユーザー配列をパース
 * @param {string} cellValue - セル値
 * @returns {Array<string>} ユーザーメール配列
 */
function parseReactionUsers(cellValue) {
  if (!cellValue || typeof cellValue !== 'string') {
    return [];
  }

  const trimmed = cellValue.trim();
  if (!trimmed) {
    return [];
  }

  // Split by delimiter and filter out empty strings
  return trimmed.split('|').filter(email => email.trim().length > 0);
}

/**
 * ユーザー配列をセル用文字列にシリアライズ
 * @param {Array<string>} users - ユーザーメール配列
 * @returns {string} セル格納用文字列
 */
function serializeReactionUsers(users) {
  if (!Array.isArray(users) || users.length === 0) {
    return '';
  }

  // Filter out empty emails and join with delimiter
  const validEmails = users.filter(email => email && email.trim().length > 0);
  return validEmails.join('|');
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
function extractReactions(row, headers, userEmail = null) {
  try {
    const reactions = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };

    // リアクション列を探してメール配列を抽出
    headers.forEach((header, index) => {
      const headerStr = String(header).toUpperCase();
      let reactionType = null;

      // ヘッダー名からリアクションタイプを特定
      if (headerStr.includes('UNDERSTAND') || headerStr.includes('理解')) {
        reactionType = 'UNDERSTAND';
      } else if (headerStr.includes('LIKE') || headerStr.includes('いいね')) {
        reactionType = 'LIKE';
      } else if (headerStr.includes('CURIOUS') || headerStr.includes('気になる')) {
        reactionType = 'CURIOUS';
      }

      if (reactionType) {
        const cellValue = row[index] || '';
        const reactionUsers = parseReactionUsers(cellValue);
        reactions[reactionType] = {
          count: reactionUsers.length,
          reacted: userEmail ? reactionUsers.includes(userEmail) : false
        };
      }
    });

    return reactions;
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
function getSpreadsheetList(options = {}) {
  // 🚀 Zero-dependency: ServiceFactory経由で初期化
  const started = Date.now();
  try {
    // ✅ GAS Best Practice: 直接API呼び出し（依存除去）
    const session = ServiceFactory.getSession();
    const currentUser = session.email;

    // オプション設定
    const {
      adminMode = false,
      maxCount = adminMode ? 20 : 25,
      includeSize = adminMode,
      includeTimestamp = true
    } = options;

    // DriveApp直接使用（効率重視）
    const files = DriveApp.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');

    // 権限テスト（必要最小限）
    let driveAccessOk = true;
    try {
      const testFiles = DriveApp.getFiles();
      driveAccessOk = testFiles.hasNext();
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

    while (files.hasNext() && count < maxCount) {
      try {
        const file = files.next();
        const fileData = {
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          lastUpdated: file.getLastUpdated()
        };

        // 管理者モード時は追加情報を含める
        if (includeSize) {
          fileData.size = file.getSize() || 0;
        }

        spreadsheets.push(fileData);
        count++;
      } catch (fileError) {
        console.warn('DataService.getSpreadsheetList: ファイル処理スキップ', fileError.message);
        continue; // エラー時はスキップして継続
      }
    }


    // ✅ シンプル形式に最適化
    const response = {
      success: true,
      spreadsheets,
      executionTime: `${Date.now() - started}ms`
    };

    // レスポンスサイズ監視（GAS制限対応）
    const responseSize = JSON.stringify(response).length;
    const responseSizeKB = Math.round(responseSize / 1024 * 100) / 100;


    // ✅ 構造チェック
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
 * 🎯 AI列分析実装 - connectToSheetInternalに統合
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} sheetName シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumnStructure(spreadsheetId, sheetName) {
  const started = Date.now();
  try {
    const paramValidation = validateSheetParams(spreadsheetId, sheetName);
    if (!paramValidation.isValid) {
      return paramValidation.errorResponse;
    }

    const connectionResult = connectToSheetInternal(spreadsheetId, sheetName);
    if (!connectionResult.success) {
      return connectionResult.errorResponse;
    }

    return {
      success: true,
      headers: connectionResult.headers,
      mapping: connectionResult.mapping,        // フロントエンド期待形式
      confidence: connectionResult.confidence,  // 分離
      executionTime: `${Date.now() - started}ms`
    };

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
      mapping: {},           // フロントエンド期待形式
      confidence: {},        // 分離
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
      mapping: {},
      confidence: {}
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
    const dataAccess = Data.open(spreadsheetId);
    const {spreadsheet} = dataAccess;

    // サービスアカウントを編集者として自動登録 (Data.openで既に処理済み)
    // Note: Data.open()内でDriveApp.getFileById(id).addEditor()が既に実行されている
    console.log('connectToSheetInternal: サービスアカウント編集者権限はData.openで処理済み');

    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: `シート "${sheetName}" が見つかりません`,
          headers: [],
          columns: [],
          mapping: {},
      confidence: {}
        }
      };
    }

    // Batch operations for performance (CLAUDE.md準拠)
    const headers = sheet.getDataRange().getValues()[0] || [];

    // AI列判定を統合実行（Zero-Dependency Architecture）
    let columnMapping = { mapping: {}, confidence: {} };
    try {
      // サンプルデータを取得してAI分析実行
      const dataRange = sheet.getDataRange();
      const allData = dataRange.getValues();
      const sampleData = allData.slice(1, Math.min(11, allData.length)); // 最大10行のサンプル

      const analysisResult = detectColumnTypes(headers, sampleData);
      columnMapping = analysisResult || { mapping: {}, confidence: {} };

      console.log('DataService.connectToSheetInternal: AI分析完了', {
        headers: headers.length,
        sampleData: sampleData.length,
        mapping: columnMapping.mapping,
        confidence: columnMapping.confidence
      });
    } catch (aiError) {
      console.warn('DataService.connectToSheetInternal: AI分析エラー', aiError.message);
      // エラー時はデフォルト値を使用
    }

    return {
      success: true,
      sheet,
      headers, // UI必須データ追加
      mapping: columnMapping.mapping,      // フロントエンド期待形式
      confidence: columnMapping.confidence // 分離
    };

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
        mapping: {},
      confidence: {}
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
    const lastColumn = sheet.getLastColumn();
    const lastRow = sheet.getLastRow();

    if (lastColumn === 0 || lastRow === 0) {
      return {
        success: false,
        errorResponse: {
          success: false,
          message: 'スプレッドシートが空です',
          headers: [],
          columns: [],
          mapping: {},
      confidence: {}
        }
      };
    }

    // 🎯 GAS Best Practice: バッチでデータ取得
    const [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues();

    // サンプルデータを取得（最大5行）
    let sampleData = [];
    const sampleRowCount = Math.min(5, lastRow - 1);
    if (sampleRowCount > 0) {
      sampleData = sheet.getRange(2, 1, sampleRowCount, lastColumn).getValues();
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
        mapping: {},
      confidence: {}
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
    const session = ServiceFactory.getSession();
    const user = Data.findUserByEmail(session.email);
    if (!user) {
      return { success: false, message: 'User not found' };
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(userId);
    const config = configResult.success ? configResult.config : {};
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
    // 🎯 サービスアカウント認証でData.open()を使用（ServiceFactoryのフォールバック回避）
    const dataAccess = Data.open(spreadsheetId);
    const {spreadsheet} = dataAccess;
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

    // 防御的プログラミング: 入力値検証
    if (!Array.isArray(headers) || headers.length === 0) {
      console.warn('DataService.detectColumnTypes: 無効なheaders', headers);
      return { columns: [], mapping: {}, confidence: {} };
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

    // 🎯 AI判定システム実装 - 接続された列の中での最適マッチ判定
    const relativeMatchingResult = performRelativeMatching(analysisResults, headers);

    // AI判定結果をマッピングに反映
    Object.entries(relativeMatchingResult.mapping).forEach(([columnType, result]) => {
      if (result.shouldMap) {
        mapping.mapping[columnType] = result.index;
        mapping.confidence[columnType] = Math.round(result.confidence);
      } else {
        // マッピング対象外でも信頼度情報は記録（フロントエンド用）
        mapping.confidence[columnType] = Math.round(result.confidence);
      }
    });

    return {
      columns,
      mapping: mapping.mapping,        // フロントエンド期待形式
      confidence: mapping.confidence   // 分離
    };

  } catch (error) {
    console.error('DataService.detectColumnTypes: エラー', {
      error: error.message,
      stack: error.stack
    });

    // エラー時のフォールバック
    return {
      columns: [],
      mapping: {},
      confidence: {}
    };
  }
}

/**
 * 🎯 AI判定システム - 接続された列の中での最適マッチ判定
 * @param {Object} analysisResults - AI分析結果
 * @param {Array} headers - 列ヘッダー
 * @returns {Object} AI判定結果
 */
function performRelativeMatching(analysisResults, headers) {
  const targetTypes = ['answer', 'reason', 'class', 'name'];
  const mapping = {};
  const usedIndices = new Set();
  const mappingStats = [];

  // 1️⃣ 各ターゲットタイプについて最高スコアを収集
  targetTypes.forEach(targetType => {
    const result = analysisResults[targetType];
    mappingStats.push({
      targetType,
      index: result.index,
      confidence: result.confidence,
      headerName: headers[result.index] || '不明'
    });
  });

  // 2️⃣ 信頼度でソートして優先順位決定
  const sortedStats = mappingStats
    .filter(stat => stat.confidence > 0)
    .sort((a, b) => b.confidence - a.confidence);

  console.log('🔍 AI判定システム - 信頼度ランキング:', sortedStats.map(stat => ({
    ターゲット: stat.targetType,
    列名: stat.headerName,
    信頼度: `${Math.round(stat.confidence)}%`,
    順位: sortedStats.indexOf(stat) + 1
  })));

  // 3️⃣ 相対的な品質評価とマッピング決定
  sortedStats.forEach((stat, rank) => {
    const { targetType, index, confidence } = stat;

    // 相対的な判定ロジック
    let shouldMap = false;
    let adjustedConfidence = confidence;

    if (rank === 0 && confidence > 25) {
      // 最高スコア: 25%以上で採用
      shouldMap = true;
      adjustedConfidence = Math.min(confidence + 15, 100); // ボーナス
    } else if (rank === 1 && confidence > 20 && !usedIndices.has(index)) {
      // 2位: 20%以上で採用（重複除く）
      shouldMap = true;
      adjustedConfidence = Math.min(confidence + 10, 100); // 小ボーナス
    } else if (rank === 2 && confidence > 15 && !usedIndices.has(index)) {
      // 3位: 15%以上で採用（重複除く）
      shouldMap = true;
      adjustedConfidence = Math.min(confidence + 5, 100); // 最小ボーナス
    } else if (confidence > 30 && !usedIndices.has(index)) {
      // 高信頼度: 順位に関わらず30%以上で採用
      shouldMap = true;
    }

    // 重複チェック
    if (shouldMap && usedIndices.has(index)) {
      console.warn(`⚠️ 列インデックス${index}の重複を検出 - ${targetType}をスキップ`);
      shouldMap = false;
    }

    if (shouldMap) {
      usedIndices.add(index);
    }

    mapping[targetType] = {
      index,
      confidence: adjustedConfidence,
      shouldMap,
      rank: rank + 1,
      originalConfidence: confidence
    };
  });

  // 4️⃣ 未割り当てのターゲットタイプを処理
  targetTypes.forEach(targetType => {
    if (!mapping[targetType]) {
      const result = analysisResults[targetType];
      mapping[targetType] = {
        index: result.index,
        confidence: result.confidence,
        shouldMap: false,
        rank: null,
        originalConfidence: result.confidence
      };
    }
  });

  // 5️⃣ 結果サマリー
  const mappedCount = Object.values(mapping).filter(m => m.shouldMap).length;
  console.log('✅ AI判定システム完了:', {
    '対象列数': headers.length,
    'マッピング成功数': mappedCount,
    '成功率': `${Math.round(mappedCount / targetTypes.length * 100)}%`,
    'マッピング詳細': Object.entries(mapping)
      .filter(([_, m]) => m.shouldMap)
      .map(([type, m]) => `${type}→列${m.index}(${Math.round(m.confidence)}%)`)
  });

  return { mapping, stats: mappingStats };
}

/**
 * 🎯 高精度AI検出システム - 5次元統計分析
 * @param {Array} headers - 列ヘッダー
 * @param {Array} sampleData - サンプルデータ
 * @returns {Object} 分析結果
 */
function performHighPrecisionAnalysis(headers, sampleData) {

  const results = {
    answer: { index: -1, confidence: 0, factors: {} },
    reason: { index: -1, confidence: 0, factors: {} },
    class: { index: -1, confidence: 0, factors: {} },
    name: { index: -1, confidence: 0, factors: {} }
  };

  headers.forEach((header, index) => {
    if (!header) return;

    const samples = sampleData.map(row => row && row[index]).filter(v => v != null && v !== '');

    // 各列タイプに対する分析を実行
    Object.keys(results).forEach(columnType => {
      const analysis = analyzeColumnForType(header, samples, index, headers, columnType);

      if (analysis.confidence > results[columnType].confidence) {
        results[columnType] = analysis;
      }
    });
  });

  // 🎯 分析結果サマリー出力
  console.info('🔍 AI列判定分析サマリー', {
    '分析対象列数': headers.length,
    'サンプル行数': sampleData.length,
    '検出結果': Object.entries(results).map(([type, result]) => ({
      列種別: type,
      検出列: result.index >= 0 ? `インデックス ${result.index} ("${headers[result.index]}")` : '未検出',
      信頼度: `${Math.round(result.confidence)}%`,
      閾値達成: result.confidence >= 60 ? '✅' : '❌'
    })),
    '統一閾値': '60%',
    '最高信頼度': `${Math.max(...Object.values(results).map(r => Math.round(r.confidence)))}%`
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

  // 🎯 ヘッダー特化AI判定システム - サンプルデータなし設計
  let headerWeight, contentWeight, linguisticWeight, contextWeight, semanticWeight;

  // サンプルデータ有無による最適化
  const hasSampleData = samples && samples.length > 0;

  // 🎯 競合検出による動的重み調整
  const hasReasonKeywords = /理由|根拠|なぜ|why|わけ|説明/.test(headerLower);
  const hasAnswerKeywords = /答え|回答|answer|意見|予想|考え/.test(headerLower);
  const isConflictCase = hasReasonKeywords && hasAnswerKeywords && (targetType === 'answer' || targetType === 'reason');

  if (isConflictCase) {
    console.debug(`🎯 競合ケース検出 [${targetType}]: "${headerLower}" - コンテキスト重み強化`);
  }

  if (headerScore >= 90) {
    // 日本語完全一致 - ヘッダー特化重視
    headerWeight = hasSampleData ? 0.5 : 0.7;    // サンプルなし時は70%
    contentWeight = hasSampleData ? 0.2 : 0.0;   // サンプルなし時は無効
    linguisticWeight = hasSampleData ? 0.15 : 0.2; // 言語パターン強化
    contextWeight = hasSampleData ? 0.1 : 0.1;   // コンテキスト維持
    semanticWeight = hasSampleData ? 0.05 : 0.0; // サンプルなし時は無効
  } else if (headerScore >= 70) {
    // 強パターンマッチ - ヘッダー・コンテキスト重視
    headerWeight = hasSampleData ? 0.4 : 0.6;    // サンプルなし時は60%
    contentWeight = hasSampleData ? 0.25 : 0.0;  // サンプルなし時は無効
    linguisticWeight = hasSampleData ? 0.2 : 0.25; // 言語分析強化
    contextWeight = hasSampleData ? 0.1 : 0.15;  // コンテキスト強化
    semanticWeight = hasSampleData ? 0.05 : 0.0; // サンプルなし時は無効
  } else {
    // 標準分析 - ヘッダー・言語・コンテキスト重視
    headerWeight = hasSampleData ? 0.3 : 0.5;    // サンプルなし時は50%
    contentWeight = hasSampleData ? 0.3 : 0.0;   // サンプルなし時は無効
    linguisticWeight = hasSampleData ? 0.25 : 0.35; // 言語分析大幅強化
    contextWeight = hasSampleData ? 0.1 : 0.15;  // コンテキスト強化
    semanticWeight = hasSampleData ? 0.05 : 0.0; // サンプルなし時は無効
  }

  // 🎯 競合時の制約付き重み最適化
  if (isConflictCase) {
    const originalWeights = { headerWeight, contentWeight, linguisticWeight, contextWeight, semanticWeight };

    // 制約付き重み最適化の実行
    const optimizedWeights = optimizeWeightsWithConstraints(originalWeights, {
      contextBoost: 2.0,
      semanticBoost: hasSampleData ? 2.0 : 1.5,
      headerReduction: 0.8
    });

    // 最適化された重みを適用
    ({
      headerWeight,
      contentWeight,
      linguisticWeight,
      contextWeight,
      semanticWeight
    } = optimizedWeights);

    console.debug(`🎯 制約付き重み最適化完了 [${targetType}]: context=${(contextWeight*100).toFixed(1)}%, semantic=${(semanticWeight*100).toFixed(1)}%`);
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

  // 🎯 強化されたデバッグ出力 - 分析プロセスの可視化
  console.info(`🤖 AI列分析詳細 [${targetType}] インデックス:${index} ヘッダー:"${header}"`, {
    最終信頼度: Math.round(finalConfidence * 100) / 100,
    '重み配分': {
      'ヘッダー': `${(headerWeight * 100).toFixed(1)}%`,
      'コンテンツ': `${(contentWeight * 100).toFixed(1)}%`,
      '言語': `${(linguisticWeight * 100).toFixed(1)}%`,
      'コンテキスト': `${(contextWeight * 100).toFixed(1)}%`,
      'セマンティック': `${(semanticWeight * 100).toFixed(1)}%`
    },
    '各要素スコア': {
      'ヘッダーパターン': Math.round(factors.headerPattern * 100) / 100,
      'コンテンツ統計': Math.round(factors.contentStatistics * 100) / 100,
      '言語パターン': Math.round(factors.linguisticPatterns * 100) / 100,
      'コンテキスト': Math.round(factors.contextualClues * 100) / 100,
      'セマンティック': Math.round(factors.semanticCharacteristics * 100) / 100
    },
    '加重後スコア': {
      'ヘッダー貢献': Math.round(factors.headerPattern * headerWeight * 100) / 100,
      'コンテンツ貢献': Math.round(factors.contentStatistics * contentWeight * 100) / 100,
      '言語貢献': Math.round(factors.linguisticPatterns * linguisticWeight * 100) / 100,
      'コンテキスト貢献': Math.round(factors.contextualClues * contextWeight * 100) / 100,
      'セマンティック貢献': Math.round(factors.semanticCharacteristics * semanticWeight * 100) / 100
    }
  });

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
      // 🎯 Composite Patterns - 複合パターンで具体的マッチング
      composite: [
        /答え.*書/, /回答.*記入/, /考え.*書/, /意見.*述べ/, /予想.*記入/,
        /選択.*理由.*含/, /答え.*説明.*含/, /回答.*詳細/ // 複合的なanswer列
      ],
      strong: [
        /回答/, /意見/, /予想/, /選択/, /choice/,
        // 🎯 教育現場パターン強化
        /予想.*しよう/, /思い.*記入/, /どのように/, /何が/, /どんな/,
        /観察.*気づいた/, /気づいた.*こと/, /わかった.*こと/, /感じた.*こと/
      ],
      medium: [
        /結果/, /result/, /値/, /value/, /内容/, /content/,
        // 🎯 教育質問文パターン
        /しよう$/, /ましょう$/, /てください$/, /について/, /に関して/
      ],
      weak: [/データ/, /data/, /情報/, /info/],
      // 🎯 Smart Conflict Patterns - 段階的減点（30%減点）
      conflict: [
        { pattern: /理由.*だけ/, penalty: 0.2 },     // 「理由だけ」→ 80%減点
        { pattern: /なぜ.*のみ/, penalty: 0.2 },     // 「なぜのみ」→ 80%減点
        { pattern: /根拠.*記載/, penalty: 0.3 },     // 「根拠記載」→ 70%減点
        { pattern: /説明.*のみ/, penalty: 0.3 }      // 「説明のみ」→ 70%減点
      ]
    },
    reason: {
      primary: [/^理由$/, /^根拠$/, /^reason$/, /^説明$/, /^答えた理由$/],
      // 🎯 Composite Patterns - 複合パターンで理由系を強化
      composite: [
        /答えた.*理由/, /選んだ.*理由/, /考えた.*理由/, /そう.*思.*理由/,
        /理由.*書/, /根拠.*教/, /なぜ.*思/, /どうして.*考/,
        /背景.*あれば/, /体験.*あれば/, /経験.*あれば/
      ],
      strong: [
        /理由/, /根拠/, /reason/, /なぜ/, /why/, /わけ/, /説明/, /explanation/
      ],
      medium: [
        /詳細/, /detail/, /背景/, /background/, /コメント/, /comment/,
        // 🎯 感情・経験パターン（answer列との競合回避）
        /体験/, /経験/, /きっかけ/
      ],
      weak: [/その他/, /other/, /備考/, /note/],
      // 🎯 Smart Conflict Patterns - answer列との競合時の減点
      conflict: [
        { pattern: /答え.*中心/, penalty: 0.3 },      // 「答え中心」→ 70%減点
        { pattern: /回答.*メイン/, penalty: 0.3 }     // 「回答メイン」→ 70%減点
      ]
    },
    class: {
      primary: [/^クラス$/, /^class$/, /^組$/, /^年組$/],
      strong: [/クラス/, /class/, /組/, /年組/, /学級/, /grade/],
      medium: [/学年/, /year/, /グループ/, /group/],
      weak: [/チーム/, /team/]
    },
    name: {
      primary: [/^名前$/, /^氏名$/, /^name$/, /^お名前$/],
      // 🎯 Composite Patterns - 複合パターンで名前系を強化
      composite: [
        /名前.*書/, /名前.*入力/, /氏名.*記入/, /お名前.*教/,
        /name.*enter/, /name.*write/, /名前.*ましょう/, /氏名.*ましょう/
      ],
      strong: [
        /名前/, /氏名/, /name/, /お名前/, /ネーム/, /ニックネーム/
      ],
      medium: [
        /ユーザー/, /user/, /学生/, /student/, /生徒/, /児童/,
        // 🎯 一般的入力パターン（複合と重複しない単体のみ）
        /記入/, /入力/
      ],
      weak: [/id/, /アカウント/, /account/]
    }
  };

  const typePatterns = patterns[targetType] || {};
  let score = 0;

  // 🎯 Smart Penalty System - 段階的減点による論理的判定
  let penaltyMultiplier = 1.0;
  const conflictPatterns = typePatterns.conflict || [];
  for (const conflictPattern of conflictPatterns) {
    if (conflictPattern.pattern.test(headerLower)) {
      penaltyMultiplier *= conflictPattern.penalty; // 段階的減点
      console.debug(`🎯 競合パターン検出 [${targetType}]: "${headerLower}" → 減点率${conflictPattern.penalty}`);
      break; // 最初の競合パターンのみ適用
    }
  }

  // 🎯 Smart Pattern Evaluation Matrix - 全パターン評価による最適判定
  const patternEvaluations = [];

  // パターンレベル定義（重み付き評価）
  const patternLevels = {
    primary: { weight: 1.1, baseScore: 85 },
    composite: { weight: 1.2, baseScore: 80 },
    strong: { weight: 1.0, baseScore: 75 },
    medium: { weight: 0.9, baseScore: 60 },
    weak: { weight: 0.8, baseScore: 35 }
  };

  // 🎯 全パターンレベルを包括的に評価
  for (const [levelName, levelConfig] of Object.entries(patternLevels)) {
    const patterns = typePatterns[levelName] || [];

    for (const pattern of patterns) {
      if (pattern.test(headerLower)) {
        let levelScore = levelConfig.baseScore * levelConfig.weight;

        // 🎯 Primary特別ボーナス処理
        if (levelName === 'primary') {
          const ultraClearKeywords = ['クラス', '名前', '氏名', 'class', 'name'];
          if (ultraClearKeywords.some(keyword => headerLower.includes(keyword.toLowerCase()))) {
            levelScore += 5; // 超明確キーワードボーナス
          }
        }

        patternEvaluations.push({
          level: levelName,
          pattern: pattern.toString(),
          score: Math.round(levelScore),
          weight: levelConfig.weight
        });

        console.debug(`🎯 パターン評価 [${targetType}]: ${levelName} "${pattern}" → ${Math.round(levelScore)}点`);
      }
    }
  }

  // 🎯 Multi-Criteria Decision Matrix (MCDM) による競合解決
  if (patternEvaluations.length > 0) {
    const maxScore = Math.max(...patternEvaluations.map(e => e.score));
    const topEvaluations = patternEvaluations.filter(e => e.score === maxScore);

    if (topEvaluations.length === 1) {
      // 単一最高点 - 明確な選択
      const [{ score: bestScore, level: bestLevel }] = topEvaluations;
      score = bestScore;
      console.debug(`🎯 単一最適パターン [${targetType}]: ${bestLevel} → ${score}点`);
    } else {
      // 同点競合 - MCDM適用
      console.debug(`🎯 同点競合検出 [${targetType}]: ${topEvaluations.length}個のパターン → MCDM適用`);

      const mcdmResult = resolveConflictWithMCDM(topEvaluations, headerLower, targetType);
      score = mcdmResult.finalScore;

      console.debug(`🎯 MCDM競合解決 [${targetType}]: ${mcdmResult.selectedPattern} → ${score}点`);
    }
  }

  // 🎯 改善された否定的パターンフィルタ - 精密な誤判定防止
  const negativePatterns = [
    // リアクション系（完全一致重視）
    { pattern: /^like$/i, penalty: 40 },
    { pattern: /^いいね$/i, penalty: 40 },
    { pattern: /^good$/i, penalty: 35 },
    { pattern: /^understand$/i, penalty: 40 },
    { pattern: /^なるほど$/i, penalty: 35 },
    { pattern: /^curious$/i, penalty: 40 },
    { pattern: /^highlight$/i, penalty: 30 },
    { pattern: /^ハイライト$/i, penalty: 30 },
    // 感情表現（部分一致）
    { pattern: /！$/, penalty: 25 },
    { pattern: /!$/, penalty: 25 },
    { pattern: /^すごい/, penalty: 20 },
    { pattern: /^amazing/i, penalty: 20 },
    // 単発アクション
    { pattern: /^yes$/i, penalty: 30 },
    { pattern: /^no$/i, penalty: 30 },
    { pattern: /^はい$/, penalty: 30 },
    { pattern: /^いいえ$/, penalty: 30 }
  ];

  // 否定的パターンにマッチした場合は適度な減点（段階的）
  for (const negItem of negativePatterns) {
    if (negItem.pattern.test(headerLower)) {
      score = Math.max(0, score - negItem.penalty); // 段階的減点（最低0点）
      break; // 最初にマッチしたパターンのみ適用
    }
  }

  // 🎯 Smart Penalty適用 - 最終スコアに段階的減点を適用
  const finalScore = Math.round(score * penaltyMultiplier);

  if (penaltyMultiplier < 1.0) {
    console.debug(`🎯 最終スコア調整 [${targetType}]: ${score} × ${penaltyMultiplier} = ${finalScore}`);
  }

  return finalScore;
}

/**
 * 🎯 Multi-Criteria Decision Matrix (MCDM) による競合解決
 * @param {Array} conflictingEvaluations 競合するパターン評価
 * @param {string} headerLower 小文字ヘッダー
 * @param {string} targetType 対象列タイプ
 * @returns {Object} MCDM解決結果
 */
function resolveConflictWithMCDM(conflictingEvaluations, headerLower, targetType) {
  // MCDM基準の重み設定
  const mcdmCriteria = {
    headerSpecificity: 0.4,   // ヘッダー特異性（具体性）
    contextualFit: 0.3,       // 文脈適合度
    semanticDistance: 0.2,    // セマンティック距離
    patternComplexity: 0.1    // パターン複雑度
  };

  const evaluationResults = conflictingEvaluations.map(evaluation => {
    // 1. ヘッダー特異性スコア
    const specificityScore = calculateHeaderSpecificity(evaluation.pattern, headerLower);

    // 2. 文脈適合度スコア
    const contextualScore = calculateContextualFit(evaluation.level, targetType, headerLower);

    // 3. セマンティック距離スコア
    const semanticScore = calculateSemanticDistance(evaluation.pattern, targetType);

    // 4. パターン複雑度スコア
    const complexityScore = calculatePatternComplexity(evaluation.pattern);

    // MCDM重み付き総合スコア計算
    const mcdmScore =
      specificityScore * mcdmCriteria.headerSpecificity +
      contextualScore * mcdmCriteria.contextualFit +
      semanticScore * mcdmCriteria.semanticDistance +
      complexityScore * mcdmCriteria.patternComplexity;

    return {
      ...evaluation,
      mcdmScore: Math.round(mcdmScore * 100) / 100,
      criteria: { specificityScore, contextualScore, semanticScore, complexityScore }
    };
  });

  // 最高MCDMスコアの選択
  const bestMcdmEvaluation = evaluationResults.reduce((best, current) =>
    current.mcdmScore > best.mcdmScore ? current : best
  );

  // 元スコア + MCDM調整による最終スコア
  const finalScore = Math.round(bestMcdmEvaluation.score * (1 + bestMcdmEvaluation.mcdmScore * 0.1));

  return {
    selectedPattern: bestMcdmEvaluation.level,
    finalScore: Math.min(finalScore, 100), // 最大100点
    mcdmDetails: bestMcdmEvaluation.criteria
  };
}

/**
 * ヘッダー特異性計算
 */
function calculateHeaderSpecificity(pattern, headerLower) {
  // パターンの具体性を評価（より具体的なパターンほど高スコア）
  const patternStr = pattern.replace(/^\/|\/$/g, ''); // 正規表現マーカー除去
  const specificityFactors = {
    exactMatch: /^\^.*\$$/.test(pattern) ? 1.0 : 0.0,        // 完全一致
    wordBoundary: /\\b/.test(pattern) ? 0.3 : 0.0,           // 単語境界
    complexPattern: /\.\*/.test(pattern) ? 0.2 : 0.0,        // 複合パターン
    lengthFactor: Math.min(patternStr.length / 20, 0.5)      // 長さ係数
  };

  return Object.values(specificityFactors).reduce((sum, factor) => sum + factor, 0);
}

/**
 * 文脈適合度計算
 */
function calculateContextualFit(patternLevel, targetType, headerLower) {
  // パターンレベルと対象タイプの適合度
  const levelTypeFit = {
    primary: { answer: 0.9, reason: 0.9, name: 1.0, class: 1.0 },
    composite: { answer: 1.0, reason: 1.0, name: 0.8, class: 0.7 },
    strong: { answer: 0.8, reason: 0.8, name: 0.7, class: 0.8 },
    medium: { answer: 0.6, reason: 0.6, name: 0.6, class: 0.6 },
    weak: { answer: 0.4, reason: 0.4, name: 0.4, class: 0.4 }
  };

  return (levelTypeFit[patternLevel] || {})[targetType] || 0.5;
}

/**
 * セマンティック距離計算
 */
function calculateSemanticDistance(pattern, targetType) {
  // パターンと対象タイプ間のセマンティック親和性
  const semanticAffinities = {
    answer: [/答え/, /回答/, /answer/, /意見/, /予想/],
    reason: [/理由/, /根拠/, /reason/, /なぜ/, /説明/],
    name: [/名前/, /氏名/, /name/, /お名前/],
    class: [/クラス/, /class/, /組/, /学級/]
  };

  const targetAffinities = semanticAffinities[targetType] || [];
  const patternStr = pattern.replace(/^\/|\/$/g, '');

  // パターンが対象タイプの親和性キーワードを含むかチェック
  const affinityScore = targetAffinities.some(affinity =>
    affinity.test(patternStr)
  ) ? 1.0 : 0.3;

  return affinityScore;
}

/**
 * パターン複雑度計算
 */
function calculatePatternComplexity(pattern) {
  // 複雑なパターンほど高い特異性を持つ
  const complexityFactors = {
    quantifiers: (/[+*?{]/.test(pattern) ? 0.3 : 0.0),      // 量詞
    characterClasses: (/\[.*\]/.test(pattern) ? 0.2 : 0.0), // 文字クラス
    alternation: (/\|/.test(pattern) ? 0.2 : 0.0),          // 選択
    lookahead: (/\(\?=/.test(pattern) ? 0.3 : 0.0)         // 先読み
  };

  return Object.values(complexityFactors).reduce((sum, factor) => sum + factor, 0);
}

/**
 * 🎯 制約付き重み最適化（Constrained Weight Optimization）
 * @param {Object} originalWeights 元の重み設定
 * @param {Object} adjustments 調整パラメータ
 * @returns {Object} 最適化された重み
 */
function optimizeWeightsWithConstraints(originalWeights, adjustments) {
  // 制約条件: Σweight = 1.0, 0.01 ≤ weight ≤ 0.7
  const MIN_WEIGHT = 0.01;
  const MAX_WEIGHT = 0.7;
  const TARGET_SUM = 1.0;

  // 初期調整の適用
  const adjustedWeights = {
    headerWeight: originalWeights.headerWeight * adjustments.headerReduction,
    contentWeight: originalWeights.contentWeight,
    linguisticWeight: originalWeights.linguisticWeight,
    contextWeight: originalWeights.contextWeight * adjustments.contextBoost,
    semanticWeight: originalWeights.semanticWeight * adjustments.semanticBoost
  };

  // 制約違反のチェックと修正
  const weightKeys = Object.keys(adjustedWeights);

  // 1. 個別制約の適用（最小・最大値）
  for (const key of weightKeys) {
    adjustedWeights[key] = Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, adjustedWeights[key]));
  }

  // 2. 合計制約の適用（Lagrange乗数法の簡易版）
  const currentSum = Object.values(adjustedWeights).reduce((sum, weight) => sum + weight, 0);

  if (Math.abs(currentSum - TARGET_SUM) > 0.001) {
    // 重み正規化が必要
    const scaleFactor = TARGET_SUM / currentSum;

    // 優先順位付き調整（重要度の低い重みから調整）
    const priorityOrder = ['linguisticWeight', 'contentWeight', 'headerWeight', 'semanticWeight', 'contextWeight'];

    for (const key of priorityOrder) {
      adjustedWeights[key] *= scaleFactor;

      // 制約範囲内に収める
      adjustedWeights[key] = Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, adjustedWeights[key]));
    }

    // 最終正規化（微調整）
    const finalSum = Object.values(adjustedWeights).reduce((sum, weight) => sum + weight, 0);
    if (Math.abs(finalSum - TARGET_SUM) > 0.01) {
      const microAdjustment = (TARGET_SUM - finalSum) / weightKeys.length;
      for (const key of weightKeys) {
        adjustedWeights[key] += microAdjustment;
        adjustedWeights[key] = Math.max(MIN_WEIGHT, Math.min(MAX_WEIGHT, adjustedWeights[key]));
      }
    }
  }

  // 3. 最終検証
  const optimizedSum = Object.values(adjustedWeights).reduce((sum, weight) => sum + weight, 0);
  console.debug(`🎯 重み最適化検証: 合計=${optimizedSum.toFixed(3)}, 目標=1.000`);

  return adjustedWeights;
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
      if (/[あいうえお]だと思(う|います)/.test(sampleText)) score += 30;
      if (/(はい|いいえ|yes|no)/.test(sampleText)) score += 25;
      if (/\d+(番|号)/.test(sampleText)) score += 20;
      if (/(選択|肢)/.test(sampleText)) score += 15;
      break;

    case 'reason':
      // 理由によく現れるパターン
      if (/(だから|なぜなら|because)/.test(sampleText)) score += 35;
      if (/(と思う|と考える|だと思います)/.test(sampleText)) score += 25;
      if (/(ため|理由|根拠)/.test(sampleText)) score += 20;
      if (/(経験|体験|感じ)/.test(sampleText)) score += 15;
      break;

    case 'class':
      // クラス情報によく現れるパターン
      if (/\d+(年|組|班)/.test(sampleText)) score += 40;
      if (/[a-z]+(class|group)/.test(sampleText)) score += 30;
      if (/(グループ|チーム)\d+/.test(sampleText)) score += 20;
      break;

    case 'name':
      // 名前によく現れるパターン
      if (/^[ぁ-んァ-ン一-龯]+$/.test(sampleText)) score += 30; // 日本語名
      if (/^[a-zA-Z\s]+$/.test(sampleText)) score += 25; // 英語名
      if (/(さん|くん|ちゃん)$/.test(sampleText)) score += 20; // 敬称
      break;
  }

  return Math.min(score, 100);
}

/**
 * 4️⃣ コンテキスト推論 - 列位置と関係性を分析
 */
function analyzeContextualClues(header, index, allHeaders, targetType) {
  let score = 0;

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
 * 5️⃣ 強化されたセマンティック分析 - 多次元意味的特徴分析
 */
function analyzeSemanticCharacteristics(samples, targetType) {
  if (!samples || samples.length === 0) return 0;

  const textSamples = samples.filter(s => typeof s === 'string' && s.trim().length > 0);
  if (textSamples.length === 0) return 0;

  let score = 0;
  const uniqueValues = [...new Set(textSamples)];
  const uniquenessRatio = uniqueValues.length / textSamples.length;

  // 🎯 新機能: 文字列長分布分析
  const lengths = textSamples.map(s => s.trim().length);
  const avgLength = lengths.reduce((a, b) => a + b, 0) / lengths.length;
  const lengthVariance = lengths.reduce((sum, len) => sum + Math.pow(len - avgLength, 2), 0) / lengths.length;
  const lengthStdDev = Math.sqrt(lengthVariance);

  // 🎯 新機能: キーワード密度分析
  const keywordDensity = analyzeKeywordDensity(textSamples, targetType);

  switch (targetType) {
    case 'answer':
      // 回答は選択肢的で重複が多い + 長さが均一
      if (uniquenessRatio <= 0.3) score += 30;
      if (uniquenessRatio <= 0.5) score += 20;
      // 文字列長の均一性（回答は短く均一な傾向）
      if (avgLength <= 20 && lengthStdDev <= 10) score += 25;
      // 数値や選択肢パターン
      if (textSamples.some(s => /^[1-9]$/.test(s))) score += 25;
      if (textSamples.some(s => /^[A-D]$/.test(s))) score += 25;
      // キーワード密度
      score += keywordDensity;
      break;

    case 'reason':
      // 理由は個別性が高く、重複が少ない + 長さにバリエーション
      if (uniquenessRatio >= 0.8) score += 35;
      if (uniquenessRatio >= 0.6) score += 25;
      // 文字列長のバリエーション（理由は長さが多様）
      if (avgLength >= 30 && lengthStdDev >= 15) score += 20;
      // 説明的な言葉
      if (textSamples.some(s => s.includes('ため'))) score += 15;
      if (textSamples.some(s => s.includes('から'))) score += 10;
      // キーワード密度
      score += keywordDensity;
      break;

    case 'class':
      // クラス情報は限定的なパターン + 短い
      if (uniquenessRatio <= 0.2) score += 40;
      if (uniquenessRatio <= 0.4) score += 25;
      // 短い文字列（クラス名は通常短い）
      if (avgLength <= 15 && lengthStdDev <= 5) score += 30;
      if (textSamples.some(s => /\d/.test(s))) score += 20;
      // キーワード密度
      score += keywordDensity;
      break;

    case 'name':
      // 名前は個別性が高い + 適度な長さで均一
      if (uniquenessRatio >= 0.7) score += 35;
      if (uniquenessRatio >= 0.5) score += 20;
      // 名前の典型的な長さ（5-20文字程度）
      if (avgLength >= 5 && avgLength <= 20 && lengthStdDev <= 8) score += 25;
      // キーワード密度
      score += keywordDensity;
      break;
  }

  return Math.min(score, 100);
}

/**
 * キーワード密度分析（新機能）
 * @param {Array} samples - サンプルデータ
 * @param {string} targetType - 対象列タイプ
 * @returns {number} キーワード密度スコア
 */
function analyzeKeywordDensity(samples, targetType) {
  const sampleText = samples.join(' ').toLowerCase();
  let densityScore = 0;

  const keywords = {
    answer: ['はい', 'いいえ', 'yes', 'no', '選択', '番', '思う', 'だと思', '考える'],
    reason: ['だから', 'なぜなら', 'because', 'ため', '理由', '根拠', '経験', '感じ'],
    class: ['年', '組', '班', 'class', 'group', 'チーム', 'クラス'],
    name: ['さん', 'くん', 'ちゃん', '先生', '氏']
  };

  const typeKeywords = keywords[targetType] || [];
  let matchCount = 0;

  typeKeywords.forEach(keyword => {
    if (sampleText.includes(keyword.toLowerCase())) {
      matchCount++;
    }
  });

  // キーワードマッチ率に基づくスコア（最大15点）
  densityScore = Math.min(15, matchCount * 3);

  return densityScore;
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
    const dataAccess = Data.open(config.spreadsheetId);
    const {spreadsheet} = dataAccess;
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
    // 🎯 Zero-Dependency: Direct Data call
    const user = Data.findUserById(userId);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    const rowIndex = typeof rowId === 'string' ? parseInt(String(rowId).replace('row_', ''), 10) : parseInt(rowId, 10);
    if (!rowIndex || rowIndex < 2) {
      return createErrorResponse('Invalid row ID');
    }

    // 🔧 CLAUDE.md準拠: 行レベルロック機構 - 同時リアクション競合防止
    const reactionKey = `reaction_${config.spreadsheetId}_${config.sheetName}_${rowIndex}`;

    // eslint-disable-next-line no-undef
    if (typeof RequestGate !== 'undefined' && !RequestGate.enter(reactionKey)) {
      return {
        success: false,
        message: '同じ行に対するリアクション処理が実行中です。しばらくお待ちください。'
      };
    } else if (typeof RequestGate === 'undefined') {
      console.warn('dsAddReaction: RequestGate not available, proceeding without row lock');
    }

    try {
      const res = processReaction(config.spreadsheetId, config.sheetName, rowIndex, reactionType, user.userEmail);
      if (res && (res.success || res.status === 'success')) {
        // フロントエンド期待形式に合わせたレスポンス
        return {
          success: true,
          reactions: res.reactions || {},
          userReaction: res.userReaction || reactionType,
          action: res.action || 'added',
          message: res.message || 'リアクションを追加しました'
        };
      }

      return {
        success: false,
        message: res?.message || 'Failed to add reaction'
      };
    } catch (error) {
      console.error('DataService.dsAddReaction: エラー', error.message);
      return createExceptionResponse(error);
    } finally {
      // eslint-disable-next-line no-undef
      if (typeof RequestGate !== 'undefined') RequestGate.exit(reactionKey);
    }
  } catch (outerError) {
    console.error('DataService.dsAddReaction outer error:', outerError.message);
    if (typeof RequestGate !== 'undefined') {
      const reactionKey = `reaction_${userId}_${rowId}`;
      // eslint-disable-next-line no-undef
      RequestGate.exit(reactionKey);
    }
    return createExceptionResponse(outerError);
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
    // 🎯 Zero-Dependency: Direct Data call
    const user = Data.findUserById(userId);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getConfigSafe(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    // updateHighlightInSheet expects 'row_#'
    const rowNumber = typeof rowId === 'string' && rowId.startsWith('row_')
      ? rowId
      : `row_${parseInt(rowId, 10)}`;

    // 🔧 CLAUDE.md準拠: 行レベルロック機構 - 同時ハイライト競合防止
    const highlightKey = `highlight_${config.spreadsheetId}_${config.sheetName}_${rowNumber}`;

    // eslint-disable-next-line no-undef
    if (typeof RequestGate !== 'undefined' && !RequestGate.enter(highlightKey)) {
      return {
        success: false,
        message: '同じ行のハイライト処理が実行中です。しばらくお待ちください。'
      };
    } else if (typeof RequestGate === 'undefined') {
      console.warn('dsToggleHighlight: RequestGate not available, proceeding without row lock');
    }

    try {
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
    } finally {
      // eslint-disable-next-line no-undef
      if (typeof RequestGate !== 'undefined') RequestGate.exit(highlightKey);
    }
  } catch (outerError) {
    console.error('DataService.dsToggleHighlight outer error:', outerError.message);
    if (typeof RequestGate !== 'undefined') {
      const highlightKey = `highlight_${userId}_${rowId}`;
      // eslint-disable-next-line no-undef
      RequestGate.exit(highlightKey);
    }
    return createExceptionResponse(outerError);
  }
}

// Expose a stable namespace for non-global access patterns
if (typeof global !== 'undefined') {
  global.DataService = {
    getUserSheetData,
    processReaction,
    addReaction: dsAddReaction,
    toggleHighlight: dsToggleHighlight,
    connectToSheetInternal
  };
} else {
  this.DataService = {
    getUserSheetData,
    processReaction,
    addReaction: dsAddReaction,
    toggleHighlight: dsToggleHighlight,
    connectToSheetInternal
  };
}
