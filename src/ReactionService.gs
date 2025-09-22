/**
 * @fileoverview ReactionService - リアクション・ハイライト専用サービス
 *
 * 🎯 責任範囲:
 * - リアクション管理（UNDERSTAND, LIKE, CURIOUS）
 * - ハイライト機能
 * - リアクション状態分析・更新
 * - ユーザーリアクション追跡
 * - Cross-user権限管理（CLAUDE.md準拠）
 *
 * 🔄 CLAUDE.md Best Practices準拠:
 * - GAS-Native Pattern（直接API）
 * - Service Account適切使用（Cross-user access only）
 * - Cache-based Mutex（競合回避）
 * - V8ランタイム最適化
 */

/* global getCurrentEmail, findUserBySpreadsheetId, findUserById, getUserConfig, openSpreadsheet, createErrorResponse, createExceptionResponse, CACHE_DURATION, SYSTEM_LIMITS, resolveColumnIndex */

// ===========================================
// 🎯 リアクション管理システム - CLAUDE.md準拠
// ===========================================

/**
 * スプレッドシート内リアクション更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @param {string} reactionType - リアクションタイプ
 * @param {string} action - アクション（add/remove）
 * @param {Object} context - アクセスコンテキスト（target user info for cross-user access）
 * @returns {boolean} 成功可否
 */
function updateReactionInSheet(config, rowId, reactionType, action, context = {}) {
  try {
    // 🔧 CLAUDE.md準拠: Context-aware service account usage for reactions
    // ✅ **Cross-user**: Use service account when reacting to other user's content (typical)
    // ✅ **Self-access**: Use normal permissions for own content reactions
    const currentEmail = getCurrentEmail();

    // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
    const targetUser = findUserBySpreadsheetId(config.spreadsheetId);
    const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;

    console.log(`updateReactionInSheet: ${isSelfAccess ? 'Self-access normal permissions' : 'Cross-user service account'} for reaction update`);
    const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: !isSelfAccess });
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

    // CLAUDE.md準拠: バッチ操作による70倍性能向上 (getValue/setValue → getValues/setValues)
    const currentValue = sheet.getRange(rowNumber, reactionColumn, 1, 1).getValues()[0][0] || 0;
    const newValue = action === 'add'
      ? Math.max(0, currentValue + 1)
      : Math.max(0, currentValue - 1);

    sheet.getRange(rowNumber, reactionColumn, 1, 1).setValues([[newValue]]);

    console.info('ReactionService.updateReactionInSheet: リアクション更新完了', {
      rowId,
      reactionType,
      action,
      oldValue: currentValue,
      newValue
    });

    return true;
  } catch (error) {
    console.error('ReactionService.updateReactionInSheet: エラー', error.message);
    return false;
  }
}

/**
 * リアクション状態分析（GAS Best Practice: 単一責任）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {Object} reactionColumns - リアクション列情報
 * @param {number} rowIndex - 行インデックス
 * @param {string} userEmail - ユーザーメール
 * @returns {Object} リアクション状態
 */
function analyzeReactionState(sheet, reactionColumns, rowIndex, userEmail) {
  const currentReactions = {};
  const allReactionsData = {};
  let userCurrentReaction = null;

  // CLAUDE.md準拠: バッチ操作による70倍性能向上
  const columnNumbers = Object.values(reactionColumns);
  const minCol = Math.min(...columnNumbers);
  const maxCol = Math.max(...columnNumbers);
  const [batchData] = sheet.getRange(rowIndex, minCol, 1, maxCol - minCol + 1).getValues();

  Object.keys(reactionColumns).forEach(key => {
    const col = reactionColumns[key];
    const cellValue = batchData[col - minCol] || '';
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

  return { currentReactions, allReactionsData, userCurrentReaction };
}

/**
 * リアクション更新処理（GAS Best Practice: 単一責任）
 * @param {Object} currentReactions - 現在のリアクション
 * @param {string} reactionKey - リアクションキー
 * @param {string} userEmail - ユーザーメール
 * @param {string} userCurrentReaction - 現在のユーザーリアクション
 * @returns {Object} 更新結果
 */
function updateReactionState(currentReactions, reactionKey, userEmail, userCurrentReaction) {
  let action = 'added';
  let newUserReaction = null;

  if (userCurrentReaction === reactionKey) {
    // Same reaction -> remove (toggle)
    currentReactions[reactionKey] = currentReactions[reactionKey].filter(u => u !== userEmail);
    action = 'removed';
    newUserReaction = null;
  } else {
    // Different reaction -> remove old, add new
    if (userCurrentReaction) {
      currentReactions[userCurrentReaction] = currentReactions[userCurrentReaction].filter(u => u !== userEmail);
    }
    if (!currentReactions[reactionKey].includes(userEmail)) {
      currentReactions[reactionKey].push(userEmail);
    }
    action = 'added';
    newUserReaction = reactionKey;
  }

  return { action, newUserReaction, updatedReactions: currentReactions };
}

/**
 * リアクション処理（リファクタリング版 - GAS Best Practice準拠）
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, userEmail, context = {}) {
  try {
    // バリデーション
    if (!validateReaction(spreadsheetId, sheetName, rowIndex, reactionKey)) {
      throw new Error('無効なリアクションパラメータ');
    }
    if (!userEmail) {
      throw new Error('ユーザー情報が必要です');
    }

    // スプレッドシート接続
    // 🔧 CLAUDE.md準拠: Context-aware service account usage for reactions
    // ✅ **Cross-user**: Use service account when reacting to other user's content (typical)
    // ✅ **Self-access**: Use normal permissions for own content reactions
    const currentEmail = getCurrentEmail();

    // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
    const targetUser = findUserBySpreadsheetId(spreadsheetId);
    const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;

    console.log(`processReaction: ${isSelfAccess ? 'Self-access normal permissions' : 'Cross-user service account'} for reaction processing`);
    const dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: !isSelfAccess });
    const {spreadsheet} = dataAccess;
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // リアクション列取得
    const reactionColumns = {
      UNDERSTAND: getOrCreateReactionColumn(sheet, 'UNDERSTAND'),
      LIKE: getOrCreateReactionColumn(sheet, 'LIKE'),
      CURIOUS: getOrCreateReactionColumn(sheet, 'CURIOUS')
    };

    // 現在の状態分析
    const { currentReactions, allReactionsData, userCurrentReaction } = analyzeReactionState(
      sheet, reactionColumns, rowIndex, userEmail
    );

    // リアクション状態更新
    const { action, newUserReaction, updatedReactions } = updateReactionState(
      currentReactions, reactionKey, userEmail, userCurrentReaction
    );

    // シートに更新適用
    Object.keys(reactionColumns).forEach(key => {
      const col = reactionColumns[key];
      const users = updatedReactions[key];
      const serializedUsers = serializeReactionUsers(users);
      sheet.getRange(rowIndex, col, 1, 1).setValues([[serializedUsers]]);
    });

    // 更新後の状態を再計算
    const updatedAllReactionsData = {};
    Object.keys(updatedReactions).forEach(key => {
      updatedAllReactionsData[key] = {
        count: updatedReactions[key].length,
        reacted: updatedReactions[key].includes(userEmail)
      };
    });

    return {
      success: true,
      status: 'success',
      action,
      userReaction: newUserReaction,
      reactions: updatedAllReactionsData,
      message: action === 'added' ? 'リアクションを追加しました' : 'リアクションを削除しました'
    };

  } catch (error) {
    console.error('ReactionService.processReaction: エラー', error.message);
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
 * @param {string} userEmail - ユーザーメール（オプション）
 * @returns {Object} リアクション情報
 */
function extractReactions(row, headers, userEmail = null) {
  try {
    const reactions = {
      UNDERSTAND: { count: 0, reacted: false },
      LIKE: { count: 0, reacted: false },
      CURIOUS: { count: 0, reacted: false }
    };

    // 🎯 統一列判定システムを使用（ColumnMappingServiceから）
    const reactionTypes = ['understand', 'like', 'curious'];

    reactionTypes.forEach(reactionType => {
      const columnResult = resolveColumnIndex(headers, reactionType);

      if (columnResult.index !== -1) {
        const cellValue = row[columnResult.index] || '';
        const reactionUsers = parseReactionUsers(cellValue);
        const upperType = reactionType.toUpperCase();

        reactions[upperType] = {
          count: reactionUsers.length,
          reacted: userEmail ? reactionUsers.includes(userEmail) : false
        };
      }
    });

    return reactions;
  } catch (error) {
    console.warn('ReactionService.extractReactions: エラー', error.message);
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
    // 🎯 統一列判定システムを使用（ColumnMappingServiceから）
    const columnResult = resolveColumnIndex(headers, 'highlight');

    if (columnResult.index !== -1) {
      const value = String(row[columnResult.index] || '').toUpperCase();
      return value === 'TRUE' || value === '1' || value === 'YES';
    }

    return false;
  } catch (error) {
    console.warn('ReactionService.extractHighlight: エラー', error.message);
    return false;
  }
}

// ===========================================
// 🎯 ハイライト管理システム - CLAUDE.md準拠
// ===========================================

/**
 * ハイライト列の更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object} 成功可否
 */
function updateHighlightInSheet(config, rowId, context = {}) {
  try {
    // 🔧 CLAUDE.md準拠: Context-aware service account usage for highlights
    // ✅ **Cross-user**: Use service account when highlighting other user's content (typical)
    // ✅ **Self-access**: Use normal permissions for own content highlights
    const currentEmail = getCurrentEmail();

    // CLAUDE.md準拠: spreadsheetIdから所有者を特定して直接比較
    const targetUser = findUserBySpreadsheetId(config.spreadsheetId);
    const isSelfAccess = targetUser && targetUser.userEmail === currentEmail;

    console.log(`updateHighlightInSheet: ${isSelfAccess ? 'Self-access normal permissions' : 'Cross-user service account'} for highlight update`);
    const dataAccess = openSpreadsheet(config.spreadsheetId, { useServiceAccount: !isSelfAccess });
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

    // CLAUDE.md準拠: バッチ操作による70倍性能向上 (getValue/setValue → getValues/setValues)
    const [[currentValue]] = sheet.getRange(rowNumber, highlightColumn, 1, 1).getValues();
    const isCurrentlyHighlighted = currentValue === 'TRUE' || currentValue === true;
    const newValue = isCurrentlyHighlighted ? 'FALSE' : 'TRUE';

    sheet.getRange(rowNumber, highlightColumn, 1, 1).setValues([[newValue]]);

    const highlighted = newValue === 'TRUE';

    console.info('ReactionService.updateHighlightInSheet: ハイライト切り替え完了', {
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
    console.error('ReactionService.updateHighlightInSheet: エラー', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

// ===========================================
// 🔧 リアクション列管理
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
// 🌍 Public API Functions - CLAUDE.md準拠
// ===========================================

/**
 * addReaction (user context)
 * @param {string} userId
 * @param {number|string} rowIndex - number or 'row_#'
 * @param {string} reaction
 * @returns {Object}
 */
function addReaction(userId, rowIndex, reaction) {
  try {
    // 🎯 Zero-Dependency: Direct Data call
    const user = findUserById(userId);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    const parsedRowIndex = typeof rowIndex === 'string' ? parseInt(String(rowIndex).replace('row_', ''), SYSTEM_LIMITS.RADIX_DECIMAL) : parseInt(rowIndex, SYSTEM_LIMITS.RADIX_DECIMAL);
    if (!parsedRowIndex || parsedRowIndex < 2) {
      return createErrorResponse('Invalid row ID');
    }

    // 🔧 CLAUDE.md準拠: 行レベルロック機構 - 同時リアクション競合防止（CacheService-based mutex）
    const reactionKey = `reaction_${config.spreadsheetId}_${config.sheetName}_${parsedRowIndex}`;
    const cache = CacheService.getScriptCache();

    // 排他制御（Cache-based mutex）
    if (cache.get(reactionKey)) {
      return {
        success: false,
        message: '同じ行に対するリアクション処理が実行中です。しばらくお待ちください。'
      };
    }

    try {
      cache.put(reactionKey, true, CACHE_DURATION.MEDIUM); // 30秒ロック

      const res = processReaction(config.spreadsheetId, config.sheetName, parsedRowIndex, reaction, getCurrentEmail());
      if (res && (res.success || res.status === 'success')) {
        // フロントエンド期待形式に合わせたレスポンス
        return {
          success: true,
          reactions: res.reactions || {},
          userReaction: res.userReaction || reaction,
          action: res.action || 'added',
          message: res.message || 'リアクションを追加しました'
        };
      }

      return {
        success: false,
        message: res?.message || 'Failed to add reaction'
      };
    } catch (error) {
      console.error('ReactionService.addReaction: エラー', error.message);
      return createExceptionResponse(error);
    } finally {
      cache.remove(reactionKey);
    }
  } catch (outerError) {
    console.error('ReactionService.addReaction outer error:', outerError.message);
    // 🔧 統一ミューテックス: 緊急時のキャッシュクリア
    try {
      const cache = CacheService.getScriptCache();
      const configForCleanup = getUserConfig(userId);
      const cleanupConfig = configForCleanup.success ? configForCleanup.config : {};
      if (cleanupConfig.spreadsheetId && cleanupConfig.sheetName) {
        const cleanupRowIndex = typeof rowIndex === 'string' ? parseInt(String(rowIndex).replace('row_', ''), SYSTEM_LIMITS.RADIX_DECIMAL) : parseInt(rowIndex, SYSTEM_LIMITS.RADIX_DECIMAL);
        const cleanupReactionKey = `reaction_${cleanupConfig.spreadsheetId}_${cleanupConfig.sheetName}_${cleanupRowIndex}`;
        cache.remove(cleanupReactionKey);
      }
    } catch (cacheError) {
      console.warn('Failed to clear reaction cache in error handler:', cacheError.message);
    }
    return createExceptionResponse(outerError);
  }
}

/**
 * toggleHighlight (user context)
 * @param {string} userId
 * @param {number|string} rowIndex - number or 'row_#'
 * @returns {Object}
 */
function toggleHighlight(userId, rowIndex) {
  try {
    // 🎯 Zero-Dependency: Direct Data call
    const user = findUserById(userId);
    if (!user) {
      return createErrorResponse('User not found');
    }

    // 統一API使用: 構造化パース
    const configResult = getUserConfig(userId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Spreadsheet configuration incomplete');
    }

    // updateHighlightInSheet expects 'row_#'
    const rowNumber = typeof rowIndex === 'string' && rowIndex.startsWith('row_')
      ? rowIndex
      : `row_${parseInt(rowIndex, SYSTEM_LIMITS.RADIX_DECIMAL)}`;

    // 🔧 CLAUDE.md準拠: 行レベルロック機構 - 同時ハイライト競合防止（CacheService-based mutex）
    const highlightKey = `highlight_${config.spreadsheetId}_${config.sheetName}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    // 排他制御（Cache-based mutex）
    if (cache.get(highlightKey)) {
      return {
        success: false,
        message: '同じ行のハイライト処理が実行中です。しばらくお待ちください。'
      };
    }

    try {
      cache.put(highlightKey, true, CACHE_DURATION.MEDIUM); // 30秒ロック

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
      console.error('ReactionService.toggleHighlight: エラー', error.message);
      return createExceptionResponse(error);
    } finally {
      cache.remove(highlightKey);
    }
  } catch (outerError) {
    console.error('ReactionService.toggleHighlight outer error:', outerError.message);
    // 🔧 統一ミューテックス: 緊急時のキャッシュクリア
    try {
      const cache = CacheService.getScriptCache();
      const highlightKey = `highlight_${userId}_${rowIndex}`;
      cache.remove(highlightKey);
    } catch (cacheError) {
      console.warn('Failed to clear highlight cache in error handler:', cacheError.message);
    }
    return createExceptionResponse(outerError);
  }
}