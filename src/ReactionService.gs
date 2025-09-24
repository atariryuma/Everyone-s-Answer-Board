/**
 * @fileoverview ReactionService - マルチテナント対応リアクション・ハイライト管理
 *
 * 🎯 GAS-Native Architecture:
 * - リアクション管理（UNDERSTAND, LIKE, CURIOUS）
 * - ハイライト機能
 * - マルチテナントセキュリティ
 * - 直接SpreadsheetApp操作（Zero-Dependency）
 * - Service Account適切使用（Cross-user access only）
 */

/* global getCurrentEmail, findUserBySpreadsheetId, findUserById, getUserConfig, openSpreadsheet, createErrorResponse, createExceptionResponse, CACHE_DURATION, SYSTEM_LIMITS, resolveColumnIndex, isAdministrator */

// ===========================================
// 🎯 リアクション管理システム - CLAUDE.md準拠
// ===========================================

// ===========================================
// 🔧 セキュリティ・監査機能
// ===========================================

/**
 * マルチテナント権限検証
 * @param {string} actorEmail - 操作者
 * @param {string} targetUserId - 対象ユーザーID
 * @returns {boolean} 権限があるかどうか
 */
function validateReactionPermission(actorEmail, targetUserId) {
  if (!actorEmail || !targetUserId) return false;

  // 管理者は全アクセス可能
  if (isAdministrator(actorEmail)) return true;

  // ボード公開設定チェック（簡易実装）
  const targetUser = findUserById(targetUserId);
  if (!targetUser) return false;

  const configResult = getUserConfig(targetUserId);
  const config = configResult.success ? configResult.config : {};

  // 公開ボードまたは自分のボードの場合は許可
  return config.isPublished || targetUser.userEmail === actorEmail;
}

/**
 * 監査ログ記録
 * @param {string} action - アクション
 * @param {Object} details - 詳細情報
 */
function logReactionAudit(action, details) {
  const timestamp = new Date().toISOString();
  const logEntry = {
    timestamp,
    action,
    actor: details.actor || 'unknown',
    target: details.target || 'unknown',
    spreadsheet: details.spreadsheetId ? `${details.spreadsheetId.substring(0, 8)}***` : 'unknown',
    details: details.extra || {}
  };

  console.log(`REACTION_AUDIT: ${JSON.stringify(logEntry)}`);
}


/**
 * 🚀 GAS-Native直接リアクション処理
 * @param {Sheet} sheet - スプレッドシートオブジェクト
 * @param {number} rowNumber - 行番号
 * @param {string} reactionType - リアクション種類
 * @param {string} actorEmail - 操作者メール
 * @returns {Object} 処理結果
 */
function processReactionDirect(sheet, rowNumber, reactionType, actorEmail) {
  const reactionTypes = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

  if (!reactionTypes.includes(reactionType)) {
    throw new Error('Invalid reaction type');
  }

  // 🎯 ヘッダー行から列位置取得
  const [headers = []] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
  const reactionColumns = {};

  reactionTypes.forEach(type => {
    const colIndex = headers.findIndex(header =>
      String(header).toUpperCase().includes(type)
    );
    reactionColumns[type] = colIndex !== -1 ? colIndex + 1 : createReactionColumn(sheet, type);
  });

  // 🔄 現在のリアクション状態を取得
  const columnIndexes = Object.values(reactionColumns);
  const minCol = Math.min(...columnIndexes);
  const maxCol = Math.max(...columnIndexes);
  const [rowData = []] = sheet.getRange(rowNumber, minCol, 1, maxCol - minCol + 1).getValues();

  const currentReactions = {};
  const updatedReactions = {};
  let userCurrentReaction = null;

  // 現在の状態解析
  reactionTypes.forEach(type => {
    const col = reactionColumns[type];
    const cellValue = rowData[col - minCol] || '';
    const users = parseReactionUsers(String(cellValue));

    currentReactions[type] = users;

    if (users.includes(actorEmail)) {
      userCurrentReaction = type;
    }
  });

  // リアクション状態更新ロジック
  let action = 'added';
  let newUserReaction = reactionType;

  reactionTypes.forEach(type => {
    updatedReactions[type] = [...currentReactions[type]];
  });

  if (userCurrentReaction === reactionType) {
    // 同じリアクション → 削除（トグル）
    updatedReactions[reactionType] = updatedReactions[reactionType].filter(u => u !== actorEmail);
    action = 'removed';
    newUserReaction = null;
  } else {
    // 異なるリアクション → 古いのを削除、新しいのを追加
    if (userCurrentReaction) {
      updatedReactions[userCurrentReaction] = updatedReactions[userCurrentReaction].filter(u => u !== actorEmail);
    }
    if (!updatedReactions[reactionType].includes(actorEmail)) {
      updatedReactions[reactionType].push(actorEmail);
    }
  }

  // 🚀 一括更新（CLAUDE.md準拠70倍性能向上）
  const updateData = [];
  reactionTypes.forEach(type => {
    const col = reactionColumns[type];
    const serialized = serializeReactionUsers(updatedReactions[type]);
    updateData.push([col, serialized]);
  });

  updateData.forEach(([col, value]) => {
    sheet.getRange(rowNumber, col, 1, 1).setValues([[value]]);
  });

  // レスポンス形式構築
  const reactions = {};
  reactionTypes.forEach(type => {
    reactions[type] = {
      count: updatedReactions[type].length,
      reacted: updatedReactions[type].includes(actorEmail)
    };
  });

  return {
    action,
    userReaction: newUserReaction,
    reactions
  };
}

/**
 * 🚀 GAS-Native直接ハイライト処理
 * @param {Sheet} sheet - スプレッドシートオブジェクト
 * @param {number} rowNumber - 行番号
 * @returns {Object} 処理結果
 */
function processHighlightDirect(sheet, rowNumber) {
  const [headers = []] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();

  // ハイライト列を探す
  const highlightCol = headers.findIndex(header =>
    String(header).toUpperCase().includes('HIGHLIGHT')
  ) + 1;

  // ハイライト列が存在しない場合はエラー
  if (highlightCol === 0) {
    console.error('processHighlightDirect: HIGHLIGHT column not found. Columns must be pre-created during data source setup.');
    throw new Error('Required HIGHLIGHT column not found. Please reconnect your data source to set up highlight columns.');
  }

  // 現在の値を取得してトグル
  const [[currentValue = '']] = sheet.getRange(rowNumber, highlightCol, 1, 1).getValues();
  const isHighlighted = String(currentValue).toUpperCase() === 'TRUE';
  const newValue = isHighlighted ? 'FALSE' : 'TRUE';

  sheet.getRange(rowNumber, highlightCol, 1, 1).setValues([[newValue]]);

  return {
    highlighted: newValue === 'TRUE'
  };
}

/**
 * リアクション列作成（必要時のみ）
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} reactionType - リアクション種類
 * @returns {number} 列番号
 */
function createReactionColumn(sheet, reactionType) {
  console.error(`createReactionColumn: Column ${reactionType} not found. Columns must be pre-created during data source setup.`);
  throw new Error(`Required reaction column '${reactionType}' not found. Please reconnect your data source to set up reaction columns.`);
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
    const columnResult = resolveColumnIndex(headers, 'HIGHLIGHT');

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


// ===========================================
// 🌍 Public API Functions - CLAUDE.md準拠
// ===========================================

/**
 * リアクション送信（マルチテナント対応・GAS-Native）
 * @param {string} targetUserId - 対象ユーザー（ボード所有者）ID
 * @param {number|string} rowIndex - 行番号または'row_#'
 * @param {string} reactionType - リアクション種類
 * @returns {Object} 処理結果
 */
function addReaction(targetUserId, rowIndex, reactionType) {
  const actorEmail = getCurrentEmail();

  try {
    // 🛡️ マルチテナント権限検証
    if (!validateReactionPermission(actorEmail, targetUserId)) {
      logReactionAudit('reaction_denied', {
        actor: actorEmail,
        target: targetUserId,
        reason: 'access_denied',
        extra: { reactionType, rowIndex }
      });
      return createErrorResponse('Access denied to target board');
    }

    // 🎯 GAS-Native: 直接データアクセス
    const targetUser = findUserById(targetUserId);
    if (!targetUser) {
      return createErrorResponse('Target user not found');
    }

    const configResult = getUserConfig(targetUserId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Board configuration incomplete');
    }

    // 行番号正規化
    const rowNumber = typeof rowIndex === 'string'
      ? parseInt(rowIndex.replace('row_', ''), 10)
      : parseInt(rowIndex, 10);

    if (!rowNumber || rowNumber < 2) {
      return createErrorResponse('Invalid row ID');
    }

    // 🔐 Cache-based行レベルロック
    const lockKey = `reaction_${config.spreadsheetId}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    if (cache.get(lockKey)) {
      return createErrorResponse('同時リアクション処理中です。お待ちください。');
    }

    try {
      cache.put(lockKey, actorEmail, CACHE_DURATION.SHORT);

      // 🔧 CLAUDE.md準拠: クロスユーザーアクセス判定
      const isSelfAccess = targetUser.userEmail === actorEmail;
      const dataAccess = openSpreadsheet(config.spreadsheetId, {
        useServiceAccount: !isSelfAccess,
        context: 'reaction_processing'
      });

      if (!dataAccess) {
        throw new Error('Failed to access target spreadsheet');
      }

      const { spreadsheet } = dataAccess;
      const sheet = spreadsheet.getSheetByName(config.sheetName);
      if (!sheet) {
        throw new Error('Target sheet not found');
      }

      // 🚀 GAS-Native: 直接リアクション処理
      const result = processReactionDirect(sheet, rowNumber, reactionType, actorEmail);

      // 📊 監査ログ
      logReactionAudit('reaction_processed', {
        actor: actorEmail,
        target: targetUser.userEmail,
        spreadsheetId: config.spreadsheetId,
        extra: {
          reactionType,
          rowNumber,
          action: result.action,
          accessMethod: isSelfAccess ? 'normal' : 'service_account'
        }
      });

      return {
        success: true,
        reactions: result.reactions,
        userReaction: result.userReaction,
        action: result.action,
        message: result.action === 'added' ? 'リアクションを追加しました' : 'リアクションを削除しました'
      };

    } finally {
      cache.remove(lockKey);
    }

  } catch (error) {
    console.error('addReaction error:', error.message);
    logReactionAudit('reaction_error', {
      actor: actorEmail,
      target: targetUserId,
      error: error.message,
      extra: { reactionType, rowIndex }
    });
    return createExceptionResponse(error);
  }
}

/**
 * ハイライト切り替え（マルチテナント対応・GAS-Native）
 * @param {string} targetUserId - 対象ユーザー（ボード所有者）ID
 * @param {number|string} rowIndex - 行番号または'row_#'
 * @returns {Object} 処理結果
 */
function toggleHighlight(targetUserId, rowIndex) {
  const actorEmail = getCurrentEmail();

  try {
    // 🛡️ マルチテナント権限検証
    if (!validateReactionPermission(actorEmail, targetUserId)) {
      logReactionAudit('highlight_denied', {
        actor: actorEmail,
        target: targetUserId,
        reason: 'access_denied',
        extra: { rowIndex }
      });
      return createErrorResponse('Access denied to target board');
    }

    const targetUser = findUserById(targetUserId);
    if (!targetUser) {
      return createErrorResponse('Target user not found');
    }

    const configResult = getUserConfig(targetUserId);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Board configuration incomplete');
    }

    const rowNumber = typeof rowIndex === 'string'
      ? parseInt(rowIndex.replace('row_', ''), 10)
      : parseInt(rowIndex, 10);

    if (!rowNumber || rowNumber < 2) {
      return createErrorResponse('Invalid row ID');
    }

    // 🔐 Cache-based行レベルロック
    const lockKey = `highlight_${config.spreadsheetId}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    if (cache.get(lockKey)) {
      return createErrorResponse('同時ハイライト処理中です。お待ちください。');
    }

    try {
      cache.put(lockKey, actorEmail, CACHE_DURATION.SHORT);

      // 🔧 CLAUDE.md準拠: クロスユーザーアクセス判定
      const isSelfAccess = targetUser.userEmail === actorEmail;
      const dataAccess = openSpreadsheet(config.spreadsheetId, {
        useServiceAccount: !isSelfAccess,
        context: 'highlight_processing'
      });

      if (!dataAccess) {
        throw new Error('Failed to access target spreadsheet');
      }

      const { spreadsheet } = dataAccess;
      const sheet = spreadsheet.getSheetByName(config.sheetName);
      if (!sheet) {
        throw new Error('Target sheet not found');
      }

      // 🚀 GAS-Native: 直接ハイライト処理
      const result = processHighlightDirect(sheet, rowNumber);

      // 📊 監査ログ
      logReactionAudit('highlight_processed', {
        actor: actorEmail,
        target: targetUser.userEmail,
        spreadsheetId: config.spreadsheetId,
        extra: {
          rowNumber,
          highlighted: result.highlighted,
          accessMethod: isSelfAccess ? 'normal' : 'service_account'
        }
      });

      return {
        success: true,
        highlighted: result.highlighted,
        message: result.highlighted ? 'ハイライトしました' : 'ハイライトを解除しました'
      };

    } finally {
      cache.remove(lockKey);
    }

  } catch (error) {
    console.error('toggleHighlight error:', error.message);
    logReactionAudit('highlight_error', {
      actor: actorEmail,
      target: targetUserId,
      error: error.message,
      extra: { rowIndex }
    });
    return createExceptionResponse(error);
  }
}