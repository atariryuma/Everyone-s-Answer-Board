/**
 * @fileoverview ReactionService - マルチテナント対応リアクション・ハイライト管理
 *
 * 🎯 GAS-Native Architecture:
 * - リアクション管理（UNDERSTAND, LIKE, CURIOUS）
 * - ハイライト機能
 * - マルチテナントセキュリティ
 * - 直接SpreadsheetApp操作（Zero-Dependency）
 * - 同一ドメイン共有設定によるアクセス管理
 */

/* global getCurrentEmail, findUserBySpreadsheetId, findUserById, getUserConfig, openSpreadsheet, createErrorResponse, createExceptionResponse, CACHE_DURATION, SYSTEM_LIMITS, isAdministrator */








/**
 * マルチテナント権限検証（事前読み込みデータ版）
 * ✅ CLAUDE.md準拠: DB呼び出しなしの権限チェックでパフォーマンス最適化
 * @param {string} actorEmail - 操作者
 * @param {Object} targetUser - 事前取得済みの対象ユーザー
 * @param {Object} config - 事前取得済みの設定
 * @returns {boolean} 権限があるかどうか
 */
function validateReactionPermissionWithPreloadedData(actorEmail, targetUser, config) {
  if (!actorEmail || !targetUser) return false;

  if (isAdministrator(actorEmail)) return true;

  return config.isPublished || targetUser.userEmail === actorEmail;
}


/**
 * 🚀 GAS-Native直接リアクション処理
 * ✅ Graceful Degradation: ヘッダー取得失敗時も継続動作
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

  const dataRange = sheet.getDataRange();
  const [headers = []] = dataRange.getValues();

  if (!headers || headers.length === 0) {
    console.warn(`⚠️ processReactionDirect: Headers unavailable (likely due to API quota). Reaction feature temporarily disabled.`, {
      rowNumber,
      reactionType,
      context: 'graceful-degradation'
    });
    return {
      action: 'unavailable',
      userReaction: null,
      reactions: {
        UNDERSTAND: { count: 0, reacted: false },
        LIKE: { count: 0, reacted: false },
        CURIOUS: { count: 0, reacted: false }
      },
      message: 'リアクション機能が一時的に利用できません。しばらくしてから再度お試しください。'
    };
  }

  const reactionColumns = {};

  reactionTypes.forEach(type => {
    const colIndex = headers.findIndex(header => String(header).toUpperCase().trim() === type);
    if (colIndex === -1) {
      console.error(`❌ processReactionDirect: Required reaction column '${type}' not found`, {
        type,
        availableHeaders: headers.map(h => String(h || '').trim()).filter(h => h).join(', '),
        headerCount: headers.length,
        systemColumns: ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT']
      });
      throw new Error(`リアクション列「${type}」が見つかりません。データソースを再接続してリアクション列を作成してください。`);
    }
    reactionColumns[type] = colIndex + 1;
  });

  const columnIndexes = Object.values(reactionColumns);
  const minCol = Math.min(...columnIndexes);
  const maxCol = Math.max(...columnIndexes);
  const [rowData = []] = sheet.getRange(rowNumber, minCol, 1, maxCol - minCol + 1).getValues();

  const currentReactions = {};
  const updatedReactions = {};
  let userCurrentReaction = null;

  reactionTypes.forEach(type => {
    const col = reactionColumns[type];
    const cellValue = rowData[col - minCol] || '';
    const cellString = String(cellValue);
    const users = cellString && cellString.trim()
      ? cellString.trim().split('|').filter(email => email.trim().length > 0)
      : [];

    currentReactions[type] = users;

    if (users.includes(actorEmail)) {
      userCurrentReaction = type;
    }
  });

  let action = 'added';
  let newUserReaction = reactionType;

  reactionTypes.forEach(type => {
    updatedReactions[type] = [...currentReactions[type]];
  });

  if (userCurrentReaction === reactionType) {
    updatedReactions[reactionType] = updatedReactions[reactionType].filter(u => u !== actorEmail);
    action = 'removed';
    newUserReaction = null;
  } else {
    if (userCurrentReaction) {
      updatedReactions[userCurrentReaction] = updatedReactions[userCurrentReaction].filter(u => u !== actorEmail);
    }
    if (!updatedReactions[reactionType].includes(actorEmail)) {
      updatedReactions[reactionType].push(actorEmail);
    }
  }


  reactionTypes.forEach(type => {
    const col = reactionColumns[type];
    const users = updatedReactions[type];
    const serialized = Array.isArray(users) && users.length > 0
      ? users.filter(email => email && email.trim().length > 0).join('|')
      : '';
    rowData[col - minCol] = serialized;
  });

  sheet.getRange(rowNumber, minCol, 1, maxCol - minCol + 1).setValues([rowData]);

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
 * ✅ Graceful Degradation: ヘッダー取得失敗時も継続動作
 * @param {Sheet} sheet - スプレッドシートオブジェクト
 * @param {number} rowNumber - 行番号
 * @returns {Object} 処理結果
 */
function processHighlightDirect(sheet, rowNumber) {
  const dataRange = sheet.getDataRange();
  const [headers = []] = dataRange.getValues();

  if (!headers || headers.length === 0) {
    console.warn(`⚠️ processHighlightDirect: Headers unavailable (likely due to API quota). Highlight feature temporarily disabled.`, {
      rowNumber,
      context: 'graceful-degradation'
    });
    return {
      highlighted: false,
      message: 'ハイライト機能が一時的に利用できません。しばらくしてから再度お試しください。'
    };
  }

  const highlightColIndex = headers.findIndex(header => String(header).toUpperCase().trim() === 'HIGHLIGHT');

  if (highlightColIndex === -1) {
    console.error(`❌ processHighlightDirect: Required HIGHLIGHT column not found`, {
      availableHeaders: headers.map(h => String(h || '').trim()).filter(h => h).join(', '),
      headerCount: headers.length,
      systemColumns: ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT']
    });
    throw new Error('ハイライト列「HIGHLIGHT」が見つかりません。データソースを再接続してハイライト列を作成してください。');
  }

  const highlightCol = highlightColIndex + 1;

  const highlightRange = sheet.getRange(rowNumber, highlightCol, 1, 1);
  const [[currentValue = '']] = highlightRange.getValues();
  const isHighlighted = String(currentValue).toUpperCase() === 'TRUE';
  const newValue = isHighlighted ? 'FALSE' : 'TRUE';

  highlightRange.setValues([[newValue]]);

  return {
    highlighted: newValue === 'TRUE'
  };
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

    const reactionTypes = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

    reactionTypes.forEach(reactionType => {
      const columnIndex = headers.findIndex(header => String(header).toUpperCase().trim() === reactionType);

      if (columnIndex !== -1) {
        const cellValue = row[columnIndex] || '';
        const reactionUsers = cellValue && typeof cellValue === 'string' && cellValue.trim()
          ? cellValue.trim().split('|').filter(email => email.trim().length > 0)
          : [];

        reactions[reactionType] = {
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
    const columnIndex = headers.findIndex(header => String(header).toUpperCase().trim() === 'HIGHLIGHT');

    if (columnIndex !== -1) {
      const value = String(row[columnIndex] || '').toUpperCase();
      return value === 'TRUE' || value === '1' || value === 'YES';
    }

    return false;
  } catch (error) {
    console.warn('ReactionService.extractHighlight: エラー', error.message);
    return false;
  }
}








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
    const isAdmin = isAdministrator(actorEmail);
    const preloadedAuth = { email: actorEmail, isAdmin };

    const targetUser = findUserById(targetUserId, {
      requestingUser: actorEmail,
      preloadedAuth,
      allowPublishedRead: true
    });
    if (!targetUser) {
      return createErrorResponse('Target user not found');
    }

    const configResult = getUserConfig(targetUserId, targetUser);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Board configuration incomplete');
    }

    if (!validateReactionPermissionWithPreloadedData(actorEmail, targetUser, config)) {
      return createErrorResponse('Access denied to target board');
    }

    const rowNumber = typeof rowIndex === 'string'
      ? parseInt(rowIndex.replace('row_', ''), 10)
      : parseInt(rowIndex, 10);

    if (!rowNumber || rowNumber < 2) {
      return createErrorResponse('Invalid row ID');
    }

    const lockKey = `reaction_${config.spreadsheetId}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    if (cache.get(lockKey)) {
      return createErrorResponse('同時リアクション処理中です。お待ちください。');
    }

    const lock = LockService.getScriptLock();

    try {
      cache.put(lockKey, actorEmail, 15);

      if (!lock.tryLock(15000)) {
        return createErrorResponse('同時処理中です。少し待ってから再度お試しください。');
      }

      const dataAccess = openSpreadsheet(config.spreadsheetId, {
        useServiceAccount: false,
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

      const result = processReactionDirect(sheet, rowNumber, reactionType, actorEmail);

      return {
        success: true,
        reactions: result.reactions,
        userReaction: result.userReaction,
        action: result.action,
        message: result.action === 'added' ? 'リアクションを追加しました' : 'リアクションを削除しました'
      };

    } finally {
      try {
        lock.releaseLock();
      } catch (unlockError) {
        console.warn('addReaction: Lock release failed:', unlockError.message);
      }

      try {
        cache.remove(lockKey);
      } catch (cacheError) {
        console.warn('addReaction: Cache cleanup failed:', cacheError.message);
      }
    }

  } catch (error) {
    console.error('addReaction error:', error.message);
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
    const isAdmin = isAdministrator(actorEmail);
    const preloadedAuth = { email: actorEmail, isAdmin };

    const targetUser = findUserById(targetUserId, {
      requestingUser: actorEmail,
      preloadedAuth,
      allowPublishedRead: true
    });
    if (!targetUser) {
      return createErrorResponse('Target user not found');
    }

    const configResult = getUserConfig(targetUserId, targetUser);
    const config = configResult.success ? configResult.config : {};
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Board configuration incomplete');
    }

    if (!validateReactionPermissionWithPreloadedData(actorEmail, targetUser, config)) {
      return createErrorResponse('Access denied to target board');
    }

    const rowNumber = typeof rowIndex === 'string'
      ? parseInt(rowIndex.replace('row_', ''), 10)
      : parseInt(rowIndex, 10);

    if (!rowNumber || rowNumber < 2) {
      return createErrorResponse('Invalid row ID');
    }

    const lockKey = `highlight_${config.spreadsheetId}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    if (cache.get(lockKey)) {
      return createErrorResponse('同時ハイライト処理中です。お待ちください。');
    }

    const lock = LockService.getScriptLock();

    try {
      cache.put(lockKey, actorEmail, 15);

      if (!lock.tryLock(15000)) {
        return createErrorResponse('同時処理中です。少し待ってから再度お試しください。');
      }

      const dataAccess = openSpreadsheet(config.spreadsheetId, {
        useServiceAccount: false,
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

      const result = processHighlightDirect(sheet, rowNumber);

      return {
        success: true,
        highlighted: result.highlighted,
        message: result.highlighted ? 'ハイライトしました' : 'ハイライトを解除しました'
      };

    } finally {
      try {
        lock.releaseLock();
      } catch (unlockError) {
        console.warn('toggleHighlight: Lock release failed:', unlockError.message);
      }

      try {
        cache.remove(lockKey);
      } catch (cacheError) {
        console.warn('toggleHighlight: Cache cleanup failed:', cacheError.message);
      }
    }

  } catch (error) {
    console.error('toggleHighlight error:', error.message);
    return createExceptionResponse(error);
  }
}
