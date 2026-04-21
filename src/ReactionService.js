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

/* global getCurrentEmail, findUserBySpreadsheetId, findUserById, findPublishedBoardOwner, getUserConfig, getConfigOrDefault, openSpreadsheet, createErrorResponse, createExceptionResponse, CACHE_DURATION, SYSTEM_LIMITS, isAdministrator */

const LOCK_TIMEOUT_MS = 10000;
const LOCK_CACHE_TTL_SECONDS = 10;





/**
 * 対象ボードに対する操作権限の検証（事前読み込みデータ版）
 *
 * 操作の性質で 2 種類の権限モデルを分離する:
 *   - viewer 操作（addReaction 等）: admin | owner | 公開ボードの閲覧者
 *   - editor 操作（toggleHighlight 等）: admin | owner のみ
 *
 * Why: 以前は両方とも「config.isPublished なら誰でも可」で、
 *      UI 側で highlight-btn を隠しているだけだった。生徒が
 *      google.script.run.toggleHighlight(...) を直接叩くと先生の
 *      ハイライトを自由に切り替えられる security bug を塞ぐ。
 *
 * Perf: isAdmin を options.isAdmin で受け取れるようにして、呼び出し側で
 *       既に計算済みの値を再利用する（isAdministrator のキャッシュヒットでも
 *       関数コール自体を省ける）。
 *
 * @param {string} actorEmail
 * @param {Object} targetUser
 * @param {Object} config
 * @param {Object} [options]
 * @param {boolean} [options.requireEditor=false] editor 権限を要求するか
 * @param {boolean} [options.isAdmin] 呼び出し側が計算済みの isAdmin（省略時は再計算）
 * @returns {boolean}
 */
function canActOnTargetBoard(actorEmail, targetUser, config, options = {}) {
  if (!actorEmail || !targetUser) return false;

  const isAdmin = typeof options.isAdmin === 'boolean'
    ? options.isAdmin
    : isAdministrator(actorEmail);
  if (isAdmin) return true;

  if (targetUser.userEmail === actorEmail) return true;

  if (options.requireEditor === true) return false;

  return Boolean(config && config.isPublished);
}


/**
 * 🚀 GAS-Native直接リアクション処理
 * ✅ Graceful Degradation: ヘッダー取得失敗時も継続動作
 * @param {Sheet} sheet - スプレッドシートオブジェクト
 * @param {number} rowNumber - 行番号
 * @param {string} reactionType - リアクション種類
 * @param {string} actorEmail - 操作者メール
 * @param {Array} [preloadedHeaders] - 呼び出し側がキャッシュ経由で取得済みのヘッダー。
 *                                     省略時は互換性のためシートから直接取得する。
 * @returns {Object} 処理結果
 */
function processReactionDirect(sheet, rowNumber, reactionType, actorEmail, preloadedHeaders) {
  const reactionTypes = ['UNDERSTAND', 'LIKE', 'CURIOUS'];

  if (!reactionTypes.includes(reactionType)) {
    throw new Error('Invalid reaction type');
  }

  // Why: ヘッダーはセッション中ほぼ変化しないため、呼び出し側がキャッシュしたものを
  //      渡してくれれば sheet.getDataRange().getValues() の全件読取を回避できる。
  //      preloadedHeaders 未指定時は従来通りシートから読む（互換動作）。
  const headers = Array.isArray(preloadedHeaders) && preloadedHeaders.length > 0
    ? preloadedHeaders
    : (sheet.getDataRange().getValues()[0] || []);

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
 * @param {Array} [preloadedHeaders] - 呼び出し側がキャッシュ経由で取得済みのヘッダー。
 * @returns {Object} 処理結果
 */
function processHighlightDirect(sheet, rowNumber, preloadedHeaders) {
  const headers = Array.isArray(preloadedHeaders) && preloadedHeaders.length > 0
    ? preloadedHeaders
    : (sheet.getDataRange().getValues()[0] || []);

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
  return executeBoardRowOperation({
    targetUserId,
    rowIndex,
    lockKeyPrefix: 'reaction',
    label: 'addReaction',
    openContext: 'reaction_processing',
    concurrentMessage: '同時リアクション処理中です。お待ちください。',
    process: (sheet, rowNumber, actorEmail, preloadedHeaders) =>
      processReactionDirect(sheet, rowNumber, reactionType, actorEmail, preloadedHeaders),
    formatSuccess: (result) => ({
      success: true,
      reactions: result.reactions,
      userReaction: result.userReaction,
      action: result.action,
      message: result.action === 'added' ? 'リアクションを追加しました' : 'リアクションを削除しました'
    })
  });
}

/**
 * ハイライト切り替え（マルチテナント対応・GAS-Native）
 * @param {string} targetUserId - 対象ユーザー（ボード所有者）ID
 * @param {number|string} rowIndex - 行番号または'row_#'
 * @returns {Object} 処理結果
 */
function toggleHighlight(targetUserId, rowIndex) {
  return executeBoardRowOperation({
    targetUserId,
    rowIndex,
    lockKeyPrefix: 'highlight',
    label: 'toggleHighlight',
    openContext: 'highlight_processing',
    concurrentMessage: '同時ハイライト処理中です。お待ちください。',
    // Why: UI で highlight-btn を非エディターに出していないだけでは不十分。
    //      google.script.run を直接呼ばれると生徒がハイライトを操作できるので、
    //      サーバー側で editor 限定のゲートを張る。
    requireEditor: true,
    process: (sheet, rowNumber, _actorEmail, preloadedHeaders) =>
      processHighlightDirect(sheet, rowNumber, preloadedHeaders),
    formatSuccess: (result) => ({
      success: true,
      highlighted: result.highlighted,
      message: result.highlighted ? 'ハイライトしました' : 'ハイライトを解除しました'
    })
  });
}

/**
 * addReaction / toggleHighlight の共通実行フレーム。
 *
 * 1. 認証・ユーザー/設定/権限の検証
 * 2. rowIndex パース
 * 3. Spreadsheet をロック外でオープン（読取のみの重い処理を競合から外す）
 * 4. CacheService による同時実行の早期検出
 * 5. LockService.tryLock → process() → formatSuccess() → finally でロック/キャッシュ解放
 *
 * @param {Object} options
 * @param {string} options.targetUserId - ボード所有者のuserId
 * @param {number|string} options.rowIndex - 行番号または 'row_#' 形式
 * @param {string} options.lockKeyPrefix - キャッシュキーのプレフィックス ('reaction' | 'highlight')
 * @param {string} options.label - ログ/エラー文脈名（'addReaction' | 'toggleHighlight'）
 * @param {string} options.openContext - openSpreadsheet の context 識別子
 * @param {string} options.concurrentMessage - cache ヒット時に返すユーザー向けメッセージ
 * @param {function(sheet, rowNumber, actorEmail): Object} options.process - クリティカルセクション内で実行する処理
 * @param {function(result): Object} options.formatSuccess - process の戻り値から API レスポンスを組み立て
 * @param {boolean} [options.requireEditor] editor 権限を要求する（true: ハイライト等の editor-only 操作）
 * @returns {Object} API レスポンス
 */
function executeBoardRowOperation(options) {
  const {
    targetUserId, rowIndex, lockKeyPrefix, label, openContext,
    concurrentMessage, process, formatSuccess, requireEditor
  } = options;
  const actorEmail = getCurrentEmail();

  try {
    const isAdmin = isAdministrator(actorEmail);
    const preloadedAuth = { email: actorEmail, isAdmin };

    const targetUser = findPublishedBoardOwner(targetUserId, actorEmail, { preloadedAuth });
    if (!targetUser) return createErrorResponse('Target user not found');

    const config = getConfigOrDefault(targetUserId, targetUser);
    if (!config.spreadsheetId || !config.sheetName) {
      return createErrorResponse('Board configuration incomplete');
    }

    if (!canActOnTargetBoard(actorEmail, targetUser, config, {
      isAdmin,
      requireEditor: requireEditor === true
    })) {
      return createErrorResponse('Access denied to target board');
    }

    const rowNumber = typeof rowIndex === 'string'
      ? parseInt(rowIndex.replace('row_', ''), 10)
      : parseInt(rowIndex, 10);
    if (!rowNumber || rowNumber < 2) return createErrorResponse('Invalid row ID');

    // Spreadsheetのオープンとシート取得は読取操作で競合しない。
    // ロック保持時間を最小化するためクリティカルセクション外で実行する。
    const dataAccess = openSpreadsheet(config.spreadsheetId, {
      useServiceAccount: false,
      context: openContext
    });
    if (!dataAccess) return createErrorResponse('Failed to access target spreadsheet');
    const sheet = dataAccess.spreadsheet.getSheetByName(config.sheetName);
    if (!sheet) return createErrorResponse('Target sheet not found');

    // Load sheet headers outside the lock using the cached helper.
    // Why: processReactionDirect / processHighlightDirect previously called
    //      sheet.getDataRange().getValues() every invocation — that reads the
    //      entire sheet just to get the header row. By resolving headers here
    //      (via getSheetHeaders, 10-minute cache) we pay that cost at most
    //      once per 10 minutes per sheet and keep the critical section lean.
    let preloadedHeaders = null;
    try {
      if (typeof getSheetHeaders === 'function') {
        const info = getSheetHeaders(sheet);
        preloadedHeaders = info && info.headers ? info.headers : null;
      }
    } catch (headerError) {
      // Fall back to process()'s built-in sheet read if the cache helper throws.
      console.warn(`${label}: getSheetHeaders failed, falling back to inline read:`, headerError.message);
    }

    const lockKey = `${lockKeyPrefix}_${config.spreadsheetId}_${rowNumber}`;
    const cache = CacheService.getScriptCache();
    if (cache.get(lockKey)) return createErrorResponse(concurrentMessage);

    const lock = LockService.getScriptLock();
    try {
      cache.put(lockKey, actorEmail, LOCK_CACHE_TTL_SECONDS);
      if (!lock.tryLock(LOCK_TIMEOUT_MS)) {
        return createErrorResponse('同時処理中です。少し待ってから再度お試しください。');
      }
      const result = process(sheet, rowNumber, actorEmail, preloadedHeaders);
      return formatSuccess(result);
    } finally {
      try { lock.releaseLock(); } catch (e) { console.warn(`${label}: Lock release failed:`, e.message); }
      try { cache.remove(lockKey); } catch (e) { console.warn(`${label}: Cache cleanup failed:`, e.message); }
    }
  } catch (error) {
    console.error(`${label} error:`, error.message);
    return createExceptionResponse(error);
  }
}
