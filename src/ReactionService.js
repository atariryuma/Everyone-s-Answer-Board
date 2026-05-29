/**
 * @fileoverview ReactionService - リアクション (UNDERSTAND/LIKE/CURIOUS) と
 *   ハイライト機能。viewer/editor で権限分離（canActOnTargetBoard）。
 */

/* global getCurrentEmail, findPublishedBoardOwner, getConfigOrDefault, openSpreadsheet, createErrorResponse, createExceptionResponse, isAdministrator, invalidateSheetHeadersCache, bumpBoardDataVersion_, isBoardCollaborator, logError_, sameEmail_ */

// TTL は process() (sheet read→modify→write の RMW) の最悪ケースより長く取る。
// 旧値 10s は、 process 内の Sheets API が 429 backoff (最大 ~60s) を踏むと lock が
// 自然失効し、 同 row への並行 write が両方通って read-modify-write が lost する穴があった。
// 35s は通常 process(~200ms-2s)+1 retry を十分カバーしつつ、 6 分で execution が kill
// された場合でも遠からず自動失効する妥協点。
const ROW_LOCK_TTL_SECONDS = 35;
const ROW_LOCK_ACQUIRE_TIMEOUT_MS = 800;  // ScriptLock critical section の最大待機時間

/**
 * row 単位の lock を ScriptCache + 短い ScriptLock critical section で acquire する。
 *
 * 設計:
 *   - ScriptCache を「lock の真の保持者」 にし、 ScriptLock は単に
 *     「cache.get/cache.put をアトミックに実行する」 ためだけに使う
 *   - critical section は ~5-10ms (cache 2 操作のみ) なので、 1000+ req/sec を捌ける
 *   - 旧設計 (process() 内まで ScriptLock 保持) は ~200ms × N で全 board が serialize
 *
 * 戻り値: true で取得成功、 false で同 row 競合 (caller は error response を返す)
 */
function acquireRowLock_(cache, lockKey, actorEmail) {
  // Fast path: cache hit → 既に lock 中なので即 reject (ScriptLock 取得すらしない)
  if (cache.get(lockKey)) return false;

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(ROW_LOCK_ACQUIRE_TIMEOUT_MS)) {
    // ScriptLock 取れない = 他 row の lock acquire と競合中。 競合は ms オーダーで完了するので
    // ここで諦めるのではなく、 caller に reject させて user に retry を促す。
    return false;
  }
  try {
    // ScriptLock 配下で再 check → put のアトミック実行
    if (cache.get(lockKey)) return false;
    cache.put(lockKey, actorEmail, ROW_LOCK_TTL_SECONDS);
    return true;
  } finally {
    try { lock.releaseLock(); } catch (_) { /* ignore */ }
  }
}

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
 */
function canActOnTargetBoard(actorEmail, targetUser, config, options = {}) {
  if (!actorEmail || !targetUser) return false;

  const isAdmin = typeof options.isAdmin === 'boolean'
    ? options.isAdmin
    : isAdministrator(actorEmail);
  if (isAdmin) return true;

  if (sameEmail_(targetUser.userEmail, actorEmail)) return true;

  // v2856+: collaborator check は editor 操作 (highlight 等) でのみ実行。
  //   viewer の通常リアクション (requireEditor!==true) では published flag だけで判定。
  //   これで 700 生徒 × polling の hot path で無駄な Drive API trigger を回避。
  if (options.requireEditor === true) {
    return typeof isBoardCollaborator === 'function'
      && isBoardCollaborator(targetUser, actorEmail);
  }

  return Boolean(config && config.isPublished);
}

/**
 * リアクション処理。ヘッダー取得失敗時は graceful degradation
 * （`unavailable` レスポンスを返してリアクション機能だけ一時無効化）。
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {string} reactionType
 * @param {string} actorEmail
 * @param {Array} [preloadedHeaders] - 呼び出し側でキャッシュ済みなら渡す（再取得を回避）
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

  // Why (lazy provisioning, 2026-05-14):
  //   旧来は「列が無ければ throw → 教師にデータソース再接続を要求」する UX 破壊フロー。
  //   ログ調査 (12h で 10+ 件の UNDERSTAND not found エラー) で「新規シートに対する初回
  //   リアクションが必ず失敗」が判明。
  //   修正: 不足している reaction 列をその場で append し、headers を更新して処理続行。
  //   side-effect は新規列追加のみ（既存データには触れない）。冪等で何度呼んでも安全。
  const reactionColumns = {};
  let missingTypes = [];

  reactionTypes.forEach(type => {
    const colIndex = headers.findIndex(header => String(header).toUpperCase().trim() === type);
    if (colIndex === -1) {
      missingTypes.push(type);
    } else {
      reactionColumns[type] = colIndex + 1;
    }
  });

  // Why (scope): 「現在の操作に必要な列」だけを provision する。HIGHLIGHT は
  //   processHighlightDirect 側で別途 lazy-provision するので、ここでは reaction 3 種のみ。
  //   余計な provisioning は test の sheet._writes 期待値を壊し、生産現場でも不必要な書込み。
  const missingForProvision = missingTypes.slice();

  if (missingForProvision.length > 0) {
    // Why fresh header re-read: 呼び出し側 (executeBoardRowOperation) は preloadedHeaders
    // を 10 分 cache から取得していて、別スレッドが直前に provisioning 済みでも
    // cache が stale だと「missing」と誤判定し、重複列を append してしまう。
    // LockService 取得後にここを通るが、cache 上の headers は別途古いままなので
    // sheet から fresh に読み直して「本当に missing か」を再確認する。
    const freshHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
    const reallyMissing = [];
    missingForProvision.forEach(type => {
      const freshIdx = freshHeaders.findIndex(h => String(h).toUpperCase().trim() === type);
      if (freshIdx === -1) {
        reallyMissing.push(type);
      } else {
        // 直前の thread が既に provision 済 → 既存列を使う
        reactionColumns[type] = freshIdx + 1;
        // local headers にも反映 (以降のロジックのため)
        while (headers.length <= freshIdx) headers.push('');
        headers[freshIdx] = type;
      }
    });

    if (reallyMissing.length > 0) {
      console.warn('processReactionDirect: lazy-provisioning reaction columns', {
        reallyMissing,
        currentHeaderCount: freshHeaders.length,
        rowNumber
      });
      try {
        // 既存最終列の右側に append。setValues 1 回で済むよう一括書き込み。
        const startCol = freshHeaders.length + 1;
        sheet.getRange(1, startCol, 1, reallyMissing.length)
             .setValues([reallyMissing]);
        // headers 配列も同期更新（同じ関数内の以降のロジックが反映を見られるように）
        reallyMissing.forEach((name, i) => {
          while (headers.length < startCol + i) headers.push('');
          headers.push(name);
          if (reactionTypes.includes(name)) {
            reactionColumns[name] = startCol + i;
          }
        });
        // ヘッダー cache を invalidate (他の concurrent callers が stale を見ないように)
        if (typeof invalidateSheetHeadersCache === 'function') {
          try { invalidateSheetHeadersCache(sheet.getParent().getId(), sheet.getName()); }
          catch (_) { /* best effort */ }
        }
      } catch (provError) {
        logError_('processReactionDirect.provisioning', provError, { reallyMissing });
        throw new Error(`リアクション列の追加に失敗しました: ${provError.message}`);
      }
    }
    // 不足解消したので missingTypes をクリア
    missingTypes = missingTypes.filter(t => reactionColumns[t] === undefined);
  }

  if (missingTypes.length > 0) {
    // ここに来るのは provisioning 後にも reactionColumns に欠けがある異常ケースのみ
    logError_('processReactionDirect', new Error('missing reaction columns after provisioning'), {
      missingTypes,
      availableHeaders: headers.map(h => String(h || '').trim()).filter(h => h).join(', ')
    });
    throw new Error(`リアクション列「${missingTypes[0]}」の作成に失敗しました。`);
  }

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

  // reaction セルのみを個別に write する。
  // 旧実装は minCol..maxCol の span を 1 回の setValues で書き戻していたため、 reaction 列が
  // 非連続 (教師が手動で列を並べ替え、 間に name/class 等が挟まる) だと、 read 時点の
  // 中間列の stale 値で上書きし、 その間に走った非 reaction 経路の更新を飲み込む恐れがあった。
  // 個別 write なら reaction 3 列以外には一切触れない (最大 3 RPC、 row lock 下で十分軽量)。
  reactionTypes.forEach(type => {
    const col = reactionColumns[type];
    const users = updatedReactions[type];
    const serialized = Array.isArray(users) && users.length > 0
      ? users.filter(email => email && email.trim().length > 0).join('|')
      : '';
    sheet.getRange(rowNumber, col).setValue(serialized);
  });

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
 * ハイライト処理。ヘッダー取得失敗時は graceful degradation。
 * @param {Sheet} sheet
 * @param {number} rowNumber
 * @param {Array} [preloadedHeaders]
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

  let highlightColIndex = headers.findIndex(header => String(header).toUpperCase().trim() === 'HIGHLIGHT');

  // Why (lazy provisioning): 旧来は throw して教師にデータソース再接続を強制していたが、
  //   新規ボードで初回 HIGHLIGHT 押下時に必ず失敗する UX 破壊フロー。
  //   processReactionDirect と同じ方針で missing column を sheet 末尾に append → 続行。
  if (highlightColIndex === -1) {
    console.warn('processHighlightDirect: lazy-provisioning HIGHLIGHT column', {
      currentHeaderCount: headers.length,
      rowNumber
    });
    try {
      const newCol = headers.length + 1;
      sheet.getRange(1, newCol, 1, 1).setValues([['HIGHLIGHT']]);
      headers.push('HIGHLIGHT');
      highlightColIndex = newCol - 1;
    } catch (provError) {
      logError_('processHighlightDirect.provisioning', provError);
      throw new Error(`ハイライト列の追加に失敗しました: ${provError.message}`);
    }
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
 * リアクション送信（マルチテナント対応）。
 * @param {string} targetUserId - ボード所有者の userId
 * @param {number|string} rowIndex - 行番号または 'row_#'
 * @param {string} reactionType
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
 * ハイライト切り替え（マルチテナント対応）。
 * @param {string} targetUserId - ボード所有者の userId
 * @param {number|string} rowIndex - 行番号または 'row_#'
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
    // accessMode 自動判定: owner = openById、viewer/admin = SA pool 経由。
    const dataAccess = openSpreadsheet(config.spreadsheetId, {
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

    // per-row lock (CAS-style)。 旧設計の `getScriptLock()` 全体 serialize は 700 人 burst で
    // 全 board がスタックする bottleneck。 新設計は:
    //   1) **ScriptLock を短い critical section** (cache check + put のみ ~5ms) で「同 row
    //      のロック取得」 をアトミックに行う
    //   2) 取得後は ScriptLock を **即座に release** し、 process() (~200ms) は lock 外で実行
    //   3) 異なる row 同士は完全並列、 同 row のみシリアライズ
    // 結果: throughput は ~5 req/sec (旧) → ~200 req/sec (新) に向上。
    const lockKey = `${lockKeyPrefix}_${config.spreadsheetId}_${rowNumber}`;
    const cache = CacheService.getScriptCache();

    const acquired = acquireRowLock_(cache, lockKey, actorEmail);
    if (!acquired) return createErrorResponse(concurrentMessage);

    try {
      // Race protection: lock acquire 中に教師が unpublish した可能性がある。 lock 取得後
      // に publish 状態を再確認し、 viewer (= non-admin non-owner) からの write を再度 gate。
      // owner/admin は publish 状態に関わらず編集可能 (canActOnTargetBoard と同じ判定基準)。
      if (!isAdmin && !sameEmail_(targetUser.userEmail, actorEmail)) {
        const refreshedUser = findPublishedBoardOwner(targetUserId, actorEmail, { preloadedAuth });
        if (!refreshedUser) {
          return createErrorResponse('ボードの公開が終了しました');
        }
      }
      const result = process(sheet, rowNumber, actorEmail, preloadedHeaders);
      // board data cache を即時 stale 化 (viewer の次 polling で fresh fetch)。
      if (typeof bumpBoardDataVersion_ === 'function') {
        try { bumpBoardDataVersion_(targetUserId); } catch (_) { /* ignore */ }
      }
      return formatSuccess(result);
    } finally {
      try { cache.remove(lockKey); } catch (e) { console.warn(`${label}: Cache cleanup failed:`, e.message); }
    }
  } catch (error) {
    logError_(label, error);
    return createExceptionResponse(error);
  }
}
