/**
 * @fileoverview DatabaseCore - users シートの CRUD と、スプレッドシートを開くときのアクセス判定。
 *
 * 先生専用の配置 (先生として実行) なので、読み書きはすべて SpreadsheetApp.openById で行う
 *   (実行者 = 先生 = すべてのシートの持ち主)。サービスアカウントも Sheets API の直叩きも使わない。
 *   誰が何を見てよいかだけを checkSpreadsheetAccess_ が決める。
 */

/* global validateEmail, CACHE_DURATION, getCurrentEmail, isAdministrator, getUserConfig, executeWithRetry_, getCachedProperty_, clearPropertyCache_, simpleHash, saveToCacheWithSizeCheck_, DEFAULT_DISPLAY_SETTINGS, safeJsonParse_, logError_, sameEmail_, withStampedeGuard_, acquireCacheLock_, releaseCacheLock_, packJsonForCache_, unpackJsonFromCache_ */

// users シート書き込みの key 単位 lock (acquireCacheLock_ / helpers.js)。
const USERS_LOCK_KEYS_ = Object.freeze({
  ROW: 'users_row_lock:',        // + userId : 1 ユーザー行の read-modify-write
  CREATE: 'users_create_lock:',  // + email  : 同じ email の二重作成防止
  STRUCT: 'users_struct_lock'    //          : 行削除 (行番号がずれる) 中は誰も書かない
});

// アクセス判定の短期 cache。key に「版」を含め、公開状態が変わったら版を上げて全 viewer 分を古くする。
const ACCESS_CACHE_KEYS_ = Object.freeze({
  VERSION: 'access_ver:',   // + spreadsheetId
  DECISION: 'access_'       // + 版 + spreadsheetId (先生以外で共有)
});

/**
 * 公開状態が変わったら呼ぶ。そのスプレッドシートのアクセス判定の cache (60 秒) を全 viewer 分古くする。
 *   Why: これがないと、公開を終了した直後の最大 60 秒、児童がまだ読めてしまう。
 *   ScriptCache は wildcard で消せないので、版の番号を上げる (cache key に版を含めている)。
 */
function invalidateAccessDecisionCache_(spreadsheetId) {
  if (!spreadsheetId || typeof CacheService === 'undefined') return;
  try {
    const cache = CacheService.getScriptCache();
    const key = ACCESS_CACHE_KEYS_.VERSION + spreadsheetId;
    const cur = Number(cache.get(key) || 0);
    cache.put(key, String(cur + 1), 3600);
  } catch (_) { /* cache 不通なら判定 cache も使えないので、古い判定が残る心配もない */ }
}

/**
 * このスプレッドシートを、いまの呼び出し元に開かせてよいか。
 *
 *   | 呼び出し元           | 公開中のボードが指すシート | 非公開のボードのシート | ボードに無いシート |
 *   |----------------------|----------------------------|------------------------|--------------------|
 *   | 先生 (管理者)        | 可                         | 可                     | 可 (作ったばかりのフォームの回答先など) |
 *   | 児童・ほかの先生      | 可                         | 不可                   | 不可               |
 *
 *   DB (先生のスプレッドシート) も同じ扱い。DB には授業モードの回答シートが全部入るので、
 *   児童が開けるのは授業モードの授業がボードにつながって公開されている間だけ。users / lessons を
 *   読むサーバの処理は openDatabase で直接開く (ここを通らない)。
 *
 * @param {string} spreadsheetId
 * @param {string} [context] - ログ識別用
 * @returns {{allowed:boolean, reason:string}}
 */
function checkSpreadsheetAccess_(spreadsheetId, context = 'unknown') {
  try {
    const currentEmail = getCurrentEmail();
    if (currentEmail && isAdministrator(currentEmail)) {
      return { allowed: true, reason: 'teacher' };
    }

    let cache = null;
    try { cache = CacheService.getScriptCache(); } catch (_) { /* cache 不通は機能的に無害 (毎回判定する) */ }
    let version = '0';
    if (cache) {
      try { version = String(cache.get(ACCESS_CACHE_KEYS_.VERSION + spreadsheetId) || '0'); } catch (_) { /* 版が読めなければ 0 */ }
    }
    // 先生以外の答えはシートだけで決まる (公開中のボードが指しているか) ので、児童ごとに分けず 1 つを共有する。
    //   Why: 児童ごとに覚えると 30 人分が別々に切れ、そのたびに users の一覧を開いて全員の設定を読んでいた。
    const cacheKey = `${ACCESS_CACHE_KEYS_.DECISION}v${version}_${spreadsheetId}`;
    if (cache) {
      try {
        const hit = cache.get(cacheKey);
        if (hit) return JSON.parse(hit);
      } catch (_) { /* 読めなければ判定し直す */ }
    }
    // 許可は 60 秒、不許可は 5 秒だけ覚える。不許可を長く覚えると、ボードを公開した直後も
    //   児童の画面が「読めない」のまま 1 分続く (公開の bump は許可側の古さだけを消す)。
    const remember = (result) => {
      if (cache) {
        try { cache.put(cacheKey, JSON.stringify(result), result.allowed ? 60 : 5); } catch (_) { /* 覚えられなくても判定は正しい */ }
      }
      return result;
    };

    const owner = findUserBySpreadsheetId_(spreadsheetId, { skipCache: true });
    if (!owner) {
      console.warn('checkSpreadsheetAccess_: どのボードにも入っていないシート:', spreadsheetId.substring(0, 8), context);
      return remember({ allowed: false, reason: 'not on any board' });
    }
    if (!isUserBoardPublished(owner)) {
      console.warn('checkSpreadsheetAccess_: 非公開のボード:', {
        currentEmail: currentEmail ? `${currentEmail.split('@')[0]}@***` : 'unknown',
        spreadsheetId: spreadsheetId.substring(0, 8),
        context
      });
      return remember({ allowed: false, reason: 'board not published' });
    }
    return remember({ allowed: true, reason: 'published board' });
  } catch (error) {
    logError_('checkSpreadsheetAccess_', error);
    return { allowed: false, reason: 'access check failed' };
  }
}

/**
 * データベース (先生のスプレッドシート) を開く。
 * @returns {GoogleAppsScript.Spreadsheet.Spreadsheet|null}
 */
function openDatabase() {
  const dbId = getCachedProperty_('DATABASE_SPREADSHEET_ID');
  if (!dbId) {
    console.warn('openDatabase: DATABASE_SPREADSHEET_ID not configured');
    return null;
  }
  try {
    return SpreadsheetApp.openById(dbId);
  } catch (error) {
    logError_('openDatabase', error);
    return null;
  }
}

/**
 * スプレッドシートを開く。開かせてよいかは checkSpreadsheetAccess_ が決める。
 * @param {string} spreadsheetId
 * @param {Object} [options]
 * @param {string} [options.context] - ログ識別用
 * @returns {{spreadsheet, getSheet}|null}
 */
function openSpreadsheet(spreadsheetId, options = {}) {
  if (!spreadsheetId || typeof spreadsheetId !== 'string') {
    console.warn('openSpreadsheet: Invalid spreadsheet ID');
    return null;
  }
  const access = checkSpreadsheetAccess_(spreadsheetId, options.context || 'openSpreadsheet');
  if (!access.allowed) {
    console.warn('openSpreadsheet: access denied:', access.reason);
    return null;
  }
  let spreadsheet;
  try {
    spreadsheet = SpreadsheetApp.openById(spreadsheetId);
  } catch (openError) {
    // 呼出元 (getUserSheetData_ 等) が 1 行 ERROR に集約する前提で WARN に降格。
    console.warn('openSpreadsheet: スプレッドシート接続失敗', {
      spreadsheetId: `${spreadsheetId.substring(0, 20)}...`,
      error: openError.message
    });
    return null;
  }
  return {
    spreadsheet,
    getSheet(sheetName) {
      if (!sheetName) {
        console.warn('openSpreadsheet.getSheet: Sheet name required');
        return null;
      }
      try {
        return spreadsheet.getSheetByName(sheetName);
      } catch (error) {
        console.warn(`openSpreadsheet.getSheet: Failed to get sheet "${sheetName}":`, error.message);
        return null;
      }
    }
  };
}

/**
 * ユーザー検索の呼び出し元メールアドレスを解決
 * @param {Object} context - アクセスコンテキスト
 * @returns {string|null} リクエストユーザーのメール
 */
function resolveRequestingUser_(context = {}) {
  if (context.requestingUser && typeof context.requestingUser === 'string' && context.requestingUser.trim()) {
    return context.requestingUser.trim();
  }
  try {
    return getCurrentEmail();
  } catch (error) {
    logError_('resolveRequestingUser', error);
    return null;
  }
}

/**
 * 対象ユーザーのボード公開状態を判定
 * @param {Object} targetUser - 対象ユーザー
 * @returns {boolean} 公開中かどうか
 */
function isUserBoardPublished(targetUser) {
  if (!targetUser) {
    return false;
  }

  try {
    if (targetUser.configJson && typeof targetUser.configJson === 'string') {
      const parsed = JSON.parse(targetUser.configJson);
      return Boolean(parsed && parsed.isPublished);
    }
  } catch (parseError) {
    console.warn('isUserBoardPublished: Failed to parse configJson:', parseError.message);
  }

  if (!targetUser.userId) {
    return false;
  }

  try {
    const configResult = getUserConfig(targetUser.userId, targetUser);
    return Boolean(configResult && configResult.success && configResult.config && configResult.config.isPublished);
  } catch (configError) {
    console.warn('isUserBoardPublished: Failed to resolve config:', configError.message);
    return false;
  }
}

/**
 * ユーザー検索のアクセス可否を判定
 * @param {Object} targetUser - 対象ユーザー
 * @param {Object} context - アクセスコンテキスト
 * @returns {boolean} アクセス許可されるか
 */
function canAccessTargetUser_(targetUser, context = {}) {
  if (!targetUser) {
    return false;
  }

  if (context.skipAccessCheck === true) {
    return true;
  }

  const requestingUser = resolveRequestingUser_(context);
  if (!requestingUser) {
    return false;
  }

  let isAdmin = false;
  if (context.preloadedAuth && typeof context.preloadedAuth.isAdmin === 'boolean') {
    isAdmin = context.preloadedAuth.isAdmin;
  } else {
    isAdmin = isAdministrator(requestingUser);
  }

  if (isAdmin) {
    return true;
  }

  if (sameEmail_(targetUser.userEmail, requestingUser)) {
    return true;
  }

  if (context.allowPublishedRead === true) {
    // 無効化されたユーザー (isActive=false) のボードは、第三者には未公開として扱う。
    //   Why: 管理者の「無効化」は公開を強制終了する意図 (main.js の doGet も同じ扱い)。
    //   ここを抜かすと doGet では止まるのに getPublishedSheetData / getNotificationUpdate /
    //   リアクション経由では読めてしまい、層ごとに結論が変わっていた。
    //   admin / 本人は上で return 済みなので、ここに来るのは第三者だけ。
    //   厳密比較にするのは、isActive 未設定 (undefined/'') の既存行を巻き込まないため。
    if (targetUser.isActive === false) return false;
    return isUserBoardPublished(targetUser);
  }

  return false;
}

/**
 * ユーザー検索結果にアクセス制御を適用
 * @param {Object|null} user - 検索結果ユーザー
 * @param {Object} context - アクセスコンテキスト
 * @param {string} operation - 操作名
 * @returns {Object|null} 許可済みユーザーまたはnull
 */
function applyUserAccessControl_(user, context = {}, operation = 'user_lookup') {
  if (!user) {
    return null;
  }

  if (canAccessTargetUser_(user, context)) {
    return user;
  }

  // 呼び出し元が「居ない」と「見せられない」を区別できるように印を付ける (中身は渡さない)。
  //   Why: 公開ボードの閲覧経路は null しか受け取れず、非公開にしただけのボードが
  //   "Target user not found" (ERROR) / 児童の画面は「ユーザーエラー」になっていた (2026-10-19)。
  if (context.accessProbe && typeof context.accessProbe === 'object') {
    context.accessProbe.denied = true;
  }

  // 公開ボードの閲覧経路 (allowPublishedRead) での拒否は「非公開 / 無効化されたボードを開いた」
  //   だけで、異常ではない。授業の直前まで非公開にする運用では毎回出るので WARN にしない。
  if (context.allowPublishedRead === true) return null;

  const requestingUser = resolveRequestingUser_(context);
  const maskedRequester = requestingUser ? `${requestingUser.split('@')[0]}@***` : 'N/A';
  const maskedTarget = user.userEmail ? `${String(user.userEmail).split('@')[0]}@***` : 'unknown';
  console.warn(`${operation}: Access denied`, {
    requestingUser: maskedRequester,
    target: maskedTarget
  });
  return null;
}

/**
 * メールアドレスでユーザーを検索（CLAUDE.md準拠 - Editor→Admin共有DB）
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} context - アクセスコンテキスト
 * @param {boolean} context.systemLookup - サーバ内部の検索 (閲覧者の権限確認をしない)
 * @param {string} context.requestingUser - リクエストユーザー（デバッグ用）
 * @returns {Object|null} User object
 */
function findUserByEmail_(email, context = {}) {
  if (!email || !validateEmail(email).isValid) return null;
  return findUserByField_('userEmail', email, { ...context, cacheKeyPrefix: 'user_by_email', label: 'findUserByEmail' });
}

/**
 * ユーザーIDでユーザーを検索（strict: admin / 自分のみ許可）
 *
 * WRITE操作や管理者専用操作で使う。公開ボードを閲覧する生徒の経路では
 * 代わりに {@link findPublishedBoardOwner_} を使うこと。これを逆に使うと
 * 生徒から null が返り、「config 取得失敗 → default config → formUrl 消失」
 * のような silent failure が起きる（実際に発生した bug）。
 *
 * @param {string} userId - ユーザーID
 * @param {Object} context - アクセスコンテキスト
 * @param {string} context.requestingUser - リクエストユーザー（debug/権限判定用）
 * @param {boolean} [context.allowPublishedRead] - 公開ボードなら非オーナーも許可
 * @param {boolean} [context.systemLookup] - サーバ内部の検索 (閲覧者の権限確認をしない)
 * @returns {Object|null} User object
 */
function findUserById_(userId, context = {}) {
  if (!userId) return null;
  return findUserByField_('userId', userId, { ...context, cacheKeyPrefix: 'user_by_id', label: 'findUserById' });
}

/**
 * 公開ボードのオーナーを viewer 視点で取得（admin / owner / 公開ボード閲覧者を許可）
 *
 * `getPublishedSheetData` / `getActiveFormInfo` / `getNotificationUpdate` など、
 * 生徒が他人のボードを見る経路で使う。これを使うと `allowPublishedRead: true` を
 * 渡し忘れる類の bug を根絶できる。
 *
 * WRITE 系や editor-only 操作では `findUserById_` + 独自の editor gate を使うこと
 * （例: deleteAnswerRow、toggleHighlight）。
 *
 * @param {string} userId - 対象ボードオーナーの userId
 * @param {string} viewerEmail - 閲覧者のメールアドレス
 * @param {Object} [extra] - 追加 context（preloadedAuth 等）
 */
function findPublishedBoardOwner_(userId, viewerEmail, extra = {}) {
  if (!userId) return null;
  return findUserById_(userId, {
    ...extra,
    requestingUser: viewerEmail,
    allowPublishedRead: true
  });
}

/**
 * GoogleアカウントID（不変ID）でユーザーを検索
 * メールアドレス変更時のフォールバック検索に使用
 * @param {string} googleId - Googleアカウントのsub値
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object|null} User object
 */
function findUserByGoogleId_(googleId, context = {}) {
  if (!googleId) return null;
  // Note: historically this path did not populate an individual cache entry, so
  // we explicitly opt out via cacheKeyPrefix=null to preserve behavior.
  return findUserByField_('googleId', googleId, { ...context, cacheKeyPrefix: null, label: 'findUserByGoogleId' });
}

/**
 * findUserBy*系の共通実装。
 * 探索は (1) individual cache → (2) getAllUsers_ キャッシュ → (3) 直接DB の順。
 * @param {string} fieldName - 検索するユーザーフィールド ('userEmail' | 'userId' | 'googleId')
 * @param {*} fieldValue - 一致させる値
 * @param {Object} context - requestingUser / preloadedAuth / systemLookup 等に加え、
 *                          内部用の cacheKeyPrefix (null で個別キャッシュ無効) と label を受け取る
 * @returns {Object|null} アクセス制御適用済みのユーザー、または null
 */
function findUserByField_(fieldName, fieldValue, context = {}) {
  const { cacheKeyPrefix = null, label = `findUserBy_${fieldName}` } = context;

  try {
    let individualCacheKey = null;
    if (cacheKeyPrefix) {
      const cacheVersion = userCacheVersion_();
      individualCacheKey = `${cacheKeyPrefix}_v${cacheVersion}_${fieldValue}`;
      try {
        const cached = CacheService.getScriptCache().get(individualCacheKey);
        if (cached) {
          const cachedUser = JSON.parse(cached);
          return applyUserAccessControl_(cachedUser, context, label);
        }
      } catch (individualCacheError) {
        logError_(label + '.individualCache', individualCacheError);
      }
    }

    try {
      const allUsers = getAllUsers_(
        { activeOnly: false },
        { ...context, systemLookup: true, skipCache: false }
      );
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find((u) => u[fieldName] === fieldValue);
        if (user) {
          const allowedUser = applyUserAccessControl_(user, context, label);
          if (!allowedUser) return null;
          if (individualCacheKey) {
            saveToCacheWithSizeCheck_(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
          }
          return allowedUser;
        }
        // 一覧 (users シート全体、activeOnly=false) が取れているのに居ないなら、居ない。
        //   Why: 以前はここから users シートを直接読み直していた (寸法 + ヘッダ + TextFinder +
        //   該当行 = 数 read)。users に居ない児童の findUserByEmail_ (checkUserAuthentication /
        //   getActiveFormInfo / validateHeaderIntegrity) が毎回それを払い、授業中の 30 人分が
        //   users シートの 429 の一因だった (2026-09-08)。一覧は users の cache の版 (user_cache_ver) 付きの
        //   cache なので、作成直後のユーザーも clearDatabaseUserCache_ の bump で次の読みから見える。
        return null;
      }
    } catch (cacheError) {
      logError_(label + '.cacheSearch', cacheError);
    }

    // Why (performance note): 旧来は sheet.getDataRange().getValues() で N×7 配列を
    //   キャッシュミス毎にロードしていた。SLO 観点では現在の users 数 (<100) では
    //   許容範囲だが、将来 1000 users 超で問題化する可能性あり。
    //   その時は TextFinder API (`sheet.createTextFinder(fieldValue)`) で列限定検索に
    //   切り替えることを検討（rowIndex 取得後に該当行のみ getValues する）。
    //   現状は読み込んだ data 配列を直近 30s memoize して連続クエリの冗長読みを削減。
    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn(`${label}: Database access failed`);
      return null;
    }
    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn(`${label}: Users sheet not found`);
      return null;
    }

    // Why (perf): users 数が増えても線形に劣化しない検索パスを優先。
    //   1. ヘッダー行のみ読み込んで列 index を確定
    //   2. TextFinder で対象列のみ exact match 検索（GAS 内部最適化済み）
    //   3. ヒットした row の getValues() で 1 行のみ取得
    //   既存の全件 getValues() フォールバックは API quota 異常や TextFinder 未対応環境
    //   (テスト sandbox 等) で動作するよう残す。
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return null;
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnIndex = headerRow.indexOf(fieldName);
    if (columnIndex === -1) {
      console.warn(`${label}: ${fieldName} column not found`);
      return null;
    }

    let foundRow = null;
    try {
      if (typeof sheet.createTextFinder === 'function') {
        const finder = sheet.createTextFinder(String(fieldValue))
          .matchEntireCell(true)
          .matchCase(true);
        // 対象列にスコープを限定（A 列固定だと userId 検索が崩れるので列 index +1 で動的指定）
        if (typeof finder.findAll === 'function') {
          const range = sheet.getRange(2, columnIndex + 1, lastRow - 1, 1);
          if (typeof finder.startFrom === 'function') {
            try { finder.startFrom(range.getCell(1, 1)); } catch (_) { /* fallback below */ }
          }
          const hits = finder.findAll();
          for (const cell of hits) {
            if (cell.getColumn() === columnIndex + 1 && cell.getRow() >= 2) {
              foundRow = cell.getRow();
              break;
            }
          }
        }
      }
    } catch (finderError) {
      console.warn(`${label}: TextFinder fallback (${finderError.message})`);
    }

    if (foundRow) {
      const row = sheet.getRange(foundRow, 1, 1, headerRow.length).getValues()[0];
      const user = createUserObjectFromRow(row, headerRow);
      const allowedUser = applyUserAccessControl_(user, context, label);
      if (!allowedUser) return null;
      if (individualCacheKey) {
        saveToCacheWithSizeCheck_(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
      }
      return allowedUser;
    }

    // フォールバック: TextFinder 未対応 (test sandbox) や hits 配列形態が想定外の場合
    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return null;

    const headers = data[0];

    for (let i = 1; i < data.length; i++) {
      if (data[i][columnIndex] === fieldValue) {
        const user = createUserObjectFromRow(data[i], headers);
        const allowedUser = applyUserAccessControl_(user, context, label);
        if (!allowedUser) return null;
        if (individualCacheKey) {
          saveToCacheWithSizeCheck_(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
        }
        return allowedUser;
      }
    }

    return null;
  } catch (error) {
    logError_(label, error);
    return null;
  }
}

/**
 * 新しいユーザーを作成（CLAUDE.md準拠 - Context-Aware）
 * @param {string} email - ユーザーのメールアドレス
 * @param {Object} initialConfig - 初期設定
 * @param {Object} context - アクセスコンテキスト
 * @param {string} [context.googleId] - Googleアカウント不変ID
 * @returns {Object|null} Created user object
 */
function createUser_(email, initialConfig = {}, context = {}) {
  // email 単位の lock。別の児童の初回ログインは互いに待たない (30 人の初回ログインが
  //   1 本ずつ直列になり、10 秒で諦めた児童が「ユーザー登録に失敗」になっていた)。
  //   同じ email の二重作成 (二重タブ) だけを 1 本に絞る。構造変更 (行の削除) の lock が立っていれば待つ。
  let lockKey = null;
  let lockToken = null;

  try {
    if (!email || !validateEmail(email).isValid) {
      console.warn('createUser: Invalid email provided:', typeof email, email);
      return null;
    }

    lockKey = USERS_LOCK_KEYS_.CREATE + String(email).trim().toLowerCase();
    lockToken = acquireCacheLock_(lockKey, { ttlSec: 30, waitMs: 10000, blockedBy: [USERS_LOCK_KEYS_.STRUCT] });
    if (!lockToken) {
      console.warn('createUser: Lock timeout - concurrent user creation detected');
      return null;
    }

    const currentEmail = getCurrentEmail();

    const existingUser = findUserByEmail_(email, {
      requestingUser: currentEmail
    });
    if (existingUser) {
      return existingUser;
    }

    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('createUser: Database access failed');
      return null;
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('createUser: Users sheet not found');
      return null;
    }

    const userId = Utilities.getUuid();
    const now = new Date().toISOString();

    const defaultConfig = {
      setupStatus: 'pending',
      isPublished: false,
      displaySettings: { ...DEFAULT_DISPLAY_SETTINGS },
      ...initialConfig
    };

    const data = sheet.getDataRange().getValues();
    // Why: 読めなかったとき headers が undefined のまま進むと TypeError になり、原因が消える。
    //   戻り値の契約は user | null なので、ここで error envelope を返してはいけない
    //   (呼び出し側の `if (!user)` を通り抜けて「Invalid user data」になっていた)。
    if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0])) {
      console.warn('createUser: users sheet read returned no rows (quota?)', { email: `${String(email).split('@')[0]}@***` });
      return null;
    }
    const [headers] = data;
    const hasCreatedAtColumn = headers.indexOf('createdAt') !== -1;
    const hasGoogleIdColumn = headers.indexOf('googleId') !== -1;
    const googleId = context.googleId || '';

    const newUserRow = new Array(headers.length).fill('');
    const setCol = (name, value) => {
      const idx = headers.indexOf(name);
      if (idx !== -1) newUserRow[idx] = value;
    };
    setCol('userId', userId);
    setCol('userEmail', email);
    if (hasGoogleIdColumn) setCol('googleId', googleId);
    setCol('isActive', true);
    setCol('configJson', JSON.stringify(defaultConfig));
    if (hasCreatedAtColumn) setCol('createdAt', now);
    setCol('lastModified', now);

    // 容量判定は読んだ実データ行数で行う。
    if (data.length >= 10000) {
      logError_('createUser', new Error('Users sheet at capacity'));
      return null;
    }

    sheet.appendRow(newUserRow);

    const user = {
      userId,
      userEmail: email,
      googleId: hasGoogleIdColumn ? googleId : undefined,
      isActive: true,
      configJson: JSON.stringify(defaultConfig),
      createdAt: hasCreatedAtColumn ? now : undefined,
      lastModified: now
    };

    clearDatabaseUserCache_();

    return user;

  } catch (error) {
    logError_('createUser', error);
    return null;
  } finally {
    releaseCacheLock_(lockKey, lockToken);
  }
}

/**
 * 全ユーザーリストを取得（CLAUDE.md準拠 - Admin専用）
 * @param {Object} options - オプション設定
 * @param {Object} context - アクセスコンテキスト
 * @returns {Array} User list
 */
function getAllUsers_(options = {}, context = {}) {
  try {
    let currentEmail, isAdmin;
    if (context.preloadedAuth) {
      const { email, isAdmin: adminFlag } = context.preloadedAuth;
      currentEmail = email;
      isAdmin = adminFlag;
    } else {
      currentEmail = getCurrentEmail();
      isAdmin = isAdministrator(currentEmail);
    }

    if (!isAdmin && !context.systemLookup) {
      console.warn('getAllUsers: Non-admin user attempting cross-user data access');
      return [];
    }

    const cacheVersion = userCacheVersion_();
    const cacheSuffix = `${simpleHash(options)}_${context.systemLookup ? 'sys' : 'norm'}`;
    const cacheKey = `all_users_v${cacheVersion}_${cacheSuffix}`;
    const skipCache = context.skipCache || false;

    // users シートを実際に読む。空 ([]) は cache しない: 429 で読めなかった結果を
    //   10 分固定すると全員が "Target user not found" になる。
    const loadUsersFromSheet = () => {
      const spreadsheet = openDatabase();
      if (!spreadsheet) {
        console.warn('getAllUsers: Database access failed');
        return [];
      }

      const sheet = spreadsheet.getSheetByName('users');
      if (!sheet) {
        console.warn('getAllUsers: Users sheet not found');
        return [];
      }

      const data = sheet.getDataRange().getValues();
      if (data.length <= 1) return []; // No data or header only

      const [headers] = data;
      const users = [];

      for (let i = 1; i < data.length; i++) {
        const row = data[i];
        const user = createUserObjectFromRow(row, headers);

        if (options.activeOnly && !user.isActive) continue;
        if (options.publishedOnly) {
          const config = safeJsonParse_(user.configJson, null);
          if (!config || !config.isPublished) continue;
        }

        users.push(user);
      }
      return users;
    };

    if (skipCache) return loadUsersFromSheet();

    // Why: 100+ユーザでJSONが100KB超えるとCacheService.putが黙って失敗する。
    //      saveToCacheWithSizeCheck_でサイズ超過を検出しログに残す。
    const isCacheable = (users) => Array.isArray(users) && users.length > 0;

    // stampede 防止: config 保存 (user_cache_ver の更新) の直後は全 viewer の個別 cache も
    //   同時に miss し、30 人分の polling が一斉に users シートを読んでいた。1 件だけが読み、
    //   残りは最大 1.2 秒待って埋まった cache を読む。stale (直前の一覧) は返さない:
    //   findUserByField_ が古い configJson を個別 cache (15 分) に固定してしまうため。
    if (typeof withStampedeGuard_ === 'function') {
      return withStampedeGuard_({
        key: cacheKey,
        ttl: CACHE_DURATION.DATABASE_LONG,
        loader: loadUsersFromSheet,
        isCacheable,
        flightKey: `all_users_flight_${cacheSuffix}`,
        flightTtl: 6,
        allowStale: false,
        // users シートの読みは混雑時 1〜3 秒かかる。1.2 秒で諦めると 29 人が自分で読みに行く。
        waitMs: 500,
        waitTries: 6,
        // 一覧は gzip で書く。利用者が 30〜50 人になると素の JSON が 100KB を超えて cache に
        //   入らなくなり、その日から全員の poll が users シートを読む (人数増加で突然遅くなる)。
        pack: (typeof packJsonForCache_ === 'function') ? packJsonForCache_ : undefined,
        unpack: (typeof unpackJsonFromCache_ === 'function') ? unpackJsonFromCache_ : undefined
      });
    }

    // test 単独 load (helpers.js 不在) のための素朴な経路。本番は必ず guard を通る。
    try {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    } catch (cacheError) {
      logError_('getAllUsers.cacheRead', cacheError);
    }
    const users = loadUsersFromSheet();
    if (isCacheable(users)) saveToCacheWithSizeCheck_(cacheKey, users, CACHE_DURATION.DATABASE_LONG);
    return users;
  } catch (error) {
    logError_('getAllUsers', error);
    return [];
  }
}

/**
 * データベースユーザーキャッシュをクリア（ユーザー作成・更新・削除時）
 * シンプルなバージョニング戦略でキャッシュを無効化
 */
function clearDatabaseUserCache_() {
  try {
    CacheService.getScriptCache().put(USER_CACHE_VERSION_KEY_, newUserCacheVersion_(), 21600);
  } catch (error) {
    logError_('clearDatabaseUserCache', error);
  }
}

// users の cache の版。key に版を含め、書き込みのたびに新しい版にして全員分を一度に古くする。
//   Why CacheService か: 以前は Script Property に置いており、児童の 5 秒おきの確認のたびに
//   プロパティを読んでいた (先生として実行なので全員分が先生 1 人の 1 日の上限に数えられる)。
//   版が消えていたら (cache の追い出し) 新しい版を作る = 全員が 1 度読み直すだけで、古い値は見ない。
//   番号を足すのではなく毎回新しい値にするので、書き込みが重なっても取りこぼさない。
const USER_CACHE_VERSION_KEY_ = 'user_cache_ver';
function newUserCacheVersion_() {
  return Date.now().toString(36) + Math.random().toString(36).slice(2, 8);
}
function userCacheVersion_() {
  try {
    const cache = CacheService.getScriptCache();
    const current = cache.get(USER_CACHE_VERSION_KEY_);
    if (current) return current;
    const fresh = newUserCacheVersion_();
    cache.put(USER_CACHE_VERSION_KEY_, fresh, 21600);
    return fresh;
  } catch (_) {
    return 'nocache';   // cache 不通なら版も効かないが、そのとき個別の cache も読めないので古い値は出ない
  }
}

/**
 * 行データからユーザーオブジェクトを作成
 * @param {Array} row - スプレッドシートの行データ
 * @param {Array} headers - ヘッダー行
 * @returns {Object} User object
 */
function createUserObjectFromRow(row, headers) {
  const user = {};
  const fieldMapping = {
    'userId': 'userId',
    'userEmail': 'userEmail',
    'googleId': 'googleId',
    'isActive': 'isActive',
    'configJson': 'configJson',
    'createdAt': 'createdAt',
    'lastModified': 'lastModified'
  };

  headers.forEach((header, index) => {
    const mappedField = fieldMapping[header];
    if (mappedField) {
      user[mappedField] = row[index];
    }
  });

  if (typeof user.isActive === 'string') {
    user.isActive = user.isActive.toLowerCase() === 'true';
  }

  return user;
}

/**
 * ユーザー情報を更新（CLAUDE.md準拠 - Context-Aware）
 * @param {string} userId - ユーザーID
 * @param {Object} updates - 更新データ
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object} {success: boolean, message?: string} Operation result
 */
function updateUser_(userId, updates, context = {}) {
  // userId 単位の lock (ScriptLock 全体ではなく)。
  //   Why: 以前は ScriptLock を users シートの read → write の間 (1〜3 秒、429 なら 15 秒超) 握って
  //   いた。ScriptLock はスクリプトで 1 本なので、教師が設定を保存するたびに児童のリアクション
  //   (同じ ScriptLock を 0.8 秒だけ待つ) が全部「同時リアクション処理中」で落ちていた。
  //   他ユーザーの更新は別の行を書くので互いに待つ必要がない。
  const lockKey = USERS_LOCK_KEYS_.ROW + String(userId || '');
  let lockToken = null;

  try {
    lockToken = acquireCacheLock_(lockKey, { ttlSec: 30, waitMs: 5000, blockedBy: [USERS_LOCK_KEYS_.STRUCT] });
    if (!lockToken) {
      console.warn('updateUser: Lock timeout - concurrent update detected');
      return createErrorResponse('Concurrent update in progress. Please retry.');
    }

    const requestingUser = context.requestingUser || getCurrentEmail();
    const targetUser = findUserById_(userId, {
      ...context,
      requestingUser
    });

    if (!targetUser) {
      console.warn('updateUser: Target user not found:', userId);
      return createErrorResponse('User not found');
    }

    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('updateUser: Database access failed');
      return createErrorResponse('Database access failed');
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('updateUser: Users sheet not found');
      return createErrorResponse('Users sheet not found');
    }

    const data = sheet.getDataRange().getValues();
    // Why: 429 で retry を使い切った getValues は [] を返す。TypeError ではなく原因を返す。
    if (!Array.isArray(data) || data.length === 0 || !Array.isArray(data[0])) {
      console.warn('updateUser: users sheet read returned no rows (quota?)', { userId });
      return createErrorResponse('users シートを読み取れませんでした (Sheets API quota の可能性)。少し待って再試行してください');
    }
    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('updateUser: UserId column not found');
      return createErrorResponse('UserId column not found');
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdColumnIndex] === userId) {
        // 楽観ロック (compare-and-set) は lock の「中」で、いま読んだ行に対して行う。
        //   Why: saveUserConfig_ は etag を cache 済の user から比べていたので、2 つの保存
        //   (教師の自動保存とフェーズ切替) が同じ etag を見て両方通り、後の書き込みが前の
        //   変更を飲み込んでいた。ここで現物の configJson の etag と比べれば、負けた側は
        //   etag_mismatch を受けて読み直せる。expectedEtag が無い呼び出しは従来どおり。
        if (context.expectedEtag) {
          const cfgIdx = headers.indexOf('configJson');
          const lmIdx = headers.indexOf('lastModified');
          let currentEtag = null;
          if (cfgIdx !== -1) {
            const parsed = safeJsonParse_(data[i][cfgIdx], null);
            currentEtag = (parsed && parsed.etag) || null;
          }
          if (!currentEtag && lmIdx !== -1 && data[i][lmIdx]) currentEtag = String(data[i][lmIdx]);
          if (currentEtag && String(currentEtag) !== String(context.expectedEtag)) {
            console.warn('updateUser: etag mismatch at write time (concurrent save)', {
              userId: `${String(userId).substring(0, 8)}***`
            });
            return createErrorResponse('Configuration has been modified by another user', null,
              { error: 'etag_mismatch', currentEtag });
          }
        }

        // 内容が同じ書き込みは省く (2026-10 監査)。
        //   Why: 書き込みのたびに clearDatabaseUserCache_ が users の cache の版 (user_cache_ver) を新しくし、児童 30 人分の
        //   user cache が一斉に落ちて次の poll が users シートを読み直す (stampede guard で 1 read に
        //   抑えてはいる)。lastAccessedAt / etag しか違わない保存 (同じ内容の再保存、同じ patch の
        //   再適用) で version を上げる理由は無い。cache 済の user ではなく、いま読んだ行と比べる
        //   (別インスタンスの 30 秒 memo が古くても、行は現物)。
        if (__userRowUnchanged_(headers, data[i], updates)) {
          const cfgIdx = headers.indexOf('configJson');
          const current = cfgIdx !== -1 ? safeJsonParse_(data[i][cfgIdx], null) : null;
          return {
            success: true,
            unchanged: true,
            currentEtag: (current && current.etag) || null,
            currentConfig: current
          };
        }

        const updateCells = [];

        Object.keys(updates).forEach(field => {
          const columnIndex = headers.indexOf(field);
          if (columnIndex !== -1) {
            updateCells.push({ col: columnIndex + 1, value: updates[field] });
          }
        });

        const lastModifiedIndex = headers.indexOf('lastModified');
        if (lastModifiedIndex !== -1) {
          updateCells.push({ col: lastModifiedIndex + 1, value: new Date().toISOString() });
        }

        if (updateCells.length > 0) {
          const cols = updateCells.map(c => c.col);
          const minCol = Math.min(...cols);
          const maxCol = Math.max(...cols);
          const colSpan = maxCol - minCol + 1;

          // 既に読んだ行 (data[i]) を下敷きにする。以前はここで同じ行をもう 1 回 API で
          //   読んでいたが、429 で [] が返ると TypeError で落ち、フェーズ切替の config patch
          //   が失敗した (2026-09-08)。values API は末尾の空セルを省くので長さを揃える。
          const base = Array.isArray(data[i]) ? data[i] : [];
          const currentRowData = [];
          for (let c = 0; c < colSpan; c++) {
            const v = base[minCol - 1 + c];
            currentRowData.push(v === undefined || v === null ? '' : v);
          }

          updateCells.forEach(({ col, value }) => {
            currentRowData[col - minCol] = value;
          });

          sheet.getRange(i + 1, minCol, 1, colSpan).setValues([currentRowData]);
        }

        clearDatabaseUserCache_();

        return { success: true };
      }
    }

    console.warn('updateUser: User not found:', userId);
    return createErrorResponse('User not found');
  } catch (error) {
    logError_('updateUser', error);
    return createErrorResponse(error.message || 'Unknown error occurred');
  } finally {
    releaseCacheLock_(lockKey, lockToken);
  }
}

// updates が、いま読んだ行と内容的に同じか (updateUser_ の no-op 判定)。
//   configJson は lastAccessedAt / etag を除いて比較する (保存のたびに必ず変わる 2 項目)。
//   それ以外の列は文字列として比較。書く列が 1 つも無ければ false (従来どおり lastModified を書く)。
function __userRowUnchanged_(headers, row, updates) {
  const fields = Object.keys(updates || {}).filter((f) => headers.indexOf(f) !== -1);
  if (fields.length === 0) return false;
  for (const field of fields) {
    const cur = row[headers.indexOf(field)];
    const next = updates[field];
    if (field === 'configJson') {
      if (!__sameConfigJson_(cur, next)) return false;
    } else {
      const asText = (v) => (v === undefined || v === null) ? '' : String(v);
      if (asText(cur) !== asText(next)) return false;
    }
  }
  return true;
}

function __sameConfigJson_(a, b) {
  const pa = safeJsonParse_(a, null);
  const pb = safeJsonParse_(b, null);
  if (!pa || !pb || typeof pa !== 'object' || typeof pb !== 'object') return false;
  const strip = (c) => {
    const o = Object.assign({}, c);
    delete o.lastAccessedAt;
    delete o.etag;
    return o;
  };
  return __stableStringify_(strip(pa)) === __stableStringify_(strip(pb));
}

// key 順に依らない JSON 文字列 (比較専用)。
function __stableStringify_(v) {
  if (v === null || typeof v !== 'object') return JSON.stringify(v === undefined ? null : v);
  if (Array.isArray(v)) return '[' + v.map(__stableStringify_).join(',') + ']';
  const keys = Object.keys(v).filter((k) => v[k] !== undefined).sort();
  return '{' + keys.map((k) => JSON.stringify(k) + ':' + __stableStringify_(v[k])).join(',') + '}';
}

/**
 * SpreadsheetIDによるユーザー検索（configJSON-based）
 * Single Source of Truth - search by spreadsheetId in configJSON
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {Object} context - アクセスコンテキスト
 * @param {boolean} context.skipCache - キャッシュをスキップ（デフォルト: false）
 * @param {number} context.cacheTtl - キャッシュTTL秒数（デフォルト: 300秒）
 * @returns {Object|null} User object or null if not found
 */
function findUserBySpreadsheetId_(spreadsheetId, context = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('findUserBySpreadsheetId: Invalid spreadsheetId provided:', typeof spreadsheetId);
      return null;
    }

    const cacheVersion = userCacheVersion_();
    const cacheKey = `user_by_sheet_v${cacheVersion}_${spreadsheetId}`;
    const skipCache = context.skipCache || false;

    if (!skipCache) {
      try {
        const cached = CacheService.getScriptCache().get(cacheKey);
        if (cached) {
          const cachedUser = JSON.parse(cached);
          return cachedUser;
        }
      } catch (cacheError) {
        console.warn('findUserBySpreadsheetId: Cache read failed:', cacheError.message);
      }
    }

    // skipCache は「SS → user の対応 (user_by_sheet, 10 分)」だけを飛ばす。users 一覧そのものは
    //   users の cache の版 (user_cache_ver) 付き cache (config 保存で新しくなる) を使う。
    //   Why: 以前は skipCache が getAllUsers_ にも伝わり、checkSpreadsheetAccess_ の
    //   cache miss (児童 × SS ごとに 60 秒) のたびに users シートを読んでいた。30 人の授業で
    //   毎分 30 read、しかもフェーズ切替で同時に切れるので burst になっていた (2026-09-08)。
    const allUsers = getAllUsers_({ activeOnly: false }, {
      ...context, skipCache: false, systemLookup: true, preloadedAuth: context.preloadedAuth
    });
    if (!Array.isArray(allUsers)) {
      console.warn('findUserBySpreadsheetId: Failed to get users list');
      return null;
    }

    // Top-level config.spreadsheetId と config.profiles[].spreadsheetId の両方を走査。
    // profile 切替時 (生徒が「導入アンケート」 等のタブをクリック) は active と異なる SS ID が
    // 渡されるため、 top-level だけでは「Target user not found」 になり viewer が deny される。
    for (const user of allUsers) {
      try {
        const configJson = user.configJson || '{}';
        const config = JSON.parse(configJson);

        const matchesTop = config.spreadsheetId === spreadsheetId;
        const matchesProfile = !matchesTop && Array.isArray(config.profiles) &&
          config.profiles.some((p) => p && p.spreadsheetId === spreadsheetId);

        if (matchesTop || matchesProfile) {
          if (!skipCache) {
            try {
              const cacheTtl = context.cacheTtl || 600; // デフォルト10分（600秒）
              CacheService.getScriptCache().put(cacheKey, JSON.stringify(user), cacheTtl);
            } catch (cacheError) {
              console.warn('findUserBySpreadsheetId: Cache write failed:', cacheError.message);
            }
          }

          return user;
        }
      } catch (parseError) {
        console.warn(`findUserBySpreadsheetId_: Failed to parse configJSON for user ${user.userId}:`, parseError.message);
        continue;
      }
    }

    // Why (negative cache 5s): 旧版は 60s null cache で「Target user not found」 が
    //   1 分間ハマる事故を起こしていた (profile 切替 transient + getAllUsers_ stale fetch)。
    //   stale SS URL の防御目的なら 5s で十分 (polling ~ 8s なので次回ヒット時に再判定)。
    //   user_cache_ver の更新で proactive invalidation も働くが、 fallback として TTL も短縮。
    if (!skipCache) {
      try {
        CacheService.getScriptCache().put(cacheKey, JSON.stringify(null), 5);
      } catch (cacheError) {
        console.warn('findUserBySpreadsheetId: Cache write failed for not found result:', cacheError.message);
      }
    }

    return null;
  } catch (error) {
    logError_('findUserBySpreadsheetId', error);
    return null;
  }
}

