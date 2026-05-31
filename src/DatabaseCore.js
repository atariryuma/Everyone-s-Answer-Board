/**
 * @fileoverview DatabaseCore - users シートの CRUD、Sheets API リトライ/サーキット
 *   ブレーカー、Service Account JWT による安全なアクセス基盤。
 */

/* global validateEmail, CACHE_DURATION, getCurrentEmail, isAdministrator, getUserConfig, executeWithRetry, getCachedProperty, clearPropertyCache, simpleHash, saveToCacheWithSizeCheck, DEFAULT_DISPLAY_SETTINGS, safeJsonParse_, logError_, sameEmail_ */

/**
 * Sheets API 呼び出しラッパー (適応型 backoff + circuit breaker + SA pool failover)。
 *
 * `saEmail` と `authResolver` を渡すと、 retry 時に元 SA が cooling 中なら別 SA の token に
 * 詰め替えて非 cooling SA で retry を続ける (auto failover)。 これがないと closure 固定
 * された accessToken が同じ quota 焼け SA に再突入して失敗する。
 *
 * @param {string} url
 * @param {Object} options - UrlFetchApp.fetch オプション (Authorization header 含む)
 * @param {string} [operationName] - ログ識別
 * @param {string} [saEmail] - 初回 SA の client_email (cooldown 登録 + failover trigger)
 * @param {Function} [authResolver] - retry 時に新 SA を解決する closure
 */
function fetchSheetsAPIWithRetry(url, options, operationName, saEmail, authResolver) {
  let retryCount = 0;
  let currentOptions = options;
  let currentSaEmail = saEmail || null;

  const cache = CacheService.getScriptCache();
  const CIRCUIT_BREAKER_KEY = 'circuit_breaker_state';
  const cachedState = cache.get(CIRCUIT_BREAKER_KEY);
  // Why: JSON.parse が想定外の値 (部分書込み、過去 schema、null fields) を返した場合に
  //   undefined > 0 が silently false になり、 circuit breaker が機能しない silent bug を防ぐ。
  const circuitState = { consecutiveErrors: 0, circuitOpenUntil: 0 };
  if (cachedState) {
    try {
      const parsed = JSON.parse(cachedState);
      if (parsed && typeof parsed === 'object') {
        circuitState.consecutiveErrors = Number(parsed.consecutiveErrors) || 0;
        circuitState.circuitOpenUntil = Number(parsed.circuitOpenUntil) || 0;
      }
    } catch (parseError) {
      console.warn('Circuit breaker state parse failed, resetting:', parseError.message);
    }
  }

  const now = Date.now();
  if (circuitState.circuitOpenUntil > now) {
    const waitTime = Math.round((circuitState.circuitOpenUntil - now) / 1000);
    throw new Error(`Circuit breaker open: API calls paused for ${waitTime}s to allow quota recovery`);
  }

  return executeWithRetry(
    () => {
      // retry 時、 元 SA が cooling なら resolver で別 SA の token に詰め替える。 これがないと
      // closure 固定された accessToken が同じ quota 焼け SA に再突入して「failed after 2 attempts」
      // で user-facing fail になる。
      if (authResolver && retryCount > 0 && currentSaEmail && isServiceAccountCoolingDown_(currentSaEmail)) {
        try {
          const fresh = authResolver();
          if (fresh && fresh.token && fresh.saEmail && fresh.saEmail !== currentSaEmail) {
            const newHeaders = Object.assign({}, currentOptions.headers || {}, { Authorization: 'Bearer ' + fresh.token });
            currentOptions = Object.assign({}, currentOptions, { headers: newHeaders });
            console.warn(`↪ ${operationName || 'SheetsAPI'}: switching SA ${currentSaEmail} → ${fresh.saEmail} (retry ${retryCount})`);
            currentSaEmail = fresh.saEmail;
          }
        } catch (_) { /* resolver 失敗は飲んで元 SA で retry */ }
      }

      const response = UrlFetchApp.fetch(url, currentOptions);
      const code = response.getResponseCode();

      if (code === 429) {
        // GAS 6 分制限内に収めるため inner sleep は 15/30/45/60s で cap。 当該 SA を 30s cooldown
        // に入れて後続 request は別 SA に逃がす (auto failover)。
        if (currentSaEmail) markServiceAccountCoolingDown_(currentSaEmail);
        const backoffTime = Math.min(15000 + (retryCount * 15000), 60000);
        console.warn(`⚠️ 429 ${operationName || 'SheetsAPI'}${currentSaEmail ? ' [' + currentSaEmail + ']' : ''}: wait ${backoffTime}ms (retry ${retryCount})`);

        circuitState.consecutiveErrors++;
        if (circuitState.consecutiveErrors >= 7) {
          const lockoutMs = circuitState.consecutiveErrors >= 15 ? 120000
            : circuitState.consecutiveErrors >= 10 ? 60000 : 30000;
          // Date.now() を再取得する。 関数冒頭の `now` は executeWithRetry 突入前の値で、
          // ここに到達する頃には 429 backoff の Utilities.sleep (最大 ~60s×複数 retry) が
          // 経過済。 stale な `now` で計算するとロックアウトが実質ほぼ即時失効し、 quota
          // 回復のための circuit pause が機能しない。
          circuitState.circuitOpenUntil = Date.now() + lockoutMs;
          console.error(`Circuit breaker activated: ${circuitState.consecutiveErrors} errors. Paused for ${lockoutMs / 1000}s.`);
        }
        cache.put(CIRCUIT_BREAKER_KEY, JSON.stringify(circuitState), 120);

        Utilities.sleep(backoffTime);
        retryCount++;
        throw new Error('Quota exceeded (429), retry with adaptive backoff');
      }

      if (code !== 200) {
        const errorText = response.getContentText();
        throw new Error(`API returned ${code}: ${errorText}`);
      }

      circuitState.consecutiveErrors = 0;
      circuitState.circuitOpenUntil = 0;
      cache.put(CIRCUIT_BREAKER_KEY, JSON.stringify(circuitState), 60);

      return response;
    },
    {
      maxRetries: 3,
      initialDelay: 2000,
      maxDelay: 20000,
      operationName: operationName || 'Sheets API call'
    }
  );
}

/*
 * Service Account pool (round-robin, with cooldown).
 *
 * Sheets API quota は SA 1 個あたり 300 read/min。 200-700 人同時アクセスでは単一 SA で焼ける
 * ため、 SERVICE_ACCOUNT_CREDS (primary) に加え SERVICE_ACCOUNT_CREDS_2..10 を Script Properties
 * に並べて N 倍化する。 200 人で 2-3 個、 500-700 人で 4-5 個推奨。 10 個以上は admin 監視が
 * 煩雑になるだけで実効効果は頭打ち。
 *
 * 設計の核:
 *  - **round-robin** (`pickServiceAccount_`): ScriptCache の共有 counter `sa_rr_counter` を
 *    pool.length で割って次 SA を pick。 heavy user (admin/教師の連投) も pool 全体に均等
 *    分散される。 user-affinity hash 方式だと heavy user が固定 SA を焼く事故が起きる。
 *  - **cooldown** (`markServiceAccountCoolingDown_`): 429 を喰った SA を 30s 除外。 同 user が
 *    同 SA で 429 を再現する無駄を防ぐ。 全 SA cooling 時は full pool で round-robin (止める
 *    より retry のほうがマシ)。
 *  - **2 段 token cache** (`getServiceAccountAccessToken_`): in-memory + ScriptCache 50min。
 *    SA 切替時の JWT 交換 overhead を実質ゼロに。
 *  - **per-SA verify cache** (`verifyServiceAccountAccess_`): 1 SA だけ DB 共有漏れがあっても
 *    pool 全体は機能継続。
 */
const SERVICE_ACCOUNT_POOL_MAX_ = 10;
const SA_COOLDOWN_TTL_SEC = 30;
const SA_TOKEN_TTL_SEC_ = 50 * 60;
const SA_VERIFY_TTL_OK_SEC_ = 600;
const SA_VERIFY_TTL_NO_SEC_ = 120;

// Cache key prefixes (集約。 scattered string literal を排除)。
const SA_CACHE_KEYS_ = {
  TOKEN: 'sa_token:',          // SA access token (per client_email, 50min TTL)
  ACCESS: 'sa_access:',         // SS x SA verify result (per ssId+saEmail, 10min ok / 2min no)
  COOLDOWN: 'sa_cooldown:',     // 429 後の cooldown フラグ (per client_email, 30s TTL)
  PICK_COUNTER: 'sa_pick:',     // admin 可視化用 pick 回数 (per client_email, 5min TTL)
  RR_COUNTER: 'sa_rr_counter',  // round-robin 共有 counter
  VALIDATION_VER: 'sa_validation_ver:', // SS publish 状態変更時の cache invalidator
  VALIDATION: 'sa_validation_'  // validation 結果 cache (key includes ver + ssId + email)
};

function serviceAccountPoolSlotKey_(n) {
  return n === 1 ? 'SERVICE_ACCOUNT_CREDS' : ('SERVICE_ACCOUNT_CREDS_' + n);
}

/**
 * SA JSON を soft parse する (壊れていたら null。 throw しない)。
 * pool loader 側は「壊れてる SA は無視して次へ」が正しい挙動。
 * @param {string} raw - JSON 文字列
 * @returns {Object|null} parsed SA or null
 */
function parseServiceAccountCredsSoft_(raw) {
  if (!raw || typeof raw !== 'string') return null;
  try {
    const sa = JSON.parse(raw);
    if (!sa || typeof sa !== 'object') return null;
    if (!sa.client_email || !sa.private_key || !sa.type) return null;
    if (!validateEmail(sa.client_email).isValid) return null;
    if (typeof sa.private_key !== 'string' ||
        (sa.private_key.indexOf('BEGIN PRIVATE KEY') === -1 &&
         sa.private_key.indexOf('BEGIN RSA PRIVATE KEY') === -1)) {
      return null;
    }
    return sa;
  } catch (_) {
    return null;
  }
}

/**
 * 設定されている SA を全て返す (primary + secondaries, validate 済)。
 * @returns {Object[]} array of SA objects (empty if none configured)
 */
function getAllServiceAccounts_() {
  const out = [];
  for (let n = 1; n <= SERVICE_ACCOUNT_POOL_MAX_; n++) {
    const sa = parseServiceAccountCredsSoft_(getCachedProperty(serviceAccountPoolSlotKey_(n)));
    if (sa) out.push(sa);
  }
  return out;
}

/** SA pick 回数を 5 分窓カウンタに記録 (admin 可視化用)。 */
function recordSaPick_(saEmail) {
  if (!saEmail || typeof CacheService === 'undefined') return;
  try {
    const cache = CacheService.getScriptCache();
    const key = SA_CACHE_KEYS_.PICK_COUNTER + saEmail;
    const cur = Number(cache.get(key) || 0);
    cache.put(key, String(cur + 1), 300);
  } catch (_) { /* ignore */ }
}

/** 429 を喰った SA を short-term cooldown (30s) に入れる。 */
function markServiceAccountCoolingDown_(saEmail) {
  if (!saEmail || typeof CacheService === 'undefined') return;
  try { CacheService.getScriptCache().put(SA_CACHE_KEYS_.COOLDOWN + saEmail, '1', SA_COOLDOWN_TTL_SEC); }
  catch (_) { /* ignore */ }
}

function isServiceAccountCoolingDown_(saEmail) {
  if (!saEmail || typeof CacheService === 'undefined') return false;
  try { return Boolean(CacheService.getScriptCache().get(SA_CACHE_KEYS_.COOLDOWN + saEmail)); }
  catch (_) { return false; }
}

/**
 * リクエストごとに SA を 1 個 round-robin で pick する。 cooldown 中の SA は候補から除外。
 * 全 SA cooling 時は full pool で round-robin (degrade して止まらないことを優先)。
 * pool.length === 0 のとき null を返す (caller は openById fallback)。
 *
 * Why round-robin: heavy user (admin/教師) が同 SA に張り付いて quota 焼け、 という現象を
 * 防ぐ。 user-hash affinity だと bucket 偏りで pool 利用率が 21/44/35% など歪み、 1 SA が
 * 集中砲火を受ける。 共有 counter を pool.length で mod すれば完全均等分散。 SA token cache
 * (50 分) は per-SA で warm 維持されるので、 SA 切替コストは ~0 (cache hit のみ)。
 *
 * @returns {Object|null} picked SA or null
 */
function pickServiceAccount_() {
  const fullPool = getAllServiceAccounts_();
  if (fullPool.length === 0) return null;
  if (fullPool.length === 1) {
    recordSaPick_(fullPool[0].client_email);
    return fullPool[0];
  }
  let pool = fullPool.filter(sa => !isServiceAccountCoolingDown_(sa.client_email));
  if (pool.length === 0) pool = fullPool;
  let idx;
  try {
    if (typeof CacheService !== 'undefined') {
      const cache = CacheService.getScriptCache();
      const RR_KEY = SA_CACHE_KEYS_.RR_COUNTER;
      const cur = Number(cache.get(RR_KEY) || 0);
      idx = cur % pool.length;
      cache.put(RR_KEY, String((cur + 1) % 1000000), 600);
    } else {
      idx = Math.floor(Math.random() * pool.length);
    }
  } catch (_) {
    idx = Math.floor(Math.random() * pool.length);
  }
  const picked = pool[idx];
  recordSaPick_(picked.client_email);
  return picked;
}

/**
 * SA pool の直近 5 分使用回数を返す。 admin が「全 SA に分散しているか」を可視化する用。
 * @returns {Object} { success, windowSec, total, coolingCount, pool: [{ clientEmail, picks, cooling }] }
 */
function getServiceAccountUsage() {
  const pool = getAllServiceAccounts_();
  if (typeof CacheService === 'undefined') {
    return { success: true, windowSec: 300, cooldownTtlSec: SA_COOLDOWN_TTL_SEC, total: 0, coolingCount: 0, pool: [] };
  }
  const cache = CacheService.getScriptCache();
  const usage = pool.map((sa) => {
    let count = 0; let cooling = false;
    try { count = Number(cache.get('sa_pick:' + sa.client_email) || 0); } catch (_) {}
    try { cooling = Boolean(cache.get('sa_cooldown:' + sa.client_email)); } catch (_) {}
    return { clientEmail: sa.client_email, picks: count, cooling };
  });
  return {
    success: true,
    windowSec: 300,
    cooldownTtlSec: SA_COOLDOWN_TTL_SEC,
    total: usage.reduce((acc, u) => acc + u.picks, 0),
    coolingCount: usage.filter(u => u.cooling).length,
    pool: usage
  };
}

/**
 * サービスアカウント認証情報を取得 (primary slot のメタ + pool 全体のサマリ)。
 * @returns {Object} { isValid, email?, type?, poolSize?, pool?: email[] }
 */
function getServiceAccount() {
  try {
    const pool = getAllServiceAccounts_();
    if (pool.length === 0) {
      return { isValid: false, error: 'No credentials found' };
    }
    const primary = pool[0];
    return {
      isValid: true,
      email: primary.client_email,
      type: primary.type,
      poolSize: pool.length,
      pool: pool.map((s) => s.client_email)
    };
  } catch (error) {
    console.warn('getServiceAccount: Invalid credentials format');
    return { isValid: false, error: 'Invalid format' };
  }
}

/**
 * SA access token を 2 段 cache (in-memory + ScriptCache) で 50 分保持。
 *
 * 段構成:
 *   1. in-memory (`saTokenCache_`): 同一 V8 instance の warm reuse 内で fastest hit
 *   2. ScriptCache: 全 invocation 共有。 cold start (新 instance) でも JWT 交換 1 回で済む。
 *      200 人朝の会の最初の数十人で N 重複 JWT 交換が発生する事故を防ぐ。
 */
const saTokenCache_ = Object.create(null);

function getServiceAccountAccessToken_(sa) {
  if (!sa || !sa.client_email || !sa.private_key) return null;
  const cacheKey = sa.client_email;
  const now = Date.now();

  // Tier 1: in-memory
  const memo = saTokenCache_[cacheKey];
  if (memo && memo.expiresAt > now) return memo.token;

  // Tier 2: ScriptCache
  const scriptCache = (typeof CacheService !== 'undefined' && CacheService.getScriptCache)
    ? CacheService.getScriptCache() : null;
  const scriptCacheKey = SA_CACHE_KEYS_.TOKEN + cacheKey;
  if (scriptCache) {
    try {
      const cached = scriptCache.get(scriptCacheKey);
      if (cached) {
        // ScriptCache 値は {t: token, e: 絶対期限epochMs}。 in-memory expiresAt は **保存済の
        //   絶対期限** で cap する。 旧実装は read 時刻 + 50min を入れていたため、 ScriptCache に
        //   49 分前から載っていた token を read すると in-memory が「あと 50 分有効」と誤認し、
        //   OAuth token の実寿命 60 分を超えて期限切れ token を配り 401 を招いていた
        //   (401 は 429 でないため SA failover も走らない)。
        let token = cached;
        let absExpiry = now + SA_TOKEN_TTL_SEC_ * 1000; // legacy plain-string fallback
        try {
          const parsed = JSON.parse(cached);
          if (parsed && parsed.t) {
            token = parsed.t;
            absExpiry = Number(parsed.e) || absExpiry;
          } else {
            // 旧形式 (plain token 文字列): 真の発行時刻が不明なため短命に倒して安全側へ。
            absExpiry = now + 60 * 1000;
          }
        } catch (_) {
          absExpiry = now + 60 * 1000;
        }
        if (absExpiry > now) {
          saTokenCache_[cacheKey] = { token, expiresAt: absExpiry };
          return token;
        }
      }
    } catch (_) { /* ignore */ }
  }

  // JWT exchange
  const nowSec = Math.floor(now / 1000);
  const jwt = createServiceAccountJWT(
    { alg: 'RS256', typ: 'JWT' },
    {
      iss: sa.client_email,
      scope: 'https://www.googleapis.com/auth/spreadsheets',
      aud: 'https://oauth2.googleapis.com/token',
      exp: nowSec + 3600,
      iat: nowSec
    },
    sa.private_key
  );
  try {
    const tokenResp = UrlFetchApp.fetch('https://oauth2.googleapis.com/token', {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      payload: `grant_type=urn:ietf:params:oauth:grant-type:jwt-bearer&assertion=${jwt}`,
      muteHttpExceptions: true
    });
    const tokenData = JSON.parse(tokenResp.getContentText());
    if (!tokenData.access_token) {
      console.warn('SA token exchange failed:', tokenData.error_description || tokenData.error);
      return null;
    }
    // OAuth token の実寿命 (expires_in, 通常 3599s)。 安全マージン 5 分を引いた絶対期限を
    //   両 tier で共有し、 read 時に in-memory が実寿命を超えて延命しないようにする。
    const lifetimeSec = Number(tokenData.expires_in) > 0 ? Number(tokenData.expires_in) : 3600;
    const absExpiry = now + Math.max(0, (lifetimeSec - 300)) * 1000;
    // ScriptCache TTL は in-memory cap (=token 実寿命) と従来の 50min の小さい方。
    const scriptTtlSec = Math.min(SA_TOKEN_TTL_SEC_, Math.max(60, lifetimeSec - 300));
    saTokenCache_[cacheKey] = { token: tokenData.access_token, expiresAt: absExpiry };
    if (scriptCache) {
      try {
        scriptCache.put(scriptCacheKey, JSON.stringify({ t: tokenData.access_token, e: absExpiry }), scriptTtlSec);
      } catch (_) {}
    }
    return tokenData.access_token;
  } catch (err) {
    console.warn('SA token exchange exception:', err && err.message);
    return null;
  }
}

/**
 * SA が実際にスプレッドシートにアクセスできるか per-SA で 1 回検証 (10 分 ok cache, 2 分 no cache)。
 * 未共有 (403) の SA は次の候補に skip → 1 個だけ共有漏れがあっても pool 全体が機能継続。
 * 429/5xx は transient なので cache しない (焼くと全 SA path が一斉に封鎖される)。
 */
function verifyServiceAccountAccess_(sheetId, accessToken, saEmail) {
  if (!sheetId || !accessToken) return false;
  const cacheKey = SA_CACHE_KEYS_.ACCESS + sheetId + ':' + (saEmail || 'unknown');
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (_) {}
  if (cache) {
    try {
      const cached = cache.get(cacheKey);
      if (cached === 'ok') return true;
      if (cached === 'no') return false;
    } catch (_) {}
  }
  try {
    const resp = UrlFetchApp.fetch(
      `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}?includeGridData=false&fields=properties.title`,
      { headers: { Authorization: `Bearer ${accessToken}` }, muteHttpExceptions: true }
    );
    const code = resp.getResponseCode();
    const ok = code === 200;
    const isTransient = code === 429 || (code >= 500 && code < 600);
    if (cache) {
      try {
        if (ok) cache.put(cacheKey, 'ok', SA_VERIFY_TTL_OK_SEC_);
        else if (!isTransient) cache.put(cacheKey, 'no', SA_VERIFY_TTL_NO_SEC_);
      } catch (_) {}
    }
    if (!ok) {
      console.warn('SA access verify failed for', sheetId.substring(0, 8) + '...',
        'code=' + code, isTransient ? '(transient, no cache)' : '');
    }
    return ok;
  } catch (err) {
    console.warn('SA access verify exception:', err && err.message);
    return false;
  }
}

/**
 * proxy が保持する SA を「動的解決」する closure を作る。
 *  - 通常: 初回 pick した SA をそのまま使う (token cache 再利用)
 *  - 初回 SA が cooldown 中: pool から非 cooling SA を 1 個再 pick + token を取得し直す
 *  - 非 cooling SA 無し (全 pool cooling): 初回 SA に戻す (完全停止より retry のほうがマシ)
 *
 * これにより同 session 内で 429 を踏んだ瞬間に別 SA に逃げられる。 旧実装は proxy が初回 token
 * を closure 固定していたため、 同 session が同 SA に張り付いて quota 焼け retry を繰り返す
 * 事故が起きていた。
 */
function makeProxyAuthResolver_(sheetId, initialToken, initialSaEmail) {
  return function currentAuth() {
    if (!initialSaEmail || !isServiceAccountCoolingDown_(initialSaEmail)) {
      return { token: initialToken, saEmail: initialSaEmail };
    }
    const fresh = pickServiceAccount_();
    if (!fresh || fresh.client_email === initialSaEmail) {
      return { token: initialToken, saEmail: initialSaEmail };
    }
    if (isServiceAccountCoolingDown_(fresh.client_email)) {
      return { token: initialToken, saEmail: initialSaEmail };
    }
    const freshToken = getServiceAccountAccessToken_(fresh);
    if (!freshToken) return { token: initialToken, saEmail: initialSaEmail };
    return { token: freshToken, saEmail: fresh.client_email };
  };
}

/**
 * 全 users.config から board の spreadsheetId を抽出 (重複除去)。
 * 旧 countBoardSpreadsheets_ (DatabaseCore) / collectAllBoardSpreadsheetIds_ (AdminApis) の統合。
 * @returns {string[]}
 */
function getAllBoardSpreadsheetIds() {
  try {
    if (typeof getAllUsers !== 'function') return [];
    const users = getAllUsers({ activeOnly: false }, { forceServiceAccount: true });
    if (!Array.isArray(users)) return [];
    const ids = new Set();
    for (const u of users) {
      if (!u || !u.configJson) continue;
      const cfg = safeJsonParse_(u.configJson, null);
      if (!cfg) continue;
      if (cfg.spreadsheetId && typeof cfg.spreadsheetId === 'string') ids.add(cfg.spreadsheetId);
      if (Array.isArray(cfg.profiles)) {
        for (const p of cfg.profiles) {
          if (p && p.spreadsheetId && typeof p.spreadsheetId === 'string') ids.add(p.spreadsheetId);
        }
      }
    }
    return Array.from(ids);
  } catch (err) {
    if (typeof logError_ === 'function') logError_('getAllBoardSpreadsheetIds', err);
    return [];
  }
}

/**
 * 公開状態変更時に呼ぶ。 sa_validation cache (`validateServiceAccountUsage` の 60s 短期 cache)
 * を該当 SS について全 viewer 分まとめて invalidate。 これがないと unpublish 直後 60 秒間
 * viewer がアクセスできる security leak が残る。
 *
 * ScriptCache には wildcard delete がないため、 version counter を bump する方式で実装する。
 * `sa_validation_*` cache key 側にこの version を含めれば、 bump 後は旧 key が全 stale になる。
 */
function invalidateSaValidationCache_(spreadsheetId) {
  if (!spreadsheetId || typeof CacheService === 'undefined') return;
  try {
    const cache = CacheService.getScriptCache();
    // `validateServiceAccountUsage` 内で email キーを含む全 viewer entry を一括 stale 化したい。
    // wildcard 不可なので、 ssId 単位の version counter を持って cache key に組み込む方式に変更。
    const key = SA_CACHE_KEYS_.VALIDATION_VER + spreadsheetId;
    const cur = Number(cache.get(key) || 0);
    cache.put(key, String(cur + 1), 3600);
  } catch (_) { /* ignore */ }
}

/** SA 1 個の token + access キャッシュをまとめて invalidate (新規登録 / 再 verify 時に呼ぶ)。 */
function invalidateSaCache_(spreadsheetId, clientEmail) {
  if (typeof CacheService === 'undefined') return;
  try {
    const cache = CacheService.getScriptCache();
    if (spreadsheetId && clientEmail) cache.remove(SA_CACHE_KEYS_.ACCESS + spreadsheetId + ':' + clientEmail);
    if (clientEmail) {
      cache.remove(SA_CACHE_KEYS_.TOKEN + clientEmail);
      delete saTokenCache_[clientEmail];
    }
  } catch (_) { /* ignore */ }
}

// ---------------------------------------------------------------------
// SA pool 管理 (admin UI 用)
// ---------------------------------------------------------------------

/** SA pool 全スロットをメタ情報のみで返す (private_key は含めない)。 */
function listServiceAccountPool() {
  const slots = [];
  for (let n = 1; n <= SERVICE_ACCOUNT_POOL_MAX_; n++) {
    const key = serviceAccountPoolSlotKey_(n);
    const raw = getCachedProperty(key);
    if (!raw) continue;
    const sa = parseServiceAccountCredsSoft_(raw);
    if (sa) {
      slots.push({
        slot: key,
        index: n,
        role: n === 1 ? 'primary' : 'secondary',
        clientEmail: sa.client_email,
        projectId: sa.project_id || null,
        privateKeyId: sa.private_key_id ? (sa.private_key_id.slice(0, 8) + '…') : null,
        valid: true
      });
    } else {
      slots.push({
        slot: key,
        index: n,
        role: n === 1 ? 'primary' : 'secondary',
        clientEmail: null,
        valid: false,
        error: 'parse-failed'
      });
    }
  }
  return {
    success: true,
    poolSize: slots.filter((s) => s.valid).length,
    maxSize: SERVICE_ACCOUNT_POOL_MAX_,
    slots
  };
}

/**
 * SA JSON を pool に追加。 空きスロットを自動選択。 同 client_email の重複は拒否。
 * DB editor 共有 → token 取得 → access verify を一括実行 (UI は verified を見る)。
 */
function addServiceAccountToPool(saJsonString) {
  try {
    const sa = parseServiceAccountCredsSoft_(saJsonString);
    if (!sa) return { success: false, error: 'INVALID_SA_JSON', message: 'SA JSON の形式が不正です' };

    const existing = getAllServiceAccounts_();
    for (let i = 0; i < existing.length; i++) {
      if (existing[i].client_email === sa.client_email) {
        return { success: false, error: 'ALREADY_REGISTERED',
          message: `${sa.client_email} は既に pool に登録されています` };
      }
    }

    let targetSlot = null;
    let targetIndex = null;
    for (let n = 1; n <= SERVICE_ACCOUNT_POOL_MAX_; n++) {
      const key = serviceAccountPoolSlotKey_(n);
      if (!getCachedProperty(key)) { targetSlot = key; targetIndex = n; break; }
    }
    if (!targetSlot) {
      return { success: false, error: 'POOL_FULL',
        message: `pool は上限 (${SERVICE_ACCOUNT_POOL_MAX_} 個) に達しています` };
    }

    PropertiesService.getScriptProperties().setProperty(targetSlot, String(saJsonString));
    clearPropertyCache(targetSlot);

    // DB 共有 + verify (fail-soft)
    let shared = false;
    let shareError = null;
    let shareHint = null;
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    try {
      if (dbId && typeof DriveApp !== 'undefined') {
        DriveApp.getFileById(dbId).addEditor(sa.client_email);
        shared = true;
      }
    } catch (err) {
      shareError = (err && err.message) || String(err);
      const msg = String(shareError).toLowerCase();
      if (msg.indexOf('does not have permission') !== -1 || msg.indexOf('権限がありません') !== -1) {
        shareHint = 'GAS 実行ユーザが DB のオーナー/編集者ではありません。 手動で DB を ' + sa.client_email + ' に編集者共有してください。';
      } else if (msg.indexOf('rate') !== -1 || msg.indexOf('429') !== -1 || msg.indexOf('quota') !== -1) {
        shareHint = 'Drive API quota 一時上限。 30 秒後に「再 verify」 を押してください。';
      } else {
        shareHint = '自動共有失敗。 必要なら DB を ' + sa.client_email + ' に編集者共有してください。';
      }
    }

    let verified = false;
    let verifyError = null;
    try {
      if (dbId) {
        invalidateSaCache_(dbId, sa.client_email);
        const accessToken = getServiceAccountAccessToken_(sa);
        if (!accessToken) {
          verifyError = 'token 取得失敗 (private_key 不正 / IAM API 未有効化)';
        } else {
          verified = verifyServiceAccountAccess_(dbId, accessToken, sa.client_email);
          if (!verified) verifyError = '共有反映待ち or DB アクセス不可';
        }
      }
    } catch (err) {
      verifyError = (err && err.message) || String(err);
    }

    // 新 SA を board SS にも展開しないと、 pool 拡張効果が出ない (新 SA は全 board で verify
    // 失敗 → primary SA に round-robin が偏る)。 admin に migrateBoardSharing 実行を促す。
    const boardsToMigrate = getAllBoardSpreadsheetIds().length;

    return {
      success: true,
      slot: targetSlot,
      index: targetIndex,
      clientEmail: sa.client_email,
      shared,
      shareError,
      shareHint,
      verified,
      verifyError,
      boardsToMigrate,
      migrationHint: boardsToMigrate > 0
        ? `${boardsToMigrate} 個の既存 board に新 SA を editor 追加するため "migrateBoardSharing" を実行してください`
        : null
    };
  } catch (err) {
    logError_('addServiceAccountToPool', err);
    return { success: false, error: 'EXCEPTION', message: (err && err.message) || String(err) };
  }
}

/** 複数 SA JSON を一括登録 (改行区切り or `}\n{` 連結 or array)。 */
function addServiceAccountsToPoolBatch(jsonInputs) {
  let arr;
  if (Array.isArray(jsonInputs)) {
    arr = jsonInputs.map(String);
  } else {
    const s = String(jsonInputs || '');
    arr = s.split(/}\s*{/).map((chunk, i, all) => {
      let c = chunk;
      if (i > 0) c = '{' + c;
      if (i < all.length - 1) c += '}';
      return c.trim();
    }).filter(Boolean);
  }
  const results = [];
  for (let i = 0; i < arr.length; i++) {
    results.push(addServiceAccountToPool(arr[i]));
  }
  const verifiedCount = results.filter((r) => r && r.success && r.verified).length;
  const registeredCount = results.filter((r) => r && r.success).length;
  return {
    success: registeredCount > 0,
    total: results.length,
    registered: registeredCount,
    verified: verifiedCount,
    results
  };
}

/** admin の「再 verify」 ボタン用。 共有反映待ちで verified=false だった SA を再確認。 */
function reverifyServiceAccountInPool(slotKey) {
  try {
    const key = String(slotKey || '');
    if (!/^SERVICE_ACCOUNT_CREDS(_\d+)?$/.test(key)) {
      return { success: false, error: 'INVALID_SLOT', message: '不正なスロット: ' + key };
    }
    const raw = getCachedProperty(key);
    if (!raw) return { success: false, error: 'NOT_FOUND', message: key + ' は登録されていません' };
    const sa = parseServiceAccountCredsSoft_(raw);
    if (!sa) return { success: false, error: 'INVALID_SA_JSON', message: 'SA JSON が壊れています' };
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) return { success: false, error: 'DB_NOT_CONFIGURED' };
    invalidateSaCache_(dbId, sa.client_email);
    const accessToken = getServiceAccountAccessToken_(sa);
    if (!accessToken) return { success: true, verified: false, error: 'TOKEN_FAILED' };
    const verified = verifyServiceAccountAccess_(dbId, accessToken, sa.client_email);
    return { success: true, verified, clientEmail: sa.client_email };
  } catch (err) {
    return { success: false, error: 'EXCEPTION', message: (err && err.message) || String(err) };
  }
}

/**
 * Secondary slot の SA を削除。 primary は削除不可 (差替は setupApp 経由)。
 * GCP 側の key revoke は別途 (誤削除からの復旧の余地を残す)。
 */
function removeServiceAccountFromPool(slotKey) {
  try {
    const key = String(slotKey || '');
    if (key === 'SERVICE_ACCOUNT_CREDS') {
      return { success: false, error: 'PRIMARY_LOCKED',
        message: 'primary (SERVICE_ACCOUNT_CREDS) は UI から削除できません。 セットアップ画面から差替えてください' };
    }
    if (!/^SERVICE_ACCOUNT_CREDS_(\d+)$/.test(key)) {
      return { success: false, error: 'INVALID_SLOT', message: `不正なスロット: ${key}` };
    }
    const n = Number(key.split('_').pop());
    if (!(n >= 2 && n <= SERVICE_ACCOUNT_POOL_MAX_)) {
      return { success: false, error: 'INVALID_SLOT', message: `範囲外スロット: ${key}` };
    }
    const existing = getCachedProperty(key);
    if (!existing) {
      return { success: false, error: 'NOT_FOUND', message: `${key} は登録されていません` };
    }
    const sa = parseServiceAccountCredsSoft_(existing);
    PropertiesService.getScriptProperties().deleteProperty(key);
    clearPropertyCache(key);
    if (sa && sa.client_email) {
      const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
      invalidateSaCache_(dbId, sa.client_email);
    }
    return { success: true, slot: key, removed: true };
  } catch (err) {
    logError_('removeServiceAccountFromPool', err);
    return { success: false, error: 'EXCEPTION', message: (err && err.message) || String(err) };
  }
}

/**
 * SS アクセスの妥当性と最適経路を判定する Security Gate + Router。
 *
 * 新セキュリティモデル (通常 Google Form 同等):
 *  - viewer は board SS への直接権限を持たない → SA pool 経由
 *  - owner は自分の SS のオーナー権限を持つ → openById 直接 (SA pool quota を節約)
 *  - admin は cross-user 操作 → SA pool
 *
 * 戻り値の `accessMode`:
 *   - `'own'` : caller が SS の owner、 SpreadsheetApp.openById で直接アクセス推奨
 *   - `'sa'`  : SA pool 経由でアクセス (admin / viewer)
 *   - `'denied'` (allowed=false のときのみ): アクセス不可
 *
 * 許可マトリクス:
 *   | caller         | DB sheet | own board | 他人の公開 board | 他人の非公開 board |
 *   |----------------|----------|-----------|------------------|--------------------|
 *   | admin          | sa       | own       | sa               | sa                 |
 *   | owner          | sa       | own       | sa               | denied             |
 *   | viewer (生徒)  | sa       | —         | sa               | denied             |
 *
 * @param {string} spreadsheetId
 * @param {boolean} useServiceAccount - false で強制 own モード (互換用)
 * @param {string} context - ログ識別用
 * @returns {{allowed:boolean, reason:string, accessMode:('own'|'sa'|'denied')}}
 */
function validateServiceAccountUsage(spreadsheetId, useServiceAccount, context = 'unknown') {
  try {
    if (!useServiceAccount) {
      return { allowed: true, reason: 'Direct access (legacy flag)', accessMode: 'own' };
    }

    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    if (spreadsheetId === dbId) {
      return { allowed: true, reason: 'DATABASE_SPREADSHEET_ID (shared resource)', accessMode: 'sa' };
    }

    const currentEmail = getCurrentEmail();

    // 短期 cache (60s) で重複 DB lookup を抑制。 cache key に caller email を含めることで
    // 同じ SS でも owner / viewer で別キャッシュ。 version counter (`sa_validation_ver:{ssId}`)
    // を組み込むことで、 publish 状態変更時に invalidateSaValidationCache_ から bump → 旧 key
    // が全 viewer 分一括 stale 化される。
    let cache = null;
    try { cache = CacheService.getScriptCache(); } catch (_) {}
    let validationVer = '0';
    if (cache) {
      try { validationVer = String(cache.get(SA_CACHE_KEYS_.VALIDATION_VER + spreadsheetId) || '0'); } catch (_) {}
    }
    const cacheKey = `sa_validation_v${validationVer}_${spreadsheetId}_${currentEmail || 'anon'}`;
    if (cache) {
      try {
        const cached = cache.get(cacheKey);
        if (cached) {
          const parsed = JSON.parse(cached);
          return {
            allowed: Boolean(parsed.allowed),
            reason: parsed.reason,
            accessMode: parsed.accessMode || (parsed.allowed ? 'sa' : 'denied')
          };
        }
      } catch (_) { /* fall through */ }
    }

    // 結果保存 helper — 4 分岐の cache.put 重複を集約。
    // Why (deny は短い TTL): allow 結果は publish 状態変化時に invalidateSaValidationCache_
    //   で proactive bump されるが、 deny 結果 (「Target user not found」「Board not
    //   published」) は SS 自体が user config に追加されただけでは bump が走らない。
    //   deny を 60s cache すると profile 追加 / 設定修正後も「アクセス不可」 が 1 分続く
    //   ため、 deny は 5s に短縮 (allow は従来通り 60s)。
    const cacheAndReturn = (result) => {
      if (cache) {
        try {
          const ttl = result.allowed ? 60 : 5;
          cache.put(cacheKey, JSON.stringify(result), ttl);
        } catch (_) {}
      }
      return result;
    };

    const targetUser = findUserBySpreadsheetId(spreadsheetId, { skipCache: true });
    if (!targetUser) {
      console.warn('SA_VALIDATION: Target user not found for spreadsheet:', spreadsheetId.substring(0, 8));
      return cacheAndReturn({ allowed: false, reason: 'Target user not found', accessMode: 'denied' });
    }

    // email は大文字小文字を区別しないので正規化して比較する。 isBoardCollaborator と揃え、
    // 保存値と session 値の case 差で owner を viewer と誤判定 (非公開ボードで誤 deny) するのを防ぐ。
    const normEmail = (e) => String(e || '').toLowerCase().trim();
    const isOwner = Boolean(targetUser.userEmail && currentEmail && normEmail(targetUser.userEmail) === normEmail(currentEmail));
    if (isOwner) {
      // 編集者は自分の SS のオーナー権限を持つ → openById で直接アクセス。
      // SA pool quota を viewer の閲覧トラフィックに温存できる。
      return cacheAndReturn({ allowed: true, reason: 'Own board (direct access)', accessMode: 'own' });
    }

    // admin は他人のボードに cross-user アクセスする必要があるので SA pool 経由。
    if (isAdministrator(currentEmail)) {
      return cacheAndReturn({ allowed: true, reason: 'Admin privileges (cross-user)', accessMode: 'sa' });
    }

    // 公開状態の判定。 `getUserConfig` 経由だと findUserById が viewer のアクセス制御で deny
    // するチキン&エッグ問題があるため、 既に取得済の `targetUser.configJson` を直接 parse する。
    // (`findUserBySpreadsheetId` は admin 経路で取得しており、 configJson は信頼できる)
    const isPublished = isUserBoardPublished(targetUser);
    if (!isPublished) {
      console.warn('SA_VALIDATION: Non-owner access to unpublished board:', {
        currentEmail: currentEmail ? `${currentEmail.split('@')[0]}@***` : 'unknown',
        spreadsheetId: spreadsheetId.substring(0, 8),
        context
      });
      return cacheAndReturn({ allowed: false, reason: 'Board not published', accessMode: 'denied' });
    }
    return cacheAndReturn({ allowed: true, reason: 'Public board access', accessMode: 'sa' });
  } catch (error) {
    logError_('SA_VALIDATION', error);
    return { allowed: false, reason: 'Validation error', accessMode: 'denied' };
  }
}

/**
 * データベーススプレッドシートを開く（CLAUDE.md準拠 - Editor→Admin共有DB）
 * DATABASE_SPREADSHEET_IDは常にサービスアカウントでアクセス（セキュリティ要件）
 * @param {Object} options - オプション設定
 * @returns {Object|null} Database spreadsheet object
 */
function openDatabase(options = {}) {
  try {
    const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
    if (!dbId) {
      console.warn('openDatabase: DATABASE_SPREADSHEET_ID not configured');
      return null;
    }

    const forceServiceAccount = true;

    const dataAccess = openSpreadsheet(dbId, {
      useServiceAccount: forceServiceAccount,
      context: 'database_access'
    });

    if (!dataAccess) {
      console.warn('openDatabase: Failed to access database via Sheets API, attempting SpreadsheetApp.openById fallback');

      try {
        const fallbackSpreadsheet = SpreadsheetApp.openById(dbId);
        return fallbackSpreadsheet;
      } catch (fallbackError) {
        logError_('openDatabase', fallbackError, { sheetsApiError: 'Failed via openSpreadsheet' });
        return null;
      }
    }

    return dataAccess?.spreadsheet || null;
  } catch (error) {
    logError_('openDatabase', error);
    return null;
  }
}

/**
 * 任意のスプレッドシートを開く。 アクセス経路は caller の役割で自動判定。
 *
 * 経路の自動振り分け (`validateServiceAccountUsage` の `accessMode` で決まる):
 *  - **owner (editor)** が自分の SS を読む → `SpreadsheetApp.openById` (own OAuth, SA quota 不要)
 *  - **viewer (生徒)** / **admin** が cross-user アクセス → SA pool 経由
 *  - **DB sheet** はどの caller でも SA pool 経由 (共有資源)
 *
 * v2782+ で旧 domain-wide sharing 経路は廃止し、 通常 Google Form 相当のセキュリティ
 * (viewer の Drive 非表示 + 直接編集不可) に統一。
 *
 * @param {string} spreadsheetId
 * @param {Object} [options]
 * @param {string} [options.context] - ログ識別用
 * @returns {{spreadsheet, auth, accessMode, getSheet}|null}
 */
function openSpreadsheet(spreadsheetId, options = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('openSpreadsheet: Invalid spreadsheet ID');
      return null;
    }

    const useServiceAccount = options.useServiceAccount !== false;

    const validation = validateServiceAccountUsage(
      spreadsheetId, useServiceAccount, options.context || 'openSpreadsheet'
    );
    if (!validation.allowed) {
      console.warn('openSpreadsheet: access denied:', validation.reason);
      return null;
    }

    // accessMode: 'own' → openById (owner's OAuth), 'sa' → SA pool
    const useSaPath = validation.accessMode === 'sa';
    let spreadsheet = null;
    let auth = null;

    try {
      if (useSaPath) {
        auth = getServiceAccount();
        if (auth.isValid) {
          spreadsheet = openSpreadsheetViaServiceAccount(spreadsheetId);
        }
        if (!spreadsheet) {
          // SA pool 全滅時の openById fallback は DB スプレッドシートに限定する。
          // user board に deploy 主を editor 追加していない (SA pool のみ) ので、
          // user board への openById は元々失敗するだけ。 ここで明示的に絞ることで
          // 「viewer は board SS に直接権限を持たない」セキュリティモデルを保証する。
          const dbId = getCachedProperty('DATABASE_SPREADSHEET_ID');
          if (dbId && spreadsheetId === dbId) {
            console.warn('openSpreadsheet: SA pool failed, falling back to openById (DB sheet only)', {
              spreadsheetId: `${spreadsheetId.substring(0, 20)}...`
            });
            spreadsheet = SpreadsheetApp.openById(spreadsheetId);
          } else {
            console.warn('openSpreadsheet: SA pool failed for user board, refusing owner fallback', {
              spreadsheetId: `${spreadsheetId.substring(0, 20)}...`
            });
            return null;
          }
        }
      } else {
        spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      }
    } catch (openError) {
      // 呼出元 (getUserSheetData 等) が 1 行 ERROR に集約する前提で WARN に降格。
      console.warn('openSpreadsheet: スプレッドシート接続失敗', {
        spreadsheetId: `${spreadsheetId.substring(0, 20)}...`,
        error: openError.message,
        accessMode: validation.accessMode,
        hasAuth: !!auth
      });
      return null;
    }

    return {
      spreadsheet,
      auth: auth || { isValid: false },
      accessMode: validation.accessMode,
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
  } catch (error) {
    logError_('openSpreadsheet', error);
    return null;
  }
}

// ---------------------------------------------------------------------
// サービスアカウント経路のヘルパー群
// openSpreadsheet内部からmodule-levelへ抽出（クロージャ未使用のため純粋な字句移動）。
// デバッグ時のスタックトレース可読性とテスト可能性の向上が目的。
// ---------------------------------------------------------------------

/**
 * SA pool から 1 個 pick → token 取得 → access verify → proxy 化。
 * verify で落ちた SA は次の候補に順次 fallback (1 個だけ DB に未共有でも継続)。
 * @param {string} sheetId
 * @returns {Object|null} SS proxy or null
 */
function openSpreadsheetViaServiceAccount(sheetId) {
  try {
    const pool = getAllServiceAccounts_();
    if (pool.length === 0) return null;

    const preferred = pickServiceAccount_();
    const startIdx = preferred ? pool.indexOf(preferred) : 0;
    const baseIdx = startIdx >= 0 ? startIdx : 0;

    for (let step = 0; step < pool.length; step++) {
      const sa = pool[(baseIdx + step) % pool.length];
      const accessToken = getServiceAccountAccessToken_(sa);
      if (!accessToken) continue;
      if (!verifyServiceAccountAccess_(sheetId, accessToken, sa.client_email)) continue;
      return createServiceAccountSpreadsheetProxy(sheetId, accessToken, sa.client_email);
    }
    return null;
  } catch (error) {
    logError_('openSpreadsheetViaServiceAccount', error);
    return null;
  }
}

function createServiceAccountJWT(header, payload, privateKey) {
  const headerB64 = Utilities.base64EncodeWebSafe(JSON.stringify(header)).replace(/=/g, '');
  const payloadB64 = Utilities.base64EncodeWebSafe(JSON.stringify(payload)).replace(/=/g, '');
  const signatureInput = `${headerB64}.${payloadB64}`;

  const signature = Utilities.computeRsaSha256Signature(signatureInput, privateKey);
  const signatureB64 = Utilities.base64EncodeWebSafe(signature).replace(/=/g, '');

  return `${signatureInput}.${signatureB64}`;
}

function createServiceAccountSpreadsheetProxy(sheetId, accessToken, saEmail) {
  const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;
  // resolveAuth: proxy が hot path で「現 SA が cooling なら別 SA に逃げる」を解決する closure。
  // SS / Sheet proxy 群はこれを共有することで session 中の 429 burst を回避する。
  const resolveAuth = makeProxyAuthResolver_(sheetId, accessToken, saEmail);

  return {
    getId: () => sheetId,
    getName: () => {
      try {
        const auth = resolveAuth();
        const response = fetchSheetsAPIWithRetry(
          `${baseUrl}?includeGridData=false&fields=properties.title`,
          { headers: { 'Authorization': `Bearer ${auth.token}` } },
          `getName(${sheetId.substring(0, 8)}...)`,
          auth.saEmail, resolveAuth
        );
        const data = JSON.parse(response.getContentText());
        return data.properties?.title || `スプレッドシート (ID: ${sheetId.substring(0, 8)}...)`;
      } catch (error) {
        console.warn('getName via API failed after retries:', error.message);
        return `スプレッドシート (ID: ${sheetId.substring(0, 8)}...)`;
      }
    },
    getSheetByName: (sheetName) => {
      const auth = resolveAuth();
      return createServiceAccountSheetProxy(sheetId, sheetName, auth.token, {}, auth.saEmail, resolveAuth);
    },
    getSheets: () => {
      try {
        const auth = resolveAuth();
        const response = fetchSheetsAPIWithRetry(
          `${baseUrl}?includeGridData=false`,
          { headers: { 'Authorization': `Bearer ${auth.token}` } },
          `getSheets(${sheetId.substring(0, 8)}...)`,
          auth.saEmail, resolveAuth
        );
        const data = JSON.parse(response.getContentText());
        const sheets = data.sheets || [];
        return sheets.map((sheetData) => {
          const properties = sheetData.properties || {};
          return createServiceAccountSheetProxy(sheetId, properties.title || 'Sheet1', auth.token, {
            sheetId: properties.sheetId,
            rowCount: properties.gridProperties?.rowCount || 1000,
            columnCount: properties.gridProperties?.columnCount || 26
          }, auth.saEmail, resolveAuth);
        });
      } catch (error) {
        console.warn('getSheets via API failed after retries:', error.message);
        return [];
      }
    }
  };
}

function createServiceAccountSheetProxy(sheetId, sheetName, accessToken, additionalInfo = {}, saEmail, parentResolveAuth) {
  const baseUrl = `https://sheets.googleapis.com/v4/spreadsheets/${sheetId}`;
  // parent から渡された resolver を使い回せば、 cooldown 解決が連鎖して同 session 内で
  // 429 burst を回避できる。 旧 API 互換のため parent 未指定時は自前 resolver を作る。
  const resolveAuth = parentResolveAuth || makeProxyAuthResolver_(sheetId, accessToken, saEmail);
  let cachedDimensions = null;

  function fetchDimensionsOnce() {
    if (cachedDimensions) return cachedDimensions;
    try {
      const auth = resolveAuth();
      const response = fetchSheetsAPIWithRetry(
        `${baseUrl}?includeGridData=false&fields=sheets(properties(title,gridProperties))`,
        { headers: { 'Authorization': `Bearer ${auth.token}` } },
        `fetchDimensions(${sheetName})`,
        auth.saEmail, resolveAuth
      );
      const data = JSON.parse(response.getContentText());
      const sheets = data.sheets || [];
      const targetSheet = sheets.find((s) => s.properties && s.properties.title === sheetName);
      cachedDimensions = {
        rowCount: targetSheet?.properties?.gridProperties?.rowCount || 1000,
        columnCount: targetSheet?.properties?.gridProperties?.columnCount || 26
      };
      return cachedDimensions;
    } catch (error) {
      console.warn(`fetchDimensions failed for ${sheetName}:`, error.message);
      return { rowCount: 1000, columnCount: 26 };
    }
  }

  // setValue(s) 共通の RAW PUT。 resource (range or sheetName) / values / label のみ可変。
  //   getRange().setValue/setValues と getDataRange().setValues の 3 箇所で共有。
  const putValuesViaApi = (resource, values, label) => {
    const auth = resolveAuth();
    return fetchSheetsAPIWithRetry(
      `${baseUrl}/values/${resource}?valueInputOption=RAW`,
      {
        method: 'PUT',
        headers: { 'Authorization': `Bearer ${auth.token}`, 'Content-Type': 'application/json' },
        payload: JSON.stringify({ values })
      },
      label,
      auth.saEmail, resolveAuth
    );
  };

  return {
    getName: () => sheetName,
    getSheetId: () => additionalInfo.sheetId || 0,
    getLastRow: () => {
      if (additionalInfo.rowCount) return additionalInfo.rowCount;
      return fetchDimensionsOnce().rowCount;
    },
    getLastColumn: () => {
      if (additionalInfo.columnCount) return additionalInfo.columnCount;
      return fetchDimensionsOnce().columnCount;
    },
    getRange: (row, col, numRows, numCols) => {
      const range = numRows && numCols
        ? `${sheetName}!R${row}C${col}:R${row + numRows - 1}C${col + numCols - 1}`
        : `${sheetName}!R${row}C${col}`;

      return {
        getValues: () => {
          try {
            const auth = resolveAuth();
            const response = fetchSheetsAPIWithRetry(
              `${baseUrl}/values/${range}`,
              { headers: { 'Authorization': `Bearer ${auth.token}` } },
              `getRange.getValues(${range})`,
              auth.saEmail, resolveAuth
            );
            const data = JSON.parse(response.getContentText());
            return data.values || [];
          } catch (error) {
            console.warn('getValues via API failed after retries:', error.message);
            return [];
          }
        },
        setValue: (value) => {
          try {
            return putValuesViaApi(range, [[value]], `setValue(${range})`);
          } catch (error) {
            console.warn('setValue via API failed after retries:', error.message);
            throw error;
          }
        },
        setValues: (values) => {
          try {
            return putValuesViaApi(range, values, `setValues(${range})`);
          } catch (error) {
            console.warn('setValues via API failed after retries:', error.message);
            throw error;
          }
        }
      };
    },
    getDataRange: () => {
      return {
        getValues: () => {
          try {
            const auth = resolveAuth();
            const response = fetchSheetsAPIWithRetry(
              `${baseUrl}/values/${sheetName}`,
              { headers: { 'Authorization': `Bearer ${auth.token}` } },
              `getDataRange(${sheetName})`,
              auth.saEmail, resolveAuth
            );
            const data = JSON.parse(response.getContentText());
            return data.values || [];
          } catch (error) {
            console.warn('getDataRange via API failed after retries:', error.message);
            return [];
          }
        },
        setValues: (values) => {
          try {
            return putValuesViaApi(sheetName, values, `getDataRange.setValues(${sheetName})`);
          } catch (error) {
            console.warn('getDataRange setValues via API failed after retries:', error.message);
            throw error;
          }
        }
      };
    },
    appendRow: (rowData) => {
      try {
        const auth = resolveAuth();
        const payload = { values: [rowData] };
        return fetchSheetsAPIWithRetry(
          `${baseUrl}/values/${sheetName}:append?valueInputOption=RAW`,
          {
            method: 'POST',
            headers: { 'Authorization': `Bearer ${auth.token}`, 'Content-Type': 'application/json' },
            payload: JSON.stringify(payload)
          },
          `appendRow(${sheetName})`,
          auth.saEmail, resolveAuth
        );
      } catch (error) {
        console.warn('appendRow via API failed after retries:', error.message);
        throw error;
      }
    }
  };
}

/**
 * ユーザー検索の呼び出し元メールアドレスを解決
 * @param {Object} context - アクセスコンテキスト
 * @returns {string|null} リクエストユーザーのメール
 */
function resolveRequestingUser(context = {}) {
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
function canAccessTargetUser(targetUser, context = {}) {
  if (!targetUser) {
    return false;
  }

  if (context.skipAccessCheck === true) {
    return true;
  }

  const requestingUser = resolveRequestingUser(context);
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
function applyUserAccessControl(user, context = {}, operation = 'user_lookup') {
  if (!user) {
    return null;
  }

  if (canAccessTargetUser(user, context)) {
    return user;
  }

  const requestingUser = resolveRequestingUser(context);
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
 * @param {boolean} context.forceServiceAccount - サービスアカウント強制使用
 * @param {string} context.requestingUser - リクエストユーザー（デバッグ用）
 * @returns {Object|null} User object
 */
function findUserByEmail(email, context = {}) {
  if (!email || !validateEmail(email).isValid) return null;
  return findUserByField('userEmail', email, { ...context, cacheKeyPrefix: 'user_by_email', label: 'findUserByEmail' });
}

/**
 * ユーザーIDでユーザーを検索（strict: admin / 自分のみ許可）
 *
 * WRITE操作や管理者専用操作で使う。公開ボードを閲覧する生徒の経路では
 * 代わりに {@link findPublishedBoardOwner} を使うこと。これを逆に使うと
 * 生徒から null が返り、「config 取得失敗 → default config → formUrl 消失」
 * のような silent failure が起きる（実際に発生した bug）。
 *
 * @param {string} userId - ユーザーID
 * @param {Object} context - アクセスコンテキスト
 * @param {string} context.requestingUser - リクエストユーザー（debug/権限判定用）
 * @param {boolean} [context.allowPublishedRead] - 公開ボードなら非オーナーも許可
 * @param {boolean} [context.forceServiceAccount] - サービスアカウント強制使用
 * @returns {Object|null} User object
 */
function findUserById(userId, context = {}) {
  if (!userId) return null;
  return findUserByField('userId', userId, { ...context, cacheKeyPrefix: 'user_by_id', label: 'findUserById' });
}

/**
 * 公開ボードのオーナーを viewer 視点で取得（admin / owner / 公開ボード閲覧者を許可）
 *
 * `getPublishedSheetData` / `getActiveFormInfo` / `getNotificationUpdate` など、
 * 生徒が他人のボードを見る経路で使う。これを使うと `allowPublishedRead: true` を
 * 渡し忘れる類の bug を根絶できる。
 *
 * WRITE 系や editor-only 操作では `findUserById` + 独自の editor gate を使うこと
 * （例: deleteAnswerRow、toggleHighlight）。
 *
 * @param {string} userId - 対象ボードオーナーの userId
 * @param {string} viewerEmail - 閲覧者のメールアドレス
 * @param {Object} [extra] - 追加 context（preloadedAuth 等）
 */
function findPublishedBoardOwner(userId, viewerEmail, extra = {}) {
  if (!userId) return null;
  return findUserById(userId, {
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
function findUserByGoogleId(googleId, context = {}) {
  if (!googleId) return null;
  // Note: historically this path did not populate an individual cache entry, so
  // we explicitly opt out via cacheKeyPrefix=null to preserve behavior.
  return findUserByField('googleId', googleId, { ...context, cacheKeyPrefix: null, label: 'findUserByGoogleId' });
}

/**
 * findUserBy*系の共通実装。
 * 探索は (1) individual cache → (2) getAllUsers キャッシュ → (3) 直接DB の順。
 * @param {string} fieldName - 検索するユーザーフィールド ('userEmail' | 'userId' | 'googleId')
 * @param {*} fieldValue - 一致させる値
 * @param {Object} context - requestingUser / preloadedAuth / forceServiceAccount 等に加え、
 *                          内部用の cacheKeyPrefix (null で個別キャッシュ無効) と label を受け取る
 * @returns {Object|null} アクセス制御適用済みのユーザー、または null
 */
function findUserByField(fieldName, fieldValue, context = {}) {
  const { cacheKeyPrefix = null, label = `findUserBy_${fieldName}` } = context;

  try {
    let individualCacheKey = null;
    if (cacheKeyPrefix) {
      const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
      individualCacheKey = `${cacheKeyPrefix}_v${cacheVersion}_${fieldValue}`;
      try {
        const cached = CacheService.getScriptCache().get(individualCacheKey);
        if (cached) {
          const cachedUser = JSON.parse(cached);
          return applyUserAccessControl(cachedUser, context, label);
        }
      } catch (individualCacheError) {
        logError_(label + '.individualCache', individualCacheError);
      }
    }

    try {
      const allUsers = getAllUsers(
        { activeOnly: false },
        { ...context, forceServiceAccount: true, skipCache: false }
      );
      if (Array.isArray(allUsers) && allUsers.length > 0) {
        const user = allUsers.find((u) => u[fieldName] === fieldValue);
        if (user) {
          const allowedUser = applyUserAccessControl(user, context, label);
          if (!allowedUser) return null;
          if (individualCacheKey) {
            saveToCacheWithSizeCheck(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
          }
          return allowedUser;
        }
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
      const allowedUser = applyUserAccessControl(user, context, label);
      if (!allowedUser) return null;
      if (individualCacheKey) {
        saveToCacheWithSizeCheck(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
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
        const allowedUser = applyUserAccessControl(user, context, label);
        if (!allowedUser) return null;
        if (individualCacheKey) {
          saveToCacheWithSizeCheck(individualCacheKey, allowedUser, CACHE_DURATION.USER_INDIVIDUAL);
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
function createUser(email, initialConfig = {}, context = {}) {
  const lock = LockService.getScriptLock();

  try {
    if (!email || !validateEmail(email).isValid) {
      console.warn('createUser: Invalid email provided:', typeof email, email);
      return null;
    }

    if (!lock.tryLock(10000)) { // 10秒待機
      console.warn('createUser: Lock timeout - concurrent user creation detected');
      return null;
    }

    const currentEmail = getCurrentEmail();

    const existingUser = findUserByEmail(email, {
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

    if (sheet.getLastRow() >= 10000) {
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

    clearDatabaseUserCache();

    return user;

  } catch (error) {
    logError_('createUser', error);
    return null;
  } finally {
    try {
      lock.releaseLock();
    } catch (unlockError) {
      console.warn('createUser: Lock release failed:', unlockError.message);
    }
  }
}

/**
 * 全ユーザーリストを取得（CLAUDE.md準拠 - Admin専用）
 * @param {Object} options - オプション設定
 * @param {Object} context - アクセスコンテキスト
 * @returns {Array} User list
 */
function getAllUsers(options = {}, context = {}) {
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

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('getAllUsers: Non-admin user attempting cross-user data access');
      return [];
    }

    const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
    const cacheKey = `all_users_v${cacheVersion}_${simpleHash(options)}_${context.forceServiceAccount ? 'sa' : 'norm'}`;
    const skipCache = context.skipCache || false;

    if (!skipCache) {
      try {
        const cached = CacheService.getScriptCache().get(cacheKey);
        if (cached) {
          return JSON.parse(cached);
        }
      } catch (cacheError) {
        logError_('getAllUsers.cacheRead', cacheError);
      }
    }

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

    if (!skipCache) {
      // Why: 100+ユーザでJSONが100KB超えるとCacheService.putが黙って失敗する。
      //      saveToCacheWithSizeCheckでサイズ超過を検出しログに残す。
      saveToCacheWithSizeCheck(cacheKey, users, CACHE_DURATION.DATABASE_LONG);
    }

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
function clearDatabaseUserCache() {
  try {
    const props = PropertiesService.getScriptProperties();
    const currentVersion = parseInt(props.getProperty('USER_CACHE_VERSION') || '0');
    const newVersion = currentVersion + 1;

    props.setProperty('USER_CACHE_VERSION', newVersion.toString());

    clearPropertyCache('USER_CACHE_VERSION');
  } catch (error) {
    logError_('clearDatabaseUserCache', error);
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
function updateUser(userId, updates, context = {}) {
  const lock = LockService.getScriptLock();

  try {
    if (!lock.tryLock(5000)) { // 5秒待機
      console.warn('updateUser: Lock timeout - concurrent update detected');
      return { success: false, message: 'Concurrent update in progress. Please retry.' };
    }

    const requestingUser = context.requestingUser || getCurrentEmail();
    const targetUser = findUserById(userId, {
      ...context,
      requestingUser
    });

    if (!targetUser) {
      console.warn('updateUser: Target user not found:', userId);
      return { success: false, message: 'User not found' };
    }

    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('updateUser: Database access failed');
      return { success: false, message: 'Database access failed' };
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('updateUser: Users sheet not found');
      return { success: false, message: 'Users sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('updateUser: UserId column not found');
      return { success: false, message: 'UserId column not found' };
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdColumnIndex] === userId) {
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

          const rangeToUpdate = sheet.getRange(i + 1, minCol, 1, colSpan);
          const [currentRowData] = rangeToUpdate.getValues();

          updateCells.forEach(({ col, value }) => {
            currentRowData[col - minCol] = value;
          });

          rangeToUpdate.setValues([currentRowData]);
        }

        clearDatabaseUserCache();

        return { success: true };
      }
    }

    console.warn('updateUser: User not found:', userId);
    return { success: false, message: 'User not found' };
  } catch (error) {
    logError_('updateUser', error);
    return { success: false, message: error.message || 'Unknown error occurred' };
  } finally {
    try {
      lock.releaseLock();
    } catch (unlockError) {
      console.warn('updateUser: Lock release failed:', unlockError.message);
    }
  }
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
function findUserBySpreadsheetId(spreadsheetId, context = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      console.warn('findUserBySpreadsheetId: Invalid spreadsheetId provided:', typeof spreadsheetId);
      return null;
    }

    const cacheVersion = getCachedProperty('USER_CACHE_VERSION') || '0';
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

    const allUsers = getAllUsers({ activeOnly: false }, { ...context, forceServiceAccount: true, preloadedAuth: context.preloadedAuth });
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
        console.warn(`findUserBySpreadsheetId: Failed to parse configJSON for user ${user.userId}:`, parseError.message);
        continue;
      }
    }

    // Why (negative cache 5s): 旧版は 60s null cache で「Target user not found」 が
    //   1 分間ハマる事故を起こしていた (profile 切替 transient + getAllUsers stale fetch)。
    //   stale SS URL の防御目的なら 5s で十分 (polling ~ 8s なので次回ヒット時に再判定)。
    //   USER_CACHE_VERSION bump で proactive invalidation も働くが、 fallback として TTL も短縮。
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

/**
 * ユーザーを削除（CLAUDE.md準拠 - Admin専用）
 * @param {string} userId - ユーザーID
 * @param {string} reason - 削除理由
 * @param {Object} context - アクセスコンテキスト
 * @returns {Object} Delete operation result
 */
function deleteUser(userId, reason = '', context = {}) {
  // Matches create/updateUser concurrency contract: without the lock, a
  // concurrent updateUser can land on a row index that delete just shifted.
  const lock = LockService.getScriptLock();

  try {
    const currentEmail = getCurrentEmail();
    const isAdmin = isAdministrator(currentEmail);

    if (!isAdmin && !context.forceServiceAccount) {
      console.warn('deleteUser: Non-admin user attempting user deletion:', userId);
      return { success: false, message: 'Insufficient permissions for user deletion' };
    }

    if (!lock.tryLock(10000)) {
      console.warn('deleteUser: Lock timeout - concurrent user modification detected');
      return { success: false, message: 'Concurrent modification in progress. Please retry.' };
    }

    const spreadsheet = openDatabase();
    if (!spreadsheet) {
      console.warn('deleteUser: Database access failed');
      return { success: false, message: 'Database access failed' };
    }

    const sheet = spreadsheet.getSheetByName('users');
    if (!sheet) {
      console.warn('deleteUser: Users sheet not found');
      return { success: false, message: 'Users sheet not found' };
    }

    const data = sheet.getDataRange().getValues();
    const [headers] = data;
    const userIdColumnIndex = headers.indexOf('userId');

    if (userIdColumnIndex === -1) {
      console.warn('deleteUser: UserId column not found');
      return { success: false, message: 'UserId column not found' };
    }

    const rowsToDelete = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][userIdColumnIndex] === userId) {
        rowsToDelete.push(i + 1); // シート行番号（1-indexed）
      }
    }

    if (rowsToDelete.length === 0) {
      console.warn('deleteUser: User not found:', userId);
      return { success: false, message: 'User not found' };
    }

    for (let i = rowsToDelete.length - 1; i >= 0; i--) {
      sheet.deleteRows(rowsToDelete[i], 1);
    }

    clearDatabaseUserCache();

    return {
      success: true,
      userId,
      deleted: true,
      deletedRows: rowsToDelete.length,
      reason: reason || 'No reason provided'
    };
  } catch (error) {
    logError_('deleteUser', error);
    return { success: false, message: error.message };
  } finally {
    try {
      lock.releaseLock();
    } catch (unlockError) {
      console.warn('deleteUser: Lock release failed:', unlockError.message);
    }
  }
}
