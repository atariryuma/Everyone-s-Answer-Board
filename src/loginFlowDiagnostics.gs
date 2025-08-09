/**
 * @fileoverview Login flow diagnostics helpers
 * 手元で実行できる軽量の計測関数群。getLoginStatus / confirmUserRegistration /
 * verifyAdminAccess を順に計測し、所要時間と結果をまとめて返します。
 */

function diagGetLoginStatus() {
  const startedAt = Date.now();
  try {
    const email = getCurrentUserEmail();
    const result = getLoginStatus();
    const duration = Date.now() - startedAt;
    infoLog('diagGetLoginStatus', { email, durationMs: duration, resultStatus: result && result.status });
    return { ok: true, email, durationMs: duration, result };
  } catch (e) {
    const duration = Date.now() - startedAt;
    errorLog('diagGetLoginStatus error', { durationMs: duration, error: e.message });
    return { ok: false, durationMs: duration, error: e.message };
  }
}

function diagConfirmRegistration() {
  const startedAt = Date.now();
  try {
    const email = getCurrentUserEmail();
    const result = confirmUserRegistration();
    const duration = Date.now() - startedAt;
    infoLog('diagConfirmRegistration', { email, durationMs: duration, resultStatus: result && result.status });
    return { ok: true, email, durationMs: duration, result };
  } catch (e) {
    const duration = Date.now() - startedAt;
    errorLog('diagConfirmRegistration error', { durationMs: duration, error: e.message });
    return { ok: false, durationMs: duration, error: e.message };
  }
}

function diagVerifyAdminAccessForCurrentUser() {
  const startedAt = Date.now();
  try {
    const email = getCurrentUserEmail();
    const user = findUserByEmail(email);
    if (!user) {
      return { ok: false, durationMs: Date.now() - startedAt, error: 'user not found', email };
    }
    const ok = verifyAdminAccess(user.userId);
    const urls = generateUserUrls(user.userId);
    const duration = Date.now() - startedAt;
    infoLog('diagVerifyAdminAccessForCurrentUser', { email, userId: user.userId, ok, durationMs: duration });
    return { ok, durationMs: duration, email, userId: user.userId, adminUrl: urls && urls.adminUrl };
  } catch (e) {
    const duration = Date.now() - startedAt;
    errorLog('diagVerifyAdminAccessForCurrentUser error', { durationMs: duration, error: e.message });
    return { ok: false, durationMs: duration, error: e.message };
  }
}

/**
 * ログインフロースモークテスト
 * 1) getLoginStatus → 2) unregisteredなら confirmUserRegistration → 3) getLoginStatus 再確認
 * 最後に verifyAdminAccess を実施し、adminUrl を返します。
 */
function diagRunLoginFlowSmokeTest() {
  const allStart = Date.now();
  const email = getCurrentUserEmail();
  const out = { email, timestamp: new Date().toISOString() };

  try {
    // 1) 初回ステータス
    const t1 = Date.now();
    const status1 = getLoginStatus();
    const d1 = Date.now() - t1;
    out.status1 = { durationMs: d1, result: status1 };

    // 2) 未登録なら登録処理
    let regDuration = 0;
    if (status1 && status1.status === 'unregistered') {
      const t2 = Date.now();
      const reg = confirmUserRegistration();
      regDuration = Date.now() - t2;
      out.registration = { durationMs: regDuration, result: reg };
    }

    // 3) 再ステータス
    const t3 = Date.now();
    const status2 = getLoginStatus();
    const d3 = Date.now() - t3;
    out.status2 = { durationMs: d3, result: status2 };

    // 4) ユーザー解決と権限確認
    const userId = (status2 && status2.userId) || (findUserByEmail(email) && findUserByEmail(email).userId) || null;
    let verify = null;
    let adminUrl = '';
    let d4 = 0;
    if (userId) {
      const t4 = Date.now();
      verify = verifyAdminAccess(userId);
      d4 = Date.now() - t4;
      const urls = generateUserUrls(userId);
      adminUrl = urls && urls.adminUrl || '';
    }
    out.verify = { userId, ok: !!verify, durationMs: d4 };
    out.adminUrl = adminUrl;

    out.totalDurationMs = Date.now() - allStart;
    infoLog('diagRunLoginFlowSmokeTest', {
      email,
      totalMs: out.totalDurationMs,
      status1Ms: d1,
      regMs: regDuration,
      status2Ms: d3,
      verifyMs: out.verify.durationMs,
      finalStatus: status2 && status2.status
    });

    return out;
  } catch (e) {
    out.totalDurationMs = Date.now() - allStart;
    out.error = e && e.message;
    errorLog('diagRunLoginFlowSmokeTest error', { email, totalMs: out.totalDurationMs, error: out.error });
    return out;
  }
}

