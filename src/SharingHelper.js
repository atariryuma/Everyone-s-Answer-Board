/**
 * @fileoverview SharingHelper - スプレッドシート共有設定 (SA-only モデル)。
 *
 * 通常 Google Form と同等のセキュリティモデル: 回答 SS は owner と SA pool だけがアクセス可能。
 * viewer (生徒) は GAS Web App 経由でのみ board を見る (SA proxy 経由)。 viewer の Drive には
 * 表示されず、 viewer が SS を直接編集することも不可。
 *
 * 旧 domain-wide sharing (`DOMAIN_WITH_LINK + EDIT`) は廃止 (v2782 refactor)。
 * 既存ボードの cleanup は `migrateBoardSharing` admin API で実施。
 */

/* global getCachedProperty, logError_, getAllServiceAccounts_, parseServiceAccountCredsSoft_, invalidateSaCache_ */

/**
 * SA pool の全 SA を SS の editor として追加する (冪等)。
 *
 * Why: DatabaseCore の Sheets REST API + JWT 経路は SA が editor でない SS で 403。
 *   pool 内の任意の SA が pick されるので、 全員を editor 登録しておかないと「pick された SA
 *   が共有漏れで verify 失敗 → 次の SA に fallback」のコストが毎回発生する。
 *
 * @param {string} spreadsheetId
 * @returns {{success:boolean, added:string[], errors:string[]}}
 */
function addServiceAccountsAsEditors(spreadsheetId) {
  const result = { success: false, added: [], errors: [] };
  try {
    const pool = (typeof getAllServiceAccounts_ === 'function') ? getAllServiceAccounts_() : [];
    if (pool.length === 0) {
      result.errors.push('SA pool not configured');
      return result;
    }
    const ss = SpreadsheetApp.openById(spreadsheetId);
    for (let i = 0; i < pool.length; i++) {
      const sa = pool[i];
      try {
        // addEditor は冪等 (既存 editor に再追加しても no-op)。 getEditors の pre-check は
        // 余分な RPC で getEditors 自体が ScriptError を投げるケースもあるため省略。
        ss.addEditor(sa.client_email);
        result.added.push(sa.client_email);
        // 過去の 'no' verify cache を吹き飛ばす (今回 add で OK になる可能性が高い)。
        if (typeof invalidateSaCache_ === 'function') {
          invalidateSaCache_(spreadsheetId, sa.client_email);
        }
      } catch (err) {
        result.errors.push(`${sa.client_email}: ${err.message}`);
      }
    }
    result.success = result.added.length > 0;
    return result;
  } catch (err) {
    if (typeof logError_ === 'function') logError_('addServiceAccountsAsEditors', err);
    else console.error('addServiceAccountsAsEditors:', err && err.message);
    result.errors.push(err && err.message ? err.message : String(err));
    return result;
  }
}

/**
 * 後方互換 wrapper: primary SA だけを editor 追加する旧 API。
 * 新規コードは `addServiceAccountsAsEditors` を使うこと。
 * @deprecated
 */
function addServiceAccountAsEditor(spreadsheetId) {
  try {
    const pool = (typeof getAllServiceAccounts_ === 'function') ? getAllServiceAccounts_() : [];
    if (pool.length === 0) {
      return { success: false, added: false, message: 'SA pool not configured' };
    }
    const primary = pool[0];
    SpreadsheetApp.openById(spreadsheetId).addEditor(primary.client_email);
    if (typeof invalidateSaCache_ === 'function') {
      invalidateSaCache_(spreadsheetId, primary.client_email);
    }
    return { success: true, added: true, saEmail: primary.client_email, message: 'Service account added as editor' };
  } catch (error) {
    if (typeof logError_ === 'function') logError_('addServiceAccountAsEditor', error);
    return { success: false, added: false, message: error.message };
  }
}

/**
 * 新規作成 SS に共有デフォルトを一括適用する。
 *   SA pool 全員を editor 追加 (必須 — 未設定だと viewer 経路が 403)。
 *
 * v2782 で domain-wide sharing は廃止。 viewer の Drive にボード SS が表示されない・
 * viewer が SS を直接編集不可、 という通常 Form 相当のセキュリティモデルに統一。
 *
 * @param {string} spreadsheetId
 * @param {string} [ownerEmail] - 後方互換用 (現在は未使用)
 * @returns {{saAdded:boolean, saEmails:string[], errors:string[]}}
 */
function applySpreadsheetSharingDefaults(spreadsheetId, ownerEmail) {
  // ownerEmail は domain-share 時代の引数。 v2782+ では使わないが、 旧 caller 互換のため残す。
  void ownerEmail;
  const result = { saAdded: false, saEmails: [], errors: [] };
  try {
    const sa = addServiceAccountsAsEditors(spreadsheetId);
    result.saAdded = !!(sa && sa.success);
    result.saEmails = (sa && sa.added) || [];
    if (sa && sa.errors && sa.errors.length) result.errors.push(...sa.errors.map((e) => 'SA: ' + e));
  } catch (saError) {
    result.errors.push('SA: ' + saError.message);
  }
  return result;
}
