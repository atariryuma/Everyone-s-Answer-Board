'use strict';

/**
 * @fileoverview ボード共同編集者 (collaborator) 認可。
 *
 * 設計判断: ボード SS (Drive file) に editor 権限 (role: writer / owner) を持つメールは
 *   owner と同じ「ボード進行操作」 を実行できる。 SS と Form は通常同じ教師チームが
 *   editor 共有するため、 SS editor チェックで実用十分にカバー (Form の publish URL からは
 *   Drive permission を取得不可なので Form 単独 editor は対象外)。
 *
 * 解放範囲 (collaborator 可):
 *   - ハイライト toggle (highlight)
 *   - リアクション操作
 *   - 公開停止 (unpublishBoard, 緊急時)
 *   - ボード閲覧 (mode=view) で editor UI 表示
 *   - lesson review (mode=review) 閲覧
 *
 * 解放しない範囲 (owner / admin 専用に維持):
 *   - ボード設定の永続変更 (publishApp / saveUserConfig / column mapping / displaySettings)
 *   - 再公開 (republishMyBoard / publish from unpublished)
 *   - プロファイル管理 (loadProfile / deleteProfile / 別名保存)
 *   - lesson の作成 / 編集 / 削除 / archive
 *   - 管理 API (SA pool, admin diagnostics, all-user 操作)
 *
 * 効率:
 *   - SA pool 経由で Drive REST API permissions.list を 1 回 fetch
 *   - 結果は ScriptCache に cache (per file × per email hash)
 *   - 付与 ('1') は短命 (2 min): editor 共有を revoke しても権限残存が最大 2 分で消える
 *     (collaborator は unpublish 等の権限を持つため、 revoke 後の stale grant 窓を短く保つ)
 *   - 拒否 ('0') は長命 (10 min): 誤って権限が増えることはないので Drive API 削減を優先
 *   - owner / admin 判定は呼び出し側で先に行うこと (Drive API 呼び出し回避)
 */

/* global pickServiceAccount_, getServiceAccountAccessToken_, getUserConfig, logError_ */

const COLLABORATOR_CACHE_TTL_OK_SEC = 120;
const COLLABORATOR_CACHE_TTL_NO_SEC = 600;

/**
 * targetUser のボード SS に viewerEmail が editor / owner 権限を持つか。
 *
 * @param {Object} targetUser - users sheet レコード (userId, userEmail, ...)
 * @param {string} viewerEmail - 判定対象のメール
 * @returns {boolean}
 */
function isBoardCollaborator(targetUser, viewerEmail) {
  if (!targetUser || !viewerEmail) return false;
  const norm = String(viewerEmail).toLowerCase().trim();
  if (!norm) return false;
  // owner 自身は collaborator ではなく owner 扱い (呼び出し側で別 path)
  if (norm === String(targetUser.userEmail || '').toLowerCase().trim()) return false;

  let res;
  try {
    res = (typeof getUserConfig === 'function') ? getUserConfig(targetUser.userId) : null;
  } catch (e) {
    if (typeof logError_ === 'function') logError_('isBoardCollaborator/config', e);
    return false;
  }
  // getUserConfig は envelope {success, config, corrupted, ...} を返す。spreadsheetId は
  //   res.config.spreadsheetId にある (envelope 直読みは常に undefined で、v2855 collaborator
  //   認可が実質無効化されていた)。他の呼び出し側 (AdminApis) と同じく .config で unwrap する。
  // 認可判定なので corrupted config では fail-closed: 破損 config の default には spreadsheetId
  //   が無く、あっても信頼できないため collaborator を認可しない (AdminApis の publish 状態変更と
  //   同じ扱い)。
  const cfg = (res && res.success && !res.corrupted) ? res.config : null;
  const ssId = cfg && cfg.spreadsheetId;
  if (!ssId) return false;

  return __hasEditorPermissionViaDrive_(ssId, norm);
}

function __hasEditorPermissionViaDrive_(fileId, emailNorm) {
  const cache = CacheService.getScriptCache();
  const key = 'collab:' + fileId + ':' + __emailHash_(emailNorm);
  const hit = cache.get(key);
  if (hit !== null) return hit === '1';

  let ok = false;
  try {
    const sa = pickServiceAccount_();
    if (!sa) return false;
    const token = getServiceAccountAccessToken_(sa);
    if (!token) return false;
    const url = 'https://www.googleapis.com/drive/v3/files/' + encodeURIComponent(fileId) +
                '/permissions?fields=permissions(emailAddress,role)&supportsAllDrives=true';
    const response = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + token },
      muteHttpExceptions: true
    });
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      const editors = (json.permissions || [])
        .filter(function (p) { return p.role === 'writer' || p.role === 'owner'; })
        .map(function (p) { return String(p.emailAddress || '').toLowerCase(); });
      ok = editors.indexOf(emailNorm) !== -1;
    }
  } catch (e) {
    if (typeof logError_ === 'function') logError_('isBoardCollaborator/drive', e);
  }
  cache.put(key, ok ? '1' : '0', ok ? COLLABORATOR_CACHE_TTL_OK_SEC : COLLABORATOR_CACHE_TTL_NO_SEC);
  return ok;
}

function __emailHash_(email) {
  // SHA-256 を使う。 旧実装は 32bit 非暗号ハッシュで、 衝突する 2 メールが
  // 同じ collab cache key を共有し「他人の editor 判定 '1'」を継承する認可リスクがあった。
  // cache key は 250 文字まで許容されるので 64 桁 hex でも余裕。
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(email), Utilities.Charset.UTF_8);
  let hex = '';
  for (let i = 0; i < bytes.length; i++) {
    // byte は -128..127 なので 0xFF マスクして 2 桁 hex に正規化
    hex += ('0' + (bytes[i] & 0xff).toString(16)).slice(-2);
  }
  return hex;
}
