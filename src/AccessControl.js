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
 *   - 結果は ScriptCache に 10 min cache (per file × per email hash)
 *   - owner / admin 判定は呼び出し側で先に行うこと (Drive API 呼び出し回避)
 */

/* global UrlFetchApp, CacheService, pickServiceAccount_, getServiceAccountAccessToken_, getUserConfig, logError_ */

const COLLABORATOR_CACHE_TTL_SEC = 600;

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

  let config;
  try {
    config = (typeof getUserConfig === 'function') ? getUserConfig(targetUser.userId) : null;
  } catch (e) {
    if (typeof logError_ === 'function') logError_('isBoardCollaborator/config', e);
    return false;
  }
  const ssId = config && config.spreadsheetId;
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
  cache.put(key, ok ? '1' : '0', COLLABORATOR_CACHE_TTL_SEC);
  return ok;
}

function __emailHash_(email) {
  let h = 0;
  for (let i = 0; i < email.length; i++) {
    h = ((h << 5) - h) + email.charCodeAt(i);
    h |= 0;
  }
  return Math.abs(h).toString(36);
}
