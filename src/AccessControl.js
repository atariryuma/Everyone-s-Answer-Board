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
 *   - 管理 API (admin diagnostics)
 *
 * 効率:
 *   - 実行者 = 先生 = 持ち主の権限で、DriveApp から編集者と持ち主を 1 回読む
 *   - 結果は ScriptCache に cache (per file × per email hash)
 *   - 付与 ('1') は短命 (2 min): editor 共有を revoke しても権限残存が最大 2 分で消える
 *     (collaborator は unpublish 等の権限を持つため、 revoke 後の stale grant 窓を短く保つ)
 *   - 拒否 ('0') は長命 (10 min): 誤って権限が増えることはないので Drive API 削減を優先
 *   - owner / admin 判定は呼び出し側で先に行うこと (Drive API 呼び出し回避)
 */

/* global getUserConfig, logError_ */

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
    // targetUser (users シートの行) を preloaded で渡す。
    //   Why: 渡さないと getUserConfig が findUserById を viewer (児童) の権限で引き直し、
    //   allowPublishedRead 無しなので必ず "findUserById: Access denied" を WARN に出しつつ
    //   default config を返していた。児童のページ表示ごとに 1 件、授業中は数十件の偽警告が
    //   本物の異常を埋めていた (2026-09-08)。行はもう手元にあるので引き直す理由がない。
    res = (typeof getUserConfig === 'function') ? getUserConfig(targetUser.userId, targetUser) : null;
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
  let cache = null;
  try { cache = CacheService.getScriptCache(); } catch (_) { cache = null; /* cache 不通は機能的に無害 (毎回 Drive を読むだけ) */ }
  const key = 'collab:' + fileId + ':' + __emailHash_(emailNorm);
  let hit = null;
  try { hit = cache ? cache.get(key) : null; } catch (_) { hit = null; }
  if (hit !== null) return hit === '1';

  let ok = false;
  // Drive が読めなかったとき (一時的な失敗) は「編集者ではない」を 10 分固定しない (cache しない)。
  let transient = false;
  try {
    // 先生として実行するので、先生のボード SS の編集者を DriveApp でそのまま読める。
    ok = __editorEmailsViaDriveApp_(fileId).indexOf(emailNorm) !== -1;
  } catch (e) {
    transient = true;
    if (typeof logError_ === 'function') logError_('isBoardCollaborator/drive', e);
  }
  if (cache && !(transient && !ok)) {
    try { cache.put(key, ok ? '1' : '0', ok ? COLLABORATOR_CACHE_TTL_OK_SEC : COLLABORATOR_CACHE_TTL_NO_SEC); }
    catch (_) { /* cache 不通は機能的に無害 (次回 Drive を読み直す) */ }
  }
  return ok;
}

// ファイルの編集者と持ち主のメール (小文字)。実行者 (先生) がそのファイルを読める前提。
function __editorEmailsViaDriveApp_(fileId) {
  const file = DriveApp.getFileById(fileId);
  const emails = file.getEditors().map(function (u) { return String(u.getEmail() || '').toLowerCase(); });
  const owner = file.getOwner();
  if (owner) emails.push(String(owner.getEmail() || '').toLowerCase());
  return emails;
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
