/**
 * @fileoverview 権限管理サービス
 * システム内のリソースへのアクセス権を検証する機能を提供します。
 */

const AuthorizationService = (function() {

  /**
   * 現在の操作ユーザーが、対象のユーザーデータにアクセスする権限を持っているか検証します。
   * @param {string} targetUserId - 操作対象のユーザーID。
   * @returns {boolean} アクセスが許可される場合は true、それ以外は false。
   * @throws {Error} 権限がない場合にエラーをスローします。
   */
  function verifyUserAccess(targetUserId) {
    const activeUserEmail = Session.getActiveUser().getEmail();

    // 1. システム管理者（デプロイユーザー）は常にアクセス可能
    if (isDeployUser()) {
      console.log(`[Auth] Access granted for ADMIN (${activeUserEmail}) to target: ${targetUserId}`);
      return true;
    }

    // 2. 本人による操作か確認
    const targetUser = findUserById(targetUserId); // データベースからユーザー情報を取得
    if (targetUser && targetUser.adminEmail === activeUserEmail) {
      console.log(`[Auth] Access granted for OWNER (${activeUserEmail}) to target: ${targetUserId}`);
      return true;
    }

    // 3. 上記以外は権限なし
    console.warn(`[Auth] Access DENIED for user (${activeUserEmail}) to target: ${targetUserId}`);
    throw new Error(`権限エラー: ${activeUserEmail} はユーザーID ${targetUserId} のデータにアクセスする権限がありません。`);
  }

  return {
    verifyUserAccess: verifyUserAccess,
  };
})();

