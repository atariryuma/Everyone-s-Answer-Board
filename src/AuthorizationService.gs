/**
 * @fileoverview 権限管理サービス
 * システム内のリソースへのアクセス権を検証する機能を提供します。
 */

var AuthorizationService = (function() {

  /**
   * メールアドレスからドメインを抽出します。
   * @param {string} email - メールアドレス
   * @returns {string} ドメイン部分
   */
  function getEmailDomain(email) {
    if (!email) return '';
    return String(email)
      .split('@')[1]
      .toLowerCase()
      .trim() || '';
  }

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

  /**
   * ボードへのアクセス権をドメイン単位で検証します。
   * @param {string} adminEmail - ボード管理者のメールアドレス
   * @returns {boolean} ドメインが一致する場合は true
   * @throws {Error} ドメインが一致しない場合にエラーをスローします
   */
  function verifyBoardAccess(adminEmail) {
    const activeUserEmail = Session.getActiveUser().getEmail();
    const activeDomain = getEmailDomain(activeUserEmail);
    const adminDomain = getEmailDomain(adminEmail);

    // ドメインの柔軟な比較: activeDomainがadminDomainで終わるかをチェック
    // これにより、サブドメインや部分的な一致を許容します。
    if (activeDomain.endsWith(adminDomain)) {
      console.log(`[Auth] Domain access granted (${activeDomain})`);
      return true;
    }

    console.warn(`[Auth] Domain access denied: ${activeDomain} vs ${adminDomain}`);
    throw new Error('権限エラー: ドメインが一致しません');
  }

  return {
    verifyUserAccess: verifyUserAccess,
    verifyBoardAccess: verifyBoardAccess,
  };
})();

