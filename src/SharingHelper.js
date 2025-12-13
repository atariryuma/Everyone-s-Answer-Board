/**
 * @fileoverview SharingHelper - スプレッドシート共有設定ヘルパー
 *
 * 🎯 責任範囲:
 * - 同一ドメイン内での共有設定
 * - サービスアカウント不使用のための共有管理
 */

/**
 * スプレッドシートを同一ドメイン内で編集可能に設定
 * ✅ CRITICAL: サービスアカウント不使用のための共有設定
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} ownerEmail - オーナーのメールアドレス
 * @returns {Object} 処理結果
 */
function setupDomainWideSharing(spreadsheetId, ownerEmail) {
  try {
    const file = DriveApp.getFileById(spreadsheetId);

    const [, domain] = ownerEmail.split('@');
    if (!domain) {
      throw new Error('Invalid email format');
    }

    const sharingAccess = file.getSharingAccess();
    const sharingPermission = file.getSharingPermission();

    if (sharingAccess === DriveApp.Access.DOMAIN || sharingAccess === DriveApp.Access.DOMAIN_WITH_LINK) {
      if (sharingPermission === DriveApp.Permission.EDIT) {
        return { success: true, message: 'Already configured' };
      }
    }

    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

    return { success: true, message: 'Domain-wide sharing configured' };

  } catch (error) {
    console.error('setupDomainWideSharing error:', error.message);
    throw error;
  }
}
