/**
 * @fileoverview SharingHelper - スプレッドシート共有設定ヘルパー
 *
 * 🎯 責任範囲:
 * - 同一ドメイン内での共有設定
 * - サービスアカウント editor 追加（DatabaseCore Sheets API JWT 経路用）
 * - 新規 SS 作成直後の共有デフォルト一括適用
 */

/* global SpreadsheetApp, DriveApp, getCachedProperty */

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

/**
 * SERVICE_ACCOUNT_CREDS の client_email を SS の editor として追加する。
 *
 * Why: DatabaseCore は Sheets REST API + Service Account JWT で SS にアクセスする
 *   (`SpreadsheetApp.openById()` ではない)。サービスアカウントが editor でない SS は
 *   403 を返し、`getPublishedSheetData` が
 *   「リクエストされたドキュメントにアクセスする権限がありません」で落ちる。
 *
 *   createTemplateForm / customizeForm が新規 SS を作るたびに必ず呼ぶこと。
 *   既存 SS の遡及修復は AdminApis.repairSpreadsheetSharing op から本関数を呼ぶ。
 *
 * @param {string} spreadsheetId
 * @returns {{success:boolean, added:boolean, saEmail?:string, message?:string}}
 */
function addServiceAccountAsEditor(spreadsheetId) {
  try {
    const credsJson = getCachedProperty('SERVICE_ACCOUNT_CREDS');
    if (!credsJson) {
      return { success: false, added: false, message: 'SERVICE_ACCOUNT_CREDS not configured' };
    }
    let saEmail = '';
    try {
      saEmail = (JSON.parse(credsJson) || {}).client_email || '';
    } catch (parseError) {
      return { success: false, added: false, message: 'Invalid SERVICE_ACCOUNT_CREDS JSON: ' + parseError.message };
    }
    if (!saEmail) {
      return { success: false, added: false, message: 'client_email not found in SERVICE_ACCOUNT_CREDS' };
    }

    // addEditor は冪等（既存 editor に再追加しても no-op）。getEditors の pre-check は
    // 余分な RPC で getEditors 自体が ScriptError を投げるケースもあるため避ける。
    SpreadsheetApp.openById(spreadsheetId).addEditor(saEmail);
    return { success: true, added: true, saEmail, message: 'Service account added as editor' };
  } catch (error) {
    console.error('addServiceAccountAsEditor error:', error.message);
    return { success: false, added: false, message: error.message };
  }
}

/**
 * 新規作成 SS に共有デフォルトを一括適用する。
 *   1) サービスアカウントを editor 追加（必須 — 未設定だと view が 403）
 *   2) ownerEmail のドメインへ DOMAIN_WITH_LINK + EDIT を設定（任意 — 失敗してもサーバ動作は保たれる）
 *
 * Why: createTemplateForm と customizeForm の resetDestination 経路で同じ処理が必要なため
 *   1 関数にまとめる。両ステップとも fail-soft（throw せず errors[] に詰めて返す）にして、
 *   呼び出し元のメインフローを止めない。
 *
 * @param {string} spreadsheetId
 * @param {string} ownerEmail - 通常は getCurrentEmail()（フォーム作成者）
 * @returns {{saAdded:boolean, domainShared:boolean, errors:string[]}}
 */
function applySpreadsheetSharingDefaults(spreadsheetId, ownerEmail) {
  const result = { saAdded: false, domainShared: false, errors: [] };

  try {
    const sa = addServiceAccountAsEditor(spreadsheetId);
    result.saAdded = !!(sa && sa.success);
    if (sa && !sa.success && sa.message) result.errors.push('SA: ' + sa.message);
  } catch (saError) {
    result.errors.push('SA: ' + saError.message);
  }

  if (ownerEmail && typeof ownerEmail === 'string' && ownerEmail.indexOf('@') > 0) {
    try {
      const dom = setupDomainWideSharing(spreadsheetId, ownerEmail);
      result.domainShared = !!(dom && dom.success);
    } catch (domError) {
      // setupDomainWideSharing は throw する設計。fail-soft で吸収。
      result.errors.push('Domain: ' + domError.message);
    }
  }

  return result;
}

