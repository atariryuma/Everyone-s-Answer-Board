/**
 * @fileoverview SharingHelper - 同一ドメイン共有 + サービスアカウント editor 追加。
 *   新規 SS 作成直後の共有デフォルト一括適用 (`applySpreadsheetSharingDefaults`)。
 */

/* global getCachedProperty */

/**
 * スプレッドシートを同一ドメイン内で編集可能に設定（サービスアカウント不使用経路用）。
 * @param {string} spreadsheetId
 * @param {string} ownerEmail
 */
function setupDomainWideSharing(spreadsheetId, ownerEmail) {
  const [, domain] = (ownerEmail || '').split('@');
  if (!domain) {
    throw new Error('Invalid email format');
  }

  try {
    const file = DriveApp.getFileById(spreadsheetId);

    const sharingAccess = file.getSharingAccess();
    const sharingPermission = file.getSharingPermission();

    if (sharingAccess === DriveApp.Access.DOMAIN || sharingAccess === DriveApp.Access.DOMAIN_WITH_LINK) {
      if (sharingPermission === DriveApp.Permission.EDIT) {
        return { success: true, message: 'Already configured' };
      }
    }

    file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);

    return { success: true, message: 'Domain-wide sharing configured' };

  } catch (driveAppError) {
    // Why (fallback): Google Workspace のドメインポリシーで DOMAIN_WITH_LINK + EDIT が
    //   禁じられているテナント（教育機関で多い）では DriveApp.setSharing が落ちる。
    //   Drive REST API では type=domain + role=reader の細かい制御ができ、多くの場合
    //   こちらは通る。fallback として REST API を試す（READER 権限のみ）。
    //
    //   ログ調査 (2026-05-14) で customizeForm 経由の新規 SS が student 環境で
    //   403 を返す事象が複数回検出されたため、自動 fallback を組み込み。
    console.warn('setupDomainWideSharing: DriveApp.setSharing failed, trying Drive REST API fallback:', driveAppError.message);
    try {
      const token = ScriptApp.getOAuthToken();
      const url = 'https://www.googleapis.com/drive/v3/files/' +
                  encodeURIComponent(spreadsheetId) + '/permissions?supportsAllDrives=true';
      const resp = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json',
        muteHttpExceptions: true,
        headers: { Authorization: 'Bearer ' + token },
        payload: JSON.stringify({
          type: 'domain',
          role: 'reader',
          domain,
          allowFileDiscovery: false
        })
      });
      const code = resp.getResponseCode();
      if (code >= 400 && code !== 409) {  // 409 = already exists (idempotent)
        const body = resp.getContentText().substring(0, 300);
        throw new Error(`Drive REST API ${code}: ${body}`);
      }
      return { success: true, message: 'Domain-wide sharing configured via REST API fallback', fallback: true };
    } catch (restError) {
      console.error('setupDomainWideSharing: both DriveApp and REST API failed', {
        driveApp: driveAppError.message,
        rest: restError.message
      });
      // 元エラーを優先 (DriveApp 側のメッセージのほうが分かりやすい)
      throw driveAppError;
    }
  }
}

/**
 * SERVICE_ACCOUNT_CREDS を parse して service account 情報を取得する shared helper。
 *
 * Why: JSON.parse(credentials) は DatabaseCore / SharingHelper / ConfigService の 4 ヶ所で
 *   重複していた DRY 違反。1 ヶ所で型・必須フィールド検証して errors を一貫させる。
 *   失敗時は throw せず {success:false, message} を返す（呼出元が分岐しやすい）。
 *
 * @returns {{success:boolean, creds?:Object, saEmail?:string, message?:string}}
 *   - success: true 時に creds (parsed JSON) と saEmail (client_email) が利用可能
 *   - success: false 時は message に原因
 */
function parseServiceAccountCreds() {
  const credsJson = (typeof getCachedProperty === 'function')
    ? getCachedProperty('SERVICE_ACCOUNT_CREDS')
    : null;
  if (!credsJson) {
    return { success: false, message: 'SERVICE_ACCOUNT_CREDS not configured' };
  }
  let creds;
  try {
    creds = JSON.parse(credsJson);
  } catch (parseError) {
    return { success: false, message: 'Invalid SERVICE_ACCOUNT_CREDS JSON: ' + parseError.message };
  }
  if (!creds || typeof creds !== 'object') {
    return { success: false, message: 'SERVICE_ACCOUNT_CREDS is not an object' };
  }
  const saEmail = typeof creds.client_email === 'string' ? creds.client_email : '';
  if (!saEmail) {
    return { success: false, message: 'client_email not found in SERVICE_ACCOUNT_CREDS' };
  }
  return { success: true, creds, saEmail };
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
    const parsed = parseServiceAccountCreds();
    if (!parsed.success) {
      return { success: false, added: false, message: parsed.message };
    }

    // addEditor は冪等（既存 editor に再追加しても no-op）。getEditors の pre-check は
    // 余分な RPC で getEditors 自体が ScriptError を投げるケースもあるため避ける。
    SpreadsheetApp.openById(spreadsheetId).addEditor(parsed.saEmail);
    return { success: true, added: true, saEmail: parsed.saEmail, message: 'Service account added as editor' };
  } catch (error) {
    logError_('addServiceAccountAsEditor', error);
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

