/**
 * @fileoverview ユーザーアクセス支援関数
 * 各ユーザーが自分専用のパネルにアクセスできるようにする
 */

/**
 * 現在のユーザー専用の管理パネルURLを取得
 * HTMLから呼び出される
 * @returns {string} 現在ユーザーの管理パネルURL
 */
function getCurrentUserAdminUrl() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('ユーザーメールアドレスを取得できませんでした');
    }
    
    const userInfo = findUserByEmail(userEmail);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません: ' + userEmail);
    }
    
    const appUrls = generateAppUrls(userInfo.userId);
    return appUrls.adminUrl;
    
  } catch (error) {
    console.error('getCurrentUserAdminUrl エラー:', error.message);
    // フォールバック：基本URL + 現在ユーザーのID
    const webAppUrl = ScriptApp.getService().getUrl();
    return webAppUrl + '?mode=admin';
  }
}

/**
 * 現在のユーザー専用の回答ボードURLを取得
 * HTMLから呼び出される
 * @returns {string} 現在ユーザーの回答ボードURL
 */
function getCurrentUserViewUrl() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      throw new Error('ユーザーメールアドレスを取得できませんでした');
    }
    
    const userInfo = findUserByEmail(userEmail);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません: ' + userEmail);
    }
    
    const appUrls = generateAppUrls(userInfo.userId);
    return appUrls.viewUrl;
    
  } catch (error) {
    console.error('getCurrentUserViewUrl エラー:', error.message);
    // フォールバック：基本URL + view mode
    const webAppUrl = ScriptApp.getService().getUrl();
    return webAppUrl + '?mode=view';
  }
}

/**
 * 現在のユーザー情報を取得（HTMLから安全に呼び出し可能）
 * @returns {object} 現在ユーザーの基本情報
 */
function getCurrentUserInfo() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return {
        authenticated: false,
        error: 'ユーザー認証が取得できませんでした'
      };
    }
    
    const userInfo = findUserByEmail(userEmail);
    if (!userInfo) {
      return {
        authenticated: true,
        email: userEmail,
        registered: false,
        message: 'ユーザー登録が必要です'
      };
    }
    
    const appUrls = generateAppUrls(userInfo.userId);
    
    return {
      authenticated: true,
      email: userEmail,
      registered: true,
      userId: userInfo.userId,
      adminUrl: appUrls.adminUrl,
      viewUrl: appUrls.viewUrl,
      isActive: userInfo.isActive
    };
    
  } catch (error) {
    console.error('getCurrentUserInfo エラー:', error.message);
    return {
      authenticated: false,
      error: error.message
    };
  }
}

/**
 * 別のユーザーのボードを表示する際の権限チェック
 * @param {string} targetUserId - 表示したいユーザーのID
 * @returns {object} アクセス権限情報
 */
function checkBoardAccess(targetUserId) {
  try {
    const currentUserEmail = Session.getActiveUser().getEmail();
    if (!currentUserEmail) {
      return {
        allowed: false,
        reason: '認証されていません'
      };
    }
    
    const currentUserInfo = findUserByEmail(currentUserEmail);
    const targetUserInfo = findUserById(targetUserId);
    
    if (!currentUserInfo || !targetUserInfo) {
      return {
        allowed: false,
        reason: 'ユーザー情報が見つかりません'
      };
    }
    
    // 自分自身のボードは常にアクセス可能
    if (currentUserInfo.userId === targetUserId) {
      return {
        allowed: true,
        reason: '所有者アクセス'
      };
    }
    
    // 同一ドメインのユーザーは閲覧可能（教育機関での共有）
    const currentDomain = getEmailDomain(currentUserEmail);
    const targetDomain = getEmailDomain(targetUserInfo.adminEmail);
    
    if (currentDomain === targetDomain && currentDomain !== '') {
      return {
        allowed: true,
        reason: '同一ドメインアクセス',
        accessLevel: 'view' // 閲覧のみ
      };
    }
    
    return {
      allowed: false,
      reason: 'アクセス権限がありません'
    };
    
  } catch (error) {
    console.error('checkBoardAccess エラー:', error.message);
    return {
      allowed: false,
      reason: 'アクセス権限チェックでエラーが発生しました'
    };
  }
}

/**
 * メールアドレスからドメインを抽出
 * @param {string} email - メールアドレス
 * @returns {string} ドメイン名
 */
function getEmailDomain(email) {
  if (!email || typeof email !== 'string') {
    return '';
  }
  const parts = email.split('@');
  return parts.length > 1 ? parts[1].toLowerCase() : '';
}