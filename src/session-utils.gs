/**
 * @fileoverview セッション管理ユーティリティ
 * アカウント間のセッション分離と整合性管理を提供
 */

/**
 * アカウント切り替え時のセッションクリーンアップ
 * 異なるアカウントでログインした際に前のセッション情報をクリア
 * @param {string} currentEmail - 現在のユーザーメール
 */
function cleanupSessionOnAccountSwitch(currentEmail) {
  try {
    console.log('セッションクリーンアップを開始: ' + currentEmail);
    
    const props = PropertiesService.getUserProperties();
    const userCache = CacheService.getUserCache();
    
    // 現在のユーザーのハッシュキーを生成
    const currentUserKey = 'CURRENT_USER_ID_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
      .map(function(byte) { return (byte + 256).toString(16).slice(-2); })
      .join('');
    
    // 古い形式のキャッシュを完全削除
    props.deleteProperty('CURRENT_USER_ID');
    
    // 他のユーザーIDキャッシュをクリア（現在のユーザー以外）
    const allProperties = props.getProperties();
    Object.keys(allProperties).forEach(function(key) {
      if (key.startsWith('CURRENT_USER_ID_') && key !== currentUserKey) {
        props.deleteProperty(key);
        console.log('削除された古いユーザーキャッシュ: ' + key);
      }
    });
    
    // ユーザーキャッシュを全面クリア
    if (userCache) {
      try {
        userCache.removeAll(['config_v3_', 'user_', 'email_', 'hdr_', 'data_', 'sheets_']);
      } catch (cacheError) {
        console.warn('ユーザーキャッシュクリア中のエラー: ' + cacheError.message);
      }
    }
    
    // スクリプトキャッシュの関連項目もクリア
    const scriptCache = CacheService.getScriptCache();
    if (scriptCache) {
      try {
        // 現在のユーザー以外のキャッシュをクリア
        ['config_v3_', 'user_', 'email_'].forEach(function(prefix) {
          scriptCache.remove(prefix + currentEmail);
        });
      } catch (scriptCacheError) {
        console.warn('スクリプトキャッシュクリア中のエラー: ' + scriptCacheError.message);
      }
    }
    
    console.log('セッションクリーンアップ完了: ' + currentEmail);
    
  } catch (error) {
    console.error('セッションクリーンアップでエラー: ' + error.message);
    // エラーが発生してもアプリケーションを停止させない
  }
}

/**
 * ユーザー認証をリセットし、ログインページURLを返す
 * @returns {string} ログインページのURL
 */
function resetUserAuthentication() {
  try {
    console.log('ユーザー認証をリセット中...');
    const userCache = CacheService.getUserCache();
    if (userCache) {
      userCache.removeAll([]); // ユーザーキャッシュを全てクリア
      console.log('ユーザーキャッシュをクリアしました。');
    }
    
    const scriptCache = CacheService.getScriptCache();
    if (scriptCache) {
      scriptCache.removeAll([]); // スクリプトキャッシュを全てクリア
      console.log('スクリプトキャッシュをクリアしました。');
    }

    // PropertiesServiceもクリアする（LAST_ACCESS_EMAILなど）
    const props = PropertiesService.getUserProperties();
    props.deleteAllProperties();
    console.log('ユーザープロパティをクリアしました。');

    // ログインページのURLを返す
    const loginPageUrl = ScriptApp.getService().getUrl();
    console.log('リセット完了。ログインページURL:', loginPageUrl);
    return loginPageUrl;
  } catch (error) {
    console.error('ユーザー認証リセット中にエラーが発生しました:', error.message);
    throw new Error('認証リセットに失敗しました: ' + error.message);
  }
}

/**
 * アカウント切り替えを検出
 * @param {string} currentEmail - 現在のユーザーメール
 * @returns {object} アカウント切り替え検出結果
 */
function detectAccountSwitch(currentEmail) {
  try {
    const props = PropertiesService.getUserProperties();
    const lastEmailKey = 'last_active_email';
    const lastEmail = props.getProperty(lastEmailKey);
    
    const isAccountSwitch = !!(lastEmail && lastEmail !== currentEmail);
    
    if (isAccountSwitch) {
      console.log('アカウント切り替えを検出:', lastEmail, '->', currentEmail);
      cleanupSessionOnAccountSwitch(currentEmail);
      clearDatabaseCache();
    }
    
    // 現在のメールアドレスを記録
    props.setProperty(lastEmailKey, currentEmail);
    
    return {
      isAccountSwitch: isAccountSwitch,
      previousEmail: lastEmail,
      currentEmail: currentEmail
    };
  } catch (error) {
    console.error('アカウント切り替え検出中にエラー:', error.message);
    return {
      isAccountSwitch: false,
      previousEmail: null,
      currentEmail: currentEmail,
      error: error.message
    };
  }
}





