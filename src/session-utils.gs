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
 * Detect account switch and handle cleanup if necessary.
 * @param {string} currentEmail - Current user's email.
 * @returns {Object} Result info.
 */
function detectAccountSwitch(currentEmail) {
  try {
    var props = PropertiesService.getUserProperties();
    var lastEmail = props.getProperty('LAST_ACCESS_EMAIL');
    var isSwitch = !!(lastEmail && lastEmail !== currentEmail);
    props.setProperty('LAST_ACCESS_EMAIL', currentEmail);
    if (isSwitch) {
      cleanupSessionOnAccountSwitch(currentEmail);
      if (typeof clearDatabaseCache === 'function') {
        clearDatabaseCache(currentEmail);
      }
    }
    return { isAccountSwitch: isSwitch, previousEmail: lastEmail };
  } catch (e) {
    console.error('detectAccountSwitch error: ' + e.message);
    return { isAccountSwitch: false };
  }
}





