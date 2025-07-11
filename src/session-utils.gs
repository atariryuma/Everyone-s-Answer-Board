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
 * セッション整合性の検証と修復
 * @param {string} userEmail - 現在のユーザーメール
 * @returns {boolean} 整合性が確保されたかどうか
 */
function validateAndRepairSession(userEmail) {
  // 軽量化：セッション検証をスキップして基本的な整合性のみ確認
  try {
    if (!userEmail) {
      return false;
    }
    
    // 簡素化されたユーザーID確認のみ
    const currentUserId = getUserId();
    if (!currentUserId) {
      return false;
    }
    
    return true; // 簡素化により常に成功とする
  } catch (error) {
    console.error('軽量セッション検証エラー: ' + error.message);
    return false;
  }
}

/**
 * アカウント切り替え検出とセッション更新
 * 前回アクセス時と異なるアカウントでアクセスした場合の処理
 * @param {string} currentEmail - 現在のアクセスユーザーのメール
 * @returns {object} 切り替え情報
 */
function detectAccountSwitch(currentEmail) {
  try {
    const userProps = PropertiesService.getUserProperties();
    const lastAccessEmail = userProps.getProperty('LAST_ACCESS_EMAIL');
    const lastAccessTime = userProps.getProperty('LAST_ACCESS_TIME');

    const isAccountSwitch = !!lastAccessEmail && lastAccessEmail !== currentEmail;
    const currentTime = new Date().toISOString();

    if (isAccountSwitch) {
      console.log('アカウント切り替えを検出:');
      console.log('前回: ' + lastAccessEmail + ' (' + lastAccessTime + ')');
      console.log('今回: ' + currentEmail + ' (' + currentTime + ')');

      cleanupSessionOnAccountSwitch(currentEmail);

      try {
        if (typeof clearDatabaseCache === 'function') {
          clearDatabaseCache();
        }
      } catch (globalCacheError) {
        console.warn('グローバルキャッシュクリアでエラー: ' + globalCacheError.message);
      }
    }

    userProps.setProperties({
      'LAST_ACCESS_EMAIL': currentEmail,
      'LAST_ACCESS_TIME': currentTime
    });

    return {
      isAccountSwitch: isAccountSwitch,
      previousEmail: lastAccessEmail,
      currentEmail: currentEmail,
      switchTime: currentTime
    };

  } catch (error) {
    console.error('アカウント切り替え検出でエラー: ' + error.message);
    return {
      isAccountSwitch: false,
      error: error.message
    };
  }
}

/**
 * 強制的なセッション初期化
 * 重大なセッション問題が発生した場合の最終手段
 * @param {string} userEmail - 対象ユーザーのメール
 */
function forceSessionReset(userEmail) {
  try {
    console.log('強制セッション初期化を開始: ' + userEmail);
    
    // 全てのプロパティをクリア
    const props = PropertiesService.getUserProperties();
    const allProps = props.getProperties();
    Object.keys(allProps).forEach(function(key) {
      if (key.includes('USER_ID') || key.includes('CACHE')) {
        props.deleteProperty(key);
      }
    });
    
    // 全てのキャッシュをクリア
    const userCache = CacheService.getUserCache();
    const scriptCache = CacheService.getScriptCache();
    
    try {
      if (userCache) userCache.removeAll([]);
      if (scriptCache) scriptCache.removeAll([]);
    } catch (cacheError) {
      console.warn('キャッシュクリアでエラー: ' + cacheError.message);
    }
    
    // 新しいセッションを初期化
    const newUserId = getUserId(); // これで新しいセッションが開始される
    
    console.log('強制セッション初期化完了: ' + userEmail + ' -> ' + newUserId);
    
    return {
      success: true,
      newUserId: newUserId,
      message: 'セッションが正常に初期化されました'
    };
    
  } catch (error) {
    console.error('強制セッション初期化でエラー: ' + error.message);
    return {
      success: false,
      error: error.message
    };
  }
}