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
 * 強制ログアウトしてログインページにサーバーサイドリダイレクト
 * Google Apps ScriptのSandbox制限を完全に回避する最適解
 * @returns {HtmlOutput} サーバーサイドリダイレクトHTML
 */
function forceLogoutAndRedirectToLogin() {
  try {
    console.log('🔄 サーバーサイド強制ログアウト開始...');
    
    // Step 1: セッションクリア処理（エラーハンドリング強化）
    try {
      const userCache = CacheService.getUserCache();
      if (userCache) {
        userCache.removeAll([]);
        console.log('✅ ユーザーキャッシュクリア完了');
      }
      
      const scriptCache = CacheService.getScriptCache();
      if (scriptCache) {
        scriptCache.removeAll([]);
        console.log('✅ スクリプトキャッシュクリア完了');
      }

      const props = PropertiesService.getUserProperties();
      props.deleteAllProperties();
      console.log('✅ ユーザープロパティクリア完了');
      
    } catch (cacheError) {
      console.warn('⚠️ キャッシュクリア中に一部エラー:', cacheError.message);
      // キャッシュクリアエラーは致命的ではないので継続
    }
    
    // Step 2: ログインページURLの生成（フォールバック付き）
    let loginUrl;
    try {
      loginUrl = getWebAppUrlCached() + '?mode=login';
      console.log('✅ ログインURL生成成功:', loginUrl);
    } catch (urlError) {
      console.warn('⚠️ WebAppURL取得失敗、フォールバック使用:', urlError.message);
      loginUrl = ScriptApp.getService().getUrl() + '?mode=login';
    }
    
    // Step 3: サーバーサイドリダイレクトHTMLの生成
    const redirectScript = `
      <script>
        console.log('🚀 サーバーサイドリダイレクト実行:', '${loginUrl}');
        
        // Google Apps Script環境に最適化されたリダイレクト
        try {
          // 最も確実な方法：window.top.location.href
          window.top.location.href = '${loginUrl}';
        } catch (topError) {
          console.warn('Top frame遷移失敗:', topError);
          try {
            // フォールバック：現在のウィンドウでリダイレクト
            window.location.href = '${loginUrl}';
          } catch (currentError) {
            console.error('リダイレクト完全失敗:', currentError);
            // 最終手段：ページリロード
            window.location.reload();
          }
        }
      </script>
      <noscript>
        <meta http-equiv="refresh" content="0; url=${loginUrl}">
        <p>リダイレクト中... <a href="${loginUrl}">こちらをクリック</a></p>
      </noscript>
    `;
    
    const htmlOutput = HtmlService.createHtmlOutput(redirectScript);
    
    // XFrameOptionsMode を安全に設定（iframe内での動作を許可）
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
      }
    } catch (frameError) {
      console.warn('XFrameOptionsMode設定失敗:', frameError.message);
    }
    
    console.log('✅ サーバーサイドリダイレクトHTML生成完了');
    return htmlOutput;
    
  } catch (error) {
    console.error('❌ サーバーサイドログアウト処理でエラー:', error.message);
    
    // エラー時のフォールバックHTML
    const fallbackScript = `
      <script>
        console.error('サーバーサイドログアウトエラー: ${error.message}');
        alert('ログアウト処理中にエラーが発生しました。ページを再読み込みします。');
        window.location.reload();
      </script>
      <p>エラーが発生しました。ページを再読み込みしています...</p>
    `;
    
    return HtmlService.createHtmlOutput(fallbackScript);
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





