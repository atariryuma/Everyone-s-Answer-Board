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
  console.log('🔄 forceLogoutAndRedirectToLogin - 関数開始');
  console.log('🔍 Function called at:', new Date().toISOString());
  console.log('🔍 Available functions check:');
  console.log('  - getWebAppUrlCached:', typeof getWebAppUrlCached);
  console.log('  - sanitizeRedirectUrl:', typeof sanitizeRedirectUrl);
  console.log('  - HtmlService:', typeof HtmlService);
  
  try {
    console.log('✅ forceLogoutAndRedirectToLogin - try block内に入りました');
    
    // Step 1: セッションクリア処理（エラーハンドリング強化）
    try {
      console.log('🧹 キャッシュクリア開始...');
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
      console.warn('⚠️ エラースタック:', cacheError.stack);
      // キャッシュクリアエラーは致命的ではないので継続
    }
    
    // Step 2: ログインページURLの生成と適切なサニタイズ
    let loginUrl;
    try {
      console.log('🔗 URL生成開始...');
      
      // getWebAppUrlCached関数の存在確認
      if (typeof getWebAppUrlCached !== 'function') {
        throw new Error('getWebAppUrlCached function not found');
      }
      
      const rawUrl = getWebAppUrlCached() + '?mode=login';
      console.log('📝 Raw URL generated:', rawUrl);
      
      // sanitizeRedirectUrl関数の存在確認
      if (typeof sanitizeRedirectUrl !== 'function') {
        throw new Error('sanitizeRedirectUrl function not found');
      }
      
      loginUrl = sanitizeRedirectUrl(rawUrl);
      console.log('✅ ログインURL生成・サニタイズ成功:', loginUrl);
      
    } catch (urlError) {
      console.warn('⚠️ WebAppURL取得失敗、フォールバック使用:', urlError.message);
      console.warn('⚠️ URLエラースタック:', urlError.stack);
      
      const fallbackUrl = ScriptApp.getService().getUrl() + '?mode=login';
      console.log('📝 Fallback URL:', fallbackUrl);
      
      try {
        loginUrl = sanitizeRedirectUrl(fallbackUrl);
      } catch (sanitizeError) {
        console.error('❌ Fallback URL sanitization failed:', sanitizeError.message);
        loginUrl = fallbackUrl; // 最終フォールバック
      }
    }
    
    console.log('🎯 Final login URL:', loginUrl);
    
    // Step 3: JavaScript文字列エスケープユーティリティ
    const escapeJavaScript = (str) => {
      if (!str) return '';
      return String(str)
        .replace(/\\/g, '\\\\')
        .replace(/'/g, "\\'")
        .replace(/"/g, '\\"')
        .replace(/\n/g, '\\n')
        .replace(/\r/g, '\\r')
        .replace(/\t/g, '\\t');
    };
    
    // Step 4: 安全なHTML生成（エスケープ済みURL使用）
    const safeLoginUrl = escapeJavaScript(loginUrl);
    console.log('🔒 Escaped login URL:', safeLoginUrl);
    
    const redirectScript = `
      <script>
        console.log('🚀 サーバーサイドリダイレクト実行:', '${safeLoginUrl}');
        
        // Google Apps Script環境に最適化されたリダイレクト
        try {
          // 最も確実な方法：window.top.location.href
          window.top.location.href = '${safeLoginUrl}';
        } catch (topError) {
          console.warn('Top frame遷移失敗:', topError);
          try {
            // フォールバック：現在のウィンドウでリダイレクト
            window.location.href = '${safeLoginUrl}';
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
    
    console.log('📄 Generated HTML script length:', redirectScript.length);
    console.log('📄 Generated HTML preview (first 200 chars):', redirectScript.substring(0, 200));
    
    // HtmlServiceの存在確認
    if (typeof HtmlService === 'undefined') {
      throw new Error('HtmlService is not available');
    }
    
    const htmlOutput = HtmlService.createHtmlOutput(redirectScript);
    console.log('✅ HtmlService.createHtmlOutput 成功');
    
    // HtmlOutputの内容確認
    if (!htmlOutput) {
      throw new Error('HtmlOutput is null or undefined');
    }
    
    // XFrameOptionsMode を安全に設定（iframe内での動作を許可）
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        console.log('✅ XFrameOptionsMode.ALLOWALL設定完了');
      }
    } catch (frameError) {
      console.warn('XFrameOptionsMode設定失敗:', frameError.message);
    }
    
    // 最終検証: HtmlOutputの内容を確認
    try {
      const outputContent = htmlOutput.getContent();
      console.log('📋 HtmlOutput content length:', outputContent ? outputContent.length : 'null/undefined');
      console.log('📋 HtmlOutput content preview:', outputContent ? outputContent.substring(0, 100) : 'NO CONTENT');
    } catch (contentError) {
      console.warn('⚠️ Cannot access HtmlOutput content:', contentError.message);
    }
    
    console.log('✅ サーバーサイドリダイレクトHTML生成完了 - 正常終了');
    return htmlOutput;
    
  } catch (error) {
    console.error('❌ サーバーサイドログアウト処理でエラー:', error.message);
    console.error('❌ エラースタック:', error.stack);
    console.error('❌ エラーの型:', typeof error);
    console.error('❌ エラーオブジェクト:', error);
    
    // Step 5: エラー時のフォールバックHTML（安全なエスケープ）
    const safeErrorMessage = String(error.message || 'Unknown error')
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/"/g, '\\"');
    
    const fallbackScript = `
      <script>
        console.error('サーバーサイドログアウトエラー: ${safeErrorMessage}');
        alert('ログアウト処理中にエラーが発生しました。\\n\\n詳細: ${safeErrorMessage}\\n\\nページを再読み込みします。');
        window.location.reload();
      </script>
      <p>エラーが発生しました: ${safeErrorMessage}</p>
      <p>ページを再読み込みしています...</p>
    `;
    
    console.log('📄 Fallback HTML generated');
    
    try {
      const fallbackOutput = HtmlService.createHtmlOutput(fallbackScript);
      console.log('✅ Fallback HtmlOutput created successfully');
      return fallbackOutput;
    } catch (fallbackError) {
      console.error('❌ Fallback HTML creation failed:', fallbackError.message);
      // 最終手段として最小限のHTML
      return HtmlService.createHtmlOutput('<script>window.location.reload();</script>');
    }
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





