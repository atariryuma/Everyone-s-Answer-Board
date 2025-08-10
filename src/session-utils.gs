/**
 * @fileoverview セッション管理ユーティリティ
 * アカウント間のセッション分離と整合性管理を提供
 */

// 回復力のある実行機構を使用
function getResilientPropertiesService() {
  return resilientExecutor.execute(
    () => PropertiesService.getUserProperties(),
    { name: 'PropertiesService.getUserProperties', idempotent: true }
  );
}

function getResilientCacheService() {
  return resilientExecutor.execute(
    () => CacheService.getUserCache(),
    { name: 'CacheService.getUserCache', idempotent: true }
  );
}

function getResilientScriptCache() {
  return resilientExecutor.execute(
    () => CacheService.getScriptCache(),
    { name: 'CacheService.getScriptCache', idempotent: true }
  );
}

/**
 * アカウント切り替え時のセッションクリーンアップ
 * 異なるアカウントでログインした際に前のセッション情報をクリア
 * @param {string} currentEmail - 現在のユーザーメール
 */
function cleanupSessionOnAccountSwitch(currentEmail) {
  try {
    debugLog('セッションクリーンアップを開始: ' + currentEmail);

    const props = getResilientPropertiesService();
    const userCache = getResilientCacheService();

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
        debugLog('削除された古いユーザーキャッシュ: ' + key);
      }
    });

    // ユーザーキャッシュを全面クリア（API修正版）
    if (userCache) {
      try {
        if (typeof userCache.removeAll === 'function') {
          userCache.removeAll();
          debugLog('ユーザーキャッシュ全クリア完了');
        } else {
          // removeAll未提供の環境。キー列挙ができないためスキップ
          warnLog('UserCache.removeAll は未サポート。スキップします');
        }
      } catch (cacheError) {
        warnLog('ユーザーキャッシュクリア中のエラー: ' + cacheError.message);
      }
    }

    // スクリプトキャッシュの関連項目もクリア
    const scriptCache = getResilientScriptCache();
    if (scriptCache) {
      try {
        // 現在のユーザー以外のキャッシュをクリア
        ['config_v3_', 'user_', 'email_'].forEach(function(prefix) {
          scriptCache.remove(prefix + currentEmail);
        });
      } catch (scriptCacheError) {
        warnLog('スクリプトキャッシュクリア中のエラー: ' + scriptCacheError.message);
      }
    }

    debugLog('セッションクリーンアップ完了: ' + currentEmail);

  } catch (error) {
    errorLog('セッションクリーンアップでエラー: ' + error.message);
    // エラーが発生してもアプリケーションを停止させない
  }
}

/**
 * ユーザー認証をリセットし、ログインページURLを返す
 * @returns {string} ログインページのURL
 */
function resetUserAuthentication() {
  try {
    debugLog('ユーザー認証をリセット中...');
    const userCache = getResilientCacheService();
    if (userCache) {
      userCache.removeAll(); // GAS API仕様に合わせて引数なし
      debugLog('ユーザーキャッシュをクリアしました。');
    }

    const scriptCache = getResilientScriptCache();
    if (scriptCache) {
      if (typeof scriptCache.removeAll === 'function') {
        scriptCache.removeAll();
        debugLog('スクリプトキャッシュをクリアしました。');
      } else {
        // 既知のキーのみ削除（全消去APIは未提供）
        try {
          const email = getCurrentUserEmail();
          ['config_v3_', 'user_', 'email_'].forEach(function(prefix) {
            scriptCache.remove(prefix + email);
          });
          debugLog('スクリプトキャッシュ: 既知のキーを削除しました');
        } catch (e) {
          warnLog('スクリプトキャッシュのキー削除に失敗: ' + e.message);
        }
      }
    }

    // PropertiesServiceもクリアする（LAST_ACCESS_EMAILなど）
    const props = getResilientPropertiesService();
    props.deleteAllProperties();
    debugLog('ユーザープロパティをクリアしました。');

    // ログインページのURLを返す
    const loginPageUrl = ScriptApp.getService().getUrl();
    debugLog('リセット完了。ログインページURL:', loginPageUrl);
    return loginPageUrl;
  } catch (error) {
    errorLog('ユーザー認証リセット中にエラーが発生しました:', error.message);
    throw new Error('認証リセットに失敗しました: ' + error.message);
  }
}

/**
 * 強制ログアウトしてログインページにサーバーサイドリダイレクト
 * Google Apps ScriptのSandbox制限を完全に回避する最適解
 * @returns {HtmlOutput} サーバーサイドリダイレクトHTML
 */
function forceLogoutAndRedirectToLogin() {
  debugLog('🔄 forceLogoutAndRedirectToLogin - 関数開始');
  debugLog('🔍 Function called at:', new Date().toISOString());
  debugLog('🔍 Available functions check:');
  debugLog('  - getWebAppUrl:', typeof getWebAppUrl);
  debugLog('  - sanitizeRedirectUrl:', typeof sanitizeRedirectUrl);
  debugLog('  - HtmlService:', typeof HtmlService);

  try {
    debugLog('✅ forceLogoutAndRedirectToLogin - try block内に入りました');

    // Step 1: セッションクリア処理（エラーハンドリング強化）
    try {
      debugLog('🧹 キャッシュクリア開始...');
      const userCache = getResilientCacheService();
      if (userCache) {
        if (typeof userCache.removeAll === 'function') {
          userCache.removeAll();
          debugLog('✅ ユーザーキャッシュクリア完了');
        } else {
          warnLog('UserCache.removeAll は未サポート。スキップします');
        }
      }

      const scriptCache = getResilientScriptCache();
      if (scriptCache) {
        if (typeof scriptCache.removeAll === 'function') {
          scriptCache.removeAll();
          debugLog('✅ スクリプトキャッシュクリア完了');
        } else {
          try {
            const email = getCurrentUserEmail();
            ['config_v3_', 'user_', 'email_'].forEach(function(prefix) {
              scriptCache.remove(prefix + email);
            });
            debugLog('スクリプトキャッシュ: 既知のキーを削除しました');
          } catch (e) {
            warnLog('スクリプトキャッシュのキー削除に失敗: ' + e.message);
          }
        }
      }

      const props = getResilientPropertiesService();
      props.deleteAllProperties();
      debugLog('✅ ユーザープロパティクリア完了');

    } catch (cacheError) {
      warnLog('⚠️ キャッシュクリア中に一部エラー:', cacheError.message);
      warnLog('⚠️ エラースタック:', cacheError.stack);
      // キャッシュクリアエラーは致命的ではないので継続
    }

    // Step 2: ログインページURLの生成と適切なサニタイズ
    let loginUrl;
    try {
      debugLog('🔗 URL生成開始...');

      // getWebAppUrl関数の存在確認
      if (typeof getWebAppUrl !== 'function') {
        throw new Error('getWebAppUrl function not found');
      }

      const rawUrl = getWebAppUrl() + '?mode=login';
      debugLog('📝 Raw URL generated:', rawUrl);

      // sanitizeRedirectUrl関数の存在確認
      if (typeof sanitizeRedirectUrl !== 'function') {
        throw new Error('sanitizeRedirectUrl function not found');
      }

      loginUrl = sanitizeRedirectUrl(rawUrl);
      debugLog('✅ ログインURL生成・サニタイズ成功:', loginUrl);

    } catch (urlError) {
      warnLog('⚠️ WebAppURL取得失敗、フォールバック使用:', urlError.message);
      warnLog('⚠️ URLエラースタック:', urlError.stack);

      const fallbackUrl = ScriptApp.getService().getUrl() + '?mode=login';
      debugLog('📝 Fallback URL:', fallbackUrl);

      try {
        loginUrl = sanitizeRedirectUrl(fallbackUrl);
      } catch (sanitizeError) {
        errorLog('❌ Fallback URL sanitization failed:', sanitizeError.message);
        loginUrl = fallbackUrl; // 最終フォールバック
      }
    }

    debugLog('🎯 Final login URL:', loginUrl);

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
    debugLog('🔒 Escaped login URL:', safeLoginUrl);

    const redirectScript = `
      <script>
        debugLog('🚀 サーバーサイドリダイレクト実行:', '${safeLoginUrl}');

        // Google Apps Script環境に最適化されたリダイレクト
        try {
          // 最も確実な方法：window.top.location.href
          window.top.location.href = '${safeLoginUrl}';
        } catch (topError) {
          warnLog('Top frame遷移失敗:', topError);
          try {
            // フォールバック：現在のウィンドウでリダイレクト
            window.location.href = '${safeLoginUrl}';
          } catch (currentError) {
            errorLog('リダイレクト完全失敗:', currentError);
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

    debugLog('📄 Generated HTML script length:', redirectScript.length);
    debugLog('📄 Generated HTML preview (first 200 chars):', redirectScript.substring(0, 200));

    // HtmlServiceの存在確認
    if (typeof HtmlService === 'undefined') {
      throw new Error('HtmlService is not available');
    }

    const htmlOutput = HtmlService.createHtmlOutput(redirectScript);
    debugLog('✅ HtmlService.createHtmlOutput 成功');

    // HtmlOutputの内容確認
    if (!htmlOutput) {
      throw new Error('HtmlOutput is null or undefined');
    }

    // XFrameOptionsMode を安全に設定（iframe内での動作を許可）
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        debugLog('✅ XFrameOptionsMode.ALLOWALL設定完了');
      }
    } catch (frameError) {
      warnLog('XFrameOptionsMode設定失敗:', frameError.message);
    }

    // 最終検証: HtmlOutputの内容を確認
    try {
      const outputContent = htmlOutput.getContent();
      debugLog('📋 HtmlOutput content length:', outputContent ? outputContent.length : 'null/undefined');
      debugLog('📋 HtmlOutput content preview:', outputContent ? outputContent.substring(0, 100) : 'NO CONTENT');
    } catch (contentError) {
      warnLog('⚠️ Cannot access HtmlOutput content:', contentError.message);
    }

    debugLog('✅ サーバーサイドリダイレクトHTML生成完了 - 正常終了');
    return htmlOutput;

  } catch (error) {
    errorLog('❌ サーバーサイドログアウト処理でエラー:', error.message);
    errorLog('❌ エラースタック:', error.stack);
    errorLog('❌ エラーの型:', typeof error);
    errorLog('❌ エラーオブジェクト:', error);

    // Step 5: エラー時のフォールバックHTML（安全なエスケープ）
    const safeErrorMessage = String(error.message || 'Unknown error')
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/"/g, '\\"');

    const fallbackScript = `
      <script>
        errorLog('サーバーサイドログアウトエラー: ${safeErrorMessage}');
        alert('ログアウト処理中にエラーが発生しました。\\n\\n詳細: ${safeErrorMessage}\\n\\nページを再読み込みします。');
        window.location.reload();
      </script>
      <p>エラーが発生しました: ${safeErrorMessage}</p>
      <p>ページを再読み込みしています...</p>
    `;

    debugLog('📄 Fallback HTML generated');

    try {
      const fallbackOutput = HtmlService.createHtmlOutput(fallbackScript);
      debugLog('✅ Fallback HtmlOutput created successfully');
      return fallbackOutput;
    } catch (fallbackError) {
      errorLog('❌ Fallback HTML creation failed:', fallbackError.message);
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
    const props = getResilientPropertiesService();
    const lastEmailKey = 'last_active_email';
    const lastEmail = props.getProperty(lastEmailKey);

    const isAccountSwitch = !!(lastEmail && lastEmail !== currentEmail);

    if (isAccountSwitch) {
      debugLog('アカウント切り替えを検出:', lastEmail, '->', currentEmail);
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
    errorLog('アカウント切り替え検出中にエラー:', error.message);
    return {
      isAccountSwitch: false,
      previousEmail: null,
      currentEmail: currentEmail,
      error: error.message
    };
  }
}




