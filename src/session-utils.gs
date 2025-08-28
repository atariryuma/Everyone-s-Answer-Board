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
 * キャッシュを安全に消去するユーティリティ
 * - removeAll() がサポートされる環境では全面削除
 * - それ以外は既知のキーやプレフィックスでピンポイント削除
 * @param {Cache} cache - GAS CacheService のキャッシュオブジェクト
 * @param {Object} options
 * @param {string} [options.label] - ログ用ラベル
 * @param {string} [options.email] - キー生成に使うメール
 * @param {string[]} [options.prefixes] - 削除対象キーのプレフィックス一覧（emailと組合せ）
 * @param {string[]} [options.keys] - 直接削除するキー一覧
 */
function clearCacheSafely(cache, options) {
  try {
    if (!cache) return { success: true, method: 'none' };
    const label = (options && options.label) || 'Cache';
    const prefixes = (options && options.prefixes) || [];
    const email = (options && options.email) || '';
    const keys = (options && options.keys) || [];

    // 1) 全消去が可能ならそれを試みる
    if (typeof cache.removeAll === 'function') {
      if (keys.length > 0) {
        // removeAll(keys[]) が使える環境
        try {
          cache.removeAll(keys);
          debugLog(label + ': removeAll(keys) でキャッシュを削除しました');
          return { success: true, method: 'removeAll(keys)' };
        } catch (e) {
          warnLog(label + ': removeAll(keys) でエラー: ' + e.message);
        }
      }

      // キー未指定でも removeAll() が使える環境（非標準だが一部で提供される可能性）
      try {
        cache.removeAll();
        debugLog(label + ': removeAll() でキャッシュを全消去しました');
        return { success: true, method: 'removeAll()' };
      } catch (e) {
        // 環境により未サポート
        warnLog(label + ': removeAll() は未サポート: ' + e.message);
      }
    }

    // 2) プレフィックス + email で個別削除（Cache.remove）
    if (typeof cache.remove === 'function' && prefixes.length > 0 && email) {
      prefixes.forEach(function(prefix) {
        try {
          cache.remove(prefix + email);
        } catch (e) {
          // 続行
          warnLog(label + ': プレフィックス削除中のエラー: ' + e.message);
        }
      });
      debugLog(label + ': 既知のプレフィックスキーを削除しました');
      return { success: true, method: 'remove(prefix+email)' };
    }

    // 3) keys があれば removeAll(keys) を最終試行
    if (typeof cache.removeAll === 'function' && keys.length > 0) {
      try {
        cache.removeAll(keys);
        debugLog(label + ': removeAll(keys) でキャッシュを削除しました');
        return { success: true, method: 'removeAll(keys)' };
      } catch (e) {
        warnLog(label + ': removeAll(keys) 最終試行で失敗: ' + e.message);
      }
    }

    // 4) 何もできない環境
    warnLog(label + ': キャッシュ全面削除 API 未提供のためスキップしました');
    return { success: false, method: 'skipped' };
  } catch (error) {
    warnLog('clearCacheSafely エラー: ' + (error && error.message));
    return { success: false, method: 'error', error: error && error.message };
  }
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

    // ユーザーキャッシュを安全にクリア
    clearCacheSafely(userCache, { label: 'UserCache' });

    // スクリプトキャッシュの関連項目もクリア
    const scriptCache = getResilientScriptCache();
    clearCacheSafely(scriptCache, { label: 'ScriptCache', email: currentEmail, prefixes: ['config_v3_', 'user_', 'email_'] });

    debugLog('セッションクリーンアップ完了: ' + currentEmail);

  } catch (error) {
    console.error('[ERROR]','セッションクリーンアップでエラー: ' + error.message);
    // エラーが発生してもアプリケーションを停止させない
  }
}

/**
 * ユーザー認証をリセットし、ログインページURLを返す
 * @returns {string} ログインページのURL
 */
function resetUserAuthentication() {
  let authResetResult = {
    cacheCleared: false,
    propertiesCleared: false,
    loginUrl: null,
    errors: []
  };
  
  try {
    debugLog('🔄 ユーザー認証をリセット開始...');
    const startTime = Date.now();
    
    // Step 1: キャッシュクリア（非致命的エラー許容）
    try {
      const userCache = CacheService.getUserCache();
      if (userCache) {
        clearCacheSafely(userCache, { label: 'UserCache' });
      }

      const scriptCache = CacheService.getScriptCache();
      if (scriptCache) {
        clearCacheSafely(scriptCache, { label: 'ScriptCache', email: getCurrentUserEmail(), prefixes: ['config_v3_', 'user_', 'email_'] });
      }
      
      authResetResult.cacheCleared = true;
      debugLog('✅ キャッシュクリア完了');
    } catch (cacheError) {
      authResetResult.errors.push(`キャッシュクリアエラー: ${cacheError.message}`);
      warnLog('⚠️ キャッシュクリアでエラーが発生しましたが、処理を続行します:', cacheError.message);
    }

    // Step 2: プロパティクリア（段階的実装）
    try {
      const props = PropertiesService.getUserProperties();
      
      // まず deleteAllProperties を試行
      if (typeof props.deleteAllProperties === 'function') {
        props.deleteAllProperties();
        authResetResult.propertiesCleared = true;
        debugLog('✅ PropertiesService.deleteAllProperties() でクリア完了');
      } else {
        // フォールバック: 重要なプロパティを個別削除
        debugLog('⚠️ deleteAllProperties が利用できません。個別削除を実行します。');
        const importantKeys = [
          'LAST_ACCESS_EMAIL',
          'lastAdminUserId', 
          'CURRENT_USER_ID',
          'USER_CACHE_KEY',
          'SESSION_ID'
        ];
        
        // 既存のプロパティをすべて取得して削除
        try {
          const allProps = props.getProperties();
          const keys = Object.keys(allProps);
          
          if (keys.length > 0) {
            debugLog(`📝 ${keys.length}個のプロパティを個別削除します:`, keys);
            
            for (const key of keys) {
              try {
                props.deleteProperty(key);
              } catch (deleteError) {
                authResetResult.errors.push(`プロパティ削除エラー(${key}): ${deleteError.message}`);
              }
            }
            
            authResetResult.propertiesCleared = true;
            debugLog('✅ 個別プロパティ削除完了');
          } else {
            debugLog('ℹ️ 削除対象のプロパティはありませんでした');
            authResetResult.propertiesCleared = true;
          }
        } catch (enumError) {
          // プロパティ一覧取得に失敗した場合、重要なキーのみ削除を試行
          debugLog('⚠️ プロパティ一覧取得に失敗。重要なキーのみ削除を試行します。');
          for (const key of importantKeys) {
            try {
              props.deleteProperty(key);
            } catch (deleteError) {
              // 個別削除エラーは記録するが処理は続行
              authResetResult.errors.push(`重要プロパティ削除エラー(${key}): ${deleteError.message}`);
            }
          }
          authResetResult.propertiesCleared = true;
        }
      }
    } catch (propsError) {
      authResetResult.errors.push(`プロパティクリアエラー: ${propsError.message}`);
      console.error('[ERROR]','❌ プロパティクリアでエラーが発生しました:', propsError.message);
      
      // プロパティクリアに失敗してもログアウト処理は続行
      warnLog('⚠️ プロパティクリアに失敗しましたが、ログアウト処理を続行します');
    }

    // Step 3: ログインページURLの取得
    try {
      const loginPageUrl = ScriptApp.getService().getUrl();
      authResetResult.loginUrl = loginPageUrl;
      
      const executionTime = Date.now() - startTime;
      debugLog(`✅ 認証リセット完了 (${executionTime}ms):`, {
        cacheCleared: authResetResult.cacheCleared,
        propertiesCleared: authResetResult.propertiesCleared,
        errorsCount: authResetResult.errors.length,
        loginUrl: loginPageUrl
      });
      
      return loginPageUrl;
    } catch (urlError) {
      authResetResult.errors.push(`URL取得エラー: ${urlError.message}`);
      throw new Error(`ログインページURL取得に失敗: ${urlError.message}`);
    }
    
  } catch (error) {
    const executionTime = Date.now() - (authResetResult.startTime || Date.now());
    console.error('[ERROR]','❌ ユーザー認証リセット中に致命的エラー:', {
      error: error.message,
      executionTime: executionTime + 'ms',
      partialResults: authResetResult
    });
    
    // エラーメッセージを詳細化
    let errorMessage = `認証リセットに失敗しました: ${error.message}`;
    if (authResetResult.errors.length > 0) {
      errorMessage += ` (追加エラー: ${authResetResult.errors.join(', ')})`;
    }
    
    throw new Error(errorMessage);
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
      
      // キャッシュクリア（非致命的エラー許容）
      try {
        const userCache = CacheService.getUserCache();
        if (userCache) {
          clearCacheSafely(userCache, { label: 'UserCache' });
        }

        const scriptCache = CacheService.getScriptCache();
        if (scriptCache) {
          clearCacheSafely(scriptCache, { label: 'ScriptCache', email: getCurrentUserEmail(), prefixes: ['config_v3_', 'user_', 'email_'] });
        }
        debugLog('✅ キャッシュクリア完了');
      } catch (cacheError) {
        warnLog('⚠️ キャッシュクリアでエラーが発生しましたが、処理を続行します:', cacheError.message);
      }

      // プロパティクリア（段階的実装）
      try {
        const props = PropertiesService.getUserProperties();
        
        // まず deleteAllProperties を試行
        if (typeof props.deleteAllProperties === 'function') {
          props.deleteAllProperties();
          debugLog('✅ PropertiesService.deleteAllProperties() でクリア完了');
        } else {
          // フォールバック: 既存のプロパティをすべて取得して削除
          debugLog('⚠️ deleteAllProperties が利用できません。個別削除を実行します。');
          try {
            const allProps = props.getProperties();
            const keys = Object.keys(allProps);
            
            if (keys.length > 0) {
              debugLog(`📝 ${keys.length}個のプロパティを個別削除します:`, keys);
              
              for (const key of keys) {
                try {
                  props.deleteProperty(key);
                } catch (deleteError) {
                  warnLog(`プロパティ削除エラー(${key}): ${deleteError.message}`);
                }
              }
              
              debugLog('✅ 個別プロパティ削除完了');
            } else {
              debugLog('ℹ️ 削除対象のプロパティはありませんでした');
            }
          } catch (enumError) {
            // プロパティ一覧取得に失敗した場合、重要なキーのみ削除を試行
            debugLog('⚠️ プロパティ一覧取得に失敗。重要なキーのみ削除を試行します。');
            const importantKeys = [
              'LAST_ACCESS_EMAIL',
              'lastAdminUserId', 
              'CURRENT_USER_ID',
              'USER_CACHE_KEY',
              'SESSION_ID'
            ];
            for (const key of importantKeys) {
              try {
                props.deleteProperty(key);
              } catch (deleteError) {
                warnLog(`重要プロパティ削除エラー(${key}): ${deleteError.message}`);
              }
            }
          }
        }
        debugLog('✅ ユーザープロパティクリア完了');
      } catch (propsError) {
        warnLog('⚠️ プロパティクリアでエラーが発生しましたが、ログアウト処理を続行します:', propsError.message);
      }

    } catch (cacheError) {
      warnLog('⚠️ セッションクリア中に一部エラー:', cacheError.message);
      warnLog('⚠️ エラースタック:', cacheError.stack);
      // セッションクリアエラーは致命的ではないので継続
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
        console.error('[ERROR]','❌ Fallback URL sanitization failed:', sanitizeError.message);
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
            console.error('[ERROR]','リダイレクト完全失敗:', currentError);
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
    console.error('[ERROR]','❌ サーバーサイドログアウト処理でエラー:', error.message);
    console.error('[ERROR]','❌ エラースタック:', error.stack);
    console.error('[ERROR]','❌ エラーの型:', typeof error);
    console.error('[ERROR]','❌ エラーオブジェクト:', error);

    // Step 5: エラー時のフォールバックHTML（安全なエスケープ）
    const safeErrorMessage = String(error.message || 'Unknown error')
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/"/g, '\\"');

    const fallbackScript = `
      <script>
        console.error('[ERROR]','サーバーサイドログアウトエラー: ${safeErrorMessage}');
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
      console.error('[ERROR]','❌ Fallback HTML creation failed:', fallbackError.message);
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
    console.error('[ERROR]','アカウント切り替え検出中にエラー:', error.message);
    return {
      isAccountSwitch: false,
      previousEmail: null,
      currentEmail: currentEmail,
      error: error.message
    };
  }
}
