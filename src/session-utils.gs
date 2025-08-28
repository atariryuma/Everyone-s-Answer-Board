/**
 * @fileoverview セッション管理ユーティリティ
 * アカウント間のセッション分離と整合性管理を提供
 */

// 回復力のある実行機構を使用
function getResilientPropertiesService() {
  return resilientExecutor.execute(() => PropertiesService.getUserProperties(), {
    name: 'PropertiesService.getUserProperties',
    idempotent: true,
  });
}

function getResilientCacheService() {
  return resilientExecutor.execute(() => CacheService.getUserCache(), {
    name: 'CacheService.getUserCache',
    idempotent: true,
  });
}

function getResilientScriptCache() {
  return resilientExecutor.execute(() => CacheService.getScriptCache(), {
    name: 'CacheService.getScriptCache',
    idempotent: true,
  });
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
          ULog.debug(label + ': removeAll(keys) でキャッシュを削除しました');
          return { success: true, method: 'removeAll(keys)' };
        } catch (e) {
          ULog.warn(label + ': removeAll(keys) でエラー: ' + e.message);
        }
      }

      // キー未指定でも removeAll() が使える環境（非標準だが一部で提供される可能性）
      try {
        cache.removeAll();
        ULog.debug(label + ': removeAll() でキャッシュを全消去しました');
        return { success: true, method: 'removeAll()' };
      } catch (e) {
        // 環境により未サポート
        ULog.warn(label + ': removeAll() は未サポート: ' + e.message);
      }
    }

    // 2) プレフィックス + email で個別削除（Cache.remove）
    if (typeof cache.remove === 'function' && prefixes.length > 0 && email) {
      prefixes.forEach(function (prefix) {
        try {
          cache.remove(prefix + email);
        } catch (e) {
          // 続行
          ULog.warn(label + ': プレフィックス削除中のエラー: ' + e.message);
        }
      });
      ULog.debug(label + ': 既知のプレフィックスキーを削除しました');
      return { success: true, method: 'remove(prefix+email)' };
    }

    // 3) keys があれば removeAll(keys) を最終試行
    if (typeof cache.removeAll === 'function' && keys.length > 0) {
      try {
        cache.removeAll(keys);
        ULog.debug(label + ': removeAll(keys) でキャッシュを削除しました');
        return { success: true, method: 'removeAll(keys)' };
      } catch (e) {
        ULog.warn(label + ': removeAll(keys) 最終試行で失敗: ' + e.message);
      }
    }

    // 4) 何もできない環境
    ULog.warn(label + ': キャッシュ全面削除 API 未提供のためスキップしました');
    return { success: false, method: 'skipped' };
  } catch (error) {
    ULog.warn('clearCacheSafely エラー: ' + (error && error.message));
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
    ULog.debug('セッションクリーンアップを開始: ' + currentEmail);

    const props = getResilientPropertiesService();
    const userCache = getResilientCacheService();

    // 現在のユーザーのハッシュキーを生成
    const currentUserKey =
      'CURRENT_USER_ID_' +
      Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, currentEmail, Utilities.Charset.UTF_8)
        .map(function (byte) {
          return (byte + 256).toString(16).slice(-2);
        })
        .join('');

    // 古い形式のキャッシュを完全削除
    props.deleteProperty('CURRENT_USER_ID');

    // 他のユーザーIDキャッシュをクリア（現在のユーザー以外）
    const allProperties = props.getProperties();
    Object.keys(allProperties).forEach(function (key) {
      if (key.startsWith('CURRENT_USER_ID_') && key !== currentUserKey) {
        props.deleteProperty(key);
        ULog.debug('削除された古いユーザーキャッシュ: ' + key);
      }
    });

    // ユーザーキャッシュを安全にクリア
    clearCacheSafely(userCache, { label: 'UserCache' });

    // スクリプトキャッシュの関連項目もクリア
    const scriptCache = getResilientScriptCache();
    clearCacheSafely(scriptCache, {
      label: 'ScriptCache',
      email: currentEmail,
      prefixes: ['config_v3_', 'user_', 'email_'],
    });

    ULog.debug('セッションクリーンアップ完了: ' + currentEmail);
  } catch (error) {
    ULog.error('セッションクリーンアップでエラー', { error: error.message }, ULog.CATEGORIES.AUTH);
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
    errors: [],
  };

  try {
    ULog.debug('🔄 ユーザー認証をリセット開始...');
    const startTime = Date.now();

    // Step 1: キャッシュクリア（非致命的エラー許容）
    try {
      const userCache = CacheService.getUserCache();
      if (userCache) {
        clearCacheSafely(userCache, { label: 'UserCache' });
      }

      const scriptCache = CacheService.getScriptCache();
      if (scriptCache) {
        clearCacheSafely(scriptCache, {
          label: 'ScriptCache',
          email: getCurrentUserEmail(),
          prefixes: ['config_v3_', 'user_', 'email_'],
        });
      }

      authResetResult.cacheCleared = true;
      ULog.debug('✅ キャッシュクリア完了');
    } catch (cacheError) {
      authResetResult.errors.push(`キャッシュクリアエラー: ${cacheError.message}`);
      ULog.warn('⚠️ キャッシュクリアでエラーが発生しましたが、処理を続行します:', cacheError.message);
    }

    // Step 2: プロパティクリア（段階的実装）
    try {
      const props = PropertiesService.getUserProperties();

      // まず deleteAllProperties を試行
      if (typeof props.deleteAllProperties === 'function') {
        props.deleteAllProperties();
        authResetResult.propertiesCleared = true;
        ULog.debug('✅ PropertiesService.deleteAllProperties() でクリア完了');
      } else {
        // フォールバック: 重要なプロパティを個別削除
        ULog.debug('⚠️ deleteAllProperties が利用できません。個別削除を実行します。');
        const importantKeys = [
          'LAST_ACCESS_EMAIL',
          'lastAdminUserId',
          'CURRENT_USER_ID',
          'USER_CACHE_KEY',
          'SESSION_ID',
        ];

        // 既存のプロパティをすべて取得して削除
        try {
          const allProps = props.getProperties();
          const keys = Object.keys(allProps);

          if (keys.length > 0) {
            ULog.debug(`📝 ${keys.length}個のプロパティを個別削除します:`, keys);

            for (const key of keys) {
              try {
                props.deleteProperty(key);
              } catch (deleteError) {
                authResetResult.errors.push(`プロパティ削除エラー(${key}): ${deleteError.message}`);
              }
            }

            authResetResult.propertiesCleared = true;
            ULog.debug('✅ 個別プロパティ削除完了');
          } else {
            ULog.debug('ℹ️ 削除対象のプロパティはありませんでした');
            authResetResult.propertiesCleared = true;
          }
        } catch (enumError) {
          // プロパティ一覧取得に失敗した場合、重要なキーのみ削除を試行
          ULog.debug('⚠️ プロパティ一覧取得に失敗。重要なキーのみ削除を試行します。');
          for (const key of importantKeys) {
            try {
              props.deleteProperty(key);
            } catch (deleteError) {
              // 個別削除エラーは記録するが処理は続行
              authResetResult.errors.push(
                `重要プロパティ削除エラー(${key}): ${deleteError.message}`
              );
            }
          }
          authResetResult.propertiesCleared = true;
        }
      }
    } catch (propsError) {
      authResetResult.errors.push(`プロパティクリアエラー: ${propsError.message}`);
      ULog.error('プロパティクリアでエラーが発生しました', { error: propsError.message }, ULog.CATEGORIES.AUTH);

      // プロパティクリアに失敗してもログアウト処理は続行
      ULog.warn('⚠️ プロパティクリアに失敗しましたが、ログアウト処理を続行します');
    }

    // Step 3: ログインページURLの取得
    try {
      const loginPageUrl = ScriptApp.getService().getUrl();
      authResetResult.loginUrl = loginPageUrl;

      const executionTime = Date.now() - startTime;
      ULog.debug(`✅ 認証リセット完了 (${executionTime}ms):`, {
        cacheCleared: authResetResult.cacheCleared,
        propertiesCleared: authResetResult.propertiesCleared,
        errorsCount: authResetResult.errors.length,
        loginUrl: loginPageUrl,
      });

      return loginPageUrl;
    } catch (urlError) {
      authResetResult.errors.push(`URL取得エラー: ${urlError.message}`);
      throw new Error(`ログインページURL取得に失敗: ${urlError.message}`);
    }
  } catch (error) {
    const executionTime = Date.now() - (authResetResult.startTime || Date.now());
    ULog.error('ユーザー認証リセット中に致命的エラー', {
      error: error.message,
      executionTime: executionTime + 'ms',
      partialResults: authResetResult,
    }, ULog.CATEGORIES.AUTH);

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
  ULog.debug('🔄 forceLogoutAndRedirectToLogin - 関数開始');
  ULog.debug('🔍 Function called at:', new Date().toISOString());
  ULog.debug('🔍 Available functions check:');
  ULog.debug('  - getWebAppUrl:', typeof getWebAppUrl);
  ULog.debug('  - sanitizeRedirectUrl:', typeof sanitizeRedirectUrl);
  ULog.debug('  - HtmlService:', typeof HtmlService);

  try {
    ULog.debug('✅ forceLogoutAndRedirectToLogin - try block内に入りました');

    // Step 1: セッションクリア処理（エラーハンドリング強化）
    try {
      ULog.debug('🧹 キャッシュクリア開始...');

      // キャッシュクリア（非致命的エラー許容）
      try {
        const userCache = CacheService.getUserCache();
        if (userCache) {
          clearCacheSafely(userCache, { label: 'UserCache' });
        }

        const scriptCache = CacheService.getScriptCache();
        if (scriptCache) {
          clearCacheSafely(scriptCache, {
            label: 'ScriptCache',
            email: getCurrentUserEmail(),
            prefixes: ['config_v3_', 'user_', 'email_'],
          });
        }
        ULog.debug('✅ キャッシュクリア完了');
      } catch (cacheError) {
        ULog.warn(
          '⚠️ キャッシュクリアでエラーが発生しましたが、処理を続行します:',
          cacheError.message
        );
      }

      // プロパティクリア（段階的実装）
      try {
        const props = PropertiesService.getUserProperties();

        // まず deleteAllProperties を試行
        if (typeof props.deleteAllProperties === 'function') {
          props.deleteAllProperties();
          ULog.debug('✅ PropertiesService.deleteAllProperties() でクリア完了');
        } else {
          // フォールバック: 既存のプロパティをすべて取得して削除
          ULog.debug('⚠️ deleteAllProperties が利用できません。個別削除を実行します。');
          try {
            const allProps = props.getProperties();
            const keys = Object.keys(allProps);

            if (keys.length > 0) {
              ULog.debug(`📝 ${keys.length}個のプロパティを個別削除します:`, keys);

              for (const key of keys) {
                try {
                  props.deleteProperty(key);
                } catch (deleteError) {
                  ULog.warn(`プロパティ削除エラー(${key}): ${deleteError.message}`);
                }
              }

              ULog.debug('✅ 個別プロパティ削除完了');
            } else {
              ULog.debug('ℹ️ 削除対象のプロパティはありませんでした');
            }
          } catch (enumError) {
            // プロパティ一覧取得に失敗した場合、重要なキーのみ削除を試行
            ULog.debug('⚠️ プロパティ一覧取得に失敗。重要なキーのみ削除を試行します。');
            const importantKeys = [
              'LAST_ACCESS_EMAIL',
              'lastAdminUserId',
              'CURRENT_USER_ID',
              'USER_CACHE_KEY',
              'SESSION_ID',
            ];
            for (const key of importantKeys) {
              try {
                props.deleteProperty(key);
              } catch (deleteError) {
                ULog.warn(`重要プロパティ削除エラー(${key}): ${deleteError.message}`);
              }
            }
          }
        }
        ULog.debug('✅ ユーザープロパティクリア完了');
      } catch (propsError) {
        ULog.warn(
          '⚠️ プロパティクリアでエラーが発生しましたが、ログアウト処理を続行します:',
          propsError.message
        );
      }
    } catch (cacheError) {
      ULog.warn('⚠️ セッションクリア中に一部エラー:', cacheError.message);
      ULog.warn('⚠️ エラースタック:', cacheError.stack);
      // セッションクリアエラーは致命的ではないので継続
    }

    // Step 2: ログインページURLの生成と適切なサニタイズ
    let loginUrl;
    try {
      ULog.debug('🔗 URL生成開始...');

      // getWebAppUrl関数の存在確認
      if (typeof getWebAppUrl !== 'function') {
        throw new Error('getWebAppUrl function not found');
      }

      const rawUrl = getWebAppUrl() + '?mode=login';
      ULog.debug('📝 Raw URL generated:', rawUrl);

      // sanitizeRedirectUrl関数の存在確認
      if (typeof sanitizeRedirectUrl !== 'function') {
        throw new Error('sanitizeRedirectUrl function not found');
      }

      loginUrl = sanitizeRedirectUrl(rawUrl);
      ULog.debug('✅ ログインURL生成・サニタイズ成功:', loginUrl);
    } catch (urlError) {
      ULog.warn('⚠️ WebAppURL取得失敗、フォールバック使用:', urlError.message);
      ULog.warn('⚠️ URLエラースタック:', urlError.stack);

      const fallbackUrl = ScriptApp.getService().getUrl() + '?mode=login';
      ULog.debug('📝 Fallback URL:', fallbackUrl);

      try {
        loginUrl = sanitizeRedirectUrl(fallbackUrl);
      } catch (sanitizeError) {
        ULog.error('Fallback URL sanitization failed', { error: sanitizeError.message }, ULog.CATEGORIES.AUTH);
        loginUrl = fallbackUrl; // 最終フォールバック
      }
    }

    ULog.debug('🎯 Final login URL:', loginUrl);

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
    ULog.debug('🔒 Escaped login URL:', safeLoginUrl);

    const redirectScript = `
      <script>
        ULog.debug('🚀 サーバーサイドリダイレクト実行:', '${safeLoginUrl}');

        // Google Apps Script環境に最適化されたリダイレクト
        try {
          // 最も確実な方法：window.top.location.href
          window.top.location.href = '${safeLoginUrl}';
        } catch (topError) {
          ULog.warn('Top frame遷移失敗:', topError);
          try {
            // フォールバック：現在のウィンドウでリダイレクト
            window.location.href = '${safeLoginUrl}';
          } catch (currentError) {
            ULog.error('リダイレクト完全失敗', { error: currentError }, ULog.CATEGORIES.AUTH);
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

    ULog.debug('📄 Generated HTML script length:', redirectScript.length);
    ULog.debug('📄 Generated HTML preview (first 200 chars):', redirectScript.substring(0, 200));

    // HtmlServiceの存在確認
    if (typeof HtmlService === 'undefined') {
      throw new Error('HtmlService is not available');
    }

    const htmlOutput = HtmlService.createHtmlOutput(redirectScript);
    ULog.debug('✅ HtmlService.createHtmlOutput 成功');

    // HtmlOutputの内容確認
    if (!htmlOutput) {
      throw new Error('HtmlOutput is null or undefined');
    }

    // XFrameOptionsMode を安全に設定（iframe内での動作を許可）
    try {
      if (HtmlService && HtmlService.XFrameOptionsMode && HtmlService.XFrameOptionsMode.ALLOWALL) {
        htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
        ULog.debug('✅ XFrameOptionsMode.ALLOWALL設定完了');
      }
    } catch (frameError) {
      ULog.warn('XFrameOptionsMode設定失敗:', frameError.message);
    }

    // 最終検証: HtmlOutputの内容を確認
    try {
      const outputContent = htmlOutput.getContent();
      ULog.debug(
        '📋 HtmlOutput content length:',
        outputContent ? outputContent.length : 'null/undefined'
      );
      ULog.debug(
        '📋 HtmlOutput content preview:',
        outputContent ? outputContent.substring(0, 100) : 'NO CONTENT'
      );
    } catch (contentError) {
      ULog.warn('⚠️ Cannot access HtmlOutput content:', contentError.message);
    }

    ULog.debug('✅ サーバーサイドリダイレクトHTML生成完了 - 正常終了');
    return htmlOutput;
  } catch (error) {
    ULog.error('サーバーサイドログアウト処理でエラー', { error: error.message }, ULog.CATEGORIES.AUTH);
    ULog.error('エラースタック', { stack: error.stack }, ULog.CATEGORIES.AUTH);
    ULog.error('エラーの型', { type: typeof error }, ULog.CATEGORIES.AUTH);
    ULog.error('エラーオブジェクト', { error }, ULog.CATEGORIES.AUTH);

    // Step 5: エラー時のフォールバックHTML（安全なエスケープ）
    const safeErrorMessage = String(error.message || 'Unknown error')
      .replace(/\\/g, '\\\\')
      .replace(/'/g, "\\'")
      .replace(/"/g, '\\"');

    const fallbackScript = `
      <script>
        ULog.error('サーバーサイドログアウトエラー', { message: safeErrorMessage }, ULog.CATEGORIES.AUTH);
        alert('ログアウト処理中にエラーが発生しました。\\n\\n詳細: ${safeErrorMessage}\\n\\nページを再読み込みします。');
        window.location.reload();
      </script>
      <p>エラーが発生しました: ${safeErrorMessage}</p>
      <p>ページを再読み込みしています...</p>
    `;

    ULog.debug('📄 Fallback HTML generated');

    try {
      const fallbackOutput = HtmlService.createHtmlOutput(fallbackScript);
      ULog.debug('✅ Fallback HtmlOutput created successfully');
      return fallbackOutput;
    } catch (fallbackError) {
      ULog.error('Fallback HTML creation failed', { error: fallbackError.message }, ULog.CATEGORIES.AUTH);
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
      ULog.debug('アカウント切り替えを検出:', lastEmail, '->', currentEmail);
      cleanupSessionOnAccountSwitch(currentEmail);
      clearDatabaseCache();
    }

    // 現在のメールアドレスを記録
    props.setProperty(lastEmailKey, currentEmail);

    return {
      isAccountSwitch: isAccountSwitch,
      previousEmail: lastEmail,
      currentEmail: currentEmail,
    };
  } catch (error) {
    ULog.error('アカウント切り替え検出中にエラー', { error: error.message }, ULog.CATEGORIES.AUTH);
    return {
      isAccountSwitch: false,
      previousEmail: null,
      currentEmail: currentEmail,
      error: error.message,
    };
  }
}
