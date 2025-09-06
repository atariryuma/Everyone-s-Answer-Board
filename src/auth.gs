/**
 * @fileoverview 認証管理 - JWTトークンキャッシュと最適化
 * GAS互換の関数ベースの実装
 */

/**
 * 完全なユーザーデータを作成する統一関数
 * @param {string} userEmail ユーザーメールアドレス
 * @returns {Object} 完全なユーザーデータオブジェクト
 */
function createCompleteUser(userEmail) {
  const userId = Utilities.getUuid();
  const timestamp = new Date().toISOString();

  // ✅ CLAUDE.md完全準拠：最小限configJSONで開始（マイグレーション不要）
  const minimalConfig = {
    // 🎯 必要最小限のみ：重複データ完全排除
    setupStatus: 'pending',
    appPublished: false,
    displaySettings: {
      showNames: false, // CLAUDE.md準拠：心理的安全性重視
      showReactions: false,
    },
    createdAt: timestamp,
    lastModified: timestamp,
  };

  console.log('新規ユーザー作成: 最適化済みconfigJSON使用', {
    userEmail,
    configSize: JSON.stringify(minimalConfig).length,
    removedFields: ['userId', 'userEmail', 'createdAt'], // DB列に移行済み
  });

  // ✅ CLAUDE.md完全準拠：5フィールドDB構造のみ返却
  return {
    userId,
    userEmail,
    isActive: 'true',  // 文字列として統一（Sheetsとの一貫性）
    configJson: JSON.stringify(minimalConfig),
    lastModified: timestamp,
  };
}

/**
 * リダイレクト関数 - X-Frame-Options対応
 * @param {string} url リダイレクト先URL
 * @returns {HtmlOutput} リダイレクトHTML
 */
function createRedirect(url) {
  const script = `
    <script>
      try {
        if (window.top && window.top.location) {
          window.top.location.href = '${url}';
        } else {
          window.location.href = '${url}';
        }
      } catch (e) {
        window.location.href = '${url}';
      }
    </script>
  `;
  return HtmlService.createHtmlOutput(script);
}

/**
 * 統一ユーザー登録処理
 * @param {string} userEmail ユーザーメールアドレス
 * @param {boolean} bypassCache キャッシュをバイパスするか（ログイン時true推奨）
 * @returns {Object} ユーザー情報
 */
function handleUserRegistration(userEmail, bypassCache = false) {
  // ログイン時（はじめるボタン）はキャッシュをバイパスして最新データを取得
  const existingUser = bypassCache ? DB.findUserByEmail(userEmail) : DB.findUserByEmail(userEmail);

  if (bypassCache) {
    console.log('🔄 handleUserRegistration: キャッシュバイパスモード');
  }

  if (existingUser) {
    // 既存ユーザー: 最終アクセス時刻のみ更新
    updateUserLastAccess(existingUser.userId);
    return existingUser;
  }

  // 新規ユーザー: 完全データ作成
  const completeUserData = createCompleteUser(userEmail);
  DB.createUser(completeUserData);

  console.log('新規ユーザー作成完了');
  return completeUserData;
}

/**
 * 🔄 ユーザーの最終アクセス時刻を更新（設定は保護）
 * @param {string} userId ユーザーID
 */
function updateUserLastAccess(userId) {
  try {
    if (!userId) {
      console.warn('updateUserLastAccess: userIdが指定されていません');
      return;
    }

    const now = new Date().toISOString();
    const currentUser = DB.findUserById(userId);

    if (!currentUser) {
      console.warn('updateUserLastAccess: ユーザーが見つかりません:', userId);
      return;
    }

    // 既存のconfigJsonを保護しつつ、lastAccessedAtのみ更新
    const existingConfig = JSON.parse(currentUser.configJson || '{}');
    const updatedConfig = { ...existingConfig, lastAccessedAt: now };

    // configJsonと lastModified のみ更新
    DB.updateUser(userId, {
      configJson: JSON.stringify(updatedConfig),
      lastModified: now,
    });

    console.log('updateUserLastAccess: 最終アクセス時刻更新完了');
  } catch (error) {
    console.error('updateUserLastAccess エラー:', error.message);
  }
}

/**
 * 🔍 設定JSONからセットアップ状況を取得
 * @param {string} configJson 設定JSON文字列
 * @returns {string} セットアップ状況
 */
function getSetupStatusFromConfig(configJson) {
  try {
    const config = JSON.parse(configJson || '{}');
    return config.setupStatus || 'pending';
  } catch (error) {
    console.warn('getSetupStatusFromConfig: JSON解析エラー:', error.message);
    return 'pending';
  }
}

/**
 * 専用の管理パネルアクセス処理（クエリパラメータ mode=admin 必須）
 * @param {string} userEmail ログインユーザーのメールアドレス
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function processAdminAccess(userEmail) {
  try {
    if (!userEmail) {
      throw new Error('ユーザーメールアドレスが指定されていません');
    }

    // 1. ユーザー情報をデータベースから直接取得（ログイン時はキャッシュバイパス）
    const userInfo = DB.findUserByEmail(userEmail);
    console.log('🔄 processLoginFlow: キャッシュバイパスでユーザー検索');

    // 2. 既存ユーザーの処理
    if (userInfo) {
      // 2a. アクティブユーザーの場合
      if (userInfo.isActive === true) {
        console.log('processLoginFlow: 既存アクティブユーザー検出');

        // 最終アクセス時刻を更新（設定は保護）
        updateUserLastAccess(userInfo.userId);

        // セットアップ状況を確認してメッセージを調整
        const setupStatus = getSetupStatusFromConfig(userInfo.configJson);
        let welcomeMessage = '管理パネルへようこそ';

        if (setupStatus === 'pending') {
          welcomeMessage = 'セットアップを続行してください';
        } else if (setupStatus === 'completed') {
          welcomeMessage = 'おかえりなさい！';
        }

        const adminUrl = buildUserAdminUrl(userInfo.userId);
        return createRedirect(adminUrl);
      }
      // 2b. 非アクティブユーザーの場合
      else {
        console.warn('processLoginFlow: 既存だが非アクティブなユーザー:', userEmail);
        return showErrorPage(
          'アカウントが無効です',
          'あなたのアカウントは現在無効化されています。管理者にお問い合わせください。'
        );
      }
    }
    // 3. 新規ユーザーの処理
    else {
      console.log('processLoginFlow: 新規ユーザー登録開始:', userEmail);

      try {
        // 統一ユーザー作成関数を使用
        const newUser = handleUserRegistration(userEmail, true); // ログイン時はキャッシュバイパス

        console.log('processLoginFlow: 新規ユーザー作成完了:', newUser.userId);

        // 🔧 修正：ログイン時は常にキャッシュバイパスで検証
        let verifiedUser = DB.findUserByEmail(userEmail);
        if (!verifiedUser) {
          console.warn('processLoginFlow: DB直接検索でもユーザー未発見、再試行');
          Utilities.sleep(200); // DB同期待機
          verifiedUser = DB.findUserByEmail(userEmail);
          if (!verifiedUser) {
            console.error('processLoginFlow: ユーザー検証失敗 - DB同期問題の可能性');
            // フォールバック: 作成したユーザー情報を直接使用
            verifiedUser = newUser;
          }
        }

        // 新規ユーザーの管理パネルへリダイレクト
        const adminUrl = buildUserAdminUrl(verifiedUser.userId);
        return createRedirect(adminUrl);
      } catch (error) {
        console.error('新規ユーザー作成失敗:', error);
        throw error;
      }
    }
  } catch (error) {
    // 構造化エラーログの出力
    const errorInfo = {
      timestamp: new Date().toISOString(),
      function: 'processLoginFlow',
      userEmail: userEmail || 'unknown',
      errorType: error.name || 'UnknownError',
      message: error.message,
      stack: error.stack,
      severity: 'high',
    };
    console.error('🚨 processLoginFlow 重大エラー:', JSON.stringify(errorInfo, null, 2));

    // ユーザーフレンドリーなエラーメッセージ
    const userMessage = error.message.includes('ユーザー')
      ? error.message
      : 'ログイン処理中に予期しないエラーが発生しました。しばらく待ってから再度お試しください。';

    return showErrorPage('ログインエラー', userMessage, error);
  }
}
