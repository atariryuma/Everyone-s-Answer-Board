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

  // ConfigOptimizerの最適化形式を使用（重複データを除去）
  const optimizedConfig = {
    title: `${userEmail}の回答ボード`,
    setupStatus: 'pending',
    formCreated: false,
    appPublished: false,
    isPublic: false,
    allowAnonymous: false,
    sheetName: null,
    columnMapping: {},
    theme: 'default',
    lastModified: timestamp,
  };

  // columnsは必要時のみ保持（デフォルト列設定）
  optimizedConfig.columns = [
    { name: 'timestamp', label: 'タイムスタンプ', type: 'datetime', required: false },
    { name: 'email', label: 'メールアドレス', type: 'email', required: false },
    { name: 'class', label: 'クラス', type: 'text', required: false },
    { name: 'opinion', label: '回答', type: 'textarea', required: true },
    { name: 'reason', label: '理由', type: 'textarea', required: false },
    { name: 'name', label: '名前', type: 'text', required: false },
  ];

  console.info('新規ユーザー作成: 最適化済みconfigJSON使用', {
    userEmail,
    optimizedSize: JSON.stringify(optimizedConfig).length,
    removedFields: ['userId', 'userEmail', 'createdAt'], // DB列に移行済み
  });

  return {
    userId,
    userEmail,
    createdAt: timestamp,
    lastAccessedAt: timestamp,
    isActive: true,
    spreadsheetId: '',
    spreadsheetUrl: '',
    configJson: JSON.stringify(optimizedConfig),
    formUrl: '',
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
 * @returns {Object} ユーザー情報
 */
function handleUserRegistration(userEmail) {
  const existingUser = DB.findUserByEmail(userEmail);

  if (existingUser) {
    // 既存ユーザー: 最終アクセス時刻のみ更新
    updateUserLastAccess(existingUser.userId);
    return existingUser;
  }

  // 新規ユーザー: 完全データ作成
  const completeUserData = createCompleteUser(userEmail);
  DB.createUser(completeUserData);

  console.log('新規ユーザー作成完了:', completeUserData.userId);
  return completeUserData;
}

/**
 * ログインフローを処理し、適切なページにリダイレクトする
 * 既存ユーザーの設定を保護しつつ、セットアップ状況に応じたメッセージを表示
 * @param {string} userEmail ログインユーザーのメールアドレス
 * @returns {HtmlOutput} 表示するHTMLコンテンツ
 */
function processLoginFlow(userEmail) {
  try {
    if (!userEmail) {
      throw new Error('ユーザーメールアドレスが指定されていません');
    }

    // 1. ユーザー情報をデータベースから取得
    const userInfo = DB.findUserByEmail(userEmail);

    // 2. 既存ユーザーの処理
    if (userInfo) {
      // 2a. アクティブユーザーの場合
      if (userInfo.isActive === true) {
        console.log('processLoginFlow: 既存アクティブユーザー:', userEmail);

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
        const newUser = handleUserRegistration(userEmail);

        console.log('processLoginFlow: 新規ユーザー作成完了:', newUser.userId);

        // キャッシュ整合性確保: 新規作成後の検証リトライ
        let verifiedUser = DB.findUserByEmail(userEmail);
        if (!verifiedUser) {
          console.warn('processLoginFlow: キャッシュ不整合検出、リトライ実行');
          Utilities.sleep(100); // キャッシュ同期待機
          verifiedUser = DB.findUserByEmail(userEmail);
          if (!verifiedUser) {
            console.error('processLoginFlow: ユーザー検証失敗 - キャッシュ問題の可能性');
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
