/**
 * @fileoverview 認証管理 - JWTトークンキャッシュと最適化
 * GAS互換の関数ベースの実装
 */

/**




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
      if (isTrue(userInfo.isActive)) {
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
        return createSecureRedirect(adminUrl, welcomeMessage);
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

      // 3a. 新規ユーザーデータを準備（初期設定でpending状態）
      const initialConfig = {
        setupStatus: 'pending',
        createdAt: new Date().toISOString(),
        formCreated: false,
        appPublished: false,
      };

      const newUser = {
        userId: Utilities.getUuid(),
        adminEmail: userEmail,
        createdAt: new Date().toISOString(),
        isActive: true, // 即時有効化
        configJson: JSON.stringify(initialConfig),
        spreadsheetId: '',
        spreadsheetUrl: '',
        lastAccessedAt: new Date().toISOString(),
      };

      // 3b. データベースに作成
      DB.createUser(newUser);
      // ConfigurationManagerで初期設定を作成
      App.getConfig().initializeUserConfig(newUser.tenantId, userEmail);
      
      if (!waitForUserRecord(newUser.tenantId, 3000, 500)) {
        console.warn('processLoginFlow: user not found after create:', newUser.tenantId);
      }
      console.log('processLoginFlow: 新規ユーザー作成完了:', newUser.tenantId);

      // 3c. 新規ユーザーの管理パネルへリダイレクト
      const adminUrl = buildUserAdminUrl(newUser.tenantId);
      return createSecureRedirect(adminUrl, 'ようこそ！セットアップを開始してください');
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
      severity: 'high', // ログインエラーは高重要度
    };
    console.error('🚨 processLoginFlow 重大エラー:', JSON.stringify(errorInfo, null, 2));

    // ユーザーフレンドリーなエラーメッセージ
    const userMessage = error.message.includes('ユーザー')
      ? error.message
      : 'ログイン処理中に予期しないエラーが発生しました。しばらく待ってから再度お試しください。';

    return showErrorPage('ログインエラー', userMessage, error);
  }
}
