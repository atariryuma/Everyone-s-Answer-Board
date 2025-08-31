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
  
  const initialConfig = {
    userId,
    userEmail,
    title: `${userEmail}の回答ボード`,
    setupStatus: 'pending',
    formCreated: false,
    appPublished: false,
    isPublic: false,
    allowAnonymous: false,
    columns: [
      { name: 'timestamp', label: 'タイムスタンプ', type: 'datetime', required: false },
      { name: 'email', label: 'メールアドレス', type: 'email', required: false },
      { name: 'class', label: 'クラス', type: 'text', required: false },
      { name: 'opinion', label: '回答', type: 'textarea', required: true },
      { name: 'reason', label: '理由', type: 'textarea', required: false },
      { name: 'name', label: '名前', type: 'text', required: false }
    ],
    theme: 'default',
    createdAt: timestamp,
    lastModified: timestamp
  };

  return {
    userId,
    userEmail,
    createdAt: timestamp,
    lastAccessedAt: timestamp,
    isActive: true,
    spreadsheetId: '',
    spreadsheetUrl: '',
    configJson: JSON.stringify(initialConfig),
    formUrl: ''
  };
}

/**
 * リダイレクト関数
 * @param {string} url リダイレクト先URL
 * @returns {HtmlOutput} リダイレクトHTML
 */
function createRedirect(url) {
  return HtmlService.createHtmlOutput('')
    .addMetaTag('http-equiv', 'refresh')
    .addMetaTag('content', `0;url=${url}`);
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
  
  // 作成確認
  if (!waitForUserRecord(completeUserData.userId, 3000, 500)) {
    throw new Error('ユーザー作成の確認がタイムアウトしました');
  }
  
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
        
        // 新規ユーザーの管理パネルへリダイレクト
        const adminUrl = buildUserAdminUrl(newUser.userId);
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
      severity: 'high'
    };
    console.error('🚨 processLoginFlow 重大エラー:', JSON.stringify(errorInfo, null, 2));

    // ユーザーフレンドリーなエラーメッセージ
    const userMessage = error.message.includes('ユーザー')
      ? error.message
      : 'ログイン処理中に予期しないエラーが発生しました。しばらく待ってから再度お試しください。';

    return showErrorPage('ログインエラー', userMessage, error);
  }
}

/**
 * ユーザー作成のロールバック処理
 * @param {string} userId ロールバック対象のユーザーID
 */
function rollbackUserCreation(userId) {
  try {
    console.log(`ユーザー作成ロールバック開始: ${userId}`);
    
    const deleted = DB.deleteUser(userId);
    
    if (deleted) {
      console.log(`ロールバック完了: ${userId}`);
    } else {
      console.warn(`ロールバック失敗 - ユーザーが見つからない: ${userId}`);
    }
  } catch (error) {
    console.error(`ロールバック処理でエラー: ${error.message}`);
  }
}
