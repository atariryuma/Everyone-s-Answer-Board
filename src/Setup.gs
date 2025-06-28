/**
 * @fileoverview StudyQuest -みんなの回答ボード- - 統合セットアップスクリプト
 * フォルダ管理、データベース作成、共有設定、プロパティ設定を自動で行う統合セットアップ機能
 */

// =================================================================
// ユーティリティ関数
// =================================================================

/**
 * デバッグログ関数
 */
function debugLog() {
  if (typeof console !== 'undefined' && console.log) {
    console.log.apply(console, arguments);
  }
}

/**
 * セキュリティ強化: ID情報をログ用にサニタイズ
 * @param {string} id - サニタイズするID
 * @return {string} サニタイズされたID
 */
function sanitizeIdForLog(id) {
  if (!id || typeof id !== 'string') return 'なし';
  return id.length > 8 ? id.substring(0, 8) + '...' : id;
}

/**
 * セキュリティ強化: DEPLOY_IDの厳密な検証
 * @param {string} deployId - 検証するDEPLOY_ID
 * @return {boolean} 有効な場合true
 */
function validateDeployId(deployId) {
  if (!deployId || typeof deployId !== 'string') return false;
  // Google Apps ScriptのDEPLOY_IDの実際の形式に合わせた厳密な検証
  return /^AKfycb[a-zA-Z0-9_-]{20,}$/.test(deployId);
}

/**
 * セキュリティ強化: ウェブアプリURLの検証
 * @param {string} url - 検証するURL
 * @return {boolean} 有効な場合true
 */
function validateWebAppUrl(url) {
  try {
    const urlObj = new URL(url);
    return urlObj.hostname === 'script.google.com' && 
           urlObj.pathname.includes('/macros/s/') &&
           /\/s\/[a-zA-Z0-9_-]+\/exec$/.test(urlObj.pathname);
  } catch {
    return false;
  }
}

/**
 * セキュリティ強化: エラーの安全なログ記録
 * @param {string} context - エラーのコンテキスト
 * @param {Error} error - エラーオブジェクト
 */
function secureLogError(context, error) {
  // 本番環境では詳細なエラーログは管理者のみがアクセス可能な場所に記録
  // ここでは基本的な情報のみをログに記録
  console.warn(`${context}: エラーが発生しました`);

  // 開発環境でのみ詳細ログを出力（DEBUG フラグがある場合）
  if (typeof DEBUG !== 'undefined' && DEBUG) {
    console.error(`Debug - ${context}:`, error.message);
  }
}

/**
 * StudyQuestの統合セットアップ関数（簡易版）
 * DEPLOY_IDを指定してセットアップを実行
 * @param {string} deployId - デプロイID（必須）
 */
function studyQuestSetupSimple(deployId) {
  if (!deployId) {
    throw new Error("DEPLOY_IDの指定が必要です。例: studyQuestSetupSimple('AKfycbxYourDeployId')");
  }
  return studyQuestSetup(deployId);
}

/**
 * StudyQuestの統合セットアップ関数
 * フォルダ作成、データベース作成、ファイル移動、プロパティ設定を自動で行う
 * @param {string} [manualDeployId] - 手動で設定するデプロイID（任意）
 */
function studyQuestSetup(manualDeployId) {
  // 認証チェック
  const currentUser = Session.getActiveUser().getEmail();
  if (!currentUser) {
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }

  const FOLDER_NAME = "StudyQuest - みんなの回答ボード"; // アプリ専用フォルダ名
  const DB_FILENAME = "StudyQuest_UserDatabase";     // データベースファイル名
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 30秒待機

  try {
    debugLog("🚀 StudyQuest 統合セットアップを開始します...");

    // ステップ1: 専用フォルダの確認と作成
    let folder;
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) {
      folder = folders.next();
      debugLog(`✅ 専用フォルダ「${FOLDER_NAME}」が既に存在します。`);
    } else {
      folder = DriveApp.createFolder(FOLDER_NAME);
      debugLog(`✅ 専用フォルダ「${FOLDER_NAME}」を新規作成しました。`);
    }

    // ステップ2: データベース（スプレッドシート）の確認と作成
    let dbFile;
    const files = folder.getFilesByName(DB_FILENAME);
    if (files.hasNext()) {
      dbFile = files.next();
      debugLog(`✅ データベースファイルがフォルダ内に既に存在します。`);
    } else {
      // マイドライブ直下も検索して、あれば移動させる
      const rootFiles = DriveApp.getRootFolder().getFilesByName(DB_FILENAME);
      if (rootFiles.hasNext()) {
        dbFile = rootFiles.next();
        dbFile.moveTo(folder);
        debugLog(`✅ 既存のデータベースファイルを専用フォルダに移動しました。`);
      } else {
        // どこにもない場合のみ新規作成
        const newDb = SpreadsheetApp.create(DB_FILENAME);
        dbFile = DriveApp.getFileById(newDb.getId());
        dbFile.moveTo(folder);
        const sheet = newDb.getSheets()[0];
        sheet.setName("users");
        // USER_DB_CONFIG.HEADERS は Code.gs で定義されていることを前提
        sheet.appendRow(["userId", "adminEmail", "spreadsheetId", "spreadsheetUrl", "createdAt", "accessToken", "configJson", "lastAccessedAt", "isActive"]);
        debugLog(`✅ データベースファイルを新規作成し、専用フォルダに配置しました。`);
      }
    }
    const dbId = dbFile.getId();
    PropertiesService.getScriptProperties().setProperty("DATABASE_ID", dbId);
    debugLog(`✅ データベースIDを設定しました: ${sanitizeIdForLog(dbId)}`);

    // ステップ2.5: データベースの共有設定
    try {
      const adminEmail = Session.getActiveUser().getEmail();
      const userDomain = adminEmail.split('@')[1];

      // データベースファイルを同一ドメイン内で閲覧可能に設定
      dbFile.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);
      debugLog(`✅ データベースファイルを同一ドメイン内で閲覧可能に設定しました。`);

      // セットアップ実行ユーザーを編集者として追加
      dbFile.addEditor(adminEmail);
      debugLog(`✅ セットアップ実行ユーザー（${adminEmail}）をデータベースの編集者として追加しました。`);

    } catch (e) {
      secureLogError('データベース共有設定', e);
    }

    // ステップ3: デプロイIDとウェブアプリURLの設定
    let deployId = manualDeployId;
    if (!deployId) {
      try {
        deployId = ScriptApp.getDeploymentId();
      } catch (e) {
        debugLog("⚠️ ScriptApp.getDeploymentId()でエラーが発生しました。手動でのDEPLOY_ID指定が必要です。");
        throw new Error("ウェブアプリとしてデプロイされていないか、DEPLOY_IDの取得に失敗しました。先にウェブアプリとしてデプロイを完了させるか、手動でDEPLOY_IDを指定してください。");
      }
    }
    
    if (!deployId) {
      throw new Error("DEPLOY_IDが取得できませんでした。先にウェブアプリとしてデプロイを完了させる必要があります。");
    }
    
    // セキュリティチェックは緩和（Google Apps Scriptの実際のDEPLOY_ID形式は様々なため）
    if (typeof deployId !== 'string' || deployId.length < 10) {
      throw new Error("無効なDEPLOY_ID形式です。");
    }
    PropertiesService.getScriptProperties().setProperty("DEPLOY_ID", deployId);
    debugLog(`✅ DEPLOY_IDを設定しました: ${sanitizeIdForLog(deployId)}`);

    const webAppUrl = `https://script.google.com/macros/s/${deployId}/exec`;
    // 基本的なURL検証のみ実施
    if (!webAppUrl.includes('script.google.com') || !webAppUrl.includes(deployId)) {
      throw new Error("生成されたウェブアプリURLが無効です。");
    }
    PropertiesService.getScriptProperties().setProperty("WEB_APP_URL", webAppUrl);
    debugLog(`✅ ウェブアプリURLを設定しました。`);

    debugLog("🎉 すべてのセットアップが正常に完了しました！");
    debugLog("---");
    debugLog("次のステップ:");
    debugLog(`1. 管理画面にアクセスして動作を確認してください: ${webAppUrl}?mode=admin`);
    debugLog(`2. 新規ユーザーとして登録テストを行ってください: ${webAppUrl}`);

  } catch (e) {
    console.error('❌ StudyQuest セットアップ中にエラーが発生しました:', e.message);
    console.error('エラーの詳細:', e.toString());
    
    // 現在のプロパティ状況を表示（デバッグ用）
    try {
      const props = PropertiesService.getScriptProperties().getProperties();
      console.log("現在のプロパティ設定:", props);
    } catch (propError) {
      console.warn("プロパティ取得エラー:", propError.message);
    }
    
    if (e.message.includes("DEPLOY_ID") || e.message.includes("デプロイ")) {
      console.log("💡 ヒント: このエラーは、ウェブアプリとして「デプロイ」を完了させる前にセットアップを実行した場合に発生します。");
      console.log("   1. Google Apps Scriptエディタで「デプロイ」→「新しいデプロイ」を選択");
      console.log("   2. 「種類の選択」でウェブアプリを選択");
      console.log("   3. デプロイ後、再度このセットアップを実行してください");
    }
    
    // エラーを再スローして呼び出し元にエラーを伝達
    throw e;
  } finally {
    lock.releaseLock();
  }
}