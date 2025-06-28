/**
 * @fileoverview StudyQuest -みんなの回答ボード- - 統合セットアップスクリプト
 * フォルダ管理、データベース作成、共有設定、プロパティ設定を自動で行う統合セットアップ機能
 *
 * NOTE: このスクリプトは、Code.gsで定義されているグローバル関数と定数に依存しています。
 * (例: USER_DB_CONFIG, debugLog, sanitizeIdForLog, validateDeployId, validateWebAppUrl, secureLogError)。
 * Code.gsが同じApps Scriptプロジェクト内にあり、その関数がグローバルスコープで定義されていることを確認してください。
 */

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
        sheet.appendRow(USER_DB_CONFIG.HEADERS); // Use USER_DB_CONFIG.HEADERS
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
    const deployId = manualDeployId || ScriptApp.getDeploymentId();
    if (!deployId || !validateDeployId(deployId)) {
      throw new Error("DEPLOY_IDの取得または検証に失敗しました。先にウェブアプリとしてデプロイを完了させる必要があります。");
    }
    PropertiesService.getScriptProperties().setProperty("DEPLOY_ID", deployId);
    debugLog(`✅ DEPLOY_IDを設定しました: ${sanitizeIdForLog(deployId)}`);

    const webAppUrl = `https://script.google.com/macros/s/${deployId}/exec`;
    if (!validateWebAppUrl(webAppUrl)) {
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
    secureLogError('StudyQuest セットアップ', e);
    debugLog("現在のプロパティ設定:", PropertiesService.getScriptProperties().getProperties());
    if (e.message.includes("DEPLOY_ID")) {
      debugLog("💡 ヒント: このエラーは、ウェブアプリとして「デプロイ」を完了させる前にセットアップを実行した場合に発生します。先にデプロイを完了させてから再度実行してください。");
    }
  } finally {
    lock.releaseLock();
  }
}
