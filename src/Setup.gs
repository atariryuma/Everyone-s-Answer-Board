/**
 * @fileoverview StudyQuest -みんなの回答ボード- - 統合セットアップスクリプト
 * フォルダ管理、データベース作成、共有設定、プロパティ設定を自動で行う。
 */

/**
 * StudyQuestの統合セットアップ関数 (エラー報告改善版)
 * フォルダ作成、データベース作成、ファイル移動、プロパティ設定を自動で行います。
 * @param {string} [manualDeployId] - 手動で設定するデプロイID（任意）
 */
function studyQuestSetup(manualDeployId) {
  // セキュリティ強化: 認証チェック
  const currentUser = Session.getActiveUser().getEmail();
  if (!currentUser) {
    throw new Error('認証が必要です。Googleアカウントでログインしてください。');
  }
  
  const FOLDER_NAME = "StudyQuest - みんなの回答ボード"; // アプリ専用フォルダ名
  const DB_FILENAME = "StudyQuest_UserDatabase";     // データベースファイル名
  const lock = LockService.getScriptLock();
  lock.waitLock(30000); // 30秒待機

  try {
    console.log("🚀 StudyQuest 統合セットアップを開始します...");

    // ステップ1: 専用フォルダの確認と作成
    let folder;
    const folders = DriveApp.getFoldersByName(FOLDER_NAME);
    if (folders.hasNext()) {
      folder = folders.next();
      console.log(`✅ 専用フォルダ「${FOLDER_NAME}」が既に存在します。`);
    } else {
      folder = DriveApp.createFolder(FOLDER_NAME);
      console.log(`✅ 専用フォルダ「${FOLDER_NAME}」を新規作成しました。`);
    }

    // ステップ2: データベース（スプレッドシート）の確認と作成
    let dbFile;
    const files = folder.getFilesByName(DB_FILENAME);
    if (files.hasNext()) {
      dbFile = files.next();
      console.log(`✅ データベースファイルがフォルダ内に既に存在します。`);
    } else {
      // マイドライブ直下も検索して、あれば移動させる
      const rootFiles = DriveApp.getRootFolder().getFilesByName(DB_FILENAME);
      if (rootFiles.hasNext()) {
        dbFile = rootFiles.next();
        dbFile.moveTo(folder);
        console.log(`✅ 既存のデータベースファイルを専用フォルダに移動しました。`);
      } else {
        // どこにもない場合のみ新規作成
        const newDb = SpreadsheetApp.create(DB_FILENAME);
        dbFile = DriveApp.getFileById(newDb.getId());
        dbFile.moveTo(folder);
        const sheet = newDb.getSheets()[0];
        sheet.setName("users");
        sheet.appendRow(["userId", "adminEmail", "spreadsheetId", "spreadsheetUrl", "createdAt", "accessToken", "configJson", "lastAccessedAt", "isActive"]);
        console.log(`✅ データベースファイルを新規作成し、専用フォルダに配置しました。`);
      }
    }
    const dbId = dbFile.getId();
    PropertiesService.getScriptProperties().setProperty("DATABASE_ID", dbId);
    console.log(`✅ データベースIDを設定しました: ${sanitizeIdForLog(dbId)}`);

    // ステップ2.5: データベースの共有設定（セキュリティ強化）
    try {
      const adminEmail = Session.getActiveUser().getEmail();
      dbFile.addEditor(adminEmail);
      console.log(`✅ 管理者に編集権限を付与しました。`);
      console.log(`ℹ️ セキュリティのため、ドメイン共有は必要に応じて手動で設定してください。`);
      
    } catch (e) {
      console.warn('⚠️ 共有設定の一部に失敗しました。手動で設定してください。');
    }


    // ステップ3: デプロイIDとウェブアプリURLの設定
    const deployId = manualDeployId || ScriptApp.getDeploymentId();
    if (!deployId || !validateDeployId(deployId)) {
      console.error("❌ DEPLOY_IDの取得または検証に失敗しました。");
      throw new Error("DEPLOY_IDの取得に失敗しました。先にウェブアプリとしてデプロイを完了させる必要があります。");
    }
    PropertiesService.getScriptProperties().setProperty("DEPLOY_ID", deployId);
    console.log(`✅ DEPLOY_IDを設定しました: ${sanitizeIdForLog(deployId)}`);

    const webAppUrl = `https://script.google.com/macros/s/${deployId}/exec`;
    if (!validateWebAppUrl(webAppUrl)) {
      throw new Error("生成されたウェブアプリURLが無効です。");
    }
    PropertiesService.getScriptProperties().setProperty("WEB_APP_URL", webAppUrl);
    console.log(`✅ ウェブアプリURLを設定しました。`);

    console.log("🎉 すべてのセットアップが正常に完了しました！");
    console.log("---");
    console.log("次のステップ:");
    console.log(`1. 管理画面にアクセスして動作を確認してください: ${webAppUrl}?mode=admin`);
    console.log(`2. 新規ユーザーとして登録テストを行ってください: ${webAppUrl}`);

  } catch (e) {
    // 【改善点】エラー内容を具体的に表示するように変更
    console.error(`❌ セットアップ中にエラーが発生しました: ${e.message}`);
    console.error(`スタックトレース: ${e.stack}`);
    // 失敗した場合でも、取得できた情報はログに出力
    const props = PropertiesService.getScriptProperties().getProperties();
    console.log("現在のプロパティ設定:", props);
    if (e.message.includes("DEPLOY_ID")) {
      console.log("💡 ヒント: このエラーは、ウェブアプリとして「デプロイ」を完了させる前にセットアップを実行した場合に発生します。先にデプロイを完了させてから再度実行してください。");
    }
  } finally {
    lock.releaseLock();
  }
}

// =================================================================
// セキュリティユーティリティ関数
// =================================================================

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
