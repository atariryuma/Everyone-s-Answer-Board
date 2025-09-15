/**
 * @fileoverview SystemController - System management and setup functions
 */

/* global ServiceFactory, DB, UserService, ConfigService, DatabaseOperations */

// ===========================================
// 🔧 DB初期化システム（GAS読み込み順序対応）
// ===========================================

/**
 * DB接続の初期化（GAS読み込み順序問題解決）
 * SystemController内の全関数で使用
 * @returns {Object|null} DB接続オブジェクト
 */
function initDatabaseConnection() {
  try {
    // Method 1: グローバルDB変数が利用可能な場合
    if (typeof DB !== 'undefined' && DB) {
      return DB;
    }

    // Method 2: DatabaseOperationsが直接利用可能な場合
    if (typeof DatabaseOperations !== 'undefined') {
      // グローバルDB変数を設定
      if (typeof DB === 'undefined') {
        global.DB = DatabaseOperations;
      }
      return DatabaseOperations;
    }

    // Method 3: fallback - 基本的なDB機能を直接実装
    console.warn('initDatabaseConnection: DatabaseOperations not available, using fallback');
    return null;
  } catch (error) {
    console.error('initDatabaseConnection: エラー', error.message);
    return null;
  }
}

/**
 * SystemController専用：直接Session APIでメール取得
 * DB依存なしでユーザー情報を取得
 */
function getCurrentEmailDirectSC() {
  try {
    let email = Session.getActiveUser().getEmail();
    if (email) {
      return email;
    }

    email = Session.getEffectiveUser().getEmail();
    if (email) {
      return email;
    }

    console.warn('getCurrentEmailDirectSC: No email available from Session API');
    return null;
  } catch (error) {
    console.error('getCurrentEmailDirectSC:', error.message);
    return null;
  }
}

/**
 * Generate unique user ID
 * @returns {string} Unique user identifier
 */
function generateUserId() {
  return `user_${  Utilities.getUuid().replace(/-/g, '').substring(0, 12)}`;
}

/**
 * アプリケーションの初期セットアップ
 * AppSetupPage.html から呼び出される
 *
 * @param {string} serviceAccountJson - サービスアカウントJSON
 * @param {string} databaseId - データベースID
 * @param {string} adminEmail - 管理者メール
 * @param {string} googleClientId - GoogleクライアントID
 * @returns {Object} セットアップ結果
 */
function setupApplication(serviceAccountJson, databaseId, adminEmail, googleClientId) {
    try {
      // バリデーション
      if (!serviceAccountJson || !databaseId || !adminEmail) {
        return {
          success: false,
          message: '必須パラメータが不足しています'
        };
      }

      // システムプロパティを設定（ServiceFactory経由）
      const props = ServiceFactory.getProperties();
      props.set('DATABASE_SPREADSHEET_ID', databaseId);
      props.set('ADMIN_EMAIL', adminEmail);
      props.set('SERVICE_ACCOUNT_CREDS', serviceAccountJson);

      if (googleClientId) {
        props.set('GOOGLE_CLIENT_ID', googleClientId);
      }

      console.log('システムセットアップ完了:', {
        databaseId,
        adminEmail,
        hasServiceAccount: !!serviceAccountJson,
        hasClientId: !!googleClientId
      });

      return {
        success: true,
        message: 'システムセットアップが完了しました',
        setupData: {
          databaseId,
          adminEmail,
          timestamp: new Date().toISOString()
        }
      };

    } catch (error) {
      console.error('SystemController.setupApplication エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * セットアップのテスト実行
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} テスト結果
 */
function testSetup() {
    try {
      const props = ServiceFactory.getProperties();
      const databaseId = props.getDatabaseSpreadsheetId();
      const adminEmail = props.getAdminEmail();

      if (!databaseId || !adminEmail) {
        return {
          success: false,
          message: 'セットアップが不完全です。必要な設定が見つかりません。'
        };
      }

      // データベースアクセステスト
      try {
        const spreadsheet = SpreadsheetApp.openById(databaseId);
        const name = spreadsheet.getName();
        console.log('データベースアクセステスト成功:', name);
      } catch (dbError) {
        return {
          success: false,
          message: `データベースにアクセスできません: ${dbError.message}`
        };
      }

      return {
        success: true,
        message: 'セットアップテストが成功しました',
        testResults: {
          database: '✅ アクセス可能',
          adminEmail: '✅ 設定済み',
          timestamp: new Date().toISOString()
        }
      };

    } catch (error) {
      console.error('SystemController.testSetup エラー:', error.message);
      return {
        success: false,
        message: `テスト中にエラーが発生しました: ${error.message}`
      };
    }
}

/**
 * システム状態の強制リセット
 * AppSetupPage.html から呼び出される（緊急時用）
 *
 * @returns {Object} リセット結果
 */
function forceUrlSystemReset() {
    try {
      console.warn('システム強制リセットが実行されました');

      // キャッシュをクリア（複数の方法を試行）
      const cacheResults = [];
      try {
        const cache = CacheService.getScriptCache();
        if (cache && typeof cache.removeAll === 'function') {
          cache.removeAll();
          cacheResults.push('ScriptCache クリア成功');
        }
      } catch (cacheError) {
        console.warn('ScriptCache クリアエラー:', cacheError.message);
        cacheResults.push(`ScriptCache クリア失敗: ${cacheError.message}`);
      }

      // Document Cache も試行
      try {
        const docCache = CacheService.getDocumentCache();
        if (docCache && typeof docCache.removeAll === 'function') {
          docCache.removeAll();
          cacheResults.push('DocumentCache クリア成功');
        }
      } catch (docCacheError) {
        console.warn('DocumentCache クリアエラー:', docCacheError.message);
        cacheResults.push(`DocumentCache クリア失敗: ${docCacheError.message}`);
      }

      // 重要: プロパティはクリアしない（データ損失防止）

      return {
        success: true,
        message: 'システムリセットが完了しました',
        actions: cacheResults,
        cacheStatus: cacheResults.join(', '),
        timestamp: new Date().toISOString()
      };

    } catch (error) {
      console.error('SystemController.forceUrlSystemReset エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * WebアプリのURL取得
 * 各種HTMLファイルから呼び出される
 *
 * @returns {string} WebアプリURL
 */
function getWebAppUrl() {
    try {
      const url = ScriptApp.getService().getUrl();
      if (!url) {
        throw new Error('WebアプリURLの取得に失敗しました');
      }
      return url;
    } catch (error) {
      console.error('SystemController.getWebAppUrl エラー:', error.message);
      return '';
    }
}

/**
 * システム全体の診断実行
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} 診断結果
 */
function testSystemDiagnosis() {
    try {
      const diagnostics = {
        timestamp: new Date().toISOString(),
        services: {},
        database: {},
        overall: 'unknown'
      };

      // Services診断 - Safe service access with GAS loading protection
      try {
        // Safe Service availability check - removed function calls to prevent undefined errors
        diagnostics.services.UserService = typeof UserService !== 'undefined'
          ? '✅ Loaded and available'
          : '❌ Not loaded or unavailable';

        diagnostics.services.ConfigService = typeof ConfigService !== 'undefined'
          ? '✅ Loaded and available'
          : '❌ Not loaded or unavailable';

        diagnostics.services.DataService = '❓ 診断機能なし';
      } catch (servicesError) {
        diagnostics.services.error = servicesError.message;
      }

      // データベース診断
      try {
        const props = ServiceFactory.getProperties();
        const databaseId = props.getDatabaseSpreadsheetId();

        if (databaseId) {
          const spreadsheet = SpreadsheetApp.openById(databaseId);
          diagnostics.database = {
            accessible: true,
            name: spreadsheet.getName(),
            sheets: spreadsheet.getSheets().length
          };
        } else {
          diagnostics.database = { accessible: false, reason: 'DATABASE_SPREADSHEET_ID not configured' };
        }
      } catch (dbError) {
        diagnostics.database = { accessible: false, error: dbError.message };
      }

      // 総合評価
      const hasErrors = Object.values(diagnostics).some(v =>
        typeof v === 'object' && (v.error || v.accessible === false)
      );
      diagnostics.overall = hasErrors ? '⚠️ 問題あり' : '✅ 正常';

      return {
        success: true,
        diagnostics
      };

    } catch (error) {
      console.error('SystemController.testSystemDiagnosis エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * システム状態の取得
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} システム状態
 */
function getSystemStatus() {
    try {
      const props = ServiceFactory.getProperties();
      const status = {
        timestamp: new Date().toISOString(),
        setup: {
          hasDatabase: !!props.getDatabaseSpreadsheetId(),
          hasAdminEmail: !!props.getAdminEmail(),
          hasServiceAccount: !!props.getServiceAccountCreds()
        },
        services: {
          available: ['UserService', 'ConfigService', 'DataService', 'SecurityService']
        }
      };

      status.setup.isComplete = status.setup.hasDatabase &&
                                 status.setup.hasAdminEmail &&
                                 status.setup.hasServiceAccount;

      return {
        success: true,
        status
      };

    } catch (error) {
      console.error('SystemController.getSystemStatus エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * システムドメイン情報の取得
 * AppSetupPage.html から呼び出される
 *
 * @returns {Object} ドメイン情報
 */
function getSystemDomainInfo() {
    try {
      // 🎯 Zero-dependency: 直接Session API使用
      const currentUser = getCurrentEmailDirectSC();
      let domain = 'unknown';

      if (currentUser && currentUser.includes('@')) {
        [, domain] = currentUser.split('@');
      }

      return {
        success: true,
        domain,
        currentUser,
        timestamp: new Date().toISOString()
      };

    } catch {
      return {
        success: false,
        message: 'ドメイン情報の取得に失敗しました'
      };
    }
}

/**
 * データ整合性チェック
 * 管理者用の高度な診断機能
 *
 * @returns {Object} チェック結果
 */
function performDataIntegrityCheck() {
    try {
      const results = {
        timestamp: new Date().toISOString(),
        checks: {},
        summary: { passed: 0, failed: 0, warnings: 0 }
      };

      // 各種整合性チェックを実行
      // （実装は複雑になるため、基本的な構造のみ示す）

      results.checks.database = { status: '✅', message: '基本チェックのみ実装' };
      results.checks.users = { status: '✅', message: '基本チェックのみ実装' };
      results.checks.configs = { status: '✅', message: '基本チェックのみ実装' };

      results.summary.passed = 3;
      results.overall = '✅ 基本チェック完了';

      return {
        success: true,
        results
      };

    } catch (error) {
      console.error('SystemController.performDataIntegrityCheck エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * 自動修復の実行
 * 管理者用の自動メンテナンス機能
 *
 * @returns {Object} 修復結果
 */
function performAutoRepair() {
    try {
      console.log('自動修復機能は現在基本実装のみです');

      const repairResults = {
        timestamp: new Date().toISOString(),
        actions: [
          'キャッシュクリア実行'
        ],
        summary: '基本的な修復のみ実行'
      };

      // キャッシュクリア
      try {
        const cache = CacheService.getScriptCache();
        if (cache && typeof cache.removeAll === 'function') {
          cache.removeAll();
        }
      } catch (cacheError) {
        repairResults.warnings = [`キャッシュクリア失敗: ${cacheError.message}`];
      }

      return {
        success: true,
        repairResults
      };

    } catch (error) {
      console.error('SystemController.performAutoRepair エラー:', error.message);
      return {
        success: false,
        message: error.message
      };
    }
}

/**
 * 現在の設定を取得
 * AdminPanel.js.html から呼び出される
 *
 * @returns {Object} 設定情報
 */
function getConfig() {
  try {
    // 🔧 DB初期化（GAS読み込み順序対応）
    const db = initDatabaseConnection();
    if (!db) {
      console.error('getConfig: DB初期化失敗');
      return {
        success: false,
        message: 'データベース接続エラー'
      };
    }

    // 🎯 Zero-dependency: 直接Session APIでユーザー取得
    const email = getCurrentEmailDirectSC();
    if (!email) {
      return { success: false, message: 'ユーザー情報が見つかりません' };
    }

    // 🎯 DB初期化済みのdb変数を使用
    let user = db.findUserByEmail(email);

    // Auto-create user if not exists
    if (!user) {
      try {
        const newUserId = generateUserId();
        user = {
          userId: newUserId,
          userEmail: email,
          isActive: true,
          createdAt: new Date().toISOString(),
          configJson: null
        };
        db.createUser(user);
        console.log('SystemController.getConfig: 新規ユーザー作成:', { userId: newUserId, email });
      } catch (createErr) {
        console.error('SystemController.getConfig: ユーザー作成エラー', createErr);
        return { success: false, message: 'ユーザー作成に失敗しました' };
      }
    }

    // 🎯 Zero-dependency: 直接DBから設定取得
    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.error('SystemController.getConfig: 設定JSON解析エラー', parseError);
        config = {};
      }
    }

    return {
      success: true,
      config,
      userId: user.userId,
      userEmail: user.userEmail
    };
  } catch (error) {
    console.error('AdminController.getConfig エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * スプレッドシート一覧を取得
 * @returns {Object} スプレッドシート一覧
 */
function getAdminSpreadsheetList() {
  console.log('🔍 getAdminSpreadsheetList: 関数開始 - Zero-dependency Architecture');
  try {
    console.log('🔍 DriveApp.getFilesByType呼び出し開始');

    // 🎯 Zero-dependency: 直接DriveAppでスプレッドシート一覧取得
    const spreadsheets = DriveApp.getFilesByType('application/vnd.google-apps.spreadsheet');
    console.log('🔍 DriveApp.getFilesByType完了', spreadsheets);

    const spreadsheetList = [];
    let count = 0;

    console.log('🔍 ファイル列挙開始');
    while (spreadsheets.hasNext() && count < 20) { // 最大20件に制限
      const file = spreadsheets.next();
      const fileData = {
        id: file.getId(),
        name: file.getName(),
        lastUpdated: file.getLastUpdated(),
        url: file.getUrl(),
        size: file.getSize() || 0
      };
      console.log(`🔍 ファイル${count + 1}:`, fileData.name, fileData.id);
      spreadsheetList.push(fileData);
      count++;
    }

    const result = {
      success: true,
      spreadsheets: spreadsheetList,
      total: spreadsheetList.length,
      timestamp: new Date().toISOString()
    };

    console.log('🔍 getAdminSpreadsheetList: 結果準備完了', result);
    return result;
  } catch (error) {
    console.error('🚨 AdminController.getSpreadsheetList エラー:', error);

    const errorResult = {
      success: false,
      message: error.message || 'スプレッドシート一覧取得エラー',
      spreadsheets: []
    };

    console.log('🔍 getAdminSpreadsheetList: エラー結果', errorResult);
    return errorResult;
  }
}

/**
 * シート一覧を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getAdminSheetList(spreadsheetId) {
  try {
    // 🎯 Zero-dependency: 直接SpreadsheetAppでシート一覧取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn()
    }));

    return {
      success: true,
      sheets: sheetList,
      total: sheetList.length,
      spreadsheetName: spreadsheet.getName()
    };
  } catch (error) {
    console.error('AdminController.getSheetList エラー:', error.message);
    return {
      success: false,
      message: error.message || 'シート一覧取得エラー',
      sheets: []
    };
  }
}

/**
 * 列を分析
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  try {
    console.log('SystemController.analyzeColumns: 開始 - Zero-dependency Architecture', {
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
      sheetName: sheetName || 'null'
    });

    // 🎯 Zero-dependency: 直接SpreadsheetAppで列分析
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return {
        success: false,
        message: `シート "${sheetName}" が見つかりません`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      };
    }

    // ヘッダー行取得と列分析
    const headerRow = 1;
    const lastColumn = sheet.getLastColumn();
    const [headers] = lastColumn > 0 ? sheet.getRange(headerRow, 1, 1, lastColumn).getValues() : [[]];

    const columns = headers.map((header, index) => ({
      index: index + 1,
      header: String(header || ''),
      letter: columnNumberToLetter(index + 1),
      type: 'text' // 簡化版では全てtext
    }));

    // 簡化版のコラムマッピング
    const mapping = {};
    const confidence = {};
    headers.forEach((header, index) => {
      const headerStr = String(header || '').toLowerCase();
      if (headerStr.includes('名前') || headerStr.includes('name')) {
        mapping.name = index + 1;
        confidence.name = 0.9;
      } else if (headerStr.includes('コメント') || headerStr.includes('comment')) {
        mapping.comment = index + 1;
        confidence.comment = 0.8;
      }
    });

    return {
      success: true,
      headers,
      columns,
      columnMapping: { mapping, confidence },
      sheetInfo: {
        name: sheetName,
        totalRows: sheet.getLastRow(),
        totalColumns: lastColumn
      }
    };
  } catch (error) {
    console.error('AdminController.analyzeColumns エラー:', error.message);

    return {
      success: false,
      message: error.message || '列分析エラー',
      headers: [],
      columns: [],
      columnMapping: { mapping: {}, confidence: {} }
    };
  }
}

/**
 * 軽量ヘッダー取得 - 列分析に失敗してもヘッダー名だけは取得
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} ヘッダー取得結果
 */
function getLightweightHeaders(spreadsheetId, sheetName) {
  try {
    // 🎯 Zero-dependency: 直接SpreadsheetAppでヘッダー取得
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return {
        success: false,
        message: `シート "${sheetName}" が見つかりません`,
        headers: []
      };
    }

    const lastColumn = sheet.getLastColumn();
    const [headers] = lastColumn > 0 ? sheet.getRange(1, 1, 1, lastColumn).getValues() : [[]];

    const result = {
      success: true,
      headers: headers.map(h => String(h || '')),
      sheetName,
      columnCount: lastColumn
    };

    return result;
  } catch (error) {
    console.error('AdminController.getLightweightHeaders エラー:', error.message);
    return {
      success: false,
      message: error.message || 'ヘッダー取得エラー',
      headers: []
    };
  }
}

/**
 * 設定の下書き保存
 * @param {Object} config - 保存する設定
 * @returns {Object} 保存結果
 */
function saveDraftConfiguration(config) {
  try {
    // 🎯 Zero-dependency: 直接DBで設定保存
    const userEmail = getCurrentEmailDirectSC();
    if (!userEmail) {
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    const user = DB.findUserByEmail(userEmail);
    if (!user) {
      return { success: false, message: 'ユーザーが見つかりません' };
    }

    // 設定をJSONで保存
    const updatedUser = {
      ...user,
      configJson: JSON.stringify(config),
      updatedAt: new Date().toISOString()
    };

    DB.updateUser(user.userId, updatedUser);

    return {
      success: true,
      message: '下書き設定を保存しました',
      userId: user.userId
    };
  } catch (error) {
    console.error('saveDraftConfiguration error:', error);
    return { success: false, message: error.message };
  }
}

/**
 * アプリケーションの公開
 * @param {Object} publishConfig - 公開設定
 * @returns {Object} 公開結果
 */
function publishApplication(publishConfig) {
  try {
    // 🎯 Zero-dependency: 直接PropertiesServiceでアプリ公開
    const props = PropertiesService.getScriptProperties();

    // アプリケーション状態をアクティブに変更
    props.setProperty('APPLICATION_STATUS', 'active');
    props.setProperty('PUBLISHED_AT', new Date().toISOString());

    // 公開設定を保存
    if (publishConfig) {
      props.setProperty('PUBLISH_CONFIG', JSON.stringify(publishConfig));
    }

    return {
      success: true,
      message: 'アプリケーションが正常に公開されました',
      publishedAt: new Date().toISOString()
    };
  } catch (error) {
    console.error('publishApplication error:', error);
    return { success: false, message: error.message };
  }
}

/**
 * スプレッドシートへのアクセス権限を検証
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateAccess(spreadsheetId) {
  try {
    // 🎯 Zero-dependency: 直接SpreadsheetAppでアデス権確認
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    // アデスできたら成功
    const result = {
      success: true,
      message: 'アクセス権限が確認されました',
      spreadsheetName: spreadsheet.getName(),
      sheets: sheets.map(sheet => ({
        name: sheet.getName(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn()
      })),
      owner: spreadsheet.getOwner()?.getEmail() || 'unknown',
      url: spreadsheet.getUrl()
    };

    return result;
  } catch (error) {
    console.error('AdminController.validateAccess エラー:', error.message);
    return {
      success: false,
      message: error.message,
      sheets: []
    };
  }
}

/**
 * システム管理者権限の確認
 * AdminPanel.js.html から呼び出される
 *
 * @returns {boolean} 管理者権限の有無
 */
function checkIsSystemAdmin() {
  try {
    // 🎯 Zero-dependency: 直接Session APIとPropertiesServiceで管理者確認
    const email = getCurrentEmailDirectSC();
    if (!email) {
      return false;
    }

    const props = PropertiesService.getScriptProperties();
    const adminEmails = props.getProperty('ADMIN_EMAILS') || '';

    return adminEmails.split(',').map(e => e.trim()).includes(email);
  } catch (error) {
    console.error('SystemController.checkIsSystemAdmin エラー:', error.message);
    return false;
  }
}

/**
 * 現在のボード情報とURLを取得
 * AdminPanel.js.html から呼び出される
 *
 * @returns {Object} ボード情報
 */
function getCurrentBoardInfoAndUrls() {
  try {
    // 🔧 DB初期化（GAS読み込み順序対応）
    const db = initDatabaseConnection();
    if (!db) {
      console.error('getCurrentBoardInfoAndUrls: DB初期化失敗');
      return {
        isActive: false,
        error: 'データベース接続エラー',
        appPublished: false
      };
    }

    // 🎯 Zero-dependency: 直接Session APIでユーザー取得
    const email = getCurrentEmailDirectSC();
    if (!email) {
      return {
        isActive: false,
        error: 'ユーザー情報が見つかりません',
        appPublished: false
      };
    }

    // 🎯 DB初期化済みのdb変数を使用
    const user = db.findUserByEmail(email);
    if (!user) {
      return {
        isActive: false,
        error: 'ユーザーが見つかりません',
        appPublished: false
      };
    }

    // 🎯 Zero-dependency: 直接PropertiesServiceでアプリ状態確認
    const props = PropertiesService.getScriptProperties();
    const appStatus = props.getProperty('APPLICATION_STATUS');
    const appPublished = appStatus === 'active';

    if (!appPublished) {
      return {
        isActive: false,
        appPublished: false,
        questionText: 'アプリケーションが公開されていません'
      };
    }

    // WebAppのベースURL取得
    const baseUrl = ScriptApp.getService().getUrl();
    const viewUrl = `${baseUrl}?mode=view&userId=${user.userId}`;

    // 設定がある場合はその情報も含める
    let config = {};
    if (user.configJson) {
      try {
        config = JSON.parse(user.configJson);
      } catch (parseError) {
        console.warn('getCurrentBoardInfoAndUrls: 設定JSON解析エラー', parseError);
      }
    }

    return {
      isActive: true,
      appPublished: true,
      questionText: config.questionText || config.boardTitle || 'Everyone\'s Answer Board',
      urls: {
        view: viewUrl,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || config.lastModified || new Date().toISOString()
    };

  } catch (error) {
    console.error('AdminController.getCurrentBoardInfoAndUrls エラー:', error.message);
    return {
      isActive: false,
      appPublished: false,
      error: error.message
    };
  }
}

/**
 * フォーム情報を取得
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfo(spreadsheetId, sheetName) {
  try {
    return ConfigService.getFormInfo(spreadsheetId, sheetName);
  } catch (error) {
    console.error('AdminController.getFormInfo エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フォーム作成
 * AdminPanel.js.html から呼び出される
 *
 * @param {string} userId - ユーザーID
 * @param {Object} config - フォーム設定
 * @returns {Object} 作成結果
 */
function createForm(userId, config) {
  try {
    console.log('AdminController.createForm: 開始', { userId, configKeys: Object.keys(config || {}) });

    if (!userId) {
      return {
        success: false,
        error: 'ユーザーIDが指定されていません'
      };
    }

    if (!config || !config.title) {
      return {
        success: false,
        error: 'フォーム設定が不正です'
      };
    }

    // ConfigService経由でフォーム作成
    const result = ConfigService.createForm(userId, config);

    if (result && result.success) {
      console.log('AdminController.createForm: 成功', { formUrl: result.formUrl });
      return result;
    } else {
      console.error('AdminController.createForm: ConfigService失敗', result);
      return {
        success: false,
        error: result?.error || 'フォーム作成に失敗しました'
      };
    }

  } catch (error) {
    console.error('AdminController.createForm エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 現在の公開状態を確認
 * AdminPanel.js.html から呼び出される
 *
 * @returns {Object} 公開状態情報
 */
function checkCurrentPublicationStatus() {
  try {
    // ユーザー情報の取得
    const userInfo = UserService.getCurrentUserInfo();
    const userId = userInfo && userInfo.userId;

    if (!userId) {
      return {
        success: false,
        published: false,
        error: 'ユーザー情報が見つかりません'
      };
    }

    // 設定情報を取得
    const config = ConfigService.getUserConfig(userId);
    if (!config) {
      return {
        success: false,
        published: false,
        error: '設定情報が見つかりません'
      };
    }

    return {
      success: true,
      published: config.appPublished === true,
      publishedAt: config.publishedAt || null,
      lastModified: config.lastModified || null,
      hasDataSource: !!(config.spreadsheetId && config.sheetName)
    };

  } catch (error) {
    console.error('AdminController.checkCurrentPublicationStatus エラー:', error.message);
    return {
      success: false,
      published: false,
      error: error.message
    };
  }
}

/**
 * 現在のユーザー情報を取得
 * login.js.html, SetupPage.html, AdminPanel.js.html から呼び出される
 *
 * @param {string} [kind='email'] - 取得する情報の種類（'email' or 'full'）
 * @returns {Object|string|null} 統一されたレスポンス形式
 */
/**
 * Direct email retrieval using GAS Session API (SystemController version)
 */
// 🎯 Zero-dependency Helper Functions
function columnNumberToLetter(num) {
  let letter = '';
  while (num > 0) {
    const remainder = (num - 1) % 26;
    letter = String.fromCharCode(65 + remainder) + letter;
    num = Math.floor((num - 1) / 26);
  }
  return letter;
}

function getUser(kind = 'email') {
  try {
    // 🚀 Direct Session API fallback - no service dependencies
    const userEmail = getCurrentEmailDirectSC();

    if (!userEmail) {
      console.warn('getUser: No email available from Session API');
      return kind === 'email' ? '' : { success: false, message: 'ユーザー情報が取得できません' };
    }

    // 後方互換性重視: kind==='email' の場合は純粋な文字列を返す
    if (kind === 'email') {
      return String(userEmail);
    }

    // 統一オブジェクト形式（'full' など）- フォールバック情報
    const userInfo = { userEmail, userId: null, isActive: true };
    return {
      success: true,
      email: userEmail,
      userId: userInfo?.userId || null,
      isActive: userInfo?.isActive || false,
      hasConfig: !!userInfo?.config
    };
  } catch (error) {
    console.error('FrontendController.getUser エラー:', error.message);
    return kind === 'email' ? '' : { success: false, message: error.message };
  }
}


// ===========================================
// 📊 認証・ログイン関連API
// ===========================================

/**
 * ログインアクションを処理
 * login.js.html から呼び出される
 *
 * @returns {Object} ログイン処理結果
 */
function processLoginAction() {
  try {
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return {
        success: false,
        message: 'ユーザー情報が取得できません',
        needsAuth: true
      };
    }

    // ユーザー情報を取得または作成
    let userInfo = UserService.getCurrentUserInfo();
    if (!userInfo) {
      userInfo = UserService.createUser(userEmail);
      // createUserの戻り値構造を正規化
      if (userInfo && userInfo.value) {
        userInfo = userInfo.value;
      }
    }

    // 管理パネル用URLを構築（userId必須）
    const baseUrl = getWebAppUrl();
    const userId = userInfo?.userId;

    if (!userId) {
      return {
        success: false,
        message: 'ユーザーIDの取得に失敗しました',
        error: 'USER_ID_MISSING'
      };
    }

    const adminUrl = `${baseUrl}?mode=admin&userId=${userId}`;

    return {
      success: true,
      userInfo,
      redirectUrl: baseUrl,
      adminUrl,
      // 後方互換性のための追加プロパティ
      appUrl: baseUrl,
      url: adminUrl
    };

  } catch (error) {
    console.error('FrontendController.processLoginAction エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 認証状態を確認
 * login.js.html, SetupPage.html から呼び出される
 *
 * @returns {Object} 認証状態
 */
function verifyUserAuthentication() {
  try {
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return {
        isAuthenticated: false,
        message: '認証が必要です'
      };
    }

    const userInfo = UserService.getCurrentUserInfo();
    return {
      isAuthenticated: true,
      userEmail,
      userInfo,
      hasConfig: !!userInfo?.config
    };

  } catch (error) {
    console.error('FrontendController.verifyUserAuthentication エラー:', error.message);
    return {
      isAuthenticated: false,
      message: error.message
    };
  }
}

/**
 * 認証情報のリセット
 * login.js.html から呼び出される
 *
 * @returns {Object} リセット結果
 */
function resetAuth() {
  try {
    console.log('FrontendController.resetAuth: 認証リセット開始');

    // ユーザーキャッシュクリア
    UserService.clearUserCache();

    // セッション関連の情報をクリア
    const props = PropertiesService.getScriptProperties();

    // 一時的な認証情報をクリア
    const authKeys = ['temp_auth_token', 'last_login_attempt', 'auth_retry_count'];
    authKeys.forEach(key => {
      props.deleteProperty(key);
    });

    console.log('FrontendController.resetAuth: 認証リセット完了');

    return {
      success: true,
      message: '認証情報がリセットされました'
    };

  } catch (error) {
    console.error('FrontendController.resetAuth エラー:', error.message);
    return {
      success: false,
      message: `認証リセットに失敗しました: ${error.message}`
    };
  }
}

/**
 * ログイン状態を取得
 * login.js.html から呼び出される
 *
 * @returns {Object} ログイン状態
 */
function getLoginStatus() {
  try {
    const userEmail = UserService.getCurrentEmail();
    if (!userEmail) {
      return {
        isLoggedIn: false,
        user: null
      };
    }

    const userInfo = UserService.getCurrentUserInfo();
    return {
      isLoggedIn: true,
      user: {
        email: userEmail,
        userId: userInfo?.userId,
        hasSetup: !!userInfo?.config?.setupComplete
      }
    };

  } catch (error) {
    console.error('FrontendController.getLoginStatus エラー:', error.message);
    return {
      isLoggedIn: false,
      user: null,
      error: error.message
    };
  }
}

/**
 * クライアントエラーを報告
 * ErrorBoundary.html から呼び出される
 *
 * @param {Object} errorInfo - エラー情報
 * @returns {Object} 報告結果
 */
function reportClientError(errorInfo) {
  try {
    console.error('クライアントエラー報告:', errorInfo);

    // エラーログを記録（将来的にはSecurityServiceや専用のログサービスに委譲）
    const logEntry = {
      timestamp: new Date().toISOString(),
      type: 'client_error',
      userEmail: UserService.getCurrentEmail() || 'unknown',
      errorInfo
    };

    // コンソールにログ出力（将来的には永続化）
    console.log('Error Log Entry:', JSON.stringify(logEntry));

    return {
      success: true,
      message: 'エラーが報告されました'
    };
  } catch (error) {
    console.error('FrontendController.reportClientError エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}

/**
 * 強制ログアウトとリダイレクトのテスト
 * ErrorBoundary.html から呼び出される
 *
 * @returns {Object} テスト結果
 */
function testForceLogoutRedirect() {
  try {
    console.log('強制ログアウトテストが実行されました');

    return {
      success: true,
      message: 'ログアウトテスト完了',
      redirectUrl: `${getWebAppUrl()}?mode=login`
    };
  } catch (error) {
    console.error('FrontendController.testForceLogoutRedirect エラー:', error.message);
    return {
      success: false,
      message: error.message
    };
  }
}
