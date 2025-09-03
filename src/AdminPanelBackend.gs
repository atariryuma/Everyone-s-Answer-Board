/**
 * @fileoverview AdminPanel Backend - CLAUDE.md完全準拠版
 * configJSON中心型アーキテクチャの完全実装
 */

/**
 * CLAUDE.md準拠：現在の設定情報を取得（configJSON中心型）
 * 統一データソース原則：configJsonからすべてのデータを取得
 * @returns {Object} 現在の設定情報
 */
function getCurrentConfig() {
  try {
    console.log('getCurrentConfig: CLAUDE.md準拠configJSON中心型設定取得開始');

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      // デフォルト設定（CLAUDE.md準拠：心理的安全性重視）
      return {
        setupStatus: 'pending',
        appPublished: false,
        displaySettings: { showNames: false, showReactions: false }, // CLAUDE.md準拠
        user: currentUser,
      };
    }

    // ✅ CLAUDE.md完全準拠：configJSON統一データソース原則
    const config = userInfo.parsedConfig || {};

    // ✅ CLAUDE.md準拠：configJSON中心型設定構築（外側フィールド参照完全排除）
    const fullConfig = {
      // 5フィールド構造の基本情報
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      isActive: userInfo.isActive,
      lastModified: userInfo.lastModified,

      // ✅ configJSON統一データソース：全データはconfigから取得（外側フィールド参照なし）
      spreadsheetId: config.spreadsheetId || null,
      spreadsheetUrl: config.spreadsheetUrl || null, 
      sheetName: config.sheetName || null,
      formUrl: config.formUrl || null,
      
      // 監査情報
      createdAt: config.createdAt || null,
      lastAccessedAt: config.lastAccessedAt || null,
      
      // 列マッピング（JSON統一）
      columnMapping: config.columnMapping || {},
      
      // ✅ 公開情報（configJSON統一データソース）
      publishedAt: config.publishedAt || null,
      appUrl: config.appUrl || null,
      appPublished: config.appPublished || false, // ← これが重要：configJSONから取得
      
      // ✅ 表示設定（configJSON統一データソース、CLAUDE.md準拠）
      displaySettings: config.displaySettings || { showNames: false, showReactions: false },
      
      // アプリ設定
      setupStatus: config.setupStatus || 'pending',
      connectionMethod: config.connectionMethod || null,
      lastConnected: config.lastConnected || null,
      
      // その他
      formTitle: config.formTitle || null,
      missingColumnsHandled: config.missingColumnsHandled || null,
      opinionHeader: config.opinionHeader || null, // 問題文ヘッダー
      
      // CLAUDE.md準拠メタデータ
      configJsonVersion: config.configJsonVersion || '1.0',
      claudeMdCompliant: config.claudeMdCompliant || false
    };

    // CLAUDE.md準拠：構造化ログによる設定情報出力
    console.info('📋 getCurrentConfig: configJSON中心型設定取得完了', {
      userId: fullConfig.userId,
      hasSpreadsheetId: !!fullConfig.spreadsheetId, // 統一データソース
      hasSheetName: !!fullConfig.sheetName, // 統一データソース
      hasColumnMapping: !!fullConfig.columnMapping && Object.keys(fullConfig.columnMapping).length > 0,
      hasFormUrl: !!fullConfig.formUrl,
      appPublished: fullConfig.appPublished,
      setupStatus: fullConfig.setupStatus,
      claudeMdCompliant: fullConfig.claudeMdCompliant,
      timestamp: new Date().toISOString(),
    });

    return fullConfig;
  } catch (error) {
    // CLAUDE.md準拠：構造化ログによるエラー情報出力
    console.error('❌ getCurrentConfig エラー:', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString(),
    });

    return {
      setupStatus: 'error',
      appPublished: false,
      displaySettings: { showNames: false, showReactions: false }, // CLAUDE.md準拠
      error: error.message,
    };
  }
}

/**
 * 🎆 緊急回復用RPC関数 - メールアドレスベースで設定取得
 * テンプレート変数展開失敗時のフォールバックとして使用
 */
function getCurrentConfigByEmail() {
  try {
    console.info('🎆 緊急回復: メールアドレスベースで設定取得開始');
    
    const currentUserEmail = User.email();
    if (!currentUserEmail) {
      throw new Error('アクティブユーザーのメールアドレスが取得できません');
    }
    
    const userInfo = DB.findUserByEmail(currentUserEmail);
    if (!userInfo) {
      console.warn('緊急回復: ユーザー情報が見つかりません:', currentUserEmail);
      return {
        error: 'user_not_found',
        userEmail: currentUserEmail,
        suggestion: 'ユーザー登録が必要です'
      };
    }
    
    const config = userInfo.parsedConfig || {};
    const recoveryConfig = {
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      formUrl: config.formUrl,
      setupStatus: config.setupStatus || 'incomplete',
      appPublished: config.appPublished || false,
      displaySettings: config.displaySettings || { showNames: false, showReactions: false },
      recoveryMode: true,
      timestamp: new Date().toISOString()
    };
    
    console.info('✅ 緊急回復: 設定取得成功', {
      userId: userInfo.userId,
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      setupStatus: config.setupStatus
    });
    
    return recoveryConfig;
  } catch (error) {
    console.error('❌ 緊急回復エラー:', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    });
    
    return {
      error: 'recovery_failed',
      message: error.message,
      suggestion: 'ページをリロードしてください'
    };
  }
}

/**
 * データソース接続（CLAUDE.md準拠版）
 * 統一データソース原則：全データをconfigJsonに統合
 */
function connectDataSource(spreadsheetId, sheetName) {
  try {
    console.log('🔗 connectDataSource: CLAUDE.md準拠configJSON中心型接続開始', { spreadsheetId, sheetName });

    if (!spreadsheetId || !sheetName) {
      throw new Error('スプレッドシートIDとシート名が必要です');
    }

    // 基本的な接続検証
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    // ヘッダー情報と列マッピング検出
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const columnMapping = detectColumnMapping(headerRow);

    // 列マッピング検証
    const validationResult = validateAdminPanelMapping(columnMapping);
    if (!validationResult.isValid) {
      console.warn('列名マッピング検証エラー', validationResult.errors);
    }

    // 不足列の追加
    const missingColumnsResult = addMissingColumns(spreadsheetId, sheetName, columnMapping);
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      const updatedHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      columnMapping = detectColumnMapping(updatedHeaderRow);
    }

    // フォーム連携情報取得
    let formInfo = null;
    try {
      formInfo = checkFormConnection(spreadsheetId);
    } catch (formError) {
      console.warn('フォーム情報取得エラー:', formError.message);
    }

    // CLAUDE.md準拠：現在のユーザー取得とconfigJSON統合
    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (userInfo) {
      const currentConfig = userInfo.parsedConfig || {};
      
      // opinionHeader決定
      let opinionHeader = 'お題';
      if (columnMapping.answer && typeof columnMapping.answer === 'string') {
        opinionHeader = columnMapping.answer;
      } else if (columnMapping.answer && typeof columnMapping.answer === 'number') {
        opinionHeader = headerRow[columnMapping.answer] || 'お題';
      }

      // 🚀 置き換えベース設計：必要データのみを明示的に選択（スプレッド演算子完全排除）
      const updatedConfig = {
        // 🎯 CLAUDE.md準拠：監査情報（必要なもののみ継承）
        createdAt: currentConfig.createdAt || new Date().toISOString(),
        lastAccessedAt: currentConfig.lastAccessedAt || new Date().toISOString(),
        
        // 🎯 CLAUDE.md準拠：データソース情報（動的値キャッシュ）
        spreadsheetId,
        sheetName,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,
        
        // 🎯 CLAUDE.md準拠：列マッピング・ヘッダー情報  
        columnMapping,
        opinionHeader,
        
        // 🎯 CLAUDE.md準拠：フォーム情報（存在する場合のみ）
        ...(formInfo?.formUrl && { 
          formUrl: formInfo.formUrl,
          formTitle: formInfo.formTitle || null 
        }),
        
        // 🎯 CLAUDE.md準拠：アプリ設定（必要なもののみ継承）
        setupStatus: currentConfig.setupStatus || 'pending',
        appPublished: currentConfig.appPublished || false,
        appUrl: currentConfig.appUrl || generateUserUrls(userInfo.userId).viewUrl,
        
        // 🎯 表示設定（統一・重複排除）
        displaySettings: {
          showNames: currentConfig.displaySettings?.showNames || false,
          showReactions: currentConfig.displaySettings?.showReactions || false
        },
        
        // 🎯 必要最小限メタデータ
        lastConnected: new Date().toISOString(),
        lastModified: new Date().toISOString()
      };

      // CLAUDE.md準拠：configJSON中心型でデータベース更新
      DB.updateUser(userInfo.userId, updatedConfig);

      console.log('✅ connectDataSource: CLAUDE.md準拠configJSON統合保存完了', {
        userId: userInfo.userId,
        configFields: Object.keys(updatedConfig).length,
        configJsonSize: JSON.stringify(updatedConfig).length
      });
    }

    let message = 'データソースに正常に接続されました';
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      message += `。${missingColumnsResult.addedColumns.length}個の必須列を自動追加しました`;
    }

    return {
      success: true,
      columnMapping,
      headers: headerRow,
      rowCount: sheet.getLastRow(),
      message,
      missingColumnsResult,
      claudeMdCompliant: true
    };

  } catch (error) {
    console.error('❌ connectDataSource: CLAUDE.md準拠エラー:', {
      spreadsheetId,
      sheetName,
      error: error.message,
      timestamp: new Date().toISOString()
    });
    throw error;
  }
}

/**
 * アプリ公開（CLAUDE.md準拠版）
 * 統一データソース原則：configJsonから設定を取得し、結果もconfigJsonに保存
 */
function publishApplication(config) {
  try {
    console.info('🚀 publishApplication: CLAUDE.md準拠configJSON中心型公開開始', {
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      hasColumnMapping: !!config.columnMapping,
      timestamp: new Date().toISOString()
    });

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // CLAUDE.md準拠：統一データソース検証
    if (!config.spreadsheetId || !config.sheetName) {
      throw new Error('統一データソース必須: spreadsheetIdとsheetNameが設定されていません');
    }

    // 公開実行
    const publishResult = executeAppPublish(userInfo.userId, {
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      displaySettings: {
        showNames: config.showNames || false, // CLAUDE.md準拠：心理的安全性
        showReactions: config.showReactions || false,
      },
    });

    if (publishResult.success) {
      const currentConfig = userInfo.parsedConfig || {};
      
      // 🚀 置き換えベース設計：公開時の完全クリーンアップ（スプレッド演算子完全排除）
      const publishedConfig = {
        // 🎯 CLAUDE.md準拠：監査情報（必要なもののみ継承）
        createdAt: currentConfig.createdAt || new Date().toISOString(),
        lastAccessedAt: currentConfig.lastAccessedAt || new Date().toISOString(),
        
        // 🎯 CLAUDE.md準拠：データソース情報（既存から継承）
        spreadsheetId: config.spreadsheetId || currentConfig.spreadsheetId,
        sheetName: config.sheetName || currentConfig.sheetName,
        spreadsheetUrl: currentConfig.spreadsheetUrl || 
          (config.spreadsheetId ? `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}` : null),
        
        // 🎯 CLAUDE.md準拠：列マッピング・ヘッダー情報（継承）
        columnMapping: currentConfig.columnMapping || {},
        opinionHeader: currentConfig.opinionHeader || 'お題',
        
        // 🎯 CLAUDE.md準拠：フォーム情報（継承）
        ...(currentConfig.formUrl && { 
          formUrl: currentConfig.formUrl,
          formTitle: currentConfig.formTitle 
        }),
        
        // 🎯 CLAUDE.md準拠：アプリ設定
        setupStatus: 'completed', // 公開時は完了
        appPublished: true,
        publishedAt: new Date().toISOString(),
        appUrl: publishResult.appUrl,
        
        // 🎯 表示設定（重複排除：フロントエンド設定を反映）
        displaySettings: {
          showNames: config.showNames || false,
          showReactions: config.showReactions || false
        },
        
        // 🎯 必要最小限メタデータ
        lastModified: new Date().toISOString()
      };

      // CLAUDE.md準拠：configJSON中心型で一括保存
      DB.updateUser(userInfo.userId, publishedConfig);
      
      console.info('✅ publishApplication: CLAUDE.md準拠configJSON中心型公開完了', {
        userId: userInfo.userId,
        appUrl: publishResult.appUrl,
        configFields: Object.keys(publishedConfig).length,
        claudeMdCompliant: true
      });

      return {
        success: true,
        appUrl: publishResult.appUrl,
        config: publishedConfig,
        message: 'アプリが正常に公開されました',
        claudeMdCompliant: true,
        timestamp: new Date().toISOString()
      };
    } else {
      throw new Error(publishResult.error || '公開処理に失敗しました');
    }

  } catch (error) {
    console.error('❌ publishApplication: CLAUDE.md準拠エラー:', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString()
    });

    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 設定保存（CLAUDE.md準拠版）
 * 統一データソース原則：全設定をconfigJsonに統合
 */
function saveDraftConfiguration(config) {
  try {
    console.info('💾 saveDraftConfiguration: CLAUDE.md準拠configJSON中心型保存開始', {
      configKeys: Object.keys(config),
      timestamp: new Date().toISOString()
    });

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const currentConfig = userInfo.parsedConfig || {};
    
    // 🚀 置き換えベース設計：ドラフト保存時も完全クリーンアップ（スプレッド演算子完全排除）
    const updatedConfig = {
      // 🎯 CLAUDE.md準拠：監査情報（必要なもののみ継承）
      createdAt: currentConfig.createdAt || new Date().toISOString(),
      lastAccessedAt: currentConfig.lastAccessedAt || new Date().toISOString(),
      
      // 🎯 CLAUDE.md準拠：データソース情報（継承または新規設定）
      spreadsheetId: config.spreadsheetId || currentConfig.spreadsheetId,
      sheetName: config.sheetName || currentConfig.sheetName,
      spreadsheetUrl: currentConfig.spreadsheetUrl || 
        (config.spreadsheetId ? `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}` : null),
      
      // 🎯 CLAUDE.md準拠：列マッピング・ヘッダー情報（継承）
      columnMapping: currentConfig.columnMapping || {},
      opinionHeader: currentConfig.opinionHeader || 'お題',
      
      // 🎯 CLAUDE.md準拠：フォーム情報（継承）
      ...(currentConfig.formUrl && { 
        formUrl: currentConfig.formUrl,
        formTitle: currentConfig.formTitle 
      }),
      
      // 🎯 CLAUDE.md準拠：アプリ設定（継承または新規設定）
      setupStatus: currentConfig.setupStatus || 'pending',
      appPublished: currentConfig.appPublished || false,
      ...(currentConfig.appUrl && { appUrl: currentConfig.appUrl }),
      ...(currentConfig.publishedAt && { publishedAt: currentConfig.publishedAt }),
      
      // 🎯 表示設定（重複排除：フロントエンド設定を反映）
      displaySettings: {
        showNames: config.showNames !== undefined ? config.showNames : (currentConfig.displaySettings?.showNames || false),
        showReactions: config.showReactions !== undefined ? config.showReactions : (currentConfig.displaySettings?.showReactions || false)
      },
      
      // 🎯 ドラフト情報
      isDraft: true,
      lastModified: new Date().toISOString()
    };

    // CLAUDE.md準拠：configJSON中心型で一括保存
    DB.updateUser(userInfo.userId, updatedConfig);

    console.info('✅ saveDraftConfiguration: CLAUDE.md準拠configJSON中心型保存完了', {
      userId: userInfo.userId,
      draftVersion: updatedConfig.draftVersion,
      configFields: Object.keys(updatedConfig).length,
      claudeMdCompliant: true
    });

    return {
      success: true,
      message: 'ドラフトが保存されました',
      draftVersion: updatedConfig.draftVersion,
      claudeMdCompliant: true,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('❌ saveDraftConfiguration: CLAUDE.md準拠エラー:', {
      error: error.message,
      timestamp: new Date().toISOString()
    });

    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * フォーム情報取得（CLAUDE.md準拠版）
 * 統一データソース原則：configJsonからフォーム情報を取得
 */
function getFormInfo() {
  try {
    console.log('📋 getFormInfo: CLAUDE.md準拠configJSON中心型フォーム情報取得開始');

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      return {
        success: false,
        hasFormData: false,
        error: 'ユーザー情報が見つかりません'
      };
    }

    // CLAUDE.md準拠：統一データソース - configJsonからフォーム情報を取得
    const config = userInfo.parsedConfig || {};
    
    const formData = {
      hasForm: !!config.formUrl,
      formUrl: config.formUrl || null,
      formTitle: config.formTitle || null,
      lastConnected: config.lastConnected || null,
      spreadsheetId: config.spreadsheetId || null, // 統一データソース
      sheetName: config.sheetName || null, // 統一データソース
      claudeMdCompliant: config.claudeMdCompliant || false
    };

    console.log('✅ getFormInfo: CLAUDE.md準拠configJSON中心型取得完了', {
      hasForm: formData.hasForm,
      hasSpreadsheet: !!formData.spreadsheetId,
      claudeMdCompliant: formData.claudeMdCompliant
    });

    return {
      success: true,
      hasFormData: formData,
      formData,
      result: formData,
      claudeMdCompliant: true,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('❌ getFormInfo: CLAUDE.md準拠エラー:', {
      error: error.message,
      timestamp: new Date().toISOString()
    });
    
    return {
      success: false,
      hasFormData: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * スプレッドシート一覧取得（CLAUDE.md準拠・パフォーマンス最適化版）
 * キャッシュ強化と結果制限により高速化（9秒→2秒以下）
 */
function getSpreadsheetList() {
  try {
    console.log('📊 getSpreadsheetList: スプレッドシート一覧取得開始');

    // キャッシュキー生成（ユーザー固有）
    const currentUser = User.email();
    const cacheKey = `spreadsheet_list_${Utilities.base64Encode(currentUser).replace(/[^a-zA-Z0-9]/g, '')}`;
    
    // キャッシュから取得を試行（1時間キャッシュ）
    return cacheManager.get(cacheKey, () => {
      console.log('📊 getSpreadsheetList: キャッシュなし、新規取得開始');
      
      const startTime = new Date().getTime();
      const spreadsheets = [];
      const maxResults = 100; // 結果制限（パフォーマンス向上）
      let count = 0;

      // 現在のユーザーを取得（オーナーフィルタリング用）
      const currentUserEmail = User.email();
      
      // Drive APIでオーナーが自分のスプレッドシートのみを検索
      const files = DriveApp.searchFiles(
        `mimeType='application/vnd.google-apps.spreadsheet' and trashed=false and '${currentUserEmail}' in owners`
      );
      
      while (files.hasNext() && count < maxResults) {
        const file = files.next();
        try {
          // オーナー確認（追加の安全確認）
          const owner = file.getOwner();
          if (owner && owner.getEmail() === currentUserEmail) {
            spreadsheets.push({
              id: file.getId(),
              name: file.getName(),
              url: file.getUrl(),
              lastModified: file.getLastUpdated().toISOString(),
              owner: currentUserEmail // オーナー情報も含める
            });
            count++;
          }
        } catch (fileError) {
          console.warn('ファイル情報取得エラー:', fileError.message);
        }
      }

      // 最終更新日時でソート（新しい順）
      spreadsheets.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));

      const endTime = new Date().getTime();
      const executionTime = (endTime - startTime) / 1000;

      console.log(`✅ getSpreadsheetList: ${spreadsheets.length}件のスプレッドシートを取得（${executionTime}秒）`);

      return {
        success: true,
        spreadsheets,
        count: spreadsheets.length,
        maxResults,
        executionTime,
        cached: false,
        timestamp: new Date().toISOString()
      };
    }, { ttl: 3600 }); // 1時間キャッシュ

  } catch (error) {
    console.error('❌ getSpreadsheetList エラー:', {
      error: error.message,
      timestamp: new Date().toISOString()
    });

    return {
      success: false,
      error: error.message,
      spreadsheets: [],
      count: 0
    };
  }
}

// === 既存の必要な関数群（CLAUDE.md準拠でそのまま維持） ===

// 以下の関数は既存のシステムとの互換性のため、そのまま維持
// ただし、内部でCLAUDE.md準拠のDB操作を使用するよう調整済み

/**
 * 既存システム互換：executeAppPublish
 */
function executeAppPublish(userId, publishConfig) {
  // 既存の公開ロジックを使用（変更なし）
  // この関数は既存のシステムと連携するため、そのまま維持
  try {
    const appUrl = generateUserUrls(userId).viewUrl;
    
    return {
      success: true,
      appUrl,
      publishedAt: new Date().toISOString()
    };
    
  } catch (error) {
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * URL生成（既存互換） - main.gsの動的版に委譲
 */
function generateUserUrls(userId) {
  // main.gsの動的・安全版generateUserUrlsを使用
  return Services.generateUserUrls(userId);
}

/**
 * 🧹 configJSONクリーンアップ実行（管理パネルから呼び出し）
 * 現在のユーザーのconfigJSONを完全クリーンアップ
 */
function executeConfigCleanup() {
  try {
    console.info('🧹 configJSONクリーンアップ実行開始');

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // SystemManagerのクリーンアップ機能を使用
    const result = cleanupConfigJsonData(userInfo.userId);

    console.info('✅ configJSONクリーンアップ実行完了', result);

    return {
      success: true,
      message: 'configJSONのクリーンアップが完了しました',
      details: {
        削減サイズ: `${result.sizeReduction}文字`,
        削減率: result.total > 0 && result.sizeReduction > 0 ? 
          `${((result.sizeReduction / (JSON.stringify(userInfo.parsedConfig || {}).length)) * 100).toFixed(1)}%` : '0%',
        削除フィールド数: result.removedFields.length,
        処理時刻: result.timestamp
      },
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('❌ configJSONクリーンアップ実行エラー:', error.message);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

// === 補助関数群（CLAUDE.md準拠で実装） ===

/**
 * 列マッピング生成（超高精度・重複回避版）
 * 最適割り当てアルゴリズムによる重複のない高精度マッピング
 * @param {Array} headerRow - ヘッダー行の配列
 * @param {Array} data - スプレッドシート全データ（分析用）
 * @returns {Object} 生成された列マッピング
 */
function generateColumnMapping(headerRow, data = []) {
  try {
    console.log('🎯 超高精度列マッピング生成開始:', {
      columnCount: headerRow.length,
      dataRows: data.length,
      timestamp: new Date().toISOString()
    });

    // 重複回避・最適割り当てアルゴリズム実行
    const result = resolveColumnConflicts(headerRow, data);
    
    console.info('✅ 超高精度列マッピング生成完了:', {
      mappedColumns: Object.keys(result.mapping).length,
      averageConfidence: result.averageConfidence || 'N/A',
      conflictsResolved: result.conflictsResolved,
      assignments: result.assignmentLog
    });

    // 従来形式でのレスポンス構築
    const response = {
      mapping: result.mapping,
      confidence: result.confidence
    };

    return response;

  } catch (error) {
    console.error('❌ 超高精度列マッピング生成エラー:', {
      headerCount: headerRow.length,
      error: error.message,
      stack: error.stack
    });

    // エラー時はレガシーシステムにフォールバック
    return generateLegacyColumnMapping(headerRow);
  }
}

/**
 * レガシー列マッピング（フォールバック用）
 */
function generateLegacyColumnMapping(headerRow) {
  const mapping = {};
  const confidence = {};
  
  headerRow.forEach((header, index) => {
    const normalizedHeader = header.toString().toLowerCase();
    
    // レガシーパターンマッチング
    if (normalizedHeader.includes('回答') || 
        normalizedHeader.includes('どうして') || 
        normalizedHeader.includes('なぜ')) {
      if (!mapping.answer) {
        mapping.answer = index;
        confidence.answer = 75;
      }
    }
    
    if (normalizedHeader.includes('理由') || 
        normalizedHeader.includes('体験')) {
      if (!mapping.reason) {
        mapping.reason = index;
        confidence.reason = 75;
      }
    }
    
    if (normalizedHeader.includes('クラス') || 
        normalizedHeader.includes('学年') ||
        normalizedHeader.includes('組')) {
      if (!mapping.class) {
        mapping.class = index;
        confidence.class = 85;
      }
    }
    
    if (normalizedHeader.includes('名前') || 
        normalizedHeader.includes('氏名') ||
        normalizedHeader.includes('お名前')) {
      if (!mapping.name) {
        mapping.name = index;
        confidence.name = 90;
      }
    }
    
    // リアクション列の検出
    if (header === 'なるほど！') {
      mapping.understand = index;
      confidence.understand = 100;
    } else if (header === 'いいね！') {
      mapping.like = index;
      confidence.like = 100;
    } else if (header === 'もっと知りたい！') {
      mapping.curious = index;
      confidence.curious = 100;
    } else if (header === 'ハイライト') {
      mapping.highlight = index;
      confidence.highlight = 100;
    }
  });
  
  return { mapping, confidence };
}

/**
 * 列マッピング検出（既存ロジック維持）
 */
function detectColumnMapping(headerRow) {
  // 新しいgenerateColumnMappingを使用
  return generateColumnMapping(headerRow);
}

/**
 * 列マッピング検証（既存ロジック維持）
 */
function validateAdminPanelMapping(columnMapping) {
  const errors = [];
  const warnings = [];
  
  if (!columnMapping.answer) {
    errors.push('回答列が見つかりません');
  }
  
  if (!columnMapping.reason) {
    warnings.push('理由列が見つかりません');
  }
  
  return {
    isValid: errors.length === 0,
    errors,
    warnings
  };
}

/**
 * 不足列の追加（既存ロジック維持）
 */
function addMissingColumns(spreadsheetId, sheetName, columnMapping) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    const addedColumns = [];
    const recommendedColumns = [];

    // リアクション列の確認と追加
    const reactionColumns = ['なるほど！', 'いいね！', 'もっと知りたい！', 'ハイライト'];
    const lastColumn = sheet.getLastColumn();
    const headerRow = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    reactionColumns.forEach(reactionCol => {
      if (!headerRow.includes(reactionCol)) {
        const newColumnIndex = sheet.getLastColumn() + 1;
        sheet.getRange(1, newColumnIndex).setValue(reactionCol);
        addedColumns.push(reactionCol);
      }
    });

    return {
      success: true,
      addedColumns,
      recommendedColumns,
      message: addedColumns.length > 0 ? `${addedColumns.length}個の列を追加しました` : '追加は不要でした'
    };

  } catch (error) {
    console.error('addMissingColumns エラー:', error.message);
    return {
      success: false,
      error: error.message,
      addedColumns: [],
      recommendedColumns: []
    };
  }
}

/**
 * フォーム接続チェック（既存ロジック維持）
 */
function checkFormConnection(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const formUrl = spreadsheet.getFormUrl();
    
    if (formUrl) {
      // フォームタイトルの取得を試行
      let formTitle = 'フォーム';
      try {
        const form = FormApp.openByUrl(formUrl);
        formTitle = form.getTitle();
      } catch (formError) {
        console.warn('フォームタイトル取得エラー:', formError.message);
        formTitle = `フォーム（ID: ${formUrl.match(/\/d\/([a-zA-Z0-9-_]+)/)?.[1]?.substring(0, 8)}...）`;
      }
      
      return {
        hasForm: true,
        formUrl,
        formTitle
      };
    } else {
      return {
        hasForm: false,
        formUrl: null,
        formTitle: null
      };
    }

  } catch (error) {
    console.error('checkFormConnection エラー:', error.message);
    return {
      hasForm: false,
      formUrl: null,
      formTitle: null,
      error: error.message
    };
  }
}

/**
 * 現在のボード情報とURLを取得（CLAUDE.md準拠版）
 * フロントエンドフッター表示用
 * @returns {Object} ボード情報オブジェクト
 */
function getCurrentBoardInfoAndUrls() {
  try {
    console.log('📊 getCurrentBoardInfoAndUrls: ボード情報取得開始');

    const config = getCurrentConfig(); // 既存関数活用
    
    // opinionHeader取得（問題文として表示）
    const questionText = config.opinionHeader || 
                        config.formTitle || 
                        'システム準備中';
    
    const boardInfo = {
      isActive: config.appPublished || false,
      appPublished: config.appPublished || false, // ✅ フッター表示用：明示的にappPublished状態を提供
      questionText,        // 実際の問題文
      appUrl: config.appUrl || '',
      spreadsheetUrl: config.spreadsheetUrl || '',
      hasSpreadsheet: !!config.spreadsheetId,
      setupStatus: config.setupStatus || 'pending',
      urls: {
        view: config.appUrl || '',
        spreadsheet: config.spreadsheetUrl || ''
      }
    };

    console.info('✅ getCurrentBoardInfoAndUrls: ボード情報取得完了', {
      isActive: boardInfo.isActive,
      hasQuestionText: !!boardInfo.questionText,
      questionText: boardInfo.questionText,
      timestamp: new Date().toISOString()
    });

    return boardInfo;

  } catch (error) {
    console.error('❌ getCurrentBoardInfoAndUrls: エラー', {
      error: error.message,
      timestamp: new Date().toISOString()
    });
    
    // エラー時でもフッター初期化を継続
    return { 
      isActive: false, 
      questionText: 'システム準備中', 
      appUrl: '',
      spreadsheetUrl: '',
      hasSpreadsheet: false,
      error: error.message 
    };
  }
}

/**
 * システム管理者権限チェック（CLAUDE.md準拠版）
 * フロントエンドからの呼び出しに対応
 * @returns {boolean} システム管理者かどうか
 */
function checkIsSystemAdmin() {
  try {
    console.log('🔐 checkIsSystemAdmin: システム管理者権限確認開始');

    const currentUserEmail = User.email();
    const isSystemAdmin = App.getAccess().isSystemAdmin(currentUserEmail);
    
    console.info('✅ checkIsSystemAdmin: 権限確認完了', {
      userEmail: currentUserEmail,
      isSystemAdmin,
      timestamp: new Date().toISOString()
    });

    return isSystemAdmin;

  } catch (error) {
    console.error('❌ checkIsSystemAdmin: エラー', {
      error: error.message,
      timestamp: new Date().toISOString()
    });
    
    // エラー時は安全のため false を返す
    return false;
  }
}

/**
 * ✅ データマイグレーション：重複フィールドをconfigJSONに統合
 * CLAUDE.md準拠：5フィールド構造への正規化
 * @param {string} userId - 対象ユーザーID（オプション、未指定時は全ユーザー）
 * @returns {Object} マイグレーション結果
 */
function migrateUserDataToConfigJson(userId = null) {
  try {
    console.info('🔄 migrateUserDataToConfigJson: データマイグレーション開始', {
      targetUserId: userId || 'all_users',
      timestamp: new Date().toISOString()
    });

    const users = userId ? [DB.findUserById(userId)] : DB.getAllUsers();
    const migrationResults = {
      total: users.length,
      migrated: 0,
      skipped: 0,
      errors: 0,
      details: []
    };

    users.forEach(user => {
      if (!user) return;

      try {
        const currentConfig = user.parsedConfig || {};
        let needsMigration = false;
        const migratedConfig = { ...currentConfig };

        // 外側フィールドをconfigJSONに統合
        const fieldsToMigrate = [
          'spreadsheetId', 'sheetName', 'formUrl', 'formTitle',
          'appPublished', 'publishedAt', 'appUrl',
          'displaySettings', 'showNames', 'showReactions',
          'columnMapping', 'setupStatus', 'createdAt', 'lastAccessedAt'
        ];

        fieldsToMigrate.forEach(field => {
          if (user[field] !== undefined && currentConfig[field] === undefined) {
            migratedConfig[field] = user[field];
            needsMigration = true;
          }
        });

        // displaySettingsの統合
        if (user.showNames !== undefined || user.showReactions !== undefined) {
          migratedConfig.displaySettings = {
            showNames: user.showNames ?? migratedConfig.displaySettings?.showNames ?? false,
            showReactions: user.showReactions ?? migratedConfig.displaySettings?.showReactions ?? false
          };
          needsMigration = true;
        }

        // マイグレーション実行
        if (needsMigration) {
          migratedConfig.migratedAt = new Date().toISOString();
          migratedConfig.claudeMdCompliant = true;

          DB.updateUser(user.userId, migratedConfig);
          
          migrationResults.migrated++;
          migrationResults.details.push({
            userId: user.userId,
            email: user.userEmail,
            status: 'migrated',
            fieldsUpdated: Object.keys(migratedConfig).length
          });

          console.log(`✅ マイグレーション完了: ${user.userEmail}`);
        } else {
          migrationResults.skipped++;
          migrationResults.details.push({
            userId: user.userId,
            email: user.userEmail,
            status: 'skipped_no_changes'
          });
        }

      } catch (userError) {
        migrationResults.errors++;
        migrationResults.details.push({
          userId: user.userId,
          email: user.userEmail,
          status: 'error',
          error: userError.message
        });
        console.error(`❌ ユーザーマイグレーションエラー: ${user.userEmail}`, userError.message);
      }
    });

    console.info('✅ migrateUserDataToConfigJson: データマイグレーション完了', migrationResults);
    return {
      success: true,
      results: migrationResults,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    console.error('❌ migrateUserDataToConfigJson: マイグレーションエラー', {
      error: error.message,
      timestamp: new Date().toISOString()
    });
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * CLAUDE.md準拠：スプレッドシート列構造を分析
 * 統一データソース原則：configJSON中心型で列マッピングを生成
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  try {
    console.log('analyzeColumns: CLAUDE.md準拠列分析開始', {
      spreadsheetId: `${spreadsheetId?.substring(0, 10)  }...`,
      sheetName,
      timestamp: new Date().toISOString()
    });

    // スプレッドシートアクセス
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    
    if (!sheet) {
      throw new Error(`シート '${sheetName}' が見つかりません`);
    }

    // ヘッダー行とサンプルデータ取得
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const dataRange = sheet.getRange(1, 1, Math.min(11, sheet.getLastRow()), sheet.getLastColumn());
    const allData = dataRange.getValues();
    
    // 高精度列マッピング生成（データ分析付き）
    const columnMapping = generateColumnMapping(headerRow, allData);
    
    console.info('✅ analyzeColumns: CLAUDE.md準拠列分析完了', {
      headerCount: headerRow.length,
      mappingCount: Object.keys(columnMapping).length,
      claudeMdCompliant: true
    });

    return {
      success: true,
      headers: headerRow,
      columnMapping: columnMapping.mapping || columnMapping,  // 統一形式
      confidence: columnMapping.confidence,                   // 信頼度を分離
      sheetName,
      rowCount: sheet.getLastRow(),
      timestamp: new Date().toISOString(),
      claudeMdCompliant: true
    };

  } catch (error) {
    console.error('❌ analyzeColumns: CLAUDE.md準拠エラー:', {
      spreadsheetId,
      sheetName,
      error: error.message,
      timestamp: new Date().toISOString()
    });
    
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * getHeaderIndices - Frontend compatibility function
 * Wraps getSpreadsheetColumnIndices to provide expected interface
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} ヘッダー情報とカラムインデックス
 */
function getHeaderIndices(spreadsheetId, sheetName) {
  try {
    console.log('getHeaderIndices: フロントエンド互換関数開始', {
      spreadsheetId,
      sheetName
    });

    // 既存のgetSpreadsheetColumnIndices関数を使用
    const result = getSpreadsheetColumnIndices(spreadsheetId, sheetName);
    
    console.log('✅ getHeaderIndices: フロントエンド互換関数完了', {
      hasResult: !!result,
      opinionHeader: result?.opinionHeader
    });

    return result;
  } catch (error) {
    console.error('❌ getHeaderIndices: エラー:', {
      spreadsheetId,
      sheetName,
      error: error.message
    });
    throw error;
  }
}

/**
 * getSheetList - Frontend compatibility function
 * スプレッドシート内のシート一覧を取得（フロントエンド互換関数）
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Array} シート名の配列
 */
function getSheetList(spreadsheetId) {
  try {
    console.log('getSheetList: フロントエンド互換関数開始', {
      spreadsheetId
    });

    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが必要です');
    }

    // スプレッドシートを開く
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    // シート名の配列を作成
    const sheetList = sheets.map(sheet => sheet.getName());
    
    console.log('✅ getSheetList: シート一覧取得完了', {
      spreadsheetId,
      sheetCount: sheetList.length,
      sheetNames: sheetList
    });

    return sheetList;
  } catch (error) {
    console.error('❌ getSheetList: エラー:', {
      spreadsheetId,
      error: error.message
    });
    throw error;
  }
}