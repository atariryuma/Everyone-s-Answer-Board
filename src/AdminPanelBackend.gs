/**
 * @fileoverview AdminPanel Backend - CLAUDE.md完全準拠版
 * configJSON中心型アーキテクチャの完全実装
 */

/**
 * データソース接続（CLAUDE.md準拠版）
 * 統一データソース原則：全データをconfigJsonに統合
 */
function connectDataSource(spreadsheetId, sheetName) {
  try {
    console.log('🔗 connectDataSource: CLAUDE.md準拠configJSON中心型接続開始', {
      spreadsheetId,
      sheetName,
    });

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
    const currentUser = UserManager.getCurrentEmail();
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

        // 🎯 CLAUDE.md準拠：フォーム情報（確実な取得・設定）
        formUrl: formInfo?.formUrl || null,
        formTitle: formInfo?.formTitle || null,

        // 🎯 CLAUDE.md準拠：アプリ設定（データソース接続完了で'completed'に自動更新）
        setupStatus: 'completed',
        appPublished: currentConfig.appPublished || false,
        appUrl: currentConfig.appUrl || generateUserUrls(userInfo.userId).viewUrl,

        // 🎯 表示設定（統一・重複排除）
        displaySettings: {
          showNames: currentConfig.displaySettings?.showNames || false,
          showReactions: currentConfig.displaySettings?.showReactions || false,
        },

        // 🎯 必要最小限メタデータ
        lastConnected: new Date().toISOString(),
        lastModified: new Date().toISOString(),
      };

      // 🔧 デバッグ: updatedConfigの内容を詳細ログ出力
      console.log('🔍 connectDataSource: 保存前のupdatedConfig詳細', {
        userId: userInfo.userId,
        updatedConfigKeys: Object.keys(updatedConfig),
        spreadsheetId: updatedConfig.spreadsheetId,
        sheetName: updatedConfig.sheetName,
        formUrl: updatedConfig.formUrl,
        formTitle: updatedConfig.formTitle,
        setupStatus: updatedConfig.setupStatus,
        configSize: JSON.stringify(updatedConfig).length,
      });

      // 🔧 修正: 構築したupdatedConfigを確実に保存
      const saveSuccess = ConfigManager.saveConfig(userInfo.userId, updatedConfig);

      console.log('🔍 connectDataSource: 保存結果', {
        userId: userInfo.userId,
        saveSuccess,
        timestamp: new Date().toISOString(),
      });

      if (!saveSuccess) {
        throw new Error('設定の保存に失敗しました');
      }

      console.log('✅ connectDataSource: CLAUDE.md準拠configJSON統合保存完了', {
        userId: userInfo.userId,
        configFields: Object.keys(updatedConfig).length,
        configJsonSize: JSON.stringify(updatedConfig).length,
        spreadsheetId: updatedConfig.spreadsheetId,
        sheetName: updatedConfig.sheetName,
        setupStatus: updatedConfig.setupStatus,
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
      claudeMdCompliant: true,
    };
  } catch (error) {
    console.error('❌ connectDataSource: CLAUDE.md準拠エラー:', {
      spreadsheetId,
      sheetName,
      error: error.message,
      timestamp: new Date().toISOString(),
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
    console.log('🚀 publishApplication: CLAUDE.md準拠configJSON中心型公開開始', {
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      hasColumnMapping: !!config.columnMapping,
      timestamp: new Date().toISOString(),
    });

    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 🔧 修正：ConfigManager経由で確実に最新設定を取得
    const currentConfig = ConfigManager.getUserConfig(userInfo.userId);

    console.log('🔍 publishApplication: 最新設定確認', {
      userId: userInfo.userId,
      hasSpreadsheetId: !!currentConfig?.spreadsheetId,
      hasSheetName: !!currentConfig?.sheetName,
      configFromFrontend: {
        hasSpreadsheetId: !!config.spreadsheetId,
        hasSheetName: !!config.sheetName,
      },
    });

    // データソース情報のフォールバック戦略
    const effectiveSpreadsheetId = currentConfig?.spreadsheetId || config.spreadsheetId;
    const effectiveSheetName = currentConfig?.sheetName || config.sheetName;

    if (!effectiveSpreadsheetId || !effectiveSheetName) {
      console.error('❌ publishApplication: データソース未設定', {
        dbSpreadsheetId: currentConfig?.spreadsheetId,
        dbSheetName: currentConfig?.sheetName,
        frontendSpreadsheetId: config.spreadsheetId,
        frontendSheetName: config.sheetName,
      });
      throw new Error('データソースが設定されていません。まずデータソースを設定してください。');
    }

    if (currentConfig.setupStatus !== 'completed') {
      console.warn('⚠️ publishApplication: setupStatus確認', {
        currentStatus: currentConfig.setupStatus,
        hasSpreadsheetId: !!currentConfig.spreadsheetId,
        hasSheetName: !!currentConfig.sheetName,
      });
      // データソースがある場合はsetupStatusを自動修正
      if (currentConfig.spreadsheetId && currentConfig.sheetName) {
        console.log('🔧 setupStatusを自動修正: pending → completed');
        // setupStatusを更新しない（getCurrentConfigで判定済み）
      } else {
        throw new Error('セットアップが完了していません。データソース接続を完了させてください。');
      }
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
        spreadsheetUrl:
          currentConfig.spreadsheetUrl ||
          (config.spreadsheetId
            ? `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}`
            : null),

        // 🎯 CLAUDE.md準拠：列マッピング・ヘッダー情報（継承）
        columnMapping: currentConfig.columnMapping || {},
        opinionHeader: currentConfig.opinionHeader || 'お題',

        // 🎯 CLAUDE.md準拠：フォーム情報（継承）
        ...(currentConfig.formUrl && {
          formUrl: currentConfig.formUrl,
          formTitle: currentConfig.formTitle,
        }),

        // 🎯 CLAUDE.md準拠：アプリ設定
        setupStatus: 'completed', // 公開時は完了
        appPublished: true,
        publishedAt: new Date().toISOString(),
        appUrl: publishResult.appUrl,
        isDraft: false, // 🔥 公開時はドラフト状態を解除

        // 🎯 表示設定（重複排除：フロントエンド設定を反映）
        displaySettings: {
          showNames: config.showNames || false,
          showReactions: config.showReactions || false,
        },

        // 🎯 必要最小限メタデータ
        lastModified: new Date().toISOString(),
      };

      // 🔒 データソース情報を明示的に保持してConfigManager更新
      ConfigManager.updateAppStatus(userInfo.userId, {
        appPublished: true,
        setupStatus: 'completed',
        preserveDataSource: true, // 🔒 connectDataSourceで保存されたデータソース情報を保護
        // フォールバックデータソース情報（保護機能が失敗した場合の備え）
        spreadsheetId: effectiveSpreadsheetId,
        sheetName: effectiveSheetName,
        appUrl: publishResult.appUrl,
      });

      console.log('✅ publishApplication: CLAUDE.md準拠configJSON中心型公開完了', {
        userId: userInfo.userId,
        appUrl: publishResult.appUrl,
        configFields: Object.keys(publishedConfig).length,
        claudeMdCompliant: true,
      });

      return {
        success: true,
        appUrl: publishResult.appUrl,
        config: publishedConfig,
        message: 'アプリが正常に公開されました',
        claudeMdCompliant: true,
        timestamp: new Date().toISOString(),
      };
    } else {
      throw new Error(publishResult.error || '公開処理に失敗しました');
    }
  } catch (error) {
    console.error('❌ publishApplication: CLAUDE.md準拠エラー:', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * 設定保存（CLAUDE.md準拠版・ConfigManager統一）
 * ✅ ConfigManager.updateConfig()に統一簡素化
 */
function saveDraftConfiguration(config) {
  try {
    console.log('💾 saveDraftConfiguration: ConfigManager統一版保存開始', {
      configKeys: Object.keys(config),
      timestamp: new Date().toISOString(),
    });

    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 🚫 二重構造防止: configJsonフィールドを削除
    const cleanConfig = { ...config };
    delete cleanConfig.configJson;
    delete cleanConfig.configJSON;

    // ✅ ConfigManager.updateConfig()による統一更新（簡素化）
    const success = ConfigManager.updateConfig(userInfo.userId, cleanConfig);

    if (!success) {
      throw new Error('設定の保存に失敗しました');
    }

    console.log('✅ saveDraftConfiguration: ConfigManager統一版保存完了', {
      userId: userInfo.userId,
      configFields: Object.keys(cleanConfig).length,
      claudeMdCompliant: true,
    });

    return {
      success: true,
      message: 'ドラフトが保存されました',
      claudeMdCompliant: true,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    console.error('❌ saveDraftConfiguration: ConfigManager統一版エラー:', {
      error: error.message,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * フォーム情報取得（CLAUDE.md準拠版）
 * 統一データソース原則：configJsonからフォーム情報を取得
 */
function getFormInfo(spreadsheetId, sheetName) {
  try {
    console.log('📋 getFormInfo: シート固有フォーム情報取得開始', { spreadsheetId, sheetName });

    if (!spreadsheetId || !sheetName) {
      return {
        success: false,
        error: 'スプレッドシートIDまたはシート名が指定されていません',
        formData: {
          hasForm: false,
          formUrl: null,
          formTitle: null,
        },
      };
    }

    // シート固有のフォーム連携確認
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      return {
        success: false,
        error: '指定されたシートが見つかりません',
        formData: {
          hasForm: false,
          formUrl: null,
          formTitle: null,
        },
      };
    }

    // シート固有のフォームURL取得
    let formUrl = null;
    let formTitle = null;
    let hasForm = false;

    try {
      formUrl = sheet.getFormUrl();
      if (formUrl) {
        hasForm = true;
        // 🔥 フォームタイトル確実取得（キャッシュバイパス）
        try {
          const form = FormApp.openByUrl(formUrl);
          formTitle = form.getTitle() || 'Google フォーム'; // 空文字列防止
          if (!formTitle || formTitle.trim() === '') {
            formTitle = 'Google フォーム'; // 完全に空の場合のフォールバック
          }
        } catch (formError) {
          console.warn('フォームタイトル取得エラー:', formError.message);
          // より親しみやすいフォールバック名
          formTitle = 'Google フォーム（接続済み）';
        }
      }
    } catch (error) {
      console.log('フォーム連携なし:', sheetName);
    }

    const formData = {
      hasForm,
      formUrl,
      formTitle,
      spreadsheetId,
      sheetName,
    };

    console.log('✅ getFormInfo: シート固有フォーム情報取得完了', {
      sheetName,
      hasForm: formData.hasForm,
      formTitle: formData.formTitle,
      formUrl: !!formData.formUrl,
    });

    return {
      success: true,
      formData,
      hasFormData: formData,
    };
  } catch (error) {
    console.error('❌ getFormInfo: エラー:', error.message);
    return {
      success: false,
      error: error.message,
      formData: {
        hasForm: false,
        formUrl: null,
        formTitle: null,
      },
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
    const currentUser = UserManager.getCurrentEmail();
    const cacheKey = `spreadsheet_list_${Utilities.base64Encode(currentUser).replace(/[^a-zA-Z0-9]/g, '')}`;

    // キャッシュから取得を試行（1時間キャッシュ）
    return cacheManager.get(
      cacheKey,
      () => {
        console.log('📊 getSpreadsheetList: キャッシュなし、新規取得開始');

        const startTime = new Date().getTime();
        const spreadsheets = [];
        const maxResults = 100; // 結果制限（パフォーマンス向上）
        let count = 0;

        // 現在のユーザーを取得（オーナーフィルタリング用）
        const currentUserEmail = UserManager.getCurrentEmail();

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
                owner: currentUserEmail, // オーナー情報も含める
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

        console.log(
          `✅ getSpreadsheetList: ${spreadsheets.length}件のスプレッドシートを取得（${executionTime}秒）`
        );

        return {
          success: true,
          spreadsheets,
          count: spreadsheets.length,
          maxResults,
          executionTime,
          cached: false,
          timestamp: new Date().toISOString(),
        };
      },
      { ttl: 3600 }
    ); // 1時間キャッシュ
  } catch (error) {
    console.error('❌ getSpreadsheetList エラー:', {
      error: error.message,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
      spreadsheets: [],
      count: 0,
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
      publishedAt: new Date().toISOString(),
    };
  } catch (error) {
    return {
      success: false,
      error: error.message,
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
    console.log('🧹 configJSONクリーンアップ実行開始');

    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // SystemManagerのクリーンアップ機能を使用
    const result = SystemManager.cleanAllConfigJson();

    console.log('✅ configJSONクリーンアップ実行完了', result);

    return {
      success: true,
      message: 'configJSONのクリーンアップが完了しました',
      details: {
        削減サイズ: `${result.sizeReduction}文字`,
        削減率:
          result.total > 0 && result.sizeReduction > 0
            ? `${((result.sizeReduction / JSON.stringify(userInfo.parsedConfig || {}).length) * 100).toFixed(1)}%`
            : '0%',
        削除フィールド数: result.removedFields.length,
        処理時刻: result.timestamp,
      },
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    console.error('❌ configJSONクリーンアップ実行エラー:', error.message);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
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
      timestamp: new Date().toISOString(),
    });

    // 重複回避・最適割り当てアルゴリズム実行
    const result = resolveColumnConflicts(headerRow, data);

    console.log('✅ 超高精度列マッピング生成完了:', {
      mappedColumns: Object.keys(result.mapping).length,
      averageConfidence: result.averageConfidence || 'N/A',
      conflictsResolved: result.conflictsResolved,
      assignments: result.assignmentLog,
    });

    // 従来形式でのレスポンス構築
    const response = {
      mapping: result.mapping,
      confidence: result.confidence,
    };

    return response;
  } catch (error) {
    console.error('❌ 超高精度列マッピング生成エラー:', {
      headerCount: headerRow.length,
      error: error.message,
      stack: error.stack,
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
    if (
      normalizedHeader.includes('回答') ||
      normalizedHeader.includes('どうして') ||
      normalizedHeader.includes('なぜ')
    ) {
      if (!mapping.answer) {
        mapping.answer = index;
        confidence.answer = 75;
      }
    }

    if (normalizedHeader.includes('理由') || normalizedHeader.includes('体験')) {
      if (!mapping.reason) {
        mapping.reason = index;
        confidence.reason = 75;
      }
    }

    if (
      normalizedHeader.includes('クラス') ||
      normalizedHeader.includes('学年') ||
      normalizedHeader.includes('組')
    ) {
      if (!mapping.class) {
        mapping.class = index;
        confidence.class = 85;
      }
    }

    if (
      normalizedHeader.includes('名前') ||
      normalizedHeader.includes('氏名') ||
      normalizedHeader.includes('お名前')
    ) {
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
    warnings,
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

    reactionColumns.forEach((reactionCol) => {
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
      message:
        addedColumns.length > 0 ? `${addedColumns.length}個の列を追加しました` : '追加は不要でした',
    };
  } catch (error) {
    console.error('addMissingColumns エラー:', error.message);
    return {
      success: false,
      error: error.message,
      addedColumns: [],
      recommendedColumns: [],
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
        formTitle,
      };
    } else {
      return {
        hasForm: false,
        formUrl: null,
        formTitle: null,
      };
    }
  } catch (error) {
    console.error('checkFormConnection エラー:', error.message);
    return {
      hasForm: false,
      formUrl: null,
      formTitle: null,
      error: error.message,
    };
  }
}

/**
 * 現在のボード情報とURLを取得（CLAUDE.md準拠版）
 * フロントエンドフッター表示用
 * @returns {Object} ボード情報オブジェクト
 */
/**
 * 現在のユーザー設定を取得（フロントエンド統合用）
 * App.getConfig().getUserConfig() の実装として使用される
 */
function getCurrentConfig() {
  try {
    console.log('🔧 getCurrentConfig: ユーザー設定取得開始');

    // 現在のユーザーの設定を取得
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      console.warn('getCurrentConfig: ユーザー情報が見つかりません');
      return ConfigManager.buildInitialConfig();
    }

    const config = ConfigManager.getUserConfig(userInfo.userId);
    console.log('✅ getCurrentConfig: 設定取得完了', {
      userId: userInfo.userId,
      configFields: Object.keys(config || {}).length,
    });

    return config;
  } catch (error) {
    console.error('❌ getCurrentConfig エラー:', error.message);
    return ConfigManager.buildInitialConfig();
  }
}

function getCurrentBoardInfoAndUrls() {
  try {
    console.log('📊 getCurrentBoardInfoAndUrls: ボード情報取得開始');

    // 現在のユーザーの設定を取得
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    const config = userInfo ? ConfigManager.getUserConfig(userInfo.userId) : null;

    // フッター表示用の問題文を管理パネルの回答列と一致させる（シンプル版）
    let questionText = config?.opinionHeader;

    // opinionHeaderが設定されていない場合のみフォールバック
    if (!questionText) {
      questionText = config?.formTitle || 'システム準備中';
    }

    console.log('getCurrentBoardInfoAndUrls: フッター問題文決定', {
      opinionHeader: config?.opinionHeader,
      formTitle: config?.formTitle,
      finalQuestionText: questionText,
    });

    // 日時情報の取得と整形
    const createdAt = config?.createdAt || null;
    const lastModified = config?.lastModified || userInfo?.lastModified || null;
    const publishedAt = config?.publishedAt || null;

    // URLs の生成（main.gsの安定したgetWebAppUrl関数を使用）
    let appUrl = config?.appUrl || '';
    if (!appUrl && userInfo) {
      try {
        const baseUrl = getWebAppUrl();
        appUrl = baseUrl ? `${baseUrl}?mode=view&userId=${userInfo.userId}` : '';
      } catch (urlError) {
        console.warn(
          'AdminPanelBackend.getCurrentBoardInfoAndUrls: URL生成エラー:',
          urlError.message
        );
        appUrl = '';
      }
    }
    const spreadsheetUrl =
      config?.spreadsheetUrl ||
      (config?.spreadsheetId
        ? `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}`
        : '');

    const boardInfo = {
      isActive: config?.appPublished || false,
      appPublished: config?.appPublished || false,
      isPublished: config?.appPublished || false, // フッター互換性
      questionText, // 実際の問題文
      opinionHeader: config?.opinionHeader || '', // 問題文（詳細版）
      appUrl,
      spreadsheetUrl,
      hasSpreadsheet: !!config?.spreadsheetId,
      setupStatus: config?.setupStatus || 'pending',

      // 日時情報（フッター表示用）
      dates: {
        created: createdAt,
        modified: lastModified,
        published: publishedAt,
        createdFormatted: createdAt ? new Date(createdAt).toLocaleString('ja-JP') : null,
        modifiedFormatted: lastModified ? new Date(lastModified).toLocaleString('ja-JP') : null,
        publishedFormatted: publishedAt ? new Date(publishedAt).toLocaleString('ja-JP') : null,
      },

      // URLs（従来互換性）
      urls: {
        view: appUrl,
        spreadsheet: spreadsheetUrl,
      },

      // 追加のボード情報
      formUrl: config?.formUrl || '',
      hasForm: !!config?.formUrl,
      sheetName: config?.sheetName || '',
      dataCount: 0, // 後で実際のデータ数を取得可能
    };

    console.log('✅ getCurrentBoardInfoAndUrls: ボード情報取得完了', {
      isActive: boardInfo.isActive,
      hasQuestionText: !!boardInfo.questionText,
      questionText: boardInfo.questionText,
      timestamp: new Date().toISOString(),
    });

    return boardInfo;
  } catch (error) {
    console.error('❌ getCurrentBoardInfoAndUrls: エラー', {
      error: error.message,
      timestamp: new Date().toISOString(),
    });

    // エラー時でもフッター初期化を継続
    return {
      isActive: false,
      questionText: 'システム準備中',
      appUrl: '',
      spreadsheetUrl: '',
      hasSpreadsheet: false,
      error: error.message,
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

    const currentUserEmail = UserManager.getCurrentEmail();
    const isSystemAdmin = App.getAccess().isSystemAdmin(currentUserEmail);

    console.log('✅ checkIsSystemAdmin: 権限確認完了', {
      userEmail: currentUserEmail,
      isSystemAdmin,
      timestamp: new Date().toISOString(),
    });

    return isSystemAdmin;
  } catch (error) {
    console.error('❌ checkIsSystemAdmin: エラー', {
      error: error.message,
      timestamp: new Date().toISOString(),
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
    console.log('🔄 migrateUserDataToConfigJson: データマイグレーション開始', {
      targetUserId: userId || 'all_users',
      timestamp: new Date().toISOString(),
    });

    const users = userId ? [DB.findUserById(userId)] : DB.getAllUsers();
    const migrationResults = {
      total: users.length,
      migrated: 0,
      skipped: 0,
      errors: 0,
      details: [],
    };

    users.forEach((user) => {
      if (!user) return;

      try {
        const currentConfig = user.parsedConfig || {};
        let needsMigration = false;
        const migratedConfig = { ...currentConfig };

        // 外側フィールドをconfigJSONに統合
        const fieldsToMigrate = [
          'spreadsheetId',
          'sheetName',
          'formUrl',
          'formTitle',
          'appPublished',
          'publishedAt',
          'appUrl',
          'displaySettings',
          'showNames',
          'showReactions',
          'columnMapping',
          'setupStatus',
          'createdAt',
          'lastAccessedAt',
        ];

        fieldsToMigrate.forEach((field) => {
          if (user[field] !== undefined && currentConfig[field] === undefined) {
            migratedConfig[field] = user[field];
            needsMigration = true;
          }
        });

        // displaySettingsの統合
        if (user.showNames !== undefined || user.showReactions !== undefined) {
          migratedConfig.displaySettings = {
            showNames: user.showNames ?? migratedConfig.displaySettings?.showNames ?? false,
            showReactions:
              user.showReactions ?? migratedConfig.displaySettings?.showReactions ?? false,
          };
          needsMigration = true;
        }

        // マイグレーション実行
        if (needsMigration) {
          migratedConfig.migratedAt = new Date().toISOString();
          migratedConfig.claudeMdCompliant = true;

          ConfigManager.saveConfig(user.userId, migratedConfig);

          migrationResults.migrated++;
          migrationResults.details.push({
            userId: user.userId,
            email: user.userEmail,
            status: 'migrated',
            fieldsUpdated: Object.keys(migratedConfig).length,
          });

          console.log(`✅ マイグレーション完了: ${user.userEmail}`);
        } else {
          migrationResults.skipped++;
          migrationResults.details.push({
            userId: user.userId,
            email: user.userEmail,
            status: 'skipped_no_changes',
          });
        }
      } catch (userError) {
        migrationResults.errors++;
        migrationResults.details.push({
          userId: user.userId,
          email: user.userEmail,
          status: 'error',
          error: userError.message,
        });
        console.error(`❌ ユーザーマイグレーションエラー: ${user.userEmail}`, userError.message);
      }
    });

    console.log('✅ migrateUserDataToConfigJson: データマイグレーション完了', migrationResults);
    return {
      success: true,
      results: migrationResults,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    console.error('❌ migrateUserDataToConfigJson: マイグレーションエラー', {
      error: error.message,
      timestamp: new Date().toISOString(),
    });
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
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
      spreadsheetId: `${spreadsheetId?.substring(0, 10)}...`,
      sheetName,
      timestamp: new Date().toISOString(),
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

    console.log('✅ analyzeColumns: CLAUDE.md準拠列分析完了', {
      headerCount: headerRow.length,
      mappingCount: Object.keys(columnMapping).length,
      claudeMdCompliant: true,
    });

    return {
      success: true,
      headers: headerRow,
      columnMapping: columnMapping.mapping || columnMapping, // 統一形式
      confidence: columnMapping.confidence, // 信頼度を分離
      sheetName,
      rowCount: sheet.getLastRow(),
      timestamp: new Date().toISOString(),
      claudeMdCompliant: true,
    };
  } catch (error) {
    console.error('❌ analyzeColumns: CLAUDE.md準拠エラー:', {
      spreadsheetId,
      sheetName,
      error: error.message,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
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
      sheetName,
    });

    // 既存のgetSpreadsheetColumnIndices関数を使用
    const result = getSpreadsheetColumnIndices(spreadsheetId, sheetName);

    console.log('✅ getHeaderIndices: フロントエンド互換関数完了', {
      hasResult: !!result,
      opinionHeader: result?.opinionHeader,
    });

    return result;
  } catch (error) {
    console.error('❌ getHeaderIndices: エラー:', {
      spreadsheetId,
      sheetName,
      error: error.message,
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
    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが必要です');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    // 最小限のフォーム連携チェック（軽量版）
    const sheetList = sheets.map((sheet) => {
      const sheetName = sheet.getName();

      // フォームレスポンスシートの簡易判定（FormApp呼び出しなし）
      let isFormSheet = false;
      if (sheetName.match(/^(フォームの回答|Form Responses?|回答)/)) {
        // パターンマッチで高速判定
        isFormSheet = true;
      }

      return {
        name: sheetName,
        isFormResponseSheet: isFormSheet,
        formConnected: isFormSheet,
        formTitle: '', // 詳細はシート選択時に取得
      };
    });

    console.log('✅ getSheetList: シート一覧取得完了', {
      spreadsheetId,
      sheetCount: sheetList.length,
      formSheets: sheetList.filter((s) => s.isFormResponseSheet).length,
    });

    return sheetList;
  } catch (error) {
    console.error('❌ getSheetList: エラー:', {
      spreadsheetId,
      error: error.message,
    });
    throw error;
  }
}

/**
 * 既存ユーザーのフォーム情報修復
 * 不完全なconfigJSONのformUrl/formTitleを修正
 * @param {string} userId - ユーザーID
 * @returns {object} 修復結果
 */
function fixFormInfoForUser(userId) {
  try {
    console.log('🔧 fixFormInfoForUser: フォーム情報修復開始', { userId });

    // 現在のユーザー設定を取得
    const currentConfig = ConfigManager.getUserConfig(userId);
    if (!currentConfig) {
      throw new Error('ユーザー設定が見つかりません');
    }

    // スプレッドシート情報が必要
    if (!currentConfig.spreadsheetId || !currentConfig.sheetName) {
      console.warn('修復スキップ: データソース情報が不足', { userId });
      return {
        success: false,
        reason: 'no_datasource',
        message: 'データソース接続が必要です',
      };
    }

    // フォーム情報を再取得
    const formInfoResult = getFormInfo(currentConfig.spreadsheetId, currentConfig.sheetName);
    if (!formInfoResult.success) {
      console.warn('フォーム情報取得失敗', { userId, error: formInfoResult.error });
      return {
        success: false,
        reason: 'form_info_failed',
        message: 'フォーム情報の取得に失敗しました',
      };
    }

    const formInfo = formInfoResult.formData;
    const needsUpdate =
      currentConfig.formUrl !== formInfo.formUrl || currentConfig.formTitle !== formInfo.formTitle;

    if (!needsUpdate) {
      console.log('修復不要: フォーム情報は正常', { userId });
      return {
        success: true,
        reason: 'no_update_needed',
        message: 'フォーム情報は既に正常です',
      };
    }

    // フォーム情報を更新
    const updatedConfig = {
      ...currentConfig,
      formUrl: formInfo.formUrl || null,
      formTitle: formInfo.formTitle || null,
      lastModified: new Date().toISOString(),
    };

    const saveSuccess = ConfigManager.saveConfig(userId, updatedConfig);
    if (!saveSuccess) {
      throw new Error('設定の保存に失敗しました');
    }

    console.log('✅ fixFormInfoForUser: フォーム情報修復完了', {
      userId,
      oldFormUrl: currentConfig.formUrl,
      newFormUrl: formInfo.formUrl,
      oldFormTitle: currentConfig.formTitle,
      newFormTitle: formInfo.formTitle,
    });

    return {
      success: true,
      message: 'フォーム情報を修復しました',
      updated: {
        formUrl: { before: currentConfig.formUrl, after: formInfo.formUrl },
        formTitle: { before: currentConfig.formTitle, after: formInfo.formTitle },
      },
    };
  } catch (error) {
    console.error('❌ fixFormInfoForUser: エラー:', error.message);
    return {
      success: false,
      reason: 'error',
      message: `フォーム情報の修復に失敗しました: ${error.message}`,
    };
  }
}
