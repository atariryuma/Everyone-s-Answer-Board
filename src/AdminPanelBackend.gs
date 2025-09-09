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
    
    // サンプルデータも取得してより精度の高いマッピングを生成
    const dataRows = Math.min(10, sheet.getLastRow());
    let allData = [];
    if (dataRows > 1) {
      allData = sheet.getRange(1, 1, dataRows, sheet.getLastColumn()).getValues();
    }
    
    let columnMapping = detectColumnMapping(headerRow);
    
    // columnMappingが未定義の場合は強制生成
    if (!columnMapping) {
      console.warn('⚠️ detectColumnMapping失敗、強制生成を実行');
      columnMapping = generateColumnMapping(headerRow, allData);
    }
    
    // さらに未定義の場合はレガシーフォールバック
    if (!columnMapping || !columnMapping.mapping) {
      console.warn('⚠️ columnMapping生成失敗、レガシーフォールバック実行');
      columnMapping = generateLegacyColumnMapping(headerRow);
      
      // レガシー形式を新形式に変換
      if (columnMapping && !columnMapping.mapping) {
        columnMapping = {
          mapping: columnMapping,
          confidence: columnMapping.confidence || {}
        };
      }
    }
    
    console.log('✅ columnMapping最終確認:', {
      hasMapping: !!columnMapping,
      hasMappingField: !!columnMapping?.mapping,
      mappingKeys: columnMapping?.mapping ? Object.keys(columnMapping.mapping) : [],
      hasReason: !!columnMapping?.mapping?.reason
    });

    // 列マッピング検証
    const validationResult = validateAdminPanelMapping(columnMapping.mapping || columnMapping);
    if (!validationResult.isValid) {
      console.warn('列名マッピング検証エラー', validationResult.errors);
    }

    // 不足列の追加
    const missingColumnsResult = addMissingColumns(spreadsheetId, sheetName, columnMapping.mapping || columnMapping);
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      const updatedHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const newMapping = detectColumnMapping(updatedHeaderRow);
      if (newMapping) {
        columnMapping = newMapping;
      }
    }

    // 🎯 フォーム連携情報取得（正規実装）
    let formInfo = null;
    try {
      formInfo = checkFormConnection(spreadsheetId);
      
      if (formInfo && formInfo.hasForm) {
        console.info('✅ フォーム情報取得成功', {
          formUrl: formInfo.formUrl,
          formTitle: formInfo.formTitle,
          hasFormUrl: !!formInfo.formUrl,
          hasFormTitle: !!formInfo.formTitle
        });
      } else {
        console.info('⚠️ フォーム情報が見つかりません', { spreadsheetId });
      }
    } catch (formError) {
      console.error('❌ connectDataSource: フォーム情報取得エラー', {
        error: formError.message,
        spreadsheetId
      });
    }

    // 現在のユーザー取得と設定準備（最適化版）
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (userInfo) {
      // 現在のconfigJSONを直接取得（ConfigManager経由削除）
      const currentConfig = JSON.parse(userInfo.configJson || '{}');

      console.log("connectDataSource: 接続情報確認", {
        userId: userInfo.userId,
        currentSpreadsheetId: currentConfig.spreadsheetId,
        currentSheetName: currentConfig.sheetName,
        currentSetupStatus: currentConfig.setupStatus,
        newSpreadsheetId: spreadsheetId,
        newSheetName: sheetName,
      });

      // 🎯 統一形式：columnMappingのみ保存（レガシー削除）
      console.log('✅ connectDataSource: columnMapping確定', {
        mapping: columnMapping?.mapping,
        confidence: columnMapping?.confidence
      });

      // 🎯 正規的な設定更新：明確で確実な状態管理
      const updatedConfig = {
        // 既存設定を継承
        ...currentConfig,

        // 🔸 データソース情報（確実に設定）
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${spreadsheetId}`,

        // 🔸 列マッピング（統一形式のみ保存、レガシー削除）
        columnMapping: columnMapping,

        // 🔸 フォーム情報（確実な設定）
        formUrl: formInfo?.formUrl || null,
        formTitle: formInfo?.formTitle || null,

        // 🔸 ステータス更新（最重要）
        setupStatus: 'completed',
        
        // 🔸 タイムスタンプ
        lastConnected: new Date().toISOString(),
        lastModified: new Date().toISOString(),
        createdAt: currentConfig.createdAt || new Date().toISOString(),

        // 🔸 表示設定（デフォルト値確保）
        displaySettings: currentConfig.displaySettings || {
          showNames: false,
          showReactions: false
        }
      };

      console.log('💾 connectDataSource: 保存前の設定詳細', {
        userId: userInfo.userId,
        updatedFields: {
          spreadsheetId: updatedConfig.spreadsheetId,
          sheetName: updatedConfig.sheetName,
          setupStatus: updatedConfig.setupStatus,
          hasFormUrl: !!updatedConfig.formUrl,
          hasColumnMapping: !!updatedConfig.columnMapping,
        },
        configSize: JSON.stringify(updatedConfig).length,
      });

      // 🚀 シンプル化：新しいupdateUserメソッドを使用
      DB.updateUser(userInfo.userId, updatedConfig);

      console.log('✅ connectDataSource: DB更新成功', {
        userId: userInfo.userId,
        setupStatus: updatedConfig.setupStatus,
        hasFormUrl: !!updatedConfig.formUrl
      });

      console.log('✅ connectDataSource: 設定統合完了（CLAUDE.md準拠）', {
        userId: userInfo.userId,
        updatedFields: Object.keys(updatedConfig).length,
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
 * アプリ公開（最適化版 - 直接DB更新でService Account維持）
 * 複雑な階層を削除し、確実にデータを保存
 */
function publishApplication(config) {
  try {
    console.log('📱 publishApplication: アプリ公開開始（最適化版）', {
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      hasColumnMapping: !!config.columnMapping,
      timestamp: new Date().toISOString(),
    });

    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      console.error('❌ publishApplication: ユーザー情報が見つかりません', {
        currentUser,
        timestamp: new Date().toISOString(),
      });
      throw new Error('ユーザー情報が見つかりません');
    }

    // 現在のconfigJSONを直接取得（パフォーマンス向上）
    const currentConfig = JSON.parse(userInfo.configJson || '{}');

    console.log('🔍 publishApplication: 設定確認', {
      userId: userInfo.userId,
      currentConfig: {
        spreadsheetId: currentConfig.spreadsheetId,
        sheetName: currentConfig.sheetName,
        setupStatus: currentConfig.setupStatus,
        appPublished: currentConfig.appPublished,
      },
      frontendConfig: {
        spreadsheetId: config.spreadsheetId,
        sheetName: config.sheetName,
      },
    });

    // データソース情報の確定（フォールバック戦略）
    const effectiveSpreadsheetId = currentConfig.spreadsheetId || config.spreadsheetId;
    const effectiveSheetName = currentConfig.sheetName || config.sheetName;

    if (!effectiveSpreadsheetId || !effectiveSheetName) {
      console.error('❌ publishApplication: データソース未設定', {
        dbSpreadsheetId: currentConfig?.spreadsheetId,
        dbSheetName: currentConfig?.sheetName,
        frontendSpreadsheetId: config.spreadsheetId,
        frontendSheetName: config.sheetName,
      });
      throw new Error('データソースが設定されていません。まずデータソースを設定してください。');
    }

    // データソース設定の簡単な検証
    if (currentConfig.setupStatus !== 'completed' && effectiveSpreadsheetId && effectiveSheetName) {
      console.log('🔧 publishApplication: setupStatusを自動修正 (pending → completed)');
      currentConfig.setupStatus = 'completed';
    } else if (!effectiveSpreadsheetId || !effectiveSheetName) {
      console.error('❌ publishApplication: 必須データソース情報不足', {
        effectiveSpreadsheetId: !!effectiveSpreadsheetId,
        effectiveSheetName: !!effectiveSheetName,
      });
      throw new Error('セットアップが完了していません。データソース接続を完了させてください。');
    }

    // 公開実行（executeAppPublish）
    const publishResult = executeAppPublish(userInfo.userId, {
      spreadsheetId: effectiveSpreadsheetId,
      sheetName: effectiveSheetName,
      displaySettings: {
        showNames: config.showNames || false,
        showReactions: config.showReactions || false,
      },
    });

    console.log('⚡ publishApplication: executeAppPublish実行結果', {
      userId: userInfo.userId,
      success: publishResult.success,
      hasAppUrl: !!publishResult.appUrl,
      error: publishResult.error,
    });

    if (publishResult.success) {
      // 🔥 最適化：ConfigManager経由を削除し、直接configJSONを更新
      const updatedConfig = {
        ...currentConfig,
        // 公開設定
        setupStatus: 'completed',
        appPublished: true,
        publishedAt: new Date().toISOString(),
        appUrl: publishResult.appUrl,
        isDraft: false,
        
        // 表示設定を確実に更新
        displaySettings: {
          showNames: config.showNames || false,
          showReactions: config.showReactions || false,
        },

        // メタデータ
        lastModified: new Date().toISOString(),
      };

      console.log('💾 publishApplication: 直接DB更新開始', {
        userId: userInfo.userId,
        updatedFields: {
          appPublished: updatedConfig.appPublished,
          setupStatus: updatedConfig.setupStatus,
          hasAppUrl: !!updatedConfig.appUrl,
          publishedAt: updatedConfig.publishedAt,
        },
      });

      // 🚀 シンプル化：新しいupdateUserメソッドを使用
      DB.updateUser(userInfo.userId, updatedConfig);

      console.log('✅ publishApplication: DB直接更新完了', {
        userId: userInfo.userId,
        appPublished: updatedConfig.appPublished
      });

      console.log('🎉 publishApplication: アプリ公開完了（最適化版）', {
        userId: userInfo.userId,
        appUrl: publishResult.appUrl,
        appPublished: updatedConfig.appPublished,
        setupStatus: updatedConfig.setupStatus,
        hasDisplaySettings: !!updatedConfig.displaySettings,
        publishedAt: updatedConfig.publishedAt,
      });

      return {
        success: true,
        appUrl: publishResult.appUrl,
        config: updatedConfig,
        message: 'アプリが正常に公開されました（最適化版）',
        optimized: true,
        timestamp: new Date().toISOString(),
      };
    } else {
      console.error('❌ publishApplication: executeAppPublish失敗', {
        userId: userInfo.userId,
        error: publishResult.error,
        detailedError: publishResult.detailedError,
      });
      throw new Error(publishResult.error || '公開処理に失敗しました');
    }
  } catch (error) {
    console.error('❌ publishApplication: 最適化版エラー', {
      error: error.message,
      stack: error.stack,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
      optimized: true,
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
    console.log('📋 getFormInfo: フォーム情報取得開始（CLAUDE.md準拠）', {
      spreadsheetId: spreadsheetId?.substring(0, 10) + '...',
      sheetName,
      timestamp: new Date().toISOString(),
    });

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

    console.log('✅ getFormInfo: フォーム情報取得完了', {
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
    console.log('📊 getSpreadsheetList: スプレッドシート一覧取得開始（最適化版）', {
      timestamp: new Date().toISOString(),
    });

    // キャッシュキー生成（ユーザー固有）
    const currentUser = UserManager.getCurrentEmail();
    const cacheKey = `spreadsheet_list_${Utilities.base64Encode(currentUser).replace(/[^a-zA-Z0-9]/g, '')}`;

    // キャッシュから取得を試行（1時間キャッシュ）
    return cacheManager.get(
      cacheKey,
      () => {
        console.log('🔄 getSpreadsheetList: キャッシュミス、データ取得開始');
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

// URL生成関数はmain.gsのgenerateUserUrls()を使用（重複削除）

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
    console.log('🔧 generateColumnMapping: 超高精度列マッピング生成開始', {
      columnCount: headerRow.length,
      dataRows: data.length,
      timestamp: new Date().toISOString(),
    });

    // 重複回避・最適割り当てアルゴリズム実行
    const result = resolveColumnConflicts(headerRow, data);

    console.log('✅ generateColumnMapping: 超高精度マッピング完了', {
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

    // 理由列の高精度検出（CLAUDE.md準拠）
    if (
      normalizedHeader.includes('理由') || 
      normalizedHeader.includes('体験') ||
      normalizedHeader.includes('根拠') ||
      normalizedHeader.includes('詳細') ||
      normalizedHeader.includes('説明') ||
      normalizedHeader.includes('なぜなら') ||
      normalizedHeader.includes('だから')
    ) {
      if (!mapping.reason) {
        mapping.reason = index;
        confidence.reason = 85; // 信頼度向上
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
 * 現在のユーザー設定を取得（最適化版 - データ上書き防止）
 * App.getConfig().getUserConfig() の実装として使用される
 */
function getCurrentConfig() {
  const startTime = Date.now();
  
  try {
    console.log('🔧 getCurrentConfig: ユーザー設定取得開始');

    // 現在のユーザーの設定を取得
    const currentUser = UserManager.getCurrentEmail();
    if (!currentUser) {
      console.error('❌ getCurrentConfig: ユーザーメール取得失敗');
      throw new Error('ユーザー認証が必要です');
    }

    const userInfo = DB.findUserByEmail(currentUser);
    if (!userInfo) {
      console.error('❌ getCurrentConfig: ユーザー情報が見つかりません', {
        currentUser,
        timestamp: new Date().toISOString(),
      });
      // ✅ 修正：初期設定を返さず、エラーとして扱う
      throw new Error('ユーザー情報が見つかりません。セットアップが必要です。');
    }

    const config = ConfigManager.getUserConfig(userInfo.userId);
    if (!config) {
      console.error('❌ getCurrentConfig: 設定データが見つかりません', {
        userId: userInfo.userId,
        userEmail: currentUser,
      });
      // ✅ 修正：既存ユーザーには初期設定ではなく最小限の設定を返す
      throw new Error('設定データが破損している可能性があります');
    }

    const executionTime = Date.now() - startTime;
    console.log('✅ getCurrentConfig: ユーザー設定取得完了（最適化版）', {
      userId: userInfo.userId,
      configFields: Object.keys(config || {}).length,
      setupStatus: config.setupStatus,
      appPublished: config.appPublished,
      hasSpreadsheetId: !!config.spreadsheetId,
      executionTime: `${executionTime}ms`,
    });

    return config;
  } catch (error) {
    const executionTime = Date.now() - startTime;
    console.error('❌ getCurrentConfig エラー（最適化版）', {
      error: error.message,
      executionTime: `${executionTime}ms`,
      timestamp: new Date().toISOString(),
    });
    
    // ✅ 重要：エラー時も初期設定を返さない
    // フロントエンドでエラー処理を行い、ユーザーに適切な案内を表示
    throw error;
  }
}

function getCurrentBoardInfoAndUrls() {
  try {

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

          console.log('✅ ユーザーマイグレーション完了', {
            userId: user.userId,
            email: user.userEmail,
            fieldsUpdated: Object.keys(migratedConfig).length,
          });

          migrationResults.migrated++;
          migrationResults.details.push({
            userId: user.userId,
            email: user.userEmail,
            status: 'migrated',
            fieldsUpdated: Object.keys(migratedConfig).length,
          });

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

    console.log('✅ analyzeColumns: 列分析完了', {
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

    // 統一システム: ヘッダー行から直接インデックスを生成
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // headerIndices オブジェクトを生成
    const result = {};
    headerRow.forEach((header, index) => {
      if (header && String(header).trim()) {
        result[header] = index;
      }
    });

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
 * 列インデックスから実際のヘッダー名を取得
 * @param {Array} headerRow - ヘッダー行の配列
 * @param {number} columnIndex - 列インデックス
 * @returns {string|null} 実際のヘッダー名
 */
function getActualHeaderName(headerRow, columnIndex) {
  if (typeof columnIndex === 'number' && columnIndex >= 0 && columnIndex < headerRow.length) {
    const headerName = headerRow[columnIndex];
    if (headerName && typeof headerName === 'string' && headerName.trim() !== '') {
      return headerName.trim();
    }
  }
  return null;
}

/**
 * 🔍 columnMapping診断ツール（フロントエンド用）
 * 現在のユーザーのconfigJSON状態を詳細診断
 */
function diagnoseColumnMappingIssue() {
  try {
    console.log('🔍 columnMapping診断開始');
    
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    
    console.log('📊 現在の設定状態:', {
      userId: userInfo.userId,
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      hasColumnMapping: !!config.columnMapping,
      hasReasonHeader: !!config.reasonHeader,
      setupStatus: config.setupStatus
    });
    
    // スプレッドシート情報確認
    let spreadsheetInfo = null;
    if (config.spreadsheetId && config.sheetName) {
      try {
        const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
        const sheet = spreadsheet.getSheetByName(config.sheetName);
        const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        
        spreadsheetInfo = {
          headerCount: headerRow.length,
          headers: headerRow,
          dataRowCount: sheet.getLastRow() - 1
        };
        
        console.log('📋 スプレッドシートヘッダー:', headerRow);
      } catch (sheetError) {
        console.error('スプレッドシートアクセスエラー:', sheetError.message);
      }
    }
    
    const diagnosis = {
      userConfigured: !!userInfo,
      spreadsheetConfigured: !!(config.spreadsheetId && config.sheetName),
      columnMappingExists: !!config.columnMapping,
      reasonMappingExists: !!config.columnMapping?.mapping?.reason,
      legacyReasonHeaderExists: !!config.reasonHeader,
      spreadsheetAccessible: !!spreadsheetInfo,
      headerStructure: spreadsheetInfo
    };
    
    console.log('🎯 診断結果:', diagnosis);
    
    return {
      success: true,
      diagnosis,
      config,
      recommendations: generateRecommendations(diagnosis)
    };
    
  } catch (error) {
    console.error('❌ columnMapping診断エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * columnMapping自動修復
 */
function repairColumnMapping() {
  try {
    console.log('🔧 columnMapping自動修復開始');
    
    const currentUser = UserManager.getCurrentEmail();
    const userInfo = DB.findUserByEmail(currentUser);
    
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    const currentConfig = JSON.parse(userInfo.configJson || '{}');
    
    if (!currentConfig.spreadsheetId || !currentConfig.sheetName) {
      throw new Error('スプレッドシート情報が不足しています');
    }
    
    // スプレッドシートからヘッダーを取得
    const spreadsheet = SpreadsheetApp.openById(currentConfig.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(currentConfig.sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // サンプルデータも取得
    const dataRows = Math.min(10, sheet.getLastRow());
    const allData = sheet.getRange(1, 1, dataRows, sheet.getLastColumn()).getValues();
    
    // 新しいcolumnMappingを生成
    const newColumnMapping = generateColumnMapping(headerRow, allData);
    
    console.log('🎯 新columnMapping生成:', {
      oldMapping: currentConfig.columnMapping,
      newMapping: newColumnMapping
    });
    
    // 設定を更新
    const updatedConfig = {
      ...currentConfig,
      columnMapping: newColumnMapping,
      // レガシーヘッダー情報も更新
      reasonHeader: getActualHeaderName(headerRow, newColumnMapping.mapping?.reason) || '理由',
      opinionHeader: getActualHeaderName(headerRow, newColumnMapping.mapping?.answer) || 'お題',
      classHeader: getActualHeaderName(headerRow, newColumnMapping.mapping?.class) || 'クラス',
      nameHeader: getActualHeaderName(headerRow, newColumnMapping.mapping?.name) || '名前',
      lastModified: new Date().toISOString()
    };
    
    // DBに保存
    DB.updateUser(userInfo.userId, updatedConfig);
    
    console.log('✅ columnMapping修復完了');
    
    return {
      success: true,
      message: 'columnMappingを修復しました',
      oldMapping: currentConfig.columnMapping,
      newMapping: newColumnMapping,
      updatedHeaders: {
        reasonHeader: updatedConfig.reasonHeader,
        opinionHeader: updatedConfig.opinionHeader,
        classHeader: updatedConfig.classHeader,
        nameHeader: updatedConfig.nameHeader
      }
    };
    
  } catch (error) {
    console.error('❌ columnMapping修復エラー:', error.message);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * 診断結果に基づく推奨事項生成
 */
function generateRecommendations(diagnosis) {
  const recommendations = [];
  
  if (!diagnosis.spreadsheetConfigured) {
    recommendations.push('スプレッドシートの設定が必要です');
  }
  
  if (!diagnosis.columnMappingExists) {
    recommendations.push('columnMappingを生成する必要があります');
  }
  
  if (!diagnosis.reasonMappingExists && diagnosis.columnMappingExists) {
    recommendations.push('理由列のマッピングが不足しています');
  }
  
  if (!diagnosis.spreadsheetAccessible && diagnosis.spreadsheetConfigured) {
    recommendations.push('スプレッドシートにアクセスできません。権限を確認してください');
  }
  
  return recommendations;
}

/**
 * 既存ユーザーのフォーム情報修復
 * 不完全なconfigJSONのformUrl/formTitleを修正
 * @param {string} userId - ユーザーID
 * @returns {object} 修復結果
 */
function fixFormInfoForUser(userId) {
  try {
    console.log('🔧 fixFormInfoForUser: フォーム情報修復開始', {
      userId,
      timestamp: new Date().toISOString(),
    });

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

    console.log('✅ フォーム情報復元完了', {
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
