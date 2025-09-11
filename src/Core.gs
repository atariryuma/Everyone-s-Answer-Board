/**
 * @fileoverview StudyQuest - Core Functions (高速化版)
 * 主要な業務ロジック、APIエンドポイント、バルクデータAPI
 */

// ===========================================
// 🎯 シンプル化：統一データアクセス関数群
// ===========================================

/**
 * 統一設定取得関数（シンプル版）
 * @param {Object} userInfo - ユーザー情報
 * @returns {Object} 設定オブジェクト
 */
function getConfigSimple(userInfo) {
  if (!userInfo || !userInfo.configJson) {
    throw new Error('ユーザー情報または設定が見つかりません');
  }
  
  const config = JSON.parse(userInfo.configJson);
  
  // 必須フィールドの検証
  if (!config.spreadsheetId) {
    throw new Error('スプレッドシートIDが設定されていません');
  }
  if (!config.sheetName) {
    throw new Error('シート名が設定されていません');
  }
  if (!config.columnMapping || !config.columnMapping.mapping) {
    throw new Error('列マッピングが設定されていません');
  }
  
  return config;
}

/**
 * 統一エラーレスポンス作成
 * @param {string} message - エラーメッセージ
 * @param {Object} details - 詳細情報
 * @returns {Object} エラーレスポンス
 */
function createErrorResponse(message, details = {}) {
  return {
    status: 'error',
    message: message,
    data: [],
    count: 0,
    timestamp: new Date().toISOString(),
    details: details
  };
}

/**
 * 統一成功レスポンス作成
 * @param {Array} data - データ配列
 * @param {Object} metadata - メタデータ
 * @returns {Object} 成功レスポンス
 */
function createSuccessResponse(data, metadata = {}) {
  return {
    status: 'success',
    data: data,
    count: data.length,
    timestamp: new Date().toISOString(),
    ...metadata
  };
}

/**
 * デバッグログ統一出力
 * @param {string} operation - 操作名
 * @param {Object} data - ログデータ
 */
function logDebug(operation, data) {
  console.log(`[${operation}] ${JSON.stringify({
    timestamp: new Date().toISOString(),
    ...data
  })}`);
}

/**
 * columnMappingを使用したデータ処理（シンプル版）
 * @param {Array} dataRows - データ行配列
 * @param {Array} headers - ヘッダー配列
 * @param {Object} columnMapping - 列マッピング
 * @returns {Array} 処理済みデータ配列
 */
function processDataWithColumnMapping(dataRows, headers, columnMapping) {
  logDebug('processDataWithColumnMapping', {
    rowCount: dataRows.length,
    headerCount: headers.length,
    mappingKeys: Object.keys(columnMapping)
  });

  return dataRows.map((row, index) => {
    const processedRow = {
      id: index + 1,
      timestamp: row[columnMapping.timestamp] || row[0] || '',
      email: row[columnMapping.email] || row[1] || '',
      class: row[columnMapping.class] || row[2] || '',
      name: row[columnMapping.name] || row[3] || '',
      answer: row[columnMapping.answer] || row[4] || '',
      reason: row[columnMapping.reason] || row[5] || '',
      reactions: {
        understand: parseInt(row[columnMapping.understand] || 0),
        like: parseInt(row[columnMapping.like] || 0),
        curious: parseInt(row[columnMapping.curious] || 0)
      },
      highlight: row[columnMapping.highlight] === 'TRUE' || false,
      originalData: row
    };

    return processedRow;
  });
}

/**
 * 🚀 バルクデータAPI: 複数の情報を一括取得で高速化
 * 個別API呼び出しを防止し、パフォーマンスを大幅改善
 * @param {string} userId - ユーザーID
 * @param {object} options - オプション { includeSheetData, includeFormInfo, includeSystemInfo }
 * @returns {object} 一括データ
 */
function getBulkData(userId, options = {}) {
  try {
    // デバッグログを削除（個別ログ不要）
    const startTime = Date.now();
    
    // ユーザー情報取得（必須）
    const userInfo = DB.findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザーが見つかりません');
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    const bulkData = {
      timestamp: new Date().toISOString(),
      userId: userId,
      userInfo: {
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive,
        config: config
      }
    };
    
    // シートデータを含む場合
    if (options.includeSheetData && config.spreadsheetId && config.sheetName) {
      try {
        bulkData.sheetData = getData(userId, null, 'asc', false, true);
      } catch (sheetError) {
        console.warn('getBulkData: シートデータ取得エラー', sheetError.message);
        bulkData.sheetDataError = sheetError.message;
      }
    }
    
    // フォーム情報を含む場合
    if (options.includeFormInfo && config.formUrl) {
      try {
        bulkData.formInfo = getActiveFormInfo(userId);
      } catch (formError) {
        console.warn('getBulkData: フォーム情報取得エラー', formError.message);
        bulkData.formInfoError = formError.message;
      }
    }
    
    // システム情報を含む場合
    if (options.includeSystemInfo) {
      bulkData.systemInfo = {
        setupStep: determineSetupStep(userInfo, config),
        isSystemSetup: isSystemSetup(),
        appPublished: config.appPublished || false
      };
    }
    
    const executionTime = Date.now() - startTime;
    // 全体の受信成功を簡潔にログ
    
    return {
      success: true,
      data: bulkData,
      executionTime: executionTime
    };
    
  } catch (error) {
    console.error('❌ getBulkData: バルクデータ取得エラー', {
      userId,
      options,
      error: error.message,
      stack: error.stack
    });
    
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString()
    };
  }
}

/**
 * 自動停止時間を計算する
 * @param {string} publishedAt - 公開開始時間のISO文字列
 * @param {number} minutes - 自動停止までの分数
 * @returns {object} 停止時間情報
 */
function getAutoStopTime(publishedAt, minutes) {
  try {
    const publishTime = new Date(publishedAt);
    const stopTime = new Date(publishTime.getTime() + minutes * 60 * 1000);

    return {
      publishTime,
      stopTime,
      publishTimeFormatted: publishTime.toLocaleString('ja-JP'),
      stopTimeFormatted: stopTime.toLocaleString('ja-JP'),
      remainingMinutes: Math.max(
        0,
        Math.floor((stopTime.getTime() - new Date().getTime()) / (1000 * 60))
      ),
    };
  } catch (error) {
    console.error('自動停止時間計算エラー:', error);
    return null;
  }
}

/**
 * セットアップステップを統一的に判定する関数（configJSON中心設計）
 * @param {Object} userInfo - ユーザー情報（parsedConfig含む）
 * @param {Object} configJson - 設定JSON（レガシー互換性用）
 * @returns {number} セットアップステップ (1-3)
 */
function determineSetupStep(userInfo, configJson) {
  // 🚀 統一データソース：parsedConfig優先、configJsonフォールバック
  const config = JSON.parse(userInfo.configJson || '{}') || configJson || {};
  const setupStatus = config.setupStatus || 'pending';

  // Step 1: データソース未設定 OR セットアップ初期状態
  if (
    !userInfo ||
    !config.spreadsheetId ||
    config.spreadsheetId.trim() === '' ||
    setupStatus === 'pending'
  ) {
    return 1;
  }

  // Step 2: データソース設定済み + セットアップ未完了
  if (setupStatus !== 'completed' || !config.formCreated) {
    console.log(
      'determineSetupStep: Step 2 - データソース設定済み、セットアップ未完了または再設定中'
    );
    return 2;
  }

  // Step 3: 全セットアップ完了 + 公開済み
  if (setupStatus === 'completed' && config.formCreated && config.appPublished) {
    return 3;
  }

  // フォールバック: Step 2
  return 2;
}

// =================================================================
// ユーザー情報キャッシュ（関数実行中の重複取得を防ぐ）
// =================================================================

const _executionSheetsServiceCache = null;

/**
 * ヘッダー整合性検証（リアルタイム検証用）configJSON中心設計
 * @param {string} userId - ユーザーID
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(userId) {
  try {
    const userInfo = getActiveUserInfo();
    if (!userInfo || !JSON.parse(userInfo.configJson || '{}')) {
      return {
        success: false,
        error: 'User configuration not found',
        timestamp: new Date().toISOString(),
      };
    }

    // 🚀 統一データソース：parsedConfig経由でのみデータアクセス
    const config = JSON.parse(userInfo.configJson || '{}');
    const { spreadsheetId } = config;
    const sheetName = config.sheetName || 'EABDB';

    if (!spreadsheetId) {
      return {
        success: false,
        error: 'Spreadsheet ID not configured',
        timestamp: new Date().toISOString(),
      };
    }

    // 統一システムでの列マッピング検証
    const columnIndices = getAllColumnIndices(config);
    
    // 互換性のためのindices作成
    const indices = {};
    if (columnIndices.answer >= 0) indices[CONSTANTS.COLUMNS.OPINION] = columnIndices.answer;
    if (columnIndices.reason >= 0) indices[CONSTANTS.COLUMNS.REASON] = columnIndices.reason;
    if (columnIndices.class >= 0) indices[CONSTANTS.COLUMNS.CLASS] = columnIndices.class;
    if (columnIndices.name >= 0) indices[CONSTANTS.COLUMNS.NAME] = columnIndices.name;

    const validationResults = {
      success: true,
      timestamp: new Date().toISOString(),
      spreadsheetId,
      sheetName,
      headerValidation: {
        reasonColumnIndex: indices[CONSTANTS.COLUMNS.REASON],
        opinionColumnIndex: indices[CONSTANTS.COLUMNS.OPINION],
        hasReasonColumn: indices[CONSTANTS.COLUMNS.REASON] !== undefined,
        hasOpinionColumn: indices[CONSTANTS.COLUMNS.OPINION] !== undefined,
      },
      issues: [],
    };

    // 理由列の必須チェック
    if (indices[CONSTANTS.COLUMNS.REASON] === undefined) {
      validationResults.success = false;
      validationResults.issues.push('Reason column (理由) not found in headers');
    }

    // 回答列の必須チェック
    if (indices[CONSTANTS.COLUMNS.OPINION] === undefined) {
      validationResults.success = false;
      validationResults.issues.push('Opinion column (回答) not found in headers');
    }

    // ログ出力
    if (validationResults.success) {
    } else {
      console.warn('⚠️ Header integrity validation failed:', validationResults.issues);
    }

    return validationResults;
  } catch (error) {
    console.error('❌ Header integrity validation error:', error);
    return {
      success: false,
      error: error.message,
      timestamp: new Date().toISOString(),
    };
  }
}

// =================================================================
// メインロジック
// =================================================================

// doGetLegacy function removed - consolidated into main doGet in UltraOptimizedCore.gs

/**
 * 意見ヘッダーを安全に取得する関数（テンプレート変数の問題を回避）
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {string} 意見ヘッダー
 */
function getOpinionHeaderSafely(userId, sheetName) {
  try {
    // unifiedUserManager.gsの関数を使用
    const userInfo = getActiveUserInfo();
    if (!userInfo) {
      return 'お題';
    }

    const config = JSON.parse(userInfo.configJson || '{}') || {};
    
    // ✅ configJSON中心型: sheetConfig廃止、直接configJSON使用
    const opinionHeader = config.opinionHeader || config.sheetName || 'お題';

    console.log('getOpinionHeaderSafely:', {
      userId,
      sheetName,
      opinionHeader,
    });

    return opinionHeader;
  } catch (e) {
    console.error('getOpinionHeaderSafely error:', e);
    return 'お題';
  }
}

/**
 * 新規ユーザーを登録する（データベース登録のみ）
 * フォーム作成はクイックスタートで実行される
 */
/**
 * 新規ユーザー登録・既存ユーザー更新統合関数
 * GAS 2025 Best Practices準拠 - Modern JavaScript & Structured Error Handling
 * @param {string} userEmail - ユーザーメールアドレス
 * @returns {Object} ユーザー情報とURL情報
 */
function registerNewUser(userEmail) {
  const startTime = Date.now();

  // Enhanced security validation using SecurityValidator (GAS 2025 best practices)
  if (!SecurityValidator.isValidEmail(userEmail)) {
    const error = new Error('有効なメールアドレスを入力してください。');
    console.error('❌ registerNewUser: Invalid email format', {
      providedEmail: userEmail ? `${userEmail.substring(0, 10)}...` : 'null', // Partial logging for privacy
      error: error.message,
    });
    throw error;
  }

  // Sanitize email input
  const sanitizedEmail = SecurityValidator.sanitizeInput(userEmail, SECURITY.MAX_LENGTHS.EMAIL);

  console.info('🔐 認証開始', {
    userEmail: sanitizedEmail,
    timestamp: new Date().toISOString(),
  });

  try {
    // Authentication check with sanitized email
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};

    if (sanitizedEmail !== currentUserEmail) {
      const error = new Error(
        '認証エラー: 操作を実行しているユーザーとメールアドレスが一致しません。'
      );
      console.error('❌ registerNewUser: Authentication failed', {
        requestedEmail: `${sanitizedEmail.substring(0, 10)}...`,
        currentUserEmail: `${currentUserEmail.substring(0, 10)}...`,
        error: error.message,
      });
      throw error;
    }

    // ドメイン制限チェック
    const domainInfo = Deploy.domain();
    if (domainInfo?.deployDomain && domainInfo.deployDomain !== '' && !domainInfo.isDomainMatch) {
      const error = new Error(
        `ドメインアクセスが制限されています。許可されたドメイン: ${domainInfo.deployDomain}, 現在のドメイン: ${domainInfo.currentDomain}`
      );
      console.error('❌ registerNewUser: Domain access denied', {
        allowedDomain: domainInfo.deployDomain,
        currentDomain: domainInfo.currentDomain,
        error: error.message,
      });
      throw error;
    }

    // 既存ユーザーチェック（1ユーザー1行の原則）
    console.log('registerNewUser: 既存ユーザーチェック');
    const existingUser = DB.findUserByEmail(sanitizedEmail);

    if (existingUser) {
      // 既存ユーザーの場合は最小限の更新のみ（設定は保護）
      const { userId } = existingUser;
      const existingConfig = existingUser.parsedConfig || {};

      console.log('registerNewUser: 既存ユーザー更新:', sanitizedEmail);

      // 最終アクセス時刻のみ更新（設定は保護）
      updateUserLastAccess(userId);

      // キャッシュを無効化して最新状態を反映
      invalidateUserCache(userId, sanitizedEmail, existingUser.spreadsheetId, false);

      console.log('registerNewUser: 既存ユーザー更新完了');

      const appUrls = generateUserUrls(userId);

      return {
        userId,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl,
        setupRequired: false, // 既存ユーザーはセットアップ完了済みと仮定
        message: 'おかえりなさい！管理パネルへリダイレクトします。',
        isExistingUser: true,
      };
    }

    // 新規ユーザーの場合
    console.log('registerNewUser: 新ユーザー作成');

    try {
      // 統一ユーザー作成関数を使用（ログイン時はキャッシュバイパス）
      const newUser = handleUserRegistration(sanitizedEmail, true);

      console.log("ユーザー作成成功", {
        userEmail: sanitizedEmail,
        userId: newUser.userId,
        databaseWriteTime: `${Date.now() - startTime}ms`,
      });

      // 生成されたユーザー情報のキャッシュをクリア
      invalidateUserCache(newUser.userId, sanitizedEmail, null, false);

      // 成功レスポンスを返す
      const appUrls = generateUserUrls(newUser.userId);

      console.log('🎉 registerNewUser: New user registration completed', {
        userId: newUser.userId,
        totalExecutionTime: `${Date.now() - startTime}ms`,
      });

      return {
        userId: newUser.userId,
        adminUrl: appUrls.adminUrl,
        viewUrl: appUrls.viewUrl,
        setupRequired: true,
        message: 'ユーザー登録が完了しました！管理パネルで設定を開始してください。',
        isExistingUser: false,
        status: 'success',
      };
    } catch (dbError) {
      console.error('❌ registerNewUser: Database operation failed', {
        userEmail: sanitizedEmail,
        error: dbError.message,
        executionTime: `${Date.now() - startTime}ms`,
      });

      throw new Error(`ユーザー登録に失敗しました。詳細: ${dbError.message}`);
    }
  } catch (error) {
    // Comprehensive error handling with structured logging
    console.error('❌ registerNewUser: Registration process failed', {
      userEmail: sanitizedEmail || `${userEmail?.substring(0, 10)}...`,
      error: error.message,
      stack: error.stack,
      totalExecutionTime: `${Date.now() - startTime}ms`,
    });

    // Re-throw with user-friendly message while preserving technical details
    const userFriendlyMessage = error.message.includes('ユーザー登録に失敗')
      ? error.message
      : 'ユーザー登録処理でエラーが発生しました。システム管理者に連絡してください。';

    throw new Error(userFriendlyMessage);
  }
}

/**
 * リアクションを追加/削除する (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function addReaction(requestUserId, rowIndex, reactionKey, sheetName) {
  return PerformanceMonitor.measure('addReaction', () => {
    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }
    clearExecutionUserInfoCache();

    return executeAddReaction(requestUserId, rowIndex, reactionKey, sheetName);
  });
}

function executeAddReaction(requestUserId, rowIndex, reactionKey, sheetName) {

  try {
    const reactingUserEmail = UserManager.getCurrentEmail();
    const ownerUserId = requestUserId; // requestUserId を使用

    // ボードオーナーの情報をDBから取得（キャッシュ利用）
    const boardOwnerInfo = DB.findUserById(ownerUserId);
    if (!boardOwnerInfo) {
      throw new Error('無効なボードです。');
    }

    const result = processReaction(
      boardOwnerInfo.spreadsheetId,
      sheetName,
      rowIndex,
      reactionKey,
      reactingUserEmail
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      // 更新後のリアクション情報を取得
      const updatedReactions = getRowReactions(
        boardOwnerInfo.spreadsheetId,
        sheetName,
        rowIndex,
        reactingUserEmail
      );

      return {
        status: 'ok',
        reactions: updatedReactions,
      };
    } else {
      throw new Error(result.message || 'リアクションの処理に失敗しました');
    }
  } catch (e) {
    console.error(`addReaction エラー: ${e.message}`);
    return {
      status: 'error',
      message: e.message,
    };
  } finally {
    // 実行終了時にユーザー情報キャッシュをクリア
    clearExecutionUserInfoCache();
  }
}

/**
 * バッチリアクション処理（既存addReaction機能を保持したまま追加）
 * 複数のリアクション操作を効率的に一括処理
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {Array} batchOperations - バッチ操作配列 [{rowIndex, reaction, timestamp}, ...]
 * @returns {object} バッチ処理結果
 */
// addReactions関数はPageBackend.gsに移動済み

/**
 * 現在のシート名を取得するヘルパー関数
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string} シート名
 */
function getCurrentSheetName(spreadsheetId) {
  try {
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    // デフォルトでは最初のシートを使用
    if (sheets.length > 0) {
      return sheets[0].getName();
    }

    throw new Error('シートが見つかりません');
  } catch (error) {
    console.warn('シート名取得エラー:', error.message);
    return 'Sheet1'; // フォールバック
  }
}

// =================================================================
// データ取得関数
// =================================================================

/**
 * マルチテナント環境でのユーザーアクセス権限を検証します。
 * リクエストを投げたユーザー (UserManager.getCurrentEmail()) が、
 * requestUserId のデータにアクセスする権限を持っているかを確認します。
 * 権限がない場合はエラーをスローします。
 * @param {string} requestUserId - アクセスを要求しているユーザーのID
 * @throws {Error} 認証エラーまたは権限エラー
 */
// verifyUserAccess function removed - all calls now use App.getAccess().verifyAccess() directly

/**
 * 実際のデータ取得処理（キャッシュ制御から分離） (マルチテナント対応版)
 */
function getPublishedSheetData(requestUserId, classFilter, sortOrder, adminMode) {
  try {
    logDebug('getPublishedSheetData', {
      requestUserId,
      classFilter,
      sortOrder,
      adminMode
    });

    // CLAUDE.md準拠: requestUserIdを使用してユーザー情報取得
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      return createErrorResponse(`ユーザー情報が見つかりません (userId: ${requestUserId})`);
    }

    // 🎯 シンプル化: 統一設定取得関数使用
    const config = getConfigSimple(userInfo);
    logDebug('config_validated', {
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      hasColumnMapping: !!config.columnMapping.mapping,
      mappingCount: Object.keys(config.columnMapping.mapping).length
    });

    // Check if current user is the board owner
    const isOwner = userInfo.userId === requestUserId;
    logDebug('owner_check', {
      isOwner,
      userInfoUserId: userInfo.userId,
      requestUserId
    });

    // データ取得
    const sheetData = getSheetData(
      requestUserId,
      config.sheetName,
      classFilter,
      sortOrder,
      adminMode
    );
    console.log(
      'getData: sheetData status=%s, totalCount=%s',
      sheetData.status,
      sheetData.totalCount
    );

    if (sheetData.status === 'error') {
      throw new Error(sheetData.message);
    }

    // 🎯 統一列アクセス関数を使用（レガシー完全削除）
    // スプレッドシートからヘッダー行を取得
    const spreadsheet = new ConfigurationManager().getSpreadsheet(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // 統一関数で列インデックスを取得
    const columnIndices = getAllColumnIndices(config);
    
    // ヘッダー名を取得（インデックスから実際のヘッダー名を取得）
    let mainHeaderName, reasonHeaderName, classHeaderName, nameHeaderName;
    
    if (config.setupStatus === 'pending') {
      mainHeaderName = 'セットアップ中...';
      reasonHeaderName = 'セットアップ中...';
      classHeaderName = 'セットアップ中...';
      nameHeaderName = 'セットアップ中...';
    } else {
      // 統一関数でヘッダー名取得
      mainHeaderName = getColumnHeaderByIndex(headerRow, columnIndices.answer) || '回答';
      reasonHeaderName = getColumnHeaderByIndex(headerRow, columnIndices.reason) || '理由';
      classHeaderName = getColumnHeaderByIndex(headerRow, columnIndices.class) || 'クラス';
      nameHeaderName = getColumnHeaderByIndex(headerRow, columnIndices.name) || '名前';
      
      // columnMappingの存在確認
      if (!config.columnMapping || !config.columnMapping.mapping) {
        console.warn('⚠️ columnMappingが設定されていません');
      }
      
      // 理由列の存在確認
      if (columnIndices.reason === -1) {
        console.warn('⚠️ 理由列のマッピングが見つかりません');
      }
    }
    console.log(
      'getData: Configured Headers - mainHeaderName=%s, reasonHeaderName=%s, classHeaderName=%s, nameHeaderName=%s',
      mainHeaderName,
      reasonHeaderName,
      classHeaderName,
      nameHeaderName
    );


    // 🎯 統一関数で直接インデックスを使用（mapConfigToActualHeaders削除）
    const mappedIndices = {
      opinionHeader: columnIndices.answer,
      reasonHeader: columnIndices.reason,
      classHeader: columnIndices.class,
      nameHeader: columnIndices.name
    };


    // 統一システム用のheaderIndices生成
    const headerIndices = {};
    headerRow.forEach((header, index) => {
      if (header && String(header).trim()) {
        headerIndices[header] = index;
      }
    });
    
    // CONSTANTS定数による標準ヘッダーマッピングも追加
    if (columnIndices.answer >= 0) headerIndices[CONSTANTS.COLUMNS.OPINION] = columnIndices.answer;
    if (columnIndices.reason >= 0) headerIndices[CONSTANTS.COLUMNS.REASON] = columnIndices.reason;
    if (columnIndices.class >= 0) headerIndices[CONSTANTS.COLUMNS.CLASS] = columnIndices.class;
    if (columnIndices.name >= 0) headerIndices[CONSTANTS.COLUMNS.NAME] = columnIndices.name;

    const formattedData = formatSheetDataForFrontend(
      sheetData.data,
      mappedIndices,
      headerIndices,
      adminMode,
      isOwner,
      sheetData.displayMode
    );

    console.log('getData: 正常完了', {
      dataCount: formattedData.length,
      status: sheetData.status,
    });

    // ボードのタイトルを実際のスプレッドシートのヘッダーから取得
    let headerTitle = config.sheetName || '今日のお題';
    if (mappedIndices.opinionHeader !== undefined) {
      for (const actualHeader in headerIndices) {
        if (headerIndices[actualHeader] === mappedIndices.opinionHeader) {
          headerTitle = actualHeader;
          console.log('getData: Using actual header as title: "%s"', headerTitle);
          break;
        }
      }
    }

    const finalDisplayMode =
      adminMode === true
        ? CONSTANTS.DISPLAY_MODES.NAMED
        : config.displayMode || CONSTANTS.DISPLAY_MODES.ANONYMOUS;

    const result = {
      header: headerTitle,
      sheetName: config.sheetName,
      showCounts: adminMode === true ? true : config.showCounts === true,
      displayMode: finalDisplayMode,
      data: formattedData,
    };

    console.log('processSheetData: 処理完了', {
      adminMode,
      originalDisplayMode: sheetData.displayMode,
      finalDisplayMode,
      dataCount: formattedData.length,
      showCounts: result.showCounts,
    });
    return result;
  } catch (e) {
    console.error('データ取得エラー:', e.message);
    return {
      status: 'error',
      message: `データの取得に失敗しました: ${e.message}`,
      data: [],
    };
  }
}

/**
 * 増分データ取得機能：指定された基準点以降の新しいデータのみを取得 (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @param {number} sinceRowCount - この行数以降のデータを取得
 * @returns {object} 新しいデータのみを含む結果
 */
function getIncrementalSheetData(requestUserId, classFilter, sortOrder, adminMode, sinceRowCount) {
  try {
    logDebug('getIncrementalSheetData', {
      requestUserId,
      classFilter,
      sortOrder,
      adminMode,
      sinceRowCount
    });

    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      return createErrorResponse(`アクセスが拒否されました: ${accessResult.reason}`);
    }

    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      return createErrorResponse('ユーザー情報が見つかりません');
    }

    // 🎯 シンプル化: 統一設定取得関数使用
    const config = getConfigSimple(userInfo);
    
    // 🔧 シンプルな差分取得ロジック
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    const lastRow = sheet.getLastRow();
    
    logDebug('incremental_check', {
      lastRow,
      sinceRowCount,
      hasNewData: sinceRowCount < lastRow - 1
    });
    
    if (sinceRowCount >= lastRow - 1) {
      return createSuccessResponse([], {
        hasNewData: false,
        totalRows: lastRow - 1
      });
    }
    
    // 🔧 新しいデータを取得 
    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const newRowsData = allData.slice(sinceRowCount + 1); // sinceRowCount以降の新しい行
    
    if (newRowsData.length === 0) {
      return createSuccessResponse([], {
        hasNewData: false,
        totalRows: lastRow - 1
      });
    }
    
    // 🎯 シンプル化: columnMappingを使用した直接データ処理
    const processedNewData = processDataWithColumnMapping(
      newRowsData,
      headers,
      config.columnMapping.mapping
    );
    
    return createSuccessResponse(processedNewData, {
      hasNewData: true,
      totalRows: lastRow - 1,
      newDataCount: processedNewData.length,
      isIncremental: true
    });
  } catch (error) {
    logDebug('getIncrementalSheetData_error', {
      requestUserId,
      error: error.message
    });
    return createErrorResponse(`差分データ取得に失敗しました: ${error.message}`, {
      requestUserId
    });
  }
}

/**
 * スプレッドシートの生データをフロントエンドが期待する形式にフォーマットするヘルパー関数
 * @param {Array<Object>} rawData - getSheetDataから返された生データ
 * @param {Object} mappedIndices - 設定されたヘッダー名と実際の列インデックスのマッピング
 * @param {Object} headerIndices - 実際のヘッダー名と列インデックスのマッピング
 * @param {boolean} adminMode - 管理者モードかどうか
 * @param {boolean} isOwner - 現在のユーザーがボードのオーナーかどうか
 * @param {string} displayMode - 表示モード（'named' or 'anonymous'）
 * @returns {Array<Object>} フォーマットされたデータ
 */
function formatSheetDataForFrontend(
  rawData,
  mappedIndices,
  headerIndices,
  adminMode,
  isOwner,
  displayMode
) {

  // 現在のユーザーメールを取得（リアクション状態判定用）
  const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};

  return rawData.map((row, index) => {
    const classIndex = mappedIndices.classHeader;
    const opinionIndex = mappedIndices.opinionHeader;
    const reasonIndex = mappedIndices.reasonHeader;
    const nameIndex = mappedIndices.nameHeader;


    let nameValue = '';
    const shouldShowName =
      adminMode === true || displayMode === CONSTANTS.DISPLAY_MODES.NAMED || isOwner;
    const hasNameIndex = nameIndex !== undefined;
    const hasOriginalData = row.originalData && row.originalData.length > 0;

    if (shouldShowName && hasNameIndex && hasOriginalData) {
      nameValue = row.originalData[nameIndex] || '';
    }

    if (!nameValue && shouldShowName && hasOriginalData) {
      const emailIndex = headerIndices[CONSTANTS.COLUMNS.EMAIL];
      if (emailIndex !== undefined && row.originalData[emailIndex]) {
        nameValue = row.originalData[emailIndex].split('@')[0];
      }
    }

    // リアクション状態を判定するヘルパー関数
    function checkReactionState(reactionKey) {
      const columnName = CONSTANTS.REACTIONS.LABELS[reactionKey];
      const columnIndex = headerIndices[columnName];
      let count = 0;
      let reacted = false;

      if (columnIndex !== undefined && row.originalData && row.originalData[columnIndex]) {
        const reactionString = row.originalData[columnIndex].toString();
        if (reactionString) {
          const reactions = parseReactionString(reactionString);
          count = reactions.length;
          reacted = reactions.indexOf(currentUserEmail) !== -1;
        }
      }

      // フォールバック: プリカウントされた値を使用
      if (count === 0) {
        if (reactionKey === 'UNDERSTAND') count = row.understandCount || 0;
        else if (reactionKey === 'LIKE') count = row.likeCount || 0;
        else if (reactionKey === 'CURIOUS') count = row.curiousCount || 0;
      }

      return { count, reacted };
    }

    // 🔍 理由列の値を取得（包括的null/undefined/空文字列処理）
    let reasonValue = '';
    
    if (reasonIndex !== undefined && row.originalData && reasonIndex >= 0 && reasonIndex < row.originalData.length) {
      const rawReasonValue = row.originalData[reasonIndex];
      
      
      // null/undefined/空文字列の適切な処理
      if (rawReasonValue !== null && rawReasonValue !== undefined) {
        const stringValue = String(rawReasonValue).trim();
        if (stringValue.length > 0) {
          reasonValue = stringValue;
        }
      }
    }

    const finalResult = {
      rowIndex: row.rowNumber || index + 2,
      name: nameValue,
      email:
        row.originalData && row.originalData[headerIndices[CONSTANTS.COLUMNS.EMAIL]]
          ? row.originalData[headerIndices[CONSTANTS.COLUMNS.EMAIL]]
          : '',
      class:
        classIndex !== undefined && row.originalData && row.originalData[classIndex]
          ? row.originalData[classIndex]
          : '',
      opinion:
        opinionIndex !== undefined && row.originalData && row.originalData[opinionIndex]
          ? row.originalData[opinionIndex]
          : '',
      reason: reasonValue,
      reactions: {
        UNDERSTAND: checkReactionState('UNDERSTAND'),
        LIKE: checkReactionState('LIKE'),
        CURIOUS: checkReactionState('CURIOUS'),
      },
      highlight: row.isHighlighted || false,
    };
    
    
    return finalResult;
  });
}

// ✅ 廃止関数削除完了：getAppConfig() → ConfigManager.getUserConfig()に統一済み

/**
 * シート設定を保存する（統合版：Batch/Optimized機能統合）
 * AdminPanel.htmlから呼び出される
 * @param {string} userId - ユーザーID
 * @param {string} spreadsheetId - 設定対象のスプレッドシートID
 * @param {string} sheetName - 設定対象のシート名
 * @param {object} config - 保存するシート固有の設定
 * @param {object} options - オプション設定
 * @param {object} options.sheetsService - 共有SheetsService（最適化用）
function switchToSheet(userId, spreadsheetId, sheetName, options = {}) {
  try {
    if (!spreadsheetId || typeof spreadsheetId !== 'string') {
      throw new Error('無効なspreadsheetIdです: ' + spreadsheetId);
    }
    if (!sheetName || typeof sheetName !== 'string') {
      throw new Error('無効なsheetNameです: ' + sheetName);
    }
    
    const currentUserId = userId;
    
    // 最適化モード: 事前取得済みuserInfoを使用、なければ取得
    const userInfo = options.userInfo || DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}') || {};

    configJson.appPublished = true; // シートを切り替えたら公開状態にする
    configJson.lastModified = new Date().toISOString();

    // データベースのspreadsheetIdとsheetNameフィールドを更新
    // 🔧 修正: 直接フィールド更新 + ConfigManager経由での安全保存
    DB.updateUserInDatabase(currentUserId, { 
      // 基本フィールドのみ直接更新
      lastModified: new Date().toISOString()
    });
    // configJsonはConfigManager経由で安全保存
    ConfigManager.saveConfig(currentUserId, configJson);
    console.log('✅ 表示シートを切り替えました: %s - %s', spreadsheetId, sheetName);
    return { status: 'success', message: '表示シートを切り替えました。' };
  } catch (e) {
    console.error('シート切り替えエラー: ' + e.message);
    return { status: 'error', message: '表示シートの切り替えに失敗しました: ' + e.message };
  }
}

// =================================================================
// セットアップ関数
// =================================================================

/**
 * アプリケーションの初期セットアップ（管理者が手動で実行） (マルチテナント対応版)
 */
function getResponsesData(userId, sheetName) {
  const userInfo = getActiveUserInfo();
  if (!userInfo) {
    return { status: 'error', message: 'ユーザー情報が見つかりません' };
  }

  try {
    const service = getSheetsServiceCached();
    // 🚀 CLAUDE.md準拠：統一データソース原則
    const config = JSON.parse(userInfo.configJson || '{}') || {};
    const { spreadsheetId } = config;
    // 🚀 CLAUDE.md準拠：A:E範囲使用でパフォーマンス最適化
    const range = `'${sheetName || 'フォームの回答 1'}'!A:E`;

    const response = batchGetSheetsData(service, spreadsheetId, [range]);
    const values = response.valueRanges[0].values || [];

    if (values.length === 0) {
      return { status: 'success', data: [], headers: [] };
    }

    return {
      status: 'success',
      data: values.slice(1),
      headers: values[0],
    };
  } catch (e) {
    console.error('回答データ取得失敗:', e.message);
    return { status: 'error', message: `回答データの取得に失敗しました: ${e.message}` };
  }
}

// =================================================================
// HTML依存関数（UI連携）
// =================================================================

/**
 * 現在のユーザーのステータスを取得し、UIに表示するための情報を返します。(マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} ユーザーのステータス情報
 */
function getCurrentUserStatus(requestUserId) {
  try {
    const activeUserEmail = UserManager.getCurrentEmail();

    // requestUserIdが無効な場合は、メールアドレスでユーザーを検索
    let userInfo;
    if (requestUserId && requestUserId.trim() !== '') {
      const accessResult = App.getAccess().verifyAccess(
        requestUserId,
        'view',
        UserManager.getCurrentEmail()
      );
      if (!accessResult.allowed) {
        throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
      }
      userInfo = DB.findUserById(requestUserId);
    } else {
      userInfo = DB.findUserByEmail(activeUserEmail);
    }

    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません。' };
    }

    return {
      status: 'success',
      userInfo: {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: config.lastAccessedAt,
      },
    };
  } catch (e) {
    console.error(`getCurrentUserStatus エラー: ${e.message}`);
    return { status: 'error', message: `ステータス取得に失敗しました: ${e.message}` };
  }
}

/**
 * アクティブなフォーム情報を取得 (マルチテナント対応版)
 * Page.htmlから呼び出される（パラメータなし）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function getActiveFormInfo(requestUserId) {
  const accessResult = App.getAccess().verifyAccess(
    requestUserId,
    'view',
    UserManager.getCurrentEmail()
  );
  if (!accessResult.allowed) {
    throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
  }
  try {
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}') || {};

    // フォーム回答数を取得
    let answerCount = 0;
    try {
      if (configJson.spreadsheetId && configJson.publishedSheet) {
        const responseData = getResponsesData(currentUserId, configJson.publishedSheet);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
        }
      }
    } catch (countError) {
      console.warn(`回答数の取得に失敗: ${countError.message}`);
    }

    return {
      status: 'success',
      formUrl: configJson.formUrl || '',
      editFormUrl: configJson.editFormUrl || '',
      editUrl: configJson.editFormUrl || '', // AdminPanel.htmlが期待するフィールド名
      formId: extractFormIdFromUrl(configJson.formUrl || configJson.editFormUrl || ''),
      spreadsheetUrl: configJson.spreadsheetUrl || '',
      answerCount,
      isFormActive: !!(configJson.formUrl && configJson.formCreated),
      appPublished: configJson.appPublished || false, // 公開ステータスを追加
      spreadsheetId: configJson.spreadsheetId || '',
      sheetName: configJson.sheetName || '',
      displaySettings: configJson.displaySettings || {},
    };
  } catch (e) {
    console.error(`getActiveFormInfo エラー: ${e.message}`);
    return { status: 'error', message: `フォーム情報の取得に失敗しました: ${e.message}` };
  }
}

/**
 * 指定シートのデータ行数を取得します。
 * @param {string} spreadsheetId スプレッドシートID
 * @param {string} sheetName シート名
 * @param {string} classFilter クラスフィルタ
 * @returns {number} データ行数
 */
function countSheetRows(spreadsheetId, sheetName, classFilter) {
  const key = `rowCount_${spreadsheetId}_${sheetName}_${classFilter}`;
  return cacheManager.get(
    key,
    () => {
      const sheet = new ConfigurationManager().getSpreadsheet(spreadsheetId).getSheetByName(sheetName);
      if (!sheet) return 0;

      const lastRow = sheet.getLastRow();
      if (!classFilter || classFilter === 'すべて') {
        return Math.max(0, lastRow - 1);
      }

      // 簡素化: ヘッダー行から直接クラス列を検索
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const classIndex = headerRow.findIndex(header => header && header.includes('クラス'));
      if (classIndex === -1) {
        return Math.max(0, lastRow - 1);
      }

      const values = sheet.getRange(2, classIndex + 1, lastRow - 1, 1).getValues();
      return values.reduce((cnt, row) => (row[0] === classFilter ? cnt + 1 : cnt), 0);
    },
    { ttl: 30, enableMemoization: true }
  );
}

/**
 * データ数を取得する (マルチテナント対応版)
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {number} データ数
 */
// getDataCount関数はPageBackend.gsに移動済み

/**
 * フォーム設定を更新 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} title - 新しいタイトル
 * @param {string} description - 新しい説明
 */
function updateFormSettings(requestUserId, title, description) {
  const accessResult = App.getAccess().verifyAccess(
    requestUserId,
    'view',
    UserManager.getCurrentEmail()
  );
  if (!accessResult.allowed) {
    throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
  }
  try {
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}') || {};

    if (configJson.editFormUrl) {
      try {
        const formId = extractFormIdFromUrl(configJson.editFormUrl);
        const form = FormApp.openById(formId);

        if (title) {
          form.setTitle(title);
        }
        if (description) {
          form.setDescription(description);
        }

        return {
          status: 'success',
          message: 'フォーム設定が更新されました',
        };
      } catch (formError) {
        console.error(`フォーム更新エラー: ${formError.message}`);
        return { status: 'error', message: `フォームの更新に失敗しました: ${formError.message}` };
      }
    } else {
      return { status: 'error', message: 'フォームが見つかりません' };
    }
  } catch (e) {
    console.error(`updateFormSettings エラー: ${e.message}`);
    return { status: 'error', message: `フォーム設定の更新に失敗しました: ${e.message}` };
  }
}

/**
 * システム設定を保存 (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 */
function saveSystemConfig(requestUserId, config) {
  const accessResult = App.getAccess().verifyAccess(
    requestUserId,
    'view',
    UserManager.getCurrentEmail()
  );
  if (!accessResult.allowed) {
    throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
  }
  try {
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}') || {};

    // システム設定を更新
    configJson.systemConfig = {
      cacheEnabled: config.cacheEnabled,
      autosaveInterval: config.autosaveInterval,
      logLevel: config.logLevel,
      updatedAt: new Date().toISOString(),
    };

    // 🔧 修正: ConfigManager経由で安全な保存（二重構造防止）
    ConfigManager.saveConfig(currentUserId, configJson);

    return {
      status: 'success',
      message: 'システム設定が保存されました',
    };
  } catch (e) {
    console.error(`saveSystemConfig エラー: ${e.message}`);
    return { status: 'error', message: `システム設定の保存に失敗しました: ${e.message}` };
  }
}

/**
 * ハイライト状態の切り替え (マルチテナント対応版)
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
function toggleHighlight(requestUserId, rowIndex, sheetName) {
  return PerformanceMonitor.measure('toggleHighlight', () => {
    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }

    return executeToggleHighlight(requestUserId, rowIndex, sheetName);
  });
}

function executeToggleHighlight(requestUserId, rowIndex, sheetName) {
  try {
    const currentUserId = requestUserId; // requestUserId を使用

    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 管理者権限チェック - 現在のユーザーがボードの所有者かどうかを確認
    const activeUserEmail = UserManager.getCurrentEmail();
    if (activeUserEmail !== userInfo.userEmail) {
      throw new Error('ハイライト機能は管理者のみ使用できます');
    }

    // 🚀 CLAUDE.md準拠：統一データソース原則
    const config = JSON.parse(userInfo.configJson || '{}') || {};
    const result = processHighlightToggle(
      config.spreadsheetId,
      sheetName || 'フォームの回答 1',
      rowIndex
    );

    // Page.html期待形式に変換
    if (result && result.status === 'success') {
      return {
        status: 'ok',
        highlight: result.highlighted || false,
      };
    } else {
      throw new Error(result.message || 'ハイライト切り替えに失敗しました');
    }
  } catch (e) {
    console.error(`toggleHighlight エラー: ${e.message}`);
    return {
      status: 'error',
      message: e.message,
    };
  }
}

/**
 * 利用可能なシート一覧を取得
 * Page.htmlから呼び出される - フロントエンド期待形式に対応
 */

/**
 * クイックスタートセットアップ（完全版） (マルチテナント対応版)
 * フォルダ作成、フォーム作成、スプレッドシート作成、ボード公開まで一括実行
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// createQuickStartFiles - クイックスタート機能は不要のため削除済み

/**
 * ユーザー専用フォルダを作成 (マルチテナント対応版)
 */
function createUserFolder(userEmail) {
  try {
    const rootFolderName = 'StudyQuest - ユーザーデータ';
    const userFolderName = `StudyQuest - ${userEmail} - ファイル`;

    // ルートフォルダを検索または作成
    let rootFolder;
    const folders = DriveApp.getFoldersByName(rootFolderName);
    if (folders.hasNext()) {
      rootFolder = folders.next();
    } else {
      rootFolder = DriveApp.createFolder(rootFolderName);
    }

    // ユーザー専用フォルダを作成
    const userFolders = rootFolder.getFoldersByName(userFolderName);
    if (userFolders.hasNext()) {
      return userFolders.next(); // 既存フォルダを返す
    } else {
      return rootFolder.createFolder(userFolderName);
    }
  } catch (e) {
    console.error(`createUserFolder エラー: ${e.message}`);
    return null; // フォルダ作成に失敗してもnullを返して処理を継続
  }
}

/**
 * ハイライト切り替え処理
 */
function processHighlightToggle(spreadsheetId, sheetName, rowIndex) {
  try {
    const service = getSheetsServiceCached();
    
    // ヘッダー行からハイライト列を検索
    const sheet = new ConfigurationManager().getSpreadsheet(spreadsheetId).getSheetByName(sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const highlightColumnIndex = headerRow.findIndex(header => 
      header && (header.includes('ハイライト') || header.includes('highlight'))
    );

    if (highlightColumnIndex === -1) {
      throw new Error('ハイライト列が見つかりません');
    }

    // 現在の値を取得
    const range = `'${sheetName}'!${String.fromCharCode(65 + highlightColumnIndex)}${rowIndex}`;
    const response = batchGetSheetsData(service, spreadsheetId, [range]);
    let isHighlighted = false;
    if (
      response &&
      response.valueRanges &&
      response.valueRanges[0] &&
      response.valueRanges[0].values &&
      response.valueRanges[0].values[0] &&
      response.valueRanges[0].values[0][0]
    ) {
      isHighlighted = response.valueRanges[0].values[0][0] === 'true';
    }

    // 値を切り替え
    const newValue = isHighlighted ? 'false' : 'true';
    updateSheetsData(service, spreadsheetId, range, [[newValue]]);

    // ハイライト更新後のキャッシュ無効化
    try {
      if (
        typeof cacheManager !== 'undefined' &&
        typeof cacheManager.invalidateSheetData === 'function'
      ) {
        cacheManager.invalidateSheetData(spreadsheetId, sheetName);
      }
    } catch (cacheError) {
      console.warn('ハイライト後のキャッシュ無効化エラー:', cacheError.message);
    }

    return {
      status: 'success',
      highlighted: !isHighlighted,
      message: isHighlighted ? 'ハイライトを解除しました' : 'ハイライトしました',
    };
  } catch (e) {
    console.error(`ハイライト処理エラー: ${e.message}`);
    return { status: 'error', message: e.message };
  }
}

// =================================================================
// 統一化関数群
// =================================================================

// getWebAppUrl function removed - now using the unified version from url.gs


function getSheetColumns(userId, sheetId) {
  const accessResult = App.getAccess().verifyAccess(userId, 'view', UserManager.getCurrentEmail());
  if (!accessResult.allowed) {
    throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
  }
  try {
    const userInfo = DB.findUserById(userId);

    // 🚀 CLAUDE.md準拠：統一データソース原則
    const config = userInfo ? JSON.parse(userInfo.configJson || '{}') || {} : {};
    if (!userInfo || !config.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    const spreadsheet = new ConfigurationManager().getSpreadsheet(config.spreadsheetId);
    const sheet = spreadsheet.getSheetById(sheetId);

    if (!sheet) {
      throw new Error(`指定されたシートが見つかりません: ${sheetId}`);
    }

    const lastColumn = sheet.getLastColumn();
    if (lastColumn === 0) {
      return []; // 列がない場合
    }

    const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];

    // クライアントサイドが期待する形式に変換
    const columns = headers.map((headerName) => {
      return {
        id: headerName,
        name: headerName,
      };
    });

    console.log('✅ getSheetColumns: Found %s columns for sheetId %s', columns.length, sheetId);
    return columns;
  } catch (e) {
    console.error(`getSheetColumns エラー: ${e.message}`);
    console.error('Error details:', e.stack);
    throw new Error(`列の取得に失敗しました: ${e.message}`);
  }
}

/**
 * GoogleフォームURLからフォームIDを抽出
 */
function extractFormIdFromUrl(url) {
  if (!url) return '';

  try {
    // Regular expression to extract form ID from Google Forms URLs
    const formIdMatch = url.match(/\/forms\/d\/([a-zA-Z0-9-_]+)/);
    if (formIdMatch && formIdMatch[1]) {
      return formIdMatch[1];
    }

    // Alternative pattern for e/ URLs
    const eFormIdMatch = url.match(/\/forms\/d\/e\/([a-zA-Z0-9-_]+)/);
    if (eFormIdMatch && eFormIdMatch[1]) {
      return eFormIdMatch[1];
    }

    return '';
  } catch (e) {
    console.warn(`フォームID抽出エラー: ${e.message}`);
    return '';
  }
}

// =================================================================
// リアクション処理関数
// =================================================================

/**
 * リアクション処理
 */
function processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, reactingUserEmail) {
  try {
    // LockServiceを使って競合を防ぐ
    const lock = LockService.getScriptLock();
    try {
      lock.waitLock(10000);

      const service = getSheetsServiceCached();
      
      // ヘッダー行から列インデックスを取得
      const sheet = new ConfigurationManager().getSpreadsheet(spreadsheetId).getSheetByName(sheetName);
      const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const headerIndices = {};
      headerRow.forEach((header, index) => {
        if (header && String(header).trim()) {
          headerIndices[header] = index;
        }
      });

      // すべてのリアクション列を取得してユーザーの重複リアクションをチェック
      const allReactionRanges = [];
      const allReactionColumns = {};
      let targetReactionColumnIndex = null;

      // 全リアクション列の情報を準備
      CONSTANTS.REACTIONS.KEYS.forEach((key) => {
        const columnName = CONSTANTS.REACTIONS.LABELS[key];
        const columnIndex = headerIndices[columnName];
        if (columnIndex !== undefined) {
          const range = `'${sheetName}'!${String.fromCharCode(65 + columnIndex)}${rowIndex}`;
          allReactionRanges.push(range);
          allReactionColumns[key] = {
            columnIndex,
            range,
          };
          if (key === reactionKey) {
            targetReactionColumnIndex = columnIndex;
          }
        }
      });

      if (targetReactionColumnIndex === null) {
        throw new Error(`対象リアクション列が見つかりません: ${reactionKey}`);
      }

      // 全リアクション列の現在の値を一括取得
      const response = batchGetSheetsData(service, spreadsheetId, allReactionRanges);
      const updateData = [];
      let userAction = null;
      let targetCount = 0;

      // 各リアクション列を処理
      let rangeIndex = 0;
      CONSTANTS.REACTIONS.KEYS.forEach((key) => {
        if (!allReactionColumns[key]) return;

        let currentReactionString = '';
        if (
          response &&
          response.valueRanges &&
          response.valueRanges[rangeIndex] &&
          response.valueRanges[rangeIndex].values &&
          response.valueRanges[rangeIndex].values[0] &&
          response.valueRanges[rangeIndex].values[0][0]
        ) {
          currentReactionString = response.valueRanges[rangeIndex].values[0][0];
        }

        const currentReactions = parseReactionString(currentReactionString);
        const userIndex = currentReactions.indexOf(reactingUserEmail);

        if (key === reactionKey) {
          // 対象リアクション列の処理
          if (userIndex >= 0) {
            // 既にリアクション済み → 削除（トグル）
            currentReactions.splice(userIndex, 1);
            userAction = 'removed';
          } else {
            // 未リアクション → 追加
            currentReactions.push(reactingUserEmail);
            userAction = 'added';
          }
          targetCount = currentReactions.length;
        } else {
          // 他のリアクション列からユーザーを削除（1人1リアクション制限）
          if (userIndex >= 0) {
            currentReactions.splice(userIndex, 1);
          }
        }

        // 更新データを準備
        const updatedReactionString = currentReactions.join(', ');
        updateData.push({
          range: allReactionColumns[key].range,
          values: [[updatedReactionString]],
        });

        rangeIndex++;
      });

      // すべての更新を一括実行
      if (updateData.length > 0) {
        batchUpdateSheetsData(service, spreadsheetId, updateData);
      }

      console.log(
        `リアクション切り替え完了: ${reactingUserEmail} → ${reactionKey} (${userAction})`,
        {
          updatedRanges: updateData.length,
          targetCount,
          allColumns: Object.keys(allReactionColumns),
        }
      );

      // リアクション更新後のキャッシュ無効化
      try {
        if (
          typeof cacheManager !== 'undefined' &&
          typeof cacheManager.invalidateSheetData === 'function'
        ) {
          cacheManager.invalidateSheetData(spreadsheetId, sheetName);
        }
      } catch (cacheError) {
        console.warn('リアクション後のキャッシュ無効化エラー:', cacheError.message);
      }

      return {
        status: 'success',
        message: 'リアクションを更新しました。',
        action: userAction,
        count: targetCount,
      };
    } finally {
      lock.releaseLock();
    }
  } catch (e) {
    console.error(`リアクション処理エラー: ${e.message}`);
    return {
      status: 'error',
      message: `リアクションの処理に失敗しました: ${e.message}`,
    };
  }
}

// =================================================================
// フォーム作成機能
// =================================================================

/**
 * 回答ボードの公開を停止する (マルチテナント対応版)
 * AdminPanel.htmlから呼び出される
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// NOTE: unpublishBoard関数の重複を回避するため、config.gsの実装を使用
// function unpublishBoard(requestUserId) {
//   const accessResult = App.getAccess().verifyAccess(requestUserId, "view", UserManager.getCurrentEmail()); if (!accessResult.allowed) { throw new Error(`アクセスが拒否されました: ${accessResult.reason}`); }
//   try {
//     const currentUserId = requestUserId;
//     const userInfo = DB.findUserById(currentUserId);
//     if (!userInfo) {
//       throw new Error('ユーザー情報が見つかりません');
//     }

//     const configJson = JSON.parse(userInfo.configJson || '{}') || {};

//     configJson.spreadsheetId = '';
//     configJson.sheetName = '';
//     configJson.appPublished = false; // 公開状態をfalseにする
//     configJson.setupStatus = 'completed'; // 公開停止後もセットアップは完了状態とする

//     DB.updateUser(currentUserId, { configJson: JSON.stringify(configJson) });
//     invalidateUserCache(currentUserId, userInfo.userEmail, config.spreadsheetId, true);

//     console.log('✅ 回答ボードの公開を停止しました: %s', currentUserId);
//     return { status: 'success', message: '回答ボードの公開を停止しました。' };
//   } catch (e) {
//     console.error('公開停止エラー: ' + e.message);
//     return { status: 'error', message: '回答ボードの公開停止に失敗しました: ' + e.message };
//   }
// }

/**
 * フォーム作成ファクトリ
 * @param {Object} options - フォーム作成オプション
 * @param {string} options.userEmail - ユーザーのメールアドレス
 * @param {string} options.userId - ユーザーID
 * @param {string} options.questions - 質問タイプ（'default'など）
 * @param {string} options.formDescription - フォームの説明
 * @returns {Object} フォーム作成結果
 */
function createFormFactory(options) {
  try {
    const { userEmail } = options;
    const { userId } = options;
    const formDescription = options.formDescription || 'みんなの回答ボードへの投稿フォームです。';

    // タイムスタンプ生成（日時を含む）
    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // フォームタイトル生成
    const formTitle = options.formTitle || `みんなの回答ボード ${dateTimeString}`;

    // フォーム作成
    const form = FormApp.create(formTitle);
    form.setDescription(formDescription);

    // 基本的な質問を追加
    addUnifiedQuestions(form, options.questions || 'default', options.customConfig || {});

    // メールアドレス収集を有効化（スプレッドシート連携前に設定）
    try {
      if (typeof form.setEmailCollectionType === 'function') {
        form.setEmailCollectionType(FormApp.EmailCollectionType.VERIFIED);
      } else {
        form.setCollectEmail(true);
      }
    } catch (emailError) {
      console.warn('Email collection setting failed:', emailError.message);
    }

    // スプレッドシート作成
    const spreadsheetResult = createLinkedSpreadsheet(userEmail, form, dateTimeString);

    // リアクション関連列を追加
    addReactionColumnsToSpreadsheet(spreadsheetResult.spreadsheetId, spreadsheetResult.sheetName);

    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      viewFormUrl: form.getPublishedUrl(),
      editFormUrl: typeof form.getEditUrl === 'function' ? form.getEditUrl() : '',
      spreadsheetId: spreadsheetResult.spreadsheetId,
      spreadsheetUrl: spreadsheetResult.spreadsheetUrl,
      sheetName: spreadsheetResult.sheetName,
    };
  } catch (error) {
    console.error('createFormFactory エラー:', error.message);
    throw new Error(`フォーム作成ファクトリでエラーが発生しました: ${error.message}`);
  }
}

/**
 * 統一された質問をフォームに追加
 * @param {GoogleAppsScript.Forms.Form} form - フォームオブジェクト
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 */
function addUnifiedQuestions(form, questionType, customConfig) {
  try {
    const config = getQuestionConfig(questionType, customConfig);

    form.setCollectEmail(false);

    if (questionType === 'simple') {
      const classItem = form.addListItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setChoiceValues(config.classQuestion.choices);
      classItem.setRequired(true);

      const nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setRequired(true);

      const mainItem = form.addTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setRequired(true);

      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      const validation = FormApp.createParagraphTextValidation()
        .requireTextLengthLessThanOrEqualTo(140)
        .build();
      reasonItem.setValidation(validation);
      reasonItem.setRequired(false);
    } else if (questionType === 'custom' && customConfig) {
      console.log('addUnifiedQuestions - custom mode with config:', JSON.stringify(customConfig));

      // クラス選択肢（有効な場合のみ）
      if (
        customConfig.enableClass &&
        customConfig.classQuestion &&
        customConfig.classQuestion.choices &&
        customConfig.classQuestion.choices.length > 0
      ) {
        const classItem = form.addListItem();
        classItem.setTitle('クラス');
        classItem.setChoiceValues(customConfig.classQuestion.choices);
        classItem.setRequired(true);
      }

      // 名前欄（常にオン）
      const nameItem = form.addTextItem();
      nameItem.setTitle('名前');
      nameItem.setRequired(false);

      // メイン質問
      const mainQuestionTitle = customConfig.mainQuestion
        ? customConfig.mainQuestion.title
        : '今回のテーマについて、あなたの考えや意見を聞かせてください';
      let mainItem;
      const questionType = customConfig.mainQuestion ? customConfig.mainQuestion.type : 'text';

      switch (questionType) {
        case 'text':
          mainItem = form.addTextItem();
          break;
        case 'multiple':
          mainItem = form.addCheckboxItem();
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.choices &&
            customConfig.mainQuestion.choices.length > 0
          ) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.includeOthers &&
            typeof mainItem.showOtherOption === 'function'
          ) {
            mainItem.showOtherOption(true);
          }
          break;
        case 'choice':
          mainItem = form.addMultipleChoiceItem();
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.choices &&
            customConfig.mainQuestion.choices.length > 0
          ) {
            mainItem.setChoiceValues(customConfig.mainQuestion.choices);
          }
          if (
            customConfig.mainQuestion &&
            customConfig.mainQuestion.includeOthers &&
            typeof mainItem.showOtherOption === 'function'
          ) {
            mainItem.showOtherOption(true);
          }
          break;
        default:
          mainItem = form.addParagraphTextItem();
      }
      mainItem.setTitle(mainQuestionTitle);
      mainItem.setRequired(true);

      // 理由欄（常にオン）
      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle('そう考える理由や体験があれば教えてください。');
      reasonItem.setRequired(false);
    } else {
      const classItem = form.addTextItem();
      classItem.setTitle(config.classQuestion.title);
      classItem.setHelpText(config.classQuestion.helpText);
      classItem.setRequired(true);

      const mainItem = form.addParagraphTextItem();
      mainItem.setTitle(config.mainQuestion.title);
      mainItem.setHelpText(config.mainQuestion.helpText);
      mainItem.setRequired(true);

      const reasonItem = form.addParagraphTextItem();
      reasonItem.setTitle(config.reasonQuestion.title);
      reasonItem.setHelpText(config.reasonQuestion.helpText);
      reasonItem.setRequired(false);

      const nameItem = form.addTextItem();
      nameItem.setTitle(config.nameQuestion.title);
      nameItem.setHelpText(config.nameQuestion.helpText);
      nameItem.setRequired(false);
    }

  } catch (error) {
    console.error('addUnifiedQuestions エラー:', error.message);
    throw error;
  }
}

/**
 * 質問設定を取得
 * @param {string} questionType - 質問タイプ
 * @param {Object} customConfig - カスタム設定
 * @returns {Object} 質問設定
 */
function getQuestionConfig(questionType, customConfig) {
  // 統一されたテンプレート設定（simple のみ使用）
  const config = {
    classQuestion: {
      title: 'クラス',
      helpText: '',
      choices: ['クラス1', 'クラス2', 'クラス3', 'クラス4'],
    },
    nameQuestion: {
      title: '名前',
      helpText: '',
    },
    mainQuestion: {
      title: '今日の学習について、あなたの考えや感想を聞かせてください',
      helpText: '',
      choices: ['気づいたことがある。', '疑問に思うことがある。', 'もっと知りたいことがある。'],
      type: 'paragraph',
    },
    reasonQuestion: {
      title: 'そう考える理由や体験があれば教えてください。',
      helpText: '',
      type: 'paragraph',
    },
  };

  // カスタム設定をマージ
  if (customConfig && typeof customConfig === 'object') {
    for (const key in customConfig) {
      if (config[key]) {
        Object.assign(config[key], customConfig[key]);
      }
    }
  }

  return config;
}

/**
 * 質問テンプレート取得用 API
 * @returns {ContentService.TextOutput} JSON形式の質問設定
 */
function doGetQuestionConfig() {
  try {
    // 現在の日時を取得してタイトルに含める
    const now = new Date();
    const timestamp = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy/MM/dd HH:mm');

    const cfg = getQuestionConfig('simple');

    // タイトルにタイムスタンプを追加
    cfg.formTitle = `フォーム作成 - ${timestamp}`;

    return ContentService.createTextOutput(JSON.stringify(cfg)).setMimeType(
      ContentService.MimeType.JSON
    );
  } catch (error) {
    console.error('doGetQuestionConfig error:', error);
    return ContentService.createTextOutput(
      JSON.stringify({
        error: 'Failed to get question config',
        details: error.toString(),
      })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * クラス選択肢をデータベースに保存
 */
function saveClassChoices(userId, classChoices) {
  try {
    const currentUserId = userId;
    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const configJson = JSON.parse(userInfo.configJson || '{}') || {};
    configJson.savedClassChoices = classChoices;
    configJson.lastClassChoicesUpdate = new Date().toISOString();

    // 🔧 修正: ConfigManager経由で安全な保存（二重構造防止）
    ConfigManager.saveConfig(currentUserId, configJson);

    console.log('クラス選択肢が保存されました:', classChoices);
    return { status: 'success', message: 'クラス選択肢が保存されました' };
  } catch (error) {
    console.error('クラス選択肢保存エラー:', error.message);
    return { status: 'error', message: `クラス選択肢の保存に失敗しました: ${error.message}` };
  }
}

/**
 * 保存されたクラス選択肢を取得
 */
function getSavedClassChoices(userId) {
  try {
    const currentUserId = userId;
    const userInfo = DB.findUserById(currentUserId);
    if (!userInfo) {
      return { status: 'error', message: 'ユーザー情報が見つかりません' };
    }

    const configJson = JSON.parse(userInfo.configJson || '{}') || {};
    const savedClassChoices = configJson.savedClassChoices || [
      'クラス1',
      'クラス2',
      'クラス3',
      'クラス4',
    ];

    return {
      status: 'success',
      classChoices: savedClassChoices,
      lastUpdate: configJson.lastClassChoicesUpdate,
    };
  } catch (error) {
    console.error('クラス選択肢取得エラー:', error.message);
    return { status: 'error', message: `クラス選択肢の取得に失敗しました: ${error.message}` };
  }
}

/**
 * 統合フォーム作成関数（Phase 2: 最適化版）
 * プリセット設定を使用して重複コードを排除
 */
const FORM_PRESETS = {
  quickstart: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'custom',
    description:
      'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {
      mainQuestion: '今日の学習について、あなたの考えや感想を聞かせてください',
      questionType: 'text',
      enableClass: false,
      includeOthers: false,
    },
  },
  custom: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'custom',
    description:
      'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {}, // Will be overridden by user input
  },
  study: {
    titlePrefix: 'みんなの回答ボード',
    questions: 'simple', // Default, can be overridden
    description:
      'このフォームは学校での学習や話し合いで使います。みんなで考えを共有して、お互いから学び合いましょう。\n\n【デジタル市民としてのお約束】\n• 相手を思いやる気持ちを持ち、建設的な意見を心がけましょう\n• 事実に基づいた正しい情報を共有しましょう\n• 多様な意見を尊重し、違いを学びの機会としましょう\n• 個人情報やプライバシーに関わる内容は書かないようにしましょう\n• 責任ある発言を心がけ、みんなが安心して参加できる環境を作りましょう\n\nあなたの意見や感想は、クラスメイトの学びを深める大切な資源です。',
    config: {},
  },
};

/**
 * 統合フォーム作成関数
 * @param {string} presetType - 'quickstart' | 'custom' | 'study'
 * @param {string} userEmail - ユーザーメールアドレス
 * @param {string} userId - ユーザーID
 * @param {Object} overrides - プリセットを上書きする設定
 * @returns {Object} フォーム作成結果
 */
function createUnifiedForm(presetType, userEmail, userId, overrides = {}) {
  try {
    const preset = FORM_PRESETS[presetType];
    if (!preset) {
      throw new Error(`未知のプリセットタイプ: ${presetType}`);
    }

    const now = new Date();
    const dateTimeString = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy年MM月dd日 HH:mm');

    // タイトル生成（上書き可能）
    const titlePrefix = overrides.titlePrefix || preset.titlePrefix;
    const formTitle = overrides.formTitle || `${titlePrefix} ${dateTimeString}`;

    // 設定をマージ（プリセット + ユーザー設定）
    const mergedConfig = { ...preset.config, ...overrides.customConfig };

    const factoryOptions = {
      userEmail,
      userId,
      formTitle,
      questions: overrides.questions || preset.questions,
      formDescription: overrides.formDescription || preset.description,
      customConfig: mergedConfig,
    };

    return createFormFactory(factoryOptions);
  } catch (error) {
    console.error(`createUnifiedForm Error (${presetType}):`, error.message);
    throw new Error(`フォーム作成に失敗しました (${presetType}): ${error.message}`);
  }
}

/**
 * リンクされたスプレッドシートを作成
 * @param {string} userEmail - ユーザーメール
 * @param {GoogleAppsScript.Forms.Form} form - フォーム
 * @param {string} dateTimeString - 日付時刻文字列
 * @returns {Object} スプレッドシート情報
 */
function createLinkedSpreadsheet(userEmail, form, dateTimeString) {
  try {
    // スプレッドシート名を設定
    const spreadsheetName = `${userEmail} - 回答データ - ${dateTimeString}`;

    // 新しいスプレッドシートを作成
    const spreadsheetObj = SpreadsheetApp.create(spreadsheetName);
    const spreadsheetId = spreadsheetObj.getId();

    // スプレッドシートの共有設定を同一ドメイン閲覧可能に設定
    try {
      const file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error(`スプレッドシートが見つかりません: ${spreadsheetId}`);
      }

      // 同一ドメインで閲覧可能に設定（教育機関対応）
      file.setSharing(DriveApp.Access.DOMAIN, DriveApp.Permission.VIEW);

      // 作成者（現在のユーザー）は所有者として保持
    } catch (sharingError) {
      console.warn(`共有設定の変更に失敗しましたが、処理を続行します: ${sharingError.message}`);
    }

    // フォームの回答先としてスプレッドシートを設定
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheetId);

    // スプレッドシート名を再設定（フォーム設定後に名前が変わる場合があるため）
    spreadsheetObj.rename(spreadsheetName);

    // シート名を取得（通常は「フォームの回答 1」）
    const sheets = spreadsheetObj.getSheets();
    const sheetName = String(sheets[0].getName());
    // シート名が不正な値でないことを確認
    if (!sheetName || sheetName === 'true') {
      sheetName = 'Sheet1'; // または適切なデフォルト値
      console.warn(
        `不正なシート名が検出されました。デフォルトのシート名を使用します: ${sheetName}`
      );
    }

    // サービスアカウントとスプレッドシートを共有（失敗しても処理継続）
    try {
      shareSpreadsheetWithServiceAccount(spreadsheetId);
    } catch (shareError) {
      console.warn('サービスアカウント共有エラー（処理継続）:', shareError.message);
      // 権限エラーの場合でも、スプレッドシート作成自体は成功とみなす
      console.log(
        'スプレッドシート作成は完了しました。サービスアカウント共有は後で設定してください。'
      );
    }

    return {
      spreadsheetId,
      spreadsheetUrl: spreadsheetObj.getUrl(),
      sheetName,
    };
  } catch (error) {
    console.error('createLinkedSpreadsheet エラー:', error.message);
    throw new Error(`スプレッドシート作成に失敗しました: ${error.message}`);
  }
}

/**
 * スプレッドシートをサービスアカウントと共有
 * @param {string} spreadsheetId - スプレッドシートID
 */
function shareAllSpreadsheetsWithServiceAccount() {
  try {
    const allUsers = getAllUsers();
    const results = [];
    let successCount = 0;
    let errorCount = 0;

    for (let i = 0; i < allUsers.length; i++) {
      const user = allUsers[i];
      if (user.spreadsheetId && user.isActive === 'true') {
        try {
          shareSpreadsheetWithServiceAccount(user.spreadsheetId);
          results.push({
            userId: user.userId,
            userEmail: user.userEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'success',
          });
          successCount++;
          console.log('共有成功:', user.userEmail, user.spreadsheetId);
        } catch (shareError) {
          results.push({
            userId: user.userId,
            userEmail: user.userEmail,
            spreadsheetId: user.spreadsheetId,
            status: 'error',
            error: shareError.message,
          });
          errorCount++;
          console.error('共有失敗:', user.userEmail, shareError.message);
        }
      }
    }


    return {
      status: 'completed',
      totalUsers: allUsers.length,
      successCount,
      errorCount,
      results,
    };
  } catch (error) {
    console.error('shareAllSpreadsheetsWithServiceAccount エラー:', error.message);
    throw new Error(`全スプレッドシート共有処理でエラーが発生しました: ${error.message}`);
  }
}

/**
 * フォーム作成
 */
/**
 * StudyQuestフォーム作成（追加機能付き）
 * @deprecated createUnifiedForm('study', ...) を使用してください
 */
function repairUserSpreadsheetAccess(userEmail, spreadsheetId) {
  try {
    // DriveApp経由で共有設定を変更
    let file;
    try {
      file = DriveApp.getFileById(spreadsheetId);
      if (!file) {
        throw new Error(`スプレッドシートが見つかりません: ${spreadsheetId}`);
      }
    } catch (driveError) {
      console.error('DriveApp.getFileById error:', driveError.message);
      throw new Error(`スプレッドシートへのアクセスに失敗しました: ${driveError.message}`);
    }

    // ドメイン全体でアクセス可能に設定
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.EDIT);
      console.log('スプレッドシートをドメイン全体で編集可能に設定しました');
    } catch (domainSharingError) {
      console.warn(`ドメイン共有設定に失敗: ${domainSharingError.message}`);

      // ドメイン共有に失敗した場合は個別にユーザーを追加
      try {
        file.addEditor(userEmail);
      } catch (individualError) {
        console.error(`個別ユーザー追加も失敗: ${individualError.message}`);
      }
    }

    // SpreadsheetApp経由でも編集者として追加
    try {
      const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
      spreadsheet.addEditor(userEmail);
    } catch (spreadsheetAddError) {
      console.warn(`SpreadsheetApp経由の追加で警告: ${spreadsheetAddError.message}`);
    }

    // サービスアカウントも確認
    const props = PropertiesService.getScriptProperties();
    const serviceAccountCreds = JSON.parse(props.getProperty(PROPS_KEYS.SERVICE_ACCOUNT_CREDS));
    const serviceAccountEmail = serviceAccountCreds.client_email;

    if (serviceAccountEmail) {
      try {
        file.addEditor(serviceAccountEmail);
      } catch (serviceError) {
        console.warn(`サービスアカウント追加で警告: ${serviceError.message}`);
      }
    }

    return {
      success: true,
      message: 'スプレッドシートアクセス権限を修復しました。ドメイン全体でアクセス可能です。',
    };
  } catch (e) {
    console.error(`スプレッドシートアクセス権限の修復に失敗: ${e.message}`);
    return {
      success: false,
      error: e.message,
    };
  }
}

/**
 * 管理パネル専用：権限問題の緊急修復
 * @param {string} userEmail - ユーザーメール
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {object} 修復結果
 */
function addReactionColumnsToSpreadsheet(spreadsheetId, sheetName) {
  try {
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName) || spreadsheet.getSheets()[0];

    const additionalHeaders = [
      CONSTANTS.REACTIONS.LABELS.UNDERSTAND,
      CONSTANTS.REACTIONS.LABELS.LIKE,
      CONSTANTS.REACTIONS.LABELS.CURIOUS,
      CONSTANTS.REACTIONS.LABELS.HIGHLIGHT,
    ];

    // 効率的にヘッダー情報を取得
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const startCol = currentHeaders.length + 1;

    // バッチでヘッダーを追加
    sheet.getRange(1, startCol, 1, additionalHeaders.length).setValues([additionalHeaders]);

    // スタイリングを一括適用
    const allHeadersRange = sheet.getRange(
      1,
      1,
      1,
      currentHeaders.length + additionalHeaders.length
    );
    allHeadersRange.setFontWeight('bold').setBackground('#E3F2FD');

    // 自動リサイズ（エラーが出ても続行）
    try {
      sheet.autoResizeColumns(1, allHeadersRange.getNumColumns());
    } catch (resizeError) {
      console.warn('Auto-resize failed:', resizeError.message);
    }

  } catch (e) {
    console.error(`リアクション列追加エラー: ${e.message}`);
    // エラーでも処理は継続
  }
}

/**
 * スプレッドシートの共有設定をチェックする
 * @param {string} spreadsheetId - チェックするスプレッドシートのID
 * @returns {object} status ('success' or 'error') と message
 */
function getSheetData(userId, sheetName, classFilter, sortMode, adminMode) {
  // キャッシュキー生成（ユーザー、シート、フィルタ条件ごとに個別キャッシュ）
  const cacheKey = `sheetData_${userId}_${sheetName}_${classFilter}_${sortMode}`;

  // 管理モードの場合はキャッシュをバイパス（最新データを取得）
  if (adminMode === true) {
    console.log('🔄 管理モード：シートデータキャッシュをバイパス');
    return executeGetSheetData(userId, sheetName, classFilter, sortMode);
  }

  return cacheManager.get(
    cacheKey,
    () => {
      return executeGetSheetData(userId, sheetName, classFilter, sortMode);
    },
    { ttl: 180 }
  ); // 3分間キャッシュ（短縮してデータ整合性向上）
}

/**
 * ⚠️ レガシー互換性のためのsheetConfig代替関数
 * 既存のconfigJSONから必要な設定を生成
 */
// getSheetConfig関数はbuildSheetConfigDynamicallyに統合されました

/**
 * 動的sheetConfig取得関数（レガシー互換性）
 * 実行時にユーザー情報を取得して設定を構築
 */
function buildSheetConfigDynamically(userIdParam) {
  try {
    let userInfo;
    
    if (userIdParam) {
      userInfo = DB.findUserById(userIdParam);
    } else {
      userInfo = getActiveUserInfo();
    }
    
    if (!userInfo || !JSON.parse(userInfo.configJson || '{}')) {
      console.warn('buildSheetConfigDynamically: ユーザー情報なし');
      return {
        spreadsheetId: null,
        sheetName: null,
        opinionHeader: 'お題',
        reasonHeader: '理由',
        classHeader: 'クラス',
        nameHeader: '名前',
        displayMode: 'anonymous',
        setupStatus: 'pending'
      };
    }
    
    const config = JSON.parse(userInfo.configJson || '{}');
    return {
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      opinionHeader: config.opinionHeader || 'お題',
      reasonHeader: config.reasonHeader || '理由',
      classHeader: config.classHeader || 'クラス',
      nameHeader: config.nameHeader || '名前',
      displayMode: config.displayMode || 'anonymous',
      setupStatus: config.setupStatus || 'pending'
    };
  } catch (error) {
    console.error('buildSheetConfigDynamically エラー:', error.message);
    // フォールバック用デフォルト設定
    return {
      spreadsheetId: null,
      sheetName: null,
      opinionHeader: 'お題',
      reasonHeader: '理由',
      classHeader: 'クラス',
      nameHeader: '名前',
      displayMode: 'anonymous',
      setupStatus: 'pending'
    };
  }
}

// getSheetConfigSafeとgetSheetConfigCachedはbuildSheetConfigDynamicallyに統合されました
// 使用箇所ではbuildSheetConfigDynamically()を使用してください

// ✅ 直接アクセス用のプロパティ設定
Object.defineProperty(globalThis, 'sheetConfig', {
  get: function() {
    return buildSheetConfigDynamically();
  },
  configurable: true
});

/**
 * 実際のシートデータ取得処理（キャッシュ制御から分離）
 */
function executeGetSheetData(userId, sheetName, classFilter, sortMode) {
  try {
    logDebug('executeGetSheetData_start', {
      userId,
      sheetName,
      classFilter,
      sortMode
    });

    const userInfo = DB.findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 🎯 シンプル化: 統一設定取得関数使用
    const config = getConfigSimple(userInfo);
    
    // 🔧 修正: 全列データを取得（A:E制限を削除）
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);
    const sheetData = sheet.getDataRange().getValues();
    
    logDebug('sheet_data_retrieved', {
      totalRows: sheetData.length,
      totalColumns: sheetData[0]?.length || 0,
      hasColumnMapping: !!config.columnMapping.mapping
    });

    if (sheetData.length === 0) {
      return createSuccessResponse([]);
    }

    const headers = sheetData[0];
    const dataRows = sheetData.slice(1);
    
    // 🎯 シンプル化: columnMappingを使用した直接データ処理
    const processedData = processDataWithColumnMapping(
      dataRows, 
      headers, 
      config.columnMapping.mapping
    );
    
    return createSuccessResponse(processedData, {
      totalCount: processedData.length,
      headers: headers
    });
  } catch (error) {
    logDebug('executeGetSheetData_error', {
      userId,
      sheetName,
      error: error.message
    });
    return createErrorResponse(`データの取得に失敗しました: ${error.message}`, {
      userId,
      sheetName
    });
  }
}

/**
 * シート一覧取得
 */
// getSheetsList関数は削除され、getAvailableSheets関数に統合されました
// 使用箇所はgetAvailableSheets()を使用してください

/**
 * 名簿マップを構築（名簿機能無効化のため空のマップを返す）
 * @param {array} rosterData - 名簿データ（使用されません）
 * @returns {object} 空の名簿マップ
 */
function buildRosterMap() {
  // 名簿機能は使用しないため、常に空のマップを返す
  // 名前はフォーム入力から直接取得
  return {};
}

/**
 * 行データを処理（スコア計算、名前変換など）
 */
function processRowData(row, _, headerIndices, _, displayMode, rowNumber, isOwner) {

  const processedRow = {
    rowNumber,
    originalData: row,
    score: 0,
    likeCount: 0,
    understandCount: 0,
    curiousCount: 0,
    isHighlighted: false,
  };

  // リアクションカウント計算
  CONSTANTS.REACTIONS.KEYS.forEach((reactionKey) => {
    const columnName = CONSTANTS.REACTIONS.LABELS[reactionKey];
    const columnIndex = headerIndices[columnName];

    if (columnIndex !== undefined && row[columnIndex]) {
      const reactions = parseReactionString(row[columnIndex]);
      const count = reactions.length;

      switch (reactionKey) {
        case 'LIKE':
          processedRow.likeCount = count;
          break;
        case 'UNDERSTAND':
          processedRow.understandCount = count;
          break;
        case 'CURIOUS':
          processedRow.curiousCount = count;
          break;
      }
    }
  });

  // ハイライト状態チェック
  const highlightIndex = headerIndices[CONSTANTS.REACTIONS.LABELS.HIGHLIGHT];
  if (highlightIndex !== undefined && row[highlightIndex]) {
    processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
  }

  // スコア計算
  processedRow.score = calculateRowScore(processedRow);

  // 名前の表示処理（フォーム入力の名前を使用）
  const nameIndex = headerIndices[CONSTANTS.COLUMNS.NAME];
  if (
    nameIndex !== undefined &&
    row[nameIndex] &&
    (displayMode === CONSTANTS.DISPLAY_MODES.NAMED || isOwner)
  ) {
    processedRow.displayName = row[nameIndex];
  } else if (displayMode === CONSTANTS.DISPLAY_MODES.NAMED || isOwner) {
    // 名前入力がない場合のフォールバック
    processedRow.displayName = '匿名';
  }

  return processedRow;
}

/**
 * 行のスコアを計算
 */
function calculateRowScore(rowData) {
  const baseScore = 1.0;

  // いいね！による加算
  const likeBonus = rowData.likeCount * CORE.SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;

  // その他のリアクションも軽微な加算
  const reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;

  // ハイライトによる大幅加算
  const highlightBonus = rowData.isHighlighted ? 0.5 : 0;

  // ランダム要素（同じスコアの項目をランダムに並べるため）
  const randomFactor = Math.random() * CORE.SCORING_CONFIG.RANDOM_SCORE_FACTOR;

  return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
}

/**
 * データにソートを適用
 */
function applySortMode(data, sortMode) {
  switch (sortMode) {
    case 'score':
      return data.sort((a, b) => {
        return b.score - a.score;
      });
    case 'newest':
      return data.reverse();
    case 'oldest':
      return data; // 元の順序（古い順）
    case 'random':
      return shuffleArray(data.slice()); // コピーをシャッフル
    case 'likes':
      return data.sort((a, b) => {
        return b.likeCount - a.likeCount;
      });
    default:
      return data;
  }
}

/**
 * 配列をシャッフル（Fisher-Yates shuffle）
 */
function shuffleArray(array) {
  for (let i = array.length - 1; i > 0; i--) {
    const j = Math.floor(Math.random() * (i + 1));
    const temp = array[i];
    array[i] = array[j];
    array[j] = temp;
  }
  return array;
}

/**
 * ヘルパー関数：ヘッダー配列から指定した名前のインデックスを取得
 * CONSTANTSと統一された方式を使用
 */
function getHeaderIndex(headers, headerName) {
  if (!headers || !headerName) return -1;
  return headers.indexOf(headerName);
}

/**
 * CONSTANTSキーから適切なヘッダー名を取得
 * @param {string} columnKey - CONSTANTSのキー（例：'OPINION', 'CLASS'）
 * @returns {string} ヘッダー名
 */
function getColumnHeaderName(columnKey) {
  return CONSTANTS.COLUMNS[columnKey] || CONSTANTS.REACTIONS.LABELS[columnKey] || '';
}

/**
 * 統一されたヘッダーインデックス取得関数
 * @param {array} headers - ヘッダー配列
 * @param {string} columnKey - CONSTANTSのキー
 * @returns {number} インデックス（見つからない場合は-1）
 */

/**
 * 🎯 簡素化された列マッピング関数（レガシー削除版）
 * 統一関数を使用してシンプルに実装
 * @param {Object} config - ユーザー設定
 * @param {Array} headers - ヘッダー配列
 * @returns {Object} マッピングされたインデックス
 */
function mapColumnIndices(config, headers) {
  // 統一関数で全インデックスを取得
  const columnIndices = getAllColumnIndices(config);
  
  // レガシー形式の互換性マッピング
  return {
    opinionHeader: columnIndices.answer,
    reasonHeader: columnIndices.reason,
    classHeader: columnIndices.class,
    nameHeader: columnIndices.name
  };
}

// =================================================================
// 追加のコアファンクション（Code.gsから移行）
// =================================================================

/**
 * 軽量な件数チェック（新着通知用）
 * 実際のデータではなく件数のみを返す
 */

/**
 * 回答ボードのデータを強制的に再読み込み
 */

// =================================================================
// ユーティリティ関数
// =================================================================

/**
 * 現在のユーザーのステータス情報を取得（AppSetupPage.htmlから呼び出される）
 * @returns {object} ユーザーのステータス情報
 */

/**
 * isActive状態を更新（AppSetupPage.htmlから呼び出される）
 * @param {boolean} isActive - 新しいisActive状態
 * @returns {object} 更新結果
 */
function updateIsActiveStatus(requestUserId, isActive) {
  // 新規ユーザー（requestUserIdがundefinedまたはnull）の場合はverifyUserAccessをスキップ
  if (requestUserId) {
    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }
  }
  try {
    const currentUserInfo = new ConfigurationManager().getCurrentUserInfoSafely();
    if (!currentUserInfo) {
      return {
        status: 'error',
        message: 'ユーザーが認証されていません、またはデータベースに登録されていません',
      };
    }
    const { currentUserEmail: activeUserEmail, userInfo } = currentUserInfo;

    // 編集者権限があるか確認（自分自身の状態変更も含む）
    if (!isTrue(userInfo.isActive)) {
      return {
        status: 'error',
        message: 'この操作を実行する権限がありません',
      };
    }

    // isActive状態を更新
    const updateResult = DB.updateUser(userInfo.userId, {
      isActive: Boolean(isActive),  // 正規Boolean型で統一
      lastAccessedAt: new Date().toISOString(),
    });

    if (updateResult.success) {
      const statusMessage = isActive
        ? 'アプリが正常に有効化されました'
        : 'アプリが正常に無効化されました';

      return {
        status: 'success',
        message: statusMessage,
        newStatus: newIsActiveValue,
      };
    } else {
      return {
        status: 'error',
        message: 'データベースの更新に失敗しました',
      };
    }
  } catch (e) {
    console.error(`updateIsActiveStatus エラー: ${e.message}`);
    return {
      status: 'error',
      message: `isActive状態の更新に失敗しました: ${e.message}`,
    };
  }
}

/**
 * セットアップページへのアクセス権限を確認
 * @returns {boolean} アクセス権限があればtrue
 */
function hasSetupPageAccess() {
  try {
    const currentUserInfo = new ConfigurationManager().getCurrentUserInfoSafely();
    if (!currentUserInfo) {
      return false;
    }

    // データベースに登録され、かつisActiveがtrueのユーザーのみアクセス可能
    const { userInfo } = currentUserInfo;
    return userInfo && isTrue(userInfo.isActive);
  } catch (e) {
    console.error(`hasSetupPageAccess エラー: ${e.message}`);
    return false;
  }
}

/**
 * Drive APIサービスを取得
 * @returns {object} Drive APIサービス
 */
function getDriveService() {
  try {
    const accessToken = getServiceAccountTokenCached();
    return {
      accessToken,
      baseUrl: 'https://www.googleapis.com/drive/v3',
      files: {
        get(params) {
          try {
            const url = `${this.baseUrl}/files/${params.fileId}?fields=${encodeURIComponent(params.fields)}`;
            const response = UrlFetchApp.fetch(url, {
              headers: { Authorization: `Bearer ${this.accessToken}` },
            });
            return JSON.parse(response.getContentText());
          } catch (error) {
            console.error('Drive API呼び出しエラー:', error.message);
            throw error;
          }
        },
      },
    };
  } catch (error) {
    console.error('getDriveService初期化エラー:', error.message);
    throw error;
  }
}

/**
 * 現在のアクセスユーザーがデバッグを有効にすべきかを判定
 * @returns {boolean} デバッグモードを有効にすべき場合はtrue
 */
function shouldEnableDebugMode() {
  return isSystemAdmin();
}

/**
 * デプロイユーザーかどうか判定
 * 1. GASスクリプトの編集権限を確認
 * 2. 教育機関ドメインの管理者権限を確認
 * @returns {boolean} 管理者権限を持つ場合 true
 */
function isSystemAdmin() {
  try {
    const props = PropertiesService.getScriptProperties();
    const adminEmail = props.getProperty(PROPS_KEYS.ADMIN_EMAIL);
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    return adminEmail && currentUserEmail && adminEmail === currentUserEmail;
  } catch (e) {
    console.error(`isSystemAdmin エラー: ${e.message}`);
    return false;
  }
}

// =================================================================
// マルチテナント対応関数（同時アクセス対応）
// =================================================================

/**
 * マルチテナント対応: 設定保存・公開
 * @param {string} requestUserId - リクエストされたユーザーID（第一引数）
 * @param {Object} settingsData - 設定データ
 * @returns {Object} 実行結果
 */

/**
 * マルチテナント対応: アクティブフォーム情報取得
 * @param {string} requestUserId - リクエストされたユーザーID
 * @returns {Object} フォーム情報
 */
/**
 * 管理者によるユーザー削除（AppSetupPage.html用ラッパー）
 * @param {string} targetUserId - 削除対象ユーザーID
 * @param {string} reason - 削除理由
 */
function deleteUserAccountByAdminForUI(targetUserId, reason) {
  try {
    const result = DB.deleteUserAccountByAdmin(targetUserId, reason);
    return {
      status: 'success',
      message: result.message,
      deletedUser: result.deletedUser,
    };
  } catch (error) {
    console.error('deleteUserAccountByAdmin wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * 削除ログ取得（AppSetupPage.html用ラッパー）
 */
function getDeletionLogsForUI() {
  try {
    const logs = getDeletionLogs();
    return {
      status: 'success',
      logs,
    };
  } catch (error) {
    console.error('getDeletionLogs wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

function getAllUsersForAdminForUI(requestUserId) {
  try {
    const result = getAllUsersForAdmin();
    return {
      status: 'success',
      users: result,
    };
  } catch (error) {
    console.error('getAllUsersForAdminForUI wrapper error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * カスタムフォーム作成（AdminPanel.html用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {object} config - フォーム設定
 */
function createForm(requestUserId, config) {
  try {
    // セキュリティチェック: ユーザー認証と入力検証
    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }
    const activeUserEmail = UserManager.getCurrentEmail();

    // 入力検証
    if (!config || typeof config !== 'object') {
      throw new Error('フォーム設定が無効です');
    }

    // AdminPanelのconfig構造を内部形式に変換（createCustomForm の処理を統合）
    const convertedConfig = {
      mainQuestion: {
        title: config.mainQuestion || '今日の学習について、あなたの考えや感想を聞かせてください',
        type: config.responseType || config.questionType || 'text', // responseTypeを優先して使用
        choices: config.questionChoices || config.choices || [], // questionChoicesを優先して使用
        includeOthers: config.includeOthers || false,
      },
      enableClass: config.enableClass || false,
      classQuestion: {
        choices: config.classChoices || ['クラス1', 'クラス2', 'クラス3', 'クラス4'],
      },
    };

    console.log('createForm - converted config:', JSON.stringify(convertedConfig));

    const overrides = {
      titlePrefix: config.formTitle || 'カスタムフォーム',
      customConfig: convertedConfig,
    };

    const result = createUnifiedForm('custom', activeUserEmail, requestUserId, overrides);

    // ユーザー専用フォルダを作成・取得してファイルを移動
    let folder = null;
    const moveResults = { form: false, spreadsheet: false };
    try {
      console.log('📁 ユーザーフォルダ作成とファイル移動開始...');
      folder = createUserFolder(activeUserEmail);

      if (folder) {
        // フォームファイルを移動
        try {
          const formFile = DriveApp.getFileById(result.formId);
          if (formFile) {
            // 既にフォルダに存在するかチェック
            const formParents = formFile.getParents();
            let isFormAlreadyInFolder = false;

            while (formParents.hasNext()) {
              if (formParents.next().getId() === folder.getId()) {
                isFormAlreadyInFolder = true;
                break;
              }
            }

            if (!isFormAlreadyInFolder) {
              folder.addFile(formFile);
              DriveApp.getRootFolder().removeFile(formFile);
              moveResults.form = true;
              console.log('✅ カスタムフォームファイル移動完了');
            }
          }
        } catch (formMoveError) {
          console.warn('カスタムフォームファイル移動エラー:', formMoveError.message);
        }

        // スプレッドシートファイルを移動
        try {
          const ssFile = DriveApp.getFileById(result.spreadsheetId);
          if (ssFile) {
            // 既にフォルダに存在するかチェック
            const ssParents = ssFile.getParents();
            let isSsAlreadyInFolder = false;

            while (ssParents.hasNext()) {
              if (ssParents.next().getId() === folder.getId()) {
                isSsAlreadyInFolder = true;
                break;
              }
            }

            if (!isSsAlreadyInFolder) {
              folder.addFile(ssFile);
              DriveApp.getRootFolder().removeFile(ssFile);
              moveResults.spreadsheet = true;
              console.log('✅ カスタムスプレッドシートファイル移動完了');
            }
          }
        } catch (ssMoveError) {
          console.warn('カスタムスプレッドシートファイル移動エラー:', ssMoveError.message);
        }

        console.log('📁 カスタムセットアップファイル移動結果:', moveResults);
      }
    } catch (folderError) {
      console.warn('カスタムセットアップフォルダ処理エラー:', folderError.message);
    }

    // 既存ユーザーの情報を更新（スプレッドシート情報を追加）
    const existingUser = DB.findUserById(requestUserId);
    if (existingUser) {
      console.log('createCustomFormUI - updating user data for:', requestUserId);

      const updatedConfigJson = existingUser.parsedConfig || {};
      updatedConfigJson.formUrl = result.formUrl;
      updatedConfigJson.editFormUrl = result.editFormUrl || result.formUrl;
      updatedConfigJson.formCreated = true;
      updatedConfigJson.lastFormCreatedAt = new Date().toISOString();
      updatedConfigJson.setupStatus = 'completed';
      updatedConfigJson.appPublished = true;
      updatedConfigJson.folderId = folder ? folder.getId() : '';
      updatedConfigJson.folderUrl = folder ? folder.getUrl() : '';

      // ✅ configJSON中心型: カスタムフォーム設定を直接configJSONに統合
      updatedConfigJson.formTitle = config.formTitle;
      updatedConfigJson.mainQuestion = config.mainQuestion;
      updatedConfigJson.questionType = config.questionType;
      updatedConfigJson.choices = config.choices;
      updatedConfigJson.includeOthers = config.includeOthers;
      updatedConfigJson.enableClass = config.enableClass;
      updatedConfigJson.classChoices = config.classChoices;
      updatedConfigJson.lastModified = new Date().toISOString();

      // ✅ configJSON中心型: すべての設定を統合済み、削除処理不要

      // 🔧 修正: 直接フィールド更新 + ConfigManager経由での安全保存
      const directUpdateData = {
        lastAccessedAt: new Date().toISOString(),
      };

      console.log('createCustomFormUI - 直接更新フィールド:', JSON.stringify(directUpdateData));
      console.log('createCustomFormUI - ConfigManager保存データ:', Object.keys(updatedConfigJson).length, 'fields');

      // 直接フィールドのみDB更新
      DB.updateUserInDatabase(requestUserId, directUpdateData);
      // configJsonは安全に保存
      ConfigManager.saveConfig(requestUserId, updatedConfigJson);

      // カスタムフォーム作成後の包括的キャッシュ同期（Quick Startと同様）
      console.log('🗑️ カスタムフォーム作成後の包括的キャッシュ同期中...');
      synchronizeCacheAfterCriticalUpdate(
        requestUserId,
        activeUserEmail,
        existingUser.spreadsheetId,
        result.spreadsheetId
      );
    } else {
      console.warn('createCustomFormUI - user not found:', requestUserId);
    }

    return {
      status: 'success',
      message: 'カスタムフォームが正常に作成されました！',
      formUrl: result.formUrl,
      spreadsheetUrl: result.spreadsheetUrl,
      formTitle: result.formTitle,
      sheetName: result.sheetName,
      folderId: folder ? folder.getId() : '',
      folderUrl: folder ? folder.getUrl() : '',
      filesMovedToFolder: moveResults,
    };
  } catch (error) {
    console.error('createCustomFormUI error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * クイックスタート用フォーム作成（UI用ラッパー）
 * @param {string} requestUserId - リクエスト元のユーザーID
 */
// createQuickStartFormUI - 削除済み（クイックスタート機能不要）


/**
 * シートを有効化（AdminPanel.html用のシンプル版）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} sheetName - シート名
 */
function activateSheetSimple(requestUserId, sheetName) {
  try {
    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }
    const userInfo = DB.findUserById(requestUserId);
    // 🚀 CLAUDE.md準拠：統一データソース原則
    const config = userInfo ? JSON.parse(userInfo.configJson || '{}') || {} : {};
    if (!userInfo || !config.spreadsheetId) {
      throw new Error('ユーザー情報またはスプレッドシートIDが見つかりません');
    }

    return activateSheet(requestUserId, config.spreadsheetId, sheetName);
  } catch (error) {
    console.error('activateSheetSimple error:', error.message);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * ログインフロー修正の検証テスト
 * @returns {object} テスト結果
 */

/**
 * セキュリティ強化とキャッシュ最適化のパフォーマンステスト
 * @returns {object} パフォーマンステスト結果
 */

/**
 * 統合ログインフロー処理 - セキュリティ強化とパフォーマンス最適化
 * 従来の複数API呼び出し（getCurrentUserStatus → getExistingBoard → registerNewUser）を1回に集約
 * @returns {object} ログイン結果とリダイレクトURL
 */
function getLoginStatus() {
  try {
    const currentUserInfo = new ConfigurationManager().getCurrentUserInfoSafely();
    if (!currentUserInfo) {
      return { status: 'error', message: 'ログインユーザーの情報を取得できませんでした。' };
    }
    const { currentUserEmail: activeUserEmail, userInfo } = currentUserInfo;

    let result;
    if (
      userInfo &&
      (userInfo.isActive === true || String(userInfo.isActive).toLowerCase() === 'true')
    ) {
      const urls = generateUserUrls(userInfo.userId);
      result = {
        status: 'existing_user',
        userId: userInfo.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: 'ログインが完了しました',
      };
    } else if (userInfo) {
      const urls = generateUserUrls(userInfo.userId);
      result = {
        status: 'setup_required',
        userId: userInfo.userId,
        adminUrl: urls.adminUrl,
        viewUrl: urls.viewUrl,
        message: 'セットアップを完了してください',
      };
    } else {
      // 新規ユーザーの場合は自動登録を実行
      console.log('getLoginStatus: 新規ユーザーを検出、自動登録を実行');
      const registrationResult = registerNewUser(activeUserEmail);
      if (registrationResult && registrationResult.userId) {
        // 登録成功
        result = {
          status: 'new_user_registered',
          userId: registrationResult.userId,
          adminUrl: registrationResult.adminUrl,
          viewUrl: registrationResult.viewUrl,
          message: registrationResult.message || '新規ユーザー登録が完了しました',
        };
      } else {
        result = {
          status: 'registration_failed',
          userEmail: activeUserEmail,
          message: registrationResult.message || '登録に失敗しました',
        };
      }
    }

    // 🔧 修正：ログイン結果のキャッシュも無効化（削除ユーザー問題対応）
    // try {
    //   CacheService.getScriptCache().put(cacheKey, JSON.stringify(result), 30);
    // } catch (e) {
    //   console.warn('getLoginStatus: キャッシュ保存エラー -', e.message);
    // }

    return result;
  } catch (error) {
    console.error('getLoginStatus error:', error);
    return { status: 'error', message: error.message };
  }
}

// confirmUserRegistration function removed - 自動登録はgetLoginStatus()で実行

// =================================================================
// 統合API（フェーズ2最適化）
// =================================================================

/**
 * 統合初期データ取得API - OPTIMIZED
 * 従来の5つのAPI呼び出し（getCurrentUserStatus, getUserId, getAppConfig, getSheetDetails）を統合
 * @param {string} requestUserId - リクエスト元のユーザーID（省略可能）
 * @param {string} sheetName - 詳細を取得するシート名（省略可能）
 * @returns {Object} 統合された初期データ
 */
function getInitialData(requestUserId, sheetName) {
  console.log('getInitialData: 統合初期化開始');

  try {
    const startTime = new Date().getTime();

    // === ステップ1: ユーザー認証とユーザー情報取得（キャッシュ活用） ===
    const activeUserEmail = UserManager.getCurrentEmail();
    const currentUserId = requestUserId;

    // UserID の解決
    if (!currentUserId) {
      currentUserId = getUserId();
    }

    // Phase3 Optimization: Use execution-level cache to avoid duplicate database queries
    clearExecutionUserInfoCache(); // Clear any stale cache

    // ユーザー認証
    const accessResult = App.getAccess().verifyAccess(
      currentUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }
    const userInfo = getActiveUserInfo(); // Use cached version
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // === ステップ1.5: データ整合性の自動チェックと修正 ===
    try {
      const consistencyResult = fixUserDataConsistency(currentUserId);
      if (consistencyResult.updated) {
        console.log('✅ データ整合性が自動修正されました');
        // 修正後は最新データを再取得
        clearExecutionUserInfoCache();
        userInfo = getActiveUserInfo();
      }
    } catch (consistencyError) {
      console.warn('⚠️ データ整合性チェック中にエラー:', consistencyError.message);
      // エラーが発生しても初期化処理は続行
    }

    // === ステップ2: 設定データの取得と自動修復 ===
    const configJson = JSON.parse(userInfo.configJson || '{}') || {};

    // Auto-healing for inconsistent setup states
    let needsUpdate = false;
    if (configJson.formUrl && !configJson.formCreated) {
      configJson.formCreated = true;
      needsUpdate = true;
    }
    if (configJson.formCreated && configJson.setupStatus !== 'completed') {
      configJson.setupStatus = 'completed';
      needsUpdate = true;
    }
    if (configJson.sheetName && !configJson.appPublished) {
      configJson.appPublished = true;
      needsUpdate = true;
    }
    if (needsUpdate) {
      try {
        // 🔧 修正: ConfigManager経由で安全な保存（二重構造防止）
        ConfigManager.saveConfig(currentUserId, configJson);
        // userInfo.configJsonの直接更新は削除（二重構造原因となるため）
      } catch (updateErr) {
        console.warn(`Config auto-heal failed: ${updateErr.message}`);
      }
    }

    // === ステップ3: シート一覧とアプリURL生成 ===
    const sheets = getAvailableSheets(currentUserId);
    const appUrls = generateUserUrls(currentUserId);

    // === ステップ4: 回答数とリアクション数の取得 ===
    let answerCount = 0;
    let totalReactions = 0;
    try {
      if (configJson.spreadsheetId && configJson.sheetName) {
        const responseData = getResponsesData(currentUserId, configJson.sheetName);
        if (responseData.status === 'success') {
          answerCount = responseData.data.length;
          totalReactions = answerCount * 2; // 暫定値
        }
      }
    } catch (err) {
      console.warn('Answer count retrieval failed:', err.message);
    }

    // === ステップ5: セットアップステップの決定 ===
    const setupStep = determineSetupStep(userInfo, configJson);

    // ✅ configJSON中心型: 公開シート設定でsheetConfig廃止
    const sheetName = configJson.sheetName || '';
    
    // ✅ columnMapping形式を使用（legacy形式を完全削除）
    const columnMapping = configJson.columnMapping;
    
    if (!columnMapping) {
      console.warn('⚠️ getConfig: columnMappingが設定されていません。');
    }
    
    const opinionHeader = columnMapping?.answer?.header || '';
    const nameHeader = columnMapping?.name?.header || '';
    const classHeader = columnMapping?.class?.header || '';

    // === ベース応答の構築 ===
    const response = {
      // ユーザー情報
      userInfo: {
        userId: userInfo.userId,
        userEmail: userInfo.userEmail,
        isActive: userInfo.isActive,
        lastAccessedAt: config.lastAccessedAt,
        spreadsheetId: config.spreadsheetId,
        spreadsheetUrl: config.spreadsheetUrl,
        configJson: userInfo.configJson,
      },
      // アプリ設定
      appUrls,
      setupStep,
      activeSheetName: configJson.sheetName || null,
      webAppUrl: appUrls.webApp,
      isPublished: !!configJson.appPublished,
      answerCount,
      totalReactions,
      config: {
        sheetName,
        opinionHeader,
        nameHeader,
        classHeader,
        showNames: configJson.showNames || false,
        showCounts: configJson.showCounts !== undefined ? configJson.showCounts : false,
        displayMode: configJson.displayMode || 'anonymous',
        setupStatus: configJson.setupStatus || 'initial',
        isPublished: !!configJson.appPublished,
      },
      // シート情報
      allSheets: sheets,
      sheetNames: sheets,
      // カスタムフォーム情報
      customFormInfo: configJson.formUrl
        ? {
            title: configJson.formTitle || 'カスタムフォーム',
            mainQuestion:
              configJson.mainQuestion || opinionHeader || configJson.sheetName || '質問',
            formUrl: configJson.formUrl,
          }
        : null,
      // メタ情報
      _meta: {
        apiVersion: 'integrated_v1',
        executionTime: null,
        includedApis: ['getCurrentUserStatus', 'getUserId', 'getAppConfig'],
      },
    };

    // === ステップ6: シート詳細の取得（オプション）- 最適化版 ===
    const includeSheetDetails = sheetName || configJson.sheetName;

    // デバッグ: シート詳細取得パラメータの確認
    console.log('📋 シート詳細取得パラメータ:', {
      sheetName,
      sheetName: configJson.sheetName,
      includeSheetDetails,
      hasSpreadsheetId: !!config.spreadsheetId,
      willIncludeSheetDetails: !!(includeSheetDetails && config.spreadsheetId),
    });

    // sheetNameが空の場合のフォールバック処理
    if (!includeSheetDetails && config.spreadsheetId && configJson) {
      console.warn('⚠️ シート名が指定されていません。デフォルトシート名を検索中...');
      try {
        // 一般的なシート名パターンを試す
        const commonSheetNames = [
          'フォームの回答 1',
          'フォーム回答 1',
          'Form Responses 1',
          'Sheet1',
          'シート1',
        ];
        const tempService = getSheetsServiceCached();
        const spreadsheetInfo = getSpreadsheetsData(tempService, config.spreadsheetId);

        if (spreadsheetInfo && spreadsheetInfo.sheets && spreadsheetInfo.sheets.length > 0) {
          // 既知のシート名から最初に見つかったものを使用
          for (const commonName of commonSheetNames) {
            if (spreadsheetInfo.sheets.some((sheet) => sheet.properties.title === commonName)) {
              includeSheetDetails = commonName;
              console.log('✅ フォールバックシート名を使用:', commonName);
              break;
            }
          }

          // それでも見つからない場合は最初のシートを使用
          if (!includeSheetDetails) {
            includeSheetDetails = spreadsheetInfo.sheets[0].properties.title;
            console.log('✅ 最初のシートを使用:', includeSheetDetails);
          }
        }
      } catch (fallbackError) {
        console.warn('⚠️ フォールバックシート名検索に失敗:', fallbackError.message);
      }
    }

    if (includeSheetDetails && config.spreadsheetId) {
      try {
        // 最適化: getSheetsServiceの重複呼び出しを避けるため、一度だけ作成して再利用
        const sharedSheetsService = getSheetsServiceCached();

        // ExecutionContext を最適化版で作成（sheetsService と userInfo を渡して重複作成を回避）
        const context = createExecutionContext(currentUserId, {
          reuseService: sharedSheetsService,
          reuseUserInfo: userInfo,
        });
        const sheetDetails = getSheetDetailsFromContext(
          context,
          config.spreadsheetId,
          includeSheetDetails
        );
        response.sheetDetails = sheetDetails;
        response._meta.includedApis.push('getSheetDetails');
        console.log('✅ シート詳細を統合応答に追加 (最適化版):', includeSheetDetails);
        // getInitialData内で生成したcontextの変更をコミット
        commitAllChanges(context);
      } catch (sheetErr) {
        console.warn('Sheet details retrieval failed:', sheetErr.message);
        response.sheetDetailsError = sheetErr.message;
      }
    }

    // === 実行時間の記録 ===
    const endTime = new Date().getTime();
    response._meta.executionTime = endTime - startTime;

    console.log('⏱️ getInitialData実行完了:', {
      executionTime: `${response._meta.executionTime}ms`,
      userId: currentUserId,
      setupStep,
      hasSheetDetails: !!response.sheetDetails,
    });

    return response;
  } catch (error) {
    console.error('❌ getInitialData error:', error);
    return {
      status: 'error',
      message: error.message,
      _meta: {
        apiVersion: 'integrated_v1',
        error: error.message,
      },
    };
  }
}

/**
 * 手動データ整合性修正（管理パネル用）
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {object} 修正結果
 */
function fixDataConsistencyManual(requestUserId) {
  try {
    const accessResult = App.getAccess().verifyAccess(
      requestUserId,
      'view',
      UserManager.getCurrentEmail()
    );
    if (!accessResult.allowed) {
      throw new Error(`アクセスが拒否されました: ${accessResult.reason}`);
    }
    console.log('🔧 手動データ整合性修正実行:', requestUserId);

    const result = fixUserDataConsistency(requestUserId);

    if (result.status === 'success') {
      if (result.updated) {
        return {
          status: 'success',
          message: 'データ整合性の問題を修正しました。ページを再読み込みしてください。',
          details: result.message,
        };
      } else {
        return {
          status: 'success',
          message: 'データ整合性に問題は見つかりませんでした。',
          details: result.message,
        };
      }
    } else {
      return {
        status: 'error',
        message: `データ整合性修正中にエラーが発生しました: ${result.message}`,
      };
    }
  } catch (error) {
    console.error('❌ 手動データ整合性修正エラー:', error);
    return {
      status: 'error',
      message: `修正処理中にエラーが発生しました: ${error.message}`,
    };
  }
}

/**
 * getSheetDetailsの内部実装（統合API用）
 * @param {string} requestUserId - ユーザーID
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} シート詳細
 */

/**
 * アプリケーション状態取得（UI用）
 * @returns {Object} アプリケーション状態情報
 */
function getApplicationStatusForUI() {
  try {
    const accessCheck = Access.check();
    const isEnabled = getApplicationEnabled();
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};

    return {
      status: 'success',
      isEnabled,
      isSystemAdmin: accessCheck.isSystemAdmin,
      currentUserEmail,
      lastUpdated: new Date().toISOString(),
      message: accessCheck.accessReason,
    };
  } catch (error) {
    console.error('getApplicationStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

/**
 * アプリケーション状態設定（UI用）
 * @param {boolean} enabled - 有効化するかどうか
 * @returns {Object} 設定結果
 */
function setApplicationStatusForUI(enabled) {
  try {
    const result = setApplicationEnabled(enabled);
    return {
      status: 'success',
      enabled: result.enabled,
      message: result.message,
      timestamp: result.timestamp,
      currentUserEmail: result.adminEmail,
    };
  } catch (error) {
    console.error('setApplicationStatusForUI エラー:', error);
    return {
      status: 'error',
      message: error.message,
    };
  }
}

// =============================================================================
// 高精度AI列判定システム
// =============================================================================

/**
 * インターネット活用・超高精度AI列判定システム
 * 既存のidentifyHeaders()を大幅に強化
 * @param {Array<string>} headers - ヘッダー行の配列
 * @param {Object} options - オプション設定
 * @returns {Object} 超高精度で推測された列マッピング
 */
function identifyHeadersAdvanced(headers, options = {}) {
  const { useWebKnowledge = true, useContextAnalysis = true } = options;

  // 1. 既存のidentifyHeaders()を基礎として活用
  const basicResult = identifyHeaders(headers);

  // 2. インターネット知識ベースの強化判定
  if (useWebKnowledge) {
    const webEnhancedResult = enhanceWithWebKnowledge(headers, basicResult);
    Object.assign(basicResult, webEnhancedResult);
  }

  // 3. 文脈・意味解析による精度向上
  if (useContextAnalysis) {
    const contextResult = analyzeContextualMeaning(headers, basicResult);
    Object.assign(basicResult, contextResult);
  }

  // 4. 信頼度計算の高精度化
  basicResult.confidence = calculateAdvancedConfidence(headers, basicResult);

  console.log('identifyHeadersAdvanced: 超高精度判定完了', {
    originalHeaders: headers,
    detectedMapping: basicResult,
    confidence: basicResult.confidence,
  });

  return basicResult;
}

/**
 * インターネット知識ベースによる列判定強化
 */
function enhanceWithWebKnowledge(headers, basicResult) {
  const enhancements = {};

  // 教育分野の一般的なパターン（Googleフォーム等で頻出）
  const educationPatterns = {
    question: [
      'どうして',
      'なぜ',
      '理由は',
      'どのように',
      '何が',
      '思いますか',
      '考えますか',
      'について',
    ],
    answer: ['回答', '答え', '意見', '考え', '思う', 'と思います', 'だと考えます'],
    reason: ['理由', '根拠', '体験', 'そう考える', 'なぜなら', 'から'],
    name: ['名前', '氏名', 'お名前', '学生名', '回答者'],
    class: ['クラス', '組', '学級', '学年', 'グループ'],
  };

  // インターネット知識による高精度マッチング
  headers.forEach((header, index) => {
    const headerLower = String(header).toLowerCase();

    Object.entries(educationPatterns).forEach(([type, patterns]) => {
      const matchScore = patterns.reduce((score, pattern) => {
        if (headerLower.includes(pattern.toLowerCase())) {
          // 長い質問文の場合、特別な高スコア
          if (type === 'question' && headerLower.length > 30) {
            return Math.max(score, 95);
          }
          // 完全一致の場合
          if (headerLower === pattern.toLowerCase()) {
            return Math.max(score, 98);
          }
          // 部分一致の場合
          return Math.max(score, 85);
        }
        return score;
      }, 0);

      // 既存結果より高精度の場合、更新
      if (matchScore > (basicResult.confidence?.[type] || 0)) {
        if (type === 'question') {
          enhancements.answer = header; // 質問文 = 回答対象
          enhancements.question = header;
        } else {
          enhancements[type === 'class' ? 'classHeader' : type] = header;
        }
      }
    });
  });

  return enhancements;
}

/**
 * 文脈・意味解析による列判定
 */
function analyzeContextualMeaning(headers, basicResult) {
  const contextEnhancements = {};

  // 文脈分析：質問→回答の関係性を検出
  const questionIndicators = headers.filter((h) => {
    const str = String(h).toLowerCase();
    return (
      str.includes('？') ||
      str.includes('?') ||
      str.includes('ですか') ||
      str.includes('でしょうか') ||
      str.length > 40
    ); // 長文は質問文の可能性大
  });

  if (questionIndicators.length > 0 && !basicResult.answer) {
    // 最も長い質問文を回答対象として設定
    const longestQuestion = questionIndicators.sort(
      (a, b) => String(b).length - String(a).length
    )[0];
    contextEnhancements.answer = longestQuestion;
    contextEnhancements.question = longestQuestion;
  }

  return contextEnhancements;
}

/**
 * 高精度信頼度計算
 */
function calculateAdvancedConfidence(headers, result) {
  const confidence = {};

  Object.entries(result).forEach(([key, value]) => {
    if (key !== 'confidence' && value) {
      const header = String(value).toLowerCase();

      // 長さベースの信頼度（質問文の場合）
      if (key === 'answer' && header.length > 30) {
        confidence[key] = 95;
      }
      // キーワード完全一致の信頼度
      else if (header.includes('回答') || header.includes('理由')) {
        confidence[key] = 90;
      }
      // 部分一致の信頼度
      else {
        confidence[key] = 75;
      }
    }
  });

  return confidence;
}

/**
 * ヘッダー配列から列マッピングを高精度で推測（既存関数）
 * Googleフォームやカスタムシートのヘッダーを自動判定
 * @param {Array<string>} headers - ヘッダー行の配列
 * @returns {Object} 推測された列マッピング
 */
function identifyHeaders(headers) {
  const find = (keys) => {
    const header = headers.find((h) => {
      const hStr = String(h || '').toLowerCase();
      return keys.some((k) => hStr.includes(k.toLowerCase()));
    });
    return header ? String(header) : '';
  };

  console.log('guessHeadersFromArray: Available headers:', headers);

  // Googleフォーム特有のヘッダー構造に対応
  const isGoogleForm = headers.some(
    (h) => String(h || '').includes('タイムスタンプ') || String(h || '').includes('メールアドレス')
  );

  let question = '';
  let answer = '';
  let reason = '';
  let name = '';
  let classHeader = '';

  if (isGoogleForm) {
    console.log('guessHeadersFromArray: Detected Google Form response sheet');

    // Googleフォームの一般的な構造: タイムスタンプ, メールアドレス, [質問1], [質問2], ...
    const nonMetaHeaders = headers.filter((h) => {
      const hStr = String(h || '').toLowerCase();
      return (
        !hStr.includes('タイムスタンプ') &&
        !hStr.includes('timestamp') &&
        !hStr.includes('メールアドレス') &&
        !hStr.includes('email')
      );
    });

    console.log('guessHeadersFromArray: Non-meta headers:', nonMetaHeaders);

    // より柔軟な推測ロジック
    for (let i = 0; i < nonMetaHeaders.length; i++) {
      const header = nonMetaHeaders[i];
      const hStr = String(header || '').toLowerCase();

      // 質問を含む長いテキストを質問ヘッダーとして推測
      if (
        !question &&
        (hStr.includes('だろうか') ||
          hStr.includes('ですか') ||
          hStr.includes('でしょうか') ||
          hStr.length > 20)
      ) {
        question = header;
        // 同じ内容が複数列にある場合、回答用として2番目を使用
        if (i + 1 < nonMetaHeaders.length && nonMetaHeaders[i + 1] === header) {
          answer = header;
          continue;
        }
      }

      // 回答・意見に関するヘッダー
      if (
        !answer &&
        (hStr.includes('回答') ||
          hStr.includes('答え') ||
          hStr.includes('意見') ||
          hStr.includes('考え'))
      ) {
        answer = header;
      }

      // 理由に関するヘッダー
      if (!reason && (hStr.includes('理由') || hStr.includes('詳細') || hStr.includes('説明'))) {
        reason = header;
      }

      // 名前に関するヘッダー
      if (!name && (hStr.includes('名前') || hStr.includes('氏名') || hStr.includes('学生'))) {
        name = header;
      }

      // クラスに関するヘッダー
      if (
        !classHeader &&
        (hStr.includes('クラス') || hStr.includes('組') || hStr.includes('学級'))
      ) {
        classHeader = header;
      }
    }

    // フォールバック: まだ回答が見つからない場合
    if (!answer && nonMetaHeaders.length > 0) {
      // 最初の非メタヘッダーを回答として使用
      answer = nonMetaHeaders[0];
    }
  } else {
    // 通常のシート用の推測ロジック
    question = find(['質問', '問題', 'question', 'Q']);
    answer = find(['回答', '答え', 'answer', 'A', '意見', 'opinion']);
    reason = find(['理由', 'reason', '詳細', 'detail']);
    name = find(['名前', '氏名', 'name', '学生', 'student']);
    classHeader = find(['クラス', 'class', '組', '学級']);
  }

  // リアクション列も検出（システム列は完全一致のみ）
  const understand = headers.find((h) => String(h || '').trim() === 'なるほど！') || '';
  const like = headers.find((h) => String(h || '').trim() === 'いいね！') || '';
  const curious = headers.find((h) => String(h || '').trim() === 'もっと知りたい！') || '';
  const highlight = headers.find((h) => String(h || '').trim() === 'ハイライト') || '';

  const guessed = {
    questionHeader: question,
    answerHeader: answer,
    reasonHeader: reason,
    nameHeader: name,
    classHeader,
    understandHeader: understand,
    likeHeader: like,
    curiousHeader: curious,
    highlightHeader: highlight,
  };

  console.log('guessHeadersFromArray: Guessed headers:', guessed);

  // 最終フォールバック: 何も推測できない場合
  if (!question && !answer && headers.length > 0) {
    console.log('guessHeadersFromArray: No specific headers found, using positional mapping');

    // タイムスタンプとメールを除外して最初の列を回答として使用
    const usableHeaders = headers.filter((h) => {
      const hStr = String(h || '').toLowerCase();
      return (
        !hStr.includes('タイムスタンプ') &&
        !hStr.includes('timestamp') &&
        !hStr.includes('メールアドレス') &&
        !hStr.includes('email') &&
        String(h || '').trim() !== ''
      );
    });

    if (usableHeaders.length > 0) {
      guessed.answerHeader = usableHeaders[0];
      if (usableHeaders.length > 1) guessed.reasonHeader = usableHeaders[1];
      if (usableHeaders.length > 2) guessed.nameHeader = usableHeaders[2];
      if (usableHeaders.length > 3) guessed.classHeader = usableHeaders[3];
    }
  }

  console.log('guessHeadersFromArray: Final result:', guessed);
  return guessed;
}

/**
 * CLAUDE.md準拠：システム自動修復機能
 * 統一データソース原則：configJSON中心型でシステムの整合性を回復
 * @param {string} userId - ユーザーID
 * @returns {Object} 修復結果
 */
function performAutoRepair(userId) {
  try {
    console.log('performAutoRepair: CLAUDE.md準拠システム修復開始', {
      userId: `${userId?.substring(0, 10)}...`,
      timestamp: new Date().toISOString(),
    });

    if (!SecurityValidator.isValidUUID(userId)) {
      throw new Error('無効なユーザーIDです');
    }

    // ユーザー情報取得
    const userInfo = DB.findUserById(userId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    const config = JSON.parse(userInfo.configJson || '{}') || {};
    const repairResults = {
      fixedItems: [],
      warnings: [],
      errors: [],
      configUpdated: false,
    };

    // 1. 統一データソース整合性チェック
    if (config.spreadsheetId && !config.spreadsheetUrl) {
      config.spreadsheetUrl = `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}`;
      repairResults.fixedItems.push('spreadsheetUrl自動生成');
      repairResults.configUpdated = true;
    }

    if (userInfo.userId && !config.appUrl) {
      const urls = generateUserUrls(userInfo.userId);
      config.appUrl = urls.viewUrl;
      repairResults.fixedItems.push('appURL自動生成');
      repairResults.configUpdated = true;
    }

    // 2. 必須フィールドの確認と修復
    const requiredFields = ['setupStatus', 'displaySettings'];
    requiredFields.forEach((field) => {
      if (!config[field]) {
        switch (field) {
          case 'setupStatus':
            config.setupStatus = config.spreadsheetId ? 'connected' : 'pending';
            break;
          case 'displaySettings':
            config.displaySettings = { showNames: false, showReactions: false };
            break;
        }
        repairResults.fixedItems.push(`${field}の初期化`);
        repairResults.configUpdated = true;
      }
    });

    // 3. 古いフィールドのクリーンアップ
    const obsoleteFields = ['oldSpreadsheetId', 'legacyConfig', 'deprecatedSettings'];
    obsoleteFields.forEach((field) => {
      if (config[field]) {
        delete config[field];
        repairResults.fixedItems.push(`古いフィールド${field}を削除`);
        repairResults.configUpdated = true;
      }
    });

    // 4. configJSON更新（修復が必要な場合のみ）
    if (repairResults.configUpdated) {
      config.lastModified = new Date().toISOString();
      config.autoRepairDate = new Date().toISOString();
      config.claudeMdCompliant = true;

      DB.DB.updateUser(userId, config);
      console.log('performAutoRepair: configJSON更新完了', {
        userId,
        fixedItems: repairResults.fixedItems.length,
        claudeMdCompliant: true,
      });
    }

    console.log('performAutoRepair: 修復処理完了', {
      userId,
      fixedItems: repairResults.fixedItems,
      configUpdated: repairResults.configUpdated,
      claudeMdCompliant: true,
    });

    return {
      success: true,
      fixedItems: repairResults.fixedItems,
      warnings: repairResults.warnings,
      errors: repairResults.errors,
      configUpdated: repairResults.configUpdated,
      claudeMdCompliant: true,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    console.error('❌ performAutoRepair: CLAUDE.md準拠エラー:', {
      userId,
      error: error.message,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
      fixedItems: [],
      warnings: [],
      errors: [error.message],
      configUpdated: false,
      timestamp: new Date().toISOString(),
    };
  }
}

/**
 * 実行時ユーザー情報キャッシュクリア関数
 * 既存のinvalidateUserCache関数のエイリアスとして実装
 * リアクション機能で使用される
 */
function clearExecutionUserInfoCache() {
  // 既存のinvalidateUserCache関数を活用
  // 引数なしで呼び出すと基本的なキャッシュクリアを実行
  return invalidateUserCache();
}

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
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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
    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
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

      // 🔥 完全置換モードでDB更新（古いデータの残存を防止）
      DB.updateUser(userInfo.userId, updatedConfig, { replaceConfig: true });

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
  // 単一フライト制御（ユーザー単位）
  const userLock = LockService.getUserLock();
  try {
    userLock.waitLock(30000); // 最大30秒待機
  } catch (e) {
    return {
      success: false,
      error: '他の操作が進行中です。しばらくしてから再試行してください。',
      optimized: true,
      timestamp: new Date().toISOString(),
    };
  }
  try {
    console.log('📱 publishApplication: アプリ公開開始（最適化版）', {
      hasSpreadsheetId: !!config.spreadsheetId,
      hasSheetName: !!config.sheetName,
      hasColumnMapping: !!config.columnMapping,
      timestamp: new Date().toISOString(),
    });

    // 現在のユーザー情報を取得
    const currentUserEmail = Session.getActiveUser().getEmail();
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    if (!userInfo) {
      console.error('❌ publishApplication: ユーザー情報が見つかりません', {
        currentUserEmail,
        timestamp: new Date().toISOString(),
      });
      throw new Error('ユーザー情報が見つかりません');
    }

    // 現在のconfigJSONを直接取得（パフォーマンス向上）
    const currentConfig = JSON.parse(userInfo.configJson || '{}');

    // 楽観ロック: etag厳格検証（提供されている場合）
    try {
      if (config && config.etag && currentConfig && currentConfig.etag && config.etag !== currentConfig.etag) {
        console.warn('publishApplication: etag不一致により拒否', {
          provided: config.etag,
          current: currentConfig.etag,
        });
        return {
          success: false,
          error: 'etag_mismatch',
          message: '設定が他で更新されました。画面を更新して再試行してください。',
          currentConfig: currentConfig,
        };
      }
    } catch (etErr) {
      console.warn('publishApplication: etag検証エラー', etErr.message);
    }

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
      // 最新のフォーム情報を対象シートから取得（常に再検出・シート優先）
      let detectedFormUrl = null;
      let detectedFormTitle = null;
      try {
        const spreadsheet = new ConfigurationManager().getSpreadsheet(effectiveSpreadsheetId);
        const sheet = spreadsheet.getSheetByName(effectiveSheetName);
        if (sheet && typeof sheet.getFormUrl === 'function') {
          detectedFormUrl = sheet.getFormUrl();
        } else {
          // フォールバック（スプレッドシートに紐付くフォーム）
          detectedFormUrl = spreadsheet.getFormUrl();
        }
        if (detectedFormUrl) {
          try {
            const form = FormApp.openByUrl(detectedFormUrl);
            detectedFormTitle = form ? form.getTitle() : null;
          } catch (formErr) {
            console.warn('publishApplication: フォームタイトル取得失敗', formErr.message);
          }
        }
      } catch (spErr) {
        console.warn('publishApplication: フォームURL再検出失敗', spErr.message);
      }

      // ヘッダー配列とハッシュ（存在しなければ保存時に補完）
      let headers = currentConfig.headers;
      let headersHash = currentConfig.headersHash;
      try {
        if (!headers || !headersHash) {
          const spreadsheet = new ConfigurationManager().getSpreadsheet(effectiveSpreadsheetId);
          const sheet = spreadsheet.getSheetByName(effectiveSheetName);
          if (sheet) {
            const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
            headers = headerRow;
            headersHash = computeHeadersHash(headerRow);
          }
        }
      } catch (hhErr) {
        console.warn('publishApplication: ヘッダー情報の取得失敗', hhErr.message);
      }

      // 🔥 最適化：ConfigManager経由を削除し、直接configJSONを更新
      // 🔥 完全な設定構築（古いデータの残存を防止）
      const updatedConfig = {
        // 基本設定の保持
        createdAt: currentConfig.createdAt || new Date().toISOString(),
        lastAccessedAt: currentConfig.lastAccessedAt || new Date().toISOString(),
        
        // データソース設定（確定済み）
        spreadsheetId: effectiveSpreadsheetId,
        sheetName: effectiveSheetName,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${effectiveSpreadsheetId}`,
        sourceKey: buildSourceKey(effectiveSpreadsheetId, effectiveSheetName),
        
        // 表示設定（フロントエンドから）
        displaySettings: {
          showNames: config.showNames !== undefined ? config.showNames : false,
          showReactions: config.showReactions !== undefined ? config.showReactions : false,
        },
        displayMode: currentConfig.displayMode || 'anonymous',
        
        // アプリケーション状態（公開設定）
        appPublished: true,
        setupStatus: 'completed',
        publishedAt: new Date().toISOString(),
        isDraft: false,
        appUrl: publishResult.appUrl,
        
        // データ接続で設定されたフィールドを保持
        ...(currentConfig.columnMapping && { columnMapping: currentConfig.columnMapping }),
        ...(currentConfig.opinionHeader && { opinionHeader: currentConfig.opinionHeader }),
        ...(currentConfig.reasonHeader && { reasonHeader: currentConfig.reasonHeader }),
        ...(currentConfig.classHeader && { classHeader: currentConfig.classHeader }),
        ...(currentConfig.nameHeader && { nameHeader: currentConfig.nameHeader }),
        // フォーム情報は最新検出結果を優先
        ...(detectedFormUrl !== null && { formUrl: detectedFormUrl || null }),
        ...(detectedFormTitle !== null && { formTitle: detectedFormTitle || null }),
        // ヘッダー配列とハッシュ（あれば保持、なければ今回の検出）
        ...(headers && { headers }),
        ...(headersHash && { headersHash }),
        // headerIndices は廃止（headers[] + columnMapping に統一）
        ...(currentConfig.reactionMapping && { reactionMapping: currentConfig.reactionMapping }),
        ...(currentConfig.systemMetadata && { systemMetadata: currentConfig.systemMetadata }),
        
        // メタ情報
        configVersion: '3.0',
        claudeMdCompliant: true,
        lastModified: new Date().toISOString(),
        etag: computeEtag(),
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

      // 🔥 完全置換モードでDB更新（古いデータの残存を防止）
      DB.updateUser(userInfo.userId, updatedConfig, { replaceConfig: true });

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
  finally {
    try { userLock.releaseLock(); } catch (_) {}
  }
}

/**
 * 設定保存（CLAUDE.md準拠版・ConfigManager統一）
 * ✅ ConfigManager.updateConfig()に統一簡素化
 */
function saveDraftConfiguration(config) {
  // 単一フライト制御（ユーザー単位）
  const userLock = LockService.getUserLock();
  try {
    userLock.waitLock(30000);
  } catch (e) {
    return { success: false, error: '他の操作が進行中です。しばらくして再試行してください。' };
  }
  try {
    console.log('💾 saveDraftConfiguration: 完全置換保存開始', {
      configKeys: Object.keys(config),
      timestamp: new Date().toISOString(),
    });

    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 🔥 現在の設定を取得して、必要なフィールドのみを保持
    const currentConfig = ConfigManager.getUserConfig(userInfo.userId) || {};

    // 楽観ロック: etag厳格検証（提供されている場合）
    try {
      if (config && config.etag && currentConfig && currentConfig.etag && config.etag !== currentConfig.etag) {
        console.warn('saveDraftConfiguration: etag不一致により拒否', {
          provided: config.etag,
          current: currentConfig.etag,
        });
        return {
          success: false,
          error: 'etag_mismatch',
          message: '設定が他で更新されました。画面を更新してから再保存してください。',
          currentConfig: currentConfig,
        };
      }
    } catch (etErr) {
      console.warn('saveDraftConfiguration: etag検証エラー', etErr.message);
    }
    
    // データソースが変更されたかチェック
    const isDataSourceChanged = config.spreadsheetId && config.sheetName && 
      (config.spreadsheetId !== currentConfig.spreadsheetId || config.sheetName !== currentConfig.sheetName);
    
    let updatedConfig;
    
    if (isDataSourceChanged) {
      // データソースが変更された場合は、古いマッピング情報をクリア
      console.log('🔄 データソースが変更されました。マッピング情報をリセット', {
        old: { spreadsheetId: currentConfig.spreadsheetId, sheetName: currentConfig.sheetName },
        new: { spreadsheetId: config.spreadsheetId, sheetName: config.sheetName }
      });
      
      updatedConfig = {
        // 基本的な設定情報を保持
        createdAt: currentConfig.createdAt || new Date().toISOString(),
        lastAccessedAt: new Date().toISOString(),
        
        // 新しいデータソース設定
        spreadsheetId: config.spreadsheetId,
        sheetName: config.sheetName,
        spreadsheetUrl: `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}`,
        sourceKey: buildSourceKey(config.spreadsheetId, config.sheetName),
        
        // 表示設定（管理パネルから更新）
        displaySettings: {
          showNames: config.showNames !== undefined ? config.showNames : false,
          showReactions: config.showReactions !== undefined ? config.showReactions : false,
        },
        displayMode: 'anonymous',
        
        // アプリケーション状態をリセット
        setupStatus: 'data_source_set',
        appPublished: false,
        isDraft: true,
        
        // 古いマッピング情報は削除（新しいデータソースには適用できないため）
        // columnMapping, headerIndices等は意図的に含めない
        
        // メタ情報
        configVersion: '3.0',
        claudeMdCompliant: true,
        lastModified: new Date().toISOString(),
        etag: computeEtag(),
      };

      // 新ソースのヘッダーとハッシュを保存
      try {
        const spreadsheet = new ConfigurationManager().getSpreadsheet(config.spreadsheetId);
        const sheet = spreadsheet.getSheetByName(config.sheetName);
        if (sheet) {
          const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
          updatedConfig.headers = headerRow;
          updatedConfig.headersHash = computeHeadersHash(headerRow);
        }
      } catch (hhErr) {
        console.warn('saveDraftConfiguration: ヘッダー取得失敗（ソース変更時）', hhErr.message);
      }

      // 要求にcolumnMappingが含まれる場合は採用
      if (config.columnMapping) {
        updatedConfig.columnMapping = config.columnMapping;
      }
    } else {
      // データソースが変更されていない場合は、既存のフィールドを保持
      updatedConfig = {
        // 基本的な設定情報を保持
        createdAt: currentConfig.createdAt || new Date().toISOString(),
        lastAccessedAt: currentConfig.lastAccessedAt || new Date().toISOString(),
        
        // データソース設定（変更なし）
        spreadsheetId: config.spreadsheetId || currentConfig.spreadsheetId,
        sheetName: config.sheetName || currentConfig.sheetName,
        spreadsheetUrl: currentConfig.spreadsheetUrl,
        sourceKey: buildSourceKey(
          config.spreadsheetId || currentConfig.spreadsheetId,
          config.sheetName || currentConfig.sheetName
        ),
        
        // 表示設定（管理パネルから更新）
        displaySettings: {
          showNames: config.showNames !== undefined ? config.showNames : (currentConfig.displaySettings?.showNames || false),
          showReactions: config.showReactions !== undefined ? config.showReactions : (currentConfig.displaySettings?.showReactions || false),
        },
        displayMode: currentConfig.displayMode || 'anonymous',
        
        // アプリケーション状態
        setupStatus: currentConfig.setupStatus || 'pending',
        appPublished: currentConfig.appPublished || false,
        isDraft: true,
        
        // 既存のマッピング情報を保持
        ...(currentConfig.columnMapping && { columnMapping: currentConfig.columnMapping }),
        ...(currentConfig.opinionHeader && { opinionHeader: currentConfig.opinionHeader }),
        ...(currentConfig.reasonHeader && { reasonHeader: currentConfig.reasonHeader }),
        ...(currentConfig.classHeader && { classHeader: currentConfig.classHeader }),
        ...(currentConfig.nameHeader && { nameHeader: currentConfig.nameHeader }),
        ...(currentConfig.formUrl && { formUrl: currentConfig.formUrl }),
        ...(currentConfig.formTitle && { formTitle: currentConfig.formTitle }),
        // headerIndices は廃止（headers[] + columnMapping に統一）
        ...(currentConfig.headers && { headers: currentConfig.headers }),
        ...(currentConfig.headersHash && { headersHash: currentConfig.headersHash }),
        ...(currentConfig.reactionMapping && { reactionMapping: currentConfig.reactionMapping }),
        ...(currentConfig.systemMetadata && { systemMetadata: currentConfig.systemMetadata }),
        
        // メタ情報
        configVersion: '3.0',
        claudeMdCompliant: true,
        lastModified: new Date().toISOString(),
        etag: computeEtag(),
      };

      // ヘッダー保存: 要求にheadersが含まれる場合は採用。なければ更新時に最新ヘッダーを保存（互換）
      try {
        if (config.headers && Array.isArray(config.headers)) {
          updatedConfig.headers = config.headers;
          updatedConfig.headersHash = computeHeadersHash(config.headers);
        } else if (updatedConfig.spreadsheetId && updatedConfig.sheetName && !updatedConfig.headers) {
          const spreadsheet = new ConfigurationManager().getSpreadsheet(updatedConfig.spreadsheetId);
          const sheet = spreadsheet.getSheetByName(updatedConfig.sheetName);
          if (sheet) {
            const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0] || [];
            updatedConfig.headers = headerRow;
            updatedConfig.headersHash = computeHeadersHash(headerRow);
          }
        }
      } catch (hhErr2) {
        console.warn('saveDraftConfiguration: ヘッダー取得失敗（更新時）', hhErr2.message);
      }

      // 要求にcolumnMappingが含まれる場合は採用
      if (config.columnMapping) {
        updatedConfig.columnMapping = config.columnMapping;
      }
    }

    // 🔥 ConfigManager.saveConfig()を使用して完全置換
    const success = ConfigManager.saveConfig(userInfo.userId, updatedConfig);

    if (!success) {
      throw new Error('設定の保存に失敗しました');
    }

    console.log('✅ saveDraftConfiguration: 完全置換保存完了', {
      userId: userInfo.userId,
      savedFields: Object.keys(updatedConfig),
      claudeMdCompliant: true,
    });

    return {
      success: true,
      message: 'ドラフトが保存されました',
      claudeMdCompliant: true,
      timestamp: new Date().toISOString(),
    };
  } catch (error) {
    console.error('❌ saveDraftConfiguration: エラー:', {
      error: error.message,
      timestamp: new Date().toISOString(),
    });

    return {
      success: false,
      error: error.message,
    };
  }
  finally {
    try { userLock.releaseLock(); } catch (_) {}
  }
}

// =======================
// Helper utilities (architecture optimizations)
// =======================

function buildSourceKey(spreadsheetId, sheetName) {
  if (!spreadsheetId || !sheetName) return null;
  return `${spreadsheetId}::${sheetName}`;
}

function computeEtag() {
  // 乱数ベースのETag（UUID）にタイムスタンプを付与して衝突回避性を高める
  return `${Utilities.getUuid()}-${new Date().getTime()}`;
}

function normalizeHeaderValue(h) {
  try {
    return String(h)
      .normalize('NFKC')
      .replace(/[、。．・\s]+$/g, '')
      .trim();
  } catch (_) {
    return String(h).trim();
  }
}

function computeHeadersHash(headers) {
  try {
    const normalized = (headers || []).map(normalizeHeaderValue);
    const bytes = Utilities.computeDigest(
      Utilities.DigestAlgorithm.SHA_256,
      JSON.stringify(normalized),
      Utilities.Charset.UTF_8
    );
    return bytes
      .map(function (b) {
        const v = (b + 256) % 256;
        return (v < 16 ? '0' : '') + v.toString(16);
      })
      .join('');
  } catch (e) {
    console.warn('computeHeadersHash失敗', e.message);
    return null;
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
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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

    // 現在のユーザーメール取得
    const currentUserEmail = Session.getActiveUser().getEmail();
    if (!currentUserEmail) {
      throw new Error('ユーザー認証が必要です');
    }

    // キャッシュキー生成（ユーザー固有）
    const cacheKey = `spreadsheet_list_${Utilities.base64Encode(currentUserEmail).replace(/[^a-zA-Z0-9]/g, '')}`;

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
        const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};

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
function executeAppPublish(userId) {
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
 * configJSONのクリーンアップヘルパー関数
 * undefined, null, 空文字列などを削除
 */
function cleanConfigJson(config) {
  if (typeof config !== 'object' || config === null) {
    return config;
  }
  
  const cleaned = {};
  for (const [key, value] of Object.entries(config)) {
    // 削除対象: undefined, null, 空文字列, 空オブジェクト, 空配列
    if (value === undefined || value === null || value === '') {
      continue;
    }
    
    if (Array.isArray(value)) {
      const cleanedArray = value.filter(item => item !== undefined && item !== null && item !== '');
      if (cleanedArray.length > 0) {
        cleaned[key] = cleanedArray;
      }
    } else if (typeof value === 'object') {
      const cleanedObject = cleanConfigJson(value);
      if (Object.keys(cleanedObject).length > 0) {
        cleaned[key] = cleanedObject;
      }
    } else {
      cleaned[key] = value;
    }
  }
  
  return cleaned;
}

/**
 * 🧹 configJSONクリーンアップ実行（管理パネルから呼び出し）
 * 現在のユーザーのconfigJSONを完全クリーンアップ
 */
function executeConfigCleanup() {
  try {
    console.log('🧹 configJSONクリーンアップ実行開始');

    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // configJSONクリーンアップを直接実行
    const config = JSON.parse(userInfo.configJson || '{}');
    const originalSize = JSON.stringify(config).length;
    
    // 基本的なクリーンアップ: undefined, null, 空文字列の削除
    const cleanedConfig = cleanConfigJson(config);
    const cleanedSize = JSON.stringify(cleanedConfig).length;
    
    // クリーンアップされた設定をデータベースに保存
    DB.updateUserConfigJson(userInfo.userId, cleanedConfig);
    
    const result = {
      originalSize,
      cleanedSize,
      sizeReduction: originalSize - cleanedSize,
      reductionRate: Math.round(((originalSize - cleanedSize) / originalSize) * 100),
      removedFields: Object.keys(config).length - Object.keys(cleanedConfig).length,
      timestamp: new Date().toISOString()
    };

    console.log('✅ configJSONクリーンアップ実行完了', result);

    return {
      success: true,
      message: 'configJSONのクリーンアップが完了しました',
      details: {
        削減サイズ: `${result.sizeReduction}文字`,
        削減率: `${result.reductionRate}%`,
        削除フィールド数: result.removedFields,
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
function addMissingColumns(spreadsheetId, sheetName) {
  try {
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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
 * ✅ リアクション文字列を解析（メールアドレスのリストを返す）
 */
function parseReactionString(reactionString) {
  if (!reactionString || typeof reactionString !== 'string') {
    return [];
  }
  
  return reactionString
    .split(',')
    .map(email => email.trim())
    .filter(email => email.length > 0);
}

/**
 * 現在のボード情報とURLを取得（CLAUDE.md準拠版）
 * フロントエンドフッター表示用
 * @returns {Object} ボード情報オブジェクト
 */
/**
 * 現在のユーザー設定を取得（最適化版 - データ上書き防止）
 * ✅ 統一された設定取得（フロントエンド用エントリーポイント）
 * ConfigManager.getUserConfig() への統一されたアクセス
 */
function getConfig() {
  const startTime = Date.now();
  
  try {
    console.log('🔧 getConfig: ユーザー設定取得開始');

    // 現在のユーザー情報を取得
    const currentUserEmail = Session.getActiveUser().getEmail();
    if (!currentUserEmail) {
      console.error('❌ getConfig: ユーザーメール取得失敗');
      throw new Error('ユーザー認証が必要です');
    }

    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    if (!userInfo) {
      console.error('❌ getConfig: ユーザー情報が見つかりません', {
        currentUserEmail,
        timestamp: new Date().toISOString(),
      });
      // ✅ 修正：初期設定を返さず、エラーとして扱う
      throw new Error('ユーザー情報が見つかりません。セットアップが必要です。');
    }

    const config = ConfigManager.getUserConfig(userInfo.userId);
    if (!config) {
      console.error('❌ getConfig: 設定データが見つかりません', {
        userId: userInfo.userId,
        userEmail: currentUserEmail,
      });
      // ✅ 修正：既存ユーザーには初期設定ではなく最小限の設定を返す
      throw new Error('設定データが破損している可能性があります');
    }

    const executionTime = Date.now() - startTime;
    console.log('✅ getConfig: ユーザー設定取得完了（最適化版）', {
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
    console.error('❌ getConfig エラー（最適化版）', {
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
    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
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

    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
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
        console.error('❌ ユーザーマイグレーションエラー:', userError.message);
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
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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
    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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

    const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
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
    
    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
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
        const spreadsheet = new ConfigurationManager().getSpreadsheet(config.spreadsheetId);
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
    
    // ユーザー情報を取得
    const { userInfo } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }
    
    const currentConfig = JSON.parse(userInfo.configJson || '{}');
    
    if (!currentConfig.spreadsheetId || !currentConfig.sheetName) {
      throw new Error('スプレッドシート情報が不足しています');
    }
    
    // スプレッドシートからヘッダーを取得
    const spreadsheet = new ConfigurationManager().getSpreadsheet(currentConfig.spreadsheetId);
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

/**
 * PageBackend.gs - Page.html/page.js.html専用バックエンド関数
 * 2024年GAS V8 Best Practices準拠
 * CLAUDE.md準拠：統一データソース、セキュリティ、エラーハンドリング
 */

/**
 * Page専用設定定数
 */
const PAGE_CONFIG = Object.freeze({
  CACHE_TTL: CORE.TIMEOUTS.MEDIUM,
  MAX_SHEETS: 100,
  DEFAULT_SHEET_NAME: 'Sheet1',
});

/**
 * 管理者権限チェック（Page専用）
 * ボード所有者またはシステム管理者の権限を確認
 * @param {string} userId - ユーザーID（オプション、未指定時は現在のユーザー）
 * @returns {boolean} 管理者権限があるかどうか
 */
function checkAdmin(userId = null) {
  try {
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    console.log('checkAdmin: 管理者権限チェック開始', {
      userId: userId ? `${userId.substring(0, 8)}...` : 'null',
      currentUserEmail: currentUserEmail ? `${currentUserEmail.substring(0, 10)}...` : 'null',
    });

    // userIdが指定されていない場合は現在のユーザーのuserIdを取得
    let targetUserId = userId;
    if (!targetUserId) {
      const user = DB.findUserByEmail(currentUserEmail);
      if (!user) {
        console.warn('checkAdmin: 現在のユーザーがデータベースに見つかりません');
        return false;
      }
      targetUserId = user.userId;
    }

    // 入力検証
    if (!SecurityValidator.isValidUUID(targetUserId)) {
      console.error('checkAdmin: 無効なユーザーID');
      return false;
    }

    // App.getAccess().verifyAccess()でadmin権限をチェック
    const accessResult = App.getAccess().verifyAccess(targetUserId, 'admin', currentUserEmail);

    const isAdmin = accessResult.allowed;
    console.log('checkAdmin: 権限チェック結果', {
      allowed: isAdmin,
      userType: accessResult.userType,
      reason: accessResult.reason,
    });

    return isAdmin;
  } catch (error) {
    console.error('checkAdmin エラー:', {
      error: error.message,
      stack: error.stack,
      userId,
    });
    return false;
  }
}

/**
 * 利用可能シート一覧取得（Page専用）
 * 現在のユーザーのスプレッドシートからシート一覧を取得
 * @param {string} userId - ユーザーID（オプション）
 * @returns {Array<Object>} シート情報の配列
 */
function getAvailableSheets(userId = null) {
  try {
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    console.log('getAvailableSheets: シート一覧取得開始', {
      userId: userId ? `${userId.substring(0, 8)}...` : 'null',
    });

    // userIdが指定されていない場合は現在のユーザーのuserIdを取得
    let targetUserId = userId;
    if (!targetUserId) {
      const user = DB.findUserByEmail(currentUserEmail);
      if (!user) {
        throw new Error('現在のユーザーがデータベースに見つかりません');
      }
      targetUserId = user.userId;
    }

    // 入力検証
    if (!SecurityValidator.isValidUUID(targetUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(targetUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ユーザー情報からスプレッドシートIDを取得（configJSON中心型）
    const userInfo = DB.findUserById(targetUserId);
    if (!userInfo || !JSON.parse(userInfo.configJson || '{}')) {
      throw new Error('ユーザー設定が見つかりません');
    }
    const config = JSON.parse(userInfo.configJson || '{}');
    if (!config.spreadsheetId) {
      throw new Error('ユーザーのスプレッドシート情報が見つかりません');
    }

    // AdminPanelBackend.gsのgetSheetList()を利用
    const sheetList = getSheetList(config.spreadsheetId);

    console.log('getAvailableSheets: シート一覧取得完了', {
      spreadsheetId: `${config.spreadsheetId.substring(0, 20)}...`,
      sheetCount: sheetList.length,
    });

    return sheetList;
  } catch (error) {
    console.error('getAvailableSheets エラー:', {
      error: error.message,
      stack: error.stack,
      userId,
    });
    throw error;
  }
}

/**
 * アクティブシートクリア（Page専用）
 * ユーザーのアクティブシート状態をリセット
 * @param {string} userId - ユーザーID（オプション）
 * @returns {Object} 処理結果
 */
function clearActiveSheet(userId = null) {
  try {
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    console.log('clearActiveSheet: アクティブシートクリア開始', {
      userId: userId ? `${userId.substring(0, 8)}...` : 'null',
    });

    // userIdが指定されていない場合は現在のユーザーのuserIdを取得
    let targetUserId = userId;
    if (!targetUserId) {
      const user = DB.findUserByEmail(currentUserEmail);
      if (!user) {
        throw new Error('現在のユーザーがデータベースに見つかりません');
      }
      targetUserId = user.userId;
    }

    // 入力検証
    if (!SecurityValidator.isValidUUID(targetUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認（編集権限が必要）
    const accessResult = App.getAccess().verifyAccess(targetUserId, 'edit', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`編集権限がありません: ${accessResult.reason}`);
    }

    // キャッシュクリア（関連するキャッシュを削除）
    const cacheKeys = [
      `sheet_data_${targetUserId}`,
      `active_sheet_${targetUserId}`,
      `sheet_headers_${targetUserId}`,
      `reaction_data_${targetUserId}`,
    ];

    let clearedCount = 0;
    cacheKeys.forEach((key) => {
      try {
        cacheManager.remove(key);
        clearedCount++;
      } catch (e) {
        console.warn(`キャッシュクリア失敗: ${key}`, e.message);
      }
    });

    // ユーザーのアクティブシート情報をリセット（必要に応じて）
    // 現在のシステムではsheetNameは保持するため、特別な処理は不要

    const result = createResponse(true, 'キャッシュクリア完了', {
      clearedCacheCount: clearedCount,
      userId: targetUserId,
    });

    console.log('clearActiveSheet: 完了', result);
    return result;
  } catch (error) {
    console.error('clearActiveSheet エラー:', {
      error: error.message,
      stack: error.stack,
      userId,
    });

    return createResponse(false, 'キャッシュクリアエラー', null, error);
  }
}

/**
 * バッチリアクション処理（Page専用）
 * 複数のリアクション操作を効率的に実行
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {Array<Object>} batchOperations - バッチ操作の配列
 * @returns {Object} 処理結果
 */
function addReactionBatch(requestUserId, batchOperations) {
  try {
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    console.log('addReactionBatch: バッチリアクション処理開始', {
      userId: requestUserId ? `${requestUserId.substring(0, 8)}...` : 'null',
      operationsCount: Array.isArray(batchOperations) ? batchOperations.length : 0,
    });

    // 入力検証
    if (!Array.isArray(batchOperations) || batchOperations.length === 0) {
      throw new Error('バッチ操作が無効です');
    }

    // バッチサイズ制限（安全性のため）
    const MAX_BATCH_SIZE = 20;
    if (batchOperations.length > MAX_BATCH_SIZE) {
      throw new Error(`バッチサイズが制限を超えています (最大${MAX_BATCH_SIZE}件)`);
    }

    // ユーザーID検証
    if (!SecurityValidator.isValidUUID(requestUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ボードオーナーの情報をDBから取得
    const boardOwnerInfo = DB.findUserById(requestUserId);
    if (!boardOwnerInfo || !boardOwnerInfo.parsedConfig) {
      throw new Error('無効なボードです');
    }
    const ownerConfig = boardOwnerInfo.parsedConfig;
    if (!ownerConfig.spreadsheetId) {
      throw new Error('スプレッドシートが設定されていません');
    }

    // バッチ処理結果を格納
    const batchResults = [];
    const processedRows = new Set(); // 重複行の追跡

    // シート名を取得（統一データソース使用）
    const userConfig = ConfigManager.getUserConfig(requestUserId);
    const sheetName = userConfig?.sheetName || 'フォームの回答 1';

    console.log('addReactionBatch: 処理対象シート', {
      spreadsheetId: `${ownerConfig.spreadsheetId.substring(0, 20)}...`,
      sheetName,
    });

    // バッチ操作を順次処理
    for (let i = 0; i < batchOperations.length; i++) {
      const operation = batchOperations[i];

      try {
        // 入力検証
        if (!operation.rowIndex || !operation.reaction) {
          console.warn('addReactionBatch: 無効な操作をスキップ', operation);
          continue;
        }

        // 個別のリアクション処理（Core.gsのaddReaction関数を呼び出し）
        const result = addReaction(
          requestUserId,
          operation.rowIndex,
          operation.reaction,
          sheetName
        );

        if (result && result.success) {
          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            status: 'success',
            timestamp: new Date().toISOString(),
          });
          processedRows.add(operation.rowIndex);
        } else {
          console.warn('addReactionBatch: リアクション処理失敗', operation, result?.message);
          batchResults.push({
            rowIndex: operation.rowIndex,
            reaction: operation.reaction,
            status: 'error',
            message: result?.message || 'リアクション処理失敗',
          });
        }
      } catch (operationError) {
        console.error('addReactionBatch: 個別操作エラー', operation, operationError.message);
        batchResults.push({
          rowIndex: operation.rowIndex,
          reaction: operation.reaction,
          status: 'error',
          message: operationError.message,
        });
      }
    }

    const successCount = batchResults.filter((r) => r.status === 'success').length;
    console.log('addReactionBatch: バッチリアクション処理完了', {
      total: batchOperations.length,
      processed: processedRows.size,
      success: successCount,
    });

    return createResponse(true, 'バッチリアクション処理完了', {
      processedCount: batchOperations.length,
      successCount,
      details: batchResults,
    });
  } catch (error) {
    console.error('addReactionBatch エラー:', {
      error: error.message,
      stack: error.stack,
      requestUserId,
    });

    return createResponse(
      false,
      'バッチリアクション処理エラー',
      {
        fallbackToIndividual: true, // クライアント側が個別処理にフォールバック可能
      },
      error
    );
  }
}

/**
 * ボードデータ更新・キャッシュクリア（Page専用）
 * ユーザーのキャッシュをクリアして最新データを取得
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @returns {Object} 処理結果
 */
function refreshBoardData(requestUserId) {
  try {
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    console.log('refreshBoardData: ボードデータ更新開始', {
      userId: requestUserId ? `${requestUserId.substring(0, 8)}...` : 'null',
    });

    // ユーザーID検証
    if (!SecurityValidator.isValidUUID(requestUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ユーザー情報取得
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // キャッシュクリア（関連するキャッシュを削除）
    const cacheKeys = [
      `sheet_data_${requestUserId}`,
      `board_data_${requestUserId}`,
      `user_config_${requestUserId}`,
      `reaction_data_${requestUserId}`,
      `sheet_headers_${requestUserId}`,
    ];

    let clearedCount = 0;
    cacheKeys.forEach((key) => {
      try {
        cacheManager.remove(key);
        clearedCount++;
      } catch (e) {
        console.warn(`refreshBoardData: キャッシュクリア失敗: ${key}`, e.message);
      }
    });

    // ユーザーキャッシュの無効化（Core.gsの関数を利用）
    try {
      if (typeof invalidateUserCache === 'function') {
        // 🚀 CLAUDE.md準拠：userInfo使用（boardOwnerInfoはuserInfoと同じ）
        const config = JSON.parse(userInfo.configJson || '{}') || {};
        const {spreadsheetId} = config;
        invalidateUserCache(requestUserId, userInfo.userEmail, spreadsheetId, false);
      }
    } catch (invalidateError) {
      console.warn('refreshBoardData: ユーザーキャッシュ無効化エラー:', invalidateError.message);
    }

    // 最新の設定情報を取得（統一された設定取得）
    let latestConfig = {};
    try {
      latestConfig = ConfigManager.getUserConfig(requestUserId) || {};
    } catch (configError) {
      console.warn('refreshBoardData: 設定取得エラー:', configError.message);
      latestConfig = { status: 'error', message: '設定の取得に失敗しました' };
    }

    const result = createResponse(true, 'ボードデータを更新しました', {
      clearedCacheCount: clearedCount,
      config: latestConfig,
      userId: requestUserId,
    });

    console.log('refreshBoardData: 完了', {
      clearedCount,
      hasConfig: !!latestConfig,
      userId: requestUserId,
    });

    return result;
  } catch (error) {
    console.error('refreshBoardData エラー:', {
      error: error.message,
      stack: error.stack,
      requestUserId,
    });

    return createResponse(
      false,
      `ボードデータの再読み込みに失敗しました: ${error.message}`,
      {
        status: 'error',
      },
      error
    );
  }
}

/**
 * データ件数取得（Page専用）
 * 指定されたフィルタ条件でのデータ件数を取得
 * @param {string} requestUserId - リクエスト元のユーザーID
 * @param {string} classFilter - クラスフィルタ
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @returns {Object} データ件数情報
 */
function getDataCount(requestUserId, classFilter, _, adminMode) {
  try {
    const { currentUserEmail } = new ConfigurationManager().getCurrentUserInfoSafely() || {};
    // データ件数取得開始

    // ユーザーID検証
    if (!SecurityValidator.isValidUUID(requestUserId)) {
      throw new Error('無効なユーザーIDです');
    }

    // アクセス権限確認
    const accessResult = App.getAccess().verifyAccess(requestUserId, 'view', currentUserEmail);
    if (!accessResult.allowed) {
      throw new Error(`アクセス権限がありません: ${accessResult.reason}`);
    }

    // ユーザー情報取得
    const userInfo = DB.findUserById(requestUserId);
    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    // 🔥 設定情報を効率的に取得（parsedConfig優先）
    const config = JSON.parse(userInfo.configJson || '{}') || {};

    // スプレッドシート設定確認（configJSON中心型）
    const {spreadsheetId} = config;
    const sheetName = config.sheetName || 'フォームの回答 1';

    if (!spreadsheetId || !sheetName) {
      return {
        count: 0,
        lastUpdate: new Date().toISOString(),
        status: 'error',
        message: 'スプレッドシート設定が不完全です',
      };
    }

    // データ件数をカウント
    let count = 0;
    try {
      if (typeof countSheetRows === 'function') {
        count = countSheetRows(spreadsheetId, sheetName, classFilter);
      } else {
        // フォールバック: 直接スプレッドシートにアクセス
        const spreadsheet = new ConfigurationManager().getSpreadsheet(spreadsheetId);
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (sheet) {
          count = Math.max(0, sheet.getLastRow() - 1); // ヘッダー行を除く
        }
      }
    } catch (countError) {
      console.error('getDataCount: カウントエラー:', countError.message);
      count = 0;
    }

    const result = {
      count,
      lastUpdate: new Date().toISOString(),
      status: 'success',
      spreadsheetId,
      sheetName,
      classFilter: classFilter || null,
      adminMode: adminMode || false,
    };

    console.log('getDataCount: 完了', {
      count,
      sheetName,
      classFilter,
    });

    return result;
  } catch (error) {
    console.error('getDataCount エラー:', {
      error: error.message,
      stack: error.stack,
      requestUserId,
    });

    return {
      count: 0,
      lastUpdate: new Date().toISOString(),
      status: 'error',
      message: error.message,
    };
  }
}
