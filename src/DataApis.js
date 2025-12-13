/**
 * @fileoverview DataApis - データ操作専用API
 *
 * ✅ V8ランタイム対応（2022年6月アップデート準拠）
 * - 関数定義は順序に関係なく呼び出し可能
 * - グローバルスコープでのコード実行を完全排除
 *
 * 依存関係（呼び出す関数）:
 * - getCurrentEmail() - main.jsで定義
 * - findUserById() - DatabaseCore.jsで定義
 * - findUserByEmail() - DatabaseCore.jsで定義
 * - getUserConfig() - ConfigService.jsで定義
 * - saveUserConfig() - ConfigService.jsで定義
 * - openSpreadsheet() - DatabaseCore.jsで定義
 * - getSheetInfo() - SystemController.jsで定義
 * - getUserSheetData() - DataService.jsで定義
 * - getBatchedAdminAuth() - main.jsで定義
 * - getQuestionText() - DataService.jsで定義
 * - getFormInfo() - DataService.jsで定義
 * - performIntegratedColumnDiagnostics() - ColumnMappingService.jsで定義
 * - setupDomainWideSharing() - helpers.jsで定義
 * - validateAccess() - DataService.jsで定義
 * - executeWithRetry() - main.jsで定義
 * - createAuthError() - helpers.jsで定義
 * - createUserNotFoundError() - helpers.jsで定義
 * - createErrorResponse() - helpers.jsで定義
 * - createExceptionResponse() - helpers.jsで定義
 *
 * 移動元: main.js
 * 移動日: 2025-12-13
 */

/* global getCurrentEmail, findUserById, findUserByEmail, getUserConfig, saveUserConfig, openSpreadsheet, getSheetInfo, getUserSheetData, getBatchedAdminAuth, getQuestionText, getFormInfo, performIntegratedColumnDiagnostics, setupDomainWideSharing, validateAccess, executeWithRetry, createAuthError, createUserNotFoundError, createErrorResponse, createExceptionResponse, DriveApp, SpreadsheetApp, ScriptApp, URL */

// ================================
// スプレッドシート操作API
// ================================

/**
 * 自分がオーナーのスプレッドシートを30件取得
 * @returns {Object} スプレッドシート一覧
 */
function getSheets() {
  const files = DriveApp.getFilesByType('application/vnd.google-apps.spreadsheet');
  const sheets = [];

  while (files.hasNext() && sheets.length < 30) {
    const file = files.next();
    sheets.push({
      id: file.getId(),
      name: file.getName(),
      url: file.getUrl()
    });
  }

  return { success: true, sheets };
}

/**
 * Validate header integrity for user's active sheet
 * @param {string} targetUserId - 対象ユーザーID（省略可能）
 * @returns {Object} 検証結果
 */
function validateHeaderIntegrity(targetUserId) {
  try {
    const currentEmail = getCurrentEmail();
    let targetUser = targetUserId ? findUserById(targetUserId, { requestingUser: currentEmail }) : null;
    if (!targetUser && currentEmail) {
      targetUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
    }

    if (!targetUser) {
      return createUserNotFoundError();
    }

    const configResult = getUserConfig(targetUser.userId);
    const config = configResult.success ? configResult.config : {};

    if (!config.spreadsheetId || !config.sheetName) {
      return {
        success: false,
        error: 'Spreadsheet configuration incomplete'
      };
    }

    const dataAccess = openSpreadsheet(config.spreadsheetId, {
      useServiceAccount: false
    });
    const sheet = dataAccess.getSheet(config.sheetName);
    if (!sheet) {
      return {
        success: false,
        error: 'Sheet not found'
      };
    }

    const { headers } = getSheetInfo(sheet);
    const normalizedHeaders = headers.map(header => String(header || '').trim());
    const emptyHeaderCount = normalizedHeaders.filter(header => header.length === 0).length;
    const valid = normalizedHeaders.length > 0 && emptyHeaderCount < normalizedHeaders.length;

    return {
      success: valid,
      valid,
      headerCount: normalizedHeaders.length,
      emptyHeaderCount,
      headers: normalizedHeaders,
      error: valid ? null : 'ヘッダー情報が不完全です'
    };
  } catch (error) {
    console.error('validateHeaderIntegrity error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get board info - simplified name
 * @returns {Object} ボード情報
 */
function getBoardInfo() {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.error('Authentication failed');
      return createAuthError();
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      console.error('User not found:', email);
      return { success: false, message: 'User not found' };
    }

    const configResult = getUserConfig(user.userId);
    const config = configResult.success ? configResult.config : {};

    const isPublished = Boolean(config.isPublished);
    const baseUrl = ScriptApp.getService().getUrl();

    return {
      success: true,
      isActive: isPublished,
      isPublished,
      questionText: getQuestionText(config, { targetUserEmail: user.userEmail }),
      urls: {
        view: `${baseUrl}?mode=view&userId=${user.userId}`,
        admin: `${baseUrl}?mode=admin&userId=${user.userId}`
      },
      lastUpdated: config.publishedAt || user.lastModified
    };
  } catch (error) {
    console.error('getBoardInfo ERROR:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * Get sheet list for spreadsheet - simplified name
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getSheetList(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      return createErrorResponse('Spreadsheet ID required');
    }

    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.warn('getSheetList: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    let dataAccess = null;

    try {
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
    } catch (accessError) {
      console.warn('getSheetList: Spreadsheet access failed:', accessError.message);
      return {
        success: false,
        error: 'スプレッドシートにアクセスできません。同一ドメイン内で編集可能に設定されているか確認してください。'
      };
    }

    if (!dataAccess || !dataAccess.spreadsheet) {
      console.warn('getSheetList: Failed to access spreadsheet:', `${spreadsheetId.substring(0, 8)}***`);
      return {
        success: false,
        error: 'スプレッドシートを開くことができませんでした。同一ドメイン内で編集可能に設定されているか確認してください。'
      };
    }

    const { spreadsheet } = dataAccess;
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map(sheet => {
      const { lastRow, lastCol } = getSheetInfo(sheet);
      return {
        name: sheet.getName(),
        id: sheet.getSheetId(),
        rowCount: lastRow,
        columnCount: lastCol
      };
    });

    return {
      success: true,
      sheets: sheetList,
      accessMethod: 'normal_permissions'
    };
  } catch (error) {
    console.error('getSheetList error:', error.message);
    return {
      success: false,
      error: `シート一覧取得エラー: ${error.message}`,
      details: error.stack
    };
  }
}

// ================================
// データ取得・保存API
// ================================

/**
 * 統合API: フロントエンド用データ取得（最適化版・クロスユーザー対応）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @param {string} targetUserId - 対象ユーザーID
 * @returns {Object} フロントエンド期待形式のデータ
 */
function getPublishedSheetData(classFilter, sortOrder, adminMode, targetUserId) {
  classFilter = classFilter || null;
  sortOrder = sortOrder || 'newest';
  adminMode = adminMode || false;
  targetUserId = targetUserId || null;

  const startTime = Date.now();

  try {
    const adminAuth = getBatchedAdminAuth({ allowNonAdmin: true });
    if (!adminAuth.success || !adminAuth.authenticated) {
      return {
        success: false,
        error: 'Authentication required',
        data: [],
        sheetName: '',
        header: '認証エラー'
      };
    }

    const { email: viewerEmail, isAdmin: isSystemAdmin } = adminAuth;

    if (targetUserId) {
      const targetUser = findUserById(targetUserId, {
        requestingUser: viewerEmail,
        preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
      });
      if (!targetUser) {
        console.error('getPublishedSheetData: Target user not found', { targetUserId, viewerEmail });
        return {
          success: false,
          error: 'Target user not found',
          data: [],
          sheetName: '',
          header: 'ユーザーエラー'
        };
      }

      const configResult = getUserConfig(targetUserId, targetUser);
      const targetUserConfig = configResult.success ? configResult.config : {};

      const options = {
        classFilter: classFilter !== 'すべて' ? classFilter : undefined,
        sortBy: sortOrder || 'newest',
        includeTimestamp: true,
        adminMode: isSystemAdmin || (targetUser.userEmail === viewerEmail),
        requestingUser: viewerEmail,
        preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
      };

      const result = getUserSheetData(targetUser.userId, options, targetUser, targetUserConfig);

      if (!result || !result.success) {
        console.error('getPublishedSheetData: getUserSheetData failed', {
          result,
          error: result?.message || 'データ取得エラー',
          totalTime: Date.now() - startTime
        });
        return {
          success: false,
          error: result?.message || 'データ取得エラー',
          data: [],
          sheetName: result?.sheetName || '',
          header: result?.header || '問題'
        };
      }

      const finalResult = {
        success: true,
        data: result.data || [],
        header: result.header || result.sheetName || '回答一覧',
        sheetName: result.sheetName || 'Sheet1'
      };

      try {
        JSON.stringify(finalResult);

        const safeResult = {
          success: true,
          data: Array.isArray(finalResult.data) ? finalResult.data.map(item => {
            if (typeof item === 'object' && item !== null) {
              const cleaned = {};
              for (const [key, value] of Object.entries(item)) {
                if (value instanceof Date) {
                  cleaned[key] = value.toISOString();
                } else if (typeof value === 'object' && value !== null) {
                  try {
                    cleaned[key] = Array.isArray(value) ? [...value] : { ...value };
                  } catch (e) {
                    cleaned[key] = String(value);
                  }
                } else {
                  cleaned[key] = value;
                }
              }
              return cleaned;
            }
            return item;
          }) : [],
          header: String(finalResult.header || '回答一覧'),
          sheetName: String(finalResult.sheetName || 'Sheet1'),
          displaySettings: targetUserConfig.displaySettings || { showNames: false, showReactions: false }
        };

        return safeResult;

      } catch (serializationError) {
        console.error('getPublishedSheetData: Serialization failed', {
          error: serializationError.message,
          dataLength: finalResult.data?.length || 0,
          headerType: typeof finalResult.header,
          sheetNameType: typeof finalResult.sheetName
        });

        return {
          success: true,
          data: [],
          header: '回答一覧（データ変換エラー）',
          sheetName: 'Sheet1',
          displaySettings: targetUserConfig.displaySettings || { showNames: false, showReactions: false }
        };
      }
    }

    const user = findUserByEmail(viewerEmail, {
      requestingUser: viewerEmail,
      adminMode: isSystemAdmin,
      ignorePermissions: isSystemAdmin,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    });

    if (!user) {
      console.error('getPublishedSheetData: User not found (self-access)', { viewerEmail });
      return {
        success: false,
        error: 'User not found',
        data: [],
        sheetName: '',
        header: 'ユーザーエラー'
      };
    }

    const configResult = getUserConfig(user.userId, user);
    const userConfig = configResult.success ? configResult.config : {};

    const options = {
      classFilter: classFilter !== 'すべて' ? classFilter : undefined,
      sortBy: sortOrder || 'newest',
      includeTimestamp: true,
      adminMode: isSystemAdmin,
      requestingUser: viewerEmail,
      preloadedAuth: { email: viewerEmail, isAdmin: isSystemAdmin }
    };

    const result = getUserSheetData(user.userId, options, user, userConfig);

    if (!result || !result.success) {
      console.error('getPublishedSheetData: getUserSheetData failed (self-access)', {
        result,
        error: result?.message || 'データ取得エラー',
        totalTime: Date.now() - startTime
      });
      return {
        success: false,
        error: result?.message || 'データ取得エラー',
        data: [],
        sheetName: result?.sheetName || '',
        header: result?.header || '問題',
      };
    }

    const finalResult = {
      success: true,
      data: result.data || [],
      header: result.header || result.sheetName || '回答一覧',
      sheetName: result.sheetName || 'Sheet1'
    };

    try {
      const testSerialization = JSON.stringify(finalResult);

      const safeResult = {
        success: true,
        data: Array.isArray(finalResult.data) ? finalResult.data.map(item => {
          if (typeof item === 'object' && item !== null) {
            const cleaned = {};
            for (const [key, value] of Object.entries(item)) {
              if (value instanceof Date) {
                cleaned[key] = value.toISOString();
              } else if (typeof value === 'object' && value !== null) {
                try {
                  cleaned[key] = Array.isArray(value) ? [...value] : { ...value };
                } catch (e) {
                  cleaned[key] = String(value);
                }
              } else {
                cleaned[key] = value;
              }
            }
            return cleaned;
          }
          return item;
        }) : [],
        header: String(finalResult.header || '回答一覧'),
        sheetName: String(finalResult.sheetName || 'Sheet1'),
        displaySettings: userConfig.displaySettings || { showNames: false, showReactions: false }
      };

      return safeResult;

    } catch (serializationError) {
      console.error('getPublishedSheetData: Serialization failed (self-access)', {
        error: serializationError.message,
        dataLength: finalResult.data?.length || 0,
        headerType: typeof finalResult.header,
        sheetNameType: typeof finalResult.sheetName
      });

      return {
        success: true,
        data: [],
        header: '回答一覧（データ変換エラー）',
        sheetName: 'Sheet1',
        displaySettings: userConfig.displaySettings || { showNames: false, showReactions: false }
      };
    }
  } catch (error) {
    console.error('getPublishedSheetData: Exception caught', {
      error: error?.message || 'Unknown error',
      stack: error?.stack,
      classFilter,
      sortOrder,
      adminMode,
      targetUserId,
      timestamp: new Date().toISOString()
    });
    return {
      success: false,
      error: error?.message || 'データ取得エラー',
      data: [],
      sheetName: '',
      header: '問題'
    };
  }
}

/**
 * データカウント取得（フロントエンド整合性のため追加）
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @returns {Object} カウント情報
 */
function getDataCount(classFilter, sortOrder, adminMode = false) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      return { error: 'Authentication required', count: 0 };
    }

    const user = findUserByEmail(email, { requestingUser: email });
    if (!user) {
      return { error: 'User not found', count: 0 };
    }

    const result = getUserSheetData(user.userId, {
      classFilter,
      sortOrder,
      adminMode
    });

    if (result.success && result.data) {
      return {
        success: true,
        count: result.data.length,
        sheetName: result.sheetName
      };
    }

    return { error: result.message || 'Failed to get count', count: 0 };
  } catch (error) {
    console.error('getDataCount error:', error.message);
    return { error: error.message, count: 0 };
  }
}

/**
 * GAS-Native統一設定保存API
 * @param {Object} config - 設定オブジェクト
 * @param {Object} options - 保存オプション { isDraft: boolean }
 * @returns {Object} 保存結果
 */
function saveConfig(config, options = {}) {
  const startTime = Date.now();

  try {
    const userEmail = getCurrentEmail();

    if (!userEmail) {
      return { success: false, message: 'ユーザー認証が必要です' };
    }

    const user = findUserByEmail(userEmail, { requestingUser: userEmail });

    if (!user) {
      return { success: false, message: 'ユーザーが見つかりません' };
    }

    const saveOptions = options.isDraft ?
      { isDraft: true } :
      { isMainConfig: true };

    const result = saveUserConfig(user.userId, config, saveOptions);

    return result;

  } catch (error) {
    const duration = Date.now() - startTime;
    console.error(`saveConfig: ERROR after ${duration}ms - ${error.message || 'Operation error'}`);
    return { success: false, message: error.message || 'エラーが発生しました' };
  }
}

/**
 * 時刻ベース統一新着通知システム - 全ユーザーロール対応
 * @param {string} targetUserId - 閲覧対象ユーザーID
 * @param {Object} options - オプション設定
 * @returns {Object} 時刻ベース統一通知更新結果
 */
function getNotificationUpdate(targetUserId, options = {}) {
  try {
    const email = getCurrentEmail();
    if (!email || !targetUserId) {
      return { success: false, message: 'Invalid request' };
    }

    const targetUser = findUserById(targetUserId, { requestingUser: email });
    if (!targetUser) {
      return { success: false, message: 'User not found' };
    }

    const userData = getUserSheetData(targetUserId, {
      includeTimestamp: true,
      classFilter: options.classFilter,
      sortBy: options.sortOrder || 'newest',
      requestingUser: email,
      targetUserEmail: targetUser.userEmail
    });

    if (!userData.success) {
      return { success: false, message: 'Data access failed' };
    }

    const lastUpdate = new Date(options.lastUpdateTime || 0);
    const newItems = userData.data.filter(item => {
      const itemTime = new Date(item.timestamp || 0);
      return itemTime > lastUpdate;
    });

    return {
      success: true,
      hasNewContent: newItems.length > 0,
      newItemsCount: newItems.length,
      timestamp: new Date().toISOString()
    };

  } catch (error) {
    return { success: false, message: error.message };
  }
}

// ================================
// データソース接続API
// ================================

/**
 * Connect to data source - API Gateway function for DataService
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} batchOperations - バッチ操作
 * @returns {Object} 接続結果
 */
function connectDataSource(spreadsheetId, sheetName, batchOperations = null) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.warn('connectDataSource: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    try {
      setupDomainWideSharing(spreadsheetId, email);
    } catch (sharingError) {
      console.warn('connectDataSource: Domain-wide sharing setup failed (non-critical):', sharingError.message);
    }


    if (batchOperations && Array.isArray(batchOperations)) {
      return processDataSourceOperations(spreadsheetId, sheetName, batchOperations);
    }

    return getColumnAnalysis(spreadsheetId, sheetName);

  } catch (error) {
    console.error('connectDataSource error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * バッチ処理でデータソース操作を実行
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} operations - 操作配列
 * @returns {Object} バッチ処理結果
 */
function processDataSourceOperations(spreadsheetId, sheetName, operations) {
  try {
    const results = {
      success: true,
      batchResults: {},
      message: '統合処理完了'
    };

    let columnAnalysisResult = null;

    for (const operation of operations) {
      switch (operation.type) {
        case 'validateAccess':
          if (!columnAnalysisResult) {
            columnAnalysisResult = getColumnAnalysis(spreadsheetId, sheetName);
          }
          results.batchResults.validation = {
            success: columnAnalysisResult.success,
            details: {
              connectionVerified: columnAnalysisResult.success,
              connectionError: columnAnalysisResult.success ? null : columnAnalysisResult.message
            }
          };
          break;
        case 'getFormInfo':
          results.batchResults.formInfo = getFormInfoInternal(spreadsheetId, sheetName);
          break;
        case 'connectDataSource': {
          if (!columnAnalysisResult) {
            columnAnalysisResult = getColumnAnalysis(spreadsheetId, sheetName);
          }

          if (columnAnalysisResult.success) {
            results.mapping = columnAnalysisResult.mapping;
            results.confidence = columnAnalysisResult.confidence;
            results.headers = columnAnalysisResult.headers;
            results.data = columnAnalysisResult.data;
            results.sheet = columnAnalysisResult.sheet;
          } else {
            results.success = false;
            results.error = columnAnalysisResult.errorResponse?.message || columnAnalysisResult.message;
          }
          break;
        }
      }
    }

    return results;
  } catch (error) {
    console.error('processDataSourceOperations error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * 列分析 - API Gateway実装（既存サービス活用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function getColumnAnalysis(spreadsheetId, sheetName) {
  try {
    const email = getCurrentEmail();
    if (!email) {
      console.warn('getColumnAnalysis: Unauthenticated access attempt');
      return {
        success: false,
        error: 'ユーザー認証が必要です'
      };
    }

    let dataAccess;
    try {
      dataAccess = openSpreadsheet(spreadsheetId, { useServiceAccount: false });
      if (!dataAccess) {
        return {
          success: false,
          error: 'スプレッドシートにアクセスできません。同一ドメイン内で編集可能に設定されているか確認してください。'
        };
      }
    } catch (accessError) {
      console.warn('getColumnAnalysis: Spreadsheet access failed for user:', `${email.split('@')[0]}@***`, accessError.message);
      return {
        success: false,
        error: 'スプレッドシートにアクセス権がありません。同一ドメイン内で編集可能に設定されているか確認してください。'
      };
    }

    const sheet = dataAccess.getSheet(sheetName);

    if (!sheet) {
      return { success: false, message: 'Sheet not found' };
    }

    const dataRange = sheet.getDataRange();
    const values = dataRange.getValues();
    const lastRow = values.length;
    const lastCol = values[0]?.length || 0;
    const headers = lastCol > 0 ? values[0] : [];

    let sampleData = [];
    if (lastRow > 1 && lastCol > 0) {
      const sampleSize = Math.min(10, lastRow - 1);
      try {
        const dataRange = sheet.getRange(2, 1, sampleSize, lastCol);
        sampleData = dataRange.getValues();
      } catch (sampleError) {
        console.warn('getColumnAnalysis: サンプルデータ取得失敗:', sampleError.message);
        sampleData = [];
      }
    }

    const diagnostics = performIntegratedColumnDiagnostics(headers, { sampleData });

    try {
      const columnSetupResult = setupReactionAndHighlightColumns(spreadsheetId, sheetName, headers);
      if (columnSetupResult.columnsAdded && columnSetupResult.columnsAdded.length > 0) {
        const updatedSheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName(sheetName);
        const updatedValues = updatedSheet.getDataRange().getValues();
        const updatedHeaders = updatedValues[0] || [];

        return {
          success: true,
          sheet: updatedSheet,
          headers: updatedHeaders,
          data: [],
          mapping: diagnostics.recommendedMapping || {},
          confidence: diagnostics.confidence || {},
          columnsAdded: columnSetupResult.columnsAdded
        };
      }
    } catch (columnError) {
      console.warn('getColumnAnalysis: Column setup failed:', columnError.message);
    }

    return {
      success: true,
      sheet,
      headers,
      data: [],
      mapping: diagnostics.recommendedMapping || {},
      confidence: diagnostics.confidence || {}
    };
  } catch (error) {
    console.error('getColumnAnalysis error:', error.message);
    return createExceptionResponse(error);
  }
}

/**
 * リアクション列とハイライト列を事前セットアップ
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Array} currentHeaders - 現在のヘッダー配列
 * @returns {Object} 追加結果
 */
function setupReactionAndHighlightColumns(spreadsheetId, sheetName, currentHeaders = []) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`Sheet '${sheetName}' not found`);
    }

    const requiredColumns = [
      'UNDERSTAND',
      'LIKE',
      'CURIOUS',
      'HIGHLIGHT'
    ];

    const columnsToAdd = [];
    const currentHeadersUpper = currentHeaders.map(h => String(h).toUpperCase());

    requiredColumns.forEach(columnName => {
      const exists = currentHeadersUpper.some(header =>
        header.includes(columnName.toUpperCase())
      );

      if (!exists) {
        columnsToAdd.push(columnName);
      }
    });

    const columnsAdded = [];

    if (columnsToAdd.length > 0) {
      const { lastCol } = getSheetInfo(sheet);

      columnsToAdd.forEach((columnName, index) => {
        const newColIndex = lastCol + index + 1;

        try {
          sheet.getRange(1, newColIndex).setValue(columnName);
          columnsAdded.push(columnName);
        } catch (colError) {
          console.error(`setupReactionAndHighlightColumns: Failed to add column '${columnName}':`, colError.message);
        }
      });
    }

    return {
      success: true,
      columnsAdded,
      totalColumns: requiredColumns.length,
      alreadyExists: requiredColumns.length - columnsToAdd.length
    };

  } catch (error) {
    console.error('setupReactionAndHighlightColumns error:', error.message);
    return {
      success: false,
      error: error.message,
      columnsAdded: []
    };
  }
}

/**
 * フォーム情報取得（バッチ処理用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} フォーム情報
 */
function getFormInfoInternal(spreadsheetId, sheetName) {
  try {
    if (typeof getFormInfo === 'function') {
      return getFormInfo(spreadsheetId, sheetName);
    } else {
      return {
        success: false,
        message: 'getFormInfo関数が初期化されていません'
      };
    }
  } catch (error) {
    const errorMessage = error && error.message ? error.message : '予期しないエラーが発生しました';
    console.error('getFormInfoInternal error:', errorMessage);
    return {
      success: false,
      message: errorMessage
    };
  }
}

/**
 * Get active form info - フロントエンドエラー修正用
 * @param {string} userId - ユーザーID
 * @returns {Object} フォーム情報
 */
function getActiveFormInfo(userId) {
  try {
    const currentEmail = getCurrentEmail();
    if (!currentEmail) {
      console.error('Authentication failed');
      return {
        success: false,
        message: 'Authentication required',
        formUrl: null,
        formTitle: null
      };
    }

    let targetUserId = userId;
    if (!targetUserId) {
      const currentUser = findUserByEmail(currentEmail, { requestingUser: currentEmail });
      if (!currentUser) {
        console.error('Current user not found:', currentEmail);
        return {
          success: false,
          message: 'User not found',
          formUrl: null,
          formTitle: null
        };
      }
      targetUserId = currentUser.userId;
    }

    const configResult = getUserConfig(targetUserId);
    const config = configResult.success ? configResult.config : {};

    const hasFormUrl = !!(config.formUrl && config.formUrl.trim());
    const isValidUrl = hasFormUrl && isValidFormUrl(config.formUrl);

    return {
      success: hasFormUrl,
      shouldShow: hasFormUrl,
      formUrl: hasFormUrl ? config.formUrl : null,
      formTitle: config.formTitle || 'フォーム',
      isValidUrl,
      message: hasFormUrl ?
        (isValidUrl ? 'Valid form found' : 'Form URL found but validation failed') :
        'No form URL configured'
    };
  } catch (error) {
    console.error('getActiveFormInfo ERROR:', error.message);
    return {
      success: false,
      shouldShow: false,
      message: error.message,
      formUrl: null,
      formTitle: null
    };
  }
}

/**
 * Check if URL is valid Google Forms URL
 * @param {string} url - URL to validate
 * @returns {boolean} True if valid Google Forms URL
 */
function isValidFormUrl(url) {
  if (!url || typeof url !== 'string') {
    return false;
  }

  try {
    const urlObj = new URL(url.trim());

    if (urlObj.protocol !== 'https:') {
      return false;
    }

    const validHosts = ['docs.google.com', 'forms.gle'];
    const isValidHost = validHosts.includes(urlObj.hostname);

    if (urlObj.hostname === 'docs.google.com') {
      return urlObj.pathname.includes('/forms/') || urlObj.pathname.includes('/viewform');
    }

    return isValidHost;
  } catch {
    return false;
  }
}

// ================================
// URL解析・検証API
// ================================

/**
 * スプレッドシートURL解析 - GAS-Native Implementation
 * @param {string} fullUrl - 完全なスプレッドシートURL（gid含む）
 * @returns {Object} 解析結果 {spreadsheetId, gid}
 */
function extractSpreadsheetInfo(fullUrl) {
  try {
    if (!fullUrl || typeof fullUrl !== 'string') {
      return {
        success: false,
        message: 'Invalid URL provided'
      };
    }

    const spreadsheetIdMatch = fullUrl.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    const gidMatch = fullUrl.match(/[#&]gid=(\d+)/);

    if (!spreadsheetIdMatch) {
      return {
        success: false,
        message: 'Invalid Google Sheets URL format'
      };
    }

    return {
      success: true,
      spreadsheetId: spreadsheetIdMatch[1],
      gid: gidMatch ? gidMatch[1] : '0'
    };
  } catch (error) {
    console.error('extractSpreadsheetInfo error:', error.message);
    return {
      success: false,
      message: `URL parsing error: ${error.message}`
    };
  }
}

/**
 * GIDからシート名取得 - Zero-Dependency + Batch Operations
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} gid - シートGID
 * @returns {string} シート名
 */
function getSheetNameFromGid(spreadsheetId, gid) {
  try {
    const spreadsheet = executeWithRetry(
      () => SpreadsheetApp.openById(spreadsheetId),
      { operationName: 'SpreadsheetApp.openById', maxRetries: 3 }
    );
    const sheets = spreadsheet.getSheets();

    const sheetInfos = sheets.map(sheet => ({
      name: sheet.getName(),
      gid: sheet.getSheetId().toString()
    }));

    const targetSheet = sheetInfos.find(info => info.gid === gid);
    const resultName = targetSheet ? targetSheet.name : sheetInfos[0]?.name || 'Sheet1';

    return resultName;

  } catch (error) {
    console.error('getSheetNameFromGid error:', error.message);
    return 'Sheet1';
  }
}

/**
 * 完全URL統合検証 - 既存API活用 + Performance最適化
 * @param {string} fullUrl - 完全なスプレッドシートURL
 * @returns {Object} 統合検証結果
 */
function validateCompleteSpreadsheetUrl(fullUrl) {
  const started = Date.now();
  try {
    const parseResult = extractSpreadsheetInfo(fullUrl);
    if (!parseResult.success) {
      return parseResult;
    }

    const { spreadsheetId, gid } = parseResult;
    const sheetName = getSheetNameFromGid(spreadsheetId, gid);
    const accessResult = validateAccess(spreadsheetId, true);

    const formResult = typeof getFormInfo === 'function' ?
      getFormInfo(spreadsheetId, sheetName) :
      { success: false, message: 'getFormInfo関数が初期化されていません' };

    const result = {
      success: true,
      spreadsheetId,
      gid,
      sheetName,
      hasAccess: accessResult.success,
      accessInfo: {
        spreadsheetName: accessResult.spreadsheetName,
        sheets: accessResult.sheets || []
      },
      formInfo: formResult,
      readyToConnect: accessResult.success && sheetName,
      executionTime: `${Date.now() - started}ms`
    };

    return result;

  } catch (error) {
    const errorResult = {
      success: false,
      message: `Complete validation error: ${error.message}`,
      error: error.message,
      executionTime: `${Date.now() - started}ms`
    };

    console.error('validateCompleteSpreadsheetUrl ERROR:', errorResult);
    return errorResult;
  }
}
