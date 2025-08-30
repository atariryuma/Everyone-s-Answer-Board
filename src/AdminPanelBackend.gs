/**
 * @fileoverview AdminPanel バックエンド関数
 * 既存のシステムと統合する最小限の実装
 */

/**
 * 管理パネル用のサーバーサイド関数群
 * 既存のunifiedUserManager、database、Core関数を活用
 */

// =============================================================================
// データソース管理
// =============================================================================

/**
 * スプレッドシート一覧を取得
 * @returns {Array<Object>} スプレッドシート情報の配列
 */
function getSpreadsheetList() {
  try {
    console.log('getSpreadsheetList: スプレッドシート一覧取得開始');

    // Google Driveからスプレッドシートを検索
    const files = DriveApp.searchFiles(
      'mimeType="application/vnd.google-apps.spreadsheet" and trashed=false'
    );

    const spreadsheets = [];
    let count = 0;
    const maxResults = 50; // パフォーマンス制限

    while (files.hasNext() && count < maxResults) {
      const file = files.next();
      spreadsheets.push({
        id: file.getId(),
        name: file.getName(),
        lastModified: file.getLastUpdated().toISOString(),
        owner: file.getOwner().getEmail(),
      });
      count++;
    }

    // 最終更新順でソート
    spreadsheets.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));

    console.log(`getSpreadsheetList: ${spreadsheets.length}個のスプレッドシートを取得`);
    return spreadsheets;
  } catch (error) {
    console.error('getSpreadsheetList エラー:', error);
    throw new Error('スプレッドシート一覧の取得に失敗しました: ' + error.message);
  }
}

/**
 * 指定されたスプレッドシートのシート一覧を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Array<Object>} シート情報の配列
 */
function getSheetList(spreadsheetId) {
  try {
    console.log('getSheetList: シート一覧取得開始', spreadsheetId);

    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが指定されていません');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();

    const sheetList = sheets.map((sheet) => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden(),
    }));

    console.log(`getSheetList: ${sheetList.length}個のシートを取得`);
    return sheetList;
  } catch (error) {
    console.error('getSheetList エラー:', error);
    throw new Error('シート一覧の取得に失敗しました: ' + error.message);
  }
}

/**
 * データソースに接続し、列マッピングを自動検出
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 接続結果と列マッピング情報
 */
function connectDataSource(spreadsheetId, sheetName) {
  try {
    console.log('connectToDataSource: データソース接続開始', { spreadsheetId, sheetName });

    if (!spreadsheetId || !sheetName) {
      throw new Error('スプレッドシートIDとシート名が必要です');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    // 既存の堅牢なヘッダー取得関数を活用（30分キャッシュ + リトライ機能）
    const headerIndices = getHeadersCached(spreadsheetId, sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Core.gsの高精度AI列マッピングを直接活用
    const columnMapping = headerIndices ? convertIndicesToMapping(headerIndices, headerRow) : mapColumns(headerRow);

    // 不足列の検出・追加
    const missingColumnsResult = addMissingColumns(spreadsheetId, sheetName, columnMapping);
    console.log('connectToDataSource: 不足列検出結果', missingColumnsResult);

    // 列が追加された場合は、ヘッダー行を再取得して列マッピングを更新
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      const updatedHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 更新後のヘッダーで再実行（キャッシュを削除して最新取得）
      cacheManager.remove(`hdr_${spreadsheetId}_${sheetName}`);
      const updatedHeaderIndices = getHeadersCached(spreadsheetId, sheetName);
      columnMapping = updatedHeaderIndices ? convertIndicesToMapping(updatedHeaderIndices, updatedHeaderRow) : mapColumns(updatedHeaderRow);
    }

    // 設定を保存（既存のユーザー管理システムを活用）
    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (userInfo) {
      // 既存システム互換の列マッピングを作成
      const compatibleMapping = convertToCompatibleMapping(columnMapping, headerRow);

      // ユーザーのスプレッドシート設定を更新
      updateUserSpreadsheetConfig(userInfo.userId, {
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        columnMapping: columnMapping, // AdminPanel用の形式
        compatibleMapping: compatibleMapping, // Core.gs互換形式
        lastConnected: new Date().toISOString(),
        connectionMethod: 'dropdown_select',
        missingColumnsHandled: missingColumnsResult
      });

      console.log('connectToDataSource: 互換形式も保存', { columnMapping, compatibleMapping });
    }

    console.log('connectToDataSource: 接続成功', columnMapping);
    
    // メッセージを統合
    let message = 'データソースに正常に接続されました';
    if (missingColumnsResult.success) {
      if (missingColumnsResult.addedColumns.length > 0) {
        message += `。${missingColumnsResult.addedColumns.length}個の必須列を自動追加しました`;
      }
      if (missingColumnsResult.recommendedColumns && missingColumnsResult.recommendedColumns.length > 0) {
        message += `。${missingColumnsResult.recommendedColumns.length}個の推奨列を手動で追加することをお勧めします`;
      }
    }
    
    return {
      success: true,
      columnMapping: columnMapping,
      headers: headerRow,              // 🔥 追加: 実際のヘッダー情報
      rowCount: sheet.getLastRow(),
      message: message,
      missingColumnsResult: missingColumnsResult
    };
  } catch (error) {
    console.error('connectToDataSource エラー:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * ヘッダー行から列マッピングを自動検出
 * @param {Array<string>} headers - ヘッダー行の配列
 * @returns {Object} 検出された列マッピング
 */
function mapColumns(headers) {
  const mapping = {
    question: null,
    answerer: null,
    reason: null,
    confidence: {},
  };

  // キーワードパターン
  const patterns = {
    question: ['質問', 'question', 'query', 'ask', 'Q'],
    answerer: ['回答者', '名前', 'name', 'answerer', 'user', 'ユーザー'],
    reason: ['理由', 'reason', 'comment', 'コメント', '備考', 'note'],
  };

  headers.forEach((header, index) => {
    const headerLower = header.toString().toLowerCase();

    Object.keys(patterns).forEach((type) => {
      const pattern = patterns[type];
      const matchCount = pattern.filter((p) => headerLower.includes(p.toLowerCase())).length;

      if (matchCount > 0) {
        const confidence = Math.min(95, 60 + matchCount * 15);
        if (!mapping[type] || confidence > (mapping.confidence[type] || 0)) {
          mapping[type] = index;
          mapping.confidence[type] = confidence;
        }
      }
    });
  });

  return mapping;
}

/**
 * 必要な列が不足していないか検出し、必要に応じて自動追加
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {Object} columnMapping - 現在の列マッピング
 * @returns {Object} 検出・追加結果
 */
function addMissingColumns(spreadsheetId, sheetName, columnMapping) {
  try {
    console.log('detectAndAddMissingColumns: 不足列の検出開始', { spreadsheetId, sheetName, columnMapping });

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    // 現在のヘッダー行を取得
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    console.log('detectAndAddMissingColumns: 現在のヘッダー', headerRow);

    // 必要な列を定義（StudyQuestシステムで使用される標準列）
    const requiredColumns = {
      'メールアドレス': 'EMAIL',
      '理由': 'REASON',
      '名前': 'NAME', 
      'なるほど！': 'UNDERSTAND',
      'いいね！': 'LIKE',
      'もっと知りたい！': 'CURIOUS',
      'ハイライト': 'HIGHLIGHT',
      'タイムスタンプ': 'TIMESTAMP'
    };

    // 不足している列を検出
    const missingColumns = [];
    const existingColumns = headerRow.map(h => String(h || '').trim());

    Object.keys(requiredColumns).forEach(requiredCol => {
      const found = existingColumns.some(existing => {
        const existingLower = existing.toLowerCase();
        const requiredLower = requiredCol.toLowerCase();
        
        // 基本的な検出ロジック（リアクション・ハイライトは完全一致）
        if (requiredCol === 'なるほど！' || requiredCol === 'いいね！' || requiredCol === 'もっと知りたい！' || requiredCol === 'ハイライト') {
          // システム列は完全一致のみ
          return existing === requiredCol;
        }
        
        // その他の列は柔軟な検出
        return existingLower.includes(requiredLower) || 
               (requiredCol === 'メールアドレス' && existingLower.includes('email')) ||
               (requiredCol === '理由' && (existingLower.includes('理由') || existingLower.includes('詳細'))) ||
               (requiredCol === '名前' && (existingLower.includes('名前') || existingLower.includes('氏名'))) ||
               (requiredCol === 'タイムスタンプ' && existingLower.includes('timestamp'));
      });

      if (!found) {
        missingColumns.push({
          columnName: requiredCol,
          systemName: requiredColumns[requiredCol],
          priority: getPriority(requiredCol)
        });
      }
    });

    console.log('detectAndAddMissingColumns: 不足列検出結果', missingColumns);

    // 不足列がない場合
    if (missingColumns.length === 0) {
      return {
        success: true,
        missingColumns: [],
        addedColumns: [],
        message: '必要な列がすべて揃っています'
      };
    }

    // 優先度順にソート
    missingColumns.sort((a, b) => a.priority - b.priority);

    // 列を自動追加（高優先度のみ）
    const addedColumns = [];
    const highPriorityColumns = missingColumns.filter(col => col.priority <= 2);

    if (highPriorityColumns.length > 0) {
      const lastColumn = sheet.getLastColumn();
      
      highPriorityColumns.forEach((colInfo, index) => {
        const newColumnIndex = lastColumn + index + 1;
        
        // 新しい列を追加
        sheet.insertColumnAfter(lastColumn + index);
        sheet.getRange(1, newColumnIndex).setValue(colInfo.columnName);
        
        addedColumns.push({
          columnName: colInfo.columnName,
          position: newColumnIndex,
          systemName: colInfo.systemName
        });
        
        console.log('detectAndAddMissingColumns: 列追加', {
          columnName: colInfo.columnName,
          position: newColumnIndex
        });
      });
    }

    // 残りの不足列（低優先度）は推奨として返す
    const recommendedColumns = missingColumns.filter(col => col.priority > 2);

    return {
      success: true,
      missingColumns: missingColumns,
      addedColumns: addedColumns,
      recommendedColumns: recommendedColumns,
      message: `${addedColumns.length}個の必須列を自動追加しました`,
      details: {
        added: addedColumns.length,
        recommended: recommendedColumns.length,
        total: missingColumns.length
      }
    };

  } catch (error) {
    console.error('detectAndAddMissingColumns エラー:', error);
    return {
      success: false,
      error: error.message,
      missingColumns: [],
      addedColumns: []
    };
  }
}

/**
 * 列の優先度を取得
 * @param {string} columnName - 列名
 * @returns {number} 優先度（低い値ほど高優先度）
 */
function getPriority(columnName) {
  const priorities = {
    'メールアドレス': 1,    // 最高優先度
    '理由': 1,            // 最高優先度  
    '名前': 2,            // 高優先度
    'タイムスタンプ': 3,    // 中優先度
    'なるほど！': 4,       // 低優先度
    'いいね！': 4,
    'もっと知りたい！': 4,
    'ハイライト': 4
  };
  
  return priorities[columnName] || 5;
}

/**
 * スプレッドシートへのアクセス権限を検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} 検証結果
 */
function validateAccess(spreadsheetId) {
  try {
    console.log('validateSpreadsheetAccess: アクセス権限検証開始', spreadsheetId);

    if (!spreadsheetId) {
      throw new Error('スプレッドシートIDが指定されていません');
    }

    // スプレッドシートのアクセス権限を確認
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const spreadsheetName = spreadsheet.getName();
    
    // シート一覧も取得
    const sheets = spreadsheet.getSheets();
    const sheetList = sheets.map((sheet) => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn(),
      hidden: sheet.isSheetHidden(),
    }));

    console.log('validateSpreadsheetAccess: アクセス権限確認成功', {
      id: spreadsheetId,
      name: spreadsheetName,
      sheets: sheetList.length
    });

    return {
      success: true,
      spreadsheetId: spreadsheetId,
      spreadsheetName: spreadsheetName,
      sheets: sheetList,
      message: 'スプレッドシートへのアクセスが確認できました'
    };
  } catch (error) {
    console.error('validateSpreadsheetAccess エラー:', error);
    return {
      success: false,
      error: error.message,
      details: 'スプレッドシートが存在しないか、アクセス権限がありません'
    };
  }
}

/**
 * スプレッドシートの列構造を分析（URL入力方式用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  try {
    console.log('analyzeSpreadsheetColumns: 列分析開始', { spreadsheetId, sheetName });

    if (!spreadsheetId || !sheetName) {
      throw new Error('スプレッドシートIDとシート名が必要です');
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    // 既存の堅牢なヘッダー取得関数を活用（30分キャッシュ + リトライ機能）
    const headerIndices = getHeadersCached(spreadsheetId, sheetName);
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Core.gsの高精度AI列マッピングを直接活用
    const columnMapping = headerIndices ? convertIndicesToMapping(headerIndices, headerRow) : mapColumns(headerRow);
    console.log('analyzeSpreadsheetColumns: 列マッピング完了', columnMapping);

    // 不足列の検出・追加
    const missingColumnsResult = addMissingColumns(spreadsheetId, sheetName, columnMapping);
    console.log('analyzeSpreadsheetColumns: 不足列検出結果', missingColumnsResult);

    // 列が追加された場合は、ヘッダー行を再取得して列マッピングを更新
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      const updatedHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      
      // 更新後のヘッダーで再実行（キャッシュを削除して最新取得）
      cacheManager.remove(`hdr_${spreadsheetId}_${sheetName}`);
      const updatedHeaderIndices = getHeadersCached(spreadsheetId, sheetName);
      columnMapping = updatedHeaderIndices ? convertIndicesToMapping(updatedHeaderIndices, updatedHeaderRow) : mapColumns(updatedHeaderRow);
      
      console.log('analyzeSpreadsheetColumns: 列追加後の更新されたマッピング', columnMapping);
    }

    // 設定を保存（既存システム互換）
    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (userInfo) {
      updateUserSpreadsheetConfig(userInfo.userId, {
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        columnMapping: columnMapping,
        lastConnected: new Date().toISOString(),
        connectionMethod: 'url_input',
        missingColumnsHandled: missingColumnsResult
      });
    }

    console.log('analyzeSpreadsheetColumns: 分析完了', columnMapping);
    
    // メッセージを統合
    let message = 'スプレッドシートの列構造を分析しました';
    if (missingColumnsResult.success) {
      if (missingColumnsResult.addedColumns.length > 0) {
        message += `。${missingColumnsResult.addedColumns.length}個の必須列を自動追加しました`;
      }
      if (missingColumnsResult.recommendedColumns && missingColumnsResult.recommendedColumns.length > 0) {
        message += `。${missingColumnsResult.recommendedColumns.length}個の推奨列を手動で追加することをお勧めします`;
      }
    }
    
    return {
      success: true,
      columnMapping: columnMapping,
      headers: headerRow,
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn(),
      message: message,
      missingColumnsResult: missingColumnsResult
    };
  } catch (error) {
    console.error('analyzeSpreadsheetColumns エラー:', error);
    return {
      success: false,
      error: error.message
    };
  }
}

/**
 * Core.gsのヘッダーインデックスをAdminPanel用マッピングに変換
 * @param {Object} headerIndices - getHeadersCachedから返されるインデックス
 * @param {Array} headerRow - ヘッダー行（表示用）
 * @returns {Object} AdminPanel用マッピング
 */
function convertIndicesToMapping(headerIndices, headerRow) {
  const mapping = {
    question: headerIndices[COLUMN_HEADERS.OPINION] || null,  // '回答' -> question として使用
    answer: headerIndices[COLUMN_HEADERS.REASON] || null,     // '理由' -> answer として使用
    nickname: headerIndices[COLUMN_HEADERS.NAME] || null,    // '名前' -> nickname として使用
    timestamp: headerIndices[COLUMN_HEADERS.TIMESTAMP] || null, // 'タイムスタンプ'
    reason: headerIndices[COLUMN_HEADERS.REASON] || null,    // '理由'
    category: headerIndices[COLUMN_HEADERS.CLASS] || null,   // 'クラス' -> category として使用
  };

  console.log('convertIndicesToMapping: 変換結果', { headerIndices, mapping });
  return mapping;
}

/**
 * guessHeadersFromArray結果をAdminPanel用マッピングに変換（廃止予定）
 * @deprecated convertIndicesToMappingを使用してください
 * @param {Object} guessedHeaders - guessHeadersFromArrayの結果
 * @param {Array} headerRow - ヘッダー行
 * @returns {Object} AdminPanel用マッピング
 */
function convertGuessedToMapping(guessedHeaders, headerRow) {
  const mapping = {
    question: null,
    answerer: null,
    reason: null,
    name: null,
    class: null,
    understand: null,
    like: null,
    curious: null,
    highlight: null,
    confidence: {}
  };

  // guessedHeadersからマッピングを作成
  if (guessedHeaders.questionHeader) {
    const questionIndex = headerRow.findIndex(h => h === guessedHeaders.questionHeader);
    if (questionIndex !== -1) {
      mapping.question = questionIndex;
      mapping.confidence.question = 95; // 高精度AI判定の信頼度
    }
  }

  if (guessedHeaders.answerHeader) {
    const answerIndex = headerRow.findIndex(h => h === guessedHeaders.answerHeader);
    if (answerIndex !== -1) {
      mapping.answerer = answerIndex;
      mapping.confidence.answerer = 95;
    }
  }

  if (guessedHeaders.reasonHeader) {
    const reasonIndex = headerRow.findIndex(h => h === guessedHeaders.reasonHeader);
    if (reasonIndex !== -1) {
      mapping.reason = reasonIndex;
      mapping.confidence.reason = 95;
    }
  }

  if (guessedHeaders.nameHeader) {
    const nameIndex = headerRow.findIndex(h => h === guessedHeaders.nameHeader);
    if (nameIndex !== -1) {
      mapping.name = nameIndex;
      mapping.confidence.name = 95;
    }
  }

  if (guessedHeaders.classHeader) {
    const classIndex = headerRow.findIndex(h => h === guessedHeaders.classHeader);
    if (classIndex !== -1) {
      mapping.class = classIndex;
      mapping.confidence.class = 95;
    }
  }

  // リアクション列のマッピング
  if (guessedHeaders.understandHeader) {
    const understandIndex = headerRow.findIndex(h => h === guessedHeaders.understandHeader);
    if (understandIndex !== -1) {
      mapping.understand = understandIndex;
      mapping.confidence.understand = 95;
    }
  }

  if (guessedHeaders.likeHeader) {
    const likeIndex = headerRow.findIndex(h => h === guessedHeaders.likeHeader);
    if (likeIndex !== -1) {
      mapping.like = likeIndex;
      mapping.confidence.like = 95;
    }
  }

  if (guessedHeaders.curiousHeader) {
    const curiousIndex = headerRow.findIndex(h => h === guessedHeaders.curiousHeader);
    if (curiousIndex !== -1) {
      mapping.curious = curiousIndex;
      mapping.confidence.curious = 95;
    }
  }

  if (guessedHeaders.highlightHeader) {
    const highlightIndex = headerRow.findIndex(h => h === guessedHeaders.highlightHeader);
    if (highlightIndex !== -1) {
      mapping.highlight = highlightIndex;
      mapping.confidence.highlight = 95;
    }
  }

  return mapping;
}

// =============================================================================
// システム監視・メトリクス
// =============================================================================

// =============================================================================
// 設定管理
// =============================================================================

/**
 * 現在の設定情報を取得（sheetName情報強化版）
 * @returns {Object} 現在の設定情報
 */
function getCurrentConfig() {
  try {
    console.log('getCurrentConfig: 設定情報取得開始');

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      return {
        setupStatus: 'pending',
        spreadsheetId: null,
        sheetName: null,
        formCreated: false,
        appPublished: false,
        lastUpdated: null,
        user: currentUser,
      };
    }

    // 既存のdetermineSetupStep関数を活用
    const configJson = getUserConfigJson(userInfo.userId);
    const setupStep = determineSetupStep(userInfo, configJson);

    // sheetName情報の拡張取得
    let sheetName = userInfo.sheetName;
    if (!sheetName && configJson) {
      // configJsonから推奨シート名を取得
      sheetName = configJson.publishedSheetName || configJson.activeSheetName || null;
    }
    
    // まだsheetNameが不足している場合、スプレッドシートから自動検出
    if (!sheetName && userInfo.spreadsheetId) {
      try {
        console.log('getCurrentConfig: シート名自動検出を実行');
        sheetName = detectActiveSheetName(userInfo.spreadsheetId);
      } catch (detectionError) {
        console.warn('getCurrentConfig: シート名自動検出失敗', detectionError.message);
        sheetName = 'フォームの回答 1'; // フォールバック
      }
    }

    const config = {
      setupStatus: getSetupStatusFromStep(setupStep),
      spreadsheetId: userInfo.spreadsheetId,
      sheetName: sheetName,              // 🔥 強化されたシート名情報
      formCreated: configJson ? configJson.formCreated : false,
      appPublished: configJson ? configJson.appPublished : false,
      lastUpdated: userInfo.lastUpdated,
      setupStep: setupStep,
      setupComplete: setupStep >= 3,     // 🔥 追加: セットアップ完了状態
      user: currentUser,
      userId: userInfo.userId,
      displaySettings: {                 // 🔥 追加: 表示設定
        showNames: configJson ? configJson.showNames !== false : true,
        showReactions: configJson ? configJson.showReactions !== false : true
      }
    };

    console.log('getCurrentConfig: 設定情報取得完了', config);
    return config;
  } catch (error) {
    console.error('getCurrentConfig エラー:', error);
    return {
      setupStatus: 'error',
      error: error.message,
      lastUpdated: new Date().toISOString(),
    };
  }
}

/**
 * スプレッドシートからアクティブなシート名を自動検出
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string} 推奨シート名
 */
function detectActiveSheetName(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    
    // 優先順位: フォームの回答 > データがあるシート > 最初のシート
    const priorityNames = ['フォームの回答 1', 'フォームの回答', 'Sheet1', 'シート1'];
    
    // 優先名から検索
    for (const priorityName of priorityNames) {
      const sheet = sheets.find(s => s.getName() === priorityName);
      if (sheet && sheet.getLastRow() > 1) { // データがあるかチェック
        console.log('detectActiveSheetName: 優先シート検出', priorityName);
        return priorityName;
      }
    }
    
    // データが最も多いシートを選択
    const sheetsWithData = sheets
      .filter(s => s.getLastRow() > 1)
      .sort((a, b) => b.getLastRow() - a.getLastRow());
    
    if (sheetsWithData.length > 0) {
      const selectedSheet = sheetsWithData[0].getName();
      console.log('detectActiveSheetName: データ最大シート選択', selectedSheet);
      return selectedSheet;
    }
    
    // フォールバック: 最初のシート
    const fallbackSheet = sheets[0].getName();
    console.log('detectActiveSheetName: フォールバック選択', fallbackSheet);
    return fallbackSheet;
    
  } catch (error) {
    console.error('detectActiveSheetName エラー:', error);
    throw error;
  }
}

/**
 * セットアップステップからステータス文字列を取得
 * @param {number} step - セットアップステップ
 * @returns {string} ステータス文字列
 */
function getSetupStatusFromStep(step) {
  switch (step) {
    case 1:
      return 'pending';
    case 2:
      return 'configuring';
    case 3:
      return 'completed';
    default:
      return 'unknown';
  }
}

/**
 * 下書き設定を保存
 * @param {Object} config - 保存する設定
 * @returns {Object} 保存結果
 */
function saveDraftConfiguration(config) {
  try {
    console.log('saveDraftConfiguration: 下書き保存開始', config);

    const currentUser = User.email();
    let userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      // 新規ユーザーの場合は作成（既存のシステムを活用）
      throw new Error('ユーザー情報が見つかりません。先にユーザー登録を行ってください。');
    }

    // 設定を更新
    const updateResult = updateUserSpreadsheetConfig(userInfo.userId, {
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      displaySettings: {
        showNames: config.showNames,
        showReactions: config.showReactions,
      },
      lastDraftSave: new Date().toISOString(),
    });

    if (updateResult.success) {
      const updatedConfig = getCurrentConfig();
      return {
        success: true,
        config: updatedConfig,
        message: '下書きが保存されました',
      };
    } else {
      throw new Error(updateResult.error);
    }
  } catch (error) {
    console.error('saveDraftConfiguration エラー:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * アプリケーションを公開
 * @param {Object} config - 公開する設定
 * @returns {Object} 公開結果
 */
function publishApplication(config) {
  try {
    console.log('publishApplication: アプリ公開開始', config);

    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (!userInfo) {
      throw new Error('ユーザー情報が見つかりません');
    }

    if (!config.spreadsheetId || !config.sheetName) {
      throw new Error('データソースが設定されていません');
    }

    // 既存の公開システムを活用（簡略化）
    const publishResult = executeAppPublish(userInfo.userId, {
      appName: config.appName,
      spreadsheetId: config.spreadsheetId,
      sheetName: config.sheetName,
      displaySettings: {
        showNames: config.showNames,
        showReactions: config.showReactions,
      },
    });

    if (publishResult.success) {
      const updatedConfig = getCurrentConfig();
      return {
        success: true,
        config: updatedConfig,
        appUrl: publishResult.appUrl,
        message: 'アプリケーションが正常に公開されました',
      };
    } else {
      throw new Error(publishResult.error);
    }
  } catch (error) {
    console.error('publishApplication エラー:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

// =============================================================================
// ヘルパー関数
// =============================================================================

/**
 * ユーザーのスプレッドシート設定を更新
 * @param {string} userId - ユーザーID
 * @param {Object} config - 更新する設定
 * @returns {Object} 更新結果
 */
function updateUserSpreadsheetConfig(userId, config) {
  try {
    const props = PropertiesService.getScriptProperties();
    const configKey = `user_config_${userId}`;

    // 既存の設定を取得
    let existingConfig = {};
    try {
      const existingConfigStr = props.getProperty(configKey);
      if (existingConfigStr) {
        existingConfig = JSON.parse(existingConfigStr);
      }
    } catch (parseError) {
      console.warn('既存設定の解析エラー:', parseError);
    }

    // 設定をマージ
    const mergedConfig = {
      ...existingConfig,
      ...config,
      lastUpdated: new Date().toISOString(),
    };

    // 設定を保存
    props.setProperty(configKey, JSON.stringify(mergedConfig));

    console.log('updateUserSpreadsheetConfig: 設定更新完了', { userId, config: mergedConfig });
    return {
      success: true,
      config: mergedConfig,
    };
  } catch (error) {
    console.error('updateUserSpreadsheetConfig エラー:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * ユーザーの設定JSONを取得
 * @param {string} userId - ユーザーID
 * @returns {Object|null} 設定JSON
 */
function getUserConfigJson(userId) {
  try {
    const props = PropertiesService.getScriptProperties();
    const configKey = `user_config_${userId}`;
    const configStr = props.getProperty(configKey);

    if (!configStr) {
      return null;
    }

    return JSON.parse(configStr);
  } catch (error) {
    console.error('getUserConfigJson エラー:', error);
    return null;
  }
}

/**
 * アプリ公開を実行（簡略化実装）
 * @param {string} userId - ユーザーID
 * @param {Object} publishConfig - 公開設定
 * @returns {Object} 公開結果
 */
function executeAppPublish(userId, publishConfig) {
  try {
    // 既存のWebアプリURLを生成または取得
    const webAppUrl = getOrCreateWebAppUrl(userId, publishConfig.appName);

    // 公開状態をPropertiesServiceに記録
    const props = PropertiesService.getScriptProperties();
    props.setProperty(`${userId}_published`, 'true');
    props.setProperty(`${userId}_app_url`, webAppUrl);
    props.setProperty(`${userId}_publish_date`, new Date().toISOString());

    // ユーザー設定に公開情報を追加
    updateUserSpreadsheetConfig(userId, {
      appPublished: true,
      appName: publishConfig.appName,
      appUrl: webAppUrl,
      publishedAt: new Date().toISOString(),
    });

    return {
      success: true,
      appUrl: webAppUrl,
      publishedAt: new Date().toISOString(),
    };
  } catch (error) {
    console.error('executeAppPublish エラー:', error);
    return {
      success: false,
      error: error.message,
    };
  }
}

/**
 * WebアプリURLを取得または作成
 * @param {string} userId - ユーザーID
 * @param {string} appName - アプリ名
 * @returns {string} WebアプリURL
 */
function getOrCreateWebAppUrl(userId, appName) {
  try {
    // 既存のWebアプリURLがあるかチェック
    const props = PropertiesService.getScriptProperties();
    const existingUrl = props.getProperty(`${userId}_app_url`);

    if (existingUrl) {
      return existingUrl;
    }

    // 新規WebアプリURLを生成（実際の実装ではScriptAppを使用）
    const scriptId = ScriptApp.getScriptId();
    const webAppUrl = 'https://script.google.com/macros/s/' + scriptId + '/exec?userId=' + userId;

    return webAppUrl;
  } catch (error) {
    console.error('getOrCreateWebAppUrl エラー:', error);
    // フォールバック
    return 'https://script.google.com/macros/s/YOUR_SCRIPT_ID/exec?userId=' + userId;
  }
}

/**
 * 管理パネルのHTMLを返す
 * @returns {HtmlOutput} HTML出力
 */
function doGet() {
  try {
    return HtmlService.createTemplateFromFile('AdminPanel')
      .evaluate()
      .setTitle('StudyQuest 管理パネル')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (error) {
    console.error('doGet エラー:', error);
    return HtmlService.createHtmlOutput(`
      <h1>エラー</h1>
      <p>管理パネルの読み込みに失敗しました: ${error.message}</p>
    `);
  }
}

/**
 * 現在のアクティブボード情報とURL生成
 * @returns {Object} ボード情報とURL
 */
function getCurrentBoardInfoAndUrls() {
  try {
    console.log('getCurrentBoardInfoAndUrls: ボード情報取得開始');

    // 現在ログイン中のユーザー情報を取得
    const userInfo = getCurrentUserInfo();
    if (!userInfo || !userInfo.userId) {
      console.warn('getCurrentBoardInfoAndUrls: ユーザー情報が見つかりません');
      return {
        isActive: false,
        questionText: 'アクティブなボードがありません',
        error: 'ユーザー情報が見つかりません',
      };
    }

    console.log('getCurrentBoardInfoAndUrls: ユーザー情報取得成功', {
      userId: userInfo.userId,
      hasSpreadsheetId: !!userInfo.spreadsheetId,
    });

    // 現在のボードデータを取得
    let boardData = null;
    let questionText = '問題読み込み中...';

    if (userInfo.spreadsheetId) {
      try {
        // ユーザーの設定からアクティブなシート名を取得
        let config = {};
        try {
          config = JSON.parse(userInfo.configJson || '{}');
        } catch (e) {
          console.warn('getCurrentBoardInfoAndUrls: 設定JSON解析エラー:', e.message);
        }

        // アクティブなシート名を決定（優先順位: publishedSheetName > activeSheetName）
        const sheetName = config.publishedSheetName || config.activeSheetName || 'フォームの回答 1';
        console.log('getCurrentBoardInfoAndUrls: 使用するシート名:', sheetName);

        boardData = getSheetData(userInfo.userId, sheetName);
        questionText = boardData?.header || '問題文が設定されていません';
        console.log('getCurrentBoardInfoAndUrls: ボードデータ取得成功', {
          hasHeader: !!boardData?.header,
        });
      } catch (error) {
        console.warn('getCurrentBoardInfoAndUrls: ボードデータ取得エラー:', error.message);
        questionText = '問題文の読み込みに失敗しました';
      }
    } else {
      questionText = 'スプレッドシートが設定されていません';
    }

    // URL生成
    const baseUrl = getWebAppUrlCached();
    if (!baseUrl) {
      console.error('getCurrentBoardInfoAndUrls: WebAppURL取得失敗');
      return {
        isActive: false,
        questionText: questionText,
        error: 'URLの生成に失敗しました',
      };
    }

    const viewUrl = `${baseUrl}?mode=view&userId=${encodeURIComponent(userInfo.userId)}`;
    const adminUrl = `${baseUrl}?mode=admin&userId=${encodeURIComponent(userInfo.userId)}`;

    const result = {
      isActive: !!userInfo.spreadsheetId,
      questionText: questionText,
      sheetName: userInfo.sheetName || 'シート名未設定',
      urls: {
        view: viewUrl, // 閲覧者向け（共有用）
        admin: adminUrl, // 管理者向け
      },
      lastUpdated: new Date().toLocaleString('ja-JP'),
      ownerEmail: userInfo.adminEmail,
    };

    console.log('getCurrentBoardInfoAndUrls: 成功', {
      isActive: result.isActive,
      hasQuestionText: !!result.questionText,
      viewUrl: result.urls.view,
    });

    return result;
  } catch (error) {
    console.error('getCurrentBoardInfoAndUrls エラー:', error);
    return {
      isActive: false,
      questionText: 'エラーが発生しました',
      error: error.message,
    };
  }
}

/**
 * システム管理者かどうかをチェック
 * @returns {boolean} システム管理者の場合はtrue
 */
function checkIsSystemAdmin() {
  try {
    console.log('checkIsSystemAdmin: システム管理者チェック開始');

    // 既存のisDeployUser関数を利用
    const isAdmin = Deploy.isUser();

    console.log('checkIsSystemAdmin: 結果', isAdmin);
    return isAdmin;
  } catch (error) {
    console.error('checkIsSystemAdmin エラー:', error);
    return false;
  }
}

/**
 * AdminPanel用の列マッピングを既存システム互換の形式に変換
 * @param {Object} columnMapping - AdminPanel用の列マッピング
 * @param {Array<string>} headerRow - ヘッダー行の配列
 * @returns {Object} Core.gs互換の形式
 */
function convertToCompatibleMapping(columnMapping, headerRow) {
  try {
    const compatibleMapping = {};

    // AdminPanel形式から既存システムのCOLUMN_HEADERS形式に変換
    const mappingConversions = {
      question: 'OPINION', // 質問 → 回答（既存システムでは意見/回答として扱う）
      answerer: 'NAME', // 回答者 → 名前
      reason: 'REASON', // 理由 → 理由
      timestamp: 'TIMESTAMP', // タイムスタンプ → タイムスタンプ
      class: 'CLASS', // クラス → クラス
    };

    // 既存のCOLUMN_HEADERSと対応する実際の列名
    const columnHeaders = {
      TIMESTAMP: 'タイムスタンプ',
      EMAIL: 'メールアドレス',
      CLASS: 'クラス',
      OPINION: '回答',
      REASON: '理由',
      NAME: '名前',
      UNDERSTAND: 'なるほど！',
      LIKE: 'いいね！',
      CURIOUS: 'もっと知りたい！',
      HIGHLIGHT: 'ハイライト',
    };

    // AdminPanel形式を既存システム形式に変換
    Object.keys(columnMapping).forEach((key) => {
      if (key === 'confidence') return; // 信頼度は変換対象外

      const columnIndex = columnMapping[key];
      if (columnIndex !== null && columnIndex !== undefined) {
        const systemKey = mappingConversions[key];
        if (systemKey) {
          compatibleMapping[columnHeaders[systemKey]] = columnIndex;
        }
      }
    });

    // 必須列の自動検出を試行
    headerRow.forEach((header, index) => {
      const headerLower = header.toString().toLowerCase();

      // メールアドレス列の検出
      if (headerLower.includes('mail') || headerLower.includes('メール')) {
        compatibleMapping[columnHeaders.EMAIL] = index;
      }

      // タイムスタンプ列の検出
      if (
        headerLower.includes('timestamp') ||
        headerLower.includes('タイムスタンプ') ||
        headerLower.includes('時刻')
      ) {
        compatibleMapping[columnHeaders.TIMESTAMP] = index;
      }
    });

    console.log('convertToCompatibleMapping: 変換結果', {
      original: columnMapping,
      compatible: compatibleMapping,
    });

    return compatibleMapping;
  } catch (error) {
    console.error('convertToCompatibleMapping エラー:', error);
    return {};
  }
}

console.log('AdminPanel.gs 読み込み完了');
