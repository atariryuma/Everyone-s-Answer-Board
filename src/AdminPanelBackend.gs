/**
 * @fileoverview AdminPanel バックエンド関数
 * 既存のシステムと統合する最小限の実装
 */

/**
 * 管理パネル用のサーバーサイド関数群
 * 既存のunifiedUserManager、database、Core関数を活用
 */

// SYSTEM_CONSTANTS の存在確認とデバッグ
/**
 * 汎用ユーザー情報取得関数
 * Services.user.getCurrentUserInfo()を活用した統一インターフェース - 全システムでgetActiveUserInfo()を使用
 */
function getActiveUserInfo() {
  return Services.user.getActiveUserInfo();
}

function debugConstants() {
  console.log('SYSTEM_CONSTANTS:', typeof SYSTEM_CONSTANTS);
  if (typeof SYSTEM_CONSTANTS !== 'undefined') {
    console.log('COLUMN_MAPPING:', typeof SYSTEM_CONSTANTS.COLUMN_MAPPING);
    console.log('COLUMN_MAPPING keys:', SYSTEM_CONSTANTS.COLUMN_MAPPING ? Object.keys(SYSTEM_CONSTANTS.COLUMN_MAPPING) : 'undefined');
  }
}

// =============================================================================
// データソース管理
// =============================================================================

/**
 * スプレッドシート一覧を取得（オーナー権限＋フォーム連携チェック）
 * @returns {Array<Object>} スプレッドシート情報の配列
 */
function getSpreadsheetList() {
  try {
    console.log('getSpreadsheetList: スプレッドシート一覧取得開始');
    
    const currentUserEmail = Session.getActiveUser().getEmail();
    console.log('現在のユーザー:', currentUserEmail);

    // Google Driveからスプレッドシートを検索
    const files = DriveApp.searchFiles(
      'mimeType="application/vnd.google-apps.spreadsheet" and trashed=false'
    );

    const spreadsheets = [];
    let count = 0;
    const maxResults = 50; // パフォーマンス制限

    while (files.hasNext() && count < maxResults) {
      const file = files.next();
      
      // オーナー権限チェック
      let ownerEmail = 'Unknown';
      let isOwner = false;
      try {
        const owner = file.getOwner();
        if (owner) {
          ownerEmail = owner.getEmail();
          isOwner = ownerEmail === currentUserEmail;
        }
      } catch (ownerError) {
        console.warn(`Owner取得エラー for file ${file.getName()}:`, ownerError.message);
        continue; // オーナー確認できない場合はスキップ
      }

      // オーナーでない場合はスキップ
      if (!isOwner) {
        continue;
      }

      // フォーム連携チェック
      let formInfo = null;
      try {
        formInfo = checkFormConnection(file.getId());
      } catch (formError) {
        console.warn(`フォーム連携チェックエラー for ${file.getName()}:`, formError.message);
      }

      // フォーム連携がない場合はスキップ
      if (!formInfo || !formInfo.hasForm) {
        continue;
      }
      
      spreadsheets.push({
        id: file.getId(),
        name: file.getName(),
        lastModified: file.getLastUpdated().toISOString(),
        owner: ownerEmail,
        isOwner: true,
        formUrl: formInfo.formUrl,
        formTitle: formInfo.formTitle,
        hasFormConnection: true
      });
      count++;
    }

    // 最終更新順でソート
    spreadsheets.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));

    console.log(`getSpreadsheetList: ${spreadsheets.length}個の条件適合スプレッドシートを取得`);
    return spreadsheets;
  } catch (error) {
    console.error('getSpreadsheetList エラー:', error);
    throw new Error('スプレッドシート一覧の取得に失敗しました: ' + error.message);
  }
}

/**
 * スプレッドシートのフォーム連携をチェック
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} フォーム連携情報
 */
function checkFormConnection(spreadsheetId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // フォーム連携されたスプレッドシートには特定のプロパティがある
    const form = FormApp.openByUrl(spreadsheet.getFormUrl());
    
    if (form) {
      const formUrl = form.getPublishedUrl();
      const formTitle = form.getTitle();
      
      return {
        hasForm: true,
        formUrl: formUrl,
        formTitle: formTitle,
        formId: form.getId()
      };
    }
    
    return { hasForm: false };
    
  } catch (error) {
    // フォーム連携がない場合、getFormUrl()でエラーが発生する
    console.log(`フォーム連携チェック: ${spreadsheetId} はフォーム非連携`);
    return { hasForm: false };
  }
}

/**
 * シートがフォーム回答用シートかチェック
 * @param {Sheet} sheet - チェック対象のシート
 * @param {Object} formInfo - フォーム連携情報
 * @returns {boolean} フォーム回答用シートかどうか
 */
function checkIfFormResponseSheet(sheet, formInfo) {
  try {
    const sheetName = sheet.getName();
    
    // フォーム連携がない場合は false
    if (!formInfo || !formInfo.hasForm) {
      return false;
    }
    
    // フォーム回答シートの特徴をチェック
    // 1. シート名が「フォームの回答」で始まる
    if (sheetName.startsWith('フォームの回答') || sheetName.startsWith('Form Responses')) {
      return true;
    }
    
    // 2. ヘッダー行の特徴をチェック（タイムスタンプ列の存在）
    if (sheet.getLastRow() > 0 && sheet.getLastColumn() > 0) {
      const headerRow = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
      
      // タイムスタンプ列があるかチェック
      const hasTimestamp = headerRow.some(header => {
        if (!header) return false;
        const headerStr = String(header).toLowerCase();
        return headerStr.includes('timestamp') || 
               headerStr.includes('タイムスタンプ') || 
               headerStr.includes('回答時刻');
      });
      
      if (hasTimestamp) {
        return true;
      }
    }
    
    return false;
  } catch (error) {
    console.warn(`フォーム回答シートチェックエラー for ${sheet.getName()}:`, error.message);
    return false;
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

    // フォーム連携情報を取得
    let formInfo = null;
    try {
      formInfo = checkFormConnection(spreadsheetId);
    } catch (formError) {
      console.warn('フォーム連携情報取得エラー:', formError.message);
    }

    const sheetList = sheets.map((sheet) => {
      const sheetName = sheet.getName();
      
      // シートがフォーム回答用シートかチェック
      const isFormResponseSheet = checkIfFormResponseSheet(sheet, formInfo);
      
      return {
        name: sheetName,
        index: sheet.getIndex(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn(),
        hidden: sheet.isSheetHidden(),
        isFormResponseSheet: isFormResponseSheet,
        formConnected: formInfo && formInfo.hasForm ? true : false,
        formTitle: formInfo && formInfo.hasForm ? formInfo.formTitle : null
      };
    });

    console.log(`getSheetList: ${sheetList.length}個のシートを取得（フォーム連携情報付き）`);
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

    // 超高精度AI列マッピングを活用（新システム）
    // 【プロダクションバグ修正】 空オブジェクト {} も truthy と評価される問題を解決
    // 問題: headerIndices が {} の場合、truthy判定でconvertIndicesToMappingが呼ばれ
    // "Cannot convert undefined or null to object" エラーが発生
    const hasValidHeaderIndices = headerIndices && 
                                  typeof headerIndices === 'object' && 
                                  Object.keys(headerIndices).length > 0;
    
    
    let columnMapping = detectColumnMapping(headerRow);

    // 列名マッピングの整合性チェック
    const validationResult = validateAdminPanelMapping(columnMapping);
    if (!validationResult.isValid) {
      console.warn('列名マッピング検証エラー', validationResult.errors);
    }
    if (validationResult.warnings.length > 0) {
      console.warn('列名マッピング警告', validationResult.warnings);
    }

    // 不足列の検出・追加
    const missingColumnsResult = addMissingColumns(spreadsheetId, sheetName, columnMapping);
    console.log('connectToDataSource: 不足列検出結果', missingColumnsResult);

    // 列が追加された場合は、ヘッダー行を再取得して列マッピングを更新
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      const updatedHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // 更新後のヘッダーで再実行（キャッシュを削除して最新取得）
      cacheManager.remove(`hdr_${spreadsheetId}_${sheetName}`);
      const updatedHeaderIndices = getHeadersCached(spreadsheetId, sheetName);
      columnMapping = detectColumnMapping(updatedHeaderRow);
    }

    // フォーム連携情報を取得してデータベースに保存
    let formInfo = null;
    try {
      formInfo = checkFormConnection(spreadsheetId);
      console.log('フォーム連携情報:', formInfo);
    } catch (formError) {
      console.warn('フォーム情報取得エラー:', formError.message);
    }

    // 設定を保存（既存のユーザー管理システムを活用）
    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);

    if (userInfo) {
      // 既存システム互換の列マッピングを作成
      const compatibleMapping = convertToCompatibleMapping(columnMapping, headerRow);

      // ユーザーのスプレッドシート設定を更新（フォームURL含む）
      const updateData = {
        spreadsheetId: spreadsheetId,
        sheetName: sheetName,
        columnMapping: columnMapping, // AdminPanel用の形式
        compatibleMapping: compatibleMapping, // Core.gs互換形式
        lastConnected: new Date().toISOString(),
        connectionMethod: 'dropdown_select',
        missingColumnsHandled: missingColumnsResult,
      };

      // フォーム情報がある場合、データベースに保存
      if (formInfo && formInfo.hasForm) {
        // データベースのformUrl列に保存
        const dbUpdateResult = updateUser(userInfo.userId, {
          formUrl: formInfo.formUrl,
          lastAccessedAt: new Date().toISOString()
        });
        
        if (dbUpdateResult) {
          console.log('フォームURL保存成功:', formInfo.formUrl);
          updateData.formTitle = formInfo.formTitle;
        } else {
          console.warn('フォームURL保存失敗');
        }
      }

      updateUserSpreadsheetConfig(userInfo.userId, updateData);

      console.log('connectToDataSource: 互換形式も保存', { columnMapping, compatibleMapping });
    }

    console.log('connectToDataSource: 接続成功', columnMapping);

    // メッセージを統合
    let message = 'データソースに正常に接続されました';
    if (missingColumnsResult.success) {
      if (missingColumnsResult.addedColumns.length > 0) {
        message += `。${missingColumnsResult.addedColumns.length}個の必須列を自動追加しました`;
      }
      if (
        missingColumnsResult.recommendedColumns &&
        missingColumnsResult.recommendedColumns.length > 0
      ) {
        message += `。${missingColumnsResult.recommendedColumns.length}個の推奨列を手動で追加することをお勧めします`;
      }
    }

    return {
      success: true,
      columnMapping: columnMapping,
      headers: headerRow, // 🔥 追加: 実際のヘッダー情報
      rowCount: sheet.getLastRow(),
      message: message,
      missingColumnsResult: missingColumnsResult,
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
 * ヘッダー行から列マッピングを自動検出（SYSTEM_CONSTANTS.COLUMN_MAPPING使用）
 * @param {Array<string>} headers - ヘッダー行の配列
 * @returns {Object} 検出された列マッピング
 */
/**
 * 超高精度列マッピング検出システム（AI統合版）
 * mapColumns から detectColumnMapping にリネーム
 */
function detectColumnMapping(headers) {
  // デバッグ: SYSTEM_CONSTANTSの存在確認
  debugConstants();
  
  // 1. SYSTEM_CONSTANTS チェック → 失敗時は基本AI活用
  if (typeof SYSTEM_CONSTANTS === 'undefined' || !SYSTEM_CONSTANTS.COLUMN_MAPPING) {
    console.log('SYSTEM_CONSTANTS.COLUMN_MAPPING not available, using basic AI detection');
    
    // 2. 高性能AI判定システム（フォールバック版）- aiPatterns活用
    const mapping = {};
    const confidence = {};
    
    // SYSTEM_CONSTANTSが利用不可時の静的パターン定義
    const fallbackPatterns = {
      answer: {
        aiPatterns: ['？', '?', 'どうして', 'なぜ', '思いますか', '考えますか'],
        alternates: ['どうして', '質問', '問題', '意見', '答え', 'なぜ', '思います', '考え'],
        minLength: 15
      },
      reason: {
        aiPatterns: ['理由', '体験', '根拠', '詳細'],
        alternates: ['理由', '根拠', '体験', 'なぜ', '詳細', '説明']
      },
      class: {
        alternates: ['クラス', '学年']
      },
      name: {
        alternates: ['名前', '氏名', 'お名前']
      }
    };
    
    // 高精度パターンマッチング
    headers.forEach((header, index) => {
      const headerLower = header.toLowerCase();
      
      Object.keys(fallbackPatterns).forEach(key => {
        const pattern = fallbackPatterns[key];
        let matchScore = 0;
        
        // aiPatterns による高精度検出
        if (pattern.aiPatterns) {
          pattern.aiPatterns.forEach(aiPattern => {
            if (headerLower.includes(aiPattern.toLowerCase())) {
              matchScore = Math.max(matchScore, 90);
            }
          });
          
          // 質問文の特別処理（長い文章 + 疑問符）
          if (key === 'answer' && pattern.minLength && header.length > pattern.minLength) {
            const hasQuestionMark = pattern.aiPatterns.some(p => header.includes(p));
            if (hasQuestionMark) {
              matchScore = Math.max(matchScore, 95); // 最高精度
            }
          }
        }
        
        // alternates による補完検出
        if (matchScore === 0 && pattern.alternates) {
          pattern.alternates.forEach(alternate => {
            if (headerLower.includes(alternate.toLowerCase())) {
              matchScore = Math.max(matchScore, 80);
            }
          });
        }
        
        // より高いスコアで置き換え
        if (matchScore > 0) {
          if (!mapping[key] || matchScore > (confidence[key] || 0)) {
            mapping[key] = index;
            confidence[key] = matchScore;
          }
        }
      });
    });
    
    // デフォルト値の設定
    ['answer', 'reason', 'class', 'name'].forEach(key => {
      if (mapping[key] === undefined) mapping[key] = null;
    });
    
    mapping.confidence = confidence;
    
    console.log('detectColumnMapping: Enhanced AI detection result', {
      mapping, 
      confidence,
      usedPatterns: 'aiPatterns + alternates + minLength'
    });
    return mapping;
  }

  // SYSTEM_CONSTANTS.COLUMN_MAPPINGベースの初期化
  const mapping = {};
  const confidence = {};

  // SYSTEM_CONSTANTS.COLUMN_MAPPINGの各列定義を初期化
  Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
    mapping[column.key] = null;
  });
  mapping.confidence = {};

  // ヘッダー検出処理
  headers.forEach((header, index) => {
    const headerLower = header.toString().toLowerCase();

    // SYSTEM_CONSTANTS.COLUMN_MAPPINGの各列を検査
    Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
      const headerName = column.header.toLowerCase();
      const fieldKey = column.key;

      // 基本マッチング（完全一致優先）
      let matchScore = 0;
      if (headerLower === headerName) {
        matchScore = 95; // 完全一致
      } else if (headerLower.includes(headerName)) {
        matchScore = 80; // 部分一致
      } else if (headerName.includes(headerLower) && headerLower.length > 2) {
        matchScore = 70; // 逆部分一致
      }

      // alternatesを使った拡張マッチング
      if (matchScore === 0 && column.alternates) {
        column.alternates.forEach((alternate) => {
          const alternateLower = alternate.toLowerCase();
          if (headerLower.includes(alternateLower)) {
            matchScore = Math.max(matchScore, 75); // alternates マッチング
          }
        });
      }

      // aiPatternsを使った高性能AI検出
      if (matchScore === 0 && column.aiPatterns) {
        column.aiPatterns.forEach((aiPattern) => {
          const aiPatternLower = aiPattern.toLowerCase();
          if (headerLower.includes(aiPatternLower)) {
            matchScore = Math.max(matchScore, 85); // aiPatterns 高精度マッチング
          }
        });
        
        // 回答列の特別処理：長い質問文 + aiパターンの組み合わせ
        if (fieldKey === 'answer' && header.length > 15) {
          const hasAIPattern = column.aiPatterns.some(p => 
            header.includes(p) || headerLower.includes(p.toLowerCase())
          );
          if (hasAIPattern) {
            matchScore = Math.max(matchScore, 92); // 質問文特別検出
          }
        }
      }

      // より高い信頼度で置き換え
      if (matchScore > 0) {
        if (!mapping[fieldKey] || matchScore > (confidence[fieldKey] || 0)) {
          mapping[fieldKey] = index;
          confidence[fieldKey] = matchScore;
        }
      }
    });
  });

  // confidenceを返り値に追加
  mapping.confidence = confidence;

  // 4. SYSTEM_CONSTANTS処理 + AI補強
  const basicMapping = performBasicSYSTEM_CONSTANTSMapping(headers);
  const aiEnhancement = identifyHeadersAdvanced(headers);
  
  // 5. AI結果で精度向上
  const enhancedMapping = mergeColumnConfidence(basicMapping, aiEnhancement, headers);
  
  console.log('detectColumnMapping: AI統合・超高精度マッピング完了', {
    headers,
    basicMapping,
    enhancedMapping,
    basicConfidence: basicMapping.confidence,
    enhancedConfidence: enhancedMapping.confidence,
    usedTechnology: 'SYSTEM_CONSTANTS + aiPatterns + Advanced AI + Internet Knowledge'
  });

  return enhancedMapping;
}

/**
 * 基本的なSYSTEM_CONSTANTSマッピング（既存処理を分離）
 */
function performBasicSYSTEM_CONSTANTSMapping(headers) {
  const mapping = {};
  const confidence = {};

  // SYSTEM_CONSTANTS.COLUMN_MAPPINGの各列定義を初期化
  Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
    mapping[column.key] = null;
  });
  mapping.confidence = {};

  // ヘッダー検出処理（既存ロジック）
  headers.forEach((header, index) => {
    const headerLower = header.toString().toLowerCase();

    // SYSTEM_CONSTANTS.COLUMN_MAPPINGの各列を検査
    Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
      const headerName = column.header.toLowerCase();
      const fieldKey = column.key;

      // 基本マッチング（完全一致優先）
      let matchScore = 0;
      if (headerLower === headerName) {
        matchScore = 95; // 完全一致
      } else if (headerLower.includes(headerName)) {
        matchScore = 80; // 部分一致
      } else if (headerName.includes(headerLower) && headerLower.length > 2) {
        matchScore = 70; // 逆部分一致
      }

      // alternatesを使った拡張マッチング
      if (matchScore === 0 && column.alternates) {
        column.alternates.forEach((alternate) => {
          const alternateLower = alternate.toLowerCase();
          if (headerLower.includes(alternateLower)) {
            matchScore = Math.max(matchScore, 75); // alternates マッチング
          }
        });
      }

      // aiPatternsを使った高性能AI検出
      if (matchScore === 0 && column.aiPatterns) {
        column.aiPatterns.forEach((aiPattern) => {
          const aiPatternLower = aiPattern.toLowerCase();
          if (headerLower.includes(aiPatternLower)) {
            matchScore = Math.max(matchScore, 85); // aiPatterns 高精度マッチング
          }
        });
        
        // 回答列の特別処理：長い質問文 + aiパターンの組み合わせ
        if (fieldKey === 'answer' && header.length > 15) {
          const hasAIPattern = column.aiPatterns.some(p => 
            header.includes(p) || headerLower.includes(p.toLowerCase())
          );
          if (hasAIPattern) {
            matchScore = Math.max(matchScore, 92); // 質問文特別検出
          }
        }
      }

      // より高い信頼度で置き換え
      if (matchScore > 0) {
        if (!mapping[fieldKey] || matchScore > (mapping.confidence[fieldKey] || 0)) {
          mapping[fieldKey] = index;
          mapping.confidence[fieldKey] = matchScore;
        }
      }
    });
  });

  return mapping;
}

/**
 * AIによるマッピング強化
 */
function mergeColumnConfidence(basicMapping, aiResult, headers) {
  const enhanced = { ...basicMapping };
  
  // 既存のconfidence値を保持（重要：0%問題の修正）
  enhanced.confidence = { ...basicMapping.confidence };
  
  
  // AI結果で既存マッピングを強化（既存confidence値を保持）
  if (aiResult.answer && (!enhanced.answer || (aiResult.confidence?.answer || 0) > (enhanced.confidence?.answer || 0))) {
    enhanced.answer = headers.indexOf(aiResult.answer);
    if (aiResult.confidence?.answer) {
      enhanced.confidence.answer = aiResult.confidence.answer;
    }
    // aiResult.confidence?.answerが無い場合はbasicMapping.confidenceの値を保持
  }
  
  if (aiResult.reason && (!enhanced.reason || (aiResult.confidence?.reason || 0) > (enhanced.confidence?.reason || 0))) {
    enhanced.reason = headers.indexOf(aiResult.reason);
    if (aiResult.confidence?.reason) {
      enhanced.confidence.reason = aiResult.confidence.reason;
    }
  }
  
  if (aiResult.classHeader && (!enhanced.class || (aiResult.confidence?.class || 0) > (enhanced.confidence?.class || 0))) {
    enhanced.class = headers.indexOf(aiResult.classHeader);
    if (aiResult.confidence?.class) {
      enhanced.confidence.class = aiResult.confidence.class;
    }
  }
  
  if (aiResult.name && (!enhanced.name || (aiResult.confidence?.name || 0) > (enhanced.confidence?.name || 0))) {
    enhanced.name = headers.indexOf(aiResult.name);
    if (aiResult.confidence?.name) {
      enhanced.confidence.name = aiResult.confidence.name;
    }
  }
  
  
  return enhanced;
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

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet) {
      throw new Error(`シート "${sheetName}" が見つかりません`);
    }

    // 現在のヘッダー行を取得
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 必要な列を定義（StudyQuestシステムで使用される標準列）
    const requiredColumns = {
      メールアドレス: 'EMAIL',
      理由: 'REASON',
      名前: 'NAME',
      'なるほど！': 'UNDERSTAND',
      'いいね！': 'LIKE',
      'もっと知りたい！': 'CURIOUS',
      ハイライト: 'HIGHLIGHT',
      タイムスタンプ: 'TIMESTAMP',
    };

    // 不足している列を検出
    const missingColumns = [];
    const existingColumns = headerRow.map((h) => String(h || '').trim());

    Object.keys(requiredColumns).forEach((requiredCol) => {
      const found = existingColumns.some((existing) => {
        const existingLower = existing.toLowerCase();
        const requiredLower = requiredCol.toLowerCase();

        // 基本的な検出ロジック（リアクション・ハイライトは完全一致）
        if (
          requiredCol === 'なるほど！' ||
          requiredCol === 'いいね！' ||
          requiredCol === 'もっと知りたい！' ||
          requiredCol === 'ハイライト'
        ) {
          // システム列は完全一致のみ
          return existing === requiredCol;
        }

        // その他の列は柔軟な検出
        return (
          existingLower.includes(requiredLower) ||
          (requiredCol === 'メールアドレス' && existingLower.includes('email')) ||
          (requiredCol === '理由' &&
            (existingLower.includes('理由') || existingLower.includes('詳細'))) ||
          (requiredCol === '名前' &&
            (existingLower.includes('名前') || existingLower.includes('氏名'))) ||
          (requiredCol === 'タイムスタンプ' && existingLower.includes('timestamp'))
        );
      });

      if (!found) {
        missingColumns.push({
          columnName: requiredCol,
          systemName: requiredColumns[requiredCol],
          priority: getPriority(requiredCol),
        });
      }
    });


    // 不足列がない場合
    if (missingColumns.length === 0) {
      return {
        success: true,
        missingColumns: [],
        addedColumns: [],
        message: '必要な列がすべて揃っています',
      };
    }

    // 優先度順にソート
    missingColumns.sort((a, b) => a.priority - b.priority);

    // 列を自動追加（高優先度のみ）
    const addedColumns = [];
    const highPriorityColumns = missingColumns.filter((col) => col.priority <= 2);

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
          systemName: colInfo.systemName,
        });

      });
    }

    // 残りの不足列（低優先度）は推奨として返す
    const recommendedColumns = missingColumns.filter((col) => col.priority > 2);

    return {
      success: true,
      missingColumns: missingColumns,
      addedColumns: addedColumns,
      recommendedColumns: recommendedColumns,
      message: `${addedColumns.length}個の必須列を自動追加しました`,
      details: {
        added: addedColumns.length,
        recommended: recommendedColumns.length,
        total: missingColumns.length,
      },
    };
  } catch (error) {
    console.error('detectAndAddMissingColumns エラー:', error);
    return {
      success: false,
      error: error.message,
      missingColumns: [],
      addedColumns: [],
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
    メールアドレス: 1, // 最高優先度
    理由: 1, // 最高優先度
    名前: 2, // 高優先度
    タイムスタンプ: 3, // 中優先度
    'なるほど！': 4, // 低優先度
    'いいね！': 4,
    'もっと知りたい！': 4,
    ハイライト: 4,
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
      sheets: sheetList.length,
    });

    return {
      success: true,
      spreadsheetId: spreadsheetId,
      spreadsheetName: spreadsheetName,
      sheets: sheetList,
      message: 'スプレッドシートへのアクセスが確認できました',
    };
  } catch (error) {
    console.error('validateSpreadsheetAccess エラー:', error);
    return {
      success: false,
      error: error.message,
      details: 'スプレッドシートが存在しないか、アクセス権限がありません',
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

    // 超高精度AI列マッピングを活用（新システム）
    // 【プロダクションバグ修正】 空オブジェクト {} も truthy と評価される問題を解決
    const hasValidHeaderIndices = headerIndices && 
                                  typeof headerIndices === 'object' && 
                                  Object.keys(headerIndices).length > 0;
    
    let columnMapping = detectColumnMapping(headerRow);

    // 列名マッピングの整合性チェック
    const validationResult = validateAdminPanelMapping(columnMapping);
    if (!validationResult.isValid) {
      console.warn('列名マッピング検証エラー', validationResult.errors);
    }
    if (validationResult.warnings.length > 0) {
      console.warn('列名マッピング警告', validationResult.warnings);
    }

    // 不足列の検出・追加
    const missingColumnsResult = addMissingColumns(spreadsheetId, sheetName, columnMapping);

    // 列が追加された場合は、ヘッダー行を再取得して列マッピングを更新
    if (missingColumnsResult.success && missingColumnsResult.addedColumns.length > 0) {
      const updatedHeaderRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // 更新後のヘッダーで再実行（キャッシュを削除して最新取得）
      cacheManager.remove(`hdr_${spreadsheetId}_${sheetName}`);
      const updatedHeaderIndices = getHeadersCached(spreadsheetId, sheetName);
      columnMapping = detectColumnMapping(updatedHeaderRow);

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
        missingColumnsHandled: missingColumnsResult,
      });
    }

    console.log('analyzeSpreadsheetColumns: 分析完了', columnMapping);

    // メッセージを統合
    let message = 'スプレッドシートの列構造を分析しました';
    if (missingColumnsResult.success) {
      if (missingColumnsResult.addedColumns.length > 0) {
        message += `。${missingColumnsResult.addedColumns.length}個の必須列を自動追加しました`;
      }
      if (
        missingColumnsResult.recommendedColumns &&
        missingColumnsResult.recommendedColumns.length > 0
      ) {
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
      missingColumnsResult: missingColumnsResult,
    };
  } catch (error) {
    console.error('analyzeSpreadsheetColumns エラー:', error);
    return {
      success: false,
      error: error.message,
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
  // nullチェックを追加
  if (!headerIndices || typeof headerIndices !== 'object') {
    console.error('convertIndicesToMapping: headerIndices is null or invalid', headerIndices);
    // エラーを投げる代わりにAI判定にフォールバック
    console.log('convertIndicesToMapping: Falling back to AI detection due to invalid headerIndices');
    return detectColumnMapping(headerRow);
  }
  
  if (!headerRow || !Array.isArray(headerRow)) {
    console.error('convertIndicesToMapping: headerRow is null or invalid', headerRow);
    throw new Error('Cannot convert undefined or null headerRow to mapping object');
  }

  // 空のheaderIndicesの場合もAI判定にフォールバック
  if (Object.keys(headerIndices).length === 0) {
    console.log('convertIndicesToMapping: Empty headerIndices, falling back to AI detection');
    return detectColumnMapping(headerRow);
  }

  // シンプル・単一定数SYSTEM_CONSTANTS.COLUMN_MAPPINGを使用
  const mapping = {};

  console.log('convertIndicesToMapping: 入力データ確認', {
    headerIndices,
    headerRow: headerRow.slice(0, 10), // 最初の10項目のみログ出力
    headerRowLength: headerRow.length
  });

  // SYSTEM_CONSTANTS の安全性チェック
  if (!SYSTEM_CONSTANTS || !SYSTEM_CONSTANTS.COLUMN_MAPPING) {
    console.warn('convertIndicesToMapping: SYSTEM_CONSTANTS.COLUMN_MAPPING is not available, falling back to AI detection');
    return detectColumnMapping(headerRow);
  }

  // 各列定義を直接使用（変換層なし）
  Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
    const headerName = column.header; // '回答', '理由' など
    const uiFieldName = column.key; // 'answer', 'reason' など

    // 直接マッチをチェック
    let columnIndex = null;
    if (headerIndices[headerName] !== undefined) {
      columnIndex = headerIndices[headerName];
    } else {
      // alternatesでの部分マッチングを試行
      if (column.alternates && Array.isArray(column.alternates)) {
        for (const alternate of column.alternates) {
          // headerIndicesの各キーに対してalternatesをチェック
          for (const [actualHeader, index] of Object.entries(headerIndices)) {
            if (actualHeader.toLowerCase().includes(alternate.toLowerCase())) {
              columnIndex = index;
              console.log(`convertIndicesToMapping: alternateマッチ ${uiFieldName}: "${actualHeader}" -> ${alternate} (index: ${index})`);
              break;
            }
          }
          if (columnIndex !== null) break;
        }
      }
      
      // 質問文がヘッダーになっている場合の特別処理（answer列）
      if (columnIndex === null && uiFieldName === 'answer') {
        for (const [actualHeader, index] of Object.entries(headerIndices)) {
          // 15文字以上で質問っぽいヘッダーを回答列として認識
          if (actualHeader.length > 15 && 
              (actualHeader.includes('？') || actualHeader.includes('?') || 
               actualHeader.includes('どうして') || actualHeader.includes('なぜ') || 
               actualHeader.includes('思います') || actualHeader.includes('考え'))) {
            columnIndex = index;
            console.log(`convertIndicesToMapping: 質問文ヘッダーを回答列として認識 ${uiFieldName}: "${actualHeader.substring(0, 30)}..." (index: ${index})`);
            break;
          }
        }
      }
      
      // 理由っぽいヘッダーの特別処理（reason列）
      if (columnIndex === null && uiFieldName === 'reason') {
        for (const [actualHeader, index] of Object.entries(headerIndices)) {
          if (actualHeader.includes('理由') || actualHeader.includes('体験') || 
              actualHeader.includes('根拠') || actualHeader.includes('なぜ')) {
            columnIndex = index;
            console.log(`convertIndicesToMapping: 理由系ヘッダーを理由列として認識 ${uiFieldName}: "${actualHeader}" (index: ${index})`);
            break;
          }
        }
      }
    }

    mapping[uiFieldName] = columnIndex;
  });

  return mapping;
}

/**
 * AdminPanel列名マッピングの整合性チェック関数
 * @param {Object} mapping - 変換されたマッピング
 * @returns {Object} チェック結果
 */
function validateAdminPanelMapping(mapping) {
  const results = {
    isValid: true,
    errors: [],
    warnings: [],
    summary: {},
  };

  // SYSTEM_CONSTANTS.COLUMN_MAPPINGに基づく動的チェック
  Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
    const fieldKey = column.key;
    const isRequired = column.required;

    if (isRequired && !mapping[fieldKey] && mapping[fieldKey] !== 0) {
      results.isValid = false;
      results.errors.push(`必須フィールド '${fieldKey}' (${column.header}) が設定されていません`);
    } else if (!isRequired && !mapping[fieldKey] && mapping[fieldKey] !== 0) {
      results.warnings.push(`推奨フィールド '${fieldKey}' (${column.header}) が設定されていません`);
    }
  });

  // 許可されたフィールドかチェック
  const allowedFields = Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).map((col) => col.key);
  Object.keys(mapping).forEach((uiField) => {
    if (!allowedFields.includes(uiField)) {
      results.warnings.push(`未知のUIフィールド '${uiField}' が含まれています`);
    }
  });

  results.summary = {
    totalFields: Object.keys(mapping).length,
    validFields: Object.keys(mapping).filter((k) => mapping[k] !== null).length,
    nullFields: Object.keys(mapping).filter((k) => mapping[k] === null).length,
  };

  console.log('validateAdminPanelMapping:', results);
  return results;
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
      sheetName: sheetName, // 🔥 強化されたシート名情報
      formCreated: configJson ? configJson.formCreated : false,
      appPublished: configJson ? configJson.appPublished : false,
      lastUpdated: userInfo.lastUpdated,
      setupStep: setupStep,
      setupComplete: setupStep >= 3, // 🔥 追加: セットアップ完了状態
      user: currentUser,
      userId: userInfo.userId,
      displaySettings: {
        // 🔥 追加: 表示設定
        showNames: configJson ? configJson.showNames !== false : true,
        showReactions: configJson ? configJson.showReactions !== false : true,
      },
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
      const sheet = sheets.find((s) => s.getName() === priorityName);
      if (sheet && sheet.getLastRow() > 1) {
        // データがあるかチェック
        console.log('detectActiveSheetName: 優先シート検出', priorityName);
        return priorityName;
      }
    }

    // データが最も多いシートを選択
    const sheetsWithData = sheets
      .filter((s) => s.getLastRow() > 1)
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
    console.log('updateUserSpreadsheetConfig: 設定更新開始', { userId, config });

    // 新しいconfigManagerを使用して設定を更新
    const updateResult = App.getConfig().updateUserConfig(userId, config);
    
    if (updateResult) {
      const updatedConfig = App.getConfig().getUserConfig(userId);
      console.log('updateUserSpreadsheetConfig: 設定更新完了', { userId, config: updatedConfig });
      return {
        success: true,
        config: updatedConfig,
      };
    } else {
      throw new Error('設定の更新に失敗しました');
    }
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
    console.log('getUserConfigJson: 設定取得開始', userId);
    
    // 新しいconfigManagerを使用して設定を取得
    const config = App.getConfig().getUserConfig(userId);
    
    console.log('getUserConfigJson: 設定取得完了', { userId, hasConfig: !!config });
    return config;
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
 * Page.html用：ユーザーの列マッピング設定を取得
 * @param {string} userId - ユーザーID（オプション）
 * @returns {Object} 列マッピング設定
 */
function getUserColumnMapping(userId = null) {
  try {
    console.log('getUserColumnMapping: 列マッピング取得開始', userId);

    // ユーザー情報の取得
    const targetUserId = userId || getActiveUserInfo()?.userId;
    if (!targetUserId) {
      console.warn('getUserColumnMapping: ユーザーIDが見つかりません');
      return {};
    }

    // 保存されている設定を取得
    const userConfig = getUserConfigJson(targetUserId);
    if (userConfig && userConfig.columnMapping) {
      console.log('getUserColumnMapping: 保存済み設定を使用', userConfig.columnMapping);
      return userConfig.columnMapping;
    }

    // 設定が見つからない場合はデフォルト（空）を返す
    console.log('getUserColumnMapping: デフォルト設定を使用');
    const defaultMapping = {};
    Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
      defaultMapping[column.key] = null;
    });

    return defaultMapping;
  } catch (error) {
    console.error('getUserColumnMapping エラー:', error);
    return {};
  }
}

// doGet: main.gsのメインエントリーポイントに統合済み（重複削除）

/**
 * 現在のアクティブボード情報とURL生成
 * @returns {Object} ボード情報とURL
 */
function getCurrentBoardInfoAndUrls() {
  try {
    console.log('getCurrentBoardInfoAndUrls: ボード情報取得開始');

    // 現在ログイン中のユーザー情報を取得
    const userInfo = getActiveUserInfo();
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
    let config = {};

    // ユーザーの設定JSON解析
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      console.warn('getCurrentBoardInfoAndUrls: 設定JSON解析エラー:', e.message);
    }

    if (userInfo.spreadsheetId) {
      try {

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
    const baseUrl = getWebAppUrl();
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
      isPublished: config?.appPublished || false,
      questionText: questionText,
      sheetName: userInfo.sheetName || 'シート名未設定',
      urls: {
        view: viewUrl, // 閲覧者向け（共有用）
        admin: adminUrl, // 管理者向け
      },
      lastUpdated: new Date().toLocaleString('ja-JP'),
      userEmail: userInfo.userEmail,
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
 * データベース最適化を実行（管理者用）
 */
function executeDataOptimization() {
  try {
    const targetUserId = '882d95c7-1fef-4739-a4b5-4ca02feaa69b';
    
    if (typeof optimizeSpecificUser === 'function') {
      const result = optimizeSpecificUser(targetUserId);
      console.info('最適化結果:', result);
      return result;
    } else {
      return {
        success: false,
        message: 'ConfigOptimizer.gs が読み込まれていません'
      };
    }
  } catch (error) {
    console.error('executeDataOptimization エラー:', error);
    return {
      success: false,
      error: error.message
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

    // SYSTEM_CONSTANTS.COLUMN_MAPPING から動的変換マップ生成（汎用化）
    const mappingConversions = {};
    Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
      // 各列のシステム内部キー（大文字）を動的生成
      mappingConversions[column.key] = column.key.toUpperCase();
    });

    // SYSTEM_CONSTANTS.COLUMN_MAPPINGから動的な列ヘッダーマップ生成（汎用化）
    const columnHeaders = {};
    Object.values(SYSTEM_CONSTANTS.COLUMN_MAPPING).forEach((column) => {
      const systemKey = column.key.toUpperCase();
      columnHeaders[systemKey] = column.header; // 例: 'ANSWER' => '回答'
    });

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
