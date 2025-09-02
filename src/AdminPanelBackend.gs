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
 * Services.user.currentを活用した統一インターフェース - 全システムでgetActiveUserInfo()を使用
 */
function getActiveUserInfo() {
  return Services.user.current;
}

function debugConstants() {
  console.log('SYSTEM_CONSTANTS:', typeof SYSTEM_CONSTANTS);
  if (typeof SYSTEM_CONSTANTS !== 'undefined') {
    console.log('COLUMN_MAPPING:', typeof SYSTEM_CONSTANTS.COLUMN_MAPPING);
    console.log(
      'COLUMN_MAPPING keys:',
      SYSTEM_CONSTANTS.COLUMN_MAPPING ? Object.keys(SYSTEM_CONSTANTS.COLUMN_MAPPING) : 'undefined'
    );
  }
}

// =============================================================================
// データソース管理
// =============================================================================

/**
 * スプレッドシート一覧を取得（管理パネル用）- シンプル版
 * @returns {Array} スプレッドシート一覧
 */
function getSpreadsheetList() {
  try {
    console.log('getSpreadsheetList: スプレッドシート一覧取得開始');

    const currentUserEmail = User.email();
    console.log('現在のユーザー:', currentUserEmail);

    // キャッシュチェック（5分間有効）
    const cacheKey = `spreadsheet_list_${currentUserEmail}`;
    const cached = CacheService.getScriptCache().get(cacheKey);
    if (cached) {
      try {
        const cacheData = JSON.parse(cached);
        const cacheAge = Date.now() - cacheData.timestamp;
        if (cacheAge < 300000) {
          console.info(`getSpreadsheetList: キャッシュヒット (${Math.round(cacheAge / 1000)}秒前)`);
          return cacheData.data;
        }
      } catch (e) {
        // キャッシュ読み込みエラーは無視して続行
      }
    }

    const startTime = Date.now();
    const spreadsheets = [];
    const maxResults = 30; // 十分な数を取得
    
    // シンプルな検索：スプレッドシートのみ、ゴミ箱以外
    const searchQuery = `mimeType="application/vnd.google-apps.spreadsheet" and trashed=false`;
    console.info('getSpreadsheetList: 検索開始');
    
    const files = DriveApp.searchFiles(searchQuery);
    const thirtyDaysAgo = new Date(Date.now() - 30 * 24 * 60 * 60 * 1000);

    let count = 0;
    let totalChecked = 0;

    while (files.hasNext() && count < maxResults && totalChecked < 200) {
      totalChecked++;
      const file = files.next();
      
      try {
        // オーナーチェック
        const owner = file.getOwner();
        if (!owner || owner.getEmail() !== currentUserEmail) {
          continue;
        }
        
        // 30日以内にアクセスしたかチェック
        const lastUpdated = file.getLastUpdated();
        if (lastUpdated < thirtyDaysAgo) {
          continue;
        }
        
        // 条件を満たすファイルを追加
        spreadsheets.push({
          id: file.getId(),
          name: file.getName(),
          lastModified: lastUpdated.toISOString(),
          owner: currentUserEmail,
          isOwner: true
        });
        count++;
        
      } catch (error) {
        // 個別ファイルのエラーは無視して続行
        continue;
      }
    }


    // 最終更新順でソート
    spreadsheets.sort((a, b) => new Date(b.lastModified) - new Date(a.lastModified));

    const executionTime = Date.now() - startTime;
    console.info(`getSpreadsheetList: ${spreadsheets.length}件取得（実行時間: ${executionTime}ms、検査数: ${totalChecked}）`);

    // 結果をキャッシュ（5分間）
    try {
      const cacheData = {
        timestamp: Date.now(),
        data: spreadsheets
      };
      CacheService.getScriptCache().put(cacheKey, JSON.stringify(cacheData), 300);
      console.info('getSpreadsheetList: キャッシュ保存完了');
    } catch (e) {
      // キャッシュ保存エラーは無視
    }

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
    // 🚀 最適化：軽量チェックを先に実行
    let spreadsheet;
    try {
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    } catch (accessError) {
      // アクセス権限がない場合は即座にfalseを返す
      return { hasForm: false };
    }

    // 🚀 最適化：getFormUrl()の例外処理を軽量化
    let formUrl;
    try {
      formUrl = spreadsheet.getFormUrl();

      // FormURLがnullや空文字の場合は即座にfalseを返す
      if (!formUrl) {
        return { hasForm: false };
      }
    } catch (error) {
      // フォーム連携がない場合、getFormUrl()でエラーが発生する
      return { hasForm: false };
    }

    // 🚀 最適化：FormApp.openByUrlは重いので必要最小限に
    try {
      const form = FormApp.openByUrl(formUrl);

      return {
        hasForm: true,
        formUrl: form.getPublishedUrl(),
        formTitle: form.getTitle(),
        formId: form.getId(),
      };
    } catch (formError) {
      // フォームが削除されている等の場合
      return { hasForm: false };
    }
  } catch (error) {
    // 予期しないエラー
    console.warn(`フォーム連携チェック予期しないエラー: ${spreadsheetId}`, error.message);
    return { hasForm: false };
  }
}

/**
 * 特定シートのフォーム連携情報を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} シート別フォーム連携情報
 */
function getSheetFormConnection(spreadsheetId, sheetName) {
  try {
    console.log(
      `getSheetFormConnection: シート別フォーム連携チェック開始 ${spreadsheetId}/${sheetName}`
    );

    // まずスプレッドシート全体のフォーム連携をチェック
    const overallFormInfo = checkFormConnection(spreadsheetId);
    if (!overallFormInfo || !overallFormInfo.hasForm) {
      return { success: false, reason: 'スプレッドシートにフォーム連携なし' };
    }

    // シート名がフォーム回答用パターンに一致するかチェック
    const isFormResponsePattern =
      sheetName.startsWith('フォームの回答') || 
      sheetName.startsWith('Form Responses') ||
      sheetName.includes('フォーム') ||
      sheetName.includes('回答');

    if (isFormResponsePattern) {
      // パターンに一致する場合、基本的にフォーム連携とみなす
      return {
        success: true,
        formData: {
          formUrl: overallFormInfo.formUrl,
          formTitle: overallFormInfo.formTitle,
          formId: overallFormInfo.formId,
          sheetName: sheetName,
          isConnectedToThisSheet: true,
        }
      };
    }

    // より詳細なフォーム情報が必要な場合のみ FormApp を使用
    try {
      let form = null;
      
      // まずIDでの取得を試行
      if (overallFormInfo.formId) {
        try {
          form = FormApp.openById(overallFormInfo.formId);
        } catch (idError) {
          console.warn('FormIDでの取得失敗:', idError.message);
        }
      }
      
      // IDで取得できない場合はURLで試行
      if (!form && overallFormInfo.formUrl) {
        try {
          form = FormApp.openByUrl(overallFormInfo.formUrl);
        } catch (urlError) {
          console.warn('FormURLでの取得失敗:', urlError.message);
        }
      }

      if (form) {
        const formDestination = form.getDestinationId();
        
        // フォームの送信先スプレッドシートが一致するか確認
        if (formDestination === spreadsheetId) {
          return {
            success: true,
            formData: {
              formUrl: overallFormInfo.formUrl,
              formTitle: overallFormInfo.formTitle,
              formId: overallFormInfo.formId,
              sheetName: sheetName,
              isConnectedToThisSheet: true,
            }
          };
        }
      }
      
      // フォームが取得できない場合でも、基本情報は返す
      return {
        success: false,
        reason: 'フォーム詳細確認不可（基本連携あり）',
        availableForm: overallFormInfo,
      };
      
    } catch (formError) {
      console.warn('フォーム詳細取得エラー:', formError.message);
      
      // エラーが発生してもフォーム連携の基本情報は提供
      return {
        success: false,
        reason: 'フォーム詳細取得失敗',
        availableForm: overallFormInfo,
      };
    }
  } catch (error) {
    console.error(`getSheetFormConnection エラー: ${spreadsheetId}/${sheetName}`, error.message);
    return { success: false, reason: 'エラー発生', error: error.message };
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
      const hasTimestamp = headerRow.some((header) => {
        if (!header) return false;
        const headerStr = String(header).toLowerCase();
        return (
          headerStr.includes('timestamp') ||
          headerStr.includes('タイムスタンプ') ||
          headerStr.includes('回答時刻')
        );
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
        formTitle: formInfo && formInfo.hasForm ? formInfo.formTitle : null,
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
    const headerIndices = getSpreadsheetHeaders(spreadsheetId, sheetName, { validate: true });
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 超高精度AI列マッピングを活用（新システム）
    // 【プロダクションバグ修正】 空オブジェクト {} も truthy と評価される問題を解決
    // 問題: headerIndices が {} の場合、truthy判定でconvertIndicesToMappingが呼ばれ
    // "Cannot convert undefined or null to object" エラーが発生
    const hasValidHeaderIndices =
      headerIndices && typeof headerIndices === 'object' && Object.keys(headerIndices).length > 0;

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
      const updatedHeaderIndices = getSpreadsheetHeaders(spreadsheetId, sheetName, { forceRefresh: true, validate: true });
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
          lastAccessedAt: new Date().toISOString(),
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
        minLength: 15,
      },
      reason: {
        aiPatterns: ['理由', '体験', '根拠', '詳細'],
        alternates: ['理由', '根拠', '体験', 'なぜ', '詳細', '説明'],
      },
      class: {
        alternates: ['クラス', '学年'],
      },
      name: {
        alternates: ['名前', '氏名', 'お名前'],
      },
    };

    // 高精度パターンマッチング
    headers.forEach((header, index) => {
      const headerLower = header.toLowerCase();

      Object.keys(fallbackPatterns).forEach((key) => {
        const pattern = fallbackPatterns[key];
        let matchScore = 0;

        // aiPatterns による高精度検出
        if (pattern.aiPatterns) {
          pattern.aiPatterns.forEach((aiPattern) => {
            if (headerLower.includes(aiPattern.toLowerCase())) {
              matchScore = Math.max(matchScore, 90);
            }
          });

          // 質問文の特別処理（長い文章 + 疑問符）
          if (key === 'answer' && pattern.minLength && header.length > pattern.minLength) {
            const hasQuestionMark = pattern.aiPatterns.some((p) => header.includes(p));
            if (hasQuestionMark) {
              matchScore = Math.max(matchScore, 95); // 最高精度
            }
          }
        }

        // alternates による補完検出
        if (matchScore === 0 && pattern.alternates) {
          pattern.alternates.forEach((alternate) => {
            if (headerLower.includes(alternate.toLowerCase())) {
              matchScore = Math.max(matchScore, 80);
            }
          });
        }

        // より高いスコアで置き換え（重複チェック付き）
        if (matchScore > 0) {
          const currentScore = confidence[key] || 0;
          
          // 重複チェック：既に他のキーで使用されているインデックスは避ける
          const isIndexAlreadyUsed = Object.keys(mapping).some(
            existingKey => existingKey !== key && mapping[existingKey] === index
          );
          
          if (matchScore > currentScore) {
            if (!isIndexAlreadyUsed || matchScore > 90) {
              // 未使用インデックス、または非常に高い信頼度の場合は割り当て
              mapping[key] = index;
              confidence[key] = matchScore;
            } else if (!mapping[key] && matchScore > 75) {
              // 現在未割り当てで高い信頼度がある場合も割り当て
              mapping[key] = index;
              confidence[key] = matchScore - 5; // 軽度の重複ペナルティ
            }
          }
        }
      });
    });

    // デフォルト値の設定
    ['answer', 'reason', 'class', 'name'].forEach((key) => {
      if (mapping[key] === undefined) mapping[key] = null;
    });

    mapping.confidence = confidence;

    console.log('detectColumnMapping: Enhanced AI detection result', {
      mapping,
      confidence,
      usedPatterns: 'aiPatterns + alternates + minLength',
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
          const hasAIPattern = column.aiPatterns.some(
            (p) => header.includes(p) || headerLower.includes(p.toLowerCase())
          );
          if (hasAIPattern) {
            matchScore = Math.max(matchScore, 92); // 質問文特別検出
          }
        }
      }

      // より高い信頼度で置き換え（重複チェック付き）
      if (matchScore > 0) {
        const currentMappedIndex = mapping[fieldKey];
        const currentScore = confidence[fieldKey] || 0;
        
        // 重複チェック：既に他のフィールドで使用されているインデックスは避ける
        const isIndexAlreadyUsed = Object.keys(mapping).some(
          key => key !== fieldKey && mapping[key] === index
        );
        
        if (matchScore > currentScore) {
          if (!isIndexAlreadyUsed || matchScore > 85) {
            // より高いスコア、または未使用インデックス、または非常に高い信頼度の場合は割り当て
            mapping[fieldKey] = index;
            confidence[fieldKey] = matchScore;
          } else if (!currentMappedIndex && matchScore > 70) {
            // 現在未割り当てで中程度の信頼度がある場合も割り当て
            mapping[fieldKey] = index;
            confidence[fieldKey] = matchScore - 10; // 重複ペナルティ
          }
        }
      }
    });
  });

  // confidenceを返り値に追加
  mapping.confidence = confidence;

  // 既にSYSTEM_CONSTANTSベースの高精度検出が完了しているため、追加処理は不要

  console.log('detectColumnMapping: 高精度マッピング完了', {
    headers,
    mapping,
    confidence,
    usedTechnology: 'SYSTEM_CONSTANTS + aiPatterns + 重複防止',
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
    const headerIndices = getSpreadsheetHeaders(spreadsheetId, sheetName, { validate: true });
    const headerRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // 超高精度AI列マッピングを活用（新システム）
    // 【プロダクションバグ修正】 空オブジェクト {} も truthy と評価される問題を解決
    const hasValidHeaderIndices =
      headerIndices && typeof headerIndices === 'object' && Object.keys(headerIndices).length > 0;

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
      const updatedHeaderIndices = getSpreadsheetHeaders(spreadsheetId, sheetName, { forceRefresh: true, validate: true });
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
 * @param {Object} headerIndices - getSpreadsheetHeadersから返されるインデックス
 * @param {Array} headerRow - ヘッダー行（表示用）
 * @returns {Object} AdminPanel用マッピング
 */
function convertIndicesToMapping(headerIndices, headerRow) {
  // nullチェックを追加
  if (!headerIndices || typeof headerIndices !== 'object') {
    console.error('convertIndicesToMapping: headerIndices is null or invalid', headerIndices);
    // エラーを投げる代わりにAI判定にフォールバック
    console.log(
      'convertIndicesToMapping: Falling back to AI detection due to invalid headerIndices'
    );
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
    headerRowLength: headerRow.length,
  });

  // SYSTEM_CONSTANTS の安全性チェック
  if (!SYSTEM_CONSTANTS || !SYSTEM_CONSTANTS.COLUMN_MAPPING) {
    console.warn(
      'convertIndicesToMapping: SYSTEM_CONSTANTS.COLUMN_MAPPING is not available, falling back to AI detection'
    );
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
              console.log(
                `convertIndicesToMapping: alternateマッチ ${uiFieldName}: "${actualHeader}" -> ${alternate} (index: ${index})`
              );
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
          if (
            actualHeader.length > 15 &&
            (actualHeader.includes('？') ||
              actualHeader.includes('?') ||
              actualHeader.includes('どうして') ||
              actualHeader.includes('なぜ') ||
              actualHeader.includes('思います') ||
              actualHeader.includes('考え'))
          ) {
            columnIndex = index;
            console.log(
              `convertIndicesToMapping: 質問文ヘッダーを回答列として認識 ${uiFieldName}: "${actualHeader.substring(0, 30)}..." (index: ${index})`
            );
            break;
          }
        }
      }

      // 理由っぽいヘッダーの特別処理（reason列）
      if (columnIndex === null && uiFieldName === 'reason') {
        for (const [actualHeader, index] of Object.entries(headerIndices)) {
          if (
            actualHeader.includes('理由') ||
            actualHeader.includes('体験') ||
            actualHeader.includes('根拠') ||
            actualHeader.includes('なぜ')
          ) {
            columnIndex = index;
            console.log(
              `convertIndicesToMapping: 理由系ヘッダーを理由列として認識 ${uiFieldName}: "${actualHeader}" (index: ${index})`
            );
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
 * 高速公開情報取得（データベースのみ、段階的ローディング用）
 * @param {string} userId ユーザーID（オプション）
 * @returns {Object} 基本公開情報
 */
function getQuickPublishedInfo(userId = null) {
  try {
    const currentUser = User.email();
    const targetUserId = userId || currentUser;
    const userInfo = userId ? DB.findUserById(userId) : DB.findUserByEmail(currentUser);

    if (!userInfo) {
      return {
        appPublished: false,
        isLoading: false,
        error: 'ユーザー情報が見つかりません'
      };
    }

    // configJsonを完全にパース
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      console.error('getQuickPublishedInfo: configJson parse error:', e);
      config = {};
    }

    // すべての情報を含めて返す
    return {
      // 基本的な公開情報
      appPublished: config.appPublished || false,
      appName: config.appName || 'アプリ名未設定',
      appUrl: config.appUrl,
      publishedAt: config.publishedAt,
      
      // configJsonからの詳細情報
      sheetName: config.sheetName,
      columnMapping: config.columnMapping,
      compatibleMapping: config.compatibleMapping,
      formTitle: config.formTitle,
      lastModified: config.lastModified,
      lastConnected: config.lastConnected,
      connectionMethod: config.connectionMethod,
      missingColumnsHandled: config.missingColumnsHandled,
      
      // データベースからの情報
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: userInfo.formUrl,
      
      isLoading: false
    };
  } catch (error) {
    console.error('getQuickPublishedInfo エラー:', error);
    return {
      appPublished: false,
      isLoading: false,
      error: error.message
    };
  }
}

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
      // デフォルト設定（最小限）
      return {
        setupStatus: 'pending',
        appPublished: false,
        displaySettings: { showNames: true, showReactions: true },
        user: currentUser,
      };
    }

    // configJsonをパース
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      console.error('getCurrentConfig: configJson parse error:', e);
      config = {};
    }

    // データベースの全情報とconfigJsonを統合
    const fullConfig = {
      // データベースの基本情報
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: userInfo.formUrl,
      isActive: userInfo.isActive,
      createdAt: userInfo.createdAt,
      lastAccessedAt: userInfo.lastAccessedAt,
      
      // configJson内の詳細設定を展開
      ...config,
      
      // configJson内の特定フィールドを明示的に設定（確実性のため）
      appName: config.appName || 'みんなの回答ボード',
      setupStatus: config.setupStatus || 'pending',
      appPublished: config.appPublished || false,
      publishedAt: config.publishedAt,
      appUrl: config.appUrl,
      sheetName: config.sheetName,
      columnMapping: config.columnMapping,
      compatibleMapping: config.compatibleMapping,
      lastConnected: config.lastConnected,
      connectionMethod: config.connectionMethod,
      missingColumnsHandled: config.missingColumnsHandled,
      formTitle: config.formTitle,
      lastModified: config.lastModified,
      
      // デフォルト表示設定
      displaySettings: config.displaySettings || { showNames: true, showReactions: true },
    };

    console.log('getCurrentConfig: 完全な設定を返します', {
      hasSheetName: !!fullConfig.sheetName,
      hasColumnMapping: !!fullConfig.columnMapping,
      hasFormTitle: !!fullConfig.formTitle,
      appPublished: fullConfig.appPublished
    });

    return fullConfig;
    
  } catch (error) {
    console.error('getCurrentConfig エラー:', error);
    return {
      setupStatus: 'error',
      appPublished: false,
      displaySettings: { showNames: true, showReactions: true },
      error: error.message,
    };
  }
}

/**
 * 詳細設定取得（スプレッドシートアクセス含む、重い処理）
 * @param {Object} userInfo ユーザー情報
 * @returns {Object} 詳細設定情報
 */
function getFullConfigWithSpreadsheetAccess(userInfo) {
  try {
    let config = {};
    try {
      config = JSON.parse(userInfo.configJson || '{}');
    } catch (e) {
      console.error('getFullConfigWithSpreadsheetAccess: JSON parse error:', e);
      config = {};
    }

    // DB情報と統合（新フィールド対応）
    const result = {
      userId: userInfo.userId,
      userEmail: userInfo.userEmail,
      spreadsheetId: userInfo.spreadsheetId,
      spreadsheetUrl: userInfo.spreadsheetUrl,
      formUrl: userInfo.formUrl,
      sheetName: userInfo.sheetName,
      publishedAt: userInfo.publishedAt,
      appUrl: userInfo.appUrl,
      lastModified: userInfo.lastModified,
      isActive: userInfo.isActive,
      createdAt: userInfo.createdAt,
      lastAccessedAt: userInfo.lastAccessedAt,
      ...config
    };

    // 列マッピングの復元
    if (userInfo.columnMappingJson) {
      try {
        result.columnMapping = JSON.parse(userInfo.columnMappingJson);
      } catch (e) {
        console.warn('columnMappingJson parse error:', e);
        result.columnMapping = {};
      }
    }

    return result;
  } catch (error) {
    console.error('getFullConfigWithSpreadsheetAccess エラー:', error);
    return {
      setupStatus: 'error',
      appPublished: false,
      displaySettings: { showNames: true, showReactions: true },
      error: error.message
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

    // シンプルな列マッピング保存（重複検出ロジック削除）

    // DB更新データ準備（新フィールド対応）
    const updateData = {
      spreadsheetId: config.spreadsheetId,
      spreadsheetUrl: config.spreadsheetId ? 
        `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}/edit` : '',
      formUrl: config.formUrl,
      sheetName: config.sheetName,
      columnMappingJson: JSON.stringify(config.columnMapping || {}),
      lastModified: new Date().toISOString()
    };
    
    // JSON最適化：重複データ除去
    const optimizedConfig = {
      appName: config.appName,
      setupStatus: config.setupStatus,
      appPublished: config.appPublished || false,
      displaySettings: config.displaySettings || { showNames: true, showReactions: true },
      appUrl: config.appUrl
    };
    
    updateData.configJson = JSON.stringify(optimizedConfig);

    // 設定を更新
    const updateResult = updateUserSpreadsheetConfig(userInfo.userId, updateData);

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

    // 公開状態設定
    config.appPublished = true;
    config.publishedAt = new Date().toISOString();
    
    // シンプルな列マッピング保存（重複検出ロジック削除）

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
      // DB情報更新（新フィールド対応）
      const updateData = {
        spreadsheetId: config.spreadsheetId,
        spreadsheetUrl: config.spreadsheetId ? 
          `https://docs.google.com/spreadsheets/d/${config.spreadsheetId}/edit` : '',
        formUrl: config.formUrl,
        sheetName: config.sheetName,
        columnMappingJson: JSON.stringify(config.columnMapping || {}),
        publishedAt: config.publishedAt,
        appUrl: publishResult.appUrl,
        lastModified: new Date().toISOString()
      };
      
      // JSON最適化
      const optimizedConfig = {
        appName: config.appName,
        setupStatus: 'completed',
        appPublished: true,
        publishedAt: config.publishedAt,
        displaySettings: config.displaySettings || { showNames: true, showReactions: true },
        appUrl: publishResult.appUrl
      };
      
      updateData.configJson = JSON.stringify(optimizedConfig);
      
      // データベース更新
      updateUser(userInfo.userId, updateData);
      
      // 公開後すぐにフッター情報を更新
      try {
        const boardInfo = getCurrentBoardInfoAndUrls();
        console.log('publishApplication: フッター情報更新完了', boardInfo);
      } catch (boardInfoError) {
        console.warn('publishApplication: フッター情報更新に失敗:', boardInfoError);
      }
      
      return {
        success: true,
        config: getCurrentConfig(),
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
      console.log('getOrCreateWebAppUrl: 既存URL使用', existingUrl);
      return existingUrl;
    }

    // WebアプリのベースURLを取得（正式な方法）
    let baseUrl;
    try {
      baseUrl = ScriptApp.getService().getUrl();
    } catch (e) {
      // フォールバック: Script IDから構築
      const scriptId = ScriptApp.getScriptId();
      baseUrl = `https://script.google.com/macros/s/${scriptId}/exec`;
    }

    // ユーザー用の回答ボードURL（Page.htmlが表示される）
    const webAppUrl = `${baseUrl}?userId=${userId}`;
    
    console.log('getOrCreateWebAppUrl: 新規URL生成', webAppUrl);
    
    // URLを保存（オプション）
    try {
      props.setProperty(`${userId}_app_url`, webAppUrl);
    } catch (saveError) {
      console.warn('URL保存失敗:', saveError.message);
    }

    return webAppUrl;
  } catch (error) {
    console.error('getOrCreateWebAppUrl エラー:', error);
    // 最終フォールバック
    const scriptId = ScriptApp.getScriptId();
    return `https://script.google.com/macros/s/${scriptId}/exec?userId=${userId}`;
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
    const currentUser = User.email();
    const userInfo = DB.findUserByEmail(currentUser);
    
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
      userEmail: userInfo.userEmail,
      hasSpreadsheetId: !!userInfo.spreadsheetId,
    });

    // configJsonを完全に解析
    let config = {};
    let boardData = null;
    let questionText = '問題読み込み中...';

    try {
      config = JSON.parse(userInfo.configJson || '{}');
      console.log('getCurrentBoardInfoAndUrls: 設定解析成功', {
        appPublished: config.appPublished,
        hasSheetName: !!config.sheetName,
        hasFormTitle: !!config.formTitle,
        appName: config.appName
      });
    } catch (e) {
      console.warn('getCurrentBoardInfoAndUrls: 設定JSON解析エラー:', e.message);
    }

    // ボードデータの取得
    if (userInfo.spreadsheetId) {
      try {
        // configJsonから保存されているシート名を優先使用
        const sheetName = config.sheetName || config.publishedSheetName || config.activeSheetName || 'フォームの回答 1';
        console.log('getCurrentBoardInfoAndUrls: 使用するシート名:', sheetName);

        boardData = getSheetData(userInfo.userId, sheetName);
        
        // 問題文はフォームタイトルまたはヘッダーから取得
        questionText = config.formTitle || boardData?.header || '問題文が設定されていません';
        
        console.log('getCurrentBoardInfoAndUrls: ボードデータ取得成功', {
          hasHeader: !!boardData?.header,
          formTitle: config.formTitle,
          dataCount: boardData?.data?.length || 0
        });
      } catch (error) {
        console.warn('getCurrentBoardInfoAndUrls: ボードデータ取得エラー:', error.message);
        // エラー時でもフォームタイトルを使用
        questionText = config.formTitle || '問題文の読み込みに失敗しました';
      }
    } else {
      questionText = config.formTitle || 'スプレッドシートが設定されていません';
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

    // URLを構築（configJsonに保存されているappUrlも考慮）
    const viewUrl = config.appUrl || `${baseUrl}?userId=${encodeURIComponent(userInfo.userId)}`;
    const adminUrl = `${baseUrl}?mode=admin&userId=${encodeURIComponent(userInfo.userId)}`;

    const result = {
      isActive: !!userInfo.spreadsheetId && userInfo.isActive !== false && userInfo.isActive !== 'FALSE',
      appPublished: config.appPublished === true,
      questionText: questionText,
      sheetName: config.sheetName || 'シート名未設定',
      appName: config.appName || '未設定',
      urls: {
        view: viewUrl,  // 閲覧者向け（共有用）
        admin: adminUrl, // 管理者向け
      },
      lastUpdated: config.lastModified || new Date().toLocaleString('ja-JP'),
      userEmail: userInfo.userEmail,
      // 追加情報
      formTitle: config.formTitle,
      totalResponses: boardData?.data?.length || 0,
      hasPublishedData: config.appPublished === true,
      publicUrl: viewUrl,
      adminUrl: adminUrl
    };

    console.log('getCurrentBoardInfoAndUrls: 成功', {
      isActive: result.isActive,
      appPublished: result.appPublished,
      hasQuestionText: !!result.questionText,
      appName: result.appName,
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
        message: 'ConfigOptimizer.gs が読み込まれていません',
      };
    }
  } catch (error) {
    console.error('executeDataOptimization エラー:', error);
    return {
      success: false,
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

/**
 * ヘッダーインデックスを取得（管理パネル用）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} ヘッダーインデックス情報
 */
function getHeaderIndices(spreadsheetId, sheetName) {
  try {
    console.log('getHeaderIndices: ヘッダー取得開始', { spreadsheetId, sheetName });
    
    // 既存の汎用関数を活用
    const headerIndices = getSpreadsheetHeaders(spreadsheetId, sheetName, { 
      validate: false,
      useCache: true 
    });
    
    console.log('getHeaderIndices: ヘッダー取得成功', {
      columnCount: Object.keys(headerIndices).length
    });
    
    return headerIndices;
    
  } catch (error) {
    console.error('getHeaderIndices エラー:', error.message);
    throw new Error(`ヘッダー取得エラー: ${error.message}`);
  }
}

console.log('AdminPanel.gs 読み込み完了');
