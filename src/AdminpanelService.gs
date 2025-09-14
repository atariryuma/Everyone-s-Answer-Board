// ===========================================
// 📊 管理パネル用API関数（main.gsから移動）
// ===========================================

/**
 * スプレッドシート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @returns {Object} スプレッドシート一覧
 */
function getSpreadsheetList() {
  const started = Date.now();
  try {
    console.log('DataService.getSpreadsheetList: 開始 - GAS独立化完了');

    // ✅ GAS Best Practice: 直接API呼び出し（依存除去）
    const currentUser = Session.getActiveUser().getEmail();
    console.log('DataService.getSpreadsheetList: ユーザー情報', { currentUser });

    // DriveApp直接使用（効率重視）
    const files = DriveApp.searchFiles('mimeType="application/vnd.google-apps.spreadsheet"');
    console.log('DataService.getSpreadsheetList: DriveApp検索完了', {
      hasFiles: typeof files !== 'undefined',
      hasNext: files.hasNext()
    });

    // 権限テスト（必要最小限）
    let driveAccessOk = true;
    try {
      const testFiles = DriveApp.getFiles();
      driveAccessOk = testFiles.hasNext();
      console.log('DataService.getSpreadsheetList: Drive権限OK');
    } catch (driveError) {
      console.error('DataService.getSpreadsheetList: Drive権限エラー', driveError.message);
      driveAccessOk = false;
    }

    if (!driveAccessOk) {
      return {
        success: false,
        message: 'Driveアクセス権限がありません',
        spreadsheets: [],
        executionTime: `${Date.now() - started}ms`
      };
    }

    // スプレッドシート取得（高速処理）
    const spreadsheets = [];
    let count = 0;
    const maxCount = 25; // GAS制限対応

    console.log('DataService.getSpreadsheetList: スプレッドシート列挙開始');

    while (files.hasNext() && count < maxCount) {
      try {
        const file = files.next();
        spreadsheets.push({
          id: file.getId(),
          name: file.getName(),
          url: file.getUrl(),
          lastUpdated: file.getLastUpdated()
        });
        count++;
      } catch (fileError) {
        console.warn('DataService.getSpreadsheetList: ファイル処理スキップ', fileError.message);
        continue; // エラー時はスキップして継続
      }
    }

    console.log('DataService.getSpreadsheetList: 列挙完了', {
      totalFound: count,
      maxReached: count >= maxCount
    });

    // ✅ google.script.run互換 - シンプル形式に最適化
    const response = {
      success: true,
      spreadsheets,
      executionTime: `${Date.now() - started}ms`
    };

    // レスポンスサイズ監視（GAS制限対応）
    const responseSize = JSON.stringify(response).length;
    const responseSizeKB = Math.round(responseSize / 1024 * 100) / 100;

    console.log('DataService.getSpreadsheetList: 成功 - google.script.run最適化', {
      spreadsheetsCount: spreadsheets.length,
      executionTime: response.executionTime,
      responseSizeKB,
      responseValid: response !== null && typeof response === 'object'
    });

    // ✅ google.script.run互換性チェック
    if (!response || typeof response !== 'object' || !Array.isArray(response.spreadsheets)) {
      console.error('DataService.getSpreadsheetList: 無効なレスポンス形式', response);
      return {
        success: false,
        spreadsheets: [],
        message: 'レスポンス形式エラー'
      };
    }

    return response;
  } catch (error) {
    console.error('DataService.getSpreadsheetList: エラー', {
      error: error.message,
      executionTime: `${Date.now() - started}ms`
    });

    return {
      success: false,
      cached: false,
      executionTime: `${Date.now() - started}ms`,
      message: error.message || 'スプレッドシート一覧取得エラー',
      spreadsheets: []
    };
  }
}

/**
 * シート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */
function getSheetList(spreadsheetId) {
  try {
    if (!spreadsheetId) {
      return {
        success: false,
        message: 'スプレッドシートIDが指定されていません'
      };
    }

    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheets = spreadsheet.getSheets();
    const sheetList = sheets.map(sheet => ({
      name: sheet.getName(),
      index: sheet.getIndex(),
      rowCount: sheet.getMaxRows(),
      columnCount: sheet.getMaxColumns()
    }));

    return {
      success: true,
      sheets: sheetList
    };
  } catch (error) {
    console.error('DataService.getSheetList エラー:', error.message);
    return {
      success: false,
      message: error.message,
      sheets: []
    };
  }
}

/**
 * 列を分析
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */
function analyzeColumns(spreadsheetId, sheetName) {
  const started = Date.now();
  try {
    console.log('DataService.analyzeColumns: 開始', {
      spreadsheetId: spreadsheetId ? `${spreadsheetId.substring(0, 10)}...` : 'null',
      sheetName: sheetName || 'null'
    });

    if (!spreadsheetId || !sheetName) {
      const errorResponse = {
        success: false,
        message: 'スプレッドシートIDとシート名が必要です',
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      };
      console.error('DataService.analyzeColumns: 必須パラメータ不足', {
        spreadsheetId: !!spreadsheetId,
        sheetName: !!sheetName
      });
      return errorResponse;
    }

    // スプレッドシート接続テスト
    let spreadsheet;
    try {
      console.log('DataService.analyzeColumns: スプレッドシート接続開始');
      spreadsheet = SpreadsheetApp.openById(spreadsheetId);
      console.log('DataService.analyzeColumns: スプレッドシート接続成功');
    } catch (openError) {
      const errorResponse = {
        success: false,
        message: `スプレッドシートへのアクセスに失敗しました: ${openError.message}`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      };
      console.error('DataService.analyzeColumns: スプレッドシート接続失敗', {
        error: openError.message,
        spreadsheetId: `${spreadsheetId.substring(0, 10)}...`
      });
      return errorResponse;
    }

    // シート取得テスト
    let sheet;
    try {
      sheet = spreadsheet.getSheetByName(sheetName);
      if (!sheet) {
        const errorResponse = {
          success: false,
          message: `シート "${sheetName}" が見つかりません`,
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        };
        console.error('DataService.analyzeColumns: シートが見つかりません', {
          sheetName
        });
        return errorResponse;
      }
      console.log('DataService.analyzeColumns: シート取得成功');
    } catch (sheetError) {
      const errorResponse = {
        success: false,
        message: `シートアクセスエラー: ${sheetError.message}`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      };
      console.error('DataService.analyzeColumns: シートアクセスエラー', {
        error: sheetError.message,
        sheetName
      });
      return errorResponse;
    }

    // ヘッダー行を取得（1行目）- 堅牢化
    let headers = [];
    let sampleData = [];
    let lastColumn = 1;
    let lastRow = 1;

    try {
      console.log('DataService.analyzeColumns: シートサイズ取得開始');
      lastColumn = sheet.getLastColumn();
      lastRow = sheet.getLastRow();
      console.log('DataService.analyzeColumns: シートサイズ', {
        lastColumn,
        lastRow
      });

      if (lastColumn === 0 || lastRow === 0) {
        const errorResponse = {
          success: false,
          message: 'スプレッドシートが空です',
          headers: [],
          columns: [],
          columnMapping: { mapping: {}, confidence: {} }
        };
        console.error('DataService.analyzeColumns: 空のスプレッドシート', {
          lastColumn,
          lastRow
        });
        return errorResponse;
      }

      // ヘッダー行取得
      console.log('DataService.analyzeColumns: ヘッダー取得開始');
      [headers] = sheet.getRange(1, 1, 1, lastColumn).getValues();
      console.log('DataService.analyzeColumns: ヘッダー取得成功', {
        headersCount: headers.length
      });

      // サンプルデータを取得（最大5行）
      const sampleRowCount = Math.min(5, lastRow - 1);
      if (sampleRowCount > 0) {
        console.log('DataService.analyzeColumns: サンプルデータ取得開始');
        sampleData = sheet.getRange(2, 1, sampleRowCount, lastColumn).getValues();
        console.log('DataService.analyzeColumns: サンプルデータ取得成功', {
          sampleRowCount: sampleData.length
        });
      }
    } catch (rangeError) {
      const errorResponse = {
        success: false,
        message: `データ取得エラー: ${rangeError.message}`,
        headers: [],
        columns: [],
        columnMapping: { mapping: {}, confidence: {} }
      };
      console.error('DataService.analyzeColumns: データ取得エラー', {
        error: rangeError.message,
        lastColumn,
        lastRow
      });
      return errorResponse;
    }

    // 列情報を分析
    const columns = headers.map((header, index) => {
      const samples = sampleData.map(row => row[index]).filter(v => v);

      // 列タイプを推測
      let type = 'text';
      if (header.toLowerCase().includes('timestamp') || header.toLowerCase().includes('日時')) {
        type = 'datetime';
      } else if (header.toLowerCase().includes('email') || header.toLowerCase().includes('メール')) {
        type = 'email';
      } else if (header.toLowerCase().includes('class') || header.toLowerCase().includes('クラス')) {
        type = 'class';
      } else if (header.toLowerCase().includes('name') || header.toLowerCase().includes('名前')) {
        type = 'name';
      } else if (samples.length > 0 && samples.every(s => !isNaN(s))) {
        type = 'number';
      }

      return {
        index,
        header,
        type,
        samples: samples.slice(0, 3) // 最大3つのサンプル
      };
    });

    // AI検出シミュレーション（シンプルなパターンマッチング）
    const mapping = { mapping: {}, confidence: {} };

    headers.forEach((header, index) => {
      const headerLower = header.toLowerCase();

      // 回答列の検出
      if (headerLower.includes('回答') || headerLower.includes('answer') || headerLower.includes('意見')) {
        mapping.mapping.answer = index;
        mapping.confidence.answer = 85;
      }

      // 理由列の検出
      if (headerLower.includes('理由') || headerLower.includes('根拠') || headerLower.includes('reason')) {
        mapping.mapping.reason = index;
        mapping.confidence.reason = 80;
      }

      // クラス列の検出
      if (headerLower.includes('クラス') || headerLower.includes('class') || headerLower.includes('組')) {
        mapping.mapping.class = index;
        mapping.confidence.class = 90;
      }

      // 名前列の検出
      if (headerLower.includes('名前') || headerLower.includes('name') || headerLower.includes('氏名')) {
        mapping.mapping.name = index;
        mapping.confidence.name = 85;
      }
    });

    const result = {
      success: true,
      headers,
      columns,
      columnMapping: mapping,
      sampleData: sampleData.slice(0, 3), // 最大3行のサンプル
      executionTime: `${Date.now() - started}ms`
    };

    console.log('DataService.analyzeColumns: 成功', {
      headersCount: headers.length,
      columnsCount: columns.length,
      mappingDetected: Object.keys(mapping.mapping).length,
      executionTime: result.executionTime
    });

    return result;

  } catch (error) {
    const errorResponse = {
      success: false,
      message: `予期しないエラーが発生しました: ${error.message}`,
      headers: [],
      columns: [],
      columnMapping: { mapping: {}, confidence: {} },
      executionTime: `${Date.now() - started}ms`
    };

    console.error('DataService.analyzeColumns: 予期しないエラー', {
      error: error.message,
      stack: error.stack,
      executionTime: errorResponse.executionTime
    });

    return errorResponse;
  }
}

// ===========================================
// 🔧 ヘルパー関数（未定義関数の実装）
// ===========================================

/**
 * リアクション列の取得または作成
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} reactionType - リアクションタイプ
 * @returns {number} 列番号
 */
function getOrCreateReactionColumn(sheet, reactionType) {
  try {
    const [headers] = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues();
    const reactionHeader = reactionType.toUpperCase();

    // 既存の列を探す
    const existingIndex = headers.findIndex(header =>
      String(header).toUpperCase().includes(reactionHeader)
    );

    if (existingIndex !== -1) {
      return existingIndex + 1; // 1-based index
    }

    // 新しい列を作成
    const newColumn = sheet.getLastColumn() + 1;
    sheet.getRange(1, newColumn).setValue(reactionHeader);
    return newColumn;
  } catch (error) {
    console.error('getOrCreateReactionColumn: エラー', error.message);
    return null;
  }
}

/**
 * ハイライト列の更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @returns {boolean} 成功可否
 */
function updateHighlightInSheet(config, rowId) {
  try {
    const spreadsheet = SpreadsheetApp.openById(config.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(config.sheetName);

    if (!sheet) {
      throw new Error('シートが見つかりません');
    }

    // 行番号抽出（row_3 → 3）
    const rowNumber = parseInt(rowId.replace('row_', ''));
    if (isNaN(rowNumber) || rowNumber < 2) {
      throw new Error('無効な行ID');
    }

    // ハイライト列の取得・作成
    const highlightColumn = getOrCreateReactionColumn(sheet, 'HIGHLIGHT');
    if (!highlightColumn) {
      throw new Error('ハイライト列の作成に失敗');
    }

    // 現在値取得・トグル
    const currentValue = sheet.getRange(rowNumber, highlightColumn).getValue();
    const isCurrentlyHighlighted = currentValue === 'TRUE' || currentValue === true;
    const newValue = isCurrentlyHighlighted ? 'FALSE' : 'TRUE';

    sheet.getRange(rowNumber, highlightColumn).setValue(newValue);

    console.info('DataService.updateHighlightInSheet: ハイライト切り替え完了', {
      rowId,
      oldValue: currentValue,
      newValue,
      highlighted: !isCurrentlyHighlighted
    });

    return true;
  } catch (error) {
    console.error('DataService.updateHighlightInSheet: エラー', error.message);
    return false;
  }
}

/**
 * リアクションパラメータ検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類
 * @returns {boolean} 検証結果
 */
function validateReactionParams(spreadsheetId, sheetName, rowIndex, reactionKey) {
  if (!spreadsheetId || !sheetName || !rowIndex || !reactionKey) {
    return false;
  }

  if (rowIndex < 2) { // ヘッダー行は1
    return false;
  }

  const validReactions = ['UNDERSTAND', 'LIKE', 'CURIOUS', 'HIGHLIGHT'];
  return validReactions.includes(reactionKey);
}
