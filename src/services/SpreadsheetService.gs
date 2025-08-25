/**
 * @fileoverview スプレッドシートサービス - シートデータの操作とリアクション管理
 */

/**
 * 現在のシート名を取得
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {string} シート名
 */
function getCurrentSheetName(spreadsheetId) {
  try {
    const spreadsheet = Spreadsheet.openById(spreadsheetId);
    const activeSheet = spreadsheet.getActiveSheet();
    return activeSheet ? activeSheet.getName() : spreadsheet.getSheets()[0].getName();
  } catch (error) {
    logError(error, 'getCurrentSheetName', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return 'Sheet1';
  }
}

/**
 * シート設定を取得
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @returns {Object} 設定
 */
function getSheetConfig(userId, sheetName) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo) return {};
    
    const config = JSON.parse(userInfo.configJson || '{}');
    return config.sheets?.[sheetName] || {};
  } catch (error) {
    logError(error, 'getSheetConfig', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return {};
  }
}

/**
 * シートデータを取得
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @returns {Array} データ配列
 */
function getSheetData(userId, sheetName, classFilter, sortOrder, adminMode) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      return [];
    }
    
    const spreadsheet = Spreadsheet.openById(userInfo.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      return [];
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }
    
    const headers = data[0];
    const rows = data.slice(1);
    
    // フィルタリング
    let filteredRows = rows;
    if (classFilter && classFilter !== 'all') {
      const classIndex = headers.indexOf('クラス') !== -1 ? headers.indexOf('クラス') : headers.indexOf('class');
      if (classIndex !== -1) {
        filteredRows = rows.filter(row => row[classIndex] === classFilter);
      }
    }
    
    // ソート
    if (sortOrder === 'newest') {
      filteredRows.reverse();
    }
    
    // データを整形
    return filteredRows.map((row, index) => {
      const obj = { rowIndex: index + 1 };
      headers.forEach((header, i) => {
        obj[header] = row[i];
      });
      return obj;
    });
    
  } catch (error) {
    logError(error, 'getSheetData', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    return [];
  }
}

/**
 * 増分データを取得
 * @param {string} userId - ユーザーID
 * @param {string} sheetName - シート名
 * @param {number} sinceRowCount - 前回取得時の行数
 * @param {string} classFilter - クラスフィルター
 * @param {string} sortOrder - ソート順
 * @param {boolean} adminMode - 管理者モード
 * @returns {Object} 増分データ
 */
function getIncrementalData(userId, sheetName, sinceRowCount, classFilter, sortOrder, adminMode) {
  try {
    const allData = getSheetData(userId, sheetName, classFilter, sortOrder, adminMode);
    
    if (allData.length <= sinceRowCount) {
      return { hasNewData: false, data: [] };
    }
    
    return {
      hasNewData: true,
      data: allData.slice(sinceRowCount),
      totalCount: allData.length
    };
    
  } catch (error) {
    logError(error, 'getIncrementalData', ERROR_SEVERITY.LOW, ERROR_CATEGORIES.DATABASE);
    return { hasNewData: false, data: [] };
  }
}

/**
 * リアクションを更新
 * @param {string} userId - ユーザーID
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionType - リアクションタイプ
 * @param {string} sheetName - シート名
 * @returns {Object} 更新結果
 */
function updateReaction(userId, rowIndex, reactionType, sheetName) {
  try {
    const userInfo = getUserInfo(userId);
    if (!userInfo || !userInfo.spreadsheetId) {
      throw new Error('スプレッドシートが設定されていません');
    }
    
    const spreadsheet = Spreadsheet.openById(userInfo.spreadsheetId);
    const sheet = spreadsheet.getSheetByName(sheetName);
    if (!sheet) {
      throw new Error('シートが見つかりません');
    }
    
    // リアクション列を探す（なければ作成）
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    let reactionCol = headers.indexOf('リアクション');
    
    if (reactionCol === -1) {
      reactionCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, reactionCol).setValue('リアクション');
    } else {
      reactionCol += 1; // 1-indexed
    }
    
    // 現在のリアクションを取得
    const cell = sheet.getRange(rowIndex + 1, reactionCol); // +1 for header row
    const currentValue = cell.getValue() || '';
    const reactions = currentValue ? JSON.parse(currentValue) : {};
    
    // リアクションを更新
    const userEmail = userInfo.adminEmail;
    if (reactions[reactionType]) {
      if (reactions[reactionType].includes(userEmail)) {
        // 既存のリアクションを削除
        reactions[reactionType] = reactions[reactionType].filter(email => email !== userEmail);
        if (reactions[reactionType].length === 0) {
          delete reactions[reactionType];
        }
      } else {
        // リアクションを追加
        reactions[reactionType].push(userEmail);
      }
    } else {
      // 新しいリアクションタイプ
      reactions[reactionType] = [userEmail];
    }
    
    // セルを更新
    cell.setValue(Object.keys(reactions).length > 0 ? JSON.stringify(reactions) : '');
    
    return {
      success: true,
      reactions: reactions,
      rowIndex: rowIndex
    };
    
  } catch (error) {
    logError(error, 'updateReaction', ERROR_SEVERITY.MEDIUM, ERROR_CATEGORIES.DATABASE);
    throw error;
  }
}

/**
 * クイックスタートフォームを作成
 * @param {string} email - ユーザーメール
 * @param {string} userId - ユーザーID
 * @returns {Object} 作成結果
 */
function createQuickStartForm(email, userId) {
  try {
    // フォームを作成
    const form = FormApp.create(`みんなの回答ボード - ${email}`);
    form.setDescription('このフォームから投稿された回答がボードに表示されます');
    
    // 基本的な質問を追加
    form.addTextItem()
      .setTitle('お名前')
      .setRequired(true);
    
    form.addTextItem()
      .setTitle('ご意見・ご回答')
      .setRequired(true);
    
    // スプレッドシートを作成してリンク
    const spreadsheet = Spreadsheet.create(`回答データ - ${email}`);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    
    // ヘッダーを設定
    const sheet = spreadsheet.getActiveSheet();
    sheet.setName('回答');
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    logError(error, 'createQuickStartForm', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}

/**
 * カスタムフォームを作成
 * @param {string} email - ユーザーメール
 * @param {string} userId - ユーザーID
 * @param {Object} config - フォーム設定
 * @returns {Object} 作成結果
 */
function createCustomForm(email, userId, config) {
  try {
    // フォームを作成
    const form = FormApp.create(config.title || `カスタムフォーム - ${email}`);
    form.setDescription(config.description || '');
    
    // カスタム質問を追加
    if (config.questions) {
      config.questions.forEach(question => {
        switch (question.type) {
          case 'text':
            form.addTextItem()
              .setTitle(question.title)
              .setRequired(question.required || false);
            break;
          case 'paragraph':
            form.addParagraphTextItem()
              .setTitle(question.title)
              .setRequired(question.required || false);
            break;
          case 'multipleChoice':
            const mcItem = form.addMultipleChoiceItem()
              .setTitle(question.title)
              .setRequired(question.required || false);
            if (question.choices) {
              mcItem.setChoiceValues(question.choices);
            }
            break;
          case 'checkbox':
            const cbItem = form.addCheckboxItem()
              .setTitle(question.title)
              .setRequired(question.required || false);
            if (question.choices) {
              cbItem.setChoiceValues(question.choices);
            }
            break;
        }
      });
    }
    
    // スプレッドシートを作成してリンク
    const spreadsheet = Spreadsheet.create(config.spreadsheetTitle || `回答データ - ${email}`);
    form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
    
    return {
      formId: form.getId(),
      formUrl: form.getPublishedUrl(),
      spreadsheetId: spreadsheet.getId(),
      spreadsheetUrl: spreadsheet.getUrl()
    };
    
  } catch (error) {
    logError(error, 'createCustomForm', ERROR_SEVERITY.HIGH, ERROR_CATEGORIES.SYSTEM);
    throw error;
  }
}