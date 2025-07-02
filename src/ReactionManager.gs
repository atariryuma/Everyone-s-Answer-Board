/**
 * @fileoverview リアクション管理 - GAS互換版
 */

/**
 * リアクション処理（最適化版）
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクションキー
 * @param {string} userEmail - ユーザーメール
 * @returns {object} 処理結果
 */
function processReactionOptimized(spreadsheetId, sheetName, rowIndex, reactionKey, userEmail) {
  var lock = LockService.getScriptLock();
  
  try {
    lock.waitLock(10000);
    
    var service = getOptimizedSheetsService();
    
    // バッチ処理でヘッダーと行データを同時取得
    var ranges = [
      sheetName + '!1:1',
      sheetName + '!A' + rowIndex + ':Z' + rowIndex
    ];
    
    var response = batchGetSheetsData(service, spreadsheetId, ranges);
    var headers = response.valueRanges[0].values ? response.valueRanges[0].values[0] : [];
    var rowData = response.valueRanges[1].values ? response.valueRanges[1].values[0] : [];
    
    var reactionColumnName = COLUMN_HEADERS[reactionKey];
    var reactionColumnIndex = headers.indexOf(reactionColumnName);
    
    if (reactionColumnIndex === -1) {
      throw new Error('リアクション列が見つかりません: ' + reactionColumnName);
    }

    // 現在のリアクション文字列を解析
    var currentReactionString = rowData[reactionColumnIndex] || '';
    var currentReactions = parseReactionStringHelper(currentReactionString);
    
    // リアクションの追加/削除
    var userIndex = currentReactions.indexOf(userEmail);
    var action;
    
    if (userIndex >= 0) {
      currentReactions.splice(userIndex, 1);
      action = 'removed';
    } else {
      currentReactions.push(userEmail);
      action = 'added';
    }
    
    // 更新
    var updatedReactionString = currentReactions.join(', ');
    var cellRange = sheetName + '!' + String.fromCharCode(65 + reactionColumnIndex) + rowIndex;
    
    batchUpdateSheetsData(service, spreadsheetId, [{
      range: cellRange,
      values: [[updatedReactionString]]
    }]);

    debugLog('リアクション更新完了: ' + userEmail + ' → ' + reactionKey + ' (' + action + ')');
    
    return { 
      status: 'ok', 
      message: 'リアクションを更新しました。',
      action: action,
      count: currentReactions.length
    };

  } catch (e) {
    console.error('リアクション更新エラー: ' + e.message);
    throw new Error('リアクションの更新に失敗しました: ' + e.message);
  } finally {
    lock.releaseLock();
  }
}

/**
 * リアクション文字列をパース
 * @param {string} val - リアクション文字列
 * @returns {string[]} メールアドレスの配列
 */
function parseReactionStringHelper(val) {
  if (!val) return [];
  return val
    .toString()
    .split(',')
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s.length > 0; });
}