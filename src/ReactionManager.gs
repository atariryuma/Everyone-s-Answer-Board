/**
 * @fileoverview リアクション管理クラス
 */

class ReactionManager {
  /**
   * リアクション処理（最適化版）
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string} sheetName - シート名
   * @param {number} rowIndex - 行インデックス
   * @param {string} reactionKey - リアクションキー
   * @param {string} userEmail - ユーザーメール
   * @returns {object} 処理結果
   */
  static async processReaction(spreadsheetId, sheetName, rowIndex, reactionKey, userEmail) {
    const lock = LockService.getScriptLock();
    
    try {
      lock.waitLock(10000);
      
      const service = DatabaseManager.getSheetsService();
      
      // バッチ処理でヘッダーと行データを同時取得
      const ranges = [
        `${sheetName}!1:1`,
        `${sheetName}!A${rowIndex}:Z${rowIndex}`
      ];
      
      const response = await service.batchGet(spreadsheetId, ranges);
      const headers = response.valueRanges[0].values ? response.valueRanges[0].values[0] : [];
      const rowData = response.valueRanges[1].values ? response.valueRanges[1].values[0] : [];
      
      const reactionColumnName = COLUMN_HEADERS[reactionKey];
      const reactionColumnIndex = headers.indexOf(reactionColumnName);
      
      if (reactionColumnIndex === -1) {
        throw new Error('リアクション列が見つかりません: ' + reactionColumnName);
      }

      // 現在のリアクション文字列を解析
      const currentReactionString = rowData[reactionColumnIndex] || '';
      const currentReactions = this._parseReactionString(currentReactionString);
      
      // リアクションの追加/削除
      const userIndex = currentReactions.indexOf(userEmail);
      let action;
      
      if (userIndex >= 0) {
        currentReactions.splice(userIndex, 1);
        action = 'removed';
      } else {
        currentReactions.push(userEmail);
        action = 'added';
      }
      
      // 更新
      const updatedReactionString = currentReactions.join(', ');
      const cellRange = `${sheetName}!${String.fromCharCode(65 + reactionColumnIndex)}${rowIndex}`;
      
      await service.batchUpdate(spreadsheetId, [{
        range: cellRange,
        values: [[updatedReactionString]]
      }]);

      debugLog(`リアクション更新完了: ${userEmail} → ${reactionKey} (${action})`);
      
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
   * @private
   * @param {string} val - リアクション文字列
   * @returns {string[]} メールアドレスの配列
   */
  static _parseReactionString(val) {
    if (!val) return [];
    return val
      .toString()
      .split(',')
      .map(s => s.trim())
      .filter(Boolean);
  }
}