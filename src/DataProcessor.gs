/**
 * @fileoverview データ処理とソート機能
 */

class DataProcessor {
  /**
   * 指定されたシートのデータを取得し、フィルタリング・ソートを適用
   * @param {string} userId - ユーザーID
   * @param {string} sheetName - シート名
   * @param {string} classFilter - クラスフィルター
   * @param {string} sortMode - ソートモード
   * @returns {object} 処理されたデータ
   */
  static async getSheetData(userId, sheetName, classFilter, sortMode) {
    try {
      const userInfo = await DatabaseManager.findUserById(userId);
      if (!userInfo) {
        throw new Error('ユーザー情報が見つかりません');
      }
      
      const spreadsheetId = userInfo.spreadsheetId;
      const service = DatabaseManager.getSheetsService();
      
      // バッチでデータ、ヘッダー、名簿を取得
      const ranges = [
        `${sheetName}!A:Z`,
        `${ROSTER_CONFIG.SHEET_NAME}!A:Z`
      ];
      
      const responses = await service.batchGet(spreadsheetId, ranges);
      const sheetData = responses.valueRanges[0].values || [];
      const rosterData = responses.valueRanges[1].values || [];
      
      if (sheetData.length === 0) {
        return {
          status: 'success',
          data: [],
          headers: [],
          totalCount: 0
        };
      }
      
      const headers = sheetData[0];
      const dataRows = sheetData.slice(1);
      
      // ヘッダーインデックスを取得（キャッシュ利用）
      const headerIndices = await CacheManager.getHeaderIndices(spreadsheetId, sheetName);
      
      // 名簿マップを作成（キャッシュ利用）
      const rosterMap = this._buildRosterMap(rosterData);
      
      // 表示モードを取得
      const configJson = JSON.parse(userInfo.configJson || '{}');
      const displayMode = configJson.displayMode || DISPLAY_MODES.ANONYMOUS;
      
      // データを並列処理で高速化
      const processedData = dataRows.map((row, index) => {
        return this._processRowData(row, headers, headerIndices, rosterMap, displayMode, index + 2);
      });
      
      // フィルタリング
      let filteredData = processedData;
      if (classFilter && classFilter !== 'all') {
        const classIndex = headerIndices[COLUMN_HEADERS.CLASS];
        if (classIndex !== undefined) {
          filteredData = processedData.filter(row => 
            row.originalData[classIndex] === classFilter
          );
        }
      }
      
      // ソート適用
      const sortedData = this._applySortMode(filteredData, sortMode || 'newest');
      
      return {
        status: 'success',
        data: sortedData,
        headers: headers,
        totalCount: sortedData.length,
        displayMode: displayMode
      };
      
    } catch (e) {
      console.error('シートデータ取得エラー: ' + e.message);
      return {
        status: 'error',
        message: 'データの取得に失敗しました: ' + e.message,
        data: [],
        headers: []
      };
    }
  }
  
  /**
   * 名簿マップを構築
   * @private
   * @param {array} rosterData - 名簿データ
   * @returns {object} 名簿マップ
   */
  static _buildRosterMap(rosterData) {
    if (rosterData.length === 0) return {};
    
    const headers = rosterData[0];
    const emailIndex = headers.indexOf(ROSTER_CONFIG.EMAIL_COLUMN);
    const nameIndex = headers.indexOf(ROSTER_CONFIG.NAME_COLUMN);
    const classIndex = headers.indexOf(ROSTER_CONFIG.CLASS_COLUMN);
    
    const rosterMap = {};
    
    for (let i = 1; i < rosterData.length; i++) {
      const row = rosterData[i];
      if (emailIndex !== -1 && row[emailIndex]) {
        rosterMap[row[emailIndex]] = {
          name: nameIndex !== -1 ? row[nameIndex] : '',
          class: classIndex !== -1 ? row[classIndex] : ''
        };
      }
    }
    
    return rosterMap;
  }
  
  /**
   * 行データを処理（スコア計算、名前変換など）
   * @private
   */
  static _processRowData(row, headers, headerIndices, rosterMap, displayMode, rowNumber) {
    const processedRow = {
      rowNumber: rowNumber,
      originalData: row,
      score: 0,
      likeCount: 0,
      understandCount: 0,
      curiousCount: 0,
      isHighlighted: false
    };
    
    // リアクションカウント計算（並列処理）
    REACTION_KEYS.forEach(reactionKey => {
      const columnName = COLUMN_HEADERS[reactionKey];
      const columnIndex = headerIndices[columnName];
      
      if (columnIndex !== undefined && row[columnIndex]) {
        const reactions = this._parseReactionString(row[columnIndex]);
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
    const highlightIndex = headerIndices[COLUMN_HEADERS.HIGHLIGHT];
    if (highlightIndex !== undefined && row[highlightIndex]) {
      processedRow.isHighlighted = row[highlightIndex].toString().toLowerCase() === 'true';
    }
    
    // スコア計算
    processedRow.score = this._calculateRowScore(processedRow);
    
    // 名前表示処理
    const emailIndex = headerIndices[COLUMN_HEADERS.EMAIL];
    if (emailIndex !== undefined && row[emailIndex] && displayMode === DISPLAY_MODES.NAMED) {
      const email = row[emailIndex];
      const rosterInfo = rosterMap[email];
      if (rosterInfo && rosterInfo.name) {
        processedRow.displayName = rosterInfo.name;
      }
    }
    
    return processedRow;
  }
  
  /**
   * 行のスコアを計算
   * @private
   */
  static _calculateRowScore(rowData) {
    const baseScore = 1.0;
    const likeBonus = rowData.likeCount * SCORING_CONFIG.LIKE_MULTIPLIER_FACTOR;
    const reactionBonus = (rowData.understandCount + rowData.curiousCount) * 0.01;
    const highlightBonus = rowData.isHighlighted ? 0.5 : 0;
    const randomFactor = Math.random() * SCORING_CONFIG.RANDOM_SCORE_FACTOR;
    
    return baseScore + likeBonus + reactionBonus + highlightBonus + randomFactor;
  }
  
  /**
   * データにソートを適用
   * @private
   */
  static _applySortMode(data, sortMode) {
    switch (sortMode) {
      case 'score':
        return data.sort((a, b) => b.score - a.score);
      case 'newest':
        return data.reverse();
      case 'oldest':
        return data;
      case 'random':
        return this._shuffleArray([...data]);
      case 'likes':
        return data.sort((a, b) => b.likeCount - a.likeCount);
      default:
        return data;
    }
  }
  
  /**
   * 配列をシャッフル（Fisher-Yates shuffle）
   * @private
   */
  static _shuffleArray(array) {
    for (let i = array.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [array[i], array[j]] = [array[j], array[i]];
    }
    return array;
  }
  
  /**
   * リアクション文字列をパース
   * @private
   */
  static _parseReactionString(val) {
    if (!val) return [];
    return val.toString().split(',').map(s => s.trim()).filter(Boolean);
  }
}