/**
 * @fileoverview データベース管理クラス - バッチ操作とキャッシュ最適化
 */

class DatabaseManager {
  static USER_CACHE_TTL = 300; // 5分
  static BATCH_SIZE = 100;
  
  /**
   * 最適化されたSheetsサービスを取得
   * @returns {object} Sheets APIサービス
   */
  static getSheetsService() {
    const accessToken = AuthManager.getServiceAccountToken();
    return new OptimizedSheetsService(accessToken);
  }
  
  /**
   * ユーザー情報を効率的に検索（キャッシュ優先）
   * @param {string} userId - ユーザーID
   * @returns {object|null} ユーザー情報
   */
  static async findUserById(userId) {
    const cacheKey = `user_${userId}`;
    const cache = CacheService.getScriptCache();
    
    // キャッシュから取得を試行
    let cachedUser = cache.get(cacheKey);
    if (cachedUser) {
      debugLog(`ユーザーキャッシュヒット: ${userId}`);
      return JSON.parse(cachedUser);
    }
    
    // データベースから取得
    const user = await this._fetchUserFromDatabase('userId', userId);
    
    if (user) {
      // キャッシュに保存
      cache.put(cacheKey, JSON.stringify(user), this.USER_CACHE_TTL);
      cache.put(`email_${user.adminEmail}`, JSON.stringify(user), this.USER_CACHE_TTL);
    }
    
    return user;
  }
  
  /**
   * メールアドレスでユーザー検索
   * @param {string} email - メールアドレス
   * @returns {object|null} ユーザー情報
   */
  static async findUserByEmail(email) {
    const cacheKey = `email_${email}`;
    const cache = CacheService.getScriptCache();
    
    // キャッシュから取得を試行
    let cachedUser = cache.get(cacheKey);
    if (cachedUser) {
      debugLog(`メールキャッシュヒット: ${email}`);
      return JSON.parse(cachedUser);
    }
    
    // データベースから取得
    const user = await this._fetchUserFromDatabase('adminEmail', email);
    
    if (user) {
      // 両方のキーでキャッシュ
      cache.put(cacheKey, JSON.stringify(user), this.USER_CACHE_TTL);
      cache.put(`user_${user.userId}`, JSON.stringify(user), this.USER_CACHE_TTL);
    }
    
    return user;
  }
  
  /**
   * データベースからユーザーを取得
   * @private
   * @param {string} field - 検索フィールド
   * @param {string} value - 検索値
   * @returns {object|null} ユーザー情報
   */
  static async _fetchUserFromDatabase(field, value) {
    try {
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      const service = this.getSheetsService();
      const sheetName = DB_SHEET_CONFIG.SHEET_NAME;
      
      const data = await service.batchGet(dbId, [`${sheetName}!A:H`]);
      const values = data.valueRanges[0].values || [];
      
      if (values.length === 0) return null;
      
      const headers = values[0];
      const fieldIndex = headers.indexOf(field);
      
      if (fieldIndex === -1) return null;
      
      for (let i = 1; i < values.length; i++) {
        if (values[i][fieldIndex] === value) {
          const user = {};
          headers.forEach((header, index) => {
            user[header] = values[i][index] || '';
          });
          return user;
        }
      }
      
      return null;
    } catch (error) {
      console.error(`ユーザー検索エラー (${field}:${value}):`, error);
      return null;
    }
  }
  
  /**
   * ユーザー情報を一括更新
   * @param {string} userId - ユーザーID
   * @param {object} updateData - 更新データ
   * @returns {object} 更新結果
   */
  static async updateUser(userId, updateData) {
    try {
      const props = PropertiesService.getScriptProperties();
      const dbId = props.getProperty(SCRIPT_PROPS_KEYS.DATABASE_SPREADSHEET_ID);
      const service = this.getSheetsService();
      const sheetName = DB_SHEET_CONFIG.SHEET_NAME;
      
      // 現在のデータを取得
      const data = await service.batchGet(dbId, [`${sheetName}!A:H`]);
      const values = data.valueRanges[0].values || [];
      
      if (values.length === 0) {
        throw new Error('データベースが空です');
      }
      
      const headers = values[0];
      const userIdIndex = headers.indexOf('userId');
      let rowIndex = -1;
      
      // ユーザーの行を特定
      for (let i = 1; i < values.length; i++) {
        if (values[i][userIdIndex] === userId) {
          rowIndex = i + 1; // 1-based index
          break;
        }
      }
      
      if (rowIndex === -1) {
        throw new Error('更新対象のユーザーが見つかりません');
      }
      
      // バッチ更新リクエストを作成
      const requests = Object.keys(updateData).map(key => {
        const colIndex = headers.indexOf(key);
        if (colIndex === -1) return null;
        
        return {
          range: `${sheetName}!${String.fromCharCode(65 + colIndex)}${rowIndex}`,
          values: [[updateData[key]]]
        };
      }).filter(Boolean);
      
      if (requests.length > 0) {
        await service.batchUpdate(dbId, requests);
      }
      
      // キャッシュを無効化
      const cache = CacheService.getScriptCache();
      cache.remove(`user_${userId}`);
      if (updateData.adminEmail) {
        cache.remove(`email_${updateData.adminEmail}`);
      }
      
      return { success: true };
    } catch (error) {
      console.error('ユーザー更新エラー:', error);
      throw error;
    }
  }
  
  /**
   * 全キャッシュをクリア
   */
  static clearAllCache() {
    const cache = CacheService.getScriptCache();
    // ユーザー関連のキャッシュをクリア
    cache.removeAll(['USER_INFO_CACHE', 'HEADER_CACHE', 'ROSTER_CACHE']);
    debugLog('データベースキャッシュをクリアしました');
  }
}

/**
 * 最適化されたSheetsサービス
 */
class OptimizedSheetsService {
  constructor(accessToken) {
    this.accessToken = accessToken;
    this.baseUrl = 'https://sheets.googleapis.com/v4/spreadsheets';
  }
  
  /**
   * バッチ取得
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {string[]} ranges - 取得範囲の配列
   * @returns {object} レスポンス
   */
  async batchGet(spreadsheetId, ranges) {
    const url = `${this.baseUrl}/${spreadsheetId}/values:batchGet?${ranges.map(range => `ranges=${encodeURIComponent(range)}`).join('&')}`;
    
    const response = UrlFetchApp.fetch(url, {
      headers: { 'Authorization': `Bearer ${this.accessToken}` }
    });
    
    return JSON.parse(response.getContentText());
  }
  
  /**
   * バッチ更新
   * @param {string} spreadsheetId - スプレッドシートID
   * @param {object[]} requests - 更新リクエストの配列
   * @returns {object} レスポンス
   */
  async batchUpdate(spreadsheetId, requests) {
    const url = `${this.baseUrl}/${spreadsheetId}/values:batchUpdate`;
    
    const response = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': `Bearer ${this.accessToken}` },
      payload: JSON.stringify({
        data: requests,
        valueInputOption: 'RAW'
      })
    });
    
    return JSON.parse(response.getContentText());
  }
}