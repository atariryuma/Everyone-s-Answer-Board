/**
 * @fileoverview URL管理クラス
 */

class UrlManager {
  static URL_CACHE_KEY = 'WEB_APP_URL';
  static URL_CACHE_TTL = 3600; // 1時間
  
  /**
   * WebアプリのURLを取得（キャッシュ利用）
   * @returns {string} WebアプリURL
   */
  static getWebAppUrl() {
    return CacheManager.get(this.URL_CACHE_KEY, () => {
      try {
        const url = ScriptApp.getService().getUrl();
        if (!url) {
          console.warn('ScriptApp.getService().getUrl()がnullを返しました');
          return this._getFallbackUrl();
        }
        
        // URLの正規化
        return url.endsWith('/') ? url.slice(0, -1) : url;
      } catch (e) {
        console.error('WebアプリURL取得エラー: ' + e.message);
        return this._getFallbackUrl();
      }
    }, 'medium');
  }
  
  /**
   * フォールバックURL生成
   * @private
   * @returns {string} フォールバックURL
   */
  static _getFallbackUrl() {
    try {
      const scriptId = ScriptApp.getScriptId();
      if (scriptId) {
        return `https://script.google.com/macros/s/${scriptId}/exec`;
      }
    } catch (e) {
      console.error('フォールバックURL生成エラー: ' + e.message);
    }
    return '';
  }
  
  /**
   * アプリケーション用のURL群を生成
   * @param {string} userId - ユーザーID
   * @returns {object} URL群
   */
  static generateAppUrls(userId) {
    try {
      const webAppUrl = this.getWebAppUrl();
      
      if (!webAppUrl) {
        return {
          webAppUrl: '',
          adminUrl: '',
          viewUrl: '',
          setupUrl: '',
          status: 'error',
          message: 'WebアプリURLが取得できませんでした'
        };
      }
      
      return {
        webAppUrl: webAppUrl,
        adminUrl: `${webAppUrl}?userId=${userId}&mode=admin`,
        viewUrl: `${webAppUrl}?userId=${userId}`,
        setupUrl: `${webAppUrl}?setup=true`,
        status: 'success'
      };
    } catch (e) {
      console.error('URL生成エラー: ' + e.message);
      return {
        webAppUrl: '',
        adminUrl: '',
        viewUrl: '',
        setupUrl: '',
        status: 'error',
        message: 'URLの生成に失敗しました: ' + e.message
      };
    }
  }
  
  /**
   * URLキャッシュをクリア
   */
  static clearUrlCache() {
    CacheManager.remove(this.URL_CACHE_KEY);
  }
}