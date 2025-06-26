/**
 * StudyQuest 共通ユーティリティライブラリ
 * 全HTMLファイルで統一して使用する共通機能
 */

class StudyQuestUtils {
  static timers = new Set();
  static messageArea = null;

  /**
   * 統一されたメッセージ表示
   * @param {string} message - 表示するメッセージ
   * @param {string} type - メッセージタイプ (success, error, warning, info)
   * @param {number} duration - 表示時間（ミリ秒）
   */
  static showMessage(message, type = 'info', duration = 4000) {
    const messageArea = document.getElementById('message-area');
    if (!messageArea) return;

    const colors = {
      success: 'text-green-400',
      error: 'text-red-400',
      warning: 'text-yellow-400',
      info: 'text-blue-400'
    };

    const icons = {
      success: '<svg class="w-4 h-4 inline mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M9 12l2 2 4-4m6 2a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>',
      error: '<svg class="w-4 h-4 inline mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 8v4m0 4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>',
      warning: '<svg class="w-4 h-4 inline mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M12 9v2m0 4h.01m-6.938 4h13.856c1.54 0 2.502-1.667 1.732-2.5L13.732 4c-.77-.833-1.964-.833-2.732 0L3.732 16c-.77.833.192 2.5 1.732 2.5z"></path></svg>',
      info: '<svg class="w-4 h-4 inline mr-2" fill="none" stroke="currentColor" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" stroke-width="2" d="M13 16h-1v-4h-1m1-4h.01M21 12a9 9 0 11-18 0 9 9 0 0118 0z"></path></svg>'
    };

    messageArea.className = `text-center min-h-[24px] ${colors[type] || 'text-gray-400'} flex items-center justify-center`;
    messageArea.innerHTML = `${icons[type] || ''}<span>${message}</span>`;

    // アニメーション効果
    if (type === 'success') {
      messageArea.classList.add('success-pulse');
      const timer1 = setTimeout(() => messageArea.classList.remove('success-pulse'), 600);
      this.timers.add(timer1);
    } else if (type === 'error') {
      messageArea.classList.add('error-shake');
      const timer2 = setTimeout(() => messageArea.classList.remove('error-shake'), 500);
      this.timers.add(timer2);
    }

    if (duration > 0) {
      const timer3 = setTimeout(() => {
        if (messageArea) {
          messageArea.innerHTML = '';
          messageArea.className = 'text-center min-h-[24px]';
        }
      }, duration);
      this.timers.add(timer3);
    }
  }

  /**
   * 統一されたクリップボードコピー機能
   * @param {string} text - コピーするテキスト
   * @param {Element} button - フィードバック表示用のボタン要素
   * @returns {Promise<boolean>} コピー成功/失敗
   */
  static async copyToClipboard(text, button = null) {
    if (!text) return false;

    try {
      if (navigator.clipboard && window.isSecureContext) {
        await navigator.clipboard.writeText(text);
      } else {
        // フォールバック: 一時的なテキストエリアを作成
        const textArea = document.createElement('textarea');
        textArea.value = text;
        textArea.style.position = 'fixed';
        textArea.style.left = '-999999px';
        textArea.style.top = '-999999px';
        document.body.appendChild(textArea);
        textArea.focus();
        textArea.select();
        document.execCommand('copy');
        document.body.removeChild(textArea);
      }

      // 成功フィードバック
      if (button) {
        this.showButtonFeedback(button, 'コピー済み', 'success');
      } else {
        this.showMessage('クリップボードにコピーしました', 'success', 2000);
      }
      return true;
    } catch (error) {
      console.error('コピーに失敗:', error);
      if (button) {
        this.showButtonFeedback(button, 'コピー失敗', 'error');
      } else {
        this.showMessage('コピーに失敗しました', 'error', 3000);
      }
      return false;
    }
  }

  /**
   * ボタンへのフィードバック表示
   * @param {Element} button - 対象ボタン
   * @param {string} message - 表示メッセージ
   * @param {string} type - フィードバックタイプ
   */
  static showButtonFeedback(button, message, type = 'info') {
    if (!button) return;

    const originalText = button.textContent;
    const originalClasses = button.className;

    const typeClasses = {
      success: 'bg-green-600',
      error: 'bg-red-600',
      info: 'bg-blue-600'
    };

    button.textContent = message;
    button.classList.add(typeClasses[type] || typeClasses.info);

    const timer = setTimeout(() => {
      button.textContent = originalText;
      button.className = originalClasses;
    }, 2000);

    this.timers.add(timer);
  }

  /**
   * HTMLエスケープ
   * @param {string} unsafe - エスケープ対象文字列
   * @returns {string} エスケープ済み文字列
   */
  static escapeHtml(unsafe) {
    if (typeof unsafe !== 'string') return '';
    
    return unsafe
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;")
      .replace(/'/g, "&#039;");
  }

  /**
   * URL検証
   * @param {string} string - 検証対象URL
   * @returns {boolean} 有効なURLかどうか
   */
  static isValidUrl(string) {
    try {
      new URL(string);
      return true;
    } catch (_) {
      return false;
    }
  }

  /**
   * 要素の存在チェック
   * @param {string} selector - セレクター
   * @returns {Element|null} 要素またはnull
   */
  static safeQuerySelector(selector) {
    try {
      return document.querySelector(selector);
    } catch (error) {
      console.warn(`Invalid selector: ${selector}`, error);
      return null;
    }
  }

  /**
   * デバウンス処理
   * @param {Function} func - 実行する関数
   * @param {number} wait - 待機時間
   * @returns {Function} デバウンス済み関数
   */
  static debounce(func, wait) {
    let timeout;
    return function executedFunction(...args) {
      const later = () => {
        clearTimeout(timeout);
        func(...args);
      };
      clearTimeout(timeout);
      timeout = setTimeout(later, wait);
    };
  }

  /**
   * ローディングオーバーレイの作成
   * @param {string} message - ローディングメッセージ
   * @returns {Element} オーバーレイ要素
   */
  static createLoadingOverlay(message = '読み込み中...') {
    const overlay = document.createElement('div');
    overlay.className = 'loading-overlay';
    overlay.innerHTML = `
      <div class="glass-panel p-6 rounded-lg text-center">
        <div class="spinner mb-4"></div>
        <p class="text-white">${this.escapeHtml(message)}</p>
      </div>
    `;
    
    overlay.style.cssText = `
      position: fixed;
      top: 0;
      left: 0;
      right: 0;
      bottom: 0;
      background: rgba(26, 27, 38, 0.9);
      backdrop-filter: blur(4px);
      display: flex;
      align-items: center;
      justify-content: center;
      z-index: 1000;
    `;

    return overlay;
  }

  /**
   * ページ離脱時のクリーンアップ
   */
  static cleanup() {
    this.timers.forEach(timer => clearTimeout(timer));
    this.timers.clear();
  }

  /**
   * 初期化 - DOMContentLoaded時に呼び出し
   */
  static init() {
    // ページ離脱時のクリーンアップ
    window.addEventListener('beforeunload', () => this.cleanup());
    
    // エラーハンドリング
    window.addEventListener('error', (event) => {
      console.error('Global error:', event.error);
    });

    // 未処理のPromise拒否をキャッチ
    window.addEventListener('unhandledrejection', (event) => {
      console.error('Unhandled promise rejection:', event.reason);
      event.preventDefault(); // コンソールへの出力を防ぐ
    });

    console.log('StudyQuest Utils initialized');
  }
}

// 自動初期化
if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', () => StudyQuestUtils.init());
} else {
  StudyQuestUtils.init();
}

// グローバルに公開
window.StudyQuestUtils = StudyQuestUtils;