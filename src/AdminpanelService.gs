// ===========================================
// 📊 管理パネル用API関数（main.gsから移動）
// ===========================================

/**
 * スプレッドシート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @returns {Object} スプレッドシート一覧
 */

/**
 * シート一覧を取得
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @returns {Object} シート一覧
 */

/**
 * 列を分析
 * AdminPanel.js.html, AppSetupPage.html から呼び出される
 *
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @returns {Object} 列分析結果
 */

// ===========================================
// 🔧 ヘルパー関数（未定義関数の実装）
// ===========================================

/**
 * リアクション列の取得または作成
 * @param {Sheet} sheet - シートオブジェクト
 * @param {string} reactionType - リアクションタイプ
 * @returns {number} 列番号
 */

/**
 * ハイライト列の更新
 * @param {Object} config - 設定
 * @param {string} rowId - 行ID
 * @returns {boolean} 成功可否
 */

/**
 * リアクションパラメータ検証
 * @param {string} spreadsheetId - スプレッドシートID
 * @param {string} sheetName - シート名
 * @param {number} rowIndex - 行インデックス
 * @param {string} reactionKey - リアクション種類
 * @returns {boolean} 検証結果
 */
