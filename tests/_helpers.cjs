// 共通テストスタブ。src/helpers.js / src/main.js のレスポンス生成 helper を
// vm.createContext() に注入するための同形 stub をまとめる。各 test ファイルが
// 個別に同じ stub を書くと shape が drift して assertion がブレるので、
// ここで「real impl と同じ shape」を一箇所に固定する。

function gasResponseStubs() {
  const createErrorResponse = (message, data = null, extraFields = {}) => ({
    success: false,
    message,
    error: message,
    ...(data !== null && data !== undefined ? { data } : {}),
    ...extraFields
  });
  const createSuccessResponse = (message, data = null, extraFields = {}) => ({
    success: true,
    message,
    ...(data !== null && data !== undefined ? { data } : {}),
    ...extraFields
  });
  return {
    createErrorResponse,
    createSuccessResponse,
    createAuthError: () => createErrorResponse('ユーザー認証が必要です'),
    createUserNotFoundError: () => createErrorResponse('ユーザーが見つかりません'),
    createAdminRequiredError: () => createErrorResponse('管理者権限が必要です'),
    createExceptionResponse: (error) => createErrorResponse(error && error.message ? error.message : 'Unknown error'),
    createDataServiceErrorResponse: (message, sheetName = '') => createErrorResponse(message, [], { headers: [], sheetName }),
    isPlainObject: (v) => Boolean(v) && typeof v === 'object' && !Array.isArray(v),
    // sameEmail_: src/DataApis.js の case-insensitive email 比較。owner 判定で大小差により
    //   所有者を viewer と誤判定するのを防ぐ。他ファイル (ReactionService/DataService/DatabaseCore)
    //   が global 参照するため stub を共有注入する。
    sameEmail_: (a, b) => String(a || '').toLowerCase().trim() === String(b || '').toLowerCase().trim(),
    logError_: () => {}, // src/helpers.js の共通エラーログ。test では silent。
    // safeJsonParse_: src/helpers.js の本物実装と同じ contract (null/空/object pass-through、 catch で fallback)。
    safeJsonParse_: (text, fallback) => {
      const fb = (fallback === undefined ? null : fallback);
      if (text == null || text === '') return fb;
      if (typeof text === 'object') return text;
      try { return JSON.parse(text); } catch (_) { return fb; }
    }
  };
}

// GAS ContentService の最小 stub。doPost / doGet が createTextOutput().setMimeType()
//   をチェーンするので両 method を提供。本物の MimeType 列挙は JSON のみで十分。
function createContentServiceStub() {
  return {
    MimeType: { JSON: 'application/json' },
    createTextOutput(body) {
      return {
        body,
        mimeType: null,
        setMimeType(mimeType) {
          this.mimeType = mimeType;
          return this;
        }
      };
    }
  };
}

module.exports = { gasResponseStubs, createContentServiceStub };
