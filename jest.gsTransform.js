/**
 * Jest用GS(Google Apps Script)ファイル変換器
 * .gsファイルを.jsファイルとして扱うための簡易トランスフォーマー
 */
const fs = require('fs');

module.exports = {
  process(sourceText, sourcePath) {
    // GASファイルの先頭にGAS APIモックを自動的に読み込む
    const mockImport = `// Auto-imported GAS mocks\nconst { SpreadsheetApp, PropertiesService, CacheService, UrlFetchApp, HtmlService, Utilities, Logger, ScriptApp, DriveApp, Session, LockService } = require('./tests/mocks/gasMocks.js');\n\n`;
    
    // .gsファイルの内容をJavaScriptとして処理
    const transformedSource = mockImport + sourceText;
    
    return {
      code: transformedSource,
      map: null,
    };
  },
};