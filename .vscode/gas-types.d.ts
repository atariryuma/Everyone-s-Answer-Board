/**
 * Google Apps Script 型定義
 * VS Code用 IntelliSense サポート
 */

// GAS グローバルオブジェクト
declare const SpreadsheetApp: any;
declare const PropertiesService: any;
declare const CacheService: any;
declare const Session: any;
declare const Utilities: any;
declare const ScriptApp: any;
declare const FormApp: any;
declare const HtmlService: any;
declare const ContentService: any;
declare const UrlFetchApp: any;
declare const DriveApp: any;
declare const GmailApp: any;
declare const CalendarApp: any;
declare const DocumentApp: any;
declare const Logger: any;
declare const Browser: any;

// プロジェクト固有のグローバル関数
declare function getCurrentUserEmail(): string;
declare function findUserById(userId: string): any;
declare function checkLoginStatus(userId: string): { isValid: boolean; message: string; userId: string };
declare function updateLoginStatus(userId: string, status: string): void;

// ログ関数
declare function logError(message: string, data?: any, category?: string): void;
declare function debugLog(message: string, data?: any, category?: string): void;
declare function warnLog(message: string, data?: any, category?: string): void;
declare function infoLog(message: string, data?: any, category?: string): void;

// 統一ユーティリティ
declare const ULog: any;
declare const UError: any;
declare const UValidate: any;
declare const UExecute: any;
declare const UData: any;
declare const UnifiedValidation: any;

// 回復力のある実行
declare const resilientExecutor: any;
declare function resilientUrlFetch(url: string, options?: any): any;
declare function resilientSpreadsheetOperation(operation: Function, operationName?: string): any;
declare function resilientCacheOperation(operation: Function, operationName?: string, fallback?: Function): any;

// ポリフィル
declare function clearTimeout(id: any): void;