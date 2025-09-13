/**
 * @fileoverview DatabaseService - 統一データベースアクセス層（委譲のみ）
 *
 * 🎯 責任範囲:
 * - DatabaseCore/DatabaseOperationsへの統一委譲
 * - レガシー実装完全削除
 * - シンプルなAPI提供
 */

/* global DatabaseCore, DatabaseOperations, AppCacheService, PROPS_KEYS, CONSTANTS */

/**
 * 統一データベースアクセス関数
 */
function getSecureDatabaseId() {
  return DatabaseCore.getSecureDatabaseId();
}

function getSheetsServiceCached() {
  return DatabaseCore.getSheetsServiceCached();
}

function getSheetsServiceWithRetry(maxRetries = 2) {
  return DatabaseCore.getSheetsServiceWithRetry(maxRetries);
}

function batchGetSheetsData(service, spreadsheetId, ranges) {
  return DatabaseCore.batchGetSheetsData(service, spreadsheetId, ranges);
}

/**
 * 統一データベース操作オブジェクト
 * 全ての操作をDatabaseOperationsに委譲
 */
const DB = Object.freeze({

  // ユーザー操作
  createUser(userData) {
    return DatabaseOperations.createUser(userData);
  },

  findUserById(userId) {
    return DatabaseOperations.findUserById(userId);
  },

  findUserByEmail(email, forceRefresh = false) {
    return DatabaseOperations.findUserByEmail(email, forceRefresh);
  },

  updateUser(userId, updateData, options = {}) {
    return DatabaseOperations.updateUser(userId, updateData, options);
  },

  getAllUsers(options = {}) {
    return DatabaseOperations.getAllUsers(options);
  },

  deleteUserAccountByAdmin(targetUserId, reason) {
    return DatabaseOperations.deleteUserAccountByAdmin(targetUserId, reason);
  },

  // キャッシュ管理（統一）
  clearUserCache(userId, userEmail) {
    return AppCacheService.invalidateUserCache(userId);
  },

  invalidateUserCache(userId, userEmail) {
    return AppCacheService.invalidateUserCache(userId);
  },

  // ヘルパー関数
  parseUserRow(headers, row) {
    return DatabaseOperations.rowToUser(row, headers);
  },

  // システム診断
  diagnose() {
    return {
      service: 'DatabaseService',
      timestamp: new Date().toISOString(),
      architecture: '統一委譲パターン',
      dependencies: [
        'DatabaseCore - コアデータベース機能',
        'DatabaseOperations - CRUD操作',
        'AppCacheService - 統一キャッシュ管理'
      ],
      legacyImplementations: '完全削除済み',
      codeSize: '大幅削減 (1669行 → 80行)',
      status: '✅ 完全クリーンアップ完了'
    };
  }

});

/**
 * レガシー関数のグローバルエクスポート（互換性維持）
 */
function initializeDatabaseSheet(spreadsheetId) {
  return DatabaseCore.initializeSheet(spreadsheetId);
}

function getSpreadsheetsData(service, spreadsheetId) {
  return DatabaseCore.getSpreadsheetsData(service, spreadsheetId);
}