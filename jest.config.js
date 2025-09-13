/**
 * Jest設定ファイル - CLAUDE.md準拠版（簡素化）
 * 🚀 configJSON中心型システムのテスト環境
 */

module.exports = {
  // テスト環境設定
  testEnvironment: 'node',
  
  // テストファイルパターン
  testMatch: [
    '<rootDir>/tests/**/*.test.js',
    '<rootDir>/tests/**/*.spec.js'
  ],
  
  // モジュール解決
  moduleNameMapper: {
    '^@/(.*)$': '<rootDir>/src/$1'
  },
  
  // カバレッジ設定
  collectCoverage: true,
  collectCoverageFrom: [
    'src/**/*.{js,gs}',
    '!src/**/*.html',
    '!node_modules/**'
  ],
  coverageReporters: [
    'text',
    'lcov',
    'html'
  ],
  coverageDirectory: 'coverage',
  
  // カバレッジ閾値（CLAUDE.md品質基準）
  coverageThreshold: {
    global: {
      branches: 70,
      functions: 70,
      lines: 70,
      statements: 70
    }
  },
  
  // テスト実行設定
  verbose: true,
  passWithNoTests: true,
  
  // タイムアウト設定
  testTimeout: 10000,
  
  // エラーハンドリング
  errorOnDeprecated: false,
  
  // モック設定
  clearMocks: true,
  restoreMocks: true,
  resetMocks: true,
  
  // 無視パターン
  testPathIgnorePatterns: [
    '/node_modules/',
    '/coverage/',
    '/dist/'
  ],
  
  // モジュール無視パターン
  modulePathIgnorePatterns: [
    '/coverage/',
    '/dist/'
  ]
};