/**
 * Jest最小設定 - clearTimeoutエラー回避版
 */

module.exports = {
  testEnvironment: 'node',
  testMatch: [
    '<rootDir>/tests/**/*.test.js'
  ],
  verbose: true,
  passWithNoTests: true,
  testTimeout: 10000,
  clearMocks: true,
  // fakeTimersを完全に無効化
  timers: 'real'
};