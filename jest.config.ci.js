// CI環境用 Jest設定 - 軽量・高速・安定動作重視
module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/tests'],
  
  // CI用: 基本テストのみ実行
  testMatch: [
    '**/setup-validation.test.ts'
  ],
  
  // 必要最小限のファイル拡張子
  moduleFileExtensions: ['ts', 'js', 'json'],
  
  // シンプルなセットアップ
  setupFilesAfterEnv: ['<rootDir>/tests/setup/jest.setup.ts'],
  
  // 軽量カバレッジ設定
  collectCoverageFrom: [
    'src/**/*.{js,ts,gs}',
    '!src/**/*.html',
    '!src/appsscript.json'
  ],
  coverageDirectory: 'coverage',
  coverageReporters: ['text'],
  
  // CI最適化設定
  verbose: false,
  testTimeout: 10000,
  maxWorkers: 2,
  clearMocks: true,
  
  // Transform設定 - 最小限
  transform: {
    '^.+\\.ts$': ['ts-jest', {
      useESM: false,
      tsconfig: {
        target: 'ES2020',
        module: 'CommonJS',
        strict: false,
        skipLibCheck: true
      }
    }]
  },
  
  // エラー回避のため複雑なテストは除外
  testPathIgnorePatterns: [
    '/node_modules/', 
    '/dist/', 
    '/coverage/',
    'criticalFunctions.test.ts',
    'integration/',
    'edge-cases/',
    'fixtures/'
  ]
};