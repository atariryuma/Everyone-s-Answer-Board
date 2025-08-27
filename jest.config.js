module.exports = {
  testEnvironment: 'node',
  roots: ['<rootDir>/src', '<rootDir>/tests'],
  testMatch: ['**/__tests__/**/*.js', '**/?(*.)+(spec|test).js'],
  moduleFileExtensions: ['js', 'gs', 'json'],
  setupFilesAfterEnv: ['<rootDir>/tests/mocks/gasMocks.js'],
  collectCoverageFrom: [
    'src/**/*.{js,gs}',
    '!src/**/*.html',
    '!src/appsscript.json',
  ],
  coverageDirectory: 'coverage',
  coverageReporters: ['text', 'lcov', 'html'],
  verbose: true,
  testPathIgnorePatterns: ['/node_modules/', '/.git/'],
  transform: {
    '^.+\\.gs$': '<rootDir>/jest.gsTransform.js',
  },
};