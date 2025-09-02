module.exports = {
  preset: 'ts-jest',
  testEnvironment: 'node',
  roots: ['<rootDir>/src', '<rootDir>/tests'],
  
  // Test file patterns - support both JS and TS
  testMatch: [
    '**/__tests__/**/*.(ts|tsx|js|jsx)',
    '**/?(*.)+(spec|test).(ts|tsx|js|jsx)'
  ],
  
  // File extensions to recognize
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'gs', 'json'],
  
  // Setup files - only GAS mocks needed
  setupFilesAfterEnv: [
    '<rootDir>/tests/mocks/gasMocks.js'
  ],
  
  // Coverage configuration - focused on core functionality
  collectCoverageFrom: [
    'src/**/*.{js,ts,gs}',
    '!src/**/*.html',
    '!src/appsscript.json',
    '!src/**/*.d.ts',
    '!src/utilities/**', // Exclude all utility scripts
    '!src/scripts/**'   // Exclude conversion scripts
  ],
  coverageDirectory: 'coverage',
  coverageReporters: ['text', 'lcov'],
  coverageThreshold: {
    global: {
      branches: 70,
      functions: 70,
      lines: 70,
      statements: 70
    }
  },
  
  // Test configuration
  verbose: true,
  testPathIgnorePatterns: ['/node_modules/', '/.git/', '/dist/'],
  
  // Transform configuration for TypeScript and GAS files
  transform: {
    '^.+\\.(ts|tsx)$': ['ts-jest', {
      useESM: false,
      tsconfig: {
        target: 'ES2020',
        module: 'CommonJS',
        esModuleInterop: true,
        allowSyntheticDefaultImports: true,
        skipLibCheck: true
      }
    }],
    '^.+\\.gs$': '<rootDir>/jest.gsTransform.js',
    '^.+\\.(js|jsx)$': ['ts-jest', {
      useESM: false
    }]
  },
  
  // Module name mapping for path aliases
  moduleNameMapper: {
    '^@src/(.*)$': '<rootDir>/src/$1',
    '^@tests/(.*)$': '<rootDir>/tests/$1',
    '^@mocks/(.*)$': '<rootDir>/tests/mocks/$1'
  },
  
  // Legacy globals removed - using modern transform configuration instead
  
  // Test timeout - reduced for focused unit tests
  testTimeout: 15000,
  
  // Clear mocks between tests
  clearMocks: true,
  restoreMocks: true,
  
  // Collect coverage from TypeScript files
  extensionsToTreatAsEsm: [],
  
  // Simplified reporters for faster execution
  reporters: ['default']
};