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
  
  // Setup files - enhanced with jest-extended
  setupFilesAfterEnv: [
    '<rootDir>/tests/setup/jest.setup.ts',
    '<rootDir>/tests/mocks/gasMocks.ts'
  ],
  
  // Coverage configuration
  collectCoverageFrom: [
    'src/**/*.{js,ts,gs}',
    '!src/**/*.html',
    '!src/appsscript.json',
    '!src/**/*.d.ts',
    '!src/utilities/RenamingScript.gs' // Exclude utility scripts
  ],
  coverageDirectory: 'coverage',
  coverageReporters: ['text', 'lcov', 'html', 'json'],
  coverageThreshold: {
    global: {
      branches: 80,
      functions: 80,
      lines: 80,
      statements: 80
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
  
  // Test timeout for longer integration tests
  testTimeout: 30000,
  
  // Clear mocks between tests
  clearMocks: true,
  restoreMocks: true,
  
  // Collect coverage from TypeScript files
  extensionsToTreatAsEsm: [],
  
  // Test result processor for better output
  reporters: [
    'default',
    ['jest-junit', {
      outputDirectory: 'coverage',
      outputName: 'junit.xml'
    }]
  ]
};