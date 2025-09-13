#!/usr/bin/env node
/**
 * Architecture Testing Script
 * Validates system architecture against README.md standards
 */

const fs = require('fs');
const path = require('path');

// Colors for output
const colors = {
  green: '\x1b[32m',
  red: '\x1b[31m',
  yellow: '\x1b[33m',
  blue: '\x1b[34m',
  reset: '\x1b[0m'
};

function log(color, message) {
  console.log(`${colors[color]}${message}${colors.reset}`);
}

function logSection(title) {
  console.log('\n' + '='.repeat(50));
  log('blue', title);
  console.log('='.repeat(50));
}

function logTest(name, passed, details = '') {
  const status = passed ? '✅ PASS' : '❌ FAIL';
  console.log(`${status} ${name}`);
  if (details) {
    console.log(`    ${details}`);
  }
}

class ArchitectureTester {
  constructor() {
    this.srcPath = path.join(__dirname, '../src');
    this.testsPath = path.join(__dirname, '../tests');
    this.results = {
      passed: 0,
      failed: 0,
      warnings: 0
    };
  }

  // Test 1: Service Layer Architecture
  testServiceLayer() {
    logSection('🏗️ Service Layer Architecture');
    
    const requiredServices = [
      'UserService.gs',
      'ConfigService.gs',
      'DataService.gs',
      'SecurityService.gs'
    ];

    for (const service of requiredServices) {
      const servicePath = path.join(this.srcPath, 'services', service);
      const exists = fs.existsSync(servicePath);
      
      if (exists) {
        const content = fs.readFileSync(servicePath, 'utf8');
        const hasObjectFreeze = content.includes('Object.freeze');
        const hasJSDoc = content.includes('/**');
        const hasDiagnose = content.includes('diagnose');
        
        logTest(service, true, `Found with Object.freeze: ${hasObjectFreeze}, JSDoc: ${hasJSDoc}, Diagnose: ${hasDiagnose}`);
        this.results.passed++;
        
        if (!hasObjectFreeze) {
          log('yellow', `    ⚠️ Warning: ${service} should use Object.freeze`);
          this.results.warnings++;
        }
      } else {
        logTest(service, false, 'Service file not found');
        this.results.failed++;
      }
    }
  }

  // Test 2: Utility Layer Architecture
  testUtilityLayer() {
    logSection('🔧 Utility Layer Architecture');
    
    const requiredUtils = [
      'helpers.gs',
      'validators.gs',
      'formatters.gs'
    ];

    for (const util of requiredUtils) {
      const utilPath = path.join(this.srcPath, 'utils', util);
      const exists = fs.existsSync(utilPath);
      
      logTest(`utils/${util}`, exists, exists ? 'Utility module found' : 'Missing utility module');
      
      if (exists) {
        this.results.passed++;
      } else {
        this.results.failed++;
      }
    }
  }

  // Test 3: Core Error Handling
  testErrorHandling() {
    logSection('⚠️ Error Handling Architecture');
    
    const errorHandlerPath = path.join(this.srcPath, 'core', 'errors.gs');
    const exists = fs.existsSync(errorHandlerPath);
    
    logTest('core/errors.gs', exists, exists ? 'Unified error handling found' : 'Missing error handler');
    
    if (exists) {
      const content = fs.readFileSync(errorHandlerPath, 'utf8');
      const hasErrorHandler = content.includes('ErrorHandler');
      const hasAutoRecovery = content.includes('autoRecovery') || content.includes('retry');
      
      logTest('ErrorHandler implementation', hasErrorHandler, hasErrorHandler ? 'ErrorHandler class found' : 'Missing ErrorHandler class');
      logTest('Auto-recovery mechanism', hasAutoRecovery, hasAutoRecovery ? 'Auto-recovery features found' : 'No auto-recovery detected');
      
      this.results.passed += exists + hasErrorHandler + hasAutoRecovery;
    } else {
      this.results.failed++;
    }
  }

  // Test 4: Constants and Configuration
  testConstants() {
    logSection('📋 Constants and Configuration');
    
    const constantsPath = path.join(this.srcPath, 'core', 'constants.gs');
    const exists = fs.existsSync(constantsPath);

    logTest('core/constants.gs', exists, exists ? 'Constants file found' : 'Missing constants file');
    
    if (exists) {
      const content = fs.readFileSync(constantsPath, 'utf8');
      const hasCONSTANTS = content.includes('const CONSTANTS');
      const hasDatabase = content.includes('DATABASE');
      const hasReactions = content.includes('REACTIONS');
      const hasColumns = content.includes('COLUMNS');
      
      logTest('CONSTANTS object', hasCONSTANTS);
      logTest('DATABASE constants', hasDatabase);
      logTest('REACTIONS constants', hasReactions);
      logTest('COLUMNS constants', hasColumns);
      
      this.results.passed += exists + hasCONSTANTS + hasDatabase + hasReactions + hasColumns;
    } else {
      this.results.failed++;
    }
  }

  // Test 5: Main Entry Point
  testMainEntryPoint() {
    logSection('🚀 Main Entry Point');
    
    const mainPath = path.join(this.srcPath, 'main.gs');
    const exists = fs.existsSync(mainPath);
    
    logTest('main.gs', exists, exists ? 'Main entry point found' : 'Missing main.gs');
    
    if (exists) {
      const content = fs.readFileSync(mainPath, 'utf8');
      const hasDoGet = content.includes('function doGet');
      const hasDoPost = content.includes('function doPost');
      const usesServices = content.includes('UserService') && content.includes('ConfigService');
      
      logTest('doGet function', hasDoGet);
      logTest('doPost function', hasDoPost);
      logTest('Uses service layer', usesServices);
      
      this.results.passed += exists + hasDoGet + hasDoPost + usesServices;
    } else {
      this.results.failed++;
    }
  }

  // Test 6: Test Coverage
  testCoverage() {
    logSection('🧪 Test Coverage');
    
    const testServices = [
      'UserService.test.js',
      'ConfigService.test.js',
      'DataService.test.js',
      'SecurityService.test.js'
    ];

    for (const testFile of testServices) {
      const testPath = path.join(this.testsPath, 'services', testFile);
      const exists = fs.existsSync(testPath);
      
      if (exists) {
        const content = fs.readFileSync(testPath, 'utf8');
        const hasDescribe = content.includes('describe(');
        const hasTests = content.includes('it(') || content.includes('test(');
        const hasIntegrationTests = content.includes('Integration Tests');
        
        logTest(testFile, true, `Describe: ${hasDescribe}, Tests: ${hasTests}, Integration: ${hasIntegrationTests}`);
        this.results.passed++;
      } else {
        logTest(testFile, false, 'Test file not found');
        this.results.failed++;
      }
    }

    // Check integration tests
    const integrationTestPath = path.join(this.testsPath, 'integration', 'system-integration.test.js');
    const hasIntegrationTests = fs.existsSync(integrationTestPath);
    logTest('Integration tests', hasIntegrationTests);
    
    if (hasIntegrationTests) {
      this.results.passed++;
    } else {
      this.results.failed++;
    }
  }

  // Test 7: File Structure Compliance
  testFileStructure() {
    logSection('📁 File Structure Compliance');
    
    const expectedStructure = {
      'src/services': 'Service layer',
      'src/utils': 'Utility layer', 
      'src/core': 'Core functionality',
      'tests/services': 'Service tests',
      'tests/integration': 'Integration tests',
      'tests/mocks': 'Test mocks',
      '.claude/commands': 'Claude commands',
      '.claude/hooks': 'Git hooks'
    };

    for (const [folder, description] of Object.entries(expectedStructure)) {
      const folderPath = path.join(__dirname, '..', folder);
      const exists = fs.existsSync(folderPath);
      
      logTest(`${folder}/`, exists, `${description} ${exists ? 'found' : 'missing'}`);
      
      if (exists) {
        this.results.passed++;
      } else {
        this.results.failed++;
      }
    }
  }

  // Test 8: Legacy Code Cleanup
  testLegacyCleanup() {
    logSection('🧹 Legacy Code Cleanup');
    
    const legacyFiles = [
      'UserManager.gs',
      'ConfigManager.gs',
      'UnifiedManager.gs',
      'auth.gs'
    ];

    let cleanupScore = 0;
    for (const legacyFile of legacyFiles) {
      const legacyPath = path.join(this.srcPath, legacyFile);
      const exists = fs.existsSync(legacyPath);
      
      logTest(`${legacyFile} removed`, !exists, exists ? 'Legacy file still present' : 'Legacy file properly removed');
      
      if (!exists) {
        cleanupScore++;
        this.results.passed++;
      } else {
        this.results.failed++;
      }
    }

    const cleanupPercent = (cleanupScore / legacyFiles.length) * 100;
    log(cleanupPercent === 100 ? 'green' : 'yellow', `Legacy cleanup: ${cleanupPercent}%`);
  }

  // Run all tests
  runAll() {
    console.log('🏗️ Architecture Testing Suite');
    console.log('Testing against README.md standards...\n');

    this.testServiceLayer();
    this.testUtilityLayer(); 
    this.testErrorHandling();
    this.testConstants();
    this.testMainEntryPoint();
    this.testCoverage();
    this.testFileStructure();
    this.testLegacyCleanup();

    // Summary
    logSection('📊 Test Summary');
    const total = this.results.passed + this.results.failed;
    const passRate = total > 0 ? ((this.results.passed / total) * 100).toFixed(1) : 0;

    console.log(`✅ Passed: ${this.results.passed}`);
    console.log(`❌ Failed: ${this.results.failed}`);
    console.log(`⚠️ Warnings: ${this.results.warnings}`);
    console.log(`📊 Pass Rate: ${passRate}%`);

    if (this.results.failed === 0) {
      log('green', '\n🎉 All architecture tests passed!');
      return true;
    } else {
      log('red', '\n⚠️ Architecture issues detected. Please address failed tests.');
      return false;
    }
  }
}

// Run the tests
const tester = new ArchitectureTester();
const success = tester.runAll();

// Exit with appropriate code
process.exit(success ? 0 : 1);