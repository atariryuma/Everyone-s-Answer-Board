#!/usr/bin/env node
/**
 * Security Review Script
 * Comprehensive security audit and vulnerability assessment
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

function logTest(name, passed, severity = 'info', details = '') {
  const statusMap = {
    critical: '🚨 CRITICAL',
    high: '❌ HIGH',
    medium: '⚠️ MEDIUM', 
    low: '💡 LOW',
    pass: '✅ PASS'
  };
  
  const status = passed ? statusMap.pass : statusMap[severity] || statusMap.high;
  console.log(`${status} ${name}`);
  
  if (details) {
    console.log(`    ${details}`);
  }
}

class SecurityReviewer {
  constructor() {
    this.srcPath = path.join(__dirname, '../src');
    this.results = {
      critical: 0,
      high: 0,
      medium: 0,
      low: 0,
      passed: 0
    };
  }

  // Security Test 1: Input Validation
  testInputValidation() {
    logSection('🛡️ Input Validation Security');
    
    const securityServicePath = path.join(this.srcPath, 'services', 'SecurityService.gs');
    
    if (!fs.existsSync(securityServicePath)) {
      logTest('SecurityService existence', false, 'critical', 'SecurityService.gs not found');
      this.results.critical++;
      return;
    }

    const content = fs.readFileSync(securityServicePath, 'utf8');
    
    // Check for XSS protection
    const hasXSSProtection = content.includes('escapeHtml') || content.includes('sanitize');
    logTest('XSS Protection', hasXSSProtection, hasXSSProtection ? 'pass' : 'high', 
      hasXSSProtection ? 'HTML escaping found' : 'No XSS protection detected');
    
    // Check for input length validation
    const hasLengthValidation = content.includes('maxLength') || content.includes('length');
    logTest('Input Length Validation', hasLengthValidation, hasLengthValidation ? 'pass' : 'medium',
      hasLengthValidation ? 'Length validation found' : 'No input length limits detected');
    
    // Check for email validation
    const hasEmailValidation = content.includes('validateEmail') || content.includes('@');
    logTest('Email Validation', hasEmailValidation, hasEmailValidation ? 'pass' : 'medium',
      hasEmailValidation ? 'Email validation found' : 'No email validation detected');
    
    if (hasXSSProtection && hasLengthValidation && hasEmailValidation) {
      this.results.passed++;
    } else {
      if (!hasXSSProtection) this.results.high++;
      if (!hasLengthValidation) this.results.medium++;
      if (!hasEmailValidation) this.results.medium++;
    }
  }

  // Security Test 2: Authentication & Authorization
  testAuthentication() {
    logSection('🔐 Authentication & Authorization');
    
    const userServicePath = path.join(this.srcPath, 'services', 'UserService.gs');
    const securityServicePath = path.join(this.srcPath, 'services', 'SecurityService.gs');
    
    if (!fs.existsSync(userServicePath)) {
      logTest('UserService existence', false, 'critical', 'UserService.gs not found');
      this.results.critical++;
      return;
    }
    
    const userContent = fs.readFileSync(userServicePath, 'utf8');
    const securityContent = fs.existsSync(securityServicePath) ? 
      fs.readFileSync(securityServicePath, 'utf8') : '';
    
    // Check for proper session handling
    const hasSessionHandling = userContent.includes('Session.getActiveUser') || 
                               userContent.includes('getCurrentEmail');
    logTest('Session Management', hasSessionHandling, hasSessionHandling ? 'pass' : 'high',
      hasSessionHandling ? 'Session handling found' : 'No session management detected');
    
    // Check for access level control
    const hasAccessControl = securityContent.includes('checkUserPermission') || 
                            securityContent.includes('getAccessLevel');
    logTest('Access Control', hasAccessControl, hasAccessControl ? 'pass' : 'high',
      hasAccessControl ? 'Access control found' : 'No access control system detected');
    
    // Check for email domain validation
    const hasDomainValidation = securityContent.includes('validateEmailAccess') ||
                               securityContent.includes('domain');
    logTest('Email Domain Validation', hasDomainValidation, hasDomainValidation ? 'pass' : 'medium',
      hasDomainValidation ? 'Domain validation found' : 'No domain validation detected');
    
    if (hasSessionHandling && hasAccessControl) {
      this.results.passed++;
    } else {
      if (!hasSessionHandling) this.results.high++;
      if (!hasAccessControl) this.results.high++;
      if (!hasDomainValidation) this.results.medium++;
    }
  }

  // Security Test 3: Data Protection
  testDataProtection() {
    logSection('📊 Data Protection');
    
    // Check for sensitive data exposure in logs
    const allFiles = this.getAllGSFiles();
    let hasSensitiveLogging = false;
    let hasProperLogging = true;
    
    for (const filePath of allFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      
      // Check for password/token logging
      if (content.includes('console.log') && 
          (content.includes('password') || content.includes('token') || content.includes('secret'))) {
        hasSensitiveLogging = true;
      }
      
      // Check for structured logging
      if (content.includes('console.log') && !content.includes('console.error')) {
        hasProperLogging = false;
      }
    }
    
    logTest('No Sensitive Data in Logs', !hasSensitiveLogging, hasSensitiveLogging ? 'high' : 'pass',
      hasSensitiveLogging ? 'Sensitive data may be logged' : 'No sensitive data logging detected');
    
    logTest('Structured Logging', hasProperLogging, hasProperLogging ? 'pass' : 'low',
      hasProperLogging ? 'Proper error logging found' : 'Consider structured logging');
    
    // Check for error message sanitization
    const errorHandlerPath = path.join(this.srcPath, 'core', 'errors.gs');
    let hasErrorSanitization = false;
    
    if (fs.existsSync(errorHandlerPath)) {
      const errorContent = fs.readFileSync(errorHandlerPath, 'utf8');
      hasErrorSanitization = errorContent.includes('sanitize') || 
                            errorContent.includes('userFriendlyMessage');
    }
    
    logTest('Error Message Sanitization', hasErrorSanitization, hasErrorSanitization ? 'pass' : 'medium',
      hasErrorSanitization ? 'Error sanitization found' : 'Error messages may expose internal info');
    
    if (!hasSensitiveLogging && hasProperLogging && hasErrorSanitization) {
      this.results.passed++;
    } else {
      if (hasSensitiveLogging) this.results.high++;
      if (!hasProperLogging) this.results.low++;
      if (!hasErrorSanitization) this.results.medium++;
    }
  }

  // Security Test 4: Code Injection Prevention  
  testCodeInjection() {
    logSection('💉 Code Injection Prevention');
    
    const allFiles = this.getAllGSFiles();
    let hasEvalUsage = false;
    let hasInnerHTMLUsage = false;
    let hasProperSanitization = false;
    
    for (const filePath of allFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      
      // Check for eval() usage
      if (content.includes('eval(')) {
        hasEvalUsage = true;
      }
      
      // Check for innerHTML usage without sanitization
      if (content.includes('innerHTML') && !content.includes('escapeHtml')) {
        hasInnerHTMLUsage = true;
      }
      
      // Check for sanitization functions
      if (content.includes('sanitize') || content.includes('escapeHtml')) {
        hasProperSanitization = true;
      }
    }
    
    logTest('No eval() Usage', !hasEvalUsage, hasEvalUsage ? 'critical' : 'pass',
      hasEvalUsage ? 'eval() usage detected - critical security risk' : 'No eval() usage found');
    
    logTest('Safe innerHTML Usage', !hasInnerHTMLUsage, hasInnerHTMLUsage ? 'high' : 'pass',
      hasInnerHTMLUsage ? 'Unsanitized innerHTML usage detected' : 'No unsafe innerHTML usage');
    
    logTest('Proper Sanitization', hasProperSanitization, hasProperSanitization ? 'pass' : 'high',
      hasProperSanitization ? 'Sanitization functions found' : 'No sanitization functions detected');
    
    if (!hasEvalUsage && !hasInnerHTMLUsage && hasProperSanitization) {
      this.results.passed++;
    } else {
      if (hasEvalUsage) this.results.critical++;
      if (hasInnerHTMLUsage) this.results.high++;
      if (!hasProperSanitization) this.results.high++;
    }
  }

  // Security Test 5: Access Control Validation
  testAccessControl() {
    logSection('🚫 Access Control Validation');
    
    const mainPath = path.join(this.srcPath, 'main.gs');
    
    if (!fs.existsSync(mainPath)) {
      logTest('Main entry point security', false, 'critical', 'main.gs not found');
      this.results.critical++;
      return;
    }
    
    const mainContent = fs.readFileSync(mainPath, 'utf8');
    
    // Check for authentication on doGet/doPost
    const hasDoGetAuth = mainContent.includes('doGet') && 
                        (mainContent.includes('UserService.getCurrentEmail') || 
                         mainContent.includes('SecurityService.validateAccess'));
    
    const hasDoPostAuth = mainContent.includes('doPost') && 
                         (mainContent.includes('validateAccess') || 
                          mainContent.includes('checkUserPermission'));
    
    logTest('doGet Authentication', hasDoGetAuth, hasDoGetAuth ? 'pass' : 'high',
      hasDoGetAuth ? 'doGet has authentication' : 'doGet may lack authentication');
    
    logTest('doPost Authentication', hasDoPostAuth, hasDoPostAuth ? 'pass' : 'high', 
      hasDoPostAuth ? 'doPost has authentication' : 'doPost may lack authentication');
    
    // Check for CSRF protection (if applicable)
    const hasCSRFProtection = mainContent.includes('csrf') || mainContent.includes('token');
    logTest('CSRF Protection', hasCSRFProtection, hasCSRFProtection ? 'pass' : 'medium',
      hasCSRFProtection ? 'CSRF protection found' : 'Consider CSRF protection for state-changing operations');
    
    if (hasDoGetAuth && hasDoPostAuth) {
      this.results.passed++;
    } else {
      if (!hasDoGetAuth) this.results.high++;
      if (!hasDoPostAuth) this.results.high++;
      if (!hasCSRFProtection) this.results.medium++;
    }
  }

  // Security Test 6: Configuration Security
  testConfigurationSecurity() {
    logSection('⚙️ Configuration Security');
    
    const allFiles = this.getAllGSFiles();
    let hasHardcodedSecrets = false;
    let usesPropertiesService = false;
    let hasSecureDefaults = true;
    
    for (const filePath of allFiles) {
      const content = fs.readFileSync(filePath, 'utf8');
      
      // Check for hardcoded secrets (simple patterns)
      const secretPatterns = [
        /password\s*[:=]\s*["'][^"']+["']/i,
        /api_key\s*[:=]\s*["'][^"']+["']/i,
        /secret\s*[:=]\s*["'][^"']+["']/i,
        /token\s*[:=]\s*["'][^"']+["']/i
      ];
      
      for (const pattern of secretPatterns) {
        if (pattern.test(content)) {
          hasHardcodedSecrets = true;
        }
      }
      
      // Check for PropertiesService usage
      if (content.includes('PropertiesService')) {
        usesPropertiesService = true;
      }
    }
    
    logTest('No Hardcoded Secrets', !hasHardcodedSecrets, hasHardcodedSecrets ? 'critical' : 'pass',
      hasHardcodedSecrets ? 'Potential hardcoded secrets detected' : 'No hardcoded secrets found');
    
    logTest('Uses PropertiesService', usesPropertiesService, usesPropertiesService ? 'pass' : 'medium',
      usesPropertiesService ? 'PropertiesService used for config' : 'Consider using PropertiesService for sensitive config');
    
    // Check core/constants.gs for secure defaults
    const constantsPath = path.join(this.srcPath, 'core', 'constants.gs');
    if (fs.existsSync(constantsPath)) {
      const constantsContent = fs.readFileSync(constantsPath, 'utf8');
      hasSecureDefaults = !constantsContent.includes('debug: true') && 
                         !constantsContent.includes('allowAll: true');
    }
    
    logTest('Secure Default Configuration', hasSecureDefaults, hasSecureDefaults ? 'pass' : 'medium',
      hasSecureDefaults ? 'Secure defaults found' : 'Check for insecure default configurations');
    
    if (!hasHardcodedSecrets && usesPropertiesService && hasSecureDefaults) {
      this.results.passed++;
    } else {
      if (hasHardcodedSecrets) this.results.critical++;
      if (!usesPropertiesService) this.results.medium++;
      if (!hasSecureDefaults) this.results.medium++;
    }
  }

  // Helper: Get all .gs files
  getAllGSFiles() {
    const files = [];
    
    function scanDirectory(dir) {
      if (!fs.existsSync(dir)) return;
      
      const items = fs.readdirSync(dir);
      
      for (const item of items) {
        const itemPath = path.join(dir, item);
        const stat = fs.statSync(itemPath);
        
        if (stat.isDirectory()) {
          scanDirectory(itemPath);
        } else if (item.endsWith('.gs')) {
          files.push(itemPath);
        }
      }
    }
    
    scanDirectory(this.srcPath);
    return files;
  }

  // Security Test 7: Dependency Security
  testDependencySecurity() {
    logSection('📦 Dependency Security');
    
    const packagePath = path.join(__dirname, '../package.json');
    
    if (!fs.existsSync(packagePath)) {
      logTest('Package.json exists', false, 'medium', 'package.json not found');
      this.results.medium++;
      return;
    }
    
    const packageContent = JSON.parse(fs.readFileSync(packagePath, 'utf8'));
    const devDeps = packageContent.devDependencies || {};
    
    // Check for security-related dev dependencies
    const hasESLint = devDeps.eslint !== undefined;
    const hasPrettier = devDeps.prettier !== undefined;
    const hasJest = devDeps.jest !== undefined;
    
    logTest('ESLint for Code Quality', hasESLint, hasESLint ? 'pass' : 'low',
      hasESLint ? 'ESLint found' : 'Consider adding ESLint for code quality');
    
    logTest('Jest for Testing', hasJest, hasJest ? 'pass' : 'medium',
      hasJest ? 'Jest testing framework found' : 'Testing framework not found');
    
    // Check for known vulnerable patterns in dependencies
    const hasOutdatedDeps = Object.values(devDeps).some(version => 
      version.includes('^0.') || version.includes('alpha') || version.includes('beta')
    );
    
    logTest('Stable Dependencies', !hasOutdatedDeps, hasOutdatedDeps ? 'low' : 'pass',
      hasOutdatedDeps ? 'Some dependencies may be unstable' : 'Dependencies appear stable');
    
    if (hasESLint && hasJest && !hasOutdatedDeps) {
      this.results.passed++;
    } else {
      if (!hasESLint) this.results.low++;
      if (!hasJest) this.results.medium++;
      if (hasOutdatedDeps) this.results.low++;
    }
  }

  // Calculate security score
  calculateSecurityScore() {
    const totalIssues = this.results.critical * 4 + this.results.high * 3 + 
                       this.results.medium * 2 + this.results.low * 1;
    const maxPossibleScore = 100;
    const deductions = Math.min(totalIssues * 5, maxPossibleScore);
    
    return Math.max(0, maxPossibleScore - deductions);
  }

  // Run all security tests
  runAll() {
    console.log('🔒 Security Review Suite');
    console.log('Comprehensive security audit and vulnerability assessment...\n');

    this.testInputValidation();
    this.testAuthentication();
    this.testDataProtection();
    this.testCodeInjection();
    this.testAccessControl();
    this.testConfigurationSecurity();
    this.testDependencySecurity();

    // Summary
    logSection('📊 Security Assessment Summary');
    
    const score = this.calculateSecurityScore();
    const totalIssues = this.results.critical + this.results.high + this.results.medium + this.results.low;
    
    console.log(`🚨 Critical Issues: ${this.results.critical}`);
    console.log(`❌ High Risk Issues: ${this.results.high}`);
    console.log(`⚠️ Medium Risk Issues: ${this.results.medium}`);
    console.log(`💡 Low Risk Issues: ${this.results.low}`);
    console.log(`✅ Security Tests Passed: ${this.results.passed}`);
    console.log(`📊 Overall Security Score: ${score}/100`);

    // Security status
    if (this.results.critical > 0) {
      log('red', '\n🚨 CRITICAL SECURITY ISSUES DETECTED! Address immediately.');
      return false;
    } else if (this.results.high > 0) {
      log('yellow', '\n⚠️ High-risk security issues found. Please review and fix.');
      return false;
    } else if (score >= 80) {
      log('green', '\n🛡️ Security review passed! Good security posture.');
      return true;
    } else {
      log('yellow', '\n⚠️ Security score below recommended threshold. Consider improvements.');
      return false;
    }
  }
}

// Run the security review
const reviewer = new SecurityReviewer();
const success = reviewer.runAll();

// Exit with appropriate code
process.exit(success ? 0 : 1);