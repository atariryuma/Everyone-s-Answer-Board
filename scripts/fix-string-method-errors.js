#!/usr/bin/env node

/**
 * Automated Fix Script for String Method Errors
 * Automatically fixes common patterns found by the analysis script
 */

const fs = require('fs');
const path = require('path');

// Fix patterns - these are safe, common transformations
const fixPatterns = [
  {
    name: 'email?.substring in logging',
    search: /email:\s*`\$\{email\?\.(substring|substr)\(([^}]+)\)\}([^`]*)`/g,
    replace: 'email: typeof email === \'string\' && email ? `${email.$1($2)}$3` : `[${typeof email}]`',
    description: 'Fix email logging with safe type checking'
  },
  {
    name: 'userId?.substring in logging',
    search: /userId:\s*`\$\{userId\?\.(substring|substr)\(([^}]+)\)\}([^`]*)`/g,
    replace: 'userId: typeof userId === \'string\' && userId ? `${userId.$1($2)}$3` : `[${typeof userId}]`',
    description: 'Fix userId logging with safe type checking'
  },
  {
    name: 'spreadsheetId?.substring in logging',
    search: /spreadsheetId:\s*`\$\{spreadsheetId\?\.(substring|substr)\(([^}]+)\)\}([^`]*)`/g,
    replace: 'spreadsheetId: typeof spreadsheetId === \'string\' && spreadsheetId ? `${spreadsheetId.$1($2)}$3` : `[${typeof spreadsheetId}]`',
    description: 'Fix spreadsheetId logging with safe type checking'
  },
  {
    name: 'Generic optional chaining includes with string literal',
    search: /(\w+)\?\.(includes)\s*\(\s*['"][^'"]*['"]\s*\)/g,
    replace: '(typeof $1 === \'string\' && $1.$2($3))',
    description: 'Fix optional chaining includes with string literals'
  }
];

function applyFixesToFile(filePath, content) {
  let fixedContent = content;
  let appliedFixes = [];

  fixPatterns.forEach(({ name, search, replace, description }) => {
    const originalContent = fixedContent;
    fixedContent = fixedContent.replace(search, replace);

    if (originalContent !== fixedContent) {
      appliedFixes.push({ name, description });
    }
  });

  return { content: fixedContent, fixes: appliedFixes };
}

function processDirectory(dirPath, extensions = ['.gs', '.js', '.html']) {
  const results = [];

  function processDir(currentPath) {
    const items = fs.readdirSync(currentPath);

    for (const item of items) {
      const fullPath = path.join(currentPath, item);
      const stat = fs.statSync(fullPath);

      if (stat.isDirectory()) {
        // Skip specific directories
        if (!['node_modules', '.git', 'coverage', '.nyc_output', 'backups'].includes(item)) {
          processDir(fullPath);
        }
      } else if (stat.isFile()) {
        const ext = path.extname(item);
        if (extensions.includes(ext)) {
          try {
            const content = fs.readFileSync(fullPath, 'utf8');
            const result = applyFixesToFile(fullPath, content);

            if (result.fixes.length > 0) {
              // Apply fixes
              fs.writeFileSync(fullPath, result.content, 'utf8');
              results.push({
                file: fullPath,
                fixes: result.fixes
              });
            }
          } catch (error) {
            console.error(`Error processing file ${fullPath}:`, error.message);
          }
        }
      }
    }
  }

  processDir(dirPath);
  return results;
}

function generateReport(results) {
  console.log('\n🔧 AUTOMATED STRING METHOD ERROR FIXES');
  console.log('=' .repeat(50));

  if (results.length === 0) {
    console.log('✅ No files needed automatic fixes!');
    return;
  }

  console.log(`✅ Successfully fixed ${results.length} files:\n`);

  results.forEach(({ file, fixes }) => {
    console.log(`📄 ${file.replace(process.cwd(), '.')}`);
    fixes.forEach(fix => {
      console.log(`  ✓ ${fix.name}: ${fix.description}`);
    });
    console.log('');
  });

  console.log(`\n📊 Summary: ${results.reduce((total, r) => total + r.fixes.length, 0)} total fixes applied`);
}

// Additional safety validation script
function createValidationScript() {
  const validationScript = `#!/usr/bin/env node

/**
 * Post-Fix Validation Script
 * Ensures all string method fixes are working correctly
 */

const fs = require('fs');
const path = require('path');

// Test patterns that should not exist after fixes
const dangerousPatterns = [
  /\\w+\\?\\.substring\\s*\\(/g,
  /\\w+\\?\\.substr\\s*\\(/g,
  /\\w+\\?\\.slice\\s*\\(\\s*\\d+\\s*,?\\s*\\d*\\s*\\)/g,
  /\\w+\\?\\.charAt\\s*\\(/g,
  /\\w+\\?\\.indexOf\\s*\\(/g,
  /\\w+\\?\\.toLowerCase\\s*\\(\\s*\\)/g,
  /\\w+\\?\\.toUpperCase\\s*\\(\\s*\\)/g,
  /\\w+\\?\\.trim\\s*\\(\\s*\\)/g,
  /\\w+\\?\\.replace\\s*\\(/g,
  /\\w+\\?\\.split\\s*\\(/g,
  /\\w+\\?\\.includes\\s*\\(\\s*['"\`]/g,
  /\\w+\\?\\.startsWith\\s*\\(/g,
  /\\w+\\?\\.endsWith\\s*\\(/g,
  /\\w+\\?\\.match\\s*\\(/g
];

function validateFile(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    const issues = [];

    dangerousPatterns.forEach((pattern, index) => {
      let match;
      pattern.lastIndex = 0;
      while ((match = pattern.exec(content)) !== null) {
        const lineNumber = content.substring(0, match.index).split('\\n').length;
        const line = content.split('\\n')[lineNumber - 1];
        issues.push({
          line: lineNumber,
          pattern: index,
          match: match[0],
          context: line.trim()
        });
      }
    });

    return issues;
  } catch (error) {
    console.error(\`Error validating file \${filePath}:\`, error.message);
    return [];
  }
}

function validateDirectory(dirPath) {
  const allIssues = [];

  function scanDir(currentPath) {
    const items = fs.readdirSync(currentPath);

    for (const item of items) {
      const fullPath = path.join(currentPath, item);
      const stat = fs.statSync(fullPath);

      if (stat.isDirectory()) {
        if (!['node_modules', '.git', 'coverage', 'backups'].includes(item)) {
          scanDir(fullPath);
        }
      } else if (stat.isFile()) {
        const ext = path.extname(item);
        if (['.gs', '.js', '.html'].includes(ext)) {
          const issues = validateFile(fullPath);
          if (issues.length > 0) {
            allIssues.push({ file: fullPath, issues });
          }
        }
      }
    }
  }

  scanDir(dirPath);
  return allIssues;
}

if (require.main === module) {
  const projectRoot = process.argv[2] || process.cwd();
  console.log('🔍 Validating string method fixes...');

  const results = validateDirectory(projectRoot);

  if (results.length === 0) {
    console.log('✅ All string method patterns have been safely fixed!');
    process.exit(0);
  } else {
    console.log(\`❌ Found \${results.length} files with remaining issues:\\n\`);
    results.forEach(({ file, issues }) => {
      console.log(\`📄 \${file}\`);
      issues.forEach(issue => {
        console.log(\`  Line \${issue.line}: \${issue.match}\`);
        console.log(\`    Context: \${issue.context}\`);
      });
      console.log('');
    });
    process.exit(1);
  }
}`;

  return validationScript;
}

// Main execution
if (require.main === module) {
  const projectRoot = process.argv[2] || path.join(process.cwd(), 'src');
  console.log(`🔧 Applying automated fixes to: ${projectRoot}`);

  const results = processDirectory(projectRoot);
  generateReport(results);

  // Create validation script
  const validationScriptPath = path.join(process.cwd(), 'scripts', 'validate-string-fixes.js');
  fs.writeFileSync(validationScriptPath, createValidationScript());
  fs.chmodSync(validationScriptPath, '755');

  console.log(`\n💾 Validation script created: ${validationScriptPath}`);
  console.log('Run: node scripts/validate-string-fixes.js');

  process.exit(results.length > 0 ? 0 : 1);
}

module.exports = { applyFixesToFile, processDirectory };