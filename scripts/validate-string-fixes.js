#!/usr/bin/env node

/**
 * Post-Fix Validation Script
 * Ensures all string method fixes are working correctly
 */

const fs = require('fs');
const path = require('path');

// Test patterns that should not exist after fixes
const dangerousPatterns = [
  /\w+\?\.substring\s*\(/g,
  /\w+\?\.substr\s*\(/g,
  /\w+\?\.slice\s*\(\s*\d+\s*,?\s*\d*\s*\)/g,
  /\w+\?\.charAt\s*\(/g,
  /\w+\?\.indexOf\s*\(/g,
  /\w+\?\.toLowerCase\s*\(\s*\)/g,
  /\w+\?\.toUpperCase\s*\(\s*\)/g,
  /\w+\?\.trim\s*\(\s*\)/g,
  /\w+\?\.replace\s*\(/g,
  /\w+\?\.split\s*\(/g,
  /\w+\?\.includes\s*\(\s*['"`]/g,
  /\w+\?\.startsWith\s*\(/g,
  /\w+\?\.endsWith\s*\(/g,
  /\w+\?\.match\s*\(/g
];

function validateFile(filePath) {
  try {
    const content = fs.readFileSync(filePath, 'utf8');
    const issues = [];

    dangerousPatterns.forEach((pattern, index) => {
      let match;
      pattern.lastIndex = 0;
      while ((match = pattern.exec(content)) !== null) {
        const lineNumber = content.substring(0, match.index).split('\n').length;
        const line = content.split('\n')[lineNumber - 1];
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
    console.error(`Error validating file ${filePath}:`, error.message);
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
    console.log(`❌ Found ${results.length} files with remaining issues:\n`);
    results.forEach(({ file, issues }) => {
      console.log(`📄 ${file}`);
      issues.forEach(issue => {
        console.log(`  Line ${issue.line}: ${issue.match}`);
        console.log(`    Context: ${issue.context}`);
      });
      console.log('');
    });
    process.exit(1);
  }
}