#!/usr/bin/env node

/**
 * Comprehensive Script to Find Potential String Method Errors
 * Searches for optional chaining with string methods that could fail if the value is not a string
 */

const fs = require('fs');
const path = require('path');

// Patterns to search for
const dangerousPatterns = [
  {
    name: 'Optional Chaining with substring',
    pattern: /(\w+)\?\.(substring|substr)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with slice (string context)',
    pattern: /(\w+)\?\.(slice)\s*\(\s*\d+\s*,?\s*\d*\s*\)/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with charAt',
    pattern: /(\w+)\?\.(charAt)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with indexOf',
    pattern: /(\w+)\?\.(indexOf|lastIndexOf)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with toLowerCase/toUpperCase',
    pattern: /(\w+)\?\.(toLowerCase|toUpperCase)\s*\(\s*\)/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with trim',
    pattern: /(\w+)\?\.(trim|trimStart|trimEnd)\s*\(\s*\)/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with replace',
    pattern: /(\w+)\?\.(replace|replaceAll)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with split',
    pattern: /(\w+)\?\.(split)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with includes (string context)',
    pattern: /(\w+)\?\.(includes)\s*\(\s*['"`]/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with startsWith/endsWith',
    pattern: /(\w+)\?\.(startsWith|endsWith)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Optional Chaining with match',
    pattern: /(\w+)\?\.(match)\s*\(/g,
    description: 'May fail if variable is not a string'
  },
  {
    name: 'Template literal with optional chaining string methods',
    pattern: /\$\{[^}]*(\w+)\?\.(substring|substr|slice|charAt|toLowerCase|toUpperCase|trim)[^}]*\}/g,
    description: 'Template literal may fail if variable is not a string'
  }
];

// Additional risky patterns without optional chaining
const riskyPatterns = [
  {
    name: 'Direct string method on potentially undefined',
    pattern: /(\w+)\.(substring|substr|slice|charAt)\s*\(/g,
    description: 'May fail if variable is null/undefined',
    checkContext: true
  }
];

function searchInFile(filePath, content) {
  const results = [];

  dangerousPatterns.forEach(({ name, pattern, description }) => {
    let match;
    pattern.lastIndex = 0; // Reset regex

    while ((match = pattern.exec(content)) !== null) {
      const lineNumber = content.substring(0, match.index).split('\n').length;
      const line = content.split('\n')[lineNumber - 1];

      results.push({
        file: filePath,
        line: lineNumber,
        column: match.index - content.lastIndexOf('\n', match.index),
        pattern: name,
        description,
        matchedText: match[0],
        variable: match[1],
        method: match[2] || match[0].match(/\.\w+/)?.[0]?.substring(1),
        fullLine: line.trim(),
        severity: 'HIGH'
      });
    }
  });

  return results;
}

function scanDirectory(dirPath, extensions = ['.gs', '.js', '.html']) {
  const results = [];

  function scanDir(currentPath) {
    const items = fs.readdirSync(currentPath);

    for (const item of items) {
      const fullPath = path.join(currentPath, item);
      const stat = fs.statSync(fullPath);

      if (stat.isDirectory()) {
        // Skip node_modules and .git directories
        if (!['node_modules', '.git', 'coverage', '.nyc_output'].includes(item)) {
          scanDir(fullPath);
        }
      } else if (stat.isFile()) {
        const ext = path.extname(item);
        if (extensions.includes(ext)) {
          try {
            const content = fs.readFileSync(fullPath, 'utf8');
            const fileResults = searchInFile(fullPath, content);
            results.push(...fileResults);
          } catch (error) {
            console.error(`Error reading file ${fullPath}:`, error.message);
          }
        }
      }
    }
  }

  scanDir(dirPath);
  return results;
}

function generateReport(results) {
  console.log('\n🔍 COMPREHENSIVE STRING METHOD ERROR ANALYSIS');
  console.log('=' .repeat(60));

  if (results.length === 0) {
    console.log('✅ No dangerous string method patterns found!');
    return;
  }

  // Group by file
  const byFile = {};
  results.forEach(result => {
    if (!byFile[result.file]) byFile[result.file] = [];
    byFile[result.file].push(result);
  });

  console.log(`\n❌ Found ${results.length} potential string method errors in ${Object.keys(byFile).length} files:\n`);

  Object.entries(byFile).forEach(([file, issues]) => {
    console.log(`📄 ${file.replace(process.cwd(), '.')}`);
    console.log('-'.repeat(50));

    issues.forEach(issue => {
      console.log(`  Line ${issue.line}: ${issue.pattern}`);
      console.log(`    🔍 Code: ${issue.fullLine}`);
      console.log(`    ⚠️  Issue: ${issue.description}`);
      console.log(`    🎯 Variable: "${issue.variable}" Method: "${issue.method}"`);
      console.log('');
    });
  });

  // Summary by pattern type
  console.log('\n📊 SUMMARY BY PATTERN TYPE:');
  console.log('-'.repeat(30));
  const patternCounts = {};
  results.forEach(r => {
    patternCounts[r.pattern] = (patternCounts[r.pattern] || 0) + 1;
  });

  Object.entries(patternCounts)
    .sort(([,a], [,b]) => b - a)
    .forEach(([pattern, count]) => {
      console.log(`${count.toString().padStart(3)} × ${pattern}`);
    });
}

function generateFixSuggestions(results) {
  console.log('\n🔧 RECOMMENDED FIXES:');
  console.log('=' .repeat(40));

  const fixes = new Map();

  results.forEach(result => {
    const key = `${result.file}:${result.line}`;
    if (!fixes.has(key)) {
      let suggestion = '';
      const variable = result.variable;
      const method = result.method;

      if (method === 'substring' || method === 'substr') {
        suggestion = `typeof ${variable} === 'string' && ${variable} ? ${variable}.${method}(...) : ''`;
      } else if (method === 'toLowerCase' || method === 'toUpperCase') {
        suggestion = `typeof ${variable} === 'string' && ${variable} ? ${variable}.${method}() : ${variable}`;
      } else if (method === 'trim') {
        suggestion = `typeof ${variable} === 'string' && ${variable} ? ${variable}.trim() : ${variable}`;
      } else {
        suggestion = `typeof ${variable} === 'string' && ${variable} ? ${variable}.${method}(...) : fallbackValue`;
      }

      fixes.set(key, {
        file: result.file,
        line: result.line,
        original: result.matchedText,
        suggestion,
        context: result.fullLine
      });
    }
  });

  fixes.forEach((fix, key) => {
    console.log(`📍 ${fix.file.replace(process.cwd(), '.')}:${fix.line}`);
    console.log(`   Original: ${fix.original}`);
    console.log(`   Suggested: ${fix.suggestion}`);
    console.log('');
  });
}

// Main execution
if (require.main === module) {
  const projectRoot = process.argv[2] || process.cwd();
  console.log(`🔍 Scanning project: ${projectRoot}`);

  const results = scanDirectory(projectRoot);
  generateReport(results);
  generateFixSuggestions(results);

  // Export results as JSON for automated processing
  const outputFile = path.join(projectRoot, 'string-method-analysis.json');
  fs.writeFileSync(outputFile, JSON.stringify({
    timestamp: new Date().toISOString(),
    totalIssues: results.length,
    results: results
  }, null, 2));

  console.log(`\n💾 Detailed results saved to: ${outputFile}`);

  // Exit with error code if issues found
  process.exit(results.length > 0 ? 1 : 0);
}

module.exports = { searchInFile, scanDirectory, generateReport };