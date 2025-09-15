#!/usr/bin/env node

/**
 * Remove duplicate functions between DataService.gs and DataController.gs
 * Keep DataService.gs versions as they are more detailed implementations
 */

const fs = require('fs');
const path = require('path');

const FUNCTIONS_TO_REMOVE_FROM_DATA_CONTROLLER = [
  'addDataReaction',
  'toggleDataHighlight',
  'getPublishedSheetData'
];

function removeFunction(content, functionName) {
  // Match function declaration and its body with proper brace counting
  const regex = new RegExp(`function\\s+${functionName}\\s*\\([^)]*\\)\\s*\\{`, 'g');
  let match = regex.exec(content);

  if (!match) {
    console.log(`⚠️  Function ${functionName} not found`);
    return content;
  }

  const startIndex = match.index;
  let braceCount = 0;
  let endIndex = startIndex;
  let inFunction = false;

  // Find the opening brace and start counting
  for (let i = startIndex; i < content.length; i++) {
    if (content[i] === '{') {
      braceCount++;
      inFunction = true;
    } else if (content[i] === '}' && inFunction) {
      braceCount--;
      if (braceCount === 0) {
        endIndex = i + 1;
        break;
      }
    }
  }

  // Find the comment block before the function
  let commentStart = startIndex;
  let beforeFunction = content.substring(0, startIndex).trimEnd();

  // Look for /** comment block
  if (beforeFunction.endsWith('*/')) {
    let searchIndex = beforeFunction.lastIndexOf('/**');
    if (searchIndex !== -1) {
      commentStart = searchIndex;
    }
  }

  const before = content.substring(0, commentStart);
  const after = content.substring(endIndex);

  console.log(`✅ Removed function: ${functionName}`);
  return before + after;
}

function main() {
  const filePath = path.join(__dirname, '../src/DataController.gs');

  try {
    let content = fs.readFileSync(filePath, 'utf8');
    console.log(`📄 Processing: ${filePath}`);
    console.log(`📊 Original size: ${content.length} characters`);

    // Remove each duplicate function from DataController.gs
    FUNCTIONS_TO_REMOVE_FROM_DATA_CONTROLLER.forEach(functionName => {
      content = removeFunction(content, functionName);
    });

    // Write back the cleaned content
    fs.writeFileSync(filePath, content, 'utf8');

    console.log(`📊 New size: ${content.length} characters`);
    console.log(`✅ Successfully removed ${FUNCTIONS_TO_REMOVE_FROM_DATA_CONTROLLER.length} duplicate functions from DataController.gs`);
    console.log(`🎯 These functions now exist only in DataService.gs (detailed implementations)`);

  } catch (error) {
    console.error('❌ Error processing file:', error.message);
    process.exit(1);
  }
}

if (require.main === module) {
  main();
}