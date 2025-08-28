#!/usr/bin/env node

const fs = require('fs');
const path = require('path');

/**
 * レスポンス形式を統一関数に変換するスクリプト
 * success: true/false を createSuccessResponse/createErrorResponse に変換
 */

function convertResponseFormat(content) {
  // パターン1: { success: true, ... } の形式を変換
  content = content.replace(
    /return\s*\{\s*success:\s*true,\s*([^}]+)\s*\};?/g,
    (match, rest) => {
      // restから個別のプロパティを抽出
      const props = rest.split(',').map(p => p.trim()).filter(p => p);
      const data = {};
      let message = null;
      
      props.forEach(prop => {
        const [key, value] = prop.split(':').map(s => s.trim());
        if (key === 'message') {
          message = value;
        } else {
          data[key] = value;
        }
      });
      
      const dataStr = Object.keys(data).length > 0 ? `{ ${Object.entries(data).map(([k, v]) => `${k}: ${v}`).join(', ')} }` : 'null';
      return `return createSuccessResponse(${dataStr}, ${message || 'null'});`;
    }
  );

  // パターン2: { success: false, ... } の形式を変換
  content = content.replace(
    /return\s*\{\s*success:\s*false,\s*([^}]+)\s*\};?/g,
    (match, rest) => {
      // restから個別のプロパティを抽出
      const props = rest.split(',').map(p => p.trim()).filter(p => p);
      const data = {};
      let message = null;
      let error = null;
      
      props.forEach(prop => {
        const [key, value] = prop.split(':').map(s => s.trim());
        if (key === 'message') {
          message = value;
        } else if (key === 'error') {
          error = value;
        } else {
          data[key] = value;
        }
      });
      
      const errorStr = error || message || '"エラーが発生しました"';
      const messageStr = message || 'null';
      const dataStr = Object.keys(data).length > 0 ? `{ ${Object.entries(data).map(([k, v]) => `${k}: ${v}`).join(', ')} }` : 'null';
      
      return `return createErrorResponse(${errorStr}, ${messageStr}, ${dataStr});`;
    }
  );

  // その他の簡単なパターン
  content = content.replace(
    /return\s*\{\s*success:\s*true\s*\};?/g,
    'return createSuccessResponse(null, null);'
  );

  content = content.replace(
    /return\s*\{\s*success:\s*false\s*\};?/g,
    'return createErrorResponse("エラーが発生しました", null, null);'
  );

  return content;
}

function processFile(filePath) {
  console.log(`Processing: ${filePath}`);
  
  try {
    const content = fs.readFileSync(filePath, 'utf-8');
    const convertedContent = convertResponseFormat(content);
    
    if (content !== convertedContent) {
      fs.writeFileSync(filePath, convertedContent);
      console.log(`✅ Updated: ${filePath}`);
    } else {
      console.log(`✓ No changes needed: ${filePath}`);
    }
  } catch (error) {
    console.error(`❌ Error processing ${filePath}:`, error.message);
  }
}

// 実行
const srcDir = path.join(__dirname, '../src');
const gasFiles = fs.readdirSync(srcDir)
  .filter(file => file.endsWith('.gs'))
  .map(file => path.join(srcDir, file));

console.log('🔧 レスポンス形式統一変換開始...\n');

gasFiles.forEach(processFile);

console.log('\n✅ レスポンス形式統一変換完了!');